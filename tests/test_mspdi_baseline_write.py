"""Phase 6.3 T118-T119 — MspdiProject baseline read/write/roundtrip tests.

Verifies:
- read_baselines returns empty list when source has no Baseline elements
- write_baseline creates new <Baseline Number=N> element on the task
- write_baseline updates existing element (no duplication)
- save() + reload() preserves baseline values (lossless roundtrip)
- unknown task UID is skipped silently
- task count unchanged after save (no data loss)
"""
import os
import shutil
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from mspdi_parser import MspdiProject


SOURCE_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _copy_to_tmp(name: str) -> str:
    """Copy source MSPDI to a writable tmp path."""
    tmp = os.path.join(tempfile.gettempdir(), name)
    shutil.copy(SOURCE_XML, tmp)
    return tmp


def test_read_baselines_empty_when_source_has_none():
    """Sample MSPDI has zero Baseline elements — read returns []."""
    proj = MspdiProject(SOURCE_XML)
    bls = proj.read_baselines(baseline_number=0)
    assert bls == []


def test_write_baseline_creates_element():
    """First call to write_baseline creates a new <Baseline> on the task."""
    src = _copy_to_tmp("p63_create.xml")
    try:
        proj = MspdiProject(src)
        # Pick a known UID from the source
        first_uid = next(iter(proj._task_elems.keys()))
        n = proj.write_baseline(0, [{
            "task_uid": first_uid,
            "baseline_start": "2026-01-01T08:00:00",
            "baseline_finish": "2026-01-31T17:00:00",
            "baseline_duration_h": 240.0,
            "baseline_work_h": 160.0,
        }])
        assert n == 1
        bls = proj.read_baselines(0)
        assert len(bls) == 1
        assert bls[0]["task_uid"] == first_uid
    finally:
        os.remove(src)


def test_write_baseline_updates_existing_no_duplicate():
    """Second write_baseline call for same task updates in place."""
    src = _copy_to_tmp("p63_update.xml")
    try:
        proj = MspdiProject(src)
        uid = next(iter(proj._task_elems.keys()))
        proj.write_baseline(0, [{
            "task_uid": uid,
            "baseline_start": "2026-01-01T08:00:00",
            "baseline_finish": "2026-01-31T17:00:00",
            "baseline_duration_h": 240.0,
            "baseline_work_h": 160.0,
        }])
        # Update with new finish
        proj.write_baseline(0, [{
            "task_uid": uid,
            "baseline_start": "2026-01-01T08:00:00",
            "baseline_finish": "2026-02-15T17:00:00",
            "baseline_duration_h": 360.0,
            "baseline_work_h": 240.0,
        }])
        bls = proj.read_baselines(0)
        # Still exactly 1 baseline for the task with number=0
        assert len(bls) == 1
        assert bls[0]["baseline_finish"] == "2026-02-15T17:00:00"
        assert bls[0]["baseline_work_h"] == 240.0
    finally:
        os.remove(src)


def test_write_baseline_save_reload_roundtrip():
    """write_baseline -> save -> new MspdiProject(out) -> read same data."""
    src = _copy_to_tmp("p63_roundtrip.xml")
    out = src.replace(".xml", "_baselined.xml")
    try:
        proj = MspdiProject(src)
        uid = next(iter(proj._task_elems.keys()))
        proj.write_baseline(0, [{
            "task_uid": uid,
            "baseline_start": "2026-03-01T08:00:00",
            "baseline_finish": "2026-03-31T17:00:00",
            "baseline_duration_h": 200.0,
            "baseline_work_h": 120.0,
        }])
        proj.save(out)
        reloaded = MspdiProject(out)
        bls = reloaded.read_baselines(0)
        assert len(bls) == 1
        assert bls[0]["task_uid"] == uid
        assert bls[0]["baseline_start"] == "2026-03-01T08:00:00"
        assert bls[0]["baseline_finish"] == "2026-03-31T17:00:00"
        # Duration/Work serialized as ISO 8601 — verify round-trip floats
        assert abs(bls[0]["baseline_duration_h"] - 200.0) < 0.5
        assert abs(bls[0]["baseline_work_h"] - 120.0) < 0.5
    finally:
        for p in (src, out):
            try:
                os.remove(p)
            except OSError:
                pass


def test_write_baseline_unknown_uid_skipped():
    """Tasks with non-existent UID are skipped silently."""
    src = _copy_to_tmp("p63_skip.xml")
    try:
        proj = MspdiProject(src)
        n = proj.write_baseline(0, [
            {"task_uid": 999_999_999, "baseline_start": "2026-01-01T00:00:00"},
        ])
        assert n == 0
        assert proj.read_baselines(0) == []
    finally:
        os.remove(src)


def test_save_preserves_task_count():
    """Save after baseline write must NOT remove tasks."""
    src = _copy_to_tmp("p63_count.xml")
    out = src.replace(".xml", "_saved.xml")
    try:
        proj = MspdiProject(src)
        original_count = len(proj._task_elems)
        if original_count > 0:
            uid = next(iter(proj._task_elems.keys()))
            proj.write_baseline(0, [{
                "task_uid": uid,
                "baseline_start": "2026-04-01T08:00:00",
                "baseline_finish": "2026-04-30T17:00:00",
                "baseline_duration_h": 100.0,
                "baseline_work_h": 80.0,
            }])
            proj.save(out)
            reloaded = MspdiProject(out)
            assert len(reloaded._task_elems) == original_count
    finally:
        for p in (src, out):
            try:
                os.remove(p)
            except OSError:
                pass


def test_write_baseline_multiple_numbers_independent():
    """Different baseline numbers stored separately on the same task."""
    src = _copy_to_tmp("p63_multi.xml")
    try:
        proj = MspdiProject(src)
        uid = next(iter(proj._task_elems.keys()))
        proj.write_baseline(0, [{
            "task_uid": uid,
            "baseline_start": "2026-01-01T08:00:00",
            "baseline_finish": "2026-01-31T17:00:00",
        }])
        proj.write_baseline(1, [{
            "task_uid": uid,
            "baseline_start": "2026-05-01T08:00:00",
            "baseline_finish": "2026-05-31T17:00:00",
        }])
        bls0 = proj.read_baselines(0)
        bls1 = proj.read_baselines(1)
        assert len(bls0) == 1 and len(bls1) == 1
        assert bls0[0]["baseline_start"] == "2026-01-01T08:00:00"
        assert bls1[0]["baseline_start"] == "2026-05-01T08:00:00"
    finally:
        os.remove(src)
