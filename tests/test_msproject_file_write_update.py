"""Test msproject_file write actions: update_task + save_as."""
import os
import shutil
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _msp_file_update_task, _msp_file_save_as, _msp_file_read_tasks,
)

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SOURCE_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


@pytest.fixture
def writable_xml(tmp_path):
    dst = tmp_path / "writable.xml"
    shutil.copy(SOURCE_XML, dst)
    return str(dst)


def test_update_task_duration(writable_xml):
    """Update T1 (id=1) duration 1d -> 5d, re-read confirms 40h."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={"duration": "5d"})
    assert r["status"] == "ok"
    r2 = _msp_file_read_tasks(file_path=writable_xml)
    t1 = next(t for t in r2["tasks"] if t["id"] == 1)
    assert t1["duration_h"] == 40.0


def test_update_task_missing_id(writable_xml):
    r = _msp_file_update_task(file_path=writable_xml, task_id=99999,
                              fields={"duration": "1d"})
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_update_task_invalid_file():
    r = _msp_file_update_task(file_path="/nonexistent.xml", task_id=1,
                              fields={"duration": "1d"})
    assert r["status"] == "error"


def test_update_task_mpp_rejected(tmp_path):
    fake_mpp = tmp_path / "test.mpp"
    fake_mpp.write_bytes(b"\x00" * 100)
    r = _msp_file_update_task(file_path=str(fake_mpp), task_id=1,
                              fields={"duration": "5d"})
    assert r["status"] == "error"
    msg = r["error"].lower()
    assert ".mpp" in msg or "binary" in msg or "not supported" in msg


# === Phase 9.3 — update_task baseline awareness ===

def test_update_task_baseline_only(writable_xml):
    """baseline_* fields routed to write_baseline (no schedule update)."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={
                                  "baseline_start": "2026-09-01T08:00:00",
                                  "baseline_finish": "2026-09-30T17:00:00",
                                  "baseline_duration_h": 200.0,
                                  "baseline_work_h": 120.0,
                              })
    assert r["status"] == "ok"
    assert r["schedule_updated"] is False
    assert r["baseline_written"] == 1
    # Verify the baseline persisted
    from mspdi_parser import MspdiProject
    proj = MspdiProject(writable_xml)
    bls = proj.read_baselines(0)
    assert any(b["task_id"] == 1
               and b["baseline_start"] == "2026-09-01T08:00:00"
               for b in bls)


def test_update_task_mixed_schedule_and_baseline(writable_xml):
    """Single call updates both schedule fields and baseline fields."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={
                                  "duration": "10d",
                                  "baseline_start": "2026-10-01T08:00:00",
                                  "baseline_finish": "2026-10-31T17:00:00",
                              })
    assert r["status"] == "ok"
    assert r["schedule_updated"] is True
    assert r["baseline_written"] == 1
    # Schedule check
    r2 = _msp_file_read_tasks(file_path=writable_xml)
    t1 = next(t for t in r2["tasks"] if t["id"] == 1)
    assert t1["duration_h"] == 80.0  # 10d * 8h
    # Baseline check
    from mspdi_parser import MspdiProject
    proj = MspdiProject(writable_xml)
    bls = proj.read_baselines(0)
    assert any(b["task_id"] == 1
               and b["baseline_finish"] == "2026-10-31T17:00:00"
               for b in bls)


def test_update_task_unknown_field_error_lists_baseline_keys(writable_xml):
    """Error message for unknown field must include baseline_* in valid list."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={"not_a_field": "x"})
    assert r["status"] == "error"
    assert "baseline_start" in r["error"]


def test_update_task_baseline_unknown_task_id(writable_xml):
    """Baseline-only update on missing task_id returns clear error."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=99_999_999,
                              fields={"baseline_start": "2026-01-01T08:00:00"})
    assert r["status"] == "error"


def test_update_task_baseline_number_param(writable_xml):
    """baseline_number=1 writes to a numbered baseline, not primary."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={"baseline_start": "2026-11-01T08:00:00"},
                              baseline_number=1)
    assert r["status"] == "ok"
    assert r["baseline_written"] == 1
    from mspdi_parser import MspdiProject
    proj = MspdiProject(writable_xml)
    # Number 0 should still be empty for this field; number 1 has the data
    bls0 = proj.read_baselines(0)
    bls1 = proj.read_baselines(1)
    assert not any(b["task_id"] == 1
                   and b["baseline_start"] == "2026-11-01T08:00:00"
                   for b in bls0)
    assert any(b["task_id"] == 1
               and b["baseline_start"] == "2026-11-01T08:00:00"
               for b in bls1)


def test_update_task_schedule_only_response_shape(writable_xml):
    """Phase 9.3 adds schedule_updated/baseline_written fields to response."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={"duration": "3d"})
    assert r["status"] == "ok"
    assert r["schedule_updated"] is True
    assert r["baseline_written"] == 0


# === Phase 10.1 — baseline read-back ===

def test_update_task_baseline_after_field_populated(writable_xml):
    """Phase 10.1: baseline_after returns the persisted baseline values."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={
                                  "baseline_start": "2026-12-01T08:00:00",
                                  "baseline_finish": "2026-12-31T17:00:00",
                                  "baseline_duration_h": 240.0,
                                  "baseline_work_h": 160.0,
                              })
    assert r["status"] == "ok"
    assert r["baseline_after"] is not None
    bl = r["baseline_after"]
    assert bl["task_id"] == 1
    assert bl["baseline_start"] == "2026-12-01T08:00:00"
    assert bl["baseline_finish"] == "2026-12-31T17:00:00"
    assert abs(bl["baseline_duration_h"] - 240.0) < 0.5
    assert abs(bl["baseline_work_h"] - 160.0) < 0.5


def test_update_task_baseline_after_none_when_schedule_only(writable_xml):
    """Phase 10.1: schedule-only updates leave baseline_after as None."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={"duration": "2d"})
    assert r["status"] == "ok"
    assert r["baseline_after"] is None


def test_save_as_xml(tmp_path):
    """Save fixture to a new path; output file exists and has size."""
    dst = tmp_path / "renamed.xml"
    r = _msp_file_save_as(file_path=SOURCE_XML, output_path=str(dst))
    assert r["status"] == "ok"
    assert os.path.exists(str(dst))
    assert r["output_path"] == str(dst)
    assert r["size_bytes"] > 0


def test_save_as_invalid_output_extension(tmp_path):
    """output_path must end .xml or .mspdi."""
    dst = tmp_path / "bad.txt"
    r = _msp_file_save_as(file_path=SOURCE_XML, output_path=str(dst))
    assert r["status"] == "error"
    assert "extension" in r["error"].lower() or "xml" in r["error"].lower()


def test_save_as_invalid_input_file():
    r = _msp_file_save_as(file_path="/nonexistent.xml",
                          output_path="/tmp/out.xml")
    assert r["status"] == "error"
