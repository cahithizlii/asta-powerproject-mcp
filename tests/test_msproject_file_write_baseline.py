"""Phase 8.2 — msproject_file write_baseline action integration tests.

Verifies that the file MCP exposes Phase 6.3 MspdiProject.write_baseline
through a new `write_baseline` action with proper validation:
- Requires file_path, output_path, baseline_data
- .xml/.mspdi only (.xer/.mpp rejected)
- Roundtrip: write -> read_baselines confirms persisted values
- Listed in unknown action error message
"""
import asyncio
import json
import os
import shutil
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_file
from mspdi_parser import MspdiProject


SOURCE_XML = os.path.join(os.path.dirname(__file__),
                           "fixtures", "sample_msp.xml")


def _call(action, **kw):
    raw = asyncio.run(msproject_file({"action": action, **kw}))
    return json.loads(raw)


def _copy_to_tmp(name: str) -> str:
    tmp = os.path.join(tempfile.gettempdir(), name)
    shutil.copy(SOURCE_XML, tmp)
    return tmp


# === validation paths ===

def test_write_baseline_missing_file_path_returns_error():
    r = _call("write_baseline", baseline_number=0,
              baseline_data=[{"task_uid": 1}],
              output_path="/tmp/out.xml")
    assert r["status"] == "error"
    assert "file_path" in r["error"]


def test_write_baseline_missing_output_path_returns_error():
    src = _copy_to_tmp("p82_no_out.xml")
    try:
        r = _call("write_baseline", file_path=src, baseline_number=0,
                  baseline_data=[{"task_uid": 1}])
        assert r["status"] == "error"
        assert "output_path" in r["error"]
    finally:
        os.remove(src)


def test_write_baseline_empty_data_returns_error():
    src = _copy_to_tmp("p82_empty.xml")
    out = src.replace(".xml", "_out.xml")
    try:
        r = _call("write_baseline", file_path=src, baseline_number=0,
                  baseline_data=[], output_path=out)
        assert r["status"] == "error"
        assert "baseline_data" in r["error"]
    finally:
        for p in (src,):
            try:
                os.remove(p)
            except OSError:
                pass


def test_write_baseline_xer_extension_rejected():
    """write_baseline supports .xml/.mspdi only."""
    r = _call("write_baseline", file_path="/some/file.xer",
              baseline_number=0,
              baseline_data=[{"task_uid": 1}],
              output_path="/tmp/out.xml")
    assert r["status"] == "error"
    assert ".mpp" in r["error"] or ".xer" in r["error"]


def test_write_baseline_bad_output_extension_rejected():
    src = _copy_to_tmp("p82_badext.xml")
    try:
        r = _call("write_baseline", file_path=src, baseline_number=0,
                  baseline_data=[{"task_uid": 1}],
                  output_path="/tmp/out.txt")
        assert r["status"] == "error"
        assert ".xml" in r["error"] or ".mspdi" in r["error"]
    finally:
        os.remove(src)


# === happy path: roundtrip ===

def test_write_baseline_roundtrip_via_dispatcher():
    """write_baseline dispatcher -> reload via MspdiProject -> values match."""
    src = _copy_to_tmp("p82_roundtrip.xml")
    out = src.replace(".xml", "_baselined.xml")
    try:
        # Pick a real UID from the source
        proj = MspdiProject(src)
        uid = next(iter(proj._task_elems.keys()))
        del proj
        r = _call("write_baseline",
                  file_path=src, baseline_number=0,
                  baseline_data=[{
                      "task_uid": uid,
                      "baseline_start": "2026-06-01T08:00:00",
                      "baseline_finish": "2026-06-30T17:00:00",
                      "baseline_duration_h": 200.0,
                      "baseline_work_h": 120.0,
                  }],
                  output_path=out)
        assert r["status"] == "ok"
        assert r["tasks_written"] == 1
        # mspdi_parser save() normalizes backslashes -> compare normalized paths
        assert os.path.normpath(r["output_path"]) == os.path.normpath(out)
        assert os.path.exists(out)
        # Reload + read_baselines
        reloaded = MspdiProject(out)
        bls = reloaded.read_baselines(0)
        assert len(bls) == 1
        assert bls[0]["task_uid"] == uid
        assert bls[0]["baseline_start"] == "2026-06-01T08:00:00"
        assert bls[0]["baseline_finish"] == "2026-06-30T17:00:00"
    finally:
        for p in (src, out):
            try:
                os.remove(p)
            except OSError:
                pass


def test_write_baseline_unknown_uid_skipped_silently():
    """Phase 6.3 contract: unknown UIDs skipped, count reflects skipped."""
    src = _copy_to_tmp("p82_unk.xml")
    out = src.replace(".xml", "_unk_out.xml")
    try:
        r = _call("write_baseline",
                  file_path=src, baseline_number=0,
                  baseline_data=[{"task_uid": 999_999_999,
                                  "baseline_start": "2026-01-01T00:00:00"}],
                  output_path=out)
        assert r["status"] == "ok"
        assert r["tasks_written"] == 0
    finally:
        for p in (src, out):
            try:
                os.remove(p)
            except OSError:
                pass


# === Phase 9.1 — auto-sync integration ===

def test_write_baseline_includes_auto_imported_field():
    """Phase 9.1: response must include auto_imported (bool) and
    reschedule_ok fields, populated by _maybe_auto_sync(output_path)."""
    src = _copy_to_tmp("p91_autosync_field.xml")
    out = src.replace(".xml", "_out.xml")
    try:
        proj = MspdiProject(src)
        uid = next(iter(proj._task_elems.keys()))
        del proj
        r = _call("write_baseline", file_path=src, baseline_number=0,
                  baseline_data=[{
                      "task_uid": uid,
                      "baseline_start": "2026-07-01T08:00:00",
                      "baseline_finish": "2026-07-31T17:00:00",
                  }],
                  output_path=out)
        assert r["status"] == "ok"
        assert "auto_imported" in r
        assert "reschedule_ok" in r
        assert isinstance(r["auto_imported"], bool)
    finally:
        for p in (src, out):
            try:
                os.remove(p)
            except OSError:
                pass


def test_write_baseline_auto_sync_safe_when_no_match():
    """When output_path doesn't match any open MSP project,
    auto_imported should be False (no error, no crash)."""
    src = _copy_to_tmp("p91_no_match.xml")
    # Use a unique output path unlikely to match any open project
    out = os.path.join(tempfile.gettempdir(),
                       "p91_unique_no_msp_match_baseline.xml")
    try:
        proj = MspdiProject(src)
        uid = next(iter(proj._task_elems.keys()))
        del proj
        r = _call("write_baseline", file_path=src, baseline_number=0,
                  baseline_data=[{"task_uid": uid,
                                  "baseline_start": "2026-08-01T08:00:00"}],
                  output_path=out)
        assert r["status"] == "ok"
        # MSP not running OR output_path isn't open in MSP -> False
        assert r["auto_imported"] is False
    finally:
        for p in (src, out):
            try:
                os.remove(p)
            except OSError:
                pass


# === dispatcher action listing ===

def test_dispatcher_unknown_action_lists_write_baseline():
    """Error message for unknown action must include write_baseline."""
    r = _call("definitely_not_an_action")
    assert r["status"] == "error"
    assert "write_baseline" in r["error"]
