"""Test Phase 5d T107 FastMCP dispatcher (msproject_xer)."""
import asyncio
import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_xer


def _run(coro):
    return asyncio.run(coro)


def _call(action, **kw):
    raw = _run(msproject_xer({"action": action, **kw}))
    return json.loads(raw)


def test_dispatcher_read_tasks(sample_cau_xer):
    p = _call("read_tasks", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 6


def test_dispatcher_read_links(sample_cau_xer):
    p = _call("read_links", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 5


def test_dispatcher_read_resources(sample_cau_xer):
    p = _call("read_resources", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 4


def test_dispatcher_read_assignments(sample_cau_xer):
    p = _call("read_assignments", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 7


def test_dispatcher_read_calendars(sample_cau_xer):
    p = _call("read_calendars", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 1
    assert p["calendars"][0]["day_hr_cnt"] == 9.0


def test_dispatcher_read_progress(sample_cau_xer):
    p = _call("read_progress", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["status_date"] == "2026-05-01"


def test_dispatcher_unknown_action(sample_cau_xer):
    p = _call("nonsense", file_path=sample_cau_xer)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_missing_file_path():
    p = _call("read_tasks")
    assert p["status"] == "error"


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================


def test_dispatcher_read_tasks_nonexistent_xer_returns_error(tmp_path):
    """Non-existent .xer path → error."""
    p = _call("read_tasks", file_path=str(tmp_path / "missing.xer"))
    assert p["status"] == "error"
    assert p["error"]


def test_dispatcher_read_tasks_directory_path_returns_error(tmp_path):
    """Passing a directory (not file) → error."""
    p = _call("read_tasks", file_path=str(tmp_path))
    assert p["status"] == "error"


def test_dispatcher_unknown_action_lists_valid_actions(sample_cau_xer):
    """Unknown action error message lists valid actions."""
    p = _call("invented_action", file_path=sample_cau_xer)
    assert p["status"] == "error"
    assert "read_tasks" in p["error"]


def test_dispatcher_read_links_nonexistent_returns_error(tmp_path):
    """Non-existent file → error from XerFile."""
    p = _call("read_links", file_path=str(tmp_path / "ghost.xer"))
    assert p["status"] == "error"


def test_dispatcher_read_calendars_nonexistent_returns_error(tmp_path):
    """read_calendars on missing file → error."""
    p = _call("read_calendars", file_path=str(tmp_path / "no.xer"))
    assert p["status"] == "error"
