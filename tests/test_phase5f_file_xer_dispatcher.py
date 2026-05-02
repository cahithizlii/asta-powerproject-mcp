"""Test Phase 5f T113: msproject_file dispatcher routes .xer end-to-end."""
import asyncio
import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_file


def _run(coro):
    return asyncio.run(coro)


def _call(action, **kw):
    raw = _run(msproject_file({"action": action, **kw}))
    return json.loads(raw)


def test_dispatcher_read_tasks_xer(sample_cau_xer):
    p = _call("read_tasks", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 6


def test_dispatcher_read_links_xer(sample_cau_xer):
    p = _call("read_links", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 5


def test_dispatcher_read_resources_xer(sample_cau_xer):
    p = _call("read_resources", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 4


def test_dispatcher_read_assignments_xer(sample_cau_xer):
    p = _call("read_assignments", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 7


def test_dispatcher_read_calendars_xer(sample_cau_xer):
    p = _call("read_calendars", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["count"] == 1


def test_dispatcher_read_baselines_xer(sample_cau_xer):
    p = _call("read_baselines", file_path=sample_cau_xer, baseline_number=0)
    assert p["status"] == "ok"
    assert len(p["tasks"]) == 6


def test_dispatcher_read_progress_xer(sample_cau_xer):
    p = _call("read_progress", file_path=sample_cau_xer)
    assert p["status"] == "ok"
    assert p["status_date"] == "2026-05-01"
