"""Test FastMCP msproject_file dispatcher (T74)."""
import asyncio
import json
import os
import shutil
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_file  # noqa: E402

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def _call(action, **params):
    raw = _run(msproject_file({"action": action, **params}))
    return json.loads(raw)


def test_dispatcher_read_tasks():
    p = _call("read_tasks", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "tasks" in p


def test_dispatcher_read_links():
    p = _call("read_links", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "links" in p


def test_dispatcher_read_resources():
    p = _call("read_resources", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "resources" in p


def test_dispatcher_read_assignments():
    p = _call("read_assignments", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "assignments" in p


def test_dispatcher_read_calendars():
    p = _call("read_calendars", file_path=MSP_XML)
    assert p["status"] == "ok"


def test_dispatcher_read_baselines():
    p = _call("read_baselines", file_path=MSP_XML)
    assert p["status"] == "ok"


def test_dispatcher_read_progress():
    p = _call("read_progress", file_path=MSP_XML)
    assert p["status"] == "ok"


def test_dispatcher_query():
    p = _call("query", file_path=MSP_XML, expression="duration_h > 8")
    assert p["status"] == "ok"
    assert p["count"] >= 1


def test_dispatcher_save_as(tmp_path):
    dst = tmp_path / "renamed.xml"
    p = _call("save_as", file_path=MSP_XML, output_path=str(dst))
    assert p["status"] == "ok"
    assert os.path.exists(str(dst))


def test_dispatcher_write_chain_add_tasks(tmp_path):
    """add_tasks via dispatcher round-trips."""
    dst = tmp_path / "writable.xml"
    shutil.copy(MSP_XML, str(dst))
    p = _call("add_tasks", file_path=str(dst),
              items=[{"name": "DispatcherT", "duration": "1d"}])
    assert p["status"] == "ok"
    assert p["count"] == 1


def test_dispatcher_unknown_action():
    p = _call("nonsense", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_missing_file_path():
    """No file_path → error (TypeError caught and translated)."""
    p = _call("read_tasks")
    assert p["status"] == "error"
