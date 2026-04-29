"""Test FastMCP msproject_progress dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_progress


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_set_task(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="DispST-T64", duration="3d")
    r = _run(msproject_progress({"action": "set_task_progress",
                                 "task_id": add_r["task_id"],
                                 "percent_complete": 50}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "percent_complete" in p["changes"]


def test_dispatcher_get_task(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="DispGT-T64", duration="2d")
    r = _run(msproject_progress({"action": "get_task_progress",
                                 "task_id": add_r["task_id"]}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "progress" in p


def test_dispatcher_summary(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    _msp_task_add_single(name="DispSumT-T64", duration="2d")
    r = _run(msproject_progress({"action": "summary"}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "project" in p
    assert "bac_h" in p["project"]


def test_dispatcher_status_date(clean_test_project):
    r = _run(msproject_progress({"action": "set_status_date",
                                 "status_date": "2026-04-29"}))
    p = json.loads(r)
    assert p["status"] == "ok"


def test_dispatcher_bulk(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    ids = [_msp_task_add_single(name=f"DispBlk{i}-T64", duration="2d")["task_id"]
           for i in range(3)]
    r = _run(msproject_progress({"action": "bulk_progress_update",
                                 "items": [{"task_id": tid, "percent_complete": 25}
                                           for tid in ids]}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert p["count"] == 3


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_progress({"action": "nonsense"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]
