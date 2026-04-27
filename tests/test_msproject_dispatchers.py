"""Test FastMCP dispatcher tools."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_task, msproject_link, msproject_schedule


def _run(coro):
    return asyncio.run(coro)


def test_task_dispatcher_add(msproject_app):
    result = _run(msproject_task({"action": "add", "name": "Disp Test", "duration": "2d"}))
    parsed = json.loads(result)
    assert parsed["status"] == "ok"
    # Cleanup
    _run(msproject_task({"action": "delete", "task_id": parsed["task_id"]}))


def test_task_dispatcher_invalid_action(msproject_app):
    result = _run(msproject_task({"action": "nonsense"}))
    parsed = json.loads(result)
    assert parsed["status"] == "error"
    assert "Unknown action" in parsed["error"]


def test_task_dispatcher_list(msproject_app):
    add_r = _run(msproject_task({"action": "add", "name": "ListTest", "duration": "1d"}))
    add_p = json.loads(add_r)
    list_r = _run(msproject_task({"action": "list"}))
    list_p = json.loads(list_r)
    assert list_p["status"] == "ok"
    names = [t["name"] for t in list_p["tasks"]]
    assert "ListTest" in names
    _run(msproject_task({"action": "delete", "task_id": add_p["task_id"]}))


def test_link_dispatcher_chain(msproject_app):
    a = json.loads(_run(msproject_task({"action": "add", "name": "LA", "duration": "1d"})))
    b = json.loads(_run(msproject_task({"action": "add", "name": "LB", "duration": "1d"})))
    c = json.loads(_run(msproject_task({"action": "add", "name": "LC", "duration": "1d"})))
    chain_r = _run(msproject_link({"action": "chain", "task_ids": [a["task_id"], b["task_id"], c["task_id"]]}))
    chain_p = json.loads(chain_r)
    assert chain_p["status"] == "ok"
    assert chain_p["links_added"] == 2
    _run(msproject_task({"action": "delete", "task_id": c["task_id"]}))
    _run(msproject_task({"action": "delete", "task_id": b["task_id"]}))
    _run(msproject_task({"action": "delete", "task_id": a["task_id"]}))


def test_schedule_dispatcher_reschedule(msproject_app):
    r = _run(msproject_schedule({"action": "reschedule"}))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"


def test_schedule_dispatcher_invalid_action(msproject_app):
    r = _run(msproject_schedule({"action": "fake"}))
    parsed = json.loads(r)
    assert parsed["status"] == "error"
