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


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142): TASK + LINK
# =============================================================================
# Most of these COM-gated negative tests run only when MS Project is open.
# They exercise validation paths in the dispatcher: bad task IDs, bad inputs,
# self-loops, bad link types. clean_test_project ensures isolation.


def test_task_dispatcher_delete_nonexistent_id_returns_error(clean_test_project):
    """Deleting task_id=-1 → error from helper (task not found)."""
    r = _run(msproject_task({"action": "delete", "task_id": -1}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert p["error"]


def test_task_dispatcher_get_nonexistent_id_returns_error(clean_test_project):
    """get with task_id=999999 (max-int-ish) → error."""
    r = _run(msproject_task({"action": "get", "task_id": 999999}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_task_dispatcher_update_nonexistent_id_returns_error(clean_test_project):
    """update with task_id=0 → error."""
    r = _run(msproject_task({"action": "update", "task_id": 0,
                             "name": "ShouldFail"}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_task_dispatcher_update_bad_start_date_returns_error(clean_test_project):
    """update with malformed start date → COM-rejected error."""
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="UpdBadDate", duration="1d")
    r = _run(msproject_task({"action": "update",
                             "task_id": add_r["task_id"],
                             "start": "not-a-real-date"}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_task_dispatcher_bulk_add_empty_items_returns_error(clean_test_project):
    """bulk_add with empty items list → error or count=0."""
    r = _run(msproject_task({"action": "bulk_add", "items": []}))
    p = json.loads(r)
    # Either error OR ok with count=0 — both acceptable; just check no crash
    assert p["status"] in ("ok", "error")
    if p["status"] == "ok":
        assert p.get("count", 0) == 0


# === LINK negative tests ===

def test_link_dispatcher_add_nonexistent_predecessor_returns_error(clean_test_project):
    """add link with predecessor_id=99999 → error."""
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="LinkDispNeg1", duration="1d")
    r = _run(msproject_link({"action": "add",
                             "predecessor_id": 99999,
                             "successor_id": add_r["task_id"]}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "predecessor" in p["error"].lower() or "not found" in p["error"].lower()


def test_link_dispatcher_add_nonexistent_successor_returns_error(clean_test_project):
    """add link with successor_id=99999 → error."""
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="LinkDispNeg2", duration="1d")
    r = _run(msproject_link({"action": "add",
                             "predecessor_id": add_r["task_id"],
                             "successor_id": 99999}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "successor" in p["error"].lower() or "not found" in p["error"].lower()


def test_link_dispatcher_delete_nonexistent_successor_returns_error(clean_test_project):
    """delete link with successor_id=99999 → error."""
    r = _run(msproject_link({"action": "delete",
                             "predecessor_id": 1,
                             "successor_id": 99999}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_link_dispatcher_chain_single_task_id_no_op(clean_test_project):
    """chain with only one task_id → no links (chain needs >=2)."""
    from msproject_mcp_core import _msp_task_add_single
    a = _msp_task_add_single(name="ChainNeg1", duration="1d")
    r = _run(msproject_link({"action": "chain",
                             "task_ids": [a["task_id"]]}))
    p = json.loads(r)
    # Either ok with 0 links or error — not a crash
    assert p["status"] in ("ok", "error")
    if p["status"] == "ok":
        assert p.get("links_added", 0) == 0


def test_link_dispatcher_chain_empty_task_ids_no_op(clean_test_project):
    """chain with empty task_ids → 0 links or error."""
    r = _run(msproject_link({"action": "chain", "task_ids": []}))
    p = json.loads(r)
    assert p["status"] in ("ok", "error")
    if p["status"] == "ok":
        assert p.get("links_added", 0) == 0
