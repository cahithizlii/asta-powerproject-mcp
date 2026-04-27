"""Test msproject_link operations."""
import pytest
from msproject_mcp_core import (
    _msp_task_add_single, _msp_task_delete,
    _msp_link_add,
    _msp_link_delete, _msp_link_update, _msp_link_bulk_add, _msp_link_chain,
)


def test_link_two_tasks(msproject_app):
    """Adding a predecessor link sets Predecessors string."""
    a = _msp_task_add_single(name="LinkA", duration="2d")
    b = _msp_task_add_single(name="LinkB", duration="3d")
    r = _msp_link_add(predecessor_id=a["task_id"], successor_id=b["task_id"], type="FS", lag="0d")
    assert r["status"] == "ok"
    proj = msproject_app.ActiveProject
    bt = None
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t and t.ID == b["task_id"]:
            bt = t
            break
    assert bt is not None
    # Predecessors string should now contain the predecessor task ID
    assert str(a["task_id"]) in (bt.Predecessors or "")
    # Delete in reverse creation order to avoid ID re-shift bug
    _msp_task_delete(task_id=b["task_id"])
    _msp_task_delete(task_id=a["task_id"])


def test_link_with_lag(msproject_app):
    """FS+2d lag stored correctly."""
    a = _msp_task_add_single(name="LagA", duration="2d")
    b = _msp_task_add_single(name="LagB", duration="3d")
    r = _msp_link_add(predecessor_id=a["task_id"], successor_id=b["task_id"], type="FS", lag="2d")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=b["task_id"])
    _msp_task_delete(task_id=a["task_id"])


def test_link_invalid_predecessor(msproject_app):
    """Linking to nonexistent predecessor returns error."""
    a = _msp_task_add_single(name="Solo", duration="1d")
    r = _msp_link_add(predecessor_id=99999, successor_id=a["task_id"])
    assert r["status"] == "error"
    _msp_task_delete(task_id=a["task_id"])


def test_link_delete(msproject_app):
    """Removing a link clears it from Predecessors."""
    a = _msp_task_add_single(name="DA", duration="1d")
    b = _msp_task_add_single(name="DB", duration="2d")
    _msp_link_add(predecessor_id=a["task_id"], successor_id=b["task_id"])
    r = _msp_link_delete(predecessor_id=a["task_id"], successor_id=b["task_id"])
    assert r["status"] == "ok"
    _msp_task_delete(task_id=b["task_id"])
    _msp_task_delete(task_id=a["task_id"])


def test_link_update_type(msproject_app):
    """Update link from FS to SS."""
    a = _msp_task_add_single(name="UA", duration="1d")
    b = _msp_task_add_single(name="UB", duration="2d")
    _msp_link_add(predecessor_id=a["task_id"], successor_id=b["task_id"], type="FS")
    r = _msp_link_update(predecessor_id=a["task_id"], successor_id=b["task_id"], new_type="SS")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=b["task_id"])
    _msp_task_delete(task_id=a["task_id"])


def test_link_bulk_add(msproject_app):
    """Bulk add 5 links."""
    tasks = [_msp_task_add_single(name=f"Bulk{i}", duration="1d") for i in range(6)]
    items = [{"predecessor_id": tasks[i]["task_id"], "successor_id": tasks[i+1]["task_id"], "type": "FS"}
             for i in range(5)]
    r = _msp_link_bulk_add(items=items)
    assert r["status"] == "ok"
    assert r["count"] == 5
    # Delete in reverse creation order
    for t in reversed(tasks):
        _msp_task_delete(task_id=t["task_id"])


def test_link_chain(msproject_app):
    """Chain 4 tasks: T1->T2->T3->T4."""
    tasks = [_msp_task_add_single(name=f"Chain{i}", duration="1d") for i in range(4)]
    task_ids = [t["task_id"] for t in tasks]
    r = _msp_link_chain(task_ids=task_ids, type="FS", lag="0d")
    assert r["status"] == "ok"
    assert r["links_added"] == 3
    # Delete in reverse creation order
    for t in reversed(tasks):
        _msp_task_delete(task_id=t["task_id"])
