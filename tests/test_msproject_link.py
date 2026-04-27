"""Test msproject_link operations."""
import pytest
from msproject_mcp_core import (
    _msp_task_add_single, _msp_task_delete,
    _msp_link_add,
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
    _msp_task_delete(task_id=a["task_id"])
    _msp_task_delete(task_id=b["task_id"])


def test_link_with_lag(msproject_app):
    """FS+2d lag stored correctly."""
    a = _msp_task_add_single(name="LagA", duration="2d")
    b = _msp_task_add_single(name="LagB", duration="3d")
    r = _msp_link_add(predecessor_id=a["task_id"], successor_id=b["task_id"], type="FS", lag="2d")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=a["task_id"])
    _msp_task_delete(task_id=b["task_id"])


def test_link_invalid_predecessor(msproject_app):
    """Linking to nonexistent predecessor returns error."""
    a = _msp_task_add_single(name="Solo", duration="1d")
    r = _msp_link_add(predecessor_id=99999, successor_id=a["task_id"])
    assert r["status"] == "error"
    _msp_task_delete(task_id=a["task_id"])
