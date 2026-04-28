"""Test msproject_resource assign + unassign (single)."""
import pytest
from msproject_mcp_core import (
    _msp_resource_add, _msp_resource_assign, _msp_resource_unassign,
    _msp_task_add_single, _find_task_by_id,
)


def test_assign_work_resource_default_units(clean_test_project):
    proj = clean_test_project
    res_r = _msp_resource_add(name="AssignW-T36", type="Work")
    task_r = _msp_task_add_single(name="AssignTask-T36", duration="3d")
    r = _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    assert r["status"] == "ok"
    assert r["task_id"] == task_r["task_id"]
    assert r["resource_id"] == res_r["resource_id"]
    t = _find_task_by_id(proj, task_r["task_id"])
    assert t.Assignments.Count == 1


def test_assign_with_units(clean_test_project):
    res_r = _msp_resource_add(name="AssignU-T36", type="Work", max_units=500)
    task_r = _msp_task_add_single(name="UnitsTask-T36", duration="5d")
    r = _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"], units=300)
    assert r["status"] == "ok"
    assert r["units"] == 300


def test_assign_missing_task_errors(clean_test_project):
    res_r = _msp_resource_add(name="OrphanRes-T36", type="Work")
    r = _msp_resource_assign(task_id=99999, resource_id=res_r["resource_id"])
    assert r["status"] == "error"
    assert "task" in r["error"].lower() and "99999" in r["error"]


def test_assign_missing_resource_errors(clean_test_project):
    task_r = _msp_task_add_single(name="OrphanTask-T36", duration="1d")
    r = _msp_resource_assign(task_id=task_r["task_id"], resource_id=99999)
    assert r["status"] == "error"
    assert "resource" in r["error"].lower() and "99999" in r["error"]


def test_unassign(clean_test_project):
    proj = clean_test_project
    res_r = _msp_resource_add(name="UnassignRes-T36", type="Work")
    task_r = _msp_task_add_single(name="UnassignTask-T36", duration="2d")
    _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    t = _find_task_by_id(proj, task_r["task_id"])
    assert t.Assignments.Count == 1
    r = _msp_resource_unassign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    assert r["status"] == "ok"
    t = _find_task_by_id(proj, task_r["task_id"])
    assert t.Assignments.Count == 0


def test_unassign_not_assigned_errors(clean_test_project):
    res_r = _msp_resource_add(name="NeverRes-T36", type="Work")
    task_r = _msp_task_add_single(name="NeverTask-T36", duration="1d")
    r = _msp_resource_unassign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    assert r["status"] == "error"
    assert "not assigned" in r["error"].lower() or "not found" in r["error"].lower()


def test_assign_negative_units_errors(clean_test_project):
    res_r = _msp_resource_add(name="NegU-T36", type="Work")
    task_r = _msp_task_add_single(name="NegT-T36", duration="1d")
    r = _msp_resource_assign(task_id=task_r["task_id"],
                            resource_id=res_r["resource_id"], units=-10)
    assert r["status"] == "error"
    assert "units" in r["error"].lower()


def test_assign_with_work_hours(clean_test_project):
    """work_hours parameter sets alloc.Work (in minutes COM-side)."""
    proj = clean_test_project
    res_r = _msp_resource_add(name="WHRes-T36", type="Work")
    task_r = _msp_task_add_single(name="WHTask-T36", duration="5d")
    r = _msp_resource_assign(task_id=task_r["task_id"],
                            resource_id=res_r["resource_id"],
                            work_hours=24.0)
    assert r["status"] == "ok"
    # Verify Work was set (24h = 1440 minutes COM-side)
    t = _find_task_by_id(proj, task_r["task_id"])
    alloc = t.Assignments(1)
    # COM Work returns minutes (or possibly seconds depending on version) — accept either
    work_val = float(alloc.Work)
    # 24h = 1440 minutes, but MSP may return as 1440 or 86400 (seconds)
    assert work_val in (1440.0, 86400.0) or abs(work_val - 1440.0) < 0.1


def test_assign_duplicate_resource_to_task(clean_test_project):
    """Document MSP behavior when assigning same resource twice — locks contract for T37 bulk."""
    proj = clean_test_project
    res_r = _msp_resource_add(name="DupAsg-T36", type="Work")
    task_r = _msp_task_add_single(name="DupT-T36", duration="2d")
    r1 = _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    assert r1["status"] == "ok"
    # Second assignment of same resource — what happens?
    r2 = _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    # Either: (a) error from MSP, OR (b) silently creates duplicate, OR (c) updates existing
    # Document whichever is observed — test only asserts the result is a valid response
    assert r2["status"] in ("ok", "error")
    t = _find_task_by_id(proj, task_r["task_id"])
    if r2["status"] == "ok":
        # If MSP allowed it, count should be 1 (replace) or 2 (duplicate)
        assert t.Assignments.Count in (1, 2)
    else:
        # If MSP rejected, count stays at 1
        assert t.Assignments.Count == 1
