"""Test msproject_progress get_assignment_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_get_assignments, _msp_progress_set_assignment,
    _msp_task_add_single, _msp_resource_add, _msp_resource_assign,
)


def test_get_empty_assignments(clean_test_project):
    """Task with no assignments → empty list."""
    add_r = _msp_task_add_single(name="EmptyAsgT-T56", duration="2d")
    r = _msp_progress_get_assignments(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["assignments"] == []


def test_get_one_assignment(clean_test_project):
    """1 resource assigned → 1-element list."""
    add_t = _msp_task_add_single(name="OneAsgT-T56", duration="3d")
    add_r = _msp_resource_add(name="X-T56", type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    r = _msp_progress_get_assignments(task_id=add_t["task_id"])
    assert r["status"] == "ok"
    assert len(r["assignments"]) == 1
    a = r["assignments"][0]
    assert a["resource_id"] == add_r["resource_id"]
    assert a["resource_name"] == "X-T56"
    assert "actual_work_h" in a
    assert "percent_work_complete" in a
    assert "units" in a


def test_get_three_assignments_after_writes(clean_test_project):
    """3 resources, write actuals to each → all 3 returned with values."""
    add_t = _msp_task_add_single(name="3AsgT-T56", duration="5d")
    rids = []
    for nm in ("COW-T56", "STL-T56", "MSN-T56"):
        ar = _msp_resource_add(name=nm, type="Work", max_units=100)
        _msp_resource_assign(task_id=add_t["task_id"],
                             resource_id=ar["resource_id"])
        rids.append(ar["resource_id"])
    # Write different hours to each
    _msp_progress_set_assignment(task_id=add_t["task_id"],
                                 resource_id=rids[0], actual_work_h=24)
    _msp_progress_set_assignment(task_id=add_t["task_id"],
                                 resource_id=rids[1], actual_work_h=16)
    _msp_progress_set_assignment(task_id=add_t["task_id"],
                                 resource_id=rids[2], actual_work_h=8)
    r = _msp_progress_get_assignments(task_id=add_t["task_id"])
    assert r["status"] == "ok"
    assert len(r["assignments"]) == 3
    by_rid = {a["resource_id"]: a for a in r["assignments"]}
    assert by_rid[rids[0]]["actual_work_h"] == 24
    assert by_rid[rids[1]]["actual_work_h"] == 16
    assert by_rid[rids[2]]["actual_work_h"] == 8


def test_get_missing_task(clean_test_project):
    r = _msp_progress_get_assignments(task_id=99999)
    assert r["status"] == "error"
