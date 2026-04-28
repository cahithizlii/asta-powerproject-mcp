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
