"""Test msproject_resource delete + list."""
import pytest
from msproject_mcp_core import (
    _msp_resource_add, _msp_resource_delete, _msp_resource_list,
    _find_resource_by_id, _parse_rate,
)


def test_delete_resource(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="ToDelete-T35", type="Work")
    initial = proj.Resources.Count
    r = _msp_resource_delete(resource_id=r1["resource_id"])
    assert r["status"] == "ok"
    assert r["deleted_id"] == r1["resource_id"]
    assert r["deleted_name"] == "ToDelete-T35"
    assert proj.Resources.Count == initial - 1


def test_delete_missing_errors(clean_test_project):
    r = _msp_resource_delete(resource_id=99999)
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_list_empty(clean_test_project):
    r = _msp_resource_list()
    assert r["status"] == "ok"
    assert r["count"] == 0
    assert r["resources"] == []


def test_list_mixed_types(clean_test_project):
    _msp_resource_add(name="W1-T35", type="Work", max_units=200, standard_rate=50)
    _msp_resource_add(name="M1-T35", type="Material", material_label="kg", standard_rate=2.5)
    _msp_resource_add(name="C1-T35", type="Cost")
    r = _msp_resource_list()
    assert r["status"] == "ok"
    assert r["count"] == 3
    types = {res["type"] for res in r["resources"]}
    assert types == {"Work", "Material", "Cost"}
    work = next(rr for rr in r["resources"] if rr["name"] == "W1-T35")
    assert work["max_units"] == 200.0
    assert abs(work["standard_rate"] - 50.0) < 0.01


def test_list_includes_assignment_count(clean_test_project):
    """Each entry has assignment_count (zero for unassigned)."""
    _msp_resource_add(name="UnassignedRes-T35", type="Work")
    r = _msp_resource_list()
    target = next(rr for rr in r["resources"] if rr["name"] == "UnassignedRes-T35")
    assert "assignment_count" in target
    assert target["assignment_count"] == 0


def test_delete_resource_with_assignments_reports_cascade(clean_test_project):
    """Deleting a resource with N active assignments returns assignments_removed=N."""
    proj = clean_test_project
    # Add resource + 2 tasks + assign to both
    res_r = _msp_resource_add(name="CascadeRes-T35", type="Work")
    res_id = res_r["resource_id"]
    # Use raw COM to assign (T36 _msp_resource_assign not available yet)
    # Add tasks via raw COM
    proj.Tasks.Add("CascadeT1-T35")
    proj.Tasks.Add("CascadeT2-T35")
    t1 = proj.Tasks(proj.Tasks.Count - 1)
    t2 = proj.Tasks(proj.Tasks.Count)
    # MS Project: assignments are created via task.Assignments.Add(TaskID, ResourceID)
    t1.Assignments.Add(t1.ID, res_id)
    t2.Assignments.Add(t2.ID, res_id)
    res_obj = _find_resource_by_id(proj, res_id)
    assert res_obj.Assignments.Count == 2

    r = _msp_resource_delete(resource_id=res_id)
    assert r["status"] == "ok"
    assert r["assignments_removed"] == 2
    # Verify cascade: resource is gone (its assignments removed silently by MS Project)
    assert _find_resource_by_id(proj, res_id) is None


def test_delete_resource_with_zero_assignments(clean_test_project):
    """assignments_removed = 0 for unassigned resource (sanity)."""
    r1 = _msp_resource_add(name="UnassRes-T35", type="Work")
    r = _msp_resource_delete(resource_id=r1["resource_id"])
    assert r["status"] == "ok"
    assert r["assignments_removed"] == 0
