"""Test msproject_resource add action — Work, Material, Cost types."""
import pytest
from msproject_mcp_core import _msp_resource_add, _find_resource_by_name, _parse_rate


def test_add_work_resource_default(clean_test_project):
    """Default add — Work type, MaxUnits=100, no rate."""
    proj = clean_test_project
    r = _msp_resource_add(name="WorkA-T33", type="Work")
    assert r["status"] == "ok"
    assert r["name"] == "WorkA-T33"
    assert r["type"] == "Work"
    res = _find_resource_by_name(proj, "WorkA-T33")
    assert res is not None
    assert int(res.Type) == 0


def test_add_work_with_rate(clean_test_project):
    """500% units = 5-person crew, $75/h standard, $112.5/h OT."""
    r = _msp_resource_add(name="WorkB-T33", type="Work", max_units=500, standard_rate=75.0, overtime_rate=112.5)
    assert r["status"] == "ok"
    proj = clean_test_project
    res = _find_resource_by_name(proj, "WorkB-T33")
    assert abs(float(res.MaxUnits) - 5.0) < 0.01  # 500% -> 5.0 fraction
    # Use _parse_rate since MSP returns localized string
    assert abs(_parse_rate(res.StandardRate) - 75.0) < 0.01
    assert abs(_parse_rate(res.OvertimeRate) - 112.5) < 0.01


def test_add_material(clean_test_project):
    r = _msp_resource_add(name="Cement-T33", type="Material", material_label="ton", standard_rate=120.0)
    assert r["status"] == "ok"
    assert r["type"] == "Material"
    proj = clean_test_project
    res = _find_resource_by_name(proj, "Cement-T33")
    assert int(res.Type) == 1
    assert res.MaterialLabel == "ton"
    assert abs(_parse_rate(res.StandardRate) - 120.0) < 0.01


def test_add_cost(clean_test_project):
    r = _msp_resource_add(name="Travel-T33", type="Cost")
    assert r["status"] == "ok"
    assert r["type"] == "Cost"
    proj = clean_test_project
    res = _find_resource_by_name(proj, "Travel-T33")
    assert int(res.Type) == 2


def test_add_invalid_type_errors(clean_test_project):
    r = _msp_resource_add(name="X-T33", type="Bogus")
    assert r["status"] == "error"
    assert "type" in r["error"].lower()


def test_add_duplicate_name_errors(clean_test_project):
    _msp_resource_add(name="DupRes-T33", type="Work")
    r = _msp_resource_add(name="DupRes-T33", type="Work")
    assert r["status"] == "error"
    assert "already exists" in r["error"].lower()
