"""Test msproject_resource update action."""
import pytest
from msproject_mcp_core import (
    _msp_resource_add, _msp_resource_update, _find_resource_by_id, _parse_rate,
)


def test_update_rename(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="OldName-T34", type="Work")
    r = _msp_resource_update(resource_id=r1["resource_id"], name="NewName-T34")
    assert r["status"] == "ok"
    assert "name" in r["changes"]
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert res.Name == "NewName-T34"


def test_update_rate_and_units(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="RateRes-T34", type="Work", max_units=100, standard_rate=50.0)
    r = _msp_resource_update(resource_id=r1["resource_id"],
                            max_units=600, standard_rate=80.0, overtime_rate=120.0)
    assert r["status"] == "ok"
    assert "max_units" in r["changes"]
    assert "standard_rate" in r["changes"]
    assert "overtime_rate" in r["changes"]
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert abs(float(res.MaxUnits) - 6.0) < 0.01
    assert abs(_parse_rate(res.StandardRate) - 80.0) < 0.01


def test_update_material_label(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="Mat-T34", type="Material", material_label="kg")
    r = _msp_resource_update(resource_id=r1["resource_id"], material_label="ton")
    assert r["status"] == "ok"
    assert "material_label" in r["changes"]
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert res.MaterialLabel == "ton"


def test_update_missing_resource_errors(clean_test_project):
    r = _msp_resource_update(resource_id=99999, name="X")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_update_rename_conflict_errors(clean_test_project):
    proj = clean_test_project
    _msp_resource_add(name="ExistingName-T34", type="Work")
    r1 = _msp_resource_add(name="ToRename-T34", type="Work")
    r = _msp_resource_update(resource_id=r1["resource_id"], name="ExistingName-T34")
    assert r["status"] == "error"
    assert "already exists" in r["error"].lower()
    # Verify NO mutation
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert res.Name == "ToRename-T34"


def test_update_negative_value_errors(clean_test_project):
    """Negative rate/units rejected pre-mutation."""
    proj = clean_test_project
    r1 = _msp_resource_add(name="NegTest-T34", type="Work", max_units=100, standard_rate=50)
    r = _msp_resource_update(resource_id=r1["resource_id"], standard_rate=-10)
    assert r["status"] == "error"
    assert "standard_rate" in r["error"].lower()
    # Verify NO mutation occurred
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert abs(_parse_rate(res.StandardRate) - 50.0) < 0.01


def test_update_combined_partial_write_protection(clean_test_project):
    """Combined rename+invalid-rate: rename must NOT happen if validation fails."""
    proj = clean_test_project
    r1 = _msp_resource_add(name="ProtectOld-T34", type="Work")
    r = _msp_resource_update(resource_id=r1["resource_id"],
                            name="ProtectNew-T34", standard_rate=-99)
    assert r["status"] == "error"
    res = _find_resource_by_id(proj, r1["resource_id"])
    # Original name preserved (no partial mutation)
    assert res.Name == "ProtectOld-T34"
