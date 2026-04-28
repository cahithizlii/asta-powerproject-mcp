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


def test_add_negative_max_units_errors(clean_test_project):
    """max_units < 0 returns error before mutation."""
    proj = clean_test_project
    initial_count = proj.Resources.Count
    r = _msp_resource_add(name="NegUnits-T33", type="Work", max_units=-100)
    assert r["status"] == "error"
    assert "max_units" in r["error"].lower()
    assert proj.Resources.Count == initial_count  # No mutation


def test_add_negative_rate_errors(clean_test_project):
    """standard_rate < 0 returns error before mutation."""
    proj = clean_test_project
    initial = proj.Resources.Count
    r = _msp_resource_add(name="NegRate-T33", type="Work", standard_rate=-50)
    assert r["status"] == "error"
    assert "standard_rate" in r["error"].lower()
    assert proj.Resources.Count == initial


def test_add_rollback_on_post_add_failure(clean_test_project, monkeypatch):
    """If property-set fails after Resources.Add, the orphan resource is deleted.

    Strategy: monkeypatch RESOURCE_TYPES so 'Work' maps to an invalid Type code (999).
    Pre-flight passes (key 'Work' still in dict), then `res.Type = 999` fails at COM
    layer. The except branch must call res.Delete() to roll back the orphan.

    Fallback: if MSP silently accepts 999, the test verifies count unchanged OR
    that the test resource is not present (which would mean rollback worked OR
    the COM layer did reject the assignment).
    """
    import msproject_mcp_core as core
    proj = clean_test_project
    initial = proj.Resources.Count

    # Override Type code to invalid value to force COM rejection on assignment
    monkeypatch.setattr(core, "RESOURCE_TYPES",
                        {"Work": 999, "Material": 1, "Cost": 2})
    r = core._msp_resource_add(name="RollbackTest-T33", type="Work")

    # Either the call errored AND was rolled back (preferred) OR the call
    # somehow succeeded (MSP accepted 999). The rollback contract requires that
    # on error the count returns to baseline.
    if r["status"] == "error":
        assert proj.Resources.Count == initial, (
            f"Orphan resource not rolled back: count {proj.Resources.Count} != "
            f"initial {initial}; Resource(s) still present indicates the rollback "
            f"path did not execute Delete()."
        )
        # Belt-and-braces: also verify by name
        assert _find_resource_by_name(proj, "RollbackTest-T33") is None, \
            "Orphan resource 'RollbackTest-T33' still findable by name after rollback"
    else:
        # MSP silently accepted Type=999 — test is inconclusive, but document it.
        pytest.skip("MSP COM accepted invalid Type=999; rollback path not exercised")
