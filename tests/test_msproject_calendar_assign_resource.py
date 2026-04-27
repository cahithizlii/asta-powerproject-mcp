"""Test msproject_calendar assign_to_resource action.

Phase 2a uses raw COM Resources.Add since the high-level Resource tool
arrives in Phase 2b. Once Phase 2b lands, we can refactor to use it.
"""
import pytest
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_assign_to_resource,
)


def _add_resource(proj, name: str) -> int:
    """Helper: add a Work resource via raw COM. Returns resource ID."""
    r = proj.Resources.Add(name)
    return r.ID


def test_assign_calendar_to_resource(clean_test_project):
    proj = clean_test_project
    _msp_calendar_create(name="ResCal-Phase2a", base_calendar="Standard")
    res_id = _add_resource(proj, "TestRes-Phase2a")
    r = _msp_calendar_assign_to_resource(resource_id=res_id, calendar_name="ResCal-Phase2a")
    assert r["status"] == "ok"
    # Verify
    res = None
    for i in range(1, proj.Resources.Count + 1):
        rr = proj.Resources(i)
        if rr is not None and rr.ID == res_id:
            res = rr
            break
    assert res is not None
    assert res.BaseCalendar == "ResCal-Phase2a"


def test_assign_missing_calendar_errors(clean_test_project):
    proj = clean_test_project
    res_id = _add_resource(proj, "OrphanRes-Phase2a")
    r = _msp_calendar_assign_to_resource(resource_id=res_id, calendar_name="NoSuch-Phase2a")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_assign_missing_resource_errors(clean_test_project):
    _msp_calendar_create(name="LonelyCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_assign_to_resource(resource_id=99999, calendar_name="LonelyCal-Phase2a")
    assert r["status"] == "error"
    assert "resource" in r["error"].lower() and "99999" in r["error"]
