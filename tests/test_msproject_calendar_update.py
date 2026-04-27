"""Test msproject_calendar update action."""
import pytest
from msproject_mcp_core import _msp_calendar_create, _msp_calendar_update, _find_calendar_by_name


def test_update_rename(clean_test_project):
    """Rename a calendar."""
    proj = clean_test_project
    _msp_calendar_create(name="OldName-Phase2a", base_calendar="Standard")
    r = _msp_calendar_update(name="OldName-Phase2a", new_name="NewName-Phase2a")
    assert r["status"] == "ok"
    assert "name" in r["changes"]
    assert _find_calendar_by_name(proj, "NewName-Phase2a") is not None
    assert _find_calendar_by_name(proj, "OldName-Phase2a") is None


def test_update_weekday_off(clean_test_project):
    """Set Sunday (weekday=1 in MSP) as non-working."""
    proj = clean_test_project
    _msp_calendar_create(name="WeekdayCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_update(name="WeekdayCal-Phase2a", weekday_off=1)
    assert r["status"] == "ok"
    assert "weekday_off" in r["changes"]
    cal = _find_calendar_by_name(proj, "WeekdayCal-Phase2a")
    sunday = cal.WeekDays(1)
    assert sunday.Working is False


def test_update_missing_calendar_errors(clean_test_project):
    r = _msp_calendar_update(name="DoesNotExist-Phase2a", new_name="X")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()
