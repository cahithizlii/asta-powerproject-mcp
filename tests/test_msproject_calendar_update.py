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


def test_update_rename_conflict_errors(clean_test_project):
    """Renaming to an already-taken name returns error and does NOT mutate."""
    proj = clean_test_project
    _msp_calendar_create(name="OriginalA-Phase2a", base_calendar="Standard")
    _msp_calendar_create(name="ExistingB-Phase2a", base_calendar="Standard")
    r = _msp_calendar_update(name="OriginalA-Phase2a", new_name="ExistingB-Phase2a")
    assert r["status"] == "error"
    assert "already exists" in r["error"].lower()
    # Verify NO mutation occurred
    assert _find_calendar_by_name(proj, "OriginalA-Phase2a") is not None


def test_update_invalid_weekday_off(clean_test_project):
    """weekday_off out of range (0, 8) returns error."""
    _msp_calendar_create(name="RangeCal-Phase2a", base_calendar="Standard")
    for bad in (0, 8, -1):
        r = _msp_calendar_update(name="RangeCal-Phase2a", weekday_off=bad)
        assert r["status"] == "error", f"weekday_off={bad} should error"
        assert "1-7" in r["error"]


def test_update_combined_rename_and_weekday(clean_test_project):
    """Both rename and weekday_off in one call → both applied, both in changes."""
    proj = clean_test_project
    _msp_calendar_create(name="ComboOld-Phase2a", base_calendar="Standard")
    r = _msp_calendar_update(name="ComboOld-Phase2a",
                             new_name="ComboNew-Phase2a", weekday_off=1)
    assert r["status"] == "ok"
    assert "name" in r["changes"]
    assert "weekday_off" in r["changes"]
    cal = _find_calendar_by_name(proj, "ComboNew-Phase2a")
    assert cal is not None
    assert cal.WeekDays(1).Working is False


def test_update_invalid_weekday_does_not_rename(clean_test_project):
    """If weekday_off invalid, rename must NOT have happened (partial-write protection)."""
    proj = clean_test_project
    _msp_calendar_create(name="ProtectOld-Phase2a", base_calendar="Standard")
    r = _msp_calendar_update(name="ProtectOld-Phase2a",
                             new_name="ProtectNew-Phase2a", weekday_off=99)
    assert r["status"] == "error"
    # Original name must still exist (no partial mutation)
    assert _find_calendar_by_name(proj, "ProtectOld-Phase2a") is not None
    assert _find_calendar_by_name(proj, "ProtectNew-Phase2a") is None
