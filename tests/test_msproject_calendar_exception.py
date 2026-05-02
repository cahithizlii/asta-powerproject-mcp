"""Test msproject_calendar add_exception action."""
import pytest
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_add_exception, _find_calendar_by_name,
)


def test_add_single_date_exception(clean_test_project):
    """Add a single-day non-working exception."""
    proj = clean_test_project
    _msp_calendar_create(name="ExCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="ExCal-Phase2a",
        exception_name="New Year",
        start="2026-01-01",
    )
    assert r["status"] == "ok"
    cal = _find_calendar_by_name(proj, "ExCal-Phase2a")
    # Verify the exception was added (Exceptions.Count >= 1)
    assert cal.Exceptions.Count >= 1


def test_add_date_range_exception(clean_test_project):
    """Add an exception spanning multiple days."""
    proj = clean_test_project
    _msp_calendar_create(name="RangeCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="RangeCal-Phase2a",
        exception_name="Spring Break",
        start="2026-03-23",
        finish="2026-03-27",
    )
    assert r["status"] == "ok"


def test_add_exception_missing_calendar_errors(clean_test_project):
    r = _msp_calendar_add_exception(
        calendar_name="NoSuchCal-Phase2a",
        exception_name="Holiday",
        start="2026-01-01",
    )
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_add_exception_invalid_date_range(clean_test_project):
    """Start > finish should error."""
    _msp_calendar_create(name="BadRangeCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="BadRangeCal-Phase2a",
        exception_name="Bad Range",
        start="2026-05-10",
        finish="2026-05-01",
    )
    assert r["status"] == "error"
    assert "start" in r["error"].lower()


def test_add_exception_invalid_date_format(clean_test_project):
    """Malformed start string returns error before mutation."""
    _msp_calendar_create(name="BadFmtCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="BadFmtCal-Phase2a",
        exception_name="Bad Format",
        start="01/01/2026",
    )
    assert r["status"] == "error"
    assert "invalid date format" in r["error"].lower()


def test_add_exception_actually_non_working(clean_test_project):
    """Verify the added exception is actually marked non-working in MSP
    (not just present). T25 holidays_uzbek depends on this contract.

    Note: MSP COM exposes shift times via ex.Shift1.Start (Shift sub-object),
    NOT ex.Shift1Start. We use both that and cal.Period(date).Working as
    independent verifications.
    """
    import pywintypes
    import datetime as _dt
    proj = clean_test_project
    _msp_calendar_create(name="VerifyCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="VerifyCal-Phase2a",
        exception_name="Verified Holiday",
        start="2026-07-04",
    )
    assert r["status"] == "ok"
    cal = _find_calendar_by_name(proj, "VerifyCal-Phase2a")
    # Find our exception (last one added)
    ex = cal.Exceptions(cal.Exceptions.Count)
    assert ex.Name == "Verified Holiday"
    # Non-working day check #1: Shift1.Start == 0 (no working time)
    s1_start = float(ex.Shift1.Start)
    assert s1_start == 0.0, f"Shift1.Start expected 0 (non-working), got {s1_start}"
    # Non-working day check #2: cal.Period(date).Working == False
    period = cal.Period(pywintypes.Time(_dt.date(2026, 7, 4)))
    assert period.Working is False, (
        f"cal.Period(2026-07-04).Working expected False, got {period.Working}"
    )


def test_add_exception_working_true_supported_phase10_3(clean_test_project):
    """Phase 10.3 — working=True now supported (was rejected in Phase 2a)."""
    _msp_calendar_create(name="WorkCal-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="WorkCal-Phase10",
        exception_name="Saturday Working Day",
        start="2026-08-01",
        working=True,
    )
    assert r["status"] == "ok"
    assert r["working"] is True


def test_add_exception_working_true_custom_hours(clean_test_project):
    """Phase 10.3 — working_hours_start/finish override default 08-17."""
    _msp_calendar_create(name="CustomHours-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="CustomHours-Phase10",
        exception_name="Half Day",
        start="2026-08-15",
        working=True,
        working_hours_start="09:00",
        working_hours_finish="13:00",
    )
    assert r["status"] == "ok"
    assert r["working"] is True


def test_add_exception_working_invalid_hours_format(clean_test_project):
    _msp_calendar_create(name="BadHours-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="BadHours-Phase10",
        exception_name="Bad",
        start="2026-08-20",
        working=True,
        working_hours_start="not-a-time",
    )
    assert r["status"] == "error"
    assert "format" in r["error"].lower() or "expected HH:MM" in r["error"]


# === Phase 10.2 — recurring exceptions ===

def test_add_exception_recurring_weekly_requires_days(clean_test_project):
    """recurrence='weekly' without days_of_week -> error."""
    _msp_calendar_create(name="WeeklyNoDays-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="WeeklyNoDays-Phase10",
        exception_name="Weekly Off",
        start="2026-09-01",
        finish="2026-12-31",
        recurrence="weekly",
    )
    assert r["status"] == "error"
    assert "days_of_week" in r["error"]


def test_add_exception_recurring_weekly_with_days(clean_test_project):
    """recurrence='weekly' + days_of_week=['fri','sat'] -> Friday+Saturday off."""
    _msp_calendar_create(name="WeeklyOff-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="WeeklyOff-Phase10",
        exception_name="Weekend Off",
        start="2026-01-01",
        finish="2026-12-31",
        recurrence="weekly",
        days_of_week=["fri", "sat"],
    )
    assert r["status"] == "ok"
    assert r["recurrence"] == "weekly"


def test_add_exception_recurring_monthly(clean_test_project):
    _msp_calendar_create(name="MonthlyOff-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="MonthlyOff-Phase10",
        exception_name="Monthly Maintenance",
        start="2026-01-15",
        finish="2026-12-31",
        recurrence="monthly",
        occurrences=12,
    )
    assert r["status"] == "ok"
    assert r["recurrence"] == "monthly"


def test_add_exception_recurring_yearly(clean_test_project):
    _msp_calendar_create(name="YearlyOff-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="YearlyOff-Phase10",
        exception_name="New Year Annual",
        start="2026-01-01",
        finish="2030-01-01",
        recurrence="yearly",
    )
    assert r["status"] == "ok"
    assert r["recurrence"] == "yearly"


def test_add_exception_recurring_invalid_value(clean_test_project):
    _msp_calendar_create(name="BadRec-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="BadRec-Phase10",
        exception_name="Bad",
        start="2026-01-01",
        recurrence="quarterly",  # not supported
    )
    assert r["status"] == "error"
    assert "recurrence" in r["error"]


def test_add_exception_recurring_unknown_day_name(clean_test_project):
    _msp_calendar_create(name="UnknownDay-Phase10", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="UnknownDay-Phase10",
        exception_name="Bad Day",
        start="2026-01-01",
        finish="2026-12-31",
        recurrence="weekly",
        days_of_week=["xyz"],
    )
    assert r["status"] == "error"
    assert "day_of_week" in r["error"].lower() or "Valid" in r["error"]
