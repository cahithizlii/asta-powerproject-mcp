"""Test msproject_calendar holidays_uzbek action."""
import pytest
import time
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_holidays_uzbek,
    _find_calendar_by_name, UZBEK_HOLIDAYS_2026,
)


def test_uzbek_holidays_added(clean_test_project):
    """All 9 Uzbek holidays added to a fresh calendar in <2s."""
    proj = clean_test_project
    _msp_calendar_create(name="UzbekCal-Phase2a", base_calendar="Standard")
    start = time.time()
    r = _msp_calendar_holidays_uzbek(calendar_name="UzbekCal-Phase2a", year=2026)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 9
    assert elapsed < 2.0, f"holidays_uzbek took {elapsed:.2f}s (target <2s)"
    cal = _find_calendar_by_name(proj, "UzbekCal-Phase2a")
    assert cal.Exceptions.Count >= 9


def test_uzbek_holidays_dates_correct(clean_test_project):
    """Returned holiday dates match UZBEK_HOLIDAYS_2026 constant."""
    _msp_calendar_create(name="UzbekDateCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_holidays_uzbek(calendar_name="UzbekDateCal-Phase2a", year=2026)
    assert r["status"] == "ok"
    returned_dates = {(h["month"], h["day"]) for h in r["holidays"]}
    expected_dates = {(m, d) for _, m, d in UZBEK_HOLIDAYS_2026}
    assert returned_dates == expected_dates


def test_uzbek_holidays_missing_calendar_errors(clean_test_project):
    r = _msp_calendar_holidays_uzbek(calendar_name="NoSuch-Phase2a", year=2026)
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()
