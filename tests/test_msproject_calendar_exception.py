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
