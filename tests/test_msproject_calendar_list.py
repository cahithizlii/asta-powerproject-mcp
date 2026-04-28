"""Test msproject_calendar list action."""
import pytest
from msproject_mcp_core import _msp_calendar_create, _msp_calendar_add_exception, _msp_calendar_list


def test_list_includes_standard(clean_test_project):
    """Default calendars (Standard, 24 Hours, Night Shift) listed."""
    r = _msp_calendar_list()
    assert r["status"] == "ok"
    names = [c["name"] for c in r["calendars"]]
    assert "Standard" in names
    assert r["count"] == len(r["calendars"]) >= 1
    # Verify new key name (renamed from "uid" in T28)
    assert all("calendar_uid" in c for c in r["calendars"])


def test_list_includes_custom(clean_test_project):
    """Custom calendar appears in list with exceptions."""
    _msp_calendar_create(name="ListCal-Phase2a", base_calendar="Standard")
    _msp_calendar_add_exception(
        calendar_name="ListCal-Phase2a",
        exception_name="Test Holiday",
        start="2026-01-01",
    )
    r = _msp_calendar_list()
    assert r["status"] == "ok"
    custom = next((c for c in r["calendars"] if c["name"] == "ListCal-Phase2a"), None)
    assert custom is not None
    assert custom["exception_count"] >= 1
