"""Test calendar helpers + UZBEK_HOLIDAYS constant."""
import pytest
from msproject_mcp_core import UZBEK_HOLIDAYS_2026, _find_calendar_by_name


def test_uzbek_holidays_count():
    """9 official Uzbek holidays."""
    assert len(UZBEK_HOLIDAYS_2026) == 9


def test_uzbek_holidays_structure():
    """Each entry: (str name, real (month, day) for 2026)."""
    from datetime import date
    for entry in UZBEK_HOLIDAYS_2026:
        assert len(entry) == 3
        name, month, day = entry
        assert isinstance(name, str) and len(name) > 0
        # Constructing the date catches invalid (month, day) like (2, 30)
        date(2026, month, day)


def test_uzbek_holidays_includes_navruz():
    """Navruz (March 21) must be in list."""
    found = [e for e in UZBEK_HOLIDAYS_2026 if e[1] == 3 and e[2] == 21]
    assert len(found) == 1
    assert found[0][0] == "Navruz"


def test_find_calendar_standard(clean_test_project):
    """_find_calendar_by_name finds 'Standard' in any project."""
    proj = clean_test_project
    cal = _find_calendar_by_name(proj, "Standard")
    assert cal is not None
    assert cal.Name == "Standard"


def test_find_calendar_missing(clean_test_project):
    """Returns None for missing name."""
    proj = clean_test_project
    cal = _find_calendar_by_name(proj, "NonExistent-XYZ")
    assert cal is None


def test_find_calendar_case_sensitive(clean_test_project):
    """_find_calendar_by_name is case-sensitive — locks the contract for T19+."""
    proj = clean_test_project
    assert _find_calendar_by_name(proj, "standard") is None
    assert _find_calendar_by_name(proj, "Standard") is not None
