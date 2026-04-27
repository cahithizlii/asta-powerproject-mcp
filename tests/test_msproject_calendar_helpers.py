"""Test calendar helpers + UZBEK_HOLIDAYS constant."""
import pytest
from msproject_mcp_core import UZBEK_HOLIDAYS_2026, _find_calendar_by_name


def test_uzbek_holidays_count():
    """9 official Uzbek holidays."""
    assert len(UZBEK_HOLIDAYS_2026) == 9


def test_uzbek_holidays_structure():
    """Each entry: (name, month, day) tuple."""
    for entry in UZBEK_HOLIDAYS_2026:
        assert len(entry) == 3
        name, month, day = entry
        assert isinstance(name, str) and len(name) > 0
        assert 1 <= month <= 12
        assert 1 <= day <= 31


def test_uzbek_holidays_includes_navruz():
    """Navruz (March 21) must be in list."""
    found = [e for e in UZBEK_HOLIDAYS_2026 if e[1] == 3 and e[2] == 21]
    assert len(found) == 1
    assert "Navruz" in found[0][0]


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
