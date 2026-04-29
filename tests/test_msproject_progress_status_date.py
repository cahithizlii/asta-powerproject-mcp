"""Test msproject_progress set_status_date action."""
import pytest
from msproject_mcp_core import _msp_progress_set_status_date


def test_set_status_date_basic(clean_test_project):
    proj = clean_test_project
    r = _msp_progress_set_status_date(status_date="2026-04-29")
    assert r["status"] == "ok"
    assert "status_date" in r
    # Verify on project
    sd = str(proj.StatusDate)
    assert "2026" in sd or "04" in sd


def test_set_status_date_invalid_format(clean_test_project):
    r = _msp_progress_set_status_date(status_date="not a date")
    assert r["status"] == "error"
