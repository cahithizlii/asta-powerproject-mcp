"""Test msproject_baseline save action."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_task_add_single, _baseline_saved_date,
)


def test_save_default_baseline_zero(clean_test_project):
    """Save baseline 0 on a project with 3 tasks → returns metadata."""
    proj = clean_test_project
    for i in range(3):
        _msp_task_add_single(name=f"SaveT{i}-T40", duration="2d")
    r = _msp_baseline_save(baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0
    assert r["task_count"] == 3
    assert r["saved_date"] is not None
    # Confirm via _baseline_saved_date helper
    assert _baseline_saved_date(proj, 0) is not None


def test_save_baseline_three(clean_test_project):
    """Save into Baseline3 (Into=13)."""
    proj = clean_test_project
    _msp_task_add_single(name="B3T-T40", duration="5d")
    r = _msp_baseline_save(baseline_number=3)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 3
    assert _baseline_saved_date(proj, 3) is not None
    # Verify Baseline 0 still unsaved
    assert _baseline_saved_date(proj, 0) is None


def test_save_invalid_baseline_number_errors(clean_test_project):
    """baseline_number 11 (out of 0-10) → error."""
    r = _msp_baseline_save(baseline_number=11)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
    assert "0-10" in r["error"]


def test_save_negative_baseline_number_errors(clean_test_project):
    r = _msp_baseline_save(baseline_number=-1)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
