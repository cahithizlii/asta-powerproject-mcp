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


def test_save_scope_selected(clean_test_project):
    """scope='selected' saves only currently-selected tasks. Use SelectAll for coverage."""
    proj = clean_test_project
    import pythoncom, win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject('MSProject.Application')
    for i in range(3):
        _msp_task_add_single(name=f"SelT{i}-T40", duration="1d")
    # Select all tasks via app COM (no UI)
    try:
        app.SelectAll()
    except Exception:
        pytest.skip("SelectAll not exposed — scope='selected' coverage deferred")
    r = _msp_baseline_save(baseline_number=2, scope="selected")
    assert r["status"] == "ok"
    assert r["baseline_number"] == 2
    # Saved date should be present (assuming SelectAll worked)
    if "warning" not in r:
        assert _baseline_saved_date(proj, 2) is not None


def test_save_invalid_scope_errors(clean_test_project):
    """scope='bogus' returns error pre-mutation."""
    r = _msp_baseline_save(baseline_number=0, scope="bogus")
    assert r["status"] == "error"
    assert "scope" in r["error"].lower()
