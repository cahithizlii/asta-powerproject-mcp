"""Test msproject_baseline clear + clear_all actions."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_clear, _msp_baseline_clear_all,
    _msp_task_add_single, _baseline_saved_date,
)


def test_clear_saved_baseline(clean_test_project):
    """Save baseline 0 then clear it."""
    proj = clean_test_project
    _msp_task_add_single(name="ClearT-T41", duration="1d")
    _msp_baseline_save(baseline_number=0)
    assert _baseline_saved_date(proj, 0) is not None
    r = _msp_baseline_clear(baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0
    assert r["was_saved_date"] is not None
    # Verify cleared
    assert _baseline_saved_date(proj, 0) is None


def test_clear_unsaved_baseline_no_op(clean_test_project):
    """Clearing an unsaved baseline → ok with was_saved_date=None."""
    r = _msp_baseline_clear(baseline_number=5)
    assert r["status"] == "ok"
    assert r["was_saved_date"] is None


def test_clear_invalid_baseline_number_errors(clean_test_project):
    r = _msp_baseline_clear(baseline_number=11)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()


def test_clear_all_three_baselines(clean_test_project):
    """Save 3 baselines, then clear_all → all empty."""
    proj = clean_test_project
    _msp_task_add_single(name="ClearAllT-T41", duration="1d")
    _msp_baseline_save(baseline_number=0)
    _msp_baseline_save(baseline_number=2)
    _msp_baseline_save(baseline_number=7)
    r = _msp_baseline_clear_all()
    assert r["status"] == "ok"
    assert sorted(r["cleared"]) == [0, 2, 7]
    assert r["count"] == 3
    # Verify all 11 are now unsaved
    for n in range(11):
        assert _baseline_saved_date(proj, n) is None


def test_clear_all_when_none_saved(clean_test_project):
    """clear_all on fresh project → ok with empty cleared list."""
    r = _msp_baseline_clear_all()
    assert r["status"] == "ok"
    assert r["cleared"] == []
    assert r["count"] == 0
