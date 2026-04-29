"""Test msproject_baseline compare_two action — baseline-to-baseline delta."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_compare_two,
    _msp_task_add_single, _msp_task_update,
)


def test_compare_two_zero_when_baselines_identical(clean_test_project):
    """Save B0, then save B1 (snapshot of same state) → variance 0."""
    add_r = _msp_task_add_single(name="EqT-T45", duration="3d")
    _msp_baseline_save(baseline_number=0)
    _msp_baseline_save(baseline_number=1)
    r = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
    assert r["status"] == "ok"
    assert r["summary"]["slipped_count"] == 0


def test_compare_two_revision_delta(clean_test_project):
    """B0 saved, task changed, B1 saved → compare_two(0,1) shows delta."""
    add_r = _msp_task_add_single(name="DeltaT-T45", duration="3d")
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=add_r["task_id"], duration="8d")  # 5 day slip
    _msp_baseline_save(baseline_number=1)
    r = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
    assert r["status"] == "ok"
    # B0→B1 shows the slip
    assert r["summary"]["slipped_count"] == 1


def test_compare_two_unsaved_baseline_a_errors(clean_test_project):
    _msp_task_add_single(name="OnlyB-T45", duration="2d")
    _msp_baseline_save(baseline_number=1)
    r = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
    assert r["status"] == "error"
    assert "baseline_a" in r["error"].lower() or "0" in r["error"]


def test_compare_two_invalid_baseline_number(clean_test_project):
    r = _msp_baseline_compare_two(baseline_a=99, baseline_b=0)
    assert r["status"] == "error"
