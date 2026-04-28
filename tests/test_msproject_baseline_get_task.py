"""Test msproject_baseline get_task_baseline action."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_get_task_baseline,
    _msp_task_add_single,
)


def test_get_task_baseline_after_save(clean_test_project):
    """Save baseline 0, read task's baseline → real values."""
    add_r = _msp_task_add_single(name="GetT-T43", duration="5d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_get_task_baseline(task_id=add_r["task_id"], baseline_number=0)
    assert r["status"] == "ok"
    assert r["task_id"] == add_r["task_id"]
    assert r["baseline_number"] == 0
    bd = r["baseline"]
    assert bd["start"] is not None
    assert bd["finish"] is not None
    assert bd["duration_h"] > 0  # 5d × 8h = 40h


def test_get_task_baseline_before_save(clean_test_project):
    """No baseline saved → MSP returns sentinel ('NA') for dates, 0 for numerics."""
    add_r = _msp_task_add_single(name="UnsavedT-T43", duration="2d")
    r = _msp_baseline_get_task_baseline(task_id=add_r["task_id"], baseline_number=0)
    assert r["status"] == "ok"
    # MSP COM returns the string 'NA' for unsaved baseline dates (not None)
    assert r["baseline"]["start"] in (None, "NA")
    assert r["baseline"]["duration_h"] == 0


def test_get_task_baseline_missing_task(clean_test_project):
    r = _msp_baseline_get_task_baseline(task_id=99999, baseline_number=0)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_get_task_baseline_invalid_baseline_number(clean_test_project):
    add_r = _msp_task_add_single(name="BadBN-T43", duration="1d")
    r = _msp_baseline_get_task_baseline(task_id=add_r["task_id"], baseline_number=99)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
