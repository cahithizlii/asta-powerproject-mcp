"""Test msproject_baseline compare action — variance reporting."""
import pytest
import time
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_compare,
    _msp_task_add_single, _msp_task_update,
)


def test_compare_no_change_zero_variance(clean_test_project):
    """Save baseline, no progress → all tasks on_time, totals=0."""
    add_r = _msp_task_add_single(name="NoVarT-T44", duration="3d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_compare(baseline_number=0)
    assert r["status"] == "ok"
    s = r["summary"]
    assert s["slipped_count"] == 0
    assert s["ahead_count"] == 0
    assert s["on_time_count"] == 1
    assert s["total_finish_drift_days"] == 0


def test_compare_slipped_task(clean_test_project):
    """Slip a task's finish by 5 days → slipped_count=1, drift>0."""
    add_r = _msp_task_add_single(name="SlippedT-T44", duration="3d")
    _msp_baseline_save(baseline_number=0)
    # Slip: extend duration by 5 days
    _msp_task_update(task_id=add_r["task_id"], duration="8d")
    r = _msp_baseline_compare(baseline_number=0)
    assert r["status"] == "ok"
    assert r["summary"]["slipped_count"] == 1
    assert r["summary"]["total_finish_drift_days"] > 0
    # Find this task in list
    task_var = next(t for t in r["tasks"] if t["id"] == add_r["task_id"])
    assert task_var["status"] == "slipped"
    assert task_var["finish_var_days"] > 0


def test_compare_threshold_filter(clean_test_project):
    """variance_threshold_days=10 → small slip not counted as slipped."""
    add_r = _msp_task_add_single(name="ThreshT-T44", duration="5d")
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=add_r["task_id"], duration="7d")  # 2 day slip
    r = _msp_baseline_compare(baseline_number=0, variance_threshold_days=10)
    assert r["status"] == "ok"
    # 2 day slip < 10 day threshold → counts as on_time
    assert r["summary"]["slipped_count"] == 0
    assert r["summary"]["on_time_count"] == 1


def test_compare_include_unchanged_false_filters(clean_test_project):
    """include_unchanged=False → tasks with 0 variance excluded from list."""
    a = _msp_task_add_single(name="UnchT1-T44", duration="2d")
    b = _msp_task_add_single(name="UnchT2-T44", duration="3d")
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=b["task_id"], duration="5d")  # only B slipped
    r = _msp_baseline_compare(baseline_number=0, include_unchanged=False)
    assert r["status"] == "ok"
    # Only B should be in tasks list
    ids = [t["id"] for t in r["tasks"]]
    assert b["task_id"] in ids
    assert a["task_id"] not in ids


def test_compare_unsaved_baseline_errors(clean_test_project):
    r = _msp_baseline_compare(baseline_number=5)
    assert r["status"] == "error"
    assert "not saved" in r["error"].lower() or "no baseline" in r["error"].lower()


def test_compare_perf_50_tasks_under_2s(clean_test_project):
    """Performance: 50-task compare must complete <2s."""
    for i in range(50):
        _msp_task_add_single(name=f"PerfT{i:02d}-T44", duration="1d")
    _msp_baseline_save(baseline_number=0)
    start = time.time()
    r = _msp_baseline_compare(baseline_number=0)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["summary"]["on_time_count"] == 50
    assert elapsed < 2.0, f"compare 50 tasks took {elapsed:.2f}s (target <2s)"
