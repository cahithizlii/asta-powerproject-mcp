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


def test_compare_task_added_after_baseline_save_no_silent_on_time(clean_test_project):
    """Task added AFTER baseline save → its baseline start/finish should be None,
    not classified silently as on_time. The compare result should still succeed
    (no exception), but the task should NOT be misclassified as 'on_time' with
    fake zero variance.

    Regression test for T44 review issue I1: 'NA' sentinel leak in
    _read_task_baseline. Before fix, MSP returned literal 'NA' for the new
    task's BaselineNStart/Finish; 'NA' is truthy → str('NA') → leaks downstream
    where _datetime_diff_days silently returned 0.0 (parsing failure swallow).

    With _msp_dt_or_none in place, baseline start/finish are properly None for
    post-baseline tasks. duration_var_h is the load-bearing signal: cur_dur_h
    is non-zero (task has a real duration) but bd['duration_h'] is 0 → variance
    surfaces, no_change is False, and the task appears in the tasks list.
    """
    pre_r = _msp_task_add_single(name="PreT-T44I1", duration="2d")
    _msp_baseline_save(baseline_number=0)
    # Add a task AFTER baseline save
    post_r = _msp_task_add_single(name="PostT-T44I1", duration="3d")
    r = _msp_baseline_compare(baseline_number=0, include_unchanged=True)
    assert r["status"] == "ok"
    # The pre-baseline task is on_time (no change)
    pre_entry = next(t for t in r["tasks"] if t["id"] == pre_r["task_id"])
    assert pre_entry["status"] == "on_time"
    # The post-baseline task: duration_var_h must be > 0 because cur_dur_h=24
    # (3d × 8h) and bd["duration_h"]=0 (no baseline saved for this task).
    post_entry = next(t for t in r["tasks"] if t["id"] == post_r["task_id"])
    assert post_entry is not None
    assert post_entry["duration_var_h"] > 0  # Real signal that it has no baseline


def test_compare_perf_100_tasks_under_3s(clean_test_project):
    """Stretch perf: 100 tasks <3s with shared helper + pre-build cache (TAIL #1).

    Phase 3a T49 stretches the perf bound from 50 (T44) to 100 tasks.
    The pre-build `real_tasks` cache in _compute_variance_set cuts ~50% of
    the COM dispatches on `t.Summary` checks vs the per-call proj.Tasks(i)
    pattern.
    """
    for i in range(100):
        _msp_task_add_single(name=f"Perf100T{i:03d}-T49", duration="1d")
    _msp_baseline_save(baseline_number=0)
    start = time.time()
    r = _msp_baseline_compare(baseline_number=0)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["summary"]["on_time_count"] == 100
    assert elapsed < 3.0, f"compare 100 tasks took {elapsed:.2f}s (target <3s)"
