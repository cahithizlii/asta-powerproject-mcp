"""Phase 6.2 T116 — time_phased_ac pure module tests.

Verifies per-task linear AC distribution replaces the prior uniform
'total_ac / past_buckets' simplification. Tests cover:
- single completed task
- single in-progress task (no actual_finish)
- partial linear interpolation between actual_start and actual_finish
- task with no actual_start (baseline_start fallback)
- multi-task cumulative across buckets
- future buckets capped at data_date
- zero actual_work
"""
import datetime as dt

from evm_math import time_phased_ac, _task_ac_at_date


# === _task_ac_at_date unit cases ===

def test_task_ac_zero_actual_work():
    t = {"actual_start": dt.date(2026, 1, 1), "actual_work": 0}
    assert _task_ac_at_date(t, dt.date(2026, 2, 1)) == 0.0


def test_task_ac_completed_returns_full_actual_work():
    t = {
        "actual_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 1, 31),
        "actual_work": 80.0,
    }
    assert _task_ac_at_date(t, dt.date(2026, 2, 28)) == 80.0


def test_task_ac_not_yet_started_returns_zero():
    t = {
        "actual_start": dt.date(2026, 3, 1),
        "actual_finish": dt.date(2026, 3, 31),
        "actual_work": 80.0,
    }
    assert _task_ac_at_date(t, dt.date(2026, 1, 15)) == 0.0


def test_task_ac_in_progress_no_finish_returns_full_actual_work():
    """In-progress task: actual_work reported is what's already spent.
    Eval at data_date >= actual_start -> full aw (not interpolated)."""
    t = {
        "actual_start": dt.date(2026, 1, 1),
        "actual_work": 50.0,
    }
    assert _task_ac_at_date(t, dt.date(2026, 2, 15)) == 50.0


def test_task_ac_linear_midpoint():
    """50% time elapsed between actual_start..actual_finish -> 50% AC."""
    t = {
        "actual_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 1, 21),  # 20 days
        "actual_work": 80.0,
    }
    # Day 10 from start -> 50% elapsed
    result = _task_ac_at_date(t, dt.date(2026, 1, 11))
    assert result == 40.0


def test_task_ac_baseline_fallback_when_actual_start_missing():
    """No actual_start -> use baseline_start as best estimate."""
    t = {
        "baseline_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 1, 31),
        "actual_work": 100.0,
    }
    assert _task_ac_at_date(t, dt.date(2026, 2, 28)) == 100.0


# === time_phased_ac aggregate cases ===

def test_time_phased_ac_single_task_cumulative():
    """One task spanning multiple buckets — cumulative AC grows linearly."""
    tasks = [{
        "actual_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 1, 31),  # 30 days
        "actual_work": 60.0,  # 2.0 per day
    }]
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 1, 11)),  # bucket 1: end day 10
        (dt.date(2026, 1, 11), dt.date(2026, 1, 21)),  # bucket 2: end day 20
        (dt.date(2026, 1, 21), dt.date(2026, 1, 31)),  # bucket 3: end day 30
    ]
    data_date = dt.date(2026, 1, 31)
    ac = time_phased_ac(tasks, buckets, data_date)
    # Bucket 1 end: day 10 of 30 -> 60 * 10/30 = 20
    # Bucket 2 end: day 20 of 30 -> 60 * 20/30 = 40
    # Bucket 3 end: day 30 of 30 (== finish) -> 60 (full)
    assert ac[0] == 20.0
    assert ac[1] == 40.0
    assert ac[2] == 60.0


def test_time_phased_ac_capped_at_data_date():
    """Future buckets after data_date should not grow AC."""
    tasks = [{
        "actual_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 3, 31),
        "actual_work": 90.0,
    }]
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 2, 1)),
        (dt.date(2026, 2, 1), dt.date(2026, 3, 1)),
        (dt.date(2026, 3, 1), dt.date(2026, 4, 1)),
    ]
    data_date = dt.date(2026, 2, 1)  # cap at end of bucket 1
    ac = time_phased_ac(tasks, buckets, data_date)
    # All buckets cap eval at data_date -> AC frozen at bucket 1 value
    assert ac[0] == ac[1] == ac[2]


def test_time_phased_ac_multi_task():
    """Two tasks running in parallel — both contribute to cumulative."""
    tasks = [
        {
            "actual_start": dt.date(2026, 1, 1),
            "actual_finish": dt.date(2026, 1, 31),
            "actual_work": 30.0,
        },
        {
            "actual_start": dt.date(2026, 1, 1),
            "actual_finish": dt.date(2026, 1, 31),
            "actual_work": 60.0,
        },
    ]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 31))]
    ac = time_phased_ac(tasks, buckets, dt.date(2026, 1, 31))
    assert ac[0] == 90.0


def test_time_phased_ac_in_progress_task_at_data_date():
    """In-progress task (no actual_finish) reports full actual_work at
    data_date — represents work already spent through reporting period."""
    tasks = [{
        "actual_start": dt.date(2026, 1, 1),
        "actual_work": 25.0,
    }]
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 1, 15)),
        (dt.date(2026, 1, 15), dt.date(2026, 1, 31)),
    ]
    ac = time_phased_ac(tasks, buckets, dt.date(2026, 1, 31))
    assert ac[0] == 25.0
    assert ac[1] == 25.0


def test_time_phased_ac_unstarted_task_zero():
    tasks = [{
        "actual_start": dt.date(2026, 6, 1),
        "actual_finish": dt.date(2026, 6, 30),
        "actual_work": 40.0,
    }]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 2, 1))]
    ac = time_phased_ac(tasks, buckets, dt.date(2026, 2, 1))
    assert ac[0] == 0.0
