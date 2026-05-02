"""Phase 9.2 — time_phased_ac_increments per-bucket delta tests.

Verifies the non-cumulative AC view:
- sum of increments == final cumulative AC
- increments[0] == cumulative[0]
- empty buckets -> []
- single bucket -> single-element list matches cumulative[0]
- cap-at-data-date behavior preserved (later buckets => 0 increment)
"""
import datetime as dt

from evm_math import time_phased_ac, time_phased_ac_increments


def test_increments_sum_equals_cumulative_final():
    tasks = [{
        "actual_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 1, 31),
        "actual_work": 60.0,
    }]
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 1, 11)),
        (dt.date(2026, 1, 11), dt.date(2026, 1, 21)),
        (dt.date(2026, 1, 21), dt.date(2026, 1, 31)),
    ]
    data_date = dt.date(2026, 1, 31)
    inc = time_phased_ac_increments(tasks, buckets, data_date)
    cum = time_phased_ac(tasks, buckets, data_date)
    assert abs(sum(inc) - cum[-1]) < 0.01


def test_increments_first_equals_cumulative_first():
    tasks = [{
        "actual_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 1, 31),
        "actual_work": 30.0,
    }]
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 1, 16)),
        (dt.date(2026, 1, 16), dt.date(2026, 1, 31)),
    ]
    data_date = dt.date(2026, 1, 31)
    inc = time_phased_ac_increments(tasks, buckets, data_date)
    cum = time_phased_ac(tasks, buckets, data_date)
    assert abs(inc[0] - cum[0]) < 0.01


def test_increments_empty_buckets_returns_empty():
    assert time_phased_ac_increments([], [], dt.date(2026, 1, 1)) == []


def test_increments_single_bucket_matches_cumulative():
    tasks = [{
        "actual_start": dt.date(2026, 1, 1),
        "actual_finish": dt.date(2026, 1, 31),
        "actual_work": 100.0,
    }]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 31))]
    data_date = dt.date(2026, 1, 31)
    inc = time_phased_ac_increments(tasks, buckets, data_date)
    assert len(inc) == 1
    assert abs(inc[0] - 100.0) < 0.01


def test_increments_after_data_date_are_zero():
    """Buckets past data_date contribute 0 (cumulative plateaus)."""
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
    inc = time_phased_ac_increments(tasks, buckets, data_date)
    # bucket 0 has the full pre-cap value, bucket 1+ should be ~0
    assert inc[0] > 0
    assert abs(inc[1]) < 0.01
    assert abs(inc[2]) < 0.01


def test_increments_staggered_tasks_show_step_pattern():
    """Three tasks finishing in distinct buckets -> each bucket gets ~one task."""
    tasks = [
        {"actual_start": dt.date(2026, 1, 1),
         "actual_finish": dt.date(2026, 1, 31), "actual_work": 30.0},
        {"actual_start": dt.date(2026, 3, 1),
         "actual_finish": dt.date(2026, 3, 31), "actual_work": 30.0},
        {"actual_start": dt.date(2026, 5, 1),
         "actual_finish": dt.date(2026, 5, 31), "actual_work": 30.0},
    ]
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 2, 1)),
        (dt.date(2026, 2, 1), dt.date(2026, 3, 1)),
        (dt.date(2026, 3, 1), dt.date(2026, 4, 1)),
        (dt.date(2026, 4, 1), dt.date(2026, 5, 1)),
        (dt.date(2026, 5, 1), dt.date(2026, 6, 1)),
    ]
    data_date = dt.date(2026, 6, 1)
    inc = time_phased_ac_increments(tasks, buckets, data_date)
    # Buckets 0/2/4 should have ~30, buckets 1/3 should have ~0
    assert abs(inc[0] - 30.0) < 0.5
    assert abs(inc[1]) < 0.5
    assert abs(inc[2] - 30.0) < 0.5
    assert abs(inc[3]) < 0.5
    assert abs(inc[4] - 30.0) < 0.5
