"""Test msproject_baseline summary action — project-level RAG status."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_summary,
    _msp_task_add_single, _msp_task_update,
)


def test_summary_green_no_slip(clean_test_project):
    for i in range(10):
        _msp_task_add_single(name=f"GT{i}-T46", duration="1d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_summary(baseline_number=0)
    assert r["status"] == "ok"
    assert r["project"]["slipped_pct"] == 0.0
    assert r["project"]["schedule_health"] == "green"


def test_summary_amber_when_5_to_20_pct_slipped(clean_test_project):
    """10 tasks, slip 1 → 10% slipped → amber."""
    ids = [_msp_task_add_single(name=f"AT{i}-T46", duration="2d")["task_id"]
           for i in range(10)]
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=ids[0], duration="10d")  # slip 1 task
    r = _msp_baseline_summary(baseline_number=0)
    assert r["status"] == "ok"
    assert 5 < r["project"]["slipped_pct"] <= 20
    assert r["project"]["schedule_health"] == "amber"


def test_summary_red_when_over_20_pct_slipped(clean_test_project):
    """5 tasks, slip 2 → 40% slipped → red."""
    ids = [_msp_task_add_single(name=f"RT{i}-T46", duration="2d")["task_id"]
           for i in range(5)]
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=ids[0], duration="8d")
    _msp_task_update(task_id=ids[1], duration="9d")
    r = _msp_baseline_summary(baseline_number=0)
    assert r["status"] == "ok"
    assert r["project"]["slipped_pct"] > 20
    assert r["project"]["schedule_health"] == "red"


def test_summary_unsaved_baseline_errors(clean_test_project):
    r = _msp_baseline_summary(baseline_number=2)
    assert r["status"] == "error"
