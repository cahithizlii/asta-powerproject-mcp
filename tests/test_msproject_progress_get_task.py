"""Test msproject_progress get_task_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_get_task, _msp_progress_set_task, _msp_task_add_single,
)


def test_get_initial_progress_zero(clean_test_project):
    """Fresh task — all progress fields 0/None."""
    add_r = _msp_task_add_single(name="GetInitT-T54", duration="3d")
    r = _msp_progress_get_task(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["task_id"] == add_r["task_id"]
    p = r["progress"]
    assert p["percent_complete"] == 0
    assert p["actual_start"] is None
    assert p["actual_finish"] is None
    assert p["actual_work_h"] == 0


def test_get_after_pct_set(clean_test_project):
    """Set 50%, read back."""
    add_r = _msp_task_add_single(name="GetSetT-T54", duration="4d")
    _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=50)
    r = _msp_progress_get_task(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["progress"]["percent_complete"] == 50
    # NOTE: actual_work_h depends on resource assignment. Without resources,
    # MSP may keep actual_work_h=0 even with percent_complete>0. Resource-bound
    # work tracking is exercised in T55+T56 (set/get_assignment_progress).
    assert r["progress"]["actual_work_h"] >= 0


def test_get_full_shape_keys_present(clean_test_project):
    """Returned dict has all 11 expected keys."""
    add_r = _msp_task_add_single(name="ShapeT-T54", duration="1d")
    r = _msp_progress_get_task(task_id=add_r["task_id"])
    p = r["progress"]
    expected_keys = {
        "percent_complete", "percent_work_complete", "physical_pct",
        "actual_start", "actual_finish", "stop", "resume",
        "actual_work_h", "remaining_work_h",
        "actual_duration_h", "remaining_duration_h",
    }
    assert expected_keys.issubset(p.keys())


def test_get_missing_task(clean_test_project):
    r = _msp_progress_get_task(task_id=99999)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()
