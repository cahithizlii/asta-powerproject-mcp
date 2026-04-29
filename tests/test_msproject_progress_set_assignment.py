"""Test msproject_progress set_assignment_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_set_assignment,
    _msp_progress_get_task,
    _msp_task_add_single, _msp_resource_add, _msp_resource_assign,
    _find_task_by_id, _get_assignment_by_resource_id,
)


def _setup_task_with_resource(task_name: str, dur: str = "5d",
                              res_name: str = "COW-T55"):
    """Helper: create task + resource + assignment."""
    add_t = _msp_task_add_single(name=task_name, duration=dur)
    add_r = _msp_resource_add(name=res_name, type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    return add_t["task_id"], add_r["resource_id"]


def test_set_assignment_actual_work(clean_test_project):
    """Write 16h actual on assignment."""
    proj = clean_test_project
    tid, rid = _setup_task_with_resource("AsgWT-T55")
    r = _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                     actual_work_h=16)
    assert r["status"] == "ok"
    assert "actual_work_h" in r["changes"]
    # Readback via direct COM (T56 will provide _msp_progress_get_assignments)
    t = _find_task_by_id(proj, tid)
    asg = _get_assignment_by_resource_id(t, rid)
    assert asg is not None
    # MSP COM ActualWork is in minutes
    assert int(asg.ActualWork) >= 950  # 16h × 60 = 960min, allow drift


def test_assignment_actual_rolls_up_to_task(clean_test_project):
    """Write 24h on assignment → task.ActualWork should reflect (single-resource case)."""
    tid, rid = _setup_task_with_resource("RollupT-T55", dur="5d",
                                          res_name="StlT-T55")
    _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                 actual_work_h=24)
    # Task-level read should see roll-up
    g = _msp_progress_get_task(task_id=tid)
    assert g["progress"]["actual_work_h"] >= 23  # MSP allow small drift


def test_set_assignment_pct_work_complete(clean_test_project):
    tid, rid = _setup_task_with_resource("PctWT-T55", dur="4d",
                                          res_name="MsnT-T55")
    r = _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                     percent_work_complete=50)
    assert r["status"] == "ok"


def test_set_assignment_missing_task_errors(clean_test_project):
    r = _msp_progress_set_assignment(task_id=99999, resource_id=1,
                                     actual_work_h=10)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_set_assignment_missing_resource_errors(clean_test_project):
    """Task exists but no assignment with that resource_id."""
    add_t = _msp_task_add_single(name="NoAsgT-T55", duration="2d")
    r = _msp_progress_set_assignment(task_id=add_t["task_id"], resource_id=99999,
                                     actual_work_h=10)
    assert r["status"] == "error"
    assert "assignment" in r["error"].lower() or "resource" in r["error"].lower()


def test_set_assignment_invalid_pct(clean_test_project):
    tid, rid = _setup_task_with_resource("BadPctAsgT-T55",
                                          res_name="EwiT-T55")
    r = _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                     percent_work_complete=150)
    assert r["status"] == "error"
