"""Test msproject_calendar assign_to_task action."""
import pytest
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_assign_to_task,
    _msp_task_add_single, _find_task_by_id,
)


def test_assign_calendar_to_task(clean_test_project):
    """Assign a custom calendar to an existing task."""
    proj = clean_test_project
    _msp_calendar_create(name="TaskCal-Phase2a", base_calendar="Standard")
    add_r = _msp_task_add_single(name="CalAssignTask", duration="3d")
    assert add_r["status"] == "ok"
    task_id = add_r["task_id"]
    r = _msp_calendar_assign_to_task(task_id=task_id, calendar_name="TaskCal-Phase2a")
    assert r["status"] == "ok"
    t = _find_task_by_id(proj, task_id)
    # MSP exposes task.Calendar as string name
    assert t.Calendar == "TaskCal-Phase2a"


def test_assign_missing_calendar_errors(clean_test_project):
    add_r = _msp_task_add_single(name="MissingCalTask", duration="1d")
    r = _msp_calendar_assign_to_task(task_id=add_r["task_id"], calendar_name="NoSuch-Phase2a")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_assign_missing_task_errors(clean_test_project):
    _msp_calendar_create(name="OrphanCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_assign_to_task(task_id=99999, calendar_name="OrphanCal-Phase2a")
    assert r["status"] == "error"
    assert "task" in r["error"].lower() and "99999" in r["error"]
