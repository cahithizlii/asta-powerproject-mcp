"""Test msproject_task add action."""
import pytest
from msproject_mcp_core import _msp_task_add_single


def test_add_single_task(msproject_app):
    """Adding 1 task -> ActiveProject.Tasks.Count incremented."""
    proj = msproject_app.ActiveProject
    initial = proj.Tasks.Count
    result = _msp_task_add_single(name="Test Task A", duration="3d")
    assert result["status"] == "ok"
    assert proj.Tasks.Count == initial + 1
    found = None
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t and t.Name == "Test Task A":
            found = t
            break
    assert found is not None
    found.Delete()


def test_add_milestone(msproject_app):
    """Milestone has 0d duration."""
    proj = msproject_app.ActiveProject
    result = _msp_task_add_single(name="MS Test", duration="0d", milestone=True)
    assert result["status"] == "ok"
    found = None
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t and t.Name == "MS Test":
            found = t
            break
    assert found is not None and found.Milestone
    found.Delete()
