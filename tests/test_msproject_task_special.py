import pytest
from msproject_mcp_core import _msp_task_add_summary, _msp_task_add_milestone, _msp_task_delete


def test_add_summary(msproject_app):
    r = _msp_task_add_summary(name="Phase 1", duration="10d")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=r["task_id"])


def test_add_milestone(msproject_app):
    r = _msp_task_add_milestone(name="Project Start MS", date="2026-04-26")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=r["task_id"])
