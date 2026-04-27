"""Test msproject_schedule operations."""
import pytest
from msproject_mcp_core import (
    _msp_task_add_single, _msp_task_delete,
    _msp_schedule_reschedule, _msp_schedule_level,
    _msp_schedule_set_data_date, _msp_schedule_protect_actuals,
)


def test_reschedule(msproject_app):
    """CalculateProject runs without error."""
    a = _msp_task_add_single(name="ReschedTask", duration="2d")
    r = _msp_schedule_reschedule()
    assert r["status"] == "ok"
    _msp_task_delete(task_id=a["task_id"])


def test_reschedule_with_report_date(msproject_app):
    """Reschedule with report_date sets StatusDate."""
    a = _msp_task_add_single(name="ReschedDate", duration="2d")
    r = _msp_schedule_reschedule(report_date="2026-04-30")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=a["task_id"])


def test_level_resources(msproject_app):
    """LevelAll runs without error (even with no resources)."""
    r = _msp_schedule_level()
    assert r["status"] == "ok"


def test_set_data_date(msproject_app):
    """Set status_date succeeds."""
    r = _msp_schedule_set_data_date(date="2026-05-01")
    assert r["status"] == "ok"
    assert "status_date" in r


def test_protect_actuals(msproject_app):
    """Protect actuals returns ok."""
    r = _msp_schedule_protect_actuals(enable=True)
    assert r["status"] == "ok"
    assert r["actuals_protected"] is True
