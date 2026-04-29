"""Test msproject_progress set_progress_by_date action.

Phase 3b T57 — bulk retroactive progress via app.UpdateProject.
Probe-confirmed signature on MSP 16.0: UpdateProject(All, UpdateDate, action).
"""
import pytest
import datetime as dt
from msproject_mcp_core import (
    _msp_progress_set_by_date, _msp_progress_get_task,
    _msp_task_add_single,
)


def test_set_by_date_after_full_duration(clean_test_project):
    """Task 5d, progress_date well past task end → 100% (or near) complete."""
    add_r = _msp_task_add_single(name="UpdT-T57", duration="5d")
    assert add_r["status"] == "ok", f"add failed: {add_r}"
    # progress_date well after task end (project starts today, 5 working days)
    target = (dt.date.today() + dt.timedelta(days=30)).isoformat()
    r = _msp_progress_set_by_date(progress_date=target)
    assert r["status"] == "ok", f"set_by_date failed: {r}"
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    assert g["status"] == "ok"
    assert g["progress"]["percent_complete"] >= 95


def test_set_by_date_partial_progress(clean_test_project):
    """Task 10d, progress_date ~5 working days in → 30-75% (MSP working-day calc)."""
    add_r = _msp_task_add_single(name="HalfT-T57", duration="10d")
    assert add_r["status"] == "ok", f"add failed: {add_r}"
    # 7 calendar days = ~5 working days
    target = (dt.date.today() + dt.timedelta(days=7)).isoformat()
    r = _msp_progress_set_by_date(progress_date=target)
    assert r["status"] == "ok", f"set_by_date failed: {r}"
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    assert g["status"] == "ok"
    pct = g["progress"]["percent_complete"]
    assert 30 <= pct <= 75, f"expected 30-75%, got {pct}"


def test_set_by_date_before_task_start(clean_test_project):
    """Task starts today, progress_date = yesterday → 0% complete."""
    # MSP project default start = today; task is 5d from today; progress yesterday → 0%.
    add_r = _msp_task_add_single(name="BeforeT-T57", duration="5d")
    assert add_r["status"] == "ok", f"add failed: {add_r}"
    target = (dt.date.today() - dt.timedelta(days=1)).isoformat()
    r = _msp_progress_set_by_date(progress_date=target)
    assert r["status"] == "ok", f"set_by_date failed: {r}"
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    assert g["status"] == "ok"
    assert g["progress"]["percent_complete"] == 0


def test_set_by_date_invalid_date_format(clean_test_project):
    """Garbage progress_date string → status=error with parse-related message."""
    r = _msp_progress_set_by_date(progress_date="not a date")
    assert r["status"] == "error"
    assert "date" in r["error"].lower() or "parse" in r["error"].lower()
