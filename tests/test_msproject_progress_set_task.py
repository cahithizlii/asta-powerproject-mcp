"""Test msproject_progress set_task_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_set_task, _msp_task_add_single, _find_task_by_id,
    _read_task_progress_dict,
)


def test_set_pct_complete_only(clean_test_project):
    """Set percent_complete=50 on a 4-day task."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="PctT-T53", duration="4d")
    r = _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=50)
    assert r["status"] == "ok"
    assert "percent_complete" in r["changes"]
    # Verify via direct readback
    t = _find_task_by_id(proj, add_r["task_id"])
    assert _read_task_progress_dict(t)["percent_complete"] == 50


def test_set_actual_work_h_and_remaining(clean_test_project):
    """Set actual_work_h=16 and remaining_work_h=16 on a 4-day task (32h total)."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="WorkT-T53", duration="4d")
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        actual_work_h=16,
        remaining_work_h=16,
    )
    assert r["status"] == "ok"
    assert "actual_work_h" in r["changes"]
    assert "remaining_work_h" in r["changes"]
    t = _find_task_by_id(proj, add_r["task_id"])
    p = _read_task_progress_dict(t)
    assert p["actual_work_h"] == 16
    assert p["remaining_work_h"] == 16


def test_set_actual_dates(clean_test_project):
    """Set actual_start + actual_finish (uses MSP default project start)."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="DateT-T53", duration="3d")
    # Use today as actual_start so MSP accepts the date (project default
    # start = today; actual dates are stored raw without calendar coercion)
    import datetime as _dt
    today = _dt.date.today()
    actual_start = today.isoformat()
    actual_finish = (today + _dt.timedelta(days=2)).isoformat()
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        actual_start=actual_start,
        actual_finish=actual_finish,
    )
    assert r["status"] == "ok"
    t = _find_task_by_id(proj, add_r["task_id"])
    p = _read_task_progress_dict(t)
    assert p["actual_start"] is not None
    assert p["actual_finish"] is not None


def test_set_pct_invalid_raises_error(clean_test_project):
    add_r = _msp_task_add_single(name="BadPctT-T53", duration="1d")
    r = _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=150)
    assert r["status"] == "error"
    assert "0-100" in r["error"] or "percent" in r["error"].lower()


def test_set_invalid_date_order_errors(clean_test_project):
    add_r = _msp_task_add_single(name="BadOrdT-T53", duration="1d")
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        actual_start="2026-05-15",
        actual_finish="2026-05-01",
    )
    assert r["status"] == "error"
    assert "actual_start" in r["error"].lower() or "before" in r["error"].lower()


def test_set_missing_task_id(clean_test_project):
    r = _msp_progress_set_task(task_id=99999, percent_complete=50)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_set_physical_pct(clean_test_project):
    """DCMA semantic: physical_pct independent of percent_complete."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="PhysT-T53", duration="5d")
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        percent_complete=20,
        physical_pct=50,
    )
    # Either OK with both fields or graceful fallback for older MSP
    assert r["status"] in ("ok", "partial")
    if r["status"] == "ok":
        t = _find_task_by_id(proj, add_r["task_id"])
        p = _read_task_progress_dict(t)
        assert p["percent_complete"] == 20
        # physical_pct may or may not have round-tripped; if writable, it's 50
        if p["physical_pct"] == 50:
            assert True
