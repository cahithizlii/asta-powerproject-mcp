"""Test msproject_progress time_phased_actual_write + _read actions."""
import pytest
import datetime as dt
from msproject_mcp_core import (
    _msp_progress_time_phased_write,
    _msp_task_add_single, _msp_resource_add, _msp_resource_assign,
)
# _msp_progress_time_phased_read is imported lazily inside read-tests so the
# file collects cleanly during T60 (write-only) commit. T61 adds the read
# helper and lazy imports succeed.
try:
    from msproject_mcp_core import _msp_progress_time_phased_read
    _READ_AVAILABLE = True
except ImportError:
    _READ_AVAILABLE = False
    _msp_progress_time_phased_read = None  # type: ignore


def _setup_5day_assignment(task_name="TPDT-T60"):
    """5d task starting today with 1 resource."""
    add_t = _msp_task_add_single(name=task_name, duration="5d")
    add_r = _msp_resource_add(name=f"TPR-{task_name}", type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    return add_t["task_id"], add_r["resource_id"]


def _today_iso(offset_days=0):
    return (dt.date.today() + dt.timedelta(days=offset_days)).isoformat()


@pytest.mark.skipif(not _READ_AVAILABLE, reason="T61 read helper not yet present")
def test_time_phased_read_empty_actual(clean_test_project):
    """Fresh assignment — daily reads return 0h actual_work for each day."""
    tid, rid = _setup_5day_assignment("ReadEmptyT-T60")
    r = _msp_progress_time_phased_read(
        task_id=tid, resource_id=rid,
        start_date=_today_iso(0), end_date=_today_iso(7),
        unit="day",
    )
    assert r["status"] == "ok"
    assert len(r["periods"]) >= 5  # at least 5 weekdays
    for p in r["periods"]:
        assert p["actual_work_h"] == 0


def test_time_phased_write_3_days(clean_test_project):
    """Write 4h+8h+6h to days 1-2-3 -> readback matches."""
    tid, rid = _setup_5day_assignment("Write3T-T60")
    periods = [
        {"start": _today_iso(0), "end": _today_iso(1), "actual_work_h": 4},
        {"start": _today_iso(1), "end": _today_iso(2), "actual_work_h": 8},
        {"start": _today_iso(2), "end": _today_iso(3), "actual_work_h": 6},
    ]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="day",
    )
    assert w["status"] in ("ok", "partial")
    assert w["written_count"] >= 3
    # Read back (skip if T61 read helper not yet available)
    if not _READ_AVAILABLE:
        pytest.skip("T61 read helper not yet present — readback portion skipped")
    r = _msp_progress_time_phased_read(
        task_id=tid, resource_id=rid,
        start_date=_today_iso(0), end_date=_today_iso(4), unit="day",
    )
    by_day = {p["period_start"][:10]: p["actual_work_h"] for p in r["periods"]}
    # Allow MSP 0.5h drift
    assert by_day.get(_today_iso(0), 0) >= 3.5
    assert by_day.get(_today_iso(1), 0) >= 7.5
    assert by_day.get(_today_iso(2), 0) >= 5.5


def test_time_phased_write_weekly(clean_test_project):
    """Weekly bucket: write 30h to week-1."""
    tid, rid = _setup_5day_assignment("WriteWkT-T60")
    periods = [{"start": _today_iso(0), "end": _today_iso(7), "actual_work_h": 30}]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="week",
    )
    assert w["status"] in ("ok", "partial")


def test_time_phased_write_no_overlap_period_handles_gracefully(clean_test_project):
    """Period far outside task date range -> MSP returns slots but they have no
    effect on assignment totals. Write should not crash; status is ok/partial/error.

    Note: MSP TimeScaleData returns valid slot count even for out-of-range
    dates (probe-confirmed). Slots are writable but the Value writes don't
    propagate to assignment.ActualWork — they're effectively no-ops. The
    helper should not raise; status just reflects whether COM rejected the
    write at any layer.
    """
    tid, rid = _setup_5day_assignment("OutT-T60")
    periods = [{"start": _today_iso(365), "end": _today_iso(366),
                "actual_work_h": 4}]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="day",
    )
    # Must not crash and must return one of the well-defined statuses
    assert w["status"] in ("ok", "partial", "error")
    assert "written_count" in w
    assert isinstance(w.get("failures"), list)


@pytest.mark.skipif(not _READ_AVAILABLE, reason="T61 read helper not yet present")
def test_time_phased_read_invalid_unit(clean_test_project):
    tid, rid = _setup_5day_assignment("BadUnitT-T60")
    r = _msp_progress_time_phased_read(
        task_id=tid, resource_id=rid,
        start_date=_today_iso(0), end_date=_today_iso(7),
        unit="quarter",
    )
    assert r["status"] == "error"
    assert "unit" in r["error"].lower()


@pytest.mark.skipif(not _READ_AVAILABLE, reason="T61 read helper not yet present")
def test_time_phased_missing_assignment(clean_test_project):
    """Task with no assignment -> error."""
    add_t = _msp_task_add_single(name="NoAsgTPD-T60", duration="3d")
    r = _msp_progress_time_phased_read(
        task_id=add_t["task_id"], resource_id=99999,
        start_date=_today_iso(0), end_date=_today_iso(4), unit="day",
    )
    assert r["status"] == "error"


def test_time_phased_write_invalid_date_format(clean_test_project):
    tid, rid = _setup_5day_assignment("BadDateT-T60")
    periods = [{"start": "not-a-date", "end": _today_iso(2),
                "actual_work_h": 4}]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="day",
    )
    assert w["status"] in ("error", "partial")
