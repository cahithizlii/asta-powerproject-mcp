"""Test progress helpers + constants."""
import pytest
from msproject_mcp_core import (
    _PROGRESS_PCT_FIELDS, _PROGRESS_WORK_FIELDS, _PROGRESS_DURATION_FIELDS,
    _PROGRESS_DATE_FIELDS, _TIMESCALE_UNIT_MAP, _PJ_TIMESCALED_ACTUAL_WORK,
    _normalize_progress_pct, _hours_to_minutes, _minutes_to_hours,
    _validate_actual_dates, _get_assignment_by_resource_id,
    _read_task_progress_dict, _msp_task_add_single,
    _msp_resource_add, _msp_resource_assign,
)


def test_progress_pct_fields_constant():
    assert "percent_complete" in _PROGRESS_PCT_FIELDS
    assert "percent_work_complete" in _PROGRESS_PCT_FIELDS
    assert "physical_pct" in _PROGRESS_PCT_FIELDS


def test_progress_work_fields_constant():
    assert "actual_work_h" in _PROGRESS_WORK_FIELDS
    assert "remaining_work_h" in _PROGRESS_WORK_FIELDS


def test_timescale_unit_map_constant():
    assert _TIMESCALE_UNIT_MAP["day"] == 8
    assert _TIMESCALE_UNIT_MAP["week"] == 6


def test_pj_timescaled_actual_work_const():
    # Probe-confirmed: pjAssignmentTimescaledActualWork == 24 on MSP 16.0
    assert _PJ_TIMESCALED_ACTUAL_WORK == 24


def test_normalize_progress_pct_int_float_str():
    assert _normalize_progress_pct(50) == 50.0
    assert _normalize_progress_pct(50.5) == 50.5
    assert _normalize_progress_pct("50") == 50.0
    assert _normalize_progress_pct("50%") == 50.0
    assert _normalize_progress_pct("50.25") == 50.25


def test_normalize_progress_pct_rejects_out_of_range():
    with pytest.raises(ValueError):
        _normalize_progress_pct(101)
    with pytest.raises(ValueError):
        _normalize_progress_pct(-0.5)
    with pytest.raises(ValueError):
        _normalize_progress_pct("not a number")


def test_hours_to_minutes_round_trip():
    assert _hours_to_minutes(8) == 480
    assert _hours_to_minutes(0.5) == 30
    assert _minutes_to_hours(480) == 8.0
    assert _minutes_to_hours(30) == 0.5


def test_validate_actual_dates_order():
    # Both None → OK
    assert _validate_actual_dates(None, None) is None
    # Only one supplied → OK (other determined by MSP)
    assert _validate_actual_dates("2026-04-01", None) is None
    assert _validate_actual_dates(None, "2026-04-15") is None
    # Both: start <= finish OK
    assert _validate_actual_dates("2026-04-01", "2026-04-15") is None
    # start > finish → error
    err = _validate_actual_dates("2026-04-15", "2026-04-01")
    assert err is not None
    assert "before" in err.lower() or "<=" in err.lower()


def test_get_assignment_by_resource_id_happy(clean_test_project):
    """Assign R to T → lookup returns Assignment object."""
    add_t = _msp_task_add_single(name="GetAsgT-T52", duration="2d")
    add_r = _msp_resource_add(name="ResX-T52", type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_t["task_id"])
    asg = _get_assignment_by_resource_id(t, add_r["resource_id"])
    assert asg is not None


def test_get_assignment_by_resource_id_missing(clean_test_project):
    """No matching assignment → returns None."""
    add_t = _msp_task_add_single(name="NoAsgT-T52", duration="1d")
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_t["task_id"])
    asg = _get_assignment_by_resource_id(t, 99999)
    assert asg is None


def test_read_task_progress_dict_initial_state(clean_test_project):
    """Fresh task — all progress zero/None."""
    add_t = _msp_task_add_single(name="ReadProgT-T52", duration="3d")
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_t["task_id"])
    p = _read_task_progress_dict(t)
    assert p["percent_complete"] == 0
    assert p["percent_work_complete"] == 0
    assert p["actual_start"] is None
    assert p["actual_finish"] is None
    assert p["actual_work_h"] == 0
    assert p["remaining_work_h"] >= 0  # 24h for 3d × 8h
    assert "physical_pct" in p
