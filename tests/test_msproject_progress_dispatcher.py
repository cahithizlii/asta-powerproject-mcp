"""Test FastMCP msproject_progress dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_progress


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_set_task(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="DispST-T64", duration="3d")
    r = _run(msproject_progress({"action": "set_task_progress",
                                 "task_id": add_r["task_id"],
                                 "percent_complete": 50}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "percent_complete" in p["changes"]


def test_dispatcher_get_task(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="DispGT-T64", duration="2d")
    r = _run(msproject_progress({"action": "get_task_progress",
                                 "task_id": add_r["task_id"]}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "progress" in p


def test_dispatcher_summary(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    _msp_task_add_single(name="DispSumT-T64", duration="2d")
    r = _run(msproject_progress({"action": "summary"}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "project" in p
    assert "bac_h" in p["project"]


def test_dispatcher_status_date(clean_test_project):
    r = _run(msproject_progress({"action": "set_status_date",
                                 "status_date": "2026-04-29"}))
    p = json.loads(r)
    assert p["status"] == "ok"


def test_dispatcher_bulk(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    ids = [_msp_task_add_single(name=f"DispBlk{i}-T64", duration="2d")["task_id"]
           for i in range(3)]
    r = _run(msproject_progress({"action": "bulk_progress_update",
                                 "items": [{"task_id": tid, "percent_complete": 25}
                                           for tid in ids]}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert p["count"] == 3


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_progress({"action": "nonsense"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================


def test_dispatcher_set_task_pct_over_100_returns_error(clean_test_project):
    """percent_complete=150 → error 'must be 0-100'."""
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="PctOver-T142", duration="2d")
    r = _run(msproject_progress({"action": "set_task_progress",
                                 "task_id": add_r["task_id"],
                                 "percent_complete": 150}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "0-100" in p["error"] or "percent" in p["error"].lower()


def test_dispatcher_set_task_pct_negative_returns_error(clean_test_project):
    """percent_complete=-5 → error."""
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="PctNeg-T142", duration="2d")
    r = _run(msproject_progress({"action": "set_task_progress",
                                 "task_id": add_r["task_id"],
                                 "percent_complete": -5}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "0-100" in p["error"] or "percent" in p["error"].lower()


def test_dispatcher_set_task_actual_dates_inverted_returns_error(clean_test_project):
    """actual_start > actual_finish → error."""
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="DateOrder-T142", duration="3d")
    r = _run(msproject_progress({"action": "set_task_progress",
                                 "task_id": add_r["task_id"],
                                 "actual_start": "2026-06-15",
                                 "actual_finish": "2026-06-01"}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_dispatcher_set_task_no_fields_returns_error(clean_test_project):
    """set_task_progress called with only task_id, no fields → error."""
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="NoFields-T142", duration="1d")
    r = _run(msproject_progress({"action": "set_task_progress",
                                 "task_id": add_r["task_id"]}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "fields" in p["error"].lower() or "no progress" in p["error"].lower()


def test_dispatcher_set_status_date_unparseable_returns_error(clean_test_project):
    """status_date='not-a-date' → error from dateutil parser."""
    r = _run(msproject_progress({"action": "set_status_date",
                                 "status_date": "blarg-not-a-date-12345"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "status_date" in p["error"].lower() or "parse" in p["error"].lower()


def test_dispatcher_get_task_nonexistent_returns_error(clean_test_project):
    """get_task_progress with task_id=99999 → error."""
    r = _run(msproject_progress({"action": "get_task_progress",
                                 "task_id": 99999}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "not found" in p["error"].lower() or "task" in p["error"].lower()


def test_dispatcher_time_phased_write_invalid_unit_returns_error(clean_test_project):
    """time_phased_actual_write with unit='hour' (not day/week) → error."""
    r = _run(msproject_progress({"action": "time_phased_actual_write",
                                 "task_id": 1, "resource_id": 1,
                                 "periods": [{"start": "2026-06-01",
                                              "end": "2026-06-02",
                                              "actual_work_h": 8}],
                                 "unit": "hour"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "unit" in p["error"].lower() or "day" in p["error"].lower()
