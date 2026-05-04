"""Test FastMCP msproject_calendar dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_calendar


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_create(clean_test_project):
    r = _run(msproject_calendar({
        "action": "create",
        "name": "DispCal-Phase2a",
        "base_calendar": "Standard",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"


def test_dispatcher_list(clean_test_project):
    r = _run(msproject_calendar({"action": "list"}))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"
    assert "calendars" in parsed


def test_dispatcher_holidays_uzbek(clean_test_project):
    _run(msproject_calendar({
        "action": "create",
        "name": "DispUzbek-Phase2a",
        "base_calendar": "Standard",
    }))
    r = _run(msproject_calendar({
        "action": "holidays_uzbek",
        "calendar_name": "DispUzbek-Phase2a",
        "year": 2026,
    }))
    parsed = json.loads(r)
    assert parsed["status"] in ("ok", "partial")
    assert parsed["count"] == 9


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_calendar({"action": "nonsense"}))
    parsed = json.loads(r)
    assert parsed["status"] == "error"
    assert "Unknown action" in parsed["error"]


def test_dispatcher_calendar_name_alias_for_add_exception(clean_test_project):
    """Dispatcher accepts 'name' as alias for 'calendar_name' on actions
    that natively expect calendar_name."""
    _run(msproject_calendar({
        "action": "create",
        "name": "AliasCal-T29",
        "base_calendar": "Standard",
    }))
    # Now use 'name' instead of 'calendar_name' for add_exception
    r = _run(msproject_calendar({
        "action": "add_exception",
        "name": "AliasCal-T29",
        "exception_name": "Aliased Holiday",
        "start": "2026-06-15",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"
    assert parsed["calendar_name"] == "AliasCal-T29"


def test_dispatcher_name_alias_for_holidays_uzbek(clean_test_project):
    """holidays_uzbek accepts 'name' instead of 'calendar_name'."""
    _run(msproject_calendar({
        "action": "create",
        "name": "AliasUzbek-T29",
        "base_calendar": "Standard",
    }))
    r = _run(msproject_calendar({
        "action": "holidays_uzbek",
        "name": "AliasUzbek-T29",
        "year": 2026,
    }))
    parsed = json.loads(r)
    assert parsed["status"] in ("ok", "partial")
    assert parsed["count"] == 9


def test_dispatcher_calendar_name_alias_reverse(clean_test_project):
    """Reverse: 'create' accepts 'calendar_name' as alias for 'name'."""
    r = _run(msproject_calendar({
        "action": "create",
        "calendar_name": "ReverseAlias-T29",
        "base_calendar": "Standard",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"


def test_dispatcher_does_not_rewrite_unrelated_keys(clean_test_project):
    """Alias logic must NOT touch keys for actions outside the alias set."""
    # 'invalid_action' is not in either alias set — 'name' must NOT be rewritten
    r = _run(msproject_calendar({
        "action": "invalid_action",
        "name": "ShouldStayName",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "error"
    assert "Unknown action" in parsed["error"]


def test_dispatcher_both_name_and_calendar_name_errors(clean_test_project):
    """Specifying both 'name' AND 'calendar_name' returns error (no silent drop)."""
    r = _run(msproject_calendar({
        "action": "list",
        "name": "X",
        "calendar_name": "Y",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "error"
    assert "either" in parsed["error"].lower() or "not both" in parsed["error"].lower()


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================


def test_dispatcher_add_exception_bad_date_format_slashes(clean_test_project):
    """Date with slashes (2026/06/15) instead of YYYY-MM-DD → error."""
    _run(msproject_calendar({
        "action": "create", "name": "BadDate1-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "BadDate1-T142",
        "exception_name": "BadDate",
        "start": "2026/06/15"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "date" in p["error"].lower() or "yyyy-mm-dd" in p["error"].lower()


def test_dispatcher_add_exception_invalid_month_returns_error(clean_test_project):
    """Date with month=13 → error."""
    _run(msproject_calendar({
        "action": "create", "name": "BadDate2-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "BadDate2-T142",
        "exception_name": "BadMonth",
        "start": "2026-13-01"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "date" in p["error"].lower() or "invalid" in p["error"].lower()


def test_dispatcher_add_exception_quarterly_recurrence_returns_error(clean_test_project):
    """recurrence='quarterly' is not in valid set."""
    _run(msproject_calendar({
        "action": "create", "name": "BadRec-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "BadRec-T142",
        "exception_name": "Q",
        "start": "2026-06-15",
        "recurrence": "quarterly"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "recurrence" in p["error"].lower()


def test_dispatcher_add_exception_weekly_without_days_of_week_returns_error(clean_test_project):
    """recurrence='weekly' but no days_of_week → error."""
    _run(msproject_calendar({
        "action": "create", "name": "WklyNoDays-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "WklyNoDays-T142",
        "exception_name": "W",
        "start": "2026-06-15",
        "recurrence": "weekly"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "days_of_week" in p["error"].lower()


def test_dispatcher_add_exception_unknown_day_typo_returns_error(clean_test_project):
    """day_of_week typo 'mondayy' → error."""
    _run(msproject_calendar({
        "action": "create", "name": "DayTypo-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "DayTypo-T142",
        "exception_name": "T",
        "start": "2026-06-15",
        "recurrence": "weekly",
        "days_of_week": ["mondayy"]}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "day_of_week" in p["error"].lower() or "mondayy" in p["error"].lower()


def test_dispatcher_add_exception_bad_hhmm_format_returns_error(clean_test_project):
    """working_hours_start='25:99' → error from _parse_hhmm_time."""
    _run(msproject_calendar({
        "action": "create", "name": "BadHHMM-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "BadHHMM-T142",
        "exception_name": "H",
        "start": "2026-06-15",
        "working": True,
        "working_hours_start": "25:99",
        "working_hours_finish": "17:00"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "working_hours" in p["error"].lower() or "hh:mm" in p["error"].lower()


def test_dispatcher_add_exception_hhmm_no_colon_returns_error(clean_test_project):
    """working_hours_start='0800' (no colon) → error."""
    _run(msproject_calendar({
        "action": "create", "name": "NoColon-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "NoColon-T142",
        "exception_name": "NC",
        "start": "2026-06-15",
        "working": True,
        "working_hours_start": "0800",
        "working_hours_finish": "17:00"}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_dispatcher_add_exception_finish_before_start_returns_error(clean_test_project):
    """finish < start → error."""
    _run(msproject_calendar({
        "action": "create", "name": "BadOrder-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "BadOrder-T142",
        "exception_name": "BO",
        "start": "2026-06-15",
        "finish": "2026-06-01"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "finish" in p["error"].lower() or "start" in p["error"].lower()


def test_dispatcher_add_exception_calendar_not_found_returns_error(clean_test_project):
    """Calendar that doesn't exist → error."""
    r = _run(msproject_calendar({
        "action": "add_exception",
        "calendar_name": "DOES_NOT_EXIST_T142",
        "exception_name": "X",
        "start": "2026-06-15"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "not found" in p["error"].lower()


def test_dispatcher_assign_to_task_nonexistent_task_returns_error(clean_test_project):
    """assign_to_task with task_id=99999 → error."""
    r = _run(msproject_calendar({
        "action": "assign_to_task",
        "task_id": 99999,
        "calendar_name": "Standard"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "task" in p["error"].lower() or "not found" in p["error"].lower()


def test_dispatcher_create_duplicate_calendar_returns_error(clean_test_project):
    """create twice with same name → error 'already exists'."""
    _run(msproject_calendar({
        "action": "create", "name": "DupCal-T142",
        "base_calendar": "Standard"}))
    r = _run(msproject_calendar({
        "action": "create", "name": "DupCal-T142",
        "base_calendar": "Standard"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "already exists" in p["error"].lower() or "exists" in p["error"].lower()
