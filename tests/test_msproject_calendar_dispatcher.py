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
