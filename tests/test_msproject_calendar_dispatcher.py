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
