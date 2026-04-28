"""Test msproject_calendar create action."""
import pytest
from msproject_mcp_core import _msp_calendar_create, _find_calendar_by_name


def test_create_from_standard(clean_test_project):
    """Create 'TestCal' calendar from Standard base."""
    proj = clean_test_project
    r = _msp_calendar_create(name="TestCal-Phase2a", base_calendar="Standard")
    assert r["status"] == "ok"
    assert r["name"] == "TestCal-Phase2a"
    cal = _find_calendar_by_name(proj, "TestCal-Phase2a")
    assert cal is not None


def test_create_duplicate_name_errors(clean_test_project):
    """Creating a calendar with existing name should error."""
    _msp_calendar_create(name="DupCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_create(name="DupCal-Phase2a", base_calendar="Standard")
    assert r["status"] == "error"
    assert "already exists" in r["error"].lower()


def test_create_missing_base_errors(clean_test_project):
    """Base calendar that doesn't exist must error cleanly."""
    r = _msp_calendar_create(name="X-Phase2a", base_calendar="NonExistentBase")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_create_succeeded_but_not_found_guard(clean_test_project, monkeypatch):
    """Cover the 'BaseCalendarCreate succeeded but not found' branch.

    Stubs app.BaseCalendarCreate to a no-op (does nothing). The pre-flight
    name check passes (calendar doesn't exist), then BaseCalendarCreate
    silently returns, then the post-create lookup fails -> guard fires.
    """
    import msproject_mcp_core as core
    real_app = core._connect_app()

    class _NoOp:
        def BaseCalendarCreate(self, **kwargs):
            return None  # silently no-op
        def __getattr__(self, name):
            return getattr(real_app, name)

    fake_app = _NoOp()
    monkeypatch.setattr(core, "_connect_app", lambda: fake_app)

    r = core._msp_calendar_create(name="GuardCal-T30", base_calendar="Standard")
    assert r["status"] == "error"
    assert "succeeded but" in r["error"].lower() or "not found" in r["error"].lower()
