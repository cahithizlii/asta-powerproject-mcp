"""Test FastMCP msproject_baseline dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_baseline


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_save(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    _msp_task_add_single(name="DispBT-T48", duration="2d")
    r = _run(msproject_baseline({"action": "save", "baseline_number": 0}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert p["baseline_number"] == 0


def test_dispatcher_list(clean_test_project):
    r = _run(msproject_baseline({"action": "list"}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "baselines" in p


def test_dispatcher_compare_chain(clean_test_project):
    """Chain: save -> compare -> variance via dispatcher."""
    from msproject_mcp_core import _msp_task_add_single
    _msp_task_add_single(name="ChainT-T48", duration="3d")
    _run(msproject_baseline({"action": "save", "baseline_number": 0}))
    r = _run(msproject_baseline({"action": "compare", "baseline_number": 0}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "summary" in p


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_baseline({"action": "nonsense"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================


def test_dispatcher_save_invalid_baseline_negative_returns_error(clean_test_project):
    """baseline_number=-1 → error 'must be 0-10'."""
    r = _run(msproject_baseline({"action": "save", "baseline_number": -1}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "baseline" in p["error"].lower()


def test_dispatcher_save_invalid_baseline_too_high_returns_error(clean_test_project):
    """baseline_number=99 → error."""
    r = _run(msproject_baseline({"action": "save", "baseline_number": 99}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "baseline" in p["error"].lower()


def test_dispatcher_save_invalid_scope_returns_error(clean_test_project):
    """scope='XX' (not 'all'/'selected') → error."""
    r = _run(msproject_baseline({"action": "save",
                                 "baseline_number": 0, "scope": "XX"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "scope" in p["error"].lower()


def test_dispatcher_clear_invalid_baseline_returns_error(clean_test_project):
    """clear with baseline_number=11 → error."""
    r = _run(msproject_baseline({"action": "clear", "baseline_number": 11}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "baseline" in p["error"].lower()


def test_dispatcher_get_task_baseline_nonexistent_task_returns_error(clean_test_project):
    """get_task_baseline with task_id=99999 → error."""
    r = _run(msproject_baseline({"action": "get_task_baseline",
                                 "task_id": 99999, "baseline_number": 0}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_dispatcher_compare_two_invalid_baseline_a_returns_error(clean_test_project):
    """compare_two with baseline_a=-1 → error."""
    r = _run(msproject_baseline({"action": "compare_two",
                                 "baseline_a": -1, "baseline_b": 0}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "baseline" in p["error"].lower()
