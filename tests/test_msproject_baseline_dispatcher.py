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
