"""Test FastMCP msproject_resource dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_resource


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_add(clean_test_project):
    r = _run(msproject_resource({"action": "add", "name": "Disp-T38", "type": "Work"}))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"
    assert parsed["type"] == "Work"


def test_dispatcher_list(clean_test_project):
    _run(msproject_resource({"action": "add", "name": "L1-T38", "type": "Work"}))
    _run(msproject_resource({"action": "add", "name": "L2-T38", "type": "Material",
                            "material_label": "kg"}))
    r = _run(msproject_resource({"action": "list"}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert p["count"] == 2


def test_dispatcher_assign_via_chain(clean_test_project):
    """Chain: add resource -> add task -> assign via dispatcher."""
    add_res = json.loads(_run(msproject_resource({"action": "add",
                                                  "name": "DispW-T38", "type": "Work"})))
    # Use task tool from Phase 1 directly (no msproject_resource for tasks)
    from msproject_mcp_core import _msp_task_add_single
    task_r = _msp_task_add_single(name="DispChainT-T38", duration="1d")
    r = _run(msproject_resource({
        "action": "assign",
        "task_id": task_r["task_id"],
        "resource_id": add_res["resource_id"],
    }))
    p = json.loads(r)
    assert p["status"] == "ok"


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_resource({"action": "nonsense"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]
