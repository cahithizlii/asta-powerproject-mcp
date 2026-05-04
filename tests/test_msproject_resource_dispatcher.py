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


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================


def test_dispatcher_add_invalid_type_returns_error(clean_test_project):
    """Resource type='XX' (not in Work/Material/Cost) → error."""
    r = _run(msproject_resource({"action": "add",
                                 "name": "BadType-T142",
                                 "type": "XX"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "type" in p["error"].lower() or "work" in p["error"].lower()


def test_dispatcher_add_negative_max_units_returns_error(clean_test_project):
    """max_units=-50 → error 'must be >= 0'."""
    r = _run(msproject_resource({"action": "add",
                                 "name": "NegMax-T142",
                                 "type": "Work",
                                 "max_units": -50}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "max_units" in p["error"].lower() or ">= 0" in p["error"]


def test_dispatcher_add_duplicate_name_returns_error(clean_test_project):
    """Calling add twice with same name → second call errors 'already exists'."""
    _run(msproject_resource({"action": "add",
                             "name": "DupName-T142", "type": "Work"}))
    r = _run(msproject_resource({"action": "add",
                                 "name": "DupName-T142", "type": "Work"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "already exists" in p["error"].lower() or "exists" in p["error"].lower()


def test_dispatcher_add_negative_standard_rate_returns_error(clean_test_project):
    """standard_rate=-100 → error 'must be >= 0'."""
    r = _run(msproject_resource({"action": "add",
                                 "name": "NegRate-T142",
                                 "type": "Work",
                                 "standard_rate": -100}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "rate" in p["error"].lower() or ">= 0" in p["error"]


def test_dispatcher_assign_to_nonexistent_resource_returns_error(clean_test_project):
    """assign with resource_id=99999 → error."""
    from msproject_mcp_core import _msp_task_add_single
    add_t = _msp_task_add_single(name="AssignTaskNeg-T142", duration="1d")
    r = _run(msproject_resource({"action": "assign",
                                 "task_id": add_t["task_id"],
                                 "resource_id": 99999}))
    p = json.loads(r)
    assert p["status"] == "error"


def test_dispatcher_delete_nonexistent_resource_returns_error(clean_test_project):
    """delete with resource_id=99999 → error."""
    r = _run(msproject_resource({"action": "delete", "resource_id": 99999}))
    p = json.loads(r)
    assert p["status"] == "error"
