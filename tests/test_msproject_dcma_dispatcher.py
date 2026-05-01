"""Test Phase 5b T93 FastMCP dispatcher (msproject_health)."""
import asyncio
import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_health

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def _call(action, **kw):
    raw = _run(msproject_health({"action": action, **kw}))
    return json.loads(raw)


def test_dispatcher_assess_all():
    p = _call("assess_all", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert len(p["rules"]) == 14


def test_dispatcher_summary():
    p = _call("summary", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert p["overall_rag"] in ("green", "amber", "red")


def test_dispatcher_drill_down():
    p = _call("drill_down", file_path=MSP_XML, rule_id=1)
    assert p["status"] == "ok"
    assert "failed_tasks" in p


def test_dispatcher_unknown_action():
    p = _call("nonsense", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_invalid_rule_id():
    p = _call("drill_down", file_path=MSP_XML, rule_id=99)
    assert p["status"] == "error"


def test_dispatcher_compare_no_snapshot():
    """compare without snapshot_path -> graceful ok with empty delta."""
    p = _call("compare", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "current" in p
    assert "delta" in p
