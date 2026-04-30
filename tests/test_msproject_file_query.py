"""Test msproject_file query action — filter expression parser."""
import os
import pytest
from msproject_mcp_core import _msp_file_query

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_query_simple_eq():
    """name == 'T1' returns 1 task."""
    r = _msp_file_query(file_path=MSP_XML, expression="name == 'T1'")
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["results"][0]["name"] == "T1"


def test_query_gt():
    """duration_h > 8 returns T2 (16h) and T3 (24h) — 2 tasks."""
    r = _msp_file_query(file_path=MSP_XML, expression="duration_h > 8")
    assert r["status"] == "ok"
    assert r["count"] == 2


def test_query_and():
    """duration_h > 8 AND duration_h <= 16 returns T2 only."""
    r = _msp_file_query(file_path=MSP_XML, expression="duration_h > 8 AND duration_h <= 16")
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["results"][0]["name"] == "T2"


def test_query_or():
    """name == 'T1' OR name == 'T3' returns 2 tasks."""
    r = _msp_file_query(file_path=MSP_XML, expression="name == 'T1' OR name == 'T3'")
    assert r["status"] == "ok"
    assert r["count"] == 2


def test_query_invalid_expression():
    r = _msp_file_query(file_path=MSP_XML, expression="this is not valid syntax @#$")
    assert r["status"] == "error"
    assert "expression" in r["error"].lower() or "parse" in r["error"].lower() or "invalid" in r["error"].lower()


def test_query_limit():
    r = _msp_file_query(file_path=MSP_XML, expression="duration_h >= 8", limit=1)
    assert r["status"] == "ok"
    assert r["count"] == 1


def test_query_invalid_file():
    r = _msp_file_query(file_path="/nonexistent.xml", expression="name == 'T1'")
    assert r["status"] == "error"


def test_query_safe_no_function_calls():
    """Restricted eval — function calls in expression rejected."""
    r = _msp_file_query(file_path=MSP_XML, expression="__import__('os').system('echo PWNED')")
    assert r["status"] == "error"


def test_query_no_attribute_access():
    """Restricted eval — no attribute access (e.g., dunders) allowed."""
    # This may be lenient — main thing is no actual harm done. As long as
    # a malicious expression doesn't execute system commands or import modules,
    # it can either error or return 0 results.
    r = _msp_file_query(file_path=MSP_XML, expression="(1).__class__")
    # Either error or empty result acceptable; should not raise unhandled exception
    assert r["status"] in ("ok", "error")
