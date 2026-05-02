"""Phase 6.1 T115b/c — _msp_evm_validate_currency_mode adapter + dispatcher.

Integration tests for the multi-source currency cross-validation adapter
and dispatcher action wiring. Verifies:
- 4-mode primary output schema
- XER currency_code extraction from ERMHDR
- XER assignments routed through RULE 3 detector
- cross_validation block shape (consensus_mode, confidence, conflicts,
  warnings, source_counts)
- Backward compatibility: legacy detect_currency_mode action still
  returns 2-mode 'cost'/'hours' shape
"""
import asyncio
import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _msp_evm_detect_currency_mode,
    _msp_evm_validate_currency_mode,
    msproject_evm,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run_evm(action, **kw):
    """Helper: run async msproject_evm dispatcher and parse JSON."""
    raw = asyncio.run(msproject_evm({"action": action, **kw}))
    return json.loads(raw)


# === XML / MSPDI source ===

def test_validate_currency_mode_xml_status_ok():
    r = _msp_evm_validate_currency_mode(file_path=MSP_XML)
    assert r["status"] == "ok"


def test_validate_currency_mode_xml_primary_mode_in_4mode_set():
    r = _msp_evm_validate_currency_mode(file_path=MSP_XML)
    assert r["primary_mode"] in ("cost", "hours", "mixed", "uncertain")


def test_validate_currency_mode_xml_no_xer_currency_code():
    """MSPDI source has no XER ERMHDR — currency_code must be None."""
    r = _msp_evm_validate_currency_mode(file_path=MSP_XML)
    assert r["currency_code"] is None
    assert r["sources"]["xer_assignments"] is None
    assert r["sources"]["currency_header"] is None


def test_validate_currency_mode_xml_cross_validation_schema():
    r = _msp_evm_validate_currency_mode(file_path=MSP_XML)
    cv = r["cross_validation"]
    assert "consensus_mode" in cv
    assert cv["confidence"] in ("high", "medium", "low")
    assert isinstance(cv["conflicts"], list)
    assert isinstance(cv["warnings"], list)
    assert isinstance(cv["source_counts"], dict)


# === XER source — uses sample_cau_xer fixture from conftest ===

def test_validate_currency_mode_xer_currency_code_extracted(sample_cau_xer):
    r = _msp_evm_validate_currency_mode(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["currency_code"] == "USD"


def test_validate_currency_mode_xer_assignments_rule3_hours(sample_cau_xer):
    """CAU sample: TASKRSRC target_cost == target_qty in all rows → hours.

    This matches the CAU project pattern from CLAUDE.md (cost not loaded
    even though USD currency code is set).
    """
    r = _msp_evm_validate_currency_mode(file_path=sample_cau_xer)
    assert r["sources"]["xer_assignments"] == "hours"


def test_validate_currency_mode_xer_sources_breakdown(sample_cau_xer):
    r = _msp_evm_validate_currency_mode(file_path=sample_cau_xer)
    sources = r["sources"]
    assert "tasks_resources" in sources
    assert "xer_assignments" in sources
    assert "currency_header" in sources
    assert sources["currency_header"] == "USD"


def test_validate_currency_mode_xer_primary_mode_resolves(sample_cau_xer):
    """Primary mode must be one of the 4 valid modes (not None / empty)."""
    r = _msp_evm_validate_currency_mode(file_path=sample_cau_xer)
    assert r["primary_mode"] in ("cost", "hours", "mixed", "uncertain")


# === Dispatcher action wiring ===

def test_dispatcher_validate_currency_mode_action(sample_cau_xer):
    r = _run_evm("validate_currency_mode", file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert "primary_mode" in r
    assert "cross_validation" in r
    assert r["currency_code"] == "USD"


def test_dispatcher_detect_currency_mode_backward_compat(sample_cau_xer):
    """Legacy detect_currency_mode action must still return 2-mode shape:
    {status, mode: 'cost'|'hours'} — schema change would break old callers.
    """
    r = _run_evm("detect_currency_mode", file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["mode"] in ("cost", "hours"), \
        "Legacy detect_currency_mode must NOT return 'mixed' or 'uncertain'"


def test_dispatcher_unknown_action_lists_validate(sample_cau_xer):
    """Error message for unknown action must list validate_currency_mode."""
    r = _run_evm("definitely_not_an_action")
    assert r["status"] == "error"
    assert "validate_currency_mode" in r["error"]
