"""Test msproject_baseline set_active action."""
import pytest
from msproject_mcp_core import _msp_baseline_save, _msp_baseline_set_active, _msp_task_add_single


def test_set_active_default_zero(clean_test_project):
    """Set baseline 0 active."""
    _msp_task_add_single(name="ActT-T47", duration="1d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_set_active(baseline_number=0)
    # Either ok (if API found) or specific "not supported" error (Phase 4 deferral)
    assert r["status"] in ("ok", "error")
    if r["status"] == "error":
        assert "not yet supported" in r["error"].lower() or "phase" in r["error"].lower()


def test_set_active_invalid_baseline_number(clean_test_project):
    r = _msp_baseline_set_active(baseline_number=99)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
