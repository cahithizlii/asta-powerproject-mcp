"""P0 #3 — BAC cross-validation (RULE 16.A, ALFB1 9x defense)."""
import pytest
from evm_math import cross_validate_bac


# ---------- pure helper ----------

def test_cross_validate_bac_match_within_tolerance():
    r = cross_validate_bac(1000.0, 1005.0, tolerance=0.01)
    assert r["match"] is True
    assert r["severity"] == "none"
    assert r["warning"] is None
    assert r["rel_diff"] <= 0.01


def test_cross_validate_bac_high_severity_9x():
    """The ALFB1 case: 277,640 reported vs 2,505,038 raw."""
    r = cross_validate_bac(277640.0, 2505038.0)
    assert r["match"] is False
    assert r["severity"] == "high"
    assert r["ratio"] == pytest.approx(2505038.0 / 277640.0, rel=1e-6)
    assert "RULE 16.A" in r["warning"]


def test_cross_validate_bac_low_severity_small_drift():
    r = cross_validate_bac(1000.0, 1050.0, tolerance=0.01)  # 5% drift
    assert r["match"] is False
    assert r["severity"] == "low"   # >1% but <=10%


def test_cross_validate_bac_no_independent_source():
    r = cross_validate_bac(1000.0, None)
    assert r["match"] is True        # cannot validate -> not a failure
    assert r["severity"] == "none"
    assert r["bac_independent"] is None


def test_cross_validate_bac_primary_zero_but_independent_present():
    """Missing rollup: pipeline BAC 0, raw sum large."""
    r = cross_validate_bac(0.0, 2505038.0)
    assert r["match"] is False
    assert r["severity"] == "high"
    assert "not" in r["warning"].lower()


def test_cross_validate_bac_custom_tolerance_passes():
    r = cross_validate_bac(1000.0, 1050.0, tolerance=0.10)  # 5% within 10%
    assert r["match"] is True


# ---------- action wired through XER fixture ----------

def test_verify_action_detects_xer_mismatch(sample_cau_xer):
    from msproject_mcp_core import _msp_evm_verify
    r = _msp_evm_verify(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["assignment_count"] == 7
    # CAU fixture: sum(target_qty) = 180+1000+360+180+180+180+360 = 2440
    assert r["bac_independent"] == pytest.approx(2440.0)
    # primary = sum(duration_h) of non-summary tasks = 180+360+180+180+360+0
    assert r["bac_primary"] == pytest.approx(1260.0)
    assert r["match"] is False
    assert r["severity"] == "high"


def test_verify_action_clean_match(tmp_path):
    """XER where each task's target_qty equals its duration_h -> match."""
    from tests._xer_fixture_builders import write_synthetic_xer
    from msproject_mcp_core import _msp_evm_verify
    content = (
        "ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
        "%T\tCALENDAR\n%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt\n"
        "%R\t1\tStd\t8.0\t40.0\n"
        "%T\tTASK\n%F\ttask_id\ttask_name\ttask_type\ttarget_drtn_hr_cnt"
        "\ttarget_start_date\ttarget_end_date\tphys_complete_pct\n"
        "%R\t1\tA\tTT_Task\t100.0\t2024-01-01 08:00\t2024-01-15 17:00\t0\n"
        "%R\t2\tB\tTT_Task\t200.0\t2024-01-16 08:00\t2024-02-01 17:00\t0\n"
        "%T\tTASKRSRC\n%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty"
        "\tact_reg_qty\ttarget_cost\tact_reg_cost\n"
        "%R\t1\t1\t10\t100.0\t0\t0\t0\n"
        "%R\t2\t2\t10\t200.0\t0\t0\t0\n"
        "%E\n"
    )
    path = write_synthetic_xer(content, "bac_clean.xer")
    try:
        r = _msp_evm_verify(file_path=path)
        assert r["bac_primary"] == pytest.approx(300.0)
        assert r["bac_independent"] == pytest.approx(300.0)
        assert r["match"] is True
        assert r["severity"] == "none"
    finally:
        import os
        os.remove(path)
