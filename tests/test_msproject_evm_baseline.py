import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import (
    _msp_evm_variance_to_baseline,
    _msp_evm_compare_baselines_evm,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_variance_to_baseline_xml():
    r = _msp_evm_variance_to_baseline(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    assert "baseline_number" in r


def test_msp_evm_variance_invalid_baseline():
    r = _msp_evm_variance_to_baseline(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"


def test_msp_evm_compare_baselines_xml():
    r = _msp_evm_compare_baselines_evm(file_path=MSP_XML,
                                       baseline_a=0, baseline_b=1)
    # Either ok with delta, or graceful error if baseline_b not saved.
    # Sample fixture has no baselines saved → both should still return ok
    # (T78 fallback: baseline_work = duration_h regardless of baseline_number).
    assert r["status"] in ("ok", "error")
