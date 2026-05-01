import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _msp_evm_compute_metrics,
    _msp_evm_forecast,
    _msp_evm_summary,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_compute_metrics_xml():
    r = _msp_evm_compute_metrics(file_path=MSP_XML)
    assert r["status"] == "ok"
    for k in ("bac", "ev", "ac", "pv", "spi", "cpi", "sv", "cv"):
        assert k in r


def test_msp_evm_forecast_xml():
    r = _msp_evm_forecast(file_path=MSP_XML)
    assert r["status"] == "ok"
    for k in ("eac_t1", "eac_t2", "eac_t3", "etc", "vac", "tcpi_bac"):
        assert k in r


def test_msp_evm_summary_xml():
    r = _msp_evm_summary(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "rag" in r
    assert r["rag"] in ("RED", "AMBER", "GREEN")
