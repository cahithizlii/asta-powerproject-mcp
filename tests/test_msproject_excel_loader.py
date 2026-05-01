"""Test Phase 5c T99 single-collect data aggregator."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import _excel_collect_full_data

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_excel_collect_full_data_xml():
    r = _excel_collect_full_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    for k in ("tasks", "evm", "dcma"):
        assert k in r


def test_excel_collect_full_data_evm_subkeys():
    r = _excel_collect_full_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    for k in ("metrics", "forecast", "earned_schedule", "rag", "time_phased"):
        assert k in r["evm"]


def test_excel_collect_full_data_dcma_subkeys():
    r = _excel_collect_full_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    for k in ("rules", "summary", "drilldowns"):
        assert k in r["dcma"]


def test_excel_collect_full_data_invalid_baseline():
    r = _excel_collect_full_data(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"


def test_excel_collect_full_data_invalid_file():
    r = _excel_collect_full_data(file_path="/nonexistent.xml")
    assert r["status"] == "error"
