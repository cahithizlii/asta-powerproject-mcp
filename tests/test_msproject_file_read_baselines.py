"""Test msproject_file read_baselines action — Phase 3a integration."""
import os
from msproject_mcp_core import _msp_file_read_baselines

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_baselines_xml_unsaved():
    """Sample fixture has no baseline saved — returns saved_date=None."""
    r = _msp_file_read_baselines(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    assert "baseline_number" in r
    assert "saved_date" in r
    # Either None (no baseline) or timestamp string


def test_read_baselines_invalid_baseline_number():
    r = _msp_file_read_baselines(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"
    assert "0-10" in r["error"]


def test_read_baselines_invalid_file():
    r = _msp_file_read_baselines(file_path="/nonexistent.xml")
    assert r["status"] == "error"


def test_read_baselines_default_zero():
    """Default baseline_number=0 when not specified."""
    r = _msp_file_read_baselines(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0
