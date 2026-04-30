"""Test msproject_file read_calendars action."""
import os
from msproject_mcp_core import _msp_file_read_calendars

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_calendars_xml():
    r = _msp_file_read_calendars(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "calendars" in r
    assert r["count"] >= 1  # at least the Standard calendar


def test_read_calendars_invalid_file():
    r = _msp_file_read_calendars(file_path="/nonexistent.xml")
    assert r["status"] == "error"
