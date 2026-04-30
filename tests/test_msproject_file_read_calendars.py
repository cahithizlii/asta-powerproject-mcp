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


def test_read_calendars_fields():
    r = _msp_file_read_calendars(file_path=MSP_XML)
    for c in r["calendars"]:
        assert "name" in c and "is_base" in c
        assert isinstance(c["is_base"], bool)


def test_read_calendars_standard_is_base():
    """Standard calendar in fixture must be a base calendar (CLAUDE.md RULE 0:
    same source data must produce identical reports across .xml and .mpp)."""
    r = _msp_file_read_calendars(file_path=MSP_XML)
    std = next((c for c in r["calendars"] if c["name"] == "Standard"), None)
    assert std is not None, "Standard calendar missing from fixture"
    assert std["is_base"] is True
