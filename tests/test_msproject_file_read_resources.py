"""Test msproject_file read_resources action."""
import os
from msproject_mcp_core import _msp_file_read_resources

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_resources_xml():
    """Sample fixture has 2 work resources R1, R2 (excluding system UID 0)."""
    r = _msp_file_read_resources(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["count"] == 2
    names = [res["name"] for res in r["resources"]]
    assert "R1" in names
    assert "R2" in names


def test_read_resources_fields():
    r = _msp_file_read_resources(file_path=MSP_XML)
    for res in r["resources"]:
        for key in ("id", "name", "type", "max_units"):
            assert key in res


def test_read_resources_invalid_file():
    r = _msp_file_read_resources(file_path="/nonexistent.xml")
    assert r["status"] == "error"
