"""Test msproject_file read_links action (XML path)."""
import os
import shutil
import sys
from msproject_mcp_core import _msp_file_read_links

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
MSP_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


def test_read_links_xml_returns_status():
    """sample_msp.xml has no links by default - count=0."""
    r = _msp_file_read_links(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "count" in r
    assert "links" in r
    assert isinstance(r["links"], list)
    assert r["count"] == 0


def test_read_links_invalid_file_errors():
    r = _msp_file_read_links(file_path="/nonexistent.xml")
    assert r["status"] == "error"


def test_read_links_xml_with_injected_link(tmp_path):
    """Clone fixture, inject T1->T2 FS link, verify read_links extracts it."""
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    from mspdi_parser import MspdiProject

    src = MSP_XML
    dst = tmp_path / "with_link.xml"
    shutil.copy(src, str(dst))

    proj = MspdiProject(str(dst))
    res = proj.add_link(predecessor_id=1, successor_id=2, link_type="FS", lag_str="0d")
    assert "error" not in res, f"add_link failed: {res}"
    proj.save(str(dst))

    r = _msp_file_read_links(file_path=str(dst))
    assert r["status"] == "ok"
    assert r["count"] == 1
    link = r["links"][0]
    assert link["from_id"] == 1
    assert link["to_id"] == 2
    assert link["type"] == "FS"
    assert "lag_days" in link
