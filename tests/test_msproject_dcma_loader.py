"""Test Phase 5b loader extensions (links + floats + constraints)."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _dcma_load_links,
    _dcma_collect_full_data,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_dcma_load_links_xml():
    """File path: reuses _msp_file_read_links."""
    links = _dcma_load_links(file_path=MSP_XML)
    assert isinstance(links, list)
    # Sample fixture has no links; expect empty
    assert len(links) == 0


def test_dcma_collect_full_data_xml():
    r = _dcma_collect_full_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    for k in ("tasks", "links", "assignments", "resources",
              "baseline", "status_date"):
        assert k in r


def test_dcma_collect_full_data_invalid_baseline():
    r = _dcma_collect_full_data(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"


def test_dcma_collect_full_data_invalid_file():
    r = _dcma_collect_full_data(file_path="/nonexistent.xml")
    assert r["status"] == "error"
