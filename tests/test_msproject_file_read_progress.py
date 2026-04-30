"""Test msproject_file read_progress action — Phase 3b integration."""
import os
from msproject_mcp_core import _msp_file_read_progress

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_progress_xml():
    """Sample fixture has no progress entered — all 0%."""
    r = _msp_file_read_progress(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "status_date" in r
    assert "tasks" in r
    for t in r["tasks"]:
        assert "id" in t
        assert "percent_complete" in t


def test_read_progress_task_count():
    """3 real tasks in fixture (summary excluded)."""
    r = _msp_file_read_progress(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert len(r["tasks"]) == 3


def test_read_progress_invalid_file():
    r = _msp_file_read_progress(file_path="/nonexistent.xml")
    assert r["status"] == "error"
