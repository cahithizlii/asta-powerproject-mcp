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


def test_read_progress_with_assignments():
    """include_assignments parameter accepted (forward compat per plan)."""
    r = _msp_file_read_progress(file_path=MSP_XML, include_assignments=True)
    assert r["status"] == "ok"
    assert "tasks" in r


def test_read_progress_actual_dates_normalized_to_none():
    """Unstarted tasks should have actual_start/finish == None (not 'N/A' string).

    CLAUDE.md RULE 5 — date comparisons must not see 'N/A' strings.
    """
    r = _msp_file_read_progress(file_path=MSP_XML)
    for t in r["tasks"]:
        # Fixture has no progress entered → all dates None
        assert t["actual_start"] is None or isinstance(t["actual_start"], str) and t["actual_start"] != "N/A"
        assert t["actual_finish"] is None or isinstance(t["actual_finish"], str) and t["actual_finish"] != "N/A"
