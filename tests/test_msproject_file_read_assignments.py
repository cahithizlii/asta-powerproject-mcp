"""Test msproject_file read_assignments action."""
import os
from msproject_mcp_core import _msp_file_read_assignments

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_assignments_xml():
    """3 tasks x 2 resources = 6 assignments."""
    r = _msp_file_read_assignments(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["count"] == 6


def test_read_assignments_fields():
    r = _msp_file_read_assignments(file_path=MSP_XML)
    for a in r["assignments"]:
        for key in ("task_id", "resource_id", "units", "work_h"):
            assert key in a


def test_read_assignments_filter_by_task():
    """Filter by task_id=1 -> only T1's assignments (2: T1-R1, T1-R2)."""
    r = _msp_file_read_assignments(file_path=MSP_XML, task_id=1)
    assert r["status"] == "ok"
    assert all(a["task_id"] == 1 for a in r["assignments"])


def test_read_assignments_invalid_file():
    r = _msp_file_read_assignments(file_path="/nonexistent.xml")
    assert r["status"] == "error"
