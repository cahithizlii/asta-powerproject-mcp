"""Test msproject_file read_tasks action (XML path)."""
import os
import pytest
from msproject_mcp_core import _msp_file_read_tasks

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
MSP_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


def test_read_tasks_xml_returns_count():
    """sample_msp.xml has 3 real tasks (T1, T2, T3) - summary task excluded."""
    r = _msp_file_read_tasks(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["count"] == 3
    assert len(r["tasks"]) == 3


def test_read_tasks_xml_has_required_fields():
    r = _msp_file_read_tasks(file_path=MSP_XML)
    t = r["tasks"][0]
    for key in ("id", "name", "duration_h", "start", "finish",
                "percent_complete", "summary"):
        assert key in t


def test_read_tasks_xml_duration_correct():
    """T1=1d -> 8h, T2=2d -> 16h, T3=3d -> 24h."""
    r = _msp_file_read_tasks(file_path=MSP_XML)
    durations = sorted([t["duration_h"] for t in r["tasks"]])
    assert durations == [8.0, 16.0, 24.0]


def test_read_tasks_with_limit():
    r = _msp_file_read_tasks(file_path=MSP_XML, limit=2)
    assert r["status"] == "ok"
    assert r["count"] == 2


def test_read_tasks_invalid_file_errors():
    r = _msp_file_read_tasks(file_path="/nonexistent.xml")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower() or "file" in r["error"].lower()
