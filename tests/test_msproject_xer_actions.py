"""Test Phase 5d T106 6 read action helpers."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _msp_xer_read_tasks, _msp_xer_read_links,
    _msp_xer_read_resources, _msp_xer_read_assignments,
    _msp_xer_read_calendars, _msp_xer_read_progress,
)


def test_xer_read_tasks(sample_cau_xer):
    r = _msp_xer_read_tasks(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 6
    assert len(r["tasks"]) == 6


def test_xer_read_tasks_filter(sample_cau_xer):
    r = _msp_xer_read_tasks(file_path=sample_cau_xer, filters={"task_type": "TT_FinMile"})
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["tasks"][0]["id"] == 1006


def test_xer_read_tasks_limit(sample_cau_xer):
    r = _msp_xer_read_tasks(file_path=sample_cau_xer, limit=3)
    assert r["count"] == 3


def test_xer_read_links(sample_cau_xer):
    r = _msp_xer_read_links(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 5


def test_xer_read_resources(sample_cau_xer):
    r = _msp_xer_read_resources(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 4


def test_xer_read_assignments(sample_cau_xer):
    r = _msp_xer_read_assignments(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 7


def test_xer_read_calendars(sample_cau_xer):
    r = _msp_xer_read_calendars(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["calendars"][0]["day_hr_cnt"] == 9.0


def test_xer_read_progress(sample_cau_xer):
    r = _msp_xer_read_progress(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["status_date"] == "2026-05-01"
    assert len(r["tasks"]) == 6


def test_xer_action_missing_file():
    r = _msp_xer_read_tasks(file_path="/no/such.xer")
    assert r["status"] == "error"


def test_xer_action_no_file_path():
    r = _msp_xer_read_tasks()
    assert r["status"] == "error"
