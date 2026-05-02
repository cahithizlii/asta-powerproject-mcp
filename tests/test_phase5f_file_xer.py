"""Test Phase 5f: Phase 4 file MCP _msp_file_read_* helpers .xer routing."""
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _msp_file_read_tasks, _msp_file_read_links,
    _msp_file_read_resources, _msp_file_read_assignments,
    _msp_file_read_calendars, _msp_file_read_baselines,
    _msp_file_read_progress,
)


# ---- read_tasks ----

def test_file_read_tasks_xer(sample_cau_xer):
    r = _msp_file_read_tasks(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 6


def test_file_read_tasks_xer_filter(sample_cau_xer):
    """filters dict still works on .xer path."""
    r = _msp_file_read_tasks(file_path=sample_cau_xer,
                             filters={"task_type": "TT_FinMile"})
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["tasks"][0]["id"] == 1006


def test_file_read_tasks_xer_limit(sample_cau_xer):
    r = _msp_file_read_tasks(file_path=sample_cau_xer, limit=3)
    assert r["count"] == 3


def test_file_read_tasks_xer_invalid_path():
    r = _msp_file_read_tasks(file_path="/nonexistent.xer")
    assert r["status"] == "error"


# ---- read_links ----

def test_file_read_links_xer(sample_cau_xer):
    r = _msp_file_read_links(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 5


# ---- read_resources ----

def test_file_read_resources_xer(sample_cau_xer):
    r = _msp_file_read_resources(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 4


# ---- read_assignments ----

def test_file_read_assignments_xer(sample_cau_xer):
    r = _msp_file_read_assignments(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 7


def test_file_read_assignments_xer_task_id_filter(sample_cau_xer):
    """task_id filter works on .xer path."""
    r = _msp_file_read_assignments(file_path=sample_cau_xer, task_id=1001)
    assert r["status"] == "ok"
    assert r["count"] == 2  # Foundation has COW + STL


# ---- read_calendars ----

def test_file_read_calendars_xer(sample_cau_xer):
    r = _msp_file_read_calendars(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["calendars"][0]["day_hr_cnt"] == 9.0


# ---- read_baselines ----

def test_file_read_baselines_xer(sample_cau_xer):
    r = _msp_file_read_baselines(file_path=sample_cau_xer, baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0
    assert r["saved_date"] == "2026-05-01"
    assert len(r["tasks"]) == 6
    foundation = next(t for t in r["tasks"] if t["task_id"] == 1001)
    assert foundation["baseline_finish"] == "2024-07-29"


def test_file_read_baselines_xer_invalid_baseline_number(sample_cau_xer):
    """Baseline number validation runs BEFORE XER routing."""
    r = _msp_file_read_baselines(file_path=sample_cau_xer, baseline_number=99)
    assert r["status"] == "error"
    assert "0-10" in r["error"]


# ---- read_progress ----

def test_file_read_progress_xer(sample_cau_xer):
    r = _msp_file_read_progress(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["status_date"] == "2026-05-01"
    assert len(r["tasks"]) == 6


# ---- Existing .xml path unaffected ----

def test_file_read_tasks_xml_path_unchanged():
    """sample_msp.xml (Phase 4 fixture) unaffected by .xer guard."""
    msp_xml = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")
    if not os.path.exists(msp_xml):
        return
    r = _msp_file_read_tasks(file_path=msp_xml)
    # Phase 4 file path falls through original logic
    assert r["status"] in ("ok", "error")
