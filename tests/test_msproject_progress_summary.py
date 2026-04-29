"""Test msproject_progress summary action — EVM-ready aggregate."""
import pytest
from msproject_mcp_core import (
    _msp_progress_summary, _msp_progress_set_task,
    _msp_task_add_single, _msp_resource_add, _msp_resource_assign,
)


def _add_task_with_resource(task_name: str, dur: str, res_name: str) -> int:
    """Helper: create task + resource + assignment so that t.Work is non-zero
    (EVM BAC requires resource-loaded work — CLAUDE.md RULE 3)."""
    add_t = _msp_task_add_single(name=task_name, duration=dur)
    add_r = _msp_resource_add(name=res_name, type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    return add_t["task_id"]


def test_summary_empty_project(clean_test_project):
    r = _msp_progress_summary()
    assert r["status"] == "ok"
    p = r["project"]
    assert p["bac_h"] == 0
    assert p["acwp_h"] == 0
    assert p["task_count"] == 0
    assert p["completed_count"] == 0
    assert p["in_progress_count"] == 0


def test_summary_no_progress_zero_acwp(clean_test_project):
    """5 tasks of 4d each (32h work each, resource-loaded) → BAC=160h, ACWP=0."""
    for i in range(5):
        _add_task_with_resource(f"NoProgT{i}-T63", "4d", f"R{i}-T63a")
    r = _msp_progress_summary()
    p = r["project"]
    assert p["bac_h"] == 160
    assert p["acwp_h"] == 0
    assert p["not_started_count"] == 5


def test_summary_partial_progress(clean_test_project):
    """5 tasks (32h each = 160h BAC, resource-loaded). 2 complete, 1 at 50%."""
    ids = []
    for i in range(5):
        tid = _add_task_with_resource(f"PartT{i}-T63", "4d", f"R{i}-T63b")
        ids.append(tid)
    _msp_progress_set_task(task_id=ids[0], percent_complete=100)
    _msp_progress_set_task(task_id=ids[1], percent_complete=100)
    _msp_progress_set_task(task_id=ids[2], percent_complete=50)
    r = _msp_progress_summary()
    p = r["project"]
    assert p["bac_h"] == 160
    assert p["completed_count"] == 2
    assert p["in_progress_count"] == 1
    assert p["not_started_count"] == 2
    # ACWP ≈ 32 + 32 + 16 = 80h
    assert 75 <= p["acwp_h"] <= 85
    # Project pct ~ 50%
    assert 45 <= p["project_percent_complete"] <= 55


def test_summary_fully_complete(clean_test_project):
    """3 tasks all 100% → project 100%."""
    ids = []
    for i in range(3):
        ar = _msp_task_add_single(name=f"DoneT{i}-T63", duration="2d")
        ids.append(ar["task_id"])
    for tid in ids:
        _msp_progress_set_task(task_id=tid, percent_complete=100)
    r = _msp_progress_summary()
    p = r["project"]
    assert p["completed_count"] == 3
    assert p["project_percent_complete"] >= 99


def test_summary_status_date_present(clean_test_project):
    """status_date is in the summary if set."""
    from msproject_mcp_core import _msp_progress_set_status_date
    _msp_task_add_single(name="StTask-T63", duration="2d")
    _msp_progress_set_status_date(status_date="2026-04-29")
    r = _msp_progress_summary()
    assert r["project"]["status_date"] is not None
