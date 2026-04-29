"""Test msproject_progress clear_progress + clear_all_progress actions."""
import pytest
from msproject_mcp_core import (
    _msp_progress_clear, _msp_progress_clear_all,
    _msp_progress_set_task, _msp_progress_get_task,
    _msp_task_add_single,
)


def test_clear_single_task_progress(clean_test_project):
    """Set 50% then clear → 0%."""
    add_r = _msp_task_add_single(name="ClearOneT-T59", duration="3d")
    _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=50)
    r = _msp_progress_clear(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["task_id"] == add_r["task_id"]
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    assert g["progress"]["percent_complete"] == 0
    assert g["progress"]["actual_start"] is None


def test_clear_unprogressed_task_idempotent(clean_test_project):
    """Clearing a task with no progress → still ok."""
    add_r = _msp_task_add_single(name="NoOpClrT-T59", duration="2d")
    r = _msp_progress_clear(task_id=add_r["task_id"])
    assert r["status"] == "ok"


def test_clear_missing_task(clean_test_project):
    r = _msp_progress_clear(task_id=99999)
    assert r["status"] == "error"


def test_clear_all_progress(clean_test_project):
    """Set progress on 3 tasks, clear_all → all reset."""
    ids = []
    for i in range(3):
        ar = _msp_task_add_single(name=f"AllClrT{i}-T59", duration="2d")
        _msp_progress_set_task(task_id=ar["task_id"], percent_complete=50)
        ids.append(ar["task_id"])
    r = _msp_progress_clear_all()
    assert r["status"] == "ok"
    assert r["cleared_count"] >= 3
    for tid in ids:
        g = _msp_progress_get_task(task_id=tid)
        assert g["progress"]["percent_complete"] == 0


def test_clear_all_when_none_progressed(clean_test_project):
    """clear_all on fresh project → ok with cleared_count 0."""
    r = _msp_progress_clear_all()
    assert r["status"] == "ok"
    assert r["cleared_count"] == 0
