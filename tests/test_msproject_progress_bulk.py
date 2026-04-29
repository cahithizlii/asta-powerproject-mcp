"""Test msproject_progress bulk_progress_update action (Phase 2b T37 hybrid pattern)."""
import pytest
import time
from msproject_mcp_core import (
    _msp_progress_bulk_update, _msp_progress_get_task,
    _msp_task_add_single,
)


def _make_n_tasks(n: int, prefix: str = "BlkT") -> list:
    """Create n tasks; return list of task IDs."""
    ids = []
    for i in range(n):
        r = _msp_task_add_single(name=f"{prefix}{i:03d}-T62", duration="2d")
        ids.append(r["task_id"])
    return ids


def test_bulk_3_items_com_direct(clean_test_project):
    """3 items → com_direct path."""
    ids = _make_n_tasks(3, "DirectT")
    items = [{"task_id": tid, "percent_complete": 25} for tid in ids]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_direct"
    assert r["count"] == 3
    for tid in ids:
        g = _msp_progress_get_task(task_id=tid)
        assert g["progress"]["percent_complete"] == 25


def test_bulk_10_items_com_batch(clean_test_project):
    """10 items → com_batch path."""
    ids = _make_n_tasks(10, "BatchT")
    items = [{"task_id": tid, "percent_complete": 50} for tid in ids]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_batch"
    assert r["count"] == 10


def test_bulk_25_items_mspdi_path(clean_test_project):
    """25 items → mspdi_bulk path (com_batch_fallback in Phase 3b)."""
    ids = _make_n_tasks(25, "MspdiT")
    items = [{"task_id": tid, "percent_complete": 30,
              "actual_work_h": 4} for tid in ids]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 25


def test_bulk_partial_failure_invalid_task_id(clean_test_project):
    """Mix of valid + invalid task IDs → status=partial."""
    ids = _make_n_tasks(3, "MixT")
    items = [{"task_id": ids[0], "percent_complete": 50},
             {"task_id": 99999, "percent_complete": 50},  # invalid
             {"task_id": ids[1], "percent_complete": 50}]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "partial"
    assert r["count"] == 2
    assert len(r["failures"]) == 1


def test_bulk_perf_50_tasks_under_3s(clean_test_project):
    """50 tasks bulk update <3s (com_batch path proxy via mspdi_bulk)."""
    ids = _make_n_tasks(50, "PerfT")
    items = [{"task_id": tid, "percent_complete": 25} for tid in ids]
    start = time.time()
    r = _msp_progress_bulk_update(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 50
    assert elapsed < 3.0, f"bulk 50 tasks took {elapsed:.2f}s (target <3s)"


def test_bulk_empty_list(clean_test_project):
    r = _msp_progress_bulk_update(items=[])
    assert r["status"] == "ok"
    assert r["path"] == "noop"
    assert r["count"] == 0
