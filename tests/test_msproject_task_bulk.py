"""Test msproject_task bulk_add hybrid routing."""
import pytest
import time
from msproject_mcp_core import _msp_task_bulk_add


def _cleanup_all_tasks(proj):
    while proj.Tasks.Count > 0:
        proj.Tasks(1).Delete()


def test_bulk_3_tasks_com_direct(msproject_app):
    """3 items -> COM direct path."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"T{i}", "duration": "1d"} for i in range(3)]
    r = _msp_task_bulk_add(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_direct"
    assert r["count"] == 3
    _cleanup_all_tasks(msproject_app.ActiveProject)


def test_bulk_15_tasks_com_batch(msproject_app):
    """15 items -> COM batch path (Calculation manual)."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"B{i}", "duration": "2d"} for i in range(15)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "com_batch"
    assert r["count"] == 15
    assert elapsed < 10, f"Too slow: {elapsed}s"
    _cleanup_all_tasks(msproject_app.ActiveProject)


def test_bulk_30_tasks_mspdi(msproject_app):
    """30 items -> MSPDI path (T9: COM batch fallback; T10 will wire FileOpen)."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"M{i}", "duration": "1d"} for i in range(30)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 30
    assert elapsed < 30, f"Too slow: {elapsed}s"
    _cleanup_all_tasks(msproject_app.ActiveProject)


def test_bulk_empty_items(msproject_app):
    """Empty list -> noop, status ok."""
    r = _msp_task_bulk_add(items=[])
    assert r["status"] == "ok"
    assert r["count"] == 0


def test_bulk_200_tasks_under_5_sec(msproject_app):
    """Performance: 200 tasks via MSPDI must finish <5 sec."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"P{i:03d}", "duration": "1d"} for i in range(200)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 200
    assert elapsed < 5.0, f"Bulk 200 took {elapsed:.2f}s (target: <5s)"
    proj = msproject_app.ActiveProject
    assert proj.Tasks.Count == 200
    _cleanup_all_tasks(msproject_app.ActiveProject)
