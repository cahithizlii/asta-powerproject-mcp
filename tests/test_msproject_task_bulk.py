"""Test msproject_task bulk_add hybrid routing.

SAFETY: All bulk tests use the `clean_test_project` fixture, which creates
a NEW empty MS Project workspace via FileNew. The user's original project
is never modified. On teardown, the test project is closed without saving.
"""
import pytest
import time
from msproject_mcp_core import _msp_task_bulk_add


def test_bulk_3_tasks_com_direct(clean_test_project):
    """3 items -> COM direct path."""
    proj = clean_test_project
    initial = proj.Tasks.Count
    items = [{"name": f"T{i}", "duration": "1d"} for i in range(3)]
    r = _msp_task_bulk_add(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_direct"
    assert r["count"] == 3
    assert proj.Tasks.Count == initial + 3


def test_bulk_15_tasks_com_batch(clean_test_project):
    """15 items -> COM batch path (Calculation manual)."""
    proj = clean_test_project
    items = [{"name": f"B{i}", "duration": "2d"} for i in range(15)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "com_batch"
    assert r["count"] == 15
    assert elapsed < 10, f"Too slow: {elapsed}s"


def test_bulk_30_tasks_mspdi(clean_test_project):
    """30 items -> MSPDI path (T9: COM batch fallback; T10 will wire FileOpen)."""
    proj = clean_test_project
    items = [{"name": f"M{i}", "duration": "1d"} for i in range(30)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 30
    assert elapsed < 30, f"Too slow: {elapsed}s"


def test_bulk_empty_items():
    """Empty list -> noop, status ok. Doesn't need clean_test_project (no creation)."""
    r = _msp_task_bulk_add(items=[])
    assert r["status"] == "ok"
    assert r["count"] == 0


def test_bulk_200_tasks_under_5_sec(clean_test_project):
    """Performance: 200 tasks via MSPDI must finish <5 sec."""
    proj = clean_test_project
    items = [{"name": f"P{i:03d}", "duration": "1d"} for i in range(200)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 200
    assert elapsed < 5.0, f"Bulk 200 took {elapsed:.2f}s (target: <5s)"
    assert proj.Tasks.Count == 200
