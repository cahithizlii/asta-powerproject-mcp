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
    """30 items -> MSPDI path. Duration preserved (Phase 2b TAIL fix)."""
    proj = clean_test_project
    items = [{"name": f"M{i}", "duration": "1d"} for i in range(30)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 30
    assert elapsed < 30, f"Too slow: {elapsed}s"
    # Phase 2b TAIL — MSP MSPDI FileOpen drops Duration; mspdi path applies
    # post-paste Duration set to restore it. 1d = 480 min (8h × 60).
    expected_min = 480
    for i in range(1, 31):
        t = proj.Tasks(i)
        if t is None or t.Summary:
            continue
        assert t.Duration == expected_min, \
            f"Task {i} ({t.Name!r}): Duration={t.Duration} != {expected_min}"


def test_bulk_20_tasks_duration_preserved_2d(clean_test_project):
    """Phase 2b TAIL regression: N=20 (just past mspdi threshold) preserves
    "2d" duration (= 960 min). MSP MSPDI FileOpen import drops Duration silently;
    fix applies post-paste t.Duration = 960 in the same batch.
    """
    proj = clean_test_project
    items = [{"name": f"D{i:02d}", "duration": "2d"} for i in range(20)]
    r = _msp_task_bulk_add(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 20
    expected_min = 960  # 2d × 8h × 60
    zero_count = 0
    for i in range(1, 21):
        t = proj.Tasks(i)
        if t is None or t.Summary:
            continue
        if t.Duration == 0:
            zero_count += 1
        else:
            assert t.Duration == expected_min, \
                f"Task {i}: Duration={t.Duration} != {expected_min}"
    assert zero_count == 0, f"{zero_count}/20 tasks had Duration=0 (fix regressed)"
    # The fix should report 0 set failures
    assert r.get("duration_set_failures", 0) == 0


def test_bulk_empty_items():
    """Empty list -> noop, status ok. Doesn't need clean_test_project (no creation)."""
    r = _msp_task_bulk_add(items=[])
    assert r["status"] == "ok"
    assert r["count"] == 0


def test_bulk_200_tasks_under_7_sec(clean_test_project):
    """Performance: 200 tasks via MSPDI <7 sec.

    Original target was <5s, raised to <7s after Phase 2b TAIL fix added
    ~1-2s post-paste Duration set overhead (200 tasks × ~5ms/task COM
    SetField in batch mode). Tradeoff: correctness over raw speed —
    pre-fix 5s budget was unreachable anyway because MSPDI path was
    falling back to slower COM batch path due to "method not available"
    error, yielding ~9s. Post-fix mspdi_bulk path runs cleanly.
    """
    proj = clean_test_project
    items = [{"name": f"P{i:03d}", "duration": "1d"} for i in range(200)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 200
    assert elapsed < 7.0, f"Bulk 200 took {elapsed:.2f}s (target: <7s)"
    assert proj.Tasks.Count == 200
    # Spot-check: durations preserved on the 200-task path too
    expected_min = 480  # 1d
    sample_idx = (1, 50, 100, 150, 200)
    for i in sample_idx:
        t = proj.Tasks(i)
        if t and not t.Summary:
            assert t.Duration == expected_min, \
                f"Task {i}: Duration={t.Duration} != {expected_min} (Phase 2b TAIL regression)"
