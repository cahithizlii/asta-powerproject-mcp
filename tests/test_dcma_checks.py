"""Test pure-math DCMA 14-Point check functions (CLAUDE.md RULE 10).

No fixtures, no COM, no MSP - fully data-driven.
"""
import pytest
from dcma_checks import (
    DCMA_RULES, _DCMA_THRESHOLDS,
    check_no_predecessor, check_no_successor,
)


# ---------- DCMA_RULES metadata ----------

def test_dcma_rules_count():
    assert len(DCMA_RULES) == 14


def test_dcma_rules_have_required_fields():
    for rule in DCMA_RULES:
        for k in ("id", "name", "threshold_label", "threshold_value", "category"):
            assert k in rule


def test_dcma_rules_ids_1_through_14():
    ids = sorted(r["id"] for r in DCMA_RULES)
    assert ids == list(range(1, 15))


def test_dcma_thresholds_count():
    """All 14 rules must have threshold metadata."""
    assert len(_DCMA_THRESHOLDS) == 14
    assert sorted(_DCMA_THRESHOLDS.keys()) == list(range(1, 15))


# ---------- check_no_predecessor (RULE 1) ----------

def _make_task(id, name="T", summary=False, predecessors=None, successors=None):
    return {"id": id, "name": name, "summary": summary,
            "predecessors": predecessors or [], "successors": successors or []}


def test_check_no_predecessor_pass():
    """0% of real tasks lack preds -> PASS.

    Note: simple implementation doesn't treat 'start tasks' specially.
    To pass this rule, structure so all real tasks reference a predecessor
    (test uses circular dummy refs to keep impl simple).
    """
    tasks = [
        _make_task(0, "Project", summary=True),       # root (excluded)
        _make_task(1, "T1", predecessors=[2]),        # all have preds
        _make_task(2, "T2", predecessors=[1]),
        _make_task(3, "T3", predecessors=[2]),
    ]
    r = check_no_predecessor(tasks)
    assert r["id"] == 1
    assert r["status"] == "pass"
    assert r["actual"] == 0.0
    assert r["total_count"] == 3


def test_check_no_predecessor_fail_high_pct():
    """50% of real tasks have no predecessor -> FAIL."""
    tasks = [
        _make_task(1, "T1", predecessors=[]),
        _make_task(2, "T2", predecessors=[]),
        _make_task(3, "T3", predecessors=[1]),
        _make_task(4, "T4", predecessors=[2]),
    ]
    r = check_no_predecessor(tasks)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)
    assert r["failed_count"] == 2
    assert r["total_count"] == 4


def test_check_no_predecessor_excludes_summaries():
    """Summary tasks NOT counted as 'real' tasks."""
    tasks = [
        _make_task(1, "Summary", summary=True, predecessors=[]),
        _make_task(2, "T2", predecessors=[1]),
        _make_task(3, "T3", predecessors=[2]),
    ]
    r = check_no_predecessor(tasks)
    # Real tasks: T2, T3 (both have preds) -> 0% no-pred -> PASS
    assert r["status"] == "pass"
    assert r["actual"] == 0.0


def test_check_no_predecessor_empty():
    r = check_no_predecessor([])
    assert r["status"] == "pass"  # vacuously true
    assert r["total_count"] == 0


def test_check_no_predecessor_returns_failed_ids():
    """Drill-down: failed_task_ids lists offending task IDs."""
    tasks = [
        _make_task(10, "T10", predecessors=[]),
        _make_task(20, "T20", predecessors=[10]),
    ]
    r = check_no_predecessor(tasks)
    assert 10 in r["failed_task_ids"]
    assert 20 not in r["failed_task_ids"]


# ---------- check_no_successor (RULE 2) ----------

def test_check_no_successor_fail_high_pct():
    tasks = [
        _make_task(1, "T1", successors=[]),
        _make_task(2, "T2", successors=[]),
        _make_task(3, "T3", successors=[1]),
        _make_task(4, "T4", successors=[2]),
    ]
    r = check_no_successor(tasks)
    assert r["id"] == 2
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)
    assert r["failed_count"] == 2


def test_check_no_successor_pass_low_pct():
    """1 of 30 tasks (3.3%) has no successor -> PASS (<5%)."""
    tasks = [_make_task(i, f"T{i}", successors=[i + 1]) for i in range(1, 30)]
    tasks.append(_make_task(30, "T30", successors=[]))  # last
    r = check_no_successor(tasks)
    assert r["status"] == "pass"
    assert r["actual"] < 5.0


def test_check_no_successor_excludes_summaries():
    tasks = [
        _make_task(1, "Sum", summary=True, successors=[]),
        _make_task(2, "T2", successors=[3]),
        _make_task(3, "T3", successors=[2]),
    ]
    r = check_no_successor(tasks)
    assert r["status"] == "pass"
    assert r["actual"] == 0.0


def test_check_no_successor_empty():
    r = check_no_successor([])
    assert r["status"] == "pass"
    assert r["total_count"] == 0
