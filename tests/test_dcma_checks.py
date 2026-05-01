"""Test pure-math DCMA 14-Point check functions (CLAUDE.md RULE 10).

No fixtures, no COM, no MSP - fully data-driven.
"""
import pytest
from dcma_checks import (
    DCMA_RULES, _DCMA_THRESHOLDS,
    check_no_predecessor, check_no_successor,
    check_leads, check_lags, check_fs_link_pct,
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


# ---------- T86: Link helpers ----------

def _link(from_id, to_id, type="FS", lag_days=0):
    return {"from_id": from_id, "to_id": to_id, "type": type, "lag_days": lag_days}


# ---------- check_leads (RULE 3: =0) ----------

def test_check_leads_pass_no_leads():
    links = [_link(1, 2, lag_days=0), _link(2, 3, lag_days=2)]
    r = check_leads(links)
    assert r["id"] == 3
    assert r["status"] == "pass"
    assert r["actual"] == 0


def test_check_leads_fail_one_lead():
    links = [_link(1, 2, lag_days=-3)]  # negative = lead
    r = check_leads(links)
    assert r["status"] == "fail"
    assert r["actual"] == 1
    assert r["failed_count"] == 1
    assert len(r["failed_links"]) == 1


def test_check_leads_fail_three_leads():
    links = [_link(1, 2, lag_days=-3), _link(2, 3, lag_days=-1),
             _link(3, 4, lag_days=-2), _link(4, 5, lag_days=0)]
    r = check_leads(links)
    assert r["status"] == "fail"
    assert r["actual"] == 3


def test_check_leads_empty():
    r = check_leads([])
    assert r["status"] == "pass"
    assert r["actual"] == 0
    assert r["total_count"] == 0


# ---------- check_lags (RULE 4: <5%) ----------

def test_check_lags_pass_low_pct():
    """1 lag in 30 links = 3.3% -> PASS."""
    links = [_link(i, i + 1, lag_days=0) for i in range(1, 30)]
    links.append(_link(30, 31, lag_days=2))
    r = check_lags(links)
    assert r["id"] == 4
    assert r["status"] == "pass"
    assert r["actual"] < 5.0


def test_check_lags_fail_high_pct():
    """3 lags in 10 links = 30% -> FAIL."""
    links = [_link(i, i + 1, lag_days=0) for i in range(1, 8)]
    links += [_link(8, 9, lag_days=3), _link(9, 10, lag_days=2),
              _link(10, 11, lag_days=1)]
    r = check_lags(links)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(30.0, rel=1e-2)


def test_check_lags_empty():
    r = check_lags([])
    assert r["status"] == "pass"
    assert r["total_count"] == 0


def test_check_lags_returns_failed_links():
    """Drill-down: failed_links contains lag info."""
    links = [_link(1, 2, lag_days=0), _link(2, 3, lag_days=5)]
    r = check_lags(links)
    assert len(r["failed_links"]) == 1
    assert r["failed_links"][0]["lag_days"] == 5


# ---------- check_fs_link_pct (RULE 5: >90%) ----------

def test_check_fs_link_pct_pass():
    """20 links, 19 FS = 95% -> PASS."""
    links = [_link(i, i + 1, type="FS") for i in range(1, 20)]
    links.append(_link(20, 21, type="SS"))  # 1 SS / 20 = 95% FS
    r = check_fs_link_pct(links)
    assert r["id"] == 5
    assert r["status"] == "pass"
    assert r["actual"] >= 90.0


def test_check_fs_link_pct_fail_too_many_non_fs():
    """5 of 10 = 50% FS -> FAIL."""
    links = [_link(i, i + 1, type="FS") for i in range(1, 6)]
    links += [_link(i, i + 1, type="SS") for i in range(6, 11)]
    r = check_fs_link_pct(links)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)


def test_check_fs_link_pct_empty():
    r = check_fs_link_pct([])
    assert r["status"] == "pass"  # vacuous
    assert r["total_count"] == 0


def test_check_fs_link_pct_case_insensitive():
    """'fs' lowercase should still count as FS."""
    links = [_link(1, 2, type="fs"), _link(2, 3, type="FS")]
    r = check_fs_link_pct(links)
    assert r["actual"] == 100.0


def test_check_fs_link_pct_boundary_exactly_90():
    """Exactly 90% should FAIL (strict >90)."""
    links = [_link(i, i + 1, type="FS") for i in range(1, 10)]  # 9 FS
    links.append(_link(10, 11, type="SS"))                       # 1 SS
    # 9/10 = 90.0 — strict >90 -> FAIL
    r = check_fs_link_pct(links)
    assert r["actual"] == 90.0
    assert r["status"] == "fail"
