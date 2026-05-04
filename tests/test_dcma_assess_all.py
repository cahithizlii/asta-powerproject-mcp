"""Test DCMA assess_all aggregator + compute_overall_rag (RULE 10 + RULE 12)."""
import pytest
from dcma_checks import assess_all, compute_overall_rag, DCMA_RULES


# ---------- assess_all returns 14-rule envelope ----------

def test_assess_all_returns_14_rules_empty_inputs():
    r = assess_all(tasks=[], links=[], assignments=[],
                   baseline=None, status_date=None)
    assert "rules" in r
    assert len(r["rules"]) == 14


def test_assess_all_rules_have_status():
    r = assess_all(tasks=[], links=[], assignments=[],
                   baseline=None, status_date=None)
    for rule in r["rules"]:
        assert rule["status"] in ("pass", "fail")
        assert rule["id"] in range(1, 15)


def test_assess_all_rules_ids_complete():
    r = assess_all(tasks=[], links=[], assignments=[],
                   baseline=None, status_date=None)
    rule_ids = sorted(rule["id"] for rule in r["rules"])
    assert rule_ids == list(range(1, 15))


def test_assess_all_summary_keys():
    r = assess_all(tasks=[], links=[], assignments=[],
                   baseline=None, status_date=None)
    assert "summary" in r
    for k in ("pass_count", "fail_count", "overall_rag", "executive_text"):
        assert k in r["summary"]


def test_assess_all_summary_counts_match():
    r = assess_all(tasks=[], links=[], assignments=[],
                   baseline=None, status_date=None)
    s = r["summary"]
    assert s["pass_count"] + s["fail_count"] == 14


def test_assess_all_with_realistic_inputs():
    """Smoke test with non-empty inputs - should run all 14 checks without error."""
    tasks = [
        {"id": 1, "name": "T1", "summary": False, "predecessors": [],
         "successors": [2], "constraint_type": 0, "total_slack_days": 5,
         "duration_h": 80, "critical": True, "start": "2026-01-01",
         "finish": "2026-01-10", "baseline_finish": "2026-01-10",
         "percent_complete": 100},
        {"id": 2, "name": "T2", "summary": False, "predecessors": [1],
         "successors": [], "constraint_type": 0, "total_slack_days": 0,
         "duration_h": 80, "critical": True, "start": "2026-01-11",
         "finish": "2026-01-20", "baseline_finish": "2026-01-20",
         "percent_complete": 50},
    ]
    links = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0}]
    assignments = [{"task_id": 1, "resource_id": 1},
                   {"task_id": 2, "resource_id": 1}]
    r = assess_all(tasks=tasks, links=links, assignments=assignments,
                   baseline=None, status_date="2026-02-01")
    assert len(r["rules"]) == 14
    assert r["summary"]["pass_count"] + r["summary"]["fail_count"] == 14


# ---------- compute_overall_rag (RAG: pass_count >= 12 GREEN, 8-11 AMBER, <8 RED) ----------

def _make_rule_results(pass_count):
    """Synthetic rules list with N passing, rest failing."""
    return [{"id": i, "status": "pass" if i <= pass_count else "fail"}
            for i in range(1, 15)]


def test_rag_green_above_12_pass():
    assert compute_overall_rag(_make_rule_results(13)) == "green"
    assert compute_overall_rag(_make_rule_results(14)) == "green"


def test_rag_green_exactly_12():
    """Boundary: exactly 12 pass -> GREEN (>=12)."""
    assert compute_overall_rag(_make_rule_results(12)) == "green"


def test_rag_amber_8_to_11_pass():
    assert compute_overall_rag(_make_rule_results(8)) == "amber"
    assert compute_overall_rag(_make_rule_results(11)) == "amber"


def test_rag_red_below_8_pass():
    assert compute_overall_rag(_make_rule_results(7)) == "red"
    assert compute_overall_rag(_make_rule_results(0)) == "red"


def test_rag_empty_rules_red():
    """No rules at all -> 0 pass -> RED."""
    assert compute_overall_rag([]) == "red"


# ---------- assess_all integration: RAG follows pass_count ----------

def test_assess_all_executive_text_present():
    r = assess_all(tasks=[], links=[], assignments=[],
                   baseline=None, status_date=None)
    assert isinstance(r["summary"]["executive_text"], str)
    assert len(r["summary"]["executive_text"]) > 0


# === Phase 11.1 T141: gap-fill ===

def test_assess_all_executive_all_pass_message():
    """All 14 checks pass -> 'All 14 DCMA rules pass. Schedule health: GREEN.' (line 474)."""
    # Need at least 1 critical real task to satisfy check_critical_path,
    # plus default thresholds to all pass.
    tasks = [{
        "id": 1, "name": "T1", "summary": False,
        "predecessors": [99], "successors": [99],  # has both pred & succ
        "constraint_type": 0,
        "total_slack_days": 0,
        "duration_h": 80,
        "critical": True,                  # satisfies critical_path
        "start": "2026-01-01", "finish": "2026-01-10",
    }]
    links = []
    assignments = [{"task_id": 1, "resource_id": 1}]
    r = assess_all(tasks=tasks, links=links, assignments=assignments,
                   baseline=None, status_date=None)
    assert r["summary"]["fail_count"] == 0
    assert r["summary"]["pass_count"] == 14
    assert "All 14 DCMA rules pass" in r["summary"]["executive_text"]


def test_assess_all_executive_more_than_5_failures_truncated():
    """Failure summary shows '...' when >5 rules fail."""
    # Build a synthetic: many no-pred/no-succ tasks -> trigger multiple fails
    tasks = []
    for i in range(1, 21):
        tasks.append({
            "id": i, "name": f"T{i}", "summary": False,
            "predecessors": [], "successors": [],   # both fail rule 1 and 2
            "constraint_type": 5,                    # hard constraint -> fail rule 6
            "total_slack_days": 100,                 # high float -> fail rule 7
            "duration_h": 5000,                      # high duration -> fail rule 9
            "critical": False,                       # fail rule 13
            "start": "2026-01-10", "finish": "2026-01-01",  # invalid date -> rule 10
        })
    # Lags > 5% to trigger rule 4; leads to trigger rule 3
    links = [
        {"from_id": 1, "to_id": 2, "type": "FS", "lag_days": -1},
        {"from_id": 2, "to_id": 3, "type": "FS", "lag_days": 5},
        {"from_id": 3, "to_id": 4, "type": "SS", "lag_days": 0},
    ]
    r = assess_all(tasks=tasks, links=links, assignments=[],
                   baseline=None, status_date="2026-12-31")
    # Should have many failures
    assert r["summary"]["fail_count"] > 5
    assert "..." in r["summary"]["executive_text"]
