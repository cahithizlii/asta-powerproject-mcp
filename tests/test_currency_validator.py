"""Phase 6.1 T115a — currency_validator.py pure module tests.

22 unit tests covering:
- detect_mode_from_xer_assignments (6)
- detect_mode_from_tasks_resources (5)
- extract_currency_code (4)
- cross_validate_modes (7)

Pure module — no fixtures, no I/O, no COM.
"""
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from currency_validator import (
    detect_mode_from_xer_assignments,
    detect_mode_from_tasks_resources,
    extract_currency_code,
    cross_validate_modes,
)


# === detect_mode_from_xer_assignments (RULE 3) ===

def test_xer_assignments_all_target_cost_equals_qty_returns_hours():
    """RULE 3: target_cost == target_qty in ALL rows -> 'hours' (not cost loaded)."""
    assignments = [
        {"task_id": 1, "target_qty": 100.0, "target_cost": 100.0},
        {"task_id": 2, "target_qty": 50.0, "target_cost": 50.0},
        {"task_id": 3, "target_qty": 200.0, "target_cost": 200.0},
    ]
    assert detect_mode_from_xer_assignments(assignments) == "hours"


def test_xer_assignments_real_cost_returns_cost():
    """target_cost differs significantly from target_qty -> real cost loading."""
    assignments = [
        {"task_id": 1, "target_qty": 100.0, "target_cost": 5000.0},
        {"task_id": 2, "target_qty": 50.0, "target_cost": 2500.0},
    ]
    assert detect_mode_from_xer_assignments(assignments) == "cost"


def test_xer_assignments_mixed_returns_mixed():
    """Some rows cost-loaded, some not -> mixed."""
    assignments = [
        {"task_id": 1, "target_qty": 100.0, "target_cost": 100.0},   # hours
        {"task_id": 2, "target_qty": 50.0, "target_cost": 5000.0},   # cost
        {"task_id": 3, "target_qty": 200.0, "target_cost": 200.0},   # hours
    ]
    assert detect_mode_from_xer_assignments(assignments) == "mixed"


def test_xer_assignments_empty_returns_uncertain():
    assert detect_mode_from_xer_assignments([]) == "uncertain"
    assert detect_mode_from_xer_assignments(None) == "uncertain"


def test_xer_assignments_zero_qty_zero_cost_returns_uncertain():
    """No cost data anywhere -> uncertain (cannot validate)."""
    assignments = [
        {"task_id": 1, "target_qty": 0.0, "target_cost": 0.0},
        {"task_id": 2, "target_qty": 0.0, "target_cost": 0.0},
    ]
    assert detect_mode_from_xer_assignments(assignments) == "uncertain"


def test_xer_assignments_zero_qty_with_cost_returns_cost():
    """target_qty == 0 but target_cost > 0 -> definitely cost loaded."""
    assignments = [
        {"task_id": 1, "target_qty": 0.0, "target_cost": 1000.0},
    ]
    assert detect_mode_from_xer_assignments(assignments) == "cost"


# === detect_mode_from_tasks_resources ===

def test_tasks_resources_no_cost_returns_hours():
    tasks = [{"cost": 0}, {"cost": 0}]
    resources = [{"cost": 0}]
    assert detect_mode_from_tasks_resources(tasks, resources) == "hours"


def test_tasks_resources_all_cost_returns_cost():
    tasks = [{"cost": 1000}, {"cost": 2000}]
    resources = [{"cost": 500}]
    assert detect_mode_from_tasks_resources(tasks, resources) == "cost"


def test_tasks_resources_partial_cost_returns_mixed():
    """Some entries with cost, some zero -> mixed."""
    tasks = [{"cost": 1000}, {"cost": 0}, {"cost": 500}]
    assert detect_mode_from_tasks_resources(tasks, []) == "mixed"


def test_tasks_resources_empty_returns_uncertain():
    assert detect_mode_from_tasks_resources([], []) == "uncertain"
    assert detect_mode_from_tasks_resources(None, None) == "uncertain"


def test_tasks_resources_invalid_cost_treated_as_missing():
    """Non-numeric cost values skipped; if all skipped -> uncertain."""
    tasks = [{"cost": "abc"}, {"cost": None}]
    assert detect_mode_from_tasks_resources(tasks, []) == "uncertain"


# === extract_currency_code ===

def test_extract_currency_code_from_xer_header():
    h = {"version": "21.12", "currency": "USD"}
    assert extract_currency_code(h) == "USD"


def test_extract_currency_code_missing_returns_none():
    assert extract_currency_code({}) is None


def test_extract_currency_code_empty_string_returns_none():
    assert extract_currency_code({"currency": ""}) is None


def test_extract_currency_code_none_input_returns_none():
    assert extract_currency_code(None) is None


# === cross_validate_modes ===

def test_cross_validate_all_agree_high_confidence():
    sources = [
        ("xer_assignments", "hours"),
        ("xer_resources", "hours"),
        ("tasks", "hours"),
    ]
    r = cross_validate_modes(sources)
    assert r["consensus_mode"] == "hours"
    assert r["confidence"] == "high"
    assert r["conflicts"] == []


def test_cross_validate_majority_with_one_dissent_medium_confidence():
    sources = [
        ("xer_assignments", "cost"),
        ("xer_resources", "cost"),
        ("tasks", "hours"),
    ]
    r = cross_validate_modes(sources)
    assert r["consensus_mode"] == "cost"
    assert r["confidence"] == "medium"
    assert len(r["conflicts"]) >= 1


def test_cross_validate_split_low_confidence():
    """Equal split -> low confidence, consensus is first-in-list mode."""
    sources = [
        ("xer_assignments", "cost"),
        ("xer_resources", "hours"),
    ]
    r = cross_validate_modes(sources)
    assert r["confidence"] == "low"
    assert r["consensus_mode"] in ("cost", "hours")


def test_cross_validate_uncertain_filtered_out_high_confidence():
    """Uncertain sources don't count in consensus."""
    sources = [
        ("xer_assignments", "hours"),
        ("xer_resources", "uncertain"),
    ]
    r = cross_validate_modes(sources)
    assert r["consensus_mode"] == "hours"
    assert r["confidence"] == "high"


def test_cross_validate_all_uncertain_returns_uncertain_low():
    sources = [
        ("xer_assignments", "uncertain"),
        ("xer_resources", "uncertain"),
    ]
    r = cross_validate_modes(sources)
    assert r["consensus_mode"] == "uncertain"
    assert r["confidence"] == "low"


def test_cross_validate_empty_returns_uncertain_low():
    r = cross_validate_modes([])
    assert r["consensus_mode"] == "uncertain"
    assert r["confidence"] == "low"


def test_cross_validate_warnings_include_mixed_source():
    """If any source reports 'mixed', a warning is generated."""
    sources = [
        ("xer_assignments", "mixed"),
        ("xer_resources", "cost"),
    ]
    r = cross_validate_modes(sources)
    assert any("mixed" in w.lower() for w in r["warnings"])


def test_cross_validate_source_counts_present():
    sources = [
        ("a", "cost"),
        ("b", "cost"),
        ("c", "hours"),
        ("d", "uncertain"),
    ]
    r = cross_validate_modes(sources)
    assert r["source_counts"]["cost"] == 2
    assert r["source_counts"]["hours"] == 1
    assert r["source_counts"]["uncertain"] == 1


# === Phase 11.1 T141: gap-fill for non-numeric / missing fields ===

def test_xer_assignments_skips_non_numeric_qty_or_cost():
    """Rows with target_qty=None or target_cost='abc' are skipped (line 62 continue)."""
    assignments = [
        {"task_id": 1, "target_qty": None, "target_cost": 100.0},      # qty None -> skip
        {"task_id": 2, "target_qty": 100.0, "target_cost": None},      # cost None -> skip
        {"task_id": 3, "target_qty": "not-a-number", "target_cost": 0},  # qty non-numeric -> skip
        {"task_id": 4, "target_qty": 50.0, "target_cost": "bad"},      # cost non-numeric -> skip
    ]
    # All rows skipped -> uncertain
    assert detect_mode_from_xer_assignments(assignments) == "uncertain"


def test_xer_assignments_zero_qty_zero_cost_skipped():
    """target_qty=0 and target_cost=0 -> no signal (skipped from else branch)."""
    assignments = [
        {"task_id": 1, "target_qty": 0.0, "target_cost": 0.0},
        {"task_id": 2, "target_qty": 0.0, "target_cost": 0.0},
    ]
    assert detect_mode_from_xer_assignments(assignments) == "uncertain"


def test_xer_assignments_zero_qty_positive_cost_signals_cost():
    """target_qty=0 with target_cost>0 -> cost signal."""
    assignments = [
        {"task_id": 1, "target_qty": 0.0, "target_cost": 500.0},
    ]
    assert detect_mode_from_xer_assignments(assignments) == "cost"


def test_tasks_resources_skips_non_numeric_resource_cost():
    """Resource entries with cost=None are skipped (line 107 continue)."""
    tasks = [{"id": 1, "cost": 1000.0}]
    resources = [
        {"id": "R1", "cost": None},          # None -> skip
        {"id": "R2", "cost": "not-numeric"},  # bad str -> skip
        {"id": "R3", "cost": 500.0},
    ]
    # tasks: 1 cost_signal; resources: R3 -> cost_signal -> 'cost'
    assert detect_mode_from_tasks_resources(tasks, resources) == "cost"


def test_tasks_resources_skips_non_numeric_task_cost():
    """Task entries with cost=None or non-numeric are skipped (line 99 continue)."""
    tasks = [
        {"id": 1, "cost": None},
        {"id": 2, "cost": "abc"},
    ]
    # No signals at all -> uncertain
    assert detect_mode_from_tasks_resources(tasks, None) == "uncertain"


def test_extract_currency_code_strips_whitespace():
    """Whitespace-only currency value normalized to None after strip."""
    assert extract_currency_code({"currency": "   "}) is None


def test_extract_currency_code_strips_real_value():
    """Currency value with surrounding whitespace gets stripped to clean code."""
    assert extract_currency_code({"currency": "  USD  "}) == "USD"
