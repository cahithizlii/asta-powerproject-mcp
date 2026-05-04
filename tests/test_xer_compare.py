"""Phase 7 T123-T125 — xer_compare.py pure module unit tests.

30 unit tests covering all 5 diff functions:
- diff_tasks (8)
- diff_links (5)
- diff_progress (6)
- diff_evm (5)
- summarize_compare (6)

Pure module — zero I/O, no fixtures.
"""
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from xer_compare import (
    diff_tasks,
    diff_links,
    diff_progress,
    diff_evm,
    summarize_compare,
    DEFAULT_TASK_FIELDS,
)


# ============================================================
# diff_tasks (8)
# ============================================================

def test_diff_tasks_empty_inputs_all_zero():
    r = diff_tasks([], [])
    assert r["added"] == []
    assert r["removed"] == []
    assert r["changed"] == []
    assert r["unchanged_count"] == 0


def test_diff_tasks_all_added():
    r = diff_tasks([], [{"id": 1, "name": "T1"}, {"id": 2, "name": "T2"}])
    assert len(r["added"]) == 2
    assert r["removed"] == []


def test_diff_tasks_all_removed():
    r = diff_tasks([{"id": 1, "name": "T1"}, {"id": 2, "name": "T2"}], [])
    assert r["added"] == []
    assert len(r["removed"]) == 2


def test_diff_tasks_identity_unchanged():
    """Same id + same fields -> unchanged_count increments."""
    a = [{"id": 1, "name": "T1", "percent_complete": 50.0,
          "baseline_start": "2026-01-01"}]
    b = [{"id": 1, "name": "T1", "percent_complete": 50.0,
          "baseline_start": "2026-01-01"}]
    r = diff_tasks(a, b)
    assert r["changed"] == []
    assert r["unchanged_count"] == 1


def test_diff_tasks_field_change_detected():
    a = [{"id": 1, "percent_complete": 25.0}]
    b = [{"id": 1, "percent_complete": 75.0}]
    r = diff_tasks(a, b)
    assert len(r["changed"]) == 1
    fc = r["changed"][0]["fields_changed"]
    assert fc["percent_complete"] == (25.0, 75.0)


def test_diff_tasks_custom_fields_only():
    """fields param restricts which fields trigger 'changed'."""
    a = [{"id": 1, "name": "Old", "percent_complete": 25.0,
          "baseline_start": "2026-01-01"}]
    b = [{"id": 1, "name": "New", "percent_complete": 25.0,
          "baseline_start": "2026-02-01"}]
    # Custom: only baseline_start tracked. name change ignored.
    r = diff_tasks(a, b, fields=["baseline_start"])
    assert len(r["changed"]) == 1
    assert "baseline_start" in r["changed"][0]["fields_changed"]
    assert "name" not in r["changed"][0]["fields_changed"]


def test_diff_tasks_default_excludes_name_field():
    """Default fields don't include 'name' — name change alone is unchanged."""
    a = [{"id": 1, "name": "Old", "percent_complete": 50.0}]
    b = [{"id": 1, "name": "Renamed", "percent_complete": 50.0}]
    r = diff_tasks(a, b)
    assert r["changed"] == []
    assert r["unchanged_count"] == 1


def test_diff_tasks_multi_field_change():
    a = [{"id": 1, "percent_complete": 25.0,
          "baseline_start": "2026-01-01", "actual_work": 100.0}]
    b = [{"id": 1, "percent_complete": 75.0,
          "baseline_start": "2026-01-15", "actual_work": 250.0}]
    r = diff_tasks(a, b)
    fc = r["changed"][0]["fields_changed"]
    assert len(fc) == 3
    assert "percent_complete" in fc
    assert "baseline_start" in fc
    assert "actual_work" in fc


# ============================================================
# diff_links (5)
# ============================================================

def test_diff_links_all_added():
    r = diff_links([], [
        {"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0},
    ])
    assert len(r["added"]) == 1
    assert r["removed"] == []


def test_diff_links_all_removed():
    r = diff_links([
        {"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0},
    ], [])
    assert r["added"] == []
    assert len(r["removed"]) == 1


def test_diff_links_lag_change_is_changed_not_replace():
    """Same identity (from,to,type), different lag -> changed."""
    a = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0}]
    b = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 5}]
    r = diff_links(a, b)
    assert r["added"] == []
    assert r["removed"] == []
    assert len(r["changed"]) == 1
    assert r["changed"][0]["lag_a"] == 0
    assert r["changed"][0]["lag_b"] == 5


def test_diff_links_type_change_replaces():
    """Different type means different identity -> removed + added."""
    a = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0}]
    b = [{"from_id": 1, "to_id": 2, "type": "SS", "lag_days": 0}]
    r = diff_links(a, b)
    assert len(r["added"]) == 1
    assert len(r["removed"]) == 1
    assert r["changed"] == []


def test_diff_links_unchanged_count():
    a = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0},
         {"from_id": 2, "to_id": 3, "type": "FS", "lag_days": 0}]
    b = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0},
         {"from_id": 2, "to_id": 3, "type": "FS", "lag_days": 0}]
    r = diff_links(a, b)
    assert r["unchanged_count"] == 2


# ============================================================
# diff_progress (6)
# ============================================================

def test_diff_progress_no_movement_returns_empty_tasks():
    a = {"status_date": "2026-01-01",
         "tasks": [{"id": 1, "percent_complete": 50.0, "actual_work": 100.0}]}
    b = {"status_date": "2026-02-01",
         "tasks": [{"id": 1, "percent_complete": 50.0, "actual_work": 100.0}]}
    r = diff_progress(a, b)
    assert r["tasks"] == []
    assert r["summary"]["count_moved"] == 0


def test_diff_progress_single_move():
    a = {"tasks": [{"id": 1, "percent_complete": 25.0, "actual_work": 50.0}]}
    b = {"tasks": [{"id": 1, "percent_complete": 75.0, "actual_work": 200.0}]}
    r = diff_progress(a, b)
    assert len(r["tasks"]) == 1
    assert r["tasks"][0]["pct_delta"] == 50.0
    assert r["tasks"][0]["aw_delta"] == 150.0


def test_diff_progress_total_delta_math():
    a = {"tasks": [{"id": 1, "percent_complete": 0, "actual_work": 0},
                   {"id": 2, "percent_complete": 0, "actual_work": 0}]}
    b = {"tasks": [{"id": 1, "percent_complete": 50, "actual_work": 100},
                   {"id": 2, "percent_complete": 30, "actual_work": 60}]}
    r = diff_progress(a, b)
    assert r["summary"]["total_pct_delta"] == 80.0
    assert r["summary"]["total_aw_delta"] == 160.0
    assert r["summary"]["count_moved"] == 2


def test_diff_progress_status_date_passthrough():
    a = {"status_date": "2026-01-15", "tasks": []}
    b = {"status_date": "2026-02-15", "tasks": []}
    r = diff_progress(a, b)
    assert r["status_date_a"] == "2026-01-15"
    assert r["status_date_b"] == "2026-02-15"


def test_diff_progress_task_only_in_b_treated_as_moved_from_zero():
    """Task absent in A -> implicit (0, 0). Appears in moved if B has values."""
    a = {"tasks": []}
    b = {"tasks": [{"id": 1, "percent_complete": 100.0, "actual_work": 80.0}]}
    r = diff_progress(a, b)
    assert len(r["tasks"]) == 1
    assert r["tasks"][0]["pct_delta"] == 100.0


def test_diff_progress_none_inputs_safe():
    r = diff_progress(None, None)
    assert r["tasks"] == []
    assert r["summary"]["count_moved"] == 0


# ============================================================
# diff_evm (5)
# ============================================================

def test_diff_evm_identity_zero_deltas():
    snap = {"bac": 100, "pv": 50, "ev": 40, "ac": 45, "spi": 0.8, "cpi": 0.89}
    r = diff_evm(snap, snap)
    assert r["bac_delta"] == 0.0
    assert r["spi_delta"] == 0.0


def test_diff_evm_partial_deltas():
    a = {"bac": 100, "ev": 40, "ac": 45, "spi": 0.8, "cpi": 0.89}
    b = {"bac": 100, "ev": 60, "ac": 50, "spi": 0.92, "cpi": 1.20}
    r = diff_evm(a, b)
    assert r["ev_delta"] == 20.0
    assert r["ac_delta"] == 5.0
    assert round(r["spi_delta"], 2) == 0.12


def test_diff_evm_none_field_returns_none_delta():
    a = {"bac": 100}
    b = {"bac": 120}
    r = diff_evm(a, b)
    assert r["bac_delta"] == 20.0
    assert r["spi_delta"] is None  # missing in both


def test_diff_evm_one_side_none_returns_none_delta():
    a = {"spi": 0.8}
    b = {}
    r = diff_evm(a, b)
    assert r["spi_delta"] is None


def test_diff_evm_empty_inputs_safe():
    r = diff_evm({}, {})
    for key in ("bac_delta", "pv_delta", "ev_delta", "ac_delta",
                "spi_delta", "cpi_delta"):
        assert r[key] is None


# ============================================================
# summarize_compare (6)
# ============================================================

def test_summarize_no_changes_headline():
    task_d = {"added": [], "removed": [], "changed": [], "unchanged_count": 5}
    link_d = {"added": [], "removed": [], "changed": [], "unchanged_count": 3}
    progress_d = {"tasks": [], "summary": {"count_moved": 0,
                  "total_pct_delta": 0, "total_aw_delta": 0}}
    evm_d = diff_evm({}, {})
    r = summarize_compare(task_d, link_d, progress_d, evm_d)
    assert r["headline"] == "no changes detected"


def test_summarize_counts_aggregation():
    task_d = {"added": [{"id": 5}], "removed": [{"id": 3}],
              "changed": [{"id": 1}], "unchanged_count": 4}
    link_d = {"added": [{}, {}], "removed": [], "changed": [{}],
              "unchanged_count": 5}
    progress_d = {"tasks": [{"id": 1}], "summary": {"count_moved": 1,
                  "total_pct_delta": 50, "total_aw_delta": 100}}
    evm_d = {"spi_a": 0.8, "spi_b": 0.85, "cpi_a": 0.9,
             "cpi_b": 0.95, "spi_delta": 0.05, "cpi_delta": 0.05,
             "ev_delta": 100}
    r = summarize_compare(task_d, link_d, progress_d, evm_d)
    c = r["counts"]
    assert c["tasks_added"] == 1
    assert c["tasks_removed"] == 1
    assert c["tasks_changed"] == 1
    assert c["links_added"] == 2
    assert c["links_changed"] == 1
    assert c["tasks_progressed"] == 1


def test_summarize_includes_spi_in_headline_when_both_present():
    evm_d = {"spi_a": 0.74, "spi_b": 0.81, "cpi_a": 1.0,
             "cpi_b": 1.05, "spi_delta": 0.07, "cpi_delta": 0.05,
             "ev_delta": 50}
    r = summarize_compare(
        {"added": [], "removed": [], "changed": [], "unchanged_count": 0},
        {"added": [], "removed": [], "changed": [], "unchanged_count": 0},
        {"tasks": [], "summary": {"count_moved": 0,
                                  "total_pct_delta": 0, "total_aw_delta": 0}},
        evm_d)
    assert "0.74" in r["headline"]
    assert "0.81" in r["headline"]


def test_summarize_omits_spi_when_missing():
    evm_d = diff_evm({}, {})
    r = summarize_compare(
        {"added": [{"id": 1}], "removed": [], "changed": [],
         "unchanged_count": 0},
        {"added": [], "removed": [], "changed": [], "unchanged_count": 0},
        {"tasks": [], "summary": {"count_moved": 0,
                                  "total_pct_delta": 0, "total_aw_delta": 0}},
        evm_d)
    assert "SPI" not in r["headline"]
    assert "1 tasks added" in r["headline"]


def test_summarize_propagates_evm_deltas():
    evm_d = {"spi_a": 0.7, "spi_b": 0.9, "cpi_a": 0.8, "cpi_b": 1.0,
             "spi_delta": 0.2, "cpi_delta": 0.2, "ev_delta": 500}
    r = summarize_compare(
        {"added": [], "removed": [], "changed": [], "unchanged_count": 0},
        {"added": [], "removed": [], "changed": [], "unchanged_count": 0},
        {"tasks": [], "summary": {"count_moved": 0,
                                  "total_pct_delta": 0, "total_aw_delta": 0}},
        evm_d)
    assert r["spi_delta"] == 0.2
    assert r["cpi_delta"] == 0.2
    assert r["ev_delta"] == 500


def test_summarize_only_progressed_in_headline():
    task_d = {"added": [], "removed": [], "changed": [], "unchanged_count": 0}
    link_d = {"added": [], "removed": [], "changed": [], "unchanged_count": 0}
    progress_d = {"tasks": [{"id": 1}, {"id": 2}, {"id": 3}],
                  "summary": {"count_moved": 3,
                              "total_pct_delta": 75, "total_aw_delta": 250}}
    evm_d = diff_evm({}, {})
    r = summarize_compare(task_d, link_d, progress_d, evm_d)
    assert "3 progressed" in r["headline"]


# ============================================================
# DEFAULT_TASK_FIELDS sanity
# ============================================================

def test_default_task_fields_include_progress_and_baseline():
    assert "percent_complete" in DEFAULT_TASK_FIELDS
    assert "baseline_start" in DEFAULT_TASK_FIELDS
    assert "actual_work" in DEFAULT_TASK_FIELDS


# === Phase 11.1 T141: gap-fill ===

def test_diff_tasks_skips_items_with_none_id():
    """_index_by skips items whose key value is None (line 42 continue)."""
    tasks_a = [
        {"id": None, "name": "ghost"},
        {"id": 1, "name": "T1", "percent_complete": 50},
    ]
    tasks_b = [
        {"id": 1, "name": "T1", "percent_complete": 75},
    ]
    r = diff_tasks(tasks_a, tasks_b)
    # Ghost task ignored; T1 marked changed
    assert len(r["changed"]) == 1
    assert r["changed"][0]["id"] == 1


def test_diff_evm_safe_sub_handles_non_numeric_strings():
    """_safe_sub returns None when values can't be converted (lines 189-190)."""
    snap_a = {"bac": "not-a-number", "pv": 100, "ev": 100, "ac": 50,
              "spi": 1.0, "cpi": 1.0}
    snap_b = {"bac": "still-not-a-number", "pv": 200, "ev": 150, "ac": 100,
              "spi": 0.75, "cpi": 0.6}
    r = diff_evm(snap_a, snap_b)
    # bac_delta forced through ValueError path -> None
    assert r["bac_delta"] is None
    # Numeric fields still compute deltas
    assert r["pv_delta"] == 100.0


def test_diff_evm_safe_sub_handles_unsupported_type():
    """_safe_sub returns None when float() raises TypeError on weird types."""
    class Weird:
        pass
    snap_a = {"ev": Weird()}
    snap_b = {"ev": 100}
    r = diff_evm(snap_a, snap_b)
    assert r["ev_delta"] is None
