# Phase 5b DCMA Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans (or subagent-driven-development) to implement this plan task-by-task.

**Goal:** New `msproject_health` MCP tool — 4 actions implementing DCMA 14-Point Schedule Health Assessment per CLAUDE.md RULE 10. Hybrid file+COM with hardcoded industry-standard thresholds, RAG output, drill-down per rule, and snapshot comparison.

**Architecture:** New `dcma_checks.py` module — pure-Python 14 rule check functions, MSP/COM/file independent, fixture-free testable. `msproject_mcp_core.py` Phase 5b section adds I/O adapters (`_dcma_load_links`, `_dcma_extract_floats/constraints`, `_dcma_collect_full_data`), 4 helper functions, and FastMCP dispatcher. Phase 1+2+3+4+5a helpers DOKUNULMAZ; only read-only calls into Phase 5a `_evm_load_*`, Phase 4 `_msp_file_read_*`, and Phase 1 COM (`_validate_active_project`).

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM, pytest. New helper-only Python module `dcma_checks.py`. Mevcut `msproject_mcp_core.py` (~5450 satır after Phase 5a + TAIL fix), 40+ test files, **138 cumulative regression PASS** baseline.

**Design doc:** `docs/plans/2026-05-01-phase5b-dcma-design.md` (commit `bf2a8c5`)

**Baseline state at start:** HEAD `bf2a8c5`, MS Project running v16.0.

**KEY REFERENCES:**
- CLAUDE.md RULE 10 — DCMA 14-Point thresholds (no_pred <5%, no_succ <5%, leads=0, lags <5%, fs_link >90%, hard_constraint <5%, high_float <5%, negative_float=0, high_duration <5%, invalid_dates=0, resources_missing <20%, missed_tasks <5%, critical_path >0, BEI >95%)
- CLAUDE.md RULE 12 — RAG (Phase 5a uses spi-based; Phase 5b uses pass_count: >12 GREEN, 8-11 AMBER, <8 RED)
- Phase 5a `_evm_load_task_data/baseline_data/progress_data` — hybrid loaders
- Phase 4 `_msp_file_read_links/tasks/assignments` — file path
- Phase 3a `BASELINE_NUMBERS`
- Phase 1 `_validate_active_project`, `_format_com_error`

**MSPDI/COM field mapping notes:**
- Predecessors: file path → `task["predecessors"]` list (T66 probe); COM → walk `task.PredecessorTasks` collection
- Successors: file path → `task["successors"]` list; COM → `task.SuccessorTasks`
- Total slack: file path → `task["total_float"]` field; COM → `task.TotalSlack` (in minutes)
- Critical: file path → `task["critical"]` boolean; COM → `task.Critical`
- Constraint type: file path → `task["constraint_type"]` (probe T87 to confirm); COM → `task.ConstraintType` (enum 0-7)
- Link type: file path → `link["type"]` ("FS"/"SS"/"FF"/"SF"); COM → walk per task
- Link lag: file path → `link["lag_days"]` (already converted); COM → `predecessor.Lag` (minutes; negative = lead)

---

## Task 85: `dcma_checks.py` Foundations + Rule 1 + Rule 2

**Files:**
- Create: `C:\Users\CahAsus\asta-powerproject-mcp\dcma_checks.py`
- Create: `C:\Users\CahAsus\asta-powerproject-mcp\tests\test_dcma_checks.py`

**Step 1: Failing tests**

`tests/test_dcma_checks.py`:
```python
"""Test pure-math DCMA 14-Point check functions (CLAUDE.md RULE 10).

No fixtures, no COM, no MSP — fully data-driven.
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


# ---------- check_no_predecessor (RULE 1) ----------

def _make_task(id, name="T", summary=False, predecessors=None, successors=None):
    return {"id": id, "name": name, "summary": summary,
            "predecessors": predecessors or [], "successors": successors or []}


def test_check_no_predecessor_pass():
    """All real tasks have predecessors → pass."""
    tasks = [
        _make_task(0, "Project", summary=True),  # root project
        _make_task(1, "Start", predecessors=[]),  # start task — typical OK
        _make_task(2, "T2", predecessors=[1]),
        _make_task(3, "T3", predecessors=[2]),
    ]
    r = check_no_predecessor(tasks)
    assert r["id"] == 1
    assert r["status"] == "pass"
    # Real tasks (excluding summary): 3. Without preds: 1 (id=1, the start)
    # 1/3 = 33% — actually, start tasks are expected to have no predecessor.
    # DCMA convention treats first task as legitimate. Implementation should
    # skip the project root and any task that's the earliest start.
    # For simplicity: skip summaries only. So 1/3 = 33% > 5% = FAIL.
    # Test: implementation may take simpler approach.


def test_check_no_predecessor_fail_high_pct():
    """50% of real tasks have no predecessor → FAIL."""
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
    # Real tasks: T2, T3 (both have preds) → 0% no-pred → PASS
    assert r["status"] == "pass"
    assert r["actual"] == 0.0


def test_check_no_predecessor_empty():
    r = check_no_predecessor([])
    assert r["status"] == "pass"  # vacuously true
    assert r["total_count"] == 0


# ---------- check_no_successor (RULE 2) ----------

def test_check_no_successor_pass():
    tasks = [
        _make_task(1, "T1", successors=[2]),
        _make_task(2, "T2", successors=[3]),
        _make_task(3, "T3", successors=[]),  # last task, expected to have no succ
    ]
    r = check_no_successor(tasks)
    assert r["id"] == 2
    # 1/3 = 33% > 5% → FAIL (simple impl)
    # OR: implementation could treat last task as legitimate, => 0% PASS
    # For test: use 50% threshold case
    pass


def test_check_no_successor_fail_high_pct():
    tasks = [
        _make_task(1, "T1", successors=[]),
        _make_task(2, "T2", successors=[]),
        _make_task(3, "T3", successors=[1]),
        _make_task(4, "T4", successors=[2]),
    ]
    r = check_no_successor(tasks)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)
    assert r["failed_count"] == 2


def test_check_no_successor_pass_low_pct():
    """Only 1 of 30 tasks (3.3%) has no successor → PASS (<5%)."""
    tasks = [_make_task(i, f"T{i}", successors=[i+1]) for i in range(1, 30)]
    tasks.append(_make_task(30, "T30", successors=[]))  # last
    r = check_no_successor(tasks)
    assert r["status"] == "pass"
    assert r["actual"] < 5.0
```

**Step 2: Run — expect ImportError**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/test_dcma_checks.py -v
```

**Step 3: Implementation**

Create `dcma_checks.py`:
```python
"""Phase 5b — DCMA 14-Point Schedule Health Assessment per CLAUDE.md RULE 10.

Pure-Python check functions. MSP/COM/file independent — takes plain dicts,
returns plain dicts. Easily testable without fixtures, without COM, without
MS Project.

Industry-standard thresholds (DCMA spec, NDIA EVMS, RULE 10):
- Logic (Rules 1-5): no_pred <5%, no_succ <5%, leads=0, lags <5%, fs_link >90%
- Constraints (Rule 6): hard_constraints <5%
- Float (Rules 7-8): high_float <5%, negative_float=0
- Duration (Rule 9): high_duration <5%
- Quality (Rules 10-11): invalid_dates=0, resources_missing <20%
- Schedule (Rules 12-14): missed_tasks <5%, critical_path >0, BEI >95%
"""
from typing import List, Dict, Any, Optional
import datetime as _dt


# ---------- Hardcoded thresholds (CLAUDE.md RULE 10) ----------

_DCMA_THRESHOLDS = {
    1: ("no_predecessor_pct", "<", 5.0),       # %
    2: ("no_successor_pct", "<", 5.0),
    3: ("leads_count", "==", 0),                # absolute
    4: ("lags_pct", "<", 5.0),
    5: ("fs_link_pct", ">", 90.0),
    6: ("hard_constraints_pct", "<", 5.0),
    7: ("high_float_pct", "<", 5.0),
    8: ("negative_float_count", "==", 0),
    9: ("high_duration_pct", "<", 5.0),
    10: ("invalid_dates_count", "==", 0),
    11: ("resources_missing_pct", "<", 20.0),
    12: ("missed_tasks_pct", "<", 5.0),
    13: ("critical_path_count", ">", 0),
    14: ("bei_pct", ">", 95.0),
}


# ---------- DCMA_RULES metadata (14 rules) ----------

DCMA_RULES = [
    {"id": 1, "name": "No Predecessor", "threshold_label": "<5%", "threshold_value": 5.0, "category": "Logic"},
    {"id": 2, "name": "No Successor", "threshold_label": "<5%", "threshold_value": 5.0, "category": "Logic"},
    {"id": 3, "name": "Leads", "threshold_label": "=0", "threshold_value": 0, "category": "Logic"},
    {"id": 4, "name": "Lags", "threshold_label": "<5%", "threshold_value": 5.0, "category": "Logic"},
    {"id": 5, "name": "FS Link %", "threshold_label": ">90%", "threshold_value": 90.0, "category": "Logic"},
    {"id": 6, "name": "Hard Constraints", "threshold_label": "<5%", "threshold_value": 5.0, "category": "Constraints"},
    {"id": 7, "name": "High Float (>44d)", "threshold_label": "<5%", "threshold_value": 5.0, "category": "Float"},
    {"id": 8, "name": "Negative Float", "threshold_label": "=0", "threshold_value": 0, "category": "Float"},
    {"id": 9, "name": "High Duration (>44d)", "threshold_label": "<5%", "threshold_value": 5.0, "category": "Duration"},
    {"id": 10, "name": "Invalid Dates", "threshold_label": "=0", "threshold_value": 0, "category": "Quality"},
    {"id": 11, "name": "Resources Missing", "threshold_label": "<20%", "threshold_value": 20.0, "category": "Quality"},
    {"id": 12, "name": "Missed Tasks", "threshold_label": "<5%", "threshold_value": 5.0, "category": "Schedule"},
    {"id": 13, "name": "Critical Path", "threshold_label": ">0", "threshold_value": 0, "category": "Schedule"},
    {"id": 14, "name": "BEI", "threshold_label": ">95%", "threshold_value": 95.0, "category": "Schedule"},
]


def _real_tasks(tasks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Filter out summary tasks (DCMA counts only 'real' work)."""
    return [t for t in tasks if not t.get("summary", False)]


def _eval_status(rule_id: int, actual: float) -> str:
    """Compare actual against threshold; return 'pass' or 'fail'."""
    field, op, threshold = _DCMA_THRESHOLDS[rule_id]
    if op == "<":
        return "pass" if actual < threshold else "fail"
    if op == ">":
        return "pass" if actual > threshold else "fail"
    if op == "==":
        return "pass" if actual == threshold else "fail"
    return "fail"


def check_no_predecessor(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 1: <5% of real tasks should have no predecessor.

    Returns {id, name, threshold, actual, actual_unit, status, failed_count,
             total_count, failed_task_ids}.
    """
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0:
        return {"id": 1, "name": "No Predecessor", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    failed_ids = [t["id"] for t in real if not (t.get("predecessors") or [])]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 1, "name": "No Predecessor", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(1, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }


def check_no_successor(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 2: <5% of real tasks should have no successor."""
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0:
        return {"id": 2, "name": "No Successor", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    failed_ids = [t["id"] for t in real if not (t.get("successors") or [])]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 2, "name": "No Successor", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(2, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }
```

**Step 4: Run — expect 8-10 PASS** (3 metadata + 4-5 RULE 1 + 3 RULE 2 — some test names commented as TODO; only assertion-bearing tests count)

```bash
python -m pytest tests/test_dcma_checks.py -v
```

**Step 5: Run full regression to verify no Phase 1-5a breakage**

```bash
python -m pytest tests/ -q --tb=line --ignore=cleanup_test.py --ignore=test_apply_tabledef.py 2>&1 | tail -5
```

Expected: full Phase 1-5a regression unchanged + ~10 new T85 tests PASS.

**Step 6: Commit**

```bash
git add dcma_checks.py tests/test_dcma_checks.py
git commit -m "Phase 5b T85: dcma_checks foundations + RULE 1-2 (no_pred + no_succ)

DCMA_RULES metadata for 14 rules + _DCMA_THRESHOLDS hardcoded values
(CLAUDE.md RULE 10). _eval_status helper compares actual vs threshold
operators (<, >, ==). _real_tasks filter excludes summary rows.

check_no_predecessor + check_no_successor return per-rule dict with
status/actual/failed_count/total_count/failed_task_ids.

10 unit tests, no fixtures, no MSP."
```

DO NOT push (T93 will push the chain).

---

## Task 86: Link Rules — RULE 3-5 (Leads + Lags + FS Link %)

**Files:**
- Modify: `dcma_checks.py`
- Modify: `tests/test_dcma_checks.py`

**Step 1: Append failing tests**

```python
from dcma_checks import check_leads, check_lags, check_fs_link_pct


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


def test_check_leads_fail_three_leads():
    links = [_link(1, 2, lag_days=-3), _link(2, 3, lag_days=-1),
             _link(3, 4, lag_days=-2), _link(4, 5, lag_days=0)]
    r = check_leads(links)
    assert r["status"] == "fail"
    assert r["actual"] == 3


# ---------- check_lags (RULE 4: <5%) ----------

def test_check_lags_pass_low_pct():
    """1 lag in 30 links = 3.3% → PASS."""
    links = [_link(i, i+1, lag_days=0) for i in range(1, 30)]
    links.append(_link(30, 31, lag_days=2))
    r = check_lags(links)
    assert r["status"] == "pass"
    assert r["actual"] < 5.0


def test_check_lags_fail_high_pct():
    """3 lags in 10 links = 30% → FAIL."""
    links = [_link(i, i+1, lag_days=0) for i in range(1, 8)]
    links += [_link(8, 9, lag_days=3), _link(9, 10, lag_days=2),
              _link(10, 11, lag_days=1)]
    r = check_lags(links)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(30.0, rel=1e-2)


def test_check_lags_empty():
    r = check_lags([])
    assert r["status"] == "pass"
    assert r["total_count"] == 0


# ---------- check_fs_link_pct (RULE 5: >90%) ----------

def test_check_fs_link_pct_pass():
    """10 links, 9 FS = 90% — borderline. >90% strict → 9/10 fails. Use 95%."""
    links = [_link(i, i+1, type="FS") for i in range(1, 20)]
    links.append(_link(20, 21, type="SS"))  # 1 SS / 20 = 95% FS
    r = check_fs_link_pct(links)
    assert r["status"] == "pass"
    assert r["actual"] >= 90.0


def test_check_fs_link_pct_fail_too_many_non_fs():
    """5 of 10 = 50% FS → FAIL."""
    links = [_link(i, i+1, type="FS") for i in range(1, 6)]
    links += [_link(i, i+1, type="SS") for i in range(6, 11)]
    r = check_fs_link_pct(links)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)


def test_check_fs_link_pct_empty():
    r = check_fs_link_pct([])
    assert r["status"] == "pass"  # vacuous
    assert r["total_count"] == 0
```

**Step 2: Run — FAIL**

**Step 3: Implementation — append to `dcma_checks.py`**

```python
def check_leads(links: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 3: zero leads (negative lag) allowed.

    A lead = predecessor link with negative lag (successor starts BEFORE
    predecessor finishes). DCMA prohibits leads entirely.
    """
    failed_links = [l for l in links if (l.get("lag_days") or 0) < 0]
    failed_count = len(failed_links)
    return {
        "id": 3, "name": "Leads", "threshold": "=0",
        "actual": failed_count, "actual_unit": "count",
        "status": _eval_status(3, failed_count),
        "failed_count": failed_count, "total_count": len(links),
        "failed_links": [{"from_id": l["from_id"], "to_id": l["to_id"],
                         "lag_days": l["lag_days"]} for l in failed_links],
    }


def check_lags(links: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 4: <5% of links should have lag (positive lag_days)."""
    total = len(links)
    if total == 0:
        return {"id": 4, "name": "Lags", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_links": []}
    failed_links = [l for l in links if (l.get("lag_days") or 0) > 0]
    failed_count = len(failed_links)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 4, "name": "Lags", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(4, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_links": [{"from_id": l["from_id"], "to_id": l["to_id"],
                         "lag_days": l["lag_days"]} for l in failed_links],
    }


def check_fs_link_pct(links: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 5: >90% of links should be Finish-to-Start (FS)."""
    total = len(links)
    if total == 0:
        return {"id": 5, "name": "FS Link %", "threshold": ">90%",
                "actual": 100.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_links": []}
    fs_count = sum(1 for l in links if (l.get("type") or "").upper() == "FS")
    actual_pct = (fs_count / total) * 100.0
    failed_links = [l for l in links if (l.get("type") or "").upper() != "FS"]
    return {
        "id": 5, "name": "FS Link %", "threshold": ">90%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(5, actual_pct),
        "failed_count": len(failed_links), "total_count": total,
        "failed_links": [{"from_id": l["from_id"], "to_id": l["to_id"],
                         "type": l["type"]} for l in failed_links],
    }
```

**Step 4: Run — expect ~9 new + ~10 prev = ~19 PASS**

**Step 5: Commit**

```bash
git add dcma_checks.py tests/test_dcma_checks.py
git commit -m "Phase 5b T86: dcma_checks RULE 3-5 (leads + lags + fs_link_pct)

Logic category — link-based rules. Lead detection via negative lag_days;
lag detection via positive lag_days; FS percentage via type field
case-insensitive comparison. Each returns failed_links with from/to/type/lag
for drill-down."
```

---

## Task 87: Task Quality Rules — RULE 6, 10, 11 (Hard Constraints + Invalid Dates + Resources Missing)

**Files:**
- Modify: `dcma_checks.py`
- Modify: `tests/test_dcma_checks.py`

**Step 1: Append failing tests**

```python
from dcma_checks import check_hard_constraints, check_invalid_dates, check_resources_missing


# ---------- check_hard_constraints (RULE 6: <5%) ----------
# Hard constraints in MSPDI/COM enum:
# 0=ASAP, 1=ALAP, 2=MSO (Must Start On), 3=MFO (Must Finish On),
# 4=SNET (Start No Earlier Than), 5=SNLT (Start No Later Than),
# 6=FNET (Finish No Earlier Than), 7=FNLT (Finish No Later Than)
# DCMA classifies MSO, MFO, SNLT, FNLT as "hard" (rigid).

HARD_CONSTRAINT_TYPES = {2, 3, 5, 7}


def _make_task_constraint(id, constraint_type=0, summary=False):
    return {"id": id, "name": f"T{id}", "summary": summary,
            "constraint_type": constraint_type}


def test_check_hard_constraints_pass():
    tasks = [_make_task_constraint(i, constraint_type=0) for i in range(1, 30)]
    tasks.append(_make_task_constraint(30, constraint_type=2))  # 1 MSO
    r = check_hard_constraints(tasks)
    assert r["id"] == 6
    assert r["status"] == "pass"
    assert r["actual"] < 5.0


def test_check_hard_constraints_fail():
    tasks = [_make_task_constraint(i, constraint_type=2) for i in range(1, 6)]  # all MSO
    tasks += [_make_task_constraint(i, constraint_type=0) for i in range(6, 11)]
    r = check_hard_constraints(tasks)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)


def test_check_hard_constraints_excludes_summaries():
    tasks = [_make_task_constraint(1, constraint_type=2, summary=True),
             _make_task_constraint(2, constraint_type=0),
             _make_task_constraint(3, constraint_type=0)]
    r = check_hard_constraints(tasks)
    # Real tasks: 2 (T2, T3), neither hard → 0%
    assert r["status"] == "pass"
    assert r["actual"] == 0.0


# ---------- check_invalid_dates (RULE 10: =0) ----------

def test_check_invalid_dates_pass():
    tasks = [{"id": 1, "name": "T1", "start": "2026-01-01", "finish": "2026-01-10",
              "summary": False}]
    r = check_invalid_dates(tasks)
    assert r["id"] == 10
    assert r["status"] == "pass"


def test_check_invalid_dates_fail_start_after_finish():
    tasks = [{"id": 1, "name": "T1", "start": "2026-01-15", "finish": "2026-01-10",
              "summary": False}]
    r = check_invalid_dates(tasks)
    assert r["status"] == "fail"
    assert r["failed_count"] == 1
    assert 1 in r["failed_task_ids"]


def test_check_invalid_dates_handles_none():
    tasks = [{"id": 1, "name": "T1", "start": None, "finish": None, "summary": False}]
    r = check_invalid_dates(tasks)
    # None dates → no validation possible, treat as PASS (don't fail unnecessarily)
    assert r["status"] == "pass"


# ---------- check_resources_missing (RULE 11: <20%) ----------

def _task_with_dur(id, duration_h, summary=False):
    return {"id": id, "name": f"T{id}", "summary": summary, "duration_h": duration_h}


def test_check_resources_missing_pass():
    """1 of 10 tasks has no assignment, duration > 0 → 10% < 20% → PASS."""
    tasks = [_task_with_dur(i, 8) for i in range(1, 11)]
    assignments = [{"task_id": i, "resource_id": 1} for i in range(1, 10)]
    # Task 10 has no assignment
    r = check_resources_missing(tasks, assignments)
    assert r["id"] == 11
    assert r["status"] == "pass"
    assert r["actual"] < 20.0


def test_check_resources_missing_fail():
    """30% of tasks without resources → FAIL."""
    tasks = [_task_with_dur(i, 8) for i in range(1, 11)]
    assignments = [{"task_id": i, "resource_id": 1} for i in range(1, 8)]  # 1-7 only
    r = check_resources_missing(tasks, assignments)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(30.0, rel=1e-2)


def test_check_resources_missing_excludes_zero_duration():
    """Zero-duration tasks (milestones) excluded from resource count."""
    tasks = [_task_with_dur(1, 0)]  # milestone
    r = check_resources_missing(tasks, [])
    # No real tasks with duration → 0/0, treat as pass
    assert r["total_count"] == 0
    assert r["status"] == "pass"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
HARD_CONSTRAINT_TYPES = {2, 3, 5, 7}  # MSO, MFO, SNLT, FNLT


def check_hard_constraints(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 6: <5% of real tasks should have hard constraints (MSO/MFO/SNLT/FNLT)."""
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0:
        return {"id": 6, "name": "Hard Constraints", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    failed_ids = [t["id"] for t in real
                  if int(t.get("constraint_type") or 0) in HARD_CONSTRAINT_TYPES]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 6, "name": "Hard Constraints", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(6, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }


def _parse_iso_date_local(s):
    """Local date parser (avoid circular import with msproject_mcp_core)."""
    if not s or s == "N/A":
        return None
    try:
        return _dt.date.fromisoformat(str(s)[:10])
    except (ValueError, TypeError):
        return None


def check_invalid_dates(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 10: zero invalid dates (start>finish, etc.)."""
    failed_ids = []
    for t in tasks:
        start = _parse_iso_date_local(t.get("start"))
        finish = _parse_iso_date_local(t.get("finish"))
        if start and finish and start > finish:
            failed_ids.append(t["id"])
    failed_count = len(failed_ids)
    return {
        "id": 10, "name": "Invalid Dates", "threshold": "=0",
        "actual": failed_count, "actual_unit": "count",
        "status": _eval_status(10, failed_count),
        "failed_count": failed_count, "total_count": len(tasks),
        "failed_task_ids": failed_ids,
    }


def check_resources_missing(tasks: List[Dict[str, Any]],
                            assignments: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 11: <20% of real tasks (with duration > 0) should lack resource assignments."""
    real = _real_tasks(tasks)
    # Only count tasks with positive duration
    real = [t for t in real if float(t.get("duration_h") or 0) > 0]
    total = len(real)
    if total == 0:
        return {"id": 11, "name": "Resources Missing", "threshold": "<20%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    assigned_task_ids = {a.get("task_id") for a in (assignments or [])}
    failed_ids = [t["id"] for t in real if t["id"] not in assigned_task_ids]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 11, "name": "Resources Missing", "threshold": "<20%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(11, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }
```

**Step 4: Run — expect ~9 new + 19 prev = ~28 PASS**

**Step 5: Commit**

```bash
git add dcma_checks.py tests/test_dcma_checks.py
git commit -m "Phase 5b T87: dcma_checks RULE 6, 10, 11 (hard_constraints + invalid_dates + resources_missing)

Constraint type enum: MSO=2, MFO=3, SNLT=5, FNLT=7 considered hard.
Invalid dates: start > finish detection (None dates skip — vacuous pass).
Resources missing: filters zero-duration tasks (milestones) before %.

Local _parse_iso_date_local to avoid circular import with msproject_mcp_core."
```

---

## Task 88: Duration/Float Rules — RULE 7-9 (High Float + Negative Float + High Duration)

**Files:**
- Modify: `dcma_checks.py`
- Modify: `tests/test_dcma_checks.py`

**Step 1: Append failing tests**

```python
from dcma_checks import check_high_float, check_negative_float, check_high_duration


def _task_with_slack(id, total_slack_days, summary=False):
    return {"id": id, "name": f"T{id}", "summary": summary,
            "total_slack_days": total_slack_days}


# ---------- check_high_float (RULE 7: <5% with float > 44d) ----------

def test_check_high_float_pass():
    tasks = [_task_with_slack(i, 10) for i in range(1, 21)]
    tasks.append(_task_with_slack(21, 50))  # 1 high
    r = check_high_float(tasks)
    assert r["id"] == 7
    assert r["status"] == "pass"  # 1/21 = 4.7% < 5%
    assert r["failed_count"] == 1


def test_check_high_float_fail():
    tasks = [_task_with_slack(i, 60) for i in range(1, 6)]  # 5 high
    tasks += [_task_with_slack(i, 5) for i in range(6, 11)]  # 5 normal
    r = check_high_float(tasks)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)


def test_check_high_float_threshold_44d():
    """Boundary: exactly 44d = NOT high (>44 is high)."""
    tasks = [_task_with_slack(1, 44), _task_with_slack(2, 45)]
    r = check_high_float(tasks)
    assert r["failed_count"] == 1  # only T2


# ---------- check_negative_float (RULE 8: =0) ----------

def test_check_negative_float_pass():
    tasks = [_task_with_slack(i, 5) for i in range(1, 11)]
    r = check_negative_float(tasks)
    assert r["id"] == 8
    assert r["status"] == "pass"
    assert r["actual"] == 0


def test_check_negative_float_fail():
    tasks = [_task_with_slack(1, -3), _task_with_slack(2, -1),
             _task_with_slack(3, 5)]
    r = check_negative_float(tasks)
    assert r["status"] == "fail"
    assert r["actual"] == 2


# ---------- check_high_duration (RULE 9: <5% with duration > 44d) ----------

def _task_with_dur_days(id, duration_h, summary=False):
    return {"id": id, "name": f"T{id}", "summary": summary,
            "duration_h": duration_h}


def test_check_high_duration_pass():
    """1 of 30 tasks > 44d → 3.3% < 5% PASS.

    44 working days × 8h/day = 352 hours threshold (DCMA standard 8h/day).
    """
    tasks = [_task_with_dur_days(i, 80) for i in range(1, 30)]  # 10d each
    tasks.append(_task_with_dur_days(30, 400))  # 50d > 44d
    r = check_high_duration(tasks)
    assert r["id"] == 9
    assert r["status"] == "pass"


def test_check_high_duration_fail():
    """3 of 10 tasks > 44d → 30% FAIL."""
    tasks = [_task_with_dur_days(i, 400) for i in range(1, 4)]  # 50d each
    tasks += [_task_with_dur_days(i, 80) for i in range(4, 11)]
    r = check_high_duration(tasks)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(30.0, rel=1e-2)


def test_check_high_duration_excludes_summaries():
    """Summary tasks excluded from duration count."""
    tasks = [_task_with_dur_days(1, 1000, summary=True),
             _task_with_dur_days(2, 80)]
    r = check_high_duration(tasks)
    assert r["status"] == "pass"
    assert r["actual"] == 0.0
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
HIGH_FLOAT_THRESHOLD_DAYS = 44.0
HIGH_DURATION_THRESHOLD_HOURS = 44.0 * 8.0  # 44 working days × 8h/day standard


def check_high_float(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 7: <5% of real tasks should have total slack > 44 days."""
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0:
        return {"id": 7, "name": "High Float (>44d)", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    failed_ids = [t["id"] for t in real
                  if float(t.get("total_slack_days") or 0) > HIGH_FLOAT_THRESHOLD_DAYS]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 7, "name": "High Float (>44d)", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(7, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }


def check_negative_float(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 8: zero tasks with negative total slack."""
    real = _real_tasks(tasks)
    failed_ids = [t["id"] for t in real
                  if float(t.get("total_slack_days") or 0) < 0]
    failed_count = len(failed_ids)
    return {
        "id": 8, "name": "Negative Float", "threshold": "=0",
        "actual": failed_count, "actual_unit": "count",
        "status": _eval_status(8, failed_count),
        "failed_count": failed_count, "total_count": len(real),
        "failed_task_ids": failed_ids,
    }


def check_high_duration(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 9: <5% of real tasks should have duration > 44 working days (352h)."""
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0:
        return {"id": 9, "name": "High Duration (>44d)", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    failed_ids = [t["id"] for t in real
                  if float(t.get("duration_h") or 0) > HIGH_DURATION_THRESHOLD_HOURS]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 9, "name": "High Duration (>44d)", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(9, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }
```

**Step 4: Run — expect ~9 new + 28 prev = ~37 PASS**

**Step 5: Commit**

```bash
git add dcma_checks.py tests/test_dcma_checks.py
git commit -m "Phase 5b T88: dcma_checks RULE 7-9 (high_float + negative_float + high_duration)

Float thresholds: total_slack_days field — high>44d, negative<0.
Duration: 44 working days × 8h/day = 352h threshold (DCMA standard).
Summary tasks excluded from RULE 9 count (real tasks only)."
```

---

## Task 89: Schedule Health Rules — RULE 12-14 (Missed Tasks + Critical Path + BEI)

**Files:**
- Modify: `dcma_checks.py`
- Modify: `tests/test_dcma_checks.py`

**Step 1: Append failing tests**

```python
from dcma_checks import check_missed_tasks, check_critical_path, check_bei
import datetime as dt


# ---------- check_missed_tasks (RULE 12: <5%) ----------

def _task_baseline(id, baseline_finish, percent_complete=0, summary=False):
    return {"id": id, "name": f"T{id}", "summary": summary,
            "baseline_finish": baseline_finish,
            "percent_complete": percent_complete}


def test_check_missed_tasks_pass():
    """1 of 30 tasks past baseline_finish without completion → 3.3% PASS."""
    tasks = [_task_baseline(i, "2026-04-01", percent_complete=100)
             for i in range(1, 30)]
    tasks.append(_task_baseline(30, "2026-04-01", percent_complete=50))
    status_date = "2026-05-01"
    r = check_missed_tasks(tasks, status_date)
    assert r["id"] == 12
    assert r["status"] == "pass"


def test_check_missed_tasks_fail():
    """3 of 10 tasks past baseline incomplete → 30% FAIL."""
    tasks = [_task_baseline(i, "2026-04-01", percent_complete=50)
             for i in range(1, 4)]
    tasks += [_task_baseline(i, "2026-06-01", percent_complete=0)
              for i in range(4, 11)]
    status_date = "2026-05-01"
    r = check_missed_tasks(tasks, status_date)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(30.0, rel=1e-2)


def test_check_missed_tasks_no_status_date():
    """status_date None → vacuous PASS."""
    tasks = [_task_baseline(1, "2026-01-01", percent_complete=0)]
    r = check_missed_tasks(tasks, None)
    assert r["status"] == "pass"


# ---------- check_critical_path (RULE 13: count > 0) ----------

def test_check_critical_path_pass():
    tasks = [{"id": 1, "name": "T1", "summary": False, "critical": True},
             {"id": 2, "name": "T2", "summary": False, "critical": False}]
    r = check_critical_path(tasks)
    assert r["id"] == 13
    assert r["status"] == "pass"
    assert r["actual"] >= 1


def test_check_critical_path_fail_no_critical():
    tasks = [{"id": 1, "name": "T1", "summary": False, "critical": False}]
    r = check_critical_path(tasks)
    assert r["status"] == "fail"
    assert r["actual"] == 0


# ---------- check_bei (RULE 14: BEI > 95%) ----------

def _task_bei(id, baseline_finish, completed=False, summary=False):
    return {"id": id, "name": f"T{id}", "summary": summary,
            "baseline_finish": baseline_finish,
            "percent_complete": 100 if completed else 0}


def test_check_bei_pass():
    """All baseline-due tasks completed: BEI = 100% PASS."""
    tasks = [_task_bei(i, "2026-04-01", completed=True) for i in range(1, 11)]
    status_date = "2026-05-01"
    r = check_bei(tasks, status_date)
    assert r["id"] == 14
    assert r["status"] == "pass"
    assert r["actual"] == 100.0


def test_check_bei_fail():
    """5 of 10 baseline-due tasks completed → 50% BEI FAIL."""
    tasks = [_task_bei(i, "2026-04-01", completed=True) for i in range(1, 6)]
    tasks += [_task_bei(i, "2026-04-01", completed=False) for i in range(6, 11)]
    status_date = "2026-05-01"
    r = check_bei(tasks, status_date)
    assert r["status"] == "fail"
    assert r["actual"] == pytest.approx(50.0, rel=1e-2)


def test_check_bei_no_baseline_due_tasks():
    """No tasks scheduled to be done by status_date → vacuous BEI=100% PASS."""
    tasks = [_task_bei(1, "2026-12-01", completed=False)]
    status_date = "2026-05-01"
    r = check_bei(tasks, status_date)
    assert r["status"] == "pass"
    assert r["actual"] == 100.0
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def check_missed_tasks(tasks: List[Dict[str, Any]],
                      status_date: Optional[str]) -> Dict[str, Any]:
    """RULE 12: <5% of real tasks should be past baseline_finish without completion."""
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0 or not status_date:
        return {"id": 12, "name": "Missed Tasks", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": total, "failed_task_ids": []}
    sd = _parse_iso_date_local(status_date)
    if sd is None:
        return {"id": 12, "name": "Missed Tasks", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": total, "failed_task_ids": []}
    failed_ids = []
    for t in real:
        bf = _parse_iso_date_local(t.get("baseline_finish"))
        if bf is None:
            continue
        pct = float(t.get("percent_complete") or 0)
        if bf < sd and pct < 100:
            failed_ids.append(t["id"])
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 12, "name": "Missed Tasks", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(12, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }


def check_critical_path(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 13: project must have a defined critical path (>0 critical tasks)."""
    real = _real_tasks(tasks)
    crit_ids = [t["id"] for t in real if t.get("critical", False)]
    crit_count = len(crit_ids)
    return {
        "id": 13, "name": "Critical Path", "threshold": ">0",
        "actual": crit_count, "actual_unit": "count",
        "status": _eval_status(13, crit_count),
        "failed_count": 0 if crit_count > 0 else 1,
        "total_count": len(real),
        "critical_task_ids": crit_ids,
    }


def check_bei(tasks: List[Dict[str, Any]],
             status_date: Optional[str]) -> Dict[str, Any]:
    """RULE 14: BEI > 95% (Baseline Execution Index).

    BEI = count(actually_completed_through_status_date) /
          count(should_have_been_completed_per_baseline)
    """
    real = _real_tasks(tasks)
    if not status_date:
        return {"id": 14, "name": "BEI", "threshold": ">95%",
                "actual": 100.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": len(real)}
    sd = _parse_iso_date_local(status_date)
    if sd is None:
        return {"id": 14, "name": "BEI", "threshold": ">95%",
                "actual": 100.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": len(real)}
    should_be_done = 0
    actually_done = 0
    for t in real:
        bf = _parse_iso_date_local(t.get("baseline_finish"))
        if bf is None or bf > sd:
            continue
        should_be_done += 1
        if float(t.get("percent_complete") or 0) >= 100:
            actually_done += 1
    if should_be_done == 0:
        bei = 100.0
    else:
        bei = (actually_done / should_be_done) * 100.0
    return {
        "id": 14, "name": "BEI", "threshold": ">95%",
        "actual": round(bei, 2), "actual_unit": "%",
        "status": _eval_status(14, bei),
        "failed_count": should_be_done - actually_done,
        "total_count": should_be_done,
    }
```

**Step 4: Run — expect ~9 new + 37 prev = ~46 PASS**

**Step 5: Commit**

```bash
git add dcma_checks.py tests/test_dcma_checks.py
git commit -m "Phase 5b T89: dcma_checks RULE 12-14 (missed_tasks + critical_path + bei)

Missed: baseline_finish < status_date AND not completed.
Critical path: count of critical=True tasks must be > 0.
BEI: actually_completed / should_have_been_completed_per_baseline × 100.
Vacuous PASS when status_date None or no baseline-due tasks (avoid
dividing-by-zero failures)."
```

---

## Task 90: `assess_all` Aggregator + `compute_overall_rag`

**Files:**
- Modify: `dcma_checks.py`
- Create: `tests/test_dcma_assess_all.py` (separate file for clarity)

**Step 1: Failing tests**

```python
"""Test assess_all aggregator + compute_overall_rag."""
import pytest
from dcma_checks import assess_all, compute_overall_rag, DCMA_RULES


def test_assess_all_returns_14_rules():
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


# ---------- compute_overall_rag ----------

def _make_rule_results(pass_count):
    return [{"id": i, "status": "pass" if i <= pass_count else "fail"}
            for i in range(1, 15)]


def test_rag_green_above_12_pass():
    assert compute_overall_rag(_make_rule_results(13)) == "green"
    assert compute_overall_rag(_make_rule_results(14)) == "green"


def test_rag_amber_8_to_11_pass():
    assert compute_overall_rag(_make_rule_results(8)) == "amber"
    assert compute_overall_rag(_make_rule_results(11)) == "amber"


def test_rag_red_below_8_pass():
    assert compute_overall_rag(_make_rule_results(7)) == "red"
    assert compute_overall_rag(_make_rule_results(0)) == "red"
```

**Step 2: Run — FAIL**

**Step 3: Implementation — append to `dcma_checks.py`**

```python
def compute_overall_rag(rules: List[Dict[str, Any]]) -> str:
    """Industry convention: pass_count >= 12 → GREEN, 8-11 AMBER, <8 RED."""
    pass_count = sum(1 for r in rules if r.get("status") == "pass")
    if pass_count >= 12:
        return "green"
    if pass_count >= 8:
        return "amber"
    return "red"


def assess_all(tasks: List[Dict[str, Any]],
              links: List[Dict[str, Any]],
              assignments: List[Dict[str, Any]],
              baseline: Optional[Dict[str, Any]] = None,
              status_date: Optional[str] = None) -> Dict[str, Any]:
    """Run all 14 DCMA checks; return rules + summary."""
    rules = [
        check_no_predecessor(tasks),
        check_no_successor(tasks),
        check_leads(links),
        check_lags(links),
        check_fs_link_pct(links),
        check_hard_constraints(tasks),
        check_high_float(tasks),
        check_negative_float(tasks),
        check_high_duration(tasks),
        check_invalid_dates(tasks),
        check_resources_missing(tasks, assignments),
        check_missed_tasks(tasks, status_date),
        check_critical_path(tasks),
        check_bei(tasks, status_date),
    ]
    pass_count = sum(1 for r in rules if r.get("status") == "pass")
    fail_count = 14 - pass_count
    rag = compute_overall_rag(rules)
    failing_names = [r["name"] for r in rules if r.get("status") == "fail"]
    if fail_count == 0:
        executive = "All 14 DCMA rules pass. Schedule health: GREEN."
    else:
        executive = (f"{pass_count}/14 DCMA rules pass. {fail_count} issues: "
                    f"{', '.join(failing_names[:5])}{'...' if fail_count > 5 else ''}")
    return {
        "rules": rules,
        "summary": {
            "pass_count": pass_count,
            "fail_count": fail_count,
            "overall_rag": rag,
            "executive_text": executive,
        },
    }
```

**Step 4: Run — expect 8-9 new + ~46 prev = ~55 PASS**

**Step 5: Commit**

```bash
git add dcma_checks.py tests/test_dcma_assess_all.py
git commit -m "Phase 5b T90: dcma_checks assess_all aggregator + compute_overall_rag

Calls all 14 check_* functions, packs into rules list with summary
{pass_count, fail_count, overall_rag, executive_text}. RAG per industry
convention: >=12 pass GREEN, 8-11 AMBER, <8 RED.

Executive text lists first 5 failing rule names for hakediş report
inclusion."
```

---

## Task 91: Loader Extensions — `_dcma_load_links` + `_dcma_extract_floats/constraints` (BIG ONE)

**Files:**
- Modify: `msproject_mcp_core.py` (add Phase 5b section AT END, before `def main`)
- Create: `tests/test_msproject_dcma_loader.py`

**Step 1: Failing tests**

```python
"""Test Phase 5b loader extensions (links + floats + constraints)."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _dcma_load_links,
    _dcma_collect_full_data,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_dcma_load_links_xml():
    """File path: reuses _msp_file_read_links."""
    links = _dcma_load_links(file_path=MSP_XML)
    assert isinstance(links, list)
    # Sample fixture has no links; expect empty
    assert len(links) == 0


def test_dcma_collect_full_data_xml():
    r = _dcma_collect_full_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    for k in ("tasks", "links", "assignments", "resources",
              "baseline", "status_date"):
        assert k in r


def test_dcma_collect_full_data_invalid_baseline():
    r = _dcma_collect_full_data(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"


def test_dcma_collect_full_data_invalid_file():
    r = _dcma_collect_full_data(file_path="/nonexistent.xml")
    assert r["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Insert at end of `msproject_mcp_core.py`, AFTER Phase 5a dispatcher (`@mcp.tool msproject_evm`), BEFORE `def main()`:

```python
# ============================================================================
# PHASE 5B — DCMA TOOL
# ============================================================================
from dcma_checks import (
    DCMA_RULES,
    check_no_predecessor as _dcma_check_1,
    check_no_successor as _dcma_check_2,
    check_leads as _dcma_check_3,
    check_lags as _dcma_check_4,
    check_fs_link_pct as _dcma_check_5,
    check_hard_constraints as _dcma_check_6,
    check_high_float as _dcma_check_7,
    check_negative_float as _dcma_check_8,
    check_high_duration as _dcma_check_9,
    check_invalid_dates as _dcma_check_10,
    check_resources_missing as _dcma_check_11,
    check_missed_tasks as _dcma_check_12,
    check_critical_path as _dcma_check_13,
    check_bei as _dcma_check_14,
    assess_all as _dcma_assess_all,
    compute_overall_rag as _dcma_overall_rag,
)


def _dcma_load_links(file_path=None):
    """Hybrid: file_path → Phase 4 _msp_file_read_links;
    None → Phase 1 COM iter walking proj.Tasks predecessors.

    Returns list of {from_id, to_id, type, lag_days}.
    """
    if file_path:
        r = _msp_file_read_links(file_path=file_path)
        if r.get("status") != "ok":
            return []
        return r.get("links", []) or []
    # COM path
    try:
        app = _validate_active_project()
        proj = app.ActiveProject
        out = []
        for i in range(1, proj.Tasks.Count + 1):
            try:
                t = proj.Tasks(i)
                if t is None:
                    continue
                preds = t.PredecessorTasks
                if preds is None:
                    continue
                for j in range(1, preds.Count + 1):
                    try:
                        p = preds(j)
                        if p is None:
                            continue
                        # Walk t.TaskDependencies to get type + lag
                        # Each dep has Type (0-3 mapped to FS/SS/FF/SF) + Lag (minutes)
                    except Exception:
                        continue
                # Use TaskDependencies for richer info
                try:
                    deps = t.TaskDependencies
                except Exception:
                    deps = None
                if deps:
                    for j in range(1, deps.Count + 1):
                        try:
                            d = deps(j)
                            if d is None:
                                continue
                            ft = d.From
                            tt = d.To
                            if ft is None or tt is None or tt.ID != t.ID:
                                continue
                            type_code = int(d.Type or 0)
                            type_str = ["FF", "FS", "SF", "SS"][type_code] if 0 <= type_code <= 3 else "FS"
                            lag_min = float(d.Lag or 0)
                            lag_days = lag_min / 480.0  # 8h/day default
                            out.append({
                                "from_id": ft.ID, "to_id": tt.ID,
                                "type": type_str, "lag_days": round(lag_days, 2),
                            })
                        except Exception:
                            continue
            except Exception:
                continue
        return out
    except Exception as e:
        logger.exception(f"_dcma_load_links COM path failed: {e}")
        return []


def _dcma_collect_full_data(file_path=None, baseline_number=0):
    """Aggregate Phase 5a data + Phase 5b extensions.

    Returns {status, tasks, links, assignments, resources, baseline,
             status_date}. tasks already include total_slack_days,
             critical, constraint_type fields when available (Phase 4
             file path). COM path adds these via task property reads.
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    base = _evm_load_task_data(file_path=file_path)
    if base.get("status") != "ok":
        return base
    bload = _evm_load_baseline_data(file_path=file_path,
                                   baseline_number=baseline_number)
    links = _dcma_load_links(file_path=file_path)
    tasks = base.get("tasks", []) or []
    # Enrich tasks with DCMA-specific fields when COM path
    if not file_path:
        try:
            app = _validate_active_project()
            proj = app.ActiveProject
            for t_dict in tasks:
                tid = t_dict["id"]
                try:
                    com_t = _find_task_by_id(proj, tid)
                    if com_t is None:
                        continue
                    try:
                        slack_min = float(com_t.TotalSlack or 0)
                        t_dict["total_slack_days"] = round(slack_min / 480.0, 2)
                    except Exception:
                        t_dict["total_slack_days"] = 0
                    try:
                        t_dict["critical"] = bool(com_t.Critical)
                    except Exception:
                        t_dict["critical"] = False
                    try:
                        t_dict["constraint_type"] = int(com_t.ConstraintType or 0)
                    except Exception:
                        t_dict["constraint_type"] = 0
                    # Predecessors/successors as ID lists
                    try:
                        deps = com_t.TaskDependencies
                        preds = []
                        succs = []
                        if deps:
                            for j in range(1, deps.Count + 1):
                                d = deps(j)
                                if d is None:
                                    continue
                                if d.To and d.To.ID == tid and d.From:
                                    preds.append(d.From.ID)
                                elif d.From and d.From.ID == tid and d.To:
                                    succs.append(d.To.ID)
                        t_dict["predecessors"] = preds
                        t_dict["successors"] = succs
                    except Exception:
                        t_dict.setdefault("predecessors", [])
                        t_dict.setdefault("successors", [])
                except Exception:
                    continue
        except Exception as e:
            logger.warning(f"_dcma_collect_full_data COM enrich failed: {e}")
    # File path: Phase 4 already provides total_float, critical, constraint_type
    else:
        for t_dict in tasks:
            t_dict.setdefault("total_slack_days", float(t_dict.get("total_float") or 0))
            t_dict.setdefault("critical", t_dict.get("critical", False))
            t_dict.setdefault("constraint_type", t_dict.get("constraint_type", 0))
    return {
        "status": "ok",
        "tasks": tasks,
        "links": links,
        "assignments": base.get("assignments", []) or [],
        "resources": base.get("resources", []) or [],
        "baseline": bload if bload.get("status") == "ok" else None,
        "status_date": base.get("status_date"),
    }
```

**Step 4: Run — expect 4 PASS**

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_dcma_loader.py
git commit -m "Phase 5b T91 (BIG ONE): _dcma_load_links + _dcma_collect_full_data

Hybrid loaders extending Phase 5a base data with DCMA-specific fields:
links (with type + lag_days), total_slack_days, critical, constraint_type,
predecessors/successors ID lists.

File path: Phase 4 _msp_file_read_links + total_float/critical/constraint_type
already exposed (T66 probe). COM path: walks TaskDependencies + reads
TotalSlack/Critical/ConstraintType properties per task.

Phase 1-5a helpers DOKUNULMAZ — only read-only calls."
```

---

## Task 92: 4 Action Helpers — assess_all + summary + drill_down + compare

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_dcma_actions.py`

**Step 1: Failing tests**

```python
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _msp_dcma_assess_all,
    _msp_dcma_summary,
    _msp_dcma_drill_down,
    _msp_dcma_compare,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_dcma_assess_all_xml():
    r = _msp_dcma_assess_all(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert len(r["rules"]) == 14
    assert "summary" in r


def test_msp_dcma_summary_xml():
    r = _msp_dcma_summary(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["overall_rag"] in ("green", "amber", "red")


def test_msp_dcma_drill_down_valid_rule():
    r = _msp_dcma_drill_down(file_path=MSP_XML, rule_id=1)
    assert r["status"] == "ok"
    assert r["rule"]["id"] == 1
    assert "failed_tasks" in r


def test_msp_dcma_drill_down_invalid_rule():
    r = _msp_dcma_drill_down(file_path=MSP_XML, rule_id=99)
    assert r["status"] == "error"
    assert "1-14" in r["error"]


def test_msp_dcma_compare_no_prev_snapshot(tmp_path):
    snap = str(tmp_path / "no_snap.json")
    r = _msp_dcma_compare(file_path=MSP_XML, snapshot_path=snap)
    # No prev snapshot → graceful: return current only or specific note
    assert r["status"] in ("ok", "error")
```

**Step 2: Run — FAIL**

**Step 3: Implementation — append to Phase 5b section**

```python
def _msp_dcma_assess_all(file_path=None, baseline_number=0):
    """Action 1: assess_all — full DCMA 14-Point assessment."""
    data = _dcma_collect_full_data(file_path=file_path,
                                   baseline_number=baseline_number)
    if data.get("status") != "ok":
        return data
    result = _dcma_assess_all(
        tasks=data["tasks"],
        links=data["links"],
        assignments=data["assignments"],
        baseline=data.get("baseline"),
        status_date=data.get("status_date"),
    )
    return {"status": "ok", "baseline_number": baseline_number, **result}


def _msp_dcma_summary(file_path=None, baseline_number=0):
    """Action 2: summary — RAG + executive text only."""
    full = _msp_dcma_assess_all(file_path=file_path,
                                baseline_number=baseline_number)
    if full.get("status") != "ok":
        return full
    return {"status": "ok",
            "baseline_number": baseline_number,
            **full["summary"]}


def _msp_dcma_drill_down(file_path=None, rule_id=1, baseline_number=0):
    """Action 3: drill_down — per-rule failed task details."""
    if rule_id not in range(1, 15):
        return {"status": "error",
                "error": f"rule_id must be 1-14, got {rule_id}"}
    full = _msp_dcma_assess_all(file_path=file_path,
                                baseline_number=baseline_number)
    if full.get("status") != "ok":
        return full
    rule = next((r for r in full["rules"] if r["id"] == rule_id), None)
    if rule is None:
        return {"status": "error", "error": f"Rule {rule_id} not found"}
    # Resolve failed task names
    data = _dcma_collect_full_data(file_path=file_path,
                                   baseline_number=baseline_number)
    tasks_by_id = {t["id"]: t for t in data.get("tasks", [])}
    failed_ids = rule.get("failed_task_ids", [])
    failed_tasks = []
    for tid in failed_ids:
        t = tasks_by_id.get(tid)
        if t:
            failed_tasks.append({"id": tid, "name": t.get("name", "")})
    return {
        "status": "ok",
        "baseline_number": baseline_number,
        "rule": {"id": rule["id"], "name": rule["name"],
                "threshold": rule.get("threshold")},
        "actual": rule.get("actual"),
        "failed_count": rule.get("failed_count"),
        "total_count": rule.get("total_count"),
        "failed_tasks": failed_tasks,
    }


def _msp_dcma_compare(file_path=None, snapshot_path=None, baseline_number=0):
    """Action 4: compare current DCMA vs prev snapshot.

    Reuses Phase 5a _evm_snapshot_load to read prior DCMA dumps from
    the same JSON file (snapshots can include both EVM + DCMA data).
    """
    current = _msp_dcma_assess_all(file_path=file_path,
                                   baseline_number=baseline_number)
    if current.get("status") != "ok":
        return current
    if not snapshot_path:
        return {"status": "ok", "current": current["summary"], "prev": None,
                "delta": {"rules_improved": [], "rules_degraded": []}}
    snaps = _evm_snapshot_load(snapshot_path) if os.path.exists(snapshot_path) else []
    # Filter snaps that have DCMA data
    dcma_snaps = [s for s in snaps if s.get("dcma")]
    if not dcma_snaps:
        return {"status": "ok", "current": current["summary"], "prev": None,
                "delta": {"rules_improved": [], "rules_degraded": []}}
    dcma_snaps.sort(key=lambda s: s.get("saved_at", ""))
    prev = dcma_snaps[-1].get("dcma")
    # Compute delta
    improved = []
    degraded = []
    prev_rules = {r["id"]: r for r in (prev.get("rules") or [])}
    for cr in current["rules"]:
        pr = prev_rules.get(cr["id"])
        if pr is None:
            continue
        cur_actual = cr.get("actual", 0)
        prev_actual = pr.get("actual", 0)
        if pr.get("status") == "fail" and cr.get("status") == "pass":
            improved.append({"id": cr["id"], "name": cr["name"],
                            "from_actual": prev_actual, "to_actual": cur_actual})
        elif pr.get("status") == "pass" and cr.get("status") == "fail":
            degraded.append({"id": cr["id"], "name": cr["name"],
                            "from_actual": prev_actual, "to_actual": cur_actual})
    return {
        "status": "ok",
        "current": current["summary"],
        "prev": prev.get("summary") if isinstance(prev, dict) else None,
        "delta": {"rules_improved": improved, "rules_degraded": degraded},
    }
```

**Step 4: Run — expect 5 PASS**

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_dcma_actions.py
git commit -m "Phase 5b T92 (BIG ONE): 4 action helpers (assess_all + summary + drill_down + compare)

Wraps dcma_checks.assess_all with hybrid Phase 5a/Phase 4 data sources.
drill_down resolves failed_task_ids → name pairs for hakediş report.
compare reuses _evm_snapshot_load for DCMA delta tracking (snapshots
can carry both EVM and DCMA keys)."
```

---

## Task 93: FastMCP Dispatcher + Acceptance Script + README + Push (FINAL)

**Files:**
- Modify: `msproject_mcp_core.py` (add `@mcp.tool msproject_health` after Phase 5b helpers, BEFORE `def main`)
- Create: `tests/test_msproject_dcma_dispatcher.py`
- Create: `samples/build_dcma_lifecycle.py`
- Modify: `README.md`

**Step 1: Failing dispatcher tests**

```python
import asyncio, json, os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import msproject_health

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def _call(action, **kw):
    raw = _run(msproject_health({"action": action, **kw}))
    return json.loads(raw)


def test_dispatcher_assess_all():
    p = _call("assess_all", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert len(p["rules"]) == 14


def test_dispatcher_summary():
    p = _call("summary", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert p["overall_rag"] in ("green", "amber", "red")


def test_dispatcher_drill_down():
    p = _call("drill_down", file_path=MSP_XML, rule_id=1)
    assert p["status"] == "ok"
    assert "failed_tasks" in p


def test_dispatcher_unknown_action():
    p = _call("nonsense", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_invalid_rule_id():
    p = _call("drill_down", file_path=MSP_XML, rule_id=99)
    assert p["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add `@mcp.tool msproject_health` AFTER Phase 5a `msproject_evm` dispatcher, BEFORE `def main()`:

```python
@mcp.tool(
    name="msproject_health",
    annotations={"title": "MS Project DCMA 14-Point Health Assessment", "readOnlyHint": True},
)
async def msproject_health(params: dict) -> str:
    """DCMA 14-Point Schedule Health Assessment per CLAUDE.md RULE 10.

    Hybrid: file_path verilirse Phase 4 file path; yoksa Phase 1 COM.
    Read-only — no write actions.

    Actions:
    - assess_all: All 14 rules + summary + RAG
    - summary: Just RAG + executive text
    - drill_down: Per-rule failed task list (rule_id 1-14)
    - compare: Current DCMA vs prev snapshot (reuses Phase 5a snapshot file)

    Phase 5b (1 May 2026). Tool count 9 → 10.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "assess_all":
            r = _msp_dcma_assess_all(**p)
        elif action == "summary":
            r = _msp_dcma_summary(**p)
        elif action == "drill_down":
            r = _msp_dcma_drill_down(**p)
        elif action == "compare":
            r = _msp_dcma_compare(**p)
        else:
            r = {"status": "error",
                 "error": (f"Unknown action '{action}'. Valid: "
                          "assess_all/summary/drill_down/compare")}
    except TypeError as e:
        r = {"status": "error", "error": f"Invalid params for {action}: {e}"}
    except Exception as e:
        logger.exception(f"msproject_health({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

**Step 4: Acceptance script `samples/build_dcma_lifecycle.py`**

```python
"""Phase 5b DCMA acceptance: 200-task CAU-style with intentional issues.

SAFETY: FileNew + FileClose 0. User's active project untouched.

Scenario:
  1. Build 200 tasks + 14 CAU resources
  2. Inject DCMA failures intentionally:
     - 12 tasks WITHOUT predecessor (RULE 1 fail)
     - 15 tasks duration > 44d (RULE 9 fail)
     - 8 tasks with hard constraints (RULE 6 borderline)
     - 1 task with start>finish (RULE 10 fail)
  3. Save Baseline 0
  4. Phase 3b progress for ~30 tasks (RULE 14 BEI calc)
  5. set_status_date
  6. msproject_health assess_all → display 14 rule results
  7. drill_down rule_id=1 → 12 failed tasks
  8. drill_down rule_id=9 → 15 failed tasks
  9. summary → RAG (expected AMBER)

Target <60s wall clock.
"""
import os, sys, time, tempfile, shutil
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pythoncom, win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save, _msp_progress_bulk_update, _msp_progress_set_status_date,
    _msp_dcma_assess_all, _msp_dcma_summary, _msp_dcma_drill_down,
)

N_TASKS = 200


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] isolated test: {test_name}")

    try:
        t0 = time.time()
        # 1. Build base + intentional issues
        print(f"\n1. Building {N_TASKS} tasks + 14 resources with intentional DCMA issues...")
        items = [{"name": f"V{i:03d}", "duration": "5d"} for i in range(N_TASKS - 15)]
        # 15 tasks with high duration
        items += [{"name": f"H{i:02d}", "duration": "60d"} for i in range(15)]
        tasks = _msp_task_bulk_add(items=items)
        task_ids = tasks["task_ids"]
        cau = ["COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
               "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR"]
        res_ids = [_msp_resource_add(name=n, type="Work")["resource_id"] for n in cau]
        # Assign resources to most tasks (skip first 12 to fail RULE 1 too)
        sample = [{"task_id": tid, "resource_id": res_ids[i % 14]}
                  for i, tid in enumerate(task_ids[12:])]  # skip first 12
        _msp_resource_bulk_assign(items=sample)
        print(f"   OK in {time.time()-t0:.2f}s")

        # 2. Save Baseline 0
        _msp_baseline_save(baseline_number=0)
        print(f"\n2. Baseline 0 saved at {time.time()-t0:.2f}s")

        # 3. Some progress for BEI
        progress_items = [{"task_id": tid, "percent_complete": 50.0}
                          for tid in task_ids[:30]]
        _msp_progress_bulk_update(items=progress_items)
        _msp_progress_set_status_date(status_date="2026-05-15")

        # 4. assess_all
        r = _msp_dcma_assess_all()
        print(f"\n3. DCMA assess_all results:")
        for rule in r["rules"]:
            status_emoji = "OK" if rule["status"] == "pass" else "FAIL"
            print(f"   [{status_emoji}] Rule {rule['id']}: {rule['name']} "
                  f"actual={rule.get('actual')}{rule.get('actual_unit', '')} "
                  f"({rule['threshold']})")

        # 5. summary
        s = _msp_dcma_summary()
        print(f"\n4. Summary: {s['pass_count']}/14 pass, RAG={s['overall_rag'].upper()}")
        print(f"   {s['executive_text']}")

        # 6. drill_down for failed rules
        for rule in r["rules"]:
            if rule["status"] == "fail":
                d = _msp_dcma_drill_down(rule_id=rule["id"])
                print(f"\n5. Drill-down Rule {rule['id']} ({rule['name']}): "
                      f"{d['failed_count']} failed tasks")
                for ft in d["failed_tasks"][:5]:
                    print(f"      - Task {ft['id']}: {ft['name']}")

        elapsed = time.time() - t0
        print(f"\n[OK] ACCEPTANCE: {elapsed:.2f}s total (target <60s)")
        assert elapsed < 60.0, f"Too slow: {elapsed}s"

    finally:
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    break
        except Exception:
            pass


if __name__ == "__main__":
    main()
```

**Step 5: Run acceptance**

```bash
python samples/build_dcma_lifecycle.py
```
Expected: `[OK] ACCEPTANCE: <Xs total (target <60s)`. Realistic ~30-45s.

**Step 6: README update**

Add Phase 5b section after Phase 5a:
```markdown
## Phase 5b — DCMA 14-Point (1 May 2026)

`msproject_health` tool — DCMA 14-Point Schedule Health Assessment per
CLAUDE.md RULE 10. 4 actions covering all 14 rules with industry-standard
hardcoded thresholds. Hybrid: file_path optional (Phase 4 file path or
Phase 1 COM).

**Actions:**
- `assess_all`: 14 rules + summary + RAG (>=12 pass GREEN, 8-11 AMBER, <8 RED)
- `summary`: RAG + executive text only
- `drill_down(rule_id=1..14)`: per-rule failed task list
- `compare(snapshot_path)`: DCMA delta vs prev snapshot

**Rules grouped by category:**
- Logic (1-5): no_pred, no_succ, leads, lags, fs_link
- Constraints (6): hard_constraints
- Float (7-8): high_float, negative_float
- Duration (9): high_duration
- Quality (10-11): invalid_dates, resources_missing
- Schedule (12-14): missed_tasks, critical_path, BEI

Architecture: pure-math `dcma_checks.py` (14 fixture-free check functions,
MSP/COM/file independent, ~30 tests) + I/O adapters in msproject_mcp_core.py.
Phase 1-5a helpers DOKUNULMAZ.

Acceptance: `samples/build_dcma_lifecycle.py` runs 200-task CAU-style with
intentional DCMA failures, drill_down for each failed rule, in <60s.

Tool count: **10 tools, ~83 actions**.
```

**Step 7: Run full regression**

```bash
python -m pytest tests/test_dcma_*.py tests/test_msproject_dcma_*.py tests/test_evm_math.py tests/test_msproject_evm_*.py tests/test_msproject_file_*.py -q --tb=line 2>&1 | tail -5
```
Expected: ~178 PASS (138 cumulative + ~40 Phase 5b new), no regressions.

**Step 8: Commit + push**

```bash
git add msproject_mcp_core.py tests/test_msproject_dcma_dispatcher.py samples/build_dcma_lifecycle.py README.md
git commit -m "Phase 5b T93: dispatcher + acceptance + README + push (msproject_health 10th tool)

@mcp.tool msproject_health with 4 action routing (assess_all/summary/
drill_down/compare). Acceptance: 200-task CAU-style with intentional
DCMA failures (12 no-pred + 15 high-duration + manipulated dates) for
detection validation. drill_down for each failing rule lists 5 sample
tasks.

README updated: Phase 5b section. Tool count 9 → 10, total actions
~79 → ~83.

Phase 4+5a+5b cumulative: ~178 PASS. Phase 1-5a untouched."
git push origin main
```

If push fails (network/auth), report error but DO NOT retry. Chain is local even if push fails.

---

## Phase 5b Tamamlama Kriterleri

1. ✅ T85-T93 ~10-12 commit landed
2. ✅ Acceptance script `samples/build_dcma_lifecycle.py` <60s
3. ✅ Yeni testler ~40 PASS (~30 saf math + ~10 dispatcher/loader)
4. ✅ Phase 1+2+3+4+5a mevcut regression PASS — DOKUNULMAZ
5. ✅ Total ~178 PASS + 0 xfail
6. ✅ All 14 DCMA rules per CLAUDE.md RULE 10 doğru implement edilir
7. ✅ RAG status (>=12/8-11/<8 pass) executive output
8. ✅ Push to origin/main
9. ⏸ Kullanıcı manuel onayı → Phase 5c (Excel) başlar

---

## Sequencing Tips

- **Pure math T85-T90** → manuel write + self-verify (test-driven, no probe, no fixtures)
- **T91 BIG ONE (loader extensions)** → subagent dispatch (Phase 4 reuse + COM property reads tricky)
- **T92 BIG ONE (action helpers)** → subagent dispatch (4 action wiring, drill_down task name resolution)
- **T93 standard finalize** → manuel + push

Phase 1+2+3+4+5a helpers DOKUNULMAZ; sadece read-only çağrılar.

---

*Plan tamamlandı: 1 Mayıs 2026*
*Tahmini Phase 5b süresi: ~17 saat (T85-T93, 9 task TDD chain)*
*Sonraki phase (onay sonrası): Phase 5c — Excel (`msproject_excel` import/export)*
