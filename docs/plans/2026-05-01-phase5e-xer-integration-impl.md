# Phase 5e XER Integration Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task (T109-T111).

**Goal:** Wire `.xer` file_path support into Phase 5a `_evm_load_task_data` + `_evm_load_baseline_data` loaders so Phase 5b DCMA + Phase 5c Excel collect helpers automatically work on XER files.

**Architecture:** Additive guard in Phase 5a loaders + new Phase 5E adapter section translating XerFile output to Phase 5a shape. NO new MCP tools. Tool count stays 12.

**Tech Stack:** Python 3.12, existing `xer_parser.XerFile` from Phase 5d, existing `_evm_load_task_data`/`_evm_load_baseline_data` from Phase 5a.

**Design doc:** `docs/plans/2026-05-01-phase5e-xer-integration-design.md` (commit `31b0770`)

**Baseline state at start:** HEAD `31b0770`, 370 cumulative PASS.

---

## Task 109: `_xer_to_evm_task_shape` Adapter + Routing in `_evm_load_task_data`

**Files:**
- Modify: `msproject_mcp_core.py` (add Phase 5E section AFTER Phase 5d xer dispatcher; add 1-line guard in `_evm_load_task_data`)
- Create: `tests/test_phase5e_adapter.py`

### Step 1: Failing tests

```python
"""Test Phase 5e XER -> Phase 5a shape adapter."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import _xer_to_evm_task_shape, _evm_load_task_data
from xer_parser import XerFile


def test_adapter_returns_phase5a_shape(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    assert r["status"] == "ok"
    for k in ("tasks", "resources", "assignments", "status_date", "project_file"):
        assert k in r


def test_adapter_tasks_have_baseline_fields(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    for t in r["tasks"]:
        for k in ("baseline_start", "baseline_finish", "baseline_work", "actual_work"):
            assert k in t


def test_adapter_baseline_equals_target_cau_pattern(sample_cau_xer):
    """CAU cost-loaded NO: baseline = current target schedule."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    assert foundation["baseline_start"] == foundation["start"]
    assert foundation["baseline_finish"] == foundation["finish"]
    assert foundation["baseline_work"] == foundation["duration_h"]


def test_adapter_actual_work_aggregated(sample_cau_xer):
    """actual_work = sum of TASKRSRC.act_reg_qty per task_id."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    # COW 180 + STL 1000 = 1180
    assert foundation["actual_work"] == 1180.0


def test_adapter_predecessors_derived(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    frame = next(t for t in r["tasks"] if t["id"] == 1002)
    assert 1001 in frame["predecessors"]


def test_adapter_successors_derived(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    assert 1002 in foundation["successors"]


def test_adapter_critical_from_zero_slack(sample_cau_xer):
    """critical = total_slack_days <= 0 (XER lacks explicit critical flag)."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    # total_float_hr_cnt=0 → critical=True
    assert foundation["critical"] is True
    walls = next(t for t in r["tasks"] if t["id"] == 1003)
    # total_float_hr_cnt=72 → critical=False
    assert walls["critical"] is False


def test_adapter_total_slack_days_present(sample_cau_xer):
    """total_slack_days needed for DCMA Rule 7-8."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    walls = next(t for t in r["tasks"] if t["id"] == 1003)
    assert walls["total_slack_days"] == walls["total_float"]


def test_adapter_excludes_summary_tasks(sample_cau_xer):
    """Summary tasks (TT_LOE/TT_WBS) excluded — but CAU fixture has none."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    # Fixture has 6 real tasks (incl. milestone) + 0 summaries
    assert len(r["tasks"]) == 6


# ---- _evm_load_task_data routing ----

def test_evm_load_task_data_routes_xer(sample_cau_xer):
    """When file_path ends .xer, returns Phase 5e adapter output."""
    r = _evm_load_task_data(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert len(r["tasks"]) == 6
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    assert "baseline_work" in foundation


def test_evm_load_task_data_xml_path_unchanged(tmp_path):
    """Existing .xml path unchanged - falls through original logic."""
    # Use sample_msp.xml fixture (has 0 tasks but loader should not crash)
    msp_xml = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")
    if not os.path.exists(msp_xml):
        return  # skip if fixture absent
    r = _evm_load_task_data(file_path=msp_xml)
    # Phase 4 file path — status ok or error depending on fixture
    assert r["status"] in ("ok", "error")
```

### Step 2: Run — expect ImportError

### Step 3: Implementation

Add Phase 5E section AFTER `msproject_xer` dispatcher (at end of Phase 5D section), BEFORE def main:

```python
# ============================================================================
# PHASE 5E - XER NATIVE INTEGRATION (Phase 5a loader extensions for .xer)
# ============================================================================


def _xer_to_evm_task_shape(xer):
    """Translate XerFile output to Phase 5a _evm_load_task_data shape.

    CAU pattern (cost-loaded NO): baseline = target schedule. baseline_work
    = duration_h, baseline_start/finish = target dates.

    Derives: predecessors/successors lists, total_slack_days,
    critical (heuristic: total_slack_days <= 0).
    """
    cals = xer.read_calendars()
    day_hr_cnt = cals[0]["day_hr_cnt"] if cals else 8.0
    raw_tasks = xer.read_tasks(day_hr_cnt=day_hr_cnt)
    links = xer.read_links()
    progress = xer.read_progress()
    assignments = xer.read_assignments()

    # Pre-aggregate actual_work per task from XER assignments
    actual_by_task = {}
    for a in assignments:
        tid = a.get("task_id")
        if tid is not None:
            actual_by_task[tid] = actual_by_task.get(tid, 0.0) + float(a.get("actual_qty") or 0)

    # Pre-build predecessor/successor maps from links
    preds_by_task = {}
    succs_by_task = {}
    for link in links:
        from_id = link.get("from_id")
        to_id = link.get("to_id")
        if to_id is not None and from_id is not None:
            preds_by_task.setdefault(to_id, []).append(from_id)
            succs_by_task.setdefault(from_id, []).append(to_id)

    # Build Phase 5a-shape task dicts (filter summaries)
    out_tasks = []
    for t in raw_tasks:
        if t.get("summary", False):
            continue
        tid = t["id"]
        slack = t.get("total_float", 0)
        out_tasks.append({
            **t,  # carry all XER fields (id, name, code, duration_h, start, finish,
                  # actual_start, actual_finish, percent_complete, total_float,
                  # summary, task_type, constraint_type, status)
            "baseline_start": t.get("start"),
            "baseline_finish": t.get("finish"),
            "baseline_work": float(t.get("duration_h") or 0),
            "actual_work": actual_by_task.get(tid, 0.0),
            "total_slack_days": slack,
            "critical": float(slack) <= 0,
            "predecessors": preds_by_task.get(tid, []),
            "successors": succs_by_task.get(tid, []),
        })

    return {
        "status": "ok",
        "tasks": out_tasks,
        "resources": xer.read_resources(),
        "assignments": assignments,
        "status_date": progress.get("status_date"),
        "project_file": xer.file_path,
    }
```

Now extend `_evm_load_task_data` (find at line ~4779) by adding ONE guard at the top:

```python
def _evm_load_task_data(file_path=None):
    """Hybrid: file_path -> Phase 4 file path; None -> Phase 1 COM path.
    ...
    """
    try:
        # Phase 5e: route .xer extension to dedicated adapter
        if file_path and isinstance(file_path, str) and file_path.lower().endswith(".xer"):
            from xer_parser import XerFile
            return _xer_to_evm_task_shape(XerFile(file_path))
        if file_path:
            # ... existing code ...
```

### Step 4-5: Run + commit

```bash
git add msproject_mcp_core.py tests/test_phase5e_adapter.py
git commit -m "Phase 5e T109: _xer_to_evm_task_shape adapter + .xer routing in _evm_load_task_data"
```

---

## Task 110: `_xer_to_evm_baseline_shape` + Routing + Integration Tests

**Files:**
- Modify: `msproject_mcp_core.py` (add baseline adapter + extend `_evm_load_baseline_data`)
- Create: `tests/test_phase5e_integration.py`

### Step 1: Failing tests

```python
"""Test Phase 5e end-to-end: .xer file_path through DCMA + Excel pipelines."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _evm_load_baseline_data,
    _msp_dcma_assess_all,
    _msp_dcma_summary,
    _msp_dcma_drill_down,
    _msp_excel_export_hakedis,
    _msp_excel_export_dcma,
)
from openpyxl import load_workbook


def test_evm_load_baseline_data_xer(sample_cau_xer):
    """XER baseline = target schedule (CAU cost-loaded NO)."""
    r = _evm_load_baseline_data(file_path=sample_cau_xer, baseline_number=0)
    assert r["status"] == "ok"
    assert "tasks" in r
    foundation = next(t for t in r["tasks"] if t.get("task_id") == 1001 or t.get("id") == 1001)
    assert "baseline_finish" in foundation or "finish" in foundation


def test_dcma_assess_all_xer(sample_cau_xer):
    """msproject_health.assess_all on .xer returns 14 rules."""
    r = _msp_dcma_assess_all(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert len(r["rules"]) == 14
    assert "summary" in r


def test_dcma_summary_xer(sample_cau_xer):
    r = _msp_dcma_summary(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["overall_rag"] in ("green", "amber", "red")


def test_dcma_drill_down_xer(sample_cau_xer):
    """drill_down for Rule 1 (no_predecessor) on XER."""
    r = _msp_dcma_drill_down(file_path=sample_cau_xer, rule_id=1)
    assert r["status"] == "ok"
    assert "failed_tasks" in r
    # Foundation has no predecessor in CAU XER fixture (first task)
    failed_ids = [t["id"] for t in r["failed_tasks"]]
    assert 1001 in failed_ids


def test_excel_export_dcma_xer(sample_cau_xer, tmp_path):
    """export_dcma on XER produces 2-sheet xlsx."""
    xlsx = tmp_path / "dcma.xlsx"
    r = _msp_excel_export_dcma(file_path=sample_cau_xer, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    assert xlsx.exists()
    wb = load_workbook(str(xlsx), read_only=True)
    assert "DCMA_Rules" in wb.sheetnames
    assert "DCMA_Failed" in wb.sheetnames


def test_excel_export_hakedis_xer(sample_cau_xer, tmp_path):
    """export_hakedis on XER produces 6-sheet workbook."""
    xlsx = tmp_path / "hak.xlsx"
    r = _msp_excel_export_hakedis(file_path=sample_cau_xer, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    wb = load_workbook(str(xlsx), read_only=True)
    for s in ("Summary", "Tasks", "EVM_Compute", "EVM_TimePhased",
              "DCMA_Rules", "DCMA_Failed"):
        assert s in wb.sheetnames


def test_dcma_xer_rule_13_critical_path(sample_cau_xer):
    """CAU XER chain has zero-slack tasks → critical_path > 0 → Rule 13 PASS."""
    r = _msp_dcma_assess_all(file_path=sample_cau_xer)
    rule13 = next(rule for rule in r["rules"] if rule["id"] == 13)
    assert rule13["status"] == "pass"
    assert rule13["actual"] >= 1


def test_dcma_xer_rule_5_fs_link_pct(sample_cau_xer):
    """All 5 links are FS → Rule 5 PASS (>90% FS)."""
    r = _msp_dcma_assess_all(file_path=sample_cau_xer)
    rule5 = next(rule for rule in r["rules"] if rule["id"] == 5)
    assert rule5["status"] == "pass"
    assert rule5["actual"] == 100.0
```

### Step 3: Implementation

Add baseline adapter + guard:

```python
def _xer_to_evm_baseline_shape(xer, baseline_number=0):
    """Translate XerFile to Phase 5a _evm_load_baseline_data shape.

    CAU pattern: baseline = target. Returns task baseline fields keyed by task_id
    (Phase 5a baseline shape).
    """
    cals = xer.read_calendars()
    day_hr_cnt = cals[0]["day_hr_cnt"] if cals else 8.0
    raw_tasks = xer.read_tasks(day_hr_cnt=day_hr_cnt)
    out = []
    for t in raw_tasks:
        if t.get("summary", False):
            continue
        out.append({
            "task_id": t["id"],
            "id": t["id"],
            "name": t.get("name", ""),
            "start": t.get("start"),
            "finish": t.get("finish"),
            "baseline_start": t.get("start"),
            "baseline_finish": t.get("finish"),
            "work_h": float(t.get("duration_h") or 0),
            "baseline_work": float(t.get("duration_h") or 0),
        })
    return {
        "status": "ok",
        "baseline_number": baseline_number,
        "tasks": out,
    }
```

Extend `_evm_load_baseline_data` (find ~line 4882) with guard at top:

```python
def _evm_load_baseline_data(file_path=None, baseline_number=0):
    """..."""
    try:
        # Phase 5e: route .xer extension
        if file_path and isinstance(file_path, str) and file_path.lower().endswith(".xer"):
            from xer_parser import XerFile
            return _xer_to_evm_baseline_shape(XerFile(file_path), baseline_number)
        # ... existing code ...
```

### Step 4-5: Commit

```bash
git commit -m "Phase 5e T110: _xer_to_evm_baseline_shape + integration tests (DCMA + Excel on .xer)"
```

---

## Task 111: Acceptance + README + Push (FINAL)

**Files:**
- Create: `samples/build_xer_dcma_excel_lifecycle.py`
- Modify: `README.md`

**Acceptance scenario:**
1. Write synthetic CAU XER to tempdir
2. `_msp_dcma_assess_all(file_path=xer)` → 14 rules
3. `_msp_excel_export_hakedis(file_path=xer, xlsx_path=...)` → 6-sheet xlsx
4. Verify
5. ≤30s

**README addition (Phase 5e section after Phase 5d):**
```markdown
## Phase 5e — XER Native Integration (1 May 2026)

Phase 5d shipped `msproject_xer` reader; Phase 5e wires it into Phase 5a
EVM + Phase 5b DCMA + Phase 5c Excel via additive `.xer` extension routing
in Phase 5a loaders. NO new tools — existing `msproject_health.assess_all`
+ `msproject_excel.export_hakedis` etc. now accept `.xer` file_path
directly. CAU XER → DCMA + hakediş end-to-end.
```

**Commit + push:**
```bash
git add msproject_mcp_core.py samples/build_xer_dcma_excel_lifecycle.py README.md
git commit -m "Phase 5e T111: acceptance + README + push (XER native integration complete)"
git push origin main
```

---

## Acceptance Criteria

1. ✅ T109-T111 chain landed
2. ✅ Acceptance ≤ 30s (XER → DCMA + Excel end-to-end)
3. ✅ Phase 1-5d regression untouched (370 PASS)
4. ✅ Push origin/main
5. ✅ Tool count stays 12

---

*Plan committed: 2026-05-01.*
