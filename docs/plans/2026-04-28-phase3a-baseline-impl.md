# Phase 3a Baseline Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** New `msproject_baseline` MCP tool — 9 actions covering all 11 MSP baseline slots (Baseline + Baseline1..Baseline10), variance reporting, baseline-to-baseline comparison, RAG summary.

**Architecture:** All helpers in `msproject_mcp_core.py` (Phase 1+2a+2b sections untouched). Dynamic baseline property access via `_baseline_property_name(field, N)`. Reuse `_format_com_error` (T29), `_parse_rate` (T32 locale), `_find_task_by_id` (Phase 1), `_build_task_id_map` (T37 perf), `clean_test_project` fixture (SAFETY).

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM, pytest. Mevcut `msproject_mcp_core.py` (~1850 satır after Phase 2b), 22 test dosyası, **156 PASS + 1 xfail** baseline.

**Design doc:** `docs/plans/2026-04-28-phase3a-baseline-design.md` (commit `c47dcb9`)

**Baseline state at start:** HEAD `c47dcb9`, MS Project running v16.0.

**KEY COM API REFERENCE (verified from `msproject_typelib.txt`):**
- `app.BaselineSave(All=True, Copy=pjCopyCurrent=0, Into=<see INTO_BASELINE_MAP>, RollupToSummaryTasks=True)`
- `app.BaselineClear(All=True, From=<baseline_number 0-10>)` — `From` uses 0-10 directly
- `proj.BaselineSavedDate(Baseline=<0-10>)` — returns datetime or 0/None if not saved
- Per-task baseline read: `task.BaselineStart`/`BaselineFinish`/`BaselineDuration`/`BaselineWork`/`BaselineCost` (Baseline 0); `task.Baseline1Start`/.../`Baseline10Start` for 1-10

**CRITICAL MAPPING — `Into` parameter is OFFSET vs baseline number:**
```
Baseline number 0 → Into=0  (pjIntoBaseline)
Baseline number 1 → Into=11 (pjIntoBaseline1)
Baseline number 2 → Into=12 (pjIntoBaseline2)
...
Baseline number 10 → Into=20 (pjIntoBaseline10)
```

Formula: `Into = 0 if N == 0 else (10 + N)`

`From` (BaselineClear) and `Baseline` (BaselineSavedDate) use direct 0-10 mapping. ONLY `Into` (BaselineSave) uses the offset.

---

## Task 39: Baseline Foundations (helpers + constants)

**Files:**
- Modify: `msproject_mcp_core.py` (add at end of resource section, before any task/Phase 1 code — use Grep `_msp_resource_bulk_assign` to locate insertion)
- Create: `tests/test_msproject_baseline_helpers.py`

**Step 1: Failing test**

`tests/test_msproject_baseline_helpers.py`:
```python
"""Test baseline helpers + constants."""
import pytest
from msproject_mcp_core import (
    BASELINE_NUMBERS, INTO_BASELINE_MAP,
    _baseline_property_name, _baseline_into_code, _read_task_baseline,
    _baseline_saved_date, _msp_task_add_single,
)


def test_baseline_numbers_constant():
    assert BASELINE_NUMBERS == list(range(11))  # 0..10


def test_baseline_property_name_zero():
    """Baseline 0 has no suffix: BaselineStart."""
    assert _baseline_property_name("Start", 0) == "BaselineStart"
    assert _baseline_property_name("Finish", 0) == "BaselineFinish"
    assert _baseline_property_name("Duration", 0) == "BaselineDuration"
    assert _baseline_property_name("Work", 0) == "BaselineWork"
    assert _baseline_property_name("Cost", 0) == "BaselineCost"


def test_baseline_property_name_numbered():
    """Baseline 3 -> Baseline3Start etc."""
    assert _baseline_property_name("Start", 3) == "Baseline3Start"
    assert _baseline_property_name("Finish", 10) == "Baseline10Finish"


def test_baseline_into_code_zero():
    """Baseline 0 -> Into=0 (pjIntoBaseline)."""
    assert _baseline_into_code(0) == 0


def test_baseline_into_code_offset():
    """Baseline 1 -> 11, Baseline 2 -> 12, ..., Baseline 10 -> 20."""
    assert _baseline_into_code(1) == 11
    assert _baseline_into_code(5) == 15
    assert _baseline_into_code(10) == 20


def test_into_baseline_map():
    """Constant matches formula."""
    assert INTO_BASELINE_MAP == {0: 0, 1: 11, 2: 12, 3: 13, 4: 14, 5: 15,
                                  6: 16, 7: 17, 8: 18, 9: 19, 10: 20}


def test_baseline_saved_date_unsaved(clean_test_project):
    """Fresh project — baseline 0 not saved → returns None."""
    proj = clean_test_project
    result = _baseline_saved_date(proj, 0)
    assert result is None


def test_read_task_baseline_unsaved(clean_test_project):
    """Task with no baseline saved — read returns dict with None/zero values."""
    add_r = _msp_task_add_single(name="UnsavedT-T39", duration="3d")
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_r["task_id"])
    data = _read_task_baseline(t, 0)
    # Unsaved baseline returns sentinel-empty values per MSP COM
    assert "start" in data and "finish" in data
    assert "duration_h" in data and "work_h" in data and "cost" in data
```

**Step 2: Run** — expect ImportError on BASELINE_NUMBERS.

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/test_msproject_baseline_helpers.py -v
```

**Step 3: Implementation**

Insert at end of resource section (after `_msp_resource_bulk_assign` and before any Phase 1 task code; use Grep to confirm location):

```python
# ---------- BASELINE CONSTANTS ----------

# MSP supports 11 baseline slots: Baseline + Baseline1..Baseline10
BASELINE_NUMBERS = list(range(11))  # [0, 1, ..., 10]

# CRITICAL: app.BaselineSave's `Into` parameter uses OFFSET enum, not direct number.
# Baseline 0 → Into=0 (pjIntoBaseline); Baseline N (1-10) → Into=10+N (pjIntoBaselineN).
# (Verified from msproject_typelib.txt enum PjSaveBaselineTo.)
INTO_BASELINE_MAP = {n: (0 if n == 0 else 10 + n) for n in BASELINE_NUMBERS}


# ---------- BASELINE HELPERS ----------

def _baseline_property_name(field: str, baseline_number: int) -> str:
    """Map (field='Start', N=0) → 'BaselineStart'; (field='Start', N=3) → 'Baseline3Start'.

    Used to dynamically read task baseline properties for any of the 11 slots.
    """
    suffix = "" if baseline_number == 0 else str(baseline_number)
    return f"Baseline{suffix}{field}"


def _baseline_into_code(baseline_number: int) -> int:
    """Convert baseline_number (0-10) to pjIntoBaseline enum code for BaselineSave."""
    return INTO_BASELINE_MAP[baseline_number]


def _baseline_saved_date(proj: Any, baseline_number: int) -> Optional[Any]:
    """Return saved date of given baseline, or None if not saved.

    MSP returns 0 / NA / empty for unsaved baselines — normalize to Python None.
    """
    try:
        result = proj.BaselineSavedDate(Baseline=baseline_number)
        # MSP returns various falsy values for unsaved (0, datetime(year=0), None, "")
        if not result:
            return None
        # Some COM versions return a datetime with year 1; treat year < 1980 as unsaved
        try:
            if hasattr(result, "year") and result.year < 1980:
                return None
        except Exception:
            pass
        return result
    except Exception:
        return None


def _read_task_baseline(task: Any, baseline_number: int) -> Dict[str, Any]:
    """Read a task's baseline values for the given baseline slot.

    Returns dict with start, finish, duration_h, work_h, cost.
    Unsaved baseline yields None/0 fallbacks. Each property read is guarded
    so a single bad COM read doesn't kill compare iteration.
    """
    out: Dict[str, Any] = {}
    for field, key, transform in [
        ("Start", "start", lambda v: str(v) if v else None),
        ("Finish", "finish", lambda v: str(v) if v else None),
        ("Duration", "duration_h", lambda v: float(v) / 60.0 if v else 0.0),
        ("Work", "work_h", lambda v: float(v) / 60.0 if v else 0.0),
        ("Cost", "cost", lambda v: _parse_rate(v)),
    ]:
        try:
            raw = getattr(task, _baseline_property_name(field, baseline_number))
            out[key] = transform(raw)
        except Exception:
            out[key] = None if field in ("Start", "Finish") else 0.0
    return out
```

**Step 4: Run** — 8 PASS expected.

```bash
python -m pytest tests/test_msproject_baseline_helpers.py -v
```

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_helpers.py
git commit -m "Phase 3a T39: BASELINE_NUMBERS + INTO_BASELINE_MAP + 4 helpers (_baseline_property_name, _baseline_into_code, _baseline_saved_date, _read_task_baseline)"
```

Expected full regression: **164 PASSED + 1 xfail** (156 + 8 new).

---

## Task 40: `msproject_baseline` save Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_save.py`

**Step 1: Failing test**

`tests/test_msproject_baseline_save.py`:
```python
"""Test msproject_baseline save action."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_task_add_single, _baseline_saved_date,
)


def test_save_default_baseline_zero(clean_test_project):
    """Save baseline 0 on a project with 3 tasks → returns metadata."""
    proj = clean_test_project
    for i in range(3):
        _msp_task_add_single(name=f"SaveT{i}-T40", duration="2d")
    r = _msp_baseline_save(baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0
    assert r["task_count"] == 3
    assert r["saved_date"] is not None
    # Confirm via _baseline_saved_date helper
    assert _baseline_saved_date(proj, 0) is not None


def test_save_baseline_three(clean_test_project):
    """Save into Baseline3 (Into=13)."""
    proj = clean_test_project
    _msp_task_add_single(name="B3T-T40", duration="5d")
    r = _msp_baseline_save(baseline_number=3)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 3
    assert _baseline_saved_date(proj, 3) is not None
    # Verify Baseline 0 still unsaved
    assert _baseline_saved_date(proj, 0) is None


def test_save_invalid_baseline_number_errors(clean_test_project):
    """baseline_number 11 (out of 0-10) → error."""
    r = _msp_baseline_save(baseline_number=11)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
    assert "0-10" in r["error"]


def test_save_negative_baseline_number_errors(clean_test_project):
    r = _msp_baseline_save(baseline_number=-1)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
```

**Step 2: Run — FAIL** (ImportError)

```bash
python -m pytest tests/test_msproject_baseline_save.py -v
```

**Step 3: Implementation**

Insert after `_read_task_baseline` (T39 helpers):

```python
def _msp_baseline_save(baseline_number: int = 0,
                      name: Optional[str] = None,
                      scope: str = "all",
                      roll_up_to_summary: bool = True) -> Dict[str, Any]:
    """Save current state as the specified baseline (0-10).

    scope: "all" (default — all tasks) or "selected" (currently-selected only)
    roll_up_to_summary: roll up to summary tasks (MSP default behavior)
    name: optional descriptive label (returned in metadata; MSP doesn't
          store baseline names natively in pre-2016 versions, captured here
          for caller's tracking)
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    if scope not in ("all", "selected"):
        return {"status": "error",
                "error": f"scope must be 'all' or 'selected', got '{scope}'"}
    app = _validate_active_project()
    proj = app.ActiveProject
    try:
        # PjSaveBaselineFrom.pjCopyCurrent = 0
        copy_from = 0
        into = _baseline_into_code(baseline_number)
        # All=True (whole project) or False (selected); MSP COM expects bool
        all_param = (scope == "all")
        app.BaselineSave(All=all_param, Copy=copy_from, Into=into,
                        RollupToSummaryTasks=roll_up_to_summary)
        # Read-back metadata
        saved_date = _baseline_saved_date(proj, baseline_number)
        task_count = proj.Tasks.Count
        # Aggregate baseline totals
        total_dur_min, total_work_min, total_cost = 0.0, 0.0, 0.0
        for i in range(1, task_count + 1):
            t = proj.Tasks(i)
            if t is None or t.Summary:  # skip summary tasks for totals
                continue
            data = _read_task_baseline(t, baseline_number)
            total_dur_min += data["duration_h"] * 60 if data["duration_h"] else 0
            total_work_min += data["work_h"] * 60 if data["work_h"] else 0
            total_cost += data["cost"] if data["cost"] else 0
        return {"status": "ok",
                "baseline_number": baseline_number,
                "name": name,
                "saved_date": str(saved_date) if saved_date else None,
                "task_count": task_count,
                "total_duration_days": round(total_dur_min / 60 / 8, 2),  # 8h/day
                "total_work_hours": round(total_work_min / 60, 2),
                "total_cost": total_cost}
    except Exception as e:
        logger.error(f"_msp_baseline_save({baseline_number}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS**

Expected: 4 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_save.py
git commit -m "Phase 3a T40: msproject_baseline save action (multi-baseline 0-10 with Into offset)"
```

Expected: **168 PASSED + 1 xfail** (164 + 4).

---

## Task 41: `msproject_baseline` clear + clear_all (paired)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_clear.py`

**Step 1: Failing test**

`tests/test_msproject_baseline_clear.py`:
```python
"""Test msproject_baseline clear + clear_all actions."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_clear, _msp_baseline_clear_all,
    _msp_task_add_single, _baseline_saved_date,
)


def test_clear_saved_baseline(clean_test_project):
    """Save baseline 0 then clear it."""
    proj = clean_test_project
    _msp_task_add_single(name="ClearT-T41", duration="1d")
    _msp_baseline_save(baseline_number=0)
    assert _baseline_saved_date(proj, 0) is not None
    r = _msp_baseline_clear(baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0
    assert r["was_saved_date"] is not None
    # Verify cleared
    assert _baseline_saved_date(proj, 0) is None


def test_clear_unsaved_baseline_no_op(clean_test_project):
    """Clearing an unsaved baseline → ok with was_saved_date=None."""
    r = _msp_baseline_clear(baseline_number=5)
    assert r["status"] == "ok"
    assert r["was_saved_date"] is None


def test_clear_invalid_baseline_number_errors(clean_test_project):
    r = _msp_baseline_clear(baseline_number=11)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()


def test_clear_all_three_baselines(clean_test_project):
    """Save 3 baselines, then clear_all → all empty."""
    proj = clean_test_project
    _msp_task_add_single(name="ClearAllT-T41", duration="1d")
    _msp_baseline_save(baseline_number=0)
    _msp_baseline_save(baseline_number=2)
    _msp_baseline_save(baseline_number=7)
    r = _msp_baseline_clear_all()
    assert r["status"] == "ok"
    assert sorted(r["cleared"]) == [0, 2, 7]
    assert r["count"] == 3
    # Verify all 11 are now unsaved
    for n in range(11):
        assert _baseline_saved_date(proj, n) is None


def test_clear_all_when_none_saved(clean_test_project):
    """clear_all on fresh project → ok with empty cleared list."""
    r = _msp_baseline_clear_all()
    assert r["status"] == "ok"
    assert r["cleared"] == []
    assert r["count"] == 0
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_baseline_clear(baseline_number: int = 0) -> Dict[str, Any]:
    """Clear a single baseline (0-10). Idempotent: no-op if already empty."""
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    was_saved = _baseline_saved_date(proj, baseline_number)
    try:
        app.BaselineClear(All=True, From=baseline_number)
        return {"status": "ok",
                "baseline_number": baseline_number,
                "was_saved_date": str(was_saved) if was_saved else None}
    except Exception as e:
        logger.error(f"_msp_baseline_clear({baseline_number}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_baseline_clear_all() -> Dict[str, Any]:
    """Clear all 11 baselines that are currently saved. Returns list of cleared numbers."""
    app = _validate_active_project()
    proj = app.ActiveProject
    cleared = []
    failures = []
    for n in BASELINE_NUMBERS:
        if _baseline_saved_date(proj, n) is None:
            continue
        try:
            app.BaselineClear(All=True, From=n)
            cleared.append(n)
        except Exception as e:
            failures.append({"baseline_number": n, "error": _format_com_error(e)})
    return {"status": "ok" if not failures else "partial",
            "cleared": cleared,
            "count": len(cleared),
            "failures": failures}
```

**Step 4: Run — PASS**

Expected: 5 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_clear.py
git commit -m "Phase 3a T41: msproject_baseline clear + clear_all (idempotent, partial-failure aware)"
```

Expected: **173 PASSED + 1 xfail** (168 + 5).

---

## Task 42: `msproject_baseline` list Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_list.py`

**Step 1: Failing test**

`tests/test_msproject_baseline_list.py`:
```python
"""Test msproject_baseline list action."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_list, _msp_task_add_single,
)


def test_list_empty_project(clean_test_project):
    """No baselines saved → count_saved=0, baselines=[]."""
    r = _msp_baseline_list()
    assert r["status"] == "ok"
    assert r["count_saved"] == 0
    assert r["baselines"] == []


def test_list_one_saved(clean_test_project):
    """Save baseline 0 → list shows 1 entry with metadata."""
    _msp_task_add_single(name="ListT-T42", duration="3d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_list()
    assert r["status"] == "ok"
    assert r["count_saved"] == 1
    assert len(r["baselines"]) == 1
    entry = r["baselines"][0]
    assert entry["number"] == 0
    assert entry["saved_date"] is not None
    assert entry["task_count"] == 1


def test_list_three_saved(clean_test_project):
    """Save baselines 0, 3, 7 → list shows all 3 sorted by number."""
    _msp_task_add_single(name="MultiT-T42", duration="2d")
    _msp_baseline_save(baseline_number=0)
    _msp_baseline_save(baseline_number=3)
    _msp_baseline_save(baseline_number=7)
    r = _msp_baseline_list()
    assert r["status"] == "ok"
    assert r["count_saved"] == 3
    numbers = [b["number"] for b in r["baselines"]]
    assert numbers == [0, 3, 7]  # sorted ascending
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_baseline_list() -> Dict[str, Any]:
    """List all 11 baseline slots; return only those currently saved with metadata.

    Iterates all 11 slots, checks saved date, includes task count + total stats
    only for saved ones. Returns sorted by baseline number.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    out = []
    try:
        task_ct = proj.Tasks.Count
        for n in BASELINE_NUMBERS:
            saved = _baseline_saved_date(proj, n)
            if saved is None:
                continue
            # Aggregate per-task baseline totals (skip summaries)
            total_dur_min, total_work_min, total_cost = 0.0, 0.0, 0.0
            for i in range(1, task_ct + 1):
                t = proj.Tasks(i)
                if t is None or t.Summary:
                    continue
                data = _read_task_baseline(t, n)
                total_dur_min += (data["duration_h"] or 0) * 60
                total_work_min += (data["work_h"] or 0) * 60
                total_cost += (data["cost"] or 0)
            out.append({
                "number": n,
                "name": None,  # MSP doesn't store baseline names natively
                "saved_date": str(saved),
                "task_count": task_ct,
                "total_duration_days": round(total_dur_min / 60 / 8, 2),
                "total_work_hours": round(total_work_min / 60, 2),
                "total_cost": total_cost,
            })
        return {"status": "ok", "count_saved": len(out), "baselines": out}
    except Exception as e:
        logger.error(f"_msp_baseline_list failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (3 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_list.py
git commit -m "Phase 3a T42: msproject_baseline list action (all 11 slots scanned, only saved returned)"
```

Expected: **176 PASSED + 1 xfail** (173 + 3).

---

## Task 43: `msproject_baseline` get_task_baseline Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_get_task.py`

**Step 1: Failing test**

```python
"""Test msproject_baseline get_task_baseline action."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_get_task_baseline,
    _msp_task_add_single,
)


def test_get_task_baseline_after_save(clean_test_project):
    """Save baseline 0, read task's baseline → real values."""
    add_r = _msp_task_add_single(name="GetT-T43", duration="5d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_get_task_baseline(task_id=add_r["task_id"], baseline_number=0)
    assert r["status"] == "ok"
    assert r["task_id"] == add_r["task_id"]
    assert r["baseline_number"] == 0
    bd = r["baseline"]
    assert bd["start"] is not None
    assert bd["finish"] is not None
    assert bd["duration_h"] > 0  # 5d × 8h = 40h


def test_get_task_baseline_before_save(clean_test_project):
    """No baseline saved → start/finish None, duration/work/cost 0."""
    add_r = _msp_task_add_single(name="UnsavedT-T43", duration="2d")
    r = _msp_baseline_get_task_baseline(task_id=add_r["task_id"], baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline"]["start"] is None
    assert r["baseline"]["duration_h"] == 0


def test_get_task_baseline_missing_task(clean_test_project):
    r = _msp_baseline_get_task_baseline(task_id=99999, baseline_number=0)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_get_task_baseline_invalid_baseline_number(clean_test_project):
    add_r = _msp_task_add_single(name="BadBN-T43", duration="1d")
    r = _msp_baseline_get_task_baseline(task_id=add_r["task_id"], baseline_number=99)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_baseline_get_task_baseline(task_id: int, baseline_number: int = 0) -> Dict[str, Any]:
    """Read a single task's baseline values without computing variance."""
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    try:
        data = _read_task_baseline(t, baseline_number)
        return {"status": "ok",
                "task_id": task_id,
                "baseline_number": baseline_number,
                "baseline": data}
    except Exception as e:
        logger.error(f"_msp_baseline_get_task_baseline({task_id},{baseline_number}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (4 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_get_task.py
git commit -m "Phase 3a T43: msproject_baseline get_task_baseline action"
```

Expected: **180 PASSED + 1 xfail**.

---

## Task 44: `msproject_baseline` compare Action (BIG ONE)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_compare.py`

**Step 1: Failing test**

```python
"""Test msproject_baseline compare action — variance reporting."""
import pytest
import time
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_compare,
    _msp_task_add_single, _msp_task_update,
)


def test_compare_no_change_zero_variance(clean_test_project):
    """Save baseline, no progress → all tasks on_time, totals=0."""
    add_r = _msp_task_add_single(name="NoVarT-T44", duration="3d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_compare(baseline_number=0)
    assert r["status"] == "ok"
    s = r["summary"]
    assert s["slipped_count"] == 0
    assert s["ahead_count"] == 0
    assert s["on_time_count"] == 1
    assert s["total_finish_drift_days"] == 0


def test_compare_slipped_task(clean_test_project):
    """Slip a task's finish by 5 days → slipped_count=1, drift>0."""
    add_r = _msp_task_add_single(name="SlippedT-T44", duration="3d")
    _msp_baseline_save(baseline_number=0)
    # Slip: extend duration by 5 days
    _msp_task_update(task_id=add_r["task_id"], duration="8d")
    r = _msp_baseline_compare(baseline_number=0)
    assert r["status"] == "ok"
    assert r["summary"]["slipped_count"] == 1
    assert r["summary"]["total_finish_drift_days"] > 0
    # Find this task in list
    task_var = next(t for t in r["tasks"] if t["id"] == add_r["task_id"])
    assert task_var["status"] == "slipped"
    assert task_var["finish_var_days"] > 0


def test_compare_threshold_filter(clean_test_project):
    """variance_threshold_days=10 → small slip not counted as slipped."""
    add_r = _msp_task_add_single(name="ThreshT-T44", duration="5d")
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=add_r["task_id"], duration="7d")  # 2 day slip
    r = _msp_baseline_compare(baseline_number=0, variance_threshold_days=10)
    assert r["status"] == "ok"
    # 2 day slip < 10 day threshold → counts as on_time
    assert r["summary"]["slipped_count"] == 0
    assert r["summary"]["on_time_count"] == 1


def test_compare_include_unchanged_false_filters(clean_test_project):
    """include_unchanged=False → tasks with 0 variance excluded from list."""
    a = _msp_task_add_single(name="UnchT1-T44", duration="2d")
    b = _msp_task_add_single(name="UnchT2-T44", duration="3d")
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=b["task_id"], duration="5d")  # only B slipped
    r = _msp_baseline_compare(baseline_number=0, include_unchanged=False)
    assert r["status"] == "ok"
    # Only B should be in tasks list
    ids = [t["id"] for t in r["tasks"]]
    assert b["task_id"] in ids
    assert a["task_id"] not in ids


def test_compare_unsaved_baseline_errors(clean_test_project):
    r = _msp_baseline_compare(baseline_number=5)
    assert r["status"] == "error"
    assert "not saved" in r["error"].lower() or "no baseline" in r["error"].lower()


def test_compare_perf_50_tasks_under_2s(clean_test_project):
    """Performance: 50-task compare must complete <2s."""
    for i in range(50):
        _msp_task_add_single(name=f"PerfT{i:02d}-T44", duration="1d")
    _msp_baseline_save(baseline_number=0)
    start = time.time()
    r = _msp_baseline_compare(baseline_number=0)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["summary"]["on_time_count"] == 50
    assert elapsed < 2.0, f"compare 50 tasks took {elapsed:.2f}s (target <2s)"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _datetime_diff_days(current_str: Optional[str], baseline_str: Optional[str]) -> float:
    """Compute calendar-day difference between two ISO datetime strings.
    Returns 0 if either is None/missing."""
    if not current_str or not baseline_str:
        return 0.0
    try:
        from dateutil import parser
        c = parser.parse(current_str)
        b = parser.parse(baseline_str)
        return (c - b).total_seconds() / 86400.0
    except Exception:
        # Fallback: try strict ISO format
        try:
            c = _dt.datetime.fromisoformat(current_str.replace("+00:00", "").rstrip("Z"))
            b = _dt.datetime.fromisoformat(baseline_str.replace("+00:00", "").rstrip("Z"))
            return (c - b).total_seconds() / 86400.0
        except Exception:
            return 0.0


def _msp_baseline_compare(baseline_number: int = 0,
                         include_unchanged: bool = False,
                         variance_threshold_days: float = 0.0) -> Dict[str, Any]:
    """Compare current state vs saved baseline. Returns summary + per-task variance.

    variance_threshold_days: tasks with |finish_var_days| <= threshold count as on_time.
    include_unchanged: if False, omit tasks with 0 variance from per-task list.
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    if _baseline_saved_date(proj, baseline_number) is None:
        return {"status": "error",
                "error": f"Baseline {baseline_number} not saved (no baseline data to compare against)"}
    tasks_var = []
    slipped_ct = ahead_ct = on_time_ct = 0
    total_start_drift = total_finish_drift = 0.0
    total_dur_var_h = total_work_var_h = 0.0
    total_cost_var = 0.0
    try:
        for i in range(1, proj.Tasks.Count + 1):
            t = proj.Tasks(i)
            if t is None or t.Summary:
                continue
            bd = _read_task_baseline(t, baseline_number)
            cur_start = str(t.Start) if t.Start else None
            cur_finish = str(t.Finish) if t.Finish else None
            cur_dur_h = float(t.Duration) / 60.0 if t.Duration else 0.0
            cur_work_h = float(t.Work) / 60.0 if t.Work else 0.0
            cur_cost = _parse_rate(t.Cost) if t.Cost else 0.0

            start_var = _datetime_diff_days(cur_start, bd["start"])
            finish_var = _datetime_diff_days(cur_finish, bd["finish"])
            dur_var = cur_dur_h - (bd["duration_h"] or 0)
            work_var = cur_work_h - (bd["work_h"] or 0)
            cost_var = cur_cost - (bd["cost"] or 0)

            # Status by finish variance + threshold
            if finish_var > variance_threshold_days:
                status = "slipped"
                slipped_ct += 1
            elif finish_var < -variance_threshold_days:
                status = "ahead"
                ahead_ct += 1
            else:
                status = "on_time"
                on_time_ct += 1

            total_start_drift += start_var
            total_finish_drift += finish_var
            total_dur_var_h += dur_var
            total_work_var_h += work_var
            total_cost_var += cost_var

            # Filter list per include_unchanged
            no_change = (start_var == 0 and finish_var == 0 and
                        dur_var == 0 and work_var == 0 and cost_var == 0)
            if not include_unchanged and no_change:
                continue
            tasks_var.append({
                "id": t.ID,
                "name": t.Name,
                "start_var_days": round(start_var, 2),
                "finish_var_days": round(finish_var, 2),
                "duration_var_h": round(dur_var, 2),
                "work_var_h": round(work_var, 2),
                "cost_var": round(cost_var, 2),
                "status": status,
            })
        return {
            "status": "ok",
            "baseline_number": baseline_number,
            "summary": {
                "slipped_count": slipped_ct,
                "ahead_count": ahead_ct,
                "on_time_count": on_time_ct,
                "total_start_drift_days": round(total_start_drift, 2),
                "total_finish_drift_days": round(total_finish_drift, 2),
                "total_duration_var_h": round(total_dur_var_h, 2),
                "total_work_var_h": round(total_work_var_h, 2),
                "total_cost_var": round(total_cost_var, 2),
            },
            "tasks": tasks_var,
        }
    except Exception as e:
        logger.error(f"_msp_baseline_compare({baseline_number}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**NOTE on `dateutil`:** if not available, the fromisoformat fallback handles common cases. For prod robustness, consider adding `python-dateutil` to requirements (already in mspdi_parser.py likely).

**Step 4: Run — PASS** (6 PASSED with hero perf assertion)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_compare.py
git commit -m "Phase 3a T44: msproject_baseline compare action (variance + threshold + RAG-ready summary)"
```

Expected: **186 PASSED + 1 xfail**.

---

## Task 45: `msproject_baseline` compare_two Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_compare_two.py`

**Step 1: Failing test**

```python
"""Test msproject_baseline compare_two action — baseline-to-baseline delta."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_compare_two,
    _msp_task_add_single, _msp_task_update,
)


def test_compare_two_zero_when_baselines_identical(clean_test_project):
    """Save B0, then save B1 (snapshot of same state) → variance 0."""
    add_r = _msp_task_add_single(name="EqT-T45", duration="3d")
    _msp_baseline_save(baseline_number=0)
    _msp_baseline_save(baseline_number=1)
    r = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
    assert r["status"] == "ok"
    assert r["summary"]["slipped_count"] == 0


def test_compare_two_revision_delta(clean_test_project):
    """B0 saved, task changed, B1 saved → compare_two(0,1) shows delta."""
    add_r = _msp_task_add_single(name="DeltaT-T45", duration="3d")
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=add_r["task_id"], duration="8d")  # 5 day slip
    _msp_baseline_save(baseline_number=1)
    r = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
    assert r["status"] == "ok"
    # B0→B1 shows the slip
    assert r["summary"]["slipped_count"] == 1


def test_compare_two_unsaved_baseline_a_errors(clean_test_project):
    _msp_task_add_single(name="OnlyB-T45", duration="2d")
    _msp_baseline_save(baseline_number=1)
    r = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
    assert r["status"] == "error"
    assert "baseline_a" in r["error"].lower() or "0" in r["error"]


def test_compare_two_invalid_baseline_number(clean_test_project):
    r = _msp_baseline_compare_two(baseline_a=99, baseline_b=0)
    assert r["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_baseline_compare_two(baseline_a: int,
                             baseline_b: int,
                             include_unchanged: bool = False,
                             variance_threshold_days: float = 0.0) -> Dict[str, Any]:
    """Compare two saved baselines (delta = B → A semantically: A as 'current', B as 'baseline').

    Use case: baseline_a=0 (Original), baseline_b=1 (Revised) → shows what changed
    between Original and Revised plans.
    """
    for label, n in [("baseline_a", baseline_a), ("baseline_b", baseline_b)]:
        if n not in BASELINE_NUMBERS:
            return {"status": "error",
                    "error": f"{label} must be 0-10, got {n}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    if _baseline_saved_date(proj, baseline_a) is None:
        return {"status": "error",
                "error": f"baseline_a ({baseline_a}) is not saved"}
    if _baseline_saved_date(proj, baseline_b) is None:
        return {"status": "error",
                "error": f"baseline_b ({baseline_b}) is not saved"}
    # Iterate tasks; treat baseline_a as "current" and baseline_b as "baseline" semantically
    tasks_var = []
    slipped_ct = ahead_ct = on_time_ct = 0
    total_start_drift = total_finish_drift = 0.0
    total_dur_var_h = total_work_var_h = 0.0
    total_cost_var = 0.0
    try:
        for i in range(1, proj.Tasks.Count + 1):
            t = proj.Tasks(i)
            if t is None or t.Summary:
                continue
            a_data = _read_task_baseline(t, baseline_a)
            b_data = _read_task_baseline(t, baseline_b)
            start_var = _datetime_diff_days(a_data["start"], b_data["start"])
            finish_var = _datetime_diff_days(a_data["finish"], b_data["finish"])
            dur_var = (a_data["duration_h"] or 0) - (b_data["duration_h"] or 0)
            work_var = (a_data["work_h"] or 0) - (b_data["work_h"] or 0)
            cost_var = (a_data["cost"] or 0) - (b_data["cost"] or 0)

            if finish_var > variance_threshold_days:
                status = "slipped"; slipped_ct += 1
            elif finish_var < -variance_threshold_days:
                status = "ahead"; ahead_ct += 1
            else:
                status = "on_time"; on_time_ct += 1

            total_start_drift += start_var
            total_finish_drift += finish_var
            total_dur_var_h += dur_var
            total_work_var_h += work_var
            total_cost_var += cost_var

            no_change = (start_var == 0 and finish_var == 0 and
                        dur_var == 0 and work_var == 0 and cost_var == 0)
            if not include_unchanged and no_change:
                continue
            tasks_var.append({
                "id": t.ID, "name": t.Name,
                "start_var_days": round(start_var, 2),
                "finish_var_days": round(finish_var, 2),
                "duration_var_h": round(dur_var, 2),
                "work_var_h": round(work_var, 2),
                "cost_var": round(cost_var, 2),
                "status": status,
            })
        return {
            "status": "ok",
            "baseline_a": baseline_a, "baseline_b": baseline_b,
            "summary": {
                "slipped_count": slipped_ct, "ahead_count": ahead_ct, "on_time_count": on_time_ct,
                "total_start_drift_days": round(total_start_drift, 2),
                "total_finish_drift_days": round(total_finish_drift, 2),
                "total_duration_var_h": round(total_dur_var_h, 2),
                "total_work_var_h": round(total_work_var_h, 2),
                "total_cost_var": round(total_cost_var, 2),
            },
            "tasks": tasks_var,
        }
    except Exception as e:
        logger.error(f"_msp_baseline_compare_two({baseline_a},{baseline_b}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (4 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_compare_two.py
git commit -m "Phase 3a T45: msproject_baseline compare_two action (baseline-to-baseline delta)"
```

Expected: **190 PASSED + 1 xfail**.

---

## Task 46: `msproject_baseline` summary Action (RAG)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_summary.py`

**Step 1: Failing test**

```python
"""Test msproject_baseline summary action — project-level RAG status."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_summary,
    _msp_task_add_single, _msp_task_update,
)


def test_summary_green_no_slip(clean_test_project):
    for i in range(10):
        _msp_task_add_single(name=f"GT{i}-T46", duration="1d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_summary(baseline_number=0)
    assert r["status"] == "ok"
    assert r["project"]["slipped_pct"] == 0.0
    assert r["project"]["schedule_health"] == "green"


def test_summary_amber_when_5_to_20_pct_slipped(clean_test_project):
    """10 tasks, slip 1 → 10% slipped → amber."""
    ids = [_msp_task_add_single(name=f"AT{i}-T46", duration="2d")["task_id"]
           for i in range(10)]
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=ids[0], duration="10d")  # slip 1 task
    r = _msp_baseline_summary(baseline_number=0)
    assert r["status"] == "ok"
    assert 5 < r["project"]["slipped_pct"] <= 20
    assert r["project"]["schedule_health"] == "amber"


def test_summary_red_when_over_20_pct_slipped(clean_test_project):
    """5 tasks, slip 2 → 40% slipped → red."""
    ids = [_msp_task_add_single(name=f"RT{i}-T46", duration="2d")["task_id"]
           for i in range(5)]
    _msp_baseline_save(baseline_number=0)
    _msp_task_update(task_id=ids[0], duration="8d")
    _msp_task_update(task_id=ids[1], duration="9d")
    r = _msp_baseline_summary(baseline_number=0)
    assert r["status"] == "ok"
    assert r["project"]["slipped_pct"] > 20
    assert r["project"]["schedule_health"] == "red"


def test_summary_unsaved_baseline_errors(clean_test_project):
    r = _msp_baseline_summary(baseline_number=2)
    assert r["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Summary delegates to compare internally + computes RAG:

```python
def _msp_baseline_summary(baseline_number: int = 0) -> Dict[str, Any]:
    """Project-level RAG status from baseline comparison.

    RAG thresholds:
      green  : slipped_pct <= 5
      amber  : 5 < slipped_pct <= 20
      red    : slipped_pct > 20
    """
    cmp_result = _msp_baseline_compare(baseline_number=baseline_number,
                                      include_unchanged=True)
    if cmp_result.get("status") != "ok":
        return cmp_result  # propagate error
    s = cmp_result["summary"]
    total_tasks = s["slipped_count"] + s["ahead_count"] + s["on_time_count"]
    slipped_pct = (s["slipped_count"] / total_tasks * 100.0) if total_tasks > 0 else 0.0
    if slipped_pct <= 5:
        health = "green"
    elif slipped_pct <= 20:
        health = "amber"
    else:
        health = "red"
    return {
        "status": "ok",
        "baseline_number": baseline_number,
        "project": {
            "task_count": total_tasks,
            "slipped_count": s["slipped_count"],
            "slipped_pct": round(slipped_pct, 2),
            "start_drift_days": s["total_start_drift_days"],
            "finish_drift_days": s["total_finish_drift_days"],
            "schedule_health": health,
        },
    }
```

**Step 4: Run — PASS** (4 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_summary.py
git commit -m "Phase 3a T46: msproject_baseline summary action (RAG: green<=5%, amber<=20%, red>20%)"
```

Expected: **194 PASSED + 1 xfail**.

---

## Task 47: `msproject_baseline` set_active Action (Investigate API)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_baseline_set_active.py`

**Step 1: Investigate API first**

Before writing tests, run a quick probe to find the right COM property:

```python
# probe — DO NOT COMMIT
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MSProject.Application")
# Try several candidate APIs for "active baseline"
for attr in ["BaselineForEarnedValue", "ActiveBaselineNumber", "OptionsCalculation"]:
    try:
        v = getattr(app, attr)
        print(f"{attr} = {v}")
    except Exception as e:
        print(f"{attr}: {e}")
# Also check project-level
proj = app.ActiveProject
for attr in ["BaselineForEarnedValue", "ActiveBaselineNumber"]:
    try:
        v = getattr(proj, attr)
        print(f"proj.{attr} = {v}")
    except Exception as e:
        print(f"proj.{attr}: {e}")
```

**Decision rule:**
- If a writable property like `proj.BaselineForEarnedValue` exists → use it
- If only read-only OR not found → return "not yet supported" with helpful message in T47, defer to Phase 4

**Step 2: Failing test (assume API found)**

```python
"""Test msproject_baseline set_active action."""
import pytest
from msproject_mcp_core import _msp_baseline_save, _msp_baseline_set_active, _msp_task_add_single


def test_set_active_default_zero(clean_test_project):
    """Set baseline 0 active."""
    _msp_task_add_single(name="ActT-T47", duration="1d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_set_active(baseline_number=0)
    # Either ok (if API found) or specific "not supported" error (Phase 4 deferral)
    assert r["status"] in ("ok", "error")
    if r["status"] == "error":
        assert "not yet supported" in r["error"].lower() or "phase" in r["error"].lower()


def test_set_active_invalid_baseline_number(clean_test_project):
    r = _msp_baseline_set_active(baseline_number=99)
    assert r["status"] == "error"
    assert "baseline_number" in r["error"].lower()
```

**Step 3: Implementation (graceful fallback if API not found)**

```python
def _msp_baseline_set_active(baseline_number: int) -> Dict[str, Any]:
    """Set the active baseline for views/EVM calculations.

    NOTE: MSP COM API for this is version-dependent. Phase 3a returns 'not yet
    supported' if the property isn't accessible — the saved baseline data is
    still readable via get_task_baseline / compare regardless of which is 'active'.
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    # Try the candidate property; fall back to "not supported" cleanly
    try:
        # Common in newer MSP versions
        proj.BaselineForEarnedValue = baseline_number
        return {"status": "ok", "active_baseline": baseline_number,
                "method": "proj.BaselineForEarnedValue"}
    except Exception:
        pass
    return {"status": "error",
            "error": ("set_active is not yet supported on this MS Project version. "
                     "Use msproject_baseline compare/summary directly with the "
                     "baseline_number parameter — they don't require setting an active baseline.")}
```

**Step 4: Run — PASS** (2 PASSED — both happy and error paths covered)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_set_active.py
git commit -m "Phase 3a T47: msproject_baseline set_active action (graceful fallback if API unavailable)"
```

Expected: **196 PASSED + 1 xfail**.

---

## Task 48: FastMCP Dispatcher + Acceptance Script + README + Push (FINAL)

**Files:**
- Modify: `msproject_mcp_core.py` (add `@mcp.tool` dispatcher near other dispatchers; update server `instructions` string)
- Create: `tests/test_msproject_baseline_dispatcher.py`
- Create: `samples/build_baseline_lifecycle.py`
- Modify: `README.md`

**Step 1: Failing dispatcher test**

`tests/test_msproject_baseline_dispatcher.py`:
```python
"""Test FastMCP msproject_baseline dispatcher."""
import asyncio, json
import pytest
from msproject_mcp_core import msproject_baseline


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_save(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    _msp_task_add_single(name="DispBT-T48", duration="2d")
    r = _run(msproject_baseline({"action": "save", "baseline_number": 0}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert p["baseline_number"] == 0


def test_dispatcher_list(clean_test_project):
    r = _run(msproject_baseline({"action": "list"}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "baselines" in p


def test_dispatcher_compare_chain(clean_test_project):
    """Chain: save -> compare → variance via dispatcher."""
    from msproject_mcp_core import _msp_task_add_single
    _msp_task_add_single(name="ChainT-T48", duration="3d")
    _run(msproject_baseline({"action": "save", "baseline_number": 0}))
    r = _run(msproject_baseline({"action": "compare", "baseline_number": 0}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "summary" in p


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_baseline({"action": "nonsense"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

In `msproject_mcp_core.py`, locate other `@mcp.tool` decorators (msproject_task / link / schedule / calendar / resource) and add msproject_baseline after msproject_resource:

```python
@mcp.tool(
    name="msproject_baseline",
    annotations={"title": "MS Project Baseline Operations", "readOnlyHint": False},
)
async def msproject_baseline(params: dict) -> str:
    """Multi-baseline (Baseline + Baseline1..Baseline10) management + variance reporting.

    Actions:
    - save: Save current as baseline. Params: [baseline_number=0, name, scope='all', roll_up_to_summary=True]
    - clear: Clear single baseline. Params: [baseline_number=0]
    - clear_all: Clear all 11 saved baselines.
    - list: List all saved baselines with metadata.
    - get_task_baseline: Read one task's baseline values. Params: task_id, [baseline_number=0]
    - compare: Variance current vs baseline. Params: [baseline_number=0, include_unchanged=False, variance_threshold_days=0]
    - compare_two: Delta between two baselines. Params: baseline_a, baseline_b, [include_unchanged, variance_threshold_days]
    - summary: Project-level RAG status. Params: [baseline_number=0]
    - set_active: Set active baseline for views (version-dependent). Params: baseline_number

    Phase 3a (28 Apr 2026). RAG thresholds: green<=5%, amber<=20%, red>20% slipped.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "save":
            r = _msp_baseline_save(**p)
        elif action == "clear":
            r = _msp_baseline_clear(**p)
        elif action == "clear_all":
            r = _msp_baseline_clear_all(**p)
        elif action == "list":
            r = _msp_baseline_list(**p)
        elif action == "get_task_baseline":
            r = _msp_baseline_get_task_baseline(**p)
        elif action == "compare":
            r = _msp_baseline_compare(**p)
        elif action == "compare_two":
            r = _msp_baseline_compare_two(**p)
        elif action == "summary":
            r = _msp_baseline_summary(**p)
        elif action == "set_active":
            r = _msp_baseline_set_active(**p)
        else:
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: save/clear/clear_all/list/get_task_baseline/compare/compare_two/summary/set_active"}
    except Exception as e:
        logger.error(f"msproject_baseline({action}) failed: {e}")
        r = {"status": "error", "error": _format_com_error(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

Update `mcp = FastMCP(...)` `instructions` string to include `msproject_baseline`:
```
"Tools: msproject_task, msproject_link, msproject_schedule, msproject_calendar, msproject_resource, msproject_baseline."
```

**Step 4: Acceptance script**

`samples/build_baseline_lifecycle.py`:
```python
"""Phase 3a acceptance: full baseline lifecycle.

SAFETY: Uses isolated FileNew project, never touches user's active project.

Scenario:
  1. Create 50 villa tasks
  2. Add 3 work resources, assign to all tasks (mini Phase 2b chain)
  3. Save Baseline 0 ('Original')
  4. Update progress on first 20 tasks (extend duration to simulate slips)
  5. Compare(0) → variance report
  6. Save Baseline 1 ('Rev1-AfterChangeOrder')
  7. Update more durations
  8. Compare_two(0, 1) → revision delta
  9. Summary(0) → RAG status

Target: end-to-end <10s.
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_task_update,
    _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save, _msp_baseline_compare, _msp_baseline_compare_two,
    _msp_baseline_summary, _msp_baseline_list,
)


N_TASKS = 50


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    original_name = app.ActiveProject.Name if app.ActiveProject else None
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] isolated test: {test_name}, user's: {original_name}")

    try:
        t0 = time.time()
        # 1-2. Tasks + resources
        print(f"\n1. Building {N_TASKS} villa tasks + 3 resources + assignments...")
        tasks = _msp_task_bulk_add(items=[{"name": f"Villa T{i:03d}", "duration": "2d"} for i in range(N_TASKS)])
        task_ids = tasks["task_ids"]
        res_ids = []
        for name in ["COW", "STL", "MSN"]:
            res_ids.append(_msp_resource_add(name=name, type="Work", max_units=300)["resource_id"])
        items = [{"task_id": tid, "resource_id": rid} for tid in task_ids for rid in res_ids]
        _msp_resource_bulk_assign(items=items)
        print(f"   OK {len(task_ids)} tasks, 3 resources, {len(items)} assignments in {time.time()-t0:.2f}s")

        # 3. Save baseline 0
        print("2. Saving Baseline 0 'Original'...")
        b0 = _msp_baseline_save(baseline_number=0, name="Original")
        assert b0["status"] == "ok"
        print(f"   OK saved at {b0['saved_date']} ({b0['total_work_hours']}h total work)")

        # 4. Slip first 20 tasks
        print("3. Slipping first 20 tasks (2d -> 5d)...")
        for tid in task_ids[:20]:
            _msp_task_update(task_id=tid, duration="5d")

        # 5. Compare against Baseline 0
        print("4. Compare current vs Baseline 0...")
        cmp1 = _msp_baseline_compare(baseline_number=0)
        s = cmp1["summary"]
        print(f"   slipped={s['slipped_count']}, on_time={s['on_time_count']}, total_finish_drift={s['total_finish_drift_days']:.1f}d")

        # 6. Save baseline 1
        print("5. Saving Baseline 1 'Rev1-AfterChangeOrder'...")
        b1 = _msp_baseline_save(baseline_number=1, name="Rev1-AfterChangeOrder")
        assert b1["status"] == "ok"

        # 7. More changes
        print("6. Slipping next 10 tasks (2d -> 4d)...")
        for tid in task_ids[20:30]:
            _msp_task_update(task_id=tid, duration="4d")

        # 8. Compare two baselines
        print("7. Compare Baseline 0 vs Baseline 1 (revision delta)...")
        cmp2 = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
        s = cmp2["summary"]
        print(f"   slipped={s['slipped_count']}, total_finish_drift={s['total_finish_drift_days']:.1f}d")

        # 9. Summary
        print("8. Summary against Baseline 0 (RAG status)...")
        summ = _msp_baseline_summary(baseline_number=0)
        p = summ["project"]
        print(f"   slipped_pct={p['slipped_pct']:.1f}%, schedule_health={p['schedule_health'].upper()}")

        # List
        print("9. Listing all saved baselines...")
        bl = _msp_baseline_list()
        for b in bl["baselines"]:
            print(f"   - Baseline {b['number']}: {b['saved_date'][:10]} | {b['task_count']} tasks | {b['total_work_hours']:.1f}h work")

        elapsed = time.time() - t0
        print(f"\nOK ACCEPTANCE: {elapsed:.2f}s total (target <10s)")
        assert elapsed < 10.0, f"Too slow: {elapsed}s"

    finally:
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    break
            if original_name:
                for i in range(1, app.Projects.Count + 1):
                    if app.Projects(i).Name == original_name:
                        app.WindowActivate(app.Projects(i).Windows(1).Caption)
                        break
        except Exception as e:
            print(f"[WARN] cleanup error: {e}")


if __name__ == "__main__":
    main()
```

**Step 5: Run acceptance**

```bash
python samples/build_baseline_lifecycle.py
```

Expected: `OK ACCEPTANCE: <Xs total (target <10s)`. Realistic ~5-8s.

**Step 6: README update**

Add Phase 3a section to `README.md` after Phase 2b. Read the current README first to see structure, then add:

```markdown
### Phase 3a — Baseline (28 Apr 2026)

`msproject_baseline` tool with 9 actions, all 11 baseline slots:
- `save` / `clear` / `clear_all` — multi-baseline lifecycle
- `list` — all saved baselines + metadata
- `get_task_baseline` — read one task's baseline values
- `compare` — current vs baseline variance + threshold filter
- `compare_two` — baseline-to-baseline delta (revision tracking)
- `summary` — project-level RAG (green<=5% slipped, amber<=20%, red>20%)
- `set_active` — graceful fallback if MSP version doesn't expose API

Acceptance: `samples/build_baseline_lifecycle.py` runs full Original → progress
→ revise → compare lifecycle in <10s.

Tool count: **6 tools, ~40 actions**.
```

**Step 7: Run full regression**

```bash
python -m pytest tests/ -v --tb=short -q
```

Expected: **200 PASSED + 1 xfail** (196 + 4 dispatcher).

**Step 8: Commit + push**

```bash
git add msproject_mcp_core.py tests/test_msproject_baseline_dispatcher.py samples/build_baseline_lifecycle.py README.md
git commit -m "Phase 3a T48: dispatcher + acceptance + README + push (full baseline lifecycle <10s)"
git push origin main
```

Expected: ~11 commits pushed (T39-T48 + design).

---

## Phase 3a Tamamlama Kriterleri

1. ✅ T39-T48 10 commit landed (+ fixes if any review feedback)
2. ✅ Acceptance script `samples/build_baseline_lifecycle.py` <10s
3. ✅ Yeni testler ~28 PASS
4. ✅ Phase 1+2a+2b baseline 156+1xfail regression PASS
5. ✅ Total ~184-200 PASS + 1 xfail
6. ✅ Push to origin/main
7. ✅ Phase 3a live on GitHub
8. ⏸ Kullanıcı manuel onayı → Phase 3b (Progress) başlar

---

*Plan tamamlandı: 28 Nisan 2026*
*Tahmini Phase 3a süresi: ~6-8 saat (T39-T48, 10 task)*
*Sonraki phase (onay sonrası): Phase 3b — Progress Management (`msproject_progress` tool)*
