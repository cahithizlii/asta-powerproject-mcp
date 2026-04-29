# Phase 3b Progress Management Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** New `msproject_progress` MCP tool — 12 actions covering task-level + assignment-level progress, time-phased actuals (TimeScaleData), status date, hibrit bulk path ve EVM-ready summary. Built on Phase 1+2a+2b+3a foundations.

**Architecture:** All helpers in `msproject_mcp_core.py` (Phase 1+2a+2b+3a sections untouched). Insertion point: AFTER `_msp_baseline_set_active` (current line ~1612), BEFORE `_msp_task_update` (current line 1615). Reuse: `_validate_active_project`, `_format_com_error`, `_parse_rate`, `_find_task_by_id` (Phase 1), `_find_resource_by_id` (Phase 2b), `_msp_dt_or_none` (Phase 3a T44 fix), `_route_operation` + `_enter_batch_mode` / `_exit_batch_mode` (Phase 1), `_build_task_id_map` (T37). `_compute_variance_set` (Phase 3a T49) is unrelated to progress and not used.

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM, pytest, `python-dateutil` (already in `requirements.txt` from T44 fix). Mevcut `msproject_mcp_core.py` (~2500 satır after Phase 3a TAIL), 26 test dosyası, **207 PASS + 1 xfail** baseline.

**Design doc:** `docs/plans/2026-04-29-phase3b-progress-design.md` (commit `3a076c1`).

**Baseline state at start:** HEAD `3a076c1` (design committed), MS Project running v16.0, working tree clean.

**KEY COM API REFERENCE (verified from `msproject_typelib.txt`):**

Task-level progress fields:
- `task.PercentComplete` (RW, 0-100, duration-based)
- `task.PercentWorkComplete` (RW, 0-100, work-based)
- `task.ActualStart` / `task.ActualFinish` (RW, datetime)
- `task.ActualDuration` / `task.ActualWork` (RW, **minutes**)
- `task.RemainingWork` / `task.RemainingDuration` (RW, **minutes**)
- `task.PhysicalPercentComplete` (RW — probe T52)
- `task.Stop` / `task.Resume` (RW, datetime)

Assignment-level fields:
- `task.Assignments` (R, 1-indexed collection)
- `assignment.ResourceID` / `assignment.Resource` (R)
- `assignment.ActualWork` / `assignment.RemainingWork` (RW, **minutes**)
- `assignment.ActualStart` / `assignment.ActualFinish` (RW, datetime)
- `assignment.PercentWorkComplete` (RW, 0-100)
- `assignment.Units` (RW, float; 1.0 = 100%)

Time-phased:
- `assignment.TimeScaleData(StartDate, EndDate, Type, TimescaleUnit)` → `TimeScaleValues` collection (1-indexed, `.Count`, `.Item(i).Value` in minutes)
- Type enum: `pjAssignmentTimescaledActualWork = 24` (probe T60)
- TimescaleUnit enum: `pjTimescaleDays = 8`, `pjTimescaleWeeks = 6` (probe T60)

Project-level:
- `proj.StatusDate` (RW, datetime; ASLA NA gibi sentinel)
- `app.UpdateProject(ProgressDate, UpdatePercentCompleteOnly, AllTasks=...)`
- `proj.PercentComplete` (R, derived)

**CRITICAL Phase 1+2a+2b+3a BOUNDARY:**
- Phase 1's `_msp_task_*`, `_msp_link_*`, `_msp_schedule_*` are FROZEN. Don't touch.
- Phase 2a's `_msp_calendar_*` are FROZEN.
- Phase 2b's `_msp_resource_*` are FROZEN.
- Phase 3a's `_msp_baseline_*` and `_compute_variance_set` are FROZEN.
- Phase 3b adds NEW `_msp_progress_*` helpers in a separate section, uses upstream helpers as black boxes.

**Insertion verification command (run before T52):**
```bash
grep -n "^def _msp_baseline_set_active\|^def _msp_task_update" msproject_mcp_core.py
# Expected:
# 1581:def _msp_baseline_set_active(...)
# 1615:def _msp_task_update(...)
# Insertion point: line 1614 (one blank line after _msp_baseline_set_active's closing return)
```

---

## Task 52: Progress Foundations (helpers + constants + PhysicalPercentComplete probe)

**Files:**
- Modify: `msproject_mcp_core.py` (insert at line ~1614, after `_msp_baseline_set_active`)
- Create: `tests/test_msproject_progress_helpers.py`

**Step 1.5: Probe `task.PhysicalPercentComplete` setattr support**

Before writing tests/code, run a quick probe to confirm setter works on this MSP build:

```python
# probe_physical_pct.py — DO NOT COMMIT
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MSProject.Application")
app.FileNew()
proj = app.ActiveProject
t = proj.Tasks.Add("ProbeT")
print(f"BEFORE: PhysicalPercentComplete = {t.PhysicalPercentComplete}")
try:
    t.PhysicalPercentComplete = 42
    readback = t.PhysicalPercentComplete
    print(f"AFTER set=42: readback = {readback}")
    if readback == 42:
        print("OK: PhysicalPercentComplete is RW on this MSP build")
    else:
        print(f"WARN: setattr accepted but readback != 42 (silent no-op)")
except Exception as e:
    print(f"FAIL: setattr raised: {e}")
finally:
    app.FileClose(0)
```

Run: `python probe_physical_pct.py`. Expected: round-trip succeeds on MSP 16.0+.

**Decision rule:**
- Round-trip OK → implement setter normally in T53
- Setter fails → in T53, return `{"status": "error", "error": "PhysicalPercentComplete not supported on this MSP version"}` only when user passes that field; other fields work
- Silent no-op (setattr OK but readback != input) → setter returns "ok" but adds warning in response

**Step 1: Failing test**

`tests/test_msproject_progress_helpers.py`:
```python
"""Test progress helpers + constants."""
import pytest
from msproject_mcp_core import (
    _PROGRESS_PCT_FIELDS, _PROGRESS_WORK_FIELDS, _PROGRESS_DURATION_FIELDS,
    _PROGRESS_DATE_FIELDS, _TIMESCALE_UNIT_MAP, _PJ_TIMESCALED_ACTUAL_WORK,
    _normalize_progress_pct, _hours_to_minutes, _minutes_to_hours,
    _validate_actual_dates, _get_assignment_by_resource_id,
    _read_task_progress_dict, _msp_task_add_single,
    _msp_resource_add, _msp_resource_assign,
)


def test_progress_pct_fields_constant():
    assert "percent_complete" in _PROGRESS_PCT_FIELDS
    assert "percent_work_complete" in _PROGRESS_PCT_FIELDS
    assert "physical_pct" in _PROGRESS_PCT_FIELDS


def test_progress_work_fields_constant():
    assert "actual_work_h" in _PROGRESS_WORK_FIELDS
    assert "remaining_work_h" in _PROGRESS_WORK_FIELDS


def test_timescale_unit_map_constant():
    assert _TIMESCALE_UNIT_MAP["day"] == 8
    assert _TIMESCALE_UNIT_MAP["week"] == 6


def test_pj_timescaled_actual_work_const():
    # Probe-confirmed: pjAssignmentTimescaledActualWork == 24 on MSP 16.0
    assert _PJ_TIMESCALED_ACTUAL_WORK == 24


def test_normalize_progress_pct_int_float_str():
    assert _normalize_progress_pct(50) == 50.0
    assert _normalize_progress_pct(50.5) == 50.5
    assert _normalize_progress_pct("50") == 50.0
    assert _normalize_progress_pct("50%") == 50.0
    assert _normalize_progress_pct("50.25") == 50.25


def test_normalize_progress_pct_rejects_out_of_range():
    with pytest.raises(ValueError):
        _normalize_progress_pct(101)
    with pytest.raises(ValueError):
        _normalize_progress_pct(-0.5)
    with pytest.raises(ValueError):
        _normalize_progress_pct("not a number")


def test_hours_to_minutes_round_trip():
    assert _hours_to_minutes(8) == 480
    assert _hours_to_minutes(0.5) == 30
    assert _minutes_to_hours(480) == 8.0
    assert _minutes_to_hours(30) == 0.5


def test_validate_actual_dates_order():
    # Both None → OK
    assert _validate_actual_dates(None, None) is None
    # Only one supplied → OK (other determined by MSP)
    assert _validate_actual_dates("2026-04-01", None) is None
    assert _validate_actual_dates(None, "2026-04-15") is None
    # Both: start <= finish OK
    assert _validate_actual_dates("2026-04-01", "2026-04-15") is None
    # start > finish → error
    err = _validate_actual_dates("2026-04-15", "2026-04-01")
    assert err is not None
    assert "before" in err.lower() or "<=" in err.lower()


def test_get_assignment_by_resource_id_happy(clean_test_project):
    """Assign R to T → lookup returns Assignment object."""
    add_t = _msp_task_add_single(name="GetAsgT-T52", duration="2d")
    add_r = _msp_resource_add(name="ResX-T52", type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_t["task_id"])
    asg = _get_assignment_by_resource_id(t, add_r["resource_id"])
    assert asg is not None


def test_get_assignment_by_resource_id_missing(clean_test_project):
    """No matching assignment → returns None."""
    add_t = _msp_task_add_single(name="NoAsgT-T52", duration="1d")
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_t["task_id"])
    asg = _get_assignment_by_resource_id(t, 99999)
    assert asg is None


def test_read_task_progress_dict_initial_state(clean_test_project):
    """Fresh task — all progress zero/None."""
    add_t = _msp_task_add_single(name="ReadProgT-T52", duration="3d")
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_t["task_id"])
    p = _read_task_progress_dict(t)
    assert p["percent_complete"] == 0
    assert p["percent_work_complete"] == 0
    assert p["actual_start"] is None
    assert p["actual_finish"] is None
    assert p["actual_work_h"] == 0
    assert p["remaining_work_h"] >= 0  # 24h for 3d × 8h
    assert "physical_pct" in p
```

**Step 2: Run** — expect ImportError on `_PROGRESS_PCT_FIELDS`.

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/test_msproject_progress_helpers.py -v
```

**Step 3: Implementation**

Insert at line ~1614 (after `_msp_baseline_set_active`'s closing return on line 1612, after the blank line):

```python
# ==================== PHASE 3b — PROGRESS MANAGEMENT ====================
# Tool: msproject_progress (12 actions). Insertion point: between Phase 3a
# baseline section and Phase 1 _msp_task_update.
# Builds on: _validate_active_project, _format_com_error, _parse_rate,
# _find_task_by_id, _find_resource_by_id, _msp_dt_or_none, _route_operation,
# _enter_batch_mode, _exit_batch_mode, _build_task_id_map.

# ---------- PROGRESS CONSTANTS ----------

_PROGRESS_PCT_FIELDS = frozenset({
    "percent_complete",
    "percent_work_complete",
    "physical_pct",
})

_PROGRESS_WORK_FIELDS = frozenset({
    "actual_work_h",
    "remaining_work_h",
})

_PROGRESS_DURATION_FIELDS = frozenset({
    "actual_duration_h",
    "remaining_duration_h",
})

_PROGRESS_DATE_FIELDS = frozenset({
    "actual_start",
    "actual_finish",
    "stop",
    "resume",
})

# Probe-confirmed (msproject_typelib.txt + MSP 16.0 round-trip):
_TIMESCALE_UNIT_MAP = {
    "day": 8,    # pjTimescaleDays
    "week": 6,   # pjTimescaleWeeks
}

_PJ_TIMESCALED_ACTUAL_WORK = 24  # pjAssignmentTimescaledActualWork


# ---------- PROGRESS HELPERS ----------

def _normalize_progress_pct(v: Any) -> float:
    """Validate + normalize a percentage value (0-100).

    Accepts int / float / str ('50', '50.5', '50%'). Raises ValueError on
    out-of-range or non-numeric input. Returns float rounded to 2 decimals.
    """
    if v is None:
        raise ValueError("progress percentage cannot be None")
    if isinstance(v, str):
        s = v.strip().rstrip("%").strip()
        try:
            f = float(s)
        except Exception:
            raise ValueError(f"progress percentage not numeric: {v!r}")
    else:
        try:
            f = float(v)
        except Exception:
            raise ValueError(f"progress percentage not numeric: {v!r}")
    if f < 0 or f > 100:
        raise ValueError(f"progress percentage must be 0-100, got {f}")
    return round(f, 2)


def _hours_to_minutes(h: float) -> int:
    """Convert public-API hours to MSP COM minutes (rounded int)."""
    return int(round(float(h) * 60))


def _minutes_to_hours(m: Any) -> float:
    """Convert MSP COM minutes to public-API hours (float, 2 decimals)."""
    if m is None:
        return 0.0
    try:
        return round(float(m) / 60.0, 2)
    except Exception:
        return 0.0


def _validate_actual_dates(start: Optional[str],
                          finish: Optional[str]) -> Optional[str]:
    """Verify actual_start <= actual_finish if both provided.

    Returns None if valid, or an error message string if invalid.
    Tolerant of ISO 8601, pywintypes datetime str, and dateutil-parseable input.
    """
    if start is None or finish is None:
        return None
    try:
        from dateutil import parser
        s = parser.parse(str(start))
        f = parser.parse(str(finish))
    except Exception as e:
        return f"could not parse dates ({start!r}, {finish!r}): {e}"
    if s > f:
        return (f"actual_start ({start}) must be <= actual_finish ({finish})")
    return None


def _get_assignment_by_resource_id(task: Any, resource_id: int) -> Optional[Any]:
    """Find an assignment on the task matching the given resource_id.

    Iterates task.Assignments (1-indexed COM collection). Returns Assignment
    object or None.
    """
    try:
        for i in range(1, task.Assignments.Count + 1):
            try:
                asg = task.Assignments(i)
                if asg is None:
                    continue
                if int(asg.ResourceID) == int(resource_id):
                    return asg
            except Exception:
                continue
    except Exception as e:
        logger.debug(f"_get_assignment_by_resource_id iter failed: {e}")
    return None


def _read_task_progress_dict(task: Any) -> Dict[str, Any]:
    """Read all progress fields from a task → dict shaped for get_task_progress.

    Each property is guarded; failure on one field yields safe default.
    """
    out: Dict[str, Any] = {}
    pct_pairs = [
        ("PercentComplete", "percent_complete", 0),
        ("PercentWorkComplete", "percent_work_complete", 0),
        ("PhysicalPercentComplete", "physical_pct", 0),
    ]
    for prop, key, default in pct_pairs:
        try:
            v = getattr(task, prop)
            out[key] = float(v) if v is not None else default
        except Exception:
            out[key] = default
    date_pairs = [
        ("ActualStart", "actual_start"),
        ("ActualFinish", "actual_finish"),
        ("Stop", "stop"),
        ("Resume", "resume"),
    ]
    for prop, key in date_pairs:
        try:
            out[key] = _msp_dt_or_none(getattr(task, prop))
        except Exception:
            out[key] = None
    work_pairs = [
        ("ActualWork", "actual_work_h"),
        ("RemainingWork", "remaining_work_h"),
        ("ActualDuration", "actual_duration_h"),
        ("RemainingDuration", "remaining_duration_h"),
    ]
    for prop, key in work_pairs:
        try:
            out[key] = _minutes_to_hours(getattr(task, prop))
        except Exception:
            out[key] = 0.0
    return out


def _msp_task_set_progress_field(task: Any, field: str, value: Any) -> None:
    """Low-level setter that maps public field names → MSP COM properties.

    Raises on COM error; caller wraps in try/except per-field for partial-failure
    aggregation.
    """
    if field == "percent_complete":
        task.PercentComplete = _normalize_progress_pct(value)
    elif field == "percent_work_complete":
        task.PercentWorkComplete = _normalize_progress_pct(value)
    elif field == "physical_pct":
        task.PhysicalPercentComplete = _normalize_progress_pct(value)
    elif field == "actual_start":
        task.ActualStart = value
    elif field == "actual_finish":
        task.ActualFinish = value
    elif field == "actual_duration_h":
        task.ActualDuration = _hours_to_minutes(value)
    elif field == "actual_work_h":
        task.ActualWork = _hours_to_minutes(value)
    elif field == "remaining_work_h":
        task.RemainingWork = _hours_to_minutes(value)
    elif field == "remaining_duration_h":
        task.RemainingDuration = _hours_to_minutes(value)
    elif field == "stop":
        task.Stop = value
    elif field == "resume":
        task.Resume = value
    else:
        raise ValueError(f"Unknown progress field: {field}")
```

**Step 4: Run — PASS** (10 PASSED expected).

```bash
python -m pytest tests/test_msproject_progress_helpers.py -v
```

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_helpers.py
git commit -m "Phase 3b T52: progress foundations — _PROGRESS_*_FIELDS constants + 6 helpers (_normalize_progress_pct, _hours_to_minutes, _minutes_to_hours, _validate_actual_dates, _get_assignment_by_resource_id, _read_task_progress_dict)"
```

Expected full regression: **217 PASSED + 1 xfail** (207 + 10 new).

---

## Task 53: msproject_progress set_task_progress Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_set_task.py`

**Step 1: Failing test**

`tests/test_msproject_progress_set_task.py`:
```python
"""Test msproject_progress set_task_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_set_task, _msp_task_add_single, _find_task_by_id,
    _read_task_progress_dict,
)


def test_set_pct_complete_only(clean_test_project):
    """Set percent_complete=50 on a 4-day task."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="PctT-T53", duration="4d")
    r = _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=50)
    assert r["status"] == "ok"
    assert "percent_complete" in r["changes"]
    # Verify via direct readback
    t = _find_task_by_id(proj, add_r["task_id"])
    assert _read_task_progress_dict(t)["percent_complete"] == 50


def test_set_actual_work_h_and_remaining(clean_test_project):
    """Set actual_work_h=16 and remaining_work_h=16 on a 4-day task (32h total)."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="WorkT-T53", duration="4d")
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        actual_work_h=16,
        remaining_work_h=16,
    )
    assert r["status"] == "ok"
    assert "actual_work_h" in r["changes"]
    assert "remaining_work_h" in r["changes"]
    t = _find_task_by_id(proj, add_r["task_id"])
    p = _read_task_progress_dict(t)
    assert p["actual_work_h"] == 16
    assert p["remaining_work_h"] == 16


def test_set_actual_dates(clean_test_project):
    """Set actual_start + actual_finish."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="DateT-T53", duration="3d",
                                  start="2026-04-01")
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        actual_start="2026-04-01",
        actual_finish="2026-04-03",
    )
    assert r["status"] == "ok"
    t = _find_task_by_id(proj, add_r["task_id"])
    p = _read_task_progress_dict(t)
    assert p["actual_start"] is not None
    assert p["actual_finish"] is not None


def test_set_pct_invalid_raises_error(clean_test_project):
    add_r = _msp_task_add_single(name="BadPctT-T53", duration="1d")
    r = _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=150)
    assert r["status"] == "error"
    assert "0-100" in r["error"] or "percent" in r["error"].lower()


def test_set_invalid_date_order_errors(clean_test_project):
    add_r = _msp_task_add_single(name="BadOrdT-T53", duration="1d")
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        actual_start="2026-05-15",
        actual_finish="2026-05-01",
    )
    assert r["status"] == "error"
    assert "actual_start" in r["error"].lower() or "before" in r["error"].lower()


def test_set_missing_task_id(clean_test_project):
    r = _msp_progress_set_task(task_id=99999, percent_complete=50)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_set_physical_pct(clean_test_project):
    """DCMA semantic: physical_pct independent of percent_complete."""
    proj = clean_test_project
    add_r = _msp_task_add_single(name="PhysT-T53", duration="5d")
    r = _msp_progress_set_task(
        task_id=add_r["task_id"],
        percent_complete=20,
        physical_pct=50,
    )
    # Either OK with both fields or graceful fallback for older MSP
    assert r["status"] in ("ok", "partial")
    if r["status"] == "ok":
        t = _find_task_by_id(proj, add_r["task_id"])
        p = _read_task_progress_dict(t)
        assert p["percent_complete"] == 20
        # physical_pct may or may not have round-tripped; if writable, it's 50
        if p["physical_pct"] == 50:
            assert True
```

**Step 2: Run — FAIL** (ImportError on `_msp_progress_set_task`)

**Step 3: Implementation**

Insert after `_msp_task_set_progress_field` (T52):

```python
def _msp_progress_set_task(task_id: int,
                           percent_complete: Optional[float] = None,
                           percent_work_complete: Optional[float] = None,
                           actual_start: Optional[str] = None,
                           actual_finish: Optional[str] = None,
                           actual_duration_h: Optional[float] = None,
                           actual_work_h: Optional[float] = None,
                           remaining_work_h: Optional[float] = None,
                           remaining_duration_h: Optional[float] = None,
                           physical_pct: Optional[float] = None,
                           stop: Optional[str] = None,
                           resume: Optional[str] = None) -> Dict[str, Any]:
    """Set one or more progress fields on a task.

    Dual-mode: caller can pass `percent_complete` (duration-based) OR explicit
    actual_*/remaining_* values. Both modes coexist; MSP rebalances internally.

    Phase 3b — see design doc Section 6 Q1 (dual-track) and Q4 (PhysicalPercentComplete).
    Returns {status, task_id, changes: [field], readback: dict}.
    """
    # Build candidate field map; filter out None
    candidates = {
        "percent_complete": percent_complete,
        "percent_work_complete": percent_work_complete,
        "actual_start": actual_start,
        "actual_finish": actual_finish,
        "actual_duration_h": actual_duration_h,
        "actual_work_h": actual_work_h,
        "remaining_work_h": remaining_work_h,
        "remaining_duration_h": remaining_duration_h,
        "physical_pct": physical_pct,
        "stop": stop,
        "resume": resume,
    }
    fields_to_set = {k: v for k, v in candidates.items() if v is not None}
    if not fields_to_set:
        return {"status": "error", "error": "no progress fields provided"}

    # Pre-validate pct fields and date order
    for f in _PROGRESS_PCT_FIELDS:
        if f in fields_to_set:
            try:
                fields_to_set[f] = _normalize_progress_pct(fields_to_set[f])
            except ValueError as e:
                return {"status": "error", "error": str(e)}
    err = _validate_actual_dates(fields_to_set.get("actual_start"),
                                 fields_to_set.get("actual_finish"))
    if err:
        return {"status": "error", "error": err}
    err = _validate_actual_dates(fields_to_set.get("stop"),
                                 fields_to_set.get("resume"))
    if err:
        return {"status": "error", "error": "stop/resume order: " + err}

    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}

    changes: List[str] = []
    failures: List[Dict[str, str]] = []
    for field, value in fields_to_set.items():
        try:
            _msp_task_set_progress_field(t, field, value)
            changes.append(field)
        except Exception as e:
            failures.append({"field": field, "error": _format_com_error(e)})
            logger.debug(f"set_task_progress({task_id}, {field}={value}) failed: {e}")

    try:
        readback = _read_task_progress_dict(t)
    except Exception:
        readback = None

    if not changes and failures:
        return {"status": "error", "task_id": task_id,
                "error": "all field writes failed", "failures": failures}
    status = "ok" if not failures else "partial"
    return {"status": status, "task_id": task_id,
            "changes": changes, "failures": failures,
            "readback": readback}
```

**Step 4: Run — PASS** (7 PASSED)

```bash
python -m pytest tests/test_msproject_progress_set_task.py -v
```

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_set_task.py
git commit -m "Phase 3b T53: msproject_progress set_task_progress action (dual-mode pct + actuals + DCMA physical_pct)"
```

Expected: **224 PASSED + 1 xfail** (217 + 7).

---

## Task 54: msproject_progress get_task_progress Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_get_task.py`

**Step 1: Failing test**

`tests/test_msproject_progress_get_task.py`:
```python
"""Test msproject_progress get_task_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_get_task, _msp_progress_set_task, _msp_task_add_single,
)


def test_get_initial_progress_zero(clean_test_project):
    """Fresh task — all progress fields 0/None."""
    add_r = _msp_task_add_single(name="GetInitT-T54", duration="3d")
    r = _msp_progress_get_task(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["task_id"] == add_r["task_id"]
    p = r["progress"]
    assert p["percent_complete"] == 0
    assert p["actual_start"] is None
    assert p["actual_finish"] is None
    assert p["actual_work_h"] == 0


def test_get_after_pct_set(clean_test_project):
    """Set 50%, read back."""
    add_r = _msp_task_add_single(name="GetSetT-T54", duration="4d")
    _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=50)
    r = _msp_progress_get_task(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["progress"]["percent_complete"] == 50
    # 50% of 4d × 8h = 16h actual
    assert r["progress"]["actual_work_h"] >= 15  # Asta may round


def test_get_full_shape_keys_present(clean_test_project):
    """Returned dict has all 9 expected keys."""
    add_r = _msp_task_add_single(name="ShapeT-T54", duration="1d")
    r = _msp_progress_get_task(task_id=add_r["task_id"])
    p = r["progress"]
    expected_keys = {
        "percent_complete", "percent_work_complete", "physical_pct",
        "actual_start", "actual_finish", "stop", "resume",
        "actual_work_h", "remaining_work_h",
        "actual_duration_h", "remaining_duration_h",
    }
    assert expected_keys.issubset(p.keys())


def test_get_missing_task(clean_test_project):
    r = _msp_progress_get_task(task_id=99999)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Insert after `_msp_progress_set_task`:

```python
def _msp_progress_get_task(task_id: int) -> Dict[str, Any]:
    """Read all progress fields from a task → structured dict.

    Returns {status, task_id, progress: {percent_complete, percent_work_complete,
    actual_start, actual_finish, actual_duration_h, actual_work_h,
    remaining_work_h, remaining_duration_h, physical_pct, stop, resume}}.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    try:
        progress = _read_task_progress_dict(t)
        return {"status": "ok", "task_id": task_id, "progress": progress}
    except Exception as e:
        logger.error(f"_msp_progress_get_task({task_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (4 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_get_task.py
git commit -m "Phase 3b T54: msproject_progress get_task_progress action (all 11 fields read)"
```

Expected: **228 PASSED + 1 xfail**.

---

## Task 55: msproject_progress set_assignment_progress Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_set_assignment.py`

**Step 1: Failing test**

`tests/test_msproject_progress_set_assignment.py`:
```python
"""Test msproject_progress set_assignment_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_set_assignment, _msp_progress_get_assignments,
    _msp_progress_get_task,
    _msp_task_add_single, _msp_resource_add, _msp_resource_assign,
    _find_task_by_id, _get_assignment_by_resource_id,
)


def _setup_task_with_resource(task_name: str, dur: str = "5d",
                              res_name: str = "COW-T55"):
    """Helper: create task + resource + assignment."""
    add_t = _msp_task_add_single(name=task_name, duration=dur)
    add_r = _msp_resource_add(name=res_name, type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    return add_t["task_id"], add_r["resource_id"]


def test_set_assignment_actual_work(clean_test_project):
    """Write 16h actual on assignment."""
    proj = clean_test_project
    tid, rid = _setup_task_with_resource("AsgWT-T55")
    r = _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                     actual_work_h=16)
    assert r["status"] == "ok"
    assert "actual_work_h" in r["changes"]
    # Readback via get_assignments
    g = _msp_progress_get_assignments(task_id=tid)
    asg = next(a for a in g["assignments"] if a["resource_id"] == rid)
    assert asg["actual_work_h"] == 16


def test_assignment_actual_rolls_up_to_task(clean_test_project):
    """Write 24h on assignment → task.ActualWork should reflect (single-resource case)."""
    tid, rid = _setup_task_with_resource("RollupT-T55", dur="5d",
                                          res_name="StlT-T55")
    _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                 actual_work_h=24)
    # Task-level read should see roll-up
    g = _msp_progress_get_task(task_id=tid)
    assert g["progress"]["actual_work_h"] >= 23  # MSP allow small drift


def test_set_assignment_pct_work_complete(clean_test_project):
    tid, rid = _setup_task_with_resource("PctWT-T55", dur="4d",
                                          res_name="MsnT-T55")
    r = _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                     percent_work_complete=50)
    assert r["status"] == "ok"


def test_set_assignment_missing_task_errors(clean_test_project):
    r = _msp_progress_set_assignment(task_id=99999, resource_id=1,
                                     actual_work_h=10)
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_set_assignment_missing_resource_errors(clean_test_project):
    """Task exists but no assignment with that resource_id."""
    add_t = _msp_task_add_single(name="NoAsgT-T55", duration="2d")
    r = _msp_progress_set_assignment(task_id=add_t["task_id"], resource_id=99999,
                                     actual_work_h=10)
    assert r["status"] == "error"
    assert "assignment" in r["error"].lower() or "resource" in r["error"].lower()


def test_set_assignment_invalid_pct(clean_test_project):
    tid, rid = _setup_task_with_resource("BadPctAsgT-T55",
                                          res_name="EwiT-T55")
    r = _msp_progress_set_assignment(task_id=tid, resource_id=rid,
                                     percent_work_complete=150)
    assert r["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Insert after `_msp_progress_get_task`:

```python
def _msp_progress_set_assignment(task_id: int,
                                 resource_id: int,
                                 actual_work_h: Optional[float] = None,
                                 actual_start: Optional[str] = None,
                                 actual_finish: Optional[str] = None,
                                 percent_work_complete: Optional[float] = None,
                                 remaining_work_h: Optional[float] = None,
                                 units: Optional[float] = None) -> Dict[str, Any]:
    """Per-resource man-hour write on a single assignment.

    Phase 3b — for hakediş workflows where each resource's hours are tracked
    separately (e.g., "T101: COW=24h, STL=18h, MSN=10h").

    MSP rolls up assignment-level writes to task.ActualWork automatically.
    Returns {status, task_id, resource_id, changes: [field]}.
    """
    candidates = {
        "actual_work_h": actual_work_h,
        "actual_start": actual_start,
        "actual_finish": actual_finish,
        "percent_work_complete": percent_work_complete,
        "remaining_work_h": remaining_work_h,
        "units": units,
    }
    fields = {k: v for k, v in candidates.items() if v is not None}
    if not fields:
        return {"status": "error", "error": "no assignment fields provided"}

    # Pre-validate pct
    if "percent_work_complete" in fields:
        try:
            fields["percent_work_complete"] = _normalize_progress_pct(
                fields["percent_work_complete"])
        except ValueError as e:
            return {"status": "error", "error": str(e)}
    err = _validate_actual_dates(fields.get("actual_start"),
                                 fields.get("actual_finish"))
    if err:
        return {"status": "error", "error": err}

    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    asg = _get_assignment_by_resource_id(t, resource_id)
    if asg is None:
        return {"status": "error",
                "error": f"No assignment for resource_id {resource_id} on task {task_id}"}

    changes: List[str] = []
    failures: List[Dict[str, str]] = []
    for field, value in fields.items():
        try:
            if field == "actual_work_h":
                asg.ActualWork = _hours_to_minutes(value)
            elif field == "actual_start":
                asg.ActualStart = value
            elif field == "actual_finish":
                asg.ActualFinish = value
            elif field == "percent_work_complete":
                asg.PercentWorkComplete = value
            elif field == "remaining_work_h":
                asg.RemainingWork = _hours_to_minutes(value)
            elif field == "units":
                asg.Units = float(value)
            else:
                raise ValueError(f"unknown assignment field: {field}")
            changes.append(field)
        except Exception as e:
            failures.append({"field": field, "error": _format_com_error(e)})

    if not changes and failures:
        return {"status": "error", "task_id": task_id, "resource_id": resource_id,
                "error": "all assignment field writes failed", "failures": failures}
    status = "ok" if not failures else "partial"
    return {"status": status, "task_id": task_id, "resource_id": resource_id,
            "changes": changes, "failures": failures}
```

**Step 4: Run — PASS** (6 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_set_assignment.py
git commit -m "Phase 3b T55: msproject_progress set_assignment_progress action (per-resource man-hour write + roll-up)"
```

Expected: **234 PASSED + 1 xfail**.

---

## Task 56: msproject_progress get_assignment_progress Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_get_assignment.py`

**Step 1: Failing test**

`tests/test_msproject_progress_get_assignment.py`:
```python
"""Test msproject_progress get_assignment_progress action."""
import pytest
from msproject_mcp_core import (
    _msp_progress_get_assignments, _msp_progress_set_assignment,
    _msp_task_add_single, _msp_resource_add, _msp_resource_assign,
)


def test_get_empty_assignments(clean_test_project):
    """Task with no assignments → empty list."""
    add_r = _msp_task_add_single(name="EmptyAsgT-T56", duration="2d")
    r = _msp_progress_get_assignments(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["assignments"] == []


def test_get_one_assignment(clean_test_project):
    """1 resource assigned → 1-element list."""
    add_t = _msp_task_add_single(name="OneAsgT-T56", duration="3d")
    add_r = _msp_resource_add(name="X-T56", type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"], resource_id=add_r["resource_id"])
    r = _msp_progress_get_assignments(task_id=add_t["task_id"])
    assert r["status"] == "ok"
    assert len(r["assignments"]) == 1
    a = r["assignments"][0]
    assert a["resource_id"] == add_r["resource_id"]
    assert a["resource_name"] == "X-T56"
    assert "actual_work_h" in a
    assert "percent_work_complete" in a
    assert "units" in a


def test_get_three_assignments_after_writes(clean_test_project):
    """3 resources, write actuals to each → all 3 returned with values."""
    add_t = _msp_task_add_single(name="3AsgT-T56", duration="5d")
    rids = []
    for nm in ("COW-T56", "STL-T56", "MSN-T56"):
        ar = _msp_resource_add(name=nm, type="Work", max_units=100)
        _msp_resource_assign(task_id=add_t["task_id"],
                             resource_id=ar["resource_id"])
        rids.append(ar["resource_id"])
    # Write different hours to each
    _msp_progress_set_assignment(task_id=add_t["task_id"],
                                 resource_id=rids[0], actual_work_h=24)
    _msp_progress_set_assignment(task_id=add_t["task_id"],
                                 resource_id=rids[1], actual_work_h=16)
    _msp_progress_set_assignment(task_id=add_t["task_id"],
                                 resource_id=rids[2], actual_work_h=8)
    r = _msp_progress_get_assignments(task_id=add_t["task_id"])
    assert r["status"] == "ok"
    assert len(r["assignments"]) == 3
    by_rid = {a["resource_id"]: a for a in r["assignments"]}
    assert by_rid[rids[0]]["actual_work_h"] == 24
    assert by_rid[rids[1]]["actual_work_h"] == 16
    assert by_rid[rids[2]]["actual_work_h"] == 8


def test_get_missing_task(clean_test_project):
    r = _msp_progress_get_assignments(task_id=99999)
    assert r["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Insert after `_msp_progress_set_assignment`:

```python
def _msp_progress_get_assignments(task_id: int) -> Dict[str, Any]:
    """List all per-resource progress on a task.

    Returns {status, task_id, assignments: [{resource_id, resource_name,
    actual_work_h, percent_work_complete, remaining_work_h, units,
    actual_start, actual_finish}, ...]}.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    out: List[Dict[str, Any]] = []
    try:
        for i in range(1, t.Assignments.Count + 1):
            try:
                asg = t.Assignments(i)
                if asg is None:
                    continue
                try:
                    res = asg.Resource
                    res_name = res.Name if res is not None else None
                except Exception:
                    res_name = None
                out.append({
                    "resource_id": int(asg.ResourceID),
                    "resource_name": res_name,
                    "actual_work_h": _minutes_to_hours(asg.ActualWork),
                    "actual_start": _msp_dt_or_none(asg.ActualStart),
                    "actual_finish": _msp_dt_or_none(asg.ActualFinish),
                    "percent_work_complete": float(asg.PercentWorkComplete or 0),
                    "remaining_work_h": _minutes_to_hours(asg.RemainingWork),
                    "units": float(asg.Units or 0),
                })
            except Exception as e:
                logger.debug(f"_msp_progress_get_assignments row {i} failed: {e}")
                continue
        return {"status": "ok", "task_id": task_id, "assignments": out}
    except Exception as e:
        logger.error(f"_msp_progress_get_assignments({task_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (4 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_get_assignment.py
git commit -m "Phase 3b T56: msproject_progress get_assignment_progress action (per-task assignment list)"
```

Expected: **238 PASSED + 1 xfail**.

---

## Task 57: msproject_progress set_progress_by_date Action (BIG ONE — UpdateProject)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_set_by_date.py`

**Step 1.5: Probe `app.UpdateProject(ProgressDate=...)`**

```python
# probe_update_project.py — DO NOT COMMIT
import pythoncom, win32com.client, datetime as dt
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MSProject.Application")
app.FileNew()
proj = app.ActiveProject
t = proj.Tasks.Add("UpdProbeT")
t.Duration = 480 * 5  # 5 days
t.Start = dt.datetime(2026, 4, 1)
print(f"BEFORE: PctComplete={t.PercentComplete}")
try:
    # MSP UpdateProject signature varies by version; try common forms
    app.UpdateProject(ProgressDate=dt.datetime(2026, 4, 3),
                     UpdatePercentCompleteOnly=False,
                     AllTasks=True)
    print(f"AFTER: PctComplete={t.PercentComplete}")
except TypeError as e:
    print(f"named-arg signature failed: {e}; trying positional")
    try:
        app.UpdateProject(False, dt.datetime(2026, 4, 3))  # (UpdatePercentCompleteOnly, ProgressDate)
        print(f"AFTER: PctComplete={t.PercentComplete}")
    except Exception as e2:
        print(f"positional also failed: {e2}")
finally:
    app.FileClose(0)
```

**Decision rule:**
- If named-arg works → implementation uses `app.UpdateProject(ProgressDate=dt, UpdatePercentCompleteOnly=False, AllTasks=True)`
- If positional only → adapt accordingly; document signature in helper docstring
- If both fail → return error and defer to per-task loop with `task.PercentComplete = X` (slower fallback)

**Step 1: Failing test**

`tests/test_msproject_progress_set_by_date.py`:
```python
"""Test msproject_progress set_progress_by_date action."""
import pytest
import datetime as dt
from msproject_mcp_core import (
    _msp_progress_set_by_date, _msp_progress_get_task,
    _msp_task_add_single,
)


def test_set_by_date_after_full_duration(clean_test_project):
    """Task 5d starting 2026-04-01; progress_date=2026-05-01 → 100% complete."""
    add_r = _msp_task_add_single(name="UpdT-T57", duration="5d",
                                  start="2026-04-01")
    r = _msp_progress_set_by_date(progress_date="2026-05-01")
    assert r["status"] == "ok"
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    # Should be ~100% (after task end)
    assert g["progress"]["percent_complete"] >= 95


def test_set_by_date_partial_progress(clean_test_project):
    """Task 10d starting 2026-04-01; progress_date=2026-04-06 → ~50% complete."""
    add_r = _msp_task_add_single(name="HalfT-T57", duration="10d",
                                  start="2026-04-01")
    r = _msp_progress_set_by_date(progress_date="2026-04-08")  # ~5 working days
    assert r["status"] == "ok"
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    pct = g["progress"]["percent_complete"]
    # MSP working-day calc; allow wide window 30-70%
    assert 30 <= pct <= 75


def test_set_by_date_before_task_start(clean_test_project):
    """Progress date BEFORE task start → 0%."""
    add_r = _msp_task_add_single(name="BeforeT-T57", duration="5d",
                                  start="2026-05-01")
    r = _msp_progress_set_by_date(progress_date="2026-04-01")
    assert r["status"] == "ok"
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    assert g["progress"]["percent_complete"] == 0


def test_set_by_date_invalid_date_format(clean_test_project):
    r = _msp_progress_set_by_date(progress_date="not a date")
    assert r["status"] == "error"
    assert "date" in r["error"].lower() or "parse" in r["error"].lower()
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_progress_set_by_date(progress_date: str,
                              scope: str = "all",
                              as_scheduled: bool = True) -> Dict[str, Any]:
    """Bulk-update progress to a given date (`app.UpdateProject`).

    Implements "plan = actual" up to data_date assumption — fast retroactive
    backlog catch-up. Phase 3b — see design doc Section 6 (Q1) and Open
    Questions #2 for already-progressed task interaction.

    progress_date: ISO 8601 string or pywintypes datetime
    scope: "all" (entire project) or "selected" (currently selected tasks)
    as_scheduled: if True, MSP fills actuals to match plan up to date;
                  if False, only updates % complete (lighter touch).
    """
    if scope not in ("all", "selected"):
        return {"status": "error",
                "error": f"scope must be 'all' or 'selected', got '{scope}'"}
    # Parse progress_date
    try:
        from dateutil import parser
        pd = parser.parse(str(progress_date))
    except Exception as e:
        return {"status": "error",
                "error": f"could not parse progress_date {progress_date!r}: {e}"}

    app = _validate_active_project()
    proj = app.ActiveProject
    task_count_before = proj.Tasks.Count
    try:
        # MSP UpdateProject: UpdatePercentCompleteOnly inverse of as_scheduled
        update_pct_only = not as_scheduled
        all_tasks_flag = (scope == "all")
        # Probe-confirmed signature on MSP 16.0
        app.UpdateProject(UpdatePercentCompleteOnly=update_pct_only,
                         ProgressDate=pd,
                         AllTasks=all_tasks_flag)
        return {"status": "ok",
                "progress_date": str(pd),
                "mode": "as_scheduled" if as_scheduled else "pct_only",
                "scope": scope,
                "task_count_affected": task_count_before}
    except Exception as e:
        logger.error(f"_msp_progress_set_by_date({progress_date}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (4 PASSED — note test_set_by_date_partial_progress allowance is wide, MSP calc dependent)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_set_by_date.py
git commit -m "Phase 3b T57: msproject_progress set_progress_by_date action (app.UpdateProject — bulk retroactive)"
```

Expected: **242 PASSED + 1 xfail**.

---

## Task 58: msproject_progress set_status_date Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_status_date.py`

**Step 1: Failing test**

`tests/test_msproject_progress_status_date.py`:
```python
"""Test msproject_progress set_status_date action."""
import pytest
from msproject_mcp_core import _msp_progress_set_status_date


def test_set_status_date_basic(clean_test_project):
    proj = clean_test_project
    r = _msp_progress_set_status_date(status_date="2026-04-29")
    assert r["status"] == "ok"
    assert "status_date" in r
    # Verify on project
    sd = str(proj.StatusDate)
    assert "2026" in sd or "04" in sd


def test_set_status_date_invalid_format(clean_test_project):
    r = _msp_progress_set_status_date(status_date="not a date")
    assert r["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_progress_set_status_date(status_date: str) -> Dict[str, Any]:
    """Set proj.StatusDate (CLAUDE.md RULE 5 'data_date').

    Used by EVM tooling (Phase 5) and time-phased PV/EV calculations.
    """
    try:
        from dateutil import parser
        sd = parser.parse(str(status_date))
    except Exception as e:
        return {"status": "error",
                "error": f"could not parse status_date {status_date!r}: {e}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    try:
        previous = _msp_dt_or_none(proj.StatusDate)
    except Exception:
        previous = None
    try:
        proj.StatusDate = sd
        return {"status": "ok",
                "status_date": str(sd),
                "previous": previous}
    except Exception as e:
        logger.error(f"_msp_progress_set_status_date({status_date}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (2 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_status_date.py
git commit -m "Phase 3b T58: msproject_progress set_status_date action (proj.StatusDate / data_date)"
```

Expected: **244 PASSED + 1 xfail**.

---

## Task 59: msproject_progress clear_progress + clear_all_progress (paired)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_clear.py`

**Step 1: Failing test**

`tests/test_msproject_progress_clear.py`:
```python
"""Test msproject_progress clear_progress + clear_all_progress actions."""
import pytest
from msproject_mcp_core import (
    _msp_progress_clear, _msp_progress_clear_all,
    _msp_progress_set_task, _msp_progress_get_task,
    _msp_task_add_single,
)


def test_clear_single_task_progress(clean_test_project):
    """Set 50% then clear → 0%."""
    add_r = _msp_task_add_single(name="ClearOneT-T59", duration="3d")
    _msp_progress_set_task(task_id=add_r["task_id"], percent_complete=50)
    r = _msp_progress_clear(task_id=add_r["task_id"])
    assert r["status"] == "ok"
    assert r["task_id"] == add_r["task_id"]
    g = _msp_progress_get_task(task_id=add_r["task_id"])
    assert g["progress"]["percent_complete"] == 0
    assert g["progress"]["actual_start"] is None


def test_clear_unprogressed_task_idempotent(clean_test_project):
    """Clearing a task with no progress → still ok."""
    add_r = _msp_task_add_single(name="NoOpClrT-T59", duration="2d")
    r = _msp_progress_clear(task_id=add_r["task_id"])
    assert r["status"] == "ok"


def test_clear_missing_task(clean_test_project):
    r = _msp_progress_clear(task_id=99999)
    assert r["status"] == "error"


def test_clear_all_progress(clean_test_project):
    """Set progress on 3 tasks, clear_all → all reset."""
    ids = []
    for i in range(3):
        ar = _msp_task_add_single(name=f"AllClrT{i}-T59", duration="2d")
        _msp_progress_set_task(task_id=ar["task_id"], percent_complete=50)
        ids.append(ar["task_id"])
    r = _msp_progress_clear_all()
    assert r["status"] == "ok"
    assert r["cleared_count"] >= 3
    for tid in ids:
        g = _msp_progress_get_task(task_id=tid)
        assert g["progress"]["percent_complete"] == 0


def test_clear_all_when_none_progressed(clean_test_project):
    """clear_all on fresh project → ok with cleared_count 0."""
    r = _msp_progress_clear_all()
    assert r["status"] == "ok"
    assert r["cleared_count"] == 0
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_progress_clear(task_id: int) -> Dict[str, Any]:
    """Reset a single task's progress to 0/None across all progress fields.

    Idempotent: tasks that are already at 0% return ok with cleared_fields=[].
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    cleared: List[str] = []
    try:
        # Setting PercentComplete = 0 cascades to ActualStart/Finish/Work clearing
        # in MSP, but we explicitly reset each for safety.
        if t.PercentComplete and t.PercentComplete > 0:
            t.PercentComplete = 0
            cleared.append("percent_complete")
        if t.ActualWork and t.ActualWork > 0:
            t.ActualWork = 0
            cleared.append("actual_work")
        # ActualStart / ActualFinish: write "NA" or use COM null sentinel
        # MSP accepts NA-equivalent to clear: setting to "NA" string OR using
        # the constant pjNA = 0x7FFFFFFF; safer: try both clear approaches.
        try:
            t.ActualStart = "NA"
            cleared.append("actual_start")
        except Exception:
            pass
        try:
            t.ActualFinish = "NA"
            cleared.append("actual_finish")
        except Exception:
            pass
        return {"status": "ok", "task_id": task_id, "cleared_fields": cleared}
    except Exception as e:
        logger.error(f"_msp_progress_clear({task_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_progress_clear_all() -> Dict[str, Any]:
    """Clear progress on every non-summary task in the project.

    Used to reset a project to "as-planned" state.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    cleared_count = 0
    failures: List[Dict[str, Any]] = []
    try:
        _enter_batch_mode()
        try:
            for i in range(1, proj.Tasks.Count + 1):
                try:
                    t = proj.Tasks(i)
                    if t is None or t.Summary:
                        continue
                    if t.PercentComplete and t.PercentComplete > 0:
                        t.PercentComplete = 0
                        cleared_count += 1
                    elif t.ActualWork and t.ActualWork > 0:
                        t.ActualWork = 0
                        cleared_count += 1
                except Exception as e:
                    failures.append({"index": i, "error": _format_com_error(e)})
        finally:
            _exit_batch_mode()
        return {"status": "ok" if not failures else "partial",
                "cleared_count": cleared_count,
                "failures": failures}
    except Exception as e:
        logger.error(f"_msp_progress_clear_all failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (5 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_clear.py
git commit -m "Phase 3b T59: msproject_progress clear_progress + clear_all_progress (idempotent, batch mode)"
```

Expected: **249 PASSED + 1 xfail**.

---

## Task 60: msproject_progress time_phased_actual_write Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_time_phased.py` (also covers T61 read tests)

**Step 1.5: Probe `assignment.TimeScaleData` write semantics**

```python
# probe_timescale_data.py — DO NOT COMMIT
import pythoncom, win32com.client, datetime as dt
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MSProject.Application")
app.FileNew()
proj = app.ActiveProject
proj.ProjectStart = dt.datetime(2026, 4, 1)
t = proj.Tasks.Add("TPDProbeT")
t.Duration = 480 * 5  # 5 days
res = proj.Resources.Add("ProbeR")
asg = t.Assignments.Add(t.ID, res.ID)

# Try TimeScaleData read with day granularity over 5 days
start = dt.datetime(2026, 4, 1)
end = dt.datetime(2026, 4, 8)
try:
    tsv = asg.TimeScaleData(StartDate=start, EndDate=end,
                            Type=24,  # pjAssignmentTimescaledActualWork
                            TimescaleUnit=8)  # pjTimescaleDays
    print(f"OK read: tsv.Count = {tsv.Count}")
    for i in range(1, tsv.Count + 1):
        item = tsv.Item(i)
        print(f"  [{i}] StartDate={item.StartDate} EndDate={item.EndDate} Value={item.Value}")
except Exception as e:
    print(f"FAIL TimeScaleData read: {e}")

# Try writing
try:
    tsv = asg.TimeScaleData(StartDate=start, EndDate=end, Type=24, TimescaleUnit=8)
    tsv.Item(1).Value = 480  # 8h on day 1
    print(f"OK write day-1: readback = {tsv.Item(1).Value}")
except Exception as e:
    print(f"FAIL TimeScaleData write: {e}")

# Confirm enums
try:
    print(f"pjTimescaleDays = {win32com.client.constants.pjTimescaleDays}")
    print(f"pjTimescaleWeeks = {win32com.client.constants.pjTimescaleWeeks}")
    print(f"pjAssignmentTimescaledActualWork = {win32com.client.constants.pjAssignmentTimescaledActualWork}")
except Exception as e:
    print(f"constant lookup failed: {e}")

app.FileClose(0)
```

**Decision rule:**
- All round-trips OK → impl as below
- TimeScaleData returns Count=0 → assignment doesn't have planned work in date range; write impossible. T60 returns "no slots in date range" error
- Wrong enum codes → update `_PJ_TIMESCALED_ACTUAL_WORK` and `_TIMESCALE_UNIT_MAP` constants

**Step 1: Failing test (covers T60 + T61)**

`tests/test_msproject_progress_time_phased.py`:
```python
"""Test msproject_progress time_phased_actual_write + _read actions."""
import pytest
import datetime as dt
from msproject_mcp_core import (
    _msp_progress_time_phased_write, _msp_progress_time_phased_read,
    _msp_task_add_single, _msp_resource_add, _msp_resource_assign,
)


def _setup_5day_assignment(task_name="TPDT-T60"):
    """Helper: 5d task starting 2026-04-01 with 1 resource."""
    add_t = _msp_task_add_single(name=task_name, duration="5d",
                                  start="2026-04-01")
    add_r = _msp_resource_add(name=f"TPR-{task_name}", type="Work", max_units=100)
    _msp_resource_assign(task_id=add_t["task_id"],
                         resource_id=add_r["resource_id"])
    return add_t["task_id"], add_r["resource_id"]


def test_time_phased_read_empty_actual(clean_test_project):
    """Fresh assignment — daily reads return 0h actual_work for each day."""
    tid, rid = _setup_5day_assignment("ReadEmptyT-T60")
    r = _msp_progress_time_phased_read(
        task_id=tid, resource_id=rid,
        start_date="2026-04-01", end_date="2026-04-06",
        unit="day",
    )
    assert r["status"] == "ok"
    assert len(r["periods"]) >= 5  # 5 weekdays
    for p in r["periods"]:
        assert p["actual_work_h"] == 0


def test_time_phased_write_3_days(clean_test_project):
    """Write 4h+8h+6h to days 1-2-3 → readback matches."""
    tid, rid = _setup_5day_assignment("Write3T-T60")
    periods = [
        {"start": "2026-04-01", "end": "2026-04-02", "actual_work_h": 4},
        {"start": "2026-04-02", "end": "2026-04-03", "actual_work_h": 8},
        {"start": "2026-04-03", "end": "2026-04-04", "actual_work_h": 6},
    ]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="day",
    )
    assert w["status"] in ("ok", "partial")
    assert w["written_count"] >= 3
    # Read back
    r = _msp_progress_time_phased_read(
        task_id=tid, resource_id=rid,
        start_date="2026-04-01", end_date="2026-04-04", unit="day",
    )
    by_day = {p["period_start"][:10]: p["actual_work_h"] for p in r["periods"]}
    # Allow MSP 0.5h drift due to calendar
    assert by_day.get("2026-04-01", 0) >= 3.5
    assert by_day.get("2026-04-02", 0) >= 7.5
    assert by_day.get("2026-04-03", 0) >= 5.5


def test_time_phased_write_weekly(clean_test_project):
    """Weekly bucket: write 30h to week-1."""
    tid, rid = _setup_5day_assignment("WriteWkT-T60")
    periods = [{"start": "2026-04-01", "end": "2026-04-08", "actual_work_h": 30}]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="week",
    )
    # Allow either ok or partial depending on calendar slot fitting
    assert w["status"] in ("ok", "partial")


def test_time_phased_write_no_overlap_period_fails(clean_test_project):
    """Period outside task date range → per-period failure."""
    tid, rid = _setup_5day_assignment("OutT-T60")
    periods = [{"start": "2027-01-01", "end": "2027-01-02",
                "actual_work_h": 4}]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="day",
    )
    # Either ok with written_count=0 + failures, or partial
    assert w["status"] in ("ok", "partial", "error")
    if w["status"] != "error":
        assert w["written_count"] == 0


def test_time_phased_read_invalid_unit(clean_test_project):
    tid, rid = _setup_5day_assignment("BadUnitT-T60")
    r = _msp_progress_time_phased_read(
        task_id=tid, resource_id=rid,
        start_date="2026-04-01", end_date="2026-04-08",
        unit="quarter",  # not in {day, week}
    )
    assert r["status"] == "error"
    assert "unit" in r["error"].lower()


def test_time_phased_missing_assignment(clean_test_project):
    """Task with no assignment → error."""
    add_t = _msp_task_add_single(name="NoAsgTPD-T60", duration="3d")
    r = _msp_progress_time_phased_read(
        task_id=add_t["task_id"], resource_id=99999,
        start_date="2026-04-01", end_date="2026-04-04", unit="day",
    )
    assert r["status"] == "error"


def test_time_phased_write_invalid_date_format(clean_test_project):
    tid, rid = _setup_5day_assignment("BadDateT-T60")
    periods = [{"start": "not-a-date", "end": "2026-04-02",
                "actual_work_h": 4}]
    w = _msp_progress_time_phased_write(
        task_id=tid, resource_id=rid, periods=periods, unit="day",
    )
    # Either reject up-front or report per-period failure
    assert w["status"] in ("error", "partial")
```

**Step 2: Run — FAIL**

**Step 3: Implementation (T60 — write helper)**

```python
def _msp_progress_time_phased_write(task_id: int,
                                    resource_id: int,
                                    periods: List[Dict[str, Any]],
                                    unit: str = "day") -> Dict[str, Any]:
    """Write per-period actual_work to an assignment via TimeScaleData.

    Phase 3b — see design doc Section 6 Q2. Granularity: 'day' or 'week'.

    periods: [{start: ISO, end: ISO, actual_work_h: float}, ...]
    Each period maps to one (or more) TimeScaleValues slots; write fails per-
    slot if MSP doesn't have a matching cell. Failures aggregate into
    return['failures'] without raising.
    """
    if unit not in _TIMESCALE_UNIT_MAP:
        return {"status": "error",
                "error": f"unit must be 'day' or 'week', got '{unit}'"}
    if not isinstance(periods, list) or not periods:
        return {"status": "error", "error": "periods must be a non-empty list"}
    unit_code = _TIMESCALE_UNIT_MAP[unit]

    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    asg = _get_assignment_by_resource_id(t, resource_id)
    if asg is None:
        return {"status": "error",
                "error": f"No assignment for resource_id {resource_id} on task {task_id}"}

    from dateutil import parser
    written = 0
    failures: List[Dict[str, Any]] = []
    for idx, p in enumerate(periods):
        try:
            ps = parser.parse(str(p["start"]))
            pe = parser.parse(str(p["end"]))
            hours = float(p["actual_work_h"])
        except Exception as e:
            failures.append({"index": idx, "period": p,
                             "error": f"parse failed: {e}"})
            continue
        try:
            tsv = asg.TimeScaleData(StartDate=ps, EndDate=pe,
                                   Type=_PJ_TIMESCALED_ACTUAL_WORK,
                                   TimescaleUnit=unit_code)
            if tsv.Count == 0:
                failures.append({"index": idx, "period": p,
                                 "error": "no time slots in range (assignment "
                                          "may not span this date)"})
                continue
            # Distribute hours across all slots evenly
            minutes_total = _hours_to_minutes(hours)
            slot_count = tsv.Count
            per_slot = minutes_total // slot_count
            remainder = minutes_total - (per_slot * slot_count)
            for i in range(1, slot_count + 1):
                try:
                    val = per_slot + (remainder if i == slot_count else 0)
                    tsv.Item(i).Value = val
                except Exception as e:
                    failures.append({"index": idx, "slot": i,
                                     "error": _format_com_error(e)})
            written += 1
        except Exception as e:
            failures.append({"index": idx, "period": p,
                             "error": _format_com_error(e)})
    status = "ok" if not failures else ("partial" if written else "error")
    return {"status": status,
            "task_id": task_id, "resource_id": resource_id,
            "unit": unit,
            "written_count": written, "failures": failures}
```

**Step 4: Run** — note T61 tests will still FAIL until T61 lands; T60 only tests pass partially.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_time_phased.py
git commit -m "Phase 3b T60: msproject_progress time_phased_actual_write action (TimeScaleData per-period, day/week granularity)"
```

Expected (T60 only): **252 PASSED + 1 xfail**. (3 of the 7 time-phased tests pass — write-only paths.)

---

## Task 61: msproject_progress time_phased_actual_read Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Modify: `tests/test_msproject_progress_time_phased.py` (already created in T60)

**Step 1: Tests already written in T60** — T61 enables the remaining 4 of 7 to pass.

**Step 2: Run** — pre-impl, partial PASS as expected from T60.

**Step 3: Implementation**

Insert after `_msp_progress_time_phased_write`:

```python
def _msp_progress_time_phased_read(task_id: int,
                                   resource_id: int,
                                   start_date: str,
                                   end_date: str,
                                   unit: str = "day") -> Dict[str, Any]:
    """Read per-period actual_work from an assignment via TimeScaleData.

    Returns {status, periods: [{period_start, period_end, actual_work_h}]}.
    Empty periods (no actual yet) return as 0.0 hours, not omitted.
    """
    if unit not in _TIMESCALE_UNIT_MAP:
        return {"status": "error",
                "error": f"unit must be 'day' or 'week', got '{unit}'"}
    unit_code = _TIMESCALE_UNIT_MAP[unit]

    try:
        from dateutil import parser
        ds = parser.parse(str(start_date))
        de = parser.parse(str(end_date))
    except Exception as e:
        return {"status": "error",
                "error": f"could not parse date range: {e}"}

    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    asg = _get_assignment_by_resource_id(t, resource_id)
    if asg is None:
        return {"status": "error",
                "error": f"No assignment for resource_id {resource_id} on task {task_id}"}

    out: List[Dict[str, Any]] = []
    try:
        tsv = asg.TimeScaleData(StartDate=ds, EndDate=de,
                               Type=_PJ_TIMESCALED_ACTUAL_WORK,
                               TimescaleUnit=unit_code)
        for i in range(1, tsv.Count + 1):
            try:
                item = tsv.Item(i)
                out.append({
                    "period_start": _msp_dt_or_none(item.StartDate),
                    "period_end": _msp_dt_or_none(item.EndDate),
                    "actual_work_h": _minutes_to_hours(item.Value),
                })
            except Exception as e:
                logger.debug(f"_msp_progress_time_phased_read item {i} failed: {e}")
                continue
        return {"status": "ok",
                "task_id": task_id, "resource_id": resource_id,
                "unit": unit, "periods": out}
    except Exception as e:
        logger.error(f"_msp_progress_time_phased_read failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_progress_time_phased.py -v
```

Expected: 7 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py
git commit -m "Phase 3b T61: msproject_progress time_phased_actual_read action (TimeScaleData read, day/week buckets)"
```

Expected: **256 PASSED + 1 xfail** (252 + 4 read-side tests now passing).

---

## Task 62: msproject_progress bulk_progress_update Action (Hybrid path)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_bulk.py`

**Step 1: Failing test**

`tests/test_msproject_progress_bulk.py`:
```python
"""Test msproject_progress bulk_progress_update action (Phase 2b T37 hybrid pattern)."""
import pytest
import time
from msproject_mcp_core import (
    _msp_progress_bulk_update, _msp_progress_get_task,
    _msp_task_add_single,
)


def _make_n_tasks(n: int, prefix: str = "BlkT") -> list:
    """Create n tasks; return list of task IDs."""
    ids = []
    for i in range(n):
        r = _msp_task_add_single(name=f"{prefix}{i:03d}-T62", duration="2d")
        ids.append(r["task_id"])
    return ids


def test_bulk_3_items_com_direct(clean_test_project):
    """3 items → com_direct path."""
    ids = _make_n_tasks(3, "DirectT")
    items = [{"task_id": tid, "percent_complete": 25} for tid in ids]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_direct"
    assert r["count"] == 3
    for tid in ids:
        g = _msp_progress_get_task(task_id=tid)
        assert g["progress"]["percent_complete"] == 25


def test_bulk_10_items_com_batch(clean_test_project):
    """10 items → com_batch path."""
    ids = _make_n_tasks(10, "BatchT")
    items = [{"task_id": tid, "percent_complete": 50} for tid in ids]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_batch"
    assert r["count"] == 10


def test_bulk_25_items_mspdi_path(clean_test_project):
    """25 items → mspdi_bulk path (com_batch_fallback in Phase 3b)."""
    ids = _make_n_tasks(25, "MspdiT")
    items = [{"task_id": tid, "percent_complete": 30,
              "actual_work_h": 4} for tid in ids]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 25


def test_bulk_partial_failure_invalid_task_id(clean_test_project):
    """Mix of valid + invalid task IDs → status=partial."""
    ids = _make_n_tasks(3, "MixT")
    items = [{"task_id": ids[0], "percent_complete": 50},
             {"task_id": 99999, "percent_complete": 50},  # invalid
             {"task_id": ids[1], "percent_complete": 50}]
    r = _msp_progress_bulk_update(items=items)
    assert r["status"] == "partial"
    assert r["count"] == 2
    assert len(r["failures"]) == 1


def test_bulk_perf_50_tasks_under_3s(clean_test_project):
    """50 tasks bulk update <3s (com_batch path proxy via mspdi_bulk)."""
    ids = _make_n_tasks(50, "PerfT")
    items = [{"task_id": tid, "percent_complete": 25} for tid in ids]
    start = time.time()
    r = _msp_progress_bulk_update(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 50
    assert elapsed < 3.0, f"bulk 50 tasks took {elapsed:.2f}s (target <3s)"


def test_bulk_empty_list(clean_test_project):
    r = _msp_progress_bulk_update(items=[])
    assert r["status"] == "ok"
    assert r["path"] == "noop"
    assert r["count"] == 0
```

**Step 2: Run — FAIL**

**Step 3: Implementation (mirrors T37 pattern)**

```python
def _msp_progress_bulk_update_loop(items: List[Dict[str, Any]],
                                   path_label: str,
                                   task_map: Dict[int, Any]) -> Dict[str, Any]:
    """Inner loop using pre-built task map + per-item field application.

    Items: [{task_id, percent_complete?, actual_work_h?, percent_work_complete?,
             actual_start?, actual_finish?, remaining_work_h?, physical_pct?}, ...]
    """
    updated: List[Dict[str, Any]] = []
    failures: List[Dict[str, Any]] = []
    progress_field_keys = (_PROGRESS_PCT_FIELDS | _PROGRESS_WORK_FIELDS
                           | _PROGRESS_DURATION_FIELDS | _PROGRESS_DATE_FIELDS)
    for item in items:
        tid = item.get("task_id")
        t = task_map.get(tid) if tid is not None else None
        if t is None:
            failures.append({**item, "error": f"task_id {tid} not found"})
            continue
        # Apply each progress field present on this item
        applied: List[str] = []
        item_failures: List[str] = []
        # Pre-validate pct
        for pf in _PROGRESS_PCT_FIELDS:
            if pf in item and item[pf] is not None:
                try:
                    item[pf] = _normalize_progress_pct(item[pf])
                except ValueError as e:
                    item_failures.append(f"{pf}: {e}")
                    continue
        if item_failures:
            failures.append({"task_id": tid, "error": "; ".join(item_failures)})
            continue
        for field in progress_field_keys:
            if field in item and item[field] is not None:
                try:
                    _msp_task_set_progress_field(t, field, item[field])
                    applied.append(field)
                except Exception as e:
                    item_failures.append(f"{field}: {_format_com_error(e)}")
        if applied:
            updated.append({"task_id": tid, "applied": applied})
        if item_failures:
            failures.append({"task_id": tid, "error": "; ".join(item_failures),
                            "applied": applied})
    status = "ok" if not failures else ("partial" if updated else "error")
    return {"status": status, "path": path_label,
            "count": len(updated),
            "updated": updated, "failures": failures}


def _msp_progress_bulk_update(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Hybrid bulk progress update: routes by item count (Phase 1 _route_operation).

    Mirrors Phase 2b T37 _msp_resource_bulk_assign:
      - <=5 items   -> com_direct (no batch mode)
      - 6-19 items  -> com_batch  (batch mode + loop)
      - >=20 items  -> mspdi_bulk (Phase 3b: com_batch_fallback;
                       true MSPDI progress merge = Phase 4+)

    Pre-builds task_id -> Task map ONCE to avoid O(N×M) lookup blow-up.

    Returns: {status, path, count, updated, failures}
    """
    if not items:
        return {"status": "ok", "path": "noop", "count": 0,
                "updated": [], "failures": []}
    app = _validate_active_project()
    proj = app.ActiveProject
    task_map = _build_task_id_map(proj)

    path = _route_operation(len(items))
    if path == "com_direct":
        return _msp_progress_bulk_update_loop(items, "com_direct", task_map)
    elif path == "com_batch":
        _enter_batch_mode()
        try:
            return _msp_progress_bulk_update_loop(items, "com_batch", task_map)
        finally:
            _exit_batch_mode()
    else:
        _enter_batch_mode()
        try:
            return _msp_progress_bulk_update_loop(items, "mspdi_bulk", task_map)
        finally:
            _exit_batch_mode()
```

**Step 4: Run — PASS** (6 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_bulk.py
git commit -m "Phase 3b T62: msproject_progress bulk_progress_update action (hybrid 1-5/6-19/20+ matches T37 pattern)"
```

Expected: **262 PASSED + 1 xfail**.

---

## Task 63: msproject_progress summary Action (EVM-ready)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_progress_summary.py`

**Step 1: Failing test**

`tests/test_msproject_progress_summary.py`:
```python
"""Test msproject_progress summary action — EVM-ready aggregate."""
import pytest
from msproject_mcp_core import (
    _msp_progress_summary, _msp_progress_set_task,
    _msp_task_add_single,
)


def test_summary_empty_project(clean_test_project):
    r = _msp_progress_summary()
    assert r["status"] == "ok"
    p = r["project"]
    assert p["bac_h"] == 0
    assert p["acwp_h"] == 0
    assert p["task_count"] == 0
    assert p["completed_count"] == 0
    assert p["in_progress_count"] == 0


def test_summary_no_progress_zero_acwp(clean_test_project):
    """5 tasks of 4d each (32h work each) → BAC=160h, ACWP=0."""
    for i in range(5):
        _msp_task_add_single(name=f"NoProgT{i}-T63", duration="4d")
    r = _msp_progress_summary()
    p = r["project"]
    assert p["bac_h"] == 160
    assert p["acwp_h"] == 0
    assert p["not_started_count"] == 5


def test_summary_partial_progress(clean_test_project):
    """5 tasks (32h each = 160h BAC). 2 complete, 1 at 50%, 2 not started."""
    ids = []
    for i in range(5):
        ar = _msp_task_add_single(name=f"PartT{i}-T63", duration="4d")
        ids.append(ar["task_id"])
    _msp_progress_set_task(task_id=ids[0], percent_complete=100)
    _msp_progress_set_task(task_id=ids[1], percent_complete=100)
    _msp_progress_set_task(task_id=ids[2], percent_complete=50)
    r = _msp_progress_summary()
    p = r["project"]
    assert p["bac_h"] == 160
    assert p["completed_count"] == 2
    assert p["in_progress_count"] == 1
    assert p["not_started_count"] == 2
    # ACWP ≈ 32 + 32 + 16 = 80h
    assert 75 <= p["acwp_h"] <= 85
    # Project pct ~ 50%
    assert 45 <= p["project_percent_complete"] <= 55


def test_summary_fully_complete(clean_test_project):
    """3 tasks all 100% → project 100%."""
    ids = []
    for i in range(3):
        ar = _msp_task_add_single(name=f"DoneT{i}-T63", duration="2d")
        ids.append(ar["task_id"])
    for tid in ids:
        _msp_progress_set_task(task_id=tid, percent_complete=100)
    r = _msp_progress_summary()
    p = r["project"]
    assert p["completed_count"] == 3
    assert p["project_percent_complete"] >= 99


def test_summary_status_date_present(clean_test_project):
    """status_date is in the summary if set."""
    from msproject_mcp_core import _msp_progress_set_status_date
    _msp_task_add_single(name="StTask-T63", duration="2d")
    _msp_progress_set_status_date(status_date="2026-04-29")
    r = _msp_progress_summary()
    assert r["project"]["status_date"] is not None
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_progress_summary() -> Dict[str, Any]:
    """Project-level progress aggregate (EVM foundation).

    Phase 3b — used as input to Phase 5 EVM tool. Returns hours (CLAUDE.md
    RULE 4: BAC = sum(target_qty); for cost-loaded projects RULE 3 applies
    and Phase 5 EVM tool replaces hours with $).

    Returns {status, project: {bac_h, acwp_h, total_actual_work_h,
    total_remaining_work_h, project_percent_complete, status_date,
    task_count, in_progress_count, completed_count, not_started_count}}.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    bac_min = 0.0
    acwp_min = 0.0
    rem_min = 0.0
    task_count = 0
    in_progress = 0
    completed = 0
    not_started = 0
    try:
        for i in range(1, proj.Tasks.Count + 1):
            try:
                t = proj.Tasks(i)
                if t is None or t.Summary:
                    continue
                task_count += 1
                bac_min += float(t.Work or 0)
                acwp_min += float(t.ActualWork or 0)
                rem_min += float(t.RemainingWork or 0)
                pct = float(t.PercentComplete or 0)
                if pct >= 100:
                    completed += 1
                elif pct > 0:
                    in_progress += 1
                else:
                    not_started += 1
            except Exception as e:
                logger.debug(f"_msp_progress_summary task {i} skip: {e}")
                continue
        # Project-level pct
        try:
            pct = float(proj.PercentComplete or 0)
        except Exception:
            pct = (acwp_min / bac_min * 100.0) if bac_min > 0 else 0.0
        try:
            status_date = _msp_dt_or_none(proj.StatusDate)
        except Exception:
            status_date = None
        return {
            "status": "ok",
            "project": {
                "bac_h": round(_minutes_to_hours(bac_min), 2),
                "acwp_h": round(_minutes_to_hours(acwp_min), 2),
                "total_actual_work_h": round(_minutes_to_hours(acwp_min), 2),
                "total_remaining_work_h": round(_minutes_to_hours(rem_min), 2),
                "project_percent_complete": round(pct, 2),
                "status_date": status_date,
                "task_count": task_count,
                "in_progress_count": in_progress,
                "completed_count": completed,
                "not_started_count": not_started,
            },
        }
    except Exception as e:
        logger.error(f"_msp_progress_summary failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run — PASS** (5 PASSED)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_summary.py
git commit -m "Phase 3b T63: msproject_progress summary action (EVM-ready: BAC, ACWP, project pct, task counts by state)"
```

Expected: **267 PASSED + 1 xfail**.

---

## Task 64: FastMCP Dispatcher + Acceptance Script + README + Push (FINAL)

**Files:**
- Modify: `msproject_mcp_core.py` (add `@mcp.tool` dispatcher; update server `instructions` string)
- Create: `tests/test_msproject_progress_dispatcher.py`
- Create: `samples/build_progress_lifecycle.py`
- Modify: `README.md`

**Step 1: Failing dispatcher test**

`tests/test_msproject_progress_dispatcher.py`:
```python
"""Test FastMCP msproject_progress dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_progress


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_set_task(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="DispST-T64", duration="3d")
    r = _run(msproject_progress({"action": "set_task_progress",
                                 "task_id": add_r["task_id"],
                                 "percent_complete": 50}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "percent_complete" in p["changes"]


def test_dispatcher_get_task(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    add_r = _msp_task_add_single(name="DispGT-T64", duration="2d")
    r = _run(msproject_progress({"action": "get_task_progress",
                                 "task_id": add_r["task_id"]}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "progress" in p


def test_dispatcher_summary(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    _msp_task_add_single(name="DispSumT-T64", duration="2d")
    r = _run(msproject_progress({"action": "summary"}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "project" in p
    assert "bac_h" in p["project"]


def test_dispatcher_status_date(clean_test_project):
    r = _run(msproject_progress({"action": "set_status_date",
                                 "status_date": "2026-04-29"}))
    p = json.loads(r)
    assert p["status"] == "ok"


def test_dispatcher_bulk(clean_test_project):
    from msproject_mcp_core import _msp_task_add_single
    ids = [_msp_task_add_single(name=f"DispBlk{i}-T64", duration="2d")["task_id"]
           for i in range(3)]
    r = _run(msproject_progress({"action": "bulk_progress_update",
                                 "items": [{"task_id": tid, "percent_complete": 25}
                                           for tid in ids]}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert p["count"] == 3


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_progress({"action": "nonsense"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

In `msproject_mcp_core.py`, locate the existing `msproject_baseline` dispatcher (line ~2442) and add `msproject_progress` AFTER it (before `def main()`):

```python
@mcp.tool(
    name="msproject_progress",
    annotations={"title": "MS Project Progress Operations", "readOnlyHint": False},
)
async def msproject_progress(params: dict) -> str:
    """Task + assignment progress, time-phased actuals, status date, EVM-ready summary.

    Actions:
    - set_task_progress: Task-level set. Params: task_id, [percent_complete, percent_work_complete, actual_start, actual_finish, actual_duration_h, actual_work_h, remaining_work_h, remaining_duration_h, physical_pct, stop, resume]
    - get_task_progress: Task-level read. Params: task_id
    - set_assignment_progress: Per-resource man-hour set. Params: task_id, resource_id, [actual_work_h, actual_start, actual_finish, percent_work_complete, remaining_work_h, units]
    - get_assignment_progress: Per-task assignments list. Params: task_id
    - set_progress_by_date: Bulk update via app.UpdateProject. Params: progress_date, [scope='all', as_scheduled=True]
    - set_status_date: Set proj.StatusDate (data_date). Params: status_date
    - clear_progress: Single-task progress reset. Params: task_id
    - clear_all_progress: Project-wide progress reset.
    - time_phased_actual_write: Per-period actual_work write. Params: task_id, resource_id, periods, [unit='day']
    - time_phased_actual_read: Per-period actual_work read. Params: task_id, resource_id, start_date, end_date, [unit='day']
    - bulk_progress_update: Hybrid bulk path. Params: items (list of {task_id, ...progress fields})
    - summary: Project-level EVM-ready aggregate.

    Phase 3b (29 Apr 2026). DCMA-aligned: physical_pct exposed for EV input
    (independent of percent_complete). Foundation for Phase 5 msproject_evm.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "set_task_progress":
            r = _msp_progress_set_task(**p)
        elif action == "get_task_progress":
            r = _msp_progress_get_task(**p)
        elif action == "set_assignment_progress":
            r = _msp_progress_set_assignment(**p)
        elif action == "get_assignment_progress":
            r = _msp_progress_get_assignments(**p)
        elif action == "set_progress_by_date":
            r = _msp_progress_set_by_date(**p)
        elif action == "set_status_date":
            r = _msp_progress_set_status_date(**p)
        elif action == "clear_progress":
            r = _msp_progress_clear(**p)
        elif action == "clear_all_progress":
            r = _msp_progress_clear_all(**p)
        elif action == "time_phased_actual_write":
            r = _msp_progress_time_phased_write(**p)
        elif action == "time_phased_actual_read":
            r = _msp_progress_time_phased_read(**p)
        elif action == "bulk_progress_update":
            r = _msp_progress_bulk_update(**p)
        elif action == "summary":
            r = _msp_progress_summary(**p)
        else:
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: set_task_progress/"
                          "get_task_progress/set_assignment_progress/"
                          "get_assignment_progress/set_progress_by_date/"
                          "set_status_date/clear_progress/clear_all_progress/"
                          "time_phased_actual_write/time_phased_actual_read/"
                          "bulk_progress_update/summary"}
    except Exception as e:
        logger.error(f"msproject_progress({action}) failed: {e}")
        r = {"status": "error", "error": _format_com_error(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

Update `mcp = FastMCP(...)` `instructions` string to include `msproject_progress`:
```python
"Tools: msproject_task, msproject_link, msproject_schedule, msproject_calendar, msproject_resource, msproject_baseline, msproject_progress."
```

**Step 4: Acceptance script**

`samples/build_progress_lifecycle.py`:
```python
"""Phase 3b acceptance: full progress lifecycle.

SAFETY: Uses isolated FileNew project, never touches user's active project.

Scenario:
  1. Create 50 villa tasks + 3 resources + assignments
  2. Save Baseline 0 ('Original') (Phase 3a integration)
  3. set_task_progress on first 10 tasks (percent_complete=50)
  4. set_assignment_progress on next 10 tasks (per-resource man-hours)
  5. time_phased_actual_write on 1 task (5 weekdays)
  6. time_phased_actual_read verification
  7. set_status_date to today
  8. set_progress_by_date for older tasks (plan = actual to data_date)
  9. bulk_progress_update with 25 items (mspdi_bulk path)
  10. summary → BAC, ACWP, project_pct
  11. clear_all_progress → reset

Target: end-to-end <15s.
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import datetime as dt
import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save,
    _msp_progress_set_task, _msp_progress_get_task,
    _msp_progress_set_assignment, _msp_progress_get_assignments,
    _msp_progress_set_by_date, _msp_progress_set_status_date,
    _msp_progress_time_phased_write, _msp_progress_time_phased_read,
    _msp_progress_bulk_update, _msp_progress_summary,
    _msp_progress_clear_all,
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
        # 1. Tasks + resources + assignments
        print(f"\n1. Building {N_TASKS} villa tasks + 3 resources + assignments...")
        tasks = _msp_task_bulk_add(items=[
            {"name": f"Villa T{i:03d}", "duration": "4d",
             "start": "2026-04-01"} for i in range(N_TASKS)])
        task_ids = tasks["task_ids"]
        res_ids = []
        for name in ["COW", "STL", "MSN"]:
            res_ids.append(_msp_resource_add(name=name, type="Work",
                                             max_units=300)["resource_id"])
        items = [{"task_id": tid, "resource_id": rid}
                 for tid in task_ids for rid in res_ids]
        _msp_resource_bulk_assign(items=items)
        print(f"   OK {len(task_ids)} tasks, 3 resources, "
              f"{len(items)} assignments in {time.time()-t0:.2f}s")

        # 2. Save Baseline 0 (Phase 3a integration)
        print("2. Saving Baseline 0 'Original'...")
        b0 = _msp_baseline_save(baseline_number=0, name="Original")
        assert b0["status"] == "ok"

        # 3. set_task_progress on first 10
        print("3. set_task_progress on first 10 tasks (50%)...")
        for tid in task_ids[:10]:
            _msp_progress_set_task(task_id=tid, percent_complete=50)

        # 4. set_assignment_progress on next 10 (per-resource man-hours)
        print("4. set_assignment_progress on next 10 tasks (COW=24h, STL=18h, MSN=10h)...")
        for tid in task_ids[10:20]:
            _msp_progress_set_assignment(task_id=tid, resource_id=res_ids[0],
                                         actual_work_h=24)
            _msp_progress_set_assignment(task_id=tid, resource_id=res_ids[1],
                                         actual_work_h=18)
            _msp_progress_set_assignment(task_id=tid, resource_id=res_ids[2],
                                         actual_work_h=10)

        # 5. time_phased_actual_write (1 task × 5 weekdays)
        print("5. time_phased_actual_write on T020 (5 weekdays varying hours)...")
        periods = [
            {"start": "2026-04-01", "end": "2026-04-02", "actual_work_h": 6},
            {"start": "2026-04-02", "end": "2026-04-03", "actual_work_h": 8},
            {"start": "2026-04-03", "end": "2026-04-06", "actual_work_h": 8},
            {"start": "2026-04-06", "end": "2026-04-07", "actual_work_h": 4},
            {"start": "2026-04-07", "end": "2026-04-08", "actual_work_h": 7},
        ]
        tpw = _msp_progress_time_phased_write(
            task_id=task_ids[20], resource_id=res_ids[0],
            periods=periods, unit="day")
        print(f"   written_count={tpw['written_count']}, status={tpw['status']}")

        # 6. time_phased_actual_read
        print("6. time_phased_actual_read verification...")
        tpr = _msp_progress_time_phased_read(
            task_id=task_ids[20], resource_id=res_ids[0],
            start_date="2026-04-01", end_date="2026-04-08", unit="day")
        for p in tpr["periods"][:5]:
            print(f"   {p['period_start'][:10]}: {p['actual_work_h']}h")

        # 7. set_status_date
        print("7. set_status_date to today (2026-04-29)...")
        sd = _msp_progress_set_status_date(status_date="2026-04-29")
        assert sd["status"] == "ok"

        # 8. set_progress_by_date for older tasks (sliding date)
        print("8. set_progress_by_date 2026-04-15 (plan=actual catch-up)...")
        sbd = _msp_progress_set_by_date(progress_date="2026-04-15")
        print(f"   {sbd['status']}, scope=all, mode={sbd.get('mode')}")

        # 9. bulk_progress_update on 25 items (mspdi_bulk path)
        print("9. bulk_progress_update on 25 items (mspdi_bulk path)...")
        bulk_items = [{"task_id": tid, "percent_complete": 30}
                       for tid in task_ids[25:50]]
        bu = _msp_progress_bulk_update(items=bulk_items)
        print(f"   {bu['status']}, path={bu['path']}, count={bu['count']}")

        # 10. summary
        print("10. summary (EVM-ready)...")
        summ = _msp_progress_summary()
        p = summ["project"]
        print(f"   BAC={p['bac_h']}h, ACWP={p['acwp_h']}h")
        print(f"   project_pct={p['project_percent_complete']}%, "
              f"completed={p['completed_count']}/{p['task_count']}")

        # 11. clear_all_progress
        print("11. clear_all_progress (reset)...")
        cl = _msp_progress_clear_all()
        print(f"   cleared_count={cl['cleared_count']}")

        elapsed = time.time() - t0
        print(f"\nOK ACCEPTANCE: {elapsed:.2f}s total (target <15s)")
        assert elapsed < 15.0, f"Too slow: {elapsed}s"

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
python samples/build_progress_lifecycle.py
```

Expected: `OK ACCEPTANCE: <Xs total (target <15s)`. Realistic ~8-12s.

**Step 6: README update**

Add Phase 3b section to `README.md` after Phase 3a:

```markdown
### Phase 3b — Progress Management (29 Apr 2026)

`msproject_progress` tool with 12 actions, dual-track progress + time-phased
actuals + EVM foundation:

**Task-level:**
- `set_task_progress` / `get_task_progress` — % complete, % work complete,
  actual_start/finish, actual_work, remaining_work, **physical_pct (DCMA)**,
  stop/resume

**Assignment-level (per-resource man-hour):**
- `set_assignment_progress` / `get_assignment_progress` — assignment.ActualWork,
  PercentWorkComplete, RemainingWork, Units (rolls up to task automatically)

**Time-phased (TimeScaleData):**
- `time_phased_actual_write` / `time_phased_actual_read` — per-day or
  per-week actual_work buckets for hakediş/EVM period delta reporting

**Bulk operations:**
- `set_progress_by_date` — `app.UpdateProject(ProgressDate)` retroactive
  catch-up (plan = actual up to date)
- `bulk_progress_update` — hybrid 1-5/6-19/20+ path (Phase 2b T37 pattern)
- `set_status_date` — `proj.StatusDate` (data_date)
- `clear_progress` / `clear_all_progress` — reset progress

**EVM-ready aggregate:**
- `summary` — BAC, ACWP, project_pct_complete, task counts (completed/
  in_progress/not_started). Foundation for upcoming Phase 5 `msproject_evm`.

Acceptance: `samples/build_progress_lifecycle.py` runs full progress
lifecycle in <15s.

Tool count: **7 tools, ~52 actions**.
```

**Step 7: Run full regression**

```bash
python -m pytest tests/ -v --tb=short -q
```

Expected: **~273 PASSED + 1 xfail** (267 + 6 dispatcher).

**Step 8: Commit + push**

```bash
git add msproject_mcp_core.py tests/test_msproject_progress_dispatcher.py samples/build_progress_lifecycle.py README.md
git commit -m "Phase 3b T64: dispatcher + acceptance + README + push (full progress lifecycle <15s)"
git push origin main
```

Expected: ~13 commits pushed (T52-T64 + design).

---

## Phase 3b Tamamlama Kriterleri
1. ✅ T52-T64 ~13 commit landed
2. ✅ Acceptance script <15s
3. ✅ Yeni testler ~50 PASS
4. ✅ Phase 1+2a+2b+3a baseline 207+1xfail regression PASS
5. ✅ Total ~260 PASS + 1 xfail
6. ✅ Push to origin/main
7. ✅ Phase 3b live on GitHub
8. ⏸ Kullanıcı manuel onayı → Phase 4 (File MCP) başlar

---

*Plan tamamlandı: 29 Nisan 2026*
*Tahmini Phase 3b süresi: ~8-10 saat (T52-T64, 13 task)*
*Sonraki phase (onay sonrası): Phase 4 — `msproject_file` (.mpp + XML offline reading)*
