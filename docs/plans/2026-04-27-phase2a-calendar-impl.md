# MS Project MCP — Phase 2a Calendar Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** `msproject_calendar` tool — 7 actions (create, update, add_exception, assign_to_task, assign_to_resource, list, holidays_uzbek). Built-in 9 Özbekistan 2026 resmi tatili. Phase 1 SAFETY pattern korunur, kullanıcının aktif projesi DOKUNULMAZ.

**Architecture:** COM-first (no MSPDI bulk for Calendar — volume düşük). Tüm helper'lar mevcut `msproject_mcp_core.py`'a eklenir. Tüm testler `clean_test_project` fixture ile izole proje açar (FileNew → test → FileClose 0). Custom calendar'lar project-scoped, izole projeyle birlikte otomatik wipe edilir.

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM, pytest. Mevcut `msproject_mcp_core.py` (~940 satır), `tests/conftest.py` (`clean_test_project` fixture).

**Design doc:** `docs/plans/2026-04-27-phase2a-calendar-design.md` (commit `f1dd1c9`)

---

## Task 18: Uzbek Holidays Constant + Calendar Helper

**Files:**
- Modify: `msproject_mcp_core.py` (yeni constant + helper, top-level)
- Create: `tests/test_msproject_calendar_helpers.py`

**Step 1: Failing test yaz**

`tests/test_msproject_calendar_helpers.py`:
```python
"""Test calendar helpers + UZBEK_HOLIDAYS constant."""
import pytest
from msproject_mcp_core import UZBEK_HOLIDAYS_2026, _find_calendar_by_name


def test_uzbek_holidays_count():
    """9 official Uzbek holidays."""
    assert len(UZBEK_HOLIDAYS_2026) == 9


def test_uzbek_holidays_structure():
    """Each entry: (name, month, day) tuple."""
    for entry in UZBEK_HOLIDAYS_2026:
        assert len(entry) == 3
        name, month, day = entry
        assert isinstance(name, str) and len(name) > 0
        assert 1 <= month <= 12
        assert 1 <= day <= 31


def test_uzbek_holidays_includes_navruz():
    """Navruz (March 21) must be in list."""
    found = [e for e in UZBEK_HOLIDAYS_2026 if e[1] == 3 and e[2] == 21]
    assert len(found) == 1
    assert "Navruz" in found[0][0]


def test_find_calendar_standard(clean_test_project):
    """_find_calendar_by_name finds 'Standard' in any project."""
    proj = clean_test_project
    cal = _find_calendar_by_name(proj, "Standard")
    assert cal is not None
    assert cal.Name == "Standard"


def test_find_calendar_missing(clean_test_project):
    """Returns None for missing name."""
    proj = clean_test_project
    cal = _find_calendar_by_name(proj, "NonExistent-XYZ")
    assert cal is None
```

**Step 2: Run test — expect FAIL (constant + helper undefined)**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/test_msproject_calendar_helpers.py -v
```

Expected: ImportError on UZBEK_HOLIDAYS_2026.

**Step 3: Add constant + helper to `msproject_mcp_core.py`**

Find a logical spot after existing helpers (near `_find_task_by_id`). Add:

```python
# ---------- CALENDAR CONSTANTS ----------

UZBEK_HOLIDAYS_2026 = [
    ("Yılbaşı", 1, 1),
    ("Vatan Müdafaası Günü", 1, 14),
    ("Kadınlar Günü", 3, 8),
    ("Navruz", 3, 21),
    ("İşçi Bayramı", 5, 1),
    ("Hatıra ve Şeref Günü", 5, 9),
    ("Bağımsızlık Günü", 9, 1),
    ("Öğretmenler Günü", 10, 1),
    ("Anayasa Günü", 12, 8),
]


# ---------- CALENDAR HELPERS ----------

def _find_calendar_by_name(proj: Any, name: str) -> Optional[Any]:
    """Locate a base calendar object in the project. Returns None if not found."""
    try:
        for i in range(1, proj.BaseCalendars.Count + 1):
            cal = proj.BaseCalendars(i)
            if cal is not None and cal.Name == name:
                return cal
    except Exception:
        pass
    return None
```

**Step 4: Run test — expect PASS**

```bash
python -m pytest tests/test_msproject_calendar_helpers.py -v
```

Expected: 5 PASSED (3 unit + 2 integration).

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_helpers.py
git commit -m "Phase 2a T18: UZBEK_HOLIDAYS_2026 constant + _find_calendar_by_name helper"
```

---

## Task 19: `msproject_calendar` create Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_calendar_create.py`

**Step 1: Failing test**

`tests/test_msproject_calendar_create.py`:
```python
"""Test msproject_calendar create action."""
import pytest
from msproject_mcp_core import _msp_calendar_create, _find_calendar_by_name


def test_create_from_standard(clean_test_project):
    """Create 'TestCal' calendar from Standard base."""
    proj = clean_test_project
    r = _msp_calendar_create(name="TestCal-Phase2a", base_calendar="Standard")
    assert r["status"] == "ok"
    assert r["name"] == "TestCal-Phase2a"
    cal = _find_calendar_by_name(proj, "TestCal-Phase2a")
    assert cal is not None


def test_create_duplicate_name_errors(clean_test_project):
    """Creating a calendar with existing name should error."""
    _msp_calendar_create(name="DupCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_create(name="DupCal-Phase2a", base_calendar="Standard")
    assert r["status"] == "error"
    assert "already exists" in r["error"].lower()


def test_create_missing_base_errors(clean_test_project):
    """Base calendar that doesn't exist must error cleanly."""
    r = _msp_calendar_create(name="X-Phase2a", base_calendar="NonExistentBase")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()
```

**Step 2: Run — FAIL**

```bash
python -m pytest tests/test_msproject_calendar_create.py -v
```

Expected: ImportError on `_msp_calendar_create`.

**Step 3: Implementation**

Add to `msproject_mcp_core.py` after the helpers from T18:

```python
def _msp_calendar_create(name: str, base_calendar: str = "Standard") -> Dict[str, Any]:
    """Create a new project-scoped base calendar copied from an existing one."""
    app = _validate_active_project()
    proj = app.ActiveProject
    # Pre-flight: name conflict
    if _find_calendar_by_name(proj, name) is not None:
        return {"status": "error", "error": f"Calendar '{name}' already exists"}
    # Pre-flight: base must exist
    if _find_calendar_by_name(proj, base_calendar) is None:
        return {"status": "error",
                "error": f"Base calendar '{base_calendar}' not found in project"}
    try:
        # app.BaseCalendarCreate creates a project-scoped base calendar
        app.BaseCalendarCreate(Name=name, FromName=base_calendar)
        cal = _find_calendar_by_name(proj, name)
        if cal is None:
            return {"status": "error",
                    "error": f"BaseCalendarCreate succeeded but '{name}' not found"}
        return {"status": "ok", "calendar_uid": cal.UniqueID, "name": name}
    except Exception as e:
        logger.error(f"_msp_calendar_create({name}) failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_create.py -v
```

Expected: 3 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_create.py
git commit -m "Phase 2a T19: msproject_calendar create action"
```

---

## Task 20: `msproject_calendar` update Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_calendar_update.py`

**Step 1: Failing test**

`tests/test_msproject_calendar_update.py`:
```python
"""Test msproject_calendar update action."""
import pytest
from msproject_mcp_core import _msp_calendar_create, _msp_calendar_update, _find_calendar_by_name


def test_update_rename(clean_test_project):
    """Rename a calendar."""
    proj = clean_test_project
    _msp_calendar_create(name="OldName-Phase2a", base_calendar="Standard")
    r = _msp_calendar_update(name="OldName-Phase2a", new_name="NewName-Phase2a")
    assert r["status"] == "ok"
    assert "name" in r["changes"]
    assert _find_calendar_by_name(proj, "NewName-Phase2a") is not None
    assert _find_calendar_by_name(proj, "OldName-Phase2a") is None


def test_update_weekday_off(clean_test_project):
    """Set Sunday (weekday=1 in MSP) as non-working."""
    proj = clean_test_project
    _msp_calendar_create(name="WeekdayCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_update(name="WeekdayCal-Phase2a", weekday_off=1)
    assert r["status"] == "ok"
    assert "weekday_off" in r["changes"]
    cal = _find_calendar_by_name(proj, "WeekdayCal-Phase2a")
    sunday = cal.WeekDays(1)
    assert sunday.Working is False


def test_update_missing_calendar_errors(clean_test_project):
    r = _msp_calendar_update(name="DoesNotExist-Phase2a", new_name="X")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()
```

**Step 2: Run — FAIL**

Expected: ImportError.

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
def _msp_calendar_update(name: str,
                         new_name: Optional[str] = None,
                         weekday_off: Optional[int] = None) -> Dict[str, Any]:
    """Rename a calendar and/or mark a weekday as non-working.

    weekday_off: 1=Sunday, 2=Monday, ..., 7=Saturday (MS Project convention).
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    cal = _find_calendar_by_name(proj, name)
    if cal is None:
        return {"status": "error", "error": f"Calendar '{name}' not found in project"}
    changes = []
    try:
        if new_name is not None and new_name != name:
            if _find_calendar_by_name(proj, new_name) is not None:
                return {"status": "error",
                        "error": f"Calendar '{new_name}' already exists"}
            cal.Name = new_name
            changes.append("name")
        if weekday_off is not None:
            if not (1 <= weekday_off <= 7):
                return {"status": "error",
                        "error": "weekday_off must be 1-7 (1=Sunday, 7=Saturday)"}
            wd = cal.WeekDays(weekday_off)
            wd.Working = False
            changes.append("weekday_off")
        return {"status": "ok", "calendar_name": new_name or name, "changes": changes}
    except Exception as e:
        logger.error(f"_msp_calendar_update({name}) failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_update.py -v
```

Expected: 3 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_update.py
git commit -m "Phase 2a T20: msproject_calendar update (rename + weekday_off)"
```

---

## Task 21: `msproject_calendar` add_exception Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_calendar_exception.py`

**Step 1: Failing test**

`tests/test_msproject_calendar_exception.py`:
```python
"""Test msproject_calendar add_exception action."""
import pytest
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_add_exception, _find_calendar_by_name,
)


def test_add_single_date_exception(clean_test_project):
    """Add a single-day non-working exception."""
    proj = clean_test_project
    _msp_calendar_create(name="ExCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="ExCal-Phase2a",
        exception_name="New Year",
        start="2026-01-01",
    )
    assert r["status"] == "ok"
    cal = _find_calendar_by_name(proj, "ExCal-Phase2a")
    # Verify the exception was added (Exceptions.Count >= 1)
    assert cal.Exceptions.Count >= 1


def test_add_date_range_exception(clean_test_project):
    """Add an exception spanning multiple days."""
    proj = clean_test_project
    _msp_calendar_create(name="RangeCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="RangeCal-Phase2a",
        exception_name="Spring Break",
        start="2026-03-23",
        finish="2026-03-27",
    )
    assert r["status"] == "ok"


def test_add_exception_missing_calendar_errors(clean_test_project):
    r = _msp_calendar_add_exception(
        calendar_name="NoSuchCal-Phase2a",
        exception_name="Holiday",
        start="2026-01-01",
    )
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_add_exception_invalid_date_range(clean_test_project):
    """Start > finish should error."""
    _msp_calendar_create(name="BadRangeCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_add_exception(
        calendar_name="BadRangeCal-Phase2a",
        exception_name="Bad Range",
        start="2026-05-10",
        finish="2026-05-01",
    )
    assert r["status"] == "error"
    assert "start" in r["error"].lower()
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
def _parse_date(date_str: str) -> _dt.date:
    """Parse 'YYYY-MM-DD' to date."""
    return _dt.datetime.strptime(date_str, "%Y-%m-%d").date()


# pjExceptionDaily = 7 (single fixed-date or range, non-recurring)
PJ_EXCEPTION_DAILY = 7


def _msp_calendar_add_exception(calendar_name: str, exception_name: str,
                                start: str,
                                finish: Optional[str] = None,
                                working: bool = False) -> Dict[str, Any]:
    """Add a non-working exception to a calendar (single date or range).

    For Phase 2a: only non-recurring exceptions (Type=pjExceptionDaily=7).
    Recurring patterns (weekly/monthly) deferred to Phase 3+.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    cal = _find_calendar_by_name(proj, calendar_name)
    if cal is None:
        return {"status": "error", "error": f"Calendar '{calendar_name}' not found in project"}
    try:
        start_d = _parse_date(start)
        finish_d = _parse_date(finish) if finish else start_d
    except ValueError as e:
        return {"status": "error", "error": f"Invalid date format (expected YYYY-MM-DD): {e}"}
    if finish_d < start_d:
        return {"status": "error", "error": "Start date must be <= finish date"}
    try:
        ex = cal.Exceptions.Add(
            Type=PJ_EXCEPTION_DAILY,
            Start=pywintypes.Time(start_d),
            Finish=pywintypes.Time(finish_d),
        )
        ex.Name = exception_name
        # working=False is default for new exceptions in MSP; explicitly handle
        if working:
            # leave default working hours from the base
            pass
        else:
            # Mark non-working: set all Shift starts to None / set Working flag false
            try:
                ex.Shift1Start = 0
                ex.Shift1Finish = 0
                ex.Shift2Start = 0
                ex.Shift2Finish = 0
                ex.Shift3Start = 0
                ex.Shift3Finish = 0
            except Exception:
                pass
        return {"status": "ok",
                "calendar_name": calendar_name,
                "exception_name": exception_name,
                "start": start,
                "finish": finish or start,
                "working": working}
    except Exception as e:
        logger.error(f"_msp_calendar_add_exception({calendar_name},{exception_name}) failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_exception.py -v
```

Expected: 4 PASSED.

If MS Project rejects Shift1Start=0 (some COM versions need NULL), wrap each shift assignment in its own try/except — the exception itself with default shifts is already non-working when Type=7 is "exception" rather than "working time exception".

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_exception.py
git commit -m "Phase 2a T21: msproject_calendar add_exception (single + range, non-recurring)"
```

---

## Task 22: `msproject_calendar` assign_to_task Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_calendar_assign_task.py`

**Step 1: Failing test**

`tests/test_msproject_calendar_assign_task.py`:
```python
"""Test msproject_calendar assign_to_task action."""
import pytest
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_assign_to_task,
    _msp_task_add_single, _find_task_by_id,
)


def test_assign_calendar_to_task(clean_test_project):
    """Assign a custom calendar to an existing task."""
    proj = clean_test_project
    _msp_calendar_create(name="TaskCal-Phase2a", base_calendar="Standard")
    add_r = _msp_task_add_single(name="CalAssignTask", duration="3d")
    assert add_r["status"] == "ok"
    task_id = add_r["task_id"]
    r = _msp_calendar_assign_to_task(task_id=task_id, calendar_name="TaskCal-Phase2a")
    assert r["status"] == "ok"
    t = _find_task_by_id(proj, task_id)
    # MSP exposes task.Calendar as string name
    assert t.Calendar == "TaskCal-Phase2a"


def test_assign_missing_calendar_errors(clean_test_project):
    add_r = _msp_task_add_single(name="MissingCalTask", duration="1d")
    r = _msp_calendar_assign_to_task(task_id=add_r["task_id"], calendar_name="NoSuch-Phase2a")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_assign_missing_task_errors(clean_test_project):
    _msp_calendar_create(name="OrphanCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_assign_to_task(task_id=99999, calendar_name="OrphanCal-Phase2a")
    assert r["status"] == "error"
    assert "task" in r["error"].lower() and "99999" in r["error"]
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
def _msp_calendar_assign_to_task(task_id: int, calendar_name: str) -> Dict[str, Any]:
    """Assign a base calendar to a specific task."""
    app = _validate_active_project()
    proj = app.ActiveProject
    cal = _find_calendar_by_name(proj, calendar_name)
    if cal is None:
        return {"status": "error",
                "error": f"Calendar '{calendar_name}' not found in project"}
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    try:
        t.Calendar = calendar_name
        return {"status": "ok", "task_id": task_id, "calendar_name": calendar_name}
    except Exception as e:
        logger.error(f"_msp_calendar_assign_to_task({task_id},{calendar_name}) failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_assign_task.py -v
```

Expected: 3 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_assign_task.py
git commit -m "Phase 2a T22: msproject_calendar assign_to_task action"
```

---

## Task 23: `msproject_calendar` assign_to_resource Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_calendar_assign_resource.py`

**Note:** Resource management is Phase 2b. For T23 we test assignment via raw COM `proj.Resources.Add(name)` to keep the test self-contained. The high-level Resource tool comes in Phase 2b.

**Step 1: Failing test**

`tests/test_msproject_calendar_assign_resource.py`:
```python
"""Test msproject_calendar assign_to_resource action.

Phase 2a uses raw COM Resources.Add since the high-level Resource tool
arrives in Phase 2b. Once Phase 2b lands, we can refactor to use it.
"""
import pytest
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_assign_to_resource,
)


def _add_resource(proj, name: str) -> int:
    """Helper: add a Work resource via raw COM. Returns resource ID."""
    r = proj.Resources.Add(name)
    return r.ID


def test_assign_calendar_to_resource(clean_test_project):
    proj = clean_test_project
    _msp_calendar_create(name="ResCal-Phase2a", base_calendar="Standard")
    res_id = _add_resource(proj, "TestRes-Phase2a")
    r = _msp_calendar_assign_to_resource(resource_id=res_id, calendar_name="ResCal-Phase2a")
    assert r["status"] == "ok"
    # Verify
    res = None
    for i in range(1, proj.Resources.Count + 1):
        rr = proj.Resources(i)
        if rr is not None and rr.ID == res_id:
            res = rr
            break
    assert res is not None
    assert res.BaseCalendar == "ResCal-Phase2a"


def test_assign_missing_calendar_errors(clean_test_project):
    proj = clean_test_project
    res_id = _add_resource(proj, "OrphanRes-Phase2a")
    r = _msp_calendar_assign_to_resource(resource_id=res_id, calendar_name="NoSuch-Phase2a")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_assign_missing_resource_errors(clean_test_project):
    _msp_calendar_create(name="LonelyCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_assign_to_resource(resource_id=99999, calendar_name="LonelyCal-Phase2a")
    assert r["status"] == "error"
    assert "resource" in r["error"].lower() and "99999" in r["error"]
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
def _find_resource_by_id(proj: Any, resource_id: int) -> Optional[Any]:
    """Locate a resource by ID. Phase 2a helper; reused/expanded in Phase 2b."""
    try:
        for i in range(1, proj.Resources.Count + 1):
            r = proj.Resources(i)
            if r is not None and r.ID == resource_id:
                return r
    except Exception:
        pass
    return None


def _msp_calendar_assign_to_resource(resource_id: int, calendar_name: str) -> Dict[str, Any]:
    """Assign a base calendar to a resource."""
    app = _validate_active_project()
    proj = app.ActiveProject
    cal = _find_calendar_by_name(proj, calendar_name)
    if cal is None:
        return {"status": "error",
                "error": f"Calendar '{calendar_name}' not found in project"}
    res = _find_resource_by_id(proj, resource_id)
    if res is None:
        return {"status": "error", "error": f"Resource ID {resource_id} not found"}
    try:
        res.BaseCalendar = calendar_name
        return {"status": "ok", "resource_id": resource_id, "calendar_name": calendar_name}
    except Exception as e:
        logger.error(f"_msp_calendar_assign_to_resource({resource_id},{calendar_name}) failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_assign_resource.py -v
```

Expected: 3 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_assign_resource.py
git commit -m "Phase 2a T23: msproject_calendar assign_to_resource (raw COM, refactored in Phase 2b)"
```

---

## Task 24: `msproject_calendar` list Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_calendar_list.py`

**Step 1: Failing test**

`tests/test_msproject_calendar_list.py`:
```python
"""Test msproject_calendar list action."""
import pytest
from msproject_mcp_core import _msp_calendar_create, _msp_calendar_add_exception, _msp_calendar_list


def test_list_includes_standard(clean_test_project):
    """Default calendars (Standard, 24 Hours, Night Shift) listed."""
    r = _msp_calendar_list()
    assert r["status"] == "ok"
    names = [c["name"] for c in r["calendars"]]
    assert "Standard" in names


def test_list_includes_custom(clean_test_project):
    """Custom calendar appears in list with exceptions."""
    _msp_calendar_create(name="ListCal-Phase2a", base_calendar="Standard")
    _msp_calendar_add_exception(
        calendar_name="ListCal-Phase2a",
        exception_name="Test Holiday",
        start="2026-01-01",
    )
    r = _msp_calendar_list()
    assert r["status"] == "ok"
    custom = next((c for c in r["calendars"] if c["name"] == "ListCal-Phase2a"), None)
    assert custom is not None
    assert custom["exception_count"] >= 1
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
def _msp_calendar_list() -> Dict[str, Any]:
    """List all base calendars in the active project with exception counts."""
    app = _validate_active_project()
    proj = app.ActiveProject
    out = []
    try:
        for i in range(1, proj.BaseCalendars.Count + 1):
            cal = proj.BaseCalendars(i)
            if cal is None:
                continue
            try:
                ex_count = cal.Exceptions.Count
            except Exception:
                ex_count = 0
            out.append({
                "uid": cal.UniqueID,
                "name": cal.Name,
                "exception_count": ex_count,
            })
        return {"status": "ok", "count": len(out), "calendars": out}
    except Exception as e:
        logger.error(f"_msp_calendar_list failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_list.py -v
```

Expected: 2 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_list.py
git commit -m "Phase 2a T24: msproject_calendar list action"
```

---

## Task 25: `msproject_calendar` holidays_uzbek Action (Hero Feature)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_calendar_uzbek.py`

**Step 1: Failing test**

`tests/test_msproject_calendar_uzbek.py`:
```python
"""Test msproject_calendar holidays_uzbek action."""
import pytest
import time
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_holidays_uzbek,
    _find_calendar_by_name, UZBEK_HOLIDAYS_2026,
)


def test_uzbek_holidays_added(clean_test_project):
    """All 9 Uzbek holidays added to a fresh calendar in <2s."""
    proj = clean_test_project
    _msp_calendar_create(name="UzbekCal-Phase2a", base_calendar="Standard")
    start = time.time()
    r = _msp_calendar_holidays_uzbek(calendar_name="UzbekCal-Phase2a", year=2026)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 9
    assert elapsed < 2.0, f"holidays_uzbek took {elapsed:.2f}s (target <2s)"
    cal = _find_calendar_by_name(proj, "UzbekCal-Phase2a")
    assert cal.Exceptions.Count >= 9


def test_uzbek_holidays_dates_correct(clean_test_project):
    """Returned holiday dates match UZBEK_HOLIDAYS_2026 constant."""
    _msp_calendar_create(name="UzbekDateCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_holidays_uzbek(calendar_name="UzbekDateCal-Phase2a", year=2026)
    assert r["status"] == "ok"
    returned_dates = {(h["month"], h["day"]) for h in r["holidays"]}
    expected_dates = {(m, d) for _, m, d in UZBEK_HOLIDAYS_2026}
    assert returned_dates == expected_dates


def test_uzbek_holidays_missing_calendar_errors(clean_test_project):
    r = _msp_calendar_holidays_uzbek(calendar_name="NoSuch-Phase2a", year=2026)
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
def _msp_calendar_holidays_uzbek(calendar_name: str, year: int = 2026) -> Dict[str, Any]:
    """Bulk-add 9 official Özbekistan public holidays to a calendar.

    Uses UZBEK_HOLIDAYS_2026 constant. Year parameter shifts year only —
    the (month, day) pairs are fixed (Navruz=21 March, Independence=1 Sep, etc.).
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    if _find_calendar_by_name(proj, calendar_name) is None:
        return {"status": "error",
                "error": f"Calendar '{calendar_name}' not found in project"}
    added = []
    failed = []
    for name, month, day in UZBEK_HOLIDAYS_2026:
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
        r = _msp_calendar_add_exception(
            calendar_name=calendar_name,
            exception_name=name,
            start=date_str,
        )
        if r.get("status") == "ok":
            added.append({"name": name, "date": date_str, "month": month, "day": day})
        else:
            failed.append({"name": name, "date": date_str, "error": r.get("error")})
    if failed:
        logger.warning(f"holidays_uzbek partial: {len(added)} added, {len(failed)} failed")
    return {
        "status": "ok" if not failed else "partial",
        "calendar_name": calendar_name,
        "year": year,
        "count": len(added),
        "holidays": added,
        "failures": failed,
    }
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_uzbek.py -v
```

Expected: 3 PASSED. `<2s` performance assertion validates COM speed.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_uzbek.py
git commit -m "Phase 2a T25: holidays_uzbek action (9 Uzbek 2026 holidays bulk-add <2s)"
```

---

## Task 26: FastMCP Dispatcher (`msproject_calendar` Tool Wiring)

**Files:**
- Modify: `msproject_mcp_core.py` (add `@mcp.tool` decorator + dispatcher)
- Create: `tests/test_msproject_calendar_dispatcher.py`

**Step 1: Failing test**

`tests/test_msproject_calendar_dispatcher.py`:
```python
"""Test FastMCP msproject_calendar dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_calendar


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_create(clean_test_project):
    r = _run(msproject_calendar({
        "action": "create",
        "name": "DispCal-Phase2a",
        "base_calendar": "Standard",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"


def test_dispatcher_list(clean_test_project):
    r = _run(msproject_calendar({"action": "list"}))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"
    assert "calendars" in parsed


def test_dispatcher_holidays_uzbek(clean_test_project):
    _run(msproject_calendar({
        "action": "create",
        "name": "DispUzbek-Phase2a",
        "base_calendar": "Standard",
    }))
    r = _run(msproject_calendar({
        "action": "holidays_uzbek",
        "calendar_name": "DispUzbek-Phase2a",
        "year": 2026,
    }))
    parsed = json.loads(r)
    assert parsed["status"] in ("ok", "partial")
    assert parsed["count"] == 9


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_calendar({"action": "nonsense"}))
    parsed = json.loads(r)
    assert parsed["status"] == "error"
    assert "Unknown action" in parsed["error"]
```

**Step 2: Run — FAIL** (msproject_calendar dispatcher not yet defined)

**Step 3: Implementation**

Add to `msproject_mcp_core.py` near the other `@mcp.tool` decorators (msproject_task / msproject_link / msproject_schedule):

```python
@mcp.tool(
    name="msproject_calendar",
    annotations={"title": "MS Project Calendar Operations", "readOnlyHint": False},
)
async def msproject_calendar(params: dict) -> str:
    """Manage project calendars in active MS Project (COM-based).

    Actions:
    - create: New base calendar from existing one. Params: name, [base_calendar="Standard"]
    - update: Rename or set weekday off. Params: name, [new_name, weekday_off=1-7]
    - add_exception: Non-working day/range. Params: calendar_name, exception_name, start (YYYY-MM-DD), [finish, working=False]
    - assign_to_task: Apply calendar to a task. Params: task_id, calendar_name
    - assign_to_resource: Apply calendar to a resource. Params: resource_id, calendar_name
    - list: List all base calendars + exception counts. Params: (none)
    - holidays_uzbek: Bulk-add 9 Özbekistan 2026 official holidays. Params: calendar_name, [year=2026]

    Phase 2a (27 Apr 2026). Resource integration arrives in Phase 2b.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "create":
            r = _msp_calendar_create(**p)
        elif action == "update":
            r = _msp_calendar_update(**p)
        elif action == "add_exception":
            r = _msp_calendar_add_exception(**p)
        elif action == "assign_to_task":
            r = _msp_calendar_assign_to_task(**p)
        elif action == "assign_to_resource":
            r = _msp_calendar_assign_to_resource(**p)
        elif action == "list":
            r = _msp_calendar_list(**p)
        elif action == "holidays_uzbek":
            r = _msp_calendar_holidays_uzbek(**p)
        else:
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: create/update/add_exception/assign_to_task/assign_to_resource/list/holidays_uzbek"}
    except Exception as e:
        logger.error(f"msproject_calendar({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

Also update the FastMCP server `instructions` string near the top of the file:
```python
mcp = FastMCP(
    "msproject_mcp",
    instructions=(
        "MS Project COM-based MCP server. Connects to running MS Project (Application='MSProject.Application'). "
        "Hybrid speed: 1-5 items COM direct, 6-19 batch, 20+ MSPDI bulk import. "
        "Tools: msproject_task, msproject_link, msproject_schedule, msproject_calendar."
    ),
)
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_msproject_calendar_dispatcher.py -v
```

Expected: 4 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_dispatcher.py
git commit -m "Phase 2a T26: FastMCP msproject_calendar dispatcher (7 actions)"
```

---

## Task 27: Phase 2a Acceptance — Uzbekistan Calendar End-to-End

**Files:**
- Create: `samples/build_uzbek_calendar.py`
- Modify: `README.md` (Phase 2a section)

**Step 1: Acceptance script**

`samples/build_uzbek_calendar.py`:
```python
"""Phase 2a acceptance: build a Uzbekistan-2026 calendar end-to-end.

SAFETY: opens an isolated MS Project via FileNew, never touches the user's
active project. Closes without saving on completion.

Steps:
  1. Create 'Uzbekistan-2026' calendar from Standard
  2. Bulk-add 9 official Özbek holidays
  3. Mark Sunday as non-working
  4. Add 1 task and assign the calendar to it
  5. List calendars and print summary
  6. Close the test project without saving
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_update, _msp_calendar_holidays_uzbek,
    _msp_calendar_assign_to_task, _msp_calendar_list,
    _msp_task_add_single,
)


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")

    # Remember user's active project
    original_name = None
    if app.ActiveProject is not None:
        original_name = app.ActiveProject.Name

    # Open isolated test project
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] Using isolated test project: {test_name}")
    print(f"[SAFE] User's project preserved: {original_name}")

    try:
        t0 = time.time()

        # 1. Create calendar
        print("\n1. Creating 'Uzbekistan-2026' calendar...")
        r = _msp_calendar_create(name="Uzbekistan-2026", base_calendar="Standard")
        assert r["status"] == "ok", r
        print(f"   OK uid={r['calendar_uid']}")

        # 2. Add 9 Uzbek holidays
        print("2. Adding 9 official Uzbek holidays...")
        r = _msp_calendar_holidays_uzbek(calendar_name="Uzbekistan-2026", year=2026)
        assert r["status"] in ("ok", "partial"), r
        print(f"   OK {r['count']} holidays added")
        for h in r["holidays"]:
            print(f"      - {h['date']}  {h['name']}")

        # 3. Sunday off (weekday=1)
        print("3. Marking Sunday as non-working...")
        r = _msp_calendar_update(name="Uzbekistan-2026", weekday_off=1)
        assert r["status"] == "ok", r
        print(f"   OK changes={r['changes']}")

        # 4. Add a task and assign the calendar
        print("4. Adding 'Hafriyat' task and assigning Uzbekistan-2026...")
        add_r = _msp_task_add_single(name="Hafriyat", duration="10d")
        assert add_r["status"] == "ok"
        r = _msp_calendar_assign_to_task(task_id=add_r["task_id"],
                                         calendar_name="Uzbekistan-2026")
        assert r["status"] == "ok", r
        print(f"   OK task_id={add_r['task_id']} -> Uzbekistan-2026")

        # 5. List
        print("5. Listing calendars...")
        r = _msp_calendar_list()
        assert r["status"] == "ok"
        for c in r["calendars"]:
            marker = "*" if c["name"] == "Uzbekistan-2026" else " "
            print(f"   {marker} {c['name']}  exceptions={c['exception_count']}")

        elapsed = time.time() - t0
        print(f"\nOK ACCEPTANCE: end-to-end in {elapsed:.2f}s (target: <5s total)")
        assert elapsed < 5.0, f"Too slow: {elapsed}s"

    finally:
        # 6. Always restore user's project (close test without saving)
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)  # 0 = pjDoNotSave
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

**Step 2: Run acceptance**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python samples/build_uzbek_calendar.py
```

Expected output: full summary, `OK ACCEPTANCE: end-to-end in <5s`. Test project auto-closes; user's original project still open.

**Step 3: Run full test suite (Phase 1 + Phase 2a regression)**

```bash
python -m pytest tests/ -v --tb=short
```

Expected: ~43 Phase 1 tests + ~22 Phase 2a tests all PASS (or skip if MS Project not running).

**Step 4: Update README.md Phase status**

Edit `README.md` — find the Phase 1 section and add a Phase 2a entry below it:

```markdown
### Phase 2a — Calendar (27 Apr 2026)

`msproject_calendar` tool with 7 actions:
- `create` — New base calendar from existing
- `update` — Rename or weekday off
- `add_exception` — Non-working day/range
- `assign_to_task` — Apply calendar to task
- `assign_to_resource` — Apply calendar to resource (full Resource tool in Phase 2b)
- `list` — All calendars + exception counts
- `holidays_uzbek` — Built-in 9 Özbek 2026 holidays bulk-add

Acceptance: `samples/build_uzbek_calendar.py` builds Uzbekistan-2026
calendar end-to-end (create + 9 holidays + Sunday off + task assignment)
in <5 sec, isolated from user's active project.
```

**Step 5: Commit + push**

```bash
git add samples/build_uzbek_calendar.py README.md
git commit -m "Phase 2a T27: end-to-end Uzbekistan calendar acceptance + README"
git push origin main
```

Expected: GitHub now shows Phase 2a commits T18-T27 on main.

**Step 6: Phase 2a kullanıcı onayı sun**

Kullanıcıya rapor:
- ✅ `msproject_calendar` tool — 7 action operational
- ✅ 9 Özbek bayramı tek call'da bulk-add
- ✅ Acceptance: Uzbekistan calendar end-to-end <5s
- ✅ Mevcut Phase 1 testleri (43) regression PASS
- ✅ Yeni Phase 2a testleri (~22) PASS
- ✅ Kullanıcının aktif projesi DOKUNULMADI
- ✅ Push → GitHub
- ⏸ Onay → Phase 2b (Resource Management) başlar

---

## Phase 2a Tamamlama Kriterleri (Re-verify)

1. ✅ `msproject_calendar` tool 7 action ile çalışıyor (T19-T25 + T26 dispatcher)
2. ✅ Acceptance script (`samples/build_uzbek_calendar.py`) PASS — Uzbekistan-2026 end-to-end <5s
3. ✅ Phase 2a yeni testleri (~22) PASS
4. ✅ Phase 1 mevcut 43 test regression PASS
5. ✅ Kullanıcının aktif projesi hiç dokunulmadı (`clean_test_project` disiplini)
6. ✅ README.md güncel
7. ✅ Commit + push GitHub
8. ⏸ Kullanıcı manuel onayı → Phase 2b başlar

---

*Plan tamamlandı: 27 Nisan 2026*
*Tahmini Phase 2a süresi: ~5-6 saat (T18-T27 = 10 task)*
*Sonraki phase (onay sonrası): Phase 2b — Resource Management (`msproject_resource` tool, ~6-8 task)*
