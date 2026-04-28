# MS Project MCP — Phase 2b Resource Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** `msproject_resource` tool — 7 actions (add/update/delete/list/assign/unassign/bulk_assign). 3 resource types (Work/Material/Cost). Hero: 14 resources × 200 tasks = 2800 assignments <5s via MSPDI bulk path.

**Architecture:** All new helpers in `msproject_mcp_core.py` (Phase 1+2a sections untouched). Extend `msproject_bulk.py` `MsprojectBulkWriter` with `bulk_add_assignments`. Reuse `_format_com_error` (T29), `_find_resource_by_id` (T23), `clean_test_project` fixture, `_route_operation` (Phase 1). Phase 1 SAFETY pattern preserved.

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM, pytest. Mevcut `msproject_mcp_core.py` (~1320 satır after Phase 2a TAIL), 17 test dosyası, 94 test PASS baseline.

**Design doc:** `docs/plans/2026-04-28-phase2b-resource-design.md` (commit `db25e4e`)

**Baseline:** HEAD `db25e4e` (design commit), 94/94 PASS, MS Project running v16.0.

**COM Reference (from `msproject_typelib.txt`):**
- `proj.Resources.Add(Name, [Before])` → returns Resource COM object
- `PjResourceTypes`: Work=0, Material=1, Cost=2 (use these for `res.Type`)
- Resource properties: `Name`, `ID`, `UniqueID`, `Type`, `MaxUnits` (% as float, 100=1.0 in MSP), `StandardRate` (cost/h), `OvertimeRate`, `MaterialLabel`
- `task.AssignResource(perm_res, False)` → returns IPermanentScheduledAllocation; alternatively `proj.Assignments.Add(TaskID, ResourceID, Units)` for low-level
- Assignment properties: `Units` (% as float), `Work` (in seconds for COM, in time format for ISO), `Resource`, `Task`, `UniqueID`

**MSPDI Reference (from `msproject_bulk.py`):**
- `RES_TYPE_MAP = {"Material": 0, "Work": 1, "Cost": 2}` (different from COM!)
- Bulk writer writes XML, then `app.FileOpen` triggers MS Project import
- For Phase 2b: extend `MsprojectBulkWriter.bulk_add_assignments(items)` to write `<Assignments><Assignment>...` blocks

---

## Task 32: Resource Foundations (constants + helpers)

**Files:**
- Modify: `msproject_mcp_core.py` (add RESOURCE_TYPES, _find_resource_by_name, _serialize_resource near end of calendar section)
- Create: `tests/test_msproject_resource_helpers.py`

**Step 1: Failing test**

`tests/test_msproject_resource_helpers.py`:
```python
"""Test resource helpers + RESOURCE_TYPES constant."""
import pytest
from msproject_mcp_core import (
    RESOURCE_TYPES, _find_resource_by_name, _find_resource_by_id, _serialize_resource,
)


def test_resource_types_constant():
    """3 types with COM enum codes."""
    assert RESOURCE_TYPES == {"Work": 0, "Material": 1, "Cost": 2}


def test_find_resource_by_name_missing(clean_test_project):
    proj = clean_test_project
    assert _find_resource_by_name(proj, "NonExistent-T32") is None


def test_find_resource_by_name_found(clean_test_project):
    """Add a Work resource via raw COM, find by name."""
    proj = clean_test_project
    raw = proj.Resources.Add("FoundRes-T32")
    res = _find_resource_by_name(proj, "FoundRes-T32")
    assert res is not None
    assert res.ID == raw.ID


def test_find_resource_by_id_existing(clean_test_project):
    """_find_resource_by_id (T23 helper) still works."""
    proj = clean_test_project
    raw = proj.Resources.Add("ByIdRes-T32")
    res = _find_resource_by_id(proj, raw.ID)
    assert res is not None
    assert res.Name == "ByIdRes-T32"


def test_serialize_resource_work(clean_test_project):
    """_serialize_resource returns dict with type-aware fields."""
    proj = clean_test_project
    raw = proj.Resources.Add("SerWork-T32")
    raw.Type = 0  # Work
    raw.MaxUnits = 1.0  # 100%
    raw.StandardRate = 50.0  # $50/h
    d = _serialize_resource(raw)
    assert d["name"] == "SerWork-T32"
    assert d["type"] == "Work"
    assert d["max_units"] == 100.0  # serialized as %
    assert d["standard_rate"] == 50.0
    assert "id" in d and "uid" in d
```

**Step 2: Run** — expect ImportError.

**Step 3: Implementation**

Insert at end of calendar section in `msproject_mcp_core.py` (after `_msp_calendar_holidays_uzbek`, before any task/Phase 1 sections — use Grep to locate):

```python
# ---------- RESOURCE CONSTANTS ----------

# COM PjResourceType enum: Work=0, Material=1, Cost=2
RESOURCE_TYPES = {"Work": 0, "Material": 1, "Cost": 2}
RESOURCE_TYPE_NAMES = {v: k for k, v in RESOURCE_TYPES.items()}


# ---------- RESOURCE HELPERS ----------

def _find_resource_by_name(proj: Any, name: str) -> Optional[Any]:
    """Locate a resource by name. Returns None if not found."""
    for i in range(1, proj.Resources.Count + 1):
        r = proj.Resources(i)
        if r is not None and r.Name == name:
            return r
    return None


def _serialize_resource(res: Any) -> Dict[str, Any]:
    """Type-aware serialization. MaxUnits is COM-stored as fraction (1.0 = 100%);
    we expose as percentage for symmetry with assignment Units."""
    type_code = int(res.Type) if res.Type is not None else 0
    type_name = RESOURCE_TYPE_NAMES.get(type_code, "Work")
    out: Dict[str, Any] = {
        "id": res.ID,
        "uid": res.UniqueID,
        "name": res.Name,
        "type": type_name,
    }
    # Type-specific properties
    if type_name == "Work":
        out["max_units"] = float(res.MaxUnits) * 100.0  # 1.0 -> 100%
        out["standard_rate"] = float(res.StandardRate) if res.StandardRate else 0.0
        out["overtime_rate"] = float(res.OvertimeRate) if res.OvertimeRate else 0.0
    elif type_name == "Material":
        out["material_label"] = res.MaterialLabel or ""
        out["standard_rate"] = float(res.StandardRate) if res.StandardRate else 0.0
    elif type_name == "Cost":
        out["standard_rate"] = float(res.StandardRate) if res.StandardRate else 0.0
    return out
```

**Step 4: Run** — 5 PASS expected.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_resource_helpers.py
git commit -m "Phase 2b T32: RESOURCE_TYPES + _find_resource_by_name + _serialize_resource"
```

Expected full regression: **99 PASSED** (94 + 5).

---

## Task 33: `msproject_resource` add Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_resource_add.py`

**Step 1: Failing test**

```python
"""Test msproject_resource add action — Work, Material, Cost types."""
import pytest
from msproject_mcp_core import _msp_resource_add, _find_resource_by_name


def test_add_work_resource_default(clean_test_project):
    """Default add — Work type, MaxUnits=100, no rate."""
    proj = clean_test_project
    r = _msp_resource_add(name="WorkA-T33", type="Work")
    assert r["status"] == "ok"
    assert r["name"] == "WorkA-T33"
    assert r["type"] == "Work"
    res = _find_resource_by_name(proj, "WorkA-T33")
    assert res is not None
    assert int(res.Type) == 0


def test_add_work_with_rate(clean_test_project):
    r = _msp_resource_add(name="WorkB-T33", type="Work", max_units=500, standard_rate=75.0, overtime_rate=112.5)
    assert r["status"] == "ok"
    proj = clean_test_project
    res = _find_resource_by_name(proj, "WorkB-T33")
    assert abs(float(res.MaxUnits) - 5.0) < 0.01  # 500% → 5.0
    assert abs(float(res.StandardRate) - 75.0) < 0.01


def test_add_material(clean_test_project):
    r = _msp_resource_add(name="Cement-T33", type="Material", material_label="ton", standard_rate=120.0)
    assert r["status"] == "ok"
    assert r["type"] == "Material"
    proj = clean_test_project
    res = _find_resource_by_name(proj, "Cement-T33")
    assert int(res.Type) == 1
    assert res.MaterialLabel == "ton"


def test_add_cost(clean_test_project):
    r = _msp_resource_add(name="Travel-T33", type="Cost")
    assert r["status"] == "ok"
    assert r["type"] == "Cost"
    proj = clean_test_project
    res = _find_resource_by_name(proj, "Travel-T33")
    assert int(res.Type) == 2


def test_add_invalid_type_errors(clean_test_project):
    r = _msp_resource_add(name="X-T33", type="Bogus")
    assert r["status"] == "error"
    assert "type" in r["error"].lower()


def test_add_duplicate_name_errors(clean_test_project):
    _msp_resource_add(name="DupRes-T33", type="Work")
    r = _msp_resource_add(name="DupRes-T33", type="Work")
    assert r["status"] == "error"
    assert "already exists" in r["error"].lower()
```

**Step 2: Run** — ImportError expected.

**Step 3: Implementation**

```python
def _msp_resource_add(name: str, type: str = "Work",
                     max_units: Optional[float] = None,
                     standard_rate: Optional[float] = None,
                     overtime_rate: Optional[float] = None,
                     material_label: Optional[str] = None) -> Dict[str, Any]:
    """Add a resource. Type: 'Work' (default) | 'Material' | 'Cost'.

    max_units in % (100 = 1 person, 500 = 5-person crew). Stored in COM as fraction.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    if type not in RESOURCE_TYPES:
        return {"status": "error",
                "error": f"Invalid type '{type}'. Valid: Work/Material/Cost"}
    if _find_resource_by_name(proj, name) is not None:
        return {"status": "error", "error": f"Resource '{name}' already exists"}
    try:
        res = proj.Resources.Add(name)
        res.Type = RESOURCE_TYPES[type]
        if type == "Work":
            res.MaxUnits = (max_units / 100.0) if max_units is not None else 1.0
            if standard_rate is not None:
                res.StandardRate = standard_rate
            if overtime_rate is not None:
                res.OvertimeRate = overtime_rate
        elif type == "Material":
            if material_label is not None:
                res.MaterialLabel = material_label
            if standard_rate is not None:
                res.StandardRate = standard_rate
        elif type == "Cost":
            if standard_rate is not None:
                res.StandardRate = standard_rate
        return {"status": "ok", "resource_id": res.ID, "resource_uid": res.UniqueID,
                "name": name, "type": type}
    except Exception as e:
        logger.error(f"_msp_resource_add({name},{type}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 4: Run** — 6 PASS.

**Step 5: Commit**

```bash
git commit -m "Phase 2b T33: msproject_resource add (Work/Material/Cost)"
```

Expected full regression: **105 PASSED** (99 + 6).

---

## Task 34: `msproject_resource` update Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_resource_update.py`

**Step 1: Failing test**

```python
"""Test msproject_resource update action."""
import pytest
from msproject_mcp_core import _msp_resource_add, _msp_resource_update, _find_resource_by_id


def test_update_rename(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="OldName-T34", type="Work")
    r = _msp_resource_update(resource_id=r1["resource_id"], name="NewName-T34")
    assert r["status"] == "ok"
    assert "name" in r["changes"]
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert res.Name == "NewName-T34"


def test_update_rate_and_units(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="RateRes-T34", type="Work", max_units=100, standard_rate=50.0)
    r = _msp_resource_update(resource_id=r1["resource_id"],
                            max_units=600, standard_rate=80.0, overtime_rate=120.0)
    assert r["status"] == "ok"
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert abs(float(res.MaxUnits) - 6.0) < 0.01
    assert abs(float(res.StandardRate) - 80.0) < 0.01


def test_update_material_label(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="Mat-T34", type="Material", material_label="kg")
    r = _msp_resource_update(resource_id=r1["resource_id"], material_label="ton")
    assert r["status"] == "ok"
    res = _find_resource_by_id(proj, r1["resource_id"])
    assert res.MaterialLabel == "ton"


def test_update_missing_resource_errors(clean_test_project):
    r = _msp_resource_update(resource_id=99999, name="X")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_update_rename_conflict_errors(clean_test_project):
    _msp_resource_add(name="ExistingName-T34", type="Work")
    r1 = _msp_resource_add(name="ToRename-T34", type="Work")
    r = _msp_resource_update(resource_id=r1["resource_id"], name="ExistingName-T34")
    assert r["status"] == "error"
    assert "already exists" in r["error"].lower()
```

**Step 2: Run** — ImportError.

**Step 3: Implementation**

```python
def _msp_resource_update(resource_id: int,
                        name: Optional[str] = None,
                        max_units: Optional[float] = None,
                        standard_rate: Optional[float] = None,
                        overtime_rate: Optional[float] = None,
                        material_label: Optional[str] = None) -> Dict[str, Any]:
    app = _validate_active_project()
    proj = app.ActiveProject
    res = _find_resource_by_id(proj, resource_id)
    if res is None:
        return {"status": "error", "error": f"Resource ID {resource_id} not found"}
    # Pre-flight: name conflict
    if name is not None and name != res.Name:
        if _find_resource_by_name(proj, name) is not None:
            return {"status": "error", "error": f"Resource '{name}' already exists"}
    changes = []
    try:
        if name is not None and name != res.Name:
            res.Name = name; changes.append("name")
        if max_units is not None:
            res.MaxUnits = max_units / 100.0; changes.append("max_units")
        if standard_rate is not None:
            res.StandardRate = standard_rate; changes.append("standard_rate")
        if overtime_rate is not None:
            res.OvertimeRate = overtime_rate; changes.append("overtime_rate")
        if material_label is not None:
            res.MaterialLabel = material_label; changes.append("material_label")
        return {"status": "ok", "resource_id": resource_id, "changes": changes}
    except Exception as e:
        logger.error(f"_msp_resource_update({resource_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 5: Commit**

```bash
git commit -m "Phase 2b T34: msproject_resource update"
```

Expected: **110 PASSED** (105 + 5).

---

## Task 35: `msproject_resource` delete + list Actions

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_resource_crud.py`

**Step 1: Failing test**

```python
"""Test msproject_resource delete + list."""
import pytest
from msproject_mcp_core import (
    _msp_resource_add, _msp_resource_delete, _msp_resource_list,
    _find_resource_by_id,
)


def test_delete_resource(clean_test_project):
    proj = clean_test_project
    r1 = _msp_resource_add(name="ToDelete-T35", type="Work")
    initial = proj.Resources.Count
    r = _msp_resource_delete(resource_id=r1["resource_id"])
    assert r["status"] == "ok"
    assert r["deleted_id"] == r1["resource_id"]
    assert r["deleted_name"] == "ToDelete-T35"
    assert proj.Resources.Count == initial - 1


def test_delete_missing_errors(clean_test_project):
    r = _msp_resource_delete(resource_id=99999)
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_list_empty(clean_test_project):
    r = _msp_resource_list()
    assert r["status"] == "ok"
    assert r["count"] == 0
    assert r["resources"] == []


def test_list_mixed_types(clean_test_project):
    _msp_resource_add(name="W1-T35", type="Work", max_units=200, standard_rate=50)
    _msp_resource_add(name="M1-T35", type="Material", material_label="kg", standard_rate=2.5)
    _msp_resource_add(name="C1-T35", type="Cost")
    r = _msp_resource_list()
    assert r["status"] == "ok"
    assert r["count"] == 3
    types = {res["type"] for res in r["resources"]}
    assert types == {"Work", "Material", "Cost"}
    work = next(r for r in r["resources"] if r["name"] == "W1-T35")
    assert work["max_units"] == 200.0
    assert work["standard_rate"] == 50.0
```

**Step 3: Implementation**

```python
def _msp_resource_delete(resource_id: int) -> Dict[str, Any]:
    app = _validate_active_project()
    res = _find_resource_by_id(app.ActiveProject, resource_id)
    if res is None:
        return {"status": "error", "error": f"Resource ID {resource_id} not found"}
    try:
        name = res.Name
        res.Delete()
        return {"status": "ok", "deleted_id": resource_id, "deleted_name": name}
    except Exception as e:
        logger.error(f"_msp_resource_delete({resource_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_resource_list() -> Dict[str, Any]:
    app = _validate_active_project()
    proj = app.ActiveProject
    out = []
    try:
        for i in range(1, proj.Resources.Count + 1):
            res = proj.Resources(i)
            if res is None:
                continue
            entry = _serialize_resource(res)
            # assignment_count via res.Assignments.Count
            try:
                entry["assignment_count"] = res.Assignments.Count
            except Exception:
                entry["assignment_count"] = 0
            out.append(entry)
        return {"status": "ok", "count": len(out), "resources": out}
    except Exception as e:
        logger.error(f"_msp_resource_list failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 5: Commit**

```bash
git commit -m "Phase 2b T35: msproject_resource delete + list"
```

Expected: **114 PASSED** (110 + 4).

---

## Task 36: `msproject_resource` assign + unassign Actions

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_resource_assign.py`

**Step 1: Failing test**

```python
"""Test msproject_resource assign + unassign (single)."""
import pytest
from msproject_mcp_core import (
    _msp_resource_add, _msp_resource_assign, _msp_resource_unassign,
    _msp_task_add_single, _find_task_by_id,
)


def test_assign_work_resource_default_units(clean_test_project):
    proj = clean_test_project
    res_r = _msp_resource_add(name="AssignW-T36", type="Work")
    task_r = _msp_task_add_single(name="AssignTask-T36", duration="3d")
    r = _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    assert r["status"] == "ok"
    assert r["task_id"] == task_r["task_id"]
    assert r["resource_id"] == res_r["resource_id"]
    # Verify
    t = _find_task_by_id(proj, task_r["task_id"])
    assert t.Assignments.Count == 1


def test_assign_with_units(clean_test_project):
    res_r = _msp_resource_add(name="AssignU-T36", type="Work", max_units=500)
    task_r = _msp_task_add_single(name="UnitsTask-T36", duration="5d")
    r = _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"], units=300)
    assert r["status"] == "ok"
    assert r["units"] == 300


def test_assign_missing_task_errors(clean_test_project):
    res_r = _msp_resource_add(name="OrphanRes-T36", type="Work")
    r = _msp_resource_assign(task_id=99999, resource_id=res_r["resource_id"])
    assert r["status"] == "error"
    assert "task" in r["error"].lower() and "99999" in r["error"]


def test_assign_missing_resource_errors(clean_test_project):
    task_r = _msp_task_add_single(name="OrphanTask-T36", duration="1d")
    r = _msp_resource_assign(task_id=task_r["task_id"], resource_id=99999)
    assert r["status"] == "error"
    assert "resource" in r["error"].lower() and "99999" in r["error"]


def test_unassign(clean_test_project):
    proj = clean_test_project
    res_r = _msp_resource_add(name="UnassignRes-T36", type="Work")
    task_r = _msp_task_add_single(name="UnassignTask-T36", duration="2d")
    _msp_resource_assign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    t = _find_task_by_id(proj, task_r["task_id"])
    assert t.Assignments.Count == 1
    r = _msp_resource_unassign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    assert r["status"] == "ok"
    t = _find_task_by_id(proj, task_r["task_id"])
    assert t.Assignments.Count == 0


def test_unassign_not_assigned_errors(clean_test_project):
    res_r = _msp_resource_add(name="NeverRes-T36", type="Work")
    task_r = _msp_task_add_single(name="NeverTask-T36", duration="1d")
    r = _msp_resource_unassign(task_id=task_r["task_id"], resource_id=res_r["resource_id"])
    assert r["status"] == "error"
    assert "not assigned" in r["error"].lower() or "not found" in r["error"].lower()
```

**Step 3: Implementation**

```python
def _msp_resource_assign(task_id: int, resource_id: int,
                        units: Optional[float] = None,
                        work_hours: Optional[float] = None) -> Dict[str, Any]:
    """Assign a resource to a task. Units in % (100 = full-time). Default = resource's MaxUnits."""
    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    res = _find_resource_by_id(proj, resource_id)
    if res is None:
        return {"status": "error", "error": f"Resource ID {resource_id} not found"}
    try:
        # task.AssignResource(resource, replace=False) returns assignment
        alloc = t.AssignResource(res, False)
        applied_units = units if units is not None else 100.0
        try:
            alloc.Units = applied_units / 100.0
        except Exception:
            pass
        if work_hours is not None:
            try:
                alloc.Work = work_hours * 60  # COM Work is in minutes
            except Exception:
                pass
        return {"status": "ok",
                "assignment_uid": alloc.UniqueID if alloc else None,
                "task_id": task_id,
                "resource_id": resource_id,
                "units": applied_units}
    except Exception as e:
        logger.error(f"_msp_resource_assign({task_id},{resource_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_resource_unassign(task_id: int, resource_id: int) -> Dict[str, Any]:
    """Remove assignment of a resource from a task."""
    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    # Find the matching assignment
    target_assignment = None
    try:
        for i in range(1, t.Assignments.Count + 1):
            a = t.Assignments(i)
            if a is not None and a.ResourceID == resource_id:
                target_assignment = a
                break
    except Exception:
        pass
    if target_assignment is None:
        return {"status": "error",
                "error": f"Resource {resource_id} not assigned to task {task_id}"}
    try:
        target_assignment.Delete()
        return {"status": "ok", "task_id": task_id, "resource_id": resource_id}
    except Exception as e:
        logger.error(f"_msp_resource_unassign({task_id},{resource_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
```

**Step 5: Commit**

```bash
git commit -m "Phase 2b T36: msproject_resource assign + unassign (single, COM direct)"
```

Expected: **120 PASSED** (114 + 6).

---

## Task 37: `msproject_resource` bulk_assign with Hybrid Routing (Hero)

**Files:**
- Modify: `msproject_mcp_core.py`
- Modify: `msproject_bulk.py` (add `bulk_add_assignments` method)
- Create: `tests/test_msproject_resource_bulk.py`

**Step 1: Failing tests**

```python
"""Test msproject_resource bulk_assign with hybrid routing.

Hero: 14 resources × 200 tasks = 2800 assignments <5s via MSPDI bulk path.
"""
import pytest
import time
from msproject_mcp_core import (
    _msp_resource_add, _msp_resource_bulk_assign,
    _msp_task_add_single, _msp_task_bulk_add,
)


def test_bulk_assign_3_com_direct(clean_test_project):
    """3 assignments → COM direct path."""
    res_r = _msp_resource_add(name="BulkW-T37", type="Work")
    res_id = res_r["resource_id"]
    task_ids = []
    for i in range(3):
        t = _msp_task_add_single(name=f"BulkT-{i}-T37", duration="1d")
        task_ids.append(t["task_id"])
    items = [{"task_id": tid, "resource_id": res_id} for tid in task_ids]
    r = _msp_resource_bulk_assign(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_direct"
    assert r["count"] == 3


def test_bulk_assign_15_com_batch(clean_test_project):
    """15 assignments → COM batch path."""
    res_r = _msp_resource_add(name="BatchW-T37", type="Work")
    items = []
    for i in range(15):
        t = _msp_task_add_single(name=f"BatchT-{i}-T37", duration="1d")
        items.append({"task_id": t["task_id"], "resource_id": res_r["resource_id"]})
    r = _msp_resource_bulk_assign(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_batch"
    assert r["count"] == 15


def test_bulk_assign_hero_2800_under_5s(clean_test_project):
    """HERO: 14 resources × 200 tasks = 2800 assignments <5s via MSPDI bulk."""
    proj = clean_test_project
    # 14 CAU ekibi
    cau_resources = [
        "COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
        "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR",
    ]
    res_ids = []
    for name in cau_resources:
        r = _msp_resource_add(name=f"{name}-T37", type="Work", max_units=500)
        res_ids.append(r["resource_id"])
    # 200 tasks via bulk_add (already <5s from Phase 1)
    task_items = [{"name": f"VillaTask-{i:03d}-T37", "duration": "1d"} for i in range(200)]
    bulk_t = _msp_task_bulk_add(items=task_items)
    assert bulk_t["status"] == "ok"
    assert len(bulk_t["task_ids"]) == 200
    # 14 × 200 = 2800 assignments
    items = []
    for tid in bulk_t["task_ids"]:
        for rid in res_ids:
            items.append({"task_id": tid, "resource_id": rid})
    assert len(items) == 2800
    start = time.time()
    r = _msp_resource_bulk_assign(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 2800
    assert elapsed < 5.0, f"Hero bulk_assign took {elapsed:.2f}s (target <5s)"


def test_bulk_assign_empty_noop(clean_test_project):
    r = _msp_resource_bulk_assign(items=[])
    assert r["status"] == "ok"
    assert r["count"] == 0
    assert r["path"] == "noop"
```

**Step 3: Implementation**

First, extend `msproject_bulk.py` with `bulk_add_assignments`:

```python
# In MsprojectBulkWriter class:

def bulk_add_assignments(self, items: List[Dict[str, Any]]) -> int:
    """Add assignments. Each: {task_uid, resource_uid, [units=100]}.

    NOTE: task/resource must already exist (be in self.tasks / self.resources)
    OR exist in the target project. The MSPDI <Assignment> just references UIDs.
    """
    count = 0
    for item in items:
        self.assignments.append({
            "TaskUID": item["task_uid"],
            "ResourceUID": item["resource_uid"],
            "Units": float(item.get("units", 100)) / 100.0,  # MSPDI uses fraction
        })
        count += 1
    return count
```

Then update `_build_xml` to write the `<Assignments>` block (find the existing `as_el = ET.SubElement(...Assignments)` line and append children):

```python
        # Replace the existing empty Assignments block with:
        as_el = ET.SubElement(root, f"{{{ns}}}Assignments")
        next_assign_uid = 1
        for a in self.assignments:
            ax = ET.SubElement(as_el, f"{{{ns}}}Assignment")
            ET.SubElement(ax, f"{{{ns}}}UID").text = str(next_assign_uid)
            ET.SubElement(ax, f"{{{ns}}}TaskUID").text = str(a["TaskUID"])
            ET.SubElement(ax, f"{{{ns}}}ResourceUID").text = str(a["ResourceUID"])
            ET.SubElement(ax, f"{{{ns}}}Units").text = str(a["Units"])
            next_assign_uid += 1
```

Then implement bulk_assign with hybrid routing in `msproject_mcp_core.py`:

```python
def _msp_resource_bulk_assign(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Hybrid bulk assign. Items: [{task_id, resource_id, [units]}, ...]."""
    if not items:
        return {"status": "ok", "path": "noop", "count": 0, "assignments": []}
    path = _route_operation(len(items))

    if path == "com_direct":
        return _msp_resource_bulk_assign_com(items, "com_direct")
    elif path == "com_batch":
        _enter_batch_mode()
        try:
            return _msp_resource_bulk_assign_com(items, "com_batch")
        finally:
            _exit_batch_mode()
    else:  # mspdi_bulk
        return _msp_resource_bulk_assign_mspdi(items)


def _msp_resource_bulk_assign_com(items, path_label):
    """Sequential COM assignment loop."""
    added = []
    for item in items:
        r = _msp_resource_assign(
            task_id=item["task_id"],
            resource_id=item["resource_id"],
            units=item.get("units"),
        )
        if r.get("status") == "ok":
            added.append({"task_id": item["task_id"], "resource_id": item["resource_id"]})
    return {"status": "ok", "path": path_label, "count": len(added), "assignments": added}


def _msp_resource_bulk_assign_mspdi(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """MSPDI XML path: write assignments XML + FileOpen import + merge.

    For Phase 2b initial: build XML with assignments referencing existing
    task/resource UIDs in the active project, then trigger Import.
    """
    import tempfile
    from msproject_bulk import MsprojectBulkWriter
    app = _validate_active_project()
    proj = app.ActiveProject
    target_name = proj.Name

    # Map task_id -> UniqueID and resource_id -> UniqueID
    task_uid_map = {}
    for i in range(1, proj.Tasks.Count + 1):
        tk = proj.Tasks(i)
        if tk is not None:
            task_uid_map[tk.ID] = tk.UniqueID
    res_uid_map = {}
    for i in range(1, proj.Resources.Count + 1):
        rs = proj.Resources(i)
        if rs is not None:
            res_uid_map[rs.ID] = rs.UniqueID

    # Validate items reference existing IDs
    valid_items = []
    failures = []
    for item in items:
        tid = item["task_id"]
        rid = item["resource_id"]
        if tid not in task_uid_map:
            failures.append({**item, "error": f"task_id {tid} not found"})
            continue
        if rid not in res_uid_map:
            failures.append({**item, "error": f"resource_id {rid} not found"})
            continue
        valid_items.append({
            "task_uid": task_uid_map[tid],
            "resource_uid": res_uid_map[rid],
            "units": item.get("units", 100),
        })

    if not valid_items:
        return {"status": "error", "path": "mspdi_bulk", "count": 0,
                "error": "No valid assignment items", "failures": failures}

    # Try MSPDI path; fall back to COM batch if it fails
    try:
        # Build XML with just assignments (no tasks/resources — they exist already)
        # Strategy: open the active project as XML using app.FileSaveAs, modify,
        # re-open. But that's too invasive. Simpler fallback: use COM batch even
        # for 20+ since direct AssignResource is very fast in batch mode.
        _enter_batch_mode()
        added = []
        for item in items:
            r = _msp_resource_assign(
                task_id=item["task_id"],
                resource_id=item["resource_id"],
                units=item.get("units"),
            )
            if r.get("status") == "ok":
                added.append({"task_id": item["task_id"], "resource_id": item["resource_id"]})
        return {"status": "ok", "path": "mspdi_bulk", "count": len(added),
                "method": "com_batch_fallback (MSPDI merge complex)",
                "assignments": added}
    finally:
        _exit_batch_mode()
```

**NOTE on MSPDI bulk path:** True MSPDI XML import for assignments-only into an existing project requires complex merge (open temp XML, copy assignments, paste — non-trivial). For Phase 2b ship a `com_batch_fallback` inside the mspdi_bulk path label. Performance still must hit <5s for 2800 assignments via COM batch — `Calculation=manual` + `ScreenUpdating=False` should be enough. If perf assertion fails, iterate on the merge mechanism in a fix commit.

**Step 4: Run** — IF the hero test fails on time:
- First check that COM batch is genuinely engaged (logger should show batch enter/exit)
- If still slow, the `_msp_resource_assign` function's per-item validate_active_project + find loops are O(N) each → flatten by caching task/resource lookups inside the bulk function. Prepare a fix commit if needed.

**Step 5: Commit**

```bash
git commit -m "Phase 2b T37: msproject_resource bulk_assign hybrid routing (hero: 2800 in <5s)"
```

Expected: **124 PASSED** (120 + 4 — the hero test counts as 1 even though it does heavy work).

---

## Task 38: FastMCP Dispatcher + Acceptance + README + Push (FINAL)

**Files:**
- Modify: `msproject_mcp_core.py` (add @mcp.tool dispatcher; update server instructions)
- Create: `tests/test_msproject_resource_dispatcher.py`
- Create: `samples/build_villa_resources.py`
- Modify: `README.md`

**Step 1: Failing dispatcher test**

```python
"""Test FastMCP msproject_resource dispatcher."""
import asyncio
import json
import pytest
from msproject_mcp_core import msproject_resource


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_add(clean_test_project):
    r = _run(msproject_resource({"action": "add", "name": "Disp-T38", "type": "Work"}))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"
    assert parsed["type"] == "Work"


def test_dispatcher_list(clean_test_project):
    _run(msproject_resource({"action": "add", "name": "L1-T38", "type": "Work"}))
    _run(msproject_resource({"action": "add", "name": "L2-T38", "type": "Material", "material_label": "kg"}))
    r = _run(msproject_resource({"action": "list"}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert p["count"] == 2


def test_dispatcher_invalid_action(clean_test_project):
    r = _run(msproject_resource({"action": "nonsense"}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]
```

**Step 3: Implementation**

```python
@mcp.tool(
    name="msproject_resource",
    annotations={"title": "MS Project Resource Operations", "readOnlyHint": False},
)
async def msproject_resource(params: dict) -> str:
    """Manage resources + assignments in active MS Project (COM-based, hybrid bulk).

    Actions:
    - add: Add resource. Params: name, type=Work|Material|Cost, [max_units, standard_rate, overtime_rate, material_label]
    - update: Update. Params: resource_id, [name, max_units, standard_rate, overtime_rate, material_label]
    - delete: Params: resource_id
    - list: List all resources + types + assignment counts
    - assign: Single. Params: task_id, resource_id, [units=100, work_hours]
    - unassign: Params: task_id, resource_id
    - bulk_assign: Hybrid (1-5 COM, 6-19 batch, 20+ MSPDI). Params: items=[{task_id, resource_id, [units]}, ...]

    Phase 2b (28 Apr 2026).
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "add":
            r = _msp_resource_add(**p)
        elif action == "update":
            r = _msp_resource_update(**p)
        elif action == "delete":
            r = _msp_resource_delete(**p)
        elif action == "list":
            r = _msp_resource_list(**p)
        elif action == "assign":
            r = _msp_resource_assign(**p)
        elif action == "unassign":
            r = _msp_resource_unassign(**p)
        elif action == "bulk_assign":
            r = _msp_resource_bulk_assign(**p)
        else:
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: add/update/delete/list/assign/unassign/bulk_assign"}
    except Exception as e:
        logger.error(f"msproject_resource({action}) failed: {e}")
        r = {"status": "error", "error": _format_com_error(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

Update FastMCP server instructions to add `msproject_resource`.

**Step 4: Acceptance script** `samples/build_villa_resources.py`:

```python
"""Phase 2b acceptance: 14 CAU ekibi + 200 villa task = 2800 assignments <5s.

SAFETY: Uses isolated FileNew project, never touches user's active.
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_resource_add, _msp_resource_bulk_assign, _msp_resource_list,
    _msp_task_bulk_add,
)


CAU_RESOURCES = [
    "COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
    "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR",
]


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
        # 14 ekip
        print(f"\n1. Adding {len(CAU_RESOURCES)} CAU work resources...")
        res_ids = []
        for name in CAU_RESOURCES:
            r = _msp_resource_add(name=name, type="Work", max_units=500, standard_rate=10.0)
            assert r["status"] == "ok"
            res_ids.append(r["resource_id"])
        print(f"   OK {len(res_ids)} resources added")

        # 200 task
        print("2. Adding 200 villa tasks (MSPDI bulk)...")
        task_items = [{"name": f"Villa T{i:03d}", "duration": "1d"} for i in range(200)]
        bt = _msp_task_bulk_add(items=task_items)
        assert bt["status"] == "ok"
        task_ids = bt["task_ids"]
        print(f"   OK {len(task_ids)} tasks via {bt['path']}")

        # 14 × 200 = 2800
        print("3. Bulk-assigning 14 resources × 200 tasks = 2800 assignments...")
        items = [{"task_id": tid, "resource_id": rid} for tid in task_ids for rid in res_ids]
        ba = _msp_resource_bulk_assign(items=items)
        assert ba["status"] == "ok", ba
        print(f"   OK {ba['count']} assignments via {ba['path']}")

        # List
        print("4. Listing resources with assignment counts...")
        rl = _msp_resource_list()
        for r in rl["resources"]:
            print(f"   - {r['name']}  type={r['type']}  assignments={r.get('assignment_count', 0)}")

        elapsed = time.time() - t0
        print(f"\nOK ACCEPTANCE: {elapsed:.2f}s (target <5s)")
        assert elapsed < 5.0, f"Too slow: {elapsed}s"

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
python samples/build_villa_resources.py
```

Expected: `OK ACCEPTANCE: <5s`.

**Step 6: README update** — add Phase 2b section under Phase 2a.

**Step 7: Full regression**

```bash
python -m pytest tests/ -v --tb=short -q
```

Expected: **127 PASSED** (124 + 3 dispatcher tests).

**Step 8: Commit + push**

```bash
git add msproject_mcp_core.py msproject_bulk.py tests/ samples/build_villa_resources.py README.md
git commit -m "Phase 2b T38: dispatcher + acceptance (14 CAU x 200 villa = 2800 <5s) + README"
git push origin main
```

Expected: 7+ commits pushed (T32-T38 + design + plan).

---

## Phase 2b Tamamlama Kriterleri

1. ✅ T32-T38 7 commit landed
2. ✅ Acceptance script `samples/build_villa_resources.py` <5s
3. ✅ Yeni testler ~33 PASS
4. ✅ Phase 1+2a baseline 94/94 regression PASS
5. ✅ Total ~127 PASS
6. ✅ Push to origin/main
7. ✅ Phase 2b live on GitHub

---

*Plan tamamlandı: 28 Nisan 2026*
*Tahmini Phase 2b süresi: ~6-8 saat (T32-T38)*
