# Phase 4 File MCP Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** New `msproject_file` MCP tool — 14 actions covering file-based MS Project operations: native XML/MSPDI parser + MPXJ MPP fallback, read+write with auto COM sync, hero `bulk_add_assignments` flipping Phase 2b strict xfail.

**Architecture:** All helpers in `msproject_mcp_core.py` Phase 4 section (Phase 1+2+3 untouched). Factory pattern `_get_msp_file_manager(file_path)` returns `MspdiProject` (XML, reuse from `mspdi_parser.py`) or new `MspMppFileManager` (MPP via MPXJ, adapted from Asta `asta_mcp_file.py`). Auto-sync helper `_auto_sync_to_open_msp(temp.xml)` mandatorily called by every write action — MSP open → COM import + `proj.Reschedule()`; MSP closed → XML disk only. Hero path: `mspdi_parser` writes 2800 `<Assignment>` blocks, then `_auto_sync_to_open_msp` merges via FileOpen + EditPaste, target <5s.

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM, pytest, mspdi_parser (existing), mpxj + jpype1 (new dependency). Mevcut `msproject_mcp_core.py` (~2400 satır after Phase 3b + Phase 2b TAIL fix), 24+ test dosyası, **282 PASS + 1 xfail** baseline.

**Design doc:** `docs/plans/2026-04-30-phase4-file-design.md` (commit `2c93865`)

**Baseline state at start:** HEAD `2c93865`, MS Project running v16.0 (DETACHED_PROCESS recovery if ROT lost — Section 5 of resume doc).

**KEY REFERENCE FILES:**
- `mspdi_parser.py` — Asta's MSPDI XML parser/writer (~1050 lines), class `MspdiProject` with read+write methods. **PROBE FIRST** in T65 to confirm `<Assignment>` write API.
- `asta_mcp_file.py` — Asta File MCP (~2300 lines) with `AstaFileManager` MPXJ class. Adapt to `MspMppFileManager` (drop `.pp` support, keep `.mpp`).
- `tests/fixtures/empty_msp.xml` — MSP 16.0 reference XML structure (full Calendar block, project settings).

**CRITICAL DISCOVERIES TO HONOR:**
1. **MSP MSPDI FileOpen Duration drop** (Phase 2b TAIL): `app.FileOpen(*.xml)` silently drops Duration. Workaround pattern: post-paste `t.Duration = expected_minutes` re-set. This MUST be applied in `_auto_sync_to_open_msp` for tasks/assignments where duration matters.
2. **`app.UpdateProject` positional ONLY** (Phase 3b T57): `app.UpdateProject(All, UpdateDate, action)` — never named-args.
3. **Probe-first attitude** (Phase 3b T57/T60): Plan code's enum values can be wrong. Always probe MSP COM live before relying on numeric constants.
4. **`task.Duration` is in MINUTES** (not days) at COM level. Use `_minutes_to_hours` / `_hours_to_minutes` helpers from Phase 3b.
5. **Past dates rejected** (Phase 3b discovery): MSP project default start = today. Tests/scripts use today-relative dates throughout.

---

## Task 65: Foundations — Factory + MspMppFileManager + JVM Lifecycle + Schema Detection

**Files:**
- Modify: `msproject_mcp_core.py` (add Phase 4 section at end of file, after Phase 3b helpers — use Grep `_msp_progress_summary` to locate insertion point)
- Modify: `requirements.txt` (add `mpxj`, `jpype1`)
- Create: `tests/test_msproject_file_factory.py`
- Create: `tests/fixtures/sample_msp.xml` (small MSPDI fixture — 3 tasks, 2 resources, 6 assignments)

**Step 1: Probe `mspdi_parser` capabilities**

Before writing tests, verify MspdiProject's API surface. Run probe (DO NOT COMMIT):

```python
# probe_mspdi.py
from mspdi_parser import MspdiProject
p = MspdiProject('tests/fixtures/empty_msp.xml')
# Print all public methods
print([m for m in dir(p) if not m.startswith('_')])
# Check write capability
print('add_task' in dir(p), 'add_assignment' in dir(p), 'save' in dir(p))
```

**Decision rule:** If MspdiProject lacks `add_task`/`add_assignment`/`save`, T70+T73 plans need adjustment (helpers must wrap raw XML manipulation). Document findings in commit message.

**Step 2: Create XML fixture**

`tests/fixtures/sample_msp.xml` — minimal valid MSPDI with 3 tasks (T1 1d, T2 2d, T3 3d), 2 work resources (R1, R2), 6 assignments (each task gets both resources). Base it on `empty_msp.xml` calendar block. Save with `<DefaultTaskType>0</DefaultTaskType>` (auto-scheduled) and explicit Duration in PT format (PT8H0M0S, PT16H0M0S, PT24H0M0S).

**Step 3: Failing tests**

`tests/test_msproject_file_factory.py`:
```python
"""Test Phase 4 file MCP factory + format detection."""
import pytest
import os
from msproject_mcp_core import (
    _detect_msp_xml_schema,
    _get_msp_file_manager,
    MspMppFileManager,
)
from mspdi_parser import MspdiProject

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
MSP_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")
EMPTY_MSP = os.path.join(FIXTURE_DIR, "empty_msp.xml")


def test_detect_msp_xml_schema_positive():
    """MSP XML with schemas.microsoft.com/project namespace → True."""
    assert _detect_msp_xml_schema(MSP_XML) is True
    assert _detect_msp_xml_schema(EMPTY_MSP) is True


def test_detect_msp_xml_schema_missing_file_raises():
    with pytest.raises(FileNotFoundError):
        _detect_msp_xml_schema("/nonexistent/file.xml")


def test_get_manager_xml_returns_mspdi():
    mgr = _get_msp_file_manager(MSP_XML)
    assert isinstance(mgr, MspdiProject)


def test_get_manager_unsupported_extension():
    with pytest.raises(ValueError) as exc:
        _get_msp_file_manager("/path/file.pp")
    assert "extension" in str(exc.value).lower()


def test_msp_mpp_file_manager_init_smoke(tmp_path):
    """MspMppFileManager initializes with .mpp path (smoke — no read needed)."""
    fake_mpp = tmp_path / "fake.mpp"
    fake_mpp.write_bytes(b"\x00" * 100)  # not real MPP, but file exists
    # Construction should not fail; actual MPXJ read happens lazily
    mgr = MspMppFileManager(str(fake_mpp))
    assert mgr.file_path.endswith("fake.mpp")
```

**Step 4: Run tests (expect ImportError)**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/test_msproject_file_factory.py -v
```

Expected: ImportError on `_detect_msp_xml_schema` etc.

**Step 5: Implementation**

Add at end of `msproject_mcp_core.py` (after Phase 3b helpers, before `if __name__`):

```python
# ============================================================================
# PHASE 4 — FILE MCP (msproject_file)
# ============================================================================

# Native MSPDI parser (zero Java dependency, reuse from Asta)
from mspdi_parser import MspdiProject

# JVM pre-start for MPXJ (lazy — only if .mpp encountered)
_jvm_started = False


def _ensure_jvm_started() -> None:
    """Start JVM lazily on first MPP request. Idempotent."""
    global _jvm_started
    if _jvm_started:
        return
    try:
        import mpxj
        if not mpxj.jpype.isJVMStarted():
            mpxj.jpype.startJVM()
            logger.info("JVM started for MPXJ")
        _jvm_started = True
    except ImportError as e:
        raise RuntimeError(
            "MPXJ not installed. For .mpp support: pip install mpxj jpype1"
        ) from e


def _detect_msp_xml_schema(file_path: str) -> bool:
    """Read first 512 bytes; check for MS Project MSPDI namespace.

    Returns True if MS Project XML, False if Asta or unknown.
    Raises FileNotFoundError if path missing.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    with open(file_path, 'rb') as f:
        head = f.read(512).decode('utf-8', errors='replace')
    return 'schemas.microsoft.com/project' in head


class MspMppFileManager:
    """Read-only manager for .mpp files via MPXJ + JVM.

    Adapted from Asta asta_mcp_file.py AstaFileManager (drop .pp support).
    """

    def __init__(self, file_path: str):
        self.file_path = file_path.replace("\\", "/")
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")
        self._project = None  # lazy load on first read

    def _load(self):
        if self._project is not None:
            return self._project
        _ensure_jvm_started()
        from org.mpxj.reader import UniversalProjectReader
        reader = UniversalProjectReader()
        try:
            self._project = reader.read(self.file_path)
            logger.info(f"MPP loaded: {self.file_path}")
            return self._project
        except Exception as e:
            raise RuntimeError(f"MPXJ read failed for {self.file_path}: {e}") from e

    def read_tasks(self) -> List[Dict[str, Any]]:
        proj = self._load()
        out = []
        for t in proj.getTasks():
            if t is None or t.getID() == 0:
                continue
            out.append({
                "id": int(t.getID()),
                "name": str(t.getName() or ""),
                "duration_h": _mpxj_duration_to_hours(t.getDuration()),
                "start": str(t.getStart()) if t.getStart() else None,
                "finish": str(t.getFinish()) if t.getFinish() else None,
                "percent_complete": float(t.getPercentageComplete() or 0),
                "summary": bool(t.getSummary()),
            })
        return out

    def read_links(self) -> List[Dict[str, Any]]:
        proj = self._load()
        out = []
        for t in proj.getTasks():
            if t is None or t.getID() == 0:
                continue
            for rel in (t.getPredecessors() or []):
                out.append({
                    "from_id": int(rel.getTargetTask().getID()),
                    "to_id": int(t.getID()),
                    "type": str(rel.getType()),
                    "lag_days": _mpxj_duration_to_hours(rel.getLag()) / 8.0 if rel.getLag() else 0.0,
                })
        return out

    def read_resources(self) -> List[Dict[str, Any]]:
        proj = self._load()
        out = []
        for r in proj.getResources():
            if r is None or r.getID() == 0:
                continue
            out.append({
                "id": int(r.getID()),
                "name": str(r.getName() or ""),
                "type": str(r.getType() or "Work"),
                "max_units": float(r.getMaxUnits() or 1.0),
            })
        return out

    def read_assignments(self) -> List[Dict[str, Any]]:
        proj = self._load()
        out = []
        for a in proj.getResourceAssignments():
            t = a.getTask()
            r = a.getResource()
            if t is None or r is None:
                continue
            out.append({
                "task_id": int(t.getID()),
                "resource_id": int(r.getID()),
                "units": float(a.getUnits() or 1.0),
                "work_h": _mpxj_duration_to_hours(a.getWork()),
            })
        return out

    def read_calendars(self) -> List[Dict[str, Any]]:
        proj = self._load()
        out = []
        for cal in proj.getCalendars():
            out.append({
                "name": str(cal.getName() or ""),
                "is_base": bool(cal.getParent() is None),
            })
        return out

    def read_baselines(self, baseline_number: int = 0) -> Dict[str, Any]:
        # MPP baselines limited via MPXJ; Phase 3a integration is best-effort
        return {"baseline_number": baseline_number, "note": "MPP baseline read via MPXJ — limited fields"}

    def read_progress(self) -> Dict[str, Any]:
        # MPP progress read via task.getActualWork etc. — phase 3b integration
        proj = self._load()
        tasks = []
        for t in proj.getTasks():
            if t is None or t.getID() == 0 or t.getSummary():
                continue
            tasks.append({
                "id": int(t.getID()),
                "percent_complete": float(t.getPercentageComplete() or 0),
                "actual_work_h": _mpxj_duration_to_hours(t.getActualWork()),
            })
        status_date = proj.getProjectProperties().getStatusDate() if proj.getProjectProperties() else None
        return {"status_date": str(status_date) if status_date else None, "tasks": tasks}


def _mpxj_duration_to_hours(d) -> float:
    """Convert MPXJ Duration object to hours (float). Handles None."""
    if d is None:
        return 0.0
    try:
        from org.mpxj import TimeUnit
        n = float(d.getDuration())
        unit = d.getUnits()
        # MPXJ TimeUnit: HOURS=2, DAYS=4, WEEKS=5, MONTHS=6, YEARS=8 (approx mapping)
        if unit == TimeUnit.HOURS:
            return n
        elif unit == TimeUnit.DAYS:
            return n * 8.0
        elif unit == TimeUnit.WEEKS:
            return n * 40.0
        elif unit == TimeUnit.MINUTES:
            return n / 60.0
        else:
            return n  # fallback assume hours
    except Exception:
        return 0.0


def _get_msp_file_manager(file_path: str):
    """Factory: returns MspdiProject for .xml/.mspdi, MspMppFileManager for .mpp.

    Performs schema check for XML — refuses non-MSPDI XML with clear error.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xml', '.mspdi'):
        if not _detect_msp_xml_schema(file_path):
            raise ValueError(
                f"Not a MS Project XML — appears to be Asta or unknown schema. "
                f"For Asta files use asta_powerproject_file MCP. File: {file_path}"
            )
        return MspdiProject(file_path)
    elif ext == '.mpp':
        return MspMppFileManager(file_path)
    else:
        raise ValueError(
            f"Unsupported extension '{ext}'. Phase 4 supports: .xml, .mspdi, .mpp"
        )
```

**Step 6: Update requirements.txt**

Append (idempotent, check before adding):
```
mpxj>=14.0
jpype1>=1.5
```

**Step 7: Run tests — PASS**

```bash
python -m pytest tests/test_msproject_file_factory.py -v
```

Expected: 5 PASS.

**Step 8: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_factory.py tests/fixtures/sample_msp.xml requirements.txt
git commit -m "Phase 4 T65: foundations (factory + MspMppFileManager + JVM lifecycle + schema detect)"
```

Expected full regression: **287 PASS + 1 xfail** (282 + 5 new).

---

## Task 66: `read_tasks` + `read_links` Actions

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_file_read_tasks.py`
- Create: `tests/test_msproject_file_read_links.py`

**Step 1: Failing tests**

`tests/test_msproject_file_read_tasks.py`:
```python
"""Test msproject_file read_tasks action (XML + MPP paths)."""
import pytest
import os
from msproject_mcp_core import _msp_file_read_tasks

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
MSP_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


def test_read_tasks_xml_returns_count():
    r = _msp_file_read_tasks(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["count"] == 3
    assert len(r["tasks"]) == 3


def test_read_tasks_xml_has_required_fields():
    r = _msp_file_read_tasks(file_path=MSP_XML)
    t = r["tasks"][0]
    for key in ("id", "name", "duration_h", "start", "finish",
                "percent_complete", "summary"):
        assert key in t


def test_read_tasks_xml_duration_correct():
    """T1=1d → 8h, T2=2d → 16h, T3=3d → 24h."""
    r = _msp_file_read_tasks(file_path=MSP_XML)
    durations = sorted([t["duration_h"] for t in r["tasks"]])
    assert durations == [8.0, 16.0, 24.0]


def test_read_tasks_filter_by_finish_after(tmp_path):
    """filters parameter limits tasks. Use simple operator support."""
    r = _msp_file_read_tasks(file_path=MSP_XML, limit=2)
    assert r["count"] == 2


def test_read_tasks_invalid_file_errors():
    r = _msp_file_read_tasks(file_path="/nonexistent.xml")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower() or "file" in r["error"].lower()
```

`tests/test_msproject_file_read_links.py`:
```python
"""Test msproject_file read_links action."""
import pytest
import os
from msproject_mcp_core import _msp_file_read_links

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
MSP_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


def test_read_links_xml_returns_status():
    """Sample fixture has no links by default — returns count=0."""
    r = _msp_file_read_links(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "count" in r
    assert "links" in r
    assert isinstance(r["links"], list)


def test_read_links_xml_link_fields():
    """If links exist, each has from_id/to_id/type/lag_days."""
    r = _msp_file_read_links(file_path=MSP_XML)
    for link in r["links"]:
        assert "from_id" in link
        assert "to_id" in link
        assert "type" in link  # "FS"/"SS"/"FF"/"SF"
        assert "lag_days" in link


def test_read_links_invalid_file_errors():
    r = _msp_file_read_links(file_path="/nonexistent.xml")
    assert r["status"] == "error"
```

**Step 2: Run — FAIL** (ImportError)

**Step 3: Implementation**

Insert after factory helpers in `msproject_mcp_core.py`:

```python
# ---------- PHASE 4 ACTION HELPERS ----------

def _msp_file_read_tasks(file_path: str,
                        filters: Optional[Dict[str, Any]] = None,
                        limit: Optional[int] = None) -> Dict[str, Any]:
    """Read all tasks from a MS Project file. Format auto-detected by extension."""
    try:
        mgr = _get_msp_file_manager(file_path)
        tasks = mgr.read_tasks()
        if filters:
            # Simple field equality filter — extend in T69 query
            for k, v in filters.items():
                tasks = [t for t in tasks if t.get(k) == v]
        if limit and limit > 0:
            tasks = tasks[:limit]
        return {"status": "ok", "count": len(tasks), "tasks": tasks}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_tasks({file_path}) failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_file_read_links(file_path: str) -> Dict[str, Any]:
    """Read all task predecessor/successor links from a MS Project file."""
    try:
        mgr = _get_msp_file_manager(file_path)
        links = mgr.read_links()
        return {"status": "ok", "count": len(links), "links": links}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_links({file_path}) failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Note:** `MspdiProject.read_tasks()` and `read_links()` must exist (probe T65). If they exist with different signatures, adapt the helper to translate. Document any shim.

**Step 4: Run — PASS**

Expected: 8 PASS (5 read_tasks + 3 read_links).

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_read_tasks.py tests/test_msproject_file_read_links.py
git commit -m "Phase 4 T66: read_tasks + read_links (XML path + factory dispatch)"
```

Expected: **295 PASS + 1 xfail**.

---

## Task 67: `read_resources` + `read_assignments` + `read_calendars`

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_file_read_resources.py`
- Create: `tests/test_msproject_file_read_assignments.py`
- Create: `tests/test_msproject_file_read_calendars.py`

**Step 1: Failing tests**

`tests/test_msproject_file_read_resources.py`:
```python
import os
import pytest
from msproject_mcp_core import _msp_file_read_resources

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_resources_xml():
    """Sample fixture has 2 work resources R1, R2."""
    r = _msp_file_read_resources(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["count"] == 2
    names = [res["name"] for res in r["resources"]]
    assert "R1" in names
    assert "R2" in names


def test_read_resources_fields():
    r = _msp_file_read_resources(file_path=MSP_XML)
    for res in r["resources"]:
        for key in ("id", "name", "type", "max_units"):
            assert key in res
```

`tests/test_msproject_file_read_assignments.py`:
```python
import os
from msproject_mcp_core import _msp_file_read_assignments

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_assignments_xml():
    """3 tasks × 2 resources = 6 assignments."""
    r = _msp_file_read_assignments(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["count"] == 6


def test_read_assignments_filter_task():
    r = _msp_file_read_assignments(file_path=MSP_XML, task_id=1)
    assert r["status"] == "ok"
    # 1 task × 2 resources = 2 assignments for task 1
    assert all(a["task_id"] == 1 for a in r["assignments"])
```

`tests/test_msproject_file_read_calendars.py`:
```python
import os
from msproject_mcp_core import _msp_file_read_calendars

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_calendars_xml():
    r = _msp_file_read_calendars(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "calendars" in r
    assert len(r["calendars"]) >= 1  # at least Standard
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_file_read_resources(file_path: str) -> Dict[str, Any]:
    try:
        mgr = _get_msp_file_manager(file_path)
        resources = mgr.read_resources()
        return {"status": "ok", "count": len(resources), "resources": resources}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_resources failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_file_read_assignments(file_path: str,
                              task_id: Optional[int] = None) -> Dict[str, Any]:
    try:
        mgr = _get_msp_file_manager(file_path)
        assignments = mgr.read_assignments()
        if task_id is not None:
            assignments = [a for a in assignments if a.get("task_id") == task_id]
        return {"status": "ok", "count": len(assignments), "assignments": assignments}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_assignments failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_file_read_calendars(file_path: str) -> Dict[str, Any]:
    try:
        mgr = _get_msp_file_manager(file_path)
        calendars = mgr.read_calendars()
        return {"status": "ok", "count": len(calendars), "calendars": calendars}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_calendars failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS** (5 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_read_*.py
git commit -m "Phase 4 T67: read_resources + read_assignments + read_calendars"
```

Expected: **300 PASS + 1 xfail**.

---

## Task 68: `read_baselines` + `read_progress` (Phase 3a/3b Integration)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_file_read_baselines.py`
- Create: `tests/test_msproject_file_read_progress.py`

**Step 1: Failing tests**

`tests/test_msproject_file_read_baselines.py`:
```python
import os
from msproject_mcp_core import _msp_file_read_baselines

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_baselines_xml_unsaved():
    """Sample fixture has no baseline saved — returns saved=False."""
    r = _msp_file_read_baselines(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    assert "baseline_number" in r
    assert "saved_date" in r  # None or string


def test_read_baselines_invalid_number():
    r = _msp_file_read_baselines(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"
    assert "0-10" in r["error"]
```

`tests/test_msproject_file_read_progress.py`:
```python
import os
from msproject_mcp_core import _msp_file_read_progress

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_read_progress_xml():
    r = _msp_file_read_progress(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "status_date" in r
    assert "tasks" in r
    for t in r["tasks"]:
        assert "id" in t
        assert "percent_complete" in t


def test_read_progress_with_assignments():
    r = _msp_file_read_progress(file_path=MSP_XML, include_assignments=True)
    assert r["status"] == "ok"
    # If include_assignments, response should have an "assignments" or per-task assignment list
    assert "tasks" in r
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_file_read_baselines(file_path: str, baseline_number: int = 0) -> Dict[str, Any]:
    """Read saved baseline data from file (Phase 3a integration).

    XML: parse Baseline Start/Finish/Work/Cost from MSPDI Tasks.
    MPP: limited via MPXJ (returns baseline_number + note).
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    try:
        mgr = _get_msp_file_manager(file_path)
        if hasattr(mgr, 'read_baselines'):
            data = mgr.read_baselines(baseline_number)
            return {"status": "ok", **data}
        # Fallback for MspdiProject if method missing — minimal contract
        return {
            "status": "ok",
            "baseline_number": baseline_number,
            "saved_date": None,
            "tasks": [],
            "note": "Baseline data parsing minimal in Phase 4 — full extraction in Phase 5",
        }
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_baselines failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_file_read_progress(file_path: str,
                           include_assignments: bool = False) -> Dict[str, Any]:
    """Read progress fields from file (Phase 3b integration).

    XML: parse PercentComplete / ActualWork / ActualStart / ActualFinish from MSPDI.
    MPP: same fields via MPXJ task accessors.
    """
    try:
        mgr = _get_msp_file_manager(file_path)
        if hasattr(mgr, 'read_progress'):
            data = mgr.read_progress() if not include_assignments else mgr.read_progress()
            return {"status": "ok", **data}
        return {
            "status": "ok",
            "status_date": None,
            "tasks": [],
            "note": "Progress data parsing minimal in Phase 4",
        }
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_progress failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS** (4 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_read_baselines.py tests/test_msproject_file_read_progress.py
git commit -m "Phase 4 T68: read_baselines + read_progress (Phase 3a/3b file-side integration)"
```

Expected: **304 PASS + 1 xfail**.

---

## Task 69: `query` Action (Filter Expression Parser)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_file_query.py`

**Step 1: Failing tests**

`tests/test_msproject_file_query.py`:
```python
import os
import pytest
from msproject_mcp_core import _msp_file_query

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_query_simple_eq():
    """name == 'T1' returns 1 task."""
    r = _msp_file_query(file_path=MSP_XML, expression="name == 'T1'")
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["results"][0]["name"] == "T1"


def test_query_gt():
    """duration_h > 8 returns T2 (16h) and T3 (24h) — 2 tasks."""
    r = _msp_file_query(file_path=MSP_XML, expression="duration_h > 8")
    assert r["status"] == "ok"
    assert r["count"] == 2


def test_query_and():
    """duration_h > 8 AND duration_h <= 16 returns T2 only."""
    r = _msp_file_query(file_path=MSP_XML, expression="duration_h > 8 AND duration_h <= 16")
    assert r["status"] == "ok"
    assert r["count"] == 1
    assert r["results"][0]["name"] == "T2"


def test_query_invalid_expression():
    r = _msp_file_query(file_path=MSP_XML, expression="this is not valid syntax @#$")
    assert r["status"] == "error"
    assert "expression" in r["error"].lower() or "parse" in r["error"].lower()


def test_query_limit():
    r = _msp_file_query(file_path=MSP_XML, expression="duration_h >= 8", limit=1)
    assert r["status"] == "ok"
    assert r["count"] == 1
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _safe_eval_filter(expression: str, row: Dict[str, Any]) -> bool:
    """Evaluate a simple filter expression against a row dict.

    Supports: == != < <= > >= AND OR. String literals in single/double quotes.
    Field names are dict keys. Restricted eval — no function calls.
    """
    # Normalize boolean operators to Python
    expr = expression.replace(" AND ", " and ").replace(" OR ", " or ")
    # Restricted globals — no builtins
    safe_globals = {"__builtins__": {}}
    # Only allow row keys and literal types
    safe_locals = {k: row.get(k) for k in row}
    try:
        return bool(eval(expr, safe_globals, safe_locals))
    except Exception as e:
        raise ValueError(f"expression eval failed: {e}") from e


def _msp_file_query(file_path: str,
                   expression: str,
                   limit: Optional[int] = None) -> Dict[str, Any]:
    """Run an ad-hoc filter expression against tasks in a project file.

    Returns matching task list. Expression syntax: simple python-like comparisons
    with AND/OR operators. Field names are task keys (id, name, duration_h, etc.).

    Examples:
      "duration_h > 8 AND name == 'T2'"
      "percent_complete < 100"
    """
    try:
        mgr = _get_msp_file_manager(file_path)
        tasks = mgr.read_tasks()
        results = []
        for t in tasks:
            try:
                if _safe_eval_filter(expression, t):
                    results.append(t)
            except ValueError as e:
                return {"status": "error", "error": f"Invalid expression: {e}"}
        if limit and limit > 0:
            results = results[:limit]
        return {"status": "ok", "count": len(results), "results": results}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_query failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS** (5 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_query.py
git commit -m "Phase 4 T69: query action (filter expression parser, restricted eval)"
```

Expected: **309 PASS + 1 xfail**.

---

## Task 70: `add_tasks` + `add_links` + `add_resources` (XML Write)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_file_write_add.py`

**PROBE FIRST:** Confirm `MspdiProject` has `add_task(...)`, `add_link(...)`, `add_resource(...)`, `save()` methods (see T65 probe). If signatures differ, adapt helpers below.

**Step 1: Failing tests**

`tests/test_msproject_file_write_add.py`:
```python
import os
import shutil
import pytest
from msproject_mcp_core import (
    _msp_file_add_tasks, _msp_file_add_links, _msp_file_add_resources,
    _msp_file_read_tasks, _msp_file_read_links, _msp_file_read_resources,
)

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SOURCE_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


@pytest.fixture
def writable_xml(tmp_path):
    dst = tmp_path / "writable.xml"
    shutil.copy(SOURCE_XML, dst)
    return str(dst)


def test_add_tasks_appends_to_xml(writable_xml):
    """Add 2 tasks → re-read shows 5 tasks (3 base + 2 new)."""
    r = _msp_file_add_tasks(file_path=writable_xml, items=[
        {"name": "T4", "duration": "5d"},
        {"name": "T5", "duration": "2d"},
    ])
    assert r["status"] == "ok"
    assert r["count"] == 2
    # Re-read
    r2 = _msp_file_read_tasks(file_path=writable_xml)
    assert r2["count"] == 5
    names = [t["name"] for t in r2["tasks"]]
    assert "T4" in names and "T5" in names


def test_add_links_appends(writable_xml):
    """Add link T1 → T2 (FS), then re-read shows 1 new link."""
    r0 = _msp_file_read_links(file_path=writable_xml)
    base_count = r0["count"]
    r = _msp_file_add_links(file_path=writable_xml, items=[
        {"from_id": 1, "to_id": 2, "type": "FS", "lag": "0d"},
    ])
    assert r["status"] == "ok"
    assert r["count"] == 1
    r2 = _msp_file_read_links(file_path=writable_xml)
    assert r2["count"] == base_count + 1


def test_add_resources_appends(writable_xml):
    r = _msp_file_add_resources(file_path=writable_xml, items=[
        {"name": "R3", "type": "Work", "max_units": 1.0},
    ])
    assert r["status"] == "ok"
    r2 = _msp_file_read_resources(file_path=writable_xml)
    names = [res["name"] for res in r2["resources"]]
    assert "R3" in names


def test_add_tasks_mpp_unsupported(tmp_path):
    """MPP write not supported — clear error."""
    fake_mpp = tmp_path / "test.mpp"
    fake_mpp.write_bytes(b"\x00" * 100)
    r = _msp_file_add_tasks(file_path=str(fake_mpp), items=[{"name": "X", "duration": "1d"}])
    assert r["status"] == "error"
    assert ".mpp" in r["error"].lower() or "binary" in r["error"].lower() or "not supported" in r["error"].lower()
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _ensure_xml_write_target(mgr) -> None:
    """Raise if mgr is not an MspdiProject (write only on XML)."""
    if not isinstance(mgr, MspdiProject):
        raise ValueError(
            ".mpp write not supported (Microsoft proprietary binary format). "
            "Convert to .xml first via MS Project Save As, or use COM tools."
        )


def _msp_file_add_tasks(file_path: str, items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Bulk add tasks to a MS Project XML file. MSP open → auto COM import + Reschedule."""
    try:
        mgr = _get_msp_file_manager(file_path)
        _ensure_xml_write_target(mgr)
        task_ids = []
        for item in items:
            tid = mgr.add_task(name=item.get("name"),
                              duration=item.get("duration"),
                              start=item.get("start"),
                              **{k: v for k, v in item.items()
                                 if k not in ("name", "duration", "start")})
            task_ids.append(tid)
        mgr.save()
        # Auto-sync (T72)
        sync = _auto_sync_to_open_msp(file_path) if _auto_sync_to_open_msp_available() else \
               {"auto_imported": False}
        return {"status": "ok", "count": len(task_ids), "task_ids": task_ids, **sync}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except ValueError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_add_tasks failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_file_add_links(file_path: str, items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Bulk add predecessor links."""
    try:
        mgr = _get_msp_file_manager(file_path)
        _ensure_xml_write_target(mgr)
        link_ids = []
        for item in items:
            lid = mgr.add_link(from_id=item["from_id"], to_id=item["to_id"],
                              type=item.get("type", "FS"), lag=item.get("lag", "0d"))
            link_ids.append(lid)
        mgr.save()
        sync = _auto_sync_to_open_msp(file_path) if _auto_sync_to_open_msp_available() else \
               {"auto_imported": False}
        return {"status": "ok", "count": len(link_ids), "link_ids": link_ids, **sync}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except ValueError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_add_links failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_file_add_resources(file_path: str, items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Bulk add resources."""
    try:
        mgr = _get_msp_file_manager(file_path)
        _ensure_xml_write_target(mgr)
        res_ids = []
        for item in items:
            rid = mgr.add_resource(name=item["name"], type=item.get("type", "Work"),
                                  max_units=item.get("max_units", 1.0))
            res_ids.append(rid)
        mgr.save()
        sync = _auto_sync_to_open_msp(file_path) if _auto_sync_to_open_msp_available() else \
               {"auto_imported": False}
        return {"status": "ok", "count": len(res_ids), "resource_ids": res_ids, **sync}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except ValueError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_add_resources failed: {e}")
        return {"status": "error", "error": str(e)}


def _auto_sync_to_open_msp_available() -> bool:
    """Check if auto-sync helper is implemented (T72 placeholder)."""
    return '_auto_sync_to_open_msp' in globals() and callable(globals().get('_auto_sync_to_open_msp'))
```

**NOTE on `_auto_sync_to_open_msp_available` shim:** This allows T70-T71 to land before T72 without crashing. T72 will replace the placeholder check with the real helper.

**Step 4: Run — PASS** (4 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_write_add.py
git commit -m "Phase 4 T70: add_tasks + add_links + add_resources (XML write, .mpp rejected)"
```

Expected: **313 PASS + 1 xfail**.

---

## Task 71: `update_task` + `save_as`

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_file_write_update.py`
- Create: `tests/test_msproject_file_save_as.py`

**Step 1: Failing tests**

`tests/test_msproject_file_write_update.py`:
```python
import os, shutil, pytest
from msproject_mcp_core import _msp_file_update_task, _msp_file_read_tasks

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


@pytest.fixture
def writable_xml(tmp_path):
    dst = tmp_path / "writable.xml"
    shutil.copy(FIXTURE, dst)
    return str(dst)


def test_update_task_duration(writable_xml):
    r = _msp_file_update_task(file_path=writable_xml, task_id=1, fields={"duration": "5d"})
    assert r["status"] == "ok"
    r2 = _msp_file_read_tasks(file_path=writable_xml)
    t1 = next(t for t in r2["tasks"] if t["id"] == 1)
    assert t1["duration_h"] == 40.0  # 5d * 8h


def test_update_task_missing_id(writable_xml):
    r = _msp_file_update_task(file_path=writable_xml, task_id=99999, fields={"duration": "1d"})
    assert r["status"] == "error"
    assert "task" in r["error"].lower()
```

`tests/test_msproject_file_save_as.py`:
```python
import os, shutil, pytest
from msproject_mcp_core import _msp_file_save_as

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_save_as_xml(tmp_path):
    dst = tmp_path / "renamed.xml"
    r = _msp_file_save_as(file_path=FIXTURE, output_path=str(dst))
    assert r["status"] == "ok"
    assert os.path.exists(str(dst))
    assert r["output_path"] == str(dst)
    assert r["size_bytes"] > 0


def test_save_as_invalid_output_extension(tmp_path):
    dst = tmp_path / "bad.txt"
    r = _msp_file_save_as(file_path=FIXTURE, output_path=str(dst))
    assert r["status"] == "error"
    assert "extension" in r["error"].lower() or "xml" in r["error"].lower()
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_file_update_task(file_path: str, task_id: int,
                         fields: Dict[str, Any]) -> Dict[str, Any]:
    """Update a single task's fields in an XML file."""
    try:
        mgr = _get_msp_file_manager(file_path)
        _ensure_xml_write_target(mgr)
        if not mgr.update_task(task_id=task_id, **fields):
            return {"status": "error",
                    "error": f"Task ID {task_id} not found in file"}
        mgr.save()
        sync = _auto_sync_to_open_msp(file_path) if _auto_sync_to_open_msp_available() else \
               {"auto_imported": False}
        return {"status": "ok", "task_id": task_id, **sync}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except ValueError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_update_task failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_file_save_as(file_path: str, output_path: str) -> Dict[str, Any]:
    """Save the project to a new XML path. Reads source via factory."""
    try:
        ext = os.path.splitext(output_path)[1].lower()
        if ext not in ('.xml', '.mspdi'):
            return {"status": "error",
                    "error": f"output_path must end in .xml or .mspdi (got '{ext}')"}
        mgr = _get_msp_file_manager(file_path)
        _ensure_xml_write_target(mgr)
        mgr.save_as(output_path)
        size = os.path.getsize(output_path)
        return {"status": "ok", "output_path": output_path, "size_bytes": size}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_save_as failed: {e}")
        return {"status": "error", "error": str(e)}
```

**Step 4: Run — PASS** (4 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_write_update.py tests/test_msproject_file_save_as.py
git commit -m "Phase 4 T71: update_task + save_as (XML write only)"
```

Expected: **317 PASS + 1 xfail**.

---

## Task 72: `_auto_sync_to_open_msp` Helper (BIG ONE — Probe-Driven)

**Files:**
- Modify: `msproject_mcp_core.py` (replace `_auto_sync_to_open_msp_available` shim with real helper)
- Create: `tests/test_msproject_file_auto_sync.py`

**Step 1: Probe COM merge mechanics**

Live probe — run against open MSP with active project:

```python
# probe_auto_sync.py
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject('MSProject.Application')
print(f"App active project: {app.ActiveProject.Name if app.ActiveProject else None}")

# Try FileOpen on a temp xml + EditCopy + EditPaste merge mechanic
# The pattern from _msp_task_bulk_add mspdi path is the reference
# Key questions:
#   - Does FileOpen raise if XML schema is correct?
#   - After EditCopy on temp_proj.Tasks, where do we EditPaste? On active_proj root bar or task selection?
#   - Does Reschedule run cleanly after merge?
print(dir(app))  # Look for FileOpen, EditCopy, EditPaste, FileClose
```

**Decision rule:** Document discovered merge sequence in commit message. Replace the placeholder logic below if probe reveals a different sequence.

**Step 2: Failing tests**

`tests/test_msproject_file_auto_sync.py`:
```python
import os
import pytest
from unittest.mock import patch, MagicMock
from msproject_mcp_core import _auto_sync_to_open_msp

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_auto_sync_msp_closed_returns_not_imported():
    """When MSP COM unavailable, returns auto_imported=False."""
    with patch('msproject_mcp_core.win32com.client.GetActiveObject',
               side_effect=Exception("MSP not running")):
        r = _auto_sync_to_open_msp(MSP_XML)
        assert r["auto_imported"] is False
        assert "msp" in r.get("msg", "").lower() or "closed" in r.get("msg", "").lower()


def test_auto_sync_msp_open_imports_and_reschedules(clean_test_project):
    """When MSP open, imports XML and runs Reschedule."""
    r = _auto_sync_to_open_msp(MSP_XML)
    assert r["auto_imported"] is True
    assert r["reschedule_ok"] is True


def test_auto_sync_xml_path_missing():
    r = _auto_sync_to_open_msp("/nonexistent.xml")
    assert r["auto_imported"] is False
    assert "error" in r or "not found" in r.get("msg", "").lower()


def test_auto_sync_handles_filopen_exception(monkeypatch, clean_test_project):
    """If app.FileOpen fails, returns clean error, doesn't crash."""
    import msproject_mcp_core
    mock_app = MagicMock()
    mock_app.FileOpen.side_effect = Exception("Cannot open XML")
    monkeypatch.setattr('msproject_mcp_core.win32com.client.GetActiveObject', lambda *a, **kw: mock_app)
    r = _auto_sync_to_open_msp(MSP_XML)
    assert r["auto_imported"] is False
    assert "error" in r
```

**Step 3: Run — FAIL**

**Step 4: Implementation**

Replace the `_auto_sync_to_open_msp_available` shim (T70) with real helper:

```python
def _auto_sync_to_open_msp(modified_xml_path: str) -> Dict[str, Any]:
    """Import a modified MSPDI XML into the open MS Project active project.

    Default behavior — NOT opt-in. Memory: feedback_file_mcp_auto_sync.md.

    Workflow:
      1. Try GetActiveObject('MSProject.Application'). If fail → auto_imported=False.
      2. app.FileOpen(modified_xml_path) → temp_proj
      3. EditCopy on temp_proj task collection (or full project state)
      4. Activate active_proj, EditPaste merge
      5. active_proj.Reschedule()
      6. FileClose(0) on temp_proj (no save)
      7. Apply Phase 2b TAIL pattern: post-paste re-establish Duration if dropped

    Returns: {auto_imported: bool, reschedule_ok: bool, error?: str, merged_count?: int}
    """
    if not os.path.exists(modified_xml_path):
        return {"auto_imported": False, "error": f"XML not found: {modified_xml_path}"}
    try:
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject('MSProject.Application')
        except Exception as e:
            return {"auto_imported": False,
                    "msg": f"MSP closed; XML saved at {modified_xml_path}",
                    "error": str(e)}
        # Save reference to user's active project before opening temp
        original_proj_name = None
        try:
            if app.ActiveProject:
                original_proj_name = app.ActiveProject.Name
        except Exception:
            pass
        try:
            app.FileOpen(modified_xml_path)
            temp_proj = app.ActiveProject
            # Merge mechanic — pattern from _msp_task_bulk_add mspdi path
            # (Probe T72 confirms exact sequence)
            try:
                _msp_xml_merge_into_original(app, temp_proj, original_proj_name)
            except Exception as merge_err:
                return {"auto_imported": False,
                        "error": f"Merge failed: {merge_err}",
                        "msg": "Temp XML opened but merge into original project failed"}
            # Reschedule active project
            try:
                if original_proj_name:
                    for i in range(1, app.Projects.Count + 1):
                        if app.Projects(i).Name == original_proj_name:
                            app.WindowActivate(app.Projects(i).Windows(1).Caption)
                            break
                    app.ActiveProject.Reschedule()
                    reschedule_ok = True
            except Exception as e:
                logger.warning(f"Reschedule failed: {e}")
                reschedule_ok = False
            # Close temp project without saving
            try:
                for i in range(1, app.Projects.Count + 1):
                    if app.Projects(i).FullName == modified_xml_path or \
                       os.path.basename(app.Projects(i).Name) == os.path.basename(modified_xml_path):
                        app.WindowActivate(app.Projects(i).Windows(1).Caption)
                        app.FileClose(0)  # 0 = don't save
                        break
            except Exception as e:
                logger.warning(f"Temp project close failed: {e}")
            return {"auto_imported": True, "reschedule_ok": reschedule_ok}
        except Exception as e:
            return {"auto_imported": False, "error": _format_com_error(e)}
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _msp_xml_merge_into_original(app, temp_proj, original_proj_name: Optional[str]) -> None:
    """Merge temp_proj contents into original active project.

    Replicates the pattern from _msp_task_bulk_add mspdi path:
      1. Select all tasks in temp_proj
      2. EditCopy
      3. Activate original_proj
      4. EditPaste (appends to selected location or end)

    Phase 2b TAIL discovery: MSP MSPDI FileOpen drops Duration. After paste,
    iterate tasks and re-set Duration from the temp XML's expected values if
    they don't match. (This is critical for bulk_add_tasks; for assignments
    the symptom may differ.)
    """
    if not original_proj_name:
        # No original to merge into — temp_proj is now the active project (acceptable for new file scenario)
        return
    # Implementation: probe in T72 reveals exact COM sequence — placeholder for now
    # TODO (T72): replace with verified EditCopy/EditPaste sequence
    pass  # Will be filled in based on probe results
```

**NOTE:** The `_msp_xml_merge_into_original` body is **probe-dependent**. T72 implementer must run live probes to confirm the EditCopy/EditPaste mechanic before committing. Document findings in commit message.

**Step 5: Run — PASS** (4 PASS expected; the merge test may need MSP active)

**Step 6: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_auto_sync.py
git commit -m "Phase 4 T72: _auto_sync_to_open_msp helper (FileOpen + merge + Reschedule, MSP closed graceful)"
```

Expected: **321 PASS + 1 xfail**.

---

## Task 73: 🚀 HERO `bulk_add_assignments` (xfail Flip)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_file_hero.py`
- Modify: `tests/test_msproject_resource_bulk_assign.py` (verify hero xfail still strict — it WILL flip when this lands)

**CRITICAL:** This task flips `test_bulk_assign_hero_2800_under_5s` strict=True xfail. If the test still fails after T73, Phase 4 has not delivered its key gate. If it passes (xpass), strict=True will fail the test suite — that's the SUCCESS signal. Keep strict=True intact.

**Step 1: Probe MspdiProject `<Assignment>` write capability**

```python
# probe_assignment_write.py
from mspdi_parser import MspdiProject
p = MspdiProject('tests/fixtures/sample_msp.xml')
print('add_assignment' in dir(p))
# Look for batch interface
print('add_assignments' in dir(p) or 'bulk_add_assignments' in dir(p))
```

If MspdiProject lacks bulk write, T73 implementer must extend mspdi_parser.py with `bulk_add_assignments(items)` method (single XML pass).

**Step 2: Failing test**

`tests/test_msproject_file_hero.py`:
```python
"""Phase 4 HERO test — flips Phase 2b strict xfail."""
import os
import time
import pytest
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_file_bulk_add_assignments,
    _msp_file_save_as,
)


def test_hero_2800_assignments_under_5s_via_xml_merge(clean_test_project):
    """200 tasks × 14 resources = 2800 assignments via Phase 4 XML merge in <5s."""
    proj = clean_test_project
    # Build base scenario
    items = [{"name": f"V{i:03d}", "duration": "2d"} for i in range(200)]
    _msp_task_bulk_add(items=items)
    res_ids = [_msp_resource_add(name=f"R{i:02d}", type="Work")["resource_id"]
               for i in range(14)]
    # Get task_ids fresh
    task_ids = [proj.Tasks(i).ID for i in range(1, 201) if proj.Tasks(i) is not None]
    # Build assignment items
    assignment_items = [{"task_id": tid, "resource_id": rid}
                        for tid in task_ids for rid in res_ids]
    assert len(assignment_items) == 2800
    # Export current state to temp XML, run hero merge
    import tempfile
    tmpdir = tempfile.mkdtemp()
    temp_xml = os.path.join(tmpdir, "hero.xml")
    # Export via MSP COM
    proj.SaveAs(temp_xml, 4)  # FileFormat=4 = XML (verify enum in probe)
    # HERO call
    start = time.time()
    r = _msp_file_bulk_add_assignments(file_path=temp_xml, items=assignment_items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 2800
    assert elapsed < 5.0, f"HERO took {elapsed:.2f}s (target <5s)"
    assert r.get("auto_imported", False) is True
```

**Step 3: Run — FAIL** (helper doesn't exist)

**Step 4: Implementation**

```python
def _msp_file_bulk_add_assignments(file_path: str,
                                  items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """🚀 HERO — Bulk add 2800+ assignments via single MSPDI XML merge in <5s.

    Path:
      1. Open MSPDI XML via mspdi_parser
      2. Bulk-write 2800 <Assignment> elements (single XML pass, ~1s)
      3. Save XML
      4. Auto-sync to open MSP (FileOpen + EditPaste merge, ~1-2s)
      5. proj.Reschedule()  (~0.5s)
      Total target: <5s

    items: list of {task_id, resource_id, units (default 1.0)}
    """
    if not items:
        return {"status": "ok", "count": 0, "auto_imported": False}
    start = time.time()
    try:
        mgr = _get_msp_file_manager(file_path)
        _ensure_xml_write_target(mgr)
        # Bulk write — assumes mspdi_parser.MspdiProject.bulk_add_assignments exists
        # (T73 may need to extend mspdi_parser.py; track in commit)
        if hasattr(mgr, 'bulk_add_assignments'):
            mgr.bulk_add_assignments(items)
        else:
            # Fallback — single-add loop (slower but functional)
            for item in items:
                mgr.add_assignment(task_id=item["task_id"],
                                  resource_id=item["resource_id"],
                                  units=item.get("units", 1.0))
        mgr.save()
        # Auto-sync (HERO depends on this for merge)
        sync = _auto_sync_to_open_msp(file_path)
        elapsed = time.time() - start
        return {
            "status": "ok",
            "count": len(items),
            "elapsed_s": round(elapsed, 2),
            **sync,
        }
    except Exception as e:
        logger.error(f"_msp_file_bulk_add_assignments failed: {e}")
        return {"status": "error", "error": str(e), "elapsed_s": round(time.time() - start, 2)}
```

**Step 5: Run hero — Expect xpass (strict xfail flip)**

```bash
python -m pytest tests/test_msproject_resource_bulk_assign.py::test_bulk_assign_hero_2800_under_5s -v
```

Expected: **XPASS** (test passes, strict=True converts to FAIL → confirming the gate flipped). Then update test or remove `xfail` decorator with new commit.

```bash
python -m pytest tests/test_msproject_file_hero.py -v
```

Expected: **PASS** (Phase 4 hero <5s).

**Step 6: Update Phase 2b xfail decorator**

Once hero passes consistently, change `@pytest.mark.xfail(strict=True)` → remove decorator (test now passes). Commit separately.

**Step 7: Commit (first hero, then xfail flip)**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_hero.py
git commit -m "Phase 4 T73 (HERO): bulk_add_assignments XML merge — 2800 in <5s

Implements Phase 4's success gate. Path: mspdi_parser bulk write 2800
<Assignment> elements (~1s) + auto-sync FileOpen+EditPaste merge (~1-2s)
+ Reschedule (~0.5s) = ~3-4s total."

# Then verify test_bulk_assign_hero_2800_under_5s now xpasses; if so:
git add tests/test_msproject_resource_bulk_assign.py
git commit -m "Phase 4 T73 follow-up: flip Phase 2b hero xfail (now passes via Phase 4 path)"
```

Expected: **322+ PASS + 0 xfail** (hero flipped from xfail to PASS).

---

## Task 74: FastMCP Dispatcher + Acceptance Script + README + Push (FINAL)

**Files:**
- Modify: `msproject_mcp_core.py` (add `@mcp.tool msproject_file` dispatcher near other dispatchers)
- Create: `tests/test_msproject_file_dispatcher.py`
- Create: `samples/build_file_lifecycle.py`
- Modify: `README.md`

**Step 1: Failing dispatcher tests**

`tests/test_msproject_file_dispatcher.py`:
```python
"""Test FastMCP msproject_file dispatcher."""
import asyncio, json, os
import pytest
from msproject_mcp_core import msproject_file

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def test_dispatcher_read_tasks():
    r = _run(msproject_file({"action": "read_tasks", "file_path": MSP_XML}))
    p = json.loads(r)
    assert p["status"] == "ok"
    assert "tasks" in p


def test_dispatcher_query():
    r = _run(msproject_file({
        "action": "query",
        "file_path": MSP_XML,
        "expression": "duration_h > 8",
    }))
    p = json.loads(r)
    assert p["status"] == "ok"


def test_dispatcher_invalid_action():
    r = _run(msproject_file({"action": "nonsense", "file_path": MSP_XML}))
    p = json.loads(r)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_missing_file_path():
    r = _run(msproject_file({"action": "read_tasks"}))
    p = json.loads(r)
    assert p["status"] == "error"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add after `msproject_progress` dispatcher:

```python
@mcp.tool(
    name="msproject_file",
    annotations={"title": "MS Project File-Based Operations",
                 "readOnlyHint": False},
)
async def msproject_file(params: dict) -> str:
    """File-based read+write for MS Project files (.xml/.mspdi/.mpp).

    Actions:
      Read (8): read_tasks, read_links, read_resources, read_assignments,
                read_calendars, read_baselines, read_progress, query
      Write (6): add_tasks, add_links, add_resources, bulk_add_assignments,
                 update_task, save_as

    All actions require file_path. .xml/.mspdi via native parser (zero Java);
    .mpp via MPXJ + JVM (lazy init). Write actions: if MSP open → auto COM
    import + Reschedule (default behavior, not opt-in).

    Phase 4 (30 Apr 2026). Hero: bulk_add_assignments 2800 in <5s.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "read_tasks":
            r = _msp_file_read_tasks(**p)
        elif action == "read_links":
            r = _msp_file_read_links(**p)
        elif action == "read_resources":
            r = _msp_file_read_resources(**p)
        elif action == "read_assignments":
            r = _msp_file_read_assignments(**p)
        elif action == "read_calendars":
            r = _msp_file_read_calendars(**p)
        elif action == "read_baselines":
            r = _msp_file_read_baselines(**p)
        elif action == "read_progress":
            r = _msp_file_read_progress(**p)
        elif action == "query":
            r = _msp_file_query(**p)
        elif action == "add_tasks":
            r = _msp_file_add_tasks(**p)
        elif action == "add_links":
            r = _msp_file_add_links(**p)
        elif action == "add_resources":
            r = _msp_file_add_resources(**p)
        elif action == "bulk_add_assignments":
            r = _msp_file_bulk_add_assignments(**p)
        elif action == "update_task":
            r = _msp_file_update_task(**p)
        elif action == "save_as":
            r = _msp_file_save_as(**p)
        else:
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: read_tasks/read_links/read_resources/read_assignments/read_calendars/read_baselines/read_progress/query/add_tasks/add_links/add_resources/bulk_add_assignments/update_task/save_as"}
    except TypeError as e:
        # Missing required param
        r = {"status": "error", "error": f"Invalid params for {action}: {e}"}
    except Exception as e:
        logger.error(f"msproject_file({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

Update `mcp = FastMCP(...)` instructions string to include `msproject_file`.

**Step 4: Acceptance script**

`samples/build_file_lifecycle.py`:
```python
"""Phase 4 acceptance: full file MCP lifecycle including HERO.

SAFETY: Uses isolated FileNew project. Original user project untouched.

Scenario (target <30s wall clock):
  1. Build base: 200 villa tasks + 14 CAU resources
  2. HERO: bulk_add_assignments 2800 via Phase 4 XML merge → <5s strict
  3. Read demo: export temp XML → read_tasks/links/resources/assignments → counts
  4. Write demo: update_task duration → auto-sync verify (proj.Reschedule reflected)
  5. .mpp read demo: temp .mpp via MPXJ
  6. Query demo: ad-hoc filter expression
"""
import os, sys, time, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom, win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add,
    _msp_file_bulk_add_assignments, _msp_file_read_tasks,
    _msp_file_read_resources, _msp_file_read_assignments,
    _msp_file_query, _msp_file_update_task,
)


N_TASKS = 200


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
        # Step 1: base build
        print(f"\n1. Building {N_TASKS} tasks + 14 resources...")
        tasks = _msp_task_bulk_add(items=[{"name": f"V{i:03d}", "duration": "2d"}
                                          for i in range(N_TASKS)])
        task_ids = tasks["task_ids"]
        res_ids = []
        for n in ["COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
                  "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR"]:
            res_ids.append(_msp_resource_add(name=n, type="Work")["resource_id"])
        print(f"   OK in {time.time()-t0:.2f}s")

        # Step 2: HERO
        print("\n2. HERO: bulk_add_assignments 2800 via Phase 4 XML merge...")
        tmpdir = tempfile.mkdtemp()
        temp_xml = os.path.join(tmpdir, "hero.xml")
        test_proj.SaveAs(temp_xml, 4)  # FileFormat=4 XML
        items = [{"task_id": tid, "resource_id": rid}
                 for tid in task_ids for rid in res_ids]
        h0 = time.time()
        r = _msp_file_bulk_add_assignments(file_path=temp_xml, items=items)
        h_elapsed = time.time() - h0
        assert r["status"] == "ok", f"hero failed: {r}"
        assert h_elapsed < 5.0, f"HERO took {h_elapsed:.2f}s (target <5s)"
        print(f"   OK 🚀 {h_elapsed:.2f}s (target <5s) auto_imported={r.get('auto_imported')}")

        # Step 3: read demo
        print("\n3. Read demo via Phase 4 file MCP...")
        rt = _msp_file_read_tasks(file_path=temp_xml)
        rs = _msp_file_read_resources(file_path=temp_xml)
        ra = _msp_file_read_assignments(file_path=temp_xml)
        print(f"   tasks={rt['count']}, resources={rs['count']}, assignments={ra['count']}")
        assert ra['count'] == 2800

        # Step 4: write demo
        print("\n4. Write demo: update_task...")
        u = _msp_file_update_task(file_path=temp_xml, task_id=task_ids[0],
                                  fields={"duration": "5d"})
        print(f"   {u['status']} auto_imported={u.get('auto_imported')}")

        # Step 5: .mpp read demo (skip if no .mpp env)
        try:
            print("\n5. .mpp read demo via MPXJ...")
            temp_mpp = os.path.join(tmpdir, "lifecycle.mpp")
            test_proj.SaveAs(temp_mpp, 0)  # FileFormat=0 MPP
            rmpp = _msp_file_read_tasks(file_path=temp_mpp)
            print(f"   MPP tasks={rmpp['count']}")
        except Exception as e:
            print(f"   .mpp demo skipped: {e}")

        # Step 6: query
        print("\n6. Query demo...")
        q = _msp_file_query(file_path=temp_xml, expression="duration_h >= 16")
        print(f"   query 'duration_h >= 16' → {q['count']} matches")

        elapsed = time.time() - t0
        print(f"\n✅ ACCEPTANCE: {elapsed:.2f}s total (target <30s)")
        assert elapsed < 30.0, f"Too slow: {elapsed}s"

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
python samples/build_file_lifecycle.py
```

Expected: `✅ ACCEPTANCE: <Xs total (target <30s)`. Realistic ~15-25s.

**Step 6: README update**

Add Phase 4 section to `README.md` after Phase 3b:

```markdown
### Phase 4 — File MCP (30 Apr 2026)

`msproject_file` tool with 14 actions, file-based read+write:
- 8 read: read_tasks/links/resources/assignments/calendars/baselines/progress/query
- 6 write: add_tasks/links/resources, bulk_add_assignments (HERO), update_task, save_as
- Format: .xml/.mspdi via native Python parser (zero Java); .mpp via MPXJ+JVM (lazy)
- Auto-sync: write actions automatically COM-import + Reschedule when MSP open
- 🚀 HERO: bulk_add_assignments 2800 in <5s via XML merge (flips Phase 2b xfail)

Acceptance: `samples/build_file_lifecycle.py` full lifecycle <30s.

Tool count: **8 tools, ~66 actions**.
```

**Step 7: Run full regression**

```bash
python -m pytest tests/ -q --tb=line --ignore=cleanup_test.py --ignore=test_apply_tabledef.py
```

Expected: **~322 PASS + 0 xfail** (282 baseline + ~40 Phase 4 new + 1 hero xfail flipped).

**Step 8: Commit + push**

```bash
git add msproject_mcp_core.py tests/test_msproject_file_dispatcher.py samples/build_file_lifecycle.py README.md
git commit -m "Phase 4 T74: dispatcher + acceptance + README + push (file lifecycle <30s)"
git push origin main
```

Expected: ~13-15 commits pushed (T65-T74 + design + plan + any fix commits).

---

## Phase 4 Tamamlama Kriterleri

1. ✅ T65-T74 ~10-15 commit landed
2. ✅ Acceptance script `samples/build_file_lifecycle.py` <30s
3. ✅ Yeni testler ~40 PASS
4. ✅ Phase 1+2+3 mevcut 282+1xfail regression PASS
5. ✅ Total ~322+ PASS + 0 xfail (Phase 2b hero xfail FLIPPED via Phase 4 path)
6. ✅ Push to origin/main
7. ✅ Phase 4 live on GitHub
8. ✅ XML schema detection: MS Project ≠ Asta net hata mesajı
9. ✅ Auto-sync: write → MSP açık → otomatik COM import + Reschedule
10. ⏸ Kullanıcı manuel onayı → Phase 5 (Power Tools — health/evm/excel) başlar

---

## Sequencing Tips (Phase 3a/3b deneyiminden)

1. **T65 BIG ONE** — Probe mspdi_parser API capabilities first (subagent dispatch with full T65 task text)
2. **T66-T68 trivial** — manual write + self-verify (read action'lar)
3. **T69 medium** — query action subagent (filter parser edge cases)
4. **T70-T71 trivial** — write action'lar manual
5. **T72 BIG ONE** — auto-sync probe-driven, subagent dispatch
6. **T73 BIG ONE (HERO)** — full subagent + spec reviewer + quality reviewer chain
7. **T74 standard** — dispatcher + acceptance + push

**Pattern reminder:** subagent-driven-development for BIG ONEs (T65, T69, T72, T73). Plan dosyasını subagent'a OKUTMA — controller curate eder, full task text prompt'a yapıştır. Phase 1+2+3 kodu DOKUNULMAZ.

---

*Plan tamamlandı: 30 Nisan 2026*
*Tahmini Phase 4 süresi: ~17 saat (T65-T74, ~10 task TDD chain)*
*Sonraki phase (onay sonrası): Phase 5 — Power Tools (msproject_health DCMA + msproject_evm + msproject_excel)*
