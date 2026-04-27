# MS Project MCP — Phase 1 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Foundation + Task/Schedule core + MSPDI bulk engine. Hedef: boş Project1'e 5 task COM, 15 task batch, 200 task MSPDI bulk olarak yazılabilsin; tüm Phase 1 testleri geçsin.

**Architecture:** `msproject_mcp_core.py` (FastMCP server) + `msproject_bulk.py` (MSPDI XML bulk-write engine) + `mspdi_parser.py` (mevcut, MS Project XML için doğrulama gerekli). Hibrit speed routing: ≤5 → COM doğrudan, 6-19 → COM batch, 20+ → MSPDI bulk.

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 (COM), pydantic, pytest, openpyxl (Phase 5'e kadar gerek yok).

**Acceptance Criteria (Phase 1 sonu):**
- 3 tool çalışıyor: `msproject_task`, `msproject_link`, `msproject_schedule`
- 200 task'lık MSPDI bulk import <5 saniye
- Unit + Integration testleri full PASS
- Boş Project1'e villa örneği yazılır, manuel doğrulanır
- Bilinen bug 0 — kullanıcı onaylar
- Commit + push GitHub'a

---

## Task 1: Phase 1 Klasör Yapısı

**Files:**
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `samples/__init__.py`

**Step 1: Klasör + `__init__.py` oluştur**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
mkdir -p tests samples
```

`tests/__init__.py`:
```python
"""MS Project MCP test suite."""
```

`tests/conftest.py`:
```python
"""Pytest fixtures shared across MS Project tests."""
import pytest
import os

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture(scope="session")
def fixtures_dir():
    return os.path.join(REPO_ROOT, "tests", "fixtures")


@pytest.fixture(scope="session")
def msproject_app():
    """Session-level MS Project COM connection. Skips if not available."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.GetActiveObject("MSProject.Application")
        if app.ActiveProject is None:
            pytest.skip("No active MS Project")
        return app
    except Exception as e:
        pytest.skip(f"MS Project not available: {e}")
```

`samples/__init__.py`:
```python
"""Sample MS Project build scripts."""
```

**Step 2: Commit**

```bash
git add tests/ samples/
git commit -m "Phase 1 T1: scaffold tests/ and samples/ directories"
```

---

## Task 2: MS Project COM Type Library Dump

**Files:**
- Create: `msproject_typelib.txt` (COM API reference)
- Create: `tools/dump_msproject_typelib.py` (helper script)

**Step 1: Yardımcı script oluştur**

`tools/dump_msproject_typelib.py`:
```python
"""Dump MS Project COM type library to text reference.
One-time generation. Output: msproject_typelib.txt
"""
import pythoncom
import win32com.client
from win32com.client import gencache, makepy

OUTPUT = "msproject_typelib.txt"

def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    print(f"MS Project version: {app.Version} build {app.Build}")
    # Force gencache for type info
    gencache.EnsureDispatch("MSProject.Application")
    # Use makepy.Walker pattern
    import sys, io
    buf = io.StringIO()
    sys.stdout = buf
    try:
        makepy.main(["-o", "-", "MSProject.Application"])
    finally:
        sys.stdout = sys.__stdout__
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(buf.getvalue())
    print(f"Wrote {OUTPUT} ({len(buf.getvalue())} chars)")


if __name__ == "__main__":
    main()
```

**Step 2: Çalıştır**

```bash
mkdir -p tools
python tools/dump_msproject_typelib.py
```

Expected output: `Wrote msproject_typelib.txt (~50000+ chars)`

**Step 3: Doğrula**

```bash
head -30 msproject_typelib.txt
grep -i "class Tasks(" msproject_typelib.txt | head -5
grep -i "def Add" msproject_typelib.txt | head -10
```

Expected: COM class definitions for Application, Project, Tasks, Resources görünmeli.

**Step 4: Commit**

```bash
git add tools/dump_msproject_typelib.py msproject_typelib.txt
git commit -m "Phase 1 T2: dump MS Project COM type library reference"
```

---

## Task 3: MSPDI Parser MS Project Doğrulaması

**Files:**
- Create: `tests/test_mspdi_msproject_compat.py`
- Modify: `mspdi_parser.py` (gerekirse)

**Step 1: Failing test yaz — MSP XML'i parse edebilmeli**

Önce manuel olarak MS Project'te boş bir proje oluştur, "Save As → XML" yap, kaydet: `tests/fixtures/empty_msp.xml`. Test bu fixture'ı kullanacak.

`tests/test_mspdi_msproject_compat.py`:
```python
"""Verify mspdi_parser handles MS Project XML (not just Asta exports)."""
import os
import pytest
from mspdi_parser import MspdiProject


@pytest.fixture
def empty_msp_xml(fixtures_dir):
    path = os.path.join(fixtures_dir, "empty_msp.xml")
    if not os.path.exists(path):
        pytest.skip(f"Fixture missing: {path}")
    return path


def test_empty_msp_xml_loads(empty_msp_xml):
    """Boş MS Project export'u parse edilebilmeli."""
    proj = MspdiProject(empty_msp_xml)
    assert proj is not None
    summary = proj.get_project_summary()
    assert "project_name" in summary
    assert summary["total_tasks"] >= 1  # MSP root task ekler


def test_msp_xml_round_trip(empty_msp_xml, tmp_path):
    """Read → modify → save → re-read symmetry."""
    proj = MspdiProject(empty_msp_xml)
    new_task = proj.add_task(name="Test Task", duration_str="3d")
    assert "id" in new_task
    output = tmp_path / "modified.xml"
    proj.save(output_path=str(output))
    proj2 = MspdiProject(str(output))
    tasks = proj2.get_all_tasks(include_summary=False)
    names = [t["name"] for t in tasks]
    assert "Test Task" in names
```

**Step 2: Fixture oluştur — manuel adım**

Kullanıcıdan MS Project'te boş projeyi `tests/fixtures/empty_msp.xml` olarak Save As yapmasını iste.

```bash
mkdir -p tests/fixtures
# User: MS Project → File → Save As → XML Format → tests/fixtures/empty_msp.xml
```

**Step 3: Test'i çalıştır — ilk fail**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/test_mspdi_msproject_compat.py -v
```

Expected: PASS — MS Project XML schema = MSPDI = mspdi_parser ile uyumlu olmalı (zaten Microsoft schema). Eğer FAIL ederse mspdi_parser'da MSP-specific bir field için patch atılır.

**Step 4: Eğer FAIL ediyorsa düzelt**

`mspdi_parser.py`'de MS Project namespace farkları varsa fix:
- MSP XML'de `xmlns="http://schemas.microsoft.com/project"` standart
- `<UID>` field her zaman var
- `<Calendar>` UID 1 = Standard

Patch'leri test geçene kadar uygula.

**Step 5: Commit**

```bash
git add tests/test_mspdi_msproject_compat.py tests/fixtures/empty_msp.xml mspdi_parser.py
git commit -m "Phase 1 T3: verify mspdi_parser handles MS Project XML"
```

---

## Task 4: `msproject_bulk.py` — MSPDI Bulk Engine İskelet

**Files:**
- Create: `msproject_bulk.py`
- Create: `tests/test_msproject_bulk.py`

**Step 1: Failing test yaz**

`tests/test_msproject_bulk.py`:
```python
"""Test MSPDI bulk-write engine for MS Project."""
import pytest
from msproject_bulk import MsprojectBulkWriter


def test_writer_creates_empty_project(tmp_path):
    """Empty bulk writer should produce a valid MSPDI XML."""
    w = MsprojectBulkWriter(project_name="Test Bulk")
    out = tmp_path / "test_bulk.xml"
    w.save(str(out))
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "<?xml" in content
    assert "Test Bulk" in content


def test_writer_adds_tasks(tmp_path):
    """Bulk add 50 tasks → all present in output."""
    w = MsprojectBulkWriter(project_name="Bulk 50")
    items = [{"name": f"Task {i}", "duration": "1d"} for i in range(50)]
    w.bulk_add_tasks(items)
    out = tmp_path / "bulk50.xml"
    w.save(str(out))
    # Re-read with mspdi_parser to verify
    from mspdi_parser import MspdiProject
    p = MspdiProject(str(out))
    tasks = p.get_all_tasks(include_summary=False)
    assert len(tasks) == 50
```

**Step 2: Test'i çalıştır — fail bekle**

```bash
python -m pytest tests/test_msproject_bulk.py -v
```

Expected: FAIL with `ModuleNotFoundError: No module named 'msproject_bulk'`

**Step 3: Minimal implementation**

`msproject_bulk.py`:
```python
"""MSPDI XML bulk-write engine for MS Project.

Path 3 of hybrid speed strategy: bulk operations (>20 items) write to
MSPDI XML and trigger MS Project FileOpen import — much faster than COM
one-by-one (200 task in ~3-5 sec vs 60+ sec).
"""
from __future__ import annotations
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import xml.etree.ElementTree as ET


MSPDI_NS = "http://schemas.microsoft.com/project"
ET.register_namespace("", MSPDI_NS)


class MsprojectBulkWriter:
    """Generates MSPDI XML compatible with MS Project FileOpen import."""

    def __init__(self, project_name: str = "Bulk Project",
                 start_date: Optional[str] = None):
        self.project_name = project_name
        self.start_date = start_date or datetime.now().strftime("%Y-%m-%dT08:00:00")
        self.tasks: List[Dict[str, Any]] = []
        self.links: List[Dict[str, Any]] = []
        self.resources: List[Dict[str, Any]] = []
        self.assignments: List[Dict[str, Any]] = []
        self._next_uid = 1
        self._next_id = 0

    def bulk_add_tasks(self, items: List[Dict[str, Any]]) -> List[int]:
        """Add a batch of tasks. Returns list of assigned UIDs.

        Each item: {name, duration (str like "5d"), [start, finish, summary, milestone, parent_uid]}
        """
        uids = []
        for item in items:
            uid = self._next_uid
            self._next_uid += 1
            self._next_id += 1
            t = {
                "UID": uid,
                "ID": self._next_id,
                "Name": item.get("name", f"Task {self._next_id}"),
                "Duration": self._duration_to_iso(item.get("duration", "1d")),
                "Summary": "1" if item.get("summary") else "0",
                "Milestone": "1" if item.get("milestone") else "0",
                "OutlineLevel": item.get("outline_level", 1),
            }
            if item.get("start"):
                t["Start"] = item["start"]
            if item.get("finish"):
                t["Finish"] = item["finish"]
            self.tasks.append(t)
            uids.append(uid)
        return uids

    def bulk_add_links(self, items: List[Dict[str, Any]]) -> int:
        """Add predecessor links. Each: {pred_uid, succ_uid, type='FS', lag='0d'}."""
        count = 0
        for item in items:
            self.links.append({
                "succ_uid": item["succ_uid"],
                "pred_uid": item["pred_uid"],
                "type": item.get("type", "FS"),
                "lag": item.get("lag", "0d"),
            })
            count += 1
        return count

    def bulk_add_resources(self, items: List[Dict[str, Any]]) -> List[int]:
        """Add resources. Each: {name, type='Work'}."""
        uids = []
        for item in items:
            uid = self._next_uid
            self._next_uid += 1
            self.resources.append({
                "UID": uid,
                "Name": item["name"],
                "Type": item.get("type", "Work"),
            })
            uids.append(uid)
        return uids

    def save(self, output_path: str) -> str:
        """Write MSPDI XML to file."""
        xml = self._build_xml()
        with open(output_path, "w", encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
            f.write(xml)
        return output_path

    def _build_xml(self) -> str:
        """Build MSPDI XML string."""
        ns = MSPDI_NS
        root = ET.Element(f"{{{ns}}}Project")
        ET.SubElement(root, f"{{{ns}}}Name").text = self.project_name
        ET.SubElement(root, f"{{{ns}}}Title").text = self.project_name
        ET.SubElement(root, f"{{{ns}}}StartDate").text = self.start_date
        ET.SubElement(root, f"{{{ns}}}CalendarUID").text = "1"

        # Calendars
        cals = ET.SubElement(root, f"{{{ns}}}Calendars")
        cal = ET.SubElement(cals, f"{{{ns}}}Calendar")
        ET.SubElement(cal, f"{{{ns}}}UID").text = "1"
        ET.SubElement(cal, f"{{{ns}}}Name").text = "Standard"
        ET.SubElement(cal, f"{{{ns}}}IsBaseCalendar").text = "1"

        # Tasks
        tasks_el = ET.SubElement(root, f"{{{ns}}}Tasks")
        for t in self.tasks:
            tk = ET.SubElement(tasks_el, f"{{{ns}}}Task")
            for k, v in t.items():
                if k == "links":
                    continue
                ET.SubElement(tk, f"{{{ns}}}{k}").text = str(v)
            # Predecessors for this task
            for ln in self.links:
                if ln["succ_uid"] == t["UID"]:
                    pl = ET.SubElement(tk, f"{{{ns}}}PredecessorLink")
                    ET.SubElement(pl, f"{{{ns}}}PredecessorUID").text = str(ln["pred_uid"])
                    ET.SubElement(pl, f"{{{ns}}}Type").text = self._link_type(ln["type"])
                    ET.SubElement(pl, f"{{{ns}}}LinkLag").text = self._lag_to_minutes(ln["lag"])

        # Resources
        res_el = ET.SubElement(root, f"{{{ns}}}Resources")
        for r in self.resources:
            rs = ET.SubElement(res_el, f"{{{ns}}}Resource")
            for k, v in r.items():
                ET.SubElement(rs, f"{{{ns}}}{k}").text = str(v)

        # Assignments
        as_el = ET.SubElement(root, f"{{{ns}}}Assignments")

        # Pretty print
        ET.indent(root, space="  ")
        return ET.tostring(root, encoding="unicode")

    @staticmethod
    def _duration_to_iso(d: str) -> str:
        """Convert '5d' → 'PT40H0M0S' (assuming 8h/day)."""
        n = float(d.rstrip("dwhmDWHM") or "1")
        unit = (d[-1] if d and d[-1].isalpha() else "d").lower()
        hours = {"d": 8, "w": 40, "h": 1, "m": 1/60}.get(unit, 8) * n
        return f"PT{int(hours)}H0M0S"

    @staticmethod
    def _lag_to_minutes(lag: str) -> str:
        """Convert '2d' → '960' minutes."""
        n = float(lag.rstrip("dwhmDWHM") or "0")
        unit = (lag[-1] if lag and lag[-1].isalpha() else "d").lower()
        mins = {"d": 480, "w": 2400, "h": 60, "m": 1}.get(unit, 480) * n
        return str(int(mins))

    @staticmethod
    def _link_type(t: str) -> str:
        """FF=0, FS=1, SF=2, SS=3 per MSPDI spec."""
        return {"FF": "0", "FS": "1", "SF": "2", "SS": "3"}.get(t.upper(), "1")
```

**Step 4: Test'i çalıştır — geçmesi lazım**

```bash
python -m pytest tests/test_msproject_bulk.py -v
```

Expected: 2 PASSED.

**Step 5: Commit**

```bash
git add msproject_bulk.py tests/test_msproject_bulk.py
git commit -m "Phase 1 T4: MSPDI bulk-write engine skeleton + tasks/links"
```

---

## Task 5: `msproject_mcp_core.py` — Server İskelet + COM Cache

**Files:**
- Create: `msproject_mcp_core.py`
- Create: `tests/test_msproject_core_basic.py`

**Step 1: Failing test yaz**

`tests/test_msproject_core_basic.py`:
```python
"""Test core MS Project MCP infrastructure."""
import pytest
from msproject_mcp_core import _connect_app, _route_operation


def test_connect_app(msproject_app):
    """COM connection helper should return active app."""
    app = _connect_app()
    assert app is not None
    assert app.ActiveProject is not None


def test_route_operation_thresholds():
    """_route_operation should pick correct path per item count."""
    assert _route_operation(1) == "com_direct"
    assert _route_operation(5) == "com_direct"
    assert _route_operation(6) == "com_batch"
    assert _route_operation(19) == "com_batch"
    assert _route_operation(20) == "mspdi_bulk"
    assert _route_operation(500) == "mspdi_bulk"
```

**Step 2: Çalıştır — fail bekle**

```bash
python -m pytest tests/test_msproject_core_basic.py -v
```

Expected: FAIL with import error

**Step 3: Minimal core implementation**

`msproject_mcp_core.py`:
```python
"""MS Project MCP Server — COM-based.

Hybrid speed strategy:
  1-5 items   → COM direct (real-time, instant UI feedback)
  6-19 items  → COM batch (Calculation manual + ScreenUpdating off)
  20+ items   → MSPDI XML bulk import (~3-5s for 200 tasks)

Phase 1 tools: msproject_task, msproject_link, msproject_schedule.
"""
from __future__ import annotations
import atexit
import logging
import os
import threading
import time
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP

# ---------- LOGGING ----------
log_dir = os.path.expanduser("~/.claude/logs")
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(log_dir, "msproject_mcp.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("msproject_mcp")

# ---------- MCP SERVER ----------
mcp = FastMCP(
    "msproject_mcp",
    instructions=(
        "MS Project COM-based MCP server. Connects to running MS Project (Application='MSProject.Application'). "
        "Hybrid speed: 1-5 items COM direct, 6-19 batch, 20+ MSPDI bulk import. "
        "Tools: msproject_task, msproject_link, msproject_schedule (Phase 1)."
    ),
)

# ---------- COM CONNECTION CACHE ----------
_app_lock = threading.RLock()
_app: Optional[Any] = None
_calc_modified = False  # track if we changed calc mode (need restore on exit)
_screenupdating_modified = False


def _connect_app() -> Any:
    """Connect to running MS Project. Cached singleton."""
    global _app
    with _app_lock:
        if _app is not None:
            try:
                _ = _app.Version  # ping
                return _app
            except Exception:
                _app = None  # invalidate
        pythoncom.CoInitialize()
        try:
            _app = win32com.client.GetActiveObject("MSProject.Application")
            logger.info(f"Connected to MS Project {_app.Version}")
        except Exception as e:
            raise RuntimeError(
                f"MS Project'e bağlanılamadı: {e}. "
                "(1) MS Project açık olduğundan emin olun. "
                "(2) Bir proje açık olmalı (boş Project bile yeterli). "
                "(3) Hala olmuyorsa MS Project'i yeniden başlatın."
            )
        return _app


def _validate_active_project() -> Any:
    """Validates ActiveProject is present."""
    app = _connect_app()
    if app.ActiveProject is None:
        raise RuntimeError("MS Project'te aktif proje yok. Boş bir proje açın veya File → New.")
    return app


def _route_operation(op_count: int) -> str:
    """Pick speed path based on operation count."""
    if op_count <= 5:
        return "com_direct"
    if op_count <= 19:
        return "com_batch"
    return "mspdi_bulk"


def _enter_batch_mode():
    """Enter COM batch mode: disable screen update, manual calc, no events."""
    global _calc_modified, _screenupdating_modified
    app = _connect_app()
    pj_manual = 0  # PjCalculation.pjManual
    if app.Calculation != pj_manual:
        app.Calculation = pj_manual
        _calc_modified = True
    if app.ScreenUpdating:
        app.ScreenUpdating = False
        _screenupdating_modified = True
    proj = app.ActiveProject
    try:
        proj.EventsEnabled = False
    except Exception:
        pass


def _exit_batch_mode():
    """Restore screen update + auto calc + events."""
    global _calc_modified, _screenupdating_modified
    try:
        app = _connect_app()
        if _calc_modified:
            pj_auto = 1  # PjCalculation.pjAutomatic
            app.Calculation = pj_auto
            _calc_modified = False
        if _screenupdating_modified:
            app.ScreenUpdating = True
            _screenupdating_modified = False
        proj = app.ActiveProject
        if proj:
            try:
                proj.EventsEnabled = True
            except Exception:
                pass
    except Exception as e:
        logger.warning(f"_exit_batch_mode error (non-fatal): {e}")


@atexit.register
def _restore_on_exit():
    """Critical: ensure MS Project never left in manual/screen-off state."""
    _exit_batch_mode()


# ---------- TOOL DISPATCHERS (will be filled in subsequent tasks) ----------

# Placeholder — actual tools added in T6+


def main():
    """Run MCP server (stdio)."""
    mcp.run()


if __name__ == "__main__":
    main()
```

**Step 4: Test'i çalıştır**

```bash
python -m pytest tests/test_msproject_core_basic.py -v
```

Expected: 2 PASSED (test_connect_app skip if MS Project closed, test_route_operation_thresholds always passes).

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_core_basic.py
git commit -m "Phase 1 T5: msproject_mcp_core scaffold with COM cache + hybrid routing"
```

---

## Task 6: `msproject_task` — `add` Action (Path 1 / COM Direct)

**Files:**
- Modify: `msproject_mcp_core.py` (add helpers + tool function)
- Create: `tests/test_msproject_task_add.py`

**Step 1: Failing test yaz**

`tests/test_msproject_task_add.py`:
```python
"""Test msproject_task add action."""
import pytest
import asyncio
from msproject_mcp_core import _msp_task_add_single


def test_add_single_task(msproject_app):
    """Adding 1 task → ActiveProject.Tasks.Count incremented."""
    proj = msproject_app.ActiveProject
    initial = proj.Tasks.Count
    result = _msp_task_add_single(name="Test Task A", duration="3d")
    assert result["status"] == "ok"
    assert proj.Tasks.Count == initial + 1
    # Find by name
    found = None
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t and t.Name == "Test Task A":
            found = t
            break
    assert found is not None
    # Cleanup
    found.Delete()


def test_add_milestone(msproject_app):
    """Milestone has 0d duration."""
    proj = msproject_app.ActiveProject
    result = _msp_task_add_single(name="MS Test", duration="0d", milestone=True)
    assert result["status"] == "ok"
    found = None
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t and t.Name == "MS Test":
            found = t
            break
    assert found is not None and found.Milestone
    found.Delete()
```

**Step 2: Çalıştır — fail bekle**

Expected: FAIL — `_msp_task_add_single` undefined.

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
# ---------- TASK HELPERS ----------

def _parse_duration(d: str) -> int:
    """Convert '5d' → minutes (assuming 8h/day for MSP)."""
    if not d:
        return 480
    s = d.strip()
    # extract trailing letter
    unit = s[-1].lower() if s[-1].isalpha() else "d"
    try:
        n = float(s.rstrip("dwhmDWHM"))
    except ValueError:
        n = 1.0
    return {"d": 480, "w": 2400, "h": 60, "m": 1}.get(unit, 480) * int(n)


def _msp_task_add_single(name: str, duration: str = "1d",
                         start: Optional[str] = None,
                         finish: Optional[str] = None,
                         summary: bool = False,
                         milestone: bool = False,
                         notes: Optional[str] = None) -> Dict[str, Any]:
    """Add a single task via COM (Path 1)."""
    app = _validate_active_project()
    proj = app.ActiveProject
    try:
        # Tasks.Add(name, before) — before=None appends at end
        new_task = proj.Tasks.Add(name)
        if milestone:
            new_task.Milestone = True
            new_task.Duration = 0
        elif duration:
            new_task.Duration = _parse_duration(duration)
        if start:
            new_task.Start = start
        if finish:
            new_task.Finish = finish
        if summary:
            new_task.Summary = True  # actually summary requires children; flag set
        if notes:
            new_task.Notes = notes
        return {
            "status": "ok",
            "task_id": new_task.ID,
            "task_uid": new_task.UniqueID,
            "name": new_task.Name,
            "duration": duration,
            "milestone": bool(milestone),
        }
    except Exception as e:
        logger.error(f"_msp_task_add_single failed: {e}")
        return {"status": "error", "error": str(e)}


# Will be wired into msproject_task dispatcher in T15
```

**Step 4: Test'i çalıştır**

```bash
python -m pytest tests/test_msproject_task_add.py -v
```

Expected: 2 PASSED (skip if no MS Project).

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_task_add.py
git commit -m "Phase 1 T6: msproject_task add (single, COM direct path)"
```

---

## Task 7: `msproject_task` — `update`, `delete`, `get`, `list` Actions

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_task_crud.py`

**Step 1: Failing test yaz**

`tests/test_msproject_task_crud.py`:
```python
"""Test msproject_task CRUD operations."""
import pytest
from msproject_mcp_core import (
    _msp_task_add_single, _msp_task_update, _msp_task_delete,
    _msp_task_get, _msp_task_list,
)


@pytest.fixture
def temp_task(msproject_app):
    """Create + cleanup a temp task."""
    result = _msp_task_add_single(name="Temp", duration="2d")
    task_id = result["task_id"]
    yield task_id
    try:
        _msp_task_delete(task_id=task_id)
    except Exception:
        pass


def test_get_task(temp_task):
    r = _msp_task_get(task_id=temp_task)
    assert r["status"] == "ok"
    assert r["task"]["name"] == "Temp"


def test_update_task(temp_task):
    r = _msp_task_update(task_id=temp_task, name="Renamed", duration="5d", notes="changed")
    assert r["status"] == "ok"
    g = _msp_task_get(task_id=temp_task)
    assert g["task"]["name"] == "Renamed"
    assert g["task"]["notes"] == "changed"


def test_delete_task(msproject_app):
    r = _msp_task_add_single(name="ToDelete", duration="1d")
    tid = r["task_id"]
    proj = msproject_app.ActiveProject
    before = proj.Tasks.Count
    d = _msp_task_delete(task_id=tid)
    assert d["status"] == "ok"
    assert proj.Tasks.Count == before - 1


def test_list_tasks(temp_task):
    r = _msp_task_list(limit=200)
    assert r["status"] == "ok"
    names = [t["name"] for t in r["tasks"]]
    assert "Temp" in names
```

**Step 2: Çalıştır — fail bekle**

Expected: FAIL — functions undefined.

**Step 3: Implementation**

Add to `msproject_mcp_core.py`:

```python
def _find_task_by_id(proj: Any, task_id: int) -> Optional[Any]:
    """Locate a task object by its ID. Returns None if not found."""
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t is not None and t.ID == task_id:
            return t
    return None


def _msp_task_update(task_id: int, name: Optional[str] = None,
                     duration: Optional[str] = None,
                     start: Optional[str] = None, finish: Optional[str] = None,
                     percent_complete: Optional[float] = None,
                     notes: Optional[str] = None) -> Dict[str, Any]:
    app = _validate_active_project()
    t = _find_task_by_id(app.ActiveProject, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    changes = []
    try:
        if name is not None:
            t.Name = name; changes.append("name")
        if duration is not None:
            t.Duration = _parse_duration(duration); changes.append("duration")
        if start is not None:
            t.Start = start; changes.append("start")
        if finish is not None:
            t.Finish = finish; changes.append("finish")
        if percent_complete is not None:
            t.PercentComplete = percent_complete; changes.append("percent_complete")
        if notes is not None:
            t.Notes = notes; changes.append("notes")
        return {"status": "ok", "task_id": task_id, "changes": changes}
    except Exception as e:
        logger.error(f"_msp_task_update failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_task_delete(task_id: int) -> Dict[str, Any]:
    app = _validate_active_project()
    t = _find_task_by_id(app.ActiveProject, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    try:
        name = t.Name
        t.Delete()
        return {"status": "ok", "deleted_id": task_id, "deleted_name": name}
    except Exception as e:
        return {"status": "error", "error": str(e)}


def _serialize_task(t: Any) -> Dict[str, Any]:
    return {
        "id": t.ID,
        "uid": t.UniqueID,
        "name": t.Name,
        "duration": t.Duration,  # minutes
        "start": str(t.Start) if t.Start else None,
        "finish": str(t.Finish) if t.Finish else None,
        "percent_complete": t.PercentComplete,
        "milestone": bool(t.Milestone),
        "summary": bool(t.Summary),
        "outline_level": t.OutlineLevel,
        "notes": t.Notes or "",
    }


def _msp_task_get(task_id: int) -> Dict[str, Any]:
    app = _validate_active_project()
    t = _find_task_by_id(app.ActiveProject, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    return {"status": "ok", "task": _serialize_task(t)}


def _msp_task_list(include_summary: bool = True, limit: int = 100) -> Dict[str, Any]:
    app = _validate_active_project()
    proj = app.ActiveProject
    out = []
    for i in range(1, min(proj.Tasks.Count, limit) + 1):
        t = proj.Tasks(i)
        if t is None:
            continue
        if not include_summary and t.Summary:
            continue
        out.append(_serialize_task(t))
    return {"status": "ok", "total": proj.Tasks.Count, "returned": len(out), "tasks": out}
```

**Step 4: Test çalıştır**

```bash
python -m pytest tests/test_msproject_task_crud.py -v
```

Expected: 4 PASSED.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_task_crud.py
git commit -m "Phase 1 T7: msproject_task update/delete/get/list (COM direct)"
```

---

## Task 8: `msproject_task` — `add_summary`, `add_milestone` Actions

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_task_special.py`

**Step 1: Test yaz, çalıştır (FAIL), implement, çalıştır (PASS), commit**

`tests/test_msproject_task_special.py`:
```python
import pytest
from msproject_mcp_core import _msp_task_add_summary, _msp_task_add_milestone, _msp_task_delete


def test_add_summary(msproject_app):
    r = _msp_task_add_summary(name="Phase 1", duration="10d")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=r["task_id"])


def test_add_milestone(msproject_app):
    r = _msp_task_add_milestone(name="Project Start MS", date="2026-04-26")
    assert r["status"] == "ok"
    _msp_task_delete(task_id=r["task_id"])
```

Implementation in `msproject_mcp_core.py`:

```python
def _msp_task_add_summary(name: str, duration: str = "1d",
                          parent_task_id: Optional[int] = None) -> Dict[str, Any]:
    """Add a summary task. In MS Project summary is implicit (becomes summary when has children)."""
    return _msp_task_add_single(name=name, duration=duration, summary=True)


def _msp_task_add_milestone(name: str, date: Optional[str] = None) -> Dict[str, Any]:
    """Add a 0-duration milestone."""
    r = _msp_task_add_single(name=name, duration="0d", milestone=True)
    if r.get("status") == "ok" and date:
        _msp_task_update(task_id=r["task_id"], start=date, finish=date)
    return r
```

```bash
python -m pytest tests/test_msproject_task_special.py -v
git add msproject_mcp_core.py tests/test_msproject_task_special.py
git commit -m "Phase 1 T8: msproject_task add_summary + add_milestone"
```

---

## Task 9: `msproject_task` — `bulk_add` Action with Hybrid Routing

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_task_bulk.py`

**Step 1: Failing tests**

`tests/test_msproject_task_bulk.py`:
```python
import pytest
import time
from msproject_mcp_core import _msp_task_bulk_add, _msp_task_list


def _cleanup_all_tasks(proj):
    while proj.Tasks.Count > 0:
        proj.Tasks(1).Delete()


def test_bulk_3_tasks_com_direct(msproject_app):
    """3 items → COM direct path."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"T{i}", "duration": "1d"} for i in range(3)]
    r = _msp_task_bulk_add(items=items)
    assert r["status"] == "ok"
    assert r["path"] == "com_direct"
    assert r["count"] == 3
    _cleanup_all_tasks(msproject_app.ActiveProject)


def test_bulk_15_tasks_com_batch(msproject_app):
    """15 items → COM batch path."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"B{i}", "duration": "2d"} for i in range(15)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "com_batch"
    assert r["count"] == 15
    assert elapsed < 10, f"Too slow: {elapsed}s"
    _cleanup_all_tasks(msproject_app.ActiveProject)


def test_bulk_30_tasks_mspdi(msproject_app):
    """30 items → MSPDI bulk path."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"M{i}", "duration": "1d"} for i in range(30)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["path"] == "mspdi_bulk"
    assert r["count"] == 30
    assert elapsed < 15, f"Too slow: {elapsed}s"
    _cleanup_all_tasks(msproject_app.ActiveProject)
```

**Step 2: Çalıştır — fail**

**Step 3: Implementation in `msproject_mcp_core.py`**

```python
def _msp_task_bulk_add_com_direct(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Path 1: ≤5 items, plain COM."""
    added = []
    for item in items:
        r = _msp_task_add_single(**item)
        if r.get("status") == "ok":
            added.append(r["task_id"])
    return {"status": "ok", "path": "com_direct", "count": len(added), "task_ids": added}


def _msp_task_bulk_add_com_batch(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Path 2: 6-19 items, COM with batch mode."""
    _enter_batch_mode()
    try:
        added = []
        for item in items:
            r = _msp_task_add_single(**item)
            if r.get("status") == "ok":
                added.append(r["task_id"])
        return {"status": "ok", "path": "com_batch", "count": len(added), "task_ids": added}
    finally:
        _exit_batch_mode()


def _msp_task_bulk_add_mspdi(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Path 3: 20+ items, MSPDI XML import."""
    import tempfile
    from msproject_bulk import MsprojectBulkWriter
    app = _validate_active_project()
    proj = app.ActiveProject

    # Get current project start for new tasks
    start_date = str(proj.ProjectStart) if proj.ProjectStart else None

    w = MsprojectBulkWriter(project_name=proj.Name or "Bulk", start_date=start_date)
    w.bulk_add_tasks(items)

    tmp = tempfile.NamedTemporaryFile(suffix=".xml", delete=False, mode="w", encoding="utf-8")
    tmp.close()
    w.save(tmp.name)

    try:
        # Insert task tree from XML into current project
        # MS Project: Tasks.InsertProject or FileOpen + copy
        # Simpler: open and merge — for Phase 1 use FileOpen → Tasks copy
        # NOTE: For now, use COM Path 2 fallback for MSPDI integration complexity.
        # Real Path 3 wiring done in Task 10 with proper Insert mechanism.
        _enter_batch_mode()
        added = []
        for item in items:
            r = _msp_task_add_single(**item)
            if r.get("status") == "ok":
                added.append(r["task_id"])
        return {"status": "ok", "path": "mspdi_bulk", "count": len(added),
                "task_ids": added, "note": "Phase 1: COM batch fallback; XML import wiring in T10"}
    finally:
        _exit_batch_mode()
        os.unlink(tmp.name) if os.path.exists(tmp.name) else None


def _msp_task_bulk_add(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Hybrid bulk add: routes by item count."""
    if not items:
        return {"status": "ok", "path": "noop", "count": 0, "task_ids": []}
    path = _route_operation(len(items))
    if path == "com_direct":
        return _msp_task_bulk_add_com_direct(items)
    elif path == "com_batch":
        return _msp_task_bulk_add_com_batch(items)
    else:  # mspdi_bulk
        return _msp_task_bulk_add_mspdi(items)
```

**Step 4: Çalıştır — pass**

```bash
python -m pytest tests/test_msproject_task_bulk.py -v -s
```

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_task_bulk.py
git commit -m "Phase 1 T9: msproject_task bulk_add with hybrid routing (COM direct/batch + MSPDI fallback)"
```

---

## Task 10: True MSPDI Path 3 Implementation

**Files:**
- Modify: `msproject_mcp_core.py`
- Modify: `tests/test_msproject_task_bulk.py` (add 100+ task perf test)

**Step 1: Failing test — 200 task <5sn**

Add to existing `tests/test_msproject_task_bulk.py`:
```python
def test_bulk_200_tasks_under_5_sec(msproject_app):
    """Performance: 200 tasks via MSPDI must finish <5 sec."""
    _cleanup_all_tasks(msproject_app.ActiveProject)
    items = [{"name": f"P{i}", "duration": "1d"} for i in range(200)]
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 200
    assert elapsed < 5.0, f"Bulk 200 took {elapsed}s (target: <5s)"
    _cleanup_all_tasks(msproject_app.ActiveProject)
```

**Step 2: Çalıştır — fail (current fallback uses COM batch which is slow)**

**Step 3: True MSPDI implementation**

Replace `_msp_task_bulk_add_mspdi` body with:

```python
def _msp_task_bulk_add_mspdi(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Path 3: 20+ items via MSPDI XML import.

    Strategy: Build temp MSPDI XML, open as new project via FileOpen,
    copy tasks into ActiveProject via clipboard or insert, close temp.
    """
    import tempfile
    from msproject_bulk import MsprojectBulkWriter
    app = _validate_active_project()
    target_proj = app.ActiveProject
    target_name = target_proj.Name

    start_date_obj = target_proj.ProjectStart
    start_date = start_date_obj.strftime("%Y-%m-%dT%H:%M:%S") if start_date_obj else None

    # Build XML
    w = MsprojectBulkWriter(project_name="_BULK_TEMP_", start_date=start_date)
    uids = w.bulk_add_tasks(items)

    tmp = tempfile.NamedTemporaryFile(suffix=".xml", delete=False, mode="w", encoding="utf-8")
    tmp.close()
    w.save(tmp.name)

    _enter_batch_mode()
    try:
        # Open temp XML as new project
        app.FileOpen(tmp.name)
        temp_proj = app.ActiveProject  # now this is _BULK_TEMP_

        # Verify
        if temp_proj.Name != "_BULK_TEMP_":
            raise RuntimeError(f"FileOpen didn't activate temp project (got {temp_proj.Name})")

        # Select all tasks in temp and copy
        # MS Project COM: SelectAll then EditCopy
        app.SelectAll()
        app.EditCopy()

        # Switch to target project
        for i in range(1, app.Projects.Count + 1):
            if app.Projects(i).Name == target_name:
                app.WindowActivate(app.Projects(i).Windows(1).Caption)
                break

        # Paste at end
        target_proj = app.ActiveProject
        # Move to last task and paste below
        last_id = target_proj.Tasks.Count
        if last_id > 0:
            app.SelectTaskField(Row=last_id + 1, Column="Name")
        else:
            app.SelectTaskField(Row=1, Column="Name")
        app.EditPaste()

        # Close temp without saving
        for i in range(1, app.Projects.Count + 1):
            if app.Projects(i).Name == "_BULK_TEMP_":
                app.WindowActivate(app.Projects(i).Windows(1).Caption)
                app.FileClose(0)  # 0 = don't save
                break

        # Reactivate target
        for i in range(1, app.Projects.Count + 1):
            if app.Projects(i).Name == target_name:
                app.WindowActivate(app.Projects(i).Windows(1).Caption)
                break

        return {
            "status": "ok",
            "path": "mspdi_bulk",
            "count": len(items),
            "method": "FileOpen + EditCopy + EditPaste",
        }
    except Exception as e:
        logger.error(f"MSPDI bulk path failed: {e}; falling back to COM batch")
        # Fallback to COM batch
        added = []
        for item in items:
            r = _msp_task_add_single(**item)
            if r.get("status") == "ok":
                added.append(r["task_id"])
        return {
            "status": "ok",
            "path": "mspdi_bulk_fallback_com",
            "count": len(added),
            "fallback_reason": str(e),
        }
    finally:
        _exit_batch_mode()
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
```

**Step 4: Test çalıştır**

```bash
python -m pytest tests/test_msproject_task_bulk.py::test_bulk_200_tasks_under_5_sec -v -s
```

If fails (XML import doesn't merge correctly), iterate on Insert/Paste mechanism. Acceptance: 200 task ≤5sn.

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_task_bulk.py
git commit -m "Phase 1 T10: true MSPDI bulk import for 20+ tasks (<5s for 200 tasks)"
```

---

## Task 11: `msproject_link` — `add` Action

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_link.py`

**Step 1: Failing test**

```python
import pytest
from msproject_mcp_core import _msp_task_add_single, _msp_link_add, _msp_task_delete


def test_link_two_tasks(msproject_app):
    a = _msp_task_add_single(name="LinkA", duration="2d")
    b = _msp_task_add_single(name="LinkB", duration="3d")
    r = _msp_link_add(predecessor_id=a["task_id"], successor_id=b["task_id"], type="FS", lag="0d")
    assert r["status"] == "ok"
    # Verify
    proj = msproject_app.ActiveProject
    bt = None
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t and t.ID == b["task_id"]:
            bt = t
            break
    assert bt and bt.Predecessors  # comma-list of pred IDs
    _msp_task_delete(task_id=a["task_id"])
    _msp_task_delete(task_id=b["task_id"])
```

**Step 2-4: Implement, test pass**

In `msproject_mcp_core.py`:
```python
PJ_LINK_TYPES = {"FF": 0, "FS": 1, "SF": 2, "SS": 3}


def _msp_link_add(predecessor_id: int, successor_id: int,
                  type: str = "FS", lag: str = "0d") -> Dict[str, Any]:
    """Add a single predecessor link via COM."""
    app = _validate_active_project()
    proj = app.ActiveProject
    pred = _find_task_by_id(proj, predecessor_id)
    succ = _find_task_by_id(proj, successor_id)
    if pred is None:
        return {"status": "error", "error": f"Predecessor {predecessor_id} not found"}
    if succ is None:
        return {"status": "error", "error": f"Successor {successor_id} not found"}
    try:
        # Append to existing predecessors string: "5FS,3SS+2d"
        existing = succ.Predecessors or ""
        lag_mins = _parse_duration(lag) if lag and lag != "0d" else 0
        lag_str = f"+{lag_mins}m" if lag_mins > 0 else (f"-{-lag_mins}m" if lag_mins < 0 else "")
        new_token = f"{predecessor_id}{type.upper()}{lag_str}"
        succ.Predecessors = (existing + "," + new_token).strip(",") if existing else new_token
        return {"status": "ok", "predecessor_id": predecessor_id,
                "successor_id": successor_id, "type": type.upper(), "lag": lag}
    except Exception as e:
        return {"status": "error", "error": str(e)}
```

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_link.py
git commit -m "Phase 1 T11: msproject_link add (single)"
```

---

## Task 12: `msproject_link` — `delete`, `update`, `bulk_add`, `chain`

**Files:**
- Modify: `msproject_mcp_core.py`
- Modify: `tests/test_msproject_link.py`

Same TDD pattern: write failing tests, implement, pass, commit.

Implementation outline:
```python
def _msp_link_delete(predecessor_id: int, successor_id: int) -> Dict[str, Any]:
    """Remove specific predecessor link."""
    # Re-build Predecessors string excluding the matching token


def _msp_link_update(predecessor_id: int, successor_id: int,
                    new_type: Optional[str] = None,
                    new_lag: Optional[str] = None) -> Dict[str, Any]:
    """Update link properties via remove+re-add."""


def _msp_link_bulk_add(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Hybrid routing for bulk links."""
    path = _route_operation(len(items))
    if path == "com_batch":
        _enter_batch_mode()
    try:
        added = 0
        for it in items:
            r = _msp_link_add(**it)
            if r.get("status") == "ok":
                added += 1
        return {"status": "ok", "path": path, "count": added}
    finally:
        if path == "com_batch":
            _exit_batch_mode()


def _msp_link_chain(task_ids: List[int], type: str = "FS",
                    lag: str = "0d") -> Dict[str, Any]:
    """Chain N tasks: T1→T2→T3→...→TN with given link type."""
    if len(task_ids) < 2:
        return {"status": "error", "error": "At least 2 tasks required"}
    added = 0
    for i in range(len(task_ids) - 1):
        r = _msp_link_add(predecessor_id=task_ids[i], successor_id=task_ids[i+1],
                          type=type, lag=lag)
        if r.get("status") == "ok":
            added += 1
    return {"status": "ok", "links_added": added, "chain_length": len(task_ids)}
```

**Test patterns:** verify delete removes, update changes type, bulk_add 30 links works, chain links 5 tasks.

```bash
git add msproject_mcp_core.py tests/test_msproject_link.py
git commit -m "Phase 1 T12: msproject_link delete/update/bulk_add/chain"
```

---

## Task 13: `msproject_schedule` — All 4 Actions

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_schedule.py`

**Step 1-4: TDD pattern for each action**

```python
def _msp_schedule_reschedule(report_date: Optional[str] = None) -> Dict[str, Any]:
    """Recalculate the project (CalculateProject)."""
    app = _validate_active_project()
    if report_date:
        app.ActiveProject.StatusDate = report_date
    app.CalculateProject()
    return {"status": "ok", "message": "Project rescheduled"}


def _msp_schedule_level(within_slack: bool = False) -> Dict[str, Any]:
    """Level resources."""
    app = _validate_active_project()
    # PjLevelOrder: pjLevelOrderID=0, pjLevelOrderStandard=1, pjLevelOrderPriority=2
    app.LevelAll(LevelOrder=1, LevelInLevelingOnly=False, LevelWithinSlack=within_slack)
    return {"status": "ok", "message": "Resources leveled"}


def _msp_schedule_set_data_date(date: str) -> Dict[str, Any]:
    """Set status date / data date."""
    app = _validate_active_project()
    app.ActiveProject.StatusDate = date
    return {"status": "ok", "status_date": date}


def _msp_schedule_protect_actuals(enable: bool = True) -> Dict[str, Any]:
    """Protect actuals from being recalculated."""
    app = _validate_active_project()
    proj = app.ActiveProject
    proj.MoveCompletedEndsBack = False
    proj.MoveCompletedEndsForward = False
    return {"status": "ok", "actuals_protected": enable}
```

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_schedule.py
git commit -m "Phase 1 T13: msproject_schedule all 4 actions (reschedule/level/data_date/protect)"
```

---

## Task 14: MCP Tool Dispatchers (FastMCP wiring)

**Files:**
- Modify: `msproject_mcp_core.py` (add @mcp.tool decorators)

**Step 1: Failing test**

`tests/test_msproject_dispatchers.py`:
```python
import asyncio
from msproject_mcp_core import msproject_task, msproject_link, msproject_schedule


def test_task_dispatcher_add(msproject_app):
    result = asyncio.run(msproject_task({"action": "add", "name": "Disp Test", "duration": "2d"}))
    import json
    parsed = json.loads(result) if isinstance(result, str) else result
    assert parsed["status"] == "ok"
```

**Step 2-4: Implement dispatchers**

```python
# ---------- TOOL DISPATCHERS ----------

@mcp.tool(
    name="msproject_task",
    annotations={"title": "MS Project Task Operations", "readOnlyHint": False},
)
async def msproject_task(params: dict) -> str:
    """Manage tasks in active MS Project (COM-based, hybrid speed routing).

    Actions:
    - add: Add task. Params: name, duration (e.g. "5d"), [start, finish, summary, milestone, notes]
    - update: Update task. Params: task_id, [name, duration, start, finish, percent_complete, notes]
    - delete: Delete task. Params: task_id
    - add_summary: Add summary task. Params: name, duration
    - add_milestone: Add 0-duration milestone. Params: name, [date]
    - get: Get task details. Params: task_id
    - list: List tasks. Params: [include_summary=true, limit=100]
    - bulk_add: Bulk add tasks. Params: items=[{name, duration, ...}, ...]
                Routes: 1-5=COM direct, 6-19=COM batch, 20+=MSPDI bulk

    Returns JSON-encoded result. Hybrid speed: 200 task bulk in ~3-5 sec.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "add":
            r = _msp_task_add_single(**p)
        elif action == "update":
            r = _msp_task_update(**p)
        elif action == "delete":
            r = _msp_task_delete(**p)
        elif action == "add_summary":
            r = _msp_task_add_summary(**p)
        elif action == "add_milestone":
            r = _msp_task_add_milestone(**p)
        elif action == "get":
            r = _msp_task_get(**p)
        elif action == "list":
            r = _msp_task_list(**p)
        elif action == "bulk_add":
            r = _msp_task_bulk_add(**p)
        else:
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: add/update/delete/add_summary/add_milestone/get/list/bulk_add"}
    except Exception as e:
        logger.error(f"msproject_task({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)


@mcp.tool(
    name="msproject_link",
    annotations={"title": "MS Project Link Operations", "readOnlyHint": False},
)
async def msproject_link(params: dict) -> str:
    """Manage predecessor/successor links."""
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "add":
            r = _msp_link_add(**p)
        elif action == "delete":
            r = _msp_link_delete(**p)
        elif action == "update":
            r = _msp_link_update(**p)
        elif action == "bulk_add":
            r = _msp_link_bulk_add(**p)
        elif action == "chain":
            r = _msp_link_chain(**p)
        else:
            r = {"status": "error", "error": f"Unknown action '{action}'"}
    except Exception as e:
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)


@mcp.tool(
    name="msproject_schedule",
    annotations={"title": "MS Project Schedule Operations", "readOnlyHint": False},
)
async def msproject_schedule(params: dict) -> str:
    """Schedule operations: reschedule, level, set_data_date, protect_actuals."""
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "reschedule":
            r = _msp_schedule_reschedule(**p)
        elif action == "level":
            r = _msp_schedule_level(**p)
        elif action == "set_data_date":
            r = _msp_schedule_set_data_date(**p)
        elif action == "protect_actuals":
            r = _msp_schedule_protect_actuals(**p)
        else:
            r = {"status": "error", "error": f"Unknown action '{action}'"}
    except Exception as e:
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_dispatchers.py
git commit -m "Phase 1 T14: FastMCP tool dispatchers (msproject_task/link/schedule)"
```

---

## Task 15: MCP Server Config Update

**Files:**
- Create: `samples/claude_mcp_config_snippet.json` (referans)

**Step 1: Config snippet hazırla**

`samples/claude_mcp_config_snippet.json`:
```json
{
  "mcpServers": {
    "asta_powerproject_mcp": {
      "command": "python",
      "args": ["C:\\Users\\CahAsus\\asta-powerproject-mcp\\asta_mcp_core.py"]
    },
    "asta_powerproject_file": {
      "command": "python",
      "args": ["C:\\Users\\CahAsus\\asta-powerproject-mcp\\asta_mcp_file.py"]
    },
    "msproject_mcp": {
      "command": "python",
      "args": ["C:\\Users\\CahAsus\\asta-powerproject-mcp\\msproject_mcp_core.py"]
    }
  }
}
```

**Step 2: Kullanıcıya talimat — Claude Code config'i güncelle**

```bash
# C:\Users\CahAsus\.claude.json içine "msproject_mcp" entry ekle
# Sonra Claude Code'u kapatıp tekrar aç
```

**Step 3: Smoke test — MCP server başlatılabilir mi**

```bash
python -c "from msproject_mcp_core import mcp; print('MCP ready:', mcp.name)"
```

Expected: `MCP ready: msproject_mcp`

**Step 4: Commit**

```bash
git add samples/claude_mcp_config_snippet.json
git commit -m "Phase 1 T15: MCP config snippet for msproject_mcp registration"
```

---

## Task 16: End-to-End Acceptance Test (200-Task Villa)

**Files:**
- Create: `samples/build_villa_msp.py`
- Create: `tests/fixtures/villa_200.csv` (data)

**Step 1: Test data hazırla**

`tests/fixtures/villa_200.csv` — 200 satır example villa task:
```csv
Name,Duration,Predecessors
Hafriyat,3d,
Temel,5d,1
... (200 rows)
```

(Kullanıcı bu CSV'yi sample data olarak hazırlar, otomatik üretilebilir.)

**Step 2: Build script**

`samples/build_villa_msp.py`:
```python
"""Build a 200-task villa project in active MS Project.
Acceptance test for Phase 1: must complete in <5 sec.
"""
import time, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from msproject_mcp_core import _msp_task_bulk_add, _msp_link_bulk_add


def main():
    items = []
    for i in range(200):
        items.append({"name": f"Villa T{i+1:03d}", "duration": "1d"})
    print(f"Building {len(items)} tasks...")
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start
    print(f"  Tasks: {r['count']} via {r['path']} in {elapsed:.2f}s")
    assert elapsed < 5.0, f"Too slow: {elapsed}s"
    print("OK Acceptance: <5 sec")


if __name__ == "__main__":
    main()
```

**Step 3: Çalıştır**

```bash
python samples/build_villa_msp.py
```

Expected: `OK Acceptance: <5 sec`. MS Project UI'de 200 task görünür.

**Step 4: Manuel doğrulama**
- MS Project UI'de scroll → 200 task var
- Her birinin duration "1d"
- Project bitiş tarihi makul

**Step 5: Cleanup ve commit**

```bash
# Kullanıcı manuel olarak Ctrl+Z ile geri alabilir veya boş projeye Save As yapar
git add samples/build_villa_msp.py tests/fixtures/villa_200.csv
git commit -m "Phase 1 T16: end-to-end acceptance test (200-task villa <5s)"
```

---

## Task 17: Phase 1 Final Suite & Push

**Step 1: Tüm test'leri çalıştır**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/ -v --tb=short
```

Expected: ALL PASS (skipped tests OK if MS Project not connected at given moment).

**Step 2: Code review checklist**
- [ ] Tüm tool'lar JSON return ediyor
- [ ] Türkçe karakter encoding sorunu yok (UTF-8)
- [ ] `atexit` mode restore çalışıyor (manuel test: bulk batch sırasında Python kill et, MS Project Calc mode otomatik düzeliyor mu)
- [ ] Hata mesajları eylem önerisi içeriyor
- [ ] Logging her tool call'da çalışıyor

**Step 3: Push'a hazırla**

```bash
git log --oneline -20  # Phase 1 commit'leri görünür olmalı
git status  # clean working tree
```

**Step 4: Push**

```bash
git push origin main
```

Expected: GitHub'a tüm Phase 1 commit'leri yüklendi.

**Step 5: Phase 1 onayı için kullanıcıya sun**

Kullanıcıya rapor:
- ✅ 3 tool çalışıyor (task, link, schedule)
- ✅ 200 task <5 sn'de yükleniyor
- ✅ Tüm testler PASS
- ✅ MCP config snippet hazır, Claude Code restart sonrası kullanıma açık
- ✅ GitHub'a push edildi

**Phase 2'ye geçiş için kullanıcı onayı beklenir.**

---

## Phase 1 Tamamlama Kriterleri (Re-verify)

1. ✅ `msproject_mcp_core.py` — FastMCP server + COM cache + hybrid routing
2. ✅ `msproject_bulk.py` — MSPDI bulk-write engine
3. ✅ 3 tool: `msproject_task` (8 action), `msproject_link` (5 action), `msproject_schedule` (4 action)
4. ✅ Unit tests passing (~25 test)
5. ✅ Integration tests passing (~10 test, requires MS Project)
6. ✅ Performance: 200 task ≤5 sn (acceptance test)
7. ✅ Bilinen bug 0
8. ✅ Commit + push GitHub
9. ⏸ Kullanıcı onayı (Phase 2'ye geçişten önce)

**Phase 2 başlangıç — kullanıcı onayı sonrası:** Resource + Calendar tool'ları (Uzbek bayramları + 6-ekip atama).

---

*Plan tamamlandı: 24 Nisan 2026*
*Tahmini Phase 1 süresi: ~12 saat (1.5 iş günü)*
