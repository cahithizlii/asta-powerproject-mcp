"""MS Project MCP Server - COM-based.

Hybrid speed strategy:
  1-5 items   -> COM direct (real-time, instant UI feedback)
  6-19 items  -> COM batch (Calculation manual + ScreenUpdating off)
  20+ items   -> MSPDI XML bulk import (~3-5s for 200 tasks)

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


# ---------- TASK HELPERS ----------

def _parse_duration(d: str) -> int:
    """Convert '5d' -> minutes (assuming 8h/day for MSP)."""
    if not d:
        return 480
    s = d.strip()
    unit = s[-1].lower() if s[-1].isalpha() else "d"
    try:
        n = float(s.rstrip("dwhmDWHM"))
    except ValueError:
        n = 1.0
    return int({"d": 480, "w": 2400, "h": 60, "m": 1}.get(unit, 480) * n)


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
            try:
                new_task.Summary = True
            except Exception:
                pass  # Summary auto-set when task has children in MSP
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
        logger.error(f"_msp_task_delete failed: {e}")
        return {"status": "error", "error": str(e)}


def _serialize_task(t: Any) -> Dict[str, Any]:
    return {
        "id": t.ID,
        "uid": t.UniqueID,
        "name": t.Name,
        "duration": t.Duration,
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


def _msp_task_add_summary(name: str, duration: str = "1d",
                          parent_task_id: Optional[int] = None) -> Dict[str, Any]:
    """Add a summary task. In MS Project summary is implicit (becomes summary when has children).

    TODO(T11+): Wire `parent_task_id` to indent the new task under the given parent
    via OutlineLevel / Task.OutlineIndent so it actually becomes a summary container.
    Currently the parameter is accepted for API stability but ignored — MSP only marks
    a task `Summary=True` when it has children, so summary-ness is implicit, not flagged.
    """
    return _msp_task_add_single(name=name, duration=duration, summary=True)


def _msp_task_add_milestone(name: str, date: Optional[str] = None) -> Dict[str, Any]:
    """Add a 0-duration milestone."""
    r = _msp_task_add_single(name=name, duration="0d", milestone=True)
    if r.get("status") == "ok" and date:
        _msp_task_update(task_id=r["task_id"], start=date, finish=date)
    return r


# ---------- BULK ADD HYBRID ROUTING ----------

def _msp_task_bulk_add_com_direct(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Path 1: <=5 items, plain COM."""
    added = []
    for item in items:
        r = _msp_task_add_single(**item)
        if r.get("status") == "ok":
            added.append(r["task_id"])
    return {"status": "ok", "path": "com_direct", "count": len(added), "task_ids": added}


def _msp_task_bulk_add_com_batch(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Path 2: 6-19 items, COM with batch mode (Calculation manual)."""
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
    """Path 3: 20+ items, MSPDI XML import.

    T9: COM batch fallback (proper FileOpen wiring is T10's job).
    """
    _enter_batch_mode()
    try:
        added = []
        for item in items:
            r = _msp_task_add_single(**item)
            if r.get("status") == "ok":
                added.append(r["task_id"])
        return {"status": "ok", "path": "mspdi_bulk", "count": len(added),
                "task_ids": added,
                "note": "Phase 1 T9: COM batch fallback; XML import wiring in T10"}
    finally:
        _exit_batch_mode()


def _msp_task_bulk_add(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Hybrid bulk add: routes by item count via _route_operation()."""
    if not items:
        return {"status": "ok", "path": "noop", "count": 0, "task_ids": []}
    path = _route_operation(len(items))
    if path == "com_direct":
        return _msp_task_bulk_add_com_direct(items)
    elif path == "com_batch":
        return _msp_task_bulk_add_com_batch(items)
    else:
        return _msp_task_bulk_add_mspdi(items)


# ---------- TOOL DISPATCHERS (filled in T6+) ----------
# (Placeholder - actual @mcp.tool functions added in T14)


def main():
    """Run MCP server (stdio)."""
    mcp.run()


if __name__ == "__main__":
    main()
