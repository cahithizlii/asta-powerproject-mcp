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

import datetime as _dt

import pythoncom
import pywintypes
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
        "Tools: msproject_task, msproject_link, msproject_schedule, msproject_calendar."
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


def _format_com_error(e: Exception) -> str:
    """Extract a human-readable message from pywintypes.com_error or fallback to str(e).

    pywintypes.com_error.args = (hresult, msg, excepinfo, argerr) where
    excepinfo = (wCode, source, description, helpFile, helpContext, scode).
    Description (excepinfo[2]) is the user-friendly message; fall back to
    msg (args[1]) or str(e) if unavailable.
    """
    try:
        if hasattr(e, "args") and len(e.args) >= 3 and isinstance(e.args[2], tuple):
            excepinfo = e.args[2]
            if len(excepinfo) >= 3 and excepinfo[2]:
                return str(excepinfo[2]).strip()
            if len(e.args) >= 2 and e.args[1]:
                return str(e.args[1]).strip()
    except Exception:
        pass
    return str(e)


def _enter_batch_mode():
    """Enter COM batch mode: disable screen update, manual calc, no events."""
    with _app_lock:
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
    with _app_lock:
        try:
            app = _connect_app()
            if _calc_modified:
                pj_auto = 1  # PjCalculation.pjAutomatic
                app.Calculation = pj_auto
            if _screenupdating_modified:
                app.ScreenUpdating = True
            proj = app.ActiveProject
            if proj:
                try:
                    proj.EventsEnabled = True
                except Exception:
                    pass
        except Exception as e:
            logger.warning(f"_exit_batch_mode error (non-fatal): {e}")
        finally:
            # Reset flags even if restore failed — otherwise stuck "modified"
            # state would prevent future entries from re-applying batch mode.
            _calc_modified = False
            _screenupdating_modified = False


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

# Calendar dispatcher: actions that natively use 'calendar_name' and accept
# 'name' as alias.
_CALENDAR_NAME_ALIAS_ACTIONS = frozenset({
    "add_exception", "assign_to_task", "assign_to_resource",
    "list", "holidays_uzbek",
})
# Calendar dispatcher: actions that natively use 'name' and accept
# 'calendar_name' as alias.
_CALENDAR_NAME_NATIVE_ACTIONS = frozenset({"create", "update"})


# ---------- CALENDAR HELPERS ----------

def _find_calendar_by_name(proj: Any, name: str) -> Optional[Any]:
    """Locate a base calendar object in the project. Returns None if not found."""
    for i in range(1, proj.BaseCalendars.Count + 1):
        cal = proj.BaseCalendars(i)
        if cal is not None and cal.Name == name:
            return cal
    return None


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
        return {"status": "ok", "calendar_uid": cal.Guid, "name": name}
    except Exception as e:
        logger.error(f"_msp_calendar_create({name}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


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

    # Pre-flight: validate ALL inputs before any mutation (no partial writes)
    do_rename = new_name is not None and new_name != name
    if do_rename and _find_calendar_by_name(proj, new_name) is not None:
        return {"status": "error",
                "error": f"Calendar '{new_name}' already exists"}
    if weekday_off is not None and not (1 <= weekday_off <= 7):
        return {"status": "error",
                "error": "weekday_off must be 1-7 (1=Sunday, 7=Saturday)"}

    changes = []
    try:
        if do_rename:
            cal.Name = new_name
            changes.append("name")
        if weekday_off is not None:
            wd = cal.WeekDays(weekday_off)
            wd.Working = False
            changes.append("weekday_off")
        return {"status": "ok", "calendar_name": new_name or name, "changes": changes}
    except Exception as e:
        logger.error(f"_msp_calendar_update({name}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


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
        return {"status": "error",
                "error": f"Calendar '{calendar_name}' not found in project"}

    if working:
        return {"status": "error",
                "error": "working=True is not yet supported (Phase 3+); only non-working exceptions are supported in Phase 2a"}

    # Pre-flight: validate ALL inputs before any mutation (no partial writes)
    try:
        start_d = _parse_date(start)
        finish_d = _parse_date(finish) if finish else start_d
    except ValueError as e:
        return {"status": "error",
                "error": f"Invalid date format (expected YYYY-MM-DD): {e}"}
    if finish_d < start_d:
        return {"status": "error",
                "error": "Start date must be <= finish date"}

    try:
        ex = cal.Exceptions.Add(
            Type=PJ_EXCEPTION_DAILY,
            Start=pywintypes.Time(start_d),
            Finish=pywintypes.Time(finish_d),
        )
        ex.Name = exception_name
        # Type=PJ_EXCEPTION_DAILY=7 already implies non-working in MSP semantics;
        # MSP 16.0 exposes shift times via ex.Shift1.Start sub-objects (not flat
        # ShiftNStart props), so any zeroing is best handled if/when working=True
        # support arrives in Phase 3+.
        # NOTE: 'working' input intentionally not echoed — always False post-T21
        # guard (working=True returns error early in pre-flight).
        return {"status": "ok",
                "calendar_name": calendar_name,
                "exception_name": exception_name,
                "start": start,
                "finish": finish or start}
    except Exception as e:
        logger.error(
            f"_msp_calendar_add_exception({calendar_name},{exception_name}) failed: {e}"
        )
        return {"status": "error", "error": _format_com_error(e)}


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
        return {"status": "error", "error": _format_com_error(e)}


def _find_resource_by_id(proj: Any, resource_id: int) -> Optional[Any]:
    """Locate a resource by ID. Phase 2a helper; reused/expanded in Phase 2b."""
    for i in range(1, proj.Resources.Count + 1):
        r = proj.Resources(i)
        if r is not None and r.ID == resource_id:
            return r
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
        return {"status": "error", "error": _format_com_error(e)}


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
                "calendar_uid": cal.Guid,
                "name": cal.Name,
                "exception_count": ex_count,
            })
        return {"status": "ok", "count": len(out), "calendars": out}
    except Exception as e:
        logger.error(f"_msp_calendar_list failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_calendar_holidays_uzbek(calendar_name: str, year: int = 2026) -> Dict[str, Any]:
    """Bulk-add 9 official Özbekistan public holidays to a calendar.

    Idempotent: name-based dedup. Re-running on a calendar that already
    has matching named exceptions skips them (returns them in `skipped`).
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    cal = _find_calendar_by_name(proj, calendar_name)
    if cal is None:
        return {"status": "error",
                "error": f"Calendar '{calendar_name}' not found in project"}

    # Pre-scan existing exception names for dedup
    existing_names = set()
    try:
        for i in range(1, cal.Exceptions.Count + 1):
            ex = cal.Exceptions(i)
            if ex is not None and ex.Name:
                existing_names.add(ex.Name)
    except Exception:
        pass  # if we can't read exceptions, treat as none-existing

    added = []
    skipped = []
    failures = []
    for name, month, day in UZBEK_HOLIDAYS_2026:
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
        entry = {"name": name, "date": date_str, "month": month, "day": day}
        if name in existing_names:
            skipped.append({**entry, "reason": "already exists"})
            continue
        r = _msp_calendar_add_exception(
            calendar_name=calendar_name,
            exception_name=name,
            start=date_str,
        )
        if r.get("status") == "ok":
            added.append(entry)
        else:
            failures.append({**entry, "error": r.get("error")})

    if failures:
        status = "partial"
    elif not added and skipped:
        status = "already_done"
    else:
        status = "ok"

    if failures:
        logger.warning(f"holidays_uzbek partial: {len(added)} added, {len(skipped)} skipped, {len(failures)} failed")

    return {
        "status": status,
        "calendar_name": calendar_name,
        "year": year,
        "count": len(added),
        "skipped_count": len(skipped),
        "holidays": added,
        "skipped": skipped,
        "failures": failures,
    }


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
    """Path 3: 20+ items via MSPDI XML import.

    Strategy:
    1. Build temp MSPDI XML with all tasks via MsprojectBulkWriter
    2. Open as new project via app.FileOpen(temp.xml) — DisplayAlerts=False
       suppresses the Import Wizard dialog. The temp project's name is the
       XML filename stem (NOT the <Name> element), so we control it via the
       temp filename.
    3. SelectAll + EditCopy in temp project
    4. Switch back to target project, SelectTaskField row=last_count+1, EditPaste
    5. EditPaste sometimes inserts a leading None placeholder row; remove it
       via app.SelectRow(Row=1) + app.EditDelete (only if Tasks(1) is None)
    6. Close temp project without saving (FileClose 0)
    7. Cleanup temp file

    Performance target: 200 tasks in <5 seconds.

    On failure, falls back to COM-batch-per-task (slower but reliable).
    """
    import tempfile
    from msproject_bulk import MsprojectBulkWriter

    app = _validate_active_project()
    target_proj = app.ActiveProject
    target_name = target_proj.Name

    # Get current project start for new tasks (use as MSPDI StartDate)
    try:
        start_date_obj = target_proj.ProjectStart
        start_date = start_date_obj.strftime("%Y-%m-%dT%H:%M:%S") if start_date_obj else None
    except Exception:
        start_date = None

    # Use a controlled temp filename so we can find the loaded project by name
    # (FileOpen names the project after the file's stem, not the XML <Name>).
    tmp_dir = tempfile.gettempdir()
    tmp_stem = f"_BULK_TEMP_{int(time.time() * 1000)}"
    tmp_path = os.path.join(tmp_dir, f"{tmp_stem}.xml")

    # Build XML
    w = MsprojectBulkWriter(project_name=tmp_stem, start_date=start_date)
    w.bulk_add_tasks(items)
    w.save(tmp_path)

    # Save & flip DisplayAlerts (this is what suppresses the Import Wizard dialog)
    prev_alerts = True
    try:
        prev_alerts = app.DisplayAlerts
    except Exception:
        pass
    try:
        app.DisplayAlerts = False
    except Exception:
        pass

    _enter_batch_mode()
    try:
        # Open temp XML as new project
        app.FileOpen(tmp_path)

        # The loaded project's name is the file stem
        temp_proj = None
        for i in range(1, app.Projects.Count + 1):
            if app.Projects(i).Name == tmp_stem:
                temp_proj = app.Projects(i)
                break
        if temp_proj is None:
            raise RuntimeError(
                f"FileOpen didn't load {tmp_stem} (Projects: "
                f"{[app.Projects(i).Name for i in range(1, app.Projects.Count + 1)]})"
            )

        # Select all tasks in temp and copy
        try:
            app.SelectAll()
        except Exception:
            # Fallback: select via row range
            app.SelectTaskField(Row=1, Column="Name", Height=temp_proj.Tasks.Count)
        app.EditCopy()

        # Switch to target project window
        target_window_found = False
        for i in range(1, app.Projects.Count + 1):
            if app.Projects(i).Name == target_name:
                try:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    target_window_found = True
                    break
                except Exception:
                    continue
        if not target_window_found:
            raise RuntimeError(f"Could not switch back to target project {target_name}")

        target_proj = app.ActiveProject
        last_count = target_proj.Tasks.Count
        paste_row = last_count + 1 if last_count > 0 else 1
        try:
            app.SelectTaskField(Row=paste_row, Column="Name")
        except Exception:
            app.SelectTaskField(Row=1, Column="Name")
        app.EditPaste()

        # Cleanup leading None placeholder rows that EditPaste sometimes inserts.
        # Use app.SelectRow(Row=1, RowRelative=False) + app.EditDelete to delete
        # the row by position (not by Tasks(N).Delete which fails on None).
        cleanup_safety = 50
        while (target_proj.Tasks.Count > 0
               and target_proj.Tasks(1) is None
               and cleanup_safety > 0):
            try:
                app.SelectRow(Row=1, RowRelative=False)
                app.EditDelete()
            except Exception as ce:
                logger.warning(f"None-row cleanup at row 1 failed: {ce}")
                break
            cleanup_safety -= 1

        # Cleanup trailing None rows too (defensive)
        cleanup_safety = 50
        while (target_proj.Tasks.Count > 0
               and target_proj.Tasks(target_proj.Tasks.Count) is None
               and cleanup_safety > 0):
            try:
                app.SelectRow(Row=target_proj.Tasks.Count, RowRelative=False)
                app.EditDelete()
            except Exception:
                break
            cleanup_safety -= 1

        # Close temp project without saving
        for i in range(1, app.Projects.Count + 1):
            if app.Projects(i).Name == tmp_stem:
                try:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)  # 0 = pjDoNotSave
                    break
                except Exception:
                    pass

        # Reactivate target window
        for i in range(1, app.Projects.Count + 1):
            if app.Projects(i).Name == target_name:
                try:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    break
                except Exception:
                    pass

        # Collect IDs of newly added tasks (everything past last_count)
        target_proj = app.ActiveProject
        added_task_ids = []
        for i in range(last_count + 1, target_proj.Tasks.Count + 1):
            t = target_proj.Tasks(i)
            if t is not None:
                added_task_ids.append(t.ID)

        return {
            "status": "ok",
            "path": "mspdi_bulk",
            "count": len(added_task_ids),
            "task_ids": added_task_ids,
            "method": "FileOpen + EditCopy + EditPaste",
        }
    except Exception as e:
        logger.error(f"MSPDI bulk path failed: {e}; falling back to COM batch")
        # Cleanup any orphan _BULK_TEMP_ project
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == tmp_stem:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    break
        except Exception:
            pass
        # Switch back to target
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == target_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    break
        except Exception:
            pass
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
            "task_ids": added,
            "fallback_reason": str(e),
        }
    finally:
        _exit_batch_mode()
        try:
            app.DisplayAlerts = prev_alerts
        except Exception:
            pass
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


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


# ---------- LINK HELPERS ----------

PJ_LINK_TYPES = {"FF": 0, "FS": 1, "SF": 2, "SS": 3}
PJ_LINK_REVERSE = {v: k for k, v in PJ_LINK_TYPES.items()}


def _msp_link_add(predecessor_id: int, successor_id: int,
                  type: str = "FS", lag: str = "0d") -> Dict[str, Any]:
    """Add a single predecessor link via COM (Path 1).

    Uses Predecessors string append: "5FS,3SS+2d" format.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    pred = _find_task_by_id(proj, predecessor_id)
    succ = _find_task_by_id(proj, successor_id)
    if pred is None:
        return {"status": "error", "error": f"Predecessor {predecessor_id} not found"}
    if succ is None:
        return {"status": "error", "error": f"Successor {successor_id} not found"}
    try:
        existing = succ.Predecessors or ""
        # Build new token: "5FS" or "5FS+2d" or "5FS-1d"
        type_str = type.upper() if type and type.upper() in PJ_LINK_TYPES else "FS"
        lag_str = ""
        if lag and lag != "0d":
            # Convert minutes to display format
            mins = _parse_duration(lag)
            if mins > 0:
                # Days if multiple of 480 mins
                if mins % 480 == 0:
                    lag_str = f"+{mins // 480}d"
                elif mins % 60 == 0:
                    lag_str = f"+{mins // 60}h"
                else:
                    lag_str = f"+{mins}m"
            elif mins < 0:
                if mins % 480 == 0:
                    lag_str = f"-{abs(mins) // 480}d"
                else:
                    lag_str = f"-{abs(mins)}m"
        new_token = f"{predecessor_id}{type_str}{lag_str}"
        succ.Predecessors = (existing + "," + new_token).strip(",") if existing else new_token
        return {
            "status": "ok",
            "predecessor_id": predecessor_id,
            "successor_id": successor_id,
            "type": type_str,
            "lag": lag,
        }
    except Exception as e:
        logger.error(f"_msp_link_add failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_link_delete(predecessor_id: int, successor_id: int) -> Dict[str, Any]:
    """Remove a specific predecessor link by rebuilding Predecessors string."""
    app = _validate_active_project()
    succ = _find_task_by_id(app.ActiveProject, successor_id)
    if succ is None:
        return {"status": "error", "error": f"Successor {successor_id} not found"}
    try:
        existing = succ.Predecessors or ""
        if not existing:
            return {"status": "ok", "message": "No predecessors to remove"}
        # Parse tokens (comma-separated): "5FS+2d", "3SS", "12FF-1d"
        tokens = [t.strip() for t in existing.split(",") if t.strip()]
        # Filter out matching predecessor (number prefix matches predecessor_id)
        kept = []
        removed_count = 0
        for tok in tokens:
            # Extract numeric prefix
            num_str = ""
            for ch in tok:
                if ch.isdigit():
                    num_str += ch
                else:
                    break
            try:
                tok_id = int(num_str) if num_str else -1
            except ValueError:
                tok_id = -1
            if tok_id == predecessor_id:
                removed_count += 1
            else:
                kept.append(tok)
        succ.Predecessors = ",".join(kept)
        if removed_count == 0:
            return {"status": "ok", "message": f"No link from {predecessor_id} to {successor_id} found"}
        return {"status": "ok", "predecessor_id": predecessor_id, "successor_id": successor_id,
                "removed_count": removed_count}
    except Exception as e:
        logger.error(f"_msp_link_delete failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_link_update(predecessor_id: int, successor_id: int,
                     new_type: Optional[str] = None,
                     new_lag: Optional[str] = None) -> Dict[str, Any]:
    """Update link properties via remove+re-add."""
    delete_result = _msp_link_delete(predecessor_id=predecessor_id, successor_id=successor_id)
    if delete_result.get("status") != "ok":
        return delete_result
    add_result = _msp_link_add(
        predecessor_id=predecessor_id,
        successor_id=successor_id,
        type=new_type or "FS",
        lag=new_lag or "0d",
    )
    return add_result


def _msp_link_bulk_add(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Hybrid routing for bulk links (uses _route_operation)."""
    if not items:
        return {"status": "ok", "path": "noop", "count": 0}
    path = _route_operation(len(items))
    if path == "com_batch" or path == "mspdi_bulk":
        _enter_batch_mode()
    try:
        added = 0
        errors = []
        for it in items:
            r = _msp_link_add(**it)
            if r.get("status") == "ok":
                added += 1
            else:
                errors.append(r)
        result = {"status": "ok", "path": path, "count": added}
        if errors:
            result["errors"] = errors[:10]  # cap to avoid huge response
        return result
    finally:
        if path in ("com_batch", "mspdi_bulk"):
            _exit_batch_mode()


def _msp_link_chain(task_ids: List[int], type: str = "FS",
                    lag: str = "0d") -> Dict[str, Any]:
    """Chain N tasks: T1->T2->T3->...->TN with given link type."""
    if len(task_ids) < 2:
        return {"status": "error", "error": "At least 2 tasks required for chain"}
    added = 0
    errors = []
    for i in range(len(task_ids) - 1):
        r = _msp_link_add(
            predecessor_id=task_ids[i],
            successor_id=task_ids[i+1],
            type=type, lag=lag,
        )
        if r.get("status") == "ok":
            added += 1
        else:
            errors.append({"index": i, "error": r.get("error")})
    return {
        "status": "ok",
        "links_added": added,
        "chain_length": len(task_ids),
        "errors": errors[:10] if errors else [],
    }


# ---------- SCHEDULE HELPERS ----------

def _to_com_date(date_str: str) -> Any:
    """Convert ISO date string ('2026-04-30') to a pywintypes.Time COM date.

    MS Project's StatusDate property requires a real VT_DATE value; passing a
    plain string raises 'argument value is not valid'. Accepts both 'YYYY-MM-DD'
    and 'YYYY-MM-DD HH:MM:SS'.
    """
    s = date_str.strip()
    fmts = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y")
    last_err: Optional[Exception] = None
    for f in fmts:
        try:
            dt = _dt.datetime.strptime(s, f)
            return pywintypes.Time(dt)
        except Exception as e:
            last_err = e
    raise ValueError(f"Unrecognized date format: {date_str!r} ({last_err})")


def _msp_schedule_reschedule(report_date: Optional[str] = None) -> Dict[str, Any]:
    """Recalculate the project (CalculateProject)."""
    app = _validate_active_project()
    try:
        if report_date:
            app.ActiveProject.StatusDate = _to_com_date(report_date)
        app.CalculateProject()
        return {"status": "ok", "message": "Project rescheduled",
                "report_date": report_date or "unchanged"}
    except Exception as e:
        logger.error(f"_msp_schedule_reschedule failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_schedule_level(within_slack: bool = False) -> Dict[str, Any]:
    """Level resources (Application.LevelNow)."""
    app = _validate_active_project()
    try:
        # MS Project COM uses LevelNow(All=True) — no LevelAll method exists.
        # LevelingOptions controls within_slack behavior.
        if within_slack:
            try:
                app.LevelingOptions(DelayInSlack=True)
            except Exception as e:
                logger.warning(f"LevelingOptions(DelayInSlack=True) failed (non-fatal): {e}")
        app.LevelNow(All=True)
        return {"status": "ok", "message": "Resources leveled",
                "within_slack": within_slack}
    except Exception as e:
        logger.error(f"_msp_schedule_level failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_schedule_set_data_date(date: str) -> Dict[str, Any]:
    """Set status date / data date."""
    app = _validate_active_project()
    try:
        app.ActiveProject.StatusDate = _to_com_date(date)
        return {"status": "ok", "status_date": date}
    except Exception as e:
        logger.error(f"_msp_schedule_set_data_date failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_schedule_protect_actuals(enable: bool = True) -> Dict[str, Any]:
    """Protect actuals from being recalculated.

    In MS Project these are Project-level properties that govern whether the
    scheduling engine moves end-of-completed and start-of-remaining parts when
    rescheduling. enable=True locks them down (no move).

      MoveCompleted          Move end of completed parts after status date back to status date
      AndMoveRemaining       And move start of remaining parts back to status date
      MoveRemaining          Move start of remaining parts before status date forward to status date
      AndMoveCompleted       And move end of completed parts forward to status date
    """
    app = _validate_active_project()
    try:
        proj = app.ActiveProject
        # When 'enable=True' = protect (don't move any completed/remaining parts)
        proj.MoveCompleted = not enable
        proj.AndMoveRemaining = not enable
        proj.MoveRemaining = not enable
        proj.AndMoveCompleted = not enable
        return {"status": "ok", "actuals_protected": bool(enable)}
    except Exception as e:
        logger.error(f"_msp_schedule_protect_actuals failed: {e}")
        return {"status": "error", "error": str(e)}


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
    - add_summary: Add summary task. Params: name, duration, [parent_task_id]
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
    """Manage predecessor/successor links.

    Actions:
    - add: Add link. Params: predecessor_id, successor_id, [type='FS', lag='0d']
    - delete: Remove link. Params: predecessor_id, successor_id
    - update: Update link (type/lag). Params: predecessor_id, successor_id, [new_type, new_lag]
    - bulk_add: Bulk links. Params: items=[{predecessor_id, successor_id, type, lag}, ...]
    - chain: Chain N tasks T1->T2->...->TN. Params: task_ids=[1,2,3,...], [type='FS', lag='0d']
    """
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
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: add/delete/update/bulk_add/chain"}
    except Exception as e:
        logger.error(f"msproject_link({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)


@mcp.tool(
    name="msproject_schedule",
    annotations={"title": "MS Project Schedule Operations", "readOnlyHint": False},
)
async def msproject_schedule(params: dict) -> str:
    """Schedule operations.

    Actions:
    - reschedule: Recalculate (CalculateProject). Params: [report_date='YYYY-MM-DD']
    - level: Resource leveling (LevelNow). Params: [within_slack=False]
    - set_data_date: Set status_date. Params: date='YYYY-MM-DD'
    - protect_actuals: Lock actuals. Params: [enable=True]
    """
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
            r = {"status": "error",
                 "error": f"Unknown action '{action}'. Valid: reschedule/level/set_data_date/protect_actuals"}
    except Exception as e:
        logger.error(f"msproject_schedule({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)


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
    - holidays_uzbek: Bulk-add 9 Ozbekistan 2026 official holidays (idempotent, name-based dedup). Params: calendar_name, [year=2026]

    Phase 2a (27 Apr 2026). Resource integration arrives in Phase 2b.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    # Alias: accept 'name' / 'calendar_name' interchangeably across actions.
    # Reject ambiguous calls where both keys are provided rather than silently
    # dropping one — surfaces caller bugs instead of hiding them.
    if "name" in p and "calendar_name" in p:
        return json.dumps(
            {"status": "error",
             "error": "Specify either 'name' or 'calendar_name', not both"},
            default=str, ensure_ascii=False,
        )
    if action in _CALENDAR_NAME_ALIAS_ACTIONS and "name" in p and "calendar_name" not in p:
        p["calendar_name"] = p.pop("name")
    elif action in _CALENDAR_NAME_NATIVE_ACTIONS and "calendar_name" in p and "name" not in p:
        p["name"] = p.pop("calendar_name")
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
        r = {"status": "error", "error": _format_com_error(e)}
    return json.dumps(r, default=str, ensure_ascii=False)


def main():
    """Run MCP server (stdio)."""
    mcp.run()


if __name__ == "__main__":
    main()
