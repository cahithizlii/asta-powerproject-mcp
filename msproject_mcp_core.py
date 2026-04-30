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
from typing import Any, Callable, Dict, List, Optional, Tuple

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
        "Tools: msproject_task, msproject_link, msproject_schedule, msproject_calendar, msproject_resource, msproject_baseline, msproject_progress."
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
    """List all base calendars in the active project with exception counts.

    Order: matches `proj.BaseCalendars` enumeration (typically insertion
    order, not sorted). If callers need lexicographic ordering, sort
    client-side on the `name` field.
    """
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
    except Exception as e:
        logger.debug(f"holidays_uzbek pre-scan failed (treating calendar as empty): {_format_com_error(e)}")

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


def _parse_rate(raw: Any) -> float:
    """Parse a MS Project rate value to float.

    MS Project COM returns rates as locale-formatted strings (e.g. '$50.00/h',
    '₺50,00/hr', '50,00 €/sa'). Strip currency symbols + per-unit suffix and
    normalize decimal separator. Returns 0.0 for empty/unparseable values.
    """
    if raw is None or raw == "":
        return 0.0
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip()
    if not s:
        return 0.0
    # Drop per-unit suffix after '/' (e.g. '/h', '/hr', '/sa', '/saat')
    if "/" in s:
        s = s.split("/", 1)[0]
    # Strip everything except digits, separators, and minus sign
    cleaned = "".join(ch for ch in s if ch.isdigit() or ch in ".,-")
    if not cleaned or cleaned in ("-", ".", ","):
        return 0.0
    # Locale-aware decimal separator: if both present, last one wins as decimal
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        # Comma-only: treat as decimal separator (TR, EU locales)
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _serialize_resource(res: Any) -> Dict[str, Any]:
    """Type-aware serialization. MaxUnits is COM-stored as fraction (1.0 = 100%);
    we expose as percentage for symmetry with assignment Units. Rates are
    locale-formatted strings from COM and are parsed to float via _parse_rate.

    Per-field reads are guarded so one bad COM value (corrupt/stale row) does
    not kill list iteration in callers like T35 resource_list. id/uid/name are
    intentionally unguarded — failure there means the resource is genuinely
    broken and the caller should know.
    """
    type_code = 0
    try:
        type_code = int(res.Type) if res.Type is not None else 0
    except Exception:
        pass
    type_name = RESOURCE_TYPE_NAMES.get(type_code, "Work")
    out: Dict[str, Any] = {
        "id": res.ID,
        "uid": res.UniqueID,
        "name": res.Name,
        "type": type_name,
    }
    # Type-specific properties — guarded so a single bad field doesn't break list iteration
    try:
        if type_name == "Work":
            out["max_units"] = float(res.MaxUnits) * 100.0  # 1.0 -> 100%
            out["standard_rate"] = _parse_rate(res.StandardRate)
            out["overtime_rate"] = _parse_rate(res.OvertimeRate)
        elif type_name == "Material":
            out["material_label"] = res.MaterialLabel or ""
            out["standard_rate"] = _parse_rate(res.StandardRate)
        elif type_name == "Cost":
            out["standard_rate"] = _parse_rate(res.StandardRate)
    except Exception:
        # Field read failed — return what we have (id/uid/name/type)
        pass
    return out


def _msp_resource_add(name: str, type: str = "Work",
                     max_units: Optional[float] = None,
                     standard_rate: Optional[float] = None,
                     overtime_rate: Optional[float] = None,
                     material_label: Optional[str] = None) -> Dict[str, Any]:
    """Add a resource. Type: 'Work' (default) | 'Material' | 'Cost'.

    max_units in % (100 = 1 person, 500 = 5-person crew). Stored in COM as
    fraction. MSP's documented ceiling is 6000% (60.0 fraction).
    standard_rate / overtime_rate in $/hour (Work) or $/unit (Material).
    material_label e.g. 'kg', 'm³', 'ton'.

    Atomic: if any post-Add property set fails, the orphan resource is
    deleted before returning error.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    # Pre-flight validation
    if type not in RESOURCE_TYPES:
        return {"status": "error",
                "error": f"Invalid type '{type}'. Valid: Work/Material/Cost"}
    if _find_resource_by_name(proj, name) is not None:
        return {"status": "error", "error": f"Resource '{name}' already exists"}
    # Pre-flight value validation (Fix 2)
    if max_units is not None and max_units < 0:
        return {"status": "error", "error": "max_units must be >= 0"}
    if standard_rate is not None and standard_rate < 0:
        return {"status": "error", "error": "standard_rate must be >= 0"}
    if overtime_rate is not None and overtime_rate < 0:
        return {"status": "error", "error": "overtime_rate must be >= 0"}
    try:
        res = proj.Resources.Add(name)
    except Exception as e:
        logger.error(f"_msp_resource_add({name},{type}) Resources.Add failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}
    # Post-Add: any failure here must rollback the orphan
    try:
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
        logger.error(f"_msp_resource_add({name},{type}) property-set failed (rolling back): {e}")
        try:
            res.Delete()
        except Exception as del_err:
            logger.warning(f"_msp_resource_add rollback delete failed: {del_err}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_resource_update(resource_id: int,
                        name: Optional[str] = None,
                        max_units: Optional[float] = None,
                        standard_rate: Optional[float] = None,
                        overtime_rate: Optional[float] = None,
                        material_label: Optional[str] = None) -> Dict[str, Any]:
    """Update a resource. Pre-flight validates ALL inputs before mutation
    (no partial writes — T20 lesson).

    max_units in % (100 = 1 person). standard_rate / overtime_rate in $/unit.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    res = _find_resource_by_id(proj, resource_id)
    if res is None:
        return {"status": "error", "error": f"Resource ID {resource_id} not found"}

    # Pre-flight: validate ALL inputs before any mutation
    do_rename = name is not None and name != res.Name
    if do_rename and _find_resource_by_name(proj, name) is not None:
        return {"status": "error", "error": f"Resource '{name}' already exists"}
    if max_units is not None and max_units < 0:
        return {"status": "error", "error": "max_units must be >= 0"}
    if standard_rate is not None and standard_rate < 0:
        return {"status": "error", "error": "standard_rate must be >= 0"}
    if overtime_rate is not None and overtime_rate < 0:
        return {"status": "error", "error": "overtime_rate must be >= 0"}

    # Type-aware property validation: each property must match the resource's type
    res_type_code = int(res.Type) if res.Type is not None else 0
    res_type = RESOURCE_TYPE_NAMES.get(res_type_code, "Work")
    if res_type == "Work":
        if material_label is not None:
            return {"status": "error",
                    "error": f"material_label not applicable to Work resource (type={res_type})"}
    elif res_type == "Material":
        if max_units is not None:
            return {"status": "error",
                    "error": f"max_units not applicable to Material resource (type={res_type})"}
        if overtime_rate is not None:
            return {"status": "error",
                    "error": f"overtime_rate not applicable to Material resource (type={res_type})"}
    elif res_type == "Cost":
        if max_units is not None:
            return {"status": "error",
                    "error": f"max_units not applicable to Cost resource (type={res_type})"}
        if overtime_rate is not None:
            return {"status": "error",
                    "error": f"overtime_rate not applicable to Cost resource (type={res_type})"}
        if material_label is not None:
            return {"status": "error",
                    "error": f"material_label not applicable to Cost resource (type={res_type})"}

    # Empty-changes no-op intentionally returns ok (mirrors T20 calendar update);
    # caller can detect via empty `changes` list.
    changes = []
    try:
        if do_rename:
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


def _msp_resource_delete(resource_id: int) -> Dict[str, Any]:
    """Delete a resource by ID. Cascades: any assignments to this resource
    are silently removed by MS Project; the count is returned in the response.
    """
    app = _validate_active_project()
    res = _find_resource_by_id(app.ActiveProject, resource_id)
    if res is None:
        return {"status": "error", "error": f"Resource ID {resource_id} not found"}
    try:
        name = res.Name
        # Capture cascade-affected assignment count BEFORE delete
        try:
            assignments_removed = int(res.Assignments.Count)
        except Exception:
            assignments_removed = 0
        res.Delete()
        return {"status": "ok", "deleted_id": resource_id,
                "deleted_name": name, "assignments_removed": assignments_removed}
    except Exception as e:
        logger.error(f"_msp_resource_delete({resource_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_resource_list() -> Dict[str, Any]:
    """List all resources in the active project with type-aware properties + assignment counts.

    Order: matches `proj.Resources` enumeration (typically insertion order, not sorted).
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    out = []
    try:
        for i in range(1, proj.Resources.Count + 1):
            res = proj.Resources(i)
            if res is None:
                continue
            entry = _serialize_resource(res)
            try:
                entry["assignment_count"] = res.Assignments.Count
            except Exception:
                entry["assignment_count"] = 0
            out.append(entry)
        return {"status": "ok", "count": len(out), "resources": out}
    except Exception as e:
        logger.error(f"_msp_resource_list failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_resource_assign(task_id: int, resource_id: int,
                        units: Optional[float] = None,
                        work_hours: Optional[float] = None) -> Dict[str, Any]:
    """Assign a resource to a task. Units in % (100 = full-time, default).

    Uses MS Project's `task.Assignments.Add(TaskID, ResourceID, [Units])`.
    work_hours overrides the auto-calculated work duration.
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    # Pre-flight validation
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    res = _find_resource_by_id(proj, resource_id)
    if res is None:
        return {"status": "error", "error": f"Resource ID {resource_id} not found"}
    if units is not None and units < 0:
        return {"status": "error", "error": "units must be >= 0"}
    if work_hours is not None and work_hours < 0:
        return {"status": "error", "error": "work_hours must be >= 0"}
    try:
        applied_units = units if units is not None else 100.0
        # task.Assignments.Add(TaskID, ResourceID, Units) — Units as fraction (1.0=100%)
        alloc = t.Assignments.Add(TaskID=task_id, ResourceID=resource_id,
                                 Units=applied_units / 100.0)
        warnings = []
        if work_hours is not None:
            try:
                # COM Work property accepts minutes
                alloc.Work = work_hours * 60.0
            except Exception as wh_err:
                warnings.append(f"work_hours not applied: {_format_com_error(wh_err)}")
        result = {"status": "ok",
                  "assignment_uid": alloc.UniqueID if alloc else None,
                  "task_id": task_id,
                  "resource_id": resource_id,
                  "units": applied_units}
        if warnings:
            result["warnings"] = warnings
        return result
    except Exception as e:
        logger.error(f"_msp_resource_assign({task_id},{resource_id}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_resource_assign_unsafe(task_obj: Any, res_obj: Any,
                               task_id: int, resource_id: int,
                               units: Optional[float] = None) -> Dict[str, Any]:
    """Internal fast-path for bulk: caller pre-resolved task_obj + res_obj.

    Skips _validate_active_project + _find_task_by_id + _find_resource_by_id —
    only safe when caller has just pre-built ID->object maps. NOT for public use.
    """
    try:
        applied_units = units if units is not None else 100.0
        if applied_units < 0:
            return {"status": "error", "error": "units must be >= 0",
                    "task_id": task_id, "resource_id": resource_id}
        # task.Assignments.Add is the only available API — proj.Assignments
        # does not exist on _IProjectDoc (probed 2026-04-28: AttributeError).
        # Measured ~9.6ms/call; not improvable via project-scoped collection.
        alloc = task_obj.Assignments.Add(TaskID=task_id, ResourceID=resource_id,
                                         Units=applied_units / 100.0)
        return {"status": "ok",
                "assignment_uid": alloc.UniqueID if alloc else None,
                "task_id": task_id,
                "resource_id": resource_id,
                "units": applied_units}
    except Exception as e:
        # Don't log per-call in bulk path — caller aggregates
        return {"status": "error",
                "task_id": task_id, "resource_id": resource_id,
                "error": _format_com_error(e)}


def _build_resource_id_map(proj: Any) -> Dict[int, Any]:
    """Pre-build resource_id -> Resource COM object map. O(N) one-time scan.

    Used by bulk_assign to avoid O(N×M) per-item lookup blow-up.
    """
    out: Dict[int, Any] = {}
    for i in range(1, proj.Resources.Count + 1):
        r = proj.Resources(i)
        if r is not None:
            out[r.ID] = r
    return out


def _build_task_id_map(proj: Any) -> Dict[int, Any]:
    """Pre-build task_id -> Task COM object map. O(N) one-time scan.

    Used by bulk_assign to avoid O(N×M) per-item lookup blow-up.
    """
    out: Dict[int, Any] = {}
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t is not None:
            out[t.ID] = t
    return out


def _msp_resource_bulk_assign_loop(items: List[Dict[str, Any]], path_label: str,
                                   task_map: Dict[int, Any],
                                   res_map: Dict[int, Any]) -> Dict[str, Any]:
    """Inner loop using pre-built maps + _assign_unsafe fast-path.

    Returns aggregated result with status / path / count / assignments / failures.
    """
    added: List[Dict[str, Any]] = []
    failures: List[Dict[str, Any]] = []
    for item in items:
        tid = item.get("task_id")
        rid = item.get("resource_id")
        t_obj = task_map.get(tid)
        r_obj = res_map.get(rid)
        if t_obj is None:
            failures.append({**item, "error": f"task_id {tid} not found"})
            continue
        if r_obj is None:
            failures.append({**item, "error": f"resource_id {rid} not found"})
            continue
        result = _msp_resource_assign_unsafe(
            task_obj=t_obj, res_obj=r_obj,
            task_id=tid, resource_id=rid,
            units=item.get("units"),
        )
        if result.get("status") == "ok":
            added.append({"task_id": tid, "resource_id": rid,
                         "assignment_uid": result.get("assignment_uid")})
        else:
            failures.append({**item, "error": result.get("error", "unknown")})
    status = "ok" if not failures else ("partial" if added else "error")
    return {"status": status, "path": path_label, "count": len(added),
            "assignments": added, "failures": failures}


def _msp_resource_bulk_assign(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Hybrid bulk-assign: routes by item count (Phase 1 _route_operation pattern).

    Items: [{task_id, resource_id, [units]}, ...]
    Routing:
      - <=5 items   -> com_direct (no batch mode)
      - 6-19 items  -> com_batch  (batch mode + loop)
      - >=20 items  -> mspdi_bulk (Phase 2b: com_batch_fallback; true MSPDI merge = Phase 3+)

    Pre-builds task_id -> Task and resource_id -> Resource maps ONCE to avoid
    O(N×M) lookup blow-up on the HERO 14×200=2800 case.

    NOTE: ID->object maps are built fresh per call. Do not cache them across
    operations that mutate Tasks or Resources collections (add/delete) —
    map will be stale.

    Returns: {status, path, count, assignments, failures}
    """
    if not items:
        return {"status": "ok", "path": "noop", "count": 0,
                "assignments": [], "failures": []}
    app = _validate_active_project()
    proj = app.ActiveProject
    # Pre-build maps ONCE — avoids O(N×M) lookup
    task_map = _build_task_id_map(proj)
    res_map = _build_resource_id_map(proj)

    path = _route_operation(len(items))
    if path == "com_direct":
        return _msp_resource_bulk_assign_loop(items, "com_direct", task_map, res_map)
    elif path == "com_batch":
        _enter_batch_mode()
        try:
            return _msp_resource_bulk_assign_loop(items, "com_batch", task_map, res_map)
        finally:
            _exit_batch_mode()
    else:  # mspdi_bulk — Phase 2b uses com_batch_fallback (true MSPDI merge is Phase 3+)
        _enter_batch_mode()
        try:
            return _msp_resource_bulk_assign_loop(items, "mspdi_bulk", task_map, res_map)
        finally:
            _exit_batch_mode()


def _msp_resource_unassign(task_id: int, resource_id: int) -> Dict[str, Any]:
    """Remove the assignment of a resource from a task."""
    app = _validate_active_project()
    proj = app.ActiveProject
    t = _find_task_by_id(proj, task_id)
    if t is None:
        return {"status": "error", "error": f"Task ID {task_id} not found"}
    # Find the matching assignment by ResourceID
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


# ---------- BASELINE CONSTANTS ----------

# MSP supports 11 baseline slots: Baseline + Baseline1..Baseline10
BASELINE_NUMBERS = list(range(11))  # [0, 1, ..., 10]

# CRITICAL: app.BaselineSave's `Into` parameter uses OFFSET enum, not direct number.
# Baseline 0 → Into=0 (pjIntoBaseline); Baseline N (1-10) → Into=10+N (pjIntoBaselineN).
# (Verified from msproject_typelib.txt enum PjSaveBaselineTo.)
INTO_BASELINE_MAP = {n: (0 if n == 0 else 10 + n) for n in BASELINE_NUMBERS}

# Module-level session-level baseline name retention (TAIL #3).
# MSP COM doesn't natively persist baseline names, so we track them in-process
# for the save->list round-trip. Keys: (project_name, baseline_number).
# Values: name strings provided to _msp_baseline_save. Cleared on baseline
# clear / clear_all. Lost on Python process restart (best-effort metadata,
# matches the pattern established in T40).
_BASELINE_NAMES: Dict[Tuple[str, int], str] = {}


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

    MSP returns various sentinels for unsaved baselines — normalize to Python None:
      - 0, None, "" (falsy)
      - "NA" string (MS Project 16.0 verified — most common case)
      - datetime with year < 1980 (some COM versions)
    """
    try:
        result = proj.BaselineSavedDate(Baseline=baseline_number)
        # MSP returns various falsy values for unsaved (0, datetime(year=0), None, "")
        if not result:
            return None
        # MS Project 16.0 returns the literal string "NA" for unsaved baselines
        if isinstance(result, str) and result.strip().upper() in ("NA", "N/A", ""):
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


def _msp_dt_or_none(v: Any) -> Optional[str]:
    """Normalize MSP baseline date sentinels ('NA', 'N/A', empty) to None.

    Tasks added AFTER a baseline was saved return the literal string 'NA' for
    BaselineNStart / BaselineNFinish. 'NA' is truthy → naive ``str(v) if v``
    leaks the sentinel downstream where date math silently yields 0 drift.

    pywintypes.datetime instances become ISO strings via ``str()``.
    """
    if not v:
        return None
    if isinstance(v, str):
        s = v.strip().upper()
        if s in ("NA", "N/A", ""):
            return None
        return v  # Already a string from MSP, return as-is
    return str(v)  # pywintypes.datetime → ISO string


def _read_task_baseline(task: Any, baseline_number: int) -> Dict[str, Any]:
    """Read a task's baseline values for the given baseline slot.

    Returns dict with start, finish, duration_h, work_h, cost.
    Unsaved baseline yields None/0 fallbacks. Each property read is guarded
    so a single bad COM read doesn't kill compare iteration.
    """
    out: Dict[str, Any] = {}
    for field, key, transform in [
        ("Start", "start", _msp_dt_or_none),
        ("Finish", "finish", _msp_dt_or_none),
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

    # Phase 1: COM mutation (atomic — abort with error on failure)
    try:
        copy_from = 0  # PjSaveBaselineFrom.pjCopyCurrent
        into = _baseline_into_code(baseline_number)
        # All=True (whole project) or False (selected); MSP COM expects bool
        all_param = (scope == "all")
        app.BaselineSave(All=all_param, Copy=copy_from, Into=into,
                        RollupToSummaryTasks=roll_up_to_summary)
    except Exception as e:
        logger.error(f"_msp_baseline_save({baseline_number}) BaselineSave failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}

    # TAIL #3: retain user-provided name for session-level list lookup.
    # If `name` is None, leave any prior retention untouched (re-saving the
    # same slot without a new name preserves the previous label — cheap memo).
    if name:
        try:
            _BASELINE_NAMES[(proj.Name, baseline_number)] = name
        except Exception:
            pass  # proj.Name read failure shouldn't fail save

    # Phase 2: metadata read-back (best-effort — save already succeeded)
    result: Dict[str, Any] = {"status": "ok",
                              "baseline_number": baseline_number,
                              "name": name}
    try:
        saved_date = _baseline_saved_date(proj, baseline_number)
        task_count = proj.Tasks.Count
        # Aggregate baseline totals (skip summary tasks)
        total_dur_min, total_work_min, total_cost = 0.0, 0.0, 0.0
        for i in range(1, task_count + 1):
            t = proj.Tasks(i)
            if t is None or t.Summary:
                continue
            data = _read_task_baseline(t, baseline_number)
            total_dur_min += (data["duration_h"] * 60) if data["duration_h"] else 0
            total_work_min += (data["work_h"] * 60) if data["work_h"] else 0
            total_cost += data["cost"] if data["cost"] else 0
        # 8h/day default — TODO: read from project calendar HoursPerDay (Phase 3b parity for CAU 9h/day)
        result.update({
            "saved_date": str(saved_date) if saved_date else None,
            "task_count": task_count,
            "total_duration_days": round(total_dur_min / 60 / 8, 2),
            "total_work_hours": round(total_work_min / 60, 2),
            "total_cost": round(total_cost, 2),  # Fix 3: rounded for caller-contract symmetry
        })
    except Exception as e:
        logger.warning(f"_msp_baseline_save({baseline_number}) metadata read-back failed (save itself succeeded): {e}")
        result["warning"] = f"metadata read failed: {_format_com_error(e)}"
    return result


def _msp_baseline_clear(baseline_number: int = 0) -> Dict[str, Any]:
    """Clear a single baseline (0-10). Idempotent: no-op if already empty.

    NOTE: BaselineClear's `From` parameter uses the PjSaveBaselineTo offset
    enum (same as BaselineSave's `Into`), NOT the direct 0-10 number.
    Empirically verified on MS Project 16.0 — passing 2 only clears baseline 0.
    Map via _baseline_into_code (B0=0, B1=11, ..., B10=20).
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    was_saved = _baseline_saved_date(proj, baseline_number)
    try:
        from_code = _baseline_into_code(baseline_number)
        app.BaselineClear(All=True, From=from_code)
        # TAIL #3: evict retained name (if any) for this slot.
        try:
            _BASELINE_NAMES.pop((proj.Name, baseline_number), None)
        except Exception:
            pass
        return {"status": "ok",
                "baseline_number": baseline_number,
                "was_saved_date": str(was_saved) if was_saved else None}
    except Exception as e:
        logger.error(f"_msp_baseline_clear({baseline_number}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_baseline_clear_all() -> Dict[str, Any]:
    """Clear all 11 baselines that are currently saved. Returns list of cleared numbers.

    Uses the PjSaveBaselineTo offset enum for `From` (see _msp_baseline_clear note).
    Skips already-unsaved baselines (loop optimization, fewer COM calls).
    """
    app = _validate_active_project()
    proj = app.ActiveProject
    cleared = []
    failures = []
    for n in BASELINE_NUMBERS:
        if _baseline_saved_date(proj, n) is None:
            continue
        try:
            from_code = _baseline_into_code(n)
            app.BaselineClear(All=True, From=from_code)
            cleared.append(n)
        except Exception as e:
            failures.append({"baseline_number": n, "error": _format_com_error(e)})
    # TAIL #3: evict retained names for ALL slots of this project.
    try:
        proj_name = proj.Name
        for key in [k for k in _BASELINE_NAMES if k[0] == proj_name]:
            _BASELINE_NAMES.pop(key, None)
    except Exception:
        pass
    return {"status": "ok" if not failures else "partial",
            "cleared": cleared,
            "count": len(cleared),
            "failures": failures}


def _msp_baseline_list() -> Dict[str, Any]:
    """List all 11 baseline slots; return only those currently saved with metadata.

    Iterates all 11 slots, checks saved date, includes task count + total stats
    only for saved ones. Returns sorted by baseline number (BASELINE_NUMBERS order).
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
                total_dur_min += (data["duration_h"] * 60) if data["duration_h"] else 0
                total_work_min += (data["work_h"] * 60) if data["work_h"] else 0
                total_cost += data["cost"] if data["cost"] else 0
            out.append({
                "number": n,
                # TAIL #3: session-level retention from _msp_baseline_save's
                # `name` arg. None when no name was ever supplied (or after
                # clear/clear_all/process restart).
                "name": _BASELINE_NAMES.get((proj.Name, n)),
                "saved_date": str(saved),
                "task_count": task_ct,
                "total_duration_days": round(total_dur_min / 60 / 8, 2),
                "total_work_hours": round(total_work_min / 60, 2),
                "total_cost": round(total_cost, 2),
            })
        return {"status": "ok", "count_saved": len(out), "baselines": out}
    except Exception as e:
        logger.error(f"_msp_baseline_list failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


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


def _datetime_diff_days(current_str: Optional[str], baseline_str: Optional[str]) -> float:
    """Compute calendar-day difference between two ISO datetime strings.
    Returns 0 if either is None/missing or unparseable (e.g. MSP 'NA' sentinel)."""
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


# Module-level constant for floating-point comparison (TAIL #4)
_VARIANCE_EPSILON = 1e-9  # ~picoseconds in days, ~picohours, ~picocents


def _read_task_current_state(task: Any) -> Dict[str, Any]:
    """Read live task fields in the same shape as `_read_task_baseline`.

    Adapter so `_compute_variance_set` can treat current state and baseline
    state uniformly. Uses `_msp_dt_or_none` for Start/Finish so 'NA' sentinels
    (impossible on live tasks but cheap defensive guard) become None.
    """
    return {
        "start": _msp_dt_or_none(task.Start) if task.Start else None,
        "finish": _msp_dt_or_none(task.Finish) if task.Finish else None,
        "duration_h": float(task.Duration) / 60.0 if task.Duration else 0.0,
        "work_h": float(task.Work) / 60.0 if task.Work else 0.0,
        "cost": _parse_rate(task.Cost) if task.Cost else 0.0,
    }


def _compute_variance_set(
    proj: Any,
    get_a: Callable[[Any], Dict[str, Any]],
    get_b: Callable[[Any], Dict[str, Any]],
    include_unchanged: bool,
    variance_threshold_days: float,
) -> Dict[str, Any]:
    """Compute per-task variance summary + list given two side-readers.

    Variance direction: `b - a` (positive = side B exceeds side A).
    For `_msp_baseline_compare`: a = saved baseline, b = current state →
    positive finish_var = slipped.
    For `_msp_baseline_compare_two`: a = baseline_a (e.g. Original),
    b = baseline_b (e.g. Revised) → positive finish_var = slip between
    revisions.

    Pre-builds the non-summary task list once (TAIL #1 perf) and uses
    EPSILON-based no_change check (TAIL #4 robustness).

    Returns {"summary": {...8 keys...}, "tasks": [...]}; caller wraps with
    status/baseline_number etc.
    """
    # Pre-build task cache ONCE (TAIL #1) — cuts ~50% COM dispatches on Summary check
    real_tasks: List[Any] = []
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t is not None and not t.Summary:
            real_tasks.append(t)

    tasks_var: List[Dict[str, Any]] = []
    slipped_ct = ahead_ct = on_time_ct = 0
    total_start_drift = total_finish_drift = 0.0
    total_dur_var_h = total_work_var_h = 0.0
    total_cost_var = 0.0

    for t in real_tasks:
        a = get_a(t)
        b = get_b(t)
        start_var = _datetime_diff_days(b["start"], a["start"])
        finish_var = _datetime_diff_days(b["finish"], a["finish"])
        dur_var = (b["duration_h"] or 0) - (a["duration_h"] or 0)
        work_var = (b["work_h"] or 0) - (a["work_h"] or 0)
        cost_var = (b["cost"] or 0) - (a["cost"] or 0)

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

        # EPSILON-based no_change (TAIL #4) — robust to float residue
        no_change = (
            abs(start_var) < _VARIANCE_EPSILON
            and abs(finish_var) < _VARIANCE_EPSILON
            and abs(dur_var) < _VARIANCE_EPSILON
            and abs(work_var) < _VARIANCE_EPSILON
            and abs(cost_var) < _VARIANCE_EPSILON
        )
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
    try:
        result = _compute_variance_set(
            proj,
            get_a=lambda t: _read_task_baseline(t, baseline_number),
            get_b=_read_task_current_state,
            include_unchanged=include_unchanged,
            variance_threshold_days=variance_threshold_days,
        )
        return {
            "status": "ok",
            "baseline_number": baseline_number,
            **result,
        }
    except Exception as e:
        logger.error(f"_msp_baseline_compare({baseline_number}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


def _msp_baseline_compare_two(baseline_a: int,
                              baseline_b: int,
                              include_unchanged: bool = False,
                              variance_threshold_days: float = 0.0) -> Dict[str, Any]:
    """Compare two saved baselines as a delta (baseline_b - baseline_a).

    Variance is computed as (baseline_b - baseline_a): the change FROM
    baseline_a TO baseline_b. Use case: baseline_a=0 (Original Plan),
    baseline_b=1 (Revised Plan) → positive finish_var_days = task slipped
    between revisions; negative = task pulled ahead between revisions.

    Both baselines must be saved (have a baseline_save_date). Per-task fields
    with None on either side (e.g. tasks added after a save) yield 0 from
    _datetime_diff_days, matching the compare-action contract.

    Args:
        baseline_a: 0-10, the "earlier" baseline slot to subtract.
        baseline_b: 0-10, the "later" baseline slot.
        include_unchanged: if False, omit zero-variance tasks from the list.
        variance_threshold_days: tasks with |finish_var_days| <= threshold
            count as on_time.
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
    try:
        result = _compute_variance_set(
            proj,
            get_a=lambda t: _read_task_baseline(t, baseline_a),
            get_b=lambda t: _read_task_baseline(t, baseline_b),
            include_unchanged=include_unchanged,
            variance_threshold_days=variance_threshold_days,
        )
        return {
            "status": "ok",
            "baseline_a": baseline_a,
            "baseline_b": baseline_b,
            **result,
        }
    except Exception as e:
        logger.error(f"_msp_baseline_compare_two({baseline_a},{baseline_b}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


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


def _msp_baseline_set_active(baseline_number: int) -> Dict[str, Any]:
    """Set the active baseline for views/EVM calculations.

    The "active baseline" controls which Baseline*N* fields the Earned Value
    engine and baseline-aware views use. T47 probe (MSP 16.0) confirmed that
    proj.EarnedValueBaseline is the real read/write property — round-trips
    exactly. Other candidates (BaselineForEarnedValue, ActiveBaselineNumber)
    silently swallow setattr without persisting on this MSP build.

    NOTE: MSP COM API for this is version-dependent. Older MSP versions may
    not expose EarnedValueBaseline; in that case the function returns a
    graceful "not yet supported" error — saved baseline data is still readable
    via get_task_baseline / compare regardless of which is "active".
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    app = _validate_active_project()
    proj = app.ActiveProject
    # Probe-confirmed property: proj.EarnedValueBaseline (read/write int, round-trips).
    try:
        proj.EarnedValueBaseline = baseline_number
        readback = proj.EarnedValueBaseline
        if readback == baseline_number:
            return {"status": "ok", "active_baseline": baseline_number,
                    "method": "proj.EarnedValueBaseline"}
    except Exception as e:
        logger.debug(f"_msp_baseline_set_active: EarnedValueBaseline failed: {e}")
    return {"status": "error",
            "error": ("set_active is not yet supported on this MS Project version. "
                     "Use msproject_baseline compare/summary directly with the "
                     "baseline_number parameter — they don't require setting an active baseline.")}


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
    "day": 4,    # pjTimescaleDays (probe-confirmed MSP 2024 — T60)
    "week": 3,   # pjTimescaleWeeks (probe-confirmed MSP 2024 — T60)
}

_PJ_TIMESCALED_ACTUAL_WORK = 10  # pjAssignmentTimescaledActualWork (probe-confirmed MSP 2024 — T60)


# ---------- PROGRESS HELPERS ----------

def _normalize_progress_pct(v: Any) -> float:
    """Validate + normalize a percentage value (0-100).

    Accepts int / float / str ('50', '50.5', '50%'). Raises ValueError on
    out-of-range or non-numeric input. Returns float rounded to 2 decimals.

    Note: ``bool`` is rejected explicitly. In Python ``bool`` is a subclass of
    ``int`` so ``True`` would otherwise coerce to 1.0 and ``False`` to 0.0 —
    JSON inputs like ``{"physical_pct": true}`` would silently produce 1%
    progress. This guard catches that class of caller bug at the boundary.
    """
    if v is None:
        raise ValueError("progress percentage cannot be None")
    if isinstance(v, bool):
        raise ValueError(
            f"progress percentage must be numeric, not bool: {v!r}"
        )
    if isinstance(v, str):
        s = v.strip().rstrip("%").strip()
        if not s:
            raise ValueError(f"progress percentage is empty string: {v!r}")
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

    Returns ``None`` if valid, or an error message string if invalid.

    Date format: ISO 8601 strict (``datetime.fromisoformat``-compatible).
    Falls back to ``dateutil.parser.parse`` with ``dayfirst=False`` for
    non-ISO inputs (e.g. the ``str()`` form of ``pywintypes.datetime``).

    Format assumption: ambiguous slash-separated dates like ``01/02/2026`` are
    interpreted as **US (Jan 2)**, never EU. Public callers should always pass
    ISO 8601 to avoid surprise.
    """
    if start is None or finish is None:
        return None

    def _parse(s: Any) -> "_dt.datetime":
        s = str(s).strip()
        if not s:
            raise ValueError("empty date string")
        try:
            return _dt.datetime.fromisoformat(
                s.replace("+00:00", "").rstrip("Z")
            )
        except Exception:
            from dateutil import parser
            return parser.parse(s, dayfirst=False)

    try:
        s_dt = _parse(start)
        f_dt = _parse(finish)
    except Exception as e:
        return f"could not parse dates ({start!r}, {finish!r}): {e}"
    if s_dt > f_dt:
        return f"actual_start ({start}) must be <= actual_finish ({finish})"
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
    # Defaults are 0.0 (float) for type stability — successful reads also use
    # float(v), so JSON output remains uniform whether the read succeeded or
    # fell back to the default.
    pct_pairs = [
        ("PercentComplete", "percent_complete", 0.0),
        ("PercentWorkComplete", "percent_work_complete", 0.0),
        ("PhysicalPercentComplete", "physical_pct", 0.0),
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


def _to_pywintypes_date(v: Any) -> Any:
    """Convert public-API date input to MSP COM-compatible pywintypes.Time.

    MSP COM rejects raw strings on date properties (``task.ActualStart = "..."``
    raises an opaque COM error). Public callers JSON-encode dates as ISO
    strings, so this helper bridges the two.

    Accepts:
      * ``None`` → returns ``None`` (caller passes through to clear / no-op)
      * ``pywintypes.datetime`` / ``pywintypes.TimeType`` → returned as-is
      * ``datetime.datetime`` → wrapped via ``pywintypes.Time``
      * ISO 8601 string (``2026-04-15`` or ``2026-04-15 08:00:00``) parsed
        with ``datetime.fromisoformat``; trailing ``Z`` and ``+00:00`` stripped
      * Non-ISO strings → fall back to ``dateutil.parser.parse`` with
        ``dayfirst=False`` so ambiguous ``01/02/2026`` is deterministic (Jan 2)

    Raises ``ValueError`` on empty string, unparseable string, or unsupported
    type.
    """
    if v is None:
        return None
    # pywintypes.datetime / pywintypes.Time → pass through unchanged
    try:
        import pywintypes
        if isinstance(v, pywintypes.TimeType):
            return v
    except ImportError:
        pass
    # datetime.datetime → wrap in pywintypes.Time
    if isinstance(v, _dt.datetime):
        import pywintypes
        return pywintypes.Time(v)
    # ISO 8601 string (preferred)
    if isinstance(v, str):
        s = v.strip()
        if not s:
            raise ValueError("date input is empty string")
        try:
            dt = _dt.datetime.fromisoformat(
                s.replace("+00:00", "").rstrip("Z")
            )
        except Exception:
            try:
                from dateutil import parser
                dt = parser.parse(s, dayfirst=False)
            except Exception as e:
                raise ValueError(f"could not parse date {v!r}: {e}")
        import pywintypes
        return pywintypes.Time(dt)
    raise ValueError(f"unsupported date type: {type(v).__name__} ({v!r})")


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
        task.ActualStart = _to_pywintypes_date(value)
    elif field == "actual_finish":
        task.ActualFinish = _to_pywintypes_date(value)
    elif field == "actual_duration_h":
        task.ActualDuration = _hours_to_minutes(value)
    elif field == "actual_work_h":
        task.ActualWork = _hours_to_minutes(value)
    elif field == "remaining_work_h":
        task.RemainingWork = _hours_to_minutes(value)
    elif field == "remaining_duration_h":
        task.RemainingDuration = _hours_to_minutes(value)
    elif field == "stop":
        task.Stop = _to_pywintypes_date(value)
    elif field == "resume":
        task.Resume = _to_pywintypes_date(value)
    else:
        raise ValueError(f"Unknown progress field: {field}")


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
                asg.ActualStart = _to_pywintypes_date(value)  # T52 fix integration
            elif field == "actual_finish":
                asg.ActualFinish = _to_pywintypes_date(value)  # T52 fix integration
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
            logger.debug(
                f"set_assignment_progress({task_id},{resource_id},{field}={value}) failed: {e}")

    if not changes and failures:
        return {"status": "error", "task_id": task_id, "resource_id": resource_id,
                "error": "all assignment field writes failed", "failures": failures}
    status = "ok" if not failures else "partial"
    return {"status": status, "task_id": task_id, "resource_id": resource_id,
            "changes": changes, "failures": failures}


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


def _msp_progress_set_by_date(progress_date: Any,
                              scope: str = "all",
                              as_scheduled: bool = True) -> Dict[str, Any]:
    """Bulk-update progress to a given date (``app.UpdateProject``).

    Implements the "plan = actual" up-to-data-date assumption — fast retroactive
    backlog catch-up. Phase 3b — see design doc Section 6 (Q1) and Open
    Questions #2 for already-progressed task interaction.

    Probe-confirmed signature on MSP 16.0::

        UpdateProject(All, UpdateDate, action)

    where ``All`` (bool): True = all tasks / False = currently-selected only;
    ``UpdateDate`` (datetime.datetime; **not** pywintypes.Time — COM rejects);
    ``action`` (int): 1 = update progress as scheduled (default), 0 =
    reschedule incomplete only (no progress write), 2 = no-op.

    Args:
        progress_date: ISO 8601 string or ``datetime.datetime``. Empty / unparseable
                       string returns ``{status: error}``.
        scope: ``"all"`` (entire project) or ``"selected"`` (currently selected
               tasks in MSP UI).
        as_scheduled: True → action=1 (write progress as scheduled up to date);
                      False → action=0 (reschedule incomplete portion only).
                      Note: MSP COM ``UpdateProject`` does not expose a pure
                      "% complete only" mode. ``as_scheduled=False`` here means
                      reschedule-incomplete (lighter touch — no actuals written).

    Returns:
        ``{status, progress_date, mode, scope, task_count_affected}`` on success;
        ``{status: error, error}`` on parse / COM failure.
    """
    if scope not in ("all", "selected"):
        return {"status": "error",
                "error": f"scope must be 'all' or 'selected', got '{scope}'"}
    # Parse progress_date → plain datetime.datetime (NOT pywintypes — COM rejects)
    try:
        if isinstance(progress_date, _dt.datetime):
            pd = progress_date
        elif isinstance(progress_date, str):
            s = progress_date.strip()
            if not s:
                return {"status": "error",
                        "error": "progress_date is empty string"}
            try:
                pd = _dt.datetime.fromisoformat(
                    s.replace("+00:00", "").rstrip("Z")
                )
            except Exception:
                from dateutil import parser
                pd = parser.parse(s, dayfirst=False)
        else:
            return {"status": "error",
                    "error": f"unsupported progress_date type: "
                             f"{type(progress_date).__name__}"}
    except Exception as e:
        return {"status": "error",
                "error": f"could not parse progress_date {progress_date!r}: {e}"}

    app = _validate_active_project()
    proj = app.ActiveProject
    task_count_before = proj.Tasks.Count
    try:
        all_tasks_flag = (scope == "all")
        action_val = 1 if as_scheduled else 0  # 1 = update as scheduled; 0 = reschedule incomplete only
        # Probe-confirmed signature on MSP 16.0: positional only, plain datetime.
        app.UpdateProject(all_tasks_flag, pd, action_val)
        return {"status": "ok",
                "progress_date": pd.isoformat(),
                "mode": "as_scheduled" if as_scheduled else "reschedule_incomplete",
                "scope": scope,
                "task_count_affected": task_count_before}
    except Exception as e:
        logger.error(f"_msp_progress_set_by_date({progress_date}) failed: {e}")
        return {"status": "error", "error": _format_com_error(e)}


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
        # ActualStart / ActualFinish: write "NA" sentinel to clear (MSP convention)
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


def _msp_progress_time_phased_write(task_id: int,
                                    resource_id: int,
                                    periods: List[Dict[str, Any]],
                                    unit: str = "day") -> Dict[str, Any]:
    """Write per-period actual_work to an assignment via TimeScaleData.

    Phase 3b T60 — see design doc Section 6 Q2. Granularity: 'day' or 'week'.

    periods: [{start: ISO, end: ISO, actual_work_h: float}, ...]
    Each period maps to one (or more) TimeScaleValues slots; write fails per-
    slot if MSP doesn't have a matching cell. Failures aggregate into
    return['failures'] without raising. COM signature uses positional args
    (StartDate, EndDate, Type, TimeScaleUnit) — note capital S in TimeScaleUnit
    per probe-confirmed signature on MSP 2024.
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
            # Positional call: (StartDate, EndDate, Type, TimeScaleUnit)
            tsv = asg.TimeScaleData(ps, pe,
                                    _PJ_TIMESCALED_ACTUAL_WORK,
                                    unit_code)
            if tsv.Count == 0:
                failures.append({"index": idx, "period": p,
                                 "error": "no time slots in range (assignment "
                                          "may not span this date)"})
                continue
            # Filter slots to those that fall WITHIN the requested period —
            # MSP returns boundary slots (e.g. the day before a 1-day query)
            # that should be skipped to honor caller intent.
            target_indices: List[int] = []
            for i in range(1, tsv.Count + 1):
                try:
                    item = tsv.Item(i)
                    item_start = item.StartDate
                    # Compare naive vs tz-aware safely by stripping tz
                    if hasattr(item_start, "tzinfo") and item_start.tzinfo is not None:
                        item_start_cmp = item_start.replace(tzinfo=None)
                    else:
                        item_start_cmp = item_start
                    ps_cmp = ps.replace(tzinfo=None) if ps.tzinfo else ps
                    pe_cmp = pe.replace(tzinfo=None) if pe.tzinfo else pe
                    if item_start_cmp >= ps_cmp and item_start_cmp < pe_cmp:
                        target_indices.append(i)
                except Exception:
                    continue
            # Fallback: if filter excluded everything, write to all slots so
            # something lands (out-of-range edge case).
            if not target_indices:
                target_indices = list(range(1, tsv.Count + 1))
            # Distribute hours across selected slots evenly
            minutes_total = _hours_to_minutes(hours)
            slot_count = len(target_indices)
            per_slot = minutes_total // slot_count
            remainder = minutes_total - (per_slot * slot_count)
            slot_failures = 0
            for n, i in enumerate(target_indices, start=1):
                try:
                    val = per_slot + (remainder if n == slot_count else 0)
                    tsv.Item(i).Value = val
                except Exception as e:
                    slot_failures += 1
                    failures.append({"index": idx, "slot": i,
                                     "error": _format_com_error(e)})
            # Only count as written if at least one slot succeeded
            if slot_failures < slot_count:
                written += 1
        except Exception as e:
            failures.append({"index": idx, "period": p,
                             "error": _format_com_error(e)})
    status = "ok" if not failures else ("partial" if written else "error")
    return {"status": status,
            "task_id": task_id, "resource_id": resource_id,
            "unit": unit,
            "written_count": written, "failures": failures}


def _msp_progress_time_phased_read(task_id: int,
                                   resource_id: int,
                                   start_date: str,
                                   end_date: str,
                                   unit: str = "day") -> Dict[str, Any]:
    """Read per-period actual_work from an assignment via TimeScaleData.

    Phase 3b T61. Returns {status, periods: [{period_start, period_end,
    actual_work_h}]}. Empty/blank slot values (no actual yet) return as 0.0
    hours, not omitted. COM signature uses positional args (StartDate,
    EndDate, Type, TimeScaleUnit) per probe-confirmed signature on MSP 2024.
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
        # Positional call: (StartDate, EndDate, Type, TimeScaleUnit)
        tsv = asg.TimeScaleData(ds, de,
                                _PJ_TIMESCALED_ACTUAL_WORK,
                                unit_code)
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

    # Pre-compute expected Duration per item (in MSP minutes) for post-paste
    # fix-up. MSP 16.0 silently drops <Duration> from MSPDI FileOpen imports
    # (Phase 2b TAIL — verified by probe across multiple XML variants:
    # adding <Manual>0</Manual>, full Calendar with WorkingTimes,
    # <MinutesPerDay>/<MinutesPerWeek> project settings, <Start>/<Finish>
    # — none restore Duration). Workaround: post-paste, re-set t.Duration
    # via COM in batch mode (~5 ms/task, ≈1s for 200 tasks).
    _expected_durations_min = [
        _parse_duration(item.get("duration", "1d")) for item in items
    ]

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

        # Collect IDs of newly added tasks AND restore Duration (MSPDI import
        # drops it — see workaround note above). The collection loop iterates
        # only valid (non-None) rows, in items order, so duration_idx maps
        # 1:1 to items[].
        target_proj = app.ActiveProject
        added_task_ids = []
        duration_set_failures = 0
        duration_idx = 0
        for i in range(last_count + 1, target_proj.Tasks.Count + 1):
            t = target_proj.Tasks(i)
            if t is None:
                continue
            if duration_idx < len(_expected_durations_min):
                try:
                    t.Duration = _expected_durations_min[duration_idx]
                except Exception as e:
                    duration_set_failures += 1
                    logger.debug(
                        f"post-paste Duration set failed at row {i} "
                        f"(idx={duration_idx}): {e}"
                    )
            added_task_ids.append(t.ID)
            duration_idx += 1

        return {
            "status": "ok",
            "path": "mspdi_bulk",
            "count": len(added_task_ids),
            "task_ids": added_task_ids,
            "method": "FileOpen + EditCopy + EditPaste + post-paste Duration set",
            "duration_set_failures": duration_set_failures,
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
    - bulk_assign: Hybrid (1-5 COM, 6-19 batch, 20+ MSPDI fallback). Params: items=[{task_id, resource_id, [units]}, ...]

    Phase 2b (28 Apr 2026). Note: bulk_assign perf is ~10ms/call due to MS Project
    COM intrinsic limit; true MSPDI assignment merge is Phase 3+.
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


# ============================================================================
# PHASE 4 - FILE MCP (msproject_file)
# ============================================================================
# Phase 4 (T65 foundations) — adds file-mode helpers separate from the COM
# code above. Phase 1-3 helpers DO NOT reference anything below this line.
# Imports/state added here:
#   * MspdiProject (native MSPDI parser, reused from Asta side)
#   * _jvm_started bool (lazy MPXJ JVM lifecycle)
#   * _detect_msp_xml_schema, _get_msp_file_manager, MspMppFileManager
# ----------------------------------------------------------------------------

# Native MSPDI parser (zero Java dependency, reuse from Asta).
from mspdi_parser import MspdiProject

# JVM pre-start for MPXJ (lazy — only if .mpp encountered).
_jvm_started = False


def _ensure_jvm_started() -> None:
    """Start JVM lazily on first MPP request. Idempotent.

    MPXJ requires JPype1+JVM to read .mpp binaries. We pay this cost only
    when the user actually opens a .mpp file. Re-invocation is a no-op.
    """
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

    Returns True if MS Project XML (schemas.microsoft.com/project), False
    if Asta or unknown schema. Raises FileNotFoundError if path missing.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    with open(file_path, 'rb') as f:
        head = f.read(512).decode('utf-8', errors='replace')
    return 'schemas.microsoft.com/project' in head


def _mpxj_duration_to_hours(d) -> float:
    """Convert MPXJ Duration object to hours (float). Handles None safely."""
    if d is None:
        return 0.0
    try:
        from org.mpxj import TimeUnit
        n = float(d.getDuration())
        unit = d.getUnits()
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
    except Exception as e:
        logger.warning(f"_mpxj_duration_to_hours failed (returning 0.0): {e}")
        return 0.0


class MspMppFileManager:
    """Read-only manager for .mpp files via MPXJ + JVM.

    Adapted from Asta asta_mcp_file.py AstaFileManager (drops .pp support,
    MPP only). MPP write is not supported by MPXJ — Phase 4 intentionally
    keeps this class read-only.

    Lazy load: __init__ does NOT touch MPXJ/JVM. The first read_*() call
    triggers JVM start + UniversalProjectReader. This makes init cheap
    (and importable on systems without Java).
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
        out: List[Dict[str, Any]] = []
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
        # T-future: cache predecessor lookups if MPP perf becomes an issue
        # (per-task Java->Python crossings are O(N*M); MPP read is cold-path
        # for typical Phase 4 use, most users hit XML)
        proj = self._load()
        out: List[Dict[str, Any]] = []
        for t in proj.getTasks():
            if t is None or t.getID() == 0:
                continue
            for rel in (t.getPredecessors() or []):
                lag = rel.getLag()
                out.append({
                    "from_id": int(rel.getTargetTask().getID()),
                    "to_id": int(t.getID()),
                    "type": str(rel.getType()),
                    "lag_days": _mpxj_duration_to_hours(lag) / 8.0 if lag else 0.0,
                })
        return out

    def read_resources(self) -> List[Dict[str, Any]]:
        proj = self._load()
        out: List[Dict[str, Any]] = []
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
        out: List[Dict[str, Any]] = []
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
        out: List[Dict[str, Any]] = []
        for cal in proj.getCalendars():
            out.append({
                "name": str(cal.getName() or ""),
                "is_base": bool(cal.getParent() is None),
            })
        return out

    def read_baselines(self, baseline_number: int = 0) -> Dict[str, Any]:
        return {
            "baseline_number": baseline_number,
            "saved_date": None,
            "tasks": [],
            "note": "MPP baseline read via MPXJ - limited fields",
        }

    def read_progress(self) -> Dict[str, Any]:
        proj = self._load()
        tasks: List[Dict[str, Any]] = []
        for t in proj.getTasks():
            if t is None or t.getID() == 0 or t.getSummary():
                continue
            tasks.append({
                "id": int(t.getID()),
                "percent_complete": float(t.getPercentageComplete() or 0),
                "actual_work_h": _mpxj_duration_to_hours(t.getActualWork()),
            })
        try:
            props = proj.getProjectProperties()
            status_date = props.getStatusDate() if props else None
        except Exception as e:
            logger.debug(f"MPP status_date read failed: {e}")
            status_date = None
        return {
            "status_date": str(status_date) if status_date else None,
            "tasks": tasks,
        }


def _get_msp_file_manager(file_path: str):
    """Factory: returns MspdiProject for .xml/.mspdi, MspMppFileManager for .mpp.

    Performs schema check for XML — refuses non-MSPDI XML with clear error
    pointing at asta_powerproject_file MCP for Asta files.

    Note on error precedence: extension is validated FIRST (cheap, no I/O)
    so callers passing an unsupported extension always see ValueError, even
    if the path doesn't exist. Existence check follows.
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ('.xml', '.mspdi', '.mpp'):
        raise ValueError(
            f"Unsupported extension '{ext}'. Phase 4 supports: .xml, .mspdi, .mpp"
        )
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    if ext in ('.xml', '.mspdi'):
        if not _detect_msp_xml_schema(file_path):
            raise ValueError(
                f"Not a MS Project XML - appears to be Asta or unknown schema. "
                f"For Asta files use asta_powerproject_file MCP. File: {file_path}"
            )
        return MspdiProject(file_path)
    # ext == '.mpp'
    return MspMppFileManager(file_path)


# ---------- PHASE 4 ACTION HELPERS ----------
#
# T66 PROBE FINDINGS (mspdi_parser.MspdiProject API):
#   * get_all_tasks() returns list of dicts with keys:
#       id, unique_id, name, duration (str "1d"), start, finish,
#       percent_complete, critical, milestone, summary, total_float,
#       notes, predecessors, successors
#   * predecessors / successors are lists of dicts:
#       {"task_id": int, "type": "FS"|"SS"|"FF"|"SF", "lag": "0d"}
#     (task_id resolved from UID by _id_by_uid; type/lag pre-formatted)
#   * NO get_links() method. get_link_chain(from_pat, to_pat) takes 2
#     regex strings - not suitable for "all links". Use task walk.
#   * MspMppFileManager (added T65) already returns the unified contract
#     {id, name, duration_h, start, finish, percent_complete, summary}
#     for tasks and {from_id, to_id, type, lag_days} for links.
#
# Adapter strategy: when manager is MspdiProject, normalize task dicts
# (rename unique_id, parse duration string -> hours) and walk
# predecessors lists to build the unified link list.


def _parse_duration_h(d) -> float:
    """Parse duration representations to hours.

    Accepts numeric (assume hours), ISO 8601 (PT8H0M0S), or
    Asta/MSPDI-style strings ('1d', '3w', '8h')."""
    if d is None:
        return 0.0
    if isinstance(d, (int, float)):
        return float(d)
    s = str(d).strip()
    if s.startswith('PT'):
        import re
        h = re.search(r'(\d+)H', s)
        m = re.search(r'(\d+)M', s)
        return (float(h.group(1)) if h else 0.0) + (float(m.group(1)) / 60 if m else 0.0)
    if s.endswith('d'):
        try:
            return float(s[:-1]) * 8.0
        except ValueError:
            return 0.0
    if s.endswith('w'):
        try:
            return float(s[:-1]) * 40.0
        except ValueError:
            return 0.0
    if s.endswith('h'):
        try:
            return float(s[:-1])
        except ValueError:
            return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _normalize_mspdi_task(raw: Dict[str, Any]) -> Dict[str, Any]:
    """Translate MspdiProject get_all_tasks() dict to unified Phase 4 task dict.

    Probe-confirmed keys: id, name, duration (str), start, finish,
    percent_complete, summary.
    """
    return {
        "id": raw["id"],
        "name": raw["name"],
        "duration_h": _parse_duration_h(raw.get("duration")),
        "start": raw.get("start"),
        "finish": raw.get("finish"),
        "percent_complete": float(raw.get("percent_complete", 0)),
        "summary": bool(raw.get("summary", False)),
    }


def _msp_file_read_tasks(file_path: str,
                         filters: Optional[Dict[str, Any]] = None,
                         limit: Optional[int] = None) -> Dict[str, Any]:
    """Read all tasks from a MS Project file. Format auto-detected by extension.

    Excludes summary tasks (root project + WBS summaries) for cleaner results.
    """
    try:
        mgr = _get_msp_file_manager(file_path)
        if isinstance(mgr, MspdiProject):
            raw_tasks = mgr.get_all_tasks()
            tasks = [_normalize_mspdi_task(t) for t in raw_tasks]
        else:
            tasks = mgr.read_tasks()
        # Filter out summary tasks (root + WBS)
        tasks = [t for t in tasks if not t.get("summary", False)]
        # Apply optional field filters
        if filters:
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


def _extract_links_from_mspdi(mgr: 'MspdiProject') -> List[Dict[str, Any]]:
    """Extract links from MspdiProject by walking tasks for predecessor info.

    Each task's predecessors list (probe-confirmed shape) is:
        [{"task_id": int, "type": "FS"|"SS"|"FF"|"SF", "lag": "0d"}, ...]
    where task_id is the predecessor's task ID. Returns unified contract
    {from_id, to_id, type, lag_days}.
    """
    links: List[Dict[str, Any]] = []
    raw_tasks = mgr.get_all_tasks()
    for t in raw_tasks:
        preds = t.get("predecessors") or []
        if not preds:
            continue
        to_id = t["id"]
        for p in preds:
            links.append({
                "from_id": p["task_id"],
                "to_id": to_id,
                "type": p.get("type", "FS"),
                "lag_days": _parse_duration_h(p.get("lag", "0d")) / 8.0,
            })
    return links


def _msp_file_read_links(file_path: str) -> Dict[str, Any]:
    """Read all task predecessor/successor links from a MS Project file.

    For MspdiProject (XML): walks tasks and extracts predecessor entries.
    For MspMppFileManager (MPP): delegates to its read_links().
    """
    try:
        mgr = _get_msp_file_manager(file_path)
        if isinstance(mgr, MspdiProject):
            links = _extract_links_from_mspdi(mgr)
        else:
            links = mgr.read_links()
        return {"status": "ok", "count": len(links), "links": links}
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.error(f"_msp_file_read_links({file_path}) failed: {e}")
        return {"status": "error", "error": str(e)}


def main():
    """Run MCP server (stdio)."""
    mcp.run()


if __name__ == "__main__":
    main()
