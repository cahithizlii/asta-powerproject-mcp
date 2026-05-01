"""Phase 5d - Pure-Python Primavera P6 XER reader.

XER format (text, typically UTF-16-LE with BOM, fallback UTF-8):
- ERMHDR <version>\\t<date>\\t<user>\\t<app>\\t<currency>
- %T <table_name>            : table marker
- %F <header1>\\t<header2>... : field names (column headers)
- %R <val1>\\t<val2>...       : data row (position-mapped to %F)
- %E                         : end of file

NO mpxj dependency. Tractable in ~400 lines pure Python.
"""
import logging
import os

logger = logging.getLogger(__name__)


class XerFile:
    """Parse a P6 XER file into structured table dicts.

    Public attributes:
        file_path: original file path string.
        header_fields: dict of ERMHDR positional fields (version/exported/user/app/currency).
        tables: dict {table_name: {"headers": [str], "rows": [{col: str}]}}.

    Public read methods (added below class body):
        read_tasks() -> [task dicts in MSP shape]
        read_links() -> [link dicts {from_id, to_id, type, lag_days}]
        read_resources() -> [resource dicts]
        read_assignments() -> [assignment dicts]
        read_calendars() -> [calendar dicts]
        read_progress() -> {status_date, tasks: [...]}
        read_project() -> {proj_id, plan_start_date, plan_end_date, ...}
    """

    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"XER file not found: {file_path}")
        self.file_path = file_path
        self.header_fields = {}
        self.tables = {}
        self._parse()

    def _read_text(self):
        """Read file with encoding auto-detect (UTF-16-LE BOM or UTF-8)."""
        with open(self.file_path, "rb") as f:
            raw = f.read()
        if raw[:2] == b"\xff\xfe":
            return raw[2:].decode("utf-16-le", errors="replace")
        if raw[:3] == b"\xef\xbb\xbf":
            return raw[3:].decode("utf-8", errors="replace")
        # No BOM - try UTF-16-LE first (P6 default), fallback UTF-8
        try:
            return raw.decode("utf-16-le")
        except UnicodeDecodeError:
            return raw.decode("utf-8", errors="replace")

    def _parse(self):
        text = self._read_text()
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        current_table = None
        for line in lines:
            if not line:
                continue
            if line.startswith("ERMHDR"):
                parts = line.split("\t")
                self.header_fields = {
                    "version": parts[1] if len(parts) > 1 else "",
                    "exported": parts[2] if len(parts) > 2 else "",
                    "user": parts[3] if len(parts) > 3 else "",
                    "app": parts[4] if len(parts) > 4 else "",
                    "currency": parts[5] if len(parts) > 5 else "",
                }
                continue
            if line.startswith("%T"):
                parts = line.split("\t", 1)
                if len(parts) == 2:
                    current_table = parts[1].strip()
                    self.tables[current_table] = {"headers": [], "rows": []}
                continue
            if line.startswith("%F"):
                if current_table is None:
                    continue
                parts = line.split("\t")
                self.tables[current_table]["headers"] = [p.strip() for p in parts[1:]]
                continue
            if line.startswith("%R"):
                if current_table is None:
                    continue
                headers = self.tables[current_table]["headers"]
                if not headers:
                    continue
                parts = line.split("\t")
                values = parts[1:]
                # Pad/truncate to match header count
                if len(values) < len(headers):
                    values = values + [""] * (len(headers) - len(values))
                row = {h: values[i] for i, h in enumerate(headers)}
                self.tables[current_table]["rows"].append(row)
                continue
            if line.startswith("%E"):
                break
            # Unknown marker - skip silently (forward-compat with new P6 markers)


# ---------- Field mapping helpers ----------

CONSTRAINT_TYPE_MAP = {
    "CS_ASAP": 0, "CS_ALAP": 1,
    "CS_MSO": 2, "CS_MFO": 3,
    "CS_MSOA": 4, "CS_MSOB": 5,
    "CS_MEOA": 6, "CS_MEOB": 7,
}

LINK_TYPE_MAP = {
    "PR_FS": "FS", "PR_SS": "SS", "PR_FF": "FF", "PR_SF": "SF",
}

# DCMA/MSP convention: Summary = WBS rollup or LOE (Level of Effort).
# Milestones (TT_Mile, TT_FinMile) are leaf tasks, NOT summaries.
SUMMARY_TASK_TYPES = {"TT_LOE", "TT_WBS"}


def _to_int(s, default=None):
    try:
        return int(float(s)) if s else default
    except (ValueError, TypeError):
        return default


def _to_float(s, default=0.0):
    try:
        return float(s) if s else default
    except (ValueError, TypeError):
        return default


def _to_iso_date(s):
    """XER dates are 'YYYY-MM-DD HH:MM' or empty. Return ISO date or None."""
    if not s or not s.strip():
        return None
    return s[:10]  # 'YYYY-MM-DD' prefix


# ---------- T103: read_tasks + read_links ----------

def _read_tasks(self, day_hr_cnt=8.0):
    """TASK section -> list of MSP-shape task dicts.

    day_hr_cnt: hours per working day (CAU = 9.0; default 8.0). Used to
    convert total_float_hr_cnt to days.
    """
    tbl = self.tables.get("TASK", {"rows": []})
    out = []
    for row in tbl["rows"]:
        ttype = row.get("task_type", "")
        out.append({
            "id": _to_int(row.get("task_id")),
            "name": row.get("task_name", ""),
            "code": row.get("task_code", ""),
            "duration_h": _to_float(row.get("target_drtn_hr_cnt")),
            "start": _to_iso_date(row.get("target_start_date")),
            "finish": _to_iso_date(row.get("target_end_date")),
            "actual_start": _to_iso_date(row.get("act_start_date")),
            "actual_finish": _to_iso_date(row.get("act_end_date")),
            "percent_complete": _to_float(row.get("phys_complete_pct")),
            "total_float": _to_float(row.get("total_float_hr_cnt")) / day_hr_cnt
                           if day_hr_cnt > 0 else 0.0,
            "summary": ttype in SUMMARY_TASK_TYPES,
            "task_type": ttype,
            "constraint_type": CONSTRAINT_TYPE_MAP.get(row.get("cstr_type", ""), 0),
            "status": row.get("status_code", ""),
        })
    return out


def _read_links(self):
    """TASKPRED section -> list of {from_id, to_id, type, lag_days}.

    XER `task_id` = successor; `pred_task_id` = predecessor. Map to MSP shape:
    from_id = predecessor, to_id = successor. Lag converted hr -> day @ 8h/day.
    """
    tbl = self.tables.get("TASKPRED", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "from_id": _to_int(row.get("pred_task_id")),
            "to_id": _to_int(row.get("task_id")),
            "type": LINK_TYPE_MAP.get(row.get("pred_type", ""), "FS"),
            "lag_days": _to_float(row.get("lag_hr_cnt")) / 8.0,
        })
    return out


XerFile.read_tasks = _read_tasks
XerFile.read_links = _read_links


# ---------- T104: read_resources + read_assignments + read_calendars ----------

def _read_resources(self):
    """RSRC section -> list of {id, name, code, type, max_units}.

    P6 RT_Labor/RT_Equip → MSP "Work"; RT_Mat → "Material".
    """
    tbl = self.tables.get("RSRC", {"rows": []})
    out = []
    for row in tbl["rows"]:
        rtype = row.get("rsrc_type", "")
        msp_type = "Material" if rtype == "RT_Mat" else "Work"
        out.append({
            "id": _to_int(row.get("rsrc_id")),
            "name": row.get("rsrc_name", ""),
            "code": row.get("rsrc_short_name", ""),
            "type": msp_type,
            "max_units": _to_float(row.get("max_qty_per_hr"), default=1.0),
        })
    return out


def _read_assignments(self):
    """TASKRSRC section -> list of {task_id, resource_id, target_qty,
    actual_qty, target_cost, actual_cost}."""
    tbl = self.tables.get("TASKRSRC", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "task_id": _to_int(row.get("task_id")),
            "resource_id": _to_int(row.get("rsrc_id")),
            "target_qty": _to_float(row.get("target_qty")),
            "actual_qty": _to_float(row.get("act_reg_qty")),
            "target_cost": _to_float(row.get("target_cost")),
            "actual_cost": _to_float(row.get("act_reg_cost")),
        })
    return out


def _read_calendars(self):
    """CALENDAR section -> list of {id, name, day_hr_cnt, week_hr_cnt}.

    clndr_data BLOB (holiday detail) NOT extracted - Phase 6 enhancement.
    """
    tbl = self.tables.get("CALENDAR", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "id": _to_int(row.get("clndr_id")),
            "name": row.get("clndr_name", ""),
            "day_hr_cnt": _to_float(row.get("day_hr_cnt"), default=8.0),
            "week_hr_cnt": _to_float(row.get("week_hr_cnt"), default=40.0),
        })
    return out


XerFile.read_resources = _read_resources
XerFile.read_assignments = _read_assignments
XerFile.read_calendars = _read_calendars


# ---------- T105: read_progress + status_date + project metadata ----------

def _read_project(self):
    """PROJECT section -> first row dict (typically only 1 project per XER)."""
    tbl = self.tables.get("PROJECT", {"rows": []})
    if not tbl["rows"]:
        return {}
    row = tbl["rows"][0]
    return {
        "proj_id": _to_int(row.get("proj_id")),
        "proj_short_name": row.get("proj_short_name", ""),
        "plan_start_date": _to_iso_date(row.get("plan_start_date")),
        "plan_end_date": _to_iso_date(row.get("plan_end_date")),
        "last_recalc_date": _to_iso_date(row.get("last_recalc_date")),
    }


def _read_progress(self):
    """Return {status_date, tasks: [{id, percent_complete, actual_work_h}]}.

    status_date = PROJECT.last_recalc_date (P6 convention — XER files do
    not track a separate 'data date' field, last recalc is the de-facto
    status date).

    actual_work_h: aggregate sum of TASKRSRC.act_reg_qty per task (XER
    stores actuals at assignment level, not task level).
    """
    proj = _read_project(self)
    # Pre-aggregate actual work per task from TASKRSRC
    actual_by_task = {}
    for asgn in self.tables.get("TASKRSRC", {"rows": []})["rows"]:
        tid = _to_int(asgn.get("task_id"))
        if tid is None:
            continue
        actual_by_task[tid] = actual_by_task.get(tid, 0.0) + _to_float(
            asgn.get("act_reg_qty"))
    progress_tasks = []
    for row in self.tables.get("TASK", {"rows": []})["rows"]:
        tid = _to_int(row.get("task_id"))
        progress_tasks.append({
            "id": tid,
            "percent_complete": _to_float(row.get("phys_complete_pct")),
            "actual_work_h": actual_by_task.get(tid, 0.0),
        })
    return {
        "status_date": (proj or {}).get("last_recalc_date"),
        "tasks": progress_tasks,
    }


XerFile.read_project = _read_project
XerFile.read_progress = _read_progress
