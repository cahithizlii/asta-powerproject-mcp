#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Native MSPDI XML Parser/Writer for Asta MCP Server.
====================================================
Zero Java/MPXJ dependency - pure Python using xml.etree.ElementTree.
Reads and writes Microsoft Project XML (MSPDI) files.
Preserves full XML structure for lossless round-trip editing.

Author: Claude AI for Cahit
Version: 1.0.0
"""

import os
import re
import copy
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
import logging

logger = logging.getLogger("mspdi_parser")

# Register default namespace to avoid ns0: prefixes in output
ET.register_namespace('', 'http://schemas.microsoft.com/project')


class MspdiProject:
    """Native MSPDI XML parser/writer. Zero Java dependency.

    Parses the full Microsoft Project XML schema and provides:
    - Complete task, link, resource, calendar, code library access
    - Fast lookup by ID and UID
    - In-place XML modification for lossless round-trip saving
    """

    NS = "http://schemas.microsoft.com/project"

    # MSPDI link type codes
    LINK_TYPES = {0: "FF", 1: "FS", 2: "SF", 3: "SS"}
    LINK_TYPE_IDS = {"FF": 0, "FS": 1, "SF": 2, "SS": 3}

    # MSPDI constraint type codes
    CONSTRAINT_TYPES = {
        0: "As Soon As Possible", 1: "As Late As Possible",
        2: "Must Start On", 3: "Must Finish On",
        4: "Start No Earlier Than", 5: "Start No Later Than",
        6: "Finish No Earlier Than", 7: "Finish No Later Than",
    }

    def __init__(self, file_path: str):
        self.file_path = file_path.replace("\\", "/")
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")

        self.tree = ET.parse(self.file_path)
        self.root = self.tree.getroot()

        # Project-level settings
        self.minutes_per_day = self._root_int("MinutesPerDay", 480)
        self.minutes_per_week = self._root_int("MinutesPerWeek", 2400)
        self.days_per_month = self._root_int("DaysPerMonth", 20)
        self.hours_per_day = self.minutes_per_day / 60.0

        # Indices (populated by _parse)
        self._tasks = {}          # UID (int) -> task dict
        self._tasks_by_id = {}    # ID (int) -> task dict
        self._uid_by_id = {}      # ID (int) -> UID (int)
        self._id_by_uid = {}      # UID (int) -> ID (int)
        self._task_elems = {}     # UID (int) -> XML Element
        self._resources = {}      # UID (int) -> resource dict
        self._resource_elems = {} # UID (int) -> XML Element
        self._assignments = []    # list of assignment dicts
        self._calendars = {}      # UID (int) -> calendar dict
        self._code_libs = {}      # FieldID (str) -> code library dict

        self._parse()

    # ------------------------------------------------------------------
    # XML helper methods
    # ------------------------------------------------------------------

    def _t(self, name: str) -> str:
        """Full tag with namespace."""
        return f"{{{self.NS}}}{name}"

    def _find(self, elem, name: str):
        return elem.find(self._t(name))

    def _findall(self, elem, name: str):
        return elem.findall(self._t(name))

    def _text(self, elem, name: str, default: str = "") -> str:
        child = elem.find(self._t(name))
        if child is not None and child.text:
            return child.text
        return default

    def _int(self, elem, name: str, default: int = 0) -> int:
        v = self._text(elem, name)
        try:
            return int(v)
        except (ValueError, TypeError):
            return default

    def _float(self, elem, name: str, default: float = 0.0) -> float:
        v = self._text(elem, name)
        try:
            return float(v)
        except (ValueError, TypeError):
            return default

    def _root_int(self, name: str, default: int = 0) -> int:
        return self._int(self.root, name, default)

    def _set_text(self, elem, name: str, value: str):
        """Set or create a child element's text."""
        child = elem.find(self._t(name))
        if child is None:
            child = ET.SubElement(elem, self._t(name))
        child.text = str(value)

    def _make_elem(self, name: str, text: str = None):
        """Create a new element with optional text."""
        elem = ET.Element(self._t(name))
        if text is not None:
            elem.text = str(text)
        return elem

    # ------------------------------------------------------------------
    # Duration / date / lag conversion
    # ------------------------------------------------------------------

    def _parse_iso_duration(self, dur_str: str) -> float:
        """Parse PT{H}H{M}M{S}S -> total hours."""
        if not dur_str or not dur_str.startswith("PT"):
            return 0.0
        m = re.match(r'PT(\d+)H(\d+)M(\d+)S', dur_str)
        if m:
            return int(m.group(1)) + int(m.group(2)) / 60.0 + int(m.group(3)) / 3600.0
        return 0.0

    def _hours_to_days(self, hours: float) -> float:
        """Working hours -> working days."""
        hpd = self.hours_per_day if self.hours_per_day > 0 else 8.0
        return hours / hpd

    def _days_to_hours(self, days: float) -> float:
        """Working days -> working hours."""
        return days * (self.hours_per_day if self.hours_per_day > 0 else 8.0)

    def _format_duration_str(self, dur_str: str) -> str:
        """Format PT3512H0M0S -> '439d' for display."""
        hours = self._parse_iso_duration(dur_str)
        days = self._hours_to_days(hours)
        if days == 0:
            return "0d"
        if days == int(days):
            return f"{int(days)}d"
        return f"{days:.1f}d"

    def _days_to_iso(self, days: float) -> str:
        """Working days -> PT{H}H0M0S."""
        hours = int(round(self._days_to_hours(days)))
        return f"PT{hours}H0M0S"

    def _parse_date(self, dt_str: str) -> str:
        """Parse '2025-06-23T08:00:00' -> '2025-06-23'."""
        if not dt_str:
            return "N/A"
        return dt_str[:10] if len(dt_str) >= 10 else dt_str

    def _lag_to_days(self, lag_val: int) -> float:
        """LinkLag (tenths of minutes) -> working days."""
        if not lag_val:
            return 0.0
        mpd = self.minutes_per_day if self.minutes_per_day > 0 else 480
        return lag_val / (mpd * 10.0)

    def _days_to_lag(self, days: float) -> int:
        """Working days -> LinkLag (tenths of minutes)."""
        mpd = self.minutes_per_day if self.minutes_per_day > 0 else 480
        return int(round(days * mpd * 10))

    def _format_lag(self, lag_val: int) -> str:
        """Format lag value for display."""
        days = self._lag_to_days(lag_val)
        if days == 0:
            return "0d"
        if days == int(days):
            return f"{int(days)}d"
        return f"{days:.1f}d"

    def _parse_duration_input(self, dur_str: str) -> float:
        """Parse user input like '10d', '2w', '80h' -> working days."""
        if not dur_str:
            return 1.0
        s = str(dur_str).strip().lower()
        if 'w' in s:
            return float(s.replace('w', '').strip()) * 5
        if 'mo' in s:
            return float(s.replace('mo', '').strip()) * self.days_per_month
        if 'h' in s:
            return float(s.replace('h', '').strip()) / self.hours_per_day
        if 'd' in s:
            return float(s.replace('d', '').strip())
        try:
            return float(s)
        except ValueError:
            return 1.0

    # ------------------------------------------------------------------
    # Full parse
    # ------------------------------------------------------------------

    def _parse(self):
        """Parse all project data from XML."""
        self._parse_code_libraries()
        self._parse_calendars_data()
        self._parse_tasks_data()
        self._parse_resources_data()
        self._parse_assignments_data()
        self._build_successors()

    def _parse_code_libraries(self):
        """Parse ExtendedAttributes -> code library definitions."""
        eas = self._find(self.root, "ExtendedAttributes")
        if eas is None:
            return
        for ea in self._findall(eas, "ExtendedAttribute"):
            fid = self._text(ea, "FieldID")
            alias = self._text(ea, "Alias")
            field_name = self._text(ea, "FieldName")

            # Extract library name from alias
            lib_name = ""
            if 'CodeLibrary:"' in alias:
                try:
                    lib_name = alias.split('"')[1]
                except IndexError:
                    lib_name = alias
            elif alias:
                lib_name = alias

            # Parse value list
            values = {}
            vl = self._find(ea, "ValueList")
            if vl is not None:
                for v in self._findall(vl, "Value"):
                    vid = self._text(v, "ID")
                    vval = self._text(v, "Value")
                    vdesc = self._text(v, "Description")
                    values[vid] = {"value": vval, "description": vdesc}

            self._code_libs[fid] = {
                "field_id": fid,
                "field_name": field_name,
                "alias": alias,
                "name": lib_name,
                "values": values,
                "element": ea,
            }

    def _parse_calendars_data(self):
        """Parse Calendars."""
        cals_elem = self._find(self.root, "Calendars")
        if cals_elem is None:
            return
        for cal in self._findall(cals_elem, "Calendar"):
            uid = self._int(cal, "UID")
            name = self._text(cal, "Name")
            is_base = self._text(cal, "IsBaseCalendar", "0") == "1"
            base_uid = self._int(cal, "BaseCalendarUID", -1)

            weekdays = []
            wd_elem = self._find(cal, "WeekDays")
            if wd_elem is not None:
                for wd in self._findall(wd_elem, "WeekDay"):
                    day_type = self._int(wd, "DayType")
                    day_working = self._text(wd, "DayWorking", "0") == "1"
                    times = []
                    wt_elem = self._find(wd, "WorkingTimes")
                    if wt_elem is not None:
                        for wt in self._findall(wt_elem, "WorkingTime"):
                            ft = self._text(wt, "FromTime")
                            tt = self._text(wt, "ToTime")
                            times.append({"from": ft, "to": tt})
                    weekdays.append({
                        "day_type": day_type,
                        "day_working": day_working,
                        "working_times": times,
                    })

            exceptions = []
            exc_elem = self._find(cal, "Exceptions")
            if exc_elem is not None:
                for ex in self._findall(exc_elem, "Exception"):
                    tp = self._find(ex, "TimePeriod")
                    from_date = ""
                    to_date = ""
                    if tp is not None:
                        from_date = self._parse_date(self._text(tp, "FromDate"))
                        to_date = self._parse_date(self._text(tp, "ToDate"))
                    exceptions.append({
                        "name": self._text(ex, "Name"),
                        "type": self._int(ex, "Type", 1),
                        "from_date": from_date,
                        "to_date": to_date,
                    })

            self._calendars[uid] = {
                "uid": uid,
                "name": name,
                "is_base": is_base,
                "base_calendar_uid": base_uid,
                "weekdays": weekdays,
                "exceptions": exceptions,
                "element": cal,
            }

    def _parse_tasks_data(self):
        """Parse Tasks with full detail."""
        tasks_elem = self._find(self.root, "Tasks")
        if tasks_elem is None:
            return
        for task_elem in self._findall(tasks_elem, "Task"):
            uid = self._int(task_elem, "UID")
            tid = self._int(task_elem, "ID")
            cal_uid = self._text(task_elem, "CalendarUID", "")

            # Resolve calendar name
            cal_name = "Default"
            if cal_uid:
                try:
                    cal_data = self._calendars.get(int(cal_uid))
                    if cal_data:
                        cal_name = cal_data["name"]
                except (ValueError, TypeError):
                    pass

            dur_raw = self._text(task_elem, "Duration")
            ts_raw = self._text(task_elem, "TotalSlack")
            fs_raw = self._text(task_elem, "FreeSlack")

            task = {
                "uid": uid,
                "id": tid,
                "name": self._text(task_elem, "Name"),
                "wbs": self._text(task_elem, "WBS"),
                "outline_number": self._text(task_elem, "OutlineNumber"),
                "outline_level": self._int(task_elem, "OutlineLevel", 0),
                "duration_raw": dur_raw,
                "duration": self._format_duration_str(dur_raw),
                "start": self._parse_date(self._text(task_elem, "Start")),
                "finish": self._parse_date(self._text(task_elem, "Finish")),
                "start_raw": self._text(task_elem, "Start"),
                "finish_raw": self._text(task_elem, "Finish"),
                "early_start": self._parse_date(self._text(task_elem, "EarlyStart")),
                "early_finish": self._parse_date(self._text(task_elem, "EarlyFinish")),
                "late_start": self._parse_date(self._text(task_elem, "LateStart")),
                "late_finish": self._parse_date(self._text(task_elem, "LateFinish")),
                "actual_start": self._parse_date(self._text(task_elem, "ActualStart")),
                "actual_finish": self._parse_date(self._text(task_elem, "ActualFinish")),
                "actual_work": self._text(task_elem, "ActualWork"),
                "remaining_work": self._text(task_elem, "RemainingWork"),
                "percent_complete": self._int(task_elem, "PercentComplete", 0),
                "milestone": self._text(task_elem, "Milestone", "0") == "1",
                "summary": self._text(task_elem, "Summary", "0") == "1",
                "critical": self._text(task_elem, "Critical", "0") == "1",
                "type": self._int(task_elem, "Type", 0),
                "constraint_type": self._int(task_elem, "ConstraintType", 0),
                "constraint_date": self._parse_date(self._text(task_elem, "ConstraintDate")),
                "calendar_uid": cal_uid,
                "calendar": cal_name,
                "notes": self._text(task_elem, "Notes"),
                "fixed_cost": self._float(task_elem, "FixedCost", 0.0),
                "cost": self._float(task_elem, "FixedCost", 0.0),
                "actual_cost": self._float(task_elem, "ActualCost", 0.0),
                "work_raw": self._text(task_elem, "Work"),
                "work": self._format_duration_str(self._text(task_elem, "Work")) if self._text(task_elem, "Work") else "0h",
                "total_float": self._format_duration_str(ts_raw) if ts_raw else "N/A",
                "free_float": self._format_duration_str(fs_raw) if fs_raw else "N/A",
                "predecessors": [],
                "successors": [],
                "codes": {},
            }

            # Parse predecessor links
            for pred_elem in self._findall(task_elem, "PredecessorLink"):
                pred_uid = self._int(pred_elem, "PredecessorUID")
                link_type_id = self._int(pred_elem, "Type", 1)
                lag_raw = self._int(pred_elem, "LinkLag", 0)
                lag_fmt = self._int(pred_elem, "LagFormat", 7)
                task["predecessors"].append({
                    "predecessor_uid": pred_uid,
                    "type_id": link_type_id,
                    "type": self.LINK_TYPES.get(link_type_id, "FS"),
                    "lag_raw": lag_raw,
                    "lag": self._format_lag(lag_raw),
                    "lag_format": lag_fmt,
                })

            # Parse extended attributes (code assignments)
            for ea_elem in self._findall(task_elem, "ExtendedAttribute"):
                fid = self._text(ea_elem, "FieldID")
                value = self._text(ea_elem, "Value")
                value_id = self._text(ea_elem, "ValueID")
                if fid in self._code_libs:
                    lib = self._code_libs[fid]
                    task["codes"][lib["name"]] = {
                        "field_id": fid,
                        "value": value,
                        "value_id": value_id,
                    }

            self._tasks[uid] = task
            self._task_elems[uid] = task_elem
            self._tasks_by_id[tid] = task
            self._uid_by_id[tid] = uid
            self._id_by_uid[uid] = tid

    def _build_successors(self):
        """Build successor index from predecessor data."""
        for uid, task in self._tasks.items():
            for pred in task["predecessors"]:
                pred_uid = pred["predecessor_uid"]
                if pred_uid in self._tasks:
                    self._tasks[pred_uid]["successors"].append({
                        "successor_uid": uid,
                        "successor_id": self._id_by_uid.get(uid, 0),
                        "type": pred["type"],
                        "lag": pred["lag"],
                    })

    def _parse_resources_data(self):
        """Parse Resources."""
        res_elem = self._find(self.root, "Resources")
        if res_elem is None:
            return
        for r in self._findall(res_elem, "Resource"):
            uid = self._int(r, "UID")
            self._resources[uid] = {
                "uid": uid,
                "id": self._int(r, "ID"),
                "name": self._text(r, "Name"),
                "type": self._int(r, "Type", 1),
                "type_name": {0: "Material", 1: "Work", 2: "Cost"}.get(self._int(r, "Type", 1), "Work"),
                "max_units": self._float(r, "MaxUnits", 1.0),
                "standard_rate": self._text(r, "StandardRate", "0"),
                "cost": self._float(r, "Cost", 0.0),
                "calendar_uid": self._text(r, "CalendarUID", ""),
                "email": self._text(r, "EmailAddress"),
                "group": self._text(r, "Group"),
            }
            self._resource_elems[uid] = r

    def _parse_assignments_data(self):
        """Parse Assignments."""
        asgn_elem = self._find(self.root, "Assignments")
        if asgn_elem is None:
            return
        for a in self._findall(asgn_elem, "Assignment"):
            task_uid = self._int(a, "TaskUID")
            res_uid = self._int(a, "ResourceUID")

            task_name = ""
            task_id = 0
            if task_uid in self._tasks:
                task_name = self._tasks[task_uid]["name"]
                task_id = self._tasks[task_uid]["id"]

            res_name = ""
            res_id = 0
            if res_uid in self._resources:
                res_name = self._resources[res_uid]["name"]
                res_id = self._resources[res_uid]["id"]

            self._assignments.append({
                "uid": self._int(a, "UID"),
                "task_uid": task_uid,
                "task_id": task_id,
                "task_name": task_name,
                "resource_uid": res_uid,
                "resource_id": res_id,
                "resource_name": res_name,
                "units": self._float(a, "Units", 1.0),
                "work_raw": self._text(a, "Work"),
                "work": self._format_duration_str(self._text(a, "Work")) if self._text(a, "Work") else "0h",
                "cost": self._float(a, "Cost", 0.0),
                "start": self._parse_date(self._text(a, "Start")),
                "finish": self._parse_date(self._text(a, "Finish")),
            })

    # ------------------------------------------------------------------
    # Helper: task dict to external format (matching AstaFileManager)
    # ------------------------------------------------------------------

    def _task_to_list_dict(self, task: dict) -> dict:
        """Convert internal task dict to the format expected by list_tasks."""
        # Resolve predecessor UIDs to IDs
        preds = []
        for p in task["predecessors"]:
            pred_id = self._id_by_uid.get(p["predecessor_uid"], p["predecessor_uid"])
            preds.append({
                "task_id": pred_id,
                "type": p["type"],
                "lag": p["lag"],
            })
        succs = []
        for s in task["successors"]:
            succ_id = s.get("successor_id", self._id_by_uid.get(s["successor_uid"], s["successor_uid"]))
            succs.append({
                "task_id": succ_id,
                "type": s["type"],
                "lag": s["lag"],
            })
        return {
            "id": task["id"],
            "unique_id": task["uid"],
            "name": task["name"],
            "duration": task["duration"],
            "start": task["start"],
            "finish": task["finish"],
            "percent_complete": task["percent_complete"],
            "critical": task["critical"],
            "milestone": task["milestone"],
            "summary": task["summary"],
            "total_float": task["total_float"],
            "notes": task["notes"],
            "predecessors": preds,
            "successors": succs,
            # T68: progress fields (data already in self._tasks; expose for adapter)
            "actual_start": task.get("actual_start"),
            "actual_finish": task.get("actual_finish"),
            "actual_work": task.get("actual_work"),
            "remaining_work": task.get("remaining_work"),
        }

    def _task_to_detail_dict(self, task: dict) -> dict:
        """Convert internal task dict to the detailed format expected by get_task."""
        d = self._task_to_list_dict(task)
        d.update({
            "early_start": task["early_start"],
            "early_finish": task["early_finish"],
            "late_start": task["late_start"],
            "late_finish": task["late_finish"],
            "actual_start": task["actual_start"],
            "actual_finish": task["actual_finish"],
            "free_float": task["free_float"],
            "calendar": task["calendar"],
            "cost": task["cost"],
            "actual_cost": task["actual_cost"],
            "work": task["work"],
            "constraint_type": self.CONSTRAINT_TYPES.get(task["constraint_type"], "N/A"),
            "constraint_date": task["constraint_date"],
            "wbs": task["wbs"],
            "outline_level": task["outline_level"],
            "codes": {k: v["value"] for k, v in task["codes"].items()},
        })
        return d

    # ==================================================================
    # READ METHODS (matching AstaFileManager interface)
    # ==================================================================

    def get_project_summary(self) -> dict:
        """Get project overview."""
        all_tasks = list(self._tasks.values())
        return {
            "file": self.file_path,
            "project_name": self._text(self.root, "Name", "Unnamed"),
            "client": self._text(self.root, "Author", "N/A"),
            "start_date": self._parse_date(self._text(self.root, "StartDate")),
            "finish_date": self._parse_date(self._text(self.root, "FinishDate")),
            "status_date": self._parse_date(self._text(self.root, "StatusDate")),
            "current_date": self._parse_date(self._text(self.root, "CurrentDate")),
            "total_tasks": len(all_tasks),
            "summary_tasks": sum(1 for t in all_tasks if t["summary"]),
            "milestones": sum(1 for t in all_tasks if t["milestone"]),
            "activities": sum(1 for t in all_tasks if not t["summary"] and not t["milestone"]),
            "critical_tasks": sum(1 for t in all_tasks if t["critical"]),
            "total_resources": len(self._resources),
            "total_assignments": len(self._assignments),
            "total_links": sum(len(t["predecessors"]) for t in all_tasks),
            "calendars": len(self._calendars),
            "code_libraries": len(set(lib["name"] for lib in self._code_libs.values() if lib["name"])),
        }

    def get_all_tasks(self, include_summary: bool = True) -> List[dict]:
        """Get all tasks in ID order."""
        tasks = sorted(self._tasks.values(), key=lambda t: t["id"])
        if not include_summary:
            tasks = [t for t in tasks if not t["summary"]]
        return [self._task_to_list_dict(t) for t in tasks]

    def get_task_by_id(self, task_id: int) -> Optional[dict]:
        """Get detailed task info by ID."""
        task = self._tasks_by_id.get(task_id)
        if not task:
            return None
        return self._task_to_detail_dict(task)

    def get_task_by_uid(self, uid: int) -> Optional[dict]:
        """Get detailed task info by UID."""
        task = self._tasks.get(uid)
        if not task:
            return None
        return self._task_to_detail_dict(task)

    def get_critical_path(self) -> List[dict]:
        """Get all critical non-summary tasks."""
        critical = []
        for task in sorted(self._tasks.values(), key=lambda t: t["id"]):
            if task["critical"] and not task["summary"]:
                critical.append({
                    "id": task["id"],
                    "name": task["name"],
                    "duration": task["duration"],
                    "start": task["start"],
                    "finish": task["finish"],
                    "total_float": task["total_float"],
                    "milestone": task["milestone"],
                })
        return critical

    def get_resources(self) -> List[dict]:
        """Get all resources."""
        resources = []
        for r in sorted(self._resources.values(), key=lambda x: x["id"]):
            resources.append({
                "id": r["id"],
                "unique_id": r["uid"],
                "name": r["name"],
                "type": r["type_name"],
                "max_units": r["max_units"],
                "standard_rate": r["standard_rate"],
                "cost": r["cost"],
                "calendar": "",
            })
        return resources

    def get_resource_assignments(self) -> List[dict]:
        """Get all resource assignments."""
        assignments = []
        for a in self._assignments:
            assignments.append({
                "task_id": a["task_id"],
                "task_name": a["task_name"],
                "resource_id": a["resource_id"],
                "resource_name": a["resource_name"],
                "units": a["units"],
                "work": a["work"],
                "cost": a["cost"],
            })
        return assignments

    def get_calendars(self) -> List[dict]:
        """Get all calendars with is_base flag (T67 review fix — was stripped)."""
        return [{"id": c["uid"], "name": c["name"], "is_base": c.get("is_base", False)}
                for c in self._calendars.values()]

    def get_wbs_tree(self, max_depth: int = 99) -> List[dict]:
        """Get WBS hierarchy tree."""
        # Build parent-child relationships using outline levels
        tasks_sorted = sorted(self._tasks.values(), key=lambda t: t["id"])
        if not tasks_sorted:
            return []

        # Use outline_level to build tree
        root_nodes = []
        stack = []  # (level, node)

        for task in tasks_sorted:
            node = {
                "id": task["id"],
                "unique_id": task["uid"],
                "name": task["name"],
                "wbs": task["wbs"],
                "outline_level": task["outline_level"],
                "summary": task["summary"],
                "milestone": task["milestone"],
                "duration": task["duration"],
                "start": task["start"],
                "finish": task["finish"],
                "level": task["outline_level"],
                "children": [],
            }

            level = task["outline_level"]

            # Pop stack until we find the parent level
            while stack and stack[-1][0] >= level:
                stack.pop()

            if stack:
                parent_node = stack[-1][1]
                if len(stack) <= max_depth:
                    parent_node["children"].append(node)
                else:
                    if "children_count" not in parent_node:
                        parent_node["children_count"] = 0
                        parent_node["children_truncated"] = True
                    parent_node["children_count"] += 1
            else:
                root_nodes.append(node)

            if task["summary"] or level < max_depth:
                stack.append((level, node))

        return root_nodes

    def get_delay_analysis(self) -> dict:
        """Analyze schedule delays."""
        delays = []
        for task in self._tasks.values():
            if task["summary"] or task["milestone"]:
                continue
            if task["actual_start"] == "N/A" and task["actual_finish"] == "N/A":
                continue

            start_slip = None
            finish_slip = None
            try:
                if task["start"] != "N/A" and task["actual_start"] != "N/A":
                    from datetime import date
                    ps = date.fromisoformat(task["start"])
                    as_ = date.fromisoformat(task["actual_start"])
                    start_slip = (as_ - ps).days
            except (ValueError, TypeError):
                pass
            try:
                if task["finish"] != "N/A" and task["actual_finish"] != "N/A":
                    from datetime import date
                    pf = date.fromisoformat(task["finish"])
                    af = date.fromisoformat(task["actual_finish"])
                    finish_slip = (af - pf).days
            except (ValueError, TypeError):
                pass

            if start_slip is not None or finish_slip is not None:
                delays.append({
                    "id": task["id"],
                    "name": task["name"],
                    "planned_start": task["start"],
                    "actual_start": task["actual_start"],
                    "start_slip_days": start_slip,
                    "planned_finish": task["finish"],
                    "actual_finish": task["actual_finish"],
                    "finish_slip_days": finish_slip,
                    "percent_complete": task["percent_complete"],
                    "critical": task["critical"],
                })

        delayed_starts = [d for d in delays if d["start_slip_days"] and d["start_slip_days"] > 0]
        delayed_finishes = [d for d in delays if d["finish_slip_days"] and d["finish_slip_days"] > 0]
        early_starts = [d for d in delays if d["start_slip_days"] and d["start_slip_days"] < 0]

        return {
            "total_with_actuals": len(delays),
            "delayed_starts": len(delayed_starts),
            "delayed_finishes": len(delayed_finishes),
            "early_starts": len(early_starts),
            "max_start_slip": max((d["start_slip_days"] for d in delays if d["start_slip_days"]), default=0),
            "max_finish_slip": max((d["finish_slip_days"] for d in delays if d["finish_slip_days"]), default=0),
            "tasks": delays,
        }

    def get_float_analysis(self) -> dict:
        """Analyze float distribution.

        Uses TotalSlack from XML if available. Falls back to Critical flag
        (critical=zero float, non-critical=unknown float) when TotalSlack is absent.
        """
        float_data = {"zero_float": 0, "low_float": 0, "medium_float": 0,
                      "high_float": 0, "has_float_data": False, "tasks": []}

        for task in sorted(self._tasks.values(), key=lambda t: t["id"]):
            if task["summary"] or task["milestone"]:
                continue

            # Try to get TotalSlack from the parsed data
            elem = self._task_elems.get(task["uid"])
            ts_raw = self._text(elem, "TotalSlack", "") if elem is not None else ""

            if ts_raw:
                float_data["has_float_data"] = True
                hours = self._parse_iso_duration(ts_raw)
                days = self._hours_to_days(hours)
                if days == 0:
                    float_data["zero_float"] += 1
                elif days <= 5:
                    float_data["low_float"] += 1
                elif days <= 20:
                    float_data["medium_float"] += 1
                else:
                    float_data["high_float"] += 1
                float_data["tasks"].append({
                    "id": task["id"],
                    "name": task["name"],
                    "total_float": self._format_duration_str(ts_raw),
                })
            else:
                # Fallback: use Critical flag
                if task["critical"]:
                    float_data["zero_float"] += 1
                    float_data["tasks"].append({
                        "id": task["id"],
                        "name": task["name"],
                        "total_float": "0d (critical)",
                    })

        if not float_data["has_float_data"] and float_data["tasks"]:
            float_data["note"] = "TotalSlack not in file. Showing critical tasks as zero-float only. Run reschedule in Asta/MS Project for full float data."

        return float_data

    def get_resource_loading(self) -> dict:
        """Analyze resource loading."""
        resource_summary = {}
        for a in self._assignments:
            res_uid = a["resource_uid"]
            if res_uid not in self._resources:
                continue
            res = self._resources[res_uid]
            if res_uid not in resource_summary:
                resource_summary[res_uid] = {
                    "id": res["id"],
                    "name": res["name"],
                    "type": res["type_name"],
                    "max_units": res["max_units"],
                    "total_work": 0,
                    "total_cost": 0,
                    "task_count": 0,
                    "tasks": [],
                }
            work_hours = self._parse_iso_duration(a["work_raw"]) if a["work_raw"] else 0
            resource_summary[res_uid]["total_work"] += work_hours
            resource_summary[res_uid]["total_cost"] += a["cost"]
            resource_summary[res_uid]["task_count"] += 1
            resource_summary[res_uid]["tasks"].append({
                "task_id": a["task_id"],
                "task_name": a["task_name"],
                "units": a["units"],
                "work": a["work"],
                "cost": a["cost"],
                "start": a["start"],
                "finish": a["finish"],
            })
        return {
            "total_resources": len(resource_summary),
            "total_assignments": len(self._assignments),
            "resources": list(resource_summary.values()),
        }

    # ==================================================================
    # NEW QUERY METHODS (not in AstaFileManager)
    # ==================================================================

    def get_code_libraries(self) -> List[dict]:
        """Get all code library definitions with their values."""
        seen = {}  # lib_name -> merged dict
        for fid, lib in self._code_libs.items():
            name = lib["name"]
            if not name:
                continue
            if name not in seen:
                seen[name] = {
                    "name": name,
                    "field_ids": [],
                    "values": [],
                }
            seen[name]["field_ids"].append(fid)
            for vid, vdata in lib["values"].items():
                # Avoid duplicates
                existing_vals = {v["value"] for v in seen[name]["values"]}
                if vdata["value"] not in existing_vals:
                    seen[name]["values"].append({
                        "id": vid,
                        "value": vdata["value"],
                        "description": vdata["description"],
                    })
        return sorted(seen.values(), key=lambda x: x["name"])

    def get_task_codes(self, task_id: int) -> dict:
        """Get all code assignments for a task."""
        task = self._tasks_by_id.get(task_id)
        if not task:
            return {"error": f"Task ID {task_id} not found"}
        return {
            "task_id": task_id,
            "task_name": task["name"],
            "codes": {k: v["value"] for k, v in task["codes"].items()},
        }

    def filter_tasks_by_code(self, library_name: str, value: str = None) -> List[dict]:
        """Filter tasks by code library name and optionally value."""
        result = []
        for task in sorted(self._tasks.values(), key=lambda t: t["id"]):
            if library_name in task["codes"]:
                code_val = task["codes"][library_name]["value"]
                if value is None or value.lower() in code_val.lower():
                    result.append({
                        "id": task["id"],
                        "name": task["name"],
                        "code_value": code_val,
                        "duration": task["duration"],
                        "start": task["start"],
                        "finish": task["finish"],
                        "milestone": task["milestone"],
                        "summary": task["summary"],
                        "critical": task["critical"],
                    })
        return result

    def get_latest_finishing(self, count: int = 20) -> List[dict]:
        """Get tasks with the latest finish dates."""
        tasks_with_dates = []
        for task in self._tasks.values():
            if task["summary"]:
                continue
            if task["finish"] and task["finish"] != "N/A":
                tasks_with_dates.append(task)

        tasks_with_dates.sort(key=lambda t: t["finish"], reverse=True)
        result = []
        for task in tasks_with_dates[:count]:
            d = self._task_to_list_dict(task)
            d["codes"] = {k: v["value"] for k, v in task["codes"].items()}
            result.append(d)
        return result

    def find_missing_links(self) -> dict:
        """Find tasks with missing predecessors or successors (open ends)."""
        no_predecessors = []
        no_successors = []
        for task in sorted(self._tasks.values(), key=lambda t: t["id"]):
            if task["summary"]:
                continue
            # Skip first/last tasks by outline level 1
            if task["outline_level"] <= 1:
                continue
            if not task["predecessors"]:
                no_predecessors.append({
                    "id": task["id"],
                    "name": task["name"],
                    "start": task["start"],
                    "finish": task["finish"],
                    "milestone": task["milestone"],
                    "critical": task["critical"],
                })
            if not task["successors"]:
                no_successors.append({
                    "id": task["id"],
                    "name": task["name"],
                    "start": task["start"],
                    "finish": task["finish"],
                    "milestone": task["milestone"],
                    "critical": task["critical"],
                })
        return {
            "no_predecessors_count": len(no_predecessors),
            "no_successors_count": len(no_successors),
            "no_predecessors": no_predecessors,
            "no_successors": no_successors,
        }

    def search_tasks(self, pattern: str, include_summary: bool = True) -> List[dict]:
        """Search tasks by name pattern (case-insensitive)."""
        pat = pattern.lower()
        result = []
        for task in sorted(self._tasks.values(), key=lambda t: t["id"]):
            if not include_summary and task["summary"]:
                continue
            if pat in task["name"].lower():
                result.append(self._task_to_list_dict(task))
        return result

    def get_link_chain(self, from_pattern: str, to_pattern: str) -> dict:
        """Trace link chains between tasks matching two name patterns.

        Finds all paths from tasks matching from_pattern to tasks matching to_pattern.
        Useful for checking design->procurement->construction chains.
        """
        from_pat = from_pattern.lower()
        to_pat = to_pattern.lower()

        from_tasks = {t["uid"] for t in self._tasks.values() if from_pat in t["name"].lower()}
        to_tasks = {t["uid"] for t in self._tasks.values() if to_pat in t["name"].lower()}

        if not from_tasks:
            return {"error": f"No tasks matching '{from_pattern}'"}
        if not to_tasks:
            return {"error": f"No tasks matching '{to_pattern}'"}

        # BFS from each from_task, looking for paths to any to_task
        chains = []
        for start_uid in from_tasks:
            # BFS
            queue = [(start_uid, [start_uid])]
            visited = {start_uid}
            while queue:
                current_uid, path = queue.pop(0)
                if current_uid in to_tasks and len(path) > 1:
                    chain_detail = []
                    for i, uid in enumerate(path):
                        t = self._tasks[uid]
                        link_info = ""
                        if i > 0:
                            prev_uid = path[i - 1]
                            for p in t["predecessors"]:
                                if p["predecessor_uid"] == prev_uid:
                                    link_info = f" [{p['type']}"
                                    if p["lag"] != "0d":
                                        link_info += f"+{p['lag']}"
                                    link_info += "]"
                                    break
                        chain_detail.append({
                            "id": t["id"],
                            "name": t["name"],
                            "link": link_info.strip(),
                        })
                    chains.append(chain_detail)
                    continue  # Don't explore beyond target

                # Follow successors
                task = self._tasks.get(current_uid)
                if task and len(path) < 15:  # Max chain depth
                    for s in task["successors"]:
                        s_uid = s["successor_uid"]
                        if s_uid not in visited:
                            visited.add(s_uid)
                            queue.append((s_uid, path + [s_uid]))

        return {
            "from_pattern": from_pattern,
            "to_pattern": to_pattern,
            "from_tasks_found": len(from_tasks),
            "to_tasks_found": len(to_tasks),
            "chains_found": len(chains),
            "chains": chains[:50],  # Limit output
        }

    def get_tasks_between_dates(self, start_after: str = None, finish_before: str = None) -> List[dict]:
        """Filter tasks by date range."""
        result = []
        for task in sorted(self._tasks.values(), key=lambda t: t["id"]):
            if task["summary"]:
                continue
            if start_after and task["start"] != "N/A" and task["start"] < start_after:
                continue
            if finish_before and task["finish"] != "N/A" and task["finish"] > finish_before:
                continue
            result.append(self._task_to_list_dict(task))
        return result

    # ==================================================================
    # WRITE METHODS
    # ==================================================================

    def _next_uid(self) -> int:
        """Get next available UID across tasks, resources, assignments."""
        max_uid = 0
        for uid in self._tasks:
            if uid > max_uid:
                max_uid = uid
        for uid in self._resources:
            if uid > max_uid:
                max_uid = uid
        return max_uid + 1

    def _next_task_id(self) -> int:
        """Get next available task ID."""
        max_id = 0
        for tid in self._tasks_by_id:
            if tid > max_id:
                max_id = tid
        return max_id + 1

    def add_task(self, name: str, duration_str: str = "1d",
                 start_date: str = None, finish_date: str = None,
                 is_milestone: bool = False, is_summary: bool = False,
                 parent_task_id: int = None, calendar_uid: int = None) -> dict:
        """Add a new task to the XML."""
        tasks_elem = self._find(self.root, "Tasks")
        if tasks_elem is None:
            tasks_elem = ET.SubElement(self.root, self._t("Tasks"))

        uid = self._next_uid()
        tid = self._next_task_id()
        days = self._parse_duration_input(duration_str)
        if is_milestone:
            days = 0

        # Determine outline level and WBS
        outline_level = 1
        parent_wbs = ""
        if parent_task_id is not None:
            parent = self._tasks_by_id.get(parent_task_id)
            if parent:
                outline_level = parent["outline_level"] + 1
                parent_wbs = parent["wbs"]

        wbs = f"{parent_wbs}.{tid}" if parent_wbs else str(tid)

        # Build start/finish dates
        if not start_date:
            start_date = self._parse_date(self._text(self.root, "StartDate"))
        if start_date == "N/A":
            start_date = datetime.now().strftime("%Y-%m-%d")
        start_dt = f"{start_date}T08:00:00"

        if finish_date:
            finish_dt = f"{finish_date}T17:00:00"
        else:
            # Approximate finish from duration (calendar-naive)
            try:
                from datetime import date, timedelta
                sd = date.fromisoformat(start_date)
                fd = sd + timedelta(days=int(days * 7 / 5))  # rough working days
                finish_dt = f"{fd.isoformat()}T17:00:00"
            except (ValueError, TypeError):
                finish_dt = start_dt

        # Determine insert position (after parent's last child, or at end)
        insert_index = len(list(self._findall(tasks_elem, "Task")))
        if parent_task_id is not None:
            # Find parent's position and insert after its last child
            for i, t_elem in enumerate(self._findall(tasks_elem, "Task")):
                t_id = self._int(t_elem, "ID")
                if t_id == parent_task_id:
                    # Find last child after this
                    insert_index = i + 1
                    parent_level = self._int(t_elem, "OutlineLevel", 0)
                    for j in range(i + 1, len(list(self._findall(tasks_elem, "Task")))):
                        next_elem = list(self._findall(tasks_elem, "Task"))[j]
                        next_level = self._int(next_elem, "OutlineLevel", 0)
                        if next_level > parent_level:
                            insert_index = j + 1
                        else:
                            break
                    break

        # Create task element
        task_elem = ET.Element(self._t("Task"))
        fields = [
            ("UID", str(uid)),
            ("ID", str(tid)),
            ("Name", name),
            ("Type", "0"),
            ("IsNull", "0"),
            ("WBS", wbs),
            ("OutlineNumber", wbs),
            ("OutlineLevel", str(outline_level)),
            ("Start", start_dt),
            ("Finish", finish_dt),
            ("Duration", self._days_to_iso(days)),
            ("DurationFormat", "7"),
            ("Milestone", "1" if is_milestone else "0"),
            ("Summary", "1" if is_summary else "0"),
            ("Critical", "0"),
            ("PercentComplete", "0"),
            ("ConstraintType", "0"),
            ("CalendarUID", str(calendar_uid) if calendar_uid else
             self._text(self.root, "CalendarUID", "1")),
        ]
        for tag, val in fields:
            sub = ET.SubElement(task_elem, self._t(tag))
            sub.text = val

        # Insert into XML tree
        all_task_elems = list(self._findall(tasks_elem, "Task"))
        if insert_index >= len(all_task_elems):
            tasks_elem.append(task_elem)
        else:
            # Find the actual XML index (Tasks may have other children)
            ref_elem = all_task_elems[insert_index]
            parent_children = list(tasks_elem)
            ref_index = parent_children.index(ref_elem)
            tasks_elem.insert(ref_index, task_elem)

        # Update internal indices
        task_dict = {
            "uid": uid, "id": tid, "name": name,
            "wbs": wbs, "outline_number": wbs, "outline_level": outline_level,
            "duration_raw": self._days_to_iso(days), "duration": f"{int(days)}d" if days == int(days) else f"{days:.1f}d",
            "start": start_date, "finish": self._parse_date(finish_dt),
            "start_raw": start_dt, "finish_raw": finish_dt,
            "early_start": "N/A", "early_finish": "N/A",
            "late_start": "N/A", "late_finish": "N/A",
            "actual_start": "N/A", "actual_finish": "N/A",
            "percent_complete": 0, "milestone": is_milestone, "summary": is_summary,
            "critical": False, "type": 0, "constraint_type": 0, "constraint_date": "N/A",
            "calendar_uid": str(calendar_uid) if calendar_uid else "1",
            "calendar": "Default", "notes": "",
            "fixed_cost": 0.0, "cost": 0.0, "actual_cost": 0.0,
            "work_raw": "", "work": "0h",
            "total_float": "N/A", "free_float": "N/A",
            "predecessors": [], "successors": [], "codes": {},
        }
        self._tasks[uid] = task_dict
        self._task_elems[uid] = task_elem
        self._tasks_by_id[tid] = task_dict
        self._uid_by_id[tid] = uid
        self._id_by_uid[uid] = tid

        return {
            "task_id": tid,
            "uid": uid,
            "name": name,
            "duration": task_dict["duration"],
            "start": start_date,
            "type": "milestone" if is_milestone else "summary" if is_summary else "task",
            "message": f"Task '{name}' added (ID: {tid}, UID: {uid})",
        }

    def _next_resource_id(self) -> int:
        """Get next available resource ID."""
        max_id = 0
        for r in self._resources.values():
            rid = r.get("id", 0)
            if rid > max_id:
                max_id = rid
        return max_id + 1

    # MSPDI Resource Type enum: 0=Material, 1=Work, 2=Cost
    _RESOURCE_TYPE_MAP = {"Material": 0, "Work": 1, "Cost": 2}

    def add_resource(self, name: str, type: str = "Work",
                     max_units: float = 1.0,
                     standard_rate: str = None) -> int:
        """Add a new resource to the XML. Phase 4 T70 extension.

        Mirrors add_task style: create XML element, append to <Resources>,
        update _resources dict + _resource_elems, return new resource ID.

        Returns the new resource ID (int), or raises ValueError on bad input.
        """
        if not name:
            raise ValueError("Resource name is required")
        type_code = self._RESOURCE_TYPE_MAP.get(type)
        if type_code is None:
            raise ValueError(
                f"Resource type must be Work/Material/Cost, got {type!r}")

        resources_elem = self._find(self.root, "Resources")
        if resources_elem is None:
            resources_elem = ET.SubElement(self.root, self._t("Resources"))

        uid = self._next_uid()
        rid = self._next_resource_id()

        # Create resource element
        resource_elem = ET.Element(self._t("Resource"))
        fields = [
            ("UID", str(uid)),
            ("ID", str(rid)),
            ("Name", name),
            ("Type", str(type_code)),
            ("IsNull", "0"),
            ("MaxUnits", f"{float(max_units):.6f}"),
        ]
        if standard_rate is not None:
            fields.append(("StandardRate", str(standard_rate)))
        for tag, val in fields:
            sub = ET.SubElement(resource_elem, self._t(tag))
            sub.text = val

        resources_elem.append(resource_elem)

        # Update internal indices
        resource_dict = {
            "uid": uid, "id": rid, "name": name,
            "type": type_code, "type_name": type,
            "max_units": float(max_units),
            "standard_rate": standard_rate or "0",
            "cost": 0.0, "calendar_uid": "", "email": "", "group": "",
        }
        self._resources[uid] = resource_dict
        self._resource_elems[uid] = resource_elem

        return rid

    def _next_assignment_uid(self) -> int:
        """Get next available assignment UID (starts at 0 for assignments)."""
        max_uid = -1
        for a in self._assignments:
            uid = a.get("uid", -1)
            if uid > max_uid:
                max_uid = uid
        return max_uid + 1

    def add_assignment(self, task_id: int, resource_id: int,
                       units: float = 1.0,
                       work_str: str = None) -> int:
        """Add a single task-resource assignment. T73 extension.

        Returns the new assignment UID. Raises ValueError on missing IDs.
        """
        # Resolve task UID
        task_uid = self._uid_by_id.get(task_id)
        if task_uid is None:
            raise ValueError(f"Task ID {task_id} not found")
        # Resolve resource UID
        resource_uid = None
        for r in self._resources.values():
            if r.get("id") == resource_id:
                resource_uid = r.get("uid")
                break
        if resource_uid is None:
            raise ValueError(f"Resource ID {resource_id} not found")

        # Get or create <Assignments> root collection
        assignments_elem = self._find(self.root, "Assignments")
        if assignments_elem is None:
            assignments_elem = ET.SubElement(self.root, self._t("Assignments"))

        a_uid = self._next_assignment_uid()

        # Use task's start/finish + work derived from task duration as defaults
        task = self._tasks_by_id.get(task_id, {})
        start = task.get("start_raw") or task.get("start") or "2026-01-01T08:00:00"
        finish = task.get("finish_raw") or task.get("finish") or start
        if work_str:
            days = self._parse_duration_input(work_str)
            work_iso = self._days_to_iso(days)
        else:
            # Default: same as task duration
            work_iso = task.get("duration_raw") or "PT8H0M0S"

        # Create <Assignment> element
        a_elem = ET.Element(self._t("Assignment"))
        fields = [
            ("UID", str(a_uid)),
            ("TaskUID", str(task_uid)),
            ("ResourceUID", str(resource_uid)),
            ("Units", f"{float(units):.2f}"),
            ("Work", work_iso),
            ("Start", start),
            ("Finish", finish),
        ]
        for tag, val in fields:
            sub = ET.SubElement(a_elem, self._t(tag))
            sub.text = val
        assignments_elem.append(a_elem)

        # Update internal index
        self._assignments.append({
            "uid": a_uid,
            "task_id": task_id, "task_uid": task_uid,
            "task_name": task.get("name", ""),
            "resource_id": resource_id, "resource_uid": resource_uid,
            "resource_name": (self._resources.get(resource_uid) or {}).get("name", ""),
            "units": float(units),
            "work": work_str or task.get("duration", ""),
            "cost": 0.0,
        })
        return a_uid

    def bulk_add_assignments(self, items: List[dict]) -> int:
        """Bulk add many task-resource assignments efficiently.

        items: list of {task_id, resource_id, [units, work]} dicts.
        Returns count added. T73 HERO path — single XML write at end.

        Performance target: 2800 assignments < 2 seconds (pure Python,
        no COM). Caller should call save() once after this.
        """
        # Pre-build resource_id -> resource_uid map (avoid O(N*M) lookups)
        rid_to_uid: Dict[int, int] = {}
        for r in self._resources.values():
            rid = r.get("id")
            uid = r.get("uid")
            if rid is not None and uid is not None:
                rid_to_uid[rid] = uid

        assignments_elem = self._find(self.root, "Assignments")
        if assignments_elem is None:
            assignments_elem = ET.SubElement(self.root, self._t("Assignments"))

        # Single pass — build all elements
        next_uid = self._next_assignment_uid()
        added = 0
        for item in items:
            task_id = item["task_id"]
            resource_id = item["resource_id"]
            task_uid = self._uid_by_id.get(task_id)
            resource_uid = rid_to_uid.get(resource_id)
            if task_uid is None or resource_uid is None:
                continue  # silently skip missing IDs in bulk
            task = self._tasks_by_id.get(task_id, {})
            start = task.get("start_raw") or "2026-01-01T08:00:00"
            finish = task.get("finish_raw") or start
            work_iso = task.get("duration_raw") or "PT8H0M0S"
            units = float(item.get("units", 1.0))

            a_elem = ET.Element(self._t("Assignment"))
            for tag, val in (
                ("UID", str(next_uid)),
                ("TaskUID", str(task_uid)),
                ("ResourceUID", str(resource_uid)),
                ("Units", f"{units:.2f}"),
                ("Work", work_iso),
                ("Start", start),
                ("Finish", finish),
            ):
                sub = ET.SubElement(a_elem, self._t(tag))
                sub.text = val
            assignments_elem.append(a_elem)

            self._assignments.append({
                "uid": next_uid,
                "task_id": task_id, "task_uid": task_uid,
                "task_name": task.get("name", ""),
                "resource_id": resource_id, "resource_uid": resource_uid,
                "resource_name": (self._resources.get(resource_uid) or {}).get("name", ""),
                "units": units,
                "work": task.get("duration", ""),
                "cost": 0.0,
            })
            next_uid += 1
            added += 1
        return added

    def update_task(self, task_id: int, name: str = None,
                    duration_str: str = None, percent_complete: float = None,
                    notes: str = None, start_date: str = None,
                    finish_date: str = None) -> dict:
        """Update task properties in the XML."""
        task = self._tasks_by_id.get(task_id)
        if not task:
            return {"error": f"Task ID {task_id} not found"}

        uid = task["uid"]
        elem = self._task_elems.get(uid)
        if elem is None:
            return {"error": f"Task element not found for UID {uid}"}

        changes = []

        if name is not None:
            self._set_text(elem, "Name", name)
            task["name"] = name
            changes.append(f"name -> '{name}'")

        if duration_str is not None:
            days = self._parse_duration_input(duration_str)
            iso = self._days_to_iso(days)
            self._set_text(elem, "Duration", iso)
            task["duration_raw"] = iso
            task["duration"] = f"{int(days)}d" if days == int(days) else f"{days:.1f}d"
            changes.append(f"duration -> {task['duration']}")

        if percent_complete is not None:
            self._set_text(elem, "PercentComplete", str(int(percent_complete)))
            task["percent_complete"] = int(percent_complete)
            changes.append(f"percent_complete -> {int(percent_complete)}%")

        if notes is not None:
            self._set_text(elem, "Notes", notes)
            task["notes"] = notes
            changes.append("notes updated")

        if start_date is not None:
            self._set_text(elem, "Start", f"{start_date}T08:00:00")
            task["start"] = start_date
            task["start_raw"] = f"{start_date}T08:00:00"
            changes.append(f"start -> {start_date}")

        if finish_date is not None:
            self._set_text(elem, "Finish", f"{finish_date}T17:00:00")
            task["finish"] = finish_date
            task["finish_raw"] = f"{finish_date}T17:00:00"
            changes.append(f"finish -> {finish_date}")

        return {
            "id": task_id,
            "updated_fields": changes,
            "changes": changes,
            "message": f"Task {task_id} updated: {', '.join(changes)}",
        }

    def delete_task(self, task_id: int) -> dict:
        """Delete a task from the XML."""
        task = self._tasks_by_id.get(task_id)
        if not task:
            return {"error": f"Task ID {task_id} not found"}

        uid = task["uid"]
        elem = self._task_elems.get(uid)
        if elem is None:
            return {"error": f"Task element not found for UID {uid}"}

        tasks_elem = self._find(self.root, "Tasks")
        if tasks_elem is not None:
            tasks_elem.remove(elem)

        # Also remove any PredecessorLink referencing this task's UID
        for other_uid, other_elem in self._task_elems.items():
            if other_uid == uid:
                continue
            for pred_elem in list(self._findall(other_elem, "PredecessorLink")):
                pred_uid = self._int(pred_elem, "PredecessorUID")
                if pred_uid == uid:
                    other_elem.remove(pred_elem)

        # Remove from indices
        task_name = task["name"]
        del self._tasks[uid]
        del self._task_elems[uid]
        if task_id in self._tasks_by_id:
            del self._tasks_by_id[task_id]
        if task_id in self._uid_by_id:
            del self._uid_by_id[task_id]
        if uid in self._id_by_uid:
            del self._id_by_uid[uid]

        # Rebuild successors (since we may have removed links)
        for t in self._tasks.values():
            t["successors"] = []
        self._build_successors()

        return {
            "deleted": True,
            "deleted_id": task_id,
            "deleted_name": task_name,
            "message": f"Task '{task_name}' (ID: {task_id}) deleted",
        }

    def add_link(self, predecessor_id: int, successor_id: int,
                 link_type: str = "FS", lag_str: str = None) -> dict:
        """Add a predecessor link."""
        pred_task = self._tasks_by_id.get(predecessor_id)
        succ_task = self._tasks_by_id.get(successor_id)
        if not pred_task:
            return {"error": f"Predecessor task ID {predecessor_id} not found"}
        if not succ_task:
            return {"error": f"Successor task ID {successor_id} not found"}

        pred_uid = pred_task["uid"]
        succ_uid = succ_task["uid"]
        type_id = self.LINK_TYPE_IDS.get(link_type.upper(), 1)

        lag_days = self._parse_duration_input(lag_str) if lag_str else 0
        lag_raw = self._days_to_lag(lag_days)

        # Add PredecessorLink element to successor task
        succ_elem = self._task_elems.get(succ_uid)
        if succ_elem is None:
            return {"error": f"Successor task element not found"}

        pred_link = ET.SubElement(succ_elem, self._t("PredecessorLink"))
        ET.SubElement(pred_link, self._t("PredecessorUID")).text = str(pred_uid)
        ET.SubElement(pred_link, self._t("Type")).text = str(type_id)
        ET.SubElement(pred_link, self._t("CrossProject")).text = "0"
        ET.SubElement(pred_link, self._t("LinkLag")).text = str(lag_raw)
        ET.SubElement(pred_link, self._t("LagFormat")).text = "7"

        # Update internal data
        pred_info = {
            "predecessor_uid": pred_uid,
            "type_id": type_id,
            "type": link_type.upper(),
            "lag_raw": lag_raw,
            "lag": self._format_lag(lag_raw),
            "lag_format": 7,
        }
        succ_task["predecessors"].append(pred_info)
        pred_task["successors"].append({
            "successor_uid": succ_uid,
            "successor_id": successor_id,
            "type": link_type.upper(),
            "lag": self._format_lag(lag_raw),
        })

        return {
            "success": True,
            "predecessor_id": predecessor_id,
            "predecessor_name": pred_task["name"],
            "successor_id": successor_id,
            "successor_name": succ_task["name"],
            "link_type": link_type.upper(),
            "lag": self._format_lag(lag_raw),
            "message": f"Link added: {pred_task['name']} -> {succ_task['name']} ({link_type.upper()}, lag={self._format_lag(lag_raw)})",
        }

    def remove_link(self, predecessor_id: int, successor_id: int) -> dict:
        """Remove a predecessor link."""
        pred_task = self._tasks_by_id.get(predecessor_id)
        succ_task = self._tasks_by_id.get(successor_id)
        if not pred_task:
            return {"error": f"Predecessor task ID {predecessor_id} not found"}
        if not succ_task:
            return {"error": f"Successor task ID {successor_id} not found"}

        pred_uid = pred_task["uid"]
        succ_uid = succ_task["uid"]
        succ_elem = self._task_elems.get(succ_uid)
        if succ_elem is None:
            return {"error": f"Successor task element not found"}

        # Find and remove PredecessorLink element
        removed = False
        for pred_elem in list(self._findall(succ_elem, "PredecessorLink")):
            if self._int(pred_elem, "PredecessorUID") == pred_uid:
                succ_elem.remove(pred_elem)
                removed = True
                break

        if not removed:
            return {"error": f"No link found from task {predecessor_id} to task {successor_id}"}

        # Update internal data
        succ_task["predecessors"] = [
            p for p in succ_task["predecessors"] if p["predecessor_uid"] != pred_uid
        ]
        pred_task["successors"] = [
            s for s in pred_task["successors"] if s["successor_uid"] != succ_uid
        ]

        return {
            "removed": True,
            "predecessor_id": predecessor_id,
            "successor_id": successor_id,
            "message": f"Link removed: {pred_task['name']} -> {succ_task['name']}",
        }

    def update_link(self, predecessor_id: int, successor_id: int,
                    new_link_type: str = None, new_lag_str: str = None) -> dict:
        """Update a link's type and/or lag."""
        pred_task = self._tasks_by_id.get(predecessor_id)
        succ_task = self._tasks_by_id.get(successor_id)
        if not pred_task:
            return {"error": f"Predecessor task ID {predecessor_id} not found"}
        if not succ_task:
            return {"error": f"Successor task ID {successor_id} not found"}

        pred_uid = pred_task["uid"]
        succ_uid = succ_task["uid"]
        succ_elem = self._task_elems.get(succ_uid)
        if succ_elem is None:
            return {"error": f"Successor task element not found"}

        # Find the PredecessorLink element
        target_elem = None
        for pred_elem in self._findall(succ_elem, "PredecessorLink"):
            if self._int(pred_elem, "PredecessorUID") == pred_uid:
                target_elem = pred_elem
                break

        if target_elem is None:
            return {"error": f"No link found from task {predecessor_id} to task {successor_id}"}

        old_type = self.LINK_TYPES.get(self._int(target_elem, "Type", 1), "FS")
        old_lag = self._format_lag(self._int(target_elem, "LinkLag", 0))

        if new_link_type:
            type_id = self.LINK_TYPE_IDS.get(new_link_type.upper(), 1)
            self._set_text(target_elem, "Type", str(type_id))
        if new_lag_str:
            lag_days = self._parse_duration_input(new_lag_str)
            lag_raw = self._days_to_lag(lag_days)
            self._set_text(target_elem, "LinkLag", str(lag_raw))

        # Rebuild predecessors/successors
        for t in self._tasks.values():
            t["predecessors"] = []
            t["successors"] = []
        # Re-parse predecessor links from XML
        for task_uid, task_elem in self._task_elems.items():
            task = self._tasks.get(task_uid)
            if not task:
                continue
            for pl in self._findall(task_elem, "PredecessorLink"):
                task["predecessors"].append({
                    "predecessor_uid": self._int(pl, "PredecessorUID"),
                    "type_id": self._int(pl, "Type", 1),
                    "type": self.LINK_TYPES.get(self._int(pl, "Type", 1), "FS"),
                    "lag_raw": self._int(pl, "LinkLag", 0),
                    "lag": self._format_lag(self._int(pl, "LinkLag", 0)),
                    "lag_format": self._int(pl, "LagFormat", 7),
                })
        self._build_successors()

        new_type_str = new_link_type.upper() if new_link_type else old_type
        new_lag_str_display = self._format_lag(self._days_to_lag(self._parse_duration_input(new_lag_str))) if new_lag_str else old_lag

        return {
            "updated": True,
            "predecessor_id": predecessor_id,
            "successor_id": successor_id,
            "old_type": old_type,
            "new_type": new_type_str,
            "old_lag": old_lag,
            "new_lag": new_lag_str_display,
            "message": f"Link updated: type {old_type}->{new_type_str}, lag {old_lag}->{new_lag_str_display}",
        }

    def update_progress(self, task_id: int, percent_complete: float = None,
                        actual_start: str = None, actual_finish: str = None) -> dict:
        """Update progress data for a task."""
        task = self._tasks_by_id.get(task_id)
        if not task:
            return {"error": f"Task ID {task_id} not found"}

        uid = task["uid"]
        elem = self._task_elems.get(uid)
        if elem is None:
            return {"error": f"Task element not found"}

        changes = []
        if percent_complete is not None:
            self._set_text(elem, "PercentComplete", str(int(percent_complete)))
            task["percent_complete"] = int(percent_complete)
            changes.append(f"percent_complete -> {int(percent_complete)}%")

        if actual_start:
            self._set_text(elem, "ActualStart", f"{actual_start}T08:00:00")
            task["actual_start"] = actual_start
            changes.append(f"actual_start -> {actual_start}")

        if actual_finish:
            self._set_text(elem, "ActualFinish", f"{actual_finish}T17:00:00")
            task["actual_finish"] = actual_finish
            changes.append(f"actual_finish -> {actual_finish}")

        return {
            "updated": True,
            "id": task_id,
            "name": task["name"],
            "changes": changes,
            "message": f"Progress updated: {', '.join(changes)}",
        }

    def bulk_update_progress(self, updates: List[dict]) -> dict:
        """Update progress for multiple tasks."""
        results = []
        for upd in updates:
            r = self.update_progress(
                task_id=upd.get("task_id"),
                percent_complete=upd.get("percent_complete"),
                actual_start=upd.get("actual_start"),
                actual_finish=upd.get("actual_finish"),
            )
            results.append(r)
        success_count = sum(1 for r in results if "error" not in r)
        return {
            "total": len(updates),
            "success": success_count,
            "failed": len(updates) - success_count,
            "details": results,
        }

    def add_summary_task(self, name: str, parent_task_id: int = None) -> dict:
        """Add a summary task."""
        return self.add_task(name, "0d", is_summary=True, parent_task_id=parent_task_id)

    def add_child_task(self, parent_task_id: int, name: str, duration_str: str = "1d") -> dict:
        """Add a child task under a parent."""
        result = self.add_task(name, duration_str, parent_task_id=parent_task_id)
        if "error" not in result:
            result["parent_id"] = parent_task_id
            parent = self._tasks_by_id.get(parent_task_id)
            if parent:
                result["parent_name"] = parent["name"]
        return result

    def assign_code(self, task_id: int, library_name: str, value: str) -> dict:
        """Assign a code library value to a task."""
        task = self._tasks_by_id.get(task_id)
        if not task:
            return {"error": f"Task ID {task_id} not found"}

        uid = task["uid"]
        elem = self._task_elems.get(uid)
        if elem is None:
            return {"error": f"Task element not found"}

        # Find the field_id for this library
        target_fid = None
        for fid, lib in self._code_libs.items():
            if lib["name"] == library_name:
                # Prefer task-level fields (188743xxx range)
                fid_int = int(fid) if fid.isdigit() else 0
                if fid_int < 200000000:
                    target_fid = fid
                    break
                if target_fid is None:
                    target_fid = fid

        if not target_fid:
            return {"error": f"Code library '{library_name}' not found"}

        # Check if already assigned - update or create
        existing = None
        for ea_elem in self._findall(elem, "ExtendedAttribute"):
            if self._text(ea_elem, "FieldID") == target_fid:
                existing = ea_elem
                break

        if existing is not None:
            self._set_text(existing, "Value", value)
        else:
            ea_elem = ET.SubElement(elem, self._t("ExtendedAttribute"))
            ET.SubElement(ea_elem, self._t("FieldID")).text = target_fid
            ET.SubElement(ea_elem, self._t("Value")).text = value

        # Update internal data
        task["codes"][library_name] = {
            "field_id": target_fid,
            "value": value,
            "value_id": "",
        }

        return {
            "success": True,
            "task_id": task_id,
            "library": library_name,
            "value": value,
            "message": f"Code '{library_name}' = '{value}' assigned to task {task_id}",
        }

    # ==================================================================
    # SAVE
    # ==================================================================

    def save(self, output_path: str = None) -> str:
        """Save the project to MSPDI XML format.

        Preserves full XML structure for lossless import into MS Project and Asta.
        """
        if not output_path:
            base = os.path.splitext(self.file_path)[0]
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"{base}_updated_{ts}.xml"
        output_path = output_path.replace("\\", "/")

        # Write XML
        self.tree.write(output_path, encoding="UTF-8", xml_declaration=True)

        # Post-process: ensure correct namespace and clean formatting
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # Remove any ns0: prefixes that ElementTree might add
            content = re.sub(r'<ns\d+:', '<', content)
            content = re.sub(r'</ns\d+:', '</', content)
            content = re.sub(r'\s+xmlns:ns\d+="[^"]*"', '', content)

            # Ensure the root element has the correct namespace
            if 'xmlns="http://schemas.microsoft.com/project"' not in content:
                content = content.replace(
                    '<Project',
                    '<Project xmlns="http://schemas.microsoft.com/project"',
                    1
                )

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)

        except Exception as e:
            logger.warning(f"Post-processing warning (file still saved): {e}")

        logger.info(f"Project saved: {output_path}")
        return output_path
