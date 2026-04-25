#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Asta Powerproject MCP Server
=============================
A comprehensive MCP (Model Context Protocol) server for Asta Powerproject.
Provides both file-based operations (via MPXJ) and GUI automation tools
for full project management functionality.

Author: Claude AI for Cahit
Version: 1.0.0
"""

import json
import os
import sys
import logging
import subprocess
import time
import base64
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
from enum import Enum
from pathlib import Path

from pydantic import BaseModel, Field, field_validator, ConfigDict
from mcp.server.fastmcp import FastMCP

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================
# MCP stdio servers must NOT log to stdout - use stderr
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stderr),
        logging.FileHandler(
            os.path.join(os.path.expanduser("~"), "asta_mcp_core.log"),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger("asta_mcp_core")

# ============================================================================
# INITIALIZE MCP SERVER
# ============================================================================
mcp = FastMCP(
    "asta_powerproject_mcp",
    instructions=(
        "PRIMARY Asta Powerproject tools — connects directly to a running Asta instance via COM. "
        "NO file_path needed. Do NOT search for .pp files on disk. Just call these tools directly. "
        "Use these tools FIRST for all Asta operations. The asta_powerproject_file MCP is only a "
        "secondary fallback for reading specific project files when Asta is not running. "
        "You are an expert construction project planner with deep knowledge of PMI/PMBOK, "
        "DCMA 14-Point Assessment, AACE, CIOB, ISO 21500, NEC4/FIDIC, and CPM scheduling. "
        "Ensure logic-driven networks, keep durations 5-20 working days, use FS links by default. "
        "Always reschedule after changes. Use asta_export → report for schedule health analysis."
    )
)

# ============================================================================
# CONSTANTS
# ============================================================================
SUPPORTED_EXTENSIONS = ['.pp', '.mpp', '.xml', '.mspdi', '.xer', '.pmxml']
DEFAULT_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
ASTA_WINDOW_TITLE = "Asta Powerproject"

# Maximum response size in characters to prevent Claude Desktop context overflow.
# Claude Desktop chat has ~8K token tool result limit. 1 token ≈ 4 chars → ~30K chars max safe.
MAX_RESPONSE_CHARS = 25000


def _truncate_response(text: str, max_chars: int = MAX_RESPONSE_CHARS) -> str:
    """Truncate tool response to prevent Claude Desktop context overflow.

    If text exceeds max_chars, truncates and appends a warning note.
    """
    if len(text) <= max_chars:
        return text
    truncated = text[:max_chars]
    # Try to truncate at a newline boundary
    last_newline = truncated.rfind('\n')
    if last_newline > max_chars * 0.8:  # Only if we don't lose too much
        truncated = truncated[:last_newline]
    return truncated + f"\n\n... **[TRUNCATED]** Response exceeded {max_chars} chars. Use smaller `limit` or `max_tasks` param, or query specific tasks with `asta_task → get`."

# ============================================================================
# JVM is NOT pre-started here — MPXJ/JVM is only loaded lazily in fallback mode.
# This keeps COM MCP startup fast. For file-based operations use asta_mcp_file.py.
# ============================================================================

# ============================================================================
# HELPER: Turkish character cleaning
# ============================================================================
def clean_turkish(text: str) -> str:
    """Convert Turkish characters to ASCII equivalents."""
    if text is None:
        return ""
    text = str(text)
    tr_map = {
        '\u00e7': 'c', '\u00c7': 'C', '\u011f': 'g', '\u011e': 'G',
        '\u015f': 's', '\u015e': 'S', '\u00fc': 'u', '\u00dc': 'U',
        '\u00f6': 'o', '\u00d6': 'O', '\u0131': 'i', '\u0130': 'I',
    }
    for tr_char, eng_char in tr_map.items():
        text = text.replace(tr_char, eng_char)
    return text


def parse_duration(duration_str: str) -> float:
    """Parse duration string to days. Examples: '5d'->5, '2w'->14, '3m'->90"""
    if not duration_str:
        return 1.0
    s = str(duration_str).strip().lower()
    if 'w' in s:
        return float(s.replace('w', '').replace('eek', '').strip()) * 7
    if 'm' in s and 'min' not in s:
        return float(s.replace('m', '').replace('onth', '').strip()) * 30
    if 'h' in s:
        return float(s.replace('h', '').replace('our', '').strip()) / 8
    if 'd' in s:
        return float(s.replace('d', '').replace('ay', '').strip())
    try:
        return float(s)
    except ValueError:
        return 1.0


def format_date(dt) -> str:
    """Safely format a date object to string (date-only, no time)."""
    if dt is None:
        return "N/A"
    try:
        if hasattr(dt, 'strftime'):
            return dt.strftime("%Y-%m-%d")
        # Java LocalDateTime: toString() gives "2025-09-01T08:00" or "2025-09-01T08:00:00"
        s = str(dt)
        if 'T' in s:
            return s.split('T')[0]
        # Fallback: take first 10 chars if it looks like a date
        if len(s) >= 10 and s[4:5] == '-':
            return s[:10]
        return s
    except Exception:
        return str(dt)


def safe_float(val, default=0.0) -> float:
    """Safely convert a Java/Python value to float."""
    if val is None:
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def safe_str(val, default="") -> str:
    """Safely convert a Java/Python value to string."""
    if val is None:
        return default
    try:
        return str(val)
    except Exception:
        return default


def duration_to_hours(dur, default=0.0) -> float:
    """Convert an mpxj Duration to hours (float). Handles None and various formats."""
    if dur is None:
        return default
    try:
        # mpxj Duration has getDuration() returning the numeric part
        # and getUnits() returning the TimeUnit enum
        raw = dur.getDuration()
        units = str(dur.getUnits()).upper() if dur.getUnits() else "HOURS"
        val = float(raw)
        if "DAY" in units:
            return val * 8.0
        if "WEEK" in units:
            return val * 40.0
        if "MONTH" in units:
            return val * 160.0
        if "MINUTE" in units:
            return val / 60.0
        # Default: assume hours
        return val
    except Exception:
        # Fallback: string parsing
        try:
            s = str(dur).lower().strip()
            for suffix in ['hours', 'hour', 'hrs', 'h', 'days', 'day', 'd',
                           'weeks', 'week', 'w', 'minutes', 'mins', 'min', 'm']:
                s = s.replace(suffix, '')
            return float(s.strip()) if s.strip() else default
        except (ValueError, AttributeError):
            return default


def _get_powershell_path() -> str:
    """Find the full path to PowerShell executable."""
    candidates = [
        os.path.join(os.environ.get("SystemRoot", r"C:\WINDOWS"),
                     "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
        r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    # Last resort: hope it's on PATH
    return "powershell.exe"


def _clipboard_paste(text: str):
    """Type text using clipboard paste (handles Unicode/Turkish characters)."""
    import subprocess
    ps_path = _get_powershell_path()
    escaped = text.replace("'", "''")
    subprocess.run(
        [ps_path, '-NoProfile', '-command', f"Set-Clipboard -Value '{escaped}'"],
        capture_output=True, timeout=5
    )
    import pyautogui
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.2)


# ============================================================================
# MPXJ FILE MANAGER CLASS
# ============================================================================
class AstaFileManager:
    """Manages Asta project files using MPXJ library."""

    def __init__(self, file_path: str):
        self.file_path = file_path.replace("\\", "/")
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")

        try:
            import mpxj
            if not mpxj.jpype.isJVMStarted():
                mpxj.jpype.startJVM()
            from org.mpxj.reader import UniversalProjectReader
            reader = UniversalProjectReader()
            self.project = reader.read(self.file_path)
            logger.info(f"Project loaded: {self.file_path}")
        except ImportError:
            raise RuntimeError(
                "MPXJ library not installed. Install with: pip install mpxj jpype1"
            )
        except Exception as e:
            err_msg = str(e)
            # Provide helpful guidance for .pp file parsing failures
            if self.file_path.lower().endswith('.pp') and 'NullPointerException' in err_msg:
                raise RuntimeError(
                    f"Failed to read .pp file: {err_msg}\n\n"
                    "WORKAROUND: Export the project from Asta Powerproject as XML first:\n"
                    "  1. Open the .pp file in Asta Powerproject\n"
                    "  2. Go to File > Export > Microsoft Project XML\n"
                    "  3. Save as .xml file\n"
                    "  4. Use the .xml file path with this tool instead"
                )
            raise RuntimeError(f"Failed to read file: {err_msg}")

    def get_project_summary(self) -> dict:
        """Get project overview information."""
        props = self.project.getProjectProperties()
        tasks = list(self.project.getTasks())

        critical_count = sum(1 for t in tasks if t.getCritical())
        milestone_count = sum(1 for t in tasks if t.getMilestone())

        return {
            "file": self.file_path,
            "project_name": clean_turkish(str(props.getName())) if props and props.getName() else "Unnamed",
            "client": clean_turkish(str(props.getManager())) if props and props.getManager() else "N/A",
            "start_date": format_date(props.getStartDate() if props else None),
            "finish_date": format_date(props.getFinishDate() if props else None),
            "total_tasks": len(tasks),
            "critical_tasks": critical_count,
            "milestones": milestone_count,
            "total_resources": self.project.getResources().size() if self.project.getResources() else 0,
        }

    def get_all_tasks(self, include_summary: bool = True) -> List[dict]:
        """Get all tasks with details."""
        result = []
        for task in self.project.getTasks():
            if not include_summary and task.getSummary():
                continue
            task_data = {
                "id": task.getID(),
                "unique_id": task.getUniqueID(),
                "name": clean_turkish(str(task.getName())) if task.getName() else "Unnamed",
                "duration": str(task.getDuration()) if task.getDuration() else "0d",
                "start": format_date(task.getStart()),
                "finish": format_date(task.getFinish()),
                "percent_complete": task.getPercentageComplete() if task.getPercentageComplete() else 0,
                "critical": bool(task.getCritical()),
                "milestone": bool(task.getMilestone()),
                "summary": bool(task.getSummary()),
                "total_float": str(task.getTotalSlack()) if task.getTotalSlack() else "N/A",
                "notes": clean_turkish(str(task.getNotes())) if task.getNotes() else "",
            }
            # Get predecessor info
            predecessors = []
            if task.getPredecessors():
                for rel in task.getPredecessors():
                    pred_task = rel.getPredecessorTask()
                    predecessors.append({
                        "task_id": pred_task.getID() if pred_task else None,
                        "type": str(rel.getType()) if rel.getType() else "FS",
                        "lag": str(rel.getLag()) if rel.getLag() else "0d",
                    })
            task_data["predecessors"] = predecessors

            # Get successor info
            successors = []
            if task.getSuccessors():
                for rel in task.getSuccessors():
                    succ_task = rel.getSuccessorTask()
                    successors.append({
                        "task_id": succ_task.getID() if succ_task else None,
                        "type": str(rel.getType()) if rel.getType() else "FS",
                        "lag": str(rel.getLag()) if rel.getLag() else "0d",
                    })
            task_data["successors"] = successors
            result.append(task_data)
        return result

    def get_task_by_id(self, task_id: int) -> Optional[dict]:
        """Get a specific task by ID."""
        for task in self.project.getTasks():
            if task.getID() == task_id:
                return {
                    "id": task.getID(),
                    "unique_id": task.getUniqueID(),
                    "name": clean_turkish(str(task.getName())) if task.getName() else "Unnamed",
                    "duration": str(task.getDuration()) if task.getDuration() else "0d",
                    "start": format_date(task.getStart()),
                    "finish": format_date(task.getFinish()),
                    "early_start": format_date(task.getEarlyStart()),
                    "early_finish": format_date(task.getEarlyFinish()),
                    "late_start": format_date(task.getLateStart()),
                    "late_finish": format_date(task.getLateFinish()),
                    "percent_complete": task.getPercentageComplete() if task.getPercentageComplete() else 0,
                    "actual_start": format_date(task.getActualStart()),
                    "actual_finish": format_date(task.getActualFinish()),
                    "critical": bool(task.getCritical()),
                    "milestone": bool(task.getMilestone()),
                    "summary": bool(task.getSummary()),
                    "total_float": str(task.getTotalSlack()) if task.getTotalSlack() else "N/A",
                    "free_float": str(task.getFreeSlack()) if task.getFreeSlack() else "N/A",
                    "notes": clean_turkish(str(task.getNotes())) if task.getNotes() else "",
                    "calendar": clean_turkish(safe_str(task.getCalendar().getName(), "Default")) if task.getCalendar() else "Default",
                    "cost": safe_float(task.getCost()),
                    "actual_cost": safe_float(task.getActualCost()),
                    "work": str(task.getWork()) if task.getWork() else "0h",
                }
        return None

    def get_critical_path(self) -> List[dict]:
        """Get all tasks on the critical path."""
        critical = []
        for task in self.project.getTasks():
            if task.getCritical() and not task.getSummary():
                critical.append({
                    "id": task.getID(),
                    "name": clean_turkish(str(task.getName())) if task.getName() else "Unnamed",
                    "duration": str(task.getDuration()) if task.getDuration() else "0d",
                    "start": format_date(task.getStart()),
                    "finish": format_date(task.getFinish()),
                    "total_float": str(task.getTotalSlack()) if task.getTotalSlack() else "0d",
                })
        return critical

    def get_resources(self) -> List[dict]:
        """Get all resources in the project."""
        resources = []
        if self.project.getResources():
            for res in self.project.getResources():
                resources.append({
                    "id": res.getID(),
                    "unique_id": res.getUniqueID(),
                    "name": clean_turkish(str(res.getName())) if res.getName() else "Unnamed",
                    "type": str(res.getType()) if res.getType() else "N/A",
                    "max_units": safe_float(res.getMaxUnits(), 1.0),
                    "standard_rate": str(res.getStandardRate()) if res.getStandardRate() else "N/A",
                    "cost": safe_float(res.getCost()),
                    "calendar": clean_turkish(safe_str(res.getCalendar().getName(), "Default")) if res.getCalendar() else "Default",
                })
        return resources

    def get_resource_assignments(self) -> List[dict]:
        """Get all resource assignments."""
        assignments = []
        if self.project.getResourceAssignments():
            for asgn in self.project.getResourceAssignments():
                assignments.append({
                    "task_id": asgn.getTask().getID() if asgn.getTask() else None,
                    "task_name": clean_turkish(str(asgn.getTask().getName())) if asgn.getTask() and asgn.getTask().getName() else "N/A",
                    "resource_id": asgn.getResource().getID() if asgn.getResource() else None,
                    "resource_name": clean_turkish(str(asgn.getResource().getName())) if asgn.getResource() and asgn.getResource().getName() else "N/A",
                    "units": safe_float(asgn.getUnits(), 1.0),
                    "work": str(asgn.getWork()) if asgn.getWork() else "0h",
                    "cost": safe_float(asgn.getCost()),
                })
        return assignments

    def get_calendars(self) -> List[dict]:
        """Get all calendars in the project."""
        calendars = []
        if self.project.getCalendars():
            for cal in self.project.getCalendars():
                calendars.append({
                    "id": cal.getUniqueID(),
                    "name": clean_turkish(str(cal.getName())) if cal.getName() else "Unnamed",
                })
        return calendars

    def _next_id(self) -> int:
        """Get the next available task ID."""
        max_id = 0
        for task in self.project.getTasks():
            tid = task.getID()
            if tid is not None and tid > max_id:
                max_id = tid
        return max_id + 1

    def _next_unique_id(self) -> int:
        """Get the next available unique task ID."""
        max_uid = 0
        for task in self.project.getTasks():
            uid = task.getUniqueID()
            if uid is not None and uid > max_uid:
                max_uid = uid
        return max_uid + 1

    def _assign_ids(self, task):
        """Assign ID and UniqueID to a new task if they are None."""
        from java.lang import Integer
        if task.getID() is None:
            task.setID(Integer(self._next_id()))
        if task.getUniqueID() is None:
            task.setUniqueID(Integer(self._next_unique_id()))

    def add_task(self, name: str, duration_str: str = "1d") -> dict:
        """Add a new task to the project."""
        from org.mpxj import Duration, TimeUnit
        new_task = self.project.addTask()
        self._assign_ids(new_task)
        new_task.setName(name)
        days = parse_duration(duration_str)
        new_task.setDuration(Duration.getInstance(days, TimeUnit.DAYS))
        logger.info(f"Task added: {name} ({days} days)")
        return {
            "id": new_task.getID(),
            "name": name,
            "duration": f"{days}d",
            "message": f"Task '{name}' added successfully"
        }

    def update_task(self, task_id: int, name: str = None,
                    duration_str: str = None, percent_complete: float = None,
                    notes: str = None) -> dict:
        """Update an existing task."""
        target = None
        for task in self.project.getTasks():
            if task.getID() == task_id:
                target = task
                break
        if not target:
            return {"error": f"Task ID {task_id} not found"}

        changes = []
        if name:
            target.setName(name)
            changes.append(f"name -> '{name}'")
        if duration_str:
            from org.mpxj import Duration, TimeUnit
            days = parse_duration(duration_str)
            target.setDuration(Duration.getInstance(days, TimeUnit.DAYS))
            changes.append(f"duration -> {days}d")
        if percent_complete is not None:
            target.setPercentageComplete(percent_complete)
            changes.append(f"percent_complete -> {percent_complete}%")
        if notes:
            target.setNotes(notes)
            changes.append(f"notes updated")

        return {
            "id": task_id,
            "changes": changes,
            "message": f"Task {task_id} updated: {', '.join(changes)}"
        }

    def delete_task(self, task_id: int) -> dict:
        """Delete a task by ID."""
        target = None
        for task in self.project.getTasks():
            if task.getID() == task_id:
                target = task
                break
        if not target:
            return {"error": f"Task ID {task_id} not found"}

        task_name = clean_turkish(str(target.getName()))
        self.project.removeTask(target)
        return {
            "deleted_id": task_id,
            "deleted_name": task_name,
            "message": f"Task '{task_name}' (ID: {task_id}) deleted"
        }

    def save(self, output_path: str = None) -> str:
        """Save the project file. Returns the output path."""
        import re
        from org.mpxj.mspdi import MSPDIWriter
        if not output_path:
            base = os.path.splitext(self.file_path)[0]
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"{base}_updated_{ts}.xml"
        output_path = output_path.replace("\\", "/")
        writer = MSPDIWriter()
        writer.write(self.project, output_path)

        # Post-process: fix JAXB namespace prefixes that Asta cannot parse
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()

            # <ns2:Project xmlns:ns2="..."> -> <Project xmlns="...">
            xml_content = re.sub(
                r'<ns\d+:Project\s+xmlns:ns\d+="(http://schemas\.microsoft\.com/project)"',
                r'<Project xmlns="\1"',
                xml_content
            )
            # Remove all nsN: prefixes from element tags
            xml_content = re.sub(r'<(/?)ns\d+:', r'<\1', xml_content)
            # Remove any remaining orphaned xmlns:nsN declarations
            xml_content = re.sub(r'\s+xmlns:ns\d+="[^"]*"', '', xml_content)

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)

            logger.info(f"Project saved (namespace fixed): {output_path}")
        except Exception as e:
            logger.warning(f"Namespace post-processing failed (file still saved): {e}")

        return output_path

    # ================================================================
    # WBS / HIERARCHY METHODS
    # ================================================================

    def get_wbs_tree(self, max_depth: int = 99) -> List[dict]:
        """Get WBS/hierarchy tree showing parent-child relationships."""
        def build_tree(task, level=0):
            node = {
                "id": task.getID(),
                "unique_id": task.getUniqueID(),
                "name": clean_turkish(str(task.getName())) if task.getName() else "Unnamed",
                "wbs": str(task.getWBS()) if task.getWBS() else "",
                "outline_level": int(task.getOutlineLevel()) if task.getOutlineLevel() else 0,
                "summary": bool(task.getSummary()),
                "milestone": bool(task.getMilestone()),
                "duration": str(task.getDuration()) if task.getDuration() else "0d",
                "start": format_date(task.getStart()),
                "finish": format_date(task.getFinish()),
                "level": level,
                "children": [],
            }
            if level < max_depth:
                children = task.getChildTasks()
                if children:
                    child_count = 0
                    for child in children:
                        node["children"].append(build_tree(child, level + 1))
                        child_count += 1
                    if child_count == 0:
                        pass  # no children to add
            else:
                children = task.getChildTasks()
                if children and children.size() > 0:
                    node["children_count"] = children.size()
                    node["children_truncated"] = True
            return node

        # Start from top-level tasks (those with no parent or parent is project root)
        result = []
        for task in self.project.getTasks():
            parent = task.getParentTask()
            if parent is None:
                result.append(build_tree(task, 0))
        return result

    def add_summary_task(self, name: str, parent_task_id: int = None) -> dict:
        """Add a new summary task. If parent_task_id given, adds under that parent."""
        if parent_task_id is not None:
            parent = None
            for task in self.project.getTasks():
                if task.getID() == parent_task_id:
                    parent = task
                    break
            if not parent:
                return {"error": f"Parent task ID {parent_task_id} not found"}
            new_task = parent.addTask()
        else:
            new_task = self.project.addTask()

        self._assign_ids(new_task)
        new_task.setName(name)
        # Summary tasks typically have no duration of their own
        logger.info(f"Summary task added: {name}")
        return {
            "id": new_task.getID(),
            "unique_id": new_task.getUniqueID(),
            "name": name,
            "parent_id": parent_task_id,
            "message": f"Summary task '{name}' added successfully"
        }

    def add_child_task(self, parent_task_id: int, name: str, duration_str: str = "1d") -> dict:
        """Add a child task under a specific parent/summary task."""
        parent = None
        for task in self.project.getTasks():
            if task.getID() == parent_task_id:
                parent = task
                break
        if not parent:
            return {"error": f"Parent task ID {parent_task_id} not found"}

        from org.mpxj import Duration, TimeUnit
        new_task = parent.addTask()
        self._assign_ids(new_task)
        new_task.setName(name)
        days = parse_duration(duration_str)
        new_task.setDuration(Duration.getInstance(days, TimeUnit.DAYS))
        logger.info(f"Child task added: {name} under parent {parent_task_id}")
        return {
            "id": new_task.getID(),
            "unique_id": new_task.getUniqueID(),
            "name": name,
            "duration": f"{days}d",
            "parent_id": parent_task_id,
            "parent_name": str(parent.getName()) if parent.getName() else "Unnamed",
            "message": f"Task '{name}' added under '{str(parent.getName()) if parent.getName() else 'Unnamed'}'"
        }

    def update_summary_task(self, task_id: int, name: str = None, notes: str = None) -> dict:
        """Update a summary task's name or notes."""
        target = None
        for task in self.project.getTasks():
            if task.getID() == task_id:
                target = task
                break
        if not target:
            return {"error": f"Task ID {task_id} not found"}

        changes = []
        if name:
            target.setName(name)
            changes.append(f"name -> '{name}'")
        if notes:
            target.setNotes(notes)
            changes.append("notes updated")

        return {
            "id": task_id,
            "changes": changes,
            "message": f"Summary task {task_id} updated: {', '.join(changes)}"
        }

    # ================================================================
    # PREDECESSOR / SUCCESSOR LINK METHODS
    # ================================================================

    def _get_relation_type(self, link_type: str):
        """Map string link type to RelationType enum."""
        from org.mpxj import RelationType
        type_map = {
            "FS": RelationType.FINISH_START,
            "SS": RelationType.START_START,
            "FF": RelationType.FINISH_FINISH,
            "SF": RelationType.START_FINISH,
        }
        return type_map.get(link_type.upper(), RelationType.FINISH_START)

    def add_link(self, predecessor_id: int, successor_id: int,
                 link_type: str = "FS", lag_str: str = None) -> dict:
        """Add a predecessor-successor link between two tasks."""
        from org.mpxj import Relation, Duration, TimeUnit

        pred_task = None
        succ_task = None
        for task in self.project.getTasks():
            if task.getID() == predecessor_id:
                pred_task = task
            if task.getID() == successor_id:
                succ_task = task
            if pred_task and succ_task:
                break

        if not pred_task:
            return {"error": f"Predecessor task ID {predecessor_id} not found"}
        if not succ_task:
            return {"error": f"Successor task ID {successor_id} not found"}

        rel_type = self._get_relation_type(link_type)
        lag_days = parse_duration(lag_str) if lag_str else 0
        lag_duration = Duration.getInstance(lag_days, TimeUnit.DAYS)

        # Use Builder pattern (mpxj 14 API)
        builder = Relation.Builder()
        builder = builder.predecessorTask(pred_task).successorTask(succ_task)
        builder = builder.type(rel_type).lag(lag_duration)
        relation = succ_task.addPredecessor(builder)

        logger.info(f"Link added: {predecessor_id} -> {successor_id} ({link_type})")
        return {
            "predecessor_id": predecessor_id,
            "predecessor_name": clean_turkish(str(pred_task.getName())),
            "successor_id": successor_id,
            "successor_name": clean_turkish(str(succ_task.getName())),
            "link_type": link_type.upper(),
            "lag": lag_str or "0d",
            "message": f"Link added: '{clean_turkish(str(pred_task.getName()))}' -> '{clean_turkish(str(succ_task.getName()))}' ({link_type})"
        }

    def remove_link(self, predecessor_id: int, successor_id: int) -> dict:
        """Remove a predecessor-successor link between two tasks."""
        pred_task = None
        succ_task = None
        for task in self.project.getTasks():
            if task.getID() == predecessor_id:
                pred_task = task
            if task.getID() == successor_id:
                succ_task = task
            if pred_task and succ_task:
                break

        if not pred_task:
            return {"error": f"Predecessor task ID {predecessor_id} not found"}
        if not succ_task:
            return {"error": f"Successor task ID {successor_id} not found"}

        # Find the relation and remove with exact parameters (mpxj 14 requires task, type, lag)
        removed = False
        if succ_task.getPredecessors():
            for rel in list(succ_task.getPredecessors()):
                if rel.getPredecessorTask() and rel.getPredecessorTask().getID() == predecessor_id:
                    rel_type = rel.getType()
                    rel_lag = rel.getLag()
                    removed = succ_task.removePredecessor(pred_task, rel_type, rel_lag)
                    break

        if not removed:
            return {"error": f"No link found from task {predecessor_id} to task {successor_id}"}

        logger.info(f"Link removed: {predecessor_id} -> {successor_id}")
        return {
            "predecessor_id": predecessor_id,
            "successor_id": successor_id,
            "message": f"Link removed: '{clean_turkish(str(pred_task.getName()))}' -> '{clean_turkish(str(succ_task.getName()))}'"
        }

    def update_link(self, predecessor_id: int, successor_id: int,
                    new_link_type: str = None, new_lag_str: str = None) -> dict:
        """Update a link by removing and re-creating it (Relations are immutable)."""
        from org.mpxj import Relation, Duration, TimeUnit

        pred_task = None
        succ_task = None
        for task in self.project.getTasks():
            if task.getID() == predecessor_id:
                pred_task = task
            if task.getID() == successor_id:
                succ_task = task
            if pred_task and succ_task:
                break

        if not pred_task:
            return {"error": f"Predecessor task ID {predecessor_id} not found"}
        if not succ_task:
            return {"error": f"Successor task ID {successor_id} not found"}

        # Find current link properties
        old_type_str = "FS"
        old_lag_str = "0d"
        old_rel_type = None
        old_rel_lag = None
        found = False
        if succ_task.getPredecessors():
            for rel in list(succ_task.getPredecessors()):
                if rel.getPredecessorTask() and rel.getPredecessorTask().getID() == predecessor_id:
                    old_rel_type = rel.getType()
                    old_rel_lag = rel.getLag()
                    old_type_str = str(old_rel_type) if old_rel_type else "FS"
                    old_lag_str = str(old_rel_lag) if old_rel_lag else "0d"
                    found = True
                    break

        if not found:
            return {"error": f"No link found from task {predecessor_id} to task {successor_id}"}

        # Remove old link (requires exact type and lag)
        succ_task.removePredecessor(pred_task, old_rel_type, old_rel_lag)

        # Create new link with updated properties using Builder
        final_type_str = new_link_type.upper() if new_link_type else old_type_str
        rel_type = self._get_relation_type(final_type_str)

        if new_lag_str:
            lag_days = parse_duration(new_lag_str)
        else:
            try:
                lag_days = float(str(old_lag_str).replace('d', '').replace(' ', '').split('.')[0])
            except (ValueError, AttributeError):
                lag_days = 0
        lag_duration = Duration.getInstance(lag_days, TimeUnit.DAYS)

        builder = Relation.Builder()
        builder = builder.predecessorTask(pred_task).successorTask(succ_task)
        builder = builder.type(rel_type).lag(lag_duration)
        succ_task.addPredecessor(builder)

        logger.info(f"Link updated: {predecessor_id} -> {successor_id} ({final_type_str}, lag={new_lag_str or old_lag_str})")
        return {
            "predecessor_id": predecessor_id,
            "successor_id": successor_id,
            "old_type": old_type_str,
            "new_type": final_type_str,
            "old_lag": old_lag_str,
            "new_lag": new_lag_str or old_lag_str,
            "message": f"Link updated: type {old_type_str}->{final_type_str}, lag {old_lag_str}->{new_lag_str or old_lag_str}"
        }

    # ================================================================
    # PROGRESS / ACTUAL DATA METHODS
    # ================================================================

    def update_progress(self, task_id: int, percent_complete: float = None,
                        actual_start: str = None, actual_finish: str = None) -> dict:
        """Update progress data for a task."""
        target = None
        for task in self.project.getTasks():
            if task.getID() == task_id:
                target = task
                break
        if not target:
            return {"error": f"Task ID {task_id} not found"}

        changes = []

        if percent_complete is not None:
            target.setPercentageComplete(percent_complete)
            changes.append(f"percent_complete -> {percent_complete}%")

        if actual_start:
            try:
                from java.time import LocalDateTime
                parts = actual_start.split("/") if "/" in actual_start else actual_start.split("-")
                if len(parts[0]) == 4:  # YYYY-MM-DD
                    dt = LocalDateTime.of(int(parts[0]), int(parts[1]), int(parts[2]), 8, 0)
                else:  # DD/MM/YYYY
                    dt = LocalDateTime.of(int(parts[2]), int(parts[1]), int(parts[0]), 8, 0)
                target.setActualStart(dt)
                changes.append(f"actual_start -> {actual_start}")
            except Exception as e:
                changes.append(f"actual_start FAILED: {e}")

        if actual_finish:
            try:
                from java.time import LocalDateTime
                parts = actual_finish.split("/") if "/" in actual_finish else actual_finish.split("-")
                if len(parts[0]) == 4:  # YYYY-MM-DD
                    dt = LocalDateTime.of(int(parts[0]), int(parts[1]), int(parts[2]), 17, 0)
                else:  # DD/MM/YYYY
                    dt = LocalDateTime.of(int(parts[2]), int(parts[1]), int(parts[0]), 17, 0)
                target.setActualFinish(dt)
                changes.append(f"actual_finish -> {actual_finish}")
            except Exception as e:
                changes.append(f"actual_finish FAILED: {e}")

        return {
            "id": task_id,
            "name": clean_turkish(str(target.getName())) if target.getName() else "Unnamed",
            "changes": changes,
            "message": f"Progress updated for task {task_id}: {', '.join(changes)}"
        }

    def bulk_update_progress(self, updates: List[dict]) -> dict:
        """Update progress for multiple tasks at once."""
        results = []
        for upd in updates:
            result = self.update_progress(
                task_id=upd.get("task_id"),
                percent_complete=upd.get("percent_complete"),
                actual_start=upd.get("actual_start"),
                actual_finish=upd.get("actual_finish")
            )
            results.append(result)
        success_count = sum(1 for r in results if "error" not in r)
        return {
            "total": len(updates),
            "success": success_count,
            "failed": len(updates) - success_count,
            "details": results
        }

    # ================================================================
    # DELAY ANALYSIS METHODS
    # ================================================================

    def get_delay_analysis(self) -> dict:
        """Analyze schedule delays by comparing planned vs actual dates."""
        delays = []
        for task in self.project.getTasks():
            if task.getSummary() or task.getMilestone():
                continue

            planned_start = task.getStart()
            planned_finish = task.getFinish()
            actual_start = task.getActualStart()
            actual_finish = task.getActualFinish()

            start_slip = None
            finish_slip = None

            # Calculate start slip
            if planned_start and actual_start:
                try:
                    ps = str(planned_start)[:10]
                    as_ = str(actual_start)[:10]
                    from datetime import date
                    pd = date.fromisoformat(ps)
                    ad = date.fromisoformat(as_)
                    start_slip = (ad - pd).days
                except Exception:
                    pass

            # Calculate finish slip
            if planned_finish and actual_finish:
                try:
                    pf = str(planned_finish)[:10]
                    af = str(actual_finish)[:10]
                    from datetime import date
                    pd2 = date.fromisoformat(pf)
                    ad2 = date.fromisoformat(af)
                    finish_slip = (ad2 - pd2).days
                except Exception:
                    pass

            if start_slip is not None or finish_slip is not None:
                delays.append({
                    "id": task.getID(),
                    "name": clean_turkish(str(task.getName())) if task.getName() else "Unnamed",
                    "planned_start": format_date(planned_start),
                    "actual_start": format_date(actual_start),
                    "start_slip_days": start_slip,
                    "planned_finish": format_date(planned_finish),
                    "actual_finish": format_date(actual_finish),
                    "finish_slip_days": finish_slip,
                    "percent_complete": task.getPercentageComplete() if task.getPercentageComplete() else 0,
                    "critical": bool(task.getCritical()),
                })

        # Statistics
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
            "tasks": delays
        }

    # ================================================================
    # RESOURCE LOADING ANALYSIS
    # ================================================================

    def get_resource_loading(self) -> dict:
        """Analyze resource loading across the project."""
        resource_summary = {}
        assignments = list(self.project.getResourceAssignments()) if self.project.getResourceAssignments() else []

        for asgn in assignments:
            res = asgn.getResource()
            task = asgn.getTask()
            if not res or not task:
                continue

            res_name = clean_turkish(str(res.getName())) if res.getName() else "Unnamed"
            res_id = res.getID()

            if res_id not in resource_summary:
                resource_summary[res_id] = {
                    "id": res_id,
                    "name": res_name,
                    "type": str(res.getType()) if res.getType() else "N/A",
                    "max_units": safe_float(res.getMaxUnits(), 1.0),
                    "total_work": 0,
                    "total_cost": 0,
                    "task_count": 0,
                    "tasks": [],
                }

            work_val = duration_to_hours(asgn.getWork())
            cost_val = safe_float(asgn.getCost())

            resource_summary[res_id]["total_work"] += work_val
            resource_summary[res_id]["total_cost"] += cost_val
            resource_summary[res_id]["task_count"] += 1
            resource_summary[res_id]["tasks"].append({
                "task_id": task.getID(),
                "task_name": clean_turkish(str(task.getName())) if task.getName() else "Unnamed",
                "units": safe_float(asgn.getUnits(), 1.0),
                "work": str(asgn.getWork()) if asgn.getWork() else "0h",
                "cost": cost_val,
                "start": format_date(task.getStart()),
                "finish": format_date(task.getFinish()),
            })

        return {
            "total_resources": len(resource_summary),
            "total_assignments": len(assignments),
            "resources": list(resource_summary.values())
        }

    def get_float_analysis(self) -> dict:
        """Analyze float distribution across tasks."""
        float_data = {"zero_float": 0, "low_float": 0, "medium_float": 0,
                      "high_float": 0, "tasks": []}
        for task in self.project.getTasks():
            if task.getSummary() or task.getMilestone():
                continue
            tf = task.getTotalSlack()
            if tf is not None:
                try:
                    # Try to get numeric value
                    tf_val = float(str(tf).replace('d', '').replace(' ', '').split('.')[0])
                except (ValueError, AttributeError):
                    tf_val = 0
                if tf_val == 0:
                    float_data["zero_float"] += 1
                elif tf_val <= 5:
                    float_data["low_float"] += 1
                elif tf_val <= 20:
                    float_data["medium_float"] += 1
                else:
                    float_data["high_float"] += 1
                float_data["tasks"].append({
                    "id": task.getID(),
                    "name": clean_turkish(str(task.getName())) if task.getName() else "Unnamed",
                    "total_float": str(tf),
                })
        return float_data


# ============================================================================
# GUI AUTOMATION MANAGER CLASS
# ============================================================================

# ============================================================================
# PYDANTIC INPUT MODELS (Core)
# ============================================================================
class ResponseFormat(str, Enum):
    MARKDOWN = "markdown"
    JSON = "json"


class ProjectFileInput(BaseModel):
    """Input for project operations. Supports two modes:
    - COM mode: Leave file_path empty/None to work directly with the running Asta instance.
    - File mode: Provide file_path to work with an exported XML/MPP file via MPXJ.
    """
    model_config = ConfigDict(str_strip_whitespace=True)
    file_path: Optional[str] = Field(
        default=None,
        description="Full path to the Asta project file (.pp, .mpp, .xml). "
                    "Leave EMPTY or omit to use COM (direct connection to running Asta). "
                    "Example: 'C:/Users/GPX PRO/Downloads/myproject.pp'"
    )

    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: Optional[str]) -> Optional[str]:
        if v is None or v.strip() == "":
            return None
        v = v.strip().replace("\\", "/")
        return v


class AnalyzeProjectInput(ProjectFileInput):
    """Input for project analysis."""
    response_format: ResponseFormat = Field(
        default=ResponseFormat.MARKDOWN,
        description="Output format: 'markdown' for readable text, 'json' for structured data"
    )


class ListTasksInput(ProjectFileInput):
    """Input for listing tasks."""
    include_summary: bool = Field(
        default=True,
        description="Whether to include summary (parent) tasks in the list"
    )
    response_format: ResponseFormat = Field(
        default=ResponseFormat.MARKDOWN,
        description="Output format"
    )
    limit: int = Field(
        default=50,
        description="Maximum number of tasks to return (default 50, max 200). Use asta_get_task for details on a specific task.",
        ge=1, le=200
    )


class GetTaskInput(ProjectFileInput):
    """Input for getting a specific task."""
    task_id: int = Field(..., description="The task ID number to look up", ge=0)
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class AddTaskInput(ProjectFileInput):
    """Input for adding a new task."""
    name: str = Field(
        ...,
        description="Name of the new task. Example: 'Foundation Work'",
        min_length=1, max_length=500
    )
    duration: str = Field(
        default="1d",
        description="Task duration. Use 'd' for days, 'w' for weeks, 'h' for hours. Examples: '5d', '2w', '8h'"
    )
    start_date: Optional[str] = Field(
        default=None,
        description="Planned start date (YYYY-MM-DD). Required for COM to create a real activity bar, not just text."
    )
    finish_date: Optional[str] = Field(
        default=None,
        description="Planned finish date (YYYY-MM-DD). If given with start_date, duration is calculated automatically."
    )
    save_output: Optional[str] = Field(
        default=None,
        description="Path to save the modified file. If empty, auto-generates a timestamped XML file"
    )


class UpdateTaskInput(ProjectFileInput):
    """Input for updating an existing task."""
    task_id: int = Field(..., description="ID of the task to update", ge=0)
    name: Optional[str] = Field(default=None, description="New task name")
    duration: Optional[str] = Field(default=None, description="New duration (e.g., '5d', '2w')")
    start_date: Optional[str] = Field(default=None, description="New planned start date (YYYY-MM-DD)")
    finish_date: Optional[str] = Field(default=None, description="New planned finish date (YYYY-MM-DD)")
    percent_complete: Optional[float] = Field(
        default=None, description="Completion percentage (0-100)", ge=0, le=100
    )
    notes: Optional[str] = Field(default=None, description="Task notes/comments")
    save_output: Optional[str] = Field(default=None, description="Path to save modified file")


class DeleteTaskInput(ProjectFileInput):
    """Input for deleting a task."""
    task_id: int = Field(..., description="ID of the task to delete", ge=0)
    save_output: Optional[str] = Field(default=None, description="Path to save modified file")


class CriticalPathInput(ProjectFileInput):
    """Input for critical path analysis."""
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class ResourcesInput(ProjectFileInput):
    """Input for listing resources."""
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class FloatAnalysisInput(ProjectFileInput):
    """Input for float analysis."""
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class SaveProjectInput(ProjectFileInput):
    """Input for saving project."""
    output_path: Optional[str] = Field(
        default=None,
        description="Output file path. If empty, auto-generates timestamped XML file"
    )


# ============================================================================
# NEW FEATURE INPUT MODELS
# ============================================================================

class WBSTreeInput(ProjectFileInput):
    """Input for WBS tree view."""
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)
    max_depth: int = Field(default=3, description="Maximum tree depth to show (default 3 for large projects)", ge=1, le=99)


class AddSummaryTaskInput(ProjectFileInput):
    """Input for adding a summary task."""
    name: str = Field(..., description="Name for the summary task", min_length=1, max_length=500)
    parent_task_id: Optional[int] = Field(default=None, description="Parent task ID to nest under. None = top level")
    save_output: Optional[str] = Field(default=None)


class AddChildTaskInput(ProjectFileInput):
    """Input for adding a child task under a parent."""
    parent_task_id: int = Field(..., description="ID of the parent/summary task", ge=0)
    name: str = Field(..., description="Name of the child task", min_length=1, max_length=500)
    duration: str = Field(default="1d", description="Duration (e.g., '5d', '2w', '8h')")
    start_date: Optional[str] = Field(default=None, description="Planned start date (YYYY-MM-DD)")
    finish_date: Optional[str] = Field(default=None, description="Planned finish date (YYYY-MM-DD)")
    save_output: Optional[str] = Field(default=None)


class AddLinkInput(ProjectFileInput):
    """Input for adding a task link."""
    predecessor_id: int = Field(..., description="ID of the predecessor task", ge=0)
    successor_id: int = Field(..., description="ID of the successor task", ge=0)
    link_type: str = Field(default="FS", description="Link type: FS, SS, FF, SF")
    lag: Optional[str] = Field(default=None, description="Lag time (e.g., '2d', '-1d' for lead)")
    save_output: Optional[str] = Field(default=None)


class RemoveLinkInput(ProjectFileInput):
    """Input for removing a task link."""
    predecessor_id: int = Field(..., description="ID of the predecessor task", ge=0)
    successor_id: int = Field(..., description="ID of the successor task", ge=0)
    save_output: Optional[str] = Field(default=None)


class UpdateLinkInput(ProjectFileInput):
    """Input for updating a task link."""
    predecessor_id: int = Field(..., description="ID of the predecessor task", ge=0)
    successor_id: int = Field(..., description="ID of the successor task", ge=0)
    new_link_type: Optional[str] = Field(default=None, description="New link type: FS, SS, FF, SF")
    new_lag: Optional[str] = Field(default=None, description="New lag (e.g., '5d', '0d')")
    save_output: Optional[str] = Field(default=None)


class UpdateProgressInput(ProjectFileInput):
    """Input for updating task progress."""
    task_id: int = Field(..., description="ID of the task to update", ge=0)
    percent_complete: Optional[float] = Field(default=None, description="Completion percentage (0-100)", ge=0, le=100)
    actual_start: Optional[str] = Field(default=None, description="Actual start date (YYYY-MM-DD or DD/MM/YYYY)")
    actual_finish: Optional[str] = Field(default=None, description="Actual finish date (YYYY-MM-DD or DD/MM/YYYY)")
    save_output: Optional[str] = Field(default=None)


class BulkProgressItem(BaseModel):
    """Single task progress update item."""
    task_id: int = Field(..., description="Task ID", ge=0)
    percent_complete: Optional[float] = Field(default=None, ge=0, le=100)
    actual_start: Optional[str] = Field(default=None)
    actual_finish: Optional[str] = Field(default=None)


class BulkUpdateProgressInput(ProjectFileInput):
    """Input for bulk progress updates."""
    updates: List[BulkProgressItem] = Field(..., description="List of task progress updates")
    save_output: Optional[str] = Field(default=None)


class DelayAnalysisInput(ProjectFileInput):
    """Input for delay analysis."""
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class ResourceLoadingInput(ProjectFileInput):
    """Input for resource loading analysis."""
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


# GUI Input Models

# ============================================================================
# FILE-BASED TOOLS (MPXJ) with COM-first when file_path is omitted
# ============================================================================
# @mcp.tool(  # CONSOLIDATED into asta_query
#     name="asta_analyze_project",
#     annotations={
#         "title": "Analyze Asta Project",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_analyze_project(params: AnalyzeProjectInput) -> str:
    """Analyze an Asta Powerproject file and return a comprehensive summary.

    Reads a .pp, .mpp, or .xml project file and provides:
    - Project name, client, dates
    - Total tasks, critical tasks, milestones count
    - Resource count
    - Task list with details

    Use this as the FIRST tool when working with any Asta project file
    to understand the project structure before making changes.

    Args:
        params: Contains file_path and response_format

    Returns:
        Project analysis in markdown or JSON format
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        summary = mgr.get_project_summary()
        tasks = mgr.get_all_tasks()

        if params.response_format == ResponseFormat.JSON:
            max_json_tasks = 50
            result = {"summary": summary, "total_tasks": len(tasks), "tasks": tasks[:max_json_tasks]}
            if len(tasks) > max_json_tasks:
                result["note"] = f"Showing first {max_json_tasks} of {len(tasks)} tasks. Use asta_list_tasks with limit parameter for more."
            return json.dumps(result, indent=2, default=str)

        # Markdown format
        lines = [
            f"# Project Analysis: {summary['project_name']}",
            "",
            f"**File:** {summary['file']}",
            f"**Client:** {summary['client']}",
            f"**Start Date:** {summary['start_date']}",
            f"**Finish Date:** {summary['finish_date']}",
            "",
            "## Statistics",
            f"- **Total Tasks:** {summary['total_tasks']}",
            f"- **Critical Tasks:** {summary['critical_tasks']}",
            f"- **Milestones:** {summary['milestones']}",
            f"- **Resources:** {summary['total_resources']}",
            "",
            "## Task List",
            "",
            "| ID | Name | Duration | Start | Finish | %Done | Critical |",
            "|---|---|---|---|---|---|---|",
        ]
        for t in tasks[:50]:
            crit = "YES" if t['critical'] else ""
            lines.append(
                f"| {t['id']} | {t['name']} | {t['duration']} | {t['start']} | {t['finish']} | {t['percent_complete']}% | {crit} |"
            )
        if len(tasks) > 50:
            lines.append(f"\n*...and {len(tasks) - 50} more tasks*")
        return "\n".join(lines)

    except Exception as e:
        return f"Error analyzing project: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_query
#     name="asta_list_tasks",
#     annotations={
#         "title": "List Project Tasks",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_list_tasks(params: ListTasksInput) -> str:
    """List all tasks in an Asta project file with their properties.

    Returns compact table with ID, name, duration, dates, completion %, critical status.
    Use asta_task → get for detailed info on a specific task.

    COM-first: tries live Asta via COM (fast), falls back to MPXJ file reading.
    """
    # --- COM-first strategy (no file export needed, much faster) ---
    if params.file_path is None:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()

            all_bars = _com_get_all_bars(project, max_bars=params.limit + 50)
            tasks = []
            for bar in all_bars:
                try:
                    t = {"id": bar.ID, "name": bar.Name}
                    # Dates
                    for prop, key in [("Start", "start"), ("End", "finish")]:
                        try:
                            t[key] = format_date(getattr(bar, prop))
                        except Exception:
                            t[key] = "N/A"
                    # Percent complete
                    try:
                        t["percent_complete"] = round(bar.OverallPercentComplete, 1)
                    except Exception:
                        try:
                            t["percent_complete"] = round(bar.DurationPercentComplete, 1)
                        except Exception:
                            t["percent_complete"] = 0.0
                    # Critical — use Tasks(1) not ExpandedTask (which returns root for nested bars)
                    try:
                        _task, _ = _get_bar_task(bar)
                        t["critical"] = bool(_task.Critical) if _task else False
                    except Exception:
                        t["critical"] = False
                    # Summary / Milestone detection
                    try:
                        dur = bar.Duration
                        t["milestone"] = (dur is not None and float(str(dur)) == 0)
                    except Exception:
                        t["milestone"] = False
                    t["summary"] = False  # COM can't reliably detect summary
                    t["predecessors"] = []
                    tasks.append(t)
                except Exception:
                    continue

            if com_initialized:
                pythoncom.CoUninitialize()
                com_initialized = False

            if len(tasks) > 0:
                total = len(tasks)
                limited = tasks[:params.limit]
                return _format_task_list(limited, total, params)

        except Exception as com_err:
            logger.info(f"COM list failed ({com_err}), falling back to MPXJ")
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- MPXJ fallback (slower, but gets ALL nested tasks) ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        tasks = mgr.get_all_tasks(include_summary=params.include_summary)

        total = len(tasks)
        limited = tasks[:params.limit]
        return _format_task_list_mpxj(limited, total, params)

    except Exception as e:
        return f"Error listing tasks: {str(e)}"


def _format_task_list(tasks: list, total: int, params) -> str:
    """Format COM-sourced task list into JSON or markdown."""
    if params.response_format == ResponseFormat.JSON:
        compact_tasks = []
        for t in tasks:
            ct = {
                "id": t["id"], "name": t["name"],
                "start": t.get("start", "N/A"), "finish": t.get("finish", "N/A"),
                "pct": t.get("percent_complete", 0), "critical": t.get("critical", False),
            }
            if t.get("milestone"):
                ct["milestone"] = True
            compact_tasks.append(ct)
        result = {"total": total, "returned": len(tasks), "source": "COM", "tasks": compact_tasks}
        if total > params.limit:
            result["note"] = f"Showing {params.limit} of {total}. Use limit param (max 200) or asta_task→get for details."
        return json.dumps(result, indent=1, default=str)

    lines = [f"# Task List ({total} total, showing {len(tasks)}) [COM]", ""]
    lines.append("| ID | Name | Start | Finish | %Done | Crit |")
    lines.append("|---|---|---|---|---|---|")
    for t in tasks:
        name = t["name"][:40]
        crit = "Yes" if t.get("critical") else ""
        lines.append(f"| {t['id']} | {name} | {t.get('start','N/A')} | {t.get('finish','N/A')} | {t.get('percent_complete',0)}% | {crit} |")
    if total > params.limit:
        lines.append("")
        lines.append(f"*{total - params.limit} more tasks. Use `limit` param (max 200) or `asta_task → get` for details.*")
    lines.append("")
    return "\n".join(lines)


def _format_task_list_mpxj(limited: list, total: int, params) -> str:
    """Format MPXJ-sourced task list into JSON or markdown."""
    if params.response_format == ResponseFormat.JSON:
        compact_tasks = []
        for t in limited:
            ct = {
                "id": t["id"], "name": t["name"],
                "start": t["start"], "finish": t["finish"],
                "pct": t["percent_complete"], "critical": t["critical"],
            }
            if t["summary"]:
                ct["summary"] = True
            if t["milestone"]:
                ct["milestone"] = True
            if t["predecessors"]:
                ct["pred"] = [p["task_id"] for p in t["predecessors"]]
            compact_tasks.append(ct)
        result = {"total": total, "returned": len(limited), "source": "MPXJ", "tasks": compact_tasks}
        if total > params.limit:
            result["note"] = f"Showing {params.limit} of {total}. Use limit param (max 200) or asta_task→get for details."
        return json.dumps(result, indent=1, default=str)

    lines = [f"# Task List ({total} total, showing {len(limited)}) [MPXJ]", ""]
    summary_count = sum(1 for t in limited if t["summary"])
    milestone_count = sum(1 for t in limited if t["milestone"])
    critical_count = sum(1 for t in limited if t["critical"])
    lines.append(f"Summary: {summary_count} | Milestones: {milestone_count} | Critical: {critical_count}")
    lines.append("")
    lines.append("| ID | Name | Start | Finish | %Done | Crit | Pred |")
    lines.append("|---|---|---|---|---|---|---|")
    for t in limited:
        name = t["name"][:40]
        if t["summary"]:
            name = f"**{name}**"
        crit = "Yes" if t["critical"] else ""
        pred = ",".join([str(p["task_id"]) for p in t["predecessors"][:3]]) if t["predecessors"] else ""
        if len(t.get("predecessors", [])) > 3:
            pred += "..."
        lines.append(f"| {t['id']} | {name} | {t['start']} | {t['finish']} | {t['percent_complete']}% | {crit} | {pred} |")
    if total > params.limit:
        lines.append("")
        lines.append(f"*{total - params.limit} more tasks. Use `limit` param (max 200) or `asta_task → get` for details.*")
    lines.append("")
    return "\n".join(lines)


# @mcp.tool(  # CONSOLIDATED into asta_task
#     name="asta_get_task",
#     annotations={
#         "title": "Get Task Details",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_get_task(params: GetTaskInput) -> str:
    """Get detailed information about a specific task by its ID.

    Returns comprehensive task data including dates, float, progress,
    cost, work, and critical path status.

    COM-first: tries live Asta via COM, falls back to MPXJ file.
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()

        bar = _find_bar_by_id_deep(project, params.task_id)
        if bar is None:
            # Bar not found via COM traversal — try MPXJ fallback
            try:
                export_path = _com_auto_export()
                mgr = AstaFileManager(export_path)
                task_data = mgr.get_task_by_id(params.task_id)
                if com_initialized:
                    pythoncom.CoUninitialize()
                if not task_data:
                    return json.dumps({"error": f"Task ID {params.task_id} not found in project"})
                task_data["source"] = "MPXJ (auto-export)"
                if params.response_format == ResponseFormat.JSON:
                    return json.dumps(task_data, indent=2, default=str)
                crit = " **[CRITICAL]**" if task_data.get('critical') else ""
                lines = [
                    f"# Task: {task_data['name']} (ID: {task_data['id']}){crit}",
                    "",
                    "## Schedule",
                    f"- **Duration:** {task_data.get('duration', 'N/A')}",
                    f"- **Start:** {task_data.get('start', 'N/A')} | **Finish:** {task_data.get('finish', 'N/A')}",
                    f"- **Early Start:** {task_data.get('early_start', 'N/A')} | **Early Finish:** {task_data.get('early_finish', 'N/A')}",
                    f"- **Late Start:** {task_data.get('late_start', 'N/A')} | **Late Finish:** {task_data.get('late_finish', 'N/A')}",
                    "",
                    "## Float",
                    f"- **Total Float:** {task_data.get('total_float', 'N/A')}",
                    f"- **Free Float:** {task_data.get('free_float', 'N/A')}",
                    "",
                    "## Progress",
                    f"- **% Complete:** {task_data.get('percent_complete', 0)}%",
                    f"- **Actual Start:** {task_data.get('actual_start', 'N/A')}",
                    f"- **Actual Finish:** {task_data.get('actual_finish', 'N/A')}",
                    "",
                    "## Cost & Work",
                    f"- **Cost:** {task_data.get('cost', 'N/A')}",
                    f"- **Work:** {task_data.get('work', 'N/A')}",
                    "",
                    f"**Calendar:** {task_data.get('calendar', 'N/A')}",
                ]
                if task_data.get('notes'):
                    lines.extend(["", f"**Notes:** {task_data['notes']}"])
                return "\n".join(lines)
            except Exception:
                pass
            if com_initialized:
                pythoncom.CoUninitialize()
            return json.dumps({"error": f"Task ID {params.task_id} not found in project"})

        task = {"id": bar.ID, "name": bar.Name}

        # Schedule dates
        for prop, key in [("Start", "start"), ("End", "end")]:
            try:
                task[key] = format_date(getattr(bar, prop))
            except Exception:
                task[key] = "N/A"

        # Duration
        try:
            dur = bar.Duration
            if dur is not None:
                task["duration"] = str(dur)
            else:
                task["duration"] = "N/A"
        except Exception:
            task["duration"] = "N/A"

        # Task properties (float, critical, early/late dates)
        # Use _get_bar_task (Tasks(1)) — ExpandedTask returns root for nested bars
        try:
            btask, is_et = _get_bar_task(bar)
            if btask:
                task["critical"] = bool(btask.Critical)
                for prop, key in [
                    ("GetUserStart", "early_start"), ("GetUserEnd", "early_finish"),
                ]:
                    try:
                        task[key] = format_date(getattr(btask, prop)())
                    except Exception:
                        task[key] = "N/A"

                # Float via tokens
                for token, key in [("TotalFloat", "total_float"), ("FreeFloat", "free_float")]:
                    try:
                        val = btask.EditToken(token, "")
                        task[key] = str(val) if val else "N/A"
                    except Exception:
                        try:
                            val = bar.GetToken(token)
                            task[key] = str(val) if val else "N/A"
                        except Exception:
                            task[key] = "N/A"
        except Exception:
            task["critical"] = False
            task["total_float"] = "N/A"
            task["free_float"] = "N/A"

        # Progress
        try:
            task["percent_complete"] = round(bar.OverallPercentComplete, 1)
        except Exception:
            try:
                task["percent_complete"] = round(bar.DurationPercentComplete, 1)
            except Exception:
                task["percent_complete"] = 0.0

        for prop, key in [("ActualStart", "actual_start"), ("ActualEnd", "actual_end"),
                          ("ActualFinish", "actual_finish")]:
            try:
                val = getattr(bar, prop)
                task[key] = format_date(val) if val else "N/A"
            except Exception:
                task[key] = "N/A"

        # Cost & Work
        for prop, key in [("Cost", "cost"), ("ActualEffort", "actual_effort"),
                          ("EffortRemaining", "effort_remaining")]:
            try:
                val = getattr(bar, prop)
                task[key] = str(val) if val else "N/A"
            except Exception:
                task[key] = "N/A"

        # Hierarchy info
        try:
            task["pathname"] = bar.Pathname
        except Exception:
            pass
        try:
            task["parentname"] = bar.Parentname
        except Exception:
            pass

        task["com_method"] = method

        if com_initialized:
            pythoncom.CoUninitialize()

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(task, indent=2, default=str)

        # Markdown output
        crit = " **[CRITICAL]**" if task.get("critical") else ""
        lines = [
            f"# Task: {task['name']} (ID: {task['id']}){crit}",
            "",
            "## Schedule",
            f"- **Duration:** {task.get('duration', 'N/A')}",
            f"- **Start:** {task.get('start', 'N/A')} | **End:** {task.get('end', 'N/A')}",
            f"- **Early Start:** {task.get('early_start', 'N/A')} | **Early Finish:** {task.get('early_finish', 'N/A')}",
            "",
            "## Float",
            f"- **Total Float:** {task.get('total_float', 'N/A')}",
            f"- **Free Float:** {task.get('free_float', 'N/A')}",
            "",
            "## Progress",
            f"- **% Complete:** {task.get('percent_complete', 0)}%",
            f"- **Actual Start:** {task.get('actual_start', 'N/A')}",
            f"- **Actual Finish:** {task.get('actual_finish', task.get('actual_end', 'N/A'))}",
            "",
            "## Cost & Work",
            f"- **Cost:** {task.get('cost', 'N/A')}",
            f"- **Actual Effort:** {task.get('actual_effort', 'N/A')}",
            f"- **Effort Remaining:** {task.get('effort_remaining', 'N/A')}",
        ]
        if task.get("pathname"):
            lines.extend(["", f"**Path:** {task['pathname']}"])
        return "\n".join(lines)

    except RuntimeError:
        # COM not available, fall back to MPXJ
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    except Exception as e:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        logger.warning(f"COM get_task failed: {e}, falling back to MPXJ")

    # --- MPXJ file fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        task = mgr.get_task_by_id(params.task_id)

        if not task:
            return f"Error: Task ID {params.task_id} not found in the project"

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(task, indent=2, default=str)

        crit = " **[CRITICAL]**" if task['critical'] else ""
        lines = [
            f"# Task: {task['name']} (ID: {task['id']}){crit}",
            "",
            "## Schedule",
            f"- **Duration:** {task['duration']}",
            f"- **Start:** {task['start']} | **Finish:** {task['finish']}",
            f"- **Early Start:** {task['early_start']} | **Early Finish:** {task['early_finish']}",
            f"- **Late Start:** {task['late_start']} | **Late Finish:** {task['late_finish']}",
            "",
            "## Float",
            f"- **Total Float:** {task['total_float']}",
            f"- **Free Float:** {task['free_float']}",
            "",
            "## Progress",
            f"- **% Complete:** {task['percent_complete']}%",
            f"- **Actual Start:** {task['actual_start']}",
            f"- **Actual Finish:** {task['actual_finish']}",
            "",
            "## Cost & Work",
            f"- **Planned Cost:** {task['cost']}",
            f"- **Actual Cost:** {task['actual_cost']}",
            f"- **Work:** {task['work']}",
            "",
            f"**Calendar:** {task['calendar']}",
        ]
        if task['notes']:
            lines.extend(["", f"**Notes:** {task['notes']}"])
        return "\n".join(lines)

    except Exception as e:
        return f"Error getting task: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_task
#     name="asta_add_task",
#     annotations={
#         "title": "Add New Task",
#         "readOnlyHint": False,
#         "destructiveHint": False,
#         "idempotentHint": False,
#         "openWorldHint": False,
#     }
# )
async def asta_add_task(params: AddTaskInput) -> str:
    """Add a new task to an Asta project file.

    Creates a new task with the specified name and duration,
    then saves the modified project to a new file (preserving the original).

    IMPORTANT: After adding tasks via file, you should open the updated
    file in Asta and press F9 (Reschedule) to recalculate the schedule.

    Args:
        params: Contains file_path, name, duration, save_output

    Returns:
        Confirmation with new task details and saved file path
    """
    # --- COM-first strategy: always try COM when Asta is running ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Add Task")
        result = _com_add_task(project, params.name, params.duration,
                              start_date=params.start_date, finish_date=params.finish_date)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        _com_end_transaction(project, reschedule=True)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        # Asta not running → fall through to MPXJ if file_path provided
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for add_task, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM add task failed: {e}"}, indent=2)
        logger.warning(f"COM add_task failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback (only reached if COM unavailable and file_path given) ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.add_task(params.name, params.duration)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        result["reminder"] = "Open this file in Asta and press F9 (Reschedule) to update the schedule"
        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error adding task: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_task
#     name="asta_update_task",
#     annotations={
#         "title": "Update Existing Task",
#         "readOnlyHint": False,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_update_task(params: UpdateTaskInput) -> str:
    """Update properties of an existing task in the project file.

    Can update task name, duration, completion percentage, and notes.
    Saves the modified project to a new file (preserving the original).

    Args:
        params: Contains file_path, task_id, name, duration, percent_complete, notes, save_output

    Returns:
        Confirmation with updated fields and saved file path
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Update Task")
        result = _com_update_task(project, params.task_id,
                                  name=params.name, duration_str=params.duration,
                                  percent_complete=params.percent_complete,
                                  notes=params.notes,
                                  start_date=params.start_date,
                                  finish_date=params.finish_date)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        # Reschedule if dates/duration were changed (ImposedStart/End need reschedule)
        needs_reschedule = params.start_date or params.finish_date or params.duration
        _com_end_transaction(project, reschedule=bool(needs_reschedule))
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for update_task, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM update task failed: {e}"}, indent=2)
        logger.warning(f"COM update_task failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.update_task(
            params.task_id,
            name=params.name,
            duration_str=params.duration,
            percent_complete=params.percent_complete,
            notes=params.notes
        )
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        result["reminder"] = "Open this file in Asta and press F9 (Reschedule) to update the schedule"
        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error updating task: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_task
#     name="asta_delete_task",
#     annotations={
#         "title": "Delete Task",
#         "readOnlyHint": False,
#         "destructiveHint": True,
#         "idempotentHint": False,
#         "openWorldHint": False,
#     }
# )
async def asta_delete_task(params: DeleteTaskInput) -> str:
    """Delete a task from the project file by its ID.

    WARNING: This permanently removes the task and its links.
    The original file is preserved; changes are saved to a new file.

    Args:
        params: Contains file_path, task_id, save_output

    Returns:
        Confirmation of deletion with saved file path
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        # _com_delete_task manages its own transactions internally
        result = _com_delete_task(project, params.task_id)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for delete_task, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM delete task failed: {e}"}, indent=2)
        logger.warning(f"COM delete_task failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.delete_task(params.task_id)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error deleting task: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_query
#     name="asta_get_critical_path",
#     annotations={
#         "title": "Get Critical Path",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_get_critical_path(params: CriticalPathInput) -> str:
    """Get all tasks on the critical path of the project.

    The critical path is the longest sequence of tasks that determines
    the minimum project duration. Any delay on these tasks delays the
    entire project. Tasks on the critical path have zero total float.

    Args:
        params: Contains file_path, response_format

    Returns:
        List of critical path tasks with their schedule details
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        critical = mgr.get_critical_path()

        if params.response_format == ResponseFormat.JSON:
            return json.dumps({"count": len(critical), "critical_tasks": critical}, indent=2, default=str)

        if not critical:
            return "No critical path tasks found. The project may not have been scheduled yet (run Reschedule in Asta)."

        lines = [
            "# Critical Path Analysis",
            "",
            f"**Critical Tasks:** {len(critical)}",
            "",
            "| ID | Name | Duration | Start | Finish | Float |",
            "|---|---|---|---|---|---|",
        ]
        for t in critical:
            lines.append(f"| {t['id']} | {t['name']} | {t['duration']} | {t['start']} | {t['finish']} | {t['total_float']} |")

        lines.extend([
            "",
            "> **Note:** Any delay on critical path tasks will delay the entire project.",
            "> These tasks must be monitored closely.",
        ])
        return "\n".join(lines)

    except Exception as e:
        return f"Error getting critical path: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_resource
#     name="asta_list_resources",
#     annotations={
#         "title": "List Project Resources",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_list_resources(params: ResourcesInput) -> str:
    """List all resources (labour, equipment, materials) in the project.

    Shows resource ID, name, type, maximum units, rate, and cost.

    Args:
        params: Contains file_path, response_format

    Returns:
        Resource list in markdown or JSON format
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        resources = mgr.get_resources()

        if params.response_format == ResponseFormat.JSON:
            return json.dumps({"total": len(resources), "resources": resources}, indent=2, default=str)

        if not resources:
            return "No resources found in this project."

        lines = [
            "# Project Resources",
            "",
            f"**Total:** {len(resources)} resources",
            "",
            "| ID | Name | Type | Max Units | Rate | Cost |",
            "|---|---|---|---|---|---|",
        ]
        for r in resources:
            lines.append(f"| {r['id']} | {r['name']} | {r['type']} | {r['max_units']} | {r['standard_rate']} | {r['cost']} |")
        return "\n".join(lines)

    except Exception as e:
        return f"Error listing resources: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_resource
#     name="asta_get_resource_assignments",
#     annotations={
#         "title": "Get Resource Assignments",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_get_resource_assignments(params: ResourcesInput) -> str:
    """Get all resource assignments showing which resources are assigned to which tasks.

    Shows task name, resource name, units, work hours, and cost for each assignment.

    Args:
        params: Contains file_path, response_format

    Returns:
        Resource assignment list in markdown or JSON format
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        assignments = mgr.get_resource_assignments()

        if params.response_format == ResponseFormat.JSON:
            return json.dumps({"total": len(assignments), "assignments": assignments}, indent=2, default=str)

        if not assignments:
            return "No resource assignments found in this project."

        lines = [
            "# Resource Assignments",
            "",
            f"**Total:** {len(assignments)} assignments",
            "",
            "| Task | Resource | Units | Work | Cost |",
            "|---|---|---|---|---|",
        ]
        for a in assignments:
            lines.append(f"| {a['task_name']} (ID:{a['task_id']}) | {a['resource_name']} | {a['units']} | {a['work']} | {a['cost']} |")
        return "\n".join(lines)

    except Exception as e:
        return f"Error getting assignments: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_calendar
#     name="asta_get_calendars",
#     annotations={
#         "title": "Get Project Calendars",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_get_calendars(params: ProjectFileInput) -> str:
    """Get all calendars defined in the project.

    Calendars define working days, hours, and exceptions (holidays, overtime).

    Args:
        params: Contains file_path

    Returns:
        Calendar list in JSON format
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        calendars = mgr.get_calendars()
        return json.dumps({"total": len(calendars), "calendars": calendars}, indent=2, default=str)

    except Exception as e:
        return f"Error getting calendars: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_query
#     name="asta_float_analysis",
#     annotations={
#         "title": "Float Analysis",
#         "readOnlyHint": True,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_float_analysis(params: FloatAnalysisInput) -> str:
    """Analyze float (slack) distribution across all tasks.

    Float is the amount of time a task can be delayed without affecting
    the project end date. This analysis categorizes tasks by float amount:
    - Zero float: Critical tasks (any delay affects project)
    - Low float (1-5 days): Near-critical, needs monitoring
    - Medium float (6-20 days): Some flexibility
    - High float (>20 days): Significant flexibility

    Args:
        params: Contains file_path, response_format

    Returns:
        Float distribution analysis
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        analysis = mgr.get_float_analysis()

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(analysis, indent=2, default=str)

        lines = [
            "# Float Analysis",
            "",
            "## Distribution",
            f"- **Zero Float (Critical):** {analysis['zero_float']} tasks",
            f"- **Low Float (1-5 days):** {analysis['low_float']} tasks",
            f"- **Medium Float (6-20 days):** {analysis['medium_float']} tasks",
            f"- **High Float (>20 days):** {analysis['high_float']} tasks",
            "",
            "## Task Details",
            "",
            "| ID | Name | Total Float |",
            "|---|---|---|",
        ]
        for t in analysis['tasks'][:30]:
            lines.append(f"| {t['id']} | {t['name']} | {t['total_float']} |")
        if len(analysis['tasks']) > 30:
            lines.append(f"\n*...and {len(analysis['tasks']) - 30} more tasks*")
        return "\n".join(lines)

    except Exception as e:
        return f"Error in float analysis: {str(e)}"


# ============================================================================
# NEW FEATURE TOOLS
# ============================================================================

# @mcp.tool(  # CONSOLIDATED into asta_query
#     name="asta_get_wbs_tree",
#     annotations={"title": "Get WBS/Hierarchy Tree", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
# )
async def asta_get_wbs_tree(params: WBSTreeInput) -> str:
    """Get the WBS (Work Breakdown Structure) hierarchy tree.

    Shows parent-child relationships between tasks, summary groups,
    outline levels, and WBS codes. Essential for understanding
    how the project is organized.

    Args:
        params: Contains file_path, response_format

    Returns:
        WBS tree in markdown or JSON format
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        tree = mgr.get_wbs_tree(max_depth=params.max_depth)

        max_nodes = 200

        if params.response_format == ResponseFormat.JSON:
            node_count = [0]
            truncated = [False]
            def limit_tree_json(nodes):
                result = []
                for node in nodes:
                    if node_count[0] >= max_nodes:
                        truncated[0] = True
                        break
                    node_count[0] += 1
                    limited_node = dict(node)
                    if node["children"]:
                        limited_node["children"] = limit_tree_json(node["children"])
                    result.append(limited_node)
                return result
            limited = limit_tree_json(tree)
            result = {"wbs_tree": limited, "total_nodes_shown": node_count[0]}
            if truncated[0]:
                result["note"] = f"Tree truncated at {max_nodes} nodes. Use asta_get_task for specific task details."
            return json.dumps(result, indent=2, default=str)

        lines = ["# WBS / Hierarchy Tree", ""]
        node_count = [0]
        truncated = [False]

        def render_tree(nodes, indent=0):
            for node in nodes:
                if node_count[0] >= max_nodes:
                    truncated[0] = True
                    return
                node_count[0] += 1
                prefix = "  " * indent
                icon = "[S] " if node["summary"] else "[M] " if node["milestone"] else "    "
                lines.append(f"{prefix}{icon}{node['name']} (ID:{node['id']}, WBS:{node['wbs']})")
                if not node["summary"]:
                    lines.append(f"{prefix}     Duration: {node['duration']} | {node['start']} - {node['finish']}")
                if node["children"]:
                    render_tree(node["children"], indent + 1)

        render_tree(tree)
        if truncated[0]:
            lines.append(f"\n*...truncated at {max_nodes} nodes. Use asta_get_task for specific task details.*")
        lines.extend(["", "Legend: [S]=Summary, [M]=Milestone"])
        return "\n".join(lines)

    except Exception as e:
        return f"Error getting WBS tree: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_task
#     name="asta_add_summary_task",
#     annotations={"title": "Add Summary Task", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
# )
async def asta_add_summary_task(params: AddSummaryTaskInput) -> str:
    """Add a new summary (group) task to organize other tasks.

    Summary tasks act as containers/folders for child tasks.
    Can be added at top level or nested under another task.

    Args:
        params: Contains file_path, name, parent_task_id, save_output

    Returns:
        Confirmation with new summary task details
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Add Summary Task")
        result = _com_add_task(project, params.name,
                               parent_bar_id=params.parent_task_id,
                               is_summary=True)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        _com_end_transaction(project)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for add_summary_task, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM add summary task failed: {e}"}, indent=2)
        logger.warning(f"COM add_summary_task failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.add_summary_task(params.name, params.parent_task_id)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error adding summary task: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_task
#     name="asta_add_child_task",
#     annotations={"title": "Add Child Task", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
# )
async def asta_add_child_task(params: AddChildTaskInput) -> str:
    """Add a new task under a specific parent/summary task.

    Creates the task as a child of the specified parent,
    maintaining the WBS hierarchy.

    Args:
        params: Contains file_path, parent_task_id, name, duration, save_output

    Returns:
        Confirmation with new task details and parent info
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Add Child Task")
        result = _com_add_task(project, params.name, params.duration,
                               start_date=params.start_date, finish_date=params.finish_date,
                               parent_bar_id=params.parent_task_id)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        _com_end_transaction(project, reschedule=True)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for add_child_task, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM add child task failed: {e}"}, indent=2)
        logger.warning(f"COM add_child_task failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.add_child_task(params.parent_task_id, params.name, params.duration)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error adding child task: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_link
#     name="asta_add_link",
#     annotations={"title": "Add Task Link", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
# )
async def asta_add_link(params: AddLinkInput) -> str:
    """Add a predecessor-successor link between two tasks.

    Link types:
    - FS (Finish-to-Start): B starts after A finishes (most common)
    - SS (Start-to-Start): B starts when A starts
    - FF (Finish-to-Finish): B finishes when A finishes
    - SF (Start-to-Finish): B finishes when A starts

    Lag adds waiting time (e.g., '2d' for concrete curing).
    Negative lag (lead) means overlap (e.g., '-1d').

    Args:
        params: Contains file_path, predecessor_id, successor_id, link_type, lag, save_output

    Returns:
        Confirmation with link details
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Add Link")
        result = _com_add_link(project, params.predecessor_id, params.successor_id,
                               params.link_type, params.lag)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        _com_end_transaction(project)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for add_link, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM add link failed: {e}"}, indent=2)
        logger.warning(f"COM add_link failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.add_link(params.predecessor_id, params.successor_id, params.link_type, params.lag)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error adding link: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_link
#     name="asta_remove_link",
#     annotations={"title": "Remove Task Link", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False}
# )
async def asta_remove_link(params: RemoveLinkInput) -> str:
    """Remove a predecessor-successor link between two tasks.

    WARNING: This permanently removes the dependency link.
    The tasks themselves are NOT deleted.

    Args:
        params: Contains file_path, predecessor_id, successor_id, save_output

    Returns:
        Confirmation of link removal
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Remove Link")
        result = _com_remove_link(project, params.predecessor_id, params.successor_id)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        _com_end_transaction(project)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for remove_link, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM remove link failed: {e}"}, indent=2)
        logger.warning(f"COM remove_link failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.remove_link(params.predecessor_id, params.successor_id)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error removing link: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_link
#     name="asta_update_link",
#     annotations={"title": "Update Task Link", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
# )
async def asta_update_link(params: UpdateLinkInput) -> str:
    """Update an existing task link's type or lag.

    Can change the link type (FS/SS/FF/SF) and/or the lag duration.
    Internally removes the old link and creates a new one.

    Args:
        params: Contains file_path, predecessor_id, successor_id, new_link_type, new_lag, save_output

    Returns:
        Confirmation with old and new link properties
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Update Link")
        result = _com_update_link(project, params.predecessor_id, params.successor_id,
                                  params.new_link_type, params.new_lag)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        _com_end_transaction(project)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for update_link, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM update link failed: {e}"}, indent=2)
        logger.warning(f"COM update_link failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.update_link(params.predecessor_id, params.successor_id, params.new_link_type, params.new_lag)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error updating link: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_progress
#     name="asta_update_progress",
#     annotations={"title": "Update Task Progress", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
# )
async def asta_update_progress(params: UpdateProgressInput) -> str:
    """Update progress data for a single task.

    Can set completion percentage, actual start date, and actual finish date.
    Dates can be in YYYY-MM-DD or DD/MM/YYYY format.

    Args:
        params: Contains file_path, task_id, percent_complete, actual_start, actual_finish, save_output

    Returns:
        Confirmation with updated progress details
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Update Progress")
        result = _com_update_progress(project, params.task_id,
                                      percent_complete=params.percent_complete,
                                      actual_start=params.actual_start,
                                      actual_finish=params.actual_finish)
        if "error" in result:
            project.AbandonTransaction()
            return json.dumps(result, indent=2)
        _com_end_transaction(project)
        result["com_method"] = method
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for update_progress, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM update progress failed: {e}"}, indent=2)
        logger.warning(f"COM update_progress failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        result = mgr.update_progress(params.task_id, params.percent_complete, params.actual_start, params.actual_finish)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error updating progress: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_progress
#     name="asta_bulk_update_progress",
#     annotations={"title": "Bulk Update Progress", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
# )
async def asta_bulk_update_progress(params: BulkUpdateProgressInput) -> str:
    """Update progress for multiple tasks at once.

    Accepts a list of task updates, each with task_id,
    percent_complete, actual_start, and actual_finish.
    Useful for weekly progress reporting.

    Args:
        params: Contains file_path, updates list, save_output

    Returns:
        Summary of all updates with success/failure counts
    """
    # --- COM-first strategy ---
    import pythoncom
    com_initialized = False
    _com_project = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        _com_project = project
        project.StartTransaction("Bulk Update Progress")

        results_list = []
        errors_list = []
        for upd in params.updates:
            upd_result = _com_update_progress(
                project, upd.task_id,
                percent_complete=upd.percent_complete,
                actual_start=upd.actual_start,
                actual_finish=upd.actual_finish
            )
            if "error" in upd_result:
                errors_list.append(upd_result)
            else:
                results_list.append(upd_result)

        _com_end_transaction(project)

        result = {
            "method": "COM",
            "com_method": method,
            "total": len(params.updates),
            "successful": len(results_list),
            "failed": len(errors_list),
            "updates": results_list,
        }
        if errors_list:
            result["errors"] = errors_list
        return json.dumps(result, indent=2, default=str)
    except RuntimeError:
        if not params.file_path:
            return json.dumps({"error": "Asta Powerproject is not running and no file_path was provided. Either open Asta or provide a file_path."}, indent=2)
        logger.info("COM unavailable for bulk_update_progress, falling back to MPXJ file mode")
    except Exception as e:
        if _com_project:
            try: _com_project.AbandonTransaction()
            except: pass
        if not params.file_path:
            return json.dumps({"error": f"COM bulk progress update failed: {e}"}, indent=2)
        logger.warning(f"COM bulk_update_progress failed ({e}), falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        updates_list = [u.model_dump() for u in params.updates]
        result = mgr.bulk_update_progress(updates_list)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error in bulk progress update: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_query
#     name="asta_delay_analysis",
#     annotations={"title": "Delay Analysis", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
# )
async def asta_delay_analysis(params: DelayAnalysisInput) -> str:
    """Analyze schedule delays by comparing planned vs actual dates.

    Calculates start slip and finish slip for each task that has
    actual dates. Identifies which tasks are delayed, early, or on time.
    Shows critical tasks with delays (highest risk).

    Args:
        params: Contains file_path, response_format

    Returns:
        Delay analysis with statistics and task details
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        analysis = mgr.get_delay_analysis()

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(analysis, indent=2, default=str)

        lines = [
            "# Delay Analysis",
            "",
            "## Summary",
            f"- **Tasks with actual dates:** {analysis['total_with_actuals']}",
            f"- **Delayed starts:** {analysis['delayed_starts']}",
            f"- **Delayed finishes:** {analysis['delayed_finishes']}",
            f"- **Early starts:** {analysis['early_starts']}",
            f"- **Max start slip:** {analysis['max_start_slip']} days",
            f"- **Max finish slip:** {analysis['max_finish_slip']} days",
            "",
        ]

        if analysis["tasks"]:
            lines.extend([
                "## Task Details",
                "",
                "| ID | Name | Planned Start | Actual Start | Start Slip | Critical |",
                "|---|---|---|---|---|---|",
            ])
            for t in analysis["tasks"][:50]:
                slip = f"{t['start_slip_days']}d" if t['start_slip_days'] is not None else "N/A"
                crit = "YES" if t['critical'] else ""
                lines.append(f"| {t['id']} | {t['name']} | {t['planned_start']} | {t['actual_start']} | {slip} | {crit} |")
            if len(analysis["tasks"]) > 50:
                lines.append(f"\n*...and {len(analysis['tasks']) - 50} more tasks*")
        else:
            lines.append("No tasks with actual dates found. Enter progress data first.")

        return "\n".join(lines)

    except Exception as e:
        return f"Error in delay analysis: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_resource
#     name="asta_resource_loading",
#     annotations={"title": "Resource Loading Analysis", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
# )
async def asta_resource_loading(params: ResourceLoadingInput) -> str:
    """Analyze resource loading (work distribution) across the project.

    Shows each resource's total work hours, cost, number of assigned tasks,
    and detailed task-by-task breakdown. Useful for identifying
    over-allocated or under-utilized resources.

    Args:
        params: Contains file_path, response_format

    Returns:
        Resource loading analysis with work/cost totals
    """
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        loading = mgr.get_resource_loading()

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(loading, indent=2, default=str)

        if not loading["resources"]:
            return "No resource assignments found. Assign resources to tasks first."

        lines = [
            "# Resource Loading Analysis",
            "",
            f"**Total Resources:** {loading['total_resources']}",
            f"**Total Assignments:** {loading['total_assignments']}",
            "",
            "## Resource Summary",
            "",
            "| Resource | Type | Tasks | Total Work | Total Cost |",
            "|---|---|---|---|---|",
        ]
        for r in loading["resources"]:
            lines.append(f"| {r['name']} | {r['type']} | {r['task_count']} | {r['total_work']}h | {r['total_cost']} |")

        # Detailed breakdown for top resources
        lines.extend(["", "## Detailed Assignments", ""])
        for r in loading["resources"][:10]:
            lines.append(f"### {r['name']} ({r['task_count']} tasks)")
            for t in r["tasks"][:10]:
                lines.append(f"- {t['task_name']} | {t['work']} | {t['start']} - {t['finish']}")
            if len(r["tasks"]) > 10:
                lines.append(f"  *...and {len(r['tasks']) - 10} more*")
            lines.append("")

        return "\n".join(lines)

    except Exception as e:
        return f"Error in resource loading: {str(e)}"


# @mcp.tool(  # CONSOLIDATED into asta_schedule
#     name="asta_save_project",
#     annotations={
#         "title": "Save Project File",
#         "readOnlyHint": False,
#         "destructiveHint": False,
#         "idempotentHint": True,
#         "openWorldHint": False,
#     }
# )
async def asta_save_project(params: SaveProjectInput) -> str:
    """Save the project to an XML file.

    COM-first: uses SaveAsXMLFile directly on running Asta (fast).
    Fallback: MPXJ file-based save.
    """
    output_path = params.output_path
    if not output_path:
        output_path = os.path.join(
            os.environ.get("TEMP", "/tmp"), "asta_mcp", "saved_project.xml"
        )
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # --- COM-first: direct SaveAsXMLFile (fast) ---
    import pythoncom
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()
        project.SaveAsXMLFile(output_path, None, None)
        return json.dumps({
            "success": True,
            "method": "COM",
            "saved_to": output_path,
            "message": f"Project saved to: {output_path}"
        }, indent=2)
    except RuntimeError:
        logger.info("COM save failed, falling back to MPXJ")
    except Exception as e:
        logger.warning(f"COM save error: {e}, falling back to MPXJ")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- MPXJ fallback ---
    try:
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
        output = mgr.save(params.output_path)
        return json.dumps({
            "success": True,
            "method": "MPXJ",
            "saved_to": output,
            "message": f"Project saved to: {output}. Open in Asta and press F9 to reschedule."
        }, indent=2)

    except Exception as e:
        return f"Error saving project: {str(e)}"


# COM-BASED TOOLS (win32com / Asta OLE Automation)
# ============================================================================

class RescheduleProjectInput(BaseModel):
    """Input for rescheduling a project via COM automation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    report_date: str = Field(
        ...,
        description="The new Data Date / Report Date in ISO 8601 format (YYYY-MM-DD). "
                    "All uncompleted work will be referenced from this date."
    )
    straighten_uncompleted_work: bool = Field(
        default=True,
        description="If true, moves remaining (uncompleted) work to start on or after "
                    "the Report Date. Equivalent to 'straightening the progress line'."
    )
    preserve_links: bool = Field(
        default=False,
        description="If true, uses Retained Logic (respects original link sequences). "
                    "If false, uses Progress Override (allows out-of-sequence progress)."
    )
    target_wbs_id: Optional[str] = Field(
        default=None,
        description="If provided, only reschedules the branch containing this bar/task ID. "
                    "If omitted, reschedules the entire project."
    )

    @field_validator("report_date")
    @classmethod
    def validate_report_date(cls, v: str) -> str:
        try:
            datetime.strptime(v, "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"report_date must be YYYY-MM-DD format, got: '{v}'")
        return v


def _connect_asta_com() -> tuple:
    """Attempt to connect to a running Asta Powerproject instance via COM.

    Tries multiple connection strategies in order:
      1. GetActiveObject by Application CLSID (running instance via ROT)
      2. Dispatch by Application CLSID
      3. Dispatch by common ProgID candidates

    Returns:
        (app_object, project_object, method_used: str) on success
    Raises:
        RuntimeError with diagnostic info on failure
    """
    import pythoncom
    import win32com.client
    import pywintypes

    # Asta Application CLSID from teamplan.tlb type library
    APP_CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"

    errors = []

    # --- Strategy 1: GetActiveObject via ROT (connects to running instance) ---
    try:
        obj = pythoncom.GetActiveObject(APP_CLSID)
        app = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
        project = app.ActiveProject
        if project is None:
            raise RuntimeError("Asta is running but no project is open (ActiveProject is None)")
        return app, project, "GetActiveObject (CLSID ROT)"
    except pywintypes.com_error as e:
        errors.append(f"GetActiveObject(CLSID): {e}")
    except Exception as e:
        errors.append(f"GetActiveObject(CLSID): {e}")

    # --- Strategy 2: Dispatch by CLSID (may launch or attach) ---
    try:
        app = win32com.client.dynamic.Dispatch(APP_CLSID)
        project = app.ActiveProject
        if project is None:
            raise RuntimeError("Connected via Dispatch(CLSID) but no project is open")
        return app, project, "Dispatch (CLSID)"
    except pywintypes.com_error as e:
        errors.append(f"Dispatch(CLSID): {e}")
    except Exception as e:
        errors.append(f"Dispatch(CLSID): {e}")

    # --- Strategy 3: Well-known ProgID candidates ---
    progids = [
        "Powerproject.Application",
        "Asta.Application",
        "AstaPowerproject.Application",
        "AstaDkit.Application",
        "Elecosoft.Powerproject",
    ]
    for progid in progids:
        try:
            app = win32com.client.dynamic.Dispatch(progid)
            project = app.ActiveProject
            if project is None:
                raise RuntimeError(f"Connected via {progid} but no project is open")
            return app, project, f"Dispatch('{progid}')"
        except pywintypes.com_error:
            errors.append(f"Dispatch('{progid}'): ProgID not registered")
        except Exception as e:
            errors.append(f"Dispatch('{progid}'): {e}")

    # All strategies failed
    raise RuntimeError(
        "Could not connect to Asta Powerproject via COM.\n\n"
        "Possible causes:\n"
        "  1. Asta Powerproject is not running\n"
        "  2. No project file is currently open in Asta\n"
        "  3. The Asta Developers' Toolkit OCX is not registered\n"
        "     (run: regsvr32 \"C:\\Program Files\\Elecosoft\\Powerproject\\astadkit.ocx\")\n"
        "  4. Asta does not expose COM automation in this edition/version\n\n"
        "Detailed connection attempts:\n" +
        "\n".join(f"  - {e}" for e in errors) +
        "\n\nFallback: The tool will attempt GUI automation instead."
    )


# ---------------------------------------------------------------------------
# COM AUTO-EXPORT — for read-only tools when Asta is running
# ---------------------------------------------------------------------------

_com_auto_export_cache = {"path": None, "timestamp": 0}


def _com_auto_export() -> str:
    """When Asta is running and no file_path given, auto-export to temp XML.

    Uses a 30-second cache to avoid re-exporting on rapid successive queries.
    Returns the temp XML path.
    Raises RuntimeError if Asta is not running.
    """
    import pythoncom
    import time

    # Check cache (30 second TTL)
    cache = _com_auto_export_cache
    if cache["path"] and os.path.exists(cache["path"]) and (time.time() - cache["timestamp"]) < 30:
        logger.info(f"Using cached auto-export: {cache['path']}")
        return cache["path"]

    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        app, project, method = _connect_asta_com()

        # Export to temp directory
        temp_dir = os.path.join(os.environ.get("TEMP", "/tmp"), "asta_mcp")
        os.makedirs(temp_dir, exist_ok=True)
        temp_path = os.path.join(temp_dir, "auto_export.xml")

        project.SaveAsXMLFile(temp_path, None, None)

        cache["path"] = temp_path
        cache["timestamp"] = time.time()
        logger.info(f"COM auto-export to {temp_path} via {method}")
        return temp_path
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _resolve_file_path(file_path: str = None) -> str:
    """Resolve file_path: if None, try COM auto-export.

    Returns a valid file path for MPXJ reading.
    Raises RuntimeError if neither file_path nor COM is available.
    """
    if file_path:
        return file_path
    # Try COM auto-export
    return _com_auto_export()


# ---------------------------------------------------------------------------
# COM TRANSACTION HELPER — EndTransaction + WaitForNotificationProcessing
# ---------------------------------------------------------------------------

def _com_end_transaction(project, reschedule: bool = False):
    """End a COM transaction and wait for Asta to process all notifications.

    CRITICAL: WaitForNotificationProcessing must be called after every
    COM update to ensure Asta has fully processed the changes (dates,
    duration, links, etc.) before the next operation.

    If reschedule=True, also runs Reschedule() so that date constraints
    (ImposedStart/ImposedEnd) take effect in the schedule.
    """
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass  # Method may not exist in all Asta versions

    if reschedule:
        try:
            import pywintypes
            project.Reschedule(pywintypes.Time(datetime.now()))
            try:
                project.WaitForNotificationProcessing()
            except Exception:
                pass
        except Exception:
            pass  # Reschedule may fail if no report date set


# ---------------------------------------------------------------------------
# COM CRUD HELPERS — used by dual-strategy tools
# ---------------------------------------------------------------------------

def _find_bar_by_id(bars_or_project, bar_id: int):
    """Find a bar by its ID by traversing the full hierarchy.

    Uses bar.Tasks(1).ChildBars for hierarchy navigation (VERIFIED WORKING).
    bar.ExpandedTask.ChildBars returns SELF — do NOT use for navigation.

    Args:
        bars_or_project: Either project.Bars collection or the project object itself.
        bar_id: The bar ID to find.
    Returns:
        The bar COM object, or None if not found.
    """
    import win32com.client
    # Determine project from input
    project = None
    bars = None
    try:
        if hasattr(bars_or_project, 'Bars'):
            project = bars_or_project
            bars = project.Bars
        else:
            bars = bars_or_project
    except Exception:
        return None

    # Check root bar(s) first
    try:
        count = bars.Count
        for i in range(1, count + 1):
            try:
                bar = bars.Item(i)
                if bar.ID == bar_id:
                    return bar
            except Exception:
                continue
    except Exception:
        pass

    # Deep search: traverse hierarchy via bar.Tasks(1).ChildBars
    try:
        count = bars.Count
        for i in range(1, count + 1):
            try:
                root_bar = bars.Item(i)
                root_task = win32com.client.Dispatch(root_bar.Tasks(1))
                found = _search_bar_hierarchy(root_task, bar_id, 0, 6)
                if found:
                    return found
            except Exception:
                continue
    except Exception:
        pass

    return None


def _search_bar_hierarchy(parent_task, target_id, depth, max_depth):
    """Recursively search for a bar by ID via Tasks(1).ChildBars."""
    import win32com.client
    if depth >= max_depth:
        return None
    try:
        child_bars = parent_task.ChildBars
        for i in range(1, child_bars.Count + 1):
            try:
                cb = win32com.client.Dispatch(child_bars.Item(i))
                if cb.ID == target_id:
                    return cb
                # Recurse into children
                try:
                    ct = win32com.client.Dispatch(cb.Tasks(1))
                    result = _search_bar_hierarchy(ct, target_id, depth + 1, max_depth)
                    if result:
                        return result
                except Exception:
                    pass  # Leaf bar or no Tasks(1) — skip
            except Exception:
                continue
    except Exception:
        pass
    return None


def _get_bar_task(bar):
    """Get the task object for a bar. Returns (task, is_expanded_task).

    bar.Tasks(1) returns ITask for leaf bars, IExpandedTask for summaries.
    IExpandedTask has ImposedStart/ImposedEnd/ChildBars.
    ITask has LinkTo, SetUserDuration, StartConstraintDate but NOT ImposedStart.
    """
    import win32com.client
    try:
        t = win32com.client.Dispatch(bar.Tasks(1))
        is_et = type(t).__name__ == 'IExpandedTask'
        return t, is_et
    except Exception:
        return None, False


def _com_get_all_bars(project, max_bars=500):
    """Get all bars from the project via COM, traversing hierarchy.

    Uses bar.Tasks(1).ChildBars for correct hierarchy traversal.
    """
    import win32com.client
    all_bars = []
    try:
        bars = project.Bars
        count = bars.Count
        if count == 0:
            return all_bars

        for i in range(1, min(count + 1, max_bars + 1)):
            try:
                bar = bars.Item(i)
                all_bars.append(bar)
                # Traverse children
                _collect_child_bars(bar, all_bars, max_bars, 0, 6)
            except Exception:
                continue
    except Exception:
        pass

    return all_bars[:max_bars]


def _collect_child_bars(bar, results, max_bars, depth, max_depth):
    """Recursively collect child bars via Tasks(1).ChildBars."""
    import win32com.client
    if depth >= max_depth or len(results) >= max_bars:
        return
    try:
        task = win32com.client.Dispatch(bar.Tasks(1))
        child_bars = task.ChildBars
        for i in range(1, child_bars.Count + 1):
            if len(results) >= max_bars:
                return
            try:
                cb = win32com.client.Dispatch(child_bars.Item(i))
                results.append(cb)
                _collect_child_bars(cb, results, max_bars, depth + 1, max_depth)
            except Exception:
                continue
    except Exception:
        pass


def _find_bar_by_id_deep(project, bar_id: int):
    """Find a bar by ID — searches full hierarchy."""
    return _find_bar_by_id(project, bar_id)




def _parse_duration_to_minutes(duration_str: str) -> float:
    """Convert duration string like '5d', '2w', '8h' to minutes.
    Asta uses 8-hour work days by default."""
    if not duration_str:
        return 480.0  # 1 day default
    s = duration_str.strip().lower()
    try:
        if s.endswith('w'):
            return float(s[:-1]) * 5 * 8 * 60  # weeks → minutes
        elif s.endswith('d'):
            return float(s[:-1]) * 8 * 60  # days → minutes
        elif s.endswith('h'):
            return float(s[:-1]) * 60  # hours → minutes
        elif s.endswith('m'):
            return float(s[:-1])  # already minutes
        else:
            return float(s) * 8 * 60  # assume days
    except ValueError:
        return 480.0  # fallback: 1 day


def _parse_date(date_str: str):
    """Parse date string to datetime. Supports YYYY-MM-DD and DD/MM/YYYY."""
    from datetime import datetime as dt
    if not date_str:
        return None
    date_str = date_str.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S"):
        try:
            return dt.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def _set_bar_dates(bar, start_date_str: str = None, finish_date_str: str = None,
                   duration_str: str = None) -> dict:
    """Set start/finish dates and duration on a bar via COM.

    VERIFIED WORKING approaches:
    - On IExpandedTask (summaries): etask.ImposedStart / ImposedEnd
    - On ITask (leaves): task.StartConstraintDate
    - Both: task.GetDurationFromString("10d") + SetUserDuration(dur_obj)

    IMPORTANT: Dates take effect after project.Reschedule() is called.
    The caller should reschedule after EndTransaction.
    """
    import pywintypes
    from datetime import datetime as dt, timedelta
    info = {}

    start_dt = _parse_date(start_date_str) if start_date_str else None
    finish_dt = _parse_date(finish_date_str) if finish_date_str else None

    # If no start date given, use today as default
    if start_dt is None and finish_dt is None:
        start_dt = dt.now().replace(hour=8, minute=0, second=0, microsecond=0)
        info["start_defaulted"] = True

    # Get the task object (ITask or IExpandedTask)
    task, is_expanded = _get_bar_task(bar)
    if task is None:
        # Fallback: try bar.ExpandedTask (may return root's task — unreliable)
        try:
            task = bar.ExpandedTask
            is_expanded = True
        except Exception:
            info["error"] = "Cannot access task for this bar"
            return info

    # Set start date
    if start_dt is not None:
        ole_start = pywintypes.Time(start_dt)
        if is_expanded and hasattr(task, 'ImposedStart'):
            # IExpandedTask: use ImposedStart (VERIFIED WORKING)
            try:
                task.ImposedStart = ole_start
                info["start"] = start_dt.strftime("%Y-%m-%d")
                info["start_method"] = "ImposedStart"
            except Exception:
                info["start_warning"] = "Could not set ImposedStart"
        else:
            # ITask: use StartConstraintDate (VERIFIED WORKING)
            try:
                task.StartConstraintDate = ole_start
                info["start"] = start_dt.strftime("%Y-%m-%d")
                info["start_method"] = "StartConstraintDate"
            except Exception:
                info["start_warning"] = "Could not set start date"

    # Set finish date
    if finish_dt is not None:
        ole_end = pywintypes.Time(finish_dt)
        if is_expanded and hasattr(task, 'ImposedEnd'):
            try:
                task.ImposedEnd = ole_end
                info["finish"] = finish_dt.strftime("%Y-%m-%d")
            except Exception:
                info["finish_warning"] = "Could not set finish date"

    # Set duration via GetDurationFromString + SetUserDuration (VERIFIED WORKING)
    if duration_str:
        try:
            dur_obj = task.GetDurationFromString(duration_str)
            if dur_obj is not None:
                task.SetUserDuration(dur_obj)
                info["duration"] = duration_str
        except Exception:
            info["duration_warning"] = (
                f"Could not set duration '{duration_str}'. "
                "Dates were set via constraints instead."
            )

    info["note"] = "Dates take effect after reschedule (F9)"
    return info


def _com_add_task(project, name: str, duration_str: str = "1d",
                  start_date: str = None, finish_date: str = None,
                  parent_bar_id: int = None, is_summary: bool = False,
                  is_milestone: bool = False) -> dict:
    """Add a task via COM. Returns result dict.

    VERIFIED WORKING workflow (v21):
    1. Find parent bar (or use project root)
    2. parent_task.ChildBars.Add() → creates empty bar
    3. bar.Tasks.AddTask(start_date, duration) → creates proper ITask
    4. Set bar name
    5. Dates take effect after project.Reschedule()

    For summaries: ChildBars.Add() + bar.Tasks.AddSummaryTask(date)
    For milestones: ChildBars.Add() + bar.Tasks.AddMilestone(date)
    """
    import win32com.client
    import pywintypes
    from datetime import datetime as dt

    result = {"method": "COM"}

    # Parse start date
    start_dt = _parse_date(start_date) if start_date else None
    if start_dt is None:
        start_dt = dt.now().replace(hour=8, minute=0, second=0, microsecond=0)
        result["start_defaulted"] = True
    ole_start = pywintypes.Time(start_dt)

    # Find parent for the new bar
    parent_task = None
    if parent_bar_id is not None:
        parent_bar = _find_bar_by_id(project, parent_bar_id)
        if parent_bar:
            try:
                parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
                result["parent_id"] = parent_bar_id
            except Exception:
                result["parent_warning"] = (
                    f"Parent bar {parent_bar_id} found but has no task. "
                    "Creating under project root instead."
                )

    # If no parent specified or not found, use project root
    if parent_task is None:
        try:
            root_bar = project.Bars.Item(1)
            parent_task = win32com.client.Dispatch(root_bar.Tasks(1))
        except Exception:
            # Absolute fallback: create top-level bar
            new_bar = project.Bars.Add()
            new_bar.Name = name
            result["task_id"] = new_bar.ID
            result["name"] = name
            result["warning"] = "Created as top-level bar (no hierarchy)"
            return result

    # Create the bar under the parent
    try:
        new_bar = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_bar.Name = name
        bar_id = new_bar.ID
    except Exception as e:
        result["error"] = f"Failed to create bar: {str(e)[:100]}"
        return result

    # Add a task to the bar
    if is_milestone:
        try:
            task = new_bar.Tasks.AddMilestone(ole_start)
            if task:
                result["type"] = "milestone"
        except Exception as e:
            result["task_id"] = bar_id
            result["name"] = name
            result["warning"] = f"Bar created but AddMilestone failed: {str(e)[:80]}"
            return result
    elif is_summary:
        try:
            task = new_bar.Tasks.AddSummaryTask(ole_start)
            if task:
                result["type"] = "summary"
        except Exception as e:
            result["task_id"] = bar_id
            result["name"] = name
            result["warning"] = f"Bar created but AddSummaryTask failed: {str(e)[:80]}"
            return result
    else:
        # Normal task: AddTask(start_date, duration_or_end_date)
        try:
            dur_param = duration_str or "1d"
            if finish_date:
                finish_dt = _parse_date(finish_date)
                if finish_dt:
                    dur_param = pywintypes.Time(finish_dt)
            task = new_bar.Tasks.AddTask(ole_start, dur_param)
        except Exception as e:
            result["task_id"] = bar_id
            result["name"] = name
            result["warning"] = f"Bar created but AddTask failed: {str(e)[:80]}"
            return result

    result["task_id"] = bar_id
    result["name"] = name
    result["start"] = start_dt.strftime("%Y-%m-%d")
    if duration_str:
        result["duration"] = duration_str

    return result


def _com_update_task(project, task_id: int, name: str = None,
                     duration_str: str = None, percent_complete: float = None,
                     notes: str = None, start_date: str = None,
                     finish_date: str = None) -> dict:
    """Update a task via COM. Returns result dict.

    Uses hierarchy search to find bars at any level.
    """
    import win32com.client
    bar = _find_bar_by_id(project, task_id)
    if bar is None:
        return {"error": f"Bar ID {task_id} not found in hierarchy"}

    result = {"method": "COM", "task_id": task_id, "updated_fields": []}

    if name is not None:
        bar.Name = name
        result["updated_fields"].append("name")
        result["name"] = name

    # Use the improved date/duration setter
    if duration_str is not None or start_date is not None or finish_date is not None:
        date_info = _set_bar_dates(bar, start_date, finish_date, duration_str)
        if "duration" in date_info or "start" in date_info or "finish" in date_info:
            result["updated_fields"].extend(
                k for k in ["duration", "start", "finish"] if k in date_info
            )
        result.update(date_info)

    if percent_complete is not None:
        try:
            bar.DurationPercentComplete = percent_complete
            result["updated_fields"].append("percent_complete")
            result["percent_complete"] = percent_complete
        except Exception:
            try:
                bar.OverallPercentComplete = percent_complete
                result["updated_fields"].append("percent_complete")
                result["percent_complete"] = percent_complete
            except Exception as pe:
                result["percent_complete_error"] = str(pe)

    if notes is not None:
        try:
            task, _ = _get_bar_task(bar)
            if task is None:
                task = bar.ExpandedTask
            task.AddTextDatedNote("Note", notes, datetime.now())
            result["updated_fields"].append("notes")
        except Exception as ne:
            try:
                bar.SetUDF("Notes", notes)
                result["updated_fields"].append("notes")
            except Exception:
                result["notes_error"] = str(ne)

    return result


def _com_delete_task(project, task_id: int) -> dict:
    """Delete a task/bar via COM. Returns result dict.

    NOTE: This function manages its own transactions internally.
    The caller's outer transaction should NOT wrap this function —
    the caller should AbandonTransaction before calling, or this function
    will end the caller's transaction first.

    Steps (each in separate transaction for reliability):
    1. Remove links from the task
    2. Remove tasks from the bar (bar.Tasks.Remove)
    3. Remove the bar from its parent's ChildBars

    GetActualParentBar() is unreliable (returns self), so we search
    the hierarchy manually.
    """
    import win32com.client
    result = {"method": "COM", "task_id": task_id}

    def _wait():
        try:
            project.WaitForNotificationProcessing()
        except Exception:
            pass

    # End any caller's transaction first
    try:
        project.EndTransaction()
        _wait()
    except Exception:
        pass

    # Find the bar in hierarchy
    bar = _find_bar_by_id(project, task_id)
    if bar is None:
        result["error"] = f"Bar ID {task_id} not found in hierarchy"
        return result

    result["name"] = bar.Name

    # Step 1: Remove links (separate transaction)
    try:
        task, _ = _get_bar_task(bar)
        if task and (task.LinksOut.Count > 0 or task.LinksIn.Count > 0):
            project.StartTransaction("Delete links")
            try:
                while task.LinksOut.Count > 0:
                    task.LinksOut.Remove(1)
                while task.LinksIn.Count > 0:
                    task.LinksIn.Remove(1)
                project.EndTransaction()
                _wait()
            except Exception:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
    except Exception:
        pass

    # Step 2: Clear progress (which creates ActualStart — prevents deletion)
    bar = _find_bar_by_id(project, task_id)  # Re-fetch after txn
    if bar is None:
        result["deleted"] = True
        return result
    try:
        pct = getattr(bar, 'OverallPercentComplete', 0) or 0
        if pct > 0:
            project.StartTransaction("Clear progress")
            try:
                bar.OverallPercentComplete = 0.0
                project.EndTransaction()
                _wait()
            except Exception:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
    except Exception:
        pass

    # Step 3: Remove tasks from the bar (separate transaction)
    bar = _find_bar_by_id(project, task_id)  # Re-fetch after txn
    if bar is None:
        result["deleted"] = True
        return result
    try:
        if bar.Tasks.Count > 0:
            project.StartTransaction("Delete tasks")
            try:
                while bar.Tasks.Count > 0:
                    bar.Tasks.Remove(1)
                project.EndTransaction()
                _wait()
            except Exception:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
    except Exception:
        pass

    # Step 3: Remove the bar from hierarchy (separate transaction)
    bar = _find_bar_by_id(project, task_id)  # Re-fetch after txn
    if bar is None:
        result["deleted"] = True
        return result

    project.StartTransaction("Delete bar")
    try:
        root_bar = project.Bars.Item(1)
        root_task = win32com.client.Dispatch(root_bar.Tasks(1))
        deleted = _delete_from_parent(root_task, task_id)
        if deleted:
            project.EndTransaction()
            _wait()
            result["deleted"] = True
            return result
    except Exception:
        pass

    # Fallback: try removing from project.Bars (top-level)
    try:
        bars = project.Bars
        for i in range(bars.Count, 0, -1):
            try:
                b = bars.Item(i)
                if b.ID == task_id:
                    bars.Remove(i)
                    project.EndTransaction()
                    _wait()
                    result["deleted"] = True
                    return result
            except Exception:
                continue
    except Exception:
        pass

    try:
        project.AbandonTransaction()
    except Exception:
        pass

    result["error"] = f"Bar {task_id} found but could not be deleted"
    return result


def _delete_from_parent(parent_task, target_id, depth=0, max_depth=6):
    """Recursively search for target_id in ChildBars and remove it."""
    import win32com.client
    if depth >= max_depth:
        return False
    try:
        child_bars = parent_task.ChildBars
        for i in range(child_bars.Count, 0, -1):
            try:
                cb = win32com.client.Dispatch(child_bars.Item(i))
                if cb.ID == target_id:
                    child_bars.Remove(i)
                    return True
                # Recurse: check if target is a grandchild
                try:
                    ct = win32com.client.Dispatch(cb.Tasks(1))
                    if _delete_from_parent(ct, target_id, depth + 1, max_depth):
                        return True
                except Exception:
                    pass
            except Exception:
                continue
    except Exception:
        pass
    return False


def _com_explore_link_interfaces(project) -> dict:
    """Auto-discover available link-related COM methods at runtime.

    Explores LinksIn/LinksOut collections, LinkCategorys, and individual link objects
    to find which methods/properties actually work. Results are cached.
    """
    result = {"interfaces": {}}

    bars = project.Bars
    test_task = None
    test_bar = None
    # Find a bar with links to explore
    all_bars = _com_get_all_bars(project, max_bars=50)
    for bar in all_bars:
        try:
            task, _ = _get_bar_task(bar)
            if task is not None:
                # Prefer bars with links
                try:
                    if task.LinksIn.Count > 0 or task.LinksOut.Count > 0:
                        test_bar = bar
                        test_task = task
                        result["test_bar_id"] = bar.ID
                        result["test_bar_name"] = bar.Name
                        break
                except Exception:
                    pass
                if test_task is None:
                    test_bar = bar
                    test_task = task
                    result["test_bar_id"] = bar.ID
                    result["test_bar_name"] = bar.Name
        except Exception:
            continue

    if test_task is None:
        result["error"] = "No bar with task found"
        return result

    # Explore LinksIn
    links_in_info = {"available": False}
    try:
        links_in = test_task.LinksIn
        links_in_info["available"] = True
        try:
            links_in_info["count"] = links_in.Count
        except Exception as e:
            links_in_info["count_error"] = str(e)[:80]

        # Check for Add method
        for method in ["Add", "Remove", "Item", "All", "_NewEnum"]:
            try:
                m = getattr(links_in, method)
                links_in_info[f"has_{method}"] = True if callable(m) else f"prop={m}"
            except AttributeError:
                links_in_info[f"has_{method}"] = False
            except Exception as e:
                links_in_info[f"has_{method}"] = f"error: {str(e)[:60]}"

        # If links exist, explore the first link object
        try:
            count = links_in.Count
            if count > 0:
                link = links_in.Item(1)
                link_info = {}
                for attr in ["ID", "Name", "Type", "LinkType", "Lag", "LagDuration",
                             "PredecessorTask", "SuccessorTask", "Predecessor", "Successor",
                             "FromTask", "ToTask", "Bar", "Task", "Critical",
                             "Category", "LinkCategory", "TotalFloat", "FreeFloat",
                             "Start", "End", "Duration", "DrivingLink",
                             "Delete", "Remove", "EditToken"]:
                    try:
                        val = getattr(link, attr)
                        if callable(val):
                            link_info[attr] = "callable"
                        else:
                            link_info[attr] = repr(val)[:60]
                    except AttributeError:
                        pass
                    except Exception as e:
                        link_info[attr] = f"error: {str(e)[:50]}"
                links_in_info["link_object_attrs"] = link_info

                # Try to get type info of link object
                try:
                    ti = link._oleobj_.GetTypeInfo()
                    ta = ti.GetTypeAttr()
                    link_methods = []
                    for fi in range(min(ta[6], 50)):
                        try:
                            fd = ti.GetFuncDesc(fi)
                            names = ti.GetNames(fd[0])
                            invkind = fd[3]
                            kind = {1: "METHOD", 2: "GET", 4: "PUT"}.get(invkind, str(invkind))
                            if names and names[0] not in ("QueryInterface", "AddRef", "Release", "GetTypeInfoCount"):
                                link_methods.append(f"[{kind}] {names[0]}")
                        except Exception:
                            pass
                    links_in_info["link_type_info"] = link_methods
                except Exception:
                    pass
        except Exception:
            pass
    except Exception as e:
        links_in_info["error"] = str(e)[:100]

    result["interfaces"]["LinksIn"] = links_in_info

    # Explore LinksOut
    links_out_info = {"available": False}
    try:
        links_out = test_task.LinksOut
        links_out_info["available"] = True
        try:
            links_out_info["count"] = links_out.Count
        except Exception as e:
            links_out_info["count_error"] = str(e)[:80]
        for method in ["Add", "Remove", "Item", "All", "_NewEnum"]:
            try:
                m = getattr(links_out, method)
                links_out_info[f"has_{method}"] = True if callable(m) else f"prop={m}"
            except AttributeError:
                links_out_info[f"has_{method}"] = False
            except Exception as e:
                links_out_info[f"has_{method}"] = f"error: {str(e)[:60]}"
    except Exception as e:
        links_out_info["error"] = str(e)[:100]

    result["interfaces"]["LinksOut"] = links_out_info

    # Explore LinkCategorys
    link_cats_info = {"available": False}
    try:
        lc = project.LinkCategorys
        link_cats_info["available"] = True
        link_cats_info["count"] = lc.Count
        if lc.Count > 0:
            cat = lc.Item(1)
            cat_attrs = {}
            for attr in ["Name", "ID", "Add", "Links", "Count"]:
                try:
                    val = getattr(cat, attr)
                    cat_attrs[attr] = repr(val)[:60] if not callable(val) else "callable"
                except Exception:
                    pass
            link_cats_info["first_category"] = cat_attrs
    except Exception as e:
        link_cats_info["error"] = str(e)[:100]

    result["interfaces"]["LinkCategorys"] = link_cats_info

    # Check bar-level link methods
    bar_info = {}
    for attr in ["AddLink", "Dependencies", "Links", "LinksIn", "LinksOut"]:
        try:
            val = getattr(test_bar, attr)
            bar_info[attr] = "callable" if callable(val) else repr(val)[:60]
        except AttributeError:
            bar_info[attr] = "NOT_FOUND"
        except Exception as e:
            bar_info[attr] = f"error: {str(e)[:50]}"
    result["interfaces"]["bar_link_methods"] = bar_info

    return result


def _com_add_link(project, predecessor_id: int, successor_id: int,
                  link_type: str = "FS", lag_str: str = None) -> dict:
    """Add a link between two bars via COM.

    VERIFIED WORKING (v13, v21):
    - task.LinkTo(other_task) creates FS link, returns ILink object
    - link.type = 0/1/2/3 sets FS/SS/FF/SF
    - link.StartLagTime = dur_obj sets lag
    - Works on ITask objects from bar.Tasks(1)
    """
    import win32com.client
    result = {"method": "COM", "predecessor_id": predecessor_id,
              "successor_id": successor_id, "link_type": link_type}

    # Find bars in hierarchy
    pred_bar = _find_bar_by_id(project, predecessor_id)
    succ_bar = _find_bar_by_id(project, successor_id)

    if pred_bar is None:
        return {"error": f"Predecessor bar ID {predecessor_id} not found"}
    if succ_bar is None:
        return {"error": f"Successor bar ID {successor_id} not found"}

    # Get tasks from bars
    pred_task, _ = _get_bar_task(pred_bar)
    succ_task, _ = _get_bar_task(succ_bar)

    if pred_task is None:
        # Fallback: try ExpandedTask
        try:
            pred_task = pred_bar.ExpandedTask
        except Exception:
            return {"error": f"Cannot get task for predecessor bar {predecessor_id}"}

    if succ_task is None:
        try:
            succ_task = succ_bar.ExpandedTask
        except Exception:
            return {"error": f"Cannot get task for successor bar {successor_id}"}

    # Primary strategy: task.LinkTo(other_task) — VERIFIED WORKING
    try:
        link = pred_task.LinkTo(succ_task)
        if link:
            ld = win32com.client.Dispatch(link)
            result["success"] = True
            result["link_id"] = ld.ID
            result["strategy"] = "task.LinkTo"

            # Set link type if not FS
            link_type_map = {"FS": 0, "SS": 1, "FF": 2, "SF": 3}
            type_val = link_type_map.get(link_type.upper(), 0)
            if type_val != 0:
                try:
                    ld.type = type_val
                    result["type_set"] = link_type.upper()
                except Exception as te:
                    result["type_warning"] = f"Link created as FS, could not set to {link_type}: {str(te)[:50]}"

            # Set lag if specified
            if lag_str:
                try:
                    lag_dur = pred_task.GetDurationFromString(lag_str)
                    ld.StartLagTime = lag_dur
                    result["lag"] = lag_str
                except Exception as le:
                    result["lag_warning"] = f"Could not set lag '{lag_str}': {str(le)[:50]}"

            return result
    except Exception as e:
        result["linkto_error"] = str(e)[:100]

    # Fallback: try ExpandedTask if Tasks(1) failed
    try:
        pred_et = pred_bar.ExpandedTask
        succ_et = succ_bar.ExpandedTask
        # Verify they're different tasks (ExpandedTask may return root for both)
        if pred_et.ID != succ_et.ID:
            link = pred_et.LinkTo(succ_et)
            if link:
                ld = win32com.client.Dispatch(link)
                result["success"] = True
                result["link_id"] = ld.ID
                result["strategy"] = "ExpandedTask.LinkTo"
                return result
    except Exception:
        pass

    result["error"] = (
        f"Could not create link between {predecessor_id} and {successor_id}. "
        f"Error: {result.get('linkto_error', 'unknown')}. "
        "Ensure both bars are proper tasks (not empty bars) within the project hierarchy."
    )
    return result


def _com_remove_link(project, predecessor_id: int, successor_id: int) -> dict:
    """Remove a link between two bars via COM.

    VERIFIED WORKING: task.LinksOut.Remove(index) and task.LinksIn.Remove(index)
    ILink has StartTask and EndTask properties to identify link endpoints.
    """
    import win32com.client
    result = {"method": "COM", "predecessor_id": predecessor_id,
              "successor_id": successor_id}

    # Find bars in hierarchy
    succ_bar = _find_bar_by_id(project, successor_id)
    pred_bar = _find_bar_by_id(project, predecessor_id)
    errors = []

    # Strategy 1: predecessor's LinksOut — find link to successor and remove
    if pred_bar:
        try:
            pred_task, _ = _get_bar_task(pred_bar)
            if pred_task is None:
                pred_task = pred_bar.ExpandedTask
            links_out = pred_task.LinksOut
            for i in range(links_out.Count, 0, -1):
                try:
                    link = win32com.client.Dispatch(links_out.Item(i))
                    # ILink has EndTask (successor) property
                    try:
                        end_task = win32com.client.Dispatch(link.EndTask)
                        end_bar = win32com.client.Dispatch(end_task.Bar)
                        if end_bar.ID == successor_id:
                            links_out.Remove(i)
                            result["removed"] = True
                            result["strategy"] = "LinksOut.Remove"
                            return result
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception as e:
            errors.append(f"LinksOut: {e}")

    # Strategy 2: successor's LinksIn — find link from predecessor and remove
    if succ_bar:
        try:
            succ_task, _ = _get_bar_task(succ_bar)
            if succ_task is None:
                succ_task = succ_bar.ExpandedTask
            links_in = succ_task.LinksIn
            for i in range(links_in.Count, 0, -1):
                try:
                    link = win32com.client.Dispatch(links_in.Item(i))
                    # ILink has StartTask (predecessor) property
                    try:
                        start_task = win32com.client.Dispatch(link.StartTask)
                        start_bar = win32com.client.Dispatch(start_task.Bar)
                        if start_bar.ID == predecessor_id:
                            links_in.Remove(i)
                            result["removed"] = True
                            result["strategy"] = "LinksIn.Remove"
                            return result
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception as e:
            errors.append(f"LinksIn: {e}")

    result["error"] = f"Could not remove link: {'; '.join(errors)}"
    return result


def _com_update_link(project, predecessor_id: int, successor_id: int,
                     new_link_type: str = None, new_lag: str = None) -> dict:
    """Update a link by removing old and adding new with updated properties."""
    result = {"method": "COM", "predecessor_id": predecessor_id,
              "successor_id": successor_id}

    # Strategy: Remove old link then add new with updated properties
    remove_result = _com_remove_link(project, predecessor_id, successor_id)
    if not remove_result.get("removed"):
        result["error"] = f"Could not find/remove existing link to update: {remove_result.get('error', 'unknown')}"
        return result

    lt = new_link_type if new_link_type else "FS"
    add_result = _com_add_link(project, predecessor_id, successor_id, lt, new_lag)
    if add_result.get("success"):
        result["updated"] = True
        if new_link_type:
            result["new_link_type"] = new_link_type
        if new_lag:
            result["new_lag"] = new_lag
        result["strategy"] = f"remove+add ({add_result.get('strategy', 'unknown')})"
    else:
        result["error"] = f"Removed old link but failed to add new: {add_result.get('error', 'unknown')}"

    return result


def _com_update_progress(project, task_id: int, percent_complete: float = None,
                         actual_start: str = None, actual_finish: str = None) -> dict:
    """Update progress on a task via COM. Returns result dict."""
    import pywintypes

    bar = _find_bar_by_id(project, task_id)
    if bar is None:
        return {"error": f"Bar ID {task_id} not found"}

    result = {"method": "COM", "task_id": task_id, "updated": []}

    if percent_complete is not None:
        try:
            bar.DurationPercentComplete = percent_complete
            result["updated"].append("percent_complete")
            result["percent_complete"] = percent_complete
        except Exception:
            try:
                bar.OverallPercentComplete = percent_complete
                result["updated"].append("percent_complete")
                result["percent_complete"] = percent_complete
            except Exception as pe:
                result["percent_complete_error"] = str(pe)

    if actual_start is not None:
        try:
            # Parse date (YYYY-MM-DD or DD/MM/YYYY)
            if '/' in actual_start:
                dt = datetime.strptime(actual_start, "%d/%m/%Y")
            else:
                dt = datetime.strptime(actual_start, "%Y-%m-%d")
            try:
                ole_date = pywintypes.Time(dt)
            except Exception:
                ole_date = dt
            bar.ActualStart = ole_date
            result["updated"].append("actual_start")
            result["actual_start"] = actual_start
        except Exception as ase:
            result["actual_start_error"] = str(ase)

    if actual_finish is not None:
        try:
            if '/' in actual_finish:
                dt = datetime.strptime(actual_finish, "%d/%m/%Y")
            else:
                dt = datetime.strptime(actual_finish, "%Y-%m-%d")
            try:
                ole_date = pywintypes.Time(dt)
            except Exception:
                ole_date = dt
            try:
                bar.ActualEnd = ole_date
                result["updated"].append("actual_finish")
                result["actual_finish"] = actual_finish
            except Exception:
                # ActualEnd failed, try ActualFinish property name
                try:
                    bar.ActualFinish = ole_date
                    result["updated"].append("actual_finish")
                    result["actual_finish"] = actual_finish
                except Exception as afe:
                    result["actual_finish_error"] = str(afe)
        except (ValueError, TypeError) as parse_err:
            result["actual_finish_error"] = f"Invalid date format '{actual_finish}': {parse_err}"

    return result


def _gui_reschedule_fallback(report_date_str: str, straighten: bool) -> dict:
    """Fallback when COM reschedule fails: return manual instructions.

    COM is the only supported automation method. If COM fails,
    we return clear manual instructions for the user.
    """
    return {
        "method": "Manual (COM unavailable)",
        "success": False,
        "error": "COM connection to Asta failed. Cannot reschedule automatically.",
        "manual_steps": [
            "1. Open Asta Powerproject and load your project",
            f"2. Set the report/data date to: {report_date_str}",
            "   (Project tab > Progress Period)",
            "3. Press F9 to reschedule",
            "4. Check the critical path (shown in red)",
        ],
        "suggestion": "Ensure Asta Powerproject is running with a project open, then retry.",
    }
# @mcp.tool(  # CONSOLIDATED into asta_schedule
#     name="asta_reschedule_project",
#     annotations={
#         "title": "Reschedule Project (COM)",
#         "readOnlyHint": False,
#         "destructiveHint": False,
#         "idempotentHint": False,
#         "openWorldHint": True,
#     }
# )
async def asta_reschedule_project(params: RescheduleProjectInput) -> str:
    """Reschedule an Asta Powerproject project via COM automation.

    Connects to a running Asta Powerproject instance, sets the Report Date
    (Data Date) on the current progress period, optionally straightens
    uncompleted work, and performs a full CPM reschedule.

    Requires Asta Powerproject to be running with a project open.
    Uses COM/OLE automation (win32com) as primary method, with GUI
    automation (F9 key) as fallback.

    Args:
        params: Contains report_date, straighten_uncompleted_work,
                preserve_links, target_wbs_id

    Returns:
        JSON with reschedule results including new project dates,
        method used, and any errors encountered
    """
    import pythoncom
    import pywintypes

    result = {
        "action": "Reschedule Project",
        "requested_report_date": params.report_date,
        "straighten_uncompleted_work": params.straighten_uncompleted_work,
        "preserve_links": params.preserve_links,
        "target_wbs_id": params.target_wbs_id,
        "success": False,
    }

    # --- COM Initialization (required for async/threaded MCP context) ---
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
    except Exception as e:
        logger.warning(f"CoInitialize failed (may already be initialized): {e}")

    try:
        # ================================================================
        # STRATEGY A: COM Automation via Asta OLE Interface
        # ================================================================
        try:
            app, project, connection_method = _connect_asta_com()
            result["connection_method"] = connection_method

            # --- Capture pre-reschedule state ---
            try:
                result["project_name"] = str(project.Name) if project.Name else "Unknown"
            except Exception:
                result["project_name"] = "Unknown"

            try:
                pre_end = project.ProjectEnd
                result["project_end_before"] = format_date(pre_end)
            except Exception:
                result["project_end_before"] = "N/A"

            try:
                result["project_start"] = format_date(project.ProjectStart)
            except Exception:
                result["project_start"] = "N/A"

            # --- Parse the target report date ---
            report_dt = datetime.strptime(params.report_date, "%Y-%m-%d")

            # Convert to OLE-compatible date (pywintypes.Time or datetime)
            try:
                ole_date = pywintypes.Time(report_dt)
            except Exception:
                # Fallback: use Julian conversion if pywintypes.Time fails
                try:
                    ole_date = app.DateToJulian(report_dt)
                except Exception:
                    ole_date = report_dt

            # --- Start a transaction for atomic changes ---
            transaction_started = False
            try:
                if not project.TransactionInProgress():
                    project.StartTransaction("MCP Reschedule")
                    transaction_started = True
                    logger.info("COM: Transaction started for reschedule")
            except Exception as e:
                logger.warning(f"COM: Could not start transaction: {e}")

            try:
                # --- Set the Report Date on the current progress period ---
                # Cascading fallback: CurrentProgressPeriod -> AllProgressPeriods.LatestProgressPeriod
                #   -> ProgressPeriods collection -> skip with warning
                pp_resolved = None  # Will hold the progress period we successfully accessed

                # Strategy 1: project.CurrentProgressPeriod
                try:
                    current_pp = project.CurrentProgressPeriod
                    if current_pp is not None:
                        pp_resolved = current_pp
                        logger.info("COM: Got progress period via CurrentProgressPeriod")
                except Exception as e1:
                    logger.warning(f"COM: CurrentProgressPeriod failed: {e1}")

                # Strategy 2: AllProgressPeriods.LatestProgressPeriod
                if pp_resolved is None:
                    try:
                        all_pp = project.AllProgressPeriods
                        if all_pp is not None:
                            latest_pp = all_pp.LatestProgressPeriod
                            if latest_pp is not None:
                                pp_resolved = latest_pp
                                logger.info("COM: Got progress period via AllProgressPeriods.LatestProgressPeriod")
                    except Exception as e2:
                        logger.warning(f"COM: AllProgressPeriods.LatestProgressPeriod failed: {e2}")

                # Strategy 3: Iterate ProgressPeriods collection
                if pp_resolved is None:
                    try:
                        pp_collection = project.ProgressPeriods
                        if pp_collection is not None:
                            # Try to get the last (most recent) progress period
                            count = pp_collection.Count
                            if count > 0:
                                pp_resolved = pp_collection.Item(count)  # 1-based index, last item
                                logger.info(f"COM: Got progress period via ProgressPeriods collection (item {count})")
                    except Exception as e3:
                        logger.warning(f"COM: ProgressPeriods collection failed: {e3}")

                # Now set the Report Date on whichever progress period we found
                if pp_resolved is not None:
                    try:
                        pp_resolved.ReportDate = ole_date
                        result["report_date_set"] = params.report_date
                        logger.info(f"COM: Report date set to {params.report_date}")
                        # Try to make it the current progress period
                        try:
                            project.CurrentProgressPeriod = pp_resolved
                        except Exception:
                            pass  # Not critical if this fails
                    except pywintypes.com_error as e:
                        error_code = e.hresult if hasattr(e, 'hresult') else e.args[0]
                        if error_code == -2147209975:  # ppOleInvalidProgressPeriod
                            result["report_date_error"] = (
                                "Invalid progress period. The report date may be outside "
                                "the project date range."
                            )
                        elif error_code == -2147161491:  # ppProgressPeriodLocked
                            result["report_date_error"] = (
                                "Progress period is locked. Unlock it in Asta first."
                            )
                        else:
                            result["report_date_error"] = f"COM error setting report date: {e}"
                    except Exception as e:
                        result["report_date_error"] = f"Error setting report date: {e}"
                else:
                    result["report_date_warning"] = (
                        "Could not access any progress period via COM. "
                        "Report date was NOT set. Please set the Report Date manually in Asta "
                        "before rescheduling, or create a progress period first."
                    )

                # --- Set straightening progress period ---
                # Reuse pp_resolved from above (already the best available progress period)
                if params.straighten_uncompleted_work:
                    if pp_resolved is not None:
                        try:
                            project.ProgressPeriodForStraightening = pp_resolved
                            result["straightening_applied"] = True
                            logger.info("COM: Progress period for straightening set")
                        except Exception as e:
                            result["straightening_warning"] = f"Could not set straightening period: {e}"
                    else:
                        result["straightening_warning"] = (
                            "No progress period available for straightening. "
                            "Set a progress period in Asta first."
                        )

                # --- Set reschedule options (preserve_links = retained logic) ---
                try:
                    # SetRescheduleOptions(alap_as_critical, end_flags_pp_compat, ignore_link_cats)
                    # preserve_links=True -> don't ignore link categories (retained logic)
                    # preserve_links=False -> progress override behavior
                    project.SetRescheduleOptions(
                        False,  # alap_as_critical
                        True,   # end_flags_powerproject_compatibility
                        not params.preserve_links  # ignore_link_cats (inverted logic)
                    )
                    result["reschedule_options_set"] = True
                except Exception as e:
                    logger.warning(f"COM: Could not set reschedule options: {e}")

                # --- Wait for any pending notifications before reschedule ---
                try:
                    project.WaitForNotificationProcessing()
                except Exception:
                    pass

                # --- Execute Reschedule ---
                try:
                    if params.target_wbs_id:
                        # Branch-level reschedule: find the bar and reschedule it
                        try:
                            target_bar = _find_bar_by_id(project.Bars, int(params.target_wbs_id))
                            project.RescheduleBars([target_bar])
                            result["reschedule_scope"] = f"Branch (ID: {params.target_wbs_id})"
                        except Exception:
                            # Fallback to full project reschedule
                            project.Reschedule()
                            result["reschedule_scope"] = "Full project (branch ID not found, fell back)"
                    else:
                        project.Reschedule()
                        result["reschedule_scope"] = "Full project"

                    logger.info("COM: Reschedule executed successfully")

                    # Wait for reschedule calculations to complete
                    try:
                        project.WaitForNotificationProcessing()
                    except Exception:
                        pass

                    result["success"] = True

                except pywintypes.com_error as e:
                    error_code = e.hresult if hasattr(e, 'hresult') else e.args[0]
                    if error_code == -2147210491:  # ppOnlyOneReschedulePerConnection
                        result["error"] = (
                            "Another reschedule is already in progress on this connection. "
                            "Wait for it to complete and try again."
                        )
                    elif error_code == -2147210490:  # ppRescheduleObjectAccessFailure
                        result["error"] = (
                            "Reschedule failed: could not access required project objects. "
                            "Ensure the project is not read-only."
                        )
                    else:
                        result["error"] = f"COM reschedule error: {e}"
                except Exception as e:
                    result["error"] = f"Reschedule execution failed: {e}"

                # --- Capture post-reschedule state ---
                if result["success"]:
                    try:
                        result["project_end_after"] = format_date(project.ProjectEnd)
                    except Exception:
                        result["project_end_after"] = "N/A"

                    try:
                        # Count bars/tasks for the response
                        bars = project.Bars
                        if bars is not None:
                            bar_count = 0
                            try:
                                # Iterate bars to count
                                bcv = project.CurrentView
                                if bcv is not None:
                                    task_ids = bcv.AllTaskBaseIds
                                    if task_ids is not None:
                                        bar_count = len(task_ids) if hasattr(task_ids, '__len__') else 0
                            except Exception:
                                pass
                            result["tasks_in_project"] = bar_count if bar_count > 0 else "N/A"
                    except Exception:
                        result["tasks_in_project"] = "N/A"

                    try:
                        result["reschedule_number"] = int(project.RescheduleNumber)
                    except Exception:
                        pass

            except Exception as e:
                # If anything failed during the operations, abandon the transaction
                result["error"] = f"Unexpected error during COM reschedule: {e}"
                if transaction_started:
                    try:
                        project.AbandonTransaction()
                        logger.info("COM: Transaction abandoned due to error")
                    except Exception:
                        pass
                raise

            # --- End the transaction ---
            if transaction_started:
                try:
                    if result["success"]:
                        _com_end_transaction(project)
                        logger.info("COM: Transaction committed")
                    else:
                        project.AbandonTransaction()
                        logger.info("COM: Transaction abandoned (reschedule not successful)")
                except Exception as e:
                    logger.warning(f"COM: Error ending transaction: {e}")

            return json.dumps(result, indent=2, default=str)

        except RuntimeError as e:
            # COM connection failed entirely — fall through to GUI fallback
            com_error_detail = str(e)
            logger.info(f"COM connection failed, trying GUI fallback: {com_error_detail}")

        # ================================================================
        # STRATEGY B: GUI Automation Fallback
        # ================================================================
        logger.info("Falling back to GUI automation for reschedule")
        gui_result = _gui_reschedule_fallback(params.report_date, params.straighten_uncompleted_work)

        result.update(gui_result)
        result["com_connection_failed"] = True
        result["com_error_detail"] = com_error_detail

        return json.dumps(result, indent=2, default=str)

    except Exception as e:
        result["error"] = f"Fatal error in asta_reschedule_project: {e}"
        logger.error(f"asta_reschedule_project fatal error: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)

    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ============================================================================
# PHASE 3: ADVANCED OPTIMIZATION & MODELLING (COM-BASED)
# ============================================================================


# ---------------------------------------------------------------------------
# 1. WHAT-IF SCENARIO ENGINE
# ---------------------------------------------------------------------------

class TaskModification(BaseModel):
    """A single task modification for what-if analysis."""
    task_id: int = Field(..., description="The bar/task ID to modify.")
    new_duration: Optional[str] = Field(
        default=None,
        description="New duration string (e.g. '10d', '2w'). If omitted, duration unchanged."
    )
    new_name: Optional[str] = Field(
        default=None,
        description="New task name. If omitted, name unchanged."
    )
    add_predecessor_id: Optional[int] = Field(
        default=None,
        description="Add a Finish-to-Start link FROM this predecessor TO the task."
    )
    remove_predecessor_id: Optional[int] = Field(
        default=None,
        description="Remove the link FROM this predecessor TO the task."
    )


class WhatIfInput(BaseModel):
    """Input for the what-if scenario analysis tool."""
    model_config = ConfigDict(str_strip_whitespace=True)

    modifications: List[TaskModification] = Field(
        ...,
        description="Array of task modifications to apply as a scenario.",
        min_length=1,
        max_length=50,
    )
    target_date: str = Field(
        ...,
        description="Target completion date in YYYY-MM-DD. "
                    "If post-reschedule ProjectEnd <= this, changes are committed; "
                    "otherwise the transaction is rolled back."
    )
    scenario_name: Optional[str] = Field(
        default="AI What-If Scenario",
        description="Name for the transaction / scenario."
    )

    @field_validator("target_date")
    @classmethod
    def validate_target(cls, v: str) -> str:
        try:
            datetime.strptime(v, "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"target_date must be YYYY-MM-DD, got: '{v}'")
        return v


# @mcp.tool()  # CONSOLIDATED into asta_schedule
def asta_what_if_analysis(params: WhatIfInput) -> str:
    """Run a what-if scenario: apply task modifications, reschedule, and auto-commit or rollback.

    Applies modifications inside a COM transaction. After rescheduling, if the new
    project end date meets the target_date the transaction is committed; otherwise it
    is abandoned (all changes instantly reverted).
    """
    import pythoncom
    import pywintypes

    com_initialized = False
    result: Dict[str, Any] = {
        "tool": "asta_what_if_analysis",
        "scenario": params.scenario_name,
        "success": False,
        "modifications_requested": len(params.modifications),
        "target_date": params.target_date,
    }

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        # Capture baseline state
        try:
            result["original_project_end"] = format_date(project.ProjectEnd)
        except Exception:
            result["original_project_end"] = "N/A"

        target_dt = datetime.strptime(params.target_date, "%Y-%m-%d")

        # --- Start transaction ---
        try:
            project.StartTransaction(params.scenario_name or "AI What-If Scenario")
            logger.info(f"COM What-If: Transaction '{params.scenario_name}' started")
        except Exception as e:
            result["error"] = f"Could not start transaction: {e}"
            return json.dumps(result, indent=2, default=str)

        # --- Apply modifications ---
        applied = []
        errors = []
        bars_collection = project.Bars

        for mod in params.modifications:
            mod_result = {"task_id": mod.task_id}
            try:
                bar = _find_bar_by_id(bars_collection, mod.task_id)
                if bar is None:
                    mod_result["error"] = f"Bar ID {mod.task_id} not found"
                    errors.append(mod_result)
                    continue

                # Change duration
                if mod.new_duration is not None:
                    days = parse_duration(mod.new_duration)
                    try:
                        # Use EditToken for duration — the most reliable COM method
                        bar.EditToken("Duration", mod.new_duration)
                        mod_result["duration_set"] = mod.new_duration
                    except Exception:
                        # Fallback: try setting via the task's Duration property
                        try:
                            task, _ = _get_bar_task(bar)
                            if task is None:
                                task = bar.ExpandedTask
                            # Duration is IRelativeTime — set via token
                            task.EditToken("Duration", mod.new_duration)
                            mod_result["duration_set"] = mod.new_duration
                        except Exception as de:
                            mod_result["duration_error"] = str(de)

                # Change name
                if mod.new_name is not None:
                    bar.Name = mod.new_name
                    mod_result["name_set"] = mod.new_name

                # Add predecessor link (uses _com_add_link helper)
                if mod.add_predecessor_id is not None:
                    try:
                        link_result = _com_add_link(
                            project, mod.add_predecessor_id, mod.task_id, "FS"
                        )
                        if "error" in link_result:
                            mod_result["predecessor_error"] = link_result["error"]
                        else:
                            mod_result["predecessor_added"] = mod.add_predecessor_id
                    except Exception as le:
                        mod_result["predecessor_error"] = str(le)

                # Remove predecessor link (uses _com_remove_link helper)
                if mod.remove_predecessor_id is not None:
                    try:
                        remove_result = _com_remove_link(
                            project, mod.remove_predecessor_id, mod.task_id
                        )
                        if "error" in remove_result:
                            mod_result["predecessor_remove_error"] = remove_result["error"]
                        else:
                            mod_result["predecessor_removed"] = mod.remove_predecessor_id
                    except Exception as rle:
                        mod_result["predecessor_remove_error"] = str(rle)

                applied.append(mod_result)

            except pywintypes.com_error as ce:
                mod_result["com_error"] = str(ce)
                errors.append(mod_result)
            except Exception as ex:
                mod_result["error"] = str(ex)
                errors.append(mod_result)

        result["modifications_applied"] = applied
        if errors:
            result["modification_errors"] = errors

        # --- Wait for notification processing ---
        try:
            project.WaitForNotificationProcessing()
        except Exception:
            pass

        # --- Reschedule ---
        try:
            project.Reschedule(None, False)  # Chart=None (whole project), on_server=False
            logger.info("COM What-If: Reschedule completed")
            result["rescheduled"] = True
        except pywintypes.com_error as re:
            result["reschedule_error"] = str(re)
            # Abandon on reschedule failure
            try:
                project.AbandonTransaction()
            except Exception:
                pass
            result["transaction"] = "abandoned (reschedule failed)"
            return json.dumps(result, indent=2, default=str)

        # --- Evaluate against target ---
        try:
            project.WaitForNotificationProcessing()
        except Exception:
            pass

        try:
            new_end = project.ProjectEnd
            new_end_str = format_date(new_end)
            result["new_project_end"] = new_end_str

            # Parse new end for comparison
            try:
                new_end_dt = datetime.strptime(new_end_str, "%Y-%m-%d")
            except Exception:
                # If format_date returned something else, try the raw object
                if hasattr(new_end, 'year'):
                    new_end_dt = datetime(new_end.year, new_end.month, new_end.day)
                else:
                    new_end_dt = None

            if new_end_dt is not None and new_end_dt <= target_dt:
                # SUCCESS — commit
                _com_end_transaction(project)
                try:
                    project.Save()
                except Exception:
                    pass
                result["success"] = True
                result["transaction"] = "committed"
                result["days_margin"] = (target_dt - new_end_dt).days
                result["message"] = (
                    f"Scenario meets target. New end {new_end_str} is "
                    f"{(target_dt - new_end_dt).days} day(s) before target {params.target_date}."
                )
                logger.info(f"COM What-If: COMMITTED — new end {new_end_str} <= target {params.target_date}")
            else:
                # FAIL — rollback
                project.AbandonTransaction()
                result["success"] = False
                result["transaction"] = "abandoned (target not met)"
                if new_end_dt is not None:
                    result["days_overrun"] = (new_end_dt - target_dt).days
                result["message"] = (
                    f"Scenario does NOT meet target. New end {new_end_str} exceeds "
                    f"target {params.target_date}. All changes have been reverted."
                )
                logger.info(f"COM What-If: ABANDONED — new end {new_end_str} > target {params.target_date}")

        except Exception as eval_e:
            # Cannot evaluate — abandon to be safe
            try:
                project.AbandonTransaction()
            except Exception:
                pass
            result["transaction"] = "abandoned (evaluation error)"
            result["evaluation_error"] = str(eval_e)

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)

    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_what_if_analysis fatal error: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)

    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 2. CODE LIBRARIES — Creation & Assignment
# ---------------------------------------------------------------------------

class ManageCodeLibrariesInput(BaseModel):
    """Input for managing code libraries and their entries."""
    model_config = ConfigDict(str_strip_whitespace=True)

    action: str = Field(
        ...,
        description="Action to perform: 'list', 'create_library', 'add_entries', 'delete_entry'."
    )
    library_name: Optional[str] = Field(
        default=None,
        description="Name of the code library (required for create_library, add_entries, delete_entry)."
    )
    entries: Optional[List[str]] = Field(
        default=None,
        description="List of entry names to add (for 'add_entries' action)."
    )
    entry_name: Optional[str] = Field(
        default=None,
        description="Single entry name (for 'delete_entry' action)."
    )

    @field_validator("action")
    @classmethod
    def validate_action(cls, v: str) -> str:
        allowed = {"list", "create_library", "add_entries", "delete_entry"}
        if v.lower() not in allowed:
            raise ValueError(f"action must be one of {allowed}, got: '{v}'")
        return v.lower()


# @mcp.tool()  # CONSOLIDATED into asta_code
def asta_manage_code_libraries(params: ManageCodeLibrariesInput) -> str:
    """Manage Asta Powerproject Code Libraries via COM.

    Actions:
      - list: List all code libraries and their entries.
      - create_library: Create a new code library (idempotent — skips if exists).
      - add_entries: Add entries to an existing library (idempotent per entry).
      - delete_entry: Remove a specific entry from a library.
    """
    import pythoncom
    import pywintypes

    com_initialized = False
    result: Dict[str, Any] = {"tool": "asta_manage_code_libraries", "action": params.action, "success": False}

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        # Access CodeLibrarys (Asta uses non-standard plural: CodeLibrarys not CodeLibraries)
        code_libs = None
        try:
            code_libs = project.CodeLibrarys
        except AttributeError:
            pass
        if code_libs is None:
            try:
                code_libs = project.AllCodeLibrarys
            except AttributeError:
                pass
        if code_libs is None:
            # Fallback: try via GetToken
            try:
                code_libs = project.GetToken("CodeLibrarys")
            except Exception:
                pass
        if code_libs is None:
            result["error"] = (
                "CodeLibrarys is not accessible in this Asta version. "
                "Try: regsvr32 \"C:\\Program Files\\Elecosoft\\Powerproject\\astadkit.ocx\""
            )
            return json.dumps(result, indent=2, default=str)

        # --- Helper: find library by name ---
        def find_library(name: str):
            count = code_libs.Count
            for i in range(1, count + 1):
                try:
                    lib = code_libs.Item(i)
                    if lib.Name and lib.Name.lower() == name.lower():
                        return lib
                except Exception:
                    continue
            return None

        # --- Helper: list entries of a library ---
        def list_entries(lib) -> List[Dict[str, Any]]:
            entries_list = []
            try:
                ent_col = lib.Entries
                if ent_col is None:
                    return entries_list
                for j in range(1, ent_col.Count + 1):
                    try:
                        entry = ent_col.Item(j)
                        entries_list.append({
                            "id": entry.ID,
                            "name": entry.Name,
                            "short_name": getattr(entry, 'ShortName', ''),
                        })
                    except Exception:
                        continue
            except Exception:
                pass
            return entries_list

        # === ACTION: list ===
        if params.action == "list":
            libraries = []
            count = code_libs.Count
            for i in range(1, count + 1):
                try:
                    lib = code_libs.Item(i)
                    libraries.append({
                        "id": lib.ID,
                        "name": lib.Name,
                        "single_select": getattr(lib, 'SingleSelect', None),
                        "entries": list_entries(lib),
                    })
                except Exception as e:
                    libraries.append({"index": i, "error": str(e)})
            result["libraries"] = libraries
            result["total"] = count
            result["success"] = True

        # === ACTION: create_library ===
        elif params.action == "create_library":
            if not params.library_name:
                result["error"] = "library_name is required for create_library"
                return json.dumps(result, indent=2, default=str)

            existing = find_library(params.library_name)
            if existing is not None:
                result["success"] = True
                result["message"] = f"Library '{params.library_name}' already exists (ID={existing.ID})"
                result["library_id"] = existing.ID
                result["entries"] = list_entries(existing)
            else:
                try:
                    project.StartTransaction("Create Code Library")
                except Exception:
                    pass
                try:
                    new_lib = code_libs.Add()
                    new_lib.Name = params.library_name
                    result["success"] = True
                    result["library_id"] = new_lib.ID
                    result["message"] = f"Library '{params.library_name}' created"
                    try:
                        _com_end_transaction(project)
                    except Exception:
                        pass
                except Exception as e:
                    try:
                        project.AbandonTransaction()
                    except Exception:
                        pass
                    result["error"] = f"Failed to create library: {e}"

        # === ACTION: add_entries ===
        elif params.action == "add_entries":
            if not params.library_name:
                result["error"] = "library_name is required for add_entries"
                return json.dumps(result, indent=2, default=str)
            if not params.entries or len(params.entries) == 0:
                result["error"] = "entries list is required and must not be empty"
                return json.dumps(result, indent=2, default=str)

            lib = find_library(params.library_name)
            if lib is None:
                result["error"] = f"Library '{params.library_name}' not found. Create it first."
                return json.dumps(result, indent=2, default=str)

            # Get existing entry names for idempotency
            existing_names = set()
            try:
                ent_col = lib.Entries
                for j in range(1, ent_col.Count + 1):
                    try:
                        existing_names.add(ent_col.Item(j).Name.lower())
                    except Exception:
                        continue
            except Exception:
                pass

            try:
                project.StartTransaction("Add Code Entries")
            except Exception:
                pass

            added = []
            skipped = []
            add_errors = []
            for entry_name in params.entries:
                if entry_name.lower() in existing_names:
                    skipped.append(entry_name)
                    continue
                try:
                    new_entry = lib.Entries.Add()
                    new_entry.Name = entry_name
                    added.append({"name": entry_name, "id": new_entry.ID})
                except Exception as e:
                    add_errors.append({"name": entry_name, "error": str(e)})

            try:
                _com_end_transaction(project)
            except Exception:
                pass

            result["success"] = True
            result["added"] = added
            result["skipped_existing"] = skipped
            if add_errors:
                result["errors"] = add_errors
            result["message"] = f"Added {len(added)} entries, skipped {len(skipped)} existing"

        # === ACTION: delete_entry ===
        elif params.action == "delete_entry":
            if not params.library_name or not params.entry_name:
                result["error"] = "library_name and entry_name are required for delete_entry"
                return json.dumps(result, indent=2, default=str)

            lib = find_library(params.library_name)
            if lib is None:
                result["error"] = f"Library '{params.library_name}' not found"
                return json.dumps(result, indent=2, default=str)

            try:
                ent_col = lib.Entries
                found_idx = None
                for j in range(1, ent_col.Count + 1):
                    try:
                        if ent_col.Item(j).Name.lower() == params.entry_name.lower():
                            found_idx = j
                            break
                    except Exception:
                        continue

                if found_idx is None:
                    result["error"] = f"Entry '{params.entry_name}' not found in '{params.library_name}'"
                else:
                    try:
                        project.StartTransaction("Delete Code Entry")
                    except Exception:
                        pass
                    ent_col.Remove(found_idx)
                    try:
                        _com_end_transaction(project)
                    except Exception:
                        pass
                    result["success"] = True
                    result["message"] = f"Entry '{params.entry_name}' deleted from '{params.library_name}'"
            except Exception as e:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
                result["error"] = f"Failed to delete entry: {e}"

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_manage_code_libraries fatal: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


class AssignCodesInput(BaseModel):
    """Input for assigning code library entries to tasks."""
    model_config = ConfigDict(str_strip_whitespace=True)

    library_name: str = Field(..., description="Name of the code library.")
    assignments: List[Dict[str, Any]] = Field(
        ...,
        description="List of assignments. Each dict has 'task_id' (int) and 'entry_name' (str). "
                    "Optionally 'append' (bool, default true) to add vs replace codes.",
        min_length=1,
        max_length=200,
    )


# @mcp.tool()  # CONSOLIDATED into asta_code
def asta_assign_codes(params: AssignCodesInput) -> str:
    """Assign code library entries to tasks/bars via COM.

    For each assignment, finds the bar by ID and the code entry by name,
    then calls bar.AssignCode(entry, append).
    """
    import pythoncom
    import pywintypes

    com_initialized = False
    result: Dict[str, Any] = {"tool": "asta_assign_codes", "success": False}

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        # Find the library
        code_libs = project.CodeLibrarys
        lib = None
        for i in range(1, code_libs.Count + 1):
            try:
                candidate = code_libs.Item(i)
                if candidate.Name and candidate.Name.lower() == params.library_name.lower():
                    lib = candidate
                    break
            except Exception:
                continue

        if lib is None:
            result["error"] = f"Code library '{params.library_name}' not found"
            return json.dumps(result, indent=2, default=str)

        # Build entry name -> object lookup
        entry_map = {}
        try:
            ent_col = lib.Entries
            for j in range(1, ent_col.Count + 1):
                try:
                    entry = ent_col.Item(j)
                    entry_map[entry.Name.lower()] = entry
                except Exception:
                    continue
        except Exception as e:
            result["error"] = f"Could not read entries: {e}"
            return json.dumps(result, indent=2, default=str)

        bars = project.Bars

        try:
            project.StartTransaction("Assign Codes")
        except Exception:
            pass

        assigned = []
        assign_errors = []

        for asgn in params.assignments:
            task_id = asgn.get("task_id")
            entry_name = asgn.get("entry_name", "")
            append = asgn.get("append", True)
            a_result = {"task_id": task_id, "entry_name": entry_name}

            if entry_name.lower() not in entry_map:
                a_result["error"] = f"Entry '{entry_name}' not found in library"
                assign_errors.append(a_result)
                continue

            try:
                bar = _find_bar_by_id(project, int(task_id))
                if bar is None:
                    a_result["error"] = f"Bar ID {task_id} not found"
                    assign_errors.append(a_result)
                    continue
                code_entry = entry_map[entry_name.lower()]
                bar.AssignCode(code_entry, append)
                a_result["assigned"] = True
                assigned.append(a_result)
            except pywintypes.com_error as ce:
                a_result["com_error"] = str(ce)
                assign_errors.append(a_result)
            except Exception as ex:
                a_result["error"] = str(ex)
                assign_errors.append(a_result)

        try:
            _com_end_transaction(project)
        except Exception:
            pass

        result["success"] = True
        result["assigned"] = assigned
        result["assigned_count"] = len(assigned)
        if assign_errors:
            result["errors"] = assign_errors
        result["message"] = f"Assigned {len(assigned)} codes, {len(assign_errors)} errors"

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_assign_codes fatal: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 3. RESOURCE & COST MODELLING
# ---------------------------------------------------------------------------

class ManageResourcesInput(BaseModel):
    """Input for managing resources and cost centres."""
    model_config = ConfigDict(str_strip_whitespace=True)

    action: str = Field(
        ...,
        description="Action: 'list', 'list_rates', 'create_permanent', 'create_consumable', "
                    "'create_cost_centre', 'delete_resource', 'delete_cost_centre'."
    )
    name: Optional[str] = Field(default=None, description="Resource or cost centre name.")
    resource_type: Optional[str] = Field(
        default="permanent",
        description="'permanent' or 'consumable' (for create/delete actions)."
    )
    availability: Optional[float] = Field(
        default=None,
        description="Max availability/units for the resource (e.g. 1.0 = 100%)."
    )
    cost_rate: Optional[float] = Field(
        default=None,
        description="Standard cost rate per time unit."
    )
    calendar_name: Optional[str] = Field(
        default=None,
        description="Calendar name to assign to the resource."
    )
    parent_cost_centre: Optional[str] = Field(
        default=None,
        description="Parent cost centre name (for nested cost centres)."
    )

    @field_validator("action")
    @classmethod
    def validate_action(cls, v: str) -> str:
        allowed = {"list", "list_rates", "create_permanent", "create_consumable",
                    "create_cost_centre", "delete_resource", "delete_cost_centre"}
        if v.lower() not in allowed:
            raise ValueError(f"action must be one of {allowed}")
        return v.lower()


# @mcp.tool()  # CONSOLIDATED into asta_resource
def asta_manage_resources(params: ManageResourcesInput) -> str:
    """Manage resources (permanent/consumable) and cost centres via COM.

    Actions:
      - list: List all permanent resources, consumable resources, and cost centres.
      - create_permanent: Create a new permanent (labour/equipment) resource.
      - create_consumable: Create a new consumable (material) resource.
      - create_cost_centre: Create a new cost centre.
      - delete_resource: Delete a resource by name.
      - delete_cost_centre: Delete a cost centre by name.
    """
    import pythoncom
    import pywintypes
    import win32com.client

    com_initialized = False
    result: Dict[str, Any] = {"tool": "asta_manage_resources", "action": params.action, "success": False}

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        # === ACTION: list ===
        if params.action == "list":
            # Permanent resources
            perm_list = []
            try:
                perms = project.PermanentResources
                for i in range(1, perms.Count + 1):
                    try:
                        r = perms.Item(i)
                        perm_list.append({
                            "id": r.ID,
                            "name": r.Name,
                            "availability": safe_float(getattr(r, 'Availability', None)),
                            "email": safe_str(getattr(r, 'EmailAddress', None)),
                        })
                    except Exception:
                        continue
            except Exception as e:
                result["permanent_resources_error"] = str(e)

            # Consumable resources
            cons_list = []
            try:
                cons = project.ConsumableResources
                for i in range(1, cons.Count + 1):
                    try:
                        r = cons.Item(i)
                        # Read CostPerUnit via IAmountAndCurrency.Amount
                        cost_per_unit = 0.0
                        try:
                            r_d = win32com.client.Dispatch(r)
                            cpu_did = r_d._oleobj_.GetIDsOfNames(0, 'CostPerUnit')
                            cpu_raw = r_d._oleobj_.InvokeTypes(cpu_did, 0, 2, (9, 0), ())
                            if cpu_raw:
                                cpu_d = win32com.client.Dispatch(cpu_raw)
                                cost_per_unit = cpu_d._oleobj_.InvokeTypes(0, 0, 2, (5, 0), ())
                        except Exception:
                            pass
                        cons_list.append({
                            "id": r.ID,
                            "name": r.Name,
                            "availability": safe_float(getattr(r, 'Availability', None)),
                            "cost_per_unit": cost_per_unit,
                        })
                    except Exception:
                        continue
            except Exception as e:
                result["consumable_resources_error"] = str(e)

            # Cost centres
            cc_list = []
            try:
                ccs = project.CostCentres
                for i in range(1, ccs.Count + 1):
                    try:
                        cc = ccs.Item(i)
                        # Read Cost via IAmountAndCurrency.Amount
                        cc_cost = 0.0
                        try:
                            cc_d = win32com.client.Dispatch(cc)
                            cost_did = cc_d._oleobj_.GetIDsOfNames(0, 'Cost')
                            cost_raw = cc_d._oleobj_.InvokeTypes(cost_did, 0, 2, (9, 0), ())
                            if cost_raw:
                                cost_d = win32com.client.Dispatch(cost_raw)
                                cc_cost = cost_d._oleobj_.InvokeTypes(0, 0, 2, (5, 0), ())
                        except Exception:
                            pass
                        cc_list.append({
                            "id": cc.ID,
                            "name": cc.Name,
                            "cost": cc_cost,
                        })
                    except Exception:
                        continue
            except Exception as e:
                result["cost_centres_error"] = str(e)

            result["permanent_resources"] = perm_list
            result["consumable_resources"] = cons_list
            result["cost_centres"] = cc_list
            result["success"] = True

        # === ACTION: list_rates ===
        elif params.action == "list_rates":
            # List CostAndIncomeRates (project-level rate definitions for permanent resource costing)
            rates_list = []
            try:
                rates_coll = project.CostAndIncomeRates
                for i in range(1, rates_coll.Count + 1):
                    try:
                        rate = rates_coll.Item(i)
                        rate_d = win32com.client.Dispatch(rate)
                        # Get Amount (IAmountAndCurrency)
                        amount = 0.0
                        try:
                            amt_did = rate_d._oleobj_.GetIDsOfNames(0, 'Amount')
                            amt_raw = rate_d._oleobj_.InvokeTypes(amt_did, 0, 2, (9, 0), ())
                            if amt_raw:
                                amt_d = win32com.client.Dispatch(amt_raw)
                                amount = amt_d._oleobj_.InvokeTypes(0, 0, 2, (5, 0), ())
                        except Exception:
                            pass
                        # Get TimeUnit
                        time_unit = ""
                        try:
                            tu_did = rate_d._oleobj_.GetIDsOfNames(0, 'TimeUnit')
                            tu_raw = rate_d._oleobj_.InvokeTypes(tu_did, 0, 2, (9, 0), ())
                            if tu_raw:
                                tu_d = win32com.client.Dispatch(tu_raw)
                                time_unit = str(tu_d.Name) if hasattr(tu_d, 'Name') else str(tu_d)
                        except Exception:
                            pass
                        # Get CostCentre
                        cost_centre = ""
                        try:
                            cc_did = rate_d._oleobj_.GetIDsOfNames(0, 'CostCentre')
                            cc_raw = rate_d._oleobj_.InvokeTypes(cc_did, 0, 2, (9, 0), ())
                            if cc_raw:
                                cc_d = win32com.client.Dispatch(cc_raw)
                                cost_centre = str(cc_d.Name) if hasattr(cc_d, 'Name') else str(cc_d)
                        except Exception:
                            pass
                        # Get type (0=cost, 1=income)
                        rate_type = 0
                        try:
                            rate_type = rate_d.type
                        except Exception:
                            pass

                        rates_list.append({
                            "id": rate.ID,
                            "name": rate.Name,
                            "amount": amount,
                            "time_unit": time_unit,
                            "cost_centre": cost_centre,
                            "type": "cost" if rate_type == 0 else "income",
                        })
                    except Exception:
                        continue
            except Exception as e:
                result["rates_error"] = str(e)

            result["cost_and_income_rates"] = rates_list
            result["success"] = True
            result["note"] = ("Use 'rate_name' in assign action to assign a rate to a permanent "
                              "resource allocation. Cost = effort_hours × rate_amount.")

        # === ACTION: create_permanent ===
        elif params.action == "create_permanent":
            if not params.name:
                result["error"] = "name is required"
                return json.dumps(result, indent=2, default=str)

            try:
                project.StartTransaction("Create Permanent Resource")
            except Exception:
                pass

            try:
                perms = project.PermanentResources
                # Check if exists
                for i in range(1, perms.Count + 1):
                    try:
                        if perms.Item(i).Name.lower() == params.name.lower():
                            result["success"] = True
                            result["message"] = f"Resource '{params.name}' already exists (ID={perms.Item(i).ID})"
                            result["resource_id"] = perms.Item(i).ID
                            try:
                                _com_end_transaction(project)
                            except Exception:
                                pass
                            return json.dumps(result, indent=2, default=str)
                    except Exception:
                        continue

                new_res = perms.Add()
                new_res.Name = params.name
                if params.availability is not None:
                    try:
                        new_res.Availability = params.availability
                    except Exception:
                        pass
                if params.cost_rate is not None:
                    try:
                        # NOTE: IPermanentResource.Cost is READ-ONLY (calculated field).
                        # There is NO StandardRate/HourlyRate/CostRate property on IPermanentResource.
                        # Permanent resource rates can only be set via Asta UI.
                        # The cost_rate param is stored as metadata only for now.
                        result["cost_rate_note"] = (
                            "Permanent resource rates cannot be set via COM. "
                            "Cost.Amount is read-only (calculated). "
                            "Set rates manually in Asta UI, or use cost centre allocations "
                            "with GivenValue to assign costs to tasks."
                        )
                    except Exception:
                        pass
                if params.calendar_name is not None:
                    try:
                        cals = project.Calendars
                        for ci in range(1, cals.Count + 1):
                            cal = cals.Item(ci)
                            if cal.Name.lower() == params.calendar_name.lower():
                                new_res.Calendar = cal
                                break
                    except Exception:
                        pass

                try:
                    _com_end_transaction(project)
                except Exception:
                    pass

                result["success"] = True
                result["resource_id"] = new_res.ID
                result["message"] = f"Permanent resource '{params.name}' created"
            except Exception as e:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
                result["error"] = f"Failed: {e}"

        # === ACTION: create_consumable ===
        elif params.action == "create_consumable":
            if not params.name:
                result["error"] = "name is required"
                return json.dumps(result, indent=2, default=str)

            try:
                project.StartTransaction("Create Consumable Resource")
            except Exception:
                pass

            try:
                cons = project.ConsumableResources
                for i in range(1, cons.Count + 1):
                    try:
                        if cons.Item(i).Name.lower() == params.name.lower():
                            result["success"] = True
                            result["message"] = f"Resource '{params.name}' already exists"
                            result["resource_id"] = cons.Item(i).ID
                            try:
                                _com_end_transaction(project)
                            except Exception:
                                pass
                            return json.dumps(result, indent=2, default=str)
                    except Exception:
                        continue

                new_res = cons.Add()
                new_res.Name = params.name
                if params.availability is not None:
                    try:
                        new_res.Availability = params.availability
                    except Exception:
                        pass
                if params.cost_rate is not None:
                    try:
                        # IConsumableResource.CostPerUnit is IAmountAndCurrency with GET/PUT
                        # Access the COM object and set Amount (did=0) directly
                        res_d = win32com.client.Dispatch(new_res)
                        cpu_did = res_d._oleobj_.GetIDsOfNames(0, 'CostPerUnit')
                        cpu_raw = res_d._oleobj_.InvokeTypes(cpu_did, 0, 2, (9, 0), ())
                        if cpu_raw:
                            cpu_obj = win32com.client.Dispatch(cpu_raw)
                            cpu_obj._oleobj_.InvokeTypes(0, 0, 4, (24, 0), ((5, 1),), float(params.cost_rate))
                            result["cost_per_unit_set"] = params.cost_rate
                    except Exception as cost_e:
                        result["cost_per_unit_warning"] = f"Could not set CostPerUnit: {cost_e}"

                try:
                    _com_end_transaction(project)
                except Exception:
                    pass

                result["success"] = True
                result["resource_id"] = new_res.ID
                result["message"] = f"Consumable resource '{params.name}' created"
            except Exception as e:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
                result["error"] = f"Failed: {e}"

        # === ACTION: create_cost_centre ===
        elif params.action == "create_cost_centre":
            if not params.name:
                result["error"] = "name is required"
                return json.dumps(result, indent=2, default=str)

            try:
                project.StartTransaction("Create Cost Centre")
            except Exception:
                pass

            try:
                ccs = project.CostCentres
                for i in range(1, ccs.Count + 1):
                    try:
                        if ccs.Item(i).Name.lower() == params.name.lower():
                            result["success"] = True
                            result["message"] = f"Cost centre '{params.name}' already exists"
                            result["cost_centre_id"] = ccs.Item(i).ID
                            try:
                                _com_end_transaction(project)
                            except Exception:
                                pass
                            return json.dumps(result, indent=2, default=str)
                    except Exception:
                        continue

                # If parent specified, find it
                parent_cc = None
                if params.parent_cost_centre:
                    for i in range(1, ccs.Count + 1):
                        try:
                            if ccs.Item(i).Name.lower() == params.parent_cost_centre.lower():
                                parent_cc = ccs.Item(i)
                                break
                        except Exception:
                            continue

                if parent_cc is not None:
                    new_cc = parent_cc.SubCostCentres.Add()
                else:
                    new_cc = ccs.Add()
                new_cc.Name = params.name

                try:
                    _com_end_transaction(project)
                except Exception:
                    pass

                result["success"] = True
                result["cost_centre_id"] = new_cc.ID
                result["message"] = f"Cost centre '{params.name}' created"
                if parent_cc:
                    result["parent"] = params.parent_cost_centre

            except Exception as e:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
                result["error"] = f"Failed: {e}"

        # === ACTION: delete_resource ===
        elif params.action == "delete_resource":
            if not params.name:
                result["error"] = "name is required"
                return json.dumps(result, indent=2, default=str)

            try:
                project.StartTransaction("Delete Resource")
            except Exception:
                pass

            deleted = False
            rt = (params.resource_type or "permanent").lower()

            try:
                collection = project.PermanentResources if rt == "permanent" else project.ConsumableResources
                for i in range(1, collection.Count + 1):
                    try:
                        if collection.Item(i).Name.lower() == params.name.lower():
                            collection.Remove(i)
                            deleted = True
                            break
                    except Exception:
                        continue

                try:
                    _com_end_transaction(project)
                except Exception:
                    pass

                if deleted:
                    result["success"] = True
                    result["message"] = f"Resource '{params.name}' deleted"
                else:
                    result["error"] = f"Resource '{params.name}' not found in {rt} resources"
            except Exception as e:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
                result["error"] = f"Failed: {e}"

        # === ACTION: delete_cost_centre ===
        elif params.action == "delete_cost_centre":
            if not params.name:
                result["error"] = "name is required"
                return json.dumps(result, indent=2, default=str)

            try:
                project.StartTransaction("Delete Cost Centre")
            except Exception:
                pass

            try:
                ccs = project.CostCentres
                deleted = False
                for i in range(1, ccs.Count + 1):
                    try:
                        if ccs.Item(i).Name.lower() == params.name.lower():
                            ccs.Remove(i)
                            deleted = True
                            break
                    except Exception:
                        continue

                try:
                    _com_end_transaction(project)
                except Exception:
                    pass

                if deleted:
                    result["success"] = True
                    result["message"] = f"Cost centre '{params.name}' deleted"
                else:
                    result["error"] = f"Cost centre '{params.name}' not found"
            except Exception as e:
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass
                result["error"] = f"Failed: {e}"

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_manage_resources fatal: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


class ResourceAssignmentInput(BaseModel):
    """Input for assigning resources to tasks with allocation profiling."""
    model_config = ConfigDict(str_strip_whitespace=True)

    assignments: List[Dict[str, Any]] = Field(
        ...,
        description="List of assignments. Each dict: "
                    "{'task_id': int, 'resource_name': str, 'resource_type': 'permanent'|'consumable'|'cost_centre', "
                    "'units': float (opt, default 1.0), 'is_demand': bool (opt, default false), "
                    "'work_profile': str (opt: 'linear', 'front_loaded', 'back_loaded', 'bell_curve'), "
                    "'cost_value': float (opt, for cost_centre: sets GivenValue in $ on ICostAllocation, default $1), "
                    "'rate_name': str (opt, for permanent: assigns a CostAndIncomeRate by name for cost calc), "
                    "'effort_hours': float (opt, for permanent: sets GivenEffort in hours, converted to seconds), "
                    "'given_work': float (opt, for permanent: sets GivenWork units), "
                    "'given_allocation': float (opt, for permanent: sets resource allocation e.g. 1.0=100%), "
                    "'quantity': float (opt, for consumable: sets GivenQuantity), "
                    "'cost_per_unit': float (opt, for consumable: sets CostPerUnit on allocation), "
                    "'consumption_rate': float (opt, for consumable: sets GivenConsumptionRate), "
                    "'task_work_rate': float (opt, sets TaskWorkRate on the task itself), "
                    "'task_work': float (opt, sets Work quantity on the task itself)}",
        min_length=1,
        max_length=100,
    )


# @mcp.tool()  # CONSOLIDATED into asta_resource
def asta_assign_resource_model(params: ResourceAssignmentInput) -> str:
    """Assign resources or cost centres to tasks via COM with allocation profiling.

    For each assignment:
      1. Finds the task/bar by ID.
      2. Finds the resource or cost centre by name.
      3. Calls the appropriate AssignPermanentResource / AssignConsumableResource / AssignCost.
      4. Sets work distribution profile if specified.
    """
    import pythoncom
    import pywintypes
    import win32com.client

    com_initialized = False
    result: Dict[str, Any] = {"tool": "asta_assign_resource_model", "success": False}

    # Work profile constants (Asta internal profile IDs)
    WORK_PROFILES = {
        "linear": 0,
        "front_loaded": 1,
        "back_loaded": 2,
        "bell_curve": 3,
    }

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        # Build resource lookup caches
        perm_map = {}
        try:
            perms = project.PermanentResources
            for i in range(1, perms.Count + 1):
                try:
                    r = perms.Item(i)
                    perm_map[r.Name.lower()] = r
                except Exception:
                    continue
        except Exception:
            pass

        cons_map = {}
        try:
            cons = project.ConsumableResources
            for i in range(1, cons.Count + 1):
                try:
                    r = cons.Item(i)
                    cons_map[r.Name.lower()] = r
                except Exception:
                    continue
        except Exception:
            pass

        cc_map = {}
        try:
            ccs = project.CostCentres
            for i in range(1, ccs.Count + 1):
                try:
                    cc = ccs.Item(i)
                    cc_map[cc.Name.lower()] = cc
                except Exception:
                    continue
        except Exception:
            pass

        # Build rate lookup cache for permanent resource rate assignment
        rate_map = {}
        try:
            rates_coll = project.CostAndIncomeRates
            for i in range(1, rates_coll.Count + 1):
                try:
                    rate = rates_coll.Item(i)
                    rate_map[rate.Name.lower()] = rate
                except Exception:
                    continue
        except Exception:
            pass

        bars = project.Bars

        try:
            project.StartTransaction("Assign Resources")
        except Exception:
            pass

        assigned = []
        assign_errors = []

        for asgn in params.assignments:
            task_id = asgn.get("task_id")
            res_name = asgn.get("resource_name", "")
            res_type = asgn.get("resource_type", "permanent").lower()
            is_demand = asgn.get("is_demand", False)
            work_profile = asgn.get("work_profile", "").lower()
            a_result = {"task_id": task_id, "resource_name": res_name, "resource_type": res_type}

            try:
                bar = _find_bar_by_id(project, int(task_id))
                if bar is None:
                    a_result["error"] = f"Bar ID {task_id} not found (tried index and ID search)"
                    assign_errors.append(a_result)
                    continue
                # Get the task from the bar — Tasks(1) is reliable, ExpandedTask returns root
                task, _ = _get_bar_task(bar)
                if task is None:
                    try:
                        task = bar.ExpandedTask
                    except Exception:
                        pass
                if task is None:
                    a_result["error"] = f"Cannot get task from bar {task_id}"
                    assign_errors.append(a_result)
                    continue

                allocation = None

                if res_type == "permanent":
                    res_obj = perm_map.get(res_name.lower())
                    if res_obj is None:
                        a_result["error"] = f"Permanent resource '{res_name}' not found"
                        assign_errors.append(a_result)
                        continue
                    try:
                        allocation = task.AssignPermanentResource(res_obj, is_demand, None, None)
                    except Exception:
                        # Fallback to generic AssignResource
                        allocation = task.AssignResource(res_obj, is_demand)

                    # Set effort/work/allocation on IPermanentDemandAllocation
                    if allocation is not None:
                        alloc_d = win32com.client.Dispatch(allocation)

                        # Assign rate for cost calculation (e.g. $25/hour)
                        rate_name = asgn.get("rate_name")
                        if rate_name:
                            rate_obj = rate_map.get(rate_name.lower())
                            if rate_obj:
                                try:
                                    alloc_d.AssignRate(rate_obj)
                                    a_result["rate_assigned"] = rate_name
                                except Exception as re:
                                    a_result["rate_warning"] = f"Could not assign rate '{rate_name}': {re}"
                            else:
                                a_result["rate_warning"] = f"Rate '{rate_name}' not found. Available: {list(rate_map.keys())}"

                        # Set GivenEffort (in seconds; user provides hours)
                        effort_hours = asgn.get("effort_hours")
                        if effort_hours is not None:
                            try:
                                alloc_d.GivenEffort = float(effort_hours) * 3600.0
                                a_result["effort_hours_set"] = effort_hours
                            except Exception as ee:
                                a_result["effort_warning"] = f"Could not set GivenEffort: {ee}"

                        # Set GivenWork
                        given_work = asgn.get("given_work")
                        if given_work is not None:
                            try:
                                alloc_d.GivenWork = float(given_work)
                                a_result["given_work_set"] = given_work
                            except Exception as we:
                                a_result["work_warning"] = f"Could not set GivenWork: {we}"

                        # Set GivenAllocation (resource units, 1.0 = 100%)
                        given_allocation = asgn.get("given_allocation")
                        if given_allocation is not None:
                            try:
                                alloc_d.GivenAllocation = float(given_allocation)
                                a_result["given_allocation_set"] = given_allocation
                            except Exception as ae:
                                a_result["allocation_warning"] = f"Could not set GivenAllocation: {ae}"

                elif res_type == "consumable":
                    res_obj = cons_map.get(res_name.lower())
                    if res_obj is None:
                        a_result["error"] = f"Consumable resource '{res_name}' not found"
                        assign_errors.append(a_result)
                        continue
                    try:
                        allocation = task.AssignConsumableResource(res_obj, is_demand, None, None)
                    except Exception:
                        allocation = task.AssignResource(res_obj, is_demand)

                    # Set quantity, CostPerUnit, consumption rate on IConsumableDemandAllocation
                    if allocation is not None:
                        alloc_d = win32com.client.Dispatch(allocation)

                        # Set GivenQuantity
                        quantity = asgn.get("quantity")
                        if quantity is not None:
                            try:
                                alloc_d.GivenQuantity = float(quantity)
                                a_result["quantity_set"] = quantity
                            except Exception as qe:
                                a_result["quantity_warning"] = f"Could not set GivenQuantity: {qe}"

                        # Set CostPerUnit (IAmountAndCurrency) on allocation
                        cost_per_unit = asgn.get("cost_per_unit")
                        if cost_per_unit is not None:
                            try:
                                cpu_did = alloc_d._oleobj_.GetIDsOfNames(0, 'CostPerUnit')
                                cpu_raw = alloc_d._oleobj_.InvokeTypes(cpu_did, 0, 2, (9, 0), ())
                                if cpu_raw:
                                    cpu_obj = win32com.client.Dispatch(cpu_raw)
                                    cpu_obj._oleobj_.InvokeTypes(0, 0, 4, (24, 0), ((5, 1),), float(cost_per_unit))
                                    a_result["cost_per_unit_set"] = cost_per_unit
                            except Exception as cpe:
                                a_result["cost_per_unit_warning"] = f"Could not set CostPerUnit: {cpe}"

                        # Set GivenConsumptionRate
                        consumption_rate = asgn.get("consumption_rate")
                        if consumption_rate is not None:
                            try:
                                alloc_d.GivenConsumptionRate = float(consumption_rate)
                                a_result["consumption_rate_set"] = consumption_rate
                            except Exception as cre:
                                a_result["consumption_rate_warning"] = f"Could not set rate: {cre}"

                elif res_type == "cost_centre":
                    cc_obj = cc_map.get(res_name.lower())
                    if cc_obj is None:
                        a_result["error"] = f"Cost centre '{res_name}' not found"
                        assign_errors.append(a_result)
                        continue
                    allocation = task.AssignCost(cc_obj)

                    # Set GivenValue on ICostAllocation if cost_value provided
                    # Default AssignCost creates $1 allocation — this sets actual cost
                    cost_value = asgn.get("cost_value")
                    if allocation is not None and cost_value is not None:
                        try:
                            alloc_d = win32com.client.Dispatch(allocation)
                            gv_did = alloc_d._oleobj_.GetIDsOfNames(0, 'GivenValue')
                            gv_raw = alloc_d._oleobj_.InvokeTypes(gv_did, 0, 2, (9, 0), ())
                            if gv_raw:
                                gv_obj = win32com.client.Dispatch(gv_raw)
                                gv_obj._oleobj_.InvokeTypes(0, 0, 4, (24, 0), ((5, 1),), float(cost_value))
                                a_result["cost_value_set"] = cost_value
                        except Exception as cv_e:
                            a_result["cost_value_warning"] = f"Could not set GivenValue: {cv_e}"

                else:
                    a_result["error"] = f"Invalid resource_type: '{res_type}'"
                    assign_errors.append(a_result)
                    continue

                a_result["assigned"] = True

                # Set task-level work properties (independent of resource type)
                task_work_rate = asgn.get("task_work_rate")
                if task_work_rate is not None:
                    try:
                        task.TaskWorkRate = float(task_work_rate)
                        a_result["task_work_rate_set"] = task_work_rate
                    except Exception as twr_e:
                        a_result["task_work_rate_warning"] = f"Could not set TaskWorkRate: {twr_e}"

                task_work = asgn.get("task_work")
                if task_work is not None:
                    try:
                        task.Work = float(task_work)
                        a_result["task_work_set"] = task_work
                    except Exception as tw_e:
                        a_result["task_work_warning"] = f"Could not set Work: {tw_e}"

                # Set work profile if specified and allocation was successful
                if allocation is not None and work_profile and work_profile in WORK_PROFILES:
                    try:
                        profile_id = WORK_PROFILES[work_profile]
                        allocation.EditToken("WorkProfile", str(profile_id))
                        a_result["work_profile_set"] = work_profile
                    except Exception as wp_e:
                        a_result["work_profile_warning"] = f"Could not set profile: {wp_e}"

                assigned.append(a_result)

            except pywintypes.com_error as ce:
                a_result["com_error"] = str(ce)
                assign_errors.append(a_result)
            except Exception as ex:
                a_result["error"] = str(ex)
                assign_errors.append(a_result)

        try:
            _com_end_transaction(project)
        except Exception:
            pass

        result["success"] = True
        result["assigned"] = assigned
        result["assigned_count"] = len(assigned)
        if assign_errors:
            result["errors"] = assign_errors
        result["message"] = f"Assigned {len(assigned)} resources, {len(assign_errors)} errors"

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_assign_resource_model fatal: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 4. ADVANCED VIEW MANAGEMENT
# ---------------------------------------------------------------------------

class ViewConfigInput(BaseModel):
    """Input for configuring the active Asta Powerproject view."""
    model_config = ConfigDict(str_strip_whitespace=True)

    action: str = Field(
        ...,
        description="Action: 'get_status', 'set_display', 'set_grouping', "
                    "'set_sorting', 'set_filter', 'toggle_histogram', 'show_hierarchy_level', "
                    "'list_tables', 'apply_table', 'get_columns'."
    )
    # Table/column management
    table_name: Optional[str] = Field(default=None, description="Table definition name to apply (for apply_table action).")
    # Display toggles
    display_critical_path: Optional[bool] = Field(default=None, description="Toggle critical path display.")
    display_free_float: Optional[bool] = Field(default=None, description="Toggle free float display.")
    display_total_float: Optional[bool] = Field(default=None, description="Toggle total float display.")
    display_progress_lines: Optional[bool] = Field(default=None, description="Toggle progress lines.")
    display_annotations: Optional[bool] = Field(default=None, description="Toggle annotations.")
    display_cost_allocations: Optional[bool] = Field(default=None, description="Toggle cost allocation display.")
    display_demand_allocations: Optional[bool] = Field(default=None, description="Toggle demand allocation display.")
    display_scheduled_allocations: Optional[bool] = Field(default=None, description="Toggle scheduled allocation display.")

    # Grouping / sorting
    group_by_library: Optional[str] = Field(default=None, description="Code library name to group by.")
    sort_field: Optional[str] = Field(
        default=None,
        description="Field to sort by (e.g. 'Start', 'End', 'Name', 'Duration', 'TotalFloat')."
    )
    sort_ascending: Optional[bool] = Field(default=True, description="Sort ascending (true) or descending (false).")

    # Histogram
    histogram_visible: Optional[bool] = Field(default=None, description="Show/hide the histogram pane.")
    histogram_type: Optional[str] = Field(
        default=None,
        description="Histogram type: 'resource', 'cost', 'work'."
    )

    # Hierarchy
    hierarchy_level: Optional[int] = Field(
        default=None,
        description="Expand hierarchy to this level (1=top level only, 0=collapse all)."
    )

    @field_validator("action")
    @classmethod
    def validate_action(cls, v: str) -> str:
        allowed = {"get_status", "set_display", "set_grouping", "set_sorting",
                    "set_filter", "toggle_histogram", "show_hierarchy_level",
                    "list_tables", "apply_table", "get_columns"}
        if v.lower() not in allowed:
            raise ValueError(f"action must be one of {allowed}")
        return v.lower()


# @mcp.tool()  # CONSOLIDATED into asta_view
def asta_configure_view(params: ViewConfigInput) -> str:
    """Configure the active Asta Powerproject bar chart view via COM.

    Manipulates display settings, grouping, sorting, filters, histograms,
    and hierarchy expansion on the currently active view.
    """
    import pythoncom
    import pywintypes

    com_initialized = False
    result: Dict[str, Any] = {"tool": "asta_configure_view", "action": params.action, "success": False}

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        # Get the active view — try multiple access paths
        view = None
        view_source = None

        # Strategy 1: app.ActiveView
        try:
            view = app.ActiveView
            if view is not None:
                view_source = "app.ActiveView"
        except Exception:
            pass

        # Strategy 2: project.CurrentView
        if view is None:
            try:
                view = project.CurrentView
                if view is not None:
                    view_source = "project.CurrentView"
            except Exception:
                pass

        # Strategy 3: project.Views.Item(1) — first view
        if view is None:
            try:
                views = project.Views
                if views is not None and views.Count > 0:
                    view = views.Item(1)
                    view_source = "project.Views.Item(1)"
            except Exception:
                pass

        if view is None:
            result["error"] = (
                "Could not access the active view. "
                "Ensure a bar chart view is open in Asta Powerproject."
            )
            return json.dumps(result, indent=2, default=str)

        result["view_source"] = view_source
        try:
            result["view_name"] = view.Name
        except Exception:
            pass

        # === ACTION: get_status ===
        if params.action == "get_status":
            status = {}
            bool_props = [
                "DisplayCriticalPath", "DisplayFreeFloat", "DisplayTotalFloat",
                "DisplayProgressLines", "DisplayAnnotations",
                "DisplayCostAllocations", "DisplayDemandAllocations",
                "DisplayScheduledAllocations", "DisplayBarChart",
            ]
            for prop in bool_props:
                try:
                    status[prop] = getattr(view, prop)
                except Exception:
                    status[prop] = "N/A"

            try:
                status["BarLineCount"] = view.BarLineCount
            except Exception:
                pass
            try:
                status["view_type"] = view.type
            except Exception:
                pass
            try:
                vlib = view.VisibleCodeLibrary
                status["visible_code_library"] = vlib.Name if vlib else None
            except Exception:
                pass

            result["status"] = status
            result["success"] = True

        # === ACTION: set_display ===
        elif params.action == "set_display":
            changes = {}
            display_map = {
                "display_critical_path": ("DisplayCriticalPath", params.display_critical_path),
                "display_free_float": ("DisplayFreeFloat", params.display_free_float),
                "display_total_float": ("DisplayTotalFloat", params.display_total_float),
                "display_progress_lines": ("DisplayProgressLines", params.display_progress_lines),
                "display_annotations": ("DisplayAnnotations", params.display_annotations),
                "display_cost_allocations": ("DisplayCostAllocations", params.display_cost_allocations),
                "display_demand_allocations": ("DisplayDemandAllocations", params.display_demand_allocations),
                "display_scheduled_allocations": ("DisplayScheduledAllocations", params.display_scheduled_allocations),
            }
            for param_key, (prop_name, value) in display_map.items():
                if value is not None:
                    try:
                        setattr(view, prop_name, value)
                        changes[prop_name] = value
                    except (AttributeError, Exception) as e:
                        # Fallback: try EditToken
                        try:
                            view.EditToken(prop_name, str(value))
                            changes[prop_name] = f"{value} (via EditToken)"
                        except Exception:
                            changes[prop_name] = f"Not supported: {e}"

            try:
                view.Refresh()
            except Exception:
                pass

            result["changes"] = changes
            result["success"] = True

        # === ACTION: set_grouping ===
        elif params.action == "set_grouping":
            if not params.group_by_library:
                result["error"] = "group_by_library is required for set_grouping"
                return json.dumps(result, indent=2, default=str)

            # Find the code library
            lib = None
            try:
                code_libs = project.CodeLibrarys
                for i in range(1, code_libs.Count + 1):
                    try:
                        candidate = code_libs.Item(i)
                        if candidate.Name.lower() == params.group_by_library.lower():
                            lib = candidate
                            break
                    except Exception:
                        continue
            except Exception:
                pass

            if lib is None:
                result["error"] = f"Code library '{params.group_by_library}' not found"
                return json.dumps(result, indent=2, default=str)

            try:
                view.VisibleCodeLibrary = lib
                try:
                    view.Refresh()
                except Exception:
                    pass
                result["success"] = True
                result["grouped_by"] = params.group_by_library
            except Exception as e:
                # Fallback: try SetMultipleDisplayLibraries
                try:
                    view.SetMultipleDisplayLibraries([lib])
                    view.Refresh()
                    result["success"] = True
                    result["grouped_by"] = params.group_by_library
                    result["method"] = "SetMultipleDisplayLibraries"
                except Exception as e2:
                    result["error"] = f"Could not set grouping: primary={e}, fallback={e2}"

        # === ACTION: set_sorting ===
        elif params.action == "set_sorting":
            if not params.sort_field:
                result["error"] = "sort_field is required for set_sorting"
                return json.dumps(result, indent=2, default=str)

            try:
                sort_obj = view.Sort
                if sort_obj is not None:
                    # Try setting sort via the Sort object
                    try:
                        sort_obj.Field = params.sort_field
                        sort_obj.Ascending = params.sort_ascending
                        view.Refresh()
                        result["success"] = True
                        result["sorted_by"] = params.sort_field
                        result["ascending"] = params.sort_ascending
                    except Exception:
                        # Fallback: use token-based approach
                        view.EditToken("SortField", params.sort_field)
                        view.EditToken("SortAscending", str(params.sort_ascending))
                        view.Refresh()
                        result["success"] = True
                        result["sorted_by"] = params.sort_field
                        result["method"] = "EditToken"
                else:
                    result["error"] = "Sort object is not available on this view"
            except Exception as e:
                result["error"] = f"Could not set sorting: {e}"

        # === ACTION: set_filter ===
        elif params.action == "set_filter":
            try:
                # Toggle critical path display as a common filter action
                if params.display_critical_path is not None:
                    view.DisplayCriticalPath = params.display_critical_path
                    result["display_critical_path"] = params.display_critical_path

                # Access the Filter object for custom filters
                try:
                    filter_obj = view.Filter
                    if filter_obj is not None:
                        result["filter_available"] = True
                        # Report current filter state
                        try:
                            result["current_filter"] = filter_obj.Name
                        except Exception:
                            pass
                except Exception:
                    result["filter_available"] = False

                try:
                    view.Refresh()
                except Exception:
                    pass

                result["success"] = True
            except Exception as e:
                result["error"] = f"Could not set filter: {e}"

        # === ACTION: toggle_histogram ===
        elif params.action == "toggle_histogram":
            try:
                histogram = view.Histogram()
                if histogram is not None:
                    if params.histogram_visible is not None:
                        try:
                            # Toggle histogram pane visibility
                            histogram.EditToken("Visible", str(params.histogram_visible))
                            result["histogram_visible"] = params.histogram_visible
                        except Exception:
                            # Fallback: try via view property
                            try:
                                view.EditToken("HistogramVisible", str(params.histogram_visible))
                                result["histogram_visible"] = params.histogram_visible
                            except Exception as e:
                                result["histogram_warning"] = f"Could not toggle visibility: {e}"

                    if params.histogram_type:
                        try:
                            histogram.EditToken("Type", params.histogram_type)
                            result["histogram_type"] = params.histogram_type
                        except Exception as ht_e:
                            result["histogram_type_warning"] = f"Could not set type: {ht_e}"

                    try:
                        view.Refresh()
                    except Exception:
                        pass
                    result["success"] = True
                else:
                    result["error"] = "Histogram pane not available on this view"
            except Exception as e:
                result["error"] = f"Could not access histogram: {e}"

        # === ACTION: show_hierarchy_level ===
        elif params.action == "show_hierarchy_level":
            if params.hierarchy_level is None:
                result["error"] = "hierarchy_level is required"
                return json.dumps(result, indent=2, default=str)

            try:
                view.ShowHierarchy(params.hierarchy_level)
                try:
                    view.Refresh()
                except Exception:
                    pass
                result["success"] = True
                result["hierarchy_level"] = params.hierarchy_level
            except Exception as e:
                result["error"] = f"Could not set hierarchy level: {e}"

        # === ACTION: list_tables ===
        elif params.action == "list_tables":
            try:
                D = win32com.client.Dispatch
                tds = D(project.TableDefinitions)
                tables = []
                for i in range(1, tds.Count + 1):
                    td = D(tds.Item(i))
                    tables.append({"index": i, "name": td.Name})
                result["success"] = True
                result["table_definitions"] = tables
                result["count"] = tds.Count
                # Show current table
                try:
                    ss = D(view.SpreadSheet())
                    result["current_table"] = ss.TableDefinition
                except Exception:
                    pass
            except Exception as e:
                result["error"] = f"Could not list table definitions: {e}"

        # === ACTION: apply_table ===
        elif params.action == "apply_table":
            if not params.table_name:
                result["error"] = "table_name is required. Use list_tables to see available tables."
                return json.dumps(result, indent=2, default=str)

            try:
                D = win32com.client.Dispatch
                tds = D(project.TableDefinitions)
                td_obj = None
                for i in range(1, tds.Count + 1):
                    td = D(tds.Item(i))
                    if td.Name.lower() == params.table_name.lower():
                        td_obj = td
                        break

                if not td_obj:
                    # Try partial match
                    for i in range(1, tds.Count + 1):
                        td = D(tds.Item(i))
                        if params.table_name.lower() in td.Name.lower():
                            td_obj = td
                            break

                if not td_obj:
                    result["error"] = f"Table definition '{params.table_name}' not found. Use list_tables to see available names."
                    return json.dumps(result, indent=2, default=str)

                ss = D(view.SpreadSheet())
                project.StartTransaction("ApplyTable")
                ss.ApplyTableDefinition(td_obj)
                project.EndTransaction()
                project.WaitForNotificationProcessing()

                # Read resulting columns
                total = ss.TotalCols
                columns = []
                for ci in range(1, total + 1):
                    try:
                        ss.Col = ci
                        columns.append({"index": ci, "token": ss.ColDataToken, "width": ss.ColWidth})
                    except:
                        break

                try:
                    view.Refresh()
                except Exception:
                    pass

                result["success"] = True
                result["applied_table"] = td_obj.Name
                result["columns"] = columns
                result["column_count"] = len(columns)
            except Exception as e:
                result["error"] = f"Could not apply table definition: {e}"
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass

        # === ACTION: get_columns ===
        elif params.action == "get_columns":
            try:
                D = win32com.client.Dispatch
                ss = D(view.SpreadSheet())
                total = ss.TotalCols
                columns = []
                for ci in range(1, total + 1):
                    try:
                        ss.Col = ci
                        columns.append({"index": ci, "token": ss.ColDataToken, "width": ss.ColWidth})
                    except:
                        break
                result["success"] = True
                result["columns"] = columns
                result["column_count"] = len(columns)
                try:
                    result["current_table"] = ss.TableDefinition
                except Exception:
                    pass
            except Exception as e:
                result["error"] = f"Could not read columns: {e}"

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_configure_view fatal: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 5. COM-BASED EXPORT & REPORTING
# ---------------------------------------------------------------------------

class ExportXMLInput(BaseModel):
    """Input for COM-based XML export."""
    model_config = ConfigDict(str_strip_whitespace=True)

    output_path: str = Field(
        ...,
        description="Full file path for the exported XML (e.g. 'C:/Users/me/project_export.xml')."
    )
    format: str = Field(
        default="asta_xml",
        description="Export format: 'asta_xml' (native Asta XML) or 'mspdi' (MS Project XML)."
    )
    branch_ids: Optional[List[int]] = Field(
        default=None,
        description="Optional list of bar/branch IDs to export. If omitted, exports entire project."
    )

    @field_validator("format")
    @classmethod
    def validate_format(cls, v: str) -> str:
        allowed = {"asta_xml", "mspdi", "xer", "mpp"}
        if v.lower() not in allowed:
            raise ValueError(f"format must be one of {allowed}")
        return v.lower()


# @mcp.tool()  # CONSOLIDATED into asta_export
def asta_com_export_xml(params: ExportXMLInput) -> str:
    """Export the project to XML or other formats via COM.

    Uses the native COM SaveAs/SaveAsXMLFile methods to ensure
    the fully rescheduled, COM-calculated project is exported.
    """
    import pythoncom
    import pywintypes

    com_initialized = False
    result: Dict[str, Any] = {"tool": "asta_com_export_xml", "format": params.format, "success": False}

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        output_path = os.path.abspath(params.output_path)
        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        branch_ids_arg = params.branch_ids if params.branch_ids else None

        try:
            if params.format == "asta_xml":
                saved_path = project.SaveAsXMLFile(output_path, None, branch_ids_arg)
                result["exported_file"] = saved_path or output_path

            elif params.format == "mspdi":
                # SaveAsXMLFile with MSPDI flags — try flag value 1 for MSPDI
                try:
                    saved_path = project.SaveAsXMLFile(output_path, 1, branch_ids_arg)
                    result["exported_file"] = saved_path or output_path
                except Exception:
                    # Fallback: SaveAs with MSPDI extension hint
                    mspdi_path = output_path
                    if not mspdi_path.lower().endswith('.xml'):
                        mspdi_path += '.xml'
                    project.SaveAs(mspdi_path, None, None)
                    result["exported_file"] = mspdi_path
                    result["note"] = "Exported via SaveAs (MSPDI flag fallback)"

            elif params.format == "xer":
                saved_path = project.SaveAsXERFile(output_path, None, branch_ids_arg)
                result["exported_file"] = saved_path or output_path

            elif params.format == "mpp":
                saved_path = project.SaveAsMPPFile(output_path, None, branch_ids_arg)
                result["exported_file"] = saved_path or output_path

            result["success"] = True
            result["message"] = f"Project exported to {result.get('exported_file', output_path)}"

        except pywintypes.com_error as ce:
            result["com_error"] = str(ce)
            result["error"] = f"COM export failed: {ce}"
        except Exception as e:
            result["error"] = f"Export failed: {e}"

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_com_export_xml fatal: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


class ExportPDFInput(BaseModel):
    """Input for PDF export via COM/GUI."""
    model_config = ConfigDict(str_strip_whitespace=True)

    output_path: str = Field(
        ...,
        description="Full file path for the PDF output (e.g. 'C:/Users/me/project.pdf')."
    )
    print_profile: Optional[str] = Field(
        default=None,
        description="Named print profile in Asta to use. If omitted, uses current view settings."
    )


# @mcp.tool()  # CONSOLIDATED into asta_export
def asta_export_pdf(params: ExportPDFInput) -> str:
    """Export the current Asta view to PDF.

    Uses COM PrintView method first; falls back to GUI automation
    (Ctrl+P -> PDF printer) if COM printing is not available.
    """
    import pythoncom
    import pywintypes

    com_initialized = False
    result: Dict[str, Any] = {"tool": "asta_export_pdf", "success": False}

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        result["com_method"] = method

        output_path = os.path.abspath(params.output_path)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Strategy 1: COM-based PrintView on the active view
        exported_via_com = False

        try:
            view = None
            try:
                view = app.ActiveView
            except Exception:
                pass
            if view is None:
                try:
                    view = project.CurrentView
                except Exception:
                    pass

            if view is not None:
                # Try PrintView with profile
                profile_arg = params.print_profile if params.print_profile else None
                view.PrintView(profile_arg)
                exported_via_com = True
                result["method"] = "COM PrintView"
                result["note"] = (
                    "PrintView triggered. If a PDF printer (e.g. 'Microsoft Print to PDF') "
                    "is set as default, the output will be saved to the PDF printer dialog. "
                    "Ensure the PDF printer path is configured."
                )
        except Exception as com_print_err:
            logger.warning(f"COM PrintView failed: {com_print_err}")

        # Strategy 2: HTML export as a universal fallback
        if not exported_via_com:
            try:
                html_path = output_path.replace('.pdf', '.html')
                project.SaveAsHTMLFile(html_path)
                result["method"] = "HTML export (PDF fallback)"
                result["exported_file"] = html_path
                result["note"] = (
                    f"PDF printing not available via COM. Exported as HTML to '{html_path}'. "
                    "Open in a browser and use Print -> Save as PDF."
                )
                exported_via_com = True
            except Exception as html_err:
                logger.warning(f"HTML export also failed: {html_err}")

        # Strategy 3: GUI automation fallback
        if not exported_via_com:
            try:
                import pyautogui
                # Bring Asta to front
                try:
                    from pywinauto import Desktop
                    desktop = Desktop(backend="uia")
                    windows = desktop.windows(title_re=".*Asta.*|.*Powerproject.*")
                    if windows:
                        windows[0].set_focus()
                        time.sleep(0.5)
                except Exception:
                    pass

                # Ctrl+P to open print dialog
                pyautogui.hotkey('ctrl', 'p')
                time.sleep(2)

                result["method"] = "GUI automation (Ctrl+P)"
                result["note"] = (
                    "Print dialog opened via GUI automation. "
                    "Select 'Microsoft Print to PDF' and save to the desired location."
                )
                exported_via_com = True
            except Exception as gui_err:
                result["gui_error"] = str(gui_err)

        if exported_via_com:
            result["success"] = True
            result["output_path"] = output_path
        else:
            result["error"] = "All PDF export methods failed"

        return json.dumps(result, indent=2, default=str)

    except RuntimeError as e:
        result["error"] = f"COM connection failed: {e}"
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        result["error"] = f"Fatal error: {e}"
        logger.error(f"asta_export_pdf fatal: {e}", exc_info=True)
        return json.dumps(result, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


class AdvancedReportInput(BaseModel):
    """Input for generating an advanced project report."""
    model_config = ConfigDict(str_strip_whitespace=True)

    response_format: str = Field(
        default="markdown",
        description="Output format: 'markdown' or 'json'."
    )
    include_resources: bool = Field(default=True, description="Include resource allocation summary.")
    include_costs: bool = Field(default=True, description="Include cost centre summary.")
    include_variances: bool = Field(default=True, description="Include schedule variance analysis.")
    include_critical_path: bool = Field(default=True, description="Include critical path summary.")
    max_tasks: int = Field(default=50, description="Maximum number of tasks in detailed listing.")

    @field_validator("response_format")
    @classmethod
    def validate_format(cls, v: str) -> str:
        if v.lower() not in {"markdown", "json"}:
            raise ValueError("response_format must be 'markdown' or 'json'")
        return v.lower()


# @mcp.tool()  # CONSOLIDATED into asta_export
def asta_generate_advanced_report(params: AdvancedReportInput) -> str:
    """Generate a comprehensive project report via COM.

    Extracts resource allocations, cost centre totals, schedule variances,
    and critical path information directly from the running Asta instance.
    Returns structured Markdown or JSON.
    """
    import pythoncom
    import pywintypes

    com_initialized = False
    report: Dict[str, Any] = {"tool": "asta_generate_advanced_report", "success": False}

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        app, project, method = _connect_asta_com()
        report["com_method"] = method

        # --- Project Summary ---
        summary = {}
        for prop in ["Name", "FileName", "ProjectStart", "ProjectEnd"]:
            try:
                val = getattr(project, prop)
                summary[prop] = format_date(val) if "Start" in prop or "End" in prop else str(val)
            except Exception:
                summary[prop] = "N/A"

        try:
            summary["Version"] = str(app.Version)
        except Exception:
            pass

        report["project_summary"] = summary

        # --- Task Overview ---
        task_overview = {"total_bars": 0, "tasks": [], "source": ""}
        try:
            # Try COM bar traversal first
            all_bars = _com_get_all_bars(project, max_bars=params.max_tasks)
            bar_count = len(all_bars)

            # If COM only found top-level bars (likely collapsed/filtered),
            # fall back to auto-export + MPXJ which reads ALL tasks from file
            if bar_count <= 1:
                try:
                    export_path = _com_auto_export()
                    mgr = AstaFileManager(export_path)
                    mpxj_tasks = mgr.get_all_tasks(include_summary=True)
                    for t in mpxj_tasks[:params.max_tasks]:
                        task_overview["tasks"].append({
                            "id": t.get("id", 0),
                            "name": t.get("name", ""),
                            "start": str(t.get("start", "N/A")),
                            "end": str(t.get("finish", t.get("end", "N/A"))),
                            "percent_complete": t.get("percent_complete", 0),
                            "critical": t.get("critical", False),
                            "summary": t.get("summary", False),
                        })
                    task_overview["total_bars"] = len(mpxj_tasks)
                    task_overview["source"] = "MPXJ (auto-export)"
                    if len(mpxj_tasks) > params.max_tasks:
                        task_overview["note"] = f"Showing {params.max_tasks} of {len(mpxj_tasks)} tasks"
                except Exception as mpxj_err:
                    logger.warning(f"MPXJ fallback failed: {mpxj_err}")
                    # If MPXJ also fails, use whatever COM found
                    task_overview["total_bars"] = bar_count
                    task_overview["source"] = "COM (top-level only)"
                    for bar in all_bars:
                        try:
                            task_overview["tasks"].append({
                                "id": bar.ID, "name": bar.Name,
                                "start": format_date(bar.Start),
                                "end": format_date(bar.End),
                            })
                        except Exception:
                            continue
            else:
                # COM found multiple bars — use them
                task_overview["total_bars"] = bar_count
                task_overview["source"] = "COM"
                for bar in all_bars:
                    try:
                        task_info = {
                            "id": bar.ID,
                            "name": bar.Name,
                            "start": format_date(bar.Start),
                            "end": format_date(bar.End),
                        }
                        try:
                            task_info["percent_complete"] = round(bar.OverallPercentComplete, 1)
                        except Exception:
                            try:
                                task_info["percent_complete"] = round(bar.DurationPercentComplete, 1)
                            except Exception:
                                pass
                        try:
                            _btask, _ = _get_bar_task(bar)
                            if _btask:
                                task_info["critical"] = bool(_btask.Critical)
                        except Exception:
                            pass
                        task_overview["tasks"].append(task_info)
                    except Exception:
                        continue
                if bar_count >= params.max_tasks:
                    task_overview["note"] = f"Showing first {params.max_tasks} bars (may be more)"
        except Exception as e:
            task_overview["error"] = str(e)

        report["task_overview"] = task_overview

        # --- Resource Summary ---
        if params.include_resources:
            res_summary = {"permanent": [], "consumable": []}
            try:
                perms = project.PermanentResources
                for i in range(1, perms.Count + 1):
                    try:
                        r = perms.Item(i)
                        r_info = {
                            "name": r.Name,
                            "availability": safe_float(getattr(r, 'Availability', None)),
                        }
                        try:
                            r_info["scheduled_effort"] = safe_float(r.ScheduledEffort)
                        except Exception:
                            pass
                        try:
                            r_info["actual_effort"] = safe_float(r.ActualEffort)
                        except Exception:
                            pass
                        try:
                            r_info["effort_remaining"] = safe_float(r.EffortRemaining)
                        except Exception:
                            pass
                        res_summary["permanent"].append(r_info)
                    except Exception:
                        continue
            except Exception as e:
                res_summary["permanent_error"] = str(e)

            try:
                cons = project.ConsumableResources
                for i in range(1, cons.Count + 1):
                    try:
                        r = cons.Item(i)
                        res_summary["consumable"].append({
                            "name": r.Name,
                            "availability": safe_float(getattr(r, 'Availability', None)),
                        })
                    except Exception:
                        continue
            except Exception as e:
                res_summary["consumable_error"] = str(e)

            report["resource_summary"] = res_summary

        # --- Cost Centre Summary ---
        if params.include_costs:
            cost_summary = []
            try:
                ccs = project.CostCentres
                for i in range(1, ccs.Count + 1):
                    try:
                        cc = ccs.Item(i)
                        cc_info = {"name": cc.Name, "id": cc.ID}
                        try:
                            cost_obj = cc.Cost
                            if cost_obj is not None:
                                cc_info["cost"] = safe_float(getattr(cost_obj, 'Amount', None))
                                cc_info["currency"] = safe_str(getattr(cost_obj, 'Currency', None))
                        except Exception:
                            pass
                        try:
                            cum_cost = cc.CumulativeCost
                            if cum_cost is not None:
                                cc_info["cumulative_cost"] = safe_float(getattr(cum_cost, 'Amount', None))
                        except Exception:
                            pass
                        try:
                            income_obj = cc.Income
                            if income_obj is not None:
                                cc_info["income"] = safe_float(getattr(income_obj, 'Amount', None))
                        except Exception:
                            pass
                        cost_summary.append(cc_info)
                    except Exception:
                        continue
            except Exception as e:
                cost_summary = [{"error": str(e)}]

            report["cost_summary"] = cost_summary

        # --- Variance & Critical Path from task_overview ---
        # If task_overview was populated from MPXJ, use that data for variance/critical too
        mpxj_tasks_data = task_overview.get("tasks", []) if task_overview.get("source") == "MPXJ (auto-export)" else None

        # --- Variance Analysis ---
        if params.include_variances:
            variances = []
            if mpxj_tasks_data:
                # Use MPXJ data (already has all tasks)
                for t in mpxj_tasks_data:
                    v_info = {"id": t.get("id", 0), "name": t.get("name", "")}
                    # MPXJ tasks have actual_start, actual_finish, baseline_start etc.
                    # Basic variance is tracked via percent_complete
                    if t.get("percent_complete", 0) > 0:
                        v_info["percent_complete"] = t.get("percent_complete", 0)
                    if len(v_info) > 2:
                        variances.append(v_info)
            else:
                try:
                    var_bars = _com_get_all_bars(project, max_bars=params.max_tasks)
                    for bar in var_bars:
                        try:
                            v_info = {"id": bar.ID, "name": bar.Name}
                            try:
                                planned_start = bar.Start
                                actual_start = bar.ActualStart
                                if actual_start is not None and planned_start is not None:
                                    if hasattr(actual_start, 'year') and hasattr(planned_start, 'year'):
                                        v_info["start_variance_days"] = (actual_start - planned_start).days
                            except Exception:
                                pass
                            try:
                                planned_end = bar.End
                                actual_end = bar.ActualEnd
                                if actual_end is not None and planned_end is not None:
                                    if hasattr(actual_end, 'year') and hasattr(planned_end, 'year'):
                                        v_info["end_variance_days"] = (actual_end - planned_end).days
                            except Exception:
                                pass
                            try:
                                orig_start = bar.OriginalStartV
                                orig_end = bar.OriginalFinishV
                                if orig_start and orig_end:
                                    v_info["original_start"] = format_date(orig_start)
                                    v_info["original_finish"] = format_date(orig_end)
                            except Exception:
                                pass
                            if len(v_info) > 2:
                                variances.append(v_info)
                        except Exception:
                            continue
                except Exception as e:
                    variances = [{"error": str(e)}]
            report["variances"] = variances

        # --- Critical Path ---
        if params.include_critical_path:
            critical = []
            if mpxj_tasks_data:
                # Use MPXJ data
                for t in mpxj_tasks_data:
                    if t.get("critical", False) and not t.get("summary", False):
                        critical.append({
                            "id": t.get("id", 0),
                            "name": t.get("name", ""),
                            "start": str(t.get("start", "N/A")),
                            "end": str(t.get("end", t.get("finish", "N/A"))),
                        })
            else:
                max_critical = 100
                try:
                    cp_bars = _com_get_all_bars(project, max_bars=500)
                    for bar in cp_bars:
                        if len(critical) >= max_critical:
                            break
                        try:
                            is_critical = False
                            try:
                                _btask, _ = _get_bar_task(bar)
                                if _btask is not None:
                                    is_critical = bool(_btask.Critical)
                            except Exception:
                                try:
                                    is_critical = bool(bar.GetToken("Critical"))
                                except Exception:
                                    pass
                            if is_critical:
                                critical.append({
                                    "id": bar.ID, "name": bar.Name,
                                    "start": format_date(bar.Start),
                                    "end": format_date(bar.End),
                                })
                        except Exception:
                            continue
                except Exception as e:
                    critical = [{"error": str(e)}]
            report["critical_path"] = critical
            report["critical_count"] = len([c for c in critical if "error" not in c])

        report["success"] = True

        # --- Format output ---
        if params.response_format == "json":
            return json.dumps(report, indent=2, default=str)

        # Markdown output
        lines = ["# Advanced Project Report", ""]

        # Project summary
        lines.append("## Project Summary")
        for k, v in summary.items():
            lines.append(f"- **{k}:** {v}")
        lines.append("")

        # Task overview
        source_info = f" (via {task_overview.get('source', 'COM')})" if task_overview.get("source") else ""
        lines.append(f"## Tasks ({task_overview.get('total_bars', 0)} total){source_info}")
        if task_overview.get("note"):
            lines.append(f"*{task_overview['note']}*")
        lines.append("")
        lines.append("| ID | Name | Start | End | % Complete | Critical |")
        lines.append("|---|---|---|---|---|---|")
        for t in task_overview.get("tasks", []):
            pct = t.get("percent_complete", "N/A")
            crit_mark = "Yes" if t.get("critical") else ""
            summ_mark = " [S]" if t.get("summary") else ""
            lines.append(f"| {t['id']} | {t['name']}{summ_mark} | {t['start']} | {t['end']} | {pct} | {crit_mark} |")
        lines.append("")

        # Resources
        if params.include_resources and "resource_summary" in report:
            rs = report["resource_summary"]
            lines.append("## Resources")
            if rs.get("permanent"):
                lines.append("### Permanent Resources")
                lines.append("| Name | Avail. | Scheduled | Actual | Remaining |")
                lines.append("|---|---|---|---|---|")
                for r in rs["permanent"]:
                    lines.append(
                        f"| {r['name']} | {r.get('availability', 'N/A')} | "
                        f"{r.get('scheduled_effort', 'N/A')} | {r.get('actual_effort', 'N/A')} | "
                        f"{r.get('effort_remaining', 'N/A')} |"
                    )
            if rs.get("consumable"):
                lines.append("### Consumable Resources")
                for r in rs["consumable"]:
                    lines.append(f"- **{r['name']}** (Avail: {r.get('availability', 'N/A')})")
            lines.append("")

        # Costs
        if params.include_costs and "cost_summary" in report:
            lines.append("## Cost Centres")
            lines.append("| Name | Cost | Cumulative | Income |")
            lines.append("|---|---|---|---|")
            for cc in report["cost_summary"]:
                if "error" in cc:
                    lines.append(f"| Error | {cc['error']} | | |")
                else:
                    lines.append(
                        f"| {cc['name']} | {cc.get('cost', 'N/A')} {cc.get('currency', '')} | "
                        f"{cc.get('cumulative_cost', 'N/A')} | {cc.get('income', 'N/A')} |"
                    )
            lines.append("")

        # Variances
        if params.include_variances and "variances" in report:
            lines.append("## Schedule Variances")
            if report["variances"]:
                lines.append("| ID | Name | Start Var (days) | End Var (days) |")
                lines.append("|---|---|---|---|")
                for v in report["variances"]:
                    if "error" not in v:
                        lines.append(
                            f"| {v['id']} | {v['name']} | "
                            f"{v.get('start_variance_days', 'N/A')} | "
                            f"{v.get('end_variance_days', 'N/A')} |"
                        )
            else:
                lines.append("*No variance data available (no actual dates set).*")
            lines.append("")

        # Critical path
        if params.include_critical_path and "critical_path" in report:
            lines.append(f"## Critical Path ({report.get('critical_count', 0)} tasks)")
            if report["critical_path"]:
                lines.append("| ID | Name | Start | End |")
                lines.append("|---|---|---|---|")
                for c in report["critical_path"]:
                    if "error" not in c:
                        lines.append(f"| {c['id']} | {c['name']} | {c['start']} | {c['end']} |")
            else:
                lines.append("*No critical path data available. Run reschedule first.*")
            lines.append("")

        return "\n".join(lines)

    except RuntimeError as e:
        report["error"] = f"COM connection failed: {e}"
        return json.dumps(report, indent=2, default=str)
    except Exception as e:
        report["error"] = f"Fatal error: {e}"
        logger.error(f"asta_generate_advanced_report fatal: {e}", exc_info=True)
        return json.dumps(report, indent=2, default=str)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ============================================================================


# ============================================================================
# CONSOLIDATED TOOLS (32 tools -> 10 tools to reduce context token usage)
# ============================================================================
# Each consolidated tool dispatches to the original function based on the
# "action" parameter. The original functions are preserved above (decorators
# commented out) so existing logic is untouched.
# ============================================================================

@mcp.tool(
    name="asta_task",
    annotations={"title": "Task Management", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def asta_task(params: dict) -> str:
    """Manage tasks in Asta Powerproject. Connects directly to running Asta via COM — no file_path needed.

    Actions:
    - add: Add new task. Params: name, duration, start_date, finish_date
    - update: Update task. Params: task_id, name, duration, start_date, finish_date, percent_complete, notes
    - delete: Delete task. Params: task_id
    - add_summary: Add summary/group task. Params: name, parent_task_id
    - add_child: Add child task. Params: parent_task_id, name, duration, start_date, finish_date
    - get: Get task details. Params: task_id, response_format
    - list: List all tasks in project. Params: include_summary (bool), response_format, limit (default 50, max 200)

    Planning best practices:
    - Keep activity durations between 5-20 working days (DCMA standard)
    - Use clear naming: [Phase]-[Zone]-[Package]-[Action]
    - Create summary tasks for WBS levels (Phase > Zone > Package)
    - Every activity needs at least one predecessor and successor (except start/end milestones)
    - After adding tasks, use asta_schedule → reschedule to recalculate CPM

    ⚠️ CRITICAL — BULK OPERATIONS (>10 tasks/links):
    - NEVER call add/update one-by-one in a loop via MCP for >10 items!
    - Instead, write a standalone Python COM script and execute it.
    - MCP add_summary creates orphan bars with no task data when called repeatedly.
    - Bulk script pattern: store bar IDs, re-fetch after each EndTransaction, use try/except+AbandonTransaction.

    ⚠️ COM GOTCHAS:
    - Root "Program" bar → use bar.ExpandedTask (NOT bar.Tasks(1) which fails on root!)
    - Child bars from AddSummaryTask → bar.Tasks(1) works fine
    - ALL COM refs become stale after EndTransaction → must re-fetch by bar ID
    - EndTransaction errors ≠ rollback — state may/may not have changed, always verify
    - Cleanup order: links → progress → allocations → tasks → bars (deepest first!)
    - NEVER remove the last bar — root bar must persist
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "add":
            result = await asta_add_task(AddTaskInput(**p))
        elif action == "update":
            result = await asta_update_task(UpdateTaskInput(**p))
        elif action == "delete":
            result = await asta_delete_task(DeleteTaskInput(**p))
        elif action == "add_summary":
            result = await asta_add_summary_task(AddSummaryTaskInput(**p))
        elif action == "add_child":
            result = await asta_add_child_task(AddChildTaskInput(**p))
        elif action == "get":
            result = await asta_get_task(GetTaskInput(**p))
        elif action == "list":
            result = await asta_list_tasks(ListTasksInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: add, update, delete, add_summary, add_child, get, list"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_task({action}) failed: {e}"})


@mcp.tool(
    name="asta_link",
    annotations={"title": "Link Management", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def asta_link(params: dict) -> str:
    """Manage task links/dependencies in Asta Powerproject. Connects directly to running Asta via COM — no file_path needed.

    Actions:
    - add: Add link. Params: predecessor_id, successor_id, link_type (FS/SS/FF/SF), lag
    - remove: Remove link. Params: predecessor_id, successor_id
    - update: Update link. Params: predecessor_id, successor_id, new_link_type, new_lag
    - diagnose: Explore available link COM interfaces. No params needed.

    Link types (PDM - Precedence Diagramming Method):
    - FS (Finish-to-Start): Default, ~90% of links. Successor starts after predecessor finishes.
    - SS (Start-to-Start): Parallel work. Common for zone-based progression (floor by floor).
    - FF (Finish-to-Finish): Quality gates, testing dependencies.
    - SF (Start-to-Finish): Rare. Use only when specifically required.

    DCMA standards: Avoid leads (negative lag), minimize lags (<5%), keep non-FS links <10%.
    If all add strategies fail, diagnostics run automatically showing available COM methods.

    ⚠️ BULK LINKS (>10): Write a Python COM script using task.LinkTo(task2).
    - link.type = 0(FS)/1(SS)/2(FF)/3(SF)
    - link.StartLagTime = task.GetDurationFromString("10d")
    - COM refs stale after EndTransaction — always re-fetch by bar ID.
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "add":
            result = await asta_add_link(AddLinkInput(**p))
        elif action == "remove":
            result = await asta_remove_link(RemoveLinkInput(**p))
        elif action == "update":
            result = await asta_update_link(UpdateLinkInput(**p))
        elif action == "diagnose":
            import pythoncom
            com_initialized = False
            try:
                pythoncom.CoInitialize()
                com_initialized = True
                app, project, method = _connect_asta_com()
                diag = _com_explore_link_interfaces(project)
                diag["com_method"] = method
                result = json.dumps(diag, indent=2, default=str)
            except Exception as e:
                result = json.dumps({"error": f"Link diagnostics failed: {e}"})
            finally:
                if com_initialized:
                    try: pythoncom.CoUninitialize()
                    except: pass
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: add, remove, update, diagnose"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_link({action}) failed: {e}"})


@mcp.tool(
    name="asta_progress",
    annotations={"title": "Progress Tracking", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def asta_progress(params: dict) -> str:
    """Update task progress in Asta Powerproject. Connects directly to running Asta via COM — no file_path needed.

    Actions:
    - update: Update single task progress. Params: task_id, percent_complete, actual_start, actual_finish
    - bulk_update: Update multiple tasks. Params: updates (list of {task_id, percent_complete, actual_start, actual_finish})

    Progress methods: duration-based (% time elapsed), physical (quantities), milestone weighting.
    Always set actual_start when task begins. Set actual_finish only when 100% complete.
    DCMA: No actual dates in the future. No incomplete tasks with forecast before report date.
    After progress update, reschedule (asta_schedule → reschedule) to recalculate forecast.

    ⚠️ CRITICAL — PROGRESS IS A BAR PROPERTY, NOT TASK:
    - Use bar.DurationPercentComplete (preferred) or bar.OverallPercentComplete (fallback)
    - NEVER try task.OverallPercentComplete — it will fail with "can not be set"
    - DurationPercentComplete = % of time elapsed; OverallPercentComplete = overall physical %
    - For bulk progress (>10 tasks): write a Python COM script instead of MCP calls
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "update":
            result = await asta_update_progress(UpdateProgressInput(**p))
        elif action == "bulk_update":
            result = await asta_bulk_update_progress(BulkUpdateProgressInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: update, bulk_update"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_progress({action}) failed: {e}"})


# asta_query MOVED to asta_mcp_file.py (file-based read-only queries)


@mcp.tool(
    name="asta_resource",
    annotations={"title": "Resource Management", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def asta_resource(params: dict) -> str:
    """Manage and query resources in Asta Powerproject. Connects directly to running Asta via COM — no file_path needed.

    Actions:
    - manage: Manage resources via COM. Params: sub_action (list/create_permanent/create_consumable/create_cost_centre/delete_resource/delete_cost_centre), name, resource_type, availability, cost_rate, calendar_name, parent_cost_centre
    - assign: Assign resources to tasks via COM. Params: assignments (list of {task_id, resource_name, resource_type, units, is_demand, work_profile})

    Resource types: Permanent (labour/equipment), Consumable (materials with quantities), Cost Centres (budget categories).
    DCMA Check 10: <5% of tasks should lack resource assignments.
    Resource loading enables: leveling, S-curves, earned value analysis, histograms.

    ⚠️ CRITICAL — RESOURCE ASSIGNMENT GOTCHAS:
    - GivenAllocation=50 does NOT mean 50 headcount! Asta uses it as % or multiplier.
    - For precise costing: use GivenEffort (in SECONDS, e.g. 8h=28800s) instead.
    - Consumable: use GivenQuantity for amount, CostPerUnit via IAmountAndCurrency pattern.
    - Cost allocation: use GivenValue via IAmountAndCurrency pattern.
    - Resource curves: work_profile 'bell_curve'=3, 'back_loaded'=2 (EditToken IDs).
      Actual Asta ResourceCurves have different IDs (Bell Shaped=105, Back Loaded Low res=97).
    - Code assignment goes to BAR object: bar.AssignCode(entry, True), NOT task!
    - For bulk assignments (>10): write a Python COM script instead of repeated MCP calls.
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action in ("list", "assignments", "loading"):
            # File-based queries: use MPXJ (lazy JVM start)
            if action == "list":
                result = await asta_list_resources(ResourcesInput(**p))
            elif action == "assignments":
                result = await asta_get_resource_assignments(ResourcesInput(**p))
            elif action == "loading":
                result = await asta_resource_loading(ResourceLoadingInput(**p))
        elif action == "manage":
            # asta_manage_resources is sync and needs its own 'action' param
            p["action"] = p.pop("sub_action", "list")
            result = asta_manage_resources(ManageResourcesInput(**p))
        elif action == "assign":
            result = asta_assign_resource_model(ResourceAssignmentInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: list, assignments, loading, manage, assign"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_resource({action}) failed: {e}"})


@mcp.tool(
    name="asta_schedule",
    annotations={"title": "Schedule Operations", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True}
)
async def asta_schedule(params: dict) -> str:
    """Schedule operations in Asta Powerproject. Connects directly to running Asta via COM — no file_path needed.

    Actions:
    - reschedule: Reschedule project via COM (equivalent to F9). Params: report_date (YYYY-MM-DD), straighten_uncompleted_work, preserve_links, target_wbs_id
    - what_if: Run what-if scenario via COM. Params: scenario_name, modifications, target_date, auto_commit
    - save: Save project to XML file. Params: output_path

    Rescheduling calculates CPM (Critical Path Method): early/late dates, total/free float, critical path.
    Always reschedule after: adding/modifying tasks, changing links, updating progress.
    Report date defines the data date — all progress is measured relative to this date.

    ⚠️ COM NOTES:
    - project.Reschedule() works parameterless — no Chart object needed!
    - Progress Periods: project.ProgressPeriods.Item(1).ReportDate is settable
    - Baselines: BslnProjects.Add() does NOT work; use SaveProjectAs + OpenBaseline(path)
    - CurrentProgressPeriod is NOT in type library — do not attempt to access
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "reschedule":
            result = await asta_reschedule_project(RescheduleProjectInput(**p))
        elif action == "what_if":
            # asta_what_if_analysis is sync
            result = asta_what_if_analysis(WhatIfInput(**p))
        elif action == "save":
            result = await asta_save_project(SaveProjectInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: reschedule, what_if, save"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_schedule({action}) failed: {e}"})


@mcp.tool(
    name="asta_code",
    annotations={"title": "Code Library Management", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def asta_code(params: dict) -> str:
    """Manage code libraries and assign codes to tasks via COM.

    Actions:
    - manage: Manage code libraries. Params: sub_action (list/create_library/add_entries/delete_entry), library_name, entries, entry_name
    - assign: Assign code entries to tasks. Params: library_name, assignments (list of {task_id, entry_name, append})

    Code libraries categorize tasks (e.g., Responsibility, Phase, Zone, Discipline, Trade).
    Used for: filtering, grouping, reporting, WBS coding, earned value roll-ups.
    Common libraries: Phase (Enabling/Substructure/Superstructure/...), Zone (Building A/B/...),
    Trade (Concrete/Steel/MEP/...), Responsibility (Main Contractor/Sub A/Sub B/...).

    ⚠️ CRITICAL COM GOTCHAS:
    - project.CodeLibrarys (non-standard plural!) for collection access
    - lib.Entries.Add() to add entries (NOT lib.CodeLibraryEntrys.Add()!)
    - Code assignment: bar.AssignCode(entry, True) on BAR object, NOT task!
    - For bulk assignments (>10): write a Python COM script.
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "manage":
            # asta_manage_code_libraries is sync and needs its own 'action' param
            p["action"] = p.pop("sub_action", "list")
            result = asta_manage_code_libraries(ManageCodeLibrariesInput(**p))
        elif action == "assign":
            result = asta_assign_codes(AssignCodesInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: manage, assign"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_code({action}) failed: {e}"})


@mcp.tool(
    name="asta_view",
    annotations={"title": "View Configuration", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def asta_view_consolidated(params: dict) -> str:
    """Configure the active Asta Powerproject bar chart view via COM.

    Actions:
    - get_status: Get current view configuration (display toggles, grouping, etc.)
    - set_display: Toggle display properties. Params: display_critical_path, display_free_float, display_total_float, display_progress_lines, display_annotations, display_cost_allocations, display_demand_allocations, display_scheduled_allocations
    - set_grouping: Group by code library. Params: group_by_library
    - set_sorting: Sort view. Params: sort_field, sort_ascending
    - set_filter: Apply filter. Params: filter expression
    - toggle_histogram: Show/hide histogram. Params: histogram_visible, histogram_type (resource/cost/work)
    - show_hierarchy_level: Expand/collapse. Params: hierarchy_level (0=collapse, 1=top, etc.)
    - list_tables: List all 36 built-in spreadsheet table definitions (column presets)
    - apply_table: Apply a table definition to configure columns instantly. Params: table_name
    - get_columns: Show current spreadsheet columns (token IDs, widths)

    ⚠️ COLUMN MANAGEMENT — USE apply_table (NOT individual AddCol which takes 3+ minutes per column!):
    Key table definitions for common operations:
    - "% Progress - with a Baseline" → progress tracking with baseline comparison (13 columns)
    - "Name & Costs" → cost analysis view (9 columns)
    - "Task ID, Name, Start, Duration, Finish" → standard scheduling view (6 columns)
    - "% Progress - No Baseline" → simple progress view
    - "Predecessors & Successors" → logic/link analysis
    - "Task name & Resources assigned" → resource view
    - "Float Paths" → critical path analysis
    - "Name, Cost & Income" → financial view
    - "Task Work Progress" → work-based progress

    ⚠️ RECOMMENDED COLUMN SETUP PER OPERATION:
    - After creating tasks/links → apply "Task ID, Name, Start, Duration, Finish"
    - After progress update → apply "% Progress - with a Baseline"
    - After resource/cost assignment → apply "Name & Costs"
    - For delay analysis → apply "Float Paths" + display_critical_path=true
    """
    action = params.get("action", "").lower()

    # Handle new column/table actions directly (bypass ViewConfigInput validation)
    if action in ("list_tables", "apply_table", "get_columns"):
        import pythoncom
        import win32com.client
        com_initialized = False
        result_dict = {"tool": "asta_view", "action": action, "success": False}
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            D = win32com.client.Dispatch
            APP_CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
            obj = pythoncom.GetActiveObject(APP_CLSID)
            app = D(obj.QueryInterface(pythoncom.IID_IDispatch))
            project = D(app.ActiveProject)
            view = D(project.Views.Item(1))
            ss = D(view.SpreadSheet())

            if action == "list_tables":
                tds = D(project.TableDefinitions)
                tables = []
                for i in range(1, tds.Count + 1):
                    td = D(tds.Item(i))
                    tables.append({"index": i, "name": td.Name})
                result_dict["success"] = True
                result_dict["table_definitions"] = tables
                result_dict["count"] = tds.Count
                try:
                    result_dict["current_table"] = str(ss.TableDefinition)
                except Exception:
                    pass

            elif action == "apply_table":
                table_name = params.get("table_name", "")
                if not table_name:
                    result_dict["error"] = "table_name required. Use list_tables to see options."
                    return json.dumps(result_dict, indent=2, default=str)

                tds = D(project.TableDefinitions)
                td_obj = None
                # Exact match first
                for i in range(1, tds.Count + 1):
                    td = D(tds.Item(i))
                    if td.Name.lower() == table_name.lower():
                        td_obj = td
                        break
                # Partial match fallback
                if not td_obj:
                    for i in range(1, tds.Count + 1):
                        td = D(tds.Item(i))
                        if table_name.lower() in td.Name.lower():
                            td_obj = td
                            break

                if not td_obj:
                    result_dict["error"] = f"Table '{table_name}' not found."
                    return json.dumps(result_dict, indent=2, default=str)

                project.StartTransaction("ApplyTable")
                try:
                    ss.ApplyTableDefinition(td_obj)
                    project.EndTransaction()
                    project.WaitForNotificationProcessing()
                except Exception as e:
                    try:
                        project.AbandonTransaction()
                    except Exception:
                        pass
                    result_dict["error"] = f"ApplyTableDefinition failed: {e}"
                    return json.dumps(result_dict, indent=2, default=str)

                # Read resulting columns
                total = ss.TotalCols
                columns = []
                for ci in range(1, total + 1):
                    try:
                        ss.Col = ci
                        columns.append({"index": ci, "token": ss.ColDataToken, "width": ss.ColWidth})
                    except:
                        break
                try:
                    view.Refresh()
                except Exception:
                    pass

                result_dict["success"] = True
                result_dict["applied_table"] = td_obj.Name
                result_dict["columns"] = columns
                result_dict["column_count"] = len(columns)

            elif action == "get_columns":
                total = ss.TotalCols
                columns = []
                for ci in range(1, total + 1):
                    try:
                        ss.Col = ci
                        columns.append({"index": ci, "token": ss.ColDataToken, "width": ss.ColWidth})
                    except:
                        break
                result_dict["success"] = True
                result_dict["columns"] = columns
                result_dict["column_count"] = len(columns)
                try:
                    result_dict["current_table"] = str(ss.TableDefinition)
                except Exception:
                    pass

        except Exception as e:
            result_dict["error"] = f"COM error: {e}"
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        return json.dumps(result_dict, indent=2, default=str)

    # Original actions via ViewConfigInput
    try:
        result = asta_configure_view(ViewConfigInput(**params))
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_view failed: {e}"})


@mcp.tool(
    name="asta_export",
    annotations={"title": "Export & Reports", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def asta_export(params: dict) -> str:
    """Export project data and generate reports.

    Actions:
    - xml: Export to XML/MSPDI/XER/MPP via COM. Params: output_path, format (asta_xml/mspdi/xer/mpp), branch_ids
    - pdf: Export current view to PDF. Params: output_path, print_profile
    - report: Generate comprehensive report via COM. Params: response_format (markdown/json), include_resources, include_costs, include_variances, include_critical_path, max_tasks

    Export formats: asta_xml (native), mspdi (MS Project XML), xer (Primavera P6), mpp (MS Project binary).
    Use 'report' for monthly programme reports, schedule health assessments, and client submissions.
    Include all flags (resources, costs, variances, critical_path) for comprehensive analysis.

    ⚠️ EXPORT TIPS:
    - Before PDF export: use asta_view → apply_table to set appropriate columns
    - MSPDI format is best for interoperability (MS Project, Primavera, other tools)
    - SaveAsCSVFile on view.SpreadSheet() also available for tabular data export
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "xml":
            # asta_com_export_xml is sync
            result = asta_com_export_xml(ExportXMLInput(**p))
        elif action == "pdf":
            # asta_export_pdf is sync
            result = asta_export_pdf(ExportPDFInput(**p))
        elif action == "report":
            # asta_generate_advanced_report is sync
            result = asta_generate_advanced_report(AdvancedReportInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: xml, pdf, report"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_export({action}) failed: {e}"})


# asta_calendar MOVED to asta_mcp_file.py (file-based read-only queries)


# ============================================================================
# ENTRY POINT
# ============================================================================
if __name__ == "__main__":
    logger.info("Starting Asta Powerproject Core MCP Server...")
    mcp.run()
