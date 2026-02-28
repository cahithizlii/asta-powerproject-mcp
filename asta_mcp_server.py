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
            os.path.join(os.path.expanduser("~"), "asta_mcp.log"),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger("asta_mcp")

# ============================================================================
# INITIALIZE MCP SERVER
# ============================================================================
mcp = FastMCP("asta_powerproject_mcp")

# ============================================================================
# CONSTANTS
# ============================================================================
SUPPORTED_EXTENSIONS = ['.pp', '.mpp', '.xml', '.mspdi', '.xer', '.pmxml']
DEFAULT_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
ASTA_WINDOW_TITLE = "Asta Powerproject"

# ============================================================================
# PRE-START JVM (so first tool call is fast)
# ============================================================================
try:
    import mpxj
    if not mpxj.jpype.isJVMStarted():
        mpxj.jpype.startJVM()
        logger.info("JVM pre-started successfully")
    else:
        logger.info("JVM already running")
except Exception as e:
    logger.error(f"Failed to pre-start JVM: {e}")

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
class AstaGUIManager:
    """Manages GUI automation for Asta Powerproject using pyautogui/pywinauto."""

    @staticmethod
    def _check_gui_libs():
        """Check if GUI automation libraries are available."""
        missing = []
        try:
            import pyautogui
        except ImportError:
            missing.append("pyautogui")
        try:
            import pywinauto
        except ImportError:
            missing.append("pywinauto")
        if missing:
            return False, f"Missing libraries: {', '.join(missing)}. Install with: pip install {' '.join(missing)}"
        return True, "OK"

    @staticmethod
    def find_asta_window():
        """Find the Asta Powerproject window."""
        try:
            import pywinauto
            from pywinauto import Desktop
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            for w in windows:
                title = w.window_text()
                if "powerproject" in title.lower() or "asta" in title.lower():
                    return {"found": True, "title": title, "handle": w.handle}
            return {"found": False, "message": "Asta Powerproject window not found. Please open the application first."}
        except Exception as e:
            return {"found": False, "message": f"Error finding window: {str(e)}"}

    @staticmethod
    def bring_to_front():
        """Bring Asta Powerproject to the foreground."""
        try:
            import pywinauto
            from pywinauto import Desktop
            desktop = Desktop(backend="uia")
            for w in desktop.windows():
                if "powerproject" in w.window_text().lower() or "asta" in w.window_text().lower():
                    w.set_focus()
                    time.sleep(0.5)
                    return {"success": True, "message": "Asta brought to foreground"}
            return {"success": False, "message": "Asta window not found"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def send_shortcut(keys: str, delay: float = 0.5):
        """Send keyboard shortcut to Asta."""
        try:
            import pyautogui
            # Bring Asta to front first
            AstaGUIManager.bring_to_front()
            time.sleep(delay)
            pyautogui.hotkey(*keys.split('+'))
            time.sleep(delay)
            return {"success": True, "message": f"Shortcut '{keys}' sent"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def click_at(x: int, y: int, clicks: int = 1, button: str = "left"):
        """Click at specific coordinates."""
        try:
            import pyautogui
            AstaGUIManager.bring_to_front()
            time.sleep(0.3)
            pyautogui.click(x=x, y=y, clicks=clicks, button=button)
            time.sleep(0.3)
            return {"success": True, "message": f"Clicked at ({x}, {y})"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def type_text(text: str, interval: float = 0.05):
        """Type text in the currently focused field."""
        try:
            import pyautogui
            _clipboard_paste(text)
            time.sleep(0.3)
            return {"success": True, "message": f"Typed: {text}"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def take_screenshot(save_path: str = None) -> dict:
        """Take a screenshot of the Asta window."""
        if not save_path:
            save_path = os.path.join(
                os.path.expanduser("~"), "Downloads",
                f"asta_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            )
        # Try pyautogui first
        try:
            import pyautogui
            screenshot = pyautogui.screenshot()
            screenshot.save(save_path)
            return {"success": True, "path": save_path, "message": f"Screenshot saved: {save_path}"}
        except Exception as pyautogui_err:
            logger.warning(f"pyautogui screenshot failed: {pyautogui_err}")

        # Fallback: use mss (if available)
        try:
            import mss
            with mss.mss() as sct:
                sct.shot(output=save_path)
            return {"success": True, "path": save_path, "message": f"Screenshot saved (mss fallback): {save_path}"}
        except ImportError:
            pass
        except Exception as mss_err:
            logger.warning(f"mss screenshot also failed: {mss_err}")

        # Final fallback: PowerShell
        try:
            ps_script = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds; "
                "$bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height); "
                "$graphics = [System.Drawing.Graphics]::FromImage($bitmap); "
                "$graphics.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size); "
                f"$bitmap.Save('{save_path.replace(chr(39), chr(39)+chr(39))}'); "
                "$graphics.Dispose(); $bitmap.Dispose()"
            )
            subprocess.run([_get_powershell_path(), '-NoProfile', '-command', ps_script],
                           capture_output=True, timeout=10)
            if os.path.exists(save_path):
                return {"success": True, "path": save_path, "message": f"Screenshot saved (PowerShell fallback): {save_path}"}
        except Exception as ps_err:
            logger.warning(f"PowerShell screenshot also failed: {ps_err}")

        return {
            "success": False,
            "message": "All screenshot methods failed. Try: pip install --upgrade Pillow pyscreeze"
        }

    @staticmethod
    def navigate_menu(menu_path: List[str], delay: float = 0.5):
        """Navigate through Asta ribbon/menu system."""
        try:
            import pyautogui
            AstaGUIManager.bring_to_front()
            time.sleep(delay)
            # Click each menu item in sequence
            for item in menu_path:
                # Use Alt key for ribbon tabs
                pyautogui.press('alt')
                time.sleep(0.3)
                _clipboard_paste(item)
                time.sleep(delay)
            return {"success": True, "message": f"Navigated: {' > '.join(menu_path)}"}
        except Exception as e:
            return {"success": False, "message": str(e)}


# ============================================================================
# PYDANTIC INPUT MODELS
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
        default=100,
        description="Maximum number of tasks to return (default 100, max 500). Use asta_get_task for details on a specific task.",
        ge=1, le=500
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
    save_output: Optional[str] = Field(
        default=None,
        description="Path to save the modified file. If empty, auto-generates a timestamped XML file"
    )


class UpdateTaskInput(ProjectFileInput):
    """Input for updating an existing task."""
    task_id: int = Field(..., description="ID of the task to update", ge=0)
    name: Optional[str] = Field(default=None, description="New task name")
    duration: Optional[str] = Field(default=None, description="New duration (e.g., '5d', '2w')")
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
class GUIShortcutInput(BaseModel):
    """Input for sending keyboard shortcuts."""
    model_config = ConfigDict(str_strip_whitespace=True)
    shortcut: str = Field(
        ...,
        description="Keyboard shortcut to send. Examples: 'ctrl+s' (save), 'F9' (reschedule), 'ctrl+z' (undo), 'ctrl+p' (print)"
    )


class GUIClickInput(BaseModel):
    """Input for clicking at coordinates."""
    x: int = Field(..., description="X coordinate (pixels from left)", ge=0)
    y: int = Field(..., description="Y coordinate (pixels from top)", ge=0)
    clicks: int = Field(default=1, description="Number of clicks (1=single, 2=double)", ge=1, le=3)
    button: str = Field(default="left", description="Mouse button: 'left', 'right', or 'middle'")


class GUITypeInput(BaseModel):
    """Input for typing text."""
    model_config = ConfigDict(str_strip_whitespace=False)
    text: str = Field(..., description="Text to type in the currently focused cell/field")
    press_enter: bool = Field(default=False, description="Press Enter after typing")


class GUIMenuInput(BaseModel):
    """Input for menu navigation."""
    tab: str = Field(
        ...,
        description="Ribbon tab name: 'Home', 'View', 'Project', 'Allocation', 'Format', 'File'"
    )
    command: Optional[str] = Field(
        default=None,
        description="Command to click within the tab (e.g., 'Reschedule', 'Link Tasks', 'Summarise')"
    )


class GUIScreenshotInput(BaseModel):
    """Input for taking screenshots."""
    save_path: Optional[str] = Field(
        default=None,
        description="File path to save screenshot. Auto-generates if empty"
    )


class GUINewProjectInput(BaseModel):
    """Input for creating a new project via GUI."""
    project_name: str = Field(..., description="Name for the new project", min_length=1)
    client_name: Optional[str] = Field(default=None, description="Client/customer name (For field)")
    contractor_name: Optional[str] = Field(default=None, description="Contractor name (By field)")
    start_date: Optional[str] = Field(
        default=None,
        description="Project start date in DD/MM/YYYY format. Uses today if empty"
    )
    template: str = Field(
        default="Construction Template",
        description="Template to use: 'Construction Template', 'Housing Template', 'Blank Project', etc."
    )


class GUIProgressInput(BaseModel):
    """Input for entering progress via GUI."""
    task_name: str = Field(..., description="Name of the task to update progress for")
    percent_complete: float = Field(
        ..., description="Completion percentage (0-100)", ge=0, le=100
    )
    actual_start: Optional[str] = Field(
        default=None,
        description="Actual start date in DD/MM/YYYY format"
    )
    actual_finish: Optional[str] = Field(
        default=None,
        description="Actual finish date in DD/MM/YYYY format"
    )


class GUIOpenFileInput(BaseModel):
    """Input for opening a project file in Asta."""
    file_path: str = Field(
        ...,
        description="Full path to the project file to open in Asta Powerproject"
    )


class GUIFilterInput(BaseModel):
    """Input for applying filters."""
    filter_type: str = Field(
        ...,
        description="Type of filter: 'critical' (critical tasks only), 'complete' (completed tasks), 'incomplete' (not completed), 'code' (by code library), 'none' (remove filter)"
    )
    code_name: Optional[str] = Field(
        default=None,
        description="Code library name (only used when filter_type='code')"
    )


class GUIBaselineInput(BaseModel):
    """Input for taking a baseline."""
    baseline_name: str = Field(
        default="Original Plan",
        description="Name for the baseline snapshot"
    )


class GUILinkTasksInput(BaseModel):
    """Input for linking tasks in GUI."""
    predecessor_row: int = Field(..., description="Row number of the predecessor task", ge=1)
    successor_row: int = Field(..., description="Row number of the successor task", ge=1)
    link_type: str = Field(
        default="FS",
        description="Link type: 'FS' (Finish-to-Start), 'SS' (Start-to-Start), 'FF' (Finish-to-Finish), 'SF' (Start-to-Finish)"
    )
    lag: Optional[str] = Field(
        default=None,
        description="Lag/lead time (e.g., '2d' for 2 day lag, '-1d' for 1 day lead)"
    )


class GUIPrintInput(BaseModel):
    """Input for printing/exporting."""
    output_type: str = Field(
        default="pdf",
        description="Output type: 'pdf', 'clipboard', 'picture', 'printer'"
    )
    save_path: Optional[str] = Field(
        default=None,
        description="File path for PDF or picture output"
    )


# ============================================================================
# FILE-BASED TOOLS (MPXJ)
# ============================================================================

@mcp.tool(
    name="asta_analyze_project",
    annotations={
        "title": "Analyze Asta Project",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
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
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_list_tasks",
    annotations={
        "title": "List Project Tasks",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_list_tasks(params: ListTasksInput) -> str:
    """List all tasks in an Asta project file with their properties.

    Returns task ID, name, duration, dates, completion %, critical status,
    and predecessor/successor relationships.

    Args:
        params: Contains file_path, include_summary, response_format

    Returns:
        Task list in markdown or JSON format
    """
    try:
        mgr = AstaFileManager(params.file_path)
        tasks = mgr.get_all_tasks(include_summary=params.include_summary)

        total = len(tasks)
        limited = tasks[:params.limit]

        if params.response_format == ResponseFormat.JSON:
            result = {"total": total, "returned": len(limited), "limit": params.limit, "tasks": limited}
            if total > params.limit:
                result["note"] = f"Showing {params.limit} of {total} tasks. Increase 'limit' or use asta_get_task for specific task details."
            return json.dumps(result, indent=2, default=str)

        lines = ["# Task List", ""]
        if total > params.limit:
            lines.append(f"**Showing {params.limit} of {total} tasks** (use `limit` parameter to see more, max 500)")
        else:
            lines.append(f"**Total:** {total} tasks")
        lines.append("")
        for t in limited:
            prefix = "[SUMMARY] " if t['summary'] else "[MILESTONE] " if t['milestone'] else ""
            crit = " **[CRITICAL]**" if t['critical'] else ""
            lines.append(f"### {prefix}{t['name']} (ID: {t['id']}){crit}")
            lines.append(f"- Duration: {t['duration']} | Start: {t['start']} | Finish: {t['finish']}")
            lines.append(f"- Progress: {t['percent_complete']}% | Float: {t['total_float']}")
            if t['predecessors']:
                pred_str = ", ".join([f"Task {p['task_id']} ({p['type']}, lag: {p['lag']})" for p in t['predecessors']])
                lines.append(f"- Predecessors: {pred_str}")
            if t['successors']:
                succ_str = ", ".join([f"Task {s['task_id']} ({s['type']}, lag: {s['lag']})" for s in t['successors']])
                lines.append(f"- Successors: {succ_str}")
            if t['notes']:
                lines.append(f"- Notes: {t['notes']}")
            lines.append("")
        return "\n".join(lines)

    except Exception as e:
        return f"Error listing tasks: {str(e)}"


@mcp.tool(
    name="asta_get_task",
    annotations={
        "title": "Get Task Details",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_get_task(params: GetTaskInput) -> str:
    """Get detailed information about a specific task by its ID.

    Returns comprehensive task data including early/late dates,
    float values, actual dates, cost, work, and calendar info.

    Args:
        params: Contains file_path, task_id, response_format

    Returns:
        Detailed task information in markdown or JSON
    """
    try:
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_add_task",
    annotations={
        "title": "Add New Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
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
    # --- COM mode (no file_path) ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Add Task")
            result = _com_add_task(project, params.name, params.duration)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM add task failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode (MPXJ) ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.add_task(params.name, params.duration)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        result["reminder"] = "Open this file in Asta and press F9 (Reschedule) to update the schedule"
        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error adding task: {str(e)}"


@mcp.tool(
    name="asta_update_task",
    annotations={
        "title": "Update Existing Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_update_task(params: UpdateTaskInput) -> str:
    """Update properties of an existing task in the project file.

    Can update task name, duration, completion percentage, and notes.
    Saves the modified project to a new file (preserving the original).

    Args:
        params: Contains file_path, task_id, name, duration, percent_complete, notes, save_output

    Returns:
        Confirmation with updated fields and saved file path
    """
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Update Task")
            result = _com_update_task(project, params.task_id,
                                      name=params.name, duration_str=params.duration,
                                      percent_complete=params.percent_complete,
                                      notes=params.notes)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM update task failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_delete_task",
    annotations={
        "title": "Delete Task",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
async def asta_delete_task(params: DeleteTaskInput) -> str:
    """Delete a task from the project file by its ID.

    WARNING: This permanently removes the task and its links.
    The original file is preserved; changes are saved to a new file.

    Args:
        params: Contains file_path, task_id, save_output

    Returns:
        Confirmation of deletion with saved file path
    """
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Delete Task")
            result = _com_delete_task(project, params.task_id)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM delete task failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.delete_task(params.task_id)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error deleting task: {str(e)}"


@mcp.tool(
    name="asta_get_critical_path",
    annotations={
        "title": "Get Critical Path",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
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
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_list_resources",
    annotations={
        "title": "List Project Resources",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_list_resources(params: ResourcesInput) -> str:
    """List all resources (labour, equipment, materials) in the project.

    Shows resource ID, name, type, maximum units, rate, and cost.

    Args:
        params: Contains file_path, response_format

    Returns:
        Resource list in markdown or JSON format
    """
    try:
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_get_resource_assignments",
    annotations={
        "title": "Get Resource Assignments",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_get_resource_assignments(params: ResourcesInput) -> str:
    """Get all resource assignments showing which resources are assigned to which tasks.

    Shows task name, resource name, units, work hours, and cost for each assignment.

    Args:
        params: Contains file_path, response_format

    Returns:
        Resource assignment list in markdown or JSON format
    """
    try:
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_get_calendars",
    annotations={
        "title": "Get Project Calendars",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_get_calendars(params: ProjectFileInput) -> str:
    """Get all calendars defined in the project.

    Calendars define working days, hours, and exceptions (holidays, overtime).

    Args:
        params: Contains file_path

    Returns:
        Calendar list in JSON format
    """
    try:
        mgr = AstaFileManager(params.file_path)
        calendars = mgr.get_calendars()
        return json.dumps({"total": len(calendars), "calendars": calendars}, indent=2, default=str)

    except Exception as e:
        return f"Error getting calendars: {str(e)}"


@mcp.tool(
    name="asta_float_analysis",
    annotations={
        "title": "Float Analysis",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
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
        mgr = AstaFileManager(params.file_path)
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

@mcp.tool(
    name="asta_get_wbs_tree",
    annotations={"title": "Get WBS/Hierarchy Tree", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
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
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_add_summary_task",
    annotations={"title": "Add Summary Task", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def asta_add_summary_task(params: AddSummaryTaskInput) -> str:
    """Add a new summary (group) task to organize other tasks.

    Summary tasks act as containers/folders for child tasks.
    Can be added at top level or nested under another task.

    Args:
        params: Contains file_path, name, parent_task_id, save_output

    Returns:
        Confirmation with new summary task details
    """
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Add Summary Task")
            result = _com_add_task(project, params.name,
                                   parent_bar_id=params.parent_task_id,
                                   is_summary=True)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM add summary task failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.add_summary_task(params.name, params.parent_task_id)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error adding summary task: {str(e)}"


@mcp.tool(
    name="asta_add_child_task",
    annotations={"title": "Add Child Task", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def asta_add_child_task(params: AddChildTaskInput) -> str:
    """Add a new task under a specific parent/summary task.

    Creates the task as a child of the specified parent,
    maintaining the WBS hierarchy.

    Args:
        params: Contains file_path, parent_task_id, name, duration, save_output

    Returns:
        Confirmation with new task details and parent info
    """
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Add Child Task")
            result = _com_add_task(project, params.name, params.duration,
                                   parent_bar_id=params.parent_task_id)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM add child task failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.add_child_task(params.parent_task_id, params.name, params.duration)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error adding child task: {str(e)}"


@mcp.tool(
    name="asta_add_link",
    annotations={"title": "Add Task Link", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
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
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Add Link")
            result = _com_add_link(project, params.predecessor_id, params.successor_id,
                                   params.link_type, params.lag)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM add link failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.add_link(params.predecessor_id, params.successor_id, params.link_type, params.lag)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error adding link: {str(e)}"


@mcp.tool(
    name="asta_remove_link",
    annotations={"title": "Remove Task Link", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False}
)
async def asta_remove_link(params: RemoveLinkInput) -> str:
    """Remove a predecessor-successor link between two tasks.

    WARNING: This permanently removes the dependency link.
    The tasks themselves are NOT deleted.

    Args:
        params: Contains file_path, predecessor_id, successor_id, save_output

    Returns:
        Confirmation of link removal
    """
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Remove Link")
            result = _com_remove_link(project, params.predecessor_id, params.successor_id)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM remove link failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.remove_link(params.predecessor_id, params.successor_id)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error removing link: {str(e)}"


@mcp.tool(
    name="asta_update_link",
    annotations={"title": "Update Task Link", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def asta_update_link(params: UpdateLinkInput) -> str:
    """Update an existing task link's type or lag.

    Can change the link type (FS/SS/FF/SF) and/or the lag duration.
    Internally removes the old link and creates a new one.

    Args:
        params: Contains file_path, predecessor_id, successor_id, new_link_type, new_lag, save_output

    Returns:
        Confirmation with old and new link properties
    """
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Update Link")
            result = _com_update_link(project, params.predecessor_id, params.successor_id,
                                      params.new_link_type, params.new_lag)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM update link failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.update_link(params.predecessor_id, params.successor_id, params.new_link_type, params.new_lag)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error updating link: {str(e)}"


@mcp.tool(
    name="asta_update_progress",
    annotations={"title": "Update Task Progress", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def asta_update_progress(params: UpdateProgressInput) -> str:
    """Update progress data for a single task.

    Can set completion percentage, actual start date, and actual finish date.
    Dates can be in YYYY-MM-DD or DD/MM/YYYY format.

    Args:
        params: Contains file_path, task_id, percent_complete, actual_start, actual_finish, save_output

    Returns:
        Confirmation with updated progress details
    """
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
            project.StartTransaction("Update Progress")
            result = _com_update_progress(project, params.task_id,
                                          percent_complete=params.percent_complete,
                                          actual_start=params.actual_start,
                                          actual_finish=params.actual_finish)
            if "error" in result:
                project.AbandonTransaction()
                return json.dumps(result, indent=2)
            project.EndTransaction()
            result["com_method"] = method
            return json.dumps(result, indent=2, default=str)
        except Exception as e:
            return json.dumps({"error": f"COM update progress failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        result = mgr.update_progress(params.task_id, params.percent_complete, params.actual_start, params.actual_finish)
        if "error" in result:
            return json.dumps(result, indent=2)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error updating progress: {str(e)}"


@mcp.tool(
    name="asta_bulk_update_progress",
    annotations={"title": "Bulk Update Progress", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
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
    # --- COM mode ---
    if not params.file_path:
        import pythoncom
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            app, project, method = _connect_asta_com()
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

            project.EndTransaction()

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
        except Exception as e:
            return json.dumps({"error": f"COM bulk progress update failed: {e}"}, indent=2)
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- File mode ---
    try:
        mgr = AstaFileManager(params.file_path)
        updates_list = [u.model_dump() for u in params.updates]
        result = mgr.bulk_update_progress(updates_list)
        output = mgr.save(params.save_output)
        result["saved_to"] = output
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error in bulk progress update: {str(e)}"


@mcp.tool(
    name="asta_delay_analysis",
    annotations={"title": "Delay Analysis", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
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
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_resource_loading",
    annotations={"title": "Resource Loading Analysis", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
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
        mgr = AstaFileManager(params.file_path)
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


@mcp.tool(
    name="asta_save_project",
    annotations={
        "title": "Save Project File",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_save_project(params: SaveProjectInput) -> str:
    """Save the project to an XML file (compatible with Asta Powerproject).

    Creates a new XML file that can be imported into Asta Powerproject.
    The original file is never overwritten.

    Args:
        params: Contains file_path and optional output_path

    Returns:
        Path to the saved file
    """
    try:
        mgr = AstaFileManager(params.file_path)
        output = mgr.save(params.output_path)
        return json.dumps({
            "success": True,
            "saved_to": output,
            "message": f"Project saved to: {output}. Open in Asta and press F9 to reschedule."
        }, indent=2)

    except Exception as e:
        return f"Error saving project: {str(e)}"


# ============================================================================
# GUI AUTOMATION TOOLS
# ============================================================================

@mcp.tool(
    name="asta_gui_check_status",
    annotations={
        "title": "Check Asta GUI Status",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_check_status() -> str:
    """Check if Asta Powerproject is running and get window information.

    Use this FIRST before any GUI automation to verify Asta is open.
    Also checks if required GUI automation libraries are installed.

    Returns:
        Status of Asta window and GUI library availability
    """
    gui = AstaGUIManager()

    # Check libraries
    libs_ok, libs_msg = gui._check_gui_libs()

    # Check window
    window_info = gui.find_asta_window()

    result = {
        "gui_libraries_installed": libs_ok,
        "gui_libraries_message": libs_msg,
        "asta_window": window_info,
    }

    if libs_ok and window_info.get("found"):
        result["status"] = "READY - Asta is running and GUI tools are available"
    elif not libs_ok:
        result["status"] = f"GUI LIBRARIES MISSING - {libs_msg}"
    else:
        result["status"] = "ASTA NOT FOUND - Please open Asta Powerproject first"

    return json.dumps(result, indent=2, default=str)


@mcp.tool(
    name="asta_gui_bring_to_front",
    annotations={
        "title": "Bring Asta to Front",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_bring_to_front() -> str:
    """Bring the Asta Powerproject window to the foreground.

    Use this before performing any GUI actions to ensure Asta is visible
    and has keyboard/mouse focus.

    Returns:
        Success or failure message
    """
    result = AstaGUIManager.bring_to_front()
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_send_shortcut",
    annotations={
        "title": "Send Keyboard Shortcut",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_send_shortcut(params: GUIShortcutInput) -> str:
    """Send a keyboard shortcut to Asta Powerproject.

    Common Asta shortcuts:
    - 'ctrl+s' = Save project
    - 'F9' = Reschedule (calculate critical path)
    - 'ctrl+z' = Undo
    - 'ctrl+y' = Redo
    - 'ctrl+p' = Print
    - 'insert' = Insert new bar/row
    - 'delete' = Delete selected item
    - 'F1' = Help

    Args:
        params: Contains the shortcut string (e.g., 'ctrl+s')

    Returns:
        Confirmation that shortcut was sent
    """
    result = AstaGUIManager.send_shortcut(params.shortcut)
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_reschedule",
    annotations={
        "title": "Reschedule Project",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_reschedule() -> str:
    """Run the Reschedule command in Asta Powerproject (F9).

    Rescheduling:
    1. Calculates the optimal start/end dates for all tasks
    2. Determines the Critical Path (shown in red)
    3. Calculates Total Float and Free Float
    4. Identifies constraint violations
    5. Finds the earliest project finish date

    IMPORTANT: Always reschedule after making changes to tasks or links.

    Returns:
        Confirmation that reschedule was triggered
    """
    # First bring Asta to front
    AstaGUIManager.bring_to_front()
    time.sleep(0.5)
    result = AstaGUIManager.send_shortcut("F9", delay=1.0)
    result["action"] = "Reschedule (F9)"
    result["note"] = "Check the Asta window for reschedule results. Critical path will be shown in red."
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_save",
    annotations={
        "title": "Save Project (GUI)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_save() -> str:
    """Save the current project in Asta Powerproject (Ctrl+S).

    Returns:
        Confirmation that save command was sent
    """
    result = AstaGUIManager.send_shortcut("ctrl+s")
    result["action"] = "Save (Ctrl+S)"
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_undo",
    annotations={
        "title": "Undo Last Action",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_undo() -> str:
    """Undo the last action in Asta Powerproject (Ctrl+Z).

    Returns:
        Confirmation that undo was triggered
    """
    result = AstaGUIManager.send_shortcut("ctrl+z")
    result["action"] = "Undo (Ctrl+Z)"
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_click",
    annotations={
        "title": "Click at Position",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_click(params: GUIClickInput) -> str:
    """Click at specific screen coordinates in Asta Powerproject.

    Use take_screenshot first to identify the correct coordinates.
    The Asta window will be brought to the foreground automatically.

    Args:
        params: Contains x, y coordinates, click count, and button type

    Returns:
        Confirmation of click action
    """
    result = AstaGUIManager.click_at(params.x, params.y, params.clicks, params.button)
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_type_text",
    annotations={
        "title": "Type Text in Asta",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_type_text(params: GUITypeInput) -> str:
    """Type text into the currently focused cell or field in Asta.

    Click on a cell first, then use this tool to enter text.
    Can optionally press Enter after typing to confirm the entry.

    Args:
        params: Contains text to type and whether to press Enter

    Returns:
        Confirmation of typed text
    """
    try:
        import pyautogui
        AstaGUIManager.bring_to_front()
        time.sleep(0.3)
        _clipboard_paste(params.text)
        if params.press_enter:
            time.sleep(0.2)
            pyautogui.press('enter')
        return json.dumps({
            "success": True,
            "typed": params.text,
            "enter_pressed": params.press_enter
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_screenshot",
    annotations={
        "title": "Take Screenshot",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_screenshot(params: GUIScreenshotInput) -> str:
    """Take a screenshot of the current Asta Powerproject screen.

    Useful for verifying the current state before and after operations.
    The screenshot is saved as a PNG file.

    Args:
        params: Optional save_path for the screenshot file

    Returns:
        Path to the saved screenshot file
    """
    result = AstaGUIManager.take_screenshot(params.save_path)
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_open_file",
    annotations={
        "title": "Open Project File in Asta",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_open_file(params: GUIOpenFileInput) -> str:
    """Open a project file in Asta Powerproject using File > Open dialog.

    This uses keyboard shortcuts to navigate the Open dialog:
    1. Sends Ctrl+O to open the file dialog
    2. Types the file path
    3. Presses Enter to open

    Args:
        params: Contains the file path to open

    Returns:
        Confirmation of open command
    """
    try:
        import pyautogui
        gui = AstaGUIManager()
        gui.bring_to_front()
        time.sleep(0.5)

        # Send Ctrl+O
        pyautogui.hotkey('ctrl', 'o')
        time.sleep(1.5)

        # Type the file path
        file_path = params.file_path.replace("/", "\\")
        _clipboard_paste(file_path)
        time.sleep(0.5)

        # Press Enter
        pyautogui.press('enter')
        time.sleep(2.0)

        return json.dumps({
            "success": True,
            "file": params.file_path,
            "message": f"Open command sent for: {params.file_path}. Check Asta window for result."
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_new_project",
    annotations={
        "title": "Create New Project (GUI)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_new_project(params: GUINewProjectInput) -> str:
    """Create a new project in Asta Powerproject using the GUI.

    Steps performed:
    1. Opens File > New dialog
    2. Selects the specified template
    3. Enters project name, client, contractor, and start date
    4. Clicks Create

    Args:
        params: Contains project_name, client_name, contractor_name, start_date, template

    Returns:
        Confirmation with instructions for manual verification
    """
    try:
        import pyautogui
        gui = AstaGUIManager()
        gui.bring_to_front()
        time.sleep(0.5)

        # Open New Project dialog: Ctrl+N
        pyautogui.hotkey('ctrl', 'n')
        time.sleep(2.0)

        return json.dumps({
            "success": True,
            "message": "New Project dialog opened. Please complete the following steps in Asta:",
            "steps": [
                f"1. Select template: '{params.template}'",
                f"2. Enter project name: '{params.project_name}'",
                f"3. Enter client (For): '{params.client_name or 'N/A'}'",
                f"4. Enter contractor (By): '{params.contractor_name or 'N/A'}'",
                f"5. Set start date: '{params.start_date or 'Today'}'",
                "6. Click 'Create'"
            ],
            "note": "Due to the complexity of the New Project dialog, some fields may need manual input. Use asta_gui_click and asta_gui_type_text for precise control."
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_take_baseline",
    annotations={
        "title": "Take Baseline Snapshot",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_take_baseline(params: GUIBaselineInput) -> str:
    """Take a baseline snapshot of the current project schedule.

    A baseline captures the current plan so you can compare actual
    progress against it later. This navigates to Project tab > Take Baseline.

    Args:
        params: Contains baseline_name

    Returns:
        Instructions for completing the baseline operation
    """
    try:
        import pyautogui
        gui = AstaGUIManager()
        gui.bring_to_front()
        time.sleep(0.5)

        return json.dumps({
            "success": True,
            "message": "To take a baseline in Asta:",
            "steps": [
                "1. Click the 'Project' tab in the ribbon",
                "2. Click 'Take Baseline' button",
                f"3. Enter baseline name: '{params.baseline_name}'",
                "4. Click OK",
            ],
            "note": "Use asta_gui_click to click specific UI elements, or follow these steps manually.",
            "keyboard_alternative": "You can also use the menu: Project > Take Baseline"
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_insert_row",
    annotations={
        "title": "Insert New Row",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_insert_row() -> str:
    """Insert a new empty row/bar at the current position in Asta.

    Sends the Insert key to create a new row in the spreadsheet.
    Click on the desired position first before using this tool.

    Returns:
        Confirmation that Insert key was pressed
    """
    result = AstaGUIManager.send_shortcut("insert")
    result["action"] = "Insert new row"
    result["note"] = "A new empty row should appear. Type the task name in the Name column."
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_delete_selected",
    annotations={
        "title": "Delete Selected Item",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_delete_selected() -> str:
    """Delete the currently selected task, link, or item in Asta.

    WARNING: This sends the Delete key. Make sure the correct item
    is selected before using this tool. Use Ctrl+Z to undo if needed.

    Returns:
        Confirmation that Delete key was pressed
    """
    result = AstaGUIManager.send_shortcut("delete")
    result["action"] = "Delete selected item"
    result["warning"] = "Use Ctrl+Z (asta_gui_undo) immediately if wrong item was deleted"
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_link_tasks",
    annotations={
        "title": "Link Tasks (GUI)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_link_tasks(params: GUILinkTasksInput) -> str:
    """Link two tasks together in Asta using keyboard block-linking.

    This selects the predecessor and successor tasks and uses
    Home tab > Link Tasks to create a Finish-to-Start link.

    For other link types (SS, FF, SF) or lag, you'll need to
    edit the link properties after creation.

    Args:
        params: Contains predecessor_row, successor_row, link_type, lag

    Returns:
        Instructions for linking the tasks
    """
    return json.dumps({
        "success": True,
        "message": "To link tasks in Asta:",
        "method_1_mouse": [
            f"1. Hover at the END of task in row {params.predecessor_row} (cursor becomes link icon)",
            f"2. Click and drag to the START of task in row {params.successor_row}",
            "3. Release to create a Finish-to-Start link",
            f"4. If lag is needed ({params.lag or 'none'}): Hold Shift while dragging to add lag time",
        ],
        "method_2_keyboard": [
            f"1. Click on row {params.predecessor_row} to select it",
            f"2. Hold Ctrl and click on row {params.successor_row} to add to selection",
            "3. Go to Home tab > Link Tasks button",
            "4. This creates FS links between selected tasks in order",
        ],
        "link_type": params.link_type,
        "lag": params.lag or "None",
        "note": f"To change link type to {params.link_type}: Double-click the link line > change Type in properties"
    }, indent=2)


@mcp.tool(
    name="asta_gui_apply_filter",
    annotations={
        "title": "Apply Filter",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_apply_filter(params: GUIFilterInput) -> str:
    """Apply a display filter in Asta Powerproject.

    Filters control which tasks are visible on the bar chart:
    - 'critical': Show only critical path tasks
    - 'complete': Show only completed tasks
    - 'incomplete': Show only tasks not yet completed
    - 'code': Filter by code library
    - 'none': Remove all filters (show everything)

    Args:
        params: Contains filter_type and optional code_name

    Returns:
        Instructions for applying the filter
    """
    filter_instructions = {
        "critical": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Critical Tasks' or similar filter",
        ],
        "complete": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Complete Tasks'",
        ],
        "incomplete": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Incomplete Tasks'",
        ],
        "code": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Codes' > 'Which Code'",
            f"4. Select code library: '{params.code_name or 'Select library'}'",
            "5. Check the desired code entries",
            "6. Click Finish",
        ],
        "none": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'No Filter'",
        ],
    }

    steps = filter_instructions.get(params.filter_type, ["Unknown filter type"])

    return json.dumps({
        "success": True,
        "filter_type": params.filter_type,
        "steps": steps,
        "note": "Use asta_gui_click to click specific buttons, or follow these steps manually"
    }, indent=2)


@mcp.tool(
    name="asta_gui_change_table",
    annotations={
        "title": "Change Spreadsheet Table",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_change_table(table_name: str = "Default") -> str:
    """Change the spreadsheet table view in Asta.

    Common table views:
    - 'Default': Standard task information
    - 'Progress - no baseline': For entering progress data
    - 'Progress - with baseline': For comparing against baseline
    - 'Resource': Resource-related columns
    - 'Cost': Cost-related columns

    Args:
        table_name: Name of the table view to switch to

    Returns:
        Instructions for changing the table
    """
    return json.dumps({
        "success": True,
        "table": table_name,
        "steps": [
            "1. Go to View tab",
            "2. Click 'Table' dropdown",
            f"3. Select '{table_name}'",
        ],
        "note": "This changes which columns are visible in the spreadsheet area"
    }, indent=2)


@mcp.tool(
    name="asta_gui_print_export",
    annotations={
        "title": "Print/Export Project",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_print_export(params: GUIPrintInput) -> str:
    """Print or export the current Asta view to PDF, picture, or printer.

    The output shows exactly what's currently visible on screen -
    so set up your view (filter, zoom, table) before printing.

    Args:
        params: Contains output_type ('pdf', 'clipboard', 'picture', 'printer')
                and optional save_path

    Returns:
        Instructions for completing the print/export
    """
    try:
        import pyautogui
        AstaGUIManager.bring_to_front()
        time.sleep(0.5)

        # Open print dialog
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(1.5)

        instructions = {
            "pdf": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Select a PDF printer (e.g., 'Microsoft Print to PDF')",
                "3. Set orientation to Landscape (recommended for Gantt charts)",
                f"4. Click Print and save to: {params.save_path or 'choose location'}",
            ],
            "clipboard": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Click the 'Clipboard' option",
                "3. Click Print",
                "4. The bar chart is now in your clipboard - paste into Word/Excel",
            ],
            "picture": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Click the 'Picture file' option",
                "3. Browse to select save location",
                f"4. Save as: {params.save_path or 'choose location'}",
            ],
            "printer": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Select your printer",
                "3. Set paper size and orientation",
                "4. Click Print",
            ],
        }

        return json.dumps({
            "success": True,
            "output_type": params.output_type,
            "steps": instructions.get(params.output_type, ["Unknown output type"]),
            "tip": "Set up your view (filter, zoom, table) BEFORE printing - what you see is what you get!"
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_zoom",
    annotations={
        "title": "Zoom Bar Chart",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_zoom(direction: str = "fit") -> str:
    """Zoom the bar chart view in Asta Powerproject.

    Args:
        direction: 'in' to zoom in, 'out' to zoom out, 'fit' to fit all tasks

    Returns:
        Instructions for zooming
    """
    zoom_instructions = {
        "in": "Use Ctrl+Mouse Wheel Up on the Date Zone area to zoom in (show more detail)",
        "out": "Use Ctrl+Mouse Wheel Down on the Date Zone area to zoom out (show more time)",
        "fit": "Go to View tab > click 'Zoom to Fit' to show all tasks in the visible area",
    }

    return json.dumps({
        "success": True,
        "direction": direction,
        "instruction": zoom_instructions.get(direction, "Unknown direction"),
        "tip": "You can also right-click the Date Zone for more zoom options"
    }, indent=2)


@mcp.tool(
    name="asta_gui_summarize_tasks",
    annotations={
        "title": "Create Summary Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_summarize_tasks(summary_name: str = "New Group") -> str:
    """Create a summary task to group selected tasks in Asta.

    Summary tasks group related activities (like a folder groups files).
    Select the tasks you want to group first, then use Home > Summarise.

    Args:
        summary_name: Name for the summary group (e.g., 'Foundation Work', 'Phase 1')

    Returns:
        Step-by-step instructions for creating the summary
    """
    return json.dumps({
        "success": True,
        "steps": [
            "1. Select the tasks you want to group:",
            "   - Click first task, then Shift+Click last task (for range)",
            "   - Or Ctrl+Click to select individual tasks",
            "2. Go to Home tab",
            "3. Click 'Summarise' button",
            "4. A summary bar appears above the selected tasks",
            f"5. Type the summary name: '{summary_name}'",
            "6. Press Enter to confirm",
        ],
        "tips": [
            "To add more tasks later: Select task > Home > Indent",
            "To remove task from group: Select task > Home > Outdent",
            "Double-click summary bar to collapse/expand the group",
        ]
    }, indent=2)


@mcp.tool(
    name="asta_gui_indent_task",
    annotations={
        "title": "Indent/Outdent Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_indent_task(direction: str = "indent") -> str:
    """Indent or outdent the selected task in the project hierarchy.

    - Indent: Makes the task a child of the task above (adds to summary group)
    - Outdent: Moves the task up one level (removes from summary group)

    Args:
        direction: 'indent' to make child, 'outdent' to move up

    Returns:
        Confirmation of indent/outdent action
    """
    return json.dumps({
        "success": True,
        "direction": direction,
        "steps": [
            "1. Select the task(s) you want to move",
            f"2. Go to Home tab > click '{direction.title()}' button",
            f"   (Arrow {'right' if direction == 'indent' else 'left'} icon)",
        ],
        "note": "Indent adds task to the summary above, Outdent removes it from its current group"
    }, indent=2)


# ============================================================================
# UTILITY / HELP TOOLS
# ============================================================================

@mcp.tool(
    name="asta_help",
    annotations={
        "title": "Asta Help Guide",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_help(topic: str = "overview") -> str:
    """Get help and guidance on Asta Powerproject topics.

    Available topics:
    - 'overview': General overview of available tools
    - 'shortcuts': Keyboard shortcuts reference
    - 'workflow': Common workflow steps
    - 'links': How to create and manage task links
    - 'progress': How to track project progress
    - 'critical_path': Understanding the critical path
    - 'resources': Working with resources and costs
    - 'printing': Printing and exporting

    Args:
        topic: Help topic to display

    Returns:
        Detailed help text for the requested topic
    """
    help_topics = {
        "overview": """# Asta Powerproject MCP Server - Overview

## File-Based Tools (read/write project files):
- **asta_analyze_project**: Analyze a project file - START HERE
- **asta_list_tasks**: List all tasks with details
- **asta_get_task**: Get detailed info on one task
- **asta_add_task**: Add a new task
- **asta_update_task**: Update task properties
- **asta_delete_task**: Delete a task
- **asta_get_critical_path**: View critical path
- **asta_list_resources**: List all resources
- **asta_get_resource_assignments**: View resource allocations
- **asta_get_calendars**: View project calendars
- **asta_float_analysis**: Analyze float distribution
- **asta_save_project**: Save project to XML

## GUI Automation Tools (control Asta on screen):
- **asta_gui_check_status**: Check if Asta is running
- **asta_gui_reschedule**: Run reschedule (F9)
- **asta_gui_save**: Save project (Ctrl+S)
- **asta_gui_screenshot**: Take screenshot
- **asta_gui_click**: Click at coordinates
- **asta_gui_type_text**: Type text in cells
- **asta_gui_send_shortcut**: Send keyboard shortcuts
- **asta_gui_open_file**: Open a project file
- **asta_gui_new_project**: Create new project
- **asta_gui_take_baseline**: Take baseline snapshot
- **asta_gui_link_tasks**: Link tasks together
- **asta_gui_apply_filter**: Apply display filters
- **asta_gui_print_export**: Print/export to PDF
""",
        "shortcuts": """# Asta Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| F9 | Reschedule (calculate critical path) |
| Ctrl+S | Save project |
| Ctrl+P | Print |
| Ctrl+Z | Undo |
| Ctrl+Y | Redo |
| Ctrl+N | New project |
| Ctrl+O | Open project |
| Insert | Insert new row |
| Delete | Delete selected item |
| F1 | Help |
| Shift+Click | Select range |
| Ctrl+Click | Add to selection |
| Shift+Drag link | Add lag to link |
""",
        "workflow": """# Common Asta Workflow

## Creating a New Schedule:
1. File > New > Select template
2. Enter project details (name, dates, client)
3. Enter tasks in Spreadsheet (Name + Duration columns)
4. Link tasks (drag from end of one to start of another)
5. Press F9 to Reschedule
6. Review Critical Path (shown in red)
7. Take a Baseline (Project > Take Baseline)

## Updating Progress:
1. View > Table > Progress - no baseline
2. Enter % Complete for each active task
3. Enter Actual Start dates for started tasks
4. Enter Actual Finish dates for completed tasks
5. Press F9 to Reschedule with progress
6. Compare with baseline

## Weekly Routine:
1. Open project file
2. Enter progress for the past week
3. Reschedule (F9)
4. Check if critical path changed
5. Review float on near-critical tasks
6. Print/export updated schedule
7. Save and backup
""",
        "links": """# Task Links in Asta

## Link Types:
- **FS (Finish-to-Start)**: B starts after A finishes (most common)
- **SS (Start-to-Start)**: B starts when A starts
- **FF (Finish-to-Finish)**: B finishes when A finishes
- **SF (Start-to-Finish)**: B finishes when A starts (rare)

## Creating Links:
1. Hover at the END of the predecessor task
2. Cursor changes to link icon
3. Click and drag to the START of successor task
4. Release to create FS link

## Adding Lag (waiting time):
- Hold Shift while dragging to add lag
- Example: 2d lag for concrete curing time

## Block Linking (multiple tasks):
1. Select tasks (Shift or Ctrl click)
2. Home > Link Tasks
3. Creates FS links in sequence
""",
        "progress": """# Progress Tracking in Asta

## Setup:
1. Take a baseline BEFORE entering progress
2. Switch to Progress table: View > Table > Progress

## Entering Progress:
- **% Complete**: Enter in 'Overall Percent Complete' column
- **Actual Start**: Enter when task actually started
- **Actual Finish**: Enter when task actually finished
- **Planned %**: System calculates based on schedule

## After Entering Progress:
1. Press F9 (Reschedule)
2. Select 'Straighten progress entry period'
3. This moves incomplete work past the report date

## Reading Results:
- **Slip**: Difference between planned and actual
- **Progress line**: Visual indicator on bar chart
- Red tasks = Critical (monitor closely!)
""",
        "critical_path": """# Critical Path in Asta

## What is the Critical Path?
The longest chain of linked tasks through the project.
Any delay on critical tasks = project delay.

## How to See It:
1. Link all tasks properly
2. Press F9 (Reschedule)
3. Critical tasks get RED outline
4. Non-critical tasks show float (blue bars)

## Float Types:
- **Total Float**: Time task can slip without affecting project end
- **Free Float**: Time task can slip without affecting next task
- **Zero Float** = Critical task

## Tips:
- Keep critical tasks under close watch
- Look for 'near-critical' tasks (low float)
- Consider adding resources to shorten critical tasks
- Use Part Critical Shading for partially critical tasks
""",
        "resources": """# Resources in Asta

## Resource Types:
- **Permanent**: Labour, equipment (reusable)
- **Consumable**: Materials (used up)

## Creating Resources:
1. View > Library Explorer
2. Navigate to Resources folder
3. Right-click > New Resource
4. Set name, type, rate, availability

## Assigning Resources:
1. Open Project View (left panel)
2. Drag resource onto task
3. Set units/quantity

## Cost Centres:
1. Create Cost Centres (Labour, Materials, Plant)
2. Define rates per resource
3. Assign to tasks
4. View cost reports
""",
        "printing": """# Printing in Asta

## Before Printing:
- Set up view (filter, zoom, columns)
- What you see = what you print

## Print Options:
1. **Ctrl+P** to open print dialog
2. Choose output: Printer, PDF, Clipboard, Picture
3. Set paper size and orientation (Landscape recommended)
4. Select border file (company template)
5. Adjust scaling (fit to pages)
6. Preview before printing

## Border Files:
- Templates with company logo, revision info
- Select in Details tab > Browse
- Embed in project for portability

## Tips:
- Use 'Fit to 1 page wide' for clean output
- Save print profiles for reuse
- Landscape orientation works best for Gantt charts
""",
    }

    return help_topics.get(topic, f"Unknown topic: '{topic}'. Available: {', '.join(help_topics.keys())}")


# ============================================================================
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
# COM CRUD HELPERS — used by dual-strategy tools
# ---------------------------------------------------------------------------

def _find_bar_by_id(bars, bar_id: int):
    """Find a bar by its ID (not index). Handles the Asta COM
    Item(index) vs Item(id) ambiguity.

    Asta COM's Bars.Item() may use 1-based index OR ID depending on version.
    This function tries direct access first, then falls back to iterating.
    """
    # Strategy 1: Direct Item(id) — works if Asta maps by ID
    try:
        bar = bars.Item(bar_id)
        if bar is not None and bar.ID == bar_id:
            return bar
    except Exception:
        pass

    # Strategy 2: Iterate all bars to find matching ID
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

    return None


def _com_add_task(project, name: str, duration_str: str = "1d",
                  parent_bar_id: int = None, is_summary: bool = False) -> dict:
    """Add a task via COM. Returns result dict."""
    import pywintypes

    bars = project.Bars
    result = {"method": "COM"}

    # If parent specified, find it and add under it
    if parent_bar_id is not None:
        parent_bar = _find_bar_by_id(bars, parent_bar_id)
        if parent_bar is None:
            return {"error": f"Parent bar ID {parent_bar_id} not found"}

        # Add new bar after parent's last child
        try:
            new_bar = bars.AddAfter(parent_bar)
        except Exception:
            new_bar = bars.Add()
    else:
        new_bar = bars.Add()

    new_bar.Name = name

    if is_summary:
        # Convert to summary task
        try:
            task = new_bar.ExpandedTask
            if task is None:
                task = new_bar.Tasks.Item(1)
            task.ConvertToSummaryTask()
            result["type"] = "summary"
        except Exception:
            result["type_warning"] = "Could not convert to summary (may already be summary)"
    else:
        # Set duration
        if duration_str and duration_str != "0d":
            try:
                new_bar.EditToken("Duration", duration_str)
            except Exception:
                try:
                    task = new_bar.ExpandedTask
                    if task is None:
                        task = new_bar.Tasks.Item(1)
                    task.EditToken("Duration", duration_str)
                except Exception as de:
                    result["duration_warning"] = f"Could not set duration: {de}"

    result["task_id"] = new_bar.ID
    result["name"] = name
    result["duration"] = duration_str
    if parent_bar_id:
        result["parent_id"] = parent_bar_id

    return result


def _com_update_task(project, task_id: int, name: str = None,
                     duration_str: str = None, percent_complete: float = None,
                     notes: str = None) -> dict:
    """Update a task via COM. Returns result dict."""
    bars = project.Bars
    bar = _find_bar_by_id(bars, task_id)
    if bar is None:
        return {"error": f"Bar ID {task_id} not found"}

    result = {"method": "COM", "task_id": task_id, "updated_fields": []}

    if name is not None:
        bar.Name = name
        result["updated_fields"].append("name")
        result["name"] = name

    if duration_str is not None:
        try:
            bar.EditToken("Duration", duration_str)
            result["updated_fields"].append("duration")
            result["duration"] = duration_str
        except Exception:
            try:
                task = bar.ExpandedTask or bar.Tasks.Item(1)
                task.EditToken("Duration", duration_str)
                result["updated_fields"].append("duration")
                result["duration"] = duration_str
            except Exception as de:
                result["duration_error"] = str(de)

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
            task = bar.ExpandedTask
            if task is None:
                task = bar.Tasks.Item(1)
            task.AddTextDatedNote("Note", notes, datetime.now())
            result["updated_fields"].append("notes")
        except Exception as ne:
            # Fallback: try SetUDF
            try:
                bar.SetUDF("Notes", notes)
                result["updated_fields"].append("notes")
            except Exception:
                result["notes_error"] = str(ne)

    return result


def _com_delete_task(project, task_id: int) -> dict:
    """Delete a task/bar via COM. Returns result dict."""
    bars = project.Bars
    result = {"method": "COM", "task_id": task_id}

    # Find the index of the bar with this ID
    count = bars.Count
    found_idx = None
    bar_name = None
    for i in range(1, count + 1):
        try:
            bar = bars.Item(i)
            if bar.ID == task_id:
                found_idx = i
                bar_name = bar.Name
                break
        except Exception:
            continue

    if found_idx is None:
        # Try direct access by ID (some COM implementations support Item(id))
        try:
            bar = bars.Item(task_id)
            if bar is not None:
                bar_name = bar.Name
                bars.Remove(task_id, True)  # delete_children=True
                result["deleted"] = True
                result["name"] = bar_name
                return result
        except Exception:
            pass
        return {"error": f"Bar ID {task_id} not found"}

    bar_name_safe = bar_name or f"ID={task_id}"
    bars.Remove(found_idx, True)  # delete_children=True
    result["deleted"] = True
    result["name"] = bar_name_safe
    return result


def _com_add_link(project, predecessor_id: int, successor_id: int,
                  link_type: str = "FS", lag_str: str = None) -> dict:
    """Add a link between two bars via COM. Returns result dict."""
    import pywintypes

    result = {"method": "COM", "predecessor_id": predecessor_id,
              "successor_id": successor_id, "link_type": link_type}

    bars = project.Bars
    pred_bar = _find_bar_by_id(bars, predecessor_id)
    succ_bar = _find_bar_by_id(bars, successor_id)

    if pred_bar is None:
        return {"error": f"Predecessor bar ID {predecessor_id} not found"}
    if succ_bar is None:
        return {"error": f"Successor bar ID {successor_id} not found"}

    # Get tasks from bars
    pred_task = None
    succ_task = None
    try:
        pred_task = pred_bar.ExpandedTask
        if pred_task is None:
            pred_task = pred_bar.Tasks.Item(1)
    except Exception:
        return {"error": f"Cannot get task from predecessor bar {predecessor_id}"}

    try:
        succ_task = succ_bar.ExpandedTask
        if succ_task is None:
            succ_task = succ_bar.Tasks.Item(1)
    except Exception:
        return {"error": f"Cannot get task from successor bar {successor_id}"}

    # Link type constants (Asta COM)
    LINK_TYPES = {"FS": 0, "SS": 1, "FF": 2, "SF": 3}
    lt_val = LINK_TYPES.get(link_type.upper(), 0)

    # Create link via project.Links collection
    try:
        links = project.Links
        new_link = links.Add(pred_task, succ_task)

        # Set link type if not FS (default)
        if lt_val != 0:
            try:
                new_link.EditToken("LinkType", str(lt_val))
            except Exception:
                try:
                    new_link.Type = lt_val
                except Exception as lte:
                    result["link_type_warning"] = f"Could not set link type: {lte}"

        # Set lag
        if lag_str:
            try:
                new_link.EditToken("Lag", lag_str)
                result["lag"] = lag_str
            except Exception:
                try:
                    new_link.EditToken("lag", lag_str)
                    result["lag"] = lag_str
                except Exception as le:
                    result["lag_warning"] = f"Could not set lag: {le}"

        result["success"] = True
        result["link_id"] = getattr(new_link, 'ID', None)

    except Exception as e:
        # Fallback: try via bar token
        try:
            succ_bar.EditToken("Predecessors", f"{predecessor_id}{link_type.upper()}")
            result["success"] = True
            result["note"] = "Link added via bar token fallback"
        except Exception as e2:
            result["error"] = f"Primary: {e}, Fallback: {e2}"

    return result


def _com_remove_link(project, predecessor_id: int, successor_id: int) -> dict:
    """Remove a link between two bars via COM. Returns result dict."""
    result = {"method": "COM", "predecessor_id": predecessor_id,
              "successor_id": successor_id}

    try:
        links = project.Links
        count = links.Count
        for i in range(1, count + 1):
            try:
                link = links.Item(i)
                pred_bar_id = None
                succ_bar_id = None
                try:
                    pred_bar_id = link.PredecessorTask.Bar.ID
                except Exception:
                    try:
                        pred_bar_id = link.GetToken("PredecessorBarID")
                    except Exception:
                        continue
                try:
                    succ_bar_id = link.SuccessorTask.Bar.ID
                except Exception:
                    try:
                        succ_bar_id = link.GetToken("SuccessorBarID")
                    except Exception:
                        continue

                if int(pred_bar_id) == predecessor_id and int(succ_bar_id) == successor_id:
                    links.Remove(i)
                    result["removed"] = True
                    return result
            except Exception:
                continue

        result["error"] = f"Link from {predecessor_id} to {successor_id} not found"
    except Exception as e:
        result["error"] = f"Could not access links: {e}"

    return result


def _com_update_link(project, predecessor_id: int, successor_id: int,
                     new_link_type: str = None, new_lag: str = None) -> dict:
    """Update a link between two bars via COM. Returns result dict."""
    result = {"method": "COM", "predecessor_id": predecessor_id,
              "successor_id": successor_id}

    LINK_TYPES = {"FS": 0, "SS": 1, "FF": 2, "SF": 3}

    try:
        links = project.Links
        count = links.Count
        for i in range(1, count + 1):
            try:
                link = links.Item(i)
                pred_bar_id = None
                succ_bar_id = None
                try:
                    pred_bar_id = link.PredecessorTask.Bar.ID
                except Exception:
                    continue
                try:
                    succ_bar_id = link.SuccessorTask.Bar.ID
                except Exception:
                    continue

                if int(pred_bar_id) == predecessor_id and int(succ_bar_id) == successor_id:
                    # Found the link — update it
                    if new_link_type:
                        lt_val = LINK_TYPES.get(new_link_type.upper(), 0)
                        try:
                            link.EditToken("LinkType", str(lt_val))
                            result["new_link_type"] = new_link_type
                        except Exception:
                            try:
                                link.Type = lt_val
                                result["new_link_type"] = new_link_type
                            except Exception as lte:
                                result["link_type_error"] = str(lte)

                    if new_lag:
                        try:
                            link.EditToken("Lag", new_lag)
                            result["new_lag"] = new_lag
                        except Exception as le:
                            result["lag_error"] = str(le)

                    result["updated"] = True
                    return result
            except Exception:
                continue

        result["error"] = f"Link from {predecessor_id} to {successor_id} not found"
    except Exception as e:
        result["error"] = f"Could not access links: {e}"

    return result


def _com_update_progress(project, task_id: int, percent_complete: float = None,
                         actual_start: str = None, actual_finish: str = None) -> dict:
    """Update progress on a task via COM. Returns result dict."""
    import pywintypes

    bars = project.Bars
    bar = _find_bar_by_id(bars, task_id)
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
            bar.ActualEnd = ole_date
            result["updated"].append("actual_finish")
            result["actual_finish"] = actual_finish
        except Exception:
            # Try ActualFinish property
            try:
                bar.ActualFinish = ole_date
                result["updated"].append("actual_finish")
                result["actual_finish"] = actual_finish
            except Exception as afe:
                result["actual_finish_error"] = str(afe)

    return result


def _gui_reschedule_fallback(report_date_str: str, straighten: bool) -> dict:
    """Fallback: perform reschedule via GUI automation (pyautogui + pywinauto).

    Steps:
      1. Bring Asta to foreground
      2. Navigate to progress period settings to set the report date
      3. Send F9 to trigger reschedule

    Returns a dict with results.
    """
    result = {
        "method": "GUI Automation (Fallback)",
        "success": False,
        "report_date_set": report_date_str,
    }

    # Step 1: Check Asta is running and bring to front
    try:
        AstaGUIManager.bring_to_front()
        time.sleep(0.5)
    except Exception as e:
        result["error"] = f"Cannot find Asta Powerproject window: {e}"
        result["suggestion"] = "Please open Asta Powerproject and load a project first."
        return result

    try:
        import pyautogui
    except ImportError:
        result["error"] = "pyautogui is not installed. Cannot perform GUI automation."
        return result

    # Step 2: Set the report date via progress period dialog
    # Navigate: Project tab > Progress Periods, or use the date in the toolbar
    # For robustness, we use keyboard navigation to the Progress Period dialog
    try:
        # Open the progress period entry: Alt+P opens Project tab area
        # In Asta, the progress period can be set via:
        #   Project tab > Progress > Set Progress Period
        # However, ribbon navigation varies by version.
        # Most reliable approach: use the existing report date field
        # if visible in the toolbar, or navigate via menus.

        # Try the ribbon path: Alt > P (Project) > then look for Progress Period
        # Asta ribbon: Home | View | Project | Format | Tools
        # Under Project tab: Progress section has "Progress Period" button

        # Navigate to Project tab
        pyautogui.hotkey('alt')
        time.sleep(0.3)
        pyautogui.press('p')  # Project tab (may vary by locale)
        time.sleep(0.5)

        # Look for Progress Period button - this opens the dialog
        # The exact key depends on the ribbon layout
        # Typically: "Progress Period" or "PP" accelerator
        pyautogui.hotkey('alt', 'p')  # Try alt+p again for the progress period group
        time.sleep(0.3)

        # Since ribbon navigation is fragile, log what we're doing
        logger.info(f"GUI Reschedule: Attempting to set report date to {report_date_str}")

        # Type the date into the progress period dialog if it opened
        # Format: convert YYYY-MM-DD to DD/MM/YYYY (Asta's typical format)
        dt = datetime.strptime(report_date_str, "%Y-%m-%d")
        asta_date = dt.strftime("%d/%m/%Y")

        # If dialog is open, type the date
        time.sleep(0.5)
        _clipboard_paste(asta_date)
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(0.5)

        # Press Escape to close any remaining dialogs
        pyautogui.press('escape')
        time.sleep(0.3)

        result["report_date_entry"] = "Attempted via ribbon navigation"

    except Exception as e:
        logger.warning(f"GUI progress period setting failed: {e}")
        result["report_date_entry"] = f"Failed: {e}"
        # Continue anyway to at least trigger reschedule

    # Step 3: Trigger reschedule with F9
    try:
        AstaGUIManager.bring_to_front()
        time.sleep(0.3)
        shortcut_result = AstaGUIManager.send_shortcut("F9", delay=2.0)
        result["reschedule_triggered"] = True
        result["success"] = True
        result["note"] = (
            "Reschedule (F9) was triggered via GUI automation. "
            "The report date may need to be verified manually in the Progress Period dialog. "
            "Check the Asta window for results - critical path tasks will be shown in red."
        )
    except Exception as e:
        result["error"] = f"Failed to trigger F9 reschedule: {e}"

    return result


@mcp.tool(
    name="asta_reschedule_project",
    annotations={
        "title": "Reschedule Project (COM)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
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
                        project.EndTransaction()
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


@mcp.tool()
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
                            task = bar.ExpandedTask
                            if task is None:
                                task = bar.Tasks.Item(1)
                            # Duration is IRelativeTime — set via token
                            task.EditToken("Duration", mod.new_duration)
                            mod_result["duration_set"] = mod.new_duration
                        except Exception as de:
                            mod_result["duration_error"] = str(de)

                # Change name
                if mod.new_name is not None:
                    bar.Name = mod.new_name
                    mod_result["name_set"] = mod.new_name

                # Add predecessor link
                if mod.add_predecessor_id is not None:
                    try:
                        pred_bar = _find_bar_by_id(bars_collection, mod.add_predecessor_id)
                        if pred_bar is not None:
                            pred_task = pred_bar.ExpandedTask
                            if pred_task is None:
                                pred_task = pred_bar.Tasks.Item(1)
                            cur_task = bar.ExpandedTask
                            if cur_task is None:
                                cur_task = bar.Tasks.Item(1)
                            # Create FS link via project.Links or task method
                            try:
                                project.Links.Add(pred_task, cur_task)
                                mod_result["predecessor_added"] = mod.add_predecessor_id
                            except Exception:
                                # Fallback: use bar-level token
                                bar.EditToken("Predecessors",
                                              f"{mod.add_predecessor_id}FS")
                                mod_result["predecessor_added"] = mod.add_predecessor_id
                    except Exception as le:
                        mod_result["predecessor_error"] = str(le)

                # Remove predecessor link
                if mod.remove_predecessor_id is not None:
                    try:
                        links = project.Links
                        link_count = links.Count
                        removed = False
                        for li in range(1, link_count + 1):
                            try:
                                link = links.Item(li)
                                pred_id = link.PredecessorTask.Bar.ID
                                succ_id = link.SuccessorTask.Bar.ID
                                if pred_id == mod.remove_predecessor_id and succ_id == mod.task_id:
                                    links.Remove(li)
                                    mod_result["predecessor_removed"] = mod.remove_predecessor_id
                                    removed = True
                                    break
                            except Exception:
                                continue
                        if not removed:
                            mod_result["predecessor_remove_warning"] = (
                                f"Link from {mod.remove_predecessor_id} to {mod.task_id} not found"
                            )
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
                project.EndTransaction()
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


@mcp.tool()
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

        # Access CodeLibraries with fallback strategies
        code_libs = None
        try:
            code_libs = project.CodeLibraries
        except AttributeError:
            pass
        if code_libs is None:
            # Fallback: try via GetToken
            try:
                code_libs = project.GetToken("CodeLibraries")
            except Exception:
                pass
        if code_libs is None:
            # Fallback: try via app
            try:
                code_libs = app.ActiveProject.CodeLibraries
            except Exception:
                pass
        if code_libs is None:
            result["error"] = (
                "CodeLibraries is not accessible in this Asta version (18.x). "
                "This feature requires Asta Developer Toolkit (astadkit.ocx) to be registered, "
                "or a newer version of Asta Powerproject with full COM API support. "
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
                        project.EndTransaction()
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
                project.EndTransaction()
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
                        project.EndTransaction()
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


@mcp.tool()
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
        code_libs = project.CodeLibraries
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
                bar = _find_bar_by_id(bars, int(task_id))
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
            project.EndTransaction()
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
        description="Action: 'list', 'create_permanent', 'create_consumable', "
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
        allowed = {"list", "create_permanent", "create_consumable",
                    "create_cost_centre", "delete_resource", "delete_cost_centre"}
        if v.lower() not in allowed:
            raise ValueError(f"action must be one of {allowed}")
        return v.lower()


@mcp.tool()
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
                        cons_list.append({
                            "id": r.ID,
                            "name": r.Name,
                            "availability": safe_float(getattr(r, 'Availability', None)),
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
                        cc_list.append({
                            "id": cc.ID,
                            "name": cc.Name,
                        })
                    except Exception:
                        continue
            except Exception as e:
                result["cost_centres_error"] = str(e)

            result["permanent_resources"] = perm_list
            result["consumable_resources"] = cons_list
            result["cost_centres"] = cc_list
            result["success"] = True

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
                                project.EndTransaction()
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
                        # Cost is IAmountAndCurrency — try setting via token
                        new_res.EditToken("Cost", str(params.cost_rate))
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
                    project.EndTransaction()
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
                                project.EndTransaction()
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

                try:
                    project.EndTransaction()
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
                                project.EndTransaction()
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
                    project.EndTransaction()
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
                    project.EndTransaction()
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
                    project.EndTransaction()
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
                    "'work_profile': str (opt: 'linear', 'front_loaded', 'back_loaded', 'bell_curve')}",
        min_length=1,
        max_length=100,
    )


@mcp.tool()
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
                bar = _find_bar_by_id(bars, int(task_id))
                if bar is None:
                    a_result["error"] = f"Bar ID {task_id} not found (tried index and ID search)"
                    assign_errors.append(a_result)
                    continue
                # Get the task from the bar
                task = None
                try:
                    task = bar.ExpandedTask
                except Exception:
                    pass
                if task is None:
                    try:
                        task = bar.Tasks.Item(1)
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

                elif res_type == "cost_centre":
                    cc_obj = cc_map.get(res_name.lower())
                    if cc_obj is None:
                        a_result["error"] = f"Cost centre '{res_name}' not found"
                        assign_errors.append(a_result)
                        continue
                    allocation = task.AssignCost(cc_obj)

                else:
                    a_result["error"] = f"Invalid resource_type: '{res_type}'"
                    assign_errors.append(a_result)
                    continue

                a_result["assigned"] = True

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
            project.EndTransaction()
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
                    "'set_sorting', 'set_filter', 'toggle_histogram', 'show_hierarchy_level'."
    )
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
                    "set_filter", "toggle_histogram", "show_hierarchy_level"}
        if v.lower() not in allowed:
            raise ValueError(f"action must be one of {allowed}")
        return v.lower()


@mcp.tool()
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
                code_libs = project.CodeLibraries
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


@mcp.tool()
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


@mcp.tool()
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


@mcp.tool()
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
        task_overview = {"total_bars": 0, "tasks": []}
        try:
            bars = project.Bars
            bar_count = bars.Count
            task_overview["total_bars"] = bar_count

            for i in range(1, min(bar_count + 1, params.max_tasks + 1)):
                try:
                    bar = bars.Item(i)
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
                    task_overview["tasks"].append(task_info)
                except Exception:
                    continue

            if bar_count > params.max_tasks:
                task_overview["note"] = f"Showing {params.max_tasks} of {bar_count} bars"
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

        # --- Variance Analysis ---
        if params.include_variances:
            variances = []
            try:
                bars = project.Bars
                for i in range(1, min(bars.Count + 1, params.max_tasks + 1)):
                    try:
                        bar = bars.Item(i)
                        v_info = {"id": bar.ID, "name": bar.Name}

                        # Check for actual dates vs planned
                        try:
                            planned_start = bar.Start
                            actual_start = bar.ActualStart
                            if actual_start is not None and planned_start is not None:
                                if hasattr(actual_start, 'year') and hasattr(planned_start, 'year'):
                                    start_delta = (actual_start - planned_start).days
                                    v_info["start_variance_days"] = start_delta
                        except Exception:
                            pass

                        try:
                            planned_end = bar.End
                            actual_end = bar.ActualEnd
                            if actual_end is not None and planned_end is not None:
                                if hasattr(actual_end, 'year') and hasattr(planned_end, 'year'):
                                    end_delta = (actual_end - planned_end).days
                                    v_info["end_variance_days"] = end_delta
                        except Exception:
                            pass

                        # Original vs current duration
                        try:
                            orig_start = bar.OriginalStartV
                            orig_end = bar.OriginalFinishV
                            if orig_start and orig_end:
                                v_info["original_start"] = format_date(orig_start)
                                v_info["original_finish"] = format_date(orig_end)
                        except Exception:
                            pass

                        if len(v_info) > 2:  # Has variance data beyond id/name
                            variances.append(v_info)
                    except Exception:
                        continue
            except Exception as e:
                variances = [{"error": str(e)}]

            report["variances"] = variances

        # --- Critical Path ---
        if params.include_critical_path:
            critical = []
            try:
                bars = project.Bars
                for i in range(1, bars.Count + 1):
                    try:
                        bar = bars.Item(i)
                        # Check if bar is on critical path via token or property
                        is_critical = False
                        try:
                            is_critical = bool(bar.GetToken("Critical"))
                        except Exception:
                            try:
                                task = bar.ExpandedTask
                                if task is None:
                                    task = bar.Tasks.Item(1)
                                # Check total float = 0
                                tf = task.GetToken("TotalFloat")
                                if tf is not None and safe_float(tf) == 0:
                                    is_critical = True
                            except Exception:
                                pass

                        if is_critical:
                            critical.append({
                                "id": bar.ID,
                                "name": bar.Name,
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
        lines.append(f"## Tasks ({task_overview.get('total_bars', 0)} total)")
        if task_overview.get("note"):
            lines.append(f"*{task_overview['note']}*")
        lines.append("")
        lines.append("| ID | Name | Start | End | % Complete |")
        lines.append("|---|---|---|---|---|")
        for t in task_overview.get("tasks", []):
            pct = t.get("percent_complete", "N/A")
            lines.append(f"| {t['id']} | {t['name']} | {t['start']} | {t['end']} | {pct} |")
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
# ENTRY POINT
# ============================================================================
if __name__ == "__main__":
    logger.info("Starting Asta Powerproject MCP Server...")
    mcp.run()
