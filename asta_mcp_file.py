#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Asta Powerproject File MCP Server
===================================
A lightweight, file/MPXJ-only MCP server for Asta Powerproject.
Provides read-only analysis tools for project files using MPXJ (Java library).
Only 3 consolidated tools: asta_query, asta_resource, asta_calendar.

Author: Claude AI for Cahit
Version: 1.0.0
"""

import json
import os
import sys
import logging
import time
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
from enum import Enum

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
            os.path.join(os.path.expanduser("~"), "asta_mcp_file.log"),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger("asta_mcp_file")

# ============================================================================
# INITIALIZE MCP SERVER
# ============================================================================
mcp = FastMCP(
    "asta_powerproject_file",
    instructions=(
        "FILE-BASED read-only analysis tools for Asta Powerproject project files (.pp/.mpp/.xml). "
        "IMPORTANT: These tools are a SECONDARY fallback — only use them when the user explicitly "
        "asks to read/analyze a specific project FILE, or when asta_powerproject_mcp (COM) tools "
        "are unavailable. For live operations with a running Asta instance, ALWAYS prefer the "
        "asta_powerproject_mcp tools (asta_task, asta_link, asta_resource, asta_schedule, etc.) "
        "which connect directly via COM — no file path needed. "
        "These file tools require a file_path parameter pointing to a .pp/.mpp/.xml file."
    )
)

# ============================================================================
# CONSTANTS
# ============================================================================
SUPPORTED_EXTENSIONS = ['.pp', '.mpp', '.xml', '.mspdi', '.xer', '.pmxml']
MAX_RESPONSE_CHARS = 25000

def _truncate_response(text: str, max_chars: int = MAX_RESPONSE_CHARS) -> str:
    """Truncate response to prevent Claude Desktop context overflow."""
    if len(text) <= max_chars:
        return text
    truncated = text[:max_chars]
    last_newline = truncated.rfind('\n')
    if last_newline > max_chars * 0.8:
        truncated = truncated[:last_newline]
    return truncated + f"\n\n... **[TRUNCATED]** Response exceeded {max_chars} chars. Use smaller `limit` or `max_tasks` param, or query specific tasks with `asta_query → get_task`."

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
# COM AUTO-EXPORT HELPERS
# ============================================================================

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
    except Exception as e:
        errors.append(f"GetActiveObject(CLSID): {e}")

    # --- Strategy 2: Dispatch by CLSID (may launch or attach) ---
    try:
        app = win32com.client.dynamic.Dispatch(APP_CLSID)
        project = app.ActiveProject
        if project is None:
            raise RuntimeError("Connected via Dispatch(CLSID) but no project is open")
        return app, project, "Dispatch (CLSID)"
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
        "\n\nFallback: Provide a file_path to use MPXJ file-based reading instead."
    )


# ---------------------------------------------------------------------------
# COM AUTO-EXPORT -- for read-only tools when Asta is running
# ---------------------------------------------------------------------------

_com_auto_export_cache = {"path": None, "timestamp": 0}


def _com_auto_export() -> str:
    """When Asta is running and no file_path given, auto-export to temp XML.

    Uses a 30-second cache to avoid re-exporting on rapid successive queries.
    Returns the temp XML path.
    Raises RuntimeError if Asta is not running.
    """
    import pythoncom

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


# ============================================================================
# PYDANTIC INPUT MODELS
# ============================================================================

class ResponseFormat(str, Enum):
    MARKDOWN = "markdown"
    JSON = "json"


class ProjectFileInput(BaseModel):
    """Input for file operations. Provide file_path to an exported project file (.xml, .mpp, .pp).
    If Asta is running and file_path is omitted, auto-exports from COM."""
    model_config = ConfigDict(str_strip_whitespace=True)
    file_path: Optional[str] = Field(
        default=None,
        description="Path to project file (.pp, .mpp, .xml). Omit to auto-export from running Asta."
    )

    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v):
        if v is None or v.strip() == "":
            return None
        return v.strip().replace("\\", "/")


class AnalyzeProjectInput(ProjectFileInput):
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class ListTasksInput(ProjectFileInput):
    include_summary: bool = Field(default=True)
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)
    limit: int = Field(default=100, ge=1, le=500)


class GetTaskInput(ProjectFileInput):
    task_id: int = Field(..., ge=0)
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class CriticalPathInput(ProjectFileInput):
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class ResourcesInput(ProjectFileInput):
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class FloatAnalysisInput(ProjectFileInput):
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class WBSTreeInput(ProjectFileInput):
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)
    max_depth: int = Field(default=3, ge=1, le=99)


class DelayAnalysisInput(ProjectFileInput):
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class ResourceLoadingInput(ProjectFileInput):
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


# ============================================================================
# READ-ONLY TOOL FUNCTIONS
# ============================================================================

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
        mgr = AstaFileManager(_resolve_file_path(params.file_path))
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


async def asta_get_wbs_tree(params: WBSTreeInput) -> str:
    """Get the WBS (Work Breakdown Structure) hierarchy tree.

    Shows parent-child relationships between tasks, summary groups,
    outline levels, and WBS codes. Essential for understanding
    how the project is organized.

    Args:
        params: Contains file_path, response_format, max_depth

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


# ============================================================================
# 3 CONSOLIDATED TOOL DISPATCHERS
# ============================================================================

@mcp.tool(
    name="asta_query",
    annotations={"title": "File-Based Project Queries", "readOnlyHint": True}
)
async def asta_query(params: dict) -> str:
    """Query and analyze Asta Powerproject data from a FILE (read-only, MPXJ-based).

    SECONDARY TOOL: Only use when reading/analyzing a specific .pp/.mpp/.xml file.
    For live operations with running Asta, use asta_export → report or asta_task → get instead.

    Actions:
    - analyze: Full project analysis. Params: file_path, response_format
    - list_tasks: List tasks. Params: file_path, include_summary, response_format, limit
    - critical_path: Get critical path. Params: file_path, response_format
    - wbs: Get WBS hierarchy. Params: file_path, response_format, max_depth
    - float: Float/slack analysis. Params: file_path, response_format
    - delay: Delay analysis. Params: file_path, response_format
    - get_task: Get task details. Params: file_path, task_id, response_format

    All actions require a file_path to a .pp/.mpp/.xml project file.
    Default list_tasks limit is 100 (max 500).
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "analyze":
            result = await asta_analyze_project(AnalyzeProjectInput(**p))
        elif action == "list_tasks":
            result = await asta_list_tasks(ListTasksInput(**p))
        elif action == "critical_path":
            result = await asta_get_critical_path(CriticalPathInput(**p))
        elif action == "wbs":
            result = await asta_get_wbs_tree(WBSTreeInput(**p))
        elif action == "float":
            result = await asta_float_analysis(FloatAnalysisInput(**p))
        elif action == "delay":
            result = await asta_delay_analysis(DelayAnalysisInput(**p))
        elif action == "get_task":
            result = await asta_get_task(GetTaskInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: analyze, list_tasks, critical_path, wbs, float, delay, get_task"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_query({action}) failed: {e}"})


@mcp.tool(
    name="asta_file_resource",
    annotations={"title": "File-Based Resource Queries", "readOnlyHint": True}
)
async def asta_resource_query(params: dict) -> str:
    """Query resource data from Asta Powerproject FILE (read-only, MPXJ-based).

    SECONDARY TOOL: Only use when reading resource data from a specific .pp/.mpp/.xml file.
    For live resource management with running Asta, use asta_resource (COM) instead.

    Actions:
    - list: List resources. Params: file_path, response_format
    - assignments: Get resource assignments. Params: file_path, response_format
    - loading: Resource loading analysis. Params: file_path, response_format

    Use 'loading' for resource histograms and S-curve data. Use 'assignments' to check DCMA Check 10
    (all tasks should have resource assignments). Resource types: Permanent (labour/equipment),
    Consumable (materials), Cost Centres (budgets).
    """
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "list":
            result = await asta_list_resources(ResourcesInput(**p))
        elif action == "assignments":
            result = await asta_get_resource_assignments(ResourcesInput(**p))
        elif action == "loading":
            result = await asta_resource_loading(ResourceLoadingInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: list, assignments, loading"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_resource({action}) failed: {e}"})


@mcp.tool(
    name="asta_calendar",
    annotations={"title": "File-Based Calendar Queries", "readOnlyHint": True}
)
async def asta_calendar_query(params: dict) -> str:
    """Get project calendars from an Asta Powerproject FILE (read-only, MPXJ-based).

    SECONDARY TOOL: Only use when reading calendar data from a specific .pp/.mpp/.xml file.
    For live operations with running Asta, use asta_powerproject_mcp (COM) tools instead.

    Actions:
    - get: Get all calendars. Params: file_path

    Requires a file_path to a .pp/.mpp/.xml project file.
    Calendars define working days, hours, and exceptions (holidays, shutdowns).
    """
    action = params.get("action", "get")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "get":
            result = await asta_get_calendars(ProjectFileInput(**p))
        else:
            return json.dumps({"error": f"Unknown action '{action}'. Valid: get"})
        return _truncate_response(result)
    except Exception as e:
        return json.dumps({"error": f"asta_calendar({action}) failed: {e}"})


# ============================================================================
# ENTRY POINT
# ============================================================================
if __name__ == "__main__":
    logger.info("Starting Asta Powerproject File MCP Server...")
    mcp.run(transport="stdio")
