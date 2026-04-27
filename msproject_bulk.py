"""MSPDI XML bulk-write engine for MS Project.

Path 3 of hybrid speed strategy: bulk operations (>20 items) write to
MSPDI XML and trigger MS Project FileOpen import — much faster than COM
one-by-one (200 task in ~3-5 sec vs 60+ sec).

Usage:
    w = MsprojectBulkWriter(project_name="Villa Plan")
    uids = w.bulk_add_tasks([{"name": "Task 1", "duration": "5d"}, ...])
    w.bulk_add_links([{"pred_uid": uids[0], "succ_uid": uids[1], "type": "FS"}])
    w.save("output.xml")
    # Then: app.FileOpen("output.xml") in MS Project COM
"""
from __future__ import annotations
from datetime import datetime
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

        Each item: {name, duration (str like "5d"), [start, finish, summary,
        milestone, parent_uid, outline_level]}
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
        ET.SubElement(root, f"{{{ns}}}Assignments")

        # Pretty print
        ET.indent(root, space="  ")
        return ET.tostring(root, encoding="unicode")

    @staticmethod
    def _duration_to_iso(d: str) -> str:
        """Convert '5d' → 'PT40H0M0S' (assuming 8h/day)."""
        n = float(d.rstrip("dwhmDWHM") or "1")
        unit = (d[-1] if d and d[-1].isalpha() else "d").lower()
        hours = {"d": 8, "w": 40, "h": 1, "m": 1 / 60}.get(unit, 8) * n
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
        """FS=1, FF=0, SS=3, SF=2 per MSPDI spec."""
        return {"FF": "0", "FS": "1", "SF": "2", "SS": "3"}.get(t.upper(), "1")
