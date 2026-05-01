"""Phase 5b - DCMA 14-Point Schedule Health Assessment per CLAUDE.md RULE 10.

Pure-Python check functions. MSP/COM/file independent - takes plain dicts,
returns plain dicts. Easily testable without fixtures, without COM, without
MS Project.

Industry-standard thresholds (DCMA spec, NDIA EVMS, RULE 10):
- Logic (Rules 1-5): no_pred <5%, no_succ <5%, leads=0, lags <5%, fs_link >90%
- Constraints (Rule 6): hard_constraints <5%
- Float (Rules 7-8): high_float <5%, negative_float=0
- Duration (Rule 9): high_duration <5%
- Quality (Rules 10-11): invalid_dates=0, resources_missing <20%
- Schedule (Rules 12-14): missed_tasks <5%, critical_path >0, BEI >95%
"""
from typing import List, Dict, Any, Optional
import datetime as _dt


# ---------- Hardcoded thresholds (CLAUDE.md RULE 10) ----------

_DCMA_THRESHOLDS = {
    1: ("no_predecessor_pct", "<", 5.0),
    2: ("no_successor_pct", "<", 5.0),
    3: ("leads_count", "==", 0),
    4: ("lags_pct", "<", 5.0),
    5: ("fs_link_pct", ">", 90.0),
    6: ("hard_constraints_pct", "<", 5.0),
    7: ("high_float_pct", "<", 5.0),
    8: ("negative_float_count", "==", 0),
    9: ("high_duration_pct", "<", 5.0),
    10: ("invalid_dates_count", "==", 0),
    11: ("resources_missing_pct", "<", 20.0),
    12: ("missed_tasks_pct", "<", 5.0),
    13: ("critical_path_count", ">", 0),
    14: ("bei_pct", ">", 95.0),
}


# ---------- DCMA_RULES metadata (14 rules) ----------

DCMA_RULES = [
    {"id": 1, "name": "No Predecessor", "threshold_label": "<5%",
     "threshold_value": 5.0, "category": "Logic"},
    {"id": 2, "name": "No Successor", "threshold_label": "<5%",
     "threshold_value": 5.0, "category": "Logic"},
    {"id": 3, "name": "Leads", "threshold_label": "=0",
     "threshold_value": 0, "category": "Logic"},
    {"id": 4, "name": "Lags", "threshold_label": "<5%",
     "threshold_value": 5.0, "category": "Logic"},
    {"id": 5, "name": "FS Link %", "threshold_label": ">90%",
     "threshold_value": 90.0, "category": "Logic"},
    {"id": 6, "name": "Hard Constraints", "threshold_label": "<5%",
     "threshold_value": 5.0, "category": "Constraints"},
    {"id": 7, "name": "High Float (>44d)", "threshold_label": "<5%",
     "threshold_value": 5.0, "category": "Float"},
    {"id": 8, "name": "Negative Float", "threshold_label": "=0",
     "threshold_value": 0, "category": "Float"},
    {"id": 9, "name": "High Duration (>44d)", "threshold_label": "<5%",
     "threshold_value": 5.0, "category": "Duration"},
    {"id": 10, "name": "Invalid Dates", "threshold_label": "=0",
     "threshold_value": 0, "category": "Quality"},
    {"id": 11, "name": "Resources Missing", "threshold_label": "<20%",
     "threshold_value": 20.0, "category": "Quality"},
    {"id": 12, "name": "Missed Tasks", "threshold_label": "<5%",
     "threshold_value": 5.0, "category": "Schedule"},
    {"id": 13, "name": "Critical Path", "threshold_label": ">0",
     "threshold_value": 0, "category": "Schedule"},
    {"id": 14, "name": "BEI", "threshold_label": ">95%",
     "threshold_value": 95.0, "category": "Schedule"},
]


def _real_tasks(tasks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Filter out summary tasks (DCMA counts only 'real' work)."""
    return [t for t in tasks if not t.get("summary", False)]


def _eval_status(rule_id: int, actual: float) -> str:
    """Compare actual against threshold; return 'pass' or 'fail'."""
    field, op, threshold = _DCMA_THRESHOLDS[rule_id]
    if op == "<":
        return "pass" if actual < threshold else "fail"
    if op == ">":
        return "pass" if actual > threshold else "fail"
    if op == "==":
        return "pass" if actual == threshold else "fail"
    return "fail"


def check_no_predecessor(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 1: <5% of real tasks should have no predecessor.

    Returns {id, name, threshold, actual, actual_unit, status, failed_count,
             total_count, failed_task_ids}.
    """
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0:
        return {"id": 1, "name": "No Predecessor", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    failed_ids = [t["id"] for t in real if not (t.get("predecessors") or [])]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 1, "name": "No Predecessor", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(1, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }


def check_no_successor(tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 2: <5% of real tasks should have no successor."""
    real = _real_tasks(tasks)
    total = len(real)
    if total == 0:
        return {"id": 2, "name": "No Successor", "threshold": "<5%",
                "actual": 0.0, "actual_unit": "%", "status": "pass",
                "failed_count": 0, "total_count": 0, "failed_task_ids": []}
    failed_ids = [t["id"] for t in real if not (t.get("successors") or [])]
    failed_count = len(failed_ids)
    actual_pct = (failed_count / total) * 100.0
    return {
        "id": 2, "name": "No Successor", "threshold": "<5%",
        "actual": round(actual_pct, 2), "actual_unit": "%",
        "status": _eval_status(2, actual_pct),
        "failed_count": failed_count, "total_count": total,
        "failed_task_ids": failed_ids,
    }
