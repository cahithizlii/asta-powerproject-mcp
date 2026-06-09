"""P0 #2 — XER forecast-finish driver / top-level WBS anomaly (RULE 16.C).

Pure functions over read_tasks()/read_wbs() output. No I/O, no COM.

Implements the ALFB1 lesson algorithm:
  1. Map every task to its top-level WBS ancestor.
  2. Compute each top-level branch's forecast finish (max forecast_finish).
  3. If the latest branch finishes > anomaly_gap_days after the next one,
     flag an anomaly and drill into its latest activity.
  4. Surface LOE tasks dragging the finish (ALFB1: PR-HG-ELEC-FD-3450, a
     single TT_LOE in Procurement > Infrastructure dragged finish 5 months).
"""
import datetime as _dt


def _iso(s):
    if not s:
        return None
    try:
        return _dt.date.fromisoformat(str(s)[:10])
    except (ValueError, TypeError):
        return None


def _days_between(later, earlier):
    a, b = _iso(later), _iso(earlier)
    if a is None or b is None:
        return None
    return (a - b).days


def _build_top_level_map(wbs):
    """Map every wbs_id -> its top-level ancestor wbs_id.

    Root = node whose parent_id is None/0 or not a known wbs id.
    Top-level = a node whose parent is a root. Returns (top_of, name_by_id).
    """
    by_id = {w["id"]: w for w in wbs if w.get("id") is not None}
    ids = set(by_id)
    name_by_id = {w["id"]: w.get("name", "") for w in wbs}

    def parent_of(wid):
        p = by_id[wid].get("parent_id")
        return p if (p in ids) else None  # None => parent is root/absent

    roots = {wid for wid in ids if parent_of(wid) is None}
    top_of = {}
    for wid in ids:
        cur = wid
        seen = set()
        while cur is not None and cur not in seen:
            seen.add(cur)
            p = parent_of(cur)
            if p is None or p in roots:
                top_of[wid] = cur  # cur is a top-level node
                break
            cur = p
        else:
            top_of[wid] = wid
    return top_of, name_by_id


def forecast_drivers(tasks, wbs, anomaly_gap_days=30, top_n=5):
    """Compute top-level WBS forecast finishes + anomaly driver (RULE 16.C).

    Args:
        tasks: read_tasks() output (needs forecast_finish, wbs_id, task_type,
               code, name, status, percent_complete, duration_h).
        wbs:   read_wbs() output (id, parent_id, name).
        anomaly_gap_days: gap (latest vs next branch) that flags an anomaly.
        top_n: how many latest-finishing activities to list.

    Returns a structured dict (see keys below). Pure; safe on empty input.
    """
    top_of, name_by_id = _build_top_level_map(wbs)

    branches = {}
    dated_tasks = []
    for t in tasks:
        ff = t.get("forecast_finish")
        if not ff:
            continue
        dated_tasks.append(t)
        wid = t.get("wbs_id")
        top = top_of.get(wid, wid)
        b = branches.setdefault(top, {
            "wbs_id": top,
            "name": name_by_id.get(top, "") or (f"WBS {top}" if top is not None
                                                else "UNKNOWN"),
            "forecast_finish": None,
            "task_count": 0,
            "driving_task": None,
        })
        b["task_count"] += 1
        if b["forecast_finish"] is None or ff > b["forecast_finish"]:
            b["forecast_finish"] = ff
            b["driving_task"] = t

    branch_list = sorted(
        branches.values(),
        key=lambda b: b["forecast_finish"] or "", reverse=True)

    result = {
        "project_forecast_finish": branch_list[0]["forecast_finish"]
            if branch_list else None,
        "branch_count": len(branch_list),
        "branches": [
            {"wbs_id": b["wbs_id"], "name": b["name"],
             "forecast_finish": b["forecast_finish"],
             "task_count": b["task_count"]}
            for b in branch_list
        ],
        "anomaly": False,
        "gap_days": None,
        "driver": None,
        "latest_tasks": [],
    }

    # Latest-finishing activities overall (driver candidates)
    latest = sorted(dated_tasks, key=lambda t: t["forecast_finish"],
                    reverse=True)[:top_n]
    result["latest_tasks"] = [
        {"id": t.get("id"), "code": t.get("code"), "name": t.get("name"),
         "task_type": t.get("task_type"),
         "wbs_name": name_by_id.get(t.get("wbs_id"), ""),
         "forecast_finish": t.get("forecast_finish"),
         "status": t.get("status")}
        for t in latest
    ]

    # Anomaly: latest branch finishes >> next branch
    if len(branch_list) >= 2:
        gap = _days_between(branch_list[0]["forecast_finish"],
                            branch_list[1]["forecast_finish"])
        result["gap_days"] = gap
        if gap is not None and gap > anomaly_gap_days:
            result["anomaly"] = True
            drv_branch = branch_list[0]
            dt = drv_branch["driving_task"] or {}
            is_loe = dt.get("task_type") == "TT_LOE"
            pct = float(dt.get("percent_complete") or 0)
            note_bits = [
                f"Top-level WBS '{drv_branch['name']}' finishes {gap} days "
                f"after the next branch '{branch_list[1]['name']}'."
            ]
            if is_loe:
                note_bits.append(
                    "Driving activity is a TT_LOE (Level of Effort) — RULE "
                    "16.C: an LOE open while the field work is done drags the "
                    "forecast finish (ALFB1 PR-HG-ELEC-FD-3450 pattern).")
            if 0 < pct < 100:
                note_bits.append(
                    f"Driving activity is {pct:.0f}% complete (in progress).")
            result["driver"] = {
                "wbs_id": drv_branch["wbs_id"],
                "wbs_name": drv_branch["name"],
                "forecast_finish": drv_branch["forecast_finish"],
                "is_loe": is_loe,
                "driving_task": {
                    "id": dt.get("id"), "code": dt.get("code"),
                    "name": dt.get("name"), "task_type": dt.get("task_type"),
                    "wbs_id": dt.get("wbs_id"),
                    "wbs_name": name_by_id.get(dt.get("wbs_id"), ""),
                    "duration_h": dt.get("duration_h"),
                    "status": dt.get("status"),
                    "percent_complete": dt.get("percent_complete"),
                    "forecast_finish": dt.get("forecast_finish"),
                },
                "note": " ".join(note_bits),
            }

    return result
