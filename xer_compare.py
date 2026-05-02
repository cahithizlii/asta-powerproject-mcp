"""Phase 7 — pure-Python snapshot diff for XER/MSPDI projects.

Powers the `msproject_compare` MCP tool. Zero I/O, no COM, no fixtures.
Adapters in msproject_mcp_core wrap `_evm_load_task_data` for either
file format and pass shaped lists/dicts here.

Public API:
    diff_tasks(tasks_a, tasks_b, fields=None) -> dict
    diff_links(links_a, links_b) -> dict
    diff_progress(progress_a, progress_b) -> dict
    diff_evm(snap_a, snap_b) -> dict
    summarize_compare(task_d, link_d, progress_d, evm_d) -> dict

Identity:
    task          : id (int)
    link          : (from_id, to_id, type) tuple
    progress_task : id (int)
"""
from typing import Any, Dict, Iterable, List, Optional, Tuple


# Default task fields scanned for change detection (override via param).
DEFAULT_TASK_FIELDS = (
    "baseline_start",
    "baseline_finish",
    "baseline_work",
    "percent_complete",
    "actual_start",
    "actual_finish",
    "actual_work",
    "duration_h",
)


def _index_by(items: Optional[Iterable[Dict[str, Any]]],
              key: str) -> Dict[Any, Dict[str, Any]]:
    """Return {item[key]: item} for items with a non-None key."""
    out: Dict[Any, Dict[str, Any]] = {}
    for x in items or ():
        k = x.get(key)
        if k is None:
            continue
        out[k] = x
    return out


def _values_differ(a: Any, b: Any) -> bool:
    """Field-level comparison. None != 0; "" treated equal to None for
    ergonomic diff (XER often emits "" while MSPDI emits None)."""
    if (a in (None, "")) and (b in (None, "")):
        return False
    return a != b


def diff_tasks(tasks_a: Optional[List[Dict[str, Any]]],
               tasks_b: Optional[List[Dict[str, Any]]],
               fields: Optional[Iterable[str]] = None) -> Dict[str, Any]:
    """Compare two task lists by id.

    Returns:
        {
          added:   tasks present in B but not A,
          removed: tasks present in A but not B,
          changed: [{id, name, fields_changed: {field: (a_val, b_val)}}],
          unchanged_count: int,
        }
    """
    field_set = tuple(fields) if fields else DEFAULT_TASK_FIELDS
    a_idx = _index_by(tasks_a, "id")
    b_idx = _index_by(tasks_b, "id")
    a_ids = set(a_idx.keys())
    b_ids = set(b_idx.keys())

    added = [b_idx[i] for i in sorted(b_ids - a_ids, key=str)]
    removed = [a_idx[i] for i in sorted(a_ids - b_ids, key=str)]

    changed: List[Dict[str, Any]] = []
    unchanged = 0
    for tid in sorted(a_ids & b_ids, key=str):
        ta, tb = a_idx[tid], b_idx[tid]
        fc: Dict[str, Tuple[Any, Any]] = {}
        for f in field_set:
            va, vb = ta.get(f), tb.get(f)
            if _values_differ(va, vb):
                fc[f] = (va, vb)
        if fc:
            changed.append({
                "id": tid,
                "name": tb.get("name") or ta.get("name"),
                "fields_changed": fc,
            })
        else:
            unchanged += 1
    return {
        "added": added,
        "removed": removed,
        "changed": changed,
        "unchanged_count": unchanged,
    }


def diff_links(links_a: Optional[List[Dict[str, Any]]],
               links_b: Optional[List[Dict[str, Any]]]) -> Dict[str, Any]:
    """Compare link lists by (from_id, to_id, type) identity.

    Lag change with identity preserved -> 'changed' (not removed/added).
    """
    def key(l: Dict[str, Any]) -> Tuple:
        return (l.get("from_id"), l.get("to_id"), l.get("type"))

    a_idx = {key(l): l for l in (links_a or ())}
    b_idx = {key(l): l for l in (links_b or ())}
    a_keys = set(a_idx.keys())
    b_keys = set(b_idx.keys())

    added = [b_idx[k] for k in sorted(b_keys - a_keys, key=str)]
    removed = [a_idx[k] for k in sorted(a_keys - b_keys, key=str)]

    changed: List[Dict[str, Any]] = []
    unchanged = 0
    for k in sorted(a_keys & b_keys, key=str):
        la, lb = a_idx[k], b_idx[k]
        lag_a = la.get("lag_days")
        lag_b = lb.get("lag_days")
        if _values_differ(lag_a, lag_b):
            changed.append({
                "from_id": k[0], "to_id": k[1], "type": k[2],
                "lag_a": lag_a, "lag_b": lag_b,
            })
        else:
            unchanged += 1
    return {
        "added": added,
        "removed": removed,
        "changed": changed,
        "unchanged_count": unchanged,
    }


def diff_progress(progress_a: Optional[Dict[str, Any]],
                  progress_b: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Compare per-task percent_complete + actual_work between snapshots.

    Returns only tasks whose pct or actual_work changed.
    """
    a = progress_a or {}
    b = progress_b or {}
    a_tasks = _index_by(a.get("tasks"), "id")
    b_tasks = _index_by(b.get("tasks"), "id")
    moved: List[Dict[str, Any]] = []
    total_pct_delta = 0.0
    total_aw_delta = 0.0
    for tid in sorted(set(a_tasks.keys()) | set(b_tasks.keys()), key=str):
        ta = a_tasks.get(tid, {})
        tb = b_tasks.get(tid, {})
        pct_a = float(ta.get("percent_complete") or 0)
        pct_b = float(tb.get("percent_complete") or 0)
        aw_a = float(ta.get("actual_work") or 0)
        aw_b = float(tb.get("actual_work") or 0)
        pct_delta = pct_b - pct_a
        aw_delta = aw_b - aw_a
        if abs(pct_delta) > 0.01 or abs(aw_delta) > 0.01:
            moved.append({
                "id": tid,
                "pct_a": pct_a, "pct_b": pct_b,
                "pct_delta": round(pct_delta, 2),
                "aw_a": aw_a, "aw_b": aw_b,
                "aw_delta": round(aw_delta, 2),
            })
            total_pct_delta += pct_delta
            total_aw_delta += aw_delta
    return {
        "status_date_a": a.get("status_date"),
        "status_date_b": b.get("status_date"),
        "tasks": moved,
        "summary": {
            "count_moved": len(moved),
            "total_pct_delta": round(total_pct_delta, 2),
            "total_aw_delta": round(total_aw_delta, 2),
        },
    }


def _safe_sub(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    try:
        return round(float(b) - float(a), 4)
    except (TypeError, ValueError):
        return None


def diff_evm(snap_a: Optional[Dict[str, Any]],
             snap_b: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Compare two EVM metric snapshots.

    Each snapshot dict expected to have keys: bac, pv, ev, ac, spi, cpi.
    Missing keys treated as None; deltas only computed when both sides
    have a numeric value.
    """
    a = snap_a or {}
    b = snap_b or {}
    return {
        "bac_a": a.get("bac"), "bac_b": b.get("bac"),
        "bac_delta": _safe_sub(a.get("bac"), b.get("bac")),
        "pv_a": a.get("pv"), "pv_b": b.get("pv"),
        "pv_delta": _safe_sub(a.get("pv"), b.get("pv")),
        "ev_a": a.get("ev"), "ev_b": b.get("ev"),
        "ev_delta": _safe_sub(a.get("ev"), b.get("ev")),
        "ac_a": a.get("ac"), "ac_b": b.get("ac"),
        "ac_delta": _safe_sub(a.get("ac"), b.get("ac")),
        "spi_a": a.get("spi"), "spi_b": b.get("spi"),
        "spi_delta": _safe_sub(a.get("spi"), b.get("spi")),
        "cpi_a": a.get("cpi"), "cpi_b": b.get("cpi"),
        "cpi_delta": _safe_sub(a.get("cpi"), b.get("cpi")),
    }


def summarize_compare(task_d: Dict[str, Any],
                      link_d: Dict[str, Any],
                      progress_d: Dict[str, Any],
                      evm_d: Dict[str, Any]) -> Dict[str, Any]:
    """High-level compare summary suitable for a single-line headline."""
    counts = {
        "tasks_added": len(task_d.get("added", [])),
        "tasks_removed": len(task_d.get("removed", [])),
        "tasks_changed": len(task_d.get("changed", [])),
        "links_added": len(link_d.get("added", [])),
        "links_removed": len(link_d.get("removed", [])),
        "links_changed": len(link_d.get("changed", [])),
        "tasks_progressed": progress_d.get("summary", {}).get(
            "count_moved", 0),
    }
    spi_a = evm_d.get("spi_a")
    spi_b = evm_d.get("spi_b")
    spi_str = ""
    if spi_a is not None and spi_b is not None:
        spi_str = f"SPI {spi_a:.2f} -> {spi_b:.2f}"
    headline_parts = []
    if counts["tasks_added"]:
        headline_parts.append(f"{counts['tasks_added']} tasks added")
    if counts["tasks_removed"]:
        headline_parts.append(f"{counts['tasks_removed']} tasks removed")
    if counts["tasks_progressed"]:
        headline_parts.append(
            f"{counts['tasks_progressed']} progressed")
    if spi_str:
        headline_parts.append(spi_str)
    headline = ", ".join(headline_parts) or "no changes detected"
    return {
        "headline": headline,
        "counts": counts,
        "spi_delta": evm_d.get("spi_delta"),
        "cpi_delta": evm_d.get("cpi_delta"),
        "ev_delta": evm_d.get("ev_delta"),
    }
