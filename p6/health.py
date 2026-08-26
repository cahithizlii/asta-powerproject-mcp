"""DCMA 14-Point assessment for P6 projects.

Thin on purpose. Every threshold and every count comes from ``dcma_checks`` --
the same module msproject_health uses -- so the same programme scores the same
whether it is read from a P6 database, an XER export or an MPP. What this file
adds is only the P6 shaping (``p6.analysis.load``) and the four actions.

Two P6 details that change results, both read from the project and reported
back rather than assumed:

* ``day_hr_cnt`` drives Rule 9's 44-working-day threshold (352h at 8h/day,
  486h on a 6x9 Uzbekistan calendar).
* ``critical`` comes from ``PROJECT.critical_drtn_hr_cnt``, P6's own critical
  float threshold, not from a hard-coded "float <= 0".
"""
from __future__ import annotations

import datetime as _dt
import json
import os
from typing import Any, Mapping

from . import analysis

DEFAULT_SNAPSHOT = os.path.expanduser("~/p6_evm_snapshots.json")


def _assess(params: Mapping[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    """Load once, run the 14 checks once. Returns (data, dcma_result)."""
    import dcma_checks

    data = analysis.load(params)
    result = dcma_checks.assess_all(
        tasks=data["tasks"],
        links=data["links"],
        assignments=data["assignments"],
        baseline=None,
        status_date=data.get("status_date"),
        day_hr_cnt=data.get("day_hr_cnt") or 8.0,
    )
    return data, result


ID_SAMPLE = 10


def _slim(rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Keep the per-rule id lists out of the 14-rule response.

    Rule 7 alone can name 898 activities here. Left whole, the envelope blows
    past the response cap and mcp_common trims the longest list it can find --
    which is ``rules`` itself, so the caller silently receives one rule
    instead of fourteen. Counts and a short sample stay; the full list is what
    ``drill_down`` is for.
    """
    out = []
    for r in rules:
        slim = dict(r)
        for key in ("failed_task_ids", "critical_task_ids", "failed_links"):
            items = slim.get(key)
            if not isinstance(items, list):
                continue
            slim[key + "_sample"] = items[:ID_SAMPLE]
            slim[key + "_total"] = len(items)
            del slim[key]
        out.append(slim)
    return out


def _context(data: Mapping[str, Any]) -> dict[str, Any]:
    """The facts a reader needs to trust the score."""
    return {
        "source": data.get("source"),
        "day_hr_cnt": data.get("day_hr_cnt"),
        "status_date": data.get("status_date"),
        "task_count": data.get("task_count"),
        "summary_task_count": data.get("summary_task_count"),
        "baseline_source": data.get("baseline_source"),
        "critical_float_threshold_days": data.get("critical_float_threshold_days"),
        "baseline_warnings": data.get("baseline_warnings"),
    }


def assess_all(params: Mapping[str, Any]) -> dict[str, Any]:
    """All 14 rules with per-rule counts, plus the summary envelope."""
    data, result = _assess(params)
    return {"action": "assess_all", **_context(data),
            "rules": _slim(result["rules"]),
            "summary": result["summary"],
            "id_lists_note": ("Kural basina tam aktivite listesi icin "
                              "action='drill_down', rule_id=<1-14>.")}


def summary(params: Mapping[str, Any]) -> dict[str, Any]:
    """RAG + executive line only."""
    data, result = _assess(params)
    return {"action": "summary", **_context(data), **result["summary"]}


def drill_down(params: Mapping[str, Any]) -> dict[str, Any]:
    """The activities that failed one rule, named -- not just counted."""
    try:
        rule_id = int(params.get("rule_id", 0))
    except (TypeError, ValueError):
        rule_id = 0
    if rule_id not in range(1, 15):
        raise analysis.AnalysisError(
            "rule_id 1-14 arasinda olmali; '%s' verildi." % params.get("rule_id"))

    data, result = _assess(params)
    rule = next((r for r in result["rules"] if r["id"] == rule_id), None)
    if rule is None:
        raise analysis.AnalysisError("Kural %s bulunamadi." % rule_id)

    limit = min(max(int(params.get("limit", 100) or 100), 1), 500)

    # Rules 3-5 fail on relationships, not activities; name the links instead.
    bad_links = rule.get("failed_links")
    if isinstance(bad_links, list):
        by_id = {t["id"]: t for t in data["tasks"]}
        # dcma_checks returns only from/to/lag; the relationship type has to
        # come back from the full link list, or every row reads "type: null".
        type_by_pair = {(l.get("from_id"), l.get("to_id")): l.get("type")
                        for l in data["links"]}

        def _name(tid):
            t = by_id.get(tid) or {}
            return {"id": tid, "code": t.get("code"), "name": t.get("name")}

        listed = [{"predecessor": _name(l.get("from_id")),
                   "successor": _name(l.get("to_id")),
                   "type": type_by_pair.get((l.get("from_id"), l.get("to_id"))),
                   "lag_days": l.get("lag_days")}
                  for l in bad_links[:limit]]
        return {
            "action": "drill_down",
            **_context(data),
            "rule": {"id": rule["id"], "name": rule["name"],
                     "threshold": rule.get("threshold"),
                     "actual": rule.get("actual"),
                     "actual_unit": rule.get("actual_unit"),
                     "status": rule.get("status")},
            "failed_count": rule.get("failed_count"),
            "total_count": rule.get("total_count"),
            "returned": len(listed),
            "truncated": len(bad_links) > len(listed),
            "links": listed,
        }

    ids = rule.get("failed_task_ids") or rule.get("critical_task_ids") or []
    by_id = {t["id"]: t for t in data["tasks"]}
    listed = []
    for tid in ids[:limit]:
        t = by_id.get(tid)
        if t is None:
            listed.append({"id": tid})
            continue
        listed.append({
            "id": tid,
            "code": t.get("code"),
            "name": t.get("name"),
            "duration_h": t.get("duration_h"),
            "total_slack_days": t.get("total_slack_days"),
            "start": t.get("start"),
            "finish": t.get("finish"),
            "forecast_finish": t.get("forecast_finish"),
            "percent_complete": t.get("percent_complete"),
        })
    return {
        "action": "drill_down",
        **_context(data),
        "rule": {"id": rule["id"], "name": rule["name"],
                 "threshold": rule.get("threshold"),
                 "actual": rule.get("actual"),
                 "actual_unit": rule.get("actual_unit"),
                 "status": rule.get("status")},
        "failed_count": rule.get("failed_count"),
        "total_count": rule.get("total_count"),
        "returned": len(listed),
        "truncated": len(ids) > len(listed),
        "tasks": listed,
    }


def compare(params: Mapping[str, Any]) -> dict[str, Any]:
    """This assessment against the last saved one -- what improved, what slipped."""
    data, result = _assess(params)
    path = params.get("snapshot_path") or DEFAULT_SNAPSHOT
    current = {"rules": result["rules"], "summary": result["summary"]}

    prev = None
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                snaps = json.load(fh)
        except (OSError, ValueError) as exc:
            raise analysis.AnalysisError(
                "Snapshot dosyasi okunamadi (%s): %s" % (path, exc))
        dcma_snaps = [s for s in snaps if isinstance(s, dict) and s.get("dcma")]
        dcma_snaps.sort(key=lambda s: s.get("saved_at", ""))
        if dcma_snaps:
            prev = dcma_snaps[-1]

    if prev is None:
        return {"action": "compare", **_context(data),
                "snapshot_path": path,
                "current": result["summary"], "previous": None,
                "note": ("Karsilastirilacak onceki DCMA snapshot'i yok; "
                         "p6_evm action='save_period_snapshot' ile kaydedin."),
                "delta": {"improved": [], "degraded": []}}

    prev_rules = {r["id"]: r for r in (prev.get("dcma", {}).get("rules") or [])}
    improved, degraded = [], []
    for cur in current["rules"]:
        old = prev_rules.get(cur["id"])
        if old is None:
            continue
        entry = {"id": cur["id"], "name": cur["name"],
                 "from_actual": old.get("actual"), "to_actual": cur.get("actual")}
        if old.get("status") == "fail" and cur.get("status") == "pass":
            improved.append(entry)
        elif old.get("status") == "pass" and cur.get("status") == "fail":
            degraded.append(entry)

    return {
        "action": "compare",
        **_context(data),
        "snapshot_path": path,
        "current": result["summary"],
        "previous": prev.get("dcma", {}).get("summary"),
        "previous_saved_at": prev.get("saved_at"),
        "previous_tag": prev.get("tag"),
        "delta": {"improved": improved, "degraded": degraded},
    }


def dcma_payload(params: Mapping[str, Any]) -> dict[str, Any]:
    """The DCMA block p6_evm stores inside a period snapshot.

    Slimmed like assess_all: a snapshot file accumulates one of these per
    period, and the activity id lists would dwarf everything else in it.
    """
    _data, result = _assess(params)
    return {"rules": _slim(result["rules"]), "summary": result["summary"]}


ACTIONS = {
    "assess_all": assess_all,
    "summary": summary,
    "drill_down": drill_down,
    "compare": compare,
}


def now_iso() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")
