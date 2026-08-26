"""Earned Value Management for P6 projects.

The arithmetic is ``evm_math`` (PMI PMBOK 8th 7.4.2, Lipke 2003 for Earned
Schedule) and ``currency_validator`` for cost-loading -- the same modules
msproject_evm calls, and the action names match it too, so a report pipeline
can point at either server without learning a second vocabulary.

What is P6-specific and therefore lives here:

* **Real baselines.** ``variance_to_baseline`` and ``compare_baselines_evm``
  take ``baseline_proj_id``; P6 stores each baseline as its own project, so
  these are true baseline comparisons rather than the single implicit
  target-date baseline an XER can carry.
* **Unit honesty.** Every response repeats ``units``, ``units_reason`` and the
  BAC each of the three bases would have produced. A BAC quoted without its
  unit is how the ALFB1 9x error survived review (RULE 16.A).

Snapshots (period_delta / trend / get_period_history and the DCMA compare)
share one JSON file, ``~/p6_evm_snapshots.json`` by default.
"""
from __future__ import annotations

import datetime as _dt
import json
import os
from typing import Any, Mapping

from . import analysis, health

DEFAULT_SNAPSHOT = health.DEFAULT_SNAPSHOT


# ---------------------------------------------------------------------------
# shared plumbing
# ---------------------------------------------------------------------------
def _units_block(data: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "units": data.get("units"),
        "units_reason": data.get("units_reason"),
        "cost_loading_mode": data.get("cost_loading_mode"),
        "candidate_bac": data.get("candidate_bac"),
        "units_warnings": data.get("units_warnings"),
    }


def _context(data: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "source": data.get("source"),
        "day_hr_cnt": data.get("day_hr_cnt"),
        "status_date": data.get("status_date"),
        "task_count": data.get("task_count"),
        "baseline_source": data.get("baseline_source"),
        "baseline_proj_id": data.get("baseline_proj_id"),
        "baseline_warnings": data.get("baseline_warnings"),
        **_units_block(data),
    }


def _metrics(params: Mapping[str, Any]) -> tuple[dict, dict, dict]:
    """Load -> aggregate -> SPI/CPI/SV/CV. Returns (data, agg, metrics)."""
    import evm_math

    data = analysis.load(params)
    agg = analysis.aggregate(data)
    metrics = evm_math.compute_metrics(bac=agg["bac"], pv=agg["pv"],
                                       ev=agg["ev"], ac=agg["ac"])
    return data, agg, metrics


def _pv_curve(data: Mapping[str, Any], bucket: str = "week"):
    """[(bucket_end, cumulative PV)] across the baseline envelope."""
    import evm_math

    start, finish = analysis.project_bounds(data)
    periods = analysis.buckets(start, finish, bucket)
    dated = [{"baseline_start": analysis.to_date(t.get("baseline_start")),
              "baseline_finish": analysis.to_date(t.get("baseline_finish")),
              "baseline_work": float(t.get("baseline_work") or 0)}
             for t in data["tasks"]]
    dated = [t for t in dated if t["baseline_start"] and t["baseline_finish"]]
    cum = evm_math.time_phased_pv(dated, periods)
    return start, finish, periods, [(p[1], v) for p, v in zip(periods, cum)]


def _completion_pct(metrics: Mapping[str, Any]) -> float:
    bac = float(metrics.get("bac") or 0)
    return (float(metrics.get("ev") or 0) / bac * 100.0) if bac > 0 else 0.0


def _round(payload: dict[str, Any], places: int = 2) -> dict[str, Any]:
    return {k: (round(v, places) if isinstance(v, float) else v)
            for k, v in payload.items()}


# ---------------------------------------------------------------------------
# snapshot store
# ---------------------------------------------------------------------------
def _snapshot_path(params: Mapping[str, Any]) -> str:
    return params.get("snapshot_path") or DEFAULT_SNAPSHOT


def _snapshot_read(path: str) -> list[dict[str, Any]]:
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError) as exc:
        raise analysis.AnalysisError(
            "Snapshot dosyasi okunamadi (%s): %s" % (path, exc))
    return data if isinstance(data, list) else []


def _snapshot_write(path: str, snaps: list[dict[str, Any]]) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(snaps, fh, ensure_ascii=False, indent=2, default=str)
    os.replace(tmp, path)


def _matching(snaps: list[dict[str, Any]], params: Mapping[str, Any]):
    proj = params.get("project_filter") or params.get("proj_short_name")
    out = [s for s in snaps if isinstance(s, dict)]
    if proj:
        out = [s for s in out if s.get("project") == proj
               or s.get("proj_id") == proj]
    return sorted(out, key=lambda s: s.get("saved_at", ""))


# ---------------------------------------------------------------------------
# actions
# ---------------------------------------------------------------------------
def compute_metrics(params: Mapping[str, Any]) -> dict[str, Any]:
    """SPI / CPI / SV / CV at the data date (RULE 4)."""
    data, agg, metrics = _metrics(params)
    return {"action": "compute_metrics", **_context(data),
            "tasks_without_baseline_dates": agg["tasks_without_baseline_dates"],
            "data_date": agg["data_date"], **_round(metrics, 4)}


def forecast(params: Mapping[str, Any]) -> dict[str, Any]:
    """EAC1/2/3, ETC, VAC, TCPI (RULE 9, PMBOK 8th 7.4.2)."""
    import evm_math

    data, agg, metrics = _metrics(params)
    fc = evm_math.forecast(bac=metrics["bac"], ev=metrics["ev"],
                           ac=metrics["ac"], cpi=metrics.get("cpi"),
                           spi=metrics.get("spi"))
    return {"action": "forecast", **_context(data),
            "data_date": agg["data_date"],
            **_round(metrics, 4), **_round(fc, 2)}


def earned_schedule(params: Mapping[str, Any]) -> dict[str, Any]:
    """AT, ES, SV(t), SPI(t) -- Lipke 2003 (RULE 8)."""
    import evm_math

    data, agg, metrics = _metrics(params)
    data_date = analysis.to_date(agg["data_date"])
    start, _finish, _periods, curve = _pv_curve(data, params.get("bucket", "week"))
    es = _earned_schedule_block(evm_math, curve, metrics["ev"], start, data_date)
    return {"action": "earned_schedule", **_context(data),
            "project_start": start.isoformat(),
            "data_date": agg["data_date"],
            "ev": round(metrics["ev"], 2),
            "pv_curve_points": len(curve), **es}


def _earned_schedule_block(evm_math, curve, ev_now, start, data_date) -> dict[str, Any]:
    """Earned Schedule, or an honest refusal when no time has elapsed.

    evm_math clamps AT to 1e-9 weeks so the division survives; that turns a
    data date sitting on the project start into SPI(t) = 0.0, which reads as
    catastrophic schedule failure when it actually means "nothing has been
    scheduled to happen yet". Below one elapsed day, report None and say why.
    """
    elapsed_days = (data_date - start).days
    if elapsed_days < 1:
        return {"at_weeks": 0.0, "es_weeks": None, "sv_t_weeks": None,
                "spi_t": None,
                "earned_schedule_note": (
                    "Veri tarihi proje baslangicindan sonra degil (gecen sure "
                    "%d gun); Earned Schedule tanimsiz." % elapsed_days)}
    es = evm_math.earned_schedule(pv_curve=curve, ev_now=ev_now,
                                  project_start=start, data_date=data_date)
    return {
        "at_weeks": round(es["at"], 3) if es.get("at") is not None else None,
        "es_weeks": round(es["es"], 3) if es.get("es") is not None else None,
        "sv_t_weeks": round(es["sv_t"], 3) if es.get("sv_t") is not None else None,
        "spi_t": round(es["spi_t"], 4) if es.get("spi_t") is not None else None,
    }


def summary(params: Mapping[str, Any]) -> dict[str, Any]:
    """RAG + completion + the two indices, for a report cover page (RULE 12)."""
    import evm_math

    data, agg, metrics = _metrics(params)
    pct = _completion_pct(metrics)
    rag = evm_math.rag_status(spi=metrics.get("spi"), completion_pct=pct)
    return {"action": "summary", **_context(data),
            "data_date": agg["data_date"],
            "rag": rag, "completion_pct": round(pct, 2),
            "bac": round(metrics["bac"], 2), "pv": round(metrics["pv"], 2),
            "ev": round(metrics["ev"], 2), "ac": round(metrics["ac"], 2),
            "spi": round(metrics["spi"], 4) if metrics.get("spi") is not None else None,
            "cpi": round(metrics["cpi"], 4) if metrics.get("cpi") is not None else None}


def time_phased_evm(params: Mapping[str, Any]) -> dict[str, Any]:
    """Cumulative PV / EV / AC per period -- the S-curve (RULE 5)."""
    import evm_math

    data = analysis.load(params)
    agg = analysis.aggregate(data)
    data_date = analysis.to_date(agg["data_date"])
    bucket = params.get("bucket", "week")
    start, finish = analysis.project_bounds(data)
    periods = analysis.buckets(start, finish, bucket)

    tasks = [{"baseline_start": analysis.to_date(t.get("baseline_start")),
              "baseline_finish": analysis.to_date(t.get("baseline_finish")),
              "baseline_work": float(t.get("baseline_work") or 0),
              "percent_complete": float(t.get("percent_complete") or 0),
              "actual_work": float(t.get("actual_work") or 0),
              "actual_start": analysis.to_date(t.get("actual_start")),
              "actual_finish": analysis.to_date(t.get("actual_finish"))}
             for t in data["tasks"]]
    dated = [t for t in tasks if t["baseline_start"] and t["baseline_finish"]]

    pv = evm_math.time_phased_pv(dated, periods)
    ev = evm_math.time_phased_ev(dated, periods, data_date)
    ac = evm_math.time_phased_ac(tasks, periods, data_date)

    limit = min(max(int(params.get("limit", 200) or 200), 1), 500)
    rows = [{"period_start": s.isoformat(), "period_end": e.isoformat(),
             "pv": round(p, 2), "ev": round(v, 2), "ac": round(a, 2)}
            for (s, e), p, v, a in zip(periods, pv, ev, ac)]
    return {"action": "time_phased_evm", **_context(data),
            "bucket": bucket, "project_start": start.isoformat(),
            "project_finish": finish.isoformat(),
            "data_date": agg["data_date"],
            "count": len(rows), "truncated": len(rows) > limit,
            "periods": rows[:limit]}


def period_delta(params: Mapping[str, Any]) -> dict[str, Any]:
    """This period's PV / EV / AC movement against the last snapshot (RULE 6)."""
    import evm_math

    data, agg, metrics = _metrics(params)
    snaps = _matching(_snapshot_read(_snapshot_path(params)), params)
    prev = snaps[-1].get("metrics") if snaps else None
    delta = evm_math.period_delta(metrics, prev)
    return {"action": "period_delta", **_context(data),
            "data_date": agg["data_date"],
            "snapshot_path": _snapshot_path(params),
            "previous_saved_at": snaps[-1].get("saved_at") if snaps else None,
            "previous_tag": snaps[-1].get("tag") if snaps else None,
            "current": _round(metrics, 2),
            **_round(delta, 2)}


def progress_data_quality(params: Mapping[str, Any]) -> dict[str, Any]:
    """Is the progress data good enough to believe the SPI? (RULE 7)"""
    import evm_math

    data, agg, metrics = _metrics(params)
    start, _f, _p, curve = _pv_curve(data, params.get("bucket", "week"))
    es = _earned_schedule_block(evm_math, curve, metrics["ev"], start,
                                analysis.to_date(agg["data_date"]))
    pct = _completion_pct(metrics)
    warnings = evm_math.progress_data_quality(
        spi_h=metrics.get("spi"), spi_t=es.get("spi_t"),
        completion_pct=pct, has_resources=bool(data.get("assignments")))

    tasks = data["tasks"]
    started_no_pct = sum(1 for t in tasks
                         if t.get("actual_start") and not float(t.get("percent_complete") or 0))
    pct_no_actual = sum(1 for t in tasks
                        if float(t.get("percent_complete") or 0) > 0
                        and not float(t.get("actual_work") or 0))
    if pct == 0:
        warnings.append({
            "warning": ("Kaydedilmis ilerleme yok (EV = 0); SPI/CPI bu veriyle "
                        "anlam tasimaz."),
            "severity": "high"})
    if started_no_pct:
        warnings.append({
            "warning": ("%d aktivitenin fiili baslangici var ama yuzde "
                        "tamamlanmasi 0." % started_no_pct),
            "severity": "medium"})
    if pct_no_actual:
        warnings.append({
            "warning": ("%d aktivitede ilerleme girilmis ama fiili is/maliyet "
                        "yok -- sessiz EV." % pct_no_actual),
            "severity": "high"})
    if data.get("units_warnings"):
        warnings.extend({"warning": w, "severity": "medium"}
                        for w in data["units_warnings"])

    return {"action": "progress_data_quality", **_context(data),
            "data_date": agg["data_date"],
            "completion_pct": round(pct, 2),
            "spi_h": round(metrics["spi"], 4) if metrics.get("spi") is not None else None,
            "spi_t": round(es["spi_t"], 4) if es.get("spi_t") is not None else None,
            "assignment_count": len(data.get("assignments") or []),
            "warning_count": len(warnings), "warnings": warnings}


def verify(params: Mapping[str, Any]) -> dict[str, Any]:
    """Cross-check the BAC against the raw assignment sum (RULE 16.A).

    This is the ALFB1 defence: a tool once reported 277.640h while the raw
    sum of target_qty was 2.505.038h. Any BAC that goes into a report should
    survive this check first.
    """
    import evm_math

    data = analysis.load(params)
    tasks, assignments = data["tasks"], data["assignments"]
    bac_primary = sum(float(t.get("baseline_work") or 0) for t in tasks)

    raw_key = "target_cost" if data.get("units") == "cost" else "target_qty"
    bac_independent = (sum(float(a.get(raw_key) or 0) for a in assignments)
                       if assignments else None)
    result = evm_math.cross_validate_bac(
        bac_primary, bac_independent,
        tolerance=float(params.get("tolerance", 0.01)))

    qty_by_task: dict[Any, float] = {}
    for a in assignments:
        tid = a.get("task_id")
        if tid is not None:
            qty_by_task[tid] = qty_by_task.get(tid, 0.0) + float(a.get(raw_key) or 0)
    zero_rollup = [{"id": t.get("id"), "code": t.get("code"),
                    "name": t.get("name"),
                    "assignment_qty": round(qty_by_task.get(t.get("id"), 0.0), 2)}
                   for t in tasks
                   if float(t.get("baseline_work") or 0) == 0
                   and qty_by_task.get(t.get("id"), 0.0) > 0]

    return {"action": "verify", **_context(data),
            "independent_field": raw_key,
            "task_count": len(tasks), "assignment_count": len(assignments),
            "zero_baseline_with_assignments_count": len(zero_rollup),
            "zero_baseline_with_assignments": zero_rollup[:50],
            **result}


def detect_currency_mode(params: Mapping[str, Any]) -> dict[str, Any]:
    """Cost-loaded or hours-only? (RULE 3)"""
    data = analysis.load(params)
    return {"action": "detect_currency_mode", **_context(data),
            "assignment_count": len(data.get("assignments") or []),
            "resource_count": len(data.get("resources") or [])}


def validate_currency_mode(params: Mapping[str, Any]) -> dict[str, Any]:
    """Do the assignment rows and the resource rows tell the same story?"""
    from currency_validator import (cross_validate_modes,
                                    detect_mode_from_tasks_resources,
                                    detect_mode_from_xer_assignments)

    data = analysis.load(params)
    sources = [
        ("TASKRSRC", detect_mode_from_xer_assignments(data.get("assignments") or [])),
        ("TASK+RSRC", detect_mode_from_tasks_resources(data.get("tasks") or [],
                                                       data.get("resources") or [])),
    ]
    return {"action": "validate_currency_mode", **_context(data),
            "sources": [{"source": s, "mode": m} for s, m in sources],
            **cross_validate_modes(sources)}


def variance_to_baseline(params: Mapping[str, Any]) -> dict[str, Any]:
    """Current dates against a real P6 baseline project, activity by activity."""
    b_id = params.get("baseline_proj_id")
    if b_id is None:
        raise analysis.AnalysisError(
            "variance_to_baseline icin baseline_proj_id zorunlu. Baseline "
            "projelerini p6_query action='list_projects' ile bulun.")
    data, agg, metrics = _metrics(params)

    rows = []
    slipped = late = early = 0
    for t in data["tasks"]:
        bf = analysis.to_date(t.get("baseline_finish"))
        ff = analysis.to_date(t.get("forecast_finish"))
        if bf is None or ff is None:
            continue
        delta = (ff - bf).days
        if delta > 0:
            late += 1
        elif delta < 0:
            early += 1
        if delta > int(params.get("slip_threshold_days", 0) or 0):
            slipped += 1
        rows.append({"id": t.get("id"), "code": t.get("code"),
                     "name": t.get("name"),
                     "baseline_finish": bf.isoformat(),
                     "forecast_finish": ff.isoformat(),
                     "variance_days": delta,
                     "percent_complete": t.get("percent_complete")})
    rows.sort(key=lambda r: r["variance_days"], reverse=True)
    limit = min(max(int(params.get("limit", 50) or 50), 1), 500)

    return {"action": "variance_to_baseline", **_context(data),
            "data_date": agg["data_date"],
            **_round(metrics, 2),
            "compared_tasks": len(rows),
            "late_tasks": late, "early_tasks": early, "slipped_tasks": slipped,
            "max_slip_days": rows[0]["variance_days"] if rows else None,
            "returned": min(len(rows), limit),
            "truncated": len(rows) > limit,
            "worst": rows[:limit]}


def compare_baselines_evm(params: Mapping[str, Any]) -> dict[str, Any]:
    """Two P6 baselines side by side -- what the re-baseline actually changed."""
    a_id = params.get("baseline_proj_id_a")
    b_id = params.get("baseline_proj_id_b")
    if a_id is None or b_id is None:
        raise analysis.AnalysisError(
            "compare_baselines_evm icin baseline_proj_id_a ve "
            "baseline_proj_id_b zorunlu.")

    def _side(bid):
        p = dict(params)
        p["baseline_proj_id"] = bid
        _d, agg, metrics = _metrics(p)
        start, finish = analysis.project_bounds(_d)
        return {"baseline_proj_id": int(bid),
                "baseline_short_name": _d.get("baseline_short_name"),
                "baseline_start": start.isoformat(),
                "baseline_finish": finish.isoformat(),
                "unmatched_tasks": _d.get("baseline_unmatched_tasks", 0),
                **_round(metrics, 2)}, _d

    side_a, data_a = _side(a_id)
    side_b, _data_b = _side(b_id)
    delta = {
        "bac": round(side_b["bac"] - side_a["bac"], 2),
        "pv": round(side_b["pv"] - side_a["pv"], 2),
        "ev": round(side_b["ev"] - side_a["ev"], 2),
        "finish_days": (_dt.date.fromisoformat(side_b["baseline_finish"])
                        - _dt.date.fromisoformat(side_a["baseline_finish"])).days,
    }
    return {"action": "compare_baselines_evm", **_context(data_a),
            "baseline_a": side_a, "baseline_b": side_b, "delta": delta}


def save_period_snapshot(params: Mapping[str, Any]) -> dict[str, Any]:
    """Freeze this period's EVM (and optionally DCMA) for the next comparison."""
    import evm_math

    data, agg, metrics = _metrics(params)
    fc = evm_math.forecast(bac=metrics["bac"], ev=metrics["ev"],
                           ac=metrics["ac"], cpi=metrics.get("cpi"),
                           spi=metrics.get("spi"))
    pct = _completion_pct(metrics)
    snap: dict[str, Any] = {
        "saved_at": health.now_iso(),
        "tag": params.get("tag"),
        "project": (data.get("project") or {}).get("proj_short_name"),
        "proj_id": (data.get("project") or {}).get("proj_id"),
        "data_date": agg["data_date"],
        "units": data.get("units"),
        "baseline_source": data.get("baseline_source"),
        "metrics": _round(metrics, 4),
        "forecast": _round(fc, 2),
        "completion_pct": round(pct, 2),
        "rag": evm_math.rag_status(spi=metrics.get("spi"), completion_pct=pct),
    }
    if params.get("include_dcma", True):
        snap["dcma"] = health.dcma_payload(params)

    path = _snapshot_path(params)
    snaps = _snapshot_read(path)
    snaps.append(snap)
    _snapshot_write(path, snaps)
    return {"action": "save_period_snapshot", "snapshot_path": path,
            "total_snapshots": len(snaps),
            "saved": {k: v for k, v in snap.items() if k != "dcma"},
            "dcma_included": "dcma" in snap}


def get_period_history(params: Mapping[str, Any]) -> dict[str, Any]:
    """List what has been frozen so far."""
    path = _snapshot_path(params)
    snaps = _matching(_snapshot_read(path), params)
    limit = min(max(int(params.get("limit", 50) or 50), 1), 500)
    rows = [{"saved_at": s.get("saved_at"), "tag": s.get("tag"),
             "project": s.get("project"), "data_date": s.get("data_date"),
             "units": s.get("units"), "rag": s.get("rag"),
             "completion_pct": s.get("completion_pct"),
             "has_dcma": bool(s.get("dcma"))}
            for s in snaps]
    return {"action": "get_period_history", "snapshot_path": path,
            "count": len(rows), "truncated": len(rows) > limit,
            "snapshots": rows[-limit:]}


def trend(params: Mapping[str, Any]) -> dict[str, Any]:
    """SPI / CPI / EAC trajectory across the saved periods."""
    path = _snapshot_path(params)
    snaps = _matching(_snapshot_read(path), params)
    series = []
    for s in snaps:
        m = s.get("metrics") or {}
        f = s.get("forecast") or {}
        series.append({"saved_at": s.get("saved_at"), "tag": s.get("tag"),
                       "data_date": s.get("data_date"),
                       "spi": m.get("spi"), "cpi": m.get("cpi"),
                       "ev": m.get("ev"), "ac": m.get("ac"),
                       "eac_t3": f.get("eac_t3"), "rag": s.get("rag"),
                       "dcma_pass": ((s.get("dcma") or {}).get("summary") or {})
                       .get("pass_count")})
    return {"action": "trend", "snapshot_path": path,
            "count": len(series), "series": series}


ACTIONS = {
    "compute_metrics": compute_metrics,
    "forecast": forecast,
    "earned_schedule": earned_schedule,
    "summary": summary,
    "time_phased_evm": time_phased_evm,
    "period_delta": period_delta,
    "progress_data_quality": progress_data_quality,
    "variance_to_baseline": variance_to_baseline,
    "compare_baselines_evm": compare_baselines_evm,
    "save_period_snapshot": save_period_snapshot,
    "get_period_history": get_period_history,
    "trend": trend,
    "detect_currency_mode": detect_currency_mode,
    "validate_currency_mode": validate_currency_mode,
    "verify": verify,
}
