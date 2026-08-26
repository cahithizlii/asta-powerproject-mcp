"""Shared analysis loader for the P6 EVM and DCMA tools.

No new maths lives here. The 14 DCMA checks stay in ``dcma_checks``, the EVM
formulae stay in ``evm_math`` and cost-loading detection stays in
``currency_validator`` -- exactly the modules msproject_mcp_core already uses,
so a P6 project and an MPP/XER of the same programme go through identical
arithmetic. This module only does the P6-specific part: read one source
(database alias or XER file) and shape it into the dicts those modules expect.

Two P6 facts drive the design and neither may be guessed (RULE 0):

* **Hours per day** comes from the project calendar (``p6.source`` resolves it
  and refuses to default to 8).
* **What "baseline" means.** P6 keeps real baselines as separate PROJECT rows,
  so ``baseline_proj_id`` loads one and matches activities by ``task_code``.
  Without one the target (planned) dates are the baseline -- the same
  convention the XER path in msproject_mcp_core uses -- and ``baseline_source``
  says so in every response.

The unit a BAC is measured in is likewise reported, never assumed: see
``resolve_units``.
"""
from __future__ import annotations

import datetime as _dt
from typing import Any, Mapping, Sequence

from . import db as p6db
from . import source as src

UNIT_CHOICES = ("auto", "duration_h", "qty", "cost")

# Task fields DCMA/EVM read; everything else is carried through untouched.
_EVM_FIELDS = ("baseline_start", "baseline_finish", "baseline_work",
               "actual_work", "total_slack_days", "critical",
               "predecessors", "successors")


class AnalysisError(RuntimeError):
    """Bad analysis request; the tools turn it into a JSON error."""


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------
def _f(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def to_date(value: Any) -> _dt.date | None:
    """ISO string / date / datetime -> date. Never raises."""
    if value is None or value == "":
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    try:
        return _dt.date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def _rollup(assignments: Sequence[Mapping[str, Any]], key: str) -> dict[int, float]:
    """Sum one assignment field per task_id."""
    out: dict[int, float] = {}
    for a in assignments:
        tid = a.get("task_id")
        if tid is None:
            continue
        out[tid] = out.get(tid, 0.0) + _f(a.get(key))
    return out


def _link_maps(links: Sequence[Mapping[str, Any]]):
    """One O(M) pass -> predecessor and successor id lists per task."""
    preds: dict[int, list[int]] = {}
    succs: dict[int, list[int]] = {}
    for link in links:
        a, b = link.get("from_id"), link.get("to_id")
        if a is None or b is None:
            continue
        preds.setdefault(b, []).append(a)
        succs.setdefault(a, []).append(b)
    return preds, succs


def _project_row(bag) -> dict[str, str]:
    rows = bag.tables.get("PROJECT", {"rows": []})["rows"]
    return rows[0] if rows else {}


def _target_vs_schedule(bag) -> dict[str, Any]:
    """How far P6's planned (target) dates have drifted from the live CPM run.

    This decides whether "no baseline given" is usable at all. P6 keeps
    Planned Start/Finish in sync with the early dates for activities that
    have not started, so in a live database the target dates are frequently
    just a copy of the current schedule -- comparing the schedule against
    itself yields SPI = 1 and a slip of zero, forever. bukhtourcity is the
    worst case: 950 of 950 activities match. Measured, never assumed.
    """
    rows = bag.tables.get("TASK", {"rows": []})["rows"]
    total = same = started = 0
    for row in rows:
        total += 1
        if row.get("act_start_date"):
            started += 1
        if (row.get("target_start_date") == row.get("early_start_date")
                and row.get("target_end_date") == row.get("early_end_date")):
            same += 1
    pct = (same / total * 100.0) if total else 0.0
    return {"task_count": total, "matching_current_schedule": same,
            "matching_pct": round(pct, 2), "started_tasks": started}


# ---------------------------------------------------------------------------
# units
# ---------------------------------------------------------------------------
def resolve_units(requested: str | None,
                  tasks: Sequence[Mapping[str, Any]],
                  assignments: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    """Decide whether BAC is measured in cost, assignment hours or duration.

    ``auto`` follows the cost-loading mode ``currency_validator`` detects and
    falls back down the chain when the richer source is empty:
    cost -> assignment qty -> task target duration. Whatever is chosen, the
    three candidate totals are returned as well, so a caller (and the user)
    can see the number the other basis would have produced instead of having
    to trust one figure -- this is the ALFB1 9x class of error (RULE 16.A).
    """
    from currency_validator import detect_mode_from_xer_assignments

    requested = (requested or "auto").lower()
    if requested not in UNIT_CHOICES:
        raise AnalysisError(
            "units '%s' gecersiz; sunlardan biri olmali: %s"
            % (requested, ", ".join(UNIT_CHOICES)))

    mode = detect_mode_from_xer_assignments(list(assignments))
    totals = {
        "cost": sum(_f(a.get("target_cost")) for a in assignments),
        "qty": sum(_f(a.get("target_qty")) for a in assignments),
        "duration_h": sum(_f(t.get("duration_h")) for t in tasks
                          if not t.get("summary")),
    }

    if requested != "auto":
        chosen = requested
        why = "parametre"
    elif mode in ("cost", "mixed") and totals["cost"] > 0:
        chosen = "cost"
        why = "cost_loading=" + mode
    elif totals["qty"] > 0:
        chosen = "qty"
        why = "atama saatleri (TASKRSRC.target_qty)"
    else:
        chosen = "duration_h"
        why = "atama yok; gorev hedef suresi"

    warnings: list[str] = []
    if chosen == "duration_h" and totals["qty"] > 0:
        warnings.append(
            "BAC gorev suresinden, AC atama saatinden geliyor; iki farkli "
            "kaynak. CPI'yi yorumlarken dikkat edin.")
    if mode in ("cost", "mixed") and totals["cost"] == 0:
        # detect_mode_from_xer_assignments reads "target_cost != target_qty"
        # as a cost signal, so an hours-only programme whose target_cost is a
        # plain 0 lands on 'cost'. Say so rather than let the label stand.
        warnings.append(
            "Cost loading '%s' raporlandi ama toplam target_cost = 0; program "
            "maliyet yuklu DEGIL, saat bazli. BAC saat cinsinden alindi." % mode)
    if mode == "mixed":
        warnings.append(
            "Cost loading 'mixed': bazi atamalarda maliyet var, bazilarinda "
            "yok (RULE 3 ihlali olasi).")
    if mode == "uncertain" and chosen != "duration_h":
        warnings.append("Cost loading belirlenemedi (atama verisi bos/sifir).")

    return {"units": chosen, "units_reason": why, "cost_loading_mode": mode,
            "candidate_bac": {k: round(v, 2) for k, v in totals.items()},
            "units_warnings": warnings}


# ---------------------------------------------------------------------------
# baseline
# ---------------------------------------------------------------------------
def _baseline_by_code(backend, baseline_proj_id: int,
                      day_hr_cnt: float) -> tuple[dict[str, dict], dict]:
    """Load a P6 baseline project and key its activities by task_code.

    A P6 baseline is a copy of the project with fresh TASK_IDs, so the join
    has to be on the activity code, not the id.
    """
    bag = p6db.load_bag(backend, baseline_proj_id)
    rows = bag.tables.get("TASK", {"rows": []})["rows"]
    if not rows:
        raise AnalysisError(
            "Baseline projesi %s icin aktivite bulunamadi." % baseline_proj_id)
    tasks = p6db.read_tasks(bag, day_hr_cnt)
    by_code = {t.get("code"): t for t in tasks if t.get("code")}
    proj = _project_row(bag)
    meta = {
        "baseline_proj_id": baseline_proj_id,
        "baseline_short_name": proj.get("proj_short_name") or None,
        "baseline_task_count": len(tasks),
        "baseline_data_date": to_date(proj.get("last_recalc_date")),
    }
    if len(by_code) != len(tasks):
        meta["baseline_duplicate_codes"] = len(tasks) - len(by_code)
    return by_code, meta


# ---------------------------------------------------------------------------
# main loader
# ---------------------------------------------------------------------------
def load(params: Mapping[str, Any]) -> dict[str, Any]:
    """Read a P6 source once; return everything DCMA and EVM need.

    Returns ``{source, day_hr_cnt, status_date, project, tasks, links,
    assignments, resources, units..., baseline_source, ...}`` where every task
    dict already carries the ``_EVM_FIELDS``.
    """
    with src.OpenSource(params) as s:
        s.require_project()
        day_hr = s.day_hr_cnt
        bag = s.bag

        tasks = p6db.read_tasks(bag, day_hr)
        links = p6db.read_links(bag, day_hr)
        assignments = p6db.read_assignments(bag)
        resources = p6db.read_resources(bag)
        project = p6db.read_project(bag)
        raw_project = _project_row(bag)

        baseline_meta: dict[str, Any] = {"baseline_source": "target_dates"}
        baseline_by_code: dict[str, dict] = {}
        drift = _target_vs_schedule(bag)
        b_id = params.get("baseline_proj_id")
        if b_id is not None:
            if s.meta.get("type") != "db":
                raise AnalysisError(
                    "baseline_proj_id yalnizca 'db' kaynaginda kullanilabilir; "
                    "XER dosyasi tek bir programi tasir.")
            baseline_by_code, meta = _baseline_by_code(
                s.backend, int(b_id), day_hr)
            baseline_meta = {"baseline_source": "baseline_project", **meta}

        source_meta = dict(s.meta)

    # ---- shaping (backend already closed; everything below is plain data) --
    status_date = (params.get("status_date")
                   or project.get("last_recalc_date"))

    # P6 carries its own critical-float threshold; do not assume zero.
    crit_hr = _f(raw_project.get("critical_drtn_hr_cnt"))
    crit_days = crit_hr / day_hr if day_hr else 0.0

    preds, succs = _link_maps(links)
    units = resolve_units(params.get("units"), tasks, assignments)
    basis = units["units"]

    if basis == "cost":
        target_by_task = _rollup(assignments, "target_cost")
        actual_by_task = _rollup(assignments, "actual_cost")
    else:
        target_by_task = _rollup(assignments, "target_qty")
        actual_by_task = _rollup(assignments, "actual_qty")

    shaped: list[dict[str, Any]] = []
    missing_baseline = 0
    for t in tasks:
        if t.get("summary"):
            continue
        tid = t["id"]
        if baseline_by_code:
            b = baseline_by_code.get(t.get("code"))
            if b is None:
                missing_baseline += 1
                b_start, b_finish = None, None
            else:
                b_start, b_finish = b.get("start"), b.get("finish")
        else:
            b_start, b_finish = t.get("start"), t.get("finish")

        if basis == "duration_h":
            baseline_work = _f(t.get("duration_h"))
        else:
            baseline_work = target_by_task.get(tid, 0.0)

        slack = _f(t.get("total_float"))
        shaped.append({
            **t,
            "baseline_start": b_start,
            "baseline_finish": b_finish,
            "baseline_work": baseline_work,
            "actual_work": actual_by_task.get(tid, 0.0),
            "total_slack_days": slack,
            "critical": slack <= crit_days,
            "predecessors": preds.get(tid, []),
            "successors": succs.get(tid, []),
        })

    if missing_baseline:
        baseline_meta["baseline_unmatched_tasks"] = missing_baseline

    baseline_warnings: list[str] = []
    if baseline_meta["baseline_source"] == "target_dates":
        baseline_meta["target_vs_current_schedule"] = drift
        if drift["matching_pct"] >= 95.0:
            baseline_warnings.append(
                "Gercek baseline verilmedi ve aktivitelerin %%%.1f'inde hedef "
                "(planned) tarihler mevcut programla AYNI. Bu bir baseline "
                "degil, programin kendisidir: SPI ~ 1 ve gecikme ~ 0 cikar. "
                "Anlamli EVM icin baseline_proj_id verin (baseline projelerini "
                "p6_query action='list_projects' ile listeleyin)."
                % drift["matching_pct"])
        elif drift["matching_pct"] >= 50.0:
            baseline_warnings.append(
                "Gercek baseline verilmedi; aktivitelerin %%%.1f'inde hedef "
                "tarihler mevcut programla ayni, dolayisiyla bu kadari icin "
                "sapma yapisal olarak sifir cikar."
                % drift["matching_pct"])
    if baseline_warnings:
        baseline_meta["baseline_warnings"] = baseline_warnings

    out: dict[str, Any] = {
        "source": source_meta,
        "day_hr_cnt": day_hr,
        "status_date": status_date,
        "project": project,
        "critical_float_threshold_days": round(crit_days, 4),
        "critical_float_threshold_source": "PROJECT.critical_drtn_hr_cnt",
        "tasks": shaped,
        "links": list(links),
        "assignments": list(assignments),
        "resources": list(resources),
        "task_count": len(shaped),
        "summary_task_count": len(tasks) - len(shaped),
        **units,
        **baseline_meta,
    }
    return out


# ---------------------------------------------------------------------------
# EVM aggregation
# ---------------------------------------------------------------------------
def aggregate(data: Mapping[str, Any]) -> dict[str, Any]:
    """BAC / PV / EV / AC at the data date.

    BAC = sum(baseline_work)
    EV  = sum(baseline_work x percent_complete / 100)
    AC  = sum(actual_work)
    PV  = evm_math linear distribution at the data date (RULE 5)
    """
    import evm_math

    tasks = data.get("tasks") or []
    bac = sum(_f(t.get("baseline_work")) for t in tasks)
    ev = sum(_f(t.get("baseline_work")) * _f(t.get("percent_complete")) / 100.0
             for t in tasks)
    ac = sum(_f(t.get("actual_work")) for t in tasks)

    data_date = to_date(data.get("status_date"))
    if data_date is None:
        raise AnalysisError(
            "Veri tarihi (status_date) okunamadi; PV hesaplanamaz. "
            "status_date parametresini acikca verin (RULE 0).")

    dated = [{"baseline_start": to_date(t.get("baseline_start")),
              "baseline_finish": to_date(t.get("baseline_finish")),
              "baseline_work": _f(t.get("baseline_work"))}
             for t in tasks]
    dated = [t for t in dated
             if t["baseline_start"] and t["baseline_finish"]]
    pv = (evm_math.time_phased_pv(dated, [(_dt.date.min, data_date)])[0]
          if dated else 0.0)

    return {"bac": bac, "pv": pv, "ev": ev, "ac": ac,
            "data_date": data_date.isoformat(),
            "tasks_without_baseline_dates": len(tasks) - len(dated)}


def project_bounds(data: Mapping[str, Any]) -> tuple[_dt.date, _dt.date]:
    """Baseline start/finish envelope across the shaped tasks."""
    starts = [d for d in (to_date(t.get("baseline_start"))
                          for t in data.get("tasks") or []) if d]
    finishes = [d for d in (to_date(t.get("baseline_finish"))
                            for t in data.get("tasks") or []) if d]
    if not starts or not finishes:
        raise AnalysisError(
            "Projede baseline baslangic/bitis tarihi olan aktivite yok.")
    return min(starts), max(finishes)


def buckets(start: _dt.date, finish: _dt.date,
            bucket: str = "week") -> list[tuple[_dt.date, _dt.date]]:
    """Period list from start to finish; 'day' | 'week' | 'month'."""
    bucket = (bucket or "week").lower()
    if bucket not in ("day", "week", "month"):
        raise AnalysisError("bucket 'day', 'week' veya 'month' olmali.")
    step = {"day": 1, "week": 7, "month": 30}[bucket]
    out: list[tuple[_dt.date, _dt.date]] = []
    cursor = start
    guard = 0
    while cursor <= finish and guard < 5000:
        end = min(cursor + _dt.timedelta(days=step - 1), finish)
        out.append((cursor, end))
        cursor = end + _dt.timedelta(days=1)
        guard += 1
    return out
