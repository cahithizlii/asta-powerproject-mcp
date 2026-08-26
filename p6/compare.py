"""Compare two P6 schedules -- revisions, baselines, or XER files.

The arithmetic is ``xer_compare``, the same module msproject_compare uses. What
this file adds is the one thing that makes the comparison valid for P6:

**The join key is ``task_code``, never ``task_id``.** P6 renumbers ids every
time a schedule crosses a boundary -- a CLI import, a baseline copy, an export
and re-import. The activity that is 3274452 in an XER is 35847 in the database
after import, and the same activity in a baseline of that project is a third
number again. ``xer_compare.diff_tasks`` matches on ``id``, so both sides are
re-keyed onto their activity codes first; otherwise every activity reads as
"removed on one side, added on the other" and the comparison is worse than
useless -- it looks precise while saying nothing.

Both sides go through ``analysis.load``, so a comparison inherits everything
that layer already enforces: hours-per-day from the calendar, percent complete
on the activity's own basis, and the BAC unit reported rather than assumed.
Mixing units across the two sides (one cost-loaded, one hours-only) is
reported as a warning rather than silently subtracted.
"""
from __future__ import annotations

from typing import Any, Mapping

from . import analysis

# Task fields worth diffing by default. Everything else is either derived
# (float, early/late dates) or noise for a revision report.
DEFAULT_FIELDS = ("name", "duration_h", "start", "finish", "forecast_finish",
                  "actual_start", "actual_finish", "percent_complete",
                  "status", "constraint_type", "total_slack_days")


class CompareError(analysis.AnalysisError):
    """Bad comparison request."""


def _side_params(params: Mapping[str, Any], which: str) -> dict[str, Any]:
    """Pull one side's source spec out of the request.

    Accepts ``a``/``b`` objects, or the flat ``proj_id_a`` / ``path_b`` style.
    """
    side = params.get(which)
    if isinstance(side, Mapping):
        out = dict(side)
    else:
        out = {}
        suffix = "_" + which
        for key, value in params.items():
            if key.endswith(suffix):
                out[key[: -len(suffix)]] = value
    if not out:
        raise CompareError(
            "'%s' tarafi icin kaynak verilmedi. Ornek: {\"a\": {\"proj_id\": 368}, "
            "\"b\": {\"baseline_proj_id\": 369}} ya da proj_id_a / path_b."
            % which)
    # A baseline is a project in its own right; comparing against one means
    # loading it as the project, not as project+baseline.
    if out.get("baseline_proj_id") is not None and out.get("proj_id") is None:
        out["proj_id"] = out.pop("baseline_proj_id")
    for shared in ("alias", "units", "day_hr_cnt"):
        if shared in params and shared not in out:
            out[shared] = params[shared]
    return out


def _rekey(data: Mapping[str, Any]) -> tuple[list[dict], list[dict], dict[str, int]]:
    """Re-key tasks and links onto activity codes so the two sides can meet."""
    code_of: dict[Any, str] = {}
    tasks: list[dict] = []
    duplicate_codes = 0
    seen: set[str] = set()
    for t in data["tasks"]:
        code = t.get("code")
        if not code:
            continue
        if code in seen:
            duplicate_codes += 1
            continue
        seen.add(code)
        code_of[t["id"]] = code
        tasks.append({**t, "id": code, "task_id_original": t["id"]})

    links = []
    unmapped = 0
    for link in data["links"]:
        a, b = code_of.get(link.get("from_id")), code_of.get(link.get("to_id"))
        if a is None or b is None:
            unmapped += 1
            continue
        links.append({**link, "from_id": a, "to_id": b})

    return tasks, links, {"duplicate_codes": duplicate_codes,
                          "links_outside_project": unmapped}


def _load_side(params: Mapping[str, Any], which: str):
    side = _side_params(params, which)
    data = analysis.load(side)
    tasks, links, notes = _rekey(data)
    return data, tasks, links, notes


def _context(a: Mapping[str, Any], b: Mapping[str, Any],
             na: Mapping[str, Any], nb: Mapping[str, Any]) -> dict[str, Any]:
    warnings: list[str] = []
    if a.get("units") != b.get("units"):
        warnings.append(
            "Iki taraf farkli birimde (%s / %s); BAC ve EV farklari "
            "karsilastirilamaz." % (a.get("units"), b.get("units")))
    if a.get("day_hr_cnt") != b.get("day_hr_cnt"):
        warnings.append(
            "Gun-saat farkli (%s / %s); sure ve float gun cinsinden "
            "karsilastirilirken dikkat." % (a.get("day_hr_cnt"), b.get("day_hr_cnt")))
    if a.get("status_date") != b.get("status_date"):
        warnings.append(
            "Veri tarihleri farkli (%s / %s) -- fark bunun bir kismini "
            "aciklayabilir." % (a.get("status_date"), b.get("status_date")))
    for label, notes in (("a", na), ("b", nb)):
        if notes.get("duplicate_codes"):
            warnings.append(
                "%s tarafinda %d aktivite kodu tekrar ediyor; ilki alindi."
                % (label, notes["duplicate_codes"]))
    return {
        "a": {"source": a.get("source"), "status_date": a.get("status_date"),
              "task_count": a.get("task_count"), "units": a.get("units"),
              "day_hr_cnt": a.get("day_hr_cnt")},
        "b": {"source": b.get("source"), "status_date": b.get("status_date"),
              "task_count": b.get("task_count"), "units": b.get("units"),
              "day_hr_cnt": b.get("day_hr_cnt")},
        "join_key": "task_code",
        "warnings": warnings,
    }


def _trim(diff: Mapping[str, Any], limit: int) -> dict[str, Any]:
    """Keep the counts, cap the lists -- a 950-activity delta blows the envelope."""
    out: dict[str, Any] = {}
    for key, value in diff.items():
        if isinstance(value, list):
            out[key + "_count"] = len(value)
            out[key] = value[:limit]
            if len(value) > limit:
                out[key + "_truncated"] = True
        else:
            out[key] = value
    return out


# ---------------------------------------------------------------------------
# actions
# ---------------------------------------------------------------------------
def tasks(params: Mapping[str, Any]) -> dict[str, Any]:
    """Added / removed / changed activities, matched on activity code."""
    import xer_compare

    a, ta, _la, na = _load_side(params, "a")
    b, tb, _lb, nb = _load_side(params, "b")
    fields = params.get("fields") or list(DEFAULT_FIELDS)
    limit = min(max(int(params.get("limit", 50) or 50), 1), 300)
    diff = xer_compare.diff_tasks(ta, tb, fields=fields)
    return {"action": "tasks", **_context(a, b, na, nb),
            "compared_fields": fields, **_trim(diff, limit)}


def links(params: Mapping[str, Any]) -> dict[str, Any]:
    """Logic changes -- relationships added, removed, or re-typed/re-lagged."""
    import xer_compare

    a, _ta, la, na = _load_side(params, "a")
    b, _tb, lb, nb = _load_side(params, "b")
    limit = min(max(int(params.get("limit", 50) or 50), 1), 300)
    diff = xer_compare.diff_links(la, lb)
    ctx = _context(a, b, na, nb)
    for label, notes in (("a", na), ("b", nb)):
        if notes.get("links_outside_project"):
            ctx["warnings"].append(
                "%s tarafinda %d bag proje disina isaret ediyor; "
                "karsilastirmaya alinmadi." % (label, notes["links_outside_project"]))
    return {"action": "links", **ctx, **_trim(diff, limit)}


def progress(params: Mapping[str, Any]) -> dict[str, Any]:
    """Which activities moved, and by how much."""
    import xer_compare

    a, ta, _la, na = _load_side(params, "a")
    b, tb, _lb, nb = _load_side(params, "b")
    limit = min(max(int(params.get("limit", 50) or 50), 1), 300)

    def snap(data, tlist):
        return {"status_date": data.get("status_date"),
                "tasks": [{"id": t["id"],
                           "percent_complete": t.get("percent_complete"),
                           "actual_work": t.get("actual_work")} for t in tlist]}

    diff = xer_compare.diff_progress(snap(a, ta), snap(b, tb))
    return {"action": "progress", **_context(a, b, na, nb), **_trim(diff, limit)}


def evm(params: Mapping[str, Any]) -> dict[str, Any]:
    """BAC / PV / EV / AC and the indices, side by side."""
    import evm_math
    import xer_compare

    def metrics(data):
        agg = analysis.aggregate(data)
        return evm_math.compute_metrics(bac=agg["bac"], pv=agg["pv"],
                                        ev=agg["ev"], ac=agg["ac"])

    a, _ta, _la, na = _load_side(params, "a")
    b, _tb, _lb, nb = _load_side(params, "b")
    ma, mb = metrics(a), metrics(b)
    return {"action": "evm", **_context(a, b, na, nb),
            "metrics_a": {k: (round(v, 4) if isinstance(v, float) else v)
                          for k, v in ma.items()},
            "metrics_b": {k: (round(v, 4) if isinstance(v, float) else v)
                          for k, v in mb.items()},
            "delta": xer_compare.diff_evm(ma, mb)}


def summary(params: Mapping[str, Any]) -> dict[str, Any]:
    """The one-paragraph answer: what changed between these two schedules."""
    import evm_math
    import xer_compare

    a, ta, la, na = _load_side(params, "a")
    b, tb, lb, nb = _load_side(params, "b")
    fields = params.get("fields") or list(DEFAULT_FIELDS)

    task_d = xer_compare.diff_tasks(ta, tb, fields=fields)
    link_d = xer_compare.diff_links(la, lb)

    def snap(data, tlist):
        return {"status_date": data.get("status_date"),
                "tasks": [{"id": t["id"],
                           "percent_complete": t.get("percent_complete"),
                           "actual_work": t.get("actual_work")} for t in tlist]}

    prog_d = xer_compare.diff_progress(snap(a, ta), snap(b, tb))

    def metrics(data):
        agg = analysis.aggregate(data)
        return evm_math.compute_metrics(bac=agg["bac"], pv=agg["pv"],
                                        ev=agg["ev"], ac=agg["ac"])

    evm_d = xer_compare.diff_evm(metrics(a), metrics(b))
    head = xer_compare.summarize_compare(task_d, link_d, prog_d, evm_d)

    # Finish-date movement is what a revision report actually leads with, and
    # it is not part of the generic summary.
    fa = {t["id"]: analysis.to_date(t.get("forecast_finish")) for t in ta}
    fb = {t["id"]: analysis.to_date(t.get("forecast_finish")) for t in tb}
    slips = [(code, (fb[code] - fa[code]).days)
             for code in fa if code in fb and fa[code] and fb[code]]
    late = [s for s in slips if s[1] > 0]
    early = [s for s in slips if s[1] < 0]
    slips.sort(key=lambda s: -s[1])
    limit = min(max(int(params.get("limit", 10) or 10), 1), 100)

    ends_a = [d for d in fa.values() if d]
    ends_b = [d for d in fb.values() if d]
    project_slip = ((max(ends_b) - max(ends_a)).days
                    if ends_a and ends_b else None)

    return {
        "action": "summary", **_context(a, b, na, nb),
        "headline": head,
        "finish_movement": {
            "compared": len(slips), "later": len(late), "earlier": len(early),
            "unchanged": len(slips) - len(late) - len(early),
            "project_finish_a": max(ends_a).isoformat() if ends_a else None,
            "project_finish_b": max(ends_b).isoformat() if ends_b else None,
            "project_slip_days": project_slip,
            "worst": [{"code": c, "slip_days": d} for c, d in slips[:limit]],
        },
    }


ACTIONS = {
    "tasks": tasks,
    "links": links,
    "progress": progress,
    "evm": evm,
    "summary": summary,
}
