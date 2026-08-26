"""Drive P6's own command line: import an XER, then repair what it drops.

P6's CLI is the only supported way to get a schedule *into* the database, and
it only ever CREATEs -- it cannot update an existing project. That limitation
is what makes it the right tool for a revision: every import lands as a new
project you can compare against the old one.

**It also silently drops resource rates, and that is not a P6 bug.** The
import action script accepts an `importConfiguration`, which P6 resolves to a
`VIEWPROP` row of `view_type='VP_IMP_OPT'`. This database has none, so P6
applies its defaults: insert the resource, zero its rate. The evidence is that
the imported RSRCRATE rows carry the correct rsrc_id, max_qty_per_hr and
start_date and only `cost_per_qty` is zero -- and the importing user is a
global superuser, so it is not a privilege either. Rather than reverse-
engineer the undocumented `view_data` encoding, `repair_costs` writes the
rates back from the source XER afterwards and *measures* the result. Same
pattern as the Cyrillic repair, same reason: a deterministic, verifiable step
beats a guess at a private format.

Passwords never appear here as literals. The caller supplies one (from the
machine's own credential store, via `password` or `P6_CLI_PASSWORD`), and it
is registered with ``mcp_common.register_secret`` so it cannot surface in any
log, echo or error this module produces.
"""
from __future__ import annotations

import os
import subprocess
import tempfile
import time
from typing import Any, Mapping

import mcp_common as mc

from . import db as p6db, write as w

LAUNCHER = os.environ.get(
    "P6_LAUNCHER",
    r"C:\Program Files\Oracle\Primavera P6\P6 Professional\24.12.0"
    r"\Primavera.CacheService.exe")


class CliError(RuntimeError):
    """The P6 command line refused or failed."""


def _password(params: Mapping[str, Any]) -> str:
    pw = params.get("password") or os.environ.get("P6_CLI_PASSWORD")
    if not pw:
        raise CliError(
            "P6 kullanici parolasi gerekli. 'password' parametresiyle ya da "
            "P6_CLI_PASSWORD ortam degiskeniyle verin -- bu modul parolayi "
            "hicbir yere yazmaz, yalnizca P6'ya gecirir.")
    mc.register_secret(pw)
    return pw


def _refuse_if_p6_running() -> None:
    """PM.exe acikken CLI calismaz; sessizce yarim is birakmaktansa reddet."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq PM.exe", "/NH"],
            capture_output=True, text=True, timeout=30).stdout
    except Exception:  # noqa: BLE001
        return
    if "PM.exe" in out:
        raise CliError(
            "P6 Professional (PM.exe) acik -- komut satiri bu haldeyken "
            "calismaz. P6'yi kapatip tekrar deneyin.")


def _action_script(xer_path: str, eps: str, work_dir: str) -> str:
    path = os.path.join(work_dir, "mcp_import_action.xml")
    body = "\n".join((
        '<?xml version="1.0" encoding="windows-1252"?>',
        "<actions>",
        "  <action>",
        "    <type>import</type>",
        "    <importFormat>XER</importFormat>",
        "    <importType>PROJECT</importType>",
        "    <importAction>CREATE</importAction>",
        "    <importTo>%s</importTo>" % eps,
        "    <importFile>%s</importFile>" % xer_path,
        "  </action>",
        "</actions>",
    ))
    with open(path, "w", encoding="ascii", errors="replace") as fh:
        fh.write(body)
    return path


def _projects(alias) -> dict[int, str]:
    backend, _info = p6db.open_backend(alias, use_snapshot=False)
    try:
        return {int(r["proj_id"]): r["proj_short_name"]
                for r in p6db.list_projects(backend, "Y")}
    finally:
        backend.close()


def import_xer(params: Mapping[str, Any]) -> dict[str, Any]:
    """Import an XER as a NEW project and report which project appeared.

    The CLI does not tell you the id it created, so the project list is taken
    before and after and the difference is the answer -- measured, not parsed
    out of a log line that may change between builds.
    """
    xer = params.get("path") or params.get("xer")
    if not xer:
        raise CliError("'path' (XER dosyasi) zorunlu.")
    if not os.path.exists(xer):
        raise CliError("XER bulunamadi: " + xer)
    if not os.path.exists(LAUNCHER):
        raise CliError("P6 launcher bulunamadi: " + LAUNCHER)
    if not params.get("confirm"):
        raise w.P6WriteError(
            "Import veritabanina YENI BIR PROJE ekler. confirm=true ile "
            "tekrar cagirin.")

    _refuse_if_p6_running()
    alias = p6db.resolve_alias(params.get("alias"))
    user = params.get("user", "admin")
    pw = _password(params)
    eps = params.get("eps", "EPS")
    timeout = int(params.get("timeout_s", 600))

    work_dir = params.get("work_dir") or tempfile.gettempdir()
    action = _action_script(xer, eps, work_dir)
    log_path = os.path.join(work_dir, "mcp_import.log")
    for stale in (log_path,):
        if os.path.exists(stale):
            os.remove(stale)

    before = _projects(alias)
    t0 = time.time()
    proc = subprocess.run(
        [LAUNCHER, "/username=" + user, "/password=" + pw,
         "/alias=" + alias.name, "/actionScript=" + action,
         "/pmlogfile=" + log_path],
        capture_output=True, text=True, timeout=timeout)
    elapsed = time.time() - t0
    after = _projects(alias)

    log_text = ""
    if os.path.exists(log_path):
        with open(log_path, "r", encoding="utf-8", errors="replace") as fh:
            log_text = fh.read()

    new_ids = sorted(set(after) - set(before))
    result: dict[str, Any] = {
        "action": "import_xer",
        "xer": os.path.abspath(xer),
        "alias": alias.name,
        "exit_code": proc.returncode,
        "elapsed_s": round(elapsed, 1),
        "projects_before": len(before), "projects_after": len(after),
        "new_proj_ids": new_ids,
        "new_projects": {i: after[i] for i in new_ids},
        "log": mc.redact(log_text)[-2000:] if log_text else None,
    }
    if not new_ids:
        result["status"] = "error"
        result["error"] = (
            "Import yeni proje olusturmadi (cikis kodu %s). P6 log'una bakin."
            % proc.returncode)
    else:
        result["proj_id"] = new_ids[-1]
        result["note"] = (
            "P6'nin CLI import'u kaynak ucretlerini SIFIRLAR "
            "(docs/P6_HANDOFF.md §5.2). Maliyetli bir program icin "
            "action='repair_costs' calistirin ve sonucunu dogrulayin.")
    return result


# ---------------------------------------------------------------------------
# cost repair
# ---------------------------------------------------------------------------
def repair_costs(params: Mapping[str, Any]) -> dict[str, Any]:
    """Write resource rates and assignment costs back from the source XER.

    Joins on the stable business keys, never on ids: resources by
    ``rsrc_short_name`` and assignments by (activity code, resource short
    name). P6 renumbers every id on import, so an id join would silently
    match the wrong rows -- the same trap the baseline copy and p6_compare
    each had to avoid.
    """
    import xer_parser

    proj_id = params.get("proj_id")
    xer = params.get("path") or params.get("xer")
    if proj_id is None or not xer:
        raise CliError("proj_id ve path (kaynak XER) zorunlu.")
    proj_id = int(proj_id)
    if not os.path.exists(xer):
        raise CliError("XER bulunamadi: " + xer)
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Kaynak ucreti onarimi")

    x = xer_parser.XerFile(xer, encoding=params.get("encoding"))

    def rows(table):
        return x.tables.get(table, {"rows": []})["rows"]

    def f(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    xer_rsrc = {r.get("rsrc_id"): r for r in rows("RSRC")}
    xer_rate = {}
    for r in rows("RSRCRATE"):
        short = (xer_rsrc.get(r.get("rsrc_id")) or {}).get("rsrc_short_name")
        if short:
            xer_rate.setdefault(short, []).append(r)
    xer_task = {r.get("task_id"): r.get("task_code") for r in rows("TASK")}
    xer_asg = {}
    for a in rows("TASKRSRC"):
        code = xer_task.get(a.get("task_id"))
        short = (xer_rsrc.get(a.get("rsrc_id")) or {}).get("rsrc_short_name")
        if code and short:
            xer_asg[(code, short)] = a

    with w.open_session(params) as s:
        w.project_exists(s, proj_id)

        s.execute("SELECT rr.rsrc_rate_id, r.rsrc_short_name, rr.cost_per_qty, "
                  "rr.start_date FROM RSRCRATE rr JOIN RSRC r "
                  "ON r.rsrc_id = rr.rsrc_id")
        rate_rows = [dict(zip([d[0] for d in s.cur.description], r))
                     for r in s.cur.fetchall()]
        rate_plan = []
        for row in rate_rows:
            cands = xer_rate.get(row["rsrc_short_name"]) or []
            if not cands:
                continue
            want = f(cands[0].get("cost_per_qty"))
            if abs(f(row["cost_per_qty"]) - want) > 1e-9:
                rate_plan.append((int(row["rsrc_rate_id"]),
                                  f(row["cost_per_qty"]), want,
                                  row["rsrc_short_name"]))

        s.execute("SELECT a.taskrsrc_id, t.task_code, r.rsrc_short_name, "
                  "a.target_qty, a.target_cost, a.act_reg_cost, a.remain_cost, "
                  "a.cost_per_qty, a.act_reg_qty, a.remain_qty "
                  "FROM TASKRSRC a JOIN TASK t ON t.task_id = a.task_id "
                  "JOIN RSRC r ON r.rsrc_id = a.rsrc_id "
                  "WHERE a.proj_id = ? AND a.delete_session_id IS NULL", proj_id)
        asg_rows = [dict(zip([d[0] for d in s.cur.description], r))
                    for r in s.cur.fetchall()]
        asg_plan = []
        for row in asg_rows:
            src = xer_asg.get((row["task_code"], row["rsrc_short_name"]))
            if src is None:
                continue
            rate = f(src.get("cost_per_qty"))
            target_cost = f(src.get("target_cost"))
            if target_cost == 0 and rate:
                target_cost = f(row["target_qty"]) * rate
            act_cost = f(row["act_reg_qty"]) * rate if rate else f(src.get("act_reg_cost"))
            remain_cost = f(row["remain_qty"]) * rate if rate else f(src.get("remain_cost"))
            changed = (abs(f(row["cost_per_qty"]) - rate) > 1e-9
                       or abs(f(row["target_cost"]) - target_cost) > 1e-9)
            if changed:
                asg_plan.append((int(row["taskrsrc_id"]), rate, target_cost,
                                 act_cost, remain_cost, row["task_code"]))

        preview = {
            "rates_to_fix": len(rate_plan),
            "assignments_to_fix": len(asg_plan),
            "rate_examples": [{"resource": r[3], "from": r[1], "to": r[2]}
                              for r in rate_plan[:5]],
            "assignment_examples": [{"task_code": a[5], "cost_per_qty": a[1],
                                     "target_cost": a[2]} for a in asg_plan[:5]],
            "source_xer": os.path.abspath(xer),
            "xer_encoding": x.encoding,
        }
        if dry:
            s.conn.rollback()
            return {"action": "repair_costs", "dry_run": True,
                    "proj_id": proj_id, **preview,
                    "note": "Hicbir sey yazilmadi. confirm=true ile calistirin."}

        for rate_id, _old, new, _short in rate_plan:
            s.execute("UPDATE RSRCRATE SET cost_per_qty = ?, update_date = ?, "
                      "update_user = ? WHERE rsrc_rate_id = ?",
                      new, s.stamp, s.user, rate_id)
        for asg_id, rate, target_cost, act_cost, remain_cost, _code in asg_plan:
            s.execute("UPDATE TASKRSRC SET cost_per_qty = ?, target_cost = ?, "
                      "act_reg_cost = ?, remain_cost = ?, update_date = ?, "
                      "update_user = ? WHERE taskrsrc_id = ?",
                      rate, target_cost, act_cost, remain_cost,
                      s.stamp, s.user, asg_id)

        s.execute("SELECT COALESCE(SUM(target_cost), 0) FROM TASKRSRC "
                  "WHERE proj_id = ? AND delete_session_id IS NULL", proj_id)
        total_cost = float(s.cur.fetchone()[0])

    xer_total = sum(f(a.get("target_cost")) for a in rows("TASKRSRC"))
    return {
        "action": "repair_costs", "proj_id": proj_id, **preview,
        "rates_fixed": len(rate_plan), "assignments_fixed": len(asg_plan),
        "project_target_cost_after": round(total_cost, 2),
        "source_xer_target_cost": round(xer_total, 2),
        "matches_source": abs(total_cost - xer_total) < 0.01,
        "note": ("Kaynak XER'in toplam target_cost'u ile veritabanindaki toplam "
                 "karsilastirildi; 'matches_source' bunun sonucudur."),
    }


ACTIONS = {
    "import_xer": import_xer,
    "repair_costs": repair_costs,
}
