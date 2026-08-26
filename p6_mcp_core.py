"""Oracle Primavera P6 Professional MCP server.

Fourth server in this repo, alongside asta_mcp_core (Asta COM),
asta_mcp_file (Asta/MSPDI files) and msproject_mcp_core (MS Project).

P6 differs from the other two applications: PM.exe exposes no automation
interface and its command line can only import/export/run report batches. The
only way to run P6's own CPM engine headlessly is its Job Service, driven
through the JOBSVC queue table -- see p6/jobs.py.

Thin by design: this file holds the tool surface and dispatch only; every bit
of logic lives in the p6 package and in the shared compute modules
(xer_parser, dcma_checks, evm_math, ...).
"""
from __future__ import annotations

import logging
import os
import re
import sys
from typing import Any, Mapping

from mcp.server.fastmcp import FastMCP

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import mcp_common as mc  # noqa: E402
from p6 import baseline, compare, db, evm, health, jobs, progress  # noqa: E402
from p6 import cli, writer  # noqa: E402
from p6 import source as src  # noqa: E402

# stderr only -- stdout belongs to the MCP stdio protocol
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(os.path.expanduser("~/p6_mcp_core.log")),
              logging.StreamHandler(sys.stderr)],
)
log = logging.getLogger("p6_mcp")

mcp = FastMCP(
    "p6_mcp",
    instructions=(
        "Oracle Primavera P6 Professional tools. P6 has no COM automation and "
        "its CLI cannot schedule; reschedule (F9) runs through the P6 Job "
        "Service queue. Use p6_job with action='schedule'. Always run "
        "action='preflight' first when something fails -- it checks the three "
        "things that actually break: SQLite alias (unsupported), stopped "
        "PrmJobSv service, and a missing USEROBS row for the job user."
    ),
)

DEFAULT_USER_ID = int(os.environ.get("P6_USER_ID", "25"))


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _alias(params: Mapping[str, Any]):
    return db.resolve_alias(params.get("alias"))


def _cursor(params: Mapping[str, Any]):
    """Writable connection. Only p6/jobs.py uses it, and only for JOBSVC."""
    alias = _alias(params)
    conn = db.connect_rw(alias)
    return alias, conn, conn.cursor()


def _proj_ids(params: Mapping[str, Any]) -> list[int]:
    raw = params.get("proj_ids")
    if raw is None:
        raw = params.get("proj_id")
    if raw is None:
        raise jobs.P6JobError("proj_id veya proj_ids zorunlu.")
    if isinstance(raw, (int, str)):
        raw = [raw]
    return [int(x) for x in raw]


def _user_id(params: Mapping[str, Any]) -> int:
    return int(params.get("user_id", DEFAULT_USER_ID))


# ---------------------------------------------------------------------------
# actions
# ---------------------------------------------------------------------------
def _act_preflight(params: Mapping[str, Any]) -> dict[str, Any]:
    alias, conn, cur = _cursor(params)
    try:
        ids = []
        try:
            ids = _proj_ids(params)
        except jobs.P6JobError:
            pass
        return jobs.preflight(cur, alias, _user_id(params), ids)
    finally:
        conn.close()


def _run(params: Mapping[str, Any], job_type: str) -> dict[str, Any]:
    alias, conn, cur = _cursor(params)
    try:
        ids = _proj_ids(params)
        user = _user_id(params)
        pre = jobs.preflight(cur, alias, user, ids)
        if not pre["ready"]:
            return {"status": "error",
                    "error": "On kontrol gecmedi; is gonderilmedi.",
                    "preflight": pre}
        if job_type in jobs.MUTATING and not params.get("confirm"):
            return {"status": "error",
                    "error": ("'%s' proje verisini degistirir ve geri alinamaz. "
                              "confirm=true ile tekrar cagirin." % job_type),
                    "job_type": job_type}
        name = params.get("job_name") or ("MCP_" + job_type.replace("JT_", ""))
        timeout = int(params.get("timeout_s", 300))
        if params.get("wait", True):
            res = jobs.run_and_wait(cur, job_type, ids, user, name, timeout,
                                    params.get("default_proj_id"))
            out = res.to_dict()
            out["alias"] = alias.name
            out["proj_ids"] = ids
            if not res.ok:
                out["status"] = "error"
                out["error"] = res.error_tr or res.status
            return out
        job_id = jobs.submit(cur, job_type, ids, user, name,
                             params.get("default_proj_id"))
        return {"status": "submitted", "job_id": job_id, "job_type": job_type,
                "alias": alias.name, "proj_ids": ids,
                "note": "wait=false verildi; durumu p6_job action='status' ile izleyin."}
    finally:
        conn.close()


def _act_status(params: Mapping[str, Any]) -> dict[str, Any]:
    _alias_, conn, cur = _cursor(params)
    try:
        job_id = int(params["job_id"])
        status, err, run = jobs.read_status(cur, job_id)
        return {"job_id": job_id, "status": status,
                "error": err or None,
                "error_tr": jobs.translate_error(err) or None,
                "last_run": str(run) if run else None,
                "log": jobs.read_log(cur, job_id) or None}
    finally:
        conn.close()


def _act_wait(params: Mapping[str, Any]) -> dict[str, Any]:
    _alias_, conn, cur = _cursor(params)
    try:
        res = jobs.wait(cur, int(params["job_id"]),
                        int(params.get("timeout_s", 300)))
        return res.to_dict()
    finally:
        conn.close()


def _act_list(params: Mapping[str, Any]) -> dict[str, Any]:
    _alias_, conn, cur = _cursor(params)
    try:
        rows = jobs.list_jobs(cur, int(params.get("limit", 25)))
        return {"count": len(rows), "jobs": rows}
    finally:
        conn.close()


def _act_cancel(params: Mapping[str, Any]) -> dict[str, Any]:
    _alias_, conn, cur = _cursor(params)
    try:
        return jobs.cancel(cur, int(params["job_id"]))
    finally:
        conn.close()


def _act_purge(params: Mapping[str, Any]) -> dict[str, Any]:
    _alias_, conn, cur = _cursor(params)
    try:
        n = jobs.purge(cur, params.get("name_like", "MCP\\_%"))
        return {"deleted": n,
                "note": "Yalnizca MCP_ ile baslayan ve bitmis isler silinir."}
    finally:
        conn.close()


def _act_service_health(params: Mapping[str, Any]) -> dict[str, Any]:
    alias = _alias(params)
    out: dict[str, Any] = {
        "alias": alias.name, "driver": alias.driver,
        "database": alias.database, "host": alias.host,
        "service": jobs.service_state(),
    }
    if alias.driver == "SQLite":
        out["blocked"] = ("SQLite alias: Job Service CALISMAZ. SQL Server veya "
                          "Oracle alias'i gerekir.")
        return out
    try:
        conn = db.connect_rw(alias)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM JOBSVC WHERE status_code IN (?, ?)",
                    jobs.JS_PENDING, jobs.JS_RUNNING)
        out["queue_active"] = int(cur.fetchone()[0])
        cur.execute("SELECT COUNT(*) FROM USEROBS WHERE user_id = ?",
                    _user_id(params))
        out["userobs_rows"] = int(cur.fetchone()[0])
        conn.close()
        out["db_reachable"] = True
    except Exception as exc:  # noqa: BLE001
        out["db_reachable"] = False
        out["db_error"] = str(exc)
    return out


def _act_job_data(params: Mapping[str, Any]) -> dict[str, Any]:
    """Show the JOB_DATA blob without submitting anything."""
    job_type = params.get("job_type", jobs.JT_SCHEDULE)
    ids = _proj_ids(params)
    return {"job_type": job_type, "proj_ids": ids,
            "job_data": jobs.build_job_data(job_type, ids,
                                            params.get("default_proj_id"))}


_ACTIONS = {
    "preflight": _act_preflight,
    "service_health": _act_service_health,
    "job_data": _act_job_data,
    "schedule": lambda p: _run(p, jobs.JT_SCHEDULE),
    "level": lambda p: _run(p, jobs.JT_LEVEL),
    "summarize": lambda p: _run(p, jobs.JT_SUMMARIZE),
    "apply_actuals": lambda p: _run(p, jobs.JT_APPLY_ACTUALS),
    "update_baseline": lambda p: _run(p, jobs.JT_UPDATE_BASELINE),
    "status": _act_status,
    "wait": _act_wait,
    "list": _act_list,
    "cancel": _act_cancel,
    "purge": _act_purge,
}


# ---------------------------------------------------------------------------
# p6_query -- read-only. Never writes, never guesses.
# ---------------------------------------------------------------------------
_SELECT_OK = re.compile(r"^\s*(select|with)\b", re.I)
_SQL_FORBIDDEN = re.compile(
    r"\b(insert|update|delete|drop|alter|create|truncate|merge|exec|execute|"
    r"grant|revoke|backup|restore|shutdown|xp_|sp_)\b", re.I)


def _q_list_projects(params: Mapping[str, Any]) -> dict[str, Any]:
    flag = "N" if params.get("action") == "list_eps" else "Y"
    with src.OpenSource(params) as s:
        if s.meta.get("type") != "db":
            raise src.SourceError("list_projects yalnizca 'db' kaynaginda calisir.")
        rows = db.list_projects(s.backend, flag)
        return {"count": len(rows), "project_flag": flag,
                "alias": s.meta.get("alias"), "projects": rows}


def _reader(params: Mapping[str, Any], name: str) -> dict[str, Any]:
    """Delegate to the xer_parser reader, from db or xer alike."""
    with src.OpenSource(params) as s:
        s.require_project()
        limit = src.clamp_limit(params)
        bag = s.bag
        if name == "read_tasks":
            data = db.read_tasks(bag, s.day_hr_cnt)
        elif name == "read_links":
            data = db.read_links(bag, s.day_hr_cnt)
        elif name == "read_resources":
            data = db.read_resources(bag)
        elif name == "read_assignments":
            data = db.read_assignments(bag)
        elif name == "read_calendars":
            data = db.read_calendars(bag)
        elif name == "read_wbs":
            data = db.read_wbs(bag)
        elif name == "read_project":
            data = db.read_project(bag)
        elif name == "read_progress":
            data = bag.read_progress() if hasattr(bag, "read_progress") \
                else __import__("xer_parser").XerFile.read_progress(bag)
        else:
            raise src.SourceError("Bilinmeyen okuma: " + name)

        out: dict[str, Any] = {"action": name, "source": s.meta}
        if isinstance(data, list):
            out["count"] = len(data)
            out["truncated"] = len(data) > limit
            out["items"] = data[:limit]
        else:
            out["data"] = data
        return out


def _q_finish_drivers(params: Mapping[str, Any]) -> dict[str, Any]:
    import xer_drivers

    with src.OpenSource(params) as s:
        s.require_project()
        tasks = db.read_tasks(s.bag, s.day_hr_cnt)
        wbs = db.read_wbs(s.bag)
        res = xer_drivers.forecast_drivers(
            tasks, wbs,
            anomaly_gap_days=int(params.get("anomaly_gap_days", 30)),
            top_n=int(params.get("top_n", 5)))
        return {"action": "finish_drivers", "source": s.meta, "result": res}


def _q_schedule_options(params: Mapping[str, Any]) -> dict[str, Any]:
    with src.OpenSource(params) as s:
        s.require_project()
        return {"action": "schedule_options", "source": s.meta,
                **db.parse_schedule_options(s.bag)}


def _q_sql(params: Mapping[str, Any]) -> dict[str, Any]:
    sql = (params.get("sql") or params.get("query") or "").strip()
    if not sql:
        raise src.SourceError("'sql' zorunlu.")
    if not _SELECT_OK.match(sql):
        raise src.SourceError("Yalnizca SELECT/WITH sorgulari calistirilabilir.")
    if ";" in sql.rstrip(";"):
        raise src.SourceError("Tek ifade calistirilabilir; ';' ile ayirmayin.")
    if _SQL_FORBIDDEN.search(sql):
        raise src.SourceError("Sorguda yazma/DDL anahtar kelimesi var; reddedildi.")
    limit = src.clamp_limit(params)
    with src.OpenSource(params) as s:
        if s.meta.get("type") != "db":
            raise src.SourceError("sql yalnizca 'db' kaynaginda calisir.")
        cols, rows = s.backend.select_named(sql, [])
        items = [list(r) for r in rows[:limit]]
        return {"action": "sql", "source": {"alias": s.meta.get("alias"),
                                            "driver": s.meta.get("driver")},
                "count": len(rows), "truncated": len(rows) > limit,
                "columns": cols, "rows": items}


def _q_db_info(params: Mapping[str, Any]) -> dict[str, Any]:
    with src.OpenSource(params) as s:
        if s.meta.get("type") != "db":
            return {"action": "db_info", "source": s.meta}
        be = s.backend
        info: dict[str, Any] = {"action": "db_info", "source": s.meta}
        counts = {}
        for t in ("PROJECT", "PROJWBS", "TASK", "TASKPRED", "TASKRSRC",
                  "RSRC", "CALENDAR", "JOBSVC", "USEROBS"):
            try:
                qo, qc = be.quote_open, be.quote_close
                rows = be.select("SELECT COUNT(*) FROM " + qo + t + qc, [])
                counts[t] = int(rows[0][0])
            except Exception:  # noqa: BLE001
                counts[t] = None
        info["table_rows"] = counts
        return info


_QUERY_ACTIONS = {
    "list_projects": _q_list_projects,
    "list_eps": _q_list_projects,
    "read_tasks": lambda p: _reader(p, "read_tasks"),
    "read_links": lambda p: _reader(p, "read_links"),
    "read_resources": lambda p: _reader(p, "read_resources"),
    "read_assignments": lambda p: _reader(p, "read_assignments"),
    "read_calendars": lambda p: _reader(p, "read_calendars"),
    "read_wbs": lambda p: _reader(p, "read_wbs"),
    "read_project": lambda p: _reader(p, "read_project"),
    "read_progress": lambda p: _reader(p, "read_progress"),
    "finish_drivers": _q_finish_drivers,
    "schedule_options": _q_schedule_options,
    "sql": _q_sql,
    "db_info": _q_db_info,
}


@mcp.tool(
    name="p6_query",
    description=(
        "P6 read-only queries, from a database alias or an XER file.\n"
        "actions: list_projects | list_eps | read_tasks | read_links | "
        "read_resources | read_assignments | read_calendars | read_wbs | "
        "read_project | read_progress | finish_drivers | schedule_options | "
        "sql | db_info\n"
        "params: type ('db' default, or 'xer'), path (for xer), alias, "
        "proj_id or proj_short_name, day_hr_cnt, limit (default 100, max 500), "
        "sql (SELECT only).\n"
        "Hours-per-day is always read from the project calendar and reported as "
        "day_hr_cnt + day_hr_cnt_source; it is never assumed. Soft-deleted rows "
        "are excluded. PMXML is not supported on purpose -- MPXJ mis-reads P6 "
        "data, so use XER."
    ),
    annotations={"readOnlyHint": True},
)
def p6_query(params: dict) -> str:
    return mc.dispatch("p6_query", params or {}, _QUERY_ACTIONS)


@mcp.tool(
    name="p6_job",
    description=(
        "P6 Job Service: headless reschedule (F9) and the other queued jobs.\n"
        "actions: schedule | level | summarize | apply_actuals | update_baseline "
        "| status | wait | list | cancel | purge | preflight | service_health | "
        "job_data\n"
        "params: proj_id or proj_ids (int/list), alias, user_id, job_name, "
        "wait (default true), timeout_s (default 300), default_proj_id, "
        "confirm (required for data-changing jobs), job_id, limit.\n"
        "'schedule' runs P6's own CPM engine and typically finishes in under "
        "10 s. Requires a SQL Server/Oracle alias, a running PrmJobSv service "
        "and a USEROBS row for the user -- run 'preflight' if a job fails."
    ),
)
def p6_job(params: dict) -> str:
    return mc.dispatch("p6_job", params or {}, _ACTIONS)


@mcp.tool(
    name="p6_health",
    description=(
        "DCMA 14-Point schedule health for a P6 project, from a database "
        "alias or an XER file.\n"
        "actions: assess_all | summary | drill_down | compare\n"
        "params: type ('db' default, or 'xer'), path, alias, proj_id or "
        "proj_short_name, baseline_proj_id, status_date, day_hr_cnt, "
        "rule_id (1-14, drill_down), limit, snapshot_path (compare).\n"
        "Rule 9's 44-day threshold uses the project calendar's hours per day, "
        "and 'critical' comes from PROJECT.critical_drtn_hr_cnt -- neither is "
        "assumed. Without baseline_proj_id the target dates act as the "
        "baseline; every response says which was used."
    ),
    annotations={"readOnlyHint": True},
)
def p6_health(params: dict) -> str:
    return mc.dispatch("p6_health", params or {}, health.ACTIONS)


@mcp.tool(
    name="p6_evm",
    description=(
        "Earned Value Management for a P6 project (PMBOK 8th 7.4.2, Lipke "
        "2003 Earned Schedule). Same action vocabulary as msproject_evm.\n"
        "actions: compute_metrics | forecast | earned_schedule | summary | "
        "time_phased_evm | period_delta | progress_data_quality | "
        "variance_to_baseline | compare_baselines_evm | save_period_snapshot "
        "| get_period_history | trend | detect_currency_mode | "
        "validate_currency_mode | verify\n"
        "params: type ('db' default, or 'xer'), path, alias, proj_id or "
        "proj_short_name, baseline_proj_id (real P6 baseline), "
        "baseline_proj_id_a/_b, units ('auto'|'cost'|'qty'|'duration_h'), "
        "status_date, bucket ('day'|'week'|'month'), snapshot_path, tag, "
        "limit, tolerance.\n"
        "Every response reports the unit the BAC is measured in plus the BAC "
        "each basis would give. Run 'verify' before quoting a BAC in a report.\n"
        "The P6 database is never written to; 'save_period_snapshot' appends "
        "to a local JSON file (~/p6_evm_snapshots.json by default)."
    ),
)
def p6_evm(params: dict) -> str:
    return mc.dispatch("p6_evm", params or {}, evm.ACTIONS)


@mcp.tool(
    name="p6_progress",
    description=(
        "Enter progress and actuals into a P6 project, and read back what is "
        "recorded.\n"
        "actions: read | set_progress | set_assignment_actuals | clear | "
        "set_data_date\n"
        "params: proj_id, updates (list of {task_code or task_id, status, "
        "actual_start, actual_finish, percent_complete, "
        "remaining_duration_h}), dry_run, confirm (required to write), "
        "schedule (run F9 after), allow_future_actuals, data_date, "
        "task_codes (clear), limit, only_started (read).\n"
        "Writes keep P6's fields consistent: status_code matches the actual "
        "dates, remaining duration and units fall to zero on completion, and "
        "percent complete is written on the basis the activity actually uses "
        "(complete_pct_type). Actual dates after the data date are refused, "
        "as P6 refuses them. Dates are NOT recalculated here -- P6's own CPM "
        "engine does that via the Job Service.\n"
        "Always run with dry_run=true first: it shows the before/after of "
        "every field without writing."
    ),
)
def p6_progress(params: dict) -> str:
    return mc.dispatch("p6_progress", params or {}, progress.ACTIONS)


@mcp.tool(
    name="p6_baseline",
    description=(
        "P6 baselines: list, create, assign, delete.\n"
        "actions: list | create | assign | delete\n"
        "params: proj_id, baseline_proj_id, base_type (default 'Initial "
        "Plan') or base_type_id, baseline_name, assign (default true), "
        "dry_run, confirm (required to write).\n"
        "A P6 baseline is a full copy of the project stored as its own "
        "PROJECT row -- 'create' copies PROJECT/PROJPROP/PROJWBS/TASK/"
        "TASKPRED/TASKRSRC in one transaction with fresh ids. Feed the "
        "resulting baseline_proj_id to p6_evm action='variance_to_baseline'; "
        "without one, EVM compares the schedule against its own planned "
        "dates, which in a live P6 database are a copy of the current "
        "schedule."
    ),
)
def p6_baseline(params: dict) -> str:
    return mc.dispatch("p6_baseline", params or {}, baseline.ACTIONS)


@mcp.tool(
    name="p6_compare",
    description=(
        "Compare two P6 schedules -- revisions, a project against one of its "
        "baselines, or two XER files.\n"
        "actions: summary | tasks | links | progress | evm\n"
        "params: a and b, each a source ({proj_id: ...} or "
        "{baseline_proj_id: ...} or {type: 'xer', path: '...'}); or the flat "
        "form proj_id_a / baseline_proj_id_b / path_b. Also fields (task "
        "fields to diff), limit, units, alias.\n"
        "Activities are matched on task_code, never task_id: P6 renumbers ids "
        "on every import and baseline copy, so an id-based diff would report "
        "every activity as both removed and added. Differing units, calendars "
        "or data dates between the two sides are reported as warnings rather "
        "than silently subtracted."
    ),
    annotations={"readOnlyHint": True},
)
def p6_compare(params: dict) -> str:
    return mc.dispatch("p6_compare", params or {}, compare.ACTIONS)


def _act_export_xer(params: Mapping[str, Any]) -> dict[str, Any]:
    result = writer.write_xer(params)
    if params.get("verify", True):
        result["verify"] = writer.verify_roundtrip(result["path"],
                                                   result["tables"])
    return result


_WRITE_ACTIONS = {"export_xer": _act_export_xer}


@mcp.tool(
    name="p6_write",
    description=(
        "Write a P6 project out to an XER file.\n"
        "actions: export_xer\n"
        "params: proj_id, path, overwrite, alias, tables (default: P6's own "
        "export order), verify (default true).\n"
        "P6's headless export job (JT_XERExport) reaches P6's export code but "
        "fails with 'File name not specified.' and the CLI action script has "
        "no export elements, so this writes the file directly from the "
        "database. UTF-16-LE with a BOM -- the one encoding proven to carry "
        "Cyrillic through the round trip. With verify=true the file is read "
        "back with our own parser and the row counts are checked against what "
        "was written; a baseline copy is exported the same way by passing its "
        "proj_id."
    ),
)
def p6_write(params: dict) -> str:
    return mc.dispatch("p6_write", params or {}, _WRITE_ACTIONS)


@mcp.tool(
    name="p6_cli",
    description=(
        "Drive P6's own command line: import an XER, then repair what it "
        "drops.\n"
        "actions: import_xer | repair_costs\n"
        "params: path (XER), proj_id (repair), alias, user, eps, password or "
        "the P6_CLI_PASSWORD environment variable, confirm (required), "
        "dry_run (repair), timeout_s, work_dir.\n"
        "The CLI only ever CREATEs, so every import lands as a NEW project -- "
        "which is what makes it the right tool for a revision. It also zeroes "
        "resource rates, because it runs with no import configuration "
        "(importConfiguration resolves to VIEWPROP rows of type VP_IMP_OPT "
        "and this database has none); 'repair_costs' writes them back from "
        "the source XER, joining on activity code and resource short name, "
        "never on ids, and reports whether the project total now matches the "
        "file. PM.exe must be closed. Passwords are registered as secrets and "
        "masked out of every log this tool returns."
    ),
)
def p6_cli(params: dict) -> str:
    return mc.dispatch("p6_cli", params or {}, cli.ACTIONS)


if __name__ == "__main__":
    log.info("p6_mcp starting")
    mcp.run()
