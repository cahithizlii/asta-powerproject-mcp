"""P6 baselines: list, create, assign, delete.

A P6 baseline is not a flag on the schedule -- it is a *separate copy of the
whole project* stored as another PROJECT row whose ``orig_proj_id`` points
back at the live one. That is why ``p6_evm action='variance_to_baseline'``
needs a ``baseline_proj_id`` and why comparing a live project against its own
"planned" dates is meaningless: in a live database P6 keeps planned dates in
step with the current schedule for anything not yet started, so the schedule
would be measured against itself (see docs/P6_HANDOFF.md section 5.3).

P6 Professional has no automation interface and its Job Service has no
"create baseline" job type, so the copy is done here, in one transaction,
with fresh ids drawn from NEXTKEY. Activity codes match across the copy by
``task_code``, which is what the EVM comparison joins on.

Copied tables: PROJECT, PROJPROP, PROJWBS, TASK, TASKPRED, TASKRSRC -- the
schedule itself. Every response says exactly what was copied and what was
not, so nobody has to assume.
"""
from __future__ import annotations

import datetime as _dt
from typing import Any, Mapping

from . import write as w

# NEXTKEY names are '<table>_<column>' in P6's own convention.
KEY_NAMES = {
    "PROJECT": ("proj_id", "project_proj_id"),
    "PROJWBS": ("wbs_id", "projwbs_wbs_id"),
    "TASK": ("task_id", "task_task_id"),
    "TASKPRED": ("task_pred_id", "taskpred_task_pred_id"),
    "TASKRSRC": ("taskrsrc_id", "taskrsrc_taskrsrc_id"),
}

COPY_TABLES = ("PROJPROP", "PROJWBS", "TASK", "TASKPRED", "TASKRSRC")

DEFAULT_BASE_TYPE = "Initial Plan"


def _base_types(session) -> dict[str, int]:
    session.execute("SELECT base_type_id, base_type FROM BASETYPE "
                    "WHERE delete_session_id IS NULL")
    return {r[1]: int(r[0]) for r in session.cur.fetchall()}


def _resolve_base_type(session, params: Mapping[str, Any]) -> tuple[int, str]:
    types = _base_types(session)
    if params.get("base_type_id") is not None:
        wanted = int(params["base_type_id"])
        for name, tid in types.items():
            if tid == wanted:
                return tid, name
        raise w.P6WriteError("base_type_id %s BASETYPE'ta yok. Secenekler: %s"
                             % (wanted, sorted(types)))
    name = params.get("base_type") or DEFAULT_BASE_TYPE
    if name not in types:
        raise w.P6WriteError("base_type '%s' yok. Secenekler: %s"
                             % (name, sorted(types)))
    return types[name], name


# ---------------------------------------------------------------------------
# actions
# ---------------------------------------------------------------------------
def list_baselines(params: Mapping[str, Any]) -> dict[str, Any]:
    """The baselines of a project, and which one is the project baseline."""
    from . import db as p6db

    alias = p6db.resolve_alias(params.get("alias"))
    backend, info = p6db.open_backend(alias, use_snapshot=False)
    try:
        qo, qc, p = backend.quote_open, backend.quote_close, backend.param
        proj_id = params.get("proj_id")
        where = "WHERE " + qo + "delete_session_id" + qc + " IS NULL"
        args: list[Any] = []
        if proj_id is not None:
            where += (" AND (" + qo + "orig_proj_id" + qc + " = " + p
                      + " OR " + qo + "proj_id" + qc + " = " + p + ")")
            args = [int(proj_id), int(proj_id)]
        rows = backend.select(
            "SELECT " + ", ".join(
                qo + c + qc for c in ("proj_id", "proj_short_name", "project_flag",
                                      "orig_proj_id", "base_type_id",
                                      "sum_base_proj_id", "last_recalc_date",
                                      "last_baseline_update_date"))
            + " FROM " + qo + "PROJECT" + qc + " " + where, args)
        types = {}
        for r in backend.select("SELECT " + qo + "base_type_id" + qc + ", "
                                + qo + "base_type" + qc + " FROM "
                                + qo + "BASETYPE" + qc, []):
            types[r[0]] = r[1]

        live, baselines = [], []
        for r in rows:
            item = {
                "proj_id": r[0], "proj_short_name": r[1], "project_flag": r[2],
                "orig_proj_id": r[3],
                "base_type_id": r[4], "base_type": types.get(r[4]),
                "assigned_project_baseline": r[5],
                "data_date": str(r[6])[:10] if r[6] else None,
                "created": str(r[7])[:19] if r[7] else None,
            }
            (baselines if r[3] else live).append(item)
        return {"action": "list", "source": {"alias": alias.name},
                "projects": live, "baseline_count": len(baselines),
                "baselines": baselines,
                "base_types": sorted(types.values())}
    finally:
        backend.close()


def revision(params: Mapping[str, Any]) -> dict[str, Any]:
    """Copy a project into a new, independent project -- a revision.

    Why in the database and not through an export/import round trip: P6's CLI
    importer **rejects a UTF-16LE XER** ("The import file is invalid.", exit
    code 6) and reads an ANSI one with the machine's own code page, so a
    Cyrillic programme comes back mangled on a Turkish-locale machine
    (measured: bukhtourcity437 "Гранит" -> "Agaieo"). The import also zeroes
    resource rates. A copy inside the database has none of those losses --
    the same machinery that makes a baseline, with the copy left as a real
    project instead.
    """
    return create(dict(params, _as_revision=True))


def create(params: Mapping[str, Any]) -> dict[str, Any]:
    """Copy a project into a new baseline of it.

    dry_run reports exactly what would be copied without writing anything.
    """
    as_revision = bool(params.get("_as_revision"))
    what = "Revizyon olusturma" if as_revision else "Baseline olusturma"
    proj_id = params.get("proj_id")
    if proj_id is None:
        raise w.P6WriteError("proj_id zorunlu.")
    proj_id = int(proj_id)
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, what)

    with w.open_session(params) as s:
        short_name = w.project_exists(s, proj_id)
        if as_revision:
            base_type_id, base_type = None, None
        else:
            base_type_id, base_type = _resolve_base_type(s, params)

        counts = {}
        for table in COPY_TABLES:
            counts[table] = int(s.scalar(
                "SELECT COUNT(*) FROM [%s] WHERE proj_id = ? "
                "AND delete_session_id IS NULL" % table, proj_id) or 0)

        default_name = "%s - %s%s" % (
            short_name, "R" if as_revision else "B",
            _dt.datetime.now().strftime("%Y%m%d-%H%M"))
        name = ((params.get("revision_name") if as_revision else None)
                or params.get("baseline_name") or default_name)
        action_name = "revision" if as_revision else "create"

        if dry:
            s.conn.rollback()
            return {"action": action_name, "dry_run": True,
                    "source_proj_id": proj_id, "source_name": short_name,
                    "new_name": name, "base_type": base_type,
                    "would_copy": counts,
                    "note": "Hicbir sey yazilmadi. confirm=true ile calistirin."}

        new_proj_id = s.reserve(KEY_NAMES["PROJECT"][1], 1)[0]

        # --- PROJECT row -------------------------------------------------
        pcols, prows = s.select_rows("PROJECT", "proj_id = ?", [proj_id])
        prow = dict(prows[0])
        prow["proj_id"] = new_proj_id
        prow["proj_short_name"] = name[:40] if "proj_short_name" in pcols else None
        # A baseline hangs off its project and stays out of the EPS; a
        # revision is a project in its own right and must appear there.
        prow["orig_proj_id"] = None if as_revision else proj_id
        prow["base_type_id"] = base_type_id
        prow["sum_base_proj_id"] = None
        prow["project_flag"] = "Y" if as_revision else "N"
        if not as_revision and "last_baseline_update_date" in pcols:
            prow["last_baseline_update_date"] = s.stamp
        s.stamp_audit(prow, pcols)
        s.insert_rows("PROJECT", pcols, [prow])

        # --- id maps -----------------------------------------------------
        maps: dict[str, dict[Any, Any]] = {}
        table_rows: dict[str, tuple[list[str], list[dict]]] = {}
        for table in COPY_TABLES:
            cols, rows = s.select_rows(
                table, "proj_id = ? AND delete_session_id IS NULL", [proj_id])
            table_rows[table] = (cols, rows)
            if table in KEY_NAMES:
                pk, key_name = KEY_NAMES[table]
                new_ids = s.reserve(key_name, len(rows))
                maps[pk] = {r[pk]: nid for r, nid in zip(rows, new_ids)}

        wbs_map = maps.get("wbs_id", {})
        task_map = maps.get("task_id", {})

        # --- copy rows ----------------------------------------------------
        written = {}
        for table in COPY_TABLES:
            cols, rows = table_rows[table]
            out = []
            for r in rows:
                row = dict(r)
                row["proj_id"] = new_proj_id
                if table in KEY_NAMES:
                    pk = KEY_NAMES[table][0]
                    row[pk] = maps[pk][r[pk]]
                if "wbs_id" in cols and r.get("wbs_id") is not None:
                    row["wbs_id"] = wbs_map.get(r["wbs_id"], r["wbs_id"])
                if table == "PROJWBS" and r.get("parent_wbs_id") is not None:
                    # The project root's parent is an EPS node outside the
                    # project; leave that one pointing where it pointed.
                    row["parent_wbs_id"] = wbs_map.get(r["parent_wbs_id"],
                                                       r["parent_wbs_id"])
                if table == "TASKPRED":
                    row["task_id"] = task_map.get(r["task_id"], r["task_id"])
                    row["pred_task_id"] = task_map.get(r["pred_task_id"],
                                                       r["pred_task_id"])
                    if r.get("pred_proj_id") == proj_id:
                        row["pred_proj_id"] = new_proj_id
                if table == "TASKRSRC" and r.get("task_id") is not None:
                    row["task_id"] = task_map.get(r["task_id"], r["task_id"])
                s.stamp_audit(row, cols)
                out.append(row)
            written[table] = s.insert_rows(table, cols, out)

        # Cross-project links whose other end is outside this project cannot
        # be remapped; say so rather than let the caller assume a clean copy.
        _cols, pred_rows = table_rows["TASKPRED"]
        external = sum(1 for r in pred_rows
                       if r.get("pred_task_id") not in task_map)

        if not as_revision and params.get("assign", True):
            s.execute("UPDATE PROJECT SET sum_base_proj_id = ?, "
                      "last_baseline_update_date = ? WHERE proj_id = ?",
                      new_proj_id, s.stamp, proj_id)

        return {
            "action": action_name,
            "baseline_proj_id": None if as_revision else new_proj_id,
            "revision_proj_id": new_proj_id if as_revision else None,
            "new_proj_id": new_proj_id,
            "new_name": name,
            "base_type": base_type, "base_type_id": base_type_id,
            "source_proj_id": proj_id, "source_name": short_name,
            "copied": written,
            "external_predecessors_left_unmapped": external,
            "assigned_as_project_baseline": (
                False if as_revision else bool(params.get("assign", True))),
            "not_copied": ["OBSPROJ", "UACCESS", "TASKMEMO", "UDFVALUE",
                           "ACTVCODE atamalari", "PROJCOST"],
            "note": ("Program tablolari kopyalandi (PROJECT/PROJPROP/PROJWBS/"
                     "TASK/TASKPRED/TASKRSRC). Erisim, not ve UDF tablolari "
                     "kopyalanmadi -- tarih/EVM karsilastirmasi icin gerekli "
                     "degil."),
        }


def assign(params: Mapping[str, Any]) -> dict[str, Any]:
    """Point a project's 'project baseline' at one of its baselines."""
    proj_id = int(params["proj_id"])
    baseline_proj_id = params.get("baseline_proj_id")
    w.require_confirm(params, "Baseline atama")
    with w.open_session(params) as s:
        w.project_exists(s, proj_id)
        if baseline_proj_id is not None:
            baseline_proj_id = int(baseline_proj_id)
            owner = s.scalar("SELECT orig_proj_id FROM PROJECT WHERE proj_id = ?",
                             baseline_proj_id)
            if owner is None:
                raise w.P6WriteError(
                    "%s bir baseline degil (orig_proj_id bos)." % baseline_proj_id)
            if int(owner) != proj_id:
                raise w.P6WriteError(
                    "Baseline %s, %s projesine ait degil (sahibi: %s)."
                    % (baseline_proj_id, proj_id, owner))
        s.execute("UPDATE PROJECT SET sum_base_proj_id = ? WHERE proj_id = ?",
                  baseline_proj_id, proj_id)
        return {"action": "assign", "proj_id": proj_id,
                "project_baseline": baseline_proj_id,
                "cleared": baseline_proj_id is None}


def delete(params: Mapping[str, Any]) -> dict[str, Any]:
    """Soft-delete a baseline the way P6 does -- stamp delete_session_id."""
    baseline_proj_id = int(params["baseline_proj_id"])
    w.require_confirm(params, "Baseline silme")
    with w.open_session(params) as s:
        owner = s.scalar("SELECT orig_proj_id FROM PROJECT WHERE proj_id = ? "
                         "AND delete_session_id IS NULL", baseline_proj_id)
        if owner is None:
            raise w.P6WriteError(
                "%s bir baseline degil ya da zaten silinmis." % baseline_proj_id)
        session_id = s.reserve("usession_session_id", 1)[0]
        removed = {}
        for table in ("TASKRSRC", "TASKPRED", "TASK", "PROJWBS", "PROJPROP",
                      "PROJECT"):
            cur = s.execute(
                "UPDATE [%s] SET delete_session_id = ?, delete_date = ? "
                "WHERE proj_id = ? AND delete_session_id IS NULL" % table,
                session_id, s.stamp, baseline_proj_id)
            removed[table] = cur.rowcount
        s.execute("UPDATE PROJECT SET sum_base_proj_id = NULL "
                  "WHERE sum_base_proj_id = ?", baseline_proj_id)
        return {"action": "delete", "baseline_proj_id": baseline_proj_id,
                "owner_proj_id": owner, "delete_session_id": session_id,
                "soft_deleted_rows": removed,
                "note": "P6 gibi yumusak silme: satirlar delete_session_id ile "
                        "isaretlendi, fiziksel olarak silinmedi."}


ACTIONS = {
    "list": list_baselines,
    "create": create,
    "revision": revision,
    "assign": assign,
    "delete": delete,
}
