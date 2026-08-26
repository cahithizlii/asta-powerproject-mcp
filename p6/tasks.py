"""Task / link / WBS / assignment CRUD for a P6 project -- Phase 6.

This is the piece that turns the server from "analyse and progress an
existing programme" into "build one": create a project from nothing, add WBS
nodes, activities, relationships and resource assignments, and let P6's own
CPM engine compute every date via the Job Service (p6_job action='schedule').
No date is ever calculated here -- new activities are written without early/
late dates on purpose, exactly like an activity added in the P6 client before
the first F9.

Structural defaults are not guessed. For a project that already has
activities, each structural column (calendar, duration type, percent-complete
type, ...) defaults to the *modal* value of that project's own activities --
new work follows the project's convention. For an empty project the defaults
come from the PROJECT row itself (def_task_type, def_complete_pct_type,
clndr_id), which P6 fills when the project is created.

Matching follows the repo-wide rule: activities are addressed by
``task_code``, never ``task_id`` (P6 renumbers ids on every boundary).

Writes go through p6.write.Session: one transaction per call, NEXTKEY id
blocks, audit stamps, ``confirm=true`` required, ``dry_run`` shows what would
be written.
"""
from __future__ import annotations

import base64
import uuid
from collections import Counter
from typing import Any, Mapping

from . import write as w

LINK_TYPES = {"FS": "PR_FS", "SS": "PR_SS", "FF": "PR_FF", "SF": "PR_SF"}
LINK_TYPES_REV = {v: k for k, v in LINK_TYPES.items()}

TASK_TYPES = {"task": "TT_Task", "milestone": "TT_Mile",
              "finish_milestone": "TT_FinMile", "loe": "TT_LOE",
              "wbs_summary": "TT_WBS", "resource_dependent": "TT_Rsrc"}

# Structural TASK columns whose default is the project's own convention.
_MODAL_COLS = ("clndr_id", "duration_type", "complete_pct_type", "task_type",
               "priority_type", "review_type")

_CURATED_TASK_DEFAULTS = {
    "phys_complete_pct": 0, "rev_fdbk_flag": "N", "lock_plan_flag": "N",
    "auto_compute_act_flag": "N", "review_type": "RV_OK",
    "status_code": "TK_NotStart", "driving_path_flag": "N",
    "priority_type": "PT_Normal", "est_wt": 1, "duration_type": "DT_FixedDrtn",
}


def _guid() -> str:
    """P6-style 22-char base64 GUID (e.g. 'Tae0dxy64ESlawOdfzndsw')."""
    return base64.b64encode(uuid.uuid4().bytes).decode().rstrip("=")


def _live(extra: str = "") -> str:
    return "delete_session_id IS NULL" + ((" AND " + extra) if extra else "")


def _find_task(s: w.Session, proj_id: int, task_code: str) -> dict[str, Any]:
    cols, rows = s.select_rows(
        "TASK", _live("proj_id = ? AND task_code = ?"), (proj_id, task_code))
    if not rows:
        raise w.P6WriteError("Aktivite bulunamadi: %s (proj %s)"
                             % (task_code, proj_id))
    return rows[0]


def _root_wbs(s: w.Session, proj_id: int) -> dict[str, Any]:
    cols, rows = s.select_rows(
        "PROJWBS", _live("proj_id = ? AND proj_node_flag = 'Y'"), (proj_id,))
    if not rows:
        raise w.P6WriteError("Projenin kok WBS dugumu yok: %s" % proj_id)
    return rows[0]


def _resolve_wbs(s: w.Session, proj_id: int, params: Mapping[str, Any]
                 ) -> dict[str, Any]:
    """wbs_id > wbs_path ('1.2.3' short names from the root) > project root."""
    if params.get("wbs_id") is not None:
        cols, rows = s.select_rows(
            "PROJWBS", _live("proj_id = ? AND wbs_id = ?"),
            (proj_id, int(params["wbs_id"])))
        if not rows:
            raise w.P6WriteError("wbs_id %s bu projede yok." % params["wbs_id"])
        return rows[0]
    path = params.get("wbs_path")
    if not path:
        return _root_wbs(s, proj_id)
    node = _root_wbs(s, proj_id)
    for part in str(path).split("."):
        cols, rows = s.select_rows(
            "PROJWBS",
            _live("proj_id = ? AND parent_wbs_id = ? AND wbs_short_name = ?"),
            (proj_id, node["wbs_id"], part.strip()))
        if not rows:
            raise w.P6WriteError(
                "WBS yolu cozulemedi: '%s' altinda '%s' yok."
                % (node["wbs_short_name"], part.strip()))
        if len(rows) > 1:
            raise w.P6WriteError(
                "WBS yolu belirsiz: '%s' altinda birden fazla '%s' var; "
                "wbs_id kullanin." % (node["wbs_short_name"], part.strip()))
        node = rows[0]
    return node


def _modal_defaults(s: w.Session, proj_id: int) -> dict[str, Any]:
    """The project's own convention: modal value per structural column."""
    cols, rows = s.select_rows("TASK", _live("proj_id = ?"), (proj_id,))
    out: dict[str, Any] = {}
    for col in _MODAL_COLS:
        values = [r.get(col) for r in rows if r.get(col) is not None]
        if values:
            out[col] = Counter(values).most_common(1)[0][0]
    return out


def _project_defaults(s: w.Session, proj_id: int) -> dict[str, Any]:
    cols, rows = s.select_rows("PROJECT", _live("proj_id = ?"), (proj_id,))
    if not rows:
        raise w.P6WriteError("Proje yok: %s" % proj_id)
    prow = rows[0]
    out = dict(_CURATED_TASK_DEFAULTS)
    if prow.get("clndr_id") is not None:
        out["clndr_id"] = prow["clndr_id"]
    if prow.get("def_task_type"):
        out["task_type"] = prow["def_task_type"]
    if prow.get("def_complete_pct_type"):
        out["complete_pct_type"] = prow["def_complete_pct_type"]
    out.update(_modal_defaults(s, proj_id))
    return out


def _next_task_code(s: w.Session, proj_id: int) -> str:
    """P6's own scheme: proj_short_name + next task_code_base step."""
    prow = s.execute(
        "SELECT proj_short_name, task_code_base, task_code_step FROM PROJECT "
        "WHERE proj_id = ?", proj_id).fetchone()
    short, base, step = prow[0], int(prow[1] or 1), int(prow[2] or 1)
    used = {r[0] for r in s.execute(
        "SELECT task_code FROM TASK WHERE proj_id = ?", proj_id).fetchall()}
    n = base
    while (short + str(n)) in used:
        n += step
    s.execute("UPDATE PROJECT SET task_code_base = ? WHERE proj_id = ?",
              n + step, proj_id)
    return short + str(n)


# ---------------------------------------------------------------------------
# WBS
# ---------------------------------------------------------------------------
def add_wbs(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    name = params.get("name") or params.get("wbs_name")
    if not name:
        raise w.P6WriteError("'name' (WBS adi) zorunlu.")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "WBS ekleme")
    with w.open_session(params) as s:
        w.project_exists(s, proj_id)
        parent = _resolve_wbs(s, proj_id, {
            "wbs_id": params.get("parent_wbs_id"),
            "wbs_path": params.get("parent_wbs_path")})
        short = str(params.get("short_name")
                    or params.get("wbs_short_name") or name)[:40]
        seq = s.scalar("SELECT ISNULL(MAX(seq_num),0) + 10 FROM PROJWBS "
                       "WHERE proj_id = ? AND parent_wbs_id = ?",
                       proj_id, parent["wbs_id"])
        row = {
            "proj_id": proj_id, "parent_wbs_id": parent["wbs_id"],
            "obs_id": parent["obs_id"], "seq_num": seq,
            "proj_node_flag": "N", "sum_data_flag": "N",
            "status_code": parent.get("status_code") or "WS_Open",
            "wbs_short_name": short, "wbs_name": str(name)[:120],
            "ev_user_pct": parent.get("ev_user_pct"),
            "ev_etc_user_value": parent.get("ev_etc_user_value"),
            "orig_cost": 0, "indep_remain_total_cost": None,
            "guid": _guid(),
        }
        if dry:
            return {"action": "add_wbs", "dry_run": True, "would_insert": {
                "parent": parent["wbs_short_name"], "wbs_short_name": short,
                "wbs_name": row["wbs_name"], "seq_num": seq}}
        wbs_id = s.reserve("projwbs_wbs_id", 1)[0]
        row["wbs_id"] = wbs_id
        cols = s.columns("PROJWBS")
        s.stamp_audit(row, cols)
        s.insert_rows("PROJWBS", [c for c in cols if c in row], [row])
        return {"action": "add_wbs", "wbs_id": wbs_id,
                "wbs_short_name": short, "wbs_name": row["wbs_name"],
                "parent_wbs_id": parent["wbs_id"],
                "parent_short_name": parent["wbs_short_name"]}


# ---------------------------------------------------------------------------
# Tasks
# ---------------------------------------------------------------------------
def add_task(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    name = params.get("name") or params.get("task_name")
    if not name:
        raise w.P6WriteError("'name' (aktivite adi) zorunlu.")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Aktivite ekleme")

    task_type_in = str(params.get("task_type") or "").strip()
    task_type = TASK_TYPES.get(task_type_in.lower(), task_type_in or None)

    duration_h = params.get("duration_h")
    if task_type in ("TT_Mile", "TT_FinMile"):
        duration_h = 0
    if duration_h is None:
        raise w.P6WriteError(
            "'duration_h' zorunlu (kilometre tasi icin task_type='milestone' "
            "verin, sure 0 kabul edilir).")
    duration_h = float(duration_h)
    if duration_h < 0:
        raise w.P6WriteError("Sure negatif olamaz.")

    with w.open_session(params) as s:
        w.project_exists(s, proj_id)
        defaults = _project_defaults(s, proj_id)
        wbs = _resolve_wbs(s, proj_id, params)
        task_code = params.get("task_code") or _next_task_code(s, proj_id)
        dup = s.scalar("SELECT COUNT(*) FROM TASK WHERE proj_id = ? AND "
                       "task_code = ? AND delete_session_id IS NULL",
                       proj_id, task_code)
        if dup:
            raise w.P6WriteError("task_code zaten var: " + str(task_code))

        row = dict(defaults)
        if task_type:
            row["task_type"] = task_type
        if params.get("clndr_id") is not None:
            row["clndr_id"] = int(params["clndr_id"])
        if "clndr_id" not in row:
            raise w.P6WriteError(
                "Takvim belirlenemedi: projede ne aktivite ne PROJECT.clndr_id "
                "var. 'clndr_id' parametresini verin.")
        row.update({
            "proj_id": proj_id, "wbs_id": wbs["wbs_id"],
            "task_code": str(task_code)[:40], "task_name": str(name)[:120],
            "remain_drtn_hr_cnt": duration_h, "target_drtn_hr_cnt": duration_h,
            "target_work_qty": 0, "remain_work_qty": 0,
            "target_equip_qty": 0, "remain_equip_qty": 0,
            "guid": _guid(),
        })
        if dry:
            return {"action": "add_task", "dry_run": True, "would_insert": {
                "task_code": row["task_code"], "task_name": row["task_name"],
                "wbs": wbs["wbs_short_name"], "duration_h": duration_h,
                "task_type": row.get("task_type"),
                "clndr_id": row.get("clndr_id"),
                "structural_defaults_from": "project modal + PROJECT row"}}
        task_id = s.reserve("task_task_id", 1)[0]
        row["task_id"] = task_id
        cols = s.columns("TASK")
        s.stamp_audit(row, cols)
        s.insert_rows("TASK", [c for c in cols if c in row], [row])
        return {"action": "add_task", "task_id": task_id,
                "task_code": row["task_code"], "task_name": row["task_name"],
                "wbs": wbs["wbs_short_name"], "duration_h": duration_h,
                "task_type": row.get("task_type"),
                "clndr_id": row.get("clndr_id"),
                "note": "Tarihler bos birakildi -- p6_job action='schedule' "
                        "calistirin, P6'nin CPM motoru hesaplasin."}


_UPDATABLE = {"task_name", "duration_h", "task_type", "clndr_id", "wbs_id",
              "wbs_path", "new_task_code", "auto_compute_actuals"}


def update_task(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    task_code = params.get("task_code")
    if not task_code:
        raise w.P6WriteError("'task_code' zorunlu.")
    fields = {k: params[k] for k in _UPDATABLE if params.get(k) is not None}
    if not fields:
        raise w.P6WriteError(
            "Degistirilecek alan yok. Kullanilabilir: " +
            ", ".join(sorted(_UPDATABLE)))
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Aktivite guncelleme")
    with w.open_session(params) as s:
        task = _find_task(s, proj_id, task_code)
        changes: dict[str, tuple[Any, Any]] = {}
        sets: dict[str, Any] = {}
        if "duration_h" in fields:
            if task["status_code"] != "TK_NotStart":
                raise w.P6WriteError(
                    "Baslamis/bitmis aktivitede sure bu araclarla degil "
                    "p6_progress ile degistirilir (kalan sure + fiili ayrimi).")
            d = float(fields["duration_h"])
            sets["target_drtn_hr_cnt"] = d
            sets["remain_drtn_hr_cnt"] = d
        if "task_name" in fields:
            sets["task_name"] = str(fields["task_name"])[:120]
        if "task_type" in fields:
            tt = str(fields["task_type"])
            sets["task_type"] = TASK_TYPES.get(tt.lower(), tt)
        if "clndr_id" in fields:
            sets["clndr_id"] = int(fields["clndr_id"])
        if "auto_compute_actuals" in fields:
            # JT_ApplyActuals only acts on activities flagged auto-compute;
            # with none flagged the whole job fails ("No projects to apply
            # actual to."). This is the switch that arms it.
            sets["auto_compute_act_flag"] = (
                "Y" if fields["auto_compute_actuals"] else "N")
        if "new_task_code" in fields:
            new_code = str(fields["new_task_code"])[:40]
            dup = s.scalar("SELECT COUNT(*) FROM TASK WHERE proj_id = ? AND "
                           "task_code = ? AND delete_session_id IS NULL",
                           proj_id, new_code)
            if dup:
                raise w.P6WriteError("task_code zaten var: " + new_code)
            sets["task_code"] = new_code
        if "wbs_id" in fields or "wbs_path" in fields:
            wbs = _resolve_wbs(s, proj_id, fields)
            sets["wbs_id"] = wbs["wbs_id"]
        for col, new in sets.items():
            old = task.get(col)
            if str(old) != str(new):
                changes[col] = (old, new)
        # A resource-loaded activity's remaining duration is recomputed by F9
        # from the assignment's remaining units (measured on bukhtourcity85:
        # 72h written, 240h came back; and here: 160h written, the 80h
        # assignment pinned it at 80h). Changing duration therefore carries
        # the assignment ledger along, exactly like p6_progress does: units
        # follow duration at each assignment's own units/time.
        assignments = []
        if "duration_h" in fields and params.get("update_assignments", True):
            d = float(fields["duration_h"])
            for trid, per_hr, cpq in s.execute(
                    "SELECT taskrsrc_id, target_qty_per_hr, cost_per_qty "
                    "FROM TASKRSRC WHERE task_id = ? AND "
                    "delete_session_id IS NULL", task["task_id"]).fetchall():
                per_hr_f = float(per_hr or 0) or 1.0
                qty = d * per_hr_f
                cost = qty * float(cpq or 0)
                assignments.append(
                    {"taskrsrc_id": int(trid), "qty": qty, "cost": cost})
        if dry:
            return {"action": "update_task", "dry_run": True,
                    "task_code": task_code,
                    "would_change": {k: {"from": v[0], "to": v[1]}
                                     for k, v in changes.items()},
                    "would_update_assignments": assignments}
        if changes:
            assign = ", ".join("[%s] = ?" % c for c in changes)
            args = [changes[c][1] for c in changes]
            s.execute(
                "UPDATE TASK SET %s, update_date = ?, update_user = ? "
                "WHERE task_id = ?" % assign,
                *args, s.stamp, s.user, task["task_id"])
            for a in assignments:
                s.execute(
                    "UPDATE TASKRSRC SET target_qty = ?, remain_qty = ?, "
                    "target_cost = ?, remain_cost = ?, update_date = ?, "
                    "update_user = ? WHERE taskrsrc_id = ?",
                    a["qty"], a["qty"], a["cost"], a["cost"],
                    s.stamp, s.user, a["taskrsrc_id"])
        warnings = []
        if ("duration_h" in fields and not assignments
                and not params.get("update_assignments", True)):
            warnings.append(
                "update_assignments=false: atama defterleri tasinmadi -- F9 "
                "kalan sureyi atamanin kalan biriminden GERI YAZAR.")
        return {"action": "update_task", "task_code": task_code,
                "changed": {k: {"from": str(v[0]), "to": str(v[1])}
                            for k, v in changes.items()},
                "assignments_updated": len(assignments),
                "unchanged": not changes,
                "warnings": warnings or None,
                "note": "Sure/takvim degistiyse p6_job action='schedule' "
                        "calistirin." if changes else None}


def delete_task(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    task_code = params.get("task_code")
    if not task_code:
        raise w.P6WriteError("'task_code' zorunlu.")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Aktivite silme")
    with w.open_session(params) as s:
        task = _find_task(s, proj_id, task_code)
        tid = task["task_id"]
        links = s.scalar("SELECT COUNT(*) FROM TASKPRED WHERE "
                         "(task_id = ? OR pred_task_id = ?) AND "
                         "delete_session_id IS NULL", tid, tid)
        rsrc = s.scalar("SELECT COUNT(*) FROM TASKRSRC WHERE task_id = ? AND "
                        "delete_session_id IS NULL", tid)
        if dry:
            return {"action": "delete_task", "dry_run": True,
                    "task_code": task_code, "would_soft_delete":
                        {"TASK": 1, "TASKPRED": links, "TASKRSRC": rsrc}}
        session_id = s.reserve("usession_session_id", 1)[0]
        for table, where in (
                ("TASKPRED", "(task_id = ? OR pred_task_id = ?)"),
                ("TASKRSRC", "(task_id = ? OR task_id = ?)"),
                ("TASK", "(task_id = ? OR task_id = ?)")):
            s.execute(
                "UPDATE [%s] SET delete_session_id = ?, delete_date = ? "
                "WHERE %s AND delete_session_id IS NULL" % (table, where),
                session_id, s.stamp, tid, tid)
        return {"action": "delete_task", "task_code": task_code,
                "soft_deleted": {"TASK": 1, "TASKPRED": links,
                                 "TASKRSRC": rsrc},
                "note": "P6 gibi yumusak silme (delete_session_id). Bag "
                        "sayilari degisti -- p6_job action='schedule' "
                        "calistirin."}


# ---------------------------------------------------------------------------
# Links
# ---------------------------------------------------------------------------
def _creates_cycle(s: w.Session, proj_id: int, pred_id: int, succ_id: int
                   ) -> bool:
    """Would pred->succ close a loop? BFS from succ through live links."""
    edges: dict[int, list[int]] = {}
    for a, b in s.execute(
            "SELECT pred_task_id, task_id FROM TASKPRED WHERE proj_id = ? "
            "AND delete_session_id IS NULL", proj_id).fetchall():
        edges.setdefault(a, []).append(b)
    seen, stack = set(), [succ_id]
    while stack:
        node = stack.pop()
        if node == pred_id:
            return True
        if node in seen:
            continue
        seen.add(node)
        stack.extend(edges.get(node, ()))
    return False


def add_link(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    pred_code = params.get("predecessor") or params.get("pred_task_code")
    succ_code = params.get("successor") or params.get("succ_task_code")
    if not pred_code or not succ_code:
        raise w.P6WriteError("'predecessor' ve 'successor' (task_code) zorunlu.")
    if pred_code == succ_code:
        raise w.P6WriteError("Aktivite kendine baglanamaz.")
    ltype_in = str(params.get("link_type") or "FS").upper()
    pred_type = LINK_TYPES.get(ltype_in) or (
        ltype_in if ltype_in in LINK_TYPES_REV else None)
    if pred_type is None:
        raise w.P6WriteError("link_type FS/SS/FF/SF olmali: " + ltype_in)
    lag_h = float(params.get("lag_h") or 0)
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Bag ekleme")
    with w.open_session(params) as s:
        pred = _find_task(s, proj_id, pred_code)
        succ = _find_task(s, proj_id, succ_code)
        dup = s.scalar(
            "SELECT COUNT(*) FROM TASKPRED WHERE task_id = ? AND "
            "pred_task_id = ? AND pred_type = ? AND delete_session_id IS NULL",
            succ["task_id"], pred["task_id"], pred_type)
        if dup:
            raise w.P6WriteError("Bu bag zaten var: %s -%s-> %s"
                                 % (pred_code, ltype_in, succ_code))
        if _creates_cycle(s, proj_id, pred["task_id"], succ["task_id"]):
            raise w.P6WriteError(
                "Bag dongu olusturur (%s zincirin sonunda %s'e donuyor) -- "
                "P6 F9'da da reddederdi." % (succ_code, pred_code))
        if dry:
            return {"action": "add_link", "dry_run": True, "would_insert": {
                "predecessor": pred_code, "successor": succ_code,
                "type": ltype_in, "lag_h": lag_h}}
        link_id = s.reserve("taskpred_task_pred_id", 1)[0]
        row = {"task_pred_id": link_id, "task_id": succ["task_id"],
               "pred_task_id": pred["task_id"], "proj_id": proj_id,
               "pred_proj_id": proj_id, "pred_type": pred_type,
               "lag_hr_cnt": lag_h}
        cols = s.columns("TASKPRED")
        s.stamp_audit(row, cols)
        s.insert_rows("TASKPRED", [c for c in cols if c in row], [row])
        return {"action": "add_link", "task_pred_id": link_id,
                "predecessor": pred_code, "successor": succ_code,
                "type": ltype_in, "lag_h": lag_h,
                "note": "p6_job action='schedule' calistirin."}


def delete_link(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    pred_code = params.get("predecessor") or params.get("pred_task_code")
    succ_code = params.get("successor") or params.get("succ_task_code")
    if not pred_code or not succ_code:
        raise w.P6WriteError("'predecessor' ve 'successor' (task_code) zorunlu.")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Bag silme")
    with w.open_session(params) as s:
        pred = _find_task(s, proj_id, pred_code)
        succ = _find_task(s, proj_id, succ_code)
        where = ("task_id = ? AND pred_task_id = ? AND "
                 "delete_session_id IS NULL")
        args = [succ["task_id"], pred["task_id"]]
        ltype_in = params.get("link_type")
        if ltype_in:
            where += " AND pred_type = ?"
            args.append(LINK_TYPES.get(str(ltype_in).upper(),
                                       str(ltype_in).upper()))
        count = s.scalar("SELECT COUNT(*) FROM TASKPRED WHERE " + where, *args)
        if not count:
            raise w.P6WriteError("Silinecek bag yok: %s -> %s"
                                 % (pred_code, succ_code))
        if dry:
            return {"action": "delete_link", "dry_run": True,
                    "would_soft_delete": count}
        session_id = s.reserve("usession_session_id", 1)[0]
        s.execute("UPDATE TASKPRED SET delete_session_id = ?, delete_date = ? "
                  "WHERE " + where, session_id, s.stamp, *args)
        return {"action": "delete_link", "predecessor": pred_code,
                "successor": succ_code, "soft_deleted": count,
                "note": "p6_job action='schedule' calistirin."}


# ---------------------------------------------------------------------------
# Resource assignment
# ---------------------------------------------------------------------------
def assign_resource(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    task_code = params.get("task_code")
    short = params.get("rsrc_short_name") or params.get("resource")
    if not task_code or not short:
        raise w.P6WriteError("'task_code' ve 'rsrc_short_name' zorunlu.")
    qty = float(params.get("target_qty") or 0)
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Kaynak atama")
    with w.open_session(params) as s:
        task = _find_task(s, proj_id, task_code)
        r = s.execute(
            "SELECT rsrc_id, rsrc_type, clndr_id FROM RSRC WHERE "
            "rsrc_short_name = ? AND delete_session_id IS NULL", short
        ).fetchone()
        if r is None:
            raise w.P6WriteError("Kaynak bulunamadi: " + str(short))
        rsrc_id, rsrc_type, r_clndr = int(r[0]), r[1], r[2]
        dup = s.scalar(
            "SELECT COUNT(*) FROM TASKRSRC WHERE task_id = ? AND rsrc_id = ? "
            "AND delete_session_id IS NULL", task["task_id"], rsrc_id)
        if dup:
            raise w.P6WriteError(
                "Atama zaten var: %s -> %s" % (short, task_code))
        cost_per_qty = params.get("cost_per_qty")
        rate = None
        if cost_per_qty is None:
            rate = s.scalar(
                "SELECT TOP 1 cost_per_qty FROM RSRCRATE WHERE rsrc_id = ? "
                "AND delete_session_id IS NULL ORDER BY start_date DESC",
                rsrc_id)
            cost_per_qty = float(rate) if rate is not None else 0.0
        cost_per_qty = float(cost_per_qty)
        target_cost = qty * cost_per_qty
        # Units/time MUST be non-zero: F9 zeroes the activity's remaining AND
        # planned duration when the assignment carries 0 units/hr (measured
        # 26.08.2026 -- an 80h DT_FixedDrtn activity came back 0h). P6's own
        # default is the resource's rate-sheet max units/time, full time = 1/hr.
        qty_per_hr = params.get("qty_per_hr")
        if qty_per_hr is None:
            qty_per_hr = s.scalar(
                "SELECT TOP 1 max_qty_per_hr FROM RSRCRATE WHERE rsrc_id = ? "
                "AND delete_session_id IS NULL ORDER BY start_date DESC",
                rsrc_id)
        qty_per_hr = float(qty_per_hr) if qty_per_hr else 1.0
        if dry:
            return {"action": "assign_resource", "dry_run": True,
                    "would_insert": {
                        "task_code": task_code, "rsrc_short_name": short,
                        "target_qty": qty, "cost_per_qty": cost_per_qty,
                        "cost_per_qty_source": ("param" if rate is None
                                                else "RSRCRATE"),
                        "qty_per_hr": qty_per_hr,
                        "target_cost": target_cost}}
        trid = s.reserve("taskrsrc_taskrsrc_id", 1)[0]
        row = {
            "taskrsrc_id": trid, "task_id": task["task_id"],
            "proj_id": proj_id, "rsrc_id": rsrc_id,
            "rsrc_type": rsrc_type or "RT_Labor",
            "cost_qty_link_flag": "Y", "rollup_dates_flag": "Y",
            "cost_per_qty_source_type": "ST_Rsrc", "ts_pend_act_end_flag": "N",
            "target_qty": qty, "remain_qty": qty,
            "target_cost": target_cost, "remain_cost": target_cost,
            "act_reg_qty": 0, "act_ot_qty": 0, "act_reg_cost": 0,
            "act_ot_cost": 0, "cost_per_qty": cost_per_qty,
            "target_qty_per_hr": qty_per_hr, "remain_qty_per_hr": qty_per_hr,
            "guid": _guid(),
        }
        cols = s.columns("TASKRSRC")
        s.stamp_audit(row, cols)
        s.insert_rows("TASKRSRC", [c for c in cols if c in row], [row])
        return {"action": "assign_resource", "taskrsrc_id": trid,
                "task_code": task_code, "rsrc_short_name": short,
                "target_qty": qty, "cost_per_qty": cost_per_qty,
                "target_cost": target_cost}


def remove_assignment(params: Mapping[str, Any]) -> dict[str, Any]:
    proj_id = int(params["proj_id"])
    task_code = params.get("task_code")
    short = params.get("rsrc_short_name") or params.get("resource")
    if not task_code or not short:
        raise w.P6WriteError("'task_code' ve 'rsrc_short_name' zorunlu.")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Atama silme")
    with w.open_session(params) as s:
        task = _find_task(s, proj_id, task_code)
        count = s.scalar(
            "SELECT COUNT(*) FROM TASKRSRC tr JOIN RSRC r ON "
            "r.rsrc_id = tr.rsrc_id WHERE tr.task_id = ? AND "
            "r.rsrc_short_name = ? AND tr.delete_session_id IS NULL",
            task["task_id"], short)
        if not count:
            raise w.P6WriteError("Atama yok: %s -> %s" % (short, task_code))
        if dry:
            return {"action": "remove_assignment", "dry_run": True,
                    "would_soft_delete": count}
        session_id = s.reserve("usession_session_id", 1)[0]
        s.execute(
            "UPDATE tr SET tr.delete_session_id = ?, tr.delete_date = ? "
            "FROM TASKRSRC tr JOIN RSRC r ON r.rsrc_id = tr.rsrc_id "
            "WHERE tr.task_id = ? AND r.rsrc_short_name = ? "
            "AND tr.delete_session_id IS NULL",
            session_id, s.stamp, task["task_id"], short)
        return {"action": "remove_assignment", "task_code": task_code,
                "rsrc_short_name": short, "soft_deleted": count}


# ---------------------------------------------------------------------------
# Project creation -- from nothing
# ---------------------------------------------------------------------------
# Values every P6 project row must carry; taken from what P6 24.12 itself
# writes when a project is created in the client (reference: live PROJECT
# rows in this schema). Anything schedule-specific is parameterised.
_PROJECT_DEFAULTS = {
    "fy_start_month_num": 1, "chng_eff_cmp_pct_flag": "N",
    "rsrc_self_add_flag": "Y", "rsrc_role_match_flag": "N",
    "allow_complete_flag": "Y", "rsrc_multi_assign_flag": "Y",
    "checkout_flag": "N", "project_flag": "Y", "step_complete_flag": "N",
    "cost_qty_recalc_flag": "N", "sum_only_flag": "N", "batch_sum_flag": "Y",
    "name_sep_char": ".", "def_complete_pct_type": "CP_Drtn",
    "def_task_type": "TT_Task", "act_this_per_link_flag": "Y",
    "act_pct_link_flag": "Y", "add_act_remain_flag": "N",
    "critical_path_type": "CT_TotFloat", "task_code_prefix_flag": "Y",
    "def_rollup_dates_flag": "Y", "rem_target_link_flag": "Y",
    "reset_planned_flag": "N", "allow_neg_act_flag": "N",
    "msp_managed_flag": "N", "msp_update_actuals_flag": "N",
    "use_project_baseline_flag": "Y", "ts_rsrc_vw_compl_asgn_flag": "N",
    "ts_rsrc_mark_act_finish_flag": "N", "ts_rsrc_vw_inact_actv_flag": "N",
    "control_updates_flag": "N", "hist_interval": "Month",
    "hist_level": "HL_None", "task_code_base": 10, "task_code_step": 10,
    "priority_num": 10, "wbs_max_sum_level": 2, "sum_assign_level":
    "SL_Taskrsrc", "critical_drtn_hr_cnt": 0, "def_duration_type":
    "DT_FixedDrtn", "guid": None, "orig_proj_id": None,
}


def _eps_node(s: w.Session, params: Mapping[str, Any]) -> dict[str, Any]:
    if params.get("eps_wbs_id") is not None:
        cols, rows = s.select_rows(
            "PROJWBS", _live("wbs_id = ?"), (int(params["eps_wbs_id"]),))
        if not rows:
            raise w.P6WriteError("eps_wbs_id yok: %s" % params["eps_wbs_id"])
        return rows[0]
    # Default: the EPS branch every existing project root hangs under.
    row = s.execute(
        "SELECT TOP 1 w.wbs_id, w.obs_id, w.wbs_short_name FROM PROJWBS w "
        "JOIN PROJECT p ON p.proj_id = w.proj_id "
        "WHERE p.project_flag = 'N' AND p.orig_proj_id IS NULL "
        "AND w.delete_session_id IS NULL AND w.parent_wbs_id IS NULL "
        "ORDER BY w.wbs_id").fetchone()
    if row is None:
        raise w.P6WriteError(
            "EPS dugumu bulunamadi; 'eps_wbs_id' parametresini verin.")
    return {"wbs_id": int(row[0]), "obs_id": row[1],
            "wbs_short_name": row[2]}


def create_project(params: Mapping[str, Any]) -> dict[str, Any]:
    short = params.get("short_name") or params.get("proj_short_name")
    name = params.get("name") or params.get("proj_name") or short
    plan_start = params.get("plan_start") or params.get("plan_start_date")
    if not short or not plan_start:
        raise w.P6WriteError("'short_name' ve 'plan_start' zorunlu.")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Proje olusturma")
    with w.open_session(params) as s:
        dup = s.scalar("SELECT COUNT(*) FROM PROJECT WHERE proj_short_name = ? "
                       "AND delete_session_id IS NULL", short)
        if dup:
            raise w.P6WriteError("proj_short_name zaten var: " + str(short))
        clndr_id = params.get("clndr_id")
        if clndr_id is None:
            clndr_id = s.scalar(
                "SELECT TOP 1 clndr_id FROM CALENDAR WHERE default_flag = 'Y' "
                "AND delete_session_id IS NULL")
        if clndr_id is None:
            raise w.P6WriteError(
                "Varsayilan takvim yok; 'clndr_id' parametresini verin.")
        eps = _eps_node(s, params)
        if dry:
            return {"action": "create_project", "dry_run": True,
                    "would_insert": {
                        "proj_short_name": short, "wbs_name": name,
                        "plan_start": str(plan_start),
                        "clndr_id": int(clndr_id),
                        "eps_node": eps["wbs_short_name"]}}
        proj_id = s.reserve("project_proj_id", 1)[0]
        wbs_id = s.reserve("projwbs_wbs_id", 1)[0]

        prow = dict(_PROJECT_DEFAULTS)
        prow.update({
            "proj_id": proj_id, "proj_short_name": str(short)[:40],
            "clndr_id": int(clndr_id), "add_date": s.stamp,
            "plan_start_date": plan_start, "last_recalc_date": plan_start,
            "add_by_name": s.user, "guid": _guid(),
        })
        pcols = s.columns("PROJECT")
        s.stamp_audit(prow, pcols)
        s.insert_rows("PROJECT", [c for c in pcols if c in prow], [prow])

        wrow = {
            "wbs_id": wbs_id, "proj_id": proj_id,
            "parent_wbs_id": eps["wbs_id"], "obs_id": eps["obs_id"],
            "seq_num": 10, "proj_node_flag": "Y", "sum_data_flag": "N",
            "status_code": "WS_Open", "wbs_short_name": str(short)[:40],
            "wbs_name": str(name)[:120], "ev_user_pct": 6,
            "est_wt": 1, "orig_cost": 0, "guid": _guid(),
        }
        wcols = s.columns("PROJWBS")
        s.stamp_audit(wrow, wcols)
        s.insert_rows("PROJWBS", [c for c in wcols if c in wrow], [wrow])

        obsproj = s.scalar("SELECT COUNT(*) FROM OBSPROJ WHERE proj_id = ?",
                           proj_id)
        return {"action": "create_project", "proj_id": proj_id,
                "proj_short_name": prow["proj_short_name"],
                "root_wbs_id": wbs_id, "clndr_id": int(clndr_id),
                "plan_start": str(plan_start),
                "eps_node": eps["wbs_short_name"],
                "obsproj_rows": obsproj,
                "note": "OBSPROJ satirini TR_PROJECT_OBSPROJ trigger'i "
                        "uretir; obsproj_rows=0 ise Job Service erisimi icin "
                        "OBS atamasini kontrol edin."}


ACTIONS = {
    "create_project": create_project,
    "add_wbs": add_wbs,
    "add_task": add_task,
    "update_task": update_task,
    "delete_task": delete_task,
    "add_link": add_link,
    "delete_link": delete_link,
    "assign_resource": assign_resource,
    "remove_assignment": remove_assignment,
}
