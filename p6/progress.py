"""Progress and actuals entry for P6 -- the write side of a status update.

P6 Professional has no automation interface, so a monthly update cannot be
driven through the application; this module writes the progress fields
directly and then lets P6's own CPM engine recompute dates through the Job
Service (``schedule=true``, or p6_job action='schedule' afterwards). Nothing
here calculates a date.

**Why this is more than an UPDATE statement.** P6 keeps an activity's state in
several fields that must agree, and a schedule where they disagree is
silently wrong rather than loudly broken:

* ``status_code`` (TK_NotStart / TK_Active / TK_Complete) has to match the
  actual dates: started means act_start_date is set, finished means both
  dates are set.
* ``complete_pct_type`` decides where percent complete actually lives.
  bukhtourcity is CP_Drtn, so writing phys_complete_pct alone changes nothing
  a P6 user would see -- remaining duration is the field that matters. Both
  are written, consistently, whichever basis the activity uses.
* Remaining duration and remaining units must fall to zero on completion, or
  the next reschedule pushes work past a finished activity.
* P6 rejects actual dates later than the data date. So does this module,
  unless the caller explicitly says otherwise.

Every action supports ``dry_run`` and shows before/after per activity. Writes
require ``confirm=true``.
"""
from __future__ import annotations

import datetime as _dt
from typing import Any, Mapping, Sequence

from . import analysis, write as w

STATUS_NOT_STARTED = "TK_NotStart"
STATUS_ACTIVE = "TK_Active"
STATUS_COMPLETE = "TK_Complete"

# Keys are lower case because lookup lower-cases the caller's value; P6's own
# codes (TK_Complete) must survive that, so they are listed lowered too.
STATUS_ALIASES = {
    "not_started": STATUS_NOT_STARTED, "notstart": STATUS_NOT_STARTED,
    "baslamadi": STATUS_NOT_STARTED, "baslamadı": STATUS_NOT_STARTED,
    STATUS_NOT_STARTED.lower(): STATUS_NOT_STARTED,
    "in_progress": STATUS_ACTIVE, "active": STATUS_ACTIVE,
    "devam": STATUS_ACTIVE, "devam_ediyor": STATUS_ACTIVE,
    STATUS_ACTIVE.lower(): STATUS_ACTIVE,
    "complete": STATUS_COMPLETE, "completed": STATUS_COMPLETE,
    "bitti": STATUS_COMPLETE, "tamamlandi": STATUS_COMPLETE,
    "tamamlandı": STATUS_COMPLETE,
    STATUS_COMPLETE.lower(): STATUS_COMPLETE,
}

# TASK columns this module is allowed to touch. Anything outside this list is
# P6's business -- dates especially, which only the scheduler may set.
TASK_FIELDS = ("status_code", "act_start_date", "act_end_date",
               "phys_complete_pct", "remain_drtn_hr_cnt",
               "act_work_qty", "remain_work_qty")


class ProgressError(w.P6WriteError):
    """Invalid progress input."""


def _parse_date(value: Any, label: str) -> _dt.datetime | None:
    if value in (None, "", "null"):
        return None
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime.combine(value, _dt.time())
    text = str(value).strip().replace("T", " ")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%d.%m.%Y %H:%M", "%d.%m.%Y"):
        try:
            return _dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ProgressError(
        "%s tarihi okunamadi: %r (YYYY-MM-DD veya YYYY-MM-DD HH:MM bekleniyor)"
        % (label, value))


def _pct(value: Any, label: str) -> float | None:
    if value is None:
        return None
    try:
        pct = float(value)
    except (TypeError, ValueError):
        raise ProgressError("%s sayi olmali: %r" % (label, value))
    if not 0.0 <= pct <= 100.0:
        raise ProgressError("%s 0-100 arasinda olmali: %s" % (label, pct))
    return pct


def _data_date(s: w.Session, proj_id: int) -> _dt.datetime | None:
    return s.scalar("SELECT last_recalc_date FROM PROJECT WHERE proj_id = ?",
                    proj_id)


def _load_tasks(s: w.Session, proj_id: int, keys: Sequence[Any]
                ) -> tuple[dict[Any, dict], dict[Any, dict]]:
    """Current state of every activity, keyed by both task_id and task_code."""
    cols = ("task_id, task_code, task_name, status_code, complete_pct_type, "
            "phys_complete_pct, target_drtn_hr_cnt, remain_drtn_hr_cnt, "
            "act_start_date, act_end_date, target_work_qty, act_work_qty, "
            "remain_work_qty, target_start_date, target_end_date")
    s.execute("SELECT %s FROM TASK WHERE proj_id = ? AND delete_session_id "
              "IS NULL" % cols, proj_id)
    names = [d[0] for d in s.cur.description]
    by_id, by_code = {}, {}
    for row in s.cur.fetchall():
        rec = dict(zip(names, row))
        by_id[int(rec["task_id"])] = rec
        by_code[rec["task_code"]] = rec
    return by_id, by_code


def _resolve_task(update: Mapping[str, Any], by_id, by_code) -> dict:
    if update.get("task_id") is not None:
        rec = by_id.get(int(update["task_id"]))
        if rec is None:
            raise ProgressError("Aktivite bulunamadi: task_id=%s"
                                % update["task_id"])
        return rec
    code = update.get("task_code") or update.get("code")
    if not code:
        raise ProgressError("Her guncelleme icin task_id veya task_code sart: %r"
                            % dict(update))
    rec = by_code.get(str(code))
    if rec is None:
        raise ProgressError("Aktivite bulunamadi: task_code=%s" % code)
    return rec


def _plan(update: Mapping[str, Any], cur: Mapping[str, Any],
          data_date: _dt.datetime | None, allow_future: bool) -> dict[str, Any]:
    """Turn one requested change into a complete, self-consistent field set."""
    raw_status = update.get("status")
    status = None
    if raw_status is not None:
        status = STATUS_ALIASES.get(str(raw_status).strip().lower())
        if status is None:
            raise ProgressError(
                "status '%s' gecersiz. Kullanilabilir: not_started, "
                "in_progress, complete" % raw_status)

    act_start = _parse_date(update.get("actual_start"), "actual_start")
    act_end = _parse_date(update.get("actual_finish")
                          or update.get("actual_end"), "actual_finish")
    pct = _pct(update.get("percent_complete"), "percent_complete")
    remain_h = update.get("remaining_duration_h")
    target_h = float(cur.get("target_drtn_hr_cnt") or 0)

    # Infer the status when the caller only gave dates or a percentage.
    if status is None:
        if act_end or pct == 100:
            status = STATUS_COMPLETE
        elif act_start or (pct or 0) > 0:
            status = STATUS_ACTIVE
        else:
            status = cur.get("status_code") or STATUS_NOT_STARTED

    if status == STATUS_NOT_STARTED:
        fields = {
            "status_code": STATUS_NOT_STARTED,
            "act_start_date": None, "act_end_date": None,
            "phys_complete_pct": 0.0,
            "remain_drtn_hr_cnt": target_h,
            "act_work_qty": 0.0,
            "remain_work_qty": float(cur.get("target_work_qty") or 0),
        }
        return fields

    if act_start is None:
        act_start = cur.get("act_start_date")
    if act_start is None:
        raise ProgressError(
            "'%s' aktivitesi baslatiliyor ama actual_start verilmedi ve "
            "kayitli fiili baslangic yok." % cur.get("task_code"))

    if status == STATUS_COMPLETE:
        if act_end is None:
            act_end = cur.get("act_end_date")
        if act_end is None:
            raise ProgressError(
                "'%s' tamamlandi olarak isaretleniyor ama actual_finish "
                "verilmedi." % cur.get("task_code"))
        if act_end < act_start:
            raise ProgressError(
                "'%s': fiili bitis (%s) fiili baslangictan (%s) once olamaz."
                % (cur.get("task_code"), act_end.date(), act_start.date()))
        return {
            "status_code": STATUS_COMPLETE,
            "act_start_date": act_start, "act_end_date": act_end,
            "phys_complete_pct": 100.0,
            "remain_drtn_hr_cnt": 0.0,
            "act_work_qty": float(cur.get("target_work_qty") or 0),
            "remain_work_qty": 0.0,
        }

    # In progress.
    if act_end is not None:
        raise ProgressError(
            "'%s' devam ediyor olarak isaretleniyor ama fiili bitis tarihi "
            "verilmis. status='complete' mi demek istediniz?"
            % cur.get("task_code"))
    if pct is None and remain_h is None:
        raise ProgressError(
            "'%s' icin percent_complete veya remaining_duration_h'dan biri "
            "sart." % cur.get("task_code"))
    if remain_h is not None:
        remain = float(remain_h)
        if remain < 0:
            raise ProgressError("remaining_duration_h negatif olamaz.")
        pct = (100.0 * (target_h - remain) / target_h) if target_h > 0 else (pct or 0.0)
        pct = max(0.0, min(100.0, pct))
    else:
        remain = target_h * (1.0 - pct / 100.0)
    if pct >= 100:
        raise ProgressError(
            "'%s': devam eden aktivite %%100 olamaz; status='complete' ve "
            "actual_finish verin." % cur.get("task_code"))
    target_work = float(cur.get("target_work_qty") or 0)
    return {
        "status_code": STATUS_ACTIVE,
        "act_start_date": act_start, "act_end_date": None,
        "phys_complete_pct": round(pct, 4),
        "remain_drtn_hr_cnt": round(remain, 4),
        "act_work_qty": round(target_work * pct / 100.0, 4),
        "remain_work_qty": round(target_work * (1 - pct / 100.0), 4),
    }


def _check_dates(fields: Mapping[str, Any], cur: Mapping[str, Any],
                 data_date, allow_future: bool) -> list[str]:
    warn = []
    if data_date is None:
        return ["Veri tarihi (PROJECT.last_recalc_date) bos; fiili tarihler "
                "veri tarihine gore denetlenemedi."]
    for key, label in (("act_start_date", "fiili baslangic"),
                       ("act_end_date", "fiili bitis")):
        value = fields.get(key)
        if value and value > data_date:
            msg = ("'%s' %s (%s) veri tarihinden (%s) sonra; P6 gelecege "
                   "fiili tarih kabul etmez."
                   % (cur.get("task_code"), label, value.date(),
                      data_date.date()))
            if not allow_future:
                raise ProgressError(msg + " allow_future_actuals=true ile "
                                          "zorlayabilirsiniz.")
            warn.append(msg)
    return warn


def _assignment_plan(rows: Sequence[Mapping[str, Any]], pct: float
                     ) -> list[tuple[int, float, float, float, float]]:
    """Move an assignment's units and cost to match the activity's completion.

    This is not bookkeeping tidiness -- it is required for the schedule to
    survive a reschedule. On a resource-loaded activity whose duration type
    links duration to units (DT_FixedDUR2 here), P6's scheduler recomputes
    remaining duration FROM the assignment's remaining units. Measured on
    bukhtourcity85: remaining duration written as 72h, F9 put it straight
    back to 240h because the assignment still said 240h remaining, while the
    identical activity with no assignment (bukhtourcity1346) kept its value.
    Leaving actuals at zero also holds AC at zero, so CPI never computes.
    """
    out = []
    frac = max(0.0, min(1.0, pct / 100.0))
    for r in rows:
        target_qty = float(r["target_qty"] or 0)
        rate = float(r["cost_per_qty"] or 0)
        target_cost = float(r["target_cost"] or 0)
        act_qty = target_qty * frac
        remain_qty = target_qty - act_qty
        if rate:
            act_cost, remain_cost = act_qty * rate, remain_qty * rate
        else:
            act_cost, remain_cost = target_cost * frac, target_cost - target_cost * frac
        out.append((int(r["taskrsrc_id"]), round(act_qty, 4),
                    round(remain_qty, 4), round(act_cost, 4),
                    round(remain_cost, 4)))
    return out


def _load_assignments(s: w.Session, proj_id: int) -> dict[int, list[dict]]:
    s.execute("SELECT taskrsrc_id, task_id, target_qty, remain_qty, act_reg_qty, "
              "target_cost, act_reg_cost, remain_cost, cost_per_qty "
              "FROM TASKRSRC WHERE proj_id = ? AND delete_session_id IS NULL",
              proj_id)
    names = [d[0] for d in s.cur.description]
    out: dict[int, list[dict]] = {}
    for row in s.cur.fetchall():
        rec = dict(zip(names, row))
        out.setdefault(int(rec["task_id"]), []).append(rec)
    return out


def _diff(cur: Mapping[str, Any], fields: Mapping[str, Any]) -> dict[str, Any]:
    out = {}
    for key, new in fields.items():
        old = cur.get(key)
        if isinstance(old, _dt.datetime) or isinstance(new, _dt.datetime):
            same = old == new
        else:
            same = abs(float(old or 0) - float(new or 0)) < 1e-6 \
                if isinstance(new, (int, float)) else old == new
        if not same:
            out[key] = {"from": str(old) if old is not None else None,
                        "to": str(new) if new is not None else None}
    return out


# ---------------------------------------------------------------------------
# actions
# ---------------------------------------------------------------------------
def read(params: Mapping[str, Any]) -> dict[str, Any]:
    """Current progress, with percent complete resolved per complete_pct_type."""
    data = analysis.load(params)
    limit = min(max(int(params.get("limit", 100) or 100), 1), 500)
    only_started = bool(params.get("only_started"))
    rows = []
    for t in data["tasks"]:
        if only_started and not (t.get("actual_start") or t.get("percent_complete")):
            continue
        rows.append({
            "id": t.get("id"), "code": t.get("code"), "name": t.get("name"),
            "status": t.get("status"),
            "percent_complete": round(float(t.get("percent_complete") or 0), 2),
            "percent_complete_stored": t.get("percent_complete_stored"),
            "actual_start": t.get("actual_start"),
            "actual_finish": t.get("actual_finish"),
            "duration_h": t.get("duration_h"),
            "forecast_finish": t.get("forecast_finish"),
        })
    started = sum(1 for t in data["tasks"] if t.get("actual_start"))
    done = sum(1 for t in data["tasks"] if float(t.get("percent_complete") or 0) >= 100)
    return {
        "action": "read", "source": data["source"],
        "status_date": data["status_date"],
        "task_count": data["task_count"],
        "started": started, "completed": done,
        "not_started": data["task_count"] - started,
        "percent_complete_basis": data.get("percent_complete_basis"),
        "count": len(rows), "truncated": len(rows) > limit,
        "tasks": rows[:limit],
    }


def set_progress(params: Mapping[str, Any]) -> dict[str, Any]:
    """Write progress onto one or more activities, consistently."""
    proj_id = params.get("proj_id")
    if proj_id is None:
        raise ProgressError("proj_id zorunlu.")
    proj_id = int(proj_id)

    updates = params.get("updates")
    if updates is None:
        single = {k: params[k] for k in
                  ("task_id", "task_code", "code", "status", "actual_start",
                   "actual_finish", "actual_end", "percent_complete",
                   "remaining_duration_h") if k in params}
        updates = [single] if single else None
    if not updates:
        raise ProgressError(
            "'updates' listesi zorunlu (ya da tek aktivite icin task_code + "
            "status/percent_complete parametreleri).")
    if not isinstance(updates, (list, tuple)):
        raise ProgressError("'updates' bir liste olmali.")
    if len(updates) > 5000:
        raise ProgressError("Tek cagride en fazla 5000 aktivite guncellenebilir.")

    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Ilerleme girisi")
    allow_future = bool(params.get("allow_future_actuals"))

    with w.open_session(params) as s:
        w.project_exists(s, proj_id)
        data_date = _data_date(s, proj_id)
        by_id, by_code = _load_tasks(s, proj_id, [])

        sync_assignments = params.get("update_assignments", True)
        asg_by_task = _load_assignments(s, proj_id) if sync_assignments else {}

        planned, warnings, changes = [], [], []
        asg_planned: list[tuple[int, float, float, float, float]] = []
        seen = set()
        for upd in updates:
            cur = _resolve_task(upd, by_id, by_code)
            tid = int(cur["task_id"])
            if tid in seen:
                raise ProgressError(
                    "Ayni aktivite listede iki kez var: %s" % cur["task_code"])
            seen.add(tid)
            fields = _plan(upd, cur, data_date, allow_future)
            warnings.extend(_check_dates(fields, cur, data_date, allow_future))
            delta = _diff(cur, fields)
            planned.append((tid, fields))

            asg_rows = asg_by_task.get(tid, [])
            asg_delta = _assignment_plan(asg_rows, fields["phys_complete_pct"]) \
                if asg_rows else []
            asg_planned.extend(asg_delta)
            if asg_rows and not sync_assignments:
                warnings.append(
                    "'%s' kaynak yuklu ama update_assignments=false verildi; "
                    "P6 reschedule'da kalan sureyi atamanin kalan birimine "
                    "gore geri yazar." % cur["task_code"])

            changes.append({
                "task_id": tid, "code": cur["task_code"],
                "name": cur["task_name"],
                "complete_pct_type": cur["complete_pct_type"],
                "changed": delta, "unchanged": not delta,
                "assignments_synced": len(asg_delta),
            })

        if dry:
            s.conn.rollback()
            return {"action": "set_progress", "dry_run": True,
                    "proj_id": proj_id, "data_date": str(data_date)[:10],
                    "requested": len(updates),
                    "would_change": sum(1 for c in changes if not c["unchanged"]),
                    "would_sync_assignments": len(asg_planned),
                    "warnings": warnings, "activities": changes,
                    "note": "Hicbir sey yazilmadi. confirm=true ile calistirin."}

        sql = ("UPDATE TASK SET " + ", ".join("%s = ?" % f for f in TASK_FIELDS)
               + ", update_date = ?, update_user = ? WHERE task_id = ?")
        for tid, fields in planned:
            s.execute(sql, *[fields[f] for f in TASK_FIELDS],
                      s.stamp, s.user, tid)

        for asg_id, aq, rq, ac, rc in asg_planned:
            s.execute(
                "UPDATE TASKRSRC SET act_reg_qty = ?, remain_qty = ?, "
                "act_reg_cost = ?, remain_cost = ?, update_date = ?, "
                "update_user = ? WHERE taskrsrc_id = ?",
                aq, rq, ac, rc, s.stamp, s.user, asg_id)

        result = {
            "action": "set_progress", "proj_id": proj_id,
            "data_date": str(data_date)[:10],
            "updated": len(planned),
            "changed": sum(1 for c in changes if not c["unchanged"]),
            "assignments_synced": len(asg_planned),
            "warnings": warnings,
            "activities": changes,
            "note": ("Tarihler yeniden hesaplanmadi. Yeni erken/gec tarihler "
                     "icin p6_job action='schedule' calistirin "
                     "(veya schedule=true verin)."),
        }
        s.conn.commit()
        s._committed = True

    if params.get("schedule"):
        result["schedule"] = _run_schedule(params, proj_id)
    return result


def set_assignment_actuals(params: Mapping[str, Any]) -> dict[str, Any]:
    """Book actual units/cost against resource assignments.

    Activity progress and assignment actuals are different books in P6: the
    first drives dates, the second drives AC in earned value. A cost report
    that quotes CPI needs this one.
    """
    proj_id = int(params["proj_id"])
    updates = params.get("updates")
    if not updates:
        raise ProgressError("'updates' listesi zorunlu.")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Fiili birim/maliyet girisi")

    with w.open_session(params) as s:
        w.project_exists(s, proj_id)
        s.execute(
            "SELECT r.taskrsrc_id, r.task_id, t.task_code, r.rsrc_id, "
            "rs.rsrc_short_name, "
            "r.target_qty, r.act_reg_qty, r.remain_qty, r.target_cost, "
            "r.act_reg_cost, r.remain_cost, r.cost_per_qty "
            "FROM TASKRSRC r JOIN TASK t ON t.task_id = r.task_id "
            "LEFT JOIN RSRC rs ON rs.rsrc_id = r.rsrc_id "
            "WHERE r.proj_id = ? AND r.delete_session_id IS NULL", proj_id)
        names = [d[0] for d in s.cur.description]
        rows = [dict(zip(names, r)) for r in s.cur.fetchall()]
        by_asg = {int(r["taskrsrc_id"]): r for r in rows}
        by_task: dict[str, list[dict]] = {}
        for r in rows:
            by_task.setdefault(r["task_code"], []).append(r)

        planned, changes = [], []
        for upd in updates:
            if upd.get("taskrsrc_id") is not None:
                target = [by_asg.get(int(upd["taskrsrc_id"]))]
                if target[0] is None:
                    raise ProgressError("Atama bulunamadi: taskrsrc_id=%s"
                                        % upd["taskrsrc_id"])
            else:
                code = upd.get("task_code") or upd.get("code")
                target = by_task.get(str(code)) or []
                if not target:
                    raise ProgressError(
                        "'%s' aktivitesinde kaynak atamasi yok." % code)
                if len(target) > 1 and upd.get("rsrc_id") is not None:
                    target = [r for r in target
                              if int(r["rsrc_id"]) == int(upd["rsrc_id"])]
                short = upd.get("rsrc_short_name") or upd.get("resource")
                if len(target) > 1 and short:
                    target = [r for r in target
                              if r.get("rsrc_short_name") == short]
                    if not target:
                        raise ProgressError(
                            "'%s' aktivitesinde '%s' kaynagina atama yok."
                            % (code, short))
                if len(target) > 1:
                    raise ProgressError(
                        "'%s' aktivitesinde %d atama var; rsrc_short_name, "
                        "rsrc_id veya taskrsrc_id ile hangisi oldugunu "
                        "belirtin." % (code, len(target)))
            rec = target[0]

            act_qty = upd.get("actual_qty")
            remain_qty = upd.get("remaining_qty")
            act_cost = upd.get("actual_cost")
            remain_cost = upd.get("remaining_cost")
            if act_qty is None and act_cost is None:
                raise ProgressError(
                    "Her guncelleme icin actual_qty veya actual_cost sart.")

            new_act_qty = float(act_qty) if act_qty is not None \
                else float(rec["act_reg_qty"] or 0)
            if remain_qty is not None:
                new_remain_qty = float(remain_qty)
            else:
                new_remain_qty = max(0.0, float(rec["target_qty"] or 0) - new_act_qty)
            # Etkin oran: atamanin kendi cost_per_qty'si, o sifirsa P6'nin
            # hedeflerinden turetilen oran (target_cost / target_qty). P6
            # global RSRCRATE'ten maliyet hesapladiginda atama kolonunu 0
            # birakabiliyor (bukhtourcity: cost_per_qty=0, target_cost=800,
            # target_qty=160) -- 0 orana dusup fiili maliyeti sabit birakmak
            # maliyet-yuklu programda AC'yi hic oynatmiyordu.
            rate = float(rec["cost_per_qty"] or 0)
            if not rate:
                tq = float(rec["target_qty"] or 0)
                if tq:
                    rate = float(rec["target_cost"] or 0) / tq
            new_act_cost = float(act_cost) if act_cost is not None \
                else (new_act_qty * rate if rate else float(rec["act_reg_cost"] or 0))
            new_remain_cost = float(remain_cost) if remain_cost is not None \
                else (new_remain_qty * rate if rate
                      else max(0.0, float(rec["target_cost"] or 0) - new_act_cost))
            for label, value in (("actual_qty", new_act_qty),
                                 ("remaining_qty", new_remain_qty),
                                 ("actual_cost", new_act_cost),
                                 ("remaining_cost", new_remain_cost)):
                if value < 0:
                    raise ProgressError("%s negatif olamaz (%s): %s"
                                        % (label, rec["task_code"], value))

            planned.append((int(rec["taskrsrc_id"]), new_act_qty, new_remain_qty,
                            new_act_cost, new_remain_cost))
            changes.append({
                "taskrsrc_id": int(rec["taskrsrc_id"]),
                "task_code": rec["task_code"],
                "target_qty": float(rec["target_qty"] or 0),
                "actual_qty": {"from": float(rec["act_reg_qty"] or 0),
                               "to": new_act_qty},
                "remaining_qty": {"from": float(rec["remain_qty"] or 0),
                                  "to": new_remain_qty},
                "actual_cost": {"from": float(rec["act_reg_cost"] or 0),
                                "to": new_act_cost},
                "cost_per_qty": rate,
            })

        if dry:
            s.conn.rollback()
            return {"action": "set_assignment_actuals", "dry_run": True,
                    "proj_id": proj_id, "requested": len(updates),
                    "assignments": changes,
                    "note": "Hicbir sey yazilmadi. confirm=true ile calistirin."}

        for asg_id, aq, rq, ac, rc in planned:
            s.execute(
                "UPDATE TASKRSRC SET act_reg_qty = ?, remain_qty = ?, "
                "act_reg_cost = ?, remain_cost = ?, update_date = ?, "
                "update_user = ? WHERE taskrsrc_id = ?",
                aq, rq, ac, rc, s.stamp, s.user, asg_id)
        return {"action": "set_assignment_actuals", "proj_id": proj_id,
                "updated": len(planned), "assignments": changes}


def clear(params: Mapping[str, Any]) -> dict[str, Any]:
    """Put activities back to not-started -- undo a status update."""
    proj_id = int(params["proj_id"])
    codes = params.get("task_codes") or params.get("codes")
    dry = bool(params.get("dry_run"))
    if not dry:
        w.require_confirm(params, "Ilerleme silme")

    with w.open_session(params) as s:
        w.project_exists(s, proj_id)
        where = "proj_id = ? AND delete_session_id IS NULL"
        args: list[Any] = [proj_id]
        if codes:
            where += " AND task_code IN (%s)" % ", ".join("?" for _ in codes)
            args += list(codes)
        else:
            where += " AND status_code <> ?"
            args.append(STATUS_NOT_STARTED)
        affected = s.scalar("SELECT COUNT(*) FROM TASK WHERE " + where, *args)
        if dry:
            s.conn.rollback()
            return {"action": "clear", "dry_run": True, "proj_id": proj_id,
                    "would_clear": affected,
                    "scope": "verilen kodlar" if codes else "ilerleme girilmis tum aktiviteler"}
        s.execute(
            "UPDATE TASK SET status_code = ?, act_start_date = NULL, "
            "act_end_date = NULL, phys_complete_pct = 0, "
            "remain_drtn_hr_cnt = target_drtn_hr_cnt, act_work_qty = 0, "
            "remain_work_qty = target_work_qty, update_date = ?, "
            "update_user = ? WHERE " + where,
            STATUS_NOT_STARTED, s.stamp, s.user, *args)
        s.execute(
            "UPDATE TASKRSRC SET act_reg_qty = 0, act_ot_qty = 0, "
            "act_reg_cost = 0, act_ot_cost = 0, remain_qty = target_qty, "
            "remain_cost = target_cost, update_date = ?, update_user = ? "
            "WHERE proj_id = ? AND delete_session_id IS NULL",
            s.stamp, s.user, proj_id)
        return {"action": "clear", "proj_id": proj_id, "cleared": affected,
                "note": "Atama fiilleri de sifirlandi. Tarihler icin "
                        "p6_job action='schedule' calistirin."}


def set_data_date(params: Mapping[str, Any]) -> dict[str, Any]:
    """Move the project data date -- the cut-off a status update reports to."""
    proj_id = int(params["proj_id"])
    new_date = _parse_date(params.get("data_date"), "data_date")
    if new_date is None:
        raise ProgressError("data_date zorunlu (YYYY-MM-DD).")
    if not params.get("dry_run"):
        w.require_confirm(params, "Veri tarihi degistirme")
    with w.open_session(params) as s:
        w.project_exists(s, proj_id)
        old = _data_date(s, proj_id)
        late = s.scalar(
            "SELECT COUNT(*) FROM TASK WHERE proj_id = ? AND "
            "delete_session_id IS NULL AND (act_start_date > ? OR act_end_date > ?)",
            proj_id, new_date, new_date)
        if params.get("dry_run"):
            s.conn.rollback()
            return {"action": "set_data_date", "dry_run": True,
                    "proj_id": proj_id, "from": str(old)[:19],
                    "to": str(new_date)[:19],
                    "actuals_after_new_data_date": late}
        if late and not params.get("allow_future_actuals"):
            raise ProgressError(
                "%d aktivitenin fiili tarihi yeni veri tarihinden sonra "
                "kaliyor; P6 bunu kabul etmez. Once ilerlemeyi duzeltin veya "
                "allow_future_actuals=true verin." % late)
        s.execute("UPDATE PROJECT SET last_recalc_date = ? WHERE proj_id = ?",
                  new_date, proj_id)
        return {"action": "set_data_date", "proj_id": proj_id,
                "from": str(old)[:19], "to": str(new_date)[:19],
                "actuals_after_data_date": late,
                "note": "Tarihleri yeniden hesaplamak icin "
                        "p6_job action='schedule' calistirin."}


def _run_schedule(params: Mapping[str, Any], proj_id: int) -> dict[str, Any]:
    """Hand the recalculation to P6's own engine."""
    from . import db as p6db, jobs

    alias = p6db.resolve_alias(params.get("alias"))
    conn = p6db.connect_rw(alias)
    try:
        cur = conn.cursor()
        res = jobs.run_and_wait(
            cur, jobs.JT_SCHEDULE, [proj_id],
            int(params.get("user_id", 25)), "MCP_Progress",
            int(params.get("timeout_s", 300)), None)
        out = res.to_dict()
        out["alias"] = alias.name
        return out
    finally:
        conn.close()


ACTIONS = {
    "read": read,
    "set_progress": set_progress,
    "set_assignment_actuals": set_assignment_actuals,
    "clear": clear,
    "set_data_date": set_data_date,
}
