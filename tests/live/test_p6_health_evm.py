r"""p6_health + p6_evm kabul testi -- canli SQL Server veritabani gerektirir.

Bu test iddiaya degil KANITA bakar: DCMA sayilarini ham SQL ile bagimsiz
sayar ve tool'un dondurdugu rakamla karsilastirir. Ayni sekilde S-egrisinin
son kumulatif PV'si BAC'a esit olmali, aksi halde zaman-fazli dagitim
bozuktur.

    set P6_TEST_PROJ_ID=368
    set P6_TEST_XER=...\bukhtourcity.xer      (istege bagli, parite testi)
    python tests\live\test_p6_health_evm.py
"""
from __future__ import annotations

import json
import os
import sys
import tempfile

sys.path.insert(0, os.environ.get(
    "P6_REPO",
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
import p6_mcp_core as srv  # noqa: E402
from p6 import db as p6db  # noqa: E402

PROJ = int(os.environ.get("P6_TEST_PROJ_ID", "368"))
XER = os.environ.get("P6_TEST_XER", "")
OUT = os.environ.get("P6_TEST_OUT", "test_p6_health_evm.txt")

lines: list[str] = []
fails: list[str] = []


def p(msg: str = "") -> None:
    print(msg)
    lines.append(msg)


def check(label: str, cond: bool, detail: str = "") -> None:
    p("   %-56s %s%s" % (label, "OK" if cond else "HATA",
                         ("  " + detail) if detail else ""))
    if not cond:
        fails.append(label)


def h(params: dict) -> dict:
    return json.loads(srv.p6_health(params))


def e(params: dict) -> dict:
    return json.loads(srv.p6_evm(params))


def raw_counts() -> dict[str, int]:
    """Bagimsiz dogrulama kaynagi: ham SQL. Tool'un kendi cikti degil."""
    alias = p6db.resolve_alias(None)
    backend, _info = p6db.open_backend(alias, use_snapshot=False)
    try:
        def one(sql: str) -> int:
            return int(backend.select(sql, [])[0][0])

        t = "TASK t WHERE t.proj_id=%d AND t.delete_session_id IS NULL" % PROJ
        link = ("TASKPRED p JOIN TASK t ON t.task_id=p.task_id "
                "AND t.delete_session_id IS NULL WHERE t.proj_id=%d "
                "AND p.delete_session_id IS NULL" % PROJ)
        return {
            "tasks": one("SELECT COUNT(*) FROM " + t),
            "links": one("SELECT COUNT(*) FROM " + link),
            "leads": one("SELECT COUNT(*) FROM " + link + " AND p.lag_hr_cnt<0"),
            "lags": one("SELECT COUNT(*) FROM " + link + " AND p.lag_hr_cnt>0"),
            "fs": one("SELECT COUNT(*) FROM " + link + " AND p.pred_type='PR_FS'"),
            "neg_float": one("SELECT COUNT(*) FROM " + t + " AND t.total_float_hr_cnt<0"),
            "asg_qty": int(float(backend.select(
                "SELECT COALESCE(SUM(target_qty),0) FROM TASKRSRC "
                "WHERE proj_id=? AND delete_session_id IS NULL".replace(
                    "?", backend.param), [PROJ])[0][0])),
        }
    finally:
        backend.close()


def main() -> int:
    raw = raw_counts()
    p("=== ham SQL referansi === %s" % raw)

    # --- 1) DCMA, ham SQL ile karsilastirmali ------------------------------
    p("=== 1) p6_health assess_all ===")
    r = h({"action": "assess_all", "proj_id": PROJ})
    check("hata yok", "error" not in r, r.get("error", ""))
    rules = {x["id"]: x for x in r["rules"]}
    check("14 kural dondu", len(r["rules"]) == 14, str(len(r["rules"])))
    check("day_hr_cnt takvimden", "CALENDAR" in str(r["source"].get("day_hr_cnt_source")),
          str(r["source"].get("day_hr_cnt_source")))
    check("R3 leads = ham SQL", rules[3]["failed_count"] == raw["leads"],
          "%s vs %s" % (rules[3]["failed_count"], raw["leads"]))
    check("R4 lags = ham SQL", rules[4]["failed_count"] == raw["lags"],
          "%s vs %s" % (rules[4]["failed_count"], raw["lags"]))
    check("R5 FS%% = ham SQL", abs(rules[5]["actual"]
                                   - raw["fs"] / raw["links"] * 100) < 0.02,
          "%s vs %.2f" % (rules[5]["actual"], raw["fs"] / raw["links"] * 100))
    check("R8 negatif float = ham SQL", rules[8]["failed_count"] == raw["neg_float"],
          "%s vs %s" % (rules[8]["failed_count"], raw["neg_float"]))
    check("aktivite sayisi = ham SQL", r["task_count"] == raw["tasks"],
          "%s vs %s" % (r["task_count"], raw["tasks"]))

    p("=== 2) baseline uyarisi ===")
    check("baseline_source bildirildi", r.get("baseline_source") == "target_dates",
          str(r.get("baseline_source")))
    drift = (r.get("baseline_warnings") or [])
    check("gercek baseline yoksa uyari veriliyor", bool(drift),
          (drift[0][:60] + "...") if drift else "UYARI YOK")

    p("=== 3) drill_down ===")
    r3 = h({"action": "drill_down", "proj_id": PROJ, "rule_id": 8, "limit": 20})
    check("R8 drill_down sayisi kuralla ayni",
          r3["failed_count"] == rules[8]["failed_count"],
          "%s" % r3["failed_count"])
    check("aktiviteler isimli donuyor",
          all("code" in t for t in r3["tasks"]) if r3["tasks"] else True)
    bad = h({"action": "drill_down", "proj_id": PROJ, "rule_id": 99})
    check("gecersiz rule_id reddediliyor", "error" in bad)

    # --- 4) EVM ------------------------------------------------------------
    p("=== 4) p6_evm verify (RULE 16.A) ===")
    v = e({"action": "verify", "proj_id": PROJ})
    check("hata yok", "error" not in v, v.get("error", ""))
    check("BAC bagimsiz kaynakla eslesti", v.get("match") is True,
          "%s vs %s" % (v.get("bac_primary"), v.get("bac_independent")))
    check("birim bildirildi", v.get("units") in ("cost", "qty", "duration_h"),
          str(v.get("units")))
    check("aday BAC'lar gosteriliyor", isinstance(v.get("candidate_bac"), dict))

    p("=== 5) S-egrisi ===")
    tp = e({"action": "time_phased_evm", "proj_id": PROJ,
            "bucket": "month", "limit": 500})
    per = tp["periods"]
    check("donem uretildi", len(per) > 0, str(len(per)))
    monotonic = all(per[i]["pv"] <= per[i + 1]["pv"] + 1e-6
                    for i in range(len(per) - 1))
    check("PV monoton artan", monotonic)
    m = e({"action": "compute_metrics", "proj_id": PROJ})
    check("son kumulatif PV = BAC", abs(per[-1]["pv"] - m["bac"]) < 0.01,
          "%s vs %s" % (per[-1]["pv"], m["bac"]))

    p("=== 6) Earned Schedule korumasi ===")
    es = e({"action": "earned_schedule", "proj_id": PROJ})
    if es.get("earned_schedule_note"):
        check("gecen sure yokken SPI(t) None", es.get("spi_t") is None,
              str(es.get("earned_schedule_note"))[:50])
    else:
        check("SPI(t) hesaplandi", es.get("spi_t") is not None, str(es.get("spi_t")))

    p("=== 7) veri kalitesi ===")
    q = e({"action": "progress_data_quality", "proj_id": PROJ})
    check("uyari listesi dondu", isinstance(q.get("warnings"), list),
          "%s uyari" % q.get("warning_count"))

    # --- 8) snapshot dongusu -----------------------------------------------
    p("=== 8) snapshot -> history -> trend -> delta ===")
    snap = os.path.join(tempfile.gettempdir(), "p6_live_snap_test.json")
    if os.path.exists(snap):
        os.remove(snap)
    base = {"proj_id": PROJ, "snapshot_path": snap}
    s1 = e(dict(base, action="save_period_snapshot", tag="LIVE1"))
    check("snapshot yazildi", s1.get("total_snapshots") == 1, str(s1.get("total_snapshots")))
    check("DCMA snapshot'a dahil", s1.get("dcma_included") is True)
    e(dict(base, action="save_period_snapshot", tag="LIVE2"))
    hist = e(dict(base, action="get_period_history"))
    check("history 2 kayit", hist.get("count") == 2, str(hist.get("count")))
    tr = e(dict(base, action="trend"))
    check("trend serisi 2", tr.get("count") == 2, str(tr.get("count")))
    dl = e(dict(base, action="period_delta"))
    check("period_delta onceki snapshot'i buldu", dl.get("previous_tag") == "LIVE2",
          str(dl.get("previous_tag")))
    cmp_ = h(dict(base, action="compare"))
    check("DCMA compare onceki snapshot'i buldu",
          cmp_.get("previous") is not None, str(cmp_.get("previous_tag")))
    os.remove(snap)

    # --- 9) db / XER paritesi ----------------------------------------------
    if XER and os.path.exists(XER):
        p("=== 9) db <-> XER yapisal parite ===")
        rx = h({"action": "assess_all", "type": "xer", "path": XER})
        for rid in (1, 2, 3, 4, 5, 9, 11):
            check("R%d db == XER" % rid,
                  rules[rid]["actual"] == rx["rules"][rid - 1]["actual"],
                  "%s vs %s" % (rules[rid]["actual"], rx["rules"][rid - 1]["actual"]))
        p("   NOT: float/kritik kurallari (7,8,13) veri tarihine baglidir; "
          "XER farkli bir tarihte alinmissa esit olmasi BEKLENMEZ.")
    else:
        p("=== 9) XER paritesi ATLANDI (P6_TEST_XER verilmedi) ===")

    p()
    p("SONUC: %d kontrol, %d hata" % (
        sum(1 for x in lines if "OK" in x or "HATA" in x), len(fails)))
    for f in fails:
        p("  HATA: " + f)
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    p("Rapor: " + os.path.abspath(OUT))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
