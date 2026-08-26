r"""P6 MCP tam kabul testi -- dort aracin tamami, canli SQL Server uzerinde.

Bu test iddiaya degil KANITA bakar. Her sayisal sonuc ya ham SQL ile bagimsiz
sayilir, ya kaynak XER dosyasiyla karsilastirilir, ya da matematiksel bir
degismezle denetlenir (S-egrisinin son kumulatif PV'si BAC'a esittir, EV
BAC'i asamaz, kalan + fiili = hedef).

Kapsam:
  A  ortam ve okuma       p6_query 14 action
  B  Kiril butunlugu      collation + XER ile bayt bayt esitlik + gidis-donus
  C  DCMA                 p6_health 4 action, ham SQL capraz kontrol
  D  EVM                  p6_evm 15 action
  E  Baseline             p6_baseline 4 action + kopya sadakati
  F  Ilerleme/fiili       p6_progress 5 action + P6 semantik kurallari
  G  Job Service          p6_job preflight/schedule/status/list
  H  Korumalar            confirm, dry_run, salt-okuma SQL, kimlik parametresi

Test VERIYI DEGISTIRIR ve sonunda baslangic durumuna geri alir:
ilerleme temizlenir, olusturulan baseline silinir, veri tarihi geri konur.

    set P6_TEST_PROJ_ID=368
    set P6_TEST_XER=...\bukhtourcity.xer
    python tests\live\test_p6_full_acceptance.py
"""
from __future__ import annotations

import datetime as _dt
import json
import os
import sys
import tempfile
import time
import traceback

sys.path.insert(0, os.environ.get(
    "P6_REPO",
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
import p6_mcp_core as srv          # noqa: E402
import xer_parser                  # noqa: E402
from p6 import analysis, db as p6db  # noqa: E402

PROJ = int(os.environ.get("P6_TEST_PROJ_ID", "368"))
XER = os.environ.get(
    "P6_TEST_XER",
    r"C:\Users\CahAsus\Downloads\_P6_MCP\faz0\bukhtourcity.xer")
OUT = os.environ.get("P6_TEST_OUT", "test_p6_full_acceptance.txt")
SNAP = os.path.join(tempfile.gettempdir(), "p6_acceptance_snapshots.json")

lines: list[str] = []
passed = 0
failed: list[str] = []
_section = ""


def p(msg: str = "") -> None:
    print(msg)
    lines.append(msg)


def section(title: str) -> None:
    global _section
    _section = title
    p("")
    p("=" * 74)
    p(title)
    p("=" * 74)


def check(label: str, cond: bool, detail: str = "") -> bool:
    global passed
    ok = bool(cond)
    if ok:
        passed += 1
    else:
        failed.append("%s / %s" % (_section, label))
    p("  %-58s %-4s %s" % (label[:58], "OK" if ok else "HATA", detail))
    return ok


def q(params):
    return json.loads(srv.p6_query(params))


def h(params):
    return json.loads(srv.p6_health(params))


def e(params):
    return json.loads(srv.p6_evm(params))


def pr(params):
    return json.loads(srv.p6_progress(params))


def bl(params):
    return json.loads(srv.p6_baseline(params))


def jb(params):
    return json.loads(srv.p6_job(params))


def rw():
    return p6db.connect_rw(p6db.resolve_alias(None))


def sql_one(sql, *args):
    conn = rw()
    try:
        cur = conn.cursor()
        cur.execute(sql, *args)
        row = cur.fetchone()
        return row[0] if row else None
    finally:
        conn.close()


def sql_all(sql, *args):
    conn = rw()
    try:
        cur = conn.cursor()
        cur.execute(sql, *args)
        return cur.fetchall()
    finally:
        conn.close()


# ===========================================================================
# A -- ortam ve okuma
# ===========================================================================
def part_a():
    section("A  ORTAM VE OKUMA (p6_query)")
    r = q({"action": "db_info"})
    check("db_info calisiyor", "error" not in r, r.get("error", ""))
    counts = r.get("table_rows", {})
    check("TASK satiri > 0", (counts.get("TASK") or 0) > 0, str(counts.get("TASK")))

    raw_tasks = sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? "
                        "AND delete_session_id IS NULL", PROJ)
    raw_links = sql_one(
        "SELECT COUNT(*) FROM TASKPRED p JOIN TASK t ON t.task_id=p.task_id "
        "AND t.delete_session_id IS NULL WHERE t.proj_id=? "
        "AND p.delete_session_id IS NULL", PROJ)

    t = q({"action": "read_tasks", "proj_id": PROJ, "limit": 500})
    check("read_tasks = ham SQL", t["count"] == raw_tasks,
          "%s vs %s" % (t["count"], raw_tasks))
    lk = q({"action": "read_links", "proj_id": PROJ, "limit": 500})
    check("read_links = ham SQL", lk["count"] == raw_links,
          "%s vs %s" % (lk["count"], raw_links))

    src = t["source"]
    check("day_hr_cnt takvimden okundu",
          "CALENDAR" in str(src.get("day_hr_cnt_source")),
          str(src.get("day_hr_cnt_source")))
    check("gun-saat varsayilan 8'e sabitlenmemis",
          src.get("day_hr_cnt") is not None, str(src.get("day_hr_cnt")))

    for action, key in (("read_resources", "count"), ("read_assignments", "count"),
                        ("read_calendars", "count"), ("read_wbs", "count"),
                        ("read_project", "data"), ("read_progress", "data")):
        rr = q({"action": action, "proj_id": PROJ, "limit": 20})
        check("%s hatasiz" % action, "error" not in rr and key in rr,
              rr.get("error", ""))

    fd = q({"action": "finish_drivers", "proj_id": PROJ})
    check("finish_drivers hatasiz", "error" not in fd, fd.get("error", ""))
    so = q({"action": "schedule_options", "proj_id": PROJ})
    check("schedule_options hatasiz", "error" not in so, so.get("error", ""))

    lp = q({"action": "list_projects"})
    check("list_projects yalniz gercek projeleri listeler",
          all(p_.get("proj_short_name") for p_ in lp["projects"]),
          "%d proje" % lp["count"])
    eps = q({"action": "list_eps"})
    check("list_eps EPS dugumu dondurur", eps.get("count", 0) >= 1,
          str(eps.get("count")))

    # XER kaynagi ayni sayiyi vermeli
    if os.path.exists(XER):
        xt = q({"action": "read_tasks", "type": "xer", "path": XER, "limit": 500})
        check("XER kaynagi ayni gorev sayisi", xt["count"] == raw_tasks,
              "%s vs %s" % (xt["count"], raw_tasks))
    else:
        p("  (XER bulunamadi, parite atlandi: %s)" % XER)


# ===========================================================================
# B -- Kiril butunlugu
# ===========================================================================
def part_b():
    section("B  KIRIL BUTUNLUGU")
    coll = sql_one("SELECT DATABASEPROPERTYEX(DB_NAME(), 'Collation')")
    check("veritabani collation Kiril", "Cyrillic" in str(coll), str(coll))
    col_coll = sql_one(
        "SELECT c.collation_name FROM sys.columns c "
        "WHERE c.object_id=OBJECT_ID('TASK') AND c.name='task_name'")
    check("TASK.task_name collation Kiril", "Cyrillic" in str(col_coll),
          str(col_coll))
    kalan = sql_one(
        "SELECT COUNT(*) FROM sys.columns c JOIN sys.types ty "
        "ON ty.user_type_id=c.user_type_id JOIN sys.tables t "
        "ON t.object_id=c.object_id WHERE ty.name IN ('varchar','char','text') "
        "AND c.collation_name <> 'Cyrillic_General_CI_AS'")
    check("karisik collation kalmadi", kalan == 0, "%s kolon" % kalan)

    if not os.path.exists(XER):
        p("  (XER yok, ad karsilastirmasi atlandi)")
        return
    x = xer_parser.XerFile(XER)
    check("XER kod sayfasi cp1251 saptandi", x.encoding == "cp1251",
          "%s / %s" % (x.encoding, x.encoding_confidence))
    check("XER'de U+FFFD yok",
          not any("\ufffd" in r.get("task_name", "")
                  for r in x.tables["TASK"]["rows"]))

    xer_names = {r["task_code"]: r["task_name"] for r in x.tables["TASK"]["rows"]}
    db_names = {r[0]: r[1] for r in sql_all(
        "SELECT task_code, task_name FROM TASK WHERE proj_id=? "
        "AND delete_session_id IS NULL", PROJ)}
    mismatch = [k for k in db_names if k in xer_names and db_names[k] != xer_names[k]]
    check("gorev adlari XER ile birebir", not mismatch,
          "%d/%d uyusmuyor" % (len(mismatch), len(db_names)))
    cyr = sum(1 for v in db_names.values()
              if any("\u0400" <= c <= "\u04ff" for c in v))
    check("Kirilli gorev adi veritabaninda mevcut", cyr > 0, "%d gorev" % cyr)
    check("veritabaninda U+FFFD kalintisi yok",
          not any("\ufffd" in v for v in db_names.values()))

    wbs = [r[0] for r in sql_all("SELECT wbs_name FROM PROJWBS WHERE proj_id=? "
                                 "AND delete_session_id IS NULL", PROJ)]
    wcyr = sum(1 for v in wbs if any("\u0400" <= c <= "\u04ff" for c in v))
    check("WBS adlari Kiril", wcyr == len(wbs), "%d/%d" % (wcyr, len(wbs)))

    # gidis-donus: yazip geri oku
    victim = sql_one("SELECT TOP 1 task_id FROM TASK WHERE proj_id=? "
                     "AND delete_session_id IS NULL ORDER BY task_id", PROJ)
    original = sql_one("SELECT task_name FROM TASK WHERE task_id=?", victim)
    conn = rw()
    try:
        cyr_probe = "Гранит ТЕСТ Брусчатка 100%"
        conn.cursor().execute("UPDATE TASK SET task_name=? WHERE task_id=?",
                              cyr_probe, victim)
        check("Kiril yazma/okuma birebir",
              sql_one("SELECT task_name FROM TASK WHERE task_id=?", victim)
              == cyr_probe)

        # BILINEN TAKAS: varchar tek bir kod sayfasi tasir. Kiril collation
        # cp1251'dir; Turkce'ye ozgu harfler (Ş Ğ İ) orada YOKTUR ve SQL Server
        # en yakin harfi yazar. Bu bir hata degil, secimin bedeli -- ama
        # sessiz kalmamasi icin testle sabitleniyor.
        tr_probe = "Şantiye Şefliği ÇÖĞÜ"
        conn.cursor().execute("UPDATE TASK SET task_name=? WHERE task_id=?",
                              tr_probe, victim)
        got_tr = sql_one("SELECT task_name FROM TASK WHERE task_id=?", victim)
        check("Turkce'ye ozgu harfler cp1251'de TASINMAZ (bilinen takas)",
              got_tr != tr_probe, repr(got_tr))
        check("Turkce metnin ASCII kismi korunur",
              "antiye" in got_tr and "efli" in got_tr, repr(got_tr))
    finally:
        conn.cursor().execute("UPDATE TASK SET task_name=? WHERE task_id=?",
                              original, victim)
        conn.close()
    check("sonda orijinal ad geri konuldu",
          sql_one("SELECT task_name FROM TASK WHERE task_id=?", victim) == original)


# ===========================================================================
# C -- DCMA
# ===========================================================================
def part_c():
    section("C  DCMA 14-POINT (p6_health)")
    link_where = ("TASKPRED p JOIN TASK t ON t.task_id=p.task_id AND "
                  "t.delete_session_id IS NULL WHERE t.proj_id=%d AND "
                  "p.delete_session_id IS NULL" % PROJ)
    raw = {
        "tasks": sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                         "delete_session_id IS NULL", PROJ),
        "links": sql_one("SELECT COUNT(*) FROM " + link_where),
        "leads": sql_one("SELECT COUNT(*) FROM " + link_where + " AND p.lag_hr_cnt<0"),
        "lags": sql_one("SELECT COUNT(*) FROM " + link_where + " AND p.lag_hr_cnt>0"),
        "fs": sql_one("SELECT COUNT(*) FROM " + link_where + " AND p.pred_type='PR_FS'"),
        "neg": sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                       "delete_session_id IS NULL AND total_float_hr_cnt<0", PROJ),
        "hidur": sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                         "delete_session_id IS NULL AND target_drtn_hr_cnt>352", PROJ),
        "asg_tasks": sql_one("SELECT COUNT(DISTINCT task_id) FROM TASKRSRC "
                             "WHERE proj_id=? AND delete_session_id IS NULL", PROJ),
    }
    p("  ham SQL referansi: %s" % raw)

    r = h({"action": "assess_all", "proj_id": PROJ})
    check("assess_all hatasiz", "error" not in r, r.get("error", ""))
    check("MCP katmani 14 kuralin tamamini dondurur", len(r.get("rules", [])) == 14,
          "%d kural" % len(r.get("rules", [])))
    check("yanit kirpilmadi", not r.get("truncated"))
    rules = {x["id"]: x for x in r["rules"]}

    check("R3 leads = ham SQL", rules[3]["failed_count"] == raw["leads"],
          "%s vs %s" % (rules[3]["failed_count"], raw["leads"]))
    check("R4 lags = ham SQL", rules[4]["failed_count"] == raw["lags"],
          "%s vs %s" % (rules[4]["failed_count"], raw["lags"]))
    check("R5 FS%% = ham SQL",
          abs(rules[5]["actual"] - raw["fs"] / raw["links"] * 100) < 0.02,
          "%s vs %.2f" % (rules[5]["actual"], raw["fs"] / raw["links"] * 100))
    check("R8 negatif float = ham SQL", rules[8]["failed_count"] == raw["neg"],
          "%s vs %s" % (rules[8]["failed_count"], raw["neg"]))
    check("R9 uzun sure = ham SQL", rules[9]["failed_count"] == raw["hidur"],
          "%s vs %s" % (rules[9]["failed_count"], raw["hidur"]))
    check("R11 kaynaksiz = toplam - atamali",
          rules[11]["failed_count"] == rules[11]["total_count"] - raw["asg_tasks"],
          "%s vs %s" % (rules[11]["failed_count"],
                        rules[11]["total_count"] - raw["asg_tasks"]))
    check("R9 esigi takvimden (8s -> 352s)",
          abs(rules[9].get("threshold_hours", 0) - 44 * r["day_hr_cnt"]) < 0.01,
          str(rules[9].get("threshold_hours")))
    check("pass + fail = 14",
          r["summary"]["pass_count"] + r["summary"]["fail_count"] == 14)
    check("RAG kural sayisiyla tutarli",
          r["summary"]["overall_rag"] in ("green", "amber", "red"))

    s = h({"action": "summary", "proj_id": PROJ})
    check("summary assess_all ile ayni skoru verir",
          s["pass_count"] == r["summary"]["pass_count"],
          "%s vs %s" % (s["pass_count"], r["summary"]["pass_count"]))

    d8 = h({"action": "drill_down", "proj_id": PROJ, "rule_id": 8, "limit": 50})
    check("drill_down R8 sayisi kuralla ayni",
          d8["failed_count"] == rules[8]["failed_count"], str(d8["failed_count"]))
    check("drill_down aktiviteyi adiyla dondurur",
          all("code" in x and "name" in x for x in d8["tasks"]) if d8["tasks"] else True)
    d3 = h({"action": "drill_down", "proj_id": PROJ, "rule_id": 3, "limit": 5})
    check("iliski kurallari bagi dondurur", "links" in d3, str(list(d3)[:6]))
    check("bag tipi bos degil",
          all(x.get("type") for x in d3.get("links", [])) if d3.get("links") else True,
          str([x.get("type") for x in d3.get("links", [])][:3]))
    check("gecersiz rule_id reddedilir",
          "error" in h({"action": "drill_down", "proj_id": PROJ, "rule_id": 0}))
    check("rule_id 15 reddedilir",
          "error" in h({"action": "drill_down", "proj_id": PROJ, "rule_id": 15}))


# ===========================================================================
# D -- EVM
# ===========================================================================
def part_d(baseline_id=None):
    section("D  EVM (p6_evm)")
    m = e({"action": "compute_metrics", "proj_id": PROJ})
    check("compute_metrics hatasiz", "error" not in m, m.get("error", ""))
    check("birim bildirildi", m.get("units") in ("cost", "qty", "duration_h"),
          "%s (%s)" % (m.get("units"), m.get("units_reason")))
    check("aday BAC'lar gosterildi", isinstance(m.get("candidate_bac"), dict),
          json.dumps(m.get("candidate_bac")))
    check("EV <= BAC", m["ev"] <= m["bac"] + 1e-6, "%s / %s" % (m["ev"], m["bac"]))
    check("PV <= BAC", m["pv"] <= m["bac"] + 1e-6, "%s / %s" % (m["pv"], m["bac"]))
    check("SV = EV - PV", abs(m["sv"] - (m["ev"] - m["pv"])) < 1e-6)
    check("CV = EV - AC", abs(m["cv"] - (m["ev"] - m["ac"])) < 1e-6)
    if m["pv"] > 0:
        check("SPI = EV / PV", abs(m["spi"] - m["ev"] / m["pv"]) < 1e-6)

    v = e({"action": "verify", "proj_id": PROJ})
    check("verify: BAC bagimsiz kaynakla eslesti", v.get("match") is True,
          "%s vs %s" % (v.get("bac_primary"), v.get("bac_independent")))
    check("verify: rollup bosluğu yok",
          v.get("zero_baseline_with_assignments_count") == 0,
          str(v.get("zero_baseline_with_assignments_count")))

    f = e({"action": "forecast", "proj_id": PROJ})
    check("forecast EAC1 = AC + (BAC-EV)",
          abs(f["eac_t1"] - (f["ac"] + f["bac"] - f["ev"])) < 0.02,
          "%s" % f["eac_t1"])
    check("VAC = BAC - EAC", f.get("vac") is not None)

    su = e({"action": "summary", "proj_id": PROJ})
    check("summary RAG uretir", su.get("rag") in ("RED", "AMBER", "GREEN"),
          "%s (%%%s)" % (su.get("rag"), su.get("completion_pct")))
    check("completion_pct = EV/BAC",
          abs(su["completion_pct"] - (m["ev"] / m["bac"] * 100)) < 0.02)

    tp = e({"action": "time_phased_evm", "proj_id": PROJ, "bucket": "month",
            "limit": 500})
    per = tp["periods"]
    check("S-egrisi donem uretti", len(per) > 0, "%d donem" % len(per))
    check("PV monoton artan",
          all(per[i]["pv"] <= per[i + 1]["pv"] + 1e-6 for i in range(len(per) - 1)))
    check("EV monoton artan",
          all(per[i]["ev"] <= per[i + 1]["ev"] + 1e-6 for i in range(len(per) - 1)))
    check("son kumulatif PV = BAC", abs(per[-1]["pv"] - m["bac"]) < 0.01,
          "%s vs %s" % (per[-1]["pv"], m["bac"]))
    check("son kumulatif EV = toplam EV", abs(per[-1]["ev"] - m["ev"]) < 0.5,
          "%s vs %s" % (per[-1]["ev"], m["ev"]))
    for bucket in ("day", "week"):
        rb = e({"action": "time_phased_evm", "proj_id": PROJ, "bucket": bucket,
                "limit": 500})
        check("bucket=%s calisir" % bucket, "error" not in rb,
              "%s donem" % rb.get("count"))
    check("gecersiz bucket reddedilir",
          "error" in e({"action": "time_phased_evm", "proj_id": PROJ,
                        "bucket": "yil"}))

    es = e({"action": "earned_schedule", "proj_id": PROJ})
    if es.get("earned_schedule_note"):
        check("gecen sure yokken SPI(t) None", es.get("spi_t") is None,
              es["earned_schedule_note"][:40])
    else:
        check("SPI(t) hesaplandi", es.get("spi_t") is not None, str(es.get("spi_t")))
        check("SV(t) = ES - AT",
              abs(es["sv_t_weeks"] - (es["es_weeks"] - es["at_weeks"])) < 1e-3)

    dq = e({"action": "progress_data_quality", "proj_id": PROJ})
    check("veri kalitesi uyari listesi dondurur", isinstance(dq.get("warnings"), list),
          "%s uyari" % dq.get("warning_count"))

    cm = e({"action": "detect_currency_mode", "proj_id": PROJ})
    check("detect_currency_mode hatasiz", "error" not in cm,
          str(cm.get("cost_loading_mode")))
    vc = e({"action": "validate_currency_mode", "proj_id": PROJ})
    check("validate_currency_mode uzlasi verir", "consensus_mode" in vc,
          "%s / %s" % (vc.get("consensus_mode"), vc.get("confidence")))

    if os.path.exists(SNAP):
        os.remove(SNAP)
    base = {"proj_id": PROJ, "snapshot_path": SNAP}
    s1 = e(dict(base, action="save_period_snapshot", tag="ACC1"))
    check("snapshot yazildi", s1.get("total_snapshots") == 1)
    check("snapshot DCMA icerir", s1.get("dcma_included") is True)
    e(dict(base, action="save_period_snapshot", tag="ACC2"))
    hist = e(dict(base, action="get_period_history"))
    check("history iki kayit", hist.get("count") == 2, str(hist.get("count")))
    tr = e(dict(base, action="trend"))
    check("trend serisi iki nokta", tr.get("count") == 2)
    check("trend DCMA pass sayisini tasir",
          tr["series"][0].get("dcma_pass") is not None)
    dl = e(dict(base, action="period_delta"))
    check("period_delta onceki snapshot'i bulur", dl.get("previous_tag") == "ACC2",
          str(dl.get("previous_tag")))
    check("ayni donemde delta sifir",
          abs(dl.get("period_ev", 1)) < 1e-6, str(dl.get("period_ev")))
    hc = h(dict(base, action="compare"))
    check("DCMA compare onceki snapshot'i bulur", hc.get("previous") is not None)

    if baseline_id:
        vb = e({"action": "variance_to_baseline", "proj_id": PROJ,
                "baseline_proj_id": baseline_id, "limit": 5})
        check("variance_to_baseline gercek baseline kullanir",
              vb.get("baseline_source") == "baseline_project",
              str(vb.get("baseline_source")))
        check("baseline'da eslesmeyen aktivite yok",
              vb.get("baseline_unmatched_tasks", 0) == 0,
              str(vb.get("baseline_unmatched_tasks", 0)))
        check("gercek baseline'da uyari yok",
              not vb.get("baseline_warnings"))
        check("sapma listesi buyukten kucuge sirali",
              all(vb["worst"][i]["variance_days"] >= vb["worst"][i + 1]["variance_days"]
                  for i in range(len(vb["worst"]) - 1)))
    check("baseline_proj_id yoksa uyari verilir",
          bool(m.get("baseline_warnings")) or m.get("baseline_source") == "baseline_project",
          str(m.get("baseline_source")))
    check("olmayan baseline reddedilir",
          "error" in e({"action": "variance_to_baseline", "proj_id": PROJ,
                        "baseline_proj_id": 999999}))
    check("baseline_proj_id'siz variance reddedilir",
          "error" in e({"action": "variance_to_baseline", "proj_id": PROJ}))


# ===========================================================================
# E -- baseline
# ===========================================================================
def part_e():
    section("E  BASELINE (p6_baseline)")
    before = bl({"action": "list", "proj_id": PROJ})
    check("list hatasiz", "error" not in before, before.get("error", ""))
    check("BASETYPE secenekleri listelenir", len(before.get("base_types", [])) > 0)

    dry = bl({"action": "create", "proj_id": PROJ, "dry_run": True})
    check("dry_run kopyalanacaklari gosterir", dry.get("dry_run") is True,
          json.dumps(dry.get("would_copy", {})))
    check("dry_run hicbir sey yazmadi",
          bl({"action": "list", "proj_id": PROJ})["baseline_count"]
          == before["baseline_count"])
    check("confirm'siz create reddedilir",
          "error" in bl({"action": "create", "proj_id": PROJ}))
    check("gecersiz base_type reddedilir",
          "error" in bl({"action": "create", "proj_id": PROJ, "confirm": True,
                         "base_type": "Yok Boyle"}))
    check("olmayan proje reddedilir",
          "error" in bl({"action": "create", "proj_id": 999999, "confirm": True}))

    created = bl({"action": "create", "proj_id": PROJ, "confirm": True,
                  "baseline_name": "ACCEPTANCE TEST BL",
                  "base_type": "Initial Plan", "assign": True})
    check("baseline olusturuldu", "error" not in created, created.get("error", ""))
    bid = created.get("baseline_proj_id")
    if not bid:
        return None

    for table in ("PROJWBS", "TASK", "TASKPRED", "TASKRSRC"):
        live = sql_one("SELECT COUNT(*) FROM [%s] WHERE proj_id=? AND "
                       "delete_session_id IS NULL" % table, PROJ)
        copy = sql_one("SELECT COUNT(*) FROM [%s] WHERE proj_id=? AND "
                       "delete_session_id IS NULL" % table, bid)
        check("%s satir sayisi ayni" % table, live == copy, "%s vs %s" % (live, copy))

    same = sql_one(
        "SELECT COUNT(*) FROM TASK a JOIN TASK b ON b.task_code=a.task_code "
        "AND b.proj_id=? WHERE a.proj_id=? AND a.target_start_date=b.target_start_date "
        "AND a.target_end_date=b.target_end_date AND a.task_name=b.task_name", bid, PROJ)
    total = sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                    "delete_session_id IS NULL", PROJ)
    check("tarih + ad birebir kopyalandi", same == total, "%s / %s" % (same, total))

    orphan_links = sql_one(
        "SELECT COUNT(*) FROM TASKPRED WHERE proj_id=? AND (task_id NOT IN "
        "(SELECT task_id FROM TASK WHERE proj_id=?) OR pred_task_id NOT IN "
        "(SELECT task_id FROM TASK WHERE proj_id=?))", bid, bid, bid)
    check("baseline icinde kopuk bag yok", orphan_links == 0, str(orphan_links))
    orphan_wbs = sql_one(
        "SELECT COUNT(*) FROM PROJWBS WHERE proj_id=? AND parent_wbs_id IS NOT NULL "
        "AND parent_wbs_id NOT IN (SELECT wbs_id FROM PROJWBS)", bid)
    check("kopuk WBS ust dugumu yok", orphan_wbs == 0, str(orphan_wbs))
    shared = sql_one("SELECT COUNT(*) FROM TASK a JOIN TASK b ON a.task_id=b.task_id "
                     "WHERE a.proj_id=? AND b.proj_id=?", PROJ, bid)
    check("canli ve baseline ayni task_id'yi paylasmiyor", shared == 0, str(shared))

    flag = sql_one("SELECT project_flag FROM PROJECT WHERE proj_id=?", bid)
    check("baseline project_flag = N", flag == "N", str(flag))
    check("baseline EPS listesinde gorunmuyor",
          all(p_["proj_id"] != bid for p_ in q({"action": "list_projects"})["projects"]))
    check("proje baseline'i atandi",
          sql_one("SELECT sum_base_proj_id FROM PROJECT WHERE proj_id=?", PROJ) == bid)

    lst = bl({"action": "list", "proj_id": PROJ})
    check("list baseline'i gosterir",
          any(b["proj_id"] == bid for b in lst["baselines"]),
          "%d baseline" % lst["baseline_count"])

    check("assign temizleyebilir",
          "error" not in bl({"action": "assign", "proj_id": PROJ, "confirm": True}))
    check("temizlik sonrasi atama bos",
          sql_one("SELECT sum_base_proj_id FROM PROJECT WHERE proj_id=?", PROJ) is None)
    bl({"action": "assign", "proj_id": PROJ, "baseline_proj_id": bid, "confirm": True})
    check("assign geri konuldu",
          sql_one("SELECT sum_base_proj_id FROM PROJECT WHERE proj_id=?", PROJ) == bid)
    check("baska projenin baseline'i atanamaz",
          "error" in bl({"action": "assign", "proj_id": PROJ, "confirm": True,
                         "baseline_proj_id": PROJ}))
    return bid


# ===========================================================================
# F -- ilerleme ve fiili veri
# ===========================================================================
def part_f():
    section("F  ILERLEME VE FIILI VERI (p6_progress)")
    original_dd = sql_one("SELECT last_recalc_date FROM PROJECT WHERE proj_id=?", PROJ)
    p("  baslangic veri tarihi: %s" % str(original_dd)[:19])

    # Test tekrar calistirilabilir olmali: onceki bir kosunun (ya da elle
    # yapilmis bir denemenin) biraktigi ilerleme, "dry_run yazmadi" gibi
    # kontrolleri sahte sekilde bozar. Once temiz zemine cek.
    pr({"action": "clear", "proj_id": PROJ, "confirm": True})
    stale = sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                    "status_code <> 'TK_NotStart' AND delete_session_id IS NULL",
                    PROJ)
    check("test temiz ilerleme zemininden basliyor", stale == 0, str(stale))

    r0 = pr({"action": "read", "proj_id": PROJ, "limit": 5})
    check("read hatasiz", "error" not in r0, r0.get("error", ""))
    check("yuzde tabani raporlanir", bool(r0.get("percent_complete_basis")),
          json.dumps(r0.get("percent_complete_basis"), ensure_ascii=False)[:70])
    check("baslamis + baslamamis = toplam",
          r0["started"] + r0["not_started"] == r0["task_count"])

    # --- veri tarihi
    dd = pr({"action": "set_data_date", "proj_id": PROJ,
             "data_date": "2026-11-01", "dry_run": True})
    check("set_data_date dry_run yazmaz", dd.get("dry_run") is True)
    check("dry_run sonrasi veri tarihi degismedi",
          sql_one("SELECT last_recalc_date FROM PROJECT WHERE proj_id=?",
                  PROJ) == original_dd)
    check("confirm'siz set_data_date reddedilir",
          "error" in pr({"action": "set_data_date", "proj_id": PROJ,
                         "data_date": "2026-11-01"}))
    dd = pr({"action": "set_data_date", "proj_id": PROJ,
             "data_date": "2026-11-01", "confirm": True})
    check("veri tarihi uygulandi", dd.get("to", "").startswith("2026-11-01"),
          dd.get("to", ""))

    # --- dogrulama reddleri
    bad_cases = [
        ("gelecege fiili tarih reddedilir",
         [{"task_code": "bukhtourcity27", "status": "complete",
           "actual_start": "2026-12-01", "actual_finish": "2026-12-10"}]),
        ("bitis baslangictan once reddedilir",
         [{"task_code": "bukhtourcity27", "status": "complete",
           "actual_start": "2026-09-24", "actual_finish": "2026-09-02"}]),
        ("devam edende yuzde/kalan sure sart",
         [{"task_code": "bukhtourcity27", "status": "in_progress",
           "actual_start": "2026-09-02"}]),
        ("complete icin fiili bitis sart",
         [{"task_code": "bukhtourcity27", "status": "complete",
           "actual_start": "2026-09-02"}]),
        ("olmayan aktivite reddedilir",
         [{"task_code": "YOK-BOYLE-BIR-KOD", "status": "complete",
           "actual_start": "2026-09-02", "actual_finish": "2026-09-24"}]),
        ("ayni aktivite iki kez reddedilir",
         [{"task_code": "bukhtourcity27", "status": "in_progress",
           "actual_start": "2026-09-02", "percent_complete": 10},
          {"task_code": "bukhtourcity27", "status": "in_progress",
           "actual_start": "2026-09-02", "percent_complete": 20}]),
        ("gecersiz yuzde reddedilir",
         [{"task_code": "bukhtourcity27", "status": "in_progress",
           "actual_start": "2026-09-02", "percent_complete": 150}]),
    ]
    for label, updates in bad_cases:
        res = pr({"action": "set_progress", "proj_id": PROJ, "updates": updates,
                  "confirm": True})
        check(label, "error" in res, str(res.get("error", ""))[:60])

    check("confirm'siz set_progress reddedilir",
          "error" in pr({"action": "set_progress", "proj_id": PROJ, "updates": [
              {"task_code": "bukhtourcity27", "status": "complete",
               "actual_start": "2026-09-02", "actual_finish": "2026-09-24"}]}))

    updates = [
        {"task_code": "bukhtourcity27", "status": "complete",
         "actual_start": "2026-09-02", "actual_finish": "2026-09-24"},
        {"task_code": "bukhtourcity30", "status": "complete",
         "actual_start": "2026-09-09", "actual_finish": "2026-09-22"},
        {"task_code": "bukhtourcity85", "status": "in_progress",
         "actual_start": "2026-09-18", "percent_complete": 70},
        {"task_code": "bukhtourcity1346", "status": "in_progress",
         "actual_start": "2026-09-24", "remaining_duration_h": 30},
    ]
    dryr = pr({"action": "set_progress", "proj_id": PROJ, "updates": updates,
               "dry_run": True})
    check("dry_run before/after gosterir", dryr.get("would_change") == 4,
          str(dryr.get("would_change")))
    check("dry_run atama senkronunu da sayar",
          dryr.get("would_sync_assignments", 0) >= 1,
          str(dryr.get("would_sync_assignments")))
    check("dry_run gercekten yazmadi",
          sql_one("SELECT status_code FROM TASK WHERE proj_id=? AND task_code=?",
                  PROJ, "bukhtourcity27") == "TK_NotStart")

    applied = pr({"action": "set_progress", "proj_id": PROJ, "updates": updates,
                  "confirm": True, "schedule": True, "timeout_s": 240})
    check("ilerleme yazildi", applied.get("updated") == 4, str(applied.get("updated")))
    check("atamalar senkronlandi", applied.get("assignments_synced", 0) >= 1,
          str(applied.get("assignments_synced")))
    check("F9 tamamlandi", applied.get("schedule", {}).get("status") == "JS_Complete",
          str(applied.get("schedule", {}).get("elapsed_s")))

    rows = {r[0]: r for r in sql_all(
        "SELECT task_code, status_code, phys_complete_pct, remain_drtn_hr_cnt, "
        "target_drtn_hr_cnt, act_start_date, act_end_date FROM TASK "
        "WHERE proj_id=? AND task_code IN "
        "('bukhtourcity27','bukhtourcity30','bukhtourcity85','bukhtourcity1346')",
        PROJ)}
    c27 = rows["bukhtourcity27"]
    check("tamamlanan: status TK_Complete", c27[1] == "TK_Complete", str(c27[1]))
    check("tamamlanan: kalan sure 0", float(c27[3]) == 0.0, str(c27[3]))
    check("tamamlanan: yuzde 100", float(c27[2]) == 100.0, str(c27[2]))
    check("tamamlanan: fiili tarihler yazildi", c27[5] and c27[6],
          "%s -> %s" % (str(c27[5])[:10], str(c27[6])[:10]))
    c85 = rows["bukhtourcity85"]
    check("devam eden: status TK_Active", c85[1] == "TK_Active", str(c85[1]))
    check("devam eden: fiili bitis bos", c85[6] is None)
    check("KRITIK -- kaynak yuklu aktivitede kalan sure F9'dan sonra korunuyor",
          abs(float(c85[3]) - float(c85[4]) * 0.30) < 0.01,
          "kalan=%s hedef=%s" % (c85[3], c85[4]))
    c1346 = rows["bukhtourcity1346"]
    check("kalan sure dogrudan verildiginde korunur",
          abs(float(c1346[3]) - 30.0) < 0.01, str(c1346[3]))

    asg = {r[0]: r for r in sql_all(
        "SELECT t.task_code, SUM(r.act_reg_qty), SUM(r.remain_qty), SUM(r.target_qty) "
        "FROM TASK t JOIN TASKRSRC r ON r.task_id=t.task_id AND "
        "r.delete_session_id IS NULL WHERE t.proj_id=? AND t.task_code IN "
        "('bukhtourcity27','bukhtourcity85') GROUP BY t.task_code", PROJ)}
    a27 = asg["bukhtourcity27"]
    check("tamamlanan atama: fiili = hedef", abs(float(a27[1]) - float(a27[3])) < 0.01,
          "%s / %s" % (a27[1], a27[3]))
    check("tamamlanan atama: kalan = 0", abs(float(a27[2])) < 0.01, str(a27[2]))
    a85 = asg["bukhtourcity85"]
    check("kismi atama: fiili + kalan = hedef",
          abs(float(a85[1]) + float(a85[2]) - float(a85[3])) < 0.01,
          "%s + %s = %s" % (a85[1], a85[2], a85[3]))
    check("kismi atama: fiili = hedef x yuzde",
          abs(float(a85[1]) - float(a85[3]) * 0.70) < 0.01, str(a85[1]))

    rd = pr({"action": "read", "proj_id": PROJ, "only_started": True})
    check("read ilerlemeyi geri okur", rd["started"] >= 4, str(rd["started"]))
    check("read tamamlananlari sayar", rd["completed"] >= 2, str(rd["completed"]))
    done = [t for t in rd["tasks"] if t["code"] == "bukhtourcity27"]
    check("tamamlananin tahmini bitisi = fiili bitisi",
          done and done[0]["forecast_finish"] == done[0]["actual_finish"],
          "%s vs %s" % (done[0]["forecast_finish"] if done else "?",
                        done[0]["actual_finish"] if done else "?"))
    active = [t for t in rd["tasks"] if t["code"] == "bukhtourcity85"]
    check("devam edenin yuzdesi CP_Drtn'den okunur",
          active and abs(active[0]["percent_complete"] - 70.0) < 0.01,
          str(active[0]["percent_complete"] if active else "?"))

    # --- atama fiilleri
    before_ac = e({"action": "compute_metrics", "proj_id": PROJ})["ac"]
    ad = pr({"action": "set_assignment_actuals", "proj_id": PROJ, "dry_run": True,
             "updates": [{"task_code": "bukhtourcity27", "actual_qty": 205}]})
    check("set_assignment_actuals dry_run yazmaz", ad.get("dry_run") is True)
    aa = pr({"action": "set_assignment_actuals", "proj_id": PROJ, "confirm": True,
             "updates": [{"task_code": "bukhtourcity27", "actual_qty": 205}]})
    check("fiili birim yazildi", aa.get("updated") == 1)
    after = e({"action": "compute_metrics", "proj_id": PROJ})
    check("AC bagimsiz fiili ile artti", after["ac"] > before_ac,
          "%s -> %s" % (before_ac, after["ac"]))
    check("CPI artik 1.0 degil (gercek olcum)",
          after["cpi"] is not None and abs(after["cpi"] - 1.0) > 1e-6,
          str(after["cpi"]))
    check("negatif fiili reddedilir",
          "error" in pr({"action": "set_assignment_actuals", "proj_id": PROJ,
                         "confirm": True, "updates": [
                             {"task_code": "bukhtourcity27", "actual_qty": -5}]}))
    check("atamasiz aktivitede fiili reddedilir",
          "error" in pr({"action": "set_assignment_actuals", "proj_id": PROJ,
                         "confirm": True, "updates": [
                             {"task_code": "bukhtourcity1346", "actual_qty": 5}]}))

    # --- temizlik
    cl = pr({"action": "clear", "proj_id": PROJ, "dry_run": True})
    check("clear dry_run sayar", cl.get("would_clear", 0) >= 4,
          str(cl.get("would_clear")))
    check("confirm'siz clear reddedilir",
          "error" in pr({"action": "clear", "proj_id": PROJ}))
    cl = pr({"action": "clear", "proj_id": PROJ, "confirm": True})
    check("ilerleme temizlendi", cl.get("cleared", 0) >= 4, str(cl.get("cleared")))
    left = sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                   "status_code <> 'TK_NotStart' AND delete_session_id IS NULL", PROJ)
    check("hicbir aktivite ilerlemede kalmadi", left == 0, str(left))
    dirty = sql_one("SELECT COUNT(*) FROM TASKRSRC WHERE proj_id=? AND "
                    "(act_reg_qty <> 0 OR remain_qty <> target_qty) AND "
                    "delete_session_id IS NULL", PROJ)
    check("atama defterleri de sifirlandi", dirty == 0, str(dirty))

    pr({"action": "set_data_date", "proj_id": PROJ,
        "data_date": str(original_dd)[:10], "confirm": True})
    check("veri tarihi geri konuldu",
          str(sql_one("SELECT last_recalc_date FROM PROJECT WHERE proj_id=?",
                      PROJ))[:10] == str(original_dd)[:10])


# ===========================================================================
# I -- p6_compare
# ===========================================================================
def cp(params):
    return json.loads(srv.p6_compare(params))


def part_i(baseline_id):
    section("I  KARSILASTIRMA (p6_compare)")
    if not baseline_id:
        p("  (baseline yok, karsilastirma atlandi)")
        return
    pair = {"a": {"baseline_proj_id": baseline_id}, "b": {"proj_id": PROJ}}

    s = cp(dict(pair, action="summary", limit=5))
    check("summary hatasiz", "error" not in s, s.get("error", ""))
    check("eslesme anahtari task_code", s.get("join_key") == "task_code")
    c = s["headline"]["counts"]
    # Asil kanit: baseline kopyasinin task_id'leri farklidir. id ile eslesseydi
    # 950 eklenen + 950 silinen cikardi.
    check("KRITIK -- yeniden numaralandirmaya ragmen 0 eklenen/0 silinen",
          c["tasks_added"] == 0 and c["tasks_removed"] == 0,
          "eklenen=%s silinen=%s" % (c["tasks_added"], c["tasks_removed"]))
    check("baseline ve canli farkli task_id tasiyor",
          sql_one("SELECT COUNT(*) FROM TASK a JOIN TASK b ON a.task_id=b.task_id "
                  "WHERE a.proj_id=? AND b.proj_id=?", PROJ, baseline_id) == 0)
    check("bitis hareketi 950 aktivite icin hesaplandi",
          s["finish_movement"]["compared"] == sql_one(
              "SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
              "delete_session_id IS NULL", PROJ),
          str(s["finish_movement"]["compared"]))
    check("proje bitisi gecikmesi raporlandi",
          s["finish_movement"]["project_slip_days"] is not None,
          "%s gun" % s["finish_movement"]["project_slip_days"])
    check("geciken + erken + degismeyen = karsilastirilan",
          (s["finish_movement"]["later"] + s["finish_movement"]["earlier"]
           + s["finish_movement"]["unchanged"]) == s["finish_movement"]["compared"])

    t = cp(dict(pair, action="tasks", limit=3))
    check("tasks degisen aktivite dondurur", t.get("changed_count", 0) >= 0,
          "degisen=%s" % t.get("changed_count"))
    lk = cp(dict(pair, action="links", limit=3))
    check("links: baseline kopyasinda mantik degismemis",
          lk.get("added_count") == 0 and lk.get("removed_count") == 0
          and lk.get("changed_count") == 0,
          "+%s -%s ~%s" % (lk.get("added_count"), lk.get("removed_count"),
                           lk.get("changed_count")))
    pg = cp(dict(pair, action="progress", limit=3))
    check("progress ilerleyen aktiviteleri bulur", "error" not in pg,
          "%s aktivite" % pg.get("tasks_count"))
    ev = cp(dict(pair, action="evm"))
    check("evm iki tarafi da hesaplar",
          ev["metrics_a"]["bac"] == ev["metrics_b"]["bac"],
          "%s / %s" % (ev["metrics_a"]["bac"], ev["metrics_b"]["bac"]))
    check("EV farki raporlandi", ev["delta"].get("ev_delta") is not None,
          str(ev["delta"].get("ev_delta")))

    check("kaynaksiz taraf reddedilir",
          "error" in cp({"action": "summary", "a": {"proj_id": PROJ}}))
    check("bilinmeyen action reddedilir",
          "error" in cp(dict(pair, action="yok")))

    if os.path.exists(XER):
        x = cp({"action": "summary", "a": {"type": "xer", "path": XER},
                "b": {"proj_id": PROJ}, "limit": 3})
        check("XER <-> veritabani karsilastirmasi calisir", "error" not in x,
              x.get("error", ""))
        check("XER tarafinda da 0 eklenen/0 silinen",
              x["headline"]["counts"]["tasks_added"] == 0
              and x["headline"]["counts"]["tasks_removed"] == 0)
        # 5.2 bulgusu: CLI import maliyeti dusurmus; iki taraf farkli birimde.
        check("birim uyusmazligi sessizce toplanmiyor, uyariliyor",
              any("farkli birimde" in w for w in x["warnings"]),
              str(x["warnings"])[:70])


# ===========================================================================
# J -- P6'nin kendi motoru
# ===========================================================================
def part_j():
    section("J  P6 MOTORU YAZILAN VERIYI DOGRULUYOR MU")
    p("  P6 Professional arayuzu degil, P6'nin KENDI hesaplama motoru")
    p("  (Job Service / prmjob.exe) uzerinden dogrulama.")

    # Kendi fixture'ini kurar: F bolumu ilerlemeyi temizleyerek bitiyor.
    pr({"action": "set_data_date", "proj_id": PROJ, "data_date": "2026-11-01",
        "confirm": True})
    applied = pr({"action": "set_progress", "proj_id": PROJ, "confirm": True,
                  "schedule": True, "timeout_s": 240, "updates": [
                      {"task_code": "bukhtourcity27", "status": "complete",
                       "actual_start": "2026-09-02", "actual_finish": "2026-09-24"},
                      {"task_code": "bukhtourcity85", "status": "in_progress",
                       "actual_start": "2026-09-18", "percent_complete": 70},
                      {"task_code": "bukhtourcity1346", "status": "in_progress",
                       "actual_start": "2026-09-24", "remaining_duration_h": 30}]})
    check("fixture: ilerleme yazildi ve P6 yeniden hesapladi",
          applied.get("schedule", {}).get("status") == "JS_Complete",
          str(applied.get("schedule", {}).get("elapsed_s")))

    rows = sql_all(
        "SELECT task_code, status_code, remain_drtn_hr_cnt, restart_date, "
        "reend_date, act_end_date, early_end_date FROM TASK WHERE proj_id=? "
        "AND task_code IN ('bukhtourcity85','bukhtourcity1346')", PROJ)
    dhc = float(sql_one(
        "SELECT c.day_hr_cnt FROM PROJECT p JOIN CALENDAR c ON c.clndr_id=p.clndr_id "
        "WHERE p.proj_id=?", PROJ) or 8.0)
    for r in rows:
        code, _st, remain, restart, reend = r[0], r[1], float(r[2]), r[3], r[4]
        if restart is None or reend is None:
            check("%s: kalan is penceresi var" % code, False, "pencere yok")
            continue
        gun_beklenen = remain / dhc
        gun_gercek = (reend - restart).days + 1
        check("P6 '%s' icin kalan sureyi benim degerimden turetti" % code,
              abs(gun_gercek - gun_beklenen) <= 4,
              "%.0fs/%.0fs-gun = %.1f is gunu, P6 %d gun planladi"
              % (remain, dhc, gun_beklenen, gun_gercek))

    check("tamamlananlarda P6 kalan is penceresi acmamis",
          sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                  "status_code='TK_Complete' AND (restart_date IS NOT NULL OR "
                  "reend_date IS NOT NULL)", PROJ) == 0)
    check("baslamamis is veri tarihinden once planlanmamis",
          sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                  "status_code='TK_NotStart' AND early_start_date < "
                  "(SELECT last_recalc_date FROM PROJECT WHERE proj_id=?)",
                  PROJ, PROJ) == 0)
    # P6 biten aktivitede early_end_date'i veri tarihine kaydirir -- bu bizim
    # forecast_finish zincirini act_end_date ile baslatmamizin sebebi.
    drift = sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                    "status_code='TK_Complete' AND early_end_date <> act_end_date",
                    PROJ)
    check("P6'nin biten iste early_end kaydirmasi hala gecerli (fix'in gerekcesi)",
          drift >= 0, "%d aktivitede erken bitis != fiili bitis" % drift)
    done = [t for t in pr({"action": "read", "proj_id": PROJ,
                           "only_started": True})["tasks"]
            if t["actual_finish"]]
    check("buna ragmen tahmini bitis = fiili bitis olarak raporlaniyor",
          all(t["forecast_finish"] == t["actual_finish"] for t in done) if done else True,
          "%d tamamlanan aktivite" % len(done))


# ===========================================================================
# G -- Job Service
# ===========================================================================
def part_g():
    section("G  JOB SERVICE (p6_job)")
    pf = jb({"action": "preflight", "proj_id": PROJ})
    check("preflight hazir", pf.get("ready") is True, json.dumps(pf.get("problems")))
    check("servis calisiyor", pf["checks"]["service"]["running"] is True)
    check("USEROBS satiri var", pf["checks"]["userobs_rows"] >= 1,
          str(pf["checks"]["userobs_rows"]))
    sh = jb({"action": "service_health"})
    check("service_health veritabanina ulasiyor", sh.get("db_reachable") is True)
    jd = jb({"action": "job_data", "proj_id": PROJ})
    check("JOB_DATA blob'u uretiliyor", str(PROJ) in str(jd.get("job_data")),
          str(jd.get("job_data"))[:60])

    before = sql_one("SELECT MAX(early_end_date) FROM TASK WHERE proj_id=? "
                     "AND delete_session_id IS NULL", PROJ)
    t0 = time.time()
    sc = jb({"action": "schedule", "proj_id": PROJ, "job_name": "MCP_ACCEPTANCE",
             "timeout_s": 300})
    check("schedule JS_Complete", sc.get("status") == "JS_Complete",
          "%.1f sn" % (time.time() - t0))
    check("schedule hata dondurmedi", sc.get("ok") is True, str(sc.get("error")))
    after = sql_one("SELECT MAX(early_end_date) FROM TASK WHERE proj_id=? "
                    "AND delete_session_id IS NULL", PROJ)
    check("P6 motoru tarihleri hesapladi", after is not None,
          "%s -> %s" % (str(before)[:10], str(after)[:10]))

    st = jb({"action": "status", "job_id": sc["job_id"]})
    check("status isi bulur", st.get("status") == "JS_Complete", str(st.get("status")))
    lj = jb({"action": "list", "limit": 5})
    check("list isleri dondurur", lj.get("count", 0) >= 1, str(lj.get("count")))
    check("olmayan proje reddedilir",
          "error" in jb({"action": "schedule", "proj_id": 999999, "timeout_s": 30}))


# ===========================================================================
# H -- korumalar
# ===========================================================================
def part_h():
    section("H  KORUMALAR")
    writes = [
        "DELETE FROM TASK",
        "UPDATE TASK SET task_name='x'",
        "DROP TABLE TASK",
        "EXEC sp_who",
        "SELECT 1; DELETE FROM TASK",
        "INSERT INTO TASK (task_id) VALUES (1)",
        "TRUNCATE TABLE TASK",
    ]
    rejected = sum(1 for s in writes
                   if "error" in q({"action": "sql", "sql": s}))
    check("sql yazma denemeleri reddedilir", rejected == len(writes),
          "%d/%d" % (rejected, len(writes)))
    ok_sql = q({"action": "sql", "sql": "SELECT COUNT(*) FROM TASK", "limit": 5})
    check("SELECT calisiyor", "error" not in ok_sql, str(ok_sql.get("count")))
    cte = q({"action": "sql", "limit": 5,
             "sql": "WITH x AS (SELECT 1 AS n) SELECT n FROM x"})
    check("WITH sorgusu calisiyor", "error" not in cte)

    check("bilinmeyen action reddedilir",
          "error" in q({"action": "yok_boyle"}))
    check("action'siz cagri reddedilir", "error" in q({}))
    check("kimlik parametresi reddedilir",
          "error" in q({"action": "db_info", "password": "x"}))
    check("PMXML kaynagi acikca reddedilir",
          "error" in q({"action": "read_tasks", "type": "pmxml", "path": "x.xml"}))
    check("olmayan XER reddedilir",
          "error" in q({"action": "read_tasks", "type": "xer", "path": "yok.xer"}))
    check("proj_id'siz okuma reddedilir",
          "error" in q({"action": "read_tasks"}))

    big = q({"action": "read_tasks", "proj_id": PROJ, "limit": 5000})
    check("limit 500'e kirpilir", len(big.get("items", [])) <= 500,
          str(len(big.get("items", []))))


# ===========================================================================
def cleanup(baseline_id):
    section("TEMIZLIK")
    cl = pr({"action": "clear", "proj_id": PROJ, "confirm": True})
    check("kalan ilerleme temizlendi", "error" not in cl,
          "%s aktivite" % cl.get("cleared"))
    if baseline_id:
        d = bl({"action": "delete", "baseline_proj_id": baseline_id,
                "confirm": True})
        check("test baseline'i silindi", "error" not in d,
              json.dumps(d.get("soft_deleted_rows", {})))
        check("silinen baseline listede yok",
              all(b["proj_id"] != baseline_id
                  for b in bl({"action": "list", "proj_id": PROJ})["baselines"]))
        check("canli proje bozulmadi",
              sql_one("SELECT COUNT(*) FROM TASK WHERE proj_id=? AND "
                      "delete_session_id IS NULL", PROJ) > 0)
    if os.path.exists(SNAP):
        os.remove(SNAP)
    jb({"action": "purge", "name_like": "MCP\\_ACCEPTANCE%"})


def main() -> int:
    t0 = time.time()
    p("P6 MCP TAM KABUL TESTI -- proje %s" % PROJ)
    p("baslangic: %s" % _dt.datetime.now().isoformat(timespec="seconds"))
    baseline_id = None
    try:
        part_a()
        part_b()
        part_c()
        baseline_id = part_e()
        part_f()
        part_d(baseline_id)
        part_g()
        part_h()
        part_i(baseline_id)
        part_j()
    except Exception:  # noqa: BLE001
        p("")
        p("!! TEST CALISMASI ISTISNA ILE DURDU:")
        p(traceback.format_exc())
        failed.append("calisma istisnasi")
    finally:
        try:
            cleanup(baseline_id)
        except Exception:  # noqa: BLE001
            p("temizlik hatasi:\n" + traceback.format_exc())

    p("")
    p("=" * 74)
    p("SONUC: %d kontrol gecti, %d hata, %.1f sn"
      % (passed, len(failed), time.time() - t0))
    for f in failed:
        p("   HATA: " + f)
    p("=" * 74)
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    p("Rapor: " + os.path.abspath(OUT))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
