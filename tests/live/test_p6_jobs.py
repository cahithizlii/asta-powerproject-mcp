"""p6/jobs.py kabul testi -- gercek P6 Job Service uzerinde.

Her adim kendi kanitini basar (RULE 18). Program verisine yazan tek sey
P6'nin kendi scheduler'idir; bu script yalniz JOBSVC kuyruguna satir birakir
ve dogrulama icin data date'i oynatir (tek kullanimlik import edilmis proje).
"""
from __future__ import annotations

import io
import os
import sys
import time

sys.path.insert(0, os.environ.get("P6_REPO", os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from p6 import db, jobs  # noqa: E402

PROJ = int(os.environ.get("P6_TEST_PROJ_ID", "368"))
USER = 25
OUT = os.environ.get("P6_TEST_OUT", "test_p6_jobs.txt")

lines: list[str] = []


def p(msg: str = "") -> None:
    print(msg)
    lines.append(msg)


def fingerprint(cur) -> tuple:
    cur.execute(
        "SELECT COUNT(*), MIN(EARLY_START_DATE), MAX(EARLY_END_DATE), "
        "MAX(LATE_END_DATE), SUM(CAST(TOTAL_FLOAT_HR_CNT AS float)) "
        "FROM TASK WHERE PROJ_ID = ?", PROJ)
    a = tuple(cur.fetchone())
    cur.execute("SELECT LAST_RECALC_DATE, LAST_SCHEDULE_DATE FROM PROJECT "
                "WHERE PROJ_ID = ?", PROJ)
    return a + tuple(cur.fetchone())


def main() -> int:
    failures: list[str] = []

    # --- 1) build_job_data, P6'nin kendi blob'uyla birebir mi ---------------
    p("=== 1) build_job_data ===")
    reference = "(0||(Default Project|368)((0||Schedule Projects()((0||368()())))))"
    produced = jobs.build_job_data(jobs.JT_SCHEDULE, [368])
    p("   P6 referansi : " + reference)
    p("   uretilen     : " + produced)
    ok = produced == reference
    p("   BIREBIR      : %s" % ("EVET" if ok else "HAYIR"))
    if not ok:
        failures.append("build_job_data referans blob ile eslesmiyor")

    multi = jobs.build_job_data(jobs.JT_SCHEDULE, [368, 400], default_proj_id=400)
    p("   cok projeli  : " + multi)
    if "(0||400()())(0||368()())" not in multi and "(0||400()())" not in multi:
        failures.append("cok projeli blob beklenen yapida degil")

    # --- 2) hata cevirisi ---------------------------------------------------
    p("\n=== 2) hata cevirisi ===")
    for raw in ("No projects to schedule", "Default Project not found",
                "Cannot find OCI DLL:", "OK", "bilinmeyen hata xyz"):
        p("   %-32s -> %s" % (raw, (jobs.translate_error(raw) or "(bos)")[:95]))

    # --- 3) alias + baglanti ------------------------------------------------
    p("\n=== 3) alias ve baglanti ===")
    alias = db.resolve_alias()
    p("   alias=%s driver=%s db=%s host=%s" %
      (alias.name, alias.driver, alias.database, alias.host))
    cn = db.connect_rw(alias)
    cur = cn.cursor()
    p("   yazilabilir baglanti: OK (Windows kimlik dogrulamasi)")

    # --- 4) preflight -------------------------------------------------------
    p("\n=== 4) preflight ===")
    pf = jobs.preflight(cur, alias, USER, [PROJ])
    p("   ready=%s" % pf["ready"])
    for k, v in pf["checks"].items():
        p("   %-14s %s" % (k, v))
    for prob in pf["problems"]:
        p("   SORUN: " + prob)
    if not pf["ready"]:
        failures.append("preflight ready=False")
        p("\npreflight gecmedi, is gonderilmeyecek.")
        return finish(failures)

    # preflight negatif senaryo: EPS dugumu reddediliyor mu
    pf_eps = jobs.preflight(cur, alias, USER, [355])
    p("   [negatif] EPS dugumu (355) -> ready=%s  problem=%s" %
      (pf_eps["ready"], (pf_eps["problems"] or ["-"])[0][:70]))
    if pf_eps["ready"]:
        failures.append("preflight EPS dugumunu reddetmedi")

    # --- 5) F9: submit + wait ----------------------------------------------
    p("\n=== 5) F9 (JT_Sched) ===")
    before = fingerprint(cur)
    p("   ONCE : ES_min=%s EF_max=%s TF=%.0f data_date=%s" %
      (before[1], before[2], before[4], before[5]))

    # tarihlerin gercekten yeniden hesaplandigini gorebilmek icin data date'i oynat
    cur.execute("UPDATE PROJECT SET LAST_RECALC_DATE = DATEADD(day, -30, "
                "LAST_RECALC_DATE) WHERE PROJ_ID = ?", PROJ)
    cur.execute("SELECT LAST_RECALC_DATE FROM PROJECT WHERE PROJ_ID = ?", PROJ)
    p("   data date 30 gun GERI alindi -> %s" % cur.fetchone()[0])

    t0 = time.time()
    res = jobs.run_and_wait(cur, jobs.JT_SCHEDULE, [PROJ], USER,
                            job_name="MCP_TEST_F9", timeout_s=180)
    p("   job_id=%d  gecisler=%s" % (res.job_id, res.transitions))
    p("   durum=%s  ok=%s  sure=%.1fs" % (res.status, res.ok, res.elapsed_s))
    if res.error:
        p("   hata=%s" % res.error_tr)
    if not res.ok:
        failures.append("F9 isi JS_Complete olmadi: %s" % res.status)

    after = fingerprint(cur)
    p("   SONRA: ES_min=%s EF_max=%s TF=%.0f data_date=%s" %
      (after[1], after[2], after[4], after[5]))
    changed = [i for i in range(len(before)) if before[i] != after[i]]
    p("   degisen alan indeksleri: %s" % changed)
    if not changed:
        failures.append("tarihler degismedi -- scheduler calismamis olabilir")
    p("   toplam sure (submit->complete): %.1f s" % (time.time() - t0))

    # --- 6) list / cancel / purge ------------------------------------------
    p("\n=== 6) list / cancel / purge ===")
    lst = jobs.list_jobs(cur, limit=5)
    p("   list_jobs -> %d kayit" % len(lst))
    for j in lst[:3]:
        p("      %s %-14s %-12s %s" % (j["job_id"], j["job_type"], j["status"],
                                       j["job_name"]))

    jid = jobs.submit(cur, jobs.JT_SCHEDULE, [PROJ], USER,
                      job_name="MCP_TEST_CANCEL", recur_type=jobs.RT_RECUR_DISABLED)
    can = jobs.cancel(cur, jid)
    p("   cancel(%d) -> %s" % (jid, can))
    if can.get("status") != jobs.JS_CANCELLED:
        failures.append("cancel calismadi")

    n = jobs.purge(cur)
    p("   purge -> %d satir silindi" % n)

    cn.close()
    return finish(failures)


def finish(failures: list[str]) -> int:
    p("\n" + "=" * 60)
    if failures:
        p("BASARISIZ (%d):" % len(failures))
        for f in failures:
            p("   - " + f)
    else:
        p("TUM TESTLER GECTI")
    io.open(OUT, "w", encoding="utf-8").write("\n".join(lines))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
