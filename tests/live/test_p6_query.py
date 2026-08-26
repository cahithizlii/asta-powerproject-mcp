"""p6_query kabul testi -- veritabani ve XER kaynagi, korumalar dahil."""
from __future__ import annotations

import io
import json
import os
import sys

sys.path.insert(0, os.environ.get("P6_REPO", os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
import p6_mcp_core as srv  # noqa: E402

XER = os.environ.get("P6_TEST_XER", "")   # canli test icin XER dosyasi
PROJ = int(os.environ.get("P6_TEST_PROJ_ID", "368"))
OUT = os.environ.get("P6_TEST_OUT", "test_p6_query.txt")

lines: list[str] = []
fails: list[str] = []


def p(msg: str = "") -> None:
    print(msg)
    lines.append(msg)


def q(params: dict) -> dict:
    return json.loads(srv.p6_query(params))


def check(label: str, cond: bool, detail: str = "") -> None:
    p("   %-52s %s%s" % (label, "OK" if cond else "HATA",
                         ("  " + detail) if detail else ""))
    if not cond:
        fails.append(label)


def main() -> int:
    # --- 1) db kaynagi -----------------------------------------------------
    p("=== 1) db kaynagi ===")
    r = q({"action": "list_projects"})
    check("list_projects 1 proje donuyor", r.get("count") == 1, str(r.get("count")))
    r = q({"action": "list_eps"})
    check("list_eps EPS dugumu donuyor", r.get("count", 0) >= 1,
          "count=%s flag=%s" % (r.get("count"), r.get("project_flag")))

    tasks = q({"action": "read_tasks", "proj_id": PROJ, "limit": 500})
    check("read_tasks 950 aktivite", tasks.get("count") == 950, str(tasks.get("count")))
    # 500 aktivite 25 KB'yi asar; json_response listeyi kucultup GECERLI JSON
    # dondurmeli ve kac satir donduruldugunu bildirmeli.
    check("truncated bayragi", tasks.get("truncated") is True)
    check("count sorgu toplamini veriyor (950)", tasks.get("count") == 950,
          str(tasks.get("count")))
    check("kesilen liste uzunlugu bildiriliyor",
          tasks.get("list_length_before_truncate") == 500,
          str(tasks.get("list_length_before_truncate")))
    check("returned <= istenen limit",
          0 < len(tasks.get("items", [])) <= 500, str(len(tasks.get("items", []))))
    check("kucultme notu var", "truncate_note" in tasks)
    src_meta = tasks.get("source", {})
    check("day_hr_cnt takvimden", "CALENDAR" in str(src_meta.get("day_hr_cnt_source")),
          str(src_meta.get("day_hr_cnt_source")))
    nones = [t for t in tasks["items"] if t.get("total_float") is None]
    check("total_float None yok (R10)", not nones, "%d adet" % len(nones))
    fc = [t for t in tasks["items"] if t.get("forecast_finish")]
    check("forecast_finish tum donen satirlarda dolu (RULE 16.B)",
          len(fc) == len(tasks["items"]), "%d/%d" % (len(fc), len(tasks["items"])))

    links = q({"action": "read_links", "proj_id": PROJ, "limit": 500})
    check("read_links 1701 bag", links.get("count") == 1701, str(links.get("count")))

    for act, expect_min in (("read_wbs", 700), ("read_resources", 1),
                            ("read_calendars", 1), ("read_assignments", 400)):
        rr = q({"action": act, "proj_id": PROJ, "limit": 500})
        check("%s calisiyor" % act, rr.get("count", 0) >= expect_min,
              "count=%s" % rr.get("count"))

    rp = q({"action": "read_project", "proj_id": PROJ})
    check("read_project data dondu", bool(rp.get("data")),
          str(rp.get("data", {}).get("proj_short_name")))
    pr = q({"action": "read_progress", "proj_id": PROJ})
    check("read_progress calisiyor", "data" in pr or "items" in pr)

    so = q({"action": "schedule_options", "proj_id": PROJ})
    check("schedule_options anahtar okudu", len(so.get("options", {})) > 10,
          "%d anahtar, trustworthy=%s" % (len(so.get("options", {})),
                                          so.get("cpm_trustworthy")))

    fd = q({"action": "finish_drivers", "proj_id": PROJ})
    check("finish_drivers calisiyor", "result" in fd,
          str(list(fd.get("result", {}))[:4]))

    di = q({"action": "db_info"})
    check("db_info tablo sayilari", di.get("table_rows", {}).get("TASK") == 950)

    # --- 2) XER kaynagi ----------------------------------------------------
    p("\n=== 2) XER kaynagi ===")
    xt = q({"action": "read_tasks", "type": "xer", "path": XER, "limit": 5})
    check("XER read_tasks 950", xt.get("count") == 950, str(xt.get("count")))
    check("XER day_hr_cnt takvimden",
          "CALENDAR" in str(xt.get("source", {}).get("day_hr_cnt_source")),
          str(xt.get("source", {}).get("day_hr_cnt_source")))
    xl = q({"action": "read_links", "type": "xer", "path": XER, "limit": 5})
    check("XER read_links 1701", xl.get("count") == 1701, str(xl.get("count")))

    # db ve xer ayni sayilari vermeli
    check("db ve XER aktivite sayisi ayni",
          xt.get("count") == tasks.get("count"),
          "xer=%s db=%s" % (xt.get("count"), tasks.get("count")))

    # --- 3) sql korumalari -------------------------------------------------
    p("\n=== 3) sql korumalari ===")
    ok = q({"action": "sql", "sql": "SELECT TOP 3 proj_id, proj_short_name FROM PROJECT"})
    check("SELECT calisiyor", ok.get("count", 0) >= 1,
          "kolonlar=%s" % ok.get("columns"))
    check("SELECT kolon adlarini donuyor",
          ok.get("columns") == ["proj_id", "proj_short_name"],
          str(ok.get("columns")))
    for bad, label in (
            ("DELETE FROM TASK", "DELETE reddedildi"),
            ("UPDATE PROJECT SET proj_short_name='x'", "UPDATE reddedildi"),
            ("SELECT 1; DROP TABLE TASK", "coklu ifade reddedildi"),
            ("SELECT * FROM TASK WHERE 1=1 -- ok\n DROP TABLE TASK", "DDL anahtar reddedildi"),
            ("EXEC sp_who", "EXEC reddedildi")):
        r = q({"action": "sql", "sql": bad})
        check(label, r.get("status") == "error", str(r.get("error"))[:60])

    # --- 4) hata yollari ---------------------------------------------------
    p("\n=== 4) hata yollari ===")
    r = q({"action": "read_tasks"})
    check("proj_id'siz okuma reddediliyor", r.get("status") == "error",
          str(r.get("error"))[:60])
    r = q({"action": "read_tasks", "type": "pmxml", "path": "x.xml"})
    check("PMXML acikca reddediliyor", r.get("status") == "error"
          and "MPXJ" in str(r.get("error")))
    r = q({"action": "read_tasks", "type": "xer", "path": "C:/yok/olmayan.xer"})
    check("olmayan dosya reddediliyor", r.get("status") == "error")
    r = q({"action": "read_tasks", "proj_id": 999999})
    check("olmayan proje reddediliyor", r.get("status") == "error",
          str(r.get("error"))[:60])
    r = q({"action": "read_tasks", "proj_short_name": "bukhtourcity", "limit": 1})
    check("proj_short_name ile cozuluyor", r.get("count") == 950)
    r = q({"action": "list_projects", "password": "x"})
    check("kimlik parametresi reddediliyor", r.get("status") == "error"
          and "Kimlik" in str(r.get("error")))

    p("\n" + "=" * 62)
    if fails:
        p("BASARISIZ (%d): %s" % (len(fails), "; ".join(fails)))
    else:
        p("TUM p6_query TESTLERI GECTI")
    io.open(OUT, "w", encoding="utf-8").write("\n".join(lines))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
