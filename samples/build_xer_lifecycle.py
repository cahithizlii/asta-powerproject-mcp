"""Phase 5d XER acceptance: parse synthetic CAU-style XER + invoke 6 read actions.

NO MS Project COM required (XER reader is pure Python). Self-contained -
builds fixture in temp dir, parses, asserts shape.

Target: <10s wall clock (pure Python, small fixture).
"""
import functools
import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
print = functools.partial(print, flush=True)

from msproject_mcp_core import (
    _msp_xer_read_tasks, _msp_xer_read_links,
    _msp_xer_read_resources, _msp_xer_read_assignments,
    _msp_xer_read_calendars, _msp_xer_read_progress,
)


SAMPLE_CAU_XER_CONTENT = """ERMHDR\t18.8\t2026-05-01\tcahit\tProject Management\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tCAU\t2024-07-08 08:00\t2028-06-20 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tCAU 6x9\t9.0\t54.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tConcrete Workers\tCOW\tRT_Labor\t10.0
%R\t102\tExtractors\tEXT\tRT_Labor\t5.0
%R\t103\tSteel\tSTL\tRT_Mat\t100.0
%R\t104\tCarpenters\tCAR\tRT_Labor\t8.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct\ttotal_float_hr_cnt\tcstr_type\tstatus_code
%R\t1001\t1\t1\t1\tA1010\tFoundation\tTT_Task\t180.0\t2024-07-08 08:00\t2024-07-29 17:00\t2024-07-08 08:00\t2024-07-29 17:00\t100.0\t0.0\tCS_ASAP\tTK_Complete
%R\t1002\t1\t1\t1\tA1020\tFrame\tTT_Task\t360.0\t2024-07-30 08:00\t2024-09-09 17:00\t2024-07-30 08:00\t\t75.0\t0.0\tCS_ASAP\tTK_Active
%R\t1003\t1\t1\t1\tA1030\tWalls\tTT_Task\t180.0\t2024-09-10 08:00\t2024-10-01 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1004\t1\t1\t1\tA1040\tRoof\tTT_Task\t180.0\t2024-10-02 08:00\t2024-10-23 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1005\t1\t1\t1\tA1050\tInterior\tTT_Task\t360.0\t2024-10-24 08:00\t2024-12-04 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1006\t1\t1\t1\tA1060\tHandover\tTT_FinMile\t0.0\t2024-12-15 17:00\t2024-12-15 17:00\t\t\t0.0\t81.0\tCS_MFO\tTK_NotStart
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\t1\t1002\t1001\tPR_FS\t0.0
%R\t2\t1003\t1002\tPR_FS\t0.0
%R\t3\t1004\t1003\tPR_FS\t0.0
%R\t4\t1005\t1004\tPR_FS\t0.0
%R\t5\t1006\t1005\tPR_FS\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t180.0\t180.0\t180.0\t180.0
%R\t2\t1001\t103\t1000.0\t1000.0\t1000.0\t1000.0
%R\t3\t1002\t101\t360.0\t270.0\t360.0\t270.0
%R\t4\t1002\t104\t180.0\t135.0\t180.0\t135.0
%R\t5\t1003\t101\t180.0\t0.0\t180.0\t0.0
%R\t6\t1004\t104\t180.0\t0.0\t180.0\t0.0
%R\t7\t1005\t102\t360.0\t0.0\t360.0\t0.0
%E
"""


def main():
    out_dir = tempfile.mkdtemp(prefix="xer_acceptance_")
    xer_path = os.path.join(out_dir, "sample_cau.xer")
    with open(xer_path, "wb") as f:
        f.write(b"\xff\xfe")  # UTF-16-LE BOM
        f.write(SAMPLE_CAU_XER_CONTENT.encode("utf-16-le"))
    print(f"[INFO] synthetic CAU XER written to {xer_path} "
          f"({os.path.getsize(xer_path)} bytes)")

    t0 = time.time()

    # 1. read_tasks
    r = _msp_xer_read_tasks(file_path=xer_path)
    assert r["status"] == "ok", r
    assert r["count"] == 6
    print(f"\n1. read_tasks: {r['count']} tasks at {time.time()-t0:.3f}s")
    for t in r["tasks"][:3]:
        print(f"   id={t['id']} code={t['code']} name={t['name']!r} "
              f"%={t['percent_complete']}")

    # 2. read_links
    r = _msp_xer_read_links(file_path=xer_path)
    assert r["status"] == "ok"
    assert r["count"] == 5
    print(f"\n2. read_links: {r['count']} FS chain links")

    # 3. read_resources
    r = _msp_xer_read_resources(file_path=xer_path)
    assert r["status"] == "ok"
    assert r["count"] == 4
    print(f"\n3. read_resources: {r['count']} CAU resources")
    for res in r["resources"]:
        print(f"   id={res['id']} code={res['code']} name={res['name']!r} "
              f"type={res['type']}")

    # 4. read_assignments
    r = _msp_xer_read_assignments(file_path=xer_path)
    assert r["status"] == "ok"
    assert r["count"] == 7
    print(f"\n4. read_assignments: {r['count']} task-resource links")

    # 5. read_calendars
    r = _msp_xer_read_calendars(file_path=xer_path)
    assert r["status"] == "ok"
    cal = r["calendars"][0]
    assert cal["day_hr_cnt"] == 9.0
    assert cal["week_hr_cnt"] == 54.0
    print(f"\n5. read_calendars: {cal['name']} "
          f"({cal['day_hr_cnt']}h/day, {cal['week_hr_cnt']}h/week)")

    # 6. read_progress
    r = _msp_xer_read_progress(file_path=xer_path)
    assert r["status"] == "ok"
    assert r["status_date"] == "2026-05-01"
    print(f"\n6. read_progress: status_date={r['status_date']}, "
          f"{len(r['tasks'])} tasks tracked")

    elapsed = time.time() - t0
    print(f"\n[OK] XER ACCEPTANCE: {elapsed:.3f}s total (target <10s)")
    assert elapsed < 10.0, f"Too slow: {elapsed}s"


if __name__ == "__main__":
    main()
