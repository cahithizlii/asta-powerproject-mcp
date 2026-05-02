"""Phase 7 acceptance — msproject_compare end-to-end lifecycle.

Builds two CAU-style XER snapshots (A: 0% progress, B: progress + 1
added task) and exercises all 5 dispatcher actions. Demonstrates the
CAU monthly hakediş use case: last month vs this month delta.

Run:
    python -m samples.build_compare_lifecycle
"""
import asyncio
import json
import os
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_compare


SNAPSHOT_A = """ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tCAU\t2026-01-01 08:00\t2026-04-30 17:00\t2026-01-01 08:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tCAU 6x9\t9.0\t54.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWRK\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA10\tFoundation\tTT_Task\t180.0\t2026-01-01 08:00\t2026-01-31 17:00\t\t\t0.0
%R\t1002\t1\t1\t1\tA20\tFrame\tTT_Task\t360.0\t2026-02-01 08:00\t2026-03-15 17:00\t\t\t0.0
%R\t1003\t1\t1\t1\tA30\tHandover\tTT_FinMile\t0.0\t2026-04-30 17:00\t2026-04-30 17:00\t\t\t0.0
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\t1\t1002\t1001\tPR_FS\t0.0
%R\t2\t1003\t1002\tPR_FS\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t180.0\t0.0\t180.0\t0.0
%R\t2\t1002\t101\t360.0\t0.0\t360.0\t0.0
%E
"""


SNAPSHOT_B = """ERMHDR\t18.8\t2026-02-01\tu\tApp\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tCAU\t2026-01-01 08:00\t2026-04-30 17:00\t2026-02-01 08:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tCAU 6x9\t9.0\t54.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWRK\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA10\tFoundation\tTT_Task\t180.0\t2026-01-01 08:00\t2026-01-31 17:00\t2026-01-01 08:00\t2026-01-31 17:00\t100.0
%R\t1002\t1\t1\t1\tA20\tFrame\tTT_Task\t360.0\t2026-02-01 08:00\t2026-03-15 17:00\t2026-02-01 08:00\t\t40.0
%R\t1003\t1\t1\t1\tA30\tHandover\tTT_FinMile\t0.0\t2026-04-30 17:00\t2026-04-30 17:00\t\t\t0.0
%R\t1004\t1\t1\t1\tA25\tWalls\tTT_Task\t180.0\t2026-03-16 08:00\t2026-04-15 17:00\t\t\t0.0
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\t1\t1002\t1001\tPR_FS\t0.0
%R\t2\t1004\t1002\tPR_FS\t0.0
%R\t3\t1003\t1004\tPR_FS\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t180.0\t180.0\t180.0\t180.0
%R\t2\t1002\t101\t360.0\t144.0\t360.0\t144.0
%R\t3\t1004\t101\t180.0\t0.0\t180.0\t0.0
%E
"""


def _write_xer(content: str, name: str) -> str:
    path = os.path.join(tempfile.gettempdir(), name)
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(content.encode("utf-16-le"))
    return path


def _call(action, **kw):
    raw = asyncio.run(msproject_compare({"action": action, **kw}))
    return json.loads(raw)


def _print_section(title: str, payload: dict, max_chars: int = 800):
    print(f"\n{'=' * 70}\n  {title}\n{'=' * 70}")
    s = json.dumps(payload, indent=2, ensure_ascii=False, default=str)
    if len(s) > max_chars:
        s = s[:max_chars] + "\n... [truncated]"
    print(s)


def main():
    a = _write_xer(SNAPSHOT_A, "p7_acceptance_a.xer")
    b = _write_xer(SNAPSHOT_B, "p7_acceptance_b.xer")
    print(f"Snapshot A (Jan 0% baseline) : {a}")
    print(f"Snapshot B (Feb progress)    : {b}")

    task_d = _call("task_delta", file_a=a, file_b=b)
    _print_section("task_delta", task_d)

    link_d = _call("link_delta", file_a=a, file_b=b)
    _print_section("link_delta", link_d)

    progress_d = _call("progress_delta", file_a=a, file_b=b)
    _print_section("progress_delta", progress_d)

    evm_d = _call("evm_delta", file_a=a, file_b=b)
    _print_section("evm_delta", evm_d)

    summary = _call("summary", file_a=a, file_b=b)
    _print_section("summary", summary)

    print(f"\n{'=' * 70}\n  ACCEPTANCE ASSERTIONS\n{'=' * 70}")
    # snapshot B has +1 task (Walls 1004)
    assert len(task_d["added"]) == 1, \
        f"Expected 1 added task, got {len(task_d['added'])}"
    assert task_d["added"][0]["id"] == 1004
    print(f"[PASS] task_delta: 1 task added (id=1004 Walls)")

    # Foundation + Frame progressed
    assert progress_d["summary"]["count_moved"] >= 2
    print(f"[PASS] progress_delta: "
          f"{progress_d['summary']['count_moved']} tasks moved")

    # EV grew (Foundation 100% + Frame 40%)
    assert evm_d["ev_b"] > (evm_d["ev_a"] or 0)
    print(f"[PASS] evm_delta: EV grew {evm_d['ev_a']} -> {evm_d['ev_b']}")

    # Summary captures changes
    assert summary["counts"]["tasks_added"] == 1
    assert "Walls" in str(task_d["added"][0]) or summary["counts"]["tasks_progressed"] >= 2
    print(f"[PASS] summary: {summary['headline']}")

    print("\n[PASS] All Phase 7 acceptance assertions passed.")

    for p in (a, b):
        try:
            os.remove(p)
        except OSError:
            pass


if __name__ == "__main__":
    main()
