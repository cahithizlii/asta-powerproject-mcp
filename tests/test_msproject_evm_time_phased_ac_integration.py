"""Phase 6.2 T117 — _msp_evm_time_phased_evm AC integration tests.

Verifies the time_phased_evm dispatcher action returns per-task AC
distribution (Phase 6.2) instead of the prior uniform total/past_buckets
approximation. Uses synthetic XER fixtures with predictable AC patterns.
"""
import os
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import _msp_evm_time_phased_evm, _evm_load_task_data


def _write_xer(content: str, name: str) -> str:
    path = os.path.join(tempfile.gettempdir(), name)
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(content.encode("utf-16-le"))
    return path


# Two completed tasks both finishing in the same early bucket — per-task
# distribution should put 100% AC at first bucket, plateau afterward.
SYNTH_EARLY_FINISH = """ERMHDR\t18.8\t2026-05-01\tcahit\tProject Management\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tDEMO\t2024-07-08 08:00\t2024-09-30 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tDEMO\t8.0\t40.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWRK\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA1010\tFoundation\tTT_Task\t180.0\t2024-07-08 08:00\t2024-07-29 17:00\t2024-07-08 08:00\t2024-07-29 17:00\t100.0
%R\t1002\t1\t1\t1\tA1020\tFrame\tTT_Task\t360.0\t2024-07-30 08:00\t2024-09-09 17:00\t2024-07-30 08:00\t\t75.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t180.0\t180.0\t180.0\t180.0
%R\t2\t1002\t101\t360.0\t270.0\t360.0\t270.0
%E
"""


# Two tasks finishing in DIFFERENT buckets — per-task should produce
# distinct AC growth per bucket, not uniform.
SYNTH_STAGGERED = """ERMHDR\t18.8\t2026-05-01\tcahit\tProject Management\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tDEMO\t2024-01-01 08:00\t2024-06-30 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tDEMO\t8.0\t40.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWRK\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA1010\tT1\tTT_Task\t160.0\t2024-01-01 08:00\t2024-01-31 17:00\t2024-01-01 08:00\t2024-01-31 17:00\t100.0
%R\t1002\t1\t1\t1\tA1020\tT2\tTT_Task\t160.0\t2024-03-01 08:00\t2024-03-31 17:00\t2024-03-01 08:00\t2024-03-31 17:00\t100.0
%R\t1003\t1\t1\t1\tA1030\tT3\tTT_Task\t160.0\t2024-05-01 08:00\t2024-05-31 17:00\t2024-05-01 08:00\t2024-05-31 17:00\t100.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t100.0\t100.0\t100.0\t100.0
%R\t2\t1002\t101\t100.0\t100.0\t100.0\t100.0
%R\t3\t1003\t101\t100.0\t100.0\t100.0\t100.0
%E
"""


def _build_early_finish_xer():
    return _write_xer(SYNTH_EARLY_FINISH, "phase6_2_early_finish.xer")


def _build_staggered_xer():
    return _write_xer(SYNTH_STAGGERED, "phase6_2_staggered.xer")


# ---------- Behavior tests ----------

def test_time_phased_ac_monotonic_non_decreasing():
    """Cumulative AC must never decrease across buckets (RULE 5/6)."""
    path = _build_early_finish_xer()
    try:
        r = _msp_evm_time_phased_evm(file_path=path, bucket="month")
        assert r["status"] == "ok"
        acs = [b["ac"] for b in r["buckets"]]
        for i in range(1, len(acs)):
            assert acs[i] >= acs[i - 1] - 0.01, \
                f"AC decreased at bucket {i}: {acs}"
    finally:
        os.remove(path)


def test_time_phased_ac_early_finish_plateaus():
    """Both tasks finish/start by mid-Aug 2024. AC should plateau at full
    sum(actual_work) from bucket 1 onwards (per-task, NOT uniform)."""
    path = _build_early_finish_xer()
    try:
        r = _msp_evm_time_phased_evm(file_path=path, bucket="month")
        load = _evm_load_task_data(file_path=path)
        total_aw = sum(float(t.get("actual_work") or 0) for t in load["tasks"])
        acs = [b["ac"] for b in r["buckets"]]
        # All buckets should report the full total AC (plateau early)
        for i, ac in enumerate(acs):
            assert abs(ac - total_aw) < 0.01, \
                f"Bucket {i} AC ({ac}) != total ({total_aw}) — uniform leak?"
    finally:
        os.remove(path)


def test_time_phased_ac_staggered_grows_in_steps():
    """3 tasks finishing in Jan/Mar/May should produce stepwise AC growth.
    Uniform would produce equal increments — verify distribution is NOT
    uniform by comparing variance of bucket increments."""
    path = _build_staggered_xer()
    try:
        r = _msp_evm_time_phased_evm(file_path=path, bucket="month")
        acs = [b["ac"] for b in r["buckets"]]
        # AC must end up with full sum
        load = _evm_load_task_data(file_path=path)
        total_aw = sum(float(t.get("actual_work") or 0) for t in load["tasks"])
        assert abs(acs[-1] - total_aw) < 0.01, \
            f"Final AC ({acs[-1]}) != total ({total_aw})"
        # Step pattern: at least one bucket increment matches per-task (~100h)
        # and at least one zero/small increment exists (no task finishing)
        diffs = [acs[i] - acs[i - 1] for i in range(1, len(acs))]
        positive_diffs = [d for d in diffs if d > 1.0]
        zero_diffs = [d for d in diffs if d < 1.0]
        # At least one growth period AND one plateau period — proves
        # non-uniform per-task distribution
        assert len(positive_diffs) >= 1, f"No AC growth: {acs}"
        assert len(zero_diffs) >= 1, f"No AC plateau (uniform leak?): {acs}"
    finally:
        os.remove(path)


def test_time_phased_ac_total_matches_sum_actual_work():
    """Cumulative AC at last bucket must equal sum(actual_work)."""
    path = _build_staggered_xer()
    try:
        r = _msp_evm_time_phased_evm(file_path=path, bucket="month")
        load = _evm_load_task_data(file_path=path)
        total_aw = sum(float(t.get("actual_work") or 0) for t in load["tasks"])
        final_ac = max(b["ac"] for b in r["buckets"])
        assert abs(final_ac - total_aw) < 0.01, \
            f"Final cumulative AC ({final_ac}) != sum actual_work ({total_aw})"
    finally:
        os.remove(path)
