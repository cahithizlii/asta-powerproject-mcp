"""Phase 11.3 — End-to-end acceptance scenarios.

Multi-tool workflow tests. Each scenario exercises 3-5 dispatcher
actions and asserts on the composite result.

Scenarios:
    1. CAU monthly hakedis chain (compare -> monthly_report -> Excel)
    2. Baseline lifecycle (file MCP write_baseline -> reload -> assert)
    3. Currency validation pipeline (XER -> validate_currency_mode)
    4. Time-phased EVM full chain (compute -> time_phased -> ac_increment)
    5. DCMA + EVM combined health (assess_all + progress_data_quality)
    6. Calendar recurring + assignment effect (Phase 10.2; COM-required)
    7. Update_task baseline awareness (Phase 9.3 + 10.1 baseline_after)
    8. Compare with baseline write back
"""
import asyncio
import json
import os
import shutil
import sys
import tempfile

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    msproject_compare,
    msproject_evm,
    msproject_health,
    msproject_file,
    _msp_evm_validate_currency_mode,
    _msp_evm_compute_metrics,
    _msp_evm_time_phased_evm,
    _msp_file_update_task,
)
from mspdi_parser import MspdiProject
from tests._xer_fixture_builders import write_synthetic_xer

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SOURCE_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


def _call(dispatcher, action, **kw):
    raw = asyncio.run(dispatcher({"action": action, **kw}))
    return json.loads(raw)


def _copy_to_tmp(name: str) -> str:
    """Copy source MSPDI to a writable tmp path."""
    tmp = os.path.join(tempfile.gettempdir(), name)
    shutil.copy(SOURCE_XML, tmp)
    return tmp


# ---------- Synthetic XER content blocks (test-private) ----------
# Modeled on tests/test_msproject_compare_dispatcher.py and
# tests/test_msproject_evm_time_phased_ac_integration.py patterns.

# Snapshot pair for Scenario 1 (monthly hakedis): A is 0% baseline,
# B has progress on first 2 tasks plus a newly added task.
SNAPSHOT_A = """ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tDEMO\t2026-01-01 08:00\t2026-04-30 17:00\t2026-01-01 08:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tStandard\t8.0\t40.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWRK\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA10\tT1\tTT_Task\t160.0\t2026-01-01 08:00\t2026-01-31 17:00\t\t\t0.0
%R\t1002\t1\t1\t1\tA20\tT2\tTT_Task\t160.0\t2026-02-01 08:00\t2026-02-28 17:00\t\t\t0.0
%R\t1003\t1\t1\t1\tA30\tT3\tTT_Task\t160.0\t2026-03-01 08:00\t2026-03-31 17:00\t\t\t0.0
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\t1\t1002\t1001\tPR_FS\t0.0
%R\t2\t1003\t1002\tPR_FS\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t160.0\t0.0\t160.0\t0.0
%R\t2\t1002\t101\t160.0\t0.0\t160.0\t0.0
%R\t3\t1003\t101\t160.0\t0.0\t160.0\t0.0
%E
"""


SNAPSHOT_B = """ERMHDR\t18.8\t2026-02-01\tu\tApp\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tDEMO\t2026-01-01 08:00\t2026-04-30 17:00\t2026-02-01 08:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tStandard\t8.0\t40.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWRK\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA10\tT1\tTT_Task\t160.0\t2026-01-01 08:00\t2026-01-31 17:00\t2026-01-01 08:00\t2026-01-31 17:00\t100.0
%R\t1002\t1\t1\t1\tA20\tT2\tTT_Task\t160.0\t2026-02-01 08:00\t2026-02-28 17:00\t2026-02-01 08:00\t\t50.0
%R\t1003\t1\t1\t1\tA30\tT3\tTT_Task\t160.0\t2026-03-01 08:00\t2026-03-31 17:00\t\t\t0.0
%R\t1004\t1\t1\t1\tA40\tT4\tTT_Task\t80.0\t2026-04-01 08:00\t2026-04-15 17:00\t\t\t0.0
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\t1\t1002\t1001\tPR_FS\t0.0
%R\t2\t1003\t1002\tPR_FS\t0.0
%R\t3\t1004\t1003\tPR_FS\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t160.0\t160.0\t160.0\t160.0
%R\t2\t1002\t101\t160.0\t80.0\t160.0\t80.0
%R\t3\t1003\t101\t160.0\t0.0\t160.0\t0.0
%R\t4\t1004\t101\t80.0\t0.0\t80.0\t0.0
%E
"""


# RULE 3 hours-mode pattern (target_cost == target_qty in every TASKRSRC).
# Used by Scenario 3 (currency validation).
HOURS_PATTERN = """ERMHDR\t18.8\t2026-05-01\tu\tApp\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tDEMO\t2026-01-01 08:00\t2026-06-30 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tStandard\t8.0\t40.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWRK\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct
%R\t2001\t1\t1\t1\tH10\tHT1\tTT_Task\t160.0\t2026-01-01 08:00\t2026-01-31 17:00\t\t\t0.0
%R\t2002\t1\t1\t1\tH20\tHT2\tTT_Task\t160.0\t2026-02-01 08:00\t2026-02-28 17:00\t\t\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t2001\t101\t160.0\t0.0\t160.0\t0.0
%R\t2\t2002\t101\t160.0\t0.0\t160.0\t0.0
%E
"""


# Staggered XER for Scenarios 4 + 5 — three completed tasks finishing
# in distinct buckets (Jan, Mar, May) so time-phased EVM has data.
SYNTH_STAGGERED = """ERMHDR\t18.8\t2026-05-01\tu\tApp\tUSD
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


# ---------- Scenario 1: CAU monthly hakedis chain ----------

def test_scenario_1_cau_monthly_hakedis_chain(tmp_path):
    """Build two snapshots, run monthly_report with Excel export,
    assert task delta + EVM RAG values + Excel file existence."""
    a = write_synthetic_xer(SNAPSHOT_A, "p113_s1_a.xer")
    b = write_synthetic_xer(SNAPSHOT_B, "p113_s1_b.xer")
    xlsx_out = str(tmp_path / "monthly.xlsx")
    try:
        r = _call(msproject_compare, "monthly_report",
                  file_a=a, file_b=b, output_excel=xlsx_out)
        assert r["status"] == "ok"
        # tasks_added counted via compare_summary -> counts dict
        counts = r.get("compare_summary", {}).get("counts", {})
        assert counts.get("tasks_added", 0) >= 1
        # RAG values are valid (RULE 12 emits uppercase RED/AMBER/GREEN)
        valid_rag = {"RED", "AMBER", "GREEN"}
        assert r["evm_a"]["rag"] in valid_rag
        assert r["evm_b"]["rag"] in valid_rag
        # Excel file written
        assert os.path.exists(xlsx_out)
    finally:
        for p in (a, b):
            try:
                os.remove(p)
            except OSError:
                pass


# ---------- Scenario 2: Baseline lifecycle (MSPDI) ----------

def test_scenario_2_baseline_lifecycle_mspdi(tmp_path):
    """write_baseline via file MCP -> reload via MspdiProject -> assert
    written values appear in read_baselines output."""
    out = str(tmp_path / "p113_s2_with_baseline.xml")
    proj = MspdiProject(SOURCE_XML)
    first_uid = next(iter(proj._task_elems.keys()))

    payload = [{
        "task_uid": first_uid,
        "baseline_start": "2026-06-01T08:00:00",
        "baseline_finish": "2026-06-30T17:00:00",
        "baseline_duration_h": 240.0,
        "baseline_work_h": 160.0,
    }]
    r = _call(msproject_file, "write_baseline",
              file_path=SOURCE_XML, baseline_number=0,
              baseline_data=payload, output_path=out)
    assert r["status"] == "ok"
    assert r["tasks_written"] == 1
    assert os.path.exists(out)

    # Reload via MspdiProject and assert
    reloaded = MspdiProject(out)
    bls = reloaded.read_baselines(0)
    assert len(bls) == 1
    assert bls[0]["task_uid"] == first_uid
    assert bls[0]["baseline_start"] == "2026-06-01T08:00:00"
    assert bls[0]["baseline_finish"] == "2026-06-30T17:00:00"


# ---------- Scenario 3: Currency validation pipeline ----------

def test_scenario_3_currency_validation_pipeline():
    """RULE 3 hours-mode XER -> validate_currency_mode returns
    primary_mode in valid set + USD currency_code + cross_validation
    consensus_mode."""
    xer = write_synthetic_xer(HOURS_PATTERN, "p113_s3_hours.xer")
    try:
        r = _msp_evm_validate_currency_mode(file_path=xer)
        assert r["status"] == "ok"
        assert r["primary_mode"] in ("cost", "hours", "mixed", "uncertain")
        assert r["currency_code"] == "USD"
        cv = r["cross_validation"]
        assert "consensus_mode" in cv
        assert cv["confidence"] in ("high", "medium", "low")
        assert "source_counts" in cv
    finally:
        try:
            os.remove(xer)
        except OSError:
            pass


# ---------- Scenario 4: Time-phased EVM full chain ----------

def test_scenario_4_time_phased_evm_full_chain():
    """compute_metrics + time_phased_evm — assert AC increment invariant
    and PV/EV monotonic non-decreasing."""
    xer = write_synthetic_xer(SYNTH_STAGGERED, "p113_s4_staggered.xer")
    try:
        cm = _msp_evm_compute_metrics(file_path=xer)
        assert cm["status"] == "ok"
        tp = _msp_evm_time_phased_evm(file_path=xer, bucket="month")
        assert tp["status"] == "ok"
        buckets = tp["buckets"]
        assert len(buckets) > 0
        # Phase 9.2 invariant: sum(ac_increment) ~= max(ac) (cumulative).
        sum_inc = sum(b["ac_increment"] for b in buckets)
        max_ac = max(b["ac"] for b in buckets)
        assert abs(sum_inc - max_ac) < 0.5, \
            f"sum(ac_increment)={sum_inc} max(ac)={max_ac}"
        # PV cumulative monotonic
        pvs = [b["pv"] for b in buckets]
        for i in range(1, len(pvs)):
            assert pvs[i] >= pvs[i - 1] - 0.01, \
                f"PV decreased at bucket {i}: {pvs}"
        # EV cumulative monotonic
        evs = [b["ev"] for b in buckets]
        for i in range(1, len(evs)):
            assert evs[i] >= evs[i - 1] - 0.01, \
                f"EV decreased at bucket {i}: {evs}"
    finally:
        try:
            os.remove(xer)
        except OSError:
            pass


# ---------- Scenario 5: DCMA + EVM combined health ----------

def test_scenario_5_dcma_plus_evm_combined_health():
    """msproject_health assess_all + msproject_evm progress_data_quality
    on the same staggered XER snapshot."""
    xer = write_synthetic_xer(SYNTH_STAGGERED, "p113_s5_staggered.xer")
    try:
        health = _call(msproject_health, "assess_all", file_path=xer)
        assert health["status"] == "ok"
        # DCMA assess_all returns rules + summary
        assert "rules" in health
        assert "summary" in health

        pdq = _call(msproject_evm, "progress_data_quality", file_path=xer)
        assert pdq["status"] == "ok"
        assert "warnings" in pdq
        assert "completion_pct" in pdq
    finally:
        try:
            os.remove(xer)
        except OSError:
            pass


# ---------- Scenario 6: Calendar recurring + assignment effect ----------

def test_scenario_6_calendar_recurring_weekly_exception(clean_test_project):
    """Phase 10.2 — create calendar, add recurring weekly exception
    (mon-fri off pattern), assert exception count > 0.

    Skipped automatically if MS Project COM is not available
    (clean_test_project fixture handles the skip via msproject_app)."""
    from msproject_mcp_core import (
        _msp_calendar_create,
        _msp_calendar_add_exception,
        _find_calendar_by_name,
    )
    proj = clean_test_project
    cal_name = "P113-S6-RecurCal"
    create = _msp_calendar_create(name=cal_name, base_calendar="Standard")
    assert create["status"] == "ok"

    add = _msp_calendar_add_exception(
        calendar_name=cal_name,
        exception_name="Weekday Off",
        start="2026-01-01",
        finish="2026-12-31",
        recurrence="weekly",
        days_of_week=["mon", "tue", "wed", "thu", "fri"],
    )
    assert add["status"] == "ok"
    assert add["recurrence"] == "weekly"

    cal = _find_calendar_by_name(proj, cal_name)
    assert cal is not None
    assert cal.Exceptions.Count > 0


# ---------- Scenario 7: Update_task baseline awareness ----------

def test_scenario_7_update_task_baseline_awareness(tmp_path):
    """Phase 9.3 + 10.1 — _msp_file_update_task with both schedule and
    baseline fields. Assert schedule_updated, baseline_written=1, and
    Phase 10.1 baseline_after read-back matches input."""
    src = str(tmp_path / "p113_s7_update.xml")
    shutil.copy(SOURCE_XML, src)

    proj = MspdiProject(src)
    # Pick a non-summary task ID
    sample_id = None
    for tid, uid in proj._uid_by_id.items():
        if tid != 0:  # skip summary task
            sample_id = tid
            break
    assert sample_id is not None, "No non-summary task found in fixture"

    fields = {
        "duration": "5d",
        "baseline_start": "2026-07-01T08:00:00",
        "baseline_finish": "2026-07-15T17:00:00",
        "baseline_duration_h": 80.0,
        "baseline_work_h": 40.0,
    }
    r = _msp_file_update_task(file_path=src, task_id=sample_id,
                              fields=fields, baseline_number=0)
    assert r["status"] == "ok"
    assert r["schedule_updated"] is True
    assert r["baseline_written"] == 1
    # Phase 10.1 read-back
    after = r.get("baseline_after")
    assert after is not None
    assert after.get("baseline_start") == "2026-07-01T08:00:00"
    assert after.get("baseline_finish") == "2026-07-15T17:00:00"


# ---------- Scenario 8: Compare with baseline write back ----------

def test_scenario_8_compare_with_baseline_write_back(tmp_path):
    """Build snapshot A (MSPDI fixture), write a baseline to a copy via
    file MCP write_baseline, then run msproject_compare task_delta on
    (A, modified). Verify the call returns ok and the persisted
    baseline survives a fresh MspdiProject reload.

    The MSPDI baseline is stored in <Baseline> elements; the file MCP
    surfaces them via _msp_file_read_baselines. Phase 4 file MCP does
    not yet propagate MspdiProject baselines into _evm_load_task_data,
    so msproject_compare task_delta only reports `unchanged_count` here
    (not 'changed' on baseline_start/finish). The compare call must
    still complete successfully; the baseline write is verified via
    direct MspdiProject.read_baselines on the modified copy.
    """
    out = str(tmp_path / "p113_s8_modified.xml")

    proj = MspdiProject(SOURCE_XML)
    target_uid = None
    target_id = None
    for tid, uid in proj._uid_by_id.items():
        if tid != 0:
            target_uid = uid
            target_id = tid
            break
    assert target_uid is not None

    payload = [{
        "task_uid": target_uid,
        "baseline_start": "2026-08-01T08:00:00",
        "baseline_finish": "2026-08-31T17:00:00",
        "baseline_duration_h": 240.0,
        "baseline_work_h": 160.0,
    }]
    write = _call(msproject_file, "write_baseline",
                  file_path=SOURCE_XML, baseline_number=0,
                  baseline_data=payload, output_path=out)
    assert write["status"] == "ok"
    assert write["tasks_written"] == 1
    assert os.path.exists(out)

    # Run msproject_compare task_delta — must complete cleanly even when
    # the MSPDI baseline element is the only differentiator.
    cmp = _call(msproject_compare, "task_delta",
                file_a=SOURCE_XML, file_b=out)
    assert cmp["status"] == "ok"
    assert "added" in cmp and "removed" in cmp and "changed" in cmp
    # Same task set on both sides -> nothing added / removed
    assert cmp["added"] == []
    assert cmp["removed"] == []
    # unchanged_count reflects the task set carried across both sides
    assert cmp["unchanged_count"] >= 1

    # Independent verification: the modified file's baseline is
    # readable and matches the payload (lossless write_baseline path).
    reloaded = MspdiProject(out)
    bls = reloaded.read_baselines(0)
    assert len(bls) == 1
    assert bls[0]["task_id"] == target_id
    assert bls[0]["baseline_start"] == "2026-08-01T08:00:00"
    assert bls[0]["baseline_finish"] == "2026-08-31T17:00:00"
