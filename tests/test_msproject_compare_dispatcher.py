"""Phase 7 T128 — msproject_compare dispatcher integration tests.

Verifies the new tool routes 5 actions through XER + MSPDI adapters.
Synthetic XER fixtures with known deltas:
- snapshot A: 3 tasks, all 0% progress
- snapshot B: 4 tasks (one added), first 2 progressed
"""
import asyncio
import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_compare
from tests._xer_fixture_builders import write_synthetic_xer as _write_xer


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


def _call(action, **kw):
    raw = asyncio.run(msproject_compare({"action": action, **kw}))
    return json.loads(raw)


# Module-level fixture paths (created once via pytest fixture below)
import pytest


@pytest.fixture(scope="module")
def xer_pair():
    a = _write_xer(SNAPSHOT_A, "p7_compare_a.xer")
    b = _write_xer(SNAPSHOT_B, "p7_compare_b.xer")
    yield a, b
    for p in (a, b):
        try:
            os.remove(p)
        except OSError:
            pass


# === task_delta ===

def test_dispatcher_task_delta_detects_added(xer_pair):
    a, b = xer_pair
    r = _call("task_delta", file_a=a, file_b=b)
    assert r["status"] == "ok"
    assert len(r["added"]) == 1
    assert r["added"][0]["id"] == 1004


def test_dispatcher_task_delta_detects_progress_change(xer_pair):
    a, b = xer_pair
    r = _call("task_delta", file_a=a, file_b=b)
    # T1 (1001) and T2 (1002) progressed → in 'changed'
    changed_ids = {c["id"] for c in r["changed"]}
    assert 1001 in changed_ids
    assert 1002 in changed_ids


def test_dispatcher_task_delta_no_removals_when_b_is_superset(xer_pair):
    a, b = xer_pair
    r = _call("task_delta", file_a=a, file_b=b)
    assert r["removed"] == []


# === link_delta ===

def test_dispatcher_link_delta_detects_new_link(xer_pair):
    a, b = xer_pair
    r = _call("link_delta", file_a=a, file_b=b)
    assert r["status"] == "ok"
    # New T4 brings 1004 -> 1003 link (snapshot B) — not in A
    assert len(r["added"]) >= 1


def test_dispatcher_link_delta_unchanged_count_positive(xer_pair):
    a, b = xer_pair
    r = _call("link_delta", file_a=a, file_b=b)
    # 2 links carried over (1001->1002, 1002->1003)
    assert r["unchanged_count"] >= 2


# === progress_delta ===

def test_dispatcher_progress_delta_count_moved_two(xer_pair):
    a, b = xer_pair
    r = _call("progress_delta", file_a=a, file_b=b)
    assert r["status"] == "ok"
    # T1 0->100, T2 0->50 = 2 movers (T4 may also count if treated as 0->0)
    assert r["summary"]["count_moved"] >= 2


def test_dispatcher_progress_delta_status_dates_passed_through(xer_pair):
    a, b = xer_pair
    r = _call("progress_delta", file_a=a, file_b=b)
    # Phase 5d XER reader uses last_recalc_date for status_date
    assert r["status_date_a"] is not None
    assert r["status_date_b"] is not None


# === evm_delta ===

def test_dispatcher_evm_delta_returns_metrics(xer_pair):
    a, b = xer_pair
    r = _call("evm_delta", file_a=a, file_b=b)
    assert r["status"] == "ok"
    # EV should grow (snapshot B has 100% + 50% complete on first 2)
    assert r["ev_b"] is not None
    assert r["ev_b"] > (r["ev_a"] or 0)


# === summary ===

def test_dispatcher_summary_includes_headline_and_counts(xer_pair):
    a, b = xer_pair
    r = _call("summary", file_a=a, file_b=b)
    assert r["status"] == "ok"
    assert "headline" in r
    assert "counts" in r
    assert r["counts"]["tasks_added"] == 1


def test_dispatcher_summary_headline_mentions_added_or_progressed(xer_pair):
    a, b = xer_pair
    r = _call("summary", file_a=a, file_b=b)
    assert ("added" in r["headline"]
            or "progressed" in r["headline"])


# === error paths ===

def test_dispatcher_unknown_action_lists_valid():
    r = _call("not_a_real_action")
    assert r["status"] == "error"
    assert "task_delta" in r["error"]


def test_dispatcher_missing_file_a_returns_error():
    r = _call("task_delta", file_a="/totally/missing/file.xer",
              file_b="/also/missing.xer")
    assert r["status"] == "error"


# === Phase 8.1 monthly_report ===

def test_dispatcher_monthly_report_bundles_compare_and_evm(xer_pair):
    a, b = xer_pair
    r = _call("monthly_report", file_a=a, file_b=b)
    assert r["status"] == "ok"
    assert "compare_summary" in r
    assert "evm_a" in r
    assert "evm_b" in r
    # EVM summary fields
    for key in ("rag", "completion_pct", "spi", "cpi"):
        assert key in r["evm_a"]
        assert key in r["evm_b"]


def test_dispatcher_monthly_report_headline_includes_rag(xer_pair):
    a, b = xer_pair
    r = _call("monthly_report", file_a=a, file_b=b)
    assert "RAG" in r["headline"]


def test_dispatcher_monthly_report_no_excel_path_when_omitted(xer_pair):
    a, b = xer_pair
    r = _call("monthly_report", file_a=a, file_b=b)
    assert r["excel_path"] is None
    assert r["excel_export"] is None


def test_dispatcher_monthly_report_writes_excel_when_path_given(xer_pair, tmp_path):
    a, b = xer_pair
    xlsx = str(tmp_path / "hakedis.xlsx")
    r = _call("monthly_report", file_a=a, file_b=b, output_excel=xlsx)
    assert r["status"] == "ok"
    assert r["excel_path"] == xlsx
    assert r["excel_export"] is not None
    assert r["excel_export"]["status"] == "ok"
    assert os.path.exists(xlsx)


def test_dispatcher_monthly_report_listed_in_unknown_action_error():
    r = _call("definitely_not_an_action")
    assert "monthly_report" in r["error"]


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================


def test_dispatcher_task_delta_missing_file_b_returns_error(xer_pair):
    """One-side missing → error from file_b loader."""
    a, _ = xer_pair
    r = _call("task_delta", file_a=a, file_b="/totally/missing/x.xer")
    assert r["status"] == "error"
    assert "file_b" in r["error"].lower() or "load" in r["error"].lower()


def test_dispatcher_unknown_action_listed_explicitly():
    """Unknown action lists all valid actions in error message."""
    r = _call("not_real")
    assert r["status"] == "error"
    assert "task_delta" in r["error"]
    assert "summary" in r["error"]


def test_dispatcher_evm_delta_missing_file_a_returns_error(xer_pair):
    """evm_delta with bad file_a → error from EVM loader."""
    _, b = xer_pair
    r = _call("evm_delta", file_a="/missing/a.xer", file_b=b)
    assert r["status"] == "error"
    err = r["error"].lower()
    assert "file_a" in err or "not found" in err


def test_dispatcher_summary_missing_file_returns_error():
    """summary with both missing → error."""
    r = _call("summary", file_a="/x/a.xer", file_b="/x/b.xer")
    assert r["status"] == "error"
    err = r["error"].lower()
    assert "file_a" in err or "not found" in err


def test_dispatcher_progress_delta_missing_file_returns_error():
    """progress_delta with one missing file → error."""
    r = _call("progress_delta", file_a="/no/a.xer", file_b="/no/b.xer")
    assert r["status"] == "error"
    err = r["error"].lower()
    assert "file_a" in err or "not found" in err
