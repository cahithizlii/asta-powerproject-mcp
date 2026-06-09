"""P0 #2 — forecast-finish driver / top-level WBS anomaly (RULE 16.C)."""
import pytest
from xer_drivers import forecast_drivers, _build_top_level_map


# ---------- pure: top-level mapping ----------

def test_top_level_map_deep_nesting():
    wbs = [
        {"id": 10, "parent_id": None, "name": "Root"},
        {"id": 11, "parent_id": 10, "name": "Construction"},
        {"id": 12, "parent_id": 10, "name": "Procurement"},
        {"id": 13, "parent_id": 12, "name": "Infrastructure"},
        {"id": 14, "parent_id": 13, "name": "Power"},
    ]
    top_of, names = _build_top_level_map(wbs)
    assert top_of[11] == 11          # direct child of root
    assert top_of[13] == 12          # 13 -> 12 (top-level Procurement)
    assert top_of[14] == 12          # deep node rolls up to Procurement
    assert names[12] == "Procurement"


# ---------- pure: anomaly with LOE driver (ALFB1 pattern) ----------

WBS_ALF = [
    {"id": 10, "parent_id": None, "name": "ALF Project"},
    {"id": 11, "parent_id": 10, "name": "Construction"},
    {"id": 12, "parent_id": 10, "name": "Procurement"},
    {"id": 13, "parent_id": 12, "name": "Infrastructure"},
]

TASKS_ALF = [
    {"id": 201, "code": "C-101", "name": "Foundation", "task_type": "TT_Task",
     "wbs_id": 11, "forecast_finish": "2026-05-01", "duration_h": 100,
     "status": "TK_Complete", "percent_complete": 100},
    {"id": 202, "code": "C-102", "name": "Superstructure", "task_type": "TT_Task",
     "wbs_id": 11, "forecast_finish": "2026-06-03", "duration_h": 200,
     "status": "TK_Active", "percent_complete": 60},
    {"id": 301, "code": "PR-HG-ELEC-FD-3450", "name": "Infra Power LOE",
     "task_type": "TT_LOE", "wbs_id": 13, "forecast_finish": "2027-01-11",
     "duration_h": 3600, "status": "TK_Active", "percent_complete": 80},
]


def test_forecast_drivers_detects_loe_anomaly():
    r = forecast_drivers(TASKS_ALF, WBS_ALF, anomaly_gap_days=30)
    assert r["project_forecast_finish"] == "2027-01-11"
    assert r["branch_count"] == 2
    assert r["anomaly"] is True
    assert r["gap_days"] > 200
    drv = r["driver"]
    assert drv["wbs_name"] == "Procurement"
    assert drv["is_loe"] is True
    assert drv["driving_task"]["code"] == "PR-HG-ELEC-FD-3450"
    assert "LOE" in drv["note"]


def test_forecast_drivers_latest_tasks_sorted():
    r = forecast_drivers(TASKS_ALF, WBS_ALF)
    lt = r["latest_tasks"]
    assert lt[0]["code"] == "PR-HG-ELEC-FD-3450"   # latest
    assert lt[-1]["forecast_finish"] <= lt[0]["forecast_finish"]


def test_forecast_drivers_gap_below_threshold_no_anomaly():
    tasks = [
        {"id": 1, "wbs_id": 11, "forecast_finish": "2026-06-01",
         "task_type": "TT_Task"},
        {"id": 2, "wbs_id": 12, "forecast_finish": "2026-06-10",
         "task_type": "TT_Task"},
    ]
    r = forecast_drivers(tasks, WBS_ALF, anomaly_gap_days=30)
    assert r["gap_days"] == 9
    assert r["anomaly"] is False
    assert r["driver"] is None


def test_forecast_drivers_single_branch_no_anomaly():
    tasks = [
        {"id": 1, "wbs_id": 11, "forecast_finish": "2026-06-01",
         "task_type": "TT_Task"},
        {"id": 2, "wbs_id": 11, "forecast_finish": "2026-06-10",
         "task_type": "TT_Task"},
    ]
    r = forecast_drivers(tasks, WBS_ALF)
    assert r["branch_count"] == 1
    assert r["anomaly"] is False


def test_forecast_drivers_empty_safe():
    r = forecast_drivers([], [])
    assert r["branch_count"] == 0
    assert r["anomaly"] is False
    assert r["project_forecast_finish"] is None


def test_forecast_drivers_skips_tasks_without_forecast():
    tasks = [{"id": 1, "wbs_id": 11, "forecast_finish": None,
              "task_type": "TT_Task"}]
    r = forecast_drivers(tasks, WBS_ALF)
    assert r["branch_count"] == 0


# ---------- action: through the XER reader ----------

def test_finish_drivers_action_cau_single_branch(sample_cau_xer):
    from msproject_mcp_core import _msp_xer_finish_drivers
    r = _msp_xer_finish_drivers(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert r["branch_count"] == 1          # all tasks under Construction
    assert r["anomaly"] is False
    assert r["project_forecast_finish"] == "2024-12-27"  # Handover reend


def test_finish_drivers_action_anomaly(tmp_path):
    from tests._xer_fixture_builders import write_synthetic_xer
    from msproject_mcp_core import _msp_xer_finish_drivers
    content = (
        "ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
        "%T\tPROJWBS\n%F\twbs_id\tparent_wbs_id\tproj_id\twbs_short_name\twbs_name\n"
        "%R\t10\t\t1\tROOT\tALF Project\n"
        "%R\t11\t10\t1\tCON\tConstruction\n"
        "%R\t12\t10\t1\tPRC\tProcurement\n"
        "%R\t13\t12\t1\tINF\tInfrastructure\n"
        "%T\tTASK\n%F\ttask_id\twbs_id\ttask_code\ttask_name\ttask_type"
        "\ttarget_end_date\treend_date\tphys_complete_pct\tstatus_code\n"
        "%R\t201\t11\tC-101\tFoundation\tTT_Task\t2026-05-01 17:00\t2026-05-01 17:00\t100\tTK_Complete\n"
        "%R\t202\t11\tC-102\tFrame\tTT_Task\t2026-06-03 17:00\t2026-06-03 17:00\t60\tTK_Active\n"
        "%R\t301\t13\tPR-HG-ELEC-FD-3450\tInfra Power\tTT_LOE\t2026-07-01 17:00\t2027-01-11 17:00\t80\tTK_Active\n"
        "%E\n"
    )
    path = write_synthetic_xer(content, "alf_drivers.xer")
    try:
        r = _msp_xer_finish_drivers(file_path=path, anomaly_gap_days=30)
        assert r["status"] == "ok"
        assert r["anomaly"] is True
        assert r["driver"]["is_loe"] is True
        assert r["driver"]["wbs_name"] == "Procurement"
        assert r["driver"]["driving_task"]["code"] == "PR-HG-ELEC-FD-3450"
        assert r["project_forecast_finish"] == "2027-01-11"
    finally:
        import os
        os.remove(path)
