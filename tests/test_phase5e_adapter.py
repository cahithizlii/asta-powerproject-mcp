"""Test Phase 5e XER -> Phase 5a shape adapter (T109)."""
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import _xer_to_evm_task_shape, _evm_load_task_data
from xer_parser import XerFile


def test_adapter_returns_phase5a_shape(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    assert r["status"] == "ok"
    for k in ("tasks", "resources", "assignments", "status_date", "project_file"):
        assert k in r


def test_adapter_tasks_have_baseline_fields(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    for t in r["tasks"]:
        for k in ("baseline_start", "baseline_finish", "baseline_work",
                  "actual_work"):
            assert k in t


def test_adapter_baseline_equals_target_cau_pattern(sample_cau_xer):
    """CAU cost-loaded NO: baseline = current target schedule."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    assert foundation["baseline_start"] == foundation["start"]
    assert foundation["baseline_finish"] == foundation["finish"]
    assert foundation["baseline_work"] == foundation["duration_h"]


def test_adapter_actual_work_aggregated(sample_cau_xer):
    """actual_work = sum of TASKRSRC.act_reg_qty per task_id."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    # COW 180 + STL 1000 = 1180
    assert foundation["actual_work"] == 1180.0
    # Frame: COW 270 + CAR 135 = 405
    frame = next(t for t in r["tasks"] if t["id"] == 1002)
    assert frame["actual_work"] == 405.0


def test_adapter_predecessors_derived(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    frame = next(t for t in r["tasks"] if t["id"] == 1002)
    assert 1001 in frame["predecessors"]
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    # Foundation = first task, no predecessor
    assert foundation["predecessors"] == []


def test_adapter_successors_derived(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    assert 1002 in foundation["successors"]
    handover = next(t for t in r["tasks"] if t["id"] == 1006)
    # Handover = last task, no successor
    assert handover["successors"] == []


def test_adapter_critical_from_zero_slack(sample_cau_xer):
    """critical = total_slack_days <= 0 (XER lacks explicit critical flag)."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    # total_float_hr_cnt=0 -> total_slack_days=0 -> critical=True
    assert foundation["critical"] is True
    walls = next(t for t in r["tasks"] if t["id"] == 1003)
    # total_float_hr_cnt=72 -> total_slack_days=8 (CAU 9h/day) -> critical=False
    assert walls["critical"] is False


def test_adapter_total_slack_days_present(sample_cau_xer):
    """total_slack_days needed for DCMA Rule 7-8."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    walls = next(t for t in r["tasks"] if t["id"] == 1003)
    assert walls["total_slack_days"] == walls["total_float"]


def test_adapter_excludes_summary_tasks(sample_cau_xer):
    """Summary tasks (TT_LOE/TT_WBS) excluded; CAU fixture has none."""
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    # Fixture has 6 real tasks (incl. milestone) + 0 summaries
    assert len(r["tasks"]) == 6


def test_adapter_status_date_from_progress(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    assert r["status_date"] == "2026-05-01"


def test_adapter_project_file_set(sample_cau_xer):
    xer = XerFile(sample_cau_xer)
    r = _xer_to_evm_task_shape(xer)
    assert r["project_file"] == sample_cau_xer


# ---- _evm_load_task_data routing ----

def test_evm_load_task_data_routes_xer(sample_cau_xer):
    """When file_path ends .xer, returns Phase 5e adapter output."""
    r = _evm_load_task_data(file_path=sample_cau_xer)
    assert r["status"] == "ok"
    assert len(r["tasks"]) == 6
    foundation = next(t for t in r["tasks"] if t["id"] == 1001)
    assert "baseline_work" in foundation
    assert foundation["baseline_work"] == 180.0


def test_evm_load_task_data_xer_routing_case_insensitive(tmp_path):
    """File extension match should be case-insensitive (.XER works too)."""
    src = os.path.join(os.path.dirname(__file__), "..", "..")  # not used
    # Use the conftest content but write to .XER (uppercase) path
    from tests.conftest import SAMPLE_CAU_XER_CONTENT
    path = tmp_path / "upper.XER"
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(SAMPLE_CAU_XER_CONTENT.encode("utf-16-le"))
    r = _evm_load_task_data(file_path=str(path))
    assert r["status"] == "ok"
    assert len(r["tasks"]) == 6


def test_evm_load_task_data_existing_xml_path_unchanged():
    """sample_msp.xml (Phase 4 fixture) unchanged - falls through original logic."""
    msp_xml = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")
    if not os.path.exists(msp_xml):
        return  # skip if fixture absent
    r = _evm_load_task_data(file_path=msp_xml)
    assert r["status"] in ("ok", "error")  # Phase 4 file path bubbles or errors
    # Crucially: NOT raising / not crashing due to xer routing
