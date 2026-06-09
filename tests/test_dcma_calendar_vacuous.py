"""P1 #5 (calendar-aware high_duration) + #6 (vacuous-link flags)."""
import pytest
from dcma_checks import (
    check_high_duration, check_leads, check_lags, check_fs_link_pct,
    assess_all,
)


# ---------- #5: calendar-aware Rule 9 ----------

def test_high_duration_default_8h_threshold_352():
    """Backward compat: default 8h/day -> 352h threshold."""
    tasks = [{"id": 1, "duration_h": 360.0}]   # > 352
    r = check_high_duration(tasks)
    assert r["threshold_hours"] == 352.0
    assert r["status"] == "fail"
    assert r["failed_count"] == 1


def test_high_duration_9h_calendar_threshold_396():
    """CAU 9h/day -> 396h threshold; 360h task no longer flagged."""
    tasks = [{"id": 1, "duration_h": 360.0}]   # < 396
    r = check_high_duration(tasks, day_hr_cnt=9.0)
    assert r["threshold_hours"] == 396.0
    assert r["status"] == "pass"
    assert r["failed_count"] == 0


def test_high_duration_9h_still_flags_truly_long():
    tasks = [{"id": 1, "duration_h": 400.0}]   # > 396
    r = check_high_duration(tasks, day_hr_cnt=9.0)
    assert r["status"] == "fail"


def test_assess_all_passes_day_hr_cnt_through():
    tasks = [{"id": i, "duration_h": 360.0, "predecessors": [1],
              "successors": [2]} for i in range(1, 5)]
    links = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0}]
    r8 = assess_all(tasks, links, [], day_hr_cnt=8.0)
    r9 = assess_all(tasks, links, [], day_hr_cnt=9.0)
    rule9_8 = next(x for x in r8["rules"] if x["id"] == 9)
    rule9_9 = next(x for x in r9["rules"] if x["id"] == 9)
    assert rule9_8["status"] == "fail"   # 360 > 352
    assert rule9_9["status"] == "pass"   # 360 < 396


# ---------- #6: vacuous-link flags ----------

def test_leads_empty_links_vacuous():
    r = check_leads([])
    assert r["status"] == "pass"
    assert r["vacuous"] is True
    assert "No links" in r["note"]


def test_lags_empty_links_vacuous():
    r = check_lags([])
    assert r["status"] == "pass"
    assert r["vacuous"] is True


def test_fs_link_empty_vacuous_not_real_100():
    r = check_fs_link_pct([])
    assert r["status"] == "pass"
    assert r["vacuous"] is True
    assert r["actual"] == 100.0          # still 100 but flagged vacuous
    assert "vacuous" in r["note"].lower()


def test_non_empty_links_have_no_vacuous_flag():
    links = [{"from_id": 1, "to_id": 2, "type": "FS", "lag_days": 0}]
    assert "vacuous" not in check_leads(links)
    assert "vacuous" not in check_fs_link_pct(links)
    assert "vacuous" not in check_lags(links)


# ---------- #5: XER calendar-aware via DCMA action ----------

def test_dcma_assess_uses_xer_calendar_day_hr(sample_cau_xer):
    """CAU XER (9h/day): Rule 9 threshold should be 396h, not 352h."""
    from msproject_mcp_core import _msp_dcma_assess_all
    r = _msp_dcma_assess_all(file_path=sample_cau_xer)
    rule9 = next(x for x in r["rules"] if x["id"] == 9)
    assert rule9["day_hr_cnt"] == 9.0
    assert rule9["threshold_hours"] == 396.0
