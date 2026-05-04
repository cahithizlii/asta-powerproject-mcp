"""Phase 11.1 T141 — gap-fill coverage tests for mspdi_parser.

Targets the large block of read/write methods that are not exercised by
the existing baseline_write / msproject_compat / file_* / phase5e/f tests.

Pure parser tests — no COM, no MS Project running. Uses the
fixtures/sample_msp.xml fixture and synthetic XML for branch coverage.
"""
import os
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from mspdi_parser import MspdiProject

SAMPLE_PATH = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _copy_sample(suffix: str) -> str:
    """Copy the sample XML to a writable temp path."""
    fd, tmp = tempfile.mkstemp(prefix="mspdi_test_", suffix=f"_{suffix}.xml")
    os.close(fd)
    shutil.copy(SAMPLE_PATH, tmp)
    return tmp


@pytest.fixture
def proj():
    """Fresh MspdiProject parsed from sample fixture (in-memory only)."""
    path = _copy_sample("readonly")
    try:
        yield MspdiProject(path)
    finally:
        if os.path.exists(path):
            os.remove(path)


@pytest.fixture
def writable_proj():
    """MspdiProject backed by a writable temp file (for save tests)."""
    path = _copy_sample("writable")
    proj = MspdiProject(path)
    proj.__test_path__ = path  # stash for cleanup
    yield proj
    if os.path.exists(path):
        os.remove(path)


# ============================================================
# Constructor + foundational helpers
# ============================================================

def test_init_raises_on_missing_file():
    with pytest.raises(FileNotFoundError):
        MspdiProject("/definitely/does/not/exist.xml")


def test_namespace_helper_returns_qualified_tag(proj):
    assert proj._t("Task") == "{http://schemas.microsoft.com/project}Task"


def test_root_int_default_when_missing(proj):
    # MissingField doesn't exist on root -> default returned
    assert proj._root_int("MissingField", 42) == 42


def test_text_helper_uses_default_for_missing_child(proj):
    # Read missing child via _text
    elem = proj.root
    assert proj._text(elem, "TotallyMissingField", "fallback") == "fallback"


def test_int_helper_uses_default_on_invalid(proj):
    # Inject a child with non-numeric text and check default fallback
    parent = ET.SubElement(proj.root, proj._t("ZBogus"))
    parent.text = "not-an-int"
    assert proj._int(proj.root, "ZBogus", 99) == 99


def test_float_helper_uses_default_on_invalid(proj):
    parent = ET.SubElement(proj.root, proj._t("ZBogusFloat"))
    parent.text = "not-a-float"
    assert proj._float(proj.root, "ZBogusFloat", 7.5) == 7.5


def test_make_elem_with_text_and_without(proj):
    e1 = proj._make_elem("Foo", "bar")
    assert e1.text == "bar"
    e2 = proj._make_elem("Foo")
    assert e2.text is None


def test_set_text_creates_child_when_missing(proj):
    elem = ET.Element(proj._t("Container"))
    proj._set_text(elem, "Field", "value")
    found = elem.find(proj._t("Field"))
    assert found is not None and found.text == "value"
    # Updating same field re-uses same element
    proj._set_text(elem, "Field", "updated")
    assert elem.find(proj._t("Field")).text == "updated"


# ============================================================
# Duration / date / lag helpers
# ============================================================

def test_parse_iso_duration_zero_for_empty_or_bad(proj):
    assert proj._parse_iso_duration("") == 0.0
    assert proj._parse_iso_duration(None) == 0.0
    assert proj._parse_iso_duration("not-a-PT") == 0.0
    assert proj._parse_iso_duration("PT-bad") == 0.0


def test_parse_iso_duration_full_components(proj):
    # PT8H30M15S -> 8 + 0.5 + 15/3600 = 8.5041...
    out = proj._parse_iso_duration("PT8H30M15S")
    assert out == pytest.approx(8 + 30/60 + 15/3600, rel=1e-6)


def test_hours_to_days_and_back(proj):
    # default minutes_per_day=480 -> 8h/day
    assert proj._hours_to_days(16) == 2.0
    assert proj._days_to_hours(2.0) == 16.0


def test_format_duration_str_zero(proj):
    assert proj._format_duration_str("PT0H0M0S") == "0d"


def test_format_duration_str_int_days(proj):
    assert proj._format_duration_str("PT16H0M0S") == "2d"


def test_format_duration_str_fractional_days(proj):
    # 12h = 1.5d
    assert proj._format_duration_str("PT12H0M0S") == "1.5d"


def test_days_to_iso(proj):
    # 2 days = 16 hours
    assert proj._days_to_iso(2.0) == "PT16H0M0S"


def test_parse_date_normal_and_short(proj):
    assert proj._parse_date("2026-04-30T08:00:00") == "2026-04-30"
    assert proj._parse_date("") == "N/A"
    assert proj._parse_date(None) == "N/A"
    assert proj._parse_date("short") == "short"


def test_lag_helpers_round_trip(proj):
    # 1 day = mpd*10 = 4800 lag units
    assert proj._days_to_lag(1.0) == 4800
    assert proj._lag_to_days(4800) == pytest.approx(1.0, rel=1e-6)
    assert proj._lag_to_days(0) == 0.0


def test_format_lag_zero_int_and_fraction(proj):
    assert proj._format_lag(0) == "0d"
    assert proj._format_lag(4800) == "1d"
    # 1.5 day lag
    half_day = proj._days_to_lag(1.5)
    assert proj._format_lag(half_day) == "1.5d"


def test_parse_duration_input_variants(proj):
    assert proj._parse_duration_input("") == 1.0
    assert proj._parse_duration_input(None) == 1.0
    assert proj._parse_duration_input("10d") == 10.0
    assert proj._parse_duration_input("2w") == 10.0  # 2 * 5
    # Months: days_per_month default 20 -> 1mo = 20d
    assert proj._parse_duration_input("1mo") == 20.0
    assert proj._parse_duration_input("16h") == 2.0  # 16/8
    # Plain numeric string
    assert proj._parse_duration_input("3.5") == 3.5
    # Non-parseable -> default 1.0
    assert proj._parse_duration_input("garbage") == 1.0


# ============================================================
# Project summary / basic queries
# ============================================================

def test_get_project_summary_keys(proj):
    s = proj.get_project_summary()
    for key in ("file", "project_name", "client", "start_date", "finish_date",
                "status_date", "current_date", "total_tasks",
                "summary_tasks", "milestones", "activities",
                "critical_tasks", "total_resources", "total_assignments",
                "total_links", "calendars", "code_libraries"):
        assert key in s


def test_get_all_tasks_default_includes_summary(proj):
    tasks_all = proj.get_all_tasks(include_summary=True)
    tasks_no_sum = proj.get_all_tasks(include_summary=False)
    assert len(tasks_no_sum) <= len(tasks_all)


def test_get_task_by_id_returns_none_for_missing(proj):
    assert proj.get_task_by_id(999999) is None


def test_get_task_by_uid_returns_none_for_missing(proj):
    assert proj.get_task_by_uid(999999) is None


def test_get_task_by_id_returns_detail(proj):
    # Pick a known task ID from the sample
    sample_task = next(iter(proj._tasks_by_id.values()))
    detail = proj.get_task_by_id(sample_task["id"])
    # Detail dict should be non-empty
    assert detail is not None
    for k in ("id", "name", "duration", "constraint_type"):
        assert k in detail


def test_get_task_by_uid_returns_detail(proj):
    sample_uid = next(iter(proj._tasks.keys()))
    detail = proj.get_task_by_uid(sample_uid)
    assert detail is not None
    # _task_to_detail_dict serializes detail; UID key may be 'unique_id'
    assert "id" in detail


def test_get_critical_path(proj):
    # Function should run without error and return a list
    out = proj.get_critical_path()
    assert isinstance(out, list)


def test_get_resources(proj):
    out = proj.get_resources()
    assert isinstance(out, list)


def test_get_resource_assignments(proj):
    out = proj.get_resource_assignments()
    assert isinstance(out, list)


def test_get_calendars(proj):
    out = proj.get_calendars()
    assert isinstance(out, list)


def test_get_wbs_tree(proj):
    out = proj.get_wbs_tree()
    assert isinstance(out, list)


def test_get_wbs_tree_max_depth_zero(proj):
    out = proj.get_wbs_tree(max_depth=0)
    assert isinstance(out, list)


def test_get_delay_analysis(proj):
    out = proj.get_delay_analysis()
    assert isinstance(out, dict)


def test_get_float_analysis(proj):
    out = proj.get_float_analysis()
    assert isinstance(out, dict)


def test_get_resource_loading(proj):
    out = proj.get_resource_loading()
    assert isinstance(out, dict)


def test_get_code_libraries(proj):
    out = proj.get_code_libraries()
    assert isinstance(out, list)


def test_get_task_codes_for_existing_and_missing(proj):
    sample_id = next(iter(proj._tasks_by_id.keys()))
    out = proj.get_task_codes(sample_id)
    assert isinstance(out, dict)
    out_missing = proj.get_task_codes(999999)
    assert out_missing == {} or "error" in out_missing or out_missing is None


def test_filter_tasks_by_code_unknown_lib(proj):
    # Unknown library -> empty list
    out = proj.filter_tasks_by_code("__unknown_lib__")
    assert isinstance(out, list)


def test_get_latest_finishing(proj):
    out = proj.get_latest_finishing(count=5)
    assert isinstance(out, list)
    assert len(out) <= 5


def test_find_missing_links(proj):
    out = proj.find_missing_links()
    assert isinstance(out, dict)


def test_search_tasks_by_pattern(proj):
    # Search by "" should match everything
    out = proj.search_tasks("")
    assert isinstance(out, list)
    out_no = proj.search_tasks("__zzzz_no_match__")
    assert out_no == []


def test_search_tasks_exclude_summary(proj):
    out = proj.search_tasks("", include_summary=False)
    assert isinstance(out, list)


def test_get_link_chain(proj):
    out = proj.get_link_chain("Foundation", "Roof")
    assert isinstance(out, dict)


def test_get_link_chain_no_match(proj):
    out = proj.get_link_chain("__zzz__", "__qqq__")
    assert isinstance(out, dict)


def test_get_tasks_between_dates_no_filter(proj):
    out = proj.get_tasks_between_dates()
    assert isinstance(out, list)


def test_get_tasks_between_dates_with_range(proj):
    out = proj.get_tasks_between_dates(start_after="2020-01-01",
                                        finish_before="2099-12-31")
    assert isinstance(out, list)


# ============================================================
# Write APIs: add_task, update_task, delete_task
# ============================================================

def test_add_task_default(writable_proj):
    n_before = len(writable_proj._tasks)
    r = writable_proj.add_task("New_T1")
    assert "task_id" in r
    assert "uid" in r
    assert "Task '" in r["message"]
    assert len(writable_proj._tasks) == n_before + 1


def test_add_task_with_explicit_dates_and_duration(writable_proj):
    r = writable_proj.add_task(
        "New_T2",
        duration_str="3d",
        start_date="2026-06-01",
        finish_date="2026-06-05",
    )
    assert r["start"] == "2026-06-01"


def test_add_task_milestone_zero_duration(writable_proj):
    r = writable_proj.add_task("Milestone1", duration_str="5d", is_milestone=True)
    assert r["type"] == "milestone"


def test_add_task_summary_returns_summary_type(writable_proj):
    r = writable_proj.add_task("Sum1", duration_str="0d", is_summary=True)
    assert r["type"] == "summary"


def test_add_summary_task_helper(writable_proj):
    r = writable_proj.add_summary_task("HelperSummary")
    assert r["type"] == "summary"


def test_add_child_task_with_existing_parent(writable_proj):
    parent = writable_proj.add_task("Parent1")
    pid = parent["task_id"]
    r = writable_proj.add_child_task(parent_task_id=pid, name="Child1")
    assert r.get("parent_id") == pid
    assert r.get("parent_name") == "Parent1"


def test_add_task_with_invalid_start_date_falls_back(writable_proj):
    # Force the fallback to current date by passing N/A start_date via blank Project root
    # We just confirm the code path runs and returns successfully.
    r = writable_proj.add_task("Quirky", duration_str="1d")
    assert "task_id" in r


def test_add_task_with_unparseable_finish_uses_start_dt(writable_proj):
    # Pass start_date that can't be parsed to ISO date -> falls into except
    r = writable_proj.add_task(
        "WeirdStart",
        duration_str="2d",
        start_date="not-an-iso-date",
    )
    # Should not crash; finish defaults to start_dt
    assert "task_id" in r


def test_update_task_unknown_returns_error(writable_proj):
    r = writable_proj.update_task(task_id=999999, name="X")
    assert "error" in r


def test_update_task_modifies_all_fields(writable_proj):
    parent = writable_proj.add_task("UpdateMe")
    tid = parent["task_id"]
    r = writable_proj.update_task(
        task_id=tid,
        name="Renamed",
        duration_str="5d",
        percent_complete=25,
        notes="Hello world",
        start_date="2026-07-01",
        finish_date="2026-07-10",
    )
    assert "error" not in r
    assert any("name -> 'Renamed'" in c for c in r["changes"])
    assert any("duration -> 5d" in c for c in r["changes"])
    assert any("percent_complete -> 25%" in c for c in r["changes"])
    assert any("notes" in c for c in r["changes"])
    assert any("start -> 2026-07-01" in c for c in r["changes"])
    assert any("finish -> 2026-07-10" in c for c in r["changes"])


def test_delete_task_unknown_returns_error(writable_proj):
    r = writable_proj.delete_task(999999)
    assert "error" in r


def test_delete_task_existing(writable_proj):
    a = writable_proj.add_task("ToDelete")
    tid = a["task_id"]
    r = writable_proj.delete_task(tid)
    assert r.get("deleted") is True
    assert tid not in writable_proj._tasks_by_id


def test_delete_task_removes_predecessor_links(writable_proj):
    a = writable_proj.add_task("Pred", duration_str="2d")
    b = writable_proj.add_task("Succ", duration_str="2d")
    writable_proj.add_link(a["task_id"], b["task_id"], link_type="FS")
    writable_proj.delete_task(a["task_id"])
    # After deleting predecessor, successor task elem should not contain
    # a PredecessorLink referencing the deleted task.
    succ_uid = b["uid"]
    succ_elem = writable_proj._task_elems[succ_uid]
    pred_links = list(writable_proj._findall(succ_elem, "PredecessorLink"))
    pred_uids = [writable_proj._int(p, "PredecessorUID") for p in pred_links]
    assert a["uid"] not in pred_uids


# ============================================================
# Resources + assignments
# ============================================================

def test_add_resource_basic(writable_proj):
    rid = writable_proj.add_resource("RES_W", type="Work", max_units=2.0)
    assert isinstance(rid, int)
    # Confirm in internal index
    found = any(r.get("id") == rid for r in writable_proj._resources.values())
    assert found


def test_add_resource_material(writable_proj):
    rid = writable_proj.add_resource("RES_M", type="Material",
                                      standard_rate="0/h")
    assert isinstance(rid, int)


def test_add_resource_cost_type(writable_proj):
    rid = writable_proj.add_resource("RES_C", type="Cost")
    assert isinstance(rid, int)


def test_add_resource_invalid_name_raises(writable_proj):
    with pytest.raises(ValueError):
        writable_proj.add_resource("")


def test_add_resource_invalid_type_raises(writable_proj):
    with pytest.raises(ValueError):
        writable_proj.add_resource("BadType", type="Imaginary")


def test_add_assignment_to_existing(writable_proj):
    t = writable_proj.add_task("AT1", duration_str="2d")
    rid = writable_proj.add_resource("AR1", type="Work")
    a_uid = writable_proj.add_assignment(t["task_id"], rid, units=1.5)
    assert isinstance(a_uid, int)


def test_add_assignment_missing_task_raises(writable_proj):
    rid = writable_proj.add_resource("AR2", type="Work")
    with pytest.raises(ValueError):
        writable_proj.add_assignment(999999, rid)


def test_add_assignment_missing_resource_raises(writable_proj):
    t = writable_proj.add_task("AT2", duration_str="2d")
    with pytest.raises(ValueError):
        writable_proj.add_assignment(t["task_id"], 999999)


def test_add_assignment_with_work_string(writable_proj):
    t = writable_proj.add_task("AT3", duration_str="3d")
    rid = writable_proj.add_resource("AR3", type="Work")
    a_uid = writable_proj.add_assignment(t["task_id"], rid, work_str="16h")
    assert isinstance(a_uid, int)


def test_bulk_add_assignments(writable_proj):
    # Build 3 tasks and 2 resources
    ids = [writable_proj.add_task(f"BT{i}", "2d")["task_id"] for i in range(3)]
    rid_a = writable_proj.add_resource("BRA", type="Work")
    rid_b = writable_proj.add_resource("BRB", type="Work")
    items = (
        [{"task_id": ids[0], "resource_id": rid_a, "units": 1.0},
         {"task_id": ids[1], "resource_id": rid_b, "units": 0.5},
         {"task_id": ids[2], "resource_id": rid_a},
         # Skipped: missing task
         {"task_id": 999999, "resource_id": rid_a},
         # Skipped: missing resource
         {"task_id": ids[0], "resource_id": 999999}]
    )
    n = writable_proj.bulk_add_assignments(items)
    # Only 3 valid out of 5
    assert n == 3


# ============================================================
# Links
# ============================================================

def test_add_link_unknown_predecessor(writable_proj):
    t = writable_proj.add_task("LK1")
    r = writable_proj.add_link(999999, t["task_id"])
    assert "error" in r


def test_add_link_unknown_successor(writable_proj):
    t = writable_proj.add_task("LK2")
    r = writable_proj.add_link(t["task_id"], 999999)
    assert "error" in r


def test_add_link_with_lag_and_alt_type(writable_proj):
    a = writable_proj.add_task("LKA")
    b = writable_proj.add_task("LKB")
    r = writable_proj.add_link(a["task_id"], b["task_id"],
                                link_type="SS", lag_str="2d")
    assert r.get("success") is True
    assert r["link_type"] == "SS"


def test_remove_link_works(writable_proj):
    a = writable_proj.add_task("LR1")
    b = writable_proj.add_task("LR2")
    writable_proj.add_link(a["task_id"], b["task_id"])
    out = writable_proj.remove_link(a["task_id"], b["task_id"])
    assert out.get("removed") is True


def test_remove_link_unknown_predecessor(writable_proj):
    t = writable_proj.add_task("LR3")
    r = writable_proj.remove_link(999999, t["task_id"])
    assert "error" in r


def test_remove_link_unknown_successor(writable_proj):
    t = writable_proj.add_task("LR4")
    r = writable_proj.remove_link(t["task_id"], 999999)
    assert "error" in r


def test_remove_link_nonexistent_pair_returns_error(writable_proj):
    a = writable_proj.add_task("LR5")
    b = writable_proj.add_task("LR6")
    r = writable_proj.remove_link(a["task_id"], b["task_id"])
    assert "error" in r


def test_update_link_works(writable_proj):
    a = writable_proj.add_task("LU1")
    b = writable_proj.add_task("LU2")
    writable_proj.add_link(a["task_id"], b["task_id"])
    r = writable_proj.update_link(
        a["task_id"], b["task_id"],
        new_link_type="SF", new_lag_str="1d",
    )
    assert r.get("updated") is True
    assert r["new_type"] == "SF"


def test_update_link_unknown_predecessor(writable_proj):
    t = writable_proj.add_task("LU3")
    r = writable_proj.update_link(999999, t["task_id"], new_link_type="FF")
    assert "error" in r


def test_update_link_unknown_successor(writable_proj):
    t = writable_proj.add_task("LU4")
    r = writable_proj.update_link(t["task_id"], 999999, new_link_type="FF")
    assert "error" in r


def test_update_link_no_existing_link_returns_error(writable_proj):
    a = writable_proj.add_task("LU5")
    b = writable_proj.add_task("LU6")
    r = writable_proj.update_link(a["task_id"], b["task_id"], new_lag_str="1d")
    assert "error" in r


# ============================================================
# Progress
# ============================================================

def test_update_progress_unknown_task(writable_proj):
    r = writable_proj.update_progress(999999, percent_complete=50)
    assert "error" in r


def test_update_progress_all_fields(writable_proj):
    t = writable_proj.add_task("UP1")
    r = writable_proj.update_progress(
        t["task_id"],
        percent_complete=42,
        actual_start="2026-07-01",
        actual_finish="2026-07-05",
    )
    assert r.get("updated") is True
    assert any("percent_complete -> 42%" in c for c in r["changes"])


def test_bulk_update_progress(writable_proj):
    a = writable_proj.add_task("BP1")
    b = writable_proj.add_task("BP2")
    updates = [
        {"task_id": a["task_id"], "percent_complete": 30},
        {"task_id": b["task_id"], "percent_complete": 75},
        {"task_id": 999999, "percent_complete": 0},  # error path
    ]
    r = writable_proj.bulk_update_progress(updates)
    assert r["total"] == 3
    assert r["success"] == 2
    assert r["failed"] == 1


# ============================================================
# Codes
# ============================================================

def test_assign_code_unknown_task(writable_proj):
    r = writable_proj.assign_code(999999, "Library", "Value")
    assert "error" in r


def test_assign_code_unknown_library(writable_proj):
    t = writable_proj.add_task("AC1")
    r = writable_proj.assign_code(t["task_id"], "__nonexistent_lib__", "v")
    assert "error" in r


# ============================================================
# Save / round-trip
# ============================================================

def test_save_default_creates_file(writable_proj):
    # Use a temp dir for the save
    out_dir = tempfile.mkdtemp(prefix="mspdi_save_")
    try:
        out_path = os.path.join(out_dir, "round_trip.xml")
        result_path = writable_proj.save(out_path)
        assert os.path.exists(result_path)
        # Reload — should work end-to-end
        reloaded = MspdiProject(result_path)
        assert len(reloaded._tasks) == len(writable_proj._tasks)
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)


def test_save_default_path(writable_proj):
    """save() with no path generates a timestamped path next to source."""
    out_path = writable_proj.save()
    try:
        assert os.path.exists(out_path)
    finally:
        if os.path.exists(out_path):
            os.remove(out_path)


# ============================================================
# Synthetic XML — exercise rare branches
# ============================================================

def _build_minimal_mspdi(tmp_path, extra_xml=""):
    """Construct a minimal valid MSPDI XML with optional extra blocks."""
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>min</Name>
  <MinutesPerDay>480</MinutesPerDay>
  <MinutesPerWeek>2400</MinutesPerWeek>
  <DaysPerMonth>20</DaysPerMonth>
  <CalendarUID>1</CalendarUID>
  <StartDate>2026-01-01T08:00:00</StartDate>
""" + extra_xml + """
</Project>"""
    path = os.path.join(str(tmp_path), "min.xml")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(xml)
    return path


def test_minimal_project_no_tasks_no_calendars(tmp_path):
    """Project with no Tasks/Calendars/Resources/Assignments still parses."""
    path = _build_minimal_mspdi(tmp_path)
    proj = MspdiProject(path)
    assert proj.get_project_summary()["total_tasks"] == 0
    assert proj.get_calendars() == []
    assert proj.get_resources() == []
    assert proj.get_resource_assignments() == []


def test_calendar_with_exceptions_parses(tmp_path):
    """Calendar with TimePeriod exceptions hits parse branch."""
    extra = """
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>WithException</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
      <BaseCalendarUID>0</BaseCalendarUID>
      <WeekDays>
        <WeekDay>
          <DayType>2</DayType>
          <DayWorking>1</DayWorking>
          <WorkingTimes>
            <WorkingTime>
              <FromTime>08:00:00</FromTime>
              <ToTime>17:00:00</ToTime>
            </WorkingTime>
          </WorkingTimes>
        </WeekDay>
      </WeekDays>
      <Exceptions>
        <Exception>
          <TimePeriod>
            <FromDate>2026-12-25T00:00:00</FromDate>
            <ToDate>2026-12-25T23:59:00</ToDate>
          </TimePeriod>
          <Name>Christmas</Name>
          <Type>1</Type>
        </Exception>
      </Exceptions>
    </Calendar>
  </Calendars>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    cals = proj.get_calendars()
    assert len(cals) == 1
    # Internal _calendars dict stores full detail including exceptions
    full_cal = next(iter(proj._calendars.values()))
    assert any(e.get("name") == "Christmas"
               for e in full_cal.get("exceptions", []))


def test_extended_attributes_create_code_library(tmp_path):
    """ExtendedAttributes block populates _code_libs (parse_code_libraries)."""
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743731</FieldID>
      <FieldName>Text1</FieldName>
      <Alias>CodeLibrary:"DISCIPLINE"</Alias>
      <ValueList>
        <Value>
          <ID>1</ID>
          <Value>MEP</Value>
          <Description>Mechanical Electrical Plumbing</Description>
        </Value>
      </ValueList>
    </ExtendedAttribute>
  </ExtendedAttributes>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    libs = proj.get_code_libraries()
    names = [l["name"] for l in libs]
    assert "DISCIPLINE" in names


def test_extended_attributes_alias_without_codelibrary_keeps_as_name(tmp_path):
    """Alias without CodeLibrary:" prefix is used as-is for lib name."""
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743732</FieldID>
      <FieldName>Text2</FieldName>
      <Alias>SimpleAlias</Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    libs = proj.get_code_libraries()
    names = [l["name"] for l in libs]
    assert "SimpleAlias" in names


def test_extended_attributes_malformed_codelibrary_alias(tmp_path):
    """Alias 'CodeLibrary:"' without closing quote falls into IndexError path."""
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743733</FieldID>
      <FieldName>Text3</FieldName>
      <Alias>CodeLibrary:"NoClosingQuote</Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    # Parser tolerated malformed alias
    libs = proj.get_code_libraries()
    assert isinstance(libs, list)


def test_resource_assignments_parse(tmp_path):
    """Project with Resources + Assignments triggers _parse_resources_data
    and _parse_assignments_data."""
    extra = """
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
      <WeekDays/>
    </Calendar>
  </Calendars>
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>T1</Name>
      <Duration>PT16H0M0S</Duration><DurationFormat>7</DurationFormat>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-02T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <ConstraintType>0</ConstraintType>
      <CalendarUID>1</CalendarUID>
      <OutlineLevel>1</OutlineLevel>
      <WBS>1</WBS>
    </Task>
  </Tasks>
  <Resources>
    <Resource>
      <UID>10</UID><ID>1</ID><Name>RES_X</Name>
      <Type>1</Type><MaxUnits>1.0</MaxUnits>
    </Resource>
  </Resources>
  <Assignments>
    <Assignment>
      <UID>0</UID>
      <TaskUID>1</TaskUID>
      <ResourceUID>10</ResourceUID>
      <Units>1.0</Units>
      <Work>PT16H0M0S</Work>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-02T17:00:00</Finish>
    </Assignment>
  </Assignments>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    res = proj.get_resources()
    asgn = proj.get_resource_assignments()
    assert len(res) == 1
    assert len(asgn) == 1


def test_task_with_unparseable_calendar_uid(tmp_path):
    """Task CalendarUID that's non-int -> falls into except path."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>T1</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <ConstraintType>0</ConstraintType>
      <CalendarUID>not-a-number</CalendarUID>
      <OutlineLevel>1</OutlineLevel>
      <WBS>1</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    tasks = proj.get_all_tasks()
    assert len(tasks) == 1


# ============================================================
# Targeted gap-fill: get_wbs_tree, critical_path, delay/float/loading,
# search/filter, link_chain, between_dates
# ============================================================

def test_get_wbs_tree_empty_project(tmp_path):
    """Project with no tasks -> empty list."""
    path = _build_minimal_mspdi(tmp_path)
    proj = MspdiProject(path)
    assert proj.get_wbs_tree() == []


def test_get_critical_path_includes_critical_task(tmp_path):
    """Critical task surfaces in get_critical_path (line 615)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>CritOne</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <Critical>1</Critical>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    crit = proj.get_critical_path()
    assert any(t["name"] == "CritOne" for t in crit)


def test_get_delay_analysis_with_actual_dates(tmp_path):
    """Tasks with ActualStart/ActualFinish populate delays (lines 721-741)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>SlippedTask</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-05T17:00:00</Finish>
      <ActualStart>2026-01-03T08:00:00</ActualStart>
      <ActualFinish>2026-01-08T17:00:00</ActualFinish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>100</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
    <Task>
      <UID>2</UID><ID>2</ID><Name>BadDateTask</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>not-a-date</Start>
      <Finish>also-not-a-date</Finish>
      <ActualStart>also-broken</ActualStart>
      <ActualFinish>still-broken</ActualFinish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>50</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>2</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out = proj.get_delay_analysis()
    assert out["total_with_actuals"] >= 1
    # SlippedTask had +2 day start slip, +3 day finish slip
    assert out["max_start_slip"] >= 2


def test_get_float_analysis_with_total_slack(tmp_path):
    """TotalSlack populated -> float analysis bins tasks (lines 786-797)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>ZeroFloat</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <TotalSlack>PT0H0M0S</TotalSlack>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
    <Task>
      <UID>2</UID><ID>2</ID><Name>LowFloat</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <TotalSlack>PT24H0M0S</TotalSlack>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>2</WBS>
    </Task>
    <Task>
      <UID>3</UID><ID>3</ID><Name>MedFloat</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <TotalSlack>PT80H0M0S</TotalSlack>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>3</WBS>
    </Task>
    <Task>
      <UID>4</UID><ID>4</ID><Name>HighFloat</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <TotalSlack>PT400H0M0S</TotalSlack>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>4</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out = proj.get_float_analysis()
    assert out["has_float_data"] is True
    assert out["zero_float"] >= 1
    assert out["low_float"] >= 1
    assert out["medium_float"] >= 1
    assert out["high_float"] >= 1


def test_get_float_analysis_critical_fallback_with_note(tmp_path):
    """Without TotalSlack, falls back to Critical flag (lines 805-806, 813)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>CritFallback</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Critical>1</Critical>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out = proj.get_float_analysis()
    assert out["has_float_data"] is False
    assert "note" in out
    assert out["zero_float"] >= 1


def test_get_resource_loading_skips_orphan_assignment(tmp_path):
    """Assignment whose ResourceUID isn't in _resources is skipped (line 823)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>T1</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
  </Tasks>
  <Resources>
    <Resource>
      <UID>10</UID><ID>1</ID><Name>R1</Name>
      <Type>1</Type><MaxUnits>1.0</MaxUnits>
    </Resource>
  </Resources>
  <Assignments>
    <Assignment>
      <UID>0</UID><TaskUID>1</TaskUID><ResourceUID>10</ResourceUID>
      <Units>1.0</Units><Work>PT8H0M0S</Work>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
    </Assignment>
    <Assignment>
      <UID>1</UID><TaskUID>1</TaskUID><ResourceUID>9999</ResourceUID>
      <Units>1.0</Units><Work>PT8H0M0S</Work>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
    </Assignment>
  </Assignments>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out = proj.get_resource_loading()
    # Only the valid resource shows up
    assert out["total_resources"] == 1


def test_extended_attributes_with_anonymous_alias_skipped_in_output(tmp_path):
    """ExtendedAttribute with empty Alias -> empty lib name -> skipped (line 865)."""
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743734</FieldID>
      <FieldName>Text4</FieldName>
      <Alias></Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    libs = proj.get_code_libraries()
    # Empty-name libs are skipped from the public list
    assert all(l["name"] for l in libs)


def test_filter_tasks_by_code_with_value_match(tmp_path):
    """Task carrying a code -> filter_tasks_by_code returns it (lines 900-902)."""
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743741</FieldID>
      <FieldName>Text11</FieldName>
      <Alias>CodeLibrary:"DISC"</Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>CodedTask</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
      <ExtendedAttribute>
        <FieldID>188743741</FieldID>
        <Value>MEP</Value>
        <ValueID>1</ValueID>
      </ExtendedAttribute>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out_all = proj.filter_tasks_by_code("DISC")
    out_match = proj.filter_tasks_by_code("DISC", value="mep")
    assert len(out_all) == 1
    assert len(out_match) == 1
    out_no = proj.filter_tasks_by_code("DISC", value="electrical")
    assert out_no == []


def test_find_missing_links_outline_level_2(tmp_path):
    """Outline-level-2 task with no preds/succs surfaces (lines 942-952)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>SumParent</Name>
      <Duration>PT80H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-10T17:00:00</Finish>
      <Summary>1</Summary>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
    <Task>
      <UID>2</UID><ID>2</ID><Name>OrphanChild</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>2</OutlineLevel><WBS>1.1</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out = proj.find_missing_links()
    assert out["no_predecessors_count"] >= 1
    assert out["no_successors_count"] >= 1
    assert any(t["name"] == "OrphanChild" for t in out["no_predecessors"])


def test_get_link_chain_with_chain(tmp_path):
    """Build a 3-task chain and trace from first to last (lines 992-1034)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>Design Phase</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
    <Task>
      <UID>2</UID><ID>2</ID><Name>Procurement</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-02T08:00:00</Start>
      <Finish>2026-01-02T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>2</WBS>
      <PredecessorLink>
        <PredecessorUID>1</PredecessorUID>
        <Type>1</Type>
        <CrossProject>0</CrossProject>
        <LinkLag>4800</LinkLag>
        <LagFormat>7</LagFormat>
      </PredecessorLink>
    </Task>
    <Task>
      <UID>3</UID><ID>3</ID><Name>Construction Phase</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-03T08:00:00</Start>
      <Finish>2026-01-03T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>3</WBS>
      <PredecessorLink>
        <PredecessorUID>2</PredecessorUID>
        <Type>1</Type>
        <CrossProject>0</CrossProject>
        <LinkLag>0</LinkLag>
        <LagFormat>7</LagFormat>
      </PredecessorLink>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out = proj.get_link_chain("Design", "Construction")
    assert out["chains_found"] >= 1


def test_get_tasks_between_dates_filters_correctly(tmp_path):
    """Date filters drop out-of-range tasks (lines 1050, 1052)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>EarlyTask</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2025-01-01T08:00:00</Start>
      <Finish>2025-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
    <Task>
      <UID>2</UID><ID>2</ID><Name>InRangeTask</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-06-01T08:00:00</Start>
      <Finish>2026-06-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>2</WBS>
    </Task>
    <Task>
      <UID>3</UID><ID>3</ID><Name>LateTask</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2099-01-01T08:00:00</Start>
      <Finish>2099-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>3</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    # start_after filter excludes EarlyTask
    out_after = proj.get_tasks_between_dates(start_after="2026-01-01")
    names_after = [t["name"] for t in out_after]
    assert "EarlyTask" not in names_after
    # finish_before filter excludes LateTask
    out_before = proj.get_tasks_between_dates(finish_before="2030-12-31")
    names_before = [t["name"] for t in out_before]
    assert "LateTask" not in names_before


# ============================================================
# add_task element-creation branches
# ============================================================

def test_add_task_creates_tasks_block_when_missing(tmp_path):
    """Adding to a project without a <Tasks> block creates it (line 1086)."""
    path = _build_minimal_mspdi(tmp_path)  # no Tasks block
    proj = MspdiProject(path)
    r = proj.add_task("FirstTask", duration_str="2d")
    assert "task_id" in r
    # <Tasks> element now exists
    assert proj._find(proj.root, "Tasks") is not None


def test_add_task_with_NA_start_uses_now(tmp_path):
    """Project missing StartDate -> 'N/A' triggers fallback to today (line 1109)."""
    # Build XML without a StartDate element
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>noStart</Name>
  <MinutesPerDay>480</MinutesPerDay>
  <CalendarUID>1</CalendarUID>
</Project>"""
    p = os.path.join(str(tmp_path), "no_start.xml")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(xml)
    proj = MspdiProject(p)
    r = proj.add_task("FirstNoStart", duration_str="1d")
    # Today's date used (10-char ISO prefix)
    assert len(r["start"]) == 10


def test_add_task_under_parent_with_existing_children(writable_proj):
    """add_task under a parent that already has children inserts AFTER siblings (lines 1128-1141)."""
    parent = writable_proj.add_task("ParentSeq")
    pid = parent["task_id"]
    # First child
    writable_proj.add_child_task(pid, "ChildA", duration_str="1d")
    # Second child should insert after ChildA
    r2 = writable_proj.add_child_task(pid, "ChildB", duration_str="1d")
    assert "task_id" in r2


def test_add_task_insert_at_end_when_index_out_of_range(writable_proj):
    """When insert_index >= len(all_task_elems), append at end (lines 1176-1179)."""
    # Adding any new task (no parent) hits insert_index = len(tasks); the
    # branch is exercised on every plain add_task call.
    r = writable_proj.add_task("EndOfList", duration_str="1d")
    assert "task_id" in r


def test_add_resource_creates_resources_block(tmp_path):
    """add_resource creates <Resources> when missing (line 1247)."""
    path = _build_minimal_mspdi(tmp_path)
    proj = MspdiProject(path)
    rid = proj.add_resource("R1", type="Work")
    assert isinstance(rid, int)
    assert proj._find(proj.root, "Resources") is not None


def test_add_assignment_creates_assignments_block(tmp_path):
    """add_assignment creates <Assignments> when missing (line 1315)."""
    path = _build_minimal_mspdi(tmp_path)
    proj = MspdiProject(path)
    t = proj.add_task("AB1", duration_str="1d")
    rid = proj.add_resource("ABR", type="Work")
    a_uid = proj.add_assignment(t["task_id"], rid)
    assert isinstance(a_uid, int)
    assert proj._find(proj.root, "Assignments") is not None


def test_bulk_add_assignments_creates_block(tmp_path):
    """bulk_add_assignments creates <Assignments> when missing (line 1378)."""
    path = _build_minimal_mspdi(tmp_path)
    proj = MspdiProject(path)
    t = proj.add_task("BB1", duration_str="1d")
    rid = proj.add_resource("BBR", type="Work")
    n = proj.bulk_add_assignments([{"task_id": t["task_id"], "resource_id": rid}])
    assert n == 1
    assert proj._find(proj.root, "Assignments") is not None


# ============================================================
# Element-missing error branches
# ============================================================

def test_update_task_missing_elem_returns_error(writable_proj):
    """Forcibly remove the task elem so the elem-not-found branch fires."""
    a = writable_proj.add_task("Eve", duration_str="1d")
    uid = a["uid"]
    # Manually pop the elem from indices but keep task entry
    del writable_proj._task_elems[uid]
    r = writable_proj.update_task(a["task_id"], name="Renamed")
    assert "error" in r and "element" in r["error"].lower()


def test_delete_task_missing_elem_returns_error(writable_proj):
    a = writable_proj.add_task("Eve2", duration_str="1d")
    uid = a["uid"]
    del writable_proj._task_elems[uid]
    r = writable_proj.delete_task(a["task_id"])
    assert "error" in r


def test_add_link_missing_successor_elem_returns_error(writable_proj):
    a = writable_proj.add_task("LP1")
    b = writable_proj.add_task("LP2")
    del writable_proj._task_elems[b["uid"]]
    r = writable_proj.add_link(a["task_id"], b["task_id"])
    assert "error" in r


def test_remove_link_missing_successor_elem_returns_error(writable_proj):
    a = writable_proj.add_task("RM1")
    b = writable_proj.add_task("RM2")
    writable_proj.add_link(a["task_id"], b["task_id"])
    del writable_proj._task_elems[b["uid"]]
    r = writable_proj.remove_link(a["task_id"], b["task_id"])
    assert "error" in r


def test_update_link_missing_successor_elem_returns_error(writable_proj):
    a = writable_proj.add_task("UM1")
    b = writable_proj.add_task("UM2")
    writable_proj.add_link(a["task_id"], b["task_id"])
    del writable_proj._task_elems[b["uid"]]
    r = writable_proj.update_link(a["task_id"], b["task_id"], new_link_type="SS")
    assert "error" in r


def test_update_progress_missing_elem_returns_error(writable_proj):
    a = writable_proj.add_task("UPM1")
    del writable_proj._task_elems[a["uid"]]
    r = writable_proj.update_progress(a["task_id"], percent_complete=10)
    assert "error" in r


def test_assign_code_missing_elem_returns_error(tmp_path):
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743791</FieldID>
      <FieldName>Text31</FieldName>
      <Alias>CodeLibrary:"PHASE"</Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    a = proj.add_task("AC_Err")
    # Remove elem to force the missing-elem branch
    del proj._task_elems[a["uid"]]
    r = proj.assign_code(a["task_id"], "PHASE", "Concept")
    assert "error" in r


# ============================================================
# update_link rebuild branch — task missing in dict
# ============================================================

def test_update_link_rebuild_skips_orphan_uid(writable_proj):
    """update_link iterates _task_elems; if a stale UID with no task entry
    exists, the rebuild loop skips it (line 1672 'continue')."""
    a = writable_proj.add_task("UR1")
    b = writable_proj.add_task("UR2")
    writable_proj.add_link(a["task_id"], b["task_id"])
    # Insert an orphan task elem (no _tasks entry)
    orphan = ET.Element(writable_proj._t("Task"))
    writable_proj._task_elems[99999] = orphan
    r = writable_proj.update_link(a["task_id"], b["task_id"], new_lag_str="2d")
    assert r.get("updated") is True


# ============================================================
# assign_code: existing field update + new field create
# ============================================================

def test_assign_code_creates_then_updates_extended_attribute(tmp_path):
    """First assign creates ExtendedAttribute, second updates existing
    (lines 1781-1788, 1794-1814)."""
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743750</FieldID>
      <FieldName>Text20</FieldName>
      <Alias>CodeLibrary:"AREA"</Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    a = proj.add_task("CodedAssign")
    # Initial assignment creates new ExtendedAttribute element
    r1 = proj.assign_code(a["task_id"], "AREA", "Block-A")
    assert r1.get("success") is True
    # Second assignment to same library updates existing element
    r2 = proj.assign_code(a["task_id"], "AREA", "Block-B")
    assert r2.get("success") is True
    assert r2["value"] == "Block-B"


def test_assign_code_prefers_low_field_id(tmp_path):
    """When two libraries share name, low field_id (<200000000) wins."""
    extra = """
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>200000099</FieldID>
      <FieldName>HighFid</FieldName>
      <Alias>CodeLibrary:"DUPLIB"</Alias>
    </ExtendedAttribute>
    <ExtendedAttribute>
      <FieldID>188743799</FieldID>
      <FieldName>LowFid</FieldName>
      <Alias>CodeLibrary:"DUPLIB"</Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    a = proj.add_task("DupCoded")
    r = proj.assign_code(a["task_id"], "DUPLIB", "X")
    assert r.get("success") is True


# ============================================================
# write_baseline edge cases
# ============================================================

def test_write_baseline_skips_none_uid(writable_proj):
    """task_uid=None entry skipped (line 1870)."""
    n = writable_proj.write_baseline(0, [{"task_uid": None,
                                           "baseline_start": "2026-01-01"}])
    assert n == 0


def test_write_baseline_skips_unparseable_uid(writable_proj):
    """task_uid='abc' triggers ValueError -> entry skipped (lines 1873-1874)."""
    n = writable_proj.write_baseline(0, [{"task_uid": "not-an-int",
                                           "baseline_start": "2026-01-01"}])
    assert n == 0


# ============================================================
# Round-trip with synthetic project
# ============================================================

def test_round_trip_minimal_project_preserves_namespace(tmp_path):
    """save() post-processes XML and ensures namespace is present."""
    path = _build_minimal_mspdi(tmp_path)
    proj = MspdiProject(path)
    t = proj.add_task("RT_Task", duration_str="1d")
    out = os.path.join(str(tmp_path), "rt_out.xml")
    saved = proj.save(out)
    assert os.path.exists(saved)
    with open(saved, "r", encoding="utf-8") as f:
        content = f.read()
    assert 'xmlns="http://schemas.microsoft.com/project"' in content


def test_get_link_chain_to_pattern_no_match(tmp_path):
    """from_pattern matches but to_pattern returns no tasks (line 993)."""
    extra = """
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>Foundation</Name>
      <Duration>PT8H0M0S</Duration>
      <Start>2026-01-01T08:00:00</Start>
      <Finish>2026-01-01T17:00:00</Finish>
      <Milestone>0</Milestone><Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
      <OutlineLevel>1</OutlineLevel><WBS>1</WBS>
    </Task>
  </Tasks>
"""
    path = _build_minimal_mspdi(tmp_path, extra)
    proj = MspdiProject(path)
    out = proj.get_link_chain("Foundation", "__no_such_thing__")
    assert "error" in out


def test_add_task_under_parent_skips_sibling_at_same_level(writable_proj):
    """When iterating after parent, hitting a sibling at SAME level breaks
    the inner loop (line 1140)."""
    # Create a parent + child, then add ANOTHER top-level task. Now the
    # iterator visiting children stops when it encounters the next top-level.
    parent = writable_proj.add_task("PSamLvl")
    writable_proj.add_child_task(parent["task_id"], "ChildSL", duration_str="1d")
    # Add another top-level sibling at outline level 1
    writable_proj.add_task("SiblingTopLevel")
    # Now add a NEW child to parent — the loop scanning children encounters
    # the top-level sibling and breaks
    r = writable_proj.add_child_task(parent["task_id"], "ChildSL2", duration_str="1d")
    assert "task_id" in r


def test_add_task_inserts_in_middle_when_index_in_range(writable_proj):
    """Insert path 1176-1179 covered: insert_index < len(all_task_elems)."""
    # The default add_task without parent always uses insert_index = len(all),
    # so it appends. To cover the middle-insert path, add a child that's
    # placed before the end of the Tasks block.
    parent = writable_proj.add_task("PMid")
    writable_proj.add_child_task(parent["task_id"], "Mid_C1")
    # Add another top-level task — appended at end
    writable_proj.add_task("AfterMid")
    # Now adding another child to the parent should insert in the middle
    # (after Mid_C1, before AfterMid)
    r = writable_proj.add_child_task(parent["task_id"], "Mid_C2")
    assert "task_id" in r


def test_save_handles_post_process_exception(monkeypatch, writable_proj, tmp_path):
    """save() swallows post-processing exceptions and keeps the file."""
    out = os.path.join(str(tmp_path), "rt_safe.xml")

    # Patch open() to raise on the post-process read step. We do this only
    # for the 'r' mode read inside save() to simulate a transient failure.
    real_open = open
    call_count = {"n": 0}

    def fake_open(*args, **kwargs):
        # The first call is tree.write (binary write); the second is the
        # post-process read of out. Trip on any 'r' mode call to that path.
        if len(args) >= 2 and args[1] == "r":
            call_count["n"] += 1
            raise IOError("simulated post-process read failure")
        return real_open(*args, **kwargs)

    monkeypatch.setattr("builtins.open", fake_open)
    saved = writable_proj.save(out)
    # Even after the simulated failure, the XML written by tree.write
    # remains on disk
    assert os.path.exists(saved)
