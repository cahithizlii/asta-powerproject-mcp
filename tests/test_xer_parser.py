"""Test pure-Python XER parser (Phase 5d)."""
import pytest
from xer_parser import XerFile


# ---------- T102: Foundations + encoding detect + table splitter ----------

def test_xer_file_parses(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    assert x is not None


def test_xer_header_fields(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    assert x.header_fields["version"] == "18.8"
    assert x.header_fields["currency"] == "USD"


def test_xer_tables_present(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    for tbl in ("PROJECT", "CALENDAR", "RSRC", "TASK", "TASKPRED", "TASKRSRC"):
        assert tbl in x.tables, f"missing table {tbl}"


def test_xer_task_table_row_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    assert len(x.tables["TASK"]["rows"]) == 6


def test_xer_task_table_headers(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    headers = x.tables["TASK"]["headers"]
    assert "task_id" in headers
    assert "task_name" in headers
    assert "phys_complete_pct" in headers


def test_xer_task_first_row_dict(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    row = x.tables["TASK"]["rows"][0]
    assert row["task_code"] == "A1010"
    assert row["task_name"] == "Foundation"
    assert row["task_type"] == "TT_Task"


def test_xer_handles_utf8_no_bom(tmp_path):
    """UTF-8 XER (no BOM) should also parse via fallback."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "u8.xer"
    path.write_bytes(content.encode("utf-8"))
    x = XerFile(str(path))
    assert "TASK" in x.tables
    assert x.tables["TASK"]["rows"][0]["task_name"] == "T1"


def test_xer_file_not_found():
    with pytest.raises(FileNotFoundError):
        XerFile("/definitely/nonexistent.xer")


def test_xer_empty_table_handled(tmp_path):
    """Table with header but no rows (empty section)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\n%E\n")
    path = tmp_path / "e.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert "TASK" in x.tables
    assert x.tables["TASK"]["rows"] == []


def test_xer_padded_row(tmp_path):
    """Row with fewer fields than headers - pad with empty strings."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\tdrtn_hr_cnt\n"
               "%R\t1\tT1\n%E\n")  # only 2 values for 3 headers
    path = tmp_path / "p.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    row = x.tables["TASK"]["rows"][0]
    assert row["task_id"] == "1"
    assert row["task_name"] == "T1"
    assert row["drtn_hr_cnt"] == ""


def test_xer_unknown_marker_skipped(tmp_path):
    """%X or other unknown markers should be silently ignored (forward-compat)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%X\tFutureMarker\n"  # unknown
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "x.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert "TASK" in x.tables
    assert len(x.tables["TASK"]["rows"]) == 1


# ---------- T103: read_tasks + read_links ----------

def test_read_tasks_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    assert len(tasks) == 6


def test_read_tasks_msp_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    t = tasks[0]
    for k in ("id", "name", "code", "duration_h", "start", "finish",
              "percent_complete", "summary", "constraint_type", "status"):
        assert k in t


def test_read_tasks_id_int(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    assert isinstance(tasks[0]["id"], int)
    assert tasks[0]["id"] == 1001


def test_read_tasks_milestone_summary_flag(sample_cau_xer):
    """TT_FinMile is a milestone, not summary."""
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    handover = next(t for t in tasks if t["id"] == 1006)
    assert handover["summary"] is False
    assert handover["task_type"] == "TT_FinMile"


def test_read_tasks_constraint_type_mapped(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    handover = next(t for t in tasks if t["id"] == 1006)
    assert handover["constraint_type"] == 3  # CS_MFO -> 3
    foundation = next(t for t in tasks if t["id"] == 1001)
    assert foundation["constraint_type"] == 0  # CS_ASAP -> 0


def test_read_tasks_percent_complete(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    foundation = next(t for t in tasks if t["id"] == 1001)
    assert foundation["percent_complete"] == 100.0
    frame = next(t for t in tasks if t["id"] == 1002)
    assert frame["percent_complete"] == 75.0


def test_read_tasks_total_float_days_default_8h(sample_cau_xer):
    """Default day_hr_cnt=8: 72h / 8 = 9 days."""
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    walls = next(t for t in tasks if t["id"] == 1003)
    assert walls["total_float"] == 9.0


def test_read_tasks_total_float_days_cau_9h(sample_cau_xer):
    """CAU calendar 9h/day: 72h / 9 = 8 days."""
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks(day_hr_cnt=9.0)
    walls = next(t for t in tasks if t["id"] == 1003)
    assert walls["total_float"] == 8.0


def test_read_tasks_dates_iso(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    foundation = next(t for t in tasks if t["id"] == 1001)
    assert foundation["start"] == "2024-07-08"
    assert foundation["finish"] == "2024-07-29"
    assert foundation["actual_finish"] == "2024-07-29"
    # Frame still in progress (no actual_finish)
    frame = next(t for t in tasks if t["id"] == 1002)
    assert frame["actual_finish"] is None


# ---- Links ----

def test_read_links_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    assert len(links) == 5


def test_read_links_msp_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    link = x.read_links()[0]
    for k in ("from_id", "to_id", "type", "lag_days"):
        assert k in link


def test_read_links_type_mapping_all_fs(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    assert all(l["type"] == "FS" for l in links)


def test_read_links_zero_lag(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    assert all(l["lag_days"] == 0 for l in links)


def test_read_links_pred_to_succ_mapping(sample_cau_xer):
    """from_id = predecessor (pred_task_id), to_id = successor (task_id)."""
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    first = links[0]
    assert first["from_id"] == 1001  # Foundation predecessor
    assert first["to_id"] == 1002    # Frame successor


def test_read_links_lag_conversion(tmp_path):
    """Lag in hours converted to days at 8h/day."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASKPRED\n%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt\n"
               "%R\t1\t2\t1\tPR_SS\t16.0\n%E\n")
    path = tmp_path / "lag.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    links = x.read_links()
    assert links[0]["lag_days"] == 2.0  # 16 / 8
    assert links[0]["type"] == "SS"


# ---------- T104: read_resources + read_assignments + read_calendars ----------

def test_read_resources_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    res = x.read_resources()
    assert len(res) == 4


def test_read_resources_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    res = x.read_resources()
    r = res[0]
    for k in ("id", "name", "code", "type", "max_units"):
        assert k in r


def test_read_resources_cow(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    res = x.read_resources()
    cow = next(r for r in res if r["id"] == 101)
    assert cow["name"] == "Concrete Workers"
    assert cow["code"] == "COW"
    assert cow["type"] == "Work"
    assert cow["max_units"] == 10.0


def test_read_resources_material_type_mapped(sample_cau_xer):
    """RT_Mat (STL) -> 'Material'."""
    x = XerFile(sample_cau_xer)
    res = x.read_resources()
    stl = next(r for r in res if r["id"] == 103)
    assert stl["type"] == "Material"


# ---- Assignments ----

def test_read_assignments_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    a = x.read_assignments()
    assert len(a) == 7


def test_read_assignments_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    a = x.read_assignments()
    item = a[0]
    for k in ("task_id", "resource_id", "target_qty", "actual_qty",
              "target_cost", "actual_cost"):
        assert k in item


def test_read_assignments_task_resource_link(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    a = x.read_assignments()
    foundation = [x for x in a if x["task_id"] == 1001]
    assert len(foundation) == 2
    assert {x["resource_id"] for x in foundation} == {101, 103}


def test_read_assignments_actual_qty(sample_cau_xer):
    """Foundation completed: actual_qty = target_qty for both res."""
    x = XerFile(sample_cau_xer)
    a = x.read_assignments()
    cow_foundation = next(x for x in a if x["task_id"] == 1001 and x["resource_id"] == 101)
    assert cow_foundation["actual_qty"] == 180.0
    assert cow_foundation["target_qty"] == 180.0
    # Walls not started: actual = 0
    cow_walls = next(x for x in a if x["task_id"] == 1003 and x["resource_id"] == 101)
    assert cow_walls["actual_qty"] == 0.0


# ---- Calendars ----

def test_read_calendars_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    cals = x.read_calendars()
    assert len(cals) == 1


def test_read_calendars_cau_9h_day(sample_cau_xer):
    """CLAUDE.md RULE 1: CAU calendar 6x9 = 54h/week, 9h/day."""
    x = XerFile(sample_cau_xer)
    cals = x.read_calendars()
    cal = cals[0]
    assert cal["day_hr_cnt"] == 9.0
    assert cal["week_hr_cnt"] == 54.0
    assert cal["name"] == "CAU 6x9"


def test_read_calendars_default_when_missing(tmp_path):
    """Empty CALENDAR rows -> day_hr_cnt defaults handled."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tCALENDAR\n%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt\n%E\n")
    path = tmp_path / "c.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    cals = x.read_calendars()
    assert cals == []


# ---------- T105: read_progress + status_date + project metadata ----------

def test_read_project_metadata(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    proj = x.read_project()
    assert proj["proj_id"] == 1
    assert proj["proj_short_name"] == "CAU"
    assert proj["plan_start_date"] == "2024-07-08"
    assert proj["plan_end_date"] == "2028-06-20"
    assert proj["last_recalc_date"] == "2026-05-01"


def test_read_project_empty_when_no_section(tmp_path):
    """Missing PROJECT section -> empty dict (no crash)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "n.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert x.read_project() == {}


def test_read_progress_status_date(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    p = x.read_progress()
    assert p["status_date"] == "2026-05-01"


def test_read_progress_task_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    p = x.read_progress()
    assert len(p["tasks"]) == 6


def test_read_progress_percent_complete(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    p = x.read_progress()
    foundation = next(t for t in p["tasks"] if t["id"] == 1001)
    assert foundation["percent_complete"] == 100.0


def test_read_progress_actual_work_aggregated(sample_cau_xer):
    """actual_work_h = sum of TASKRSRC.act_reg_qty per task."""
    x = XerFile(sample_cau_xer)
    p = x.read_progress()
    # Foundation: COW 180 + STL 1000 = 1180
    foundation = next(t for t in p["tasks"] if t["id"] == 1001)
    assert foundation["actual_work_h"] == 1180.0
    # Frame: COW 270 + CAR 135 = 405
    frame = next(t for t in p["tasks"] if t["id"] == 1002)
    assert frame["actual_work_h"] == 405.0
    # Walls: not started, all 0
    walls = next(t for t in p["tasks"] if t["id"] == 1003)
    assert walls["actual_work_h"] == 0.0


def test_read_progress_no_project_table(tmp_path):
    """XER without PROJECT -> status_date None."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "n.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    p = x.read_progress()
    assert p["status_date"] is None
    assert len(p["tasks"]) == 1


# === Phase 11.1 T141: gap-fill ===

def test_xer_handles_utf8_bom(tmp_path):
    """UTF-8 BOM (EF BB BF) is detected and stripped (line 51)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1_utf8bom\n%E\n")
    path = tmp_path / "u8bom.xer"
    path.write_bytes(b"\xef\xbb\xbf" + content.encode("utf-8"))
    x = XerFile(str(path))
    assert "TASK" in x.tables
    assert x.tables["TASK"]["rows"][0]["task_name"] == "T1_utf8bom"


def test_xer_skips_unknown_marker(tmp_path):
    """Unknown markers like '%Z' are silently skipped (line 102/103)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%Z unknown marker line\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "u.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    # Parser tolerated the %Z marker
    assert x.tables["TASK"]["rows"][0]["task_name"] == "T1"


def test_xer_skips_blank_lines(tmp_path):
    """Empty lines are skipped (line 64)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "\n\n"  # two blank lines
               "%T\tTASK\n%F\ttask_id\ttask_name\n"
               "\n"   # blank line in middle
               "%R\t1\tT1\n%E\n")
    path = tmp_path / "b.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert x.tables["TASK"]["rows"][0]["task_name"] == "T1"


def test_xer_skips_field_marker_without_table(tmp_path):
    """`%F` before any `%T` is ignored (line 83 continue)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%F\torphan_field\n"  # %F before any %T
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "of.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert x.tables["TASK"]["rows"][0]["task_name"] == "T1"


def test_xer_skips_row_marker_without_table(tmp_path):
    """`%R` before any `%T` is ignored (line 89 continue)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%R\torphan_value\n"  # %R before any %T
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "or.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert x.tables["TASK"]["rows"][0]["task_name"] == "T1"


def test_xer_skips_row_when_headers_empty(tmp_path):
    """`%R` after `%T` but before `%F` (no headers yet) is skipped (line 92)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n"
               "%R\t1\tT1\n"   # %R before %F (headers still empty)
               "%F\ttask_id\ttask_name\n"
               "%R\t2\tT2\n%E\n")
    path = tmp_path / "rb.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    # First %R was dropped; second should be present
    assert len(x.tables["TASK"]["rows"]) == 1
    assert x.tables["TASK"]["rows"][0]["task_name"] == "T2"


def test_to_int_handles_invalid_string(sample_cau_xer):
    """_to_int returns default on non-numeric strings (lines 127-128)."""
    from xer_parser import _to_int
    assert _to_int("not-a-number") is None
    assert _to_int("not-a-number", default=42) == 42
    assert _to_int(None) is None
    assert _to_int("") is None
    assert _to_int("123") == 123
    assert _to_int("123.45") == 123


def test_to_float_handles_invalid_string(sample_cau_xer):
    """_to_float returns default on non-numeric strings (lines 134-135)."""
    from xer_parser import _to_float
    assert _to_float("not-a-number") == 0.0
    assert _to_float("not-a-number", default=99.0) == 99.0
    assert _to_float("") == 0.0
    assert _to_float("3.14") == 3.14


def test_parse_clndr_data_handles_overflow_and_invalid_serials():
    """_parse_clndr_data skips serials that ValueError or OverflowError."""
    from xer_parser import _parse_clndr_data
    # None / empty -> []
    assert _parse_clndr_data(None) == []
    assert _parse_clndr_data("") == []
    # No exception block -> []
    assert _parse_clndr_data("(no exceptions here)") == []
    # Valid block
    blob = "(d|44197|f|0)(d|44204|f|1)"  # 2021-01-01 holiday and an override
    out = _parse_clndr_data(blob)
    assert len(out) == 2
    assert out[0]["working"] is False
    assert out[1]["working"] is True
    # Overflow path: extremely huge serial that overflows timedelta
    big_blob = f"(d|999999999999999|f|0)"
    # Should not raise; just skip the bad entry
    out2 = _parse_clndr_data(big_blob)
    assert out2 == []


def test_read_progress_skips_assignments_with_no_task_id(tmp_path):
    """TASKRSRC rows with missing task_id are skipped (line 329)."""
    content = (
        "ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
        "%T\tTASK\n%F\ttask_id\ttask_name\tphys_complete_pct\n"
        "%R\t1\tT1\t50\n"
        "%T\tTASKRSRC\n%F\ttask_id\trsrc_id\tact_reg_qty\n"
        "%R\t\t10\t100\n"   # task_id missing -> skipped
        "%R\t1\t10\t250\n"
        "%E\n"
    )
    path = tmp_path / "skip.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    p = x.read_progress()
    # Only the 250 from the second assignment counted
    t1 = next(t for t in p["tasks"] if t["id"] == 1)
    assert t1["actual_work_h"] == 250.0


def test_read_tasks_zero_day_hr_cnt_yields_zero_total_float(tmp_path):
    """day_hr_cnt=0 -> total_float is 0 (avoids div-by-zero in line 167-168)."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\ttotal_float_hr_cnt\n"
               "%R\t1\tT1\t72\n%E\n")
    path = tmp_path / "z.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    tasks = x.read_tasks(day_hr_cnt=0.0)
    assert tasks[0]["total_float"] == 0.0


# === P0 #1: reend_date / forecast_finish vs target_finish (RULE 16.B) ===

def test_read_tasks_forecast_fields_present(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    t = x.read_tasks()[0]
    for k in ("target_finish", "early_finish", "late_finish",
              "forecast_finish", "wbs_id"):
        assert k in t


def test_forecast_finish_uses_reend_not_target(sample_cau_xer):
    """Frame (in progress) forecast slips past baseline target."""
    x = XerFile(sample_cau_xer)
    frame = next(t for t in x.read_tasks() if t["id"] == 1002)
    assert frame["finish"] == "2024-09-09"          # baseline target (unchanged)
    assert frame["target_finish"] == "2024-09-09"
    assert frame["forecast_finish"] == "2024-09-20"  # reend_date (later)
    assert frame["forecast_finish"] != frame["target_finish"]


def test_forecast_finish_backward_compat_finish_is_target(sample_cau_xer):
    """`finish` MUST stay = target_end_date (no breakage for old consumers)."""
    x = XerFile(sample_cau_xer)
    for t in x.read_tasks():
        assert t["finish"] == t["target_finish"]


def test_forecast_finish_fallback_chain(tmp_path):
    """No reend_date -> falls back to early_end_date -> act -> target."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\ttarget_end_date"
               "\tearly_end_date\n"
               "%R\t1\tEarlyOnly\t2024-01-10 17:00\t2024-01-20 17:00\n"
               "%R\t2\tTargetOnly\t2024-02-10 17:00\t\n"
               "%E\n")
    path = tmp_path / "fc.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    tasks = {t["id"]: t for t in x.read_tasks()}
    assert tasks[1]["forecast_finish"] == "2024-01-20"  # early_end_date
    assert tasks[2]["forecast_finish"] == "2024-02-10"  # target fallback


def test_forecast_finish_prefers_reend_over_early(tmp_path):
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\ttarget_end_date"
               "\treend_date\tearly_end_date\n"
               "%R\t1\tT\t2024-01-10 17:00\t2024-03-01 17:00\t2024-02-01 17:00\n"
               "%E\n")
    path = tmp_path / "re.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert x.read_tasks()[0]["forecast_finish"] == "2024-03-01"


# === P0 #2 foundation: read_wbs (PROJWBS) ===

def test_read_wbs_count_and_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    wbs = x.read_wbs()
    assert len(wbs) == 2
    for k in ("id", "parent_id", "name", "code", "proj_id"):
        assert k in wbs[0]


def test_read_wbs_values(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    wbs = {w["id"]: w for w in x.read_wbs()}
    assert wbs[1]["name"] == "Construction"
    assert wbs[1]["parent_id"] == 0
    assert wbs[0]["name"] == "CAU Project"


def test_read_wbs_empty_when_no_section(tmp_path):
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n")
    path = tmp_path / "nw.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert x.read_wbs() == []
