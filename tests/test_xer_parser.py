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
