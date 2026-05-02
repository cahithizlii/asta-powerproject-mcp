"""Phase 6.4 T120-T121 — XER clndr_data exception/holiday parser tests.

Verifies:
- _parse_clndr_data returns [] on empty/None input
- single holiday pattern (f|0) -> working=False
- single working-exception pattern (f|1) -> working=True
- multiple patterns extracted in order
- Excel serial conversion correctness (44562 -> 2022-01-01)
- read_calendars integration: exceptions field present
- read_calendars when clndr_data column absent -> exceptions=[]
"""
import datetime as dt
import os
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from xer_parser import XerFile, _parse_clndr_data


# === _parse_clndr_data unit cases ===

def test_parse_clndr_empty_returns_empty_list():
    assert _parse_clndr_data("") == []
    assert _parse_clndr_data(None) == []


def test_parse_clndr_no_exception_block_returns_empty():
    """BLOB without (d|<n>|f|<bit>) pattern -> []."""
    blob = "(0||CalendarData()(0||DaysOfWeek()(0||1())))"
    assert _parse_clndr_data(blob) == []


def test_parse_clndr_single_holiday():
    """f|0 -> working: False (non-working day)."""
    blob = "(0||Exceptions(0||1(d|44562|f|0)))"
    result = _parse_clndr_data(blob)
    assert len(result) == 1
    assert result[0]["working"] is False
    assert result[0]["date"] == "2022-01-01"


def test_parse_clndr_single_working_exception():
    """f|1 -> working: True (working day override)."""
    blob = "(0||Exceptions(0||1(d|44562|f|1)))"
    result = _parse_clndr_data(blob)
    assert len(result) == 1
    assert result[0]["working"] is True


def test_parse_clndr_multiple_exceptions():
    blob = ("(0||Exceptions"
            "(0||1(d|44562|f|0))"
            "(0||2(d|44563|f|0))"
            "(0||3(d|44564|f|1)))")
    result = _parse_clndr_data(blob)
    assert len(result) == 3
    assert result[0]["date"] == "2022-01-01"
    assert result[1]["date"] == "2022-01-02"
    assert result[2]["date"] == "2022-01-03"
    assert result[0]["working"] is False
    assert result[1]["working"] is False
    assert result[2]["working"] is True


def test_parse_clndr_excel_serial_conversion_verified():
    """Excel serial epoch is 1899-12-30 (Lotus quirk preserved)."""
    # 1 -> 1899-12-31, 2 -> 1900-01-01
    blob_min = "(d|2|f|0)"
    r = _parse_clndr_data(blob_min)
    assert r[0]["date"] == "1900-01-01"


def test_parse_clndr_invalid_serial_skipped():
    """Non-integer serial chars between digits — regex won't match."""
    blob = "(d|abc|f|0)(d|44562|f|0)"
    r = _parse_clndr_data(blob)
    assert len(r) == 1  # Only valid one extracted


# === read_calendars integration ===

def test_read_calendars_no_clndr_data_column():
    """When CALENDAR rows lack clndr_data column, exceptions=[]."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tCALENDAR\n"
               "%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt\n"
               "%R\t1\tStandard\t8.0\t40.0\n"
               "%E\n")
    path = os.path.join(tempfile.gettempdir(), "p64_no_clndr.xer")
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(content.encode("utf-16-le"))
    try:
        xf = XerFile(path)
        cals = xf.read_calendars()
        assert len(cals) == 1
        assert cals[0]["exceptions"] == []
    finally:
        os.remove(path)


def test_read_calendars_with_holiday_blob():
    """clndr_data column with exception block -> exceptions populated."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tCALENDAR\n"
               "%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt\tclndr_data\n"
               "%R\t1\tCAU\t9.0\t54.0\t"
               "(0||Exceptions(0||1(d|44562|f|0)))\n"
               "%E\n")
    path = os.path.join(tempfile.gettempdir(), "p64_with_holiday.xer")
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(content.encode("utf-16-le"))
    try:
        xf = XerFile(path)
        cals = xf.read_calendars()
        assert len(cals) == 1
        assert len(cals[0]["exceptions"]) == 1
        assert cals[0]["exceptions"][0]["date"] == "2022-01-01"
        assert cals[0]["exceptions"][0]["working"] is False
    finally:
        os.remove(path)


def test_read_calendars_preserves_basic_fields():
    """Phase 6.4 must NOT regress day_hr_cnt / week_hr_cnt fields."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tCALENDAR\n"
               "%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt\tclndr_data\n"
               "%R\t1\tCAU 6x9\t9.0\t54.0\t\n"
               "%E\n")
    path = os.path.join(tempfile.gettempdir(), "p64_basic.xer")
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(content.encode("utf-16-le"))
    try:
        xf = XerFile(path)
        cals = xf.read_calendars()
        assert cals[0]["id"] == 1
        assert cals[0]["name"] == "CAU 6x9"
        assert cals[0]["day_hr_cnt"] == 9.0
        assert cals[0]["week_hr_cnt"] == 54.0
        assert cals[0]["exceptions"] == []
    finally:
        os.remove(path)
