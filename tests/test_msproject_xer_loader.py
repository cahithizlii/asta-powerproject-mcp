"""Test Phase 5d T106 single-collect aggregator."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import _xer_collect_full_data


def test_xer_collect_full_data_xml(sample_cau_xer):
    r = _xer_collect_full_data(sample_cau_xer)
    assert r["status"] == "ok"
    for k in ("tasks", "links", "resources", "assignments",
              "calendars", "progress", "project"):
        assert k in r


def test_xer_collect_uses_calendar_day_hr_cnt(sample_cau_xer):
    """First calendar's day_hr_cnt (CAU=9.0) should drive total_float conversion."""
    r = _xer_collect_full_data(sample_cau_xer)
    walls = next(t for t in r["tasks"] if t["id"] == 1003)
    # 72h / 9h/day = 8 days
    assert walls["total_float"] == 8.0


def test_xer_collect_no_file_path():
    r = _xer_collect_full_data(None)
    assert r["status"] == "error"


def test_xer_collect_file_not_found():
    r = _xer_collect_full_data("/definitely/nonexistent.xer")
    assert r["status"] == "error"


def test_xer_collect_falls_back_8h_when_no_calendar(tmp_path):
    """No CALENDAR section -> default 8h/day for total_float."""
    content = ("ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n"
               "%T\tTASK\n%F\ttask_id\ttask_name\ttotal_float_hr_cnt\n"
               "%R\t1\tT1\t72.0\n%E\n")
    path = tmp_path / "n.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    r = _xer_collect_full_data(str(path))
    assert r["status"] == "ok"
    assert r["tasks"][0]["total_float"] == 9.0  # 72/8
