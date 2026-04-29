"""Test msproject_baseline list action."""
import pytest
from msproject_mcp_core import (
    _msp_baseline_save, _msp_baseline_list, _msp_task_add_single,
)


def test_list_empty_project(clean_test_project):
    """No baselines saved → count_saved=0, baselines=[]."""
    r = _msp_baseline_list()
    assert r["status"] == "ok"
    assert r["count_saved"] == 0
    assert r["baselines"] == []


def test_list_one_saved(clean_test_project):
    """Save baseline 0 → list shows 1 entry with metadata."""
    _msp_task_add_single(name="ListT-T42", duration="3d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_list()
    assert r["status"] == "ok"
    assert r["count_saved"] == 1
    assert len(r["baselines"]) == 1
    entry = r["baselines"][0]
    assert entry["number"] == 0
    assert entry["saved_date"] is not None
    assert entry["task_count"] == 1


def test_list_three_saved(clean_test_project):
    """Save baselines 0, 3, 7 → list shows all 3 sorted by number."""
    _msp_task_add_single(name="MultiT-T42", duration="2d")
    _msp_baseline_save(baseline_number=0)
    _msp_baseline_save(baseline_number=3)
    _msp_baseline_save(baseline_number=7)
    r = _msp_baseline_list()
    assert r["status"] == "ok"
    assert r["count_saved"] == 3
    numbers = [b["number"] for b in r["baselines"]]
    assert numbers == [0, 3, 7]  # sorted ascending


def test_list_returns_user_provided_name(clean_test_project):
    """Save with name='Original' -> list shows the name (TAIL #3)."""
    _msp_task_add_single(name="NamedT-T50", duration="2d")
    _msp_baseline_save(baseline_number=0, name="Original")
    r = _msp_baseline_list()
    assert r["status"] == "ok"
    assert r["count_saved"] == 1
    assert r["baselines"][0]["name"] == "Original"


def test_list_returns_none_when_no_name_provided(clean_test_project):
    """Save without name -> list shows name=None (no false retention)."""
    _msp_task_add_single(name="UnnamedT-T50", duration="2d")
    _msp_baseline_save(baseline_number=0)
    r = _msp_baseline_list()
    assert r["baselines"][0]["name"] is None
