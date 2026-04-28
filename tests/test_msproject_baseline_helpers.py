"""Test baseline helpers + constants."""
import pytest
from msproject_mcp_core import (
    BASELINE_NUMBERS, INTO_BASELINE_MAP,
    _baseline_property_name, _baseline_into_code, _read_task_baseline,
    _baseline_saved_date, _msp_task_add_single,
)


def test_baseline_numbers_constant():
    assert BASELINE_NUMBERS == list(range(11))  # 0..10


def test_baseline_property_name_zero():
    """Baseline 0 has no suffix: BaselineStart."""
    assert _baseline_property_name("Start", 0) == "BaselineStart"
    assert _baseline_property_name("Finish", 0) == "BaselineFinish"
    assert _baseline_property_name("Duration", 0) == "BaselineDuration"
    assert _baseline_property_name("Work", 0) == "BaselineWork"
    assert _baseline_property_name("Cost", 0) == "BaselineCost"


def test_baseline_property_name_numbered():
    """Baseline 3 -> Baseline3Start etc."""
    assert _baseline_property_name("Start", 3) == "Baseline3Start"
    assert _baseline_property_name("Finish", 10) == "Baseline10Finish"


def test_baseline_into_code_zero():
    """Baseline 0 -> Into=0 (pjIntoBaseline)."""
    assert _baseline_into_code(0) == 0


def test_baseline_into_code_offset():
    """Baseline 1 -> 11, Baseline 2 -> 12, ..., Baseline 10 -> 20."""
    assert _baseline_into_code(1) == 11
    assert _baseline_into_code(5) == 15
    assert _baseline_into_code(10) == 20


def test_into_baseline_map():
    """Constant matches formula."""
    assert INTO_BASELINE_MAP == {0: 0, 1: 11, 2: 12, 3: 13, 4: 14, 5: 15,
                                  6: 16, 7: 17, 8: 18, 9: 19, 10: 20}


def test_baseline_saved_date_unsaved(clean_test_project):
    """Fresh project — baseline 0 not saved → returns None."""
    proj = clean_test_project
    result = _baseline_saved_date(proj, 0)
    assert result is None


def test_read_task_baseline_unsaved(clean_test_project):
    """Task with no baseline saved — read returns dict with None/zero values."""
    add_r = _msp_task_add_single(name="UnsavedT-T39", duration="3d")
    proj = clean_test_project
    from msproject_mcp_core import _find_task_by_id
    t = _find_task_by_id(proj, add_r["task_id"])
    data = _read_task_baseline(t, 0)
    # Unsaved baseline returns sentinel-empty values per MSP COM
    assert "start" in data and "finish" in data
    assert "duration_h" in data and "work_h" in data and "cost" in data
