"""Test hybrid file + COM data source adapter."""
import os
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _evm_load_task_data,
    _evm_load_progress_data,
    _evm_load_baseline_data,
    _evm_detect_currency_mode,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_evm_load_task_data_xml():
    """file_path -> reads via Phase 4 helpers, returns task list."""
    r = _evm_load_task_data(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert len(r["tasks"]) == 3  # sample fixture
    for t in r["tasks"]:
        for k in ("id", "name", "duration_h"):
            assert k in t


def test_evm_load_task_data_xml_includes_resources():
    r = _evm_load_task_data(file_path=MSP_XML)
    assert "resources" in r
    assert len(r["resources"]) == 2  # R1, R2


def test_evm_load_progress_data_xml():
    r = _evm_load_progress_data(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "tasks" in r


def test_evm_load_baseline_data_xml():
    r = _evm_load_baseline_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0


def test_evm_load_baseline_data_invalid_number():
    r = _evm_load_baseline_data(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"
    assert "0-10" in r["error"]


def test_evm_detect_currency_mode_hours():
    """Sample fixture has zero costs -> hours mode."""
    load = _evm_load_task_data(file_path=MSP_XML)
    mode = _evm_detect_currency_mode(load["tasks"], load["resources"])
    assert mode == "hours"
