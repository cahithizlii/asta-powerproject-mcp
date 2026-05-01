"""Test Phase 5c T99 export action helpers."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _msp_excel_export_hakedis, _msp_excel_export_tasks,
    _msp_excel_export_evm, _msp_excel_export_dcma,
)
from openpyxl import load_workbook

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_export_hakedis_creates_workbook(tmp_path):
    xlsx = tmp_path / "hak.xlsx"
    r = _msp_excel_export_hakedis(file_path=MSP_XML, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    assert xlsx.exists()
    wb = load_workbook(str(xlsx), read_only=True)
    for s in ("Summary", "Tasks", "EVM_Compute", "EVM_TimePhased",
              "DCMA_Rules", "DCMA_Failed"):
        assert s in wb.sheetnames


def test_export_hakedis_missing_xlsx_path():
    r = _msp_excel_export_hakedis(file_path=MSP_XML)
    assert r["status"] == "error"
    assert "xlsx_path" in r["error"]


def test_export_tasks_creates_single_sheet(tmp_path):
    xlsx = tmp_path / "tasks.xlsx"
    r = _msp_excel_export_tasks(file_path=MSP_XML, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    wb = load_workbook(str(xlsx), read_only=True)
    assert "Tasks" in wb.sheetnames


def test_export_evm_creates_two_sheets(tmp_path):
    xlsx = tmp_path / "evm.xlsx"
    r = _msp_excel_export_evm(file_path=MSP_XML, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    wb = load_workbook(str(xlsx), read_only=True)
    assert "EVM_Compute" in wb.sheetnames
    assert "EVM_TimePhased" in wb.sheetnames


def test_export_dcma_creates_rules_sheet(tmp_path):
    xlsx = tmp_path / "dcma.xlsx"
    r = _msp_excel_export_dcma(file_path=MSP_XML, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    wb = load_workbook(str(xlsx), read_only=True)
    assert "DCMA_Rules" in wb.sheetnames


def test_export_evm_invalid_bucket(tmp_path):
    """bucket must be day/week/month."""
    xlsx = tmp_path / "evm.xlsx"
    r = _msp_excel_export_evm(file_path=MSP_XML, xlsx_path=str(xlsx),
                              bucket="quarterly")
    assert r["status"] == "error"
