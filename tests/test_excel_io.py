"""Test excel_io pure-Python helpers (CLAUDE.md RULE 14 brand)."""
import pytest
from openpyxl import Workbook
from excel_io import (
    BRAND_LACIVERT, BRAND_NAVY, BRAND_TURQUOISE, BRAND_LABEL_GRAY,
    ZEBRA, RAG_GREEN, RAG_AMBER, RAG_RED,
    apply_header_style, apply_rag_fill, apply_zebra_fill,
)


def test_brand_constants_exist():
    assert BRAND_LACIVERT == "FF0B1F4D"
    assert BRAND_NAVY == "FF3D4663"
    assert BRAND_TURQUOISE == "FF39B4CC"
    assert ZEBRA == "FFF0F3F8"


def test_rag_colors_distinct():
    assert RAG_GREEN.startswith("FF")
    assert RAG_AMBER.startswith("FF")
    assert RAG_RED.startswith("FF")
    assert len({RAG_GREEN, RAG_AMBER, RAG_RED}) == 3


def test_apply_header_style():
    wb = Workbook()
    ws = wb.active
    ws.append(["A", "B", "C"])
    apply_header_style(ws, row=1)
    cell = ws.cell(row=1, column=1)
    assert cell.font.bold is True
    assert cell.fill.start_color.value == BRAND_LACIVERT
    assert cell.font.color.value == "FFFFFFFF"
    assert cell.font.name == "Calibri"
    assert cell.font.size == 11


def test_apply_header_style_all_columns():
    wb = Workbook()
    ws = wb.active
    ws.append(["A", "B", "C", "D"])
    apply_header_style(ws, row=1)
    for col in range(1, 5):
        assert ws.cell(row=1, column=col).font.bold is True


def test_apply_rag_fill_pass():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="OK")
    apply_rag_fill(ws.cell(row=1, column=1), status="pass")
    assert ws.cell(row=1, column=1).fill.start_color.value == RAG_GREEN


def test_apply_rag_fill_fail():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="FAIL")
    apply_rag_fill(ws.cell(row=1, column=1), status="fail")
    assert ws.cell(row=1, column=1).fill.start_color.value == RAG_RED


def test_apply_rag_fill_amber():
    wb = Workbook()
    ws = wb.active
    cell = ws.cell(row=1, column=1, value="WARN")
    apply_rag_fill(cell, status="amber")
    assert cell.fill.start_color.value == RAG_AMBER


def test_apply_rag_fill_unknown_status_no_fill():
    """Unknown status string -> no fill applied (defensive)."""
    wb = Workbook()
    ws = wb.active
    cell = ws.cell(row=1, column=1, value="x")
    apply_rag_fill(cell, status="unknown_status")
    assert cell.fill.fill_type in (None, "none")


def test_apply_zebra_fill_even_row():
    wb = Workbook()
    ws = wb.active
    cell = ws.cell(row=2, column=1, value="x")
    apply_zebra_fill(cell, row_idx=2)
    assert cell.fill.start_color.value == ZEBRA


def test_apply_zebra_fill_odd_row_no_fill():
    wb = Workbook()
    ws = wb.active
    cell = ws.cell(row=3, column=1, value="x")
    apply_zebra_fill(cell, row_idx=3)
    assert cell.fill.fill_type in (None, "none")
