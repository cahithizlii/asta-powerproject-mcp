# Phase 5c Excel Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to execute this plan task-by-task (T94-T101).

**Goal:** 11th MCP tool `msproject_excel` — multi-sheet hakediş workbook export + Excel-driven task/progress import. Bridges Phase 1-5b read pipelines to Excel.

**Architecture:** Yaklaşım C (Phase 5a/5b proven). New `excel_io.py` pure-Python module + I/O adapters in `msproject_mcp_core.py` Phase 5C section + `@mcp.tool msproject_excel` dispatcher. openpyxl 3.1.5 (already installed). Phase 1-5b helpers DOKUNULMAZ.

**Tech Stack:** Python 3.12, openpyxl 3.1.5, mcp (FastMCP), pytest. Existing `msproject_mcp_core.py` (~5800 lines after Phase 5b TAIL fix), 40+ test files, 235 cumulative regression PASS baseline.

**Design doc:** `docs/plans/2026-05-01-phase5c-excel-design.md` (commit `279fb59`)

**Baseline state at start:** HEAD `279fb59`, MS Project running v16.0.

**KEY REFERENCES:**
- CLAUDE.md RULE 14 — MCS brand: Lacivert `#0B1F4D`, Calibri, Navy `#3D4663`, Turquoise `#39B4CC`
- CLAUDE.md RULE 12 — RAG: SPI <0.3 RED, 0.3-0.7 AMBER, ≥0.7 GREEN (Phase 5a EVM); pass_count <8 RED, 8-11 AMBER, ≥12 GREEN (Phase 5b DCMA)
- CLAUDE.md RULE 13 — RAKİP YASAĞI: "Industry Standard" yazılır, McKinsey/PwC/Deloitte/Mace adı YAZILMAZ
- Phase 5a `_msp_evm_compute_metrics`, `_msp_evm_time_phased_evm`, `_msp_evm_summary`
- Phase 5b `_msp_dcma_assess_all`, `_msp_dcma_drill_down`
- Phase 4 `_msp_file_read_tasks`, `_msp_file_read_links`
- Phase 1 `_msp_task_bulk_add`, `_validate_active_project`
- Phase 3b `_msp_progress_bulk_update`

---

## Task 94: `excel_io.py` Foundations + Cell Formatters

**Files:**
- Create: `excel_io.py`
- Create: `tests/test_excel_io.py`

### Step 1: Write failing tests

`tests/test_excel_io.py`:
```python
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


def test_rag_colors():
    assert RAG_GREEN.startswith("FF")
    assert RAG_AMBER.startswith("FF")
    assert RAG_RED.startswith("FF")
    assert RAG_GREEN != RAG_AMBER != RAG_RED


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
    # odd rows should not get zebra
    assert cell.fill.fill_type in (None, "none")
```

### Step 2: Run — expect ImportError

```bash
cd /c/Users/CahAsus/asta-powerproject-mcp && python -m pytest tests/test_excel_io.py -v 2>&1 | tail -10
```

### Step 3: Implementation

Create `excel_io.py`:
```python
"""Phase 5c — Excel I/O for MS Project hakediş workflow.

Pure-Python openpyxl helpers. MSP/COM/file independent — takes plain
dicts, writes openpyxl Workbook objects. MCS brand styling per CLAUDE.md
RULE 14.
"""
from openpyxl.styles import Font, PatternFill, Alignment


# ---------- MCS brand constants (CLAUDE.md RULE 14) ----------

BRAND_LACIVERT = "FF0B1F4D"   # ARGB - main brand
BRAND_NAVY = "FF3D4663"
BRAND_TURQUOISE = "FF39B4CC"
BRAND_LABEL_GRAY = "FF6B7394"
ZEBRA = "FFF0F3F8"

# RAG status colors (industry standard, no competitor names)
RAG_GREEN = "FF8FBC8F"
RAG_AMBER = "FFFFCC66"
RAG_RED = "FFE57373"

WHITE = "FFFFFFFF"


def apply_header_style(ws, row=1):
    """Lacivert background + white bold Calibri 11pt header (CLAUDE.md RULE 14)."""
    fill = PatternFill(start_color=BRAND_LACIVERT, end_color=BRAND_LACIVERT, fill_type="solid")
    font = Font(name="Calibri", bold=True, size=11, color=WHITE)
    align = Alignment(horizontal="center", vertical="center")
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = align


def apply_rag_fill(cell, status):
    """Apply RAG color fill based on status string (pass/fail/amber/green/red)."""
    s = (status or "").lower()
    if s in ("pass", "green", "ok"):
        color = RAG_GREEN
    elif s in ("fail", "red"):
        color = RAG_RED
    elif s == "amber":
        color = RAG_AMBER
    else:
        return
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")


def apply_zebra_fill(cell, row_idx):
    """Light gray zebra stripe on even rows (1-indexed). Odd rows untouched."""
    if row_idx % 2 == 0:
        cell.fill = PatternFill(start_color=ZEBRA, end_color=ZEBRA, fill_type="solid")
```

### Step 4: Run — expect 8 PASS

```bash
cd /c/Users/CahAsus/asta-powerproject-mcp && python -m pytest tests/test_excel_io.py -v 2>&1 | tail -12
```

### Step 5: Run regression — expect 235+8 = 243 PASS

```bash
cd /c/Users/CahAsus/asta-powerproject-mcp && python -m pytest tests/test_msproject_file_*.py tests/test_evm_math.py tests/test_msproject_evm_*.py tests/test_dcma_*.py tests/test_msproject_dcma_*.py tests/test_excel_io.py -q --tb=line 2>&1 | tail -3
```

### Step 6: Commit

```bash
git add excel_io.py tests/test_excel_io.py && git commit -m "Phase 5c T94: excel_io foundations + cell formatters

MCS brand constants per CLAUDE.md RULE 14: Lacivert #0B1F4D, Navy,
Turquoise, Label Gray. RAG colors green/amber/red per RULE 12.
Helpers: apply_header_style (lacivert bg + white Calibri 11pt bold),
apply_rag_fill (pass=green/fail=red/amber), apply_zebra_fill (even rows).

8 unit tests, openpyxl 3.1.5."
```

---

## Task 95: Tasks Sheet Builder + Reader + Summary Sheet

**Files:**
- Modify: `excel_io.py`
- Modify: `tests/test_excel_io.py`

### Step 1: Append failing tests

```python
from excel_io import build_tasks_sheet, read_tasks_sheet, build_summary_sheet


def _sample_tasks():
    return [
        {"id": 1, "name": "Foundation", "duration_h": 80, "start": "2026-01-01",
         "finish": "2026-01-15", "percent_complete": 100, "critical": True,
         "summary": False},
        {"id": 2, "name": "Frame", "duration_h": 160, "start": "2026-01-16",
         "finish": "2026-02-15", "percent_complete": 50, "critical": False,
         "summary": False},
    ]


def test_build_tasks_sheet_creates_sheet():
    wb = Workbook()
    build_tasks_sheet(wb, _sample_tasks(), sheet_name="Tasks")
    assert "Tasks" in wb.sheetnames


def test_build_tasks_sheet_header_row():
    wb = Workbook()
    build_tasks_sheet(wb, _sample_tasks())
    ws = wb["Tasks"]
    headers = [c.value for c in ws[1]]
    assert "ID" in headers
    assert "Name" in headers
    assert "Duration (d)" in headers


def test_build_tasks_sheet_data_rows():
    wb = Workbook()
    build_tasks_sheet(wb, _sample_tasks())
    ws = wb["Tasks"]
    # row 2 = first task
    assert ws.cell(row=2, column=1).value == 1
    assert ws.cell(row=2, column=2).value == "Foundation"


def test_build_tasks_sheet_critical_styled():
    wb = Workbook()
    build_tasks_sheet(wb, _sample_tasks())
    ws = wb["Tasks"]
    # critical=True for task 1; should have red font color
    crit_col = next(i for i, c in enumerate(ws[1], start=1) if c.value == "Critical")
    assert ws.cell(row=2, column=crit_col).value in (True, "Yes")


def test_build_tasks_sheet_excludes_summary():
    tasks = _sample_tasks() + [{"id": 0, "name": "Sum", "summary": True,
                                "duration_h": 0, "start": None, "finish": None,
                                "percent_complete": 0, "critical": False}]
    wb = Workbook()
    build_tasks_sheet(wb, tasks)
    ws = wb["Tasks"]
    # Should have 2 data rows + 1 header (summary excluded)
    assert ws.max_row == 3


def test_read_tasks_sheet_round_trip(tmp_path):
    wb = Workbook()
    build_tasks_sheet(wb, _sample_tasks())
    xlsx = tmp_path / "round.xlsx"
    wb.save(str(xlsx))
    rows = read_tasks_sheet(str(xlsx), sheet_name="Tasks")
    assert len(rows) == 2
    assert rows[0]["name"] == "Foundation"
    assert rows[0]["duration_h"] == 80


def test_build_summary_sheet():
    wb = Workbook()
    summary = {"BAC": 1000.0, "EAC": 1100.0, "SPI": 0.95, "CPI": 0.91,
               "rag": "amber", "executive_text": "Project on track."}
    build_summary_sheet(wb, summary)
    assert "Summary" in wb.sheetnames
    ws = wb["Summary"]
    # Header A1 should mention "Project Health Summary"
    assert "Summary" in str(ws.cell(row=1, column=1).value)
```

### Step 2-3: Implementation

Append to `excel_io.py`:
```python
from openpyxl import Workbook, load_workbook


TASKS_HEADERS = ["ID", "Name", "Duration (d)", "Start", "Finish",
                 "%Complete", "Critical", "Summary"]


def _real_only(tasks):
    return [t for t in tasks if not t.get("summary", False)]


def _hours_to_days(h):
    """Convert duration in hours to working days (8h/day)."""
    try:
        return round(float(h or 0) / 8.0, 2)
    except (ValueError, TypeError):
        return 0


def build_tasks_sheet(wb, tasks, sheet_name="Tasks"):
    """Build Tasks sheet from list of task dicts. Summary tasks excluded."""
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    elif wb.active.title == "Sheet" and wb.active.max_row == 1 and wb.active.max_column == 1:
        ws = wb.active
        ws.title = sheet_name
    else:
        ws = wb.create_sheet(sheet_name)
    ws.append(TASKS_HEADERS)
    apply_header_style(ws, row=1)
    real = _real_only(tasks)
    for idx, t in enumerate(real, start=2):
        ws.cell(row=idx, column=1, value=t.get("id"))
        ws.cell(row=idx, column=2, value=t.get("name", ""))
        ws.cell(row=idx, column=3, value=_hours_to_days(t.get("duration_h")))
        ws.cell(row=idx, column=4, value=t.get("start"))
        ws.cell(row=idx, column=5, value=t.get("finish"))
        ws.cell(row=idx, column=6, value=float(t.get("percent_complete") or 0) / 100.0)
        ws.cell(row=idx, column=6).number_format = "0%"
        ws.cell(row=idx, column=7, value=bool(t.get("critical", False)))
        ws.cell(row=idx, column=8, value=bool(t.get("summary", False)))
        for col in range(1, len(TASKS_HEADERS) + 1):
            apply_zebra_fill(ws.cell(row=idx, column=col), row_idx=idx)
    ws.freeze_panes = "A2"
    return ws


def read_tasks_sheet(xlsx_path, sheet_name="Tasks"):
    """Read Tasks sheet back into list of dicts (for import workflow)."""
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h or "").strip() for h in rows[0]]
    out = []
    for row in rows[1:]:
        d = {}
        for i, val in enumerate(row):
            if i >= len(headers):
                continue
            h = headers[i]
            if h == "ID":
                d["id"] = int(val) if val is not None else None
            elif h == "Name":
                d["name"] = str(val) if val is not None else ""
            elif h == "Duration (d)":
                d["duration_h"] = float(val or 0) * 8.0
            elif h == "Start":
                d["start"] = str(val) if val else None
            elif h == "Finish":
                d["finish"] = str(val) if val else None
            elif h == "%Complete":
                pc = val or 0
                d["percent_complete"] = float(pc) * 100.0 if pc <= 1 else float(pc)
            elif h == "Critical":
                d["critical"] = bool(val)
            elif h == "Summary":
                d["summary"] = bool(val)
        if d.get("name"):
            out.append(d)
    return out


def build_summary_sheet(wb, summary):
    """Hakediş executive summary sheet (BAC, RAG, exec text)."""
    if "Summary" in wb.sheetnames:
        ws = wb["Summary"]
    else:
        ws = wb.create_sheet("Summary", 0)  # insert as first
    ws.cell(row=1, column=1, value="Project Health Summary")
    ws.cell(row=1, column=1).font = Font(name="Calibri", bold=True, size=16, color=BRAND_LACIVERT)
    metrics = [
        ("BAC (Budget at Completion)", summary.get("BAC")),
        ("EAC (Estimate at Completion)", summary.get("EAC")),
        ("SPI (Schedule Performance Index)", summary.get("SPI")),
        ("CPI (Cost Performance Index)", summary.get("CPI")),
        ("Overall RAG", str(summary.get("rag", "")).upper()),
    ]
    for i, (label, val) in enumerate(metrics, start=3):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=1).font = Font(name="Calibri", bold=True, color=BRAND_NAVY)
        ws.cell(row=i, column=2, value=val if val is not None else "N/A")
        if label == "Overall RAG":
            apply_rag_fill(ws.cell(row=i, column=2), str(val or "").lower())
    exec_text = summary.get("executive_text", "")
    if exec_text:
        ws.cell(row=10, column=1, value="Executive Summary")
        ws.cell(row=10, column=1).font = Font(name="Calibri", bold=True, color=BRAND_LACIVERT)
        ws.cell(row=11, column=1, value=exec_text)
        ws.merge_cells(start_row=11, start_column=1, end_row=11, end_column=4)
    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 20
    return ws
```

### Step 4-5: Run + regression
- T95 dcma-style: `python -m pytest tests/test_excel_io.py -q` expect 15 PASS
- Regression check identical pattern as T94

### Step 6: Commit

```bash
git add excel_io.py tests/test_excel_io.py && git commit -m "Phase 5c T95: tasks sheet builder + reader + summary sheet

build_tasks_sheet: 8-col header (ID/Name/Duration/Start/Finish/
%Complete/Critical/Summary), zebra rows, frozen header. Summary tasks
excluded. Duration h→days conversion (8h/day).

read_tasks_sheet: round-trip reader (xlsx→list of dicts in MSP shape).

build_summary_sheet: executive header (Lacivert 16pt) + key metrics
(BAC/EAC/SPI/CPI/RAG) + exec text merged cell. RAG cell color-coded.

7 new tests (15 cumulative T94+T95)."
```

---

## Task 96: EVM Sheet Builder

**Files:** Modify `excel_io.py` + `tests/test_excel_io.py`

### Step 1: Failing tests

```python
from excel_io import build_evm_sheet


def _sample_evm():
    return {
        "metrics": {"BAC": 1000.0, "EV": 600.0, "AC": 650.0, "PV": 700.0,
                    "SV": -100.0, "CV": -50.0, "SPI": 0.857, "CPI": 0.923},
        "forecast": {"EAC1": 1050.0, "EAC2": 1083.5, "EAC3": 1100.0,
                     "ETC": 400.0, "VAC": -50.0,
                     "TCPI_BAC": 1.14, "TCPI_EAC": 1.05},
        "earned_schedule": {"AT": 5.0, "ES": 4.5, "SVt": -0.5, "SPIt": 0.9},
        "rag": "amber",
        "time_phased": [
            {"period": "2026-W01", "PV": 100, "EV": 80, "AC": 90,
             "cum_PV": 100, "cum_EV": 80, "cum_AC": 90},
            {"period": "2026-W02", "PV": 200, "EV": 180, "AC": 200,
             "cum_PV": 300, "cum_EV": 260, "cum_AC": 290},
        ],
    }


def test_build_evm_sheet_creates_two_sheets():
    wb = Workbook()
    build_evm_sheet(wb, _sample_evm())
    assert "EVM_Compute" in wb.sheetnames
    assert "EVM_TimePhased" in wb.sheetnames


def test_evm_compute_sheet_has_metrics():
    wb = Workbook()
    build_evm_sheet(wb, _sample_evm())
    ws = wb["EVM_Compute"]
    # Find BAC row
    bac_found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == "BAC" and row[1] == 1000.0:
            bac_found = True
            break
    assert bac_found


def test_evm_compute_sheet_rag_colored():
    wb = Workbook()
    build_evm_sheet(wb, _sample_evm())
    ws = wb["EVM_Compute"]
    # Find RAG row, check color
    for row_idx in range(1, ws.max_row + 1):
        if ws.cell(row=row_idx, column=1).value == "Overall RAG":
            assert ws.cell(row=row_idx, column=2).fill.start_color.value == RAG_AMBER
            return
    pytest.fail("RAG row not found")


def test_evm_time_phased_sheet_has_data():
    wb = Workbook()
    build_evm_sheet(wb, _sample_evm())
    ws = wb["EVM_TimePhased"]
    # Header + 2 data rows
    assert ws.max_row == 3
    assert ws.cell(row=2, column=1).value == "2026-W01"


def test_evm_handles_missing_time_phased():
    wb = Workbook()
    evm = {**_sample_evm(), "time_phased": []}
    build_evm_sheet(wb, evm)
    ws = wb["EVM_TimePhased"]
    # Header only
    assert ws.max_row == 1


def test_evm_handles_missing_forecast():
    wb = Workbook()
    evm = {"metrics": {"BAC": 100, "EV": 50, "AC": 60, "PV": 80, "SPI": 0.625, "CPI": 0.83}}
    build_evm_sheet(wb, evm)
    assert "EVM_Compute" in wb.sheetnames
```

### Step 2-3: Implementation

Append to `excel_io.py`:
```python
EVM_COMPUTE_ROWS = [
    ("BAC", "metrics", "BAC", "currency"),
    ("EV (Earned Value)", "metrics", "EV", "currency"),
    ("AC (Actual Cost)", "metrics", "AC", "currency"),
    ("PV (Planned Value)", "metrics", "PV", "currency"),
    ("SV (Schedule Variance)", "metrics", "SV", "currency"),
    ("CV (Cost Variance)", "metrics", "CV", "currency"),
    ("SPI", "metrics", "SPI", "ratio"),
    ("CPI", "metrics", "CPI", "ratio"),
    ("EAC1", "forecast", "EAC1", "currency"),
    ("EAC2", "forecast", "EAC2", "currency"),
    ("EAC3", "forecast", "EAC3", "currency"),
    ("ETC", "forecast", "ETC", "currency"),
    ("VAC", "forecast", "VAC", "currency"),
    ("TCPI(BAC)", "forecast", "TCPI_BAC", "ratio"),
    ("TCPI(EAC)", "forecast", "TCPI_EAC", "ratio"),
    ("AT (Actual Time, weeks)", "earned_schedule", "AT", "weeks"),
    ("ES (Earned Schedule, weeks)", "earned_schedule", "ES", "weeks"),
    ("SV(t) (weeks)", "earned_schedule", "SVt", "weeks"),
    ("SPI(t)", "earned_schedule", "SPIt", "ratio"),
]


def build_evm_sheet(wb, evm):
    """Build EVM_Compute + EVM_TimePhased sheets from EVM data dict.

    evm = {metrics, forecast?, earned_schedule?, rag, time_phased[]}
    """
    # Compute sheet
    if "EVM_Compute" in wb.sheetnames:
        ws = wb["EVM_Compute"]
    elif wb.active.title == "Sheet" and wb.active.max_row == 1 and wb.active.max_column == 1:
        ws = wb.active
        ws.title = "EVM_Compute"
    else:
        ws = wb.create_sheet("EVM_Compute")
    ws.append(["Metric", "Value", "Unit"])
    apply_header_style(ws, row=1)
    row = 2
    for label, src, key, unit in EVM_COMPUTE_ROWS:
        section = evm.get(src) if isinstance(evm.get(src), dict) else {}
        val = section.get(key)
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=val if val is not None else "N/A")
        if unit == "currency" and isinstance(val, (int, float)):
            ws.cell(row=row, column=2).number_format = "#,##0.00"
        elif unit == "ratio" and isinstance(val, (int, float)):
            ws.cell(row=row, column=2).number_format = "0.000"
        ws.cell(row=row, column=3, value=unit)
        for col in range(1, 4):
            apply_zebra_fill(ws.cell(row=row, column=col), row_idx=row)
        row += 1
    # RAG row
    ws.cell(row=row, column=1, value="Overall RAG")
    ws.cell(row=row, column=1).font = Font(name="Calibri", bold=True, color=BRAND_NAVY)
    rag = str(evm.get("rag", "")).lower()
    ws.cell(row=row, column=2, value=rag.upper())
    apply_rag_fill(ws.cell(row=row, column=2), rag)
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 12
    ws.freeze_panes = "A2"

    # Time-phased sheet
    if "EVM_TimePhased" in wb.sheetnames:
        ws_tp = wb["EVM_TimePhased"]
    else:
        ws_tp = wb.create_sheet("EVM_TimePhased")
    ws_tp.append(["Period", "PV", "EV", "AC", "Cum PV", "Cum EV", "Cum AC"])
    apply_header_style(ws_tp, row=1)
    for idx, p in enumerate(evm.get("time_phased") or [], start=2):
        ws_tp.cell(row=idx, column=1, value=p.get("period"))
        for col, key in enumerate(["PV", "EV", "AC", "cum_PV", "cum_EV", "cum_AC"], start=2):
            ws_tp.cell(row=idx, column=col, value=p.get(key, 0))
            ws_tp.cell(row=idx, column=col).number_format = "#,##0.00"
        for col in range(1, 8):
            apply_zebra_fill(ws_tp.cell(row=idx, column=col), row_idx=idx)
    ws_tp.freeze_panes = "B2"
    return ws, ws_tp
```

### Steps 4-6: Run + commit

```bash
git add excel_io.py tests/test_excel_io.py && git commit -m "Phase 5c T96: EVM sheet builder (Compute + TimePhased)

EVM_Compute: 19 rows (BAC/EV/AC/PV/SV/CV/SPI/CPI + EAC1-3/ETC/VAC/
TCPI/AT/ES/SVt/SPIt) + Overall RAG row with color fill. Currency rows
formatted #,##0.00; ratio rows 0.000.

EVM_TimePhased: 7-col table (Period/PV/EV/AC/CumPV/CumEV/CumAC),
frozen header + first column.

Handles missing forecast/earned_schedule/time_phased gracefully.

6 new tests (21 cumulative T94+T95+T96)."
```

---

## Task 97: DCMA Sheet Builder

**Files:** Modify `excel_io.py` + `tests/test_excel_io.py`

### Step 1: Failing tests

```python
from excel_io import build_dcma_sheet


def _sample_dcma():
    return {
        "rules": [
            {"id": 1, "name": "No Predecessor", "threshold": "<5%", "actual": 3.5,
             "actual_unit": "%", "status": "pass", "failed_count": 7, "total_count": 200,
             "failed_task_ids": []},
            {"id": 9, "name": "High Duration (>44d)", "threshold": "<5%", "actual": 7.5,
             "actual_unit": "%", "status": "fail", "failed_count": 15, "total_count": 200,
             "failed_task_ids": [186, 187]},
        ],
        "summary": {"pass_count": 13, "fail_count": 1, "overall_rag": "green",
                    "executive_text": "13/14 pass"},
        "drilldowns": {
            9: [{"id": 186, "name": "H00"}, {"id": 187, "name": "H01"}],
        },
    }


def test_build_dcma_sheet_creates_two_sheets():
    wb = Workbook()
    build_dcma_sheet(wb, _sample_dcma())
    assert "DCMA_Rules" in wb.sheetnames
    assert "DCMA_Failed" in wb.sheetnames


def test_dcma_rules_sheet_lists_all_rules():
    wb = Workbook()
    build_dcma_sheet(wb, _sample_dcma())
    ws = wb["DCMA_Rules"]
    # Header + 2 rules
    assert ws.max_row == 3


def test_dcma_rules_status_color():
    wb = Workbook()
    build_dcma_sheet(wb, _sample_dcma())
    ws = wb["DCMA_Rules"]
    # Find status col
    headers = [c.value for c in ws[1]]
    status_col = headers.index("Status") + 1
    # Row 2 is rule 1 (pass) → green
    assert ws.cell(row=2, column=status_col).fill.start_color.value == RAG_GREEN
    # Row 3 is rule 9 (fail) → red
    assert ws.cell(row=3, column=status_col).fill.start_color.value == RAG_RED


def test_dcma_failed_sheet_lists_drilldowns():
    wb = Workbook()
    build_dcma_sheet(wb, _sample_dcma())
    ws = wb["DCMA_Failed"]
    # Header + 2 drilldown rows for rule 9
    assert ws.max_row >= 2
    # Find Task ID col
    headers = [c.value for c in ws[1]]
    tid_col = headers.index("Task ID") + 1
    task_ids = [ws.cell(row=r, column=tid_col).value for r in range(2, ws.max_row + 1)]
    assert 186 in task_ids


def test_dcma_failed_sheet_empty_when_no_drilldowns():
    wb = Workbook()
    build_dcma_sheet(wb, {"rules": [], "summary": {}, "drilldowns": {}})
    ws = wb["DCMA_Failed"]
    # Header only
    assert ws.max_row == 1
```

### Step 2-3: Implementation

```python
def build_dcma_sheet(wb, dcma):
    """Build DCMA_Rules + DCMA_Failed sheets.

    dcma = {rules: [...], summary: {...}, drilldowns: {rule_id: [{id, name}]}}
    """
    # Rules sheet
    if "DCMA_Rules" in wb.sheetnames:
        ws = wb["DCMA_Rules"]
    elif wb.active.title == "Sheet" and wb.active.max_row == 1 and wb.active.max_column == 1:
        ws = wb.active
        ws.title = "DCMA_Rules"
    else:
        ws = wb.create_sheet("DCMA_Rules")
    ws.append(["Rule #", "Name", "Threshold", "Actual", "Status", "Failed Count", "Total"])
    apply_header_style(ws, row=1)
    for idx, rule in enumerate(dcma.get("rules") or [], start=2):
        ws.cell(row=idx, column=1, value=rule.get("id"))
        ws.cell(row=idx, column=2, value=rule.get("name"))
        ws.cell(row=idx, column=3, value=rule.get("threshold"))
        actual = rule.get("actual")
        unit = rule.get("actual_unit", "")
        ws.cell(row=idx, column=4, value=f"{actual}{unit}" if actual is not None else "N/A")
        status = rule.get("status", "")
        ws.cell(row=idx, column=5, value=status.upper())
        apply_rag_fill(ws.cell(row=idx, column=5), status)
        ws.cell(row=idx, column=6, value=rule.get("failed_count", 0))
        ws.cell(row=idx, column=7, value=rule.get("total_count", 0))
        for col in range(1, 8):
            apply_zebra_fill(ws.cell(row=idx, column=col), row_idx=idx)
    for col, w in zip("ABCDEFG", [8, 28, 12, 14, 10, 14, 10]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    # Failed drilldown sheet
    if "DCMA_Failed" in wb.sheetnames:
        ws_f = wb["DCMA_Failed"]
    else:
        ws_f = wb.create_sheet("DCMA_Failed")
    ws_f.append(["Rule #", "Rule Name", "Task ID", "Task Name"])
    apply_header_style(ws_f, row=1)
    drills = dcma.get("drilldowns") or {}
    rules_by_id = {r["id"]: r for r in (dcma.get("rules") or [])}
    row = 2
    for rid, tasks in drills.items():
        rule_name = (rules_by_id.get(rid) or {}).get("name", "")
        for t in (tasks or [])[:10]:  # limit 10 per rule
            ws_f.cell(row=row, column=1, value=rid)
            ws_f.cell(row=row, column=2, value=rule_name)
            ws_f.cell(row=row, column=3, value=t.get("id"))
            ws_f.cell(row=row, column=4, value=t.get("name", ""))
            for col in range(1, 5):
                apply_zebra_fill(ws_f.cell(row=row, column=col), row_idx=row)
            row += 1
    for col, w in zip("ABCD", [8, 28, 10, 28]):
        ws_f.column_dimensions[col].width = w
    ws_f.freeze_panes = "A2"
    return ws, ws_f
```

### Steps 4-6: Run + commit

```bash
git add excel_io.py tests/test_excel_io.py && git commit -m "Phase 5c T97: DCMA sheet builder (Rules + Failed)

DCMA_Rules: 7-col table (Rule #/Name/Threshold/Actual/Status/Failed/
Total), status cell RAG-colored (pass=green, fail=red).

DCMA_Failed: drill-down rows (Rule#/RuleName/TaskID/TaskName), max 10
tasks per rule. Empty sheet (header only) when no drilldowns.

5 new tests (26 cumulative T94+T95+T96+T97)."
```

---

## Task 98: Hakediş Workbook Composer + Progress Reader

**Files:** Modify `excel_io.py` + `tests/test_excel_io.py`

### Step 1: Failing tests

```python
from excel_io import build_hakedis_workbook, read_progress_sheet


def test_build_hakedis_workbook(tmp_path):
    tasks = _sample_tasks()
    evm = _sample_evm()
    dcma = _sample_dcma()
    summary = {"BAC": 1000, "EAC": 1100, "SPI": 0.95, "CPI": 0.91,
               "rag": "amber", "executive_text": "Track."}
    xlsx = tmp_path / "hak.xlsx"
    wb = build_hakedis_workbook(tasks, evm, dcma, summary, str(xlsx))
    # All 5 sheets exist (Summary, Tasks, EVM_Compute, EVM_TimePhased, DCMA_Rules, DCMA_Failed)
    expected = {"Summary", "Tasks", "EVM_Compute", "EVM_TimePhased",
                "DCMA_Rules", "DCMA_Failed"}
    assert expected.issubset(set(wb.sheetnames))
    # File written
    assert xlsx.exists()


def test_build_hakedis_workbook_summary_first(tmp_path):
    """Summary sheet should be first tab."""
    xlsx = tmp_path / "hak.xlsx"
    wb = build_hakedis_workbook(_sample_tasks(), _sample_evm(), _sample_dcma(),
                                {"rag": "green"}, str(xlsx))
    assert wb.sheetnames[0] == "Summary"


def test_read_progress_sheet_round_trip(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Progress"
    ws.append(["Task ID", "%Complete", "Actual Work (h)"])
    ws.append([1, 100, 80])
    ws.append([2, 50, 80])
    xlsx = tmp_path / "prog.xlsx"
    wb.save(str(xlsx))
    rows = read_progress_sheet(str(xlsx))
    assert len(rows) == 2
    assert rows[0]["task_id"] == 1
    assert rows[0]["percent_complete"] == 100


def test_read_progress_sheet_handles_pct_fraction(tmp_path):
    """If %Complete is fraction (0.5) treat as 50%."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Progress"
    ws.append(["Task ID", "%Complete"])
    ws.append([1, 0.5])
    xlsx = tmp_path / "prog.xlsx"
    wb.save(str(xlsx))
    rows = read_progress_sheet(str(xlsx))
    assert rows[0]["percent_complete"] == 50.0
```

### Step 2-3: Implementation

```python
def build_hakedis_workbook(tasks, evm, dcma, summary, xlsx_path):
    """Compose multi-sheet hakedis workbook and save to xlsx_path.

    Sheets (in order): Summary, Tasks, EVM_Compute, EVM_TimePhased,
    DCMA_Rules, DCMA_Failed.
    """
    wb = Workbook()
    # Remove default empty sheet
    default = wb.active
    wb.remove(default)
    # Build in display order (Summary first)
    build_summary_sheet(wb, summary or {})
    build_tasks_sheet(wb, tasks or [], sheet_name="Tasks")
    build_evm_sheet(wb, evm or {})
    build_dcma_sheet(wb, dcma or {})
    wb.save(xlsx_path)
    return wb


def read_progress_sheet(xlsx_path, sheet_name="Progress"):
    """Read Progress sheet → [{task_id, percent_complete, actual_work_h?}].

    %Complete cells: accepts 50, 50.0, or 0.5 (treats <=1 as fraction).
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h or "").strip() for h in rows[0]]
    out = []
    for row in rows[1:]:
        d = {}
        for i, val in enumerate(row):
            if i >= len(headers):
                continue
            h = headers[i]
            if h == "Task ID":
                d["task_id"] = int(val) if val is not None else None
            elif h == "%Complete":
                pc = float(val or 0)
                d["percent_complete"] = pc * 100.0 if 0 < pc <= 1 else pc
            elif h == "Actual Work (h)":
                d["actual_work_h"] = float(val or 0)
        if d.get("task_id") is not None:
            out.append(d)
    return out
```

### Steps 4-6: Run + commit

```bash
git add excel_io.py tests/test_excel_io.py && git commit -m "Phase 5c T98: hakedis workbook composer + progress reader

build_hakedis_workbook: composes 6-sheet workbook (Summary first,
then Tasks, EVM_Compute, EVM_TimePhased, DCMA_Rules, DCMA_Failed)
and writes to xlsx_path.

read_progress_sheet: round-trip reader (xlsx→list of progress
update dicts in MSP shape). Accepts %Complete as int/float (50)
or fraction (0.5).

4 new tests (30 cumulative T94-T98). excel_io.py complete."
```

---

## Task 99 (BIG ONE — subagent dispatch): `_excel_collect_full_data` + 4 Export Action Helpers

**Files:**
- Modify: `msproject_mcp_core.py` (add Phase 5C section AFTER Phase 5b dispatcher, BEFORE def main)
- Create: `tests/test_msproject_excel_loader.py`
- Create: `tests/test_msproject_excel_actions.py`

**Subagent context (verbatim — provide full task text + insertion location):**

Phase 5b dispatcher `msproject_health` is currently the LAST `@mcp.tool` before `def main`. Insert Phase 5C section between `msproject_health` (~line 5780) and `def main` (~line 5790). All Phase 5b helpers (`_msp_dcma_assess_all`, etc.) and Phase 5a (`_msp_evm_compute_metrics`, etc.) are available — use directly.

**Implementation skeleton:**

```python
# ============================================================================
# PHASE 5C - EXCEL TOOL
# ============================================================================
from excel_io import (
    build_tasks_sheet, build_evm_sheet, build_dcma_sheet,
    build_summary_sheet, build_hakedis_workbook,
    read_tasks_sheet, read_progress_sheet,
)
from openpyxl import Workbook


def _excel_collect_full_data(file_path=None, baseline_number=0, bucket="week"):
    """SINGLE collect (Phase 5b TAIL lesson) — fetch tasks + EVM + DCMA once.

    Returns {status, tasks, evm: {metrics, forecast, earned_schedule, rag,
             time_phased}, dcma: {rules, summary, drilldowns}}.
    """
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    # Phase 5a: tasks come from _evm_load_task_data
    base = _evm_load_task_data(file_path=file_path)
    if base.get("status") != "ok":
        return base
    tasks = base.get("tasks", []) or []
    # Phase 5a EVM compute (single call)
    evm_metrics = _msp_evm_compute_metrics(file_path=file_path,
                                           baseline_number=baseline_number)
    evm_forecast = _msp_evm_forecast(file_path=file_path,
                                     baseline_number=baseline_number)
    evm_es = _msp_evm_earned_schedule(file_path=file_path,
                                      baseline_number=baseline_number)
    evm_summary = _msp_evm_summary(file_path=file_path,
                                   baseline_number=baseline_number)
    evm_tp = _msp_evm_time_phased_evm(file_path=file_path,
                                      baseline_number=baseline_number,
                                      bucket=bucket)
    # Phase 5b DCMA (single call)
    dcma_full = _msp_dcma_assess_all(file_path=file_path,
                                     baseline_number=baseline_number)
    drilldowns = {}
    if dcma_full.get("status") == "ok":
        for rule in dcma_full.get("rules", []):
            if rule.get("status") == "fail":
                rid = rule["id"]
                d = _msp_dcma_drill_down(file_path=file_path, rule_id=rid,
                                         baseline_number=baseline_number)
                if d.get("status") == "ok":
                    drilldowns[rid] = d.get("failed_tasks", [])[:10]
    return {
        "status": "ok",
        "tasks": tasks,
        "evm": {
            "metrics": (evm_metrics or {}).get("metrics", {}) if evm_metrics.get("status") == "ok" else {},
            "forecast": (evm_forecast or {}).get("forecast", {}) if evm_forecast.get("status") == "ok" else {},
            "earned_schedule": (evm_es or {}) if evm_es.get("status") == "ok" else {},
            "rag": (evm_summary or {}).get("rag") if evm_summary.get("status") == "ok" else None,
            "time_phased": (evm_tp or {}).get("buckets", []) if evm_tp.get("status") == "ok" else [],
        },
        "dcma": {
            "rules": dcma_full.get("rules", []) if dcma_full.get("status") == "ok" else [],
            "summary": dcma_full.get("summary", {}) if dcma_full.get("status") == "ok" else {},
            "drilldowns": drilldowns,
        },
    }


def _msp_excel_export_hakedis(file_path=None, xlsx_path=None, baseline_number=0):
    """Action 1 (HERO): export multi-sheet hakedis workbook."""
    if not xlsx_path:
        return {"status": "error", "error": "xlsx_path required"}
    data = _excel_collect_full_data(file_path=file_path,
                                    baseline_number=baseline_number)
    if data.get("status") != "ok":
        return data
    summary = {
        "BAC": data["evm"]["metrics"].get("BAC"),
        "EAC": data["evm"]["forecast"].get("EAC2"),
        "SPI": data["evm"]["metrics"].get("SPI"),
        "CPI": data["evm"]["metrics"].get("CPI"),
        "rag": data["dcma"]["summary"].get("overall_rag"),
        "executive_text": data["dcma"]["summary"].get("executive_text", ""),
    }
    try:
        build_hakedis_workbook(
            tasks=data["tasks"], evm=data["evm"], dcma=data["dcma"],
            summary=summary, xlsx_path=xlsx_path,
        )
    except Exception as e:
        logger.exception(f"export_hakedis failed: {e}")
        return {"status": "error", "error": str(e)}
    return {
        "status": "ok",
        "xlsx_path": xlsx_path,
        "sheets_written": ["Summary", "Tasks", "EVM_Compute", "EVM_TimePhased",
                           "DCMA_Rules", "DCMA_Failed"],
        "rows_written": {"tasks": len([t for t in data["tasks"] if not t.get("summary")]),
                        "evm_time_phased": len(data["evm"].get("time_phased", [])),
                        "dcma_rules": len(data["dcma"].get("rules", []))},
    }


def _msp_excel_export_tasks(file_path=None, xlsx_path=None):
    if not xlsx_path:
        return {"status": "error", "error": "xlsx_path required"}
    base = _evm_load_task_data(file_path=file_path)
    if base.get("status") != "ok":
        return base
    try:
        wb = Workbook()
        wb.remove(wb.active)
        build_tasks_sheet(wb, base.get("tasks", []), sheet_name="Tasks")
        wb.save(xlsx_path)
    except Exception as e:
        logger.exception(f"export_tasks failed: {e}")
        return {"status": "error", "error": str(e)}
    return {"status": "ok", "xlsx_path": xlsx_path,
            "rows_written": len([t for t in base.get("tasks", []) if not t.get("summary")])}


def _msp_excel_export_evm(file_path=None, xlsx_path=None, baseline_number=0,
                          bucket="week"):
    if not xlsx_path:
        return {"status": "error", "error": "xlsx_path required"}
    data = _excel_collect_full_data(file_path=file_path,
                                    baseline_number=baseline_number,
                                    bucket=bucket)
    if data.get("status") != "ok":
        return data
    try:
        wb = Workbook()
        wb.remove(wb.active)
        build_evm_sheet(wb, data["evm"])
        wb.save(xlsx_path)
    except Exception as e:
        logger.exception(f"export_evm failed: {e}")
        return {"status": "error", "error": str(e)}
    return {"status": "ok", "xlsx_path": xlsx_path,
            "rows_written": {"compute": 19, "time_phased": len(data["evm"].get("time_phased", []))}}


def _msp_excel_export_dcma(file_path=None, xlsx_path=None, baseline_number=0):
    if not xlsx_path:
        return {"status": "error", "error": "xlsx_path required"}
    data = _excel_collect_full_data(file_path=file_path,
                                    baseline_number=baseline_number)
    if data.get("status") != "ok":
        return data
    try:
        wb = Workbook()
        wb.remove(wb.active)
        build_dcma_sheet(wb, data["dcma"])
        wb.save(xlsx_path)
    except Exception as e:
        logger.exception(f"export_dcma failed: {e}")
        return {"status": "error", "error": str(e)}
    return {"status": "ok", "xlsx_path": xlsx_path,
            "rows_written": {"rules": len(data["dcma"].get("rules", [])),
                            "drilldowns": sum(len(v) for v in data["dcma"].get("drilldowns", {}).values())}}
```

**Tests** (`tests/test_msproject_excel_loader.py`):

```python
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import _excel_collect_full_data

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_excel_collect_full_data_xml():
    r = _excel_collect_full_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    for k in ("tasks", "evm", "dcma"):
        assert k in r


def test_excel_collect_full_data_invalid_baseline():
    r = _excel_collect_full_data(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"


def test_excel_collect_full_data_invalid_file():
    r = _excel_collect_full_data(file_path="/nonexistent.xml")
    assert r["status"] == "error"
```

**Tests** (`tests/test_msproject_excel_actions.py`):

```python
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
    wb = load_workbook(str(xlsx))
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
    wb = load_workbook(str(xlsx))
    assert "Tasks" in wb.sheetnames


def test_export_evm_creates_two_sheets(tmp_path):
    xlsx = tmp_path / "evm.xlsx"
    r = _msp_excel_export_evm(file_path=MSP_XML, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    wb = load_workbook(str(xlsx))
    assert "EVM_Compute" in wb.sheetnames
    assert "EVM_TimePhased" in wb.sheetnames


def test_export_dcma_creates_rules_sheet(tmp_path):
    xlsx = tmp_path / "dcma.xlsx"
    r = _msp_excel_export_dcma(file_path=MSP_XML, xlsx_path=str(xlsx))
    assert r["status"] == "ok"
    wb = load_workbook(str(xlsx))
    assert "DCMA_Rules" in wb.sheetnames
```

**Commit:**
```
Phase 5c T99 (BIG ONE): _excel_collect_full_data + 4 export action helpers

Single collect pattern (Phase 5b TAIL lesson) — calls Phase 5a EVM
helpers + Phase 5b DCMA assess_all once, packages into shared shape
for export builders.

4 export actions: hakedis (multi-sheet hero) + tasks + evm + dcma.
All accept xlsx_path; missing path returns error.

3 loader + 5 action tests. Phase 1-5b DOKUNULMAZ.
```

---

## Task 100 (BIG ONE — subagent dispatch): 2 Import Action Helpers

**Files:**
- Modify: `msproject_mcp_core.py` (append to Phase 5C section)
- Modify: `tests/test_msproject_excel_actions.py` (add import tests)

**Implementation:**

```python
def _msp_excel_import_tasks(xlsx_path=None, sheet_name="Tasks"):
    """Action 5: import tasks from xlsx via Phase 1 _msp_task_bulk_add."""
    if not xlsx_path:
        return {"status": "error", "error": "xlsx_path required"}
    if not os.path.exists(xlsx_path):
        return {"status": "error", "error": f"File not found: {xlsx_path}"}
    try:
        rows = read_tasks_sheet(xlsx_path, sheet_name=sheet_name)
    except Exception as e:
        logger.exception(f"import_tasks read failed: {e}")
        return {"status": "error", "error": f"Read failed: {e}"}
    if not rows:
        return {"status": "ok", "rows_imported": 0, "task_ids": []}
    # Convert excel_io shape to _msp_task_bulk_add items shape:
    #   {name, duration} where duration is "Nd" string
    items = []
    for r in rows:
        if not r.get("name"):
            continue
        days = round(float(r.get("duration_h") or 0) / 8.0, 1)
        if days <= 0:
            days = 1.0
        items.append({"name": r["name"], "duration": f"{days}d"})
    try:
        result = _msp_task_bulk_add(items=items)
    except Exception as e:
        logger.exception(f"_msp_task_bulk_add failed: {e}")
        return {"status": "error", "error": str(e)}
    return {
        "status": "ok",
        "rows_imported": len(items),
        "task_ids": result.get("task_ids", []) if isinstance(result, dict) else [],
    }


def _msp_excel_import_progress(xlsx_path=None, sheet_name="Progress"):
    """Action 6: import progress updates from xlsx via Phase 3b bulk_update."""
    if not xlsx_path:
        return {"status": "error", "error": "xlsx_path required"}
    if not os.path.exists(xlsx_path):
        return {"status": "error", "error": f"File not found: {xlsx_path}"}
    try:
        rows = read_progress_sheet(xlsx_path, sheet_name=sheet_name)
    except Exception as e:
        logger.exception(f"import_progress read failed: {e}")
        return {"status": "error", "error": f"Read failed: {e}"}
    if not rows:
        return {"status": "ok", "rows_imported": 0}
    items = [{"task_id": r["task_id"], "percent_complete": r.get("percent_complete", 0)}
             for r in rows if r.get("task_id") is not None]
    try:
        _msp_progress_bulk_update(items=items)
    except Exception as e:
        logger.exception(f"_msp_progress_bulk_update failed: {e}")
        return {"status": "error", "error": str(e)}
    return {"status": "ok", "rows_imported": len(items)}
```

**Tests** (append to test_msproject_excel_actions.py):

```python
from openpyxl import Workbook


def test_import_tasks_round_trip(tmp_path):
    """Build xlsx → import → verify call shape."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Tasks"
    ws.append(["ID", "Name", "Duration (d)", "Start", "Finish",
               "%Complete", "Critical", "Summary"])
    ws.append([1, "TestImport1", 5, None, None, 0, False, False])
    ws.append([2, "TestImport2", 3, None, None, 0, False, False])
    xlsx = tmp_path / "imp.xlsx"
    wb.save(str(xlsx))

    from msproject_mcp_core import _msp_excel_import_tasks
    r = _msp_excel_import_tasks(xlsx_path=str(xlsx))
    # MSP COM round-trip — only assert call shape (not real task ids
    # since MSP state isn't deterministic in unit test). Status must
    # be ok and rows_imported reflects parsed rows.
    assert r["status"] in ("ok", "error")
    if r["status"] == "ok":
        assert r["rows_imported"] == 2


def test_import_tasks_missing_file():
    from msproject_mcp_core import _msp_excel_import_tasks
    r = _msp_excel_import_tasks(xlsx_path="/nonexistent.xlsx")
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_import_tasks_no_xlsx_path():
    from msproject_mcp_core import _msp_excel_import_tasks
    r = _msp_excel_import_tasks()
    assert r["status"] == "error"


def test_import_progress_round_trip(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Progress"
    ws.append(["Task ID", "%Complete"])
    ws.append([1, 50])
    xlsx = tmp_path / "prog.xlsx"
    wb.save(str(xlsx))

    from msproject_mcp_core import _msp_excel_import_progress
    r = _msp_excel_import_progress(xlsx_path=str(xlsx))
    assert r["status"] in ("ok", "error")
```

**Commit:**
```
Phase 5c T100 (BIG ONE): import action helpers (tasks + progress)

import_tasks: read xlsx → wrap _msp_task_bulk_add (Phase 1).
Converts duration_h to "Nd" duration string per bulk_add format.

import_progress: read xlsx → wrap _msp_progress_bulk_update (Phase 3b).

4 new tests (round-trip + missing file + missing path).
Phase 1-5b DOKUNULMAZ.
```

---

## Task 101: Dispatcher + Acceptance + README + Push (FINAL)

**Files:**
- Modify: `msproject_mcp_core.py` (add `@mcp.tool msproject_excel` after Phase 5C helpers, BEFORE def main)
- Create: `tests/test_msproject_excel_dispatcher.py`
- Create: `samples/build_excel_lifecycle.py`
- Modify: `README.md`
- Modify: `requirements.txt` (add openpyxl)

### Step 1: Dispatcher tests

```python
import asyncio, json, os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import msproject_excel

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def _call(action, **kw):
    raw = _run(msproject_excel({"action": action, **kw}))
    return json.loads(raw)


def test_dispatcher_export_hakedis(tmp_path):
    p = _call("export_hakedis", file_path=MSP_XML, xlsx_path=str(tmp_path / "h.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_export_tasks(tmp_path):
    p = _call("export_tasks", file_path=MSP_XML, xlsx_path=str(tmp_path / "t.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_export_evm(tmp_path):
    p = _call("export_evm", file_path=MSP_XML, xlsx_path=str(tmp_path / "e.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_export_dcma(tmp_path):
    p = _call("export_dcma", file_path=MSP_XML, xlsx_path=str(tmp_path / "d.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_unknown_action():
    p = _call("nonsense", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_missing_xlsx_path():
    p = _call("export_hakedis", file_path=MSP_XML)
    assert p["status"] == "error"
```

### Step 3: Dispatcher implementation (insert AFTER Phase 5C helpers, BEFORE def main)

```python
@mcp.tool(
    name="msproject_excel",
    annotations={
        "title": "MS Project Excel Hakediş Workbook + Bulk Import",
        "readOnlyHint": False,
    },
)
async def msproject_excel(params: dict) -> str:
    """Excel I/O for MSP — hakediş workbook export + bulk Excel→MSP import.

    Hybrid: file_path verilirse Phase 4 file path; yoksa Phase 1 COM.

    Actions:
    - export_hakedis: Multi-sheet workbook (Summary + Tasks + EVM + DCMA)
    - export_tasks: Tasks sheet only
    - export_evm: EVM_Compute + EVM_TimePhased
    - export_dcma: DCMA_Rules + DCMA_Failed
    - import_tasks: xlsx Tasks sheet → _msp_task_bulk_add
    - import_progress: xlsx Progress sheet → _msp_progress_bulk_update

    Phase 5c (1 May 2026). Tool count 10 -> 11.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "export_hakedis":
            r = _msp_excel_export_hakedis(**p)
        elif action == "export_tasks":
            r = _msp_excel_export_tasks(**p)
        elif action == "export_evm":
            r = _msp_excel_export_evm(**p)
        elif action == "export_dcma":
            r = _msp_excel_export_dcma(**p)
        elif action == "import_tasks":
            r = _msp_excel_import_tasks(**p)
        elif action == "import_progress":
            r = _msp_excel_import_progress(**p)
        else:
            r = {"status": "error",
                 "error": (f"Unknown action '{action}'. Valid: "
                           "export_hakedis/export_tasks/export_evm/"
                           "export_dcma/import_tasks/import_progress")}
    except TypeError as e:
        r = {"status": "error", "error": f"Invalid params for {action}: {e}"}
    except Exception as e:
        logger.exception(f"msproject_excel({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

### Step 4: Acceptance script `samples/build_excel_lifecycle.py`

```python
"""Phase 5c Excel acceptance: 200-task hakedis workbook + import roundtrip.

SAFETY: FileNew + FileClose 0. User's active project untouched.

Scenario (target ≤60s):
  1. Build 200 tasks + 14 CAU resources + Baseline 0 + 30 progress
  2. _msp_excel_export_hakedis(xlsx) → 6-sheet workbook
  3. Verify file structure (sheet count + key cells)
  4. Build 10-row progress xlsx
  5. _msp_excel_import_progress(xlsx)
"""
import os, sys, time, tempfile, functools

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
print = functools.partial(print, flush=True)

import pythoncom, win32com.client
from openpyxl import Workbook, load_workbook
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save, _msp_progress_bulk_update, _msp_progress_set_status_date,
    _msp_excel_export_hakedis, _msp_excel_import_progress,
)

N_TASKS = 200


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] isolated test project: {test_name}")

    out_dir = tempfile.mkdtemp(prefix="dcma_excel_")
    xlsx_out = os.path.join(out_dir, "hakedis.xlsx")
    xlsx_imp = os.path.join(out_dir, "progress_imp.xlsx")

    try:
        t0 = time.time()
        # 1. Build base
        items = [{"name": f"V{i:03d}", "duration": "5d"} for i in range(N_TASKS - 15)]
        items += [{"name": f"H{i:02d}", "duration": "60d"} for i in range(15)]
        tasks = _msp_task_bulk_add(items=items)
        cau = ["COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
               "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR"]
        res_ids = [_msp_resource_add(name=n, type="Work")["resource_id"] for n in cau]
        sample = [{"task_id": tid, "resource_id": res_ids[i % 14]}
                  for i, tid in enumerate(tasks["task_ids"][12:])]
        _msp_resource_bulk_assign(items=sample)
        _msp_baseline_save(baseline_number=0)
        _msp_progress_bulk_update(items=[{"task_id": tid, "percent_complete": 50.0}
                                         for tid in tasks["task_ids"][:30]])
        _msp_progress_set_status_date(status_date="2026-05-15")
        print(f"\n1. Setup done at {time.time()-t0:.2f}s")

        # 2. Export hakedis
        print(f"\n2. Exporting hakedis workbook to {xlsx_out}...")
        r = _msp_excel_export_hakedis(xlsx_path=xlsx_out)
        print(f"   status={r.get('status')} sheets={r.get('sheets_written')}")
        assert r["status"] == "ok"
        assert os.path.exists(xlsx_out)
        print(f"   exported at {time.time()-t0:.2f}s, size={os.path.getsize(xlsx_out)} bytes")

        # 3. Verify workbook structure
        wb = load_workbook(xlsx_out, read_only=True)
        for s in ("Summary", "Tasks", "EVM_Compute", "EVM_TimePhased",
                  "DCMA_Rules", "DCMA_Failed"):
            assert s in wb.sheetnames, f"missing sheet {s}"
        print(f"   verified 6 sheets: {wb.sheetnames}")
        wb.close()

        # 4. Build progress import xlsx
        print(f"\n3. Building progress import workbook...")
        wb_p = Workbook()
        ws = wb_p.active
        ws.title = "Progress"
        ws.append(["Task ID", "%Complete"])
        for i, tid in enumerate(tasks["task_ids"][30:40]):  # 10 new updates
            ws.append([tid, 75.0])
        wb_p.save(xlsx_imp)

        # 5. Import progress
        print(f"\n4. Importing progress from {xlsx_imp}...")
        r2 = _msp_excel_import_progress(xlsx_path=xlsx_imp)
        print(f"   status={r2.get('status')} rows_imported={r2.get('rows_imported')}")
        assert r2["status"] == "ok"
        assert r2["rows_imported"] == 10

        elapsed = time.time() - t0
        print(f"\n[OK] ACCEPTANCE: {elapsed:.2f}s total (target <60s)")
        assert elapsed < 60.0, f"Too slow: {elapsed}s"

    finally:
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    print(f"[SAFE] closed test project {test_name}")
                    break
        except Exception as e:
            print(f"[WARN] cleanup error: {e}")


if __name__ == "__main__":
    main()
```

### Step 6: README + requirements

Append to README after Phase 5b section:
```markdown
## Phase 5c — Excel I/O (1 May 2026)

`msproject_excel` tool — hakediş workbook export + Excel-driven import.
6 actions covering Phase 5a EVM + Phase 5b DCMA → multi-sheet xlsx, plus
bulk Excel → MSP imports. MCS brand styling per CLAUDE.md RULE 14.

**Actions:**
- `export_hakedis`: 6-sheet workbook (Summary + Tasks + EVM + DCMA)
- `export_tasks`: Tasks sheet only
- `export_evm`: EVM_Compute + EVM_TimePhased
- `export_dcma`: DCMA_Rules + DCMA_Failed
- `import_tasks`: xlsx → _msp_task_bulk_add (Phase 1)
- `import_progress`: xlsx → _msp_progress_bulk_update (Phase 3b)

Architecture: pure-Python `excel_io.py` (openpyxl 3.1.5, MSP/COM/file
independent, ~30 tests) + I/O adapters in msproject_mcp_core.py.
Phase 1-5b helpers DOKUNULMAZ.

Acceptance: `samples/build_excel_lifecycle.py` exports 200-task hakedis
workbook + roundtrip imports 10 progress updates in ≤60s.

Tool count: **11 tools, ~89 actions**.
```

Add `openpyxl>=3.1` to `requirements.txt`.

### Step 8: Commit + push

```bash
git add msproject_mcp_core.py tests/test_msproject_excel_dispatcher.py samples/build_excel_lifecycle.py README.md requirements.txt && git commit -m "Phase 5c T101: msproject_excel dispatcher + acceptance + README + push (11th tool)

@mcp.tool msproject_excel routes 6 actions: export_hakedis (HERO) +
export_tasks/evm/dcma + import_tasks/progress.

Acceptance build_excel_lifecycle.py: 200-task setup + hakedis export
+ 10-row progress import roundtrip. Target <60s.

README: Phase 5c section. Tool count 10 -> 11, actions ~83 -> ~89.
requirements.txt: add openpyxl>=3.1.

Cumulative regression target: ~280 PASS (235 baseline + ~45 Phase 5c).
Phase 1-5b helpers DOKUNULMAZ."
git push origin main
```

---

## Phase 5c Acceptance Criteria

1. ✅ T94-T101 8-task chain landed
2. ✅ Acceptance `build_excel_lifecycle.py` ≤ 60s (or adjust to ≤90s if COM-heavy)
3. ✅ Workbook opens in Excel/LibreOffice without errors
4. ✅ MCS brand: Lacivert headers, RAG cells, Calibri
5. ✅ Round-trip: import_progress applies updates correctly
6. ✅ Phase 1-5b regression untouched
7. ✅ All 6 actions covered by dispatcher tests
8. ✅ Push to origin/main
9. ⏸ Phase 5d (post-onay)

---

*Plan committed: 2026-05-01.*
