# Phase 5c — `msproject_excel` Design (1 May 2026)

## Goal

11th MCP tool: `msproject_excel`. Multi-sheet hakediş workbook export + Excel-driven task/progress import. Bridges MSP project state (Phase 1-5b read pipelines) to Excel — the construction industry's reporting lingua franca and the user's daily deliverable format.

## Background — Why now

Phase 5a EVM and Phase 5b DCMA produce structured rule/metric data, but consumption today is Claude chat output. Construction stakeholders (project owners, contractors, hakediş approvers) consume Excel. Ship Excel I/O to close the report-loop:

- **Export:** EVM metrics + DCMA assessment + task list → multi-sheet xlsx ready for hakediş submission.
- **Import:** Excel task templates + progress updates → MSP via existing `_msp_task_bulk_add` / `_msp_progress_bulk_update`.

## Brainstorming decisions (1 May 2026 session)

- **Q1 Scope:** B+ (6-action) — Hero `export_hakediş` + 3 granular exports + 2 imports. Round-trip enabled, scope manageable.
- **Q2 Template style:** MCS brand (Lacivert `#0B1F4D` + Calibri, RAG color cells) per CLAUDE.md RULE 14 — header bandlar + zebra rows. RULE 13 rakip yasağı: "Industry Standard" ifadeleri kullanılacak.
- **Q3 Import strategy:** Excel-only → MSP. Roundtrip xlsx ↔ msp file deferred to Phase 5d (would require xml/mpp write integration).
- **Library:** `openpyxl 3.1.5` (already installed). Pure Python, no Excel install. mspdi_parser pattern.
- **Architecture:** Yaklaşım C — `excel_io.py` pure module + I/O adapters in core + dispatcher.

## Tool surface — `msproject_excel` 6 actions

| # | Action | Inputs | Output sheet(s) |
|---|---|---|---|
| 1 | `export_hakediş` 🚀 HERO | `file_path`, `xlsx_path`, `baseline_number=0`, `status_date?` | `Tasks` + `EVM` + `DCMA` + `Summary` |
| 2 | `export_tasks` | `file_path`, `xlsx_path`, `filters?` | `Tasks` |
| 3 | `export_evm` | `file_path`, `xlsx_path`, `baseline_number=0`, `bucket="week"` | `EVM_Compute` + `EVM_TimePhased` |
| 4 | `export_dcma` | `file_path`, `xlsx_path`, `baseline_number=0` | `DCMA_Rules` + `DCMA_Failed` |
| 5 | `import_tasks` | `xlsx_path`, `sheet_name="Tasks"` | wraps `_msp_task_bulk_add` |
| 6 | `import_progress` | `xlsx_path`, `sheet_name="Progress"` | wraps `_msp_progress_bulk_update` |

All export actions accept optional `file_path` (Phase 4 file path) or fall back to Phase 1 COM (active project). All return `{"status": "ok"|"error", "xlsx_path", "rows_written" | "rows_imported", ...}`.

## Architecture

```
excel_io.py (NEW, pure Python, ~400 lines, no MSP/COM dependency)
├── Constants: BRAND_LACIVERT="0B1F4D", BRAND_NAVY="3D4663", RAG_GREEN/AMBER/RED
├── Helpers: _header_style(), _rag_fill(status), _zebra_fill(row_idx)
├── Sheet builders (take dicts, return openpyxl.Workbook):
│   ├── build_tasks_sheet(wb, tasks, sheet_name="Tasks")
│   ├── build_evm_sheet(wb, metrics, time_phased, sheet_name="EVM")
│   ├── build_dcma_sheet(wb, rules, failed_drilldowns, sheet_name="DCMA")
│   └── build_summary_sheet(wb, summary)  # exec text + RAG + key counts
├── Workbook composers:
│   └── build_hakediş_workbook(tasks, evm, dcma, summary, xlsx_path)
└── Readers (Excel → list of dicts, MSP-shape):
    ├── read_tasks_sheet(xlsx_path, sheet_name) → [{name, duration, ...}]
    └── read_progress_sheet(xlsx_path, sheet_name) → [{task_id, percent_complete, ...}]

msproject_mcp_core.py PHASE 5C SECTION (after Phase 5b dispatcher, before def main)
├── from excel_io import build_*, read_*
├── _excel_collect_full_data(file_path, baseline_number, bucket)
│     → {tasks, evm_metrics, evm_time_phased, dcma_assess, summary}
│     SINGLE COLLECT (Phase 5b TAIL lesson — avoid N+1 fetches)
├── _msp_excel_export_hakediş(file_path, xlsx_path, baseline_number=0)
├── _msp_excel_export_tasks(file_path, xlsx_path, filters=None)
├── _msp_excel_export_evm(file_path, xlsx_path, baseline_number=0, bucket="week")
├── _msp_excel_export_dcma(file_path, xlsx_path, baseline_number=0)
├── _msp_excel_import_tasks(xlsx_path, sheet_name="Tasks")
├── _msp_excel_import_progress(xlsx_path, sheet_name="Progress")
└── @mcp.tool msproject_excel dispatcher (6 actions)
```

**Phase 1+2+3+4+5a+5b helpers DOKUNULMAZ.** Read-only calls into Phase 5a `_msp_evm_*` (compute_metrics, time_phased_evm), Phase 5b `_msp_dcma_assess_all`, Phase 4 `_msp_file_read_tasks`. Imports use Phase 1 `_msp_task_bulk_add` / Phase 3b `_msp_progress_bulk_update`.

## Data flow

### Export (typical — `export_hakediş`)

```
file_path / COM
    ↓
_excel_collect_full_data — SINGLE collect via existing helpers:
    tasks = _msp_evm.compute_metrics() + Phase 4 read_tasks
    evm = _msp_evm_compute_metrics + _msp_evm_time_phased_evm + _msp_evm_summary
    dcma = _msp_dcma_assess_all
    summary = {pass_count, fail_count, rag, exec_text, BAC, EAC, SPI, CPI}
    ↓
build_hakediş_workbook(...)
    ↓ (openpyxl in-memory)
wb.save(xlsx_path)
    ↓
return {"status": "ok", "xlsx_path", "sheets_written": [...], "rows_written": {...}}
```

### Import (`import_tasks`)

```
xlsx_path
    ↓
read_tasks_sheet(xlsx_path, "Tasks") → [{name, duration, ...}]
    ↓
_msp_task_bulk_add(items=...) — Phase 1 helper, MSP COM
    ↓
return {"status": "ok", "rows_imported": N, "task_ids": [...]}
```

## Sheet templates

### Tasks sheet
| Col A | Col B | Col C | Col D | Col E | Col F | Col G | Col H |
|---|---|---|---|---|---|---|---|
| ID | Name | Duration (d) | Start | Finish | %Complete | Critical | Resources |

Header: Lacivert bg + white bold Calibri 11pt + frozen row 1.
Zebra: alternating Light Gray `#F0F3F8`.
Critical column: red text if True.
%Complete: number_format "0%".
Date columns: number_format "yyyy-mm-dd".

### EVM sheet
**EVM_Compute (top, rows 1-12):**
- Headers: Metric / Value / Unit
- Rows: BAC, EV, AC, PV, SV, CV, SPI, CPI, EAC1, EAC2, EAC3, ETC, VAC, TCPI(BAC), TCPI(EAC), RAG
- RAG cell: green/amber/red fill per status

**EVM_TimePhased (separate sheet, rows 1-N):**
| Period | PV | EV | AC | Cumulative PV | Cumulative EV | Cumulative AC |

### DCMA sheet
**DCMA_Rules:**
| Rule # | Name | Threshold | Actual | Status | Failed Count |
RAG fill: pass=green, fail=red. 14 rows.

**DCMA_Failed:**
| Rule # | Rule Name | Task ID | Task Name |
Concatenated drill-down for ALL failed rules (limit 10 per rule).

### Summary sheet (hakediş executive)
- A1: "Project Health Summary" Calibri Bold 16 Lacivert
- A3-B6: Key metrics (BAC, SPI, CPI, RAG)
- A8-A20: Executive text (DCMA pass/fail counts, missed tasks, BEI)

## Cell formatting (MCS brand)

```python
LACIVERT = "FF0B1F4D"  # ARGB
NAVY = "FF3D4663"
TURQUOISE = "FF39B4CC"
LABEL_GRAY = "FF6B7394"
ZEBRA = "FFF0F3F8"
RAG_GREEN = "FF8FBC8F"
RAG_AMBER = "FFFFCC66"
RAG_RED = "FFE57373"
```

Header style: `Font(name="Calibri", bold=True, color="FFFFFFFF", size=11)` + `PatternFill(start_color=LACIVERT, end_color=LACIVERT, fill_type="solid")`.

## Error handling

- Unknown `xlsx_path` directory → `{"status": "error", "error": "Directory not found: ..."}`
- Existing xlsx file → overwrite without prompt (caller's responsibility — match Phase 5a snapshot pattern)
- COM unavailable + no file_path → `{"status": "error", "error": "Active MS Project session required..."}` (via `_validate_active_project`)
- Import: row missing required column → `{"status": "error", "error": "Sheet '...' missing required column 'name' at row N"}`
- Empty sheet → `rows_imported=0`, status="ok" (warn-but-pass per Phase 5a vacuous pattern)

## Testing

### Unit tests — `tests/test_excel_io.py` (~30 tests, fixture-free where possible)
- Cell formatters: `_header_style`, `_rag_fill("pass")` returns green
- Sheet builders: build_tasks_sheet from synthetic dict → check sheet contents
- Workbook composers: build_hakediş_workbook from synthetic data → load via openpyxl + assert structure
- Readers: write a fixture xlsx in conftest, read it back, assert rows

### Loader tests — `tests/test_msproject_excel_loader.py` (~5 tests)
- `_excel_collect_full_data` from `sample_msp.xml` → status=ok, all keys present
- baseline_number invalid → error
- file_path invalid → error

### Action tests — `tests/test_msproject_excel_actions.py` (~6-8 tests)
- export_hakediş from sample_msp.xml → file exists, has 4 sheets
- export_tasks → 1 sheet "Tasks", header row Lacivert
- import_tasks from a fixture xlsx → returns task_ids list (use `tmp_path` fixture)
- import_progress similar

### Dispatcher tests — `tests/test_msproject_excel_dispatcher.py` (~5 tests)
- 6 actions valid + unknown action error

### Acceptance — `samples/build_excel_lifecycle.py`
**Scenario:**
1. Build 200-task CAU + 14 resources + Baseline 0 + 30 progress (Phase 5a/5b script reuse)
2. `_msp_excel_export_hakediş(xlsx_path=temp.xlsx)` → multi-sheet workbook
3. Open workbook, verify sheet count + key cells (BAC, RAG color, DCMA pass count)
4. Build a smaller xlsx with 10 progress updates (sheet "Progress")
5. `_msp_excel_import_progress(xlsx_path)` → re-applies updates
6. Total time ≤ 60s target.

**SAFETY pattern:** FileNew + FileClose 0 — Phase 5a/5b convention.

## Sequencing — T94-T101 (8 task chain)

| Task | Content | Pattern | Approx tests |
|---|---|---|---|
| T94 | `excel_io.py` foundations: constants, _header_style, _rag_fill, _zebra_fill | Manuel saf math | 6-8 |
| T95 | build_tasks_sheet + build_summary_sheet + read_tasks_sheet | Manuel | 6-8 |
| T96 | build_evm_sheet (Compute + TimePhased) + build_evm_workbook | Manuel | 6 |
| T97 | build_dcma_sheet (Rules + Failed) | Manuel | 5 |
| T98 | build_hakediş_workbook composer + read_progress_sheet | Manuel | 4-5 |
| T99 | `_excel_collect_full_data` + 4 export action helpers | **Subagent BIG ONE** | 5 loader + 5 actions |
| T100 | 2 import action helpers + dispatcher tests prep | **Subagent BIG ONE** | 4 actions + 5 dispatcher |
| T101 | `@mcp.tool msproject_excel` dispatcher + acceptance + README + push | Manuel finalize | 5 dispatcher tests |

**Estimated chain:** ~40-50 new tests. Cumulative regression target: 235 + ~45 = ~280 PASS.

## Acceptance criteria

1. ✅ T94-T101 9-12 commit chain landed
2. ✅ Acceptance `samples/build_excel_lifecycle.py` ≤ 60s @ 200 tasks
3. ✅ Multi-sheet hakediş.xlsx opens in Excel/LibreOffice without errors
4. ✅ Cell formatting matches MCS brand (Lacivert headers, RAG color)
5. ✅ Round-trip: import_tasks from exported workbook re-creates project
6. ✅ Phase 1-5b regression untouched
7. ✅ All 6 actions covered by dispatcher tests
8. ✅ Push to origin/main
9. ⏸ Phase 5d (post-onay) — Word/PDF report generation? Or XER reader?

## Risks / TAIL fix öngörüleri

1. **`requirements.txt` missing openpyxl** → add to deps
2. **Datetime serialization** — openpyxl wants Python `datetime`, mspdi returns ISO strings → reuse `_parse_iso_date_local` (already in dcma_checks.py)
3. **Phase 5b TAIL lesson** — `_excel_collect_full_data` MUST do single-collect; do NOT call `_msp_evm_*` and `_msp_dcma_assess_all` separately if it triggers redundant Phase 5a/5b data fetches. Cache once, reuse.
4. **Cell style cost** — applying styles cell-by-cell is slow at 1000+ rows. Use openpyxl's pattern: define style once, assign to range. If 200-task export exceeds budget, batch styling.
5. **CLAUDE.md RULE 13** — sheet titles "Industry Standard EVM Metrics", "DCMA 14-Point Assessment" — McKinsey/PwC/Mace/Deloitte names YASAK.
6. **CLAUDE.md RULE 12** — RAG status: SPI<0.3 RED, 0.3-0.7 AMBER, ≥0.7 GREEN (Phase 5a uses; Phase 5b uses pass_count).
7. **Brand color consistency** — share Lacivert constant with future Word/PDF skill (Phase 5d candidate).

## Out of scope (deferred)

- xlsx → MSP file (.xml/.mpp) write — would require mspdi_parser xml writer extension (Phase 5d candidate)
- Excel formula cells (e.g., `=SPI*BAC`) — keep values flat, simpler & cross-app safe
- Pivot tables, charts — Phase 5d/5e if needed for stakeholder ask
- Macro / VBA — never (security, cross-platform)
- Conditional formatting via Excel built-in (vs flat cell color) — flat color simpler

---

*Design committed: 2026-05-01. Next step: writing-plans skill → T94-T101 detailed implementation plan.*
