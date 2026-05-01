# Construction Planning MCP Suite

Three Model Context Protocol servers for construction project planning across
Asta Powerproject and Microsoft Project.

## Servers

| Server | Module | Purpose |
|---|---|---|
| `asta_powerproject_mcp` | `asta_mcp_core.py` | Live COM control of Asta Powerproject (8 tools): tasks, links, progress, resources, schedule, codes, views, exports. |
| `asta_powerproject_file` | `asta_mcp_file.py` | File-based read access via MPXJ/MSPDI for Asta `.pp/.xer/.xml` (4 tools): query, resources, calendar, edit. |
| `msproject_mcp` | `msproject_mcp_core.py` | Microsoft Project COM with hybrid bulk routing (7 tools, ~52 actions): `msproject_task`, `msproject_link`, `msproject_schedule`, `msproject_calendar`, `msproject_resource`, `msproject_baseline`, `msproject_progress`. |

## Phase 1 Status — MS Project MCP

Phase 1 delivers a COM-based Microsoft Project MCP with hybrid bulk routing:

- **com_direct**: 1-9 ops via direct COM calls (lowest latency).
- **com_batch**: 10-19 ops with batch mode (manual calc, screen-update off, events disabled) for steady throughput.
- **mspdi_bulk**: 20+ ops via in-memory MSPDI XML build + single COM `OpenProject` import — 200 tasks in **2.30 s** measured (target was <5 s).

All 17 actions across the three tools are operational with full pytest coverage
(43/43 passing) including a 200-task end-to-end villa acceptance test.

## Phase 2a Status — Calendar (27 Apr 2026)

`msproject_calendar` tool with 7 actions:

- `create` — New base calendar from existing
- `update` — Rename or weekday off
- `add_exception` — Non-working day/range
- `assign_to_task` — Apply calendar to task
- `assign_to_resource` — Apply calendar to resource (full Resource tool in Phase 2b)
- `list` — All calendars + exception counts
- `holidays_uzbek` — Built-in 9 Özbek 2026 holidays bulk-add (idempotent, name-based dedup)

Acceptance: [`samples/build_uzbek_calendar.py`](samples/build_uzbek_calendar.py)
builds a Uzbekistan-2026 calendar end-to-end (create + 9 holidays + Sunday off
+ task assignment) in <5 sec, isolated from the user's active project.

Full pytest coverage: **83/83 passing** (43 Phase 1 + 40 Phase 2a).

## Phase 2b Status — Resource Management (28 Apr 2026)

`msproject_resource` tool with 7 actions:

- `add` — Add resource (Work / Material / Cost types, type-discriminated property surface)
- `update` — Rename, set rates, units, material label
- `delete` — Remove resource (cascade-aware: returns assignments_removed count)
- `list` — All resources with type, properties, assignment counts
- `assign` — Single assignment via task.Assignments.Add API
- `unassign` — Remove specific assignment by task+resource
- `bulk_assign` — Hybrid routing (1-5 COM direct, 6-19 batch, 20+ MSPDI bulk)

Acceptance: [`samples/build_villa_resources.py`](samples/build_villa_resources.py)
builds 14 CAU resources + 50 villa tasks + 700 assignments end-to-end (~13s
total, ~16ms/assignment via MS Project COM).

**Performance note:** True MSPDI assignment bulk merge for 2800+ assignments
in <5s is Phase 3+ scope. Pure-COM `Assignments.Add` is intrinsically
~10-16ms/call regardless of routing path; the hero target (`14 × 200 = 2800`
in <5s) is xfail until Phase 3+ implements native MSPDI assignment merge.

Full pytest coverage: **156/156 + 1 xfail** (83 Phase 1+2a + 73 Phase 2b).
Tool count after Phase 2b: **5 tools, ~31 actions**.

## Phase 3a Status — Baseline (28 Apr 2026)

`msproject_baseline` tool with 9 actions, all 11 baseline slots (Baseline + Baseline1..Baseline10):

- `save` / `clear` / `clear_all` — multi-baseline lifecycle
- `list` — all saved baselines + metadata (saved date, task count, totals)
- `get_task_baseline` — read one task's baseline values
- `compare` — current vs baseline variance + threshold filter
- `compare_two` — baseline-to-baseline delta (revision tracking)
- `summary` — project-level RAG (green<=5% slipped, amber<=20%, red>20%)
- `set_active` — graceful fallback if MSP version doesn't expose API

Acceptance: [`samples/build_baseline_lifecycle.py`](samples/build_baseline_lifecycle.py)
runs full Original -> progress -> revise -> compare lifecycle (40 tasks, 3 resources,
120 assignments, 2 baselines, 2 compares, summary, list) in <10s, isolated from the
user's active project.

Full pytest coverage: **203/203 + 1 xfail** (156 Phase 1+2a+2b + 47 Phase 3a).
Tool count after Phase 3a: **6 tools, ~40 actions**.

## Phase 3b — Progress Management (29 Apr 2026)

`msproject_progress` tool with 12 actions, dual-track progress + time-phased
actuals + EVM foundation:

**Task-level:**
- `set_task_progress` / `get_task_progress` — % complete, % work complete,
  actual_start/finish, actual_work, remaining_work, **physical_pct (DCMA)**,
  stop/resume

**Assignment-level (per-resource man-hour):**
- `set_assignment_progress` / `get_assignment_progress` — assignment.ActualWork,
  PercentWorkComplete, RemainingWork, Units (rolls up to task automatically)

**Time-phased (TimeScaleData):**
- `time_phased_actual_write` / `time_phased_actual_read` — per-day or
  per-week actual_work buckets for hakediş/EVM period delta reporting

**Bulk operations:**
- `set_progress_by_date` — `app.UpdateProject(ProgressDate)` retroactive
  catch-up (plan = actual up to date)
- `bulk_progress_update` — hybrid 1-5/6-19/20+ path (Phase 2b T37 pattern)
- `set_status_date` — `proj.StatusDate` (data_date)
- `clear_progress` / `clear_all_progress` — reset progress

**EVM-ready aggregate:**
- `summary` — BAC, ACWP, project_pct_complete, task counts (completed/
  in_progress/not_started). Foundation for upcoming Phase 5 `msproject_evm`.

Acceptance: [`samples/build_progress_lifecycle.py`](samples/build_progress_lifecycle.py)
runs full progress lifecycle in <15s (measured 5.74s on dev box).

Tool count: **7 tools, ~52 actions**.

## Phase 4 — File MCP (30 Apr 2026)

`msproject_file` tool — file-based read+write for MS Project files
(`.xml`/`.mspdi`/`.mpp`). 14 actions covering read, write, query, and
bulk assignment HERO. **MS Project does not need to be running** for
file operations; if it is and a project's `FullName` matches the
edited file_path, write actions automatically `FileClose+FileOpen+
Reschedule` for clean reload (conservative auto-sync — never touches
unrelated projects).

**Read (8):**
- `read_tasks` / `read_links` / `read_resources` / `read_assignments`
  / `read_calendars` — unified contract dicts via factory dispatch
- `read_baselines` (Phase 3a integration) — minimal contract on XML
  path until Phase 5 extends `mspdi_parser.add_baseline`
- `read_progress` (Phase 3b integration) — % complete + actual_start/
  finish (N/A sentinel normalized to None for EVM safety) +
  actual_work_h + status_date
- `query` — restricted-eval filter expressions (==, !=, <, <=, >, >=,
  AND, OR). Sandboxed: empty `__builtins__` + forbidden-token preflight
  (rejects `__`, `import`, `exec`, `eval`, `lambda`, `:=`, etc.)

**Write (6, XML only — `.mpp` is Microsoft proprietary):**
- `add_tasks` / `add_links` / `add_resources` — bulk add via
  `mspdi_parser` (extended in T67/T68/T70 to expose `is_base`,
  `actual_*`, `status_date`, `add_resource`)
- `bulk_add_assignments` — **🚀 HERO**: 2800 assignments in <5s via
  single XML write pass (pure Python, no COM crossing)
- `update_task` — duration/name/percent_complete/notes/start/finish
- `save_as` — copy to new `.xml`/`.mspdi` path

**Format dispatch:**
- `.xml`/`.mspdi` → native `MspdiProject` (zero Java dependency)
- `.mpp` → `MspMppFileManager` (MPXJ + JPype JVM, lazy init —
  read-only)
- XML schema sniff (`_detect_msp_xml_schema`) refuses non-MSPDI XML
  with informative redirect to Asta MCP

Acceptance: [`samples/build_file_lifecycle.py`](samples/build_file_lifecycle.py)
runs full read+write+HERO lifecycle in <30s (measured 0.57s on dev
box; HERO 2800 assignments alone 0.07s).

Tool count: **8 tools, 14 + ~52 = ~66 actions**.

## Phase 5a — EVM (30 Apr 2026)

`msproject_evm` tool — Earned Value Management per PMI PMBOK 8th
§ 7.4.2 + Lipke 2003 Earned Schedule. 13 actions covering
CLAUDE.md RULE 4-9 + RULE 12 RAG + RULE 3 currency mode auto-detect.
Hybrid: `file_path` verilirse Phase 4 file path; yoksa Phase 1 COM.

**Compute (4):** compute_metrics (SPI/CPI/SV/CV), forecast (EAC1/2/3,
ETC, VAC, TCPI), earned_schedule (AT, ES, SV(t), SPI(t)), summary (RAG)

**Time-Phased (2):** time_phased_evm (PV/EV/AC per bucket
day/week/month), period_delta (RULE 6 haftalik delta)

**Data Quality (1):** progress_data_quality (RULE 7 SPI(h) vs SPI(t))

**Baseline Integration (2):** variance_to_baseline (Phase 3a),
compare_baselines_evm (B_a vs B_b)

**History (3):** save_period_snapshot, get_period_history, trend
(JSON-backed at `~/msproject_evm_snapshots.json`)

**Setup (1):** detect_currency_mode (RULE 3)

Architecture: pure-math `evm_math.py` (RULE 4-9 implementations,
MSP/COM/file independent) + I/O adapters in msproject_mcp_core.py.
Phase 1-4 helpers DOKUNULMAZ.

Acceptance: `samples/build_evm_lifecycle.py` runs 200-task CAU-style
hero with 4 snapshots in <30s.

Tool count: **9 tools, ~79 actions** (Phase 4 14 + Phase 3b 12 +
Phase 3a 9 + Phase 2b 7 + Phase 2a 7 + Phase 1 6 + 1 progress).

## Phase 5b — DCMA 14-Point (1 May 2026)

`msproject_health` tool — DCMA 14-Point Schedule Health Assessment per
CLAUDE.md RULE 10. 4 actions covering all 14 rules with industry-standard
hardcoded thresholds. Hybrid: `file_path` verilirse Phase 4 file path;
yoksa Phase 1 COM. Read-only.

**Actions:**
- `assess_all`: 14 rules + summary + RAG (>=12 pass GREEN, 8-11 AMBER, <8 RED)
- `summary`: RAG + executive text only
- `drill_down(rule_id=1..14)`: per-rule failed task list
- `compare(snapshot_path)`: DCMA delta vs prev snapshot (reuses Phase 5a snapshot file)

**Rules grouped by category:**
- Logic (1-5): no_pred, no_succ, leads, lags, fs_link
- Constraints (6): hard_constraints
- Float (7-8): high_float, negative_float
- Duration (9): high_duration
- Quality (10-11): invalid_dates, resources_missing
- Schedule (12-14): missed_tasks, critical_path, BEI

Architecture: pure-math `dcma_checks.py` (14 fixture-free check functions
+ aggregator + RAG, MSP/COM/file independent, ~82 tests) + I/O adapters
(`_dcma_load_links`, `_dcma_collect_full_data`, 4 `_msp_dcma_*` action
helpers) in msproject_mcp_core.py. Phase 1-5a helpers DOKUNULMAZ.

Acceptance: `samples/build_dcma_lifecycle.py` runs 200-task CAU-style with
intentional DCMA failures (12 no-predecessor + 15 high-duration + 12
unassigned), drill-down per failing rule, in <60s.

Tool count: **10 tools, ~83 actions**.

## Phase 5c — Excel I/O (1 May 2026)

`msproject_excel` tool — multi-sheet hakediş workbook export + Excel-driven
import. 6 actions covering Phase 5a EVM + Phase 5b DCMA → multi-sheet
xlsx, plus bulk Excel → MSP imports for tasks and progress. MCS brand
styling per CLAUDE.md RULE 14 (Lacivert Calibri headers, RAG color cells,
zebra rows).

**Actions:**
- `export_hakedis`: 6-sheet workbook (Summary + Tasks + EVM_Compute +
  EVM_TimePhased + DCMA_Rules + DCMA_Failed)
- `export_tasks`: Tasks sheet only (filter-friendly schedule export)
- `export_evm`: EVM_Compute (BAC/EV/AC/SPI/CPI/EAC/TCPI/ES/RAG) +
  EVM_TimePhased (PV/EV/AC + cumulative per period)
- `export_dcma`: DCMA_Rules (14 rule statuses with RAG color) +
  DCMA_Failed (drill-down task list, capped 10/rule)
- `import_tasks`: xlsx Tasks sheet → `_msp_task_bulk_add` (Phase 1)
- `import_progress`: xlsx Progress sheet → `_msp_progress_bulk_update` (Phase 3b)

Architecture: pure-Python `excel_io.py` (openpyxl 3.1.5, MSP/COM/file
independent, 41 unit tests) + I/O adapters in msproject_mcp_core.py
(`_excel_collect_full_data` single-collect aggregator + 6 action helpers).
Phase 1-5b helpers DOKUNULMAZ; only read-only calls.

Acceptance: `samples/build_excel_lifecycle.py` exports 200-task hakediş
workbook + roundtrip imports 10 progress updates in <90s.

Tool count: **11 tools, ~89 actions**.

## Quick Start

1. Install deps: `pip install -r requirements.txt`
2. Open Asta Powerproject and/or MS Project so a live COM endpoint exists.
3. Register the servers with Claude Desktop / your MCP client — see
   [`samples/claude_mcp_config_snippet.json`](samples/claude_mcp_config_snippet.json)
   for the exact `mcpServers` block to merge into your config.
4. Restart the client; the three servers (`asta_powerproject_mcp`,
   `asta_powerproject_file`, `msproject_mcp`) appear as tool sources.

## Testing

```
python -m pytest tests/ -v
```

Tests automatically skip if MS Project / Asta is not running on the host.

## Repository Layout

- `msproject_mcp_core.py` — MS Project MCP (Phase 1).
- `msproject_bulk.py` — MSPDI bulk-import writer.
- `mspdi_parser.py` — Native MSPDI XML reader/writer (shared by both ecosystems).
- `asta_mcp_core.py` — Asta Powerproject COM MCP.
- `asta_mcp_file.py` — Asta file-based MCP.
- `tools/` — One-shot helper scripts (typelib dump, fixture generator).
- `tests/` — Pytest suite. Fixtures in `tests/fixtures/` (see fixtures README).
- `docs/plans/` — Phase 1 design + implementation plan.
- `samples/` — End-to-end build scripts and MCP config snippet.
