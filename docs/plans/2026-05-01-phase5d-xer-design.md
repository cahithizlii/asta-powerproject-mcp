# Phase 5d — `msproject_xer` Primavera P6 Reader Design (1 May 2026)

## Goal

12th MCP tool: `msproject_xer`. Pure-Python Primavera P6 XER file reader. Bridges CAU-style XER projects (recurring user workflow) into the Phase 5a EVM + Phase 5b DCMA + Phase 5c Excel pipelines.

## Background

CAU Hospital Kaba İşler is THE recurring real-world project (memory: BAC 5,058,787 hours, 14 CAU resources, Uzbekistan 6×9 calendar, baseline = XER, cost-loaded NO). Phases 5a/5b/5c can ingest MS Project (.xml/.mpp) and act on COM directly, but XER files force conversion-via-Asta or mpxj/Java-bridge today. A native pure-Python XER reader closes the loop.

## Brainstorming decisions (1 May 2026)

- **Scope:** read-only (6 actions). XER write deferred (format complex, MS Project cannot write XER, P6 not user's authoring tool — only read needed).
- **Phase 5a integration:** New `_xer_collect_full_data` adapter (Yaklaşım C, parallel of Phase 5b/5c). Phase 5a `_evm_load_task_data` DOKUNULMAZ.
- **Fixture:** Synthetic CAU-style XER (privacy + test stability), built in `tests/fixtures/sample_cau.xer`.
- **Library:** Pure Python only. NO mpxj dependency. XER format is text (UTF-16-LE BOM tab-delimited tables), tractable in ~400-500 lines.

## XER format reference

```
ERMHDR\t<version>\t<exported>\t<user>\t<app>\t<currency>\n
%T\tTASK\n
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\t...
%R\t<row_data tab-separated>\n
%R\t<row_data tab-separated>\n
%T\tTASKPRED\n
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt\n
%R\t...
%T\tRSRC\n
...
%E\n
```

- `%T` = table name marker
- `%F` = field names (column headers)
- `%R` = row data (tab-separated, position-mapped to %F headers)
- `%E` = end of file
- Encoding: typically UTF-16-LE with BOM (`\xff\xfe`); fallback UTF-8

Key tables:
- `TASK` — tasks: task_id, task_code, task_name, target_drtn_hr_cnt, target_start_date, target_end_date, act_start_date, act_end_date, phys_complete_pct, total_float_hr_cnt, status_code, task_type, cstr_type
- `TASKPRED` — links: task_id, pred_task_id, pred_type (PR_FS|PR_SS|PR_FF|PR_SF), lag_hr_cnt
- `RSRC` — resources: rsrc_id, rsrc_name, rsrc_type, max_qty_per_hr
- `TASKRSRC` — assignments: task_id, rsrc_id, target_qty, act_reg_qty, target_cost, act_reg_cost
- `CALENDAR` — calendars: clndr_id, clndr_name, day_hr_cnt, week_hr_cnt, clndr_data
- `PROJECT` — project metadata: proj_id, plan_start_date, plan_end_date, last_recalc_date (= status_date)

## Tool surface — `msproject_xer` 6 read actions

| # | Action | Inputs | Output |
|---|---|---|---|
| 1 | `read_tasks` | `file_path`, `filters?`, `limit?` | List of MSP-shape task dicts |
| 2 | `read_links` | `file_path` | List of `{from_id, to_id, type, lag_days}` |
| 3 | `read_resources` | `file_path` | List of `{id, name, type, max_units}` |
| 4 | `read_assignments` | `file_path` | List of `{task_id, resource_id, target_qty, actual_qty}` |
| 5 | `read_calendars` | `file_path` | List of `{id, name, day_hr_cnt, week_hr_cnt}` |
| 6 | `read_progress` | `file_path` | `{status_date, tasks: [{id, percent_complete, actual_work_h}]}` |

All actions return `{"status": "ok", "count": N, "<key>": [...]}` or `{"status": "error", "error": "..."}`.

## Architecture (Yaklaşım C — 4th application)

```
xer_parser.py (NEW, pure Python, ~400-500 lines)
├── XerFile class:
│   ├── __init__(file_path) — open + detect encoding (UTF-16-LE BOM or UTF-8)
│   ├── _parse() — split into tables by %T markers
│   └── tables: dict {table_name: {"headers": [...], "rows": [{col: val}]}}
├── Public methods (return MSP-shape dicts):
│   ├── read_tasks() → [{id, name, duration_h, start, finish, percent_complete, ...}]
│   ├── read_links() → [{from_id, to_id, type, lag_days}]
│   ├── read_resources() → [{id, name, type, max_units}]
│   ├── read_assignments() → [{task_id, resource_id, target_qty, actual_qty}]
│   ├── read_calendars() → [{id, name, day_hr_cnt, week_hr_cnt}]
│   └── read_progress() → {status_date, tasks: [...]}
└── Helpers: _parse_date(s), _parse_float(s), _link_type_map (PR_FS->FS), etc.

msproject_mcp_core.py PHASE 5D SECTION (after Phase 5c excel dispatcher, before def main)
├── from xer_parser import XerFile
├── _xer_collect_full_data(file_path, baseline_number=0)
│     → {tasks, links, assignments, resources, calendars, status_date}
│     SINGLE collect (Phase 5b/5c TAIL lesson)
├── 6 _msp_xer_read_* action helpers (each thin wrapper around XerFile methods)
└── @mcp.tool msproject_xer dispatcher (6 actions)
```

**Phase 1+2+3+4+5a+5b+5c helpers DOKUNULMAZ.** Read-only.

## Data flow (read_tasks example)

```
file_path (.xer)
    ↓
XerFile(file_path) — detect encoding, parse tables
    ↓
xer.read_tasks() → list of dicts in MSP shape
    ↓
optional filter via _filter_tasks (Phase 4 reuse)
    ↓
return {"status": "ok", "count": N, "tasks": [...]}
```

## Field mapping (XER → MSP-shape)

| XER field | MSP shape key | Notes |
|---|---|---|
| `task_id` | `id` | int |
| `task_code` | `code` | string (e.g. "A1010") |
| `task_name` | `name` | string |
| `target_drtn_hr_cnt` | `duration_h` | float, hours |
| `target_start_date` | `start` | ISO date "YYYY-MM-DD" |
| `target_end_date` | `finish` | ISO date |
| `act_start_date` | `actual_start` | ISO date or None |
| `act_end_date` | `actual_finish` | ISO date or None |
| `phys_complete_pct` | `percent_complete` | float 0-100 |
| `total_float_hr_cnt` | `total_float` | float, days (XER hours / day_hr_cnt) |
| `task_type` | `task_type` | TT_Task/TT_Mile/TT_FinMile/TT_LOE — `summary` flag if TT_LOE |
| `cstr_type` | `constraint_type` | CS_MSO/CS_MFO/etc → enum 0-7 (DCMA Rule 6) |
| `status_code` | `status` | TK_Active/TK_Complete/TK_NotStart |

| XER link field | MSP shape key |
|---|---|
| `task_id` | `to_id` (successor) |
| `pred_task_id` | `from_id` (predecessor) |
| `pred_type` | `type` (PR_FS→"FS", PR_SS→"SS", etc.) |
| `lag_hr_cnt` | `lag_days` (hours / 8) |

## Synthetic CAU-style fixture

`tests/fixtures/sample_cau.xer`:
- 10 tasks (mix of WORK + 2 milestones)
- 9 FS links (chain)
- 4 CAU resources (COW, EXT, STL, CAR)
- 1 calendar (54h/week, 9h/day)
- 1 baseline-equivalent project section
- Synthetic dates: 2024-07-08 (CAU plan start) → ~2024-12-15

UTF-16-LE BOM, tab-delimited per spec. Generated programmatically in conftest if needed (avoid binary blob in repo).

## Error handling

- File not found → `{"status": "error", "error": "File not found: ..."}`
- Encoding detection failure → fallback UTF-8, then error if both fail
- Missing required table (TASK absent) → `{"status": "error", "error": "XER missing TASK table"}`
- Malformed row (column count mismatch) → skip row + warning log, continue parse
- Empty XER (only ERMHDR + %E) → `{"status": "ok", "count": 0, "<key>": []}`

## Testing

### Unit tests — `tests/test_xer_parser.py` (~25 tests, fixture-based)
- Encoding detection: UTF-16-LE BOM, UTF-8, fallback
- Table parsing: `%T`/`%F`/`%R` markers
- read_tasks: count, task_id type, percent_complete float
- read_links: type mapping (PR_FS→FS), lag conversion
- read_resources: max_units present
- read_assignments: task_id and resource_id linkage
- read_calendars: day_hr_cnt parsed (CAU 9h/day)
- read_progress: status_date from PROJECT.last_recalc_date
- Error: missing file, malformed XER

### Integration tests — `tests/test_msproject_xer_*.py`
- `_xer_collect_full_data` from sample_cau.xer
- 6 dispatcher tests

### Acceptance — `samples/build_xer_lifecycle.py`
**Scenario:**
1. Parse `sample_cau.xer` (10 tasks)
2. `_msp_xer_read_tasks` → list
3. `_msp_xer_read_links` + read_resources + read_assignments + read_calendars + read_progress
4. Bonus: pipeline through to `assess_all` DCMA via temporary file_path support
5. Total time ≤ 30s (small fixture, fast pure-Python parse)

## Sequencing — T102-T108

| Task | Content | Pattern |
|---|---|---|
| T102 | `xer_parser.py` foundations: `XerFile.__init__`, `_parse` table splitter, encoding detect | Manuel |
| T103 | `read_tasks` + `read_links` (TASK + TASKPRED parse + field mapping) | Manuel |
| T104 | `read_resources` + `read_assignments` + `read_calendars` | Manuel |
| T105 | `read_progress` + status_date + synthetic fixture generator | Manuel |
| T106 | `_xer_collect_full_data` + 6 action helpers (BIG ONE) | **Subagent** |
| T107 | `@mcp.tool msproject_xer` dispatcher + dispatcher tests | Manuel |
| T108 | Acceptance + README + push | Manuel finalize |

**Estimated chain:** ~30-35 new tests. Cumulative regression target: 303 + ~33 = ~336 PASS.

## Acceptance criteria

1. ✅ T102-T108 7-task chain landed
2. ✅ Acceptance `build_xer_lifecycle.py` ≤ 30s @ 10-task fixture
3. ✅ All 6 actions return well-formed MSP-shape dicts
4. ✅ Phase 1-5c regression untouched
5. ✅ Push to origin/main
6. ⏸ Phase 5e: integrate XER into `_evm_load_task_data` route (or deferred to Phase 6 polish)

## Risks / TAIL fix öngörüleri

1. **Encoding detection fragility** — synthetic fixture controlled (UTF-16-LE BOM); real XER files vary
2. **Numeric format edge cases** — empty string vs "0" vs "0.0" vs None; defensive `_parse_float` returns 0.0 for empty
3. **Date format variants** — typically `YYYY-MM-DD HH:MM` but P6 versions differ; defensive parse with fallback to None
4. **Calendar `clndr_data` is opaque blob** — XER calendar data is BLOB encoded; only extract `day_hr_cnt`, `week_hr_cnt`, `clndr_name` (calendar holiday detail deferred to Phase 6)
5. **Constraint type mapping** — XER `CS_MSO` → MSP enum 2, etc. Documented mapping table needed
6. **Acceptance fixture must be in repo** — synthetic XER blob (UTF-16-LE bytes); committed as binary, ~5KB

## Out of scope (deferred)

- XER write — P6 not user's authoring tool, write complex
- Calendar holiday detail extraction (clndr_data BLOB) — Phase 6 enhancement
- TASKACTV (activity codes) — P6 specific, no MSP analogue
- WBS hierarchy reconstruction — current XER reader treats tasks flat (CAU project memory: WBS not heavily used)
- Cost rate tables, `RSRCRATE` — Phase 6 if cost-loaded XER demands

---

*Design committed: 2026-05-01. Next step: writing-plans skill → T102-T108 implementation plan.*
