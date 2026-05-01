# Phase 5e — XER Native Integration Design (1 May 2026)

## Goal

Unlock CAU XER value: enable `msproject_health.assess_all(file_path="cau.xer")` and `msproject_excel.export_hakedis(file_path="cau.xer")` end-to-end. Phase 5d shipped XER reader; Phase 5e wires it into Phase 5a EVM + Phase 5b DCMA + Phase 5c Excel pipelines via minimal additive routing in Phase 5a loaders.

## Background — the integration gap

Phase 5d shipped `msproject_xer` with 6 read-only actions. But Phase 5a `_evm_load_task_data(file_path=...)` (which underpins Phase 5b DCMA + Phase 5c Excel collect helpers) only knows `.xml`/`.mpp`. Calling `assess_all(file_path="cau.xer")` today fails because Phase 4 file readers (`_msp_file_read_tasks`) don't recognize XER format.

Result: user must read XER, manually transform data, pass through unrelated tool. Killer feature unrealized.

## Brainstorming decision (1 May 2026)

- **Option A chosen** (vs B shim / C unified tool / D xer_only convenience): minimal additive `.xer` extension routing in Phase 5a loaders. DOKUNULMAZ "spirit preserved" — existing `.xml`/`.mpp`/COM paths untouched, regression maintained.
- Phase 5b/5c collect helpers automatically benefit (call `_evm_load_task_data` already).
- Cost-loaded NO XER (CAU pattern): baseline = current schedule (target_* fields = baseline_* fields).

## Tool surface — NO new tools

Phase 5e changes ONE thing: existing `assess_all`/`summary`/`drill_down`/`compare`/`export_hakedis`/`export_evm`/`export_dcma` actions ALL accept `.xer` file_path argument now. Tool count stays 12.

## Architecture

```
msproject_mcp_core.py PHASE 5E SECTION (after Phase 5d msproject_xer dispatcher,
                                        before def main):
├── _xer_to_evm_task_shape(xer, day_hr_cnt) -> {status, tasks, resources,
│      assignments, status_date, project_file}
│      Translates XerFile output to Phase 5a _evm_load_task_data return shape.
│      Adds baseline_start/finish/work + actual_work (from XER assignments).
└── _xer_to_evm_baseline_shape(xer, day_hr_cnt, baseline_number) -> {status, tasks: [...]}
       Translates XerFile output to Phase 5a _evm_load_baseline_data return shape.
       CAU pattern: baseline = target (cost-loaded NO).

msproject_mcp_core.py Phase 5A loader EXTENSIONS (additive only):
├── _evm_load_task_data: ADD top guard
│      if file_path and file_path.lower().endswith('.xer'):
│          return _xer_to_evm_task_shape(XerFile(file_path), ...)
│      # existing code unchanged
└── _evm_load_baseline_data: ADD top guard, same pattern
```

Phase 1+2+3+4+5b+5c+5d helpers DOKUNULMAZ. Phase 5a `_evm_load_task_data` + `_evm_load_baseline_data` get ONE guard line each (additive). Existing tests for `.xml`/`.mpp`/COM continue to pass.

## XER → Phase 5a shape mapping

| Phase 5a expected key | XER source |
|---|---|
| `tasks[i].id` | `XerFile.read_tasks()[i]["id"]` |
| `tasks[i].name` | XerFile read_tasks name |
| `tasks[i].duration_h` | XerFile target_drtn_hr_cnt |
| `tasks[i].start` | XerFile target_start_date (ISO) |
| `tasks[i].finish` | XerFile target_end_date (ISO) |
| `tasks[i].percent_complete` | XerFile phys_complete_pct |
| `tasks[i].summary` | XerFile (TT_LOE/TT_WBS only) |
| `tasks[i].baseline_start` | = `tasks[i].start` (CAU baseline = target) |
| `tasks[i].baseline_finish` | = `tasks[i].finish` |
| `tasks[i].baseline_work` | = `tasks[i].duration_h` |
| `tasks[i].actual_work` | sum(TASKRSRC.act_reg_qty per task_id) |
| `tasks[i].constraint_type` | XerFile constraint_type (already mapped 0-7) |
| `tasks[i].total_slack_days` | = `tasks[i].total_float` (DCMA Rule 7-8) |
| `tasks[i].critical` | derived: `total_slack_days <= 0` (XER lacks explicit critical flag) |
| `tasks[i].predecessors` | derived from `read_links()`: list of pred from_id where to_id == this task |
| `tasks[i].successors` | derived from `read_links()`: list of succ to_id where from_id == this task |
| `resources` | XerFile.read_resources() — already in MSP shape |
| `assignments` | XerFile.read_assignments() — already in MSP shape |
| `status_date` | XerFile.read_progress()["status_date"] |
| `project_file` | the file_path string |

## Critical fields for DCMA Rules

- **Rule 1 (no_pred)**: needs `predecessors` list per task → derive from XerFile.read_links()
- **Rule 2 (no_succ)**: needs `successors` list → same source
- **Rule 6 (hard constraints)**: `constraint_type` (already mapped via CONSTRAINT_TYPE_MAP)
- **Rule 7 (high float)**: `total_slack_days` (= XER total_float in days)
- **Rule 8 (negative float)**: same
- **Rule 9 (high duration)**: `duration_h` (already present)
- **Rule 10 (invalid dates)**: `start`/`finish` ISO strings (already)
- **Rule 11 (resources missing)**: `assignments` list (already)
- **Rule 12 (missed tasks)**: `baseline_finish` + `percent_complete` + `status_date` (all derivable)
- **Rule 13 (critical path)**: derive `critical = total_slack_days <= 0`
- **Rule 14 (BEI)**: same as Rule 12 inputs

## Critical fields for Excel hakediş

- **Tasks sheet**: id/name/duration/start/finish/%complete/critical/summary — all available
- **EVM_Compute**: BAC/EV/AC/PV/SPI/CPI — derived from task baseline_work + percent_complete + assignments. CAU cost-loaded NO → SPI(h) (hours-based) computed correctly
- **DCMA_Rules**: 14 rules from above mapping
- **Summary**: BAC + EAC + RAG — derived

## Data flow (post-Phase 5e)

```
file_path = "cau.xer"
    ↓
msproject_health.assess_all(file_path) [Phase 5b dispatcher]
    ↓
_msp_dcma_assess_all → _dcma_collect_full_data
    ↓
_evm_load_task_data(file_path) [Phase 5a + new XER guard]
    ↓ (file_path.endswith('.xer') → True)
_xer_to_evm_task_shape(XerFile(file_path)) [Phase 5e adapter]
    ↓
returns Phase 5a-shape dict
    ↓
_dcma_assess_all (pure math) → 14 rules + RAG
    ↓
JSON response
```

## Error handling

- Invalid XER (parse fails) → `{"status": "error", "error": "XER parse failed: ..."}`
- Missing CALENDAR section → default 8h/day, no crash
- Empty TASK section → `{"status": "ok", "tasks": [], ...}` (vacuous PASS for all DCMA rules)

## Testing

### Adapter tests — `tests/test_phase5e_xer_to_evm.py` (~10 tests)
- `_xer_to_evm_task_shape` returns Phase 5a shape with all expected keys
- baseline_* fields populated
- predecessors/successors derived correctly
- critical flag from total_slack_days
- actual_work aggregated from XER assignments

### Integration tests — `tests/test_phase5e_xer_integration.py` (~8 tests)
- `_msp_dcma_assess_all(file_path=cau.xer)` returns 14 rules
- `_msp_excel_export_hakedis(file_path=cau.xer, xlsx_path=...)` produces 6-sheet xlsx
- Phase 1-5d regression: 370 PASS still

### Acceptance — `samples/build_xer_dcma_excel_lifecycle.py`
**Scenario:**
1. Write synthetic CAU XER to tempdir
2. `msproject_health.assess_all(file_path=xer)` → display 14 DCMA rules
3. `msproject_excel.export_hakedis(file_path=xer, xlsx_path=...)` → 6-sheet workbook
4. Verify workbook structure
5. Total time ≤ 30s (XER pure-Python parse + DCMA pure-math + xlsx I/O all fast)

## Sequencing — T109-T111

| Task | Content | Pattern |
|---|---|---|
| T109 | `_xer_to_evm_task_shape` adapter + extend `_evm_load_task_data` with .xer guard + adapter tests | Manuel |
| T110 | `_xer_to_evm_baseline_shape` adapter + extend `_evm_load_baseline_data` with .xer guard + integration tests | Manuel |
| T111 | E2E acceptance (CAU XER → assess_all DCMA + export_hakedis Excel) + README + push | Manuel finalize |

**Estimated chain:** ~18 new tests. Cumulative regression target: 370 + ~18 = ~388 PASS.

## Acceptance criteria

1. ✅ T109-T111 3-task chain landed
2. ✅ Acceptance ≤ 30s (CAU XER → DCMA + Excel end-to-end)
3. ✅ All 14 DCMA rules return for XER input
4. ✅ Hakediş xlsx 6 sheets created
5. ✅ Phase 1-5d regression untouched (370 PASS)
6. ✅ Push to origin/main

## Risks / TAIL fix öngörüleri

1. **`_evm_load_task_data` modification** — additive only. Phase 1-5d test suite must stay 370 PASS.
2. **CAU cost-loaded NO** — `target_cost == target_qty` ⇒ EVM cost-mode false-positive. `_evm_detect_currency_mode` already handles this (CLAUDE.md RULE 3).
3. **`critical` derivation** — XER doesn't store explicit critical flag. Heuristic: `critical = (total_slack_days <= 0)`. Acceptable for DCMA Rule 13 (count > 0); chain with zero slack tasks → at least 1 critical.
4. **`predecessors`/`successors` derivation** — O(N×M) where N=tasks, M=links. For CAU 200-task with ~250 links: 50K ops, ~0.05s. Negligible.
5. **`baseline_work` source** — CAU baseline = target (cost-loaded NO assumption). For cost-loaded XERs (P6 standard), `target_qty` from TASKRSRC may be more accurate. Phase 6 enhancement candidate.

## Out of scope (deferred)

- XER → Phase 4 file MCP routing (`_msp_file_read_tasks` etc.) — symmetry with Phase 4. Not needed for killer feature; Phase 6 candidate
- XER write — P6 not user's authoring tool
- Calendar holiday detail (clndr_data BLOB) — Phase 6
- Multi-baseline support — XER has only 1 implicit baseline; multi-baseline requires XER history files (deferred)

---

*Design committed: 2026-05-01. Next step: writing-plans skill → T109-T111 implementation plan.*
