# Phase 8 — Convenience Actions (T131-T134)

> Polish wave 2. Bundles existing tools into ergonomic single calls.

## Goal

Phase 7 shipped `msproject_compare` for delta analysis but the
CAU monthly hakediş workflow still requires three tool calls:
1. `msproject_compare summary` — what changed
2. `msproject_evm summary` — current EVM state
3. `msproject_excel export_hakedis` — generate Excel report

Phase 8.1 collapses those into one `msproject_compare monthly_report`
call. Phase 8.2 exposes Phase 6.3's `MspdiProject.write_baseline`
through the file MCP so callers don't have to instantiate the parser
class directly.

## Sub-phases

### 8.1 `monthly_report` convenience action

`msproject_compare`'a 6. action eklenir:

```
{action: 'monthly_report',
 file_a: 'last_month.xer',
 file_b: 'this_month.xer',
 baseline_number?: 0,
 output_excel?: 'hakedis.xlsx'}  # optional Excel side-effect
```

Bundles internally:
- `_msp_compare_summary` (Phase 7)
- `_msp_evm_summary` for both files
- Optional: `_msp_excel_export_hakedis` writing to output_excel

Returns:
```json
{
  "status": "ok",
  "compare_summary": { ... },           // Phase 7 summary
  "evm_a": { rag, completion_pct, spi, cpi },
  "evm_b": { rag, completion_pct, spi, cpi },
  "excel_path": "...",                  // present iff output_excel given
  "headline": "1 added, 2 progressed, RAG amber->amber"
}
```

### 8.2 `write_baseline` file MCP action

`msproject_file`'a yeni action eklenir:

```
{action: 'write_baseline',
 file_path: 'project.xml',
 baseline_number: 0,
 baseline_data: [
    {task_uid: 1, baseline_start: '...', baseline_finish: '...',
     baseline_duration_h: 240.0, baseline_work_h: 160.0},
    ...
 ],
 output_path: 'project_with_baseline.xml'}  # required (no in-place write)
```

Reuses Phase 6.3 `MspdiProject.write_baseline` + `save()`.

Returns:
```json
{
  "status": "ok",
  "tasks_written": N,
  "output_path": "...",
  "baseline_number": 0
}
```

## Tasks

| Task | Scope |
|---|---|
| T131 | _msp_compare_monthly_report adapter + dispatcher action |
| T132 | monthly_report integration tests (3-5 tests) |
| T133 | _msp_file_write_baseline adapter + dispatcher action |
| T134 | write_baseline integration tests + commit + push |

## Backward compat

- Both tools (`msproject_compare`, `msproject_file`) gain ONE new
  action each — no schema/return changes for existing actions
- Phase 7 `summary` action behavior preserved
- Phase 4-6 file MCP actions preserved
- New Excel export path optional — omitting param keeps behavior
  pure read

## Tests

- `tests/test_msproject_compare_dispatcher.py` extended: 3-4 monthly_report cases
- `tests/test_msproject_file_dispatcher_baseline.py` (NEW): 4-5 write_baseline cases

## Regression target

522 → ~545 PASS (+23). Zero regression.
