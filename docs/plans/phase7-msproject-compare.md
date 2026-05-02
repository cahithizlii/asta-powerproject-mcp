# Phase 7 — msproject_compare (T123-T130)

> First new tool since Phase 5d. Tool count 12 → 13.

## Goal

CAU monthly hakediş workflow needs to answer: *"What changed between
last month's baseline XER and this month's progress XER?"* — added/
removed tasks, schedule slips, progress jumps, EVM deltas.

Phase 7 ships `msproject_compare` — a pure-diff tool over two file
snapshots, no COM dependency, hybrid XER + MSPDI input.

## Strategy

Yaklaşım C (5th application):
- `xer_compare.py` — pure diff math, fixture-free
- `msproject_mcp_core.py` — 5 adapters wrapping `_evm_load_task_data`
  for both file paths
- Dispatcher: new `msproject_compare` `@mcp.tool` with 5 actions

## Pure module API (`xer_compare.py`)

```python
diff_tasks(tasks_a, tasks_b, fields=None) -> {
    added: [task],     # task in B not in A (matched by id)
    removed: [task],   # task in A not in B
    changed: [{id, name, fields_changed: {field: (a, b)}}],
    unchanged_count: int,
}

diff_links(links_a, links_b) -> {
    added: [link],
    removed: [link],
    # identity = (from_id, to_id, type); changed lag captured in changed
    changed: [{from_id, to_id, type, lag_a, lag_b}],
    unchanged_count: int,
}

diff_progress(progress_a, progress_b) -> {
    status_date_a, status_date_b,
    tasks: [{id, pct_a, pct_b, pct_delta,
             aw_a, aw_b, aw_delta}],   # only tasks that moved
    summary: {total_pct_delta, total_aw_delta, count_moved},
}

diff_evm(snap_a, snap_b) -> {
    bac_delta, pv_delta, ev_delta, ac_delta,
    spi_a, spi_b, spi_delta,
    cpi_a, cpi_b, cpi_delta,
}

summarize_compare(task_diff, link_diff, progress_diff, evm_diff) -> {
    headline: '5 tasks added, 12 progressed, SPI 0.74 -> 0.81',
    counts: {tasks_added, tasks_removed, tasks_changed,
             links_added, links_removed, links_changed,
             tasks_progressed},
    spi_delta, cpi_delta, ev_delta,
}
```

## Adapters (in `msproject_mcp_core.py`)

```python
_msp_compare_tasks(file_path_a, file_path_b, fields=None)
_msp_compare_links(file_path_a, file_path_b)
_msp_compare_progress(file_path_a, file_path_b)
_msp_compare_evm(file_path_a, file_path_b, baseline_number=0)
_msp_compare_summary(file_path_a, file_path_b)
```

All use `_evm_load_task_data(file_path)` for shape (XER + MSPDI hybrid
already proven in Phase 5e routing).

## Dispatcher

`msproject_compare` — new `@mcp.tool`:

```
{action: 'task_delta'|'link_delta'|'progress_delta'|'evm_delta'|'summary',
 file_a: '/path/last.xer',
 file_b: '/path/this.xer',
 fields?: ['baseline_start','baseline_finish','duration_h']  # task_delta only
}
```

## Identity rules

| Object | Identity | Notes |
|---|---|---|
| Task | `id` | int from MSP/XER |
| Link | `(from_id, to_id, type)` | type FF/FS/SS/SF |
| Progress entry | `id` | task_id |

## Default change-detect fields (tasks)

`baseline_start`, `baseline_finish`, `baseline_work`,
`percent_complete`, `actual_start`, `actual_finish`, `actual_work`,
`duration_h`. User can override via `fields` param.

## Tests

### Unit (~30 tests, `tests/test_xer_compare.py`)
- `diff_tasks` (8): empty/empty, all-add, all-remove, identity-no-change,
  field change detection, custom fields, name change ignored unless
  in fields, multi-field change
- `diff_links` (5): all-add, all-remove, lag change → changed not removed,
  type change → removed+added (different identity), unchanged
- `diff_progress` (6): no movement, single move, total delta math,
  status_date passthrough, missing in B, no progress data
- `diff_evm` (5): identity, partial deltas, None handling, division-safe
- `summarize_compare` (6): headline format, count aggregation, sign
  preservation, empty diffs

### Integration (~10 tests, `tests/test_msproject_compare_dispatcher.py`)
- 5 actions × {XER vs XER, MSPDI vs MSPDI} smoke tests
- Bad path / missing file error returns
- Identity files diff = unchanged_count > 0, added/removed = 0

### Acceptance (1 script, `samples/build_compare_lifecycle.py`)
- Build snapshot A (XER 0% progress)
- Build snapshot B (XER 50% progress, +1 task)
- Run all 5 dispatcher actions
- Assert: B has +1 added, progress moved, SPI/CPI deltas computed

## Tasks

| Task | Scope |
|---|---|
| T123 | xer_compare.diff_tasks pure + 8 unit tests |
| T124 | diff_links + diff_progress + 11 unit tests |
| T125 | diff_evm + summarize_compare + 11 unit tests |
| T126 | 5 adapters in msproject_mcp_core |
| T127 | msproject_compare @mcp.tool dispatcher (13th tool) |
| T128 | 10 integration tests |
| T129 | acceptance script + samples |
| T130 | commit + push + memory update |

## Backward compat

- New tool: zero impact on existing 12 tools
- New pure module: zero shared dependencies
- `_evm_load_task_data` reused read-only (DOKUNULMAZ)

## Regression target

479 → ~530 PASS (+50). Phase 7 target zero regression.
