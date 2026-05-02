# Phase 10 — Wrap-up + Phase 2a Back-fill (T135-T140)

> Closes the marathon-summary backlog: housekeeping (test/sample
> consolidation), Phase 2a deferred items (calendar recurring + working
> exceptions), and one polish (update_task baseline read-back).

## Sub-phases

### 10.1 update_task baseline read-back (T135) — smallest, first
After write, response includes the current baseline values that ended
up persisted (read-back via `MspdiProject.read_baselines(num)`).

**Adapter change:** `_msp_file_update_task` — when baseline_written>0,
re-read the task's baseline element after `mgr.save()` and embed in
response as `baseline_after`.

**Tests:** 2 new — schema check + value match after roundtrip.

### 10.2 Calendar recurring exceptions (T136) — Phase 2a backlog
`msproject_mcp_core.py:354-365` deferred recurring (weekly/monthly)
calendar exceptions to "Phase 3+". Adds support via MS Project COM
`Calendar.Exceptions.Add(...)` with `Type=pjException{Daily,Weekly,
Monthly,Yearly}` and recurring fields.

**Tests:** Daily/Weekly/Monthly/Yearly cases — 4 new (skip if MSP COM
not available, follows existing fixture pattern).

### 10.3 Calendar working=True exceptions (T137) — Phase 2a backlog
Phase 2a only supported `working=False` exceptions. COM uses
`pjExceptionWorking` (Type=4) for working-day overrides. Removes the
`working=True is not yet supported` guard, adds proper handling.

**Tests:** working=True single date, range, mixed working/non-working
in same call — 3 new.

### 10.4 Sample script unification (T138) — housekeeping
`samples/build_*.py` repeats: REPO_ROOT setup, `_write_xer` synthetic
XER, `asyncio.run` wrapper, `_print_section` helper. Extract to
`samples/_lifecycle_common.py`.

**Affected files:** build_currency_validation_lifecycle.py,
build_compare_lifecycle.py.

**Backward compat:** Each sample remains executable as
`python -m samples.build_X_lifecycle` (no top-level API change).

### 10.5 Test consolidation (T139) — housekeeping
Multiple integration test files build synthetic XERs with the same
template. Extract to `tests/_xer_fixture_builders.py`:
- `build_synthetic_xer(content_str, name) -> path`
- `make_cau_minimal_xer()` — 2 tasks, simple chain
- `make_staggered_xer()` — 3 tasks at different periods
- `make_early_finish_xer()` — same template as Phase 6.2 + 9.2

**Affected files:** test_msproject_evm_time_phased_ac_integration.py,
test_msproject_compare_dispatcher.py.

**Backward compat:** Tests behavior unchanged — only fixture creation
moves out. PASS count must remain 552.

## Tasks

| Task | Scope | Risk | Tests |
|---|---|---|---|
| T135 | update_task baseline read-back | low | +2 |
| T136 | Calendar recurring exceptions | medium (COM) | +4 (COM-skip) |
| T137 | Calendar working=True exceptions | medium (COM) | +3 (COM-skip) |
| T138 | Sample script unification | low (refactor) | 0 |
| T139 | Test consolidation | medium (refactor) | 0 (count preserved) |
| T140 | Commit + push + memory update | n/a | n/a |

## Backward compat

- All Phase 1-9 dispatcher actions preserved
- Sample script CLI usage unchanged
- Test PASS count unchanged after refactors (552 minimum)
- Phase 2a calendar adapter signature stays — recurring/working passed
  through `recurring=` and `working=` params (additive)

## Regression target

552 → ~561 PASS (+9 from 10.1+10.2+10.3; refactors zero-delta).
