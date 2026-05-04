# Phase 11 Follow-ups (Non-blocking polish)

> Tracks nice-to-have polish items surfaced by sub-phase code reviews
> that were not blocking and were intentionally deferred.
>
> Phase 11 final verdict: **READY TO MERGE** (verified 2026-05-04 by
> final code reviewer). These items are low-priority maintenance.

## Source

Each item below was raised by a code-reviewer subagent during Phase 11
sub-phase reviews and explicitly marked **non-blocking** /
**suggestion** / **APPROVED WITH SUGGESTIONS** by that reviewer.

---

## From Phase 11.1 review (commit `cbf1d9d`)

### F1. Three-way disjunction in `test_mspdi_parser_coverage.py:289`
```python
assert out_missing == {} or "error" in out_missing or out_missing is None
```
Pin the actual return contract. A three-way disjunction hides API
uncertainty.

### F2. Fallback date assertions in `test_mspdi_parser_coverage.py:389-404`
`test_add_task_with_invalid_start_date_falls_back` and
`_with_unparseable_finish_uses_start_dt` only assert `"task_id" in r`.
Also assert `r["start"]` is the fallback date to verify the branch did
its job.

### F3. Flat-plateau test in `test_evm_math.py:434-445`
`test_earned_schedule_with_flat_plateau_curve_works` docstring admits
line 280 isn't actually exercised. Either delete the test or construct
an input where `ev_now` lands inside the flat segment.

### F4. Mid-file import in `test_dcma_checks.py:737`
`import dcma_checks as _dcma_mod` mid-file — move to module-level
imports block.

### F5. Unused stash in `test_mspdi_parser_coverage.py:50`
`proj.__test_path__ = path` is harmless but unused (cleanup uses
closure capture). Remove.

### F6. Monkeypatch scope documentation in `test_mspdi_parser_coverage.py:1598-1618`
`test_save_handles_post_process_exception` monkeypatches `builtins.open`
globally. Document why the discriminator (`args[1] == "r"`) is
sufficient, or scope tighter.

---

## From Phase 11.2 review (commit `e7199c1`)

Important items already addressed in commit `c8c6db5` (Phase 11.5).

### F7. "Either ok or error" softened tests
- `test_link_dispatcher_chain_single_task_id_no_op`
- `test_link_dispatcher_chain_empty_task_ids_no_op`
- `test_task_dispatcher_bulk_add_empty_items_returns_error`

These accept both outcomes. Defensible (empty input is genuinely a
no-op) but loses a tooth — won't catch a regression that flips the
contract. Pin to canonical behavior with a comment.

### F8. COM-gated test fragility in marathon mode
COM tests don't `pytest.skip()` defensively when MSP unavailable;
they rely on `clean_test_project` fixture's auto-skip. Working as
designed, but marathon-mode 33-error pattern shows COM tests get
fragile in long runs. Consider `--reruns 1` for that subset (would
require `pytest-rerunfailures` dev dep).

### F9. `_T142` suffix documentation
Calendar/resource entity names use `_T142` suffix to avoid collision.
Worth documenting once at module-top instead of in 15 docstrings.

---

## From Phase 11.3 review (commit `913cc08`)

### F10. Scenario 8 contract extension
When `_msp_file_read_baselines` for MSPDI is extended to propagate
baseline_start/finish into `_evm_load_task_data`, replace the
`unchanged_count >= 1` check with `len(cmp["changed"]) >= 1` and
assert the changed entry references baseline_start. Currently
documented inline (Scenario 8 docstring lines 396-402).

---

## From Phase 11.4 review (commit `45baee5`)

Important items (zero-divisor properties) already addressed in `c8c6db5`.

### F11. Property 9 trivial generator
`test_property_xer_assignments_empty_uncertain` uses `st.just([])` —
single-value generator, runs only one example. Functionally correct
but adds zero fuzz coverage. Either drop `@given` (make it a regular
test) or broaden the strategy to include "empty-equivalent" inputs
(e.g., lists with all-None cost/qty fields).

### F12. Property 10 has no `@given`
`test_property_summarize_compare_no_changes_headline` is a single-
example test in a property-test file. Consider broadening to
`@given(st.text())` on `headline` field absence, or relabel as a
deterministic regression test.

---

## From Phase 11 final review (`c8c6db5` end-state)

### F13. dcma_checks 91% coverage rationale
A one-line comment in `.coveragerc` or a sentence in the design doc
clarifying which `dcma_checks.py` branches are intentionally excluded
would prevent a future reviewer from re-litigating the 91% number.

---

## Triage suggestion

| Priority | Items |
|---|---|
| **P3 (next maintenance pass)** | F1, F2, F3, F4, F5 — small test polish |
| **P3 (next maintenance pass)** | F7, F8, F9 — test contract pinning |
| **P3 (next maintenance pass)** | F11, F12 — property test broadening |
| **Deferred (depends on future feature)** | F10 — Scenario 8, gated on file MCP MSPDI baseline propagation |
| **Documentation** | F6, F13 — annotation only |

None of these block any current work. Address opportunistically.
