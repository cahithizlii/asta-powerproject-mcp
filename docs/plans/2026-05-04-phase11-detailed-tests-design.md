# Phase 11 — Detailed Test Suite Expansion (Design Doc)

> **Status:** Design approved 2026-05-04. Implementation plan to follow
> via writing-plans skill.
> **Scope:** All 10 prior phases (1-10), all 13 tools.
> **Strategy:** Sub-phased (Approach 2) — 4 atomic commits, A→B→C→D.

## Goal

Existing suite is wide (851 collected, 569 default regression) but
uneven in depth: pure modules well-covered, dispatcher error paths
sparse, no measurable coverage metric, no property-based invariants,
no end-to-end multi-tool workflows.

Phase 11 closes those gaps in 4 atomic sub-phases.

## Constraints (from approval)

- **Coverage target: 98% line coverage** on test-eligible code
  (COM-only branches excluded via `# pragma: no cover` or
  `.coveragerc`)
- **`hypothesis` dependency:** approved for Phase 11.4
- **Approach:** sub-phased — 4 commits, each independently shippable

---

## Sub-phase 11.1 — Coverage Gap-Fill

### Tooling
- New dev dependency: `pytest-cov`
- New file: `.coveragerc` — excludes COM-only modules / branches
- Exclude criteria:
  - `# pragma: no cover` markers on COM-only branches
  - File-level exclude for pure-COM helpers (`_validate_active_project`,
    raw `Application.ActiveProject` calls)
  - `if __name__ == "__main__":` blocks
- Include scope:
  - **Pure modules** (highest priority): `evm_math.py`,
    `currency_validator.py`, `dcma_checks.py`, `xer_parser.py`,
    `xer_compare.py`, `mspdi_parser.py`
  - **Adapter helpers** (file MCP, EVM/DCMA/Excel loaders, XER
    shape adapters)

### Process
1. Install `pytest-cov`, write `.coveragerc`
2. Baseline: `pytest --cov=. --cov-report=term-missing` — capture %
3. Identify gaps via term-missing report
4. Write tests targeting uncovered branches
5. Re-measure until pure modules ≥ 98%

### Output
- ~30-50 new tests in existing test files
- `.coveragerc`
- Suite: 569 → ~610

---

## Sub-phase 11.2 — Edge Case + Negative Path

### Per-tool focus

| Tool | Negative test priorities |
|---|---|
| `msproject_task` | invalid task_id, malformed duration string, summary update via leaf API, max-int IDs |
| `msproject_link` | self-loop, type='XX', non-existent task_id, FS+lag overflow |
| `msproject_calendar` | bad date format, day_of_week typo, recurring without dates, working_hours bad HH:MM |
| `msproject_resource` | duplicate name, negative units, type='XX', missing resource_id on assign |
| `msproject_baseline` | invalid baseline_number, missing source, double-clear |
| `msproject_progress` | pct > 100, negative actual_work, future actual_start |
| `msproject_file` | corrupted XML, .pp/.xer rejection, partial schema, file_path None |
| `msproject_evm` | missing baseline, zero BAC, divide-by-zero in SPI/CPI, status_date in future |
| `msproject_health` (DCMA) | empty link list, all-summary tasks, no critical path |
| `msproject_excel` | unwritable path, missing fixtures, output ext mismatch |
| `msproject_xer` | malformed ERMHDR, missing tables, encoding fallback fail |
| `msproject_compare` | identical files, one-side empty, missing baseline_number |

### Convention
- Place new tests in existing per-tool test files (no new files needed)
- Test name pattern: `test_<action>_<edge_case>_<expected>`
- Each error returns `{status: 'error', error: <message>}` —
  assert both shape and key substring

### Output
- ~80 new tests, distributed across existing per-tool files
- Suite: 610 → ~690

---

## Sub-phase 11.3 — End-to-End Acceptance Scenarios

### Workflows (one test function per scenario, ~1-2 assertions)

1. **CAU monthly hakediş chain** — XER A → XER B → `monthly_report` →
   assert headline + counts + EVM deltas + Excel exists if path given
2. **Baseline lifecycle** — MSPDI parse → `write_baseline` → save →
   re-read via `read_baselines` → assert variance from baseline
3. **Currency validation pipeline** — XER load →
   `validate_currency_mode` → assert primary_mode + currency_code +
   cross_validation consensus
4. **Time-phased EVM full chain** — load → compute_metrics →
   time_phased_evm → period_delta → assert pv/ev/ac coherence +
   ac_increment sums
5. **DCMA + EVM combined health** — `assess_all` (DCMA) +
   `progress_data_quality` (EVM) consistency
6. **Calendar recurring + assignment effect** — create calendar →
   add_exception(weekly mon-fri off) → assign to task → assertion
7. **Update_task baseline awareness** — update with baseline_* fields
   → assert baseline_after present + values match input (Phase 10.1)
8. **Compare with baseline write back** — original XER →
   `write_baseline` to MSPDI → `compare task_delta` →
   diff captures baseline change

### File location
- New file: `tests/test_phase11_e2e_acceptance.py`

### Output
- ~10-15 test functions, suite: 690 → ~705

---

## Sub-phase 11.4 — Property-Based / Fuzz

### Tooling
- New dev dependency: `hypothesis`

### Properties (all on pure modules)

1. `time_phased_ac(tasks, buckets, dd)` — cumulative monotonic
   non-decreasing for any tasks/buckets
2. `time_phased_ac_increments` — sum(increments) ==
   cumulative[-1] (round-trip)
3. `diff_tasks(a, b)` — partition invariant:
   `len(added) + len(removed) + len(changed) + unchanged_count ==
   |union of ids|`
4. `diff_links` — symmetric: `diff(A, B).removed ==
   diff(B, A).added` (set-equal)
5. `cross_validate_modes(sources)` — idempotent:
   calling twice → same result
6. `extract_currency_code(d)` — None-safe for arbitrary dict input
   (no exception)
7. `_evm_compute(bac, pv, ev, ac)` — when bac>0,pv>0,ev>=0,ac>=0:
   SPI/CPI are finite numbers
8. `_parse_clndr_data(blob)` — any string input → no exception,
   list output
9. `currency_validator.detect_mode_from_xer_assignments` —
   empty input → "uncertain" (deterministic)
10. `xer_compare.summarize_compare` — empty diffs →
    "no changes detected" headline

### File location
- New file: `tests/test_phase11_properties.py`

### Output
- ~15-20 hypothesis-based tests, suite: 705 → **~720**

---

## Total Phase 11 estimate

| Sub-phase | New tests | Suite end | Risk |
|---|---|---|---|
| 11.1 Coverage gap-fill | ~40 | ~610 | Low |
| 11.2 Edge cases | ~80 | ~690 | Low |
| 11.3 E2E acceptance | ~12 | ~702 | Medium (multi-tool deps) |
| 11.4 Property-based | ~18 | ~720 | Medium (Hypothesis learning) |
| **Total** | **~150** | **~720** | — |

---

## Backward compatibility

- Zero changes to production code (pure addition of tests + dev deps)
- All 569 prior PASS tests preserved unchanged
- New tools NOT added (test-only phase)
- DOKUNULMAZ contracts honored across Phase 1-10

---

## DOKUNULMAZ (preserved through Phase 11)

- All 13 tool dispatcher action sets
- Pure module public API (`evm_math.*`, `currency_validator.*`,
  `dcma_checks.*`, `xer_parser.*`, `xer_compare.*`,
  `mspdi_parser.*`)
- Phase 4-10 file MCP write paths
- Phase 7+8 compare adapters

---

## Next step

Implementation plan via writing-plans skill — atomic per-sub-phase,
verifiable per commit.
