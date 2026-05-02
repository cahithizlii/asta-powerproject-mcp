# Phase 6.2 — True Per-Period AC Distribution (T116-T117)

> Polish task. Phase 5a `time_phased_evm` AC simplification eliminated.

## Goal

Phase 5a `_msp_evm_time_phased_evm` distributed total AC uniformly across
past buckets:

```python
total_ac = sum(actual_work)
ac_per_bucket = total_ac / past_buckets   # uniform leak
```

This violated RULE 5/6: a task that finished in week 2 was reported as
spending its actual_work spread across all weeks 1..N. The fix uses
per-task linear distribution across `actual_start`..`actual_finish`.

## Deliverables

### `evm_math.py` (additive)
- `_task_ac_at_date(task, eval_date) -> float` — single task AC at a date
- `time_phased_ac(tasks, buckets, data_date) -> List[float]` — cumulative
  AC at each bucket end, capped at data_date

### `msproject_mcp_core.py`
- Import `time_phased_ac as _evm_tp_ac`
- `_msp_evm_time_phased_evm` enriched dict gains `actual_start`,
  `actual_finish`, `actual_work` fields
- Uniform `total_ac / past_buckets` block REPLACED by `_evm_tp_ac` call

## Distribution rules

| Task state | Behavior |
|---|---|
| `actual_work == 0` | contributes 0 to all buckets |
| `actual_start` missing | falls back to `baseline_start` (preserves total) |
| `actual_finish` missing (in-progress) | full `actual_work` from `actual_start` onwards (work spent ≤ data_date) |
| `actual_finish` ≤ eval_date | full `actual_work` |
| `actual_start` > eval_date | 0 |
| Otherwise | linear: `aw × elapsed / duration` |

## Tests

### `tests/test_evm_math_time_phased_ac.py` (11 unit)
- 6 `_task_ac_at_date` cases (zero, completed, unstarted, in-progress,
  midpoint, baseline fallback)
- 5 `time_phased_ac` cases (cumulative, data_date cap, multi-task,
  in-progress, unstarted)

### `tests/test_msproject_evm_time_phased_ac_integration.py` (4 integration)
- Synthetic XER: AC monotonic non-decreasing
- Early-finish XER: plateau at total AC from bucket 1
- Staggered XER: stepwise growth + plateau (rejects uniform pattern)
- Final cum AC == sum(actual_work)

## Backward compat

- `_msp_evm_time_phased_evm` return shape unchanged
- `bucket=day/week/month` validation unchanged
- All 3 prior `test_msp_evm_time_phased_*` tests continue to PASS

## Regression

447 → 462 PASS (+15, zero regression). Full suite: 35s.
