# Phase 6.3 — MSPDI Baseline Read/Write Extension (T118-T119)

> Polish task. `MspdiProject` had `save()` infrastructure but no
> baseline parse/write logic.

## Goal

Pre-Phase 6.3 mspdi_parser.py status:
- 1865 lines, parses tasks/links/resources/calendars/code libraries
- `save()` writes XML tree + namespace post-processing
- **Zero baseline support** — neither read nor write

Phase 6.3 adds minimum viable read+write+roundtrip:
- `read_baselines(baseline_number=0) -> List[dict]`
- `write_baseline(baseline_number, baseline_data) -> int`
- `save()` (existing) preserves the new Baseline elements

## Deliverables

### `mspdi_parser.py` (additive — class extension)
- `MspdiProject.read_baselines(baseline_number)` — parses
  `<Task><Baseline>` children matching the requested Number
- `MspdiProject.write_baseline(baseline_number, data)` — creates or
  updates `<Baseline Number=N>` per task UID

### `tests/test_mspdi_baseline_write.py` — 7 tests
- read_baselines empty on baseline-less source
- write_baseline creates new element
- write_baseline updates existing (no duplication)
- write_baseline + save + reload roundtrip preserves values
- unknown UID skipped silently
- save preserves task count (no data loss)
- multiple baseline numbers independent

## MSPDI Baseline element schema

```xml
<Task>
  <UID>1</UID>
  <ID>1</ID>
  ...
  <Baseline>
    <Number>0</Number>
    <Start>2026-01-01T08:00:00</Start>
    <Finish>2026-01-31T17:00:00</Finish>
    <Duration>PT240H0M0S</Duration>
    <Work>PT160H0M0S</Work>
  </Baseline>
</Task>
```

Number=0 is the primary baseline. Numbers 1-10 are numbered baselines.

## Backward compat

- `MspdiProject` constructor unchanged
- `save()` unchanged (extends to baseline elements automatically since
  they're in `self.tree`)
- All 1865 lines of pre-existing code preserved
- Phase 4 file MCP `_msp_file_read_baselines` adapter NOT modified
  (separate code path for now — direct integration deferred)

## Distribution rules

| Field | Required | Format | Notes |
|---|---|---|---|
| `task_uid` | yes | int | UID from MSPDI parse |
| `baseline_start` | optional | ISO 8601 | only set if provided |
| `baseline_finish` | optional | ISO 8601 | only set if provided |
| `baseline_duration_h` | optional | hours float | converted to PT{H}H{M}M0S |
| `baseline_work_h` | optional | hours float | converted to PT{H}H{M}M0S |

## Tests

7 tests, ~0.17s. All paths exercise real ElementTree + filesystem
roundtrip (no mocks).

## Regression

462 → 469 PASS (+7, zero regression).
