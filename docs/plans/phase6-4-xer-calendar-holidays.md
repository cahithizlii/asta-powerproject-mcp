# Phase 6.4 — XER Calendar Holiday Detail (T120-T122)

> Polish task. xer_parser.py read_calendars previously dropped the
> clndr_data BLOB; Phase 6.4 best-effort extracts exception/holiday
> dates.

## Goal

Phase 5d xer_parser.py comment:
> "clndr_data BLOB (holiday detail) NOT extracted - Phase 6 enhancement."

That deferral lands here. Primavera P6 stores per-calendar exception
dates inside the `clndr_data` text BLOB on each CALENDAR row. The
format is proprietary (no public spec), but a single regex pattern
covers the exception block reliably.

## Approach

### Pattern
P6 emits exception entries as:

```
(0||Exceptions
  (0||1(d|<excel_serial>|f|<bit>))
  (0||2(d|<excel_serial>|f|<bit>))
  ...
)
```

- `<excel_serial>` — Excel-style date serial with the Lotus quirk
  preserved (epoch base = 1899-12-30, so serial 2 = 1900-01-01)
- `<bit>` = 0 means non-working day (holiday); 1 means working-day
  override

### Implementation

`_parse_clndr_data(blob)` runs a single regex
`\(d\|(\d+)\|f\|(\d)` over the BLOB and yields
`{date: 'YYYY-MM-DD', working: bool}` per match. Tolerant of
unparseable serials, missing block, empty/None input.

`read_calendars` adds `exceptions: List[dict]` to each calendar dict.
All prior fields (`id`, `name`, `day_hr_cnt`, `week_hr_cnt`)
unchanged.

## Deliverables

### `xer_parser.py`
- Module-level `_parse_clndr_data(blob) -> List[dict]`
- `read_calendars` body extended with `exceptions` field
- New `import datetime as _dt`, `import re as _re`,
  `_CLNDR_EXCEPT_RE`, `_CLNDR_EXCEL_EPOCH` constants

### `tests/test_xer_calendar_holidays.py` (9 tests)
- 7 unit tests covering empty input, no-block, single holiday,
  single working override, multiple, Excel serial verification,
  invalid serial skip
- 2 integration tests via `XerFile.read_calendars()` (with and
  without `clndr_data` column; preserves prior fields)

## Backward compat

- `read_calendars` return shape **adds** `exceptions` (additive,
  prior keys preserved)
- All 45 prior `test_xer_parser.py` tests continue to PASS
- Phase 5e XER routing through `_evm_load_task_data` unaffected
- Phase 5f file MCP `_msp_file_read_calendars` reads the same shape
  via `_xer_to_calendar_shape` (Phase 6.4 adds field, no break)

## Tests

9 tests, ~0.21s. Inline synthetic XER fixtures (no external file)
for predictable BLOB content.

## Regression

469 → 478 PASS (+9, zero regression).

---

## Phase 6 closure summary

| Sub-phase | Tasks | Tests added | Cum PASS | Commit |
|---|---|---|---|---|
| 6.1 currency cross-validation | T115a-d | +34 | 447 | e6f8a1e |
| 6.2 per-task AC distribution | T116-T117 | +15 | 462 | 2d240fb |
| 6.3 MSPDI baseline read/write | T118-T119 | +7 | 469 | a9d1d81 |
| 6.4 XER calendar holidays | T120-T122 | +9 | 478 | (pending) |

Total Phase 6: **+65 PASS** (413 → 478), zero regression across all
4 sub-phases. Suite runtime ~32s.

### What changed (high level)
- Currency mode detection: 2-mode → 4-mode + cross-validation +
  ERMHDR currency code extraction
- Time-phased AC: uniform total/N → per-task linear distribution
- MSPDI: zero baseline support → read/write/roundtrip
- XER calendar: header-only → header + exception/holiday list

### What did NOT change
- Tool count (still 12)
- Dispatcher action count delta: +1 (validate_currency_mode)
- Phase 1-5f core logic — all DOKUNULMAZ contracts honored
- Backward compat verified via:
  - `_evm_detect_currency_mode` (legacy 2-mode return)
  - `dispatcher.detect_currency_mode` action shape
  - `_msp_evm_time_phased_evm` return schema
  - `MspdiProject` constructor + save() unchanged
  - `XerFile.read_calendars` keys preserved
