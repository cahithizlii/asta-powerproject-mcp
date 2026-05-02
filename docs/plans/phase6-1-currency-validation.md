# Phase 6.1 — Currency Mode Cross-Validation (T115)

> Polish task. Phase 5a (`_evm_detect_currency_mode`) genişletilir.
> RULE 3 ile direkt entegrasyon, multi-source cross-validation eklenir.

## Goal

Mevcut Phase 5a currency detection (sadece "cost"/"hours") yetersiz:
1. **RULE 3 pattern'i kontrol etmiyor** — XER `target_cost == target_qty` her satırda → cost loaded değil.
2. **Mixed mode yok** — kısmi cost loading durumu raporlanmıyor.
3. **Currency code yok** — XER ERMHDR.currency okunmuyor.
4. **Cross-validation yok** — task vs resource vs assignment source'ları çakışırsa görünmüyor.
5. **Pure-side eksik** — currency mantığı `msproject_mcp_core.py` içinde, Yaklaşım C ihlali.

## Deliverables

### Yeni dosyalar
- `currency_validator.py` — pure module (zero dependency)
- `tests/test_currency_validator.py` — pure module unit tests
- `tests/test_msproject_evm_currency_validation.py` — adapter + dispatcher integration

### Mevcut dosya değişiklikleri (additive only)
- `msproject_mcp_core.py`:
  - `_evm_detect_currency_mode` → delegate to pure module (geri uyumlu return)
  - Yeni `_msp_evm_validate_currency_mode` adapter
  - Dispatcher: yeni `validate_currency_mode` action
  - Mevcut `detect_currency_mode` action **korunur** (backward compat)

## Pure module API

```python
# currency_validator.py

def detect_mode_from_xer_assignments(assignments) -> str
    """RULE 3: target_cost vs target_qty pattern.
    All target_cost == target_qty -> 'hours' (not cost loaded)
    target_cost varies independently with target_qty > 0 -> 'cost'
    Some rows match, some don't -> 'mixed'
    No data -> 'uncertain'
    """

def detect_mode_from_tasks_resources(tasks, resources) -> str
    """Aggregate task.cost + resource.cost fields.
    All cost > 0 -> 'cost'
    All cost == 0 -> 'hours'
    Some > 0, some == 0 -> 'mixed'
    No valid data -> 'uncertain'
    """

def extract_currency_code(xer_header_fields) -> str | None
    """ERMHDR.currency field (3-letter code) or None."""

def cross_validate_modes(sources) -> dict
    """sources: [(source_name, mode), ...]
    Returns: {
        consensus_mode: 'cost'|'hours'|'mixed'|'uncertain',
        confidence: 'high'|'medium'|'low',
        conflicts: [(source_a, mode_a, source_b, mode_b)],
        warnings: [str],
        source_counts: {mode: count},
    }
    """
```

## Confidence model

| Source agreement | Confidence |
|---|---|
| All sources agree (1 mode, uncertain hariç) | high |
| Majority ≥66% agree | medium |
| Split / no majority | low |

`uncertain` sources counted out (filtered before consensus). Empty after filter → `(uncertain, low)`.

## Backward compat

`_evm_detect_currency_mode(tasks, resources)` mevcut return contract:
- Eski: "cost" / "hours" döner
- Yeni: pure module 4-mode döner, eski wrapper:
  - "mixed" → "cost" (cost data var, geri uyumlu)
  - "uncertain" → "hours" (no cost = same as no data)
- Eski test `test_msp_evm_currency_xml` **kırılmaz**.

`detect_currency_mode` dispatcher action **korunur** — output schema değişmez.

## Yeni dispatcher action

`validate_currency_mode`:
```json
{
  "status": "ok",
  "primary_mode": "cost",
  "currency_code": "USD",
  "cross_validation": {
    "consensus_mode": "cost",
    "confidence": "high",
    "conflicts": [],
    "warnings": [],
    "source_counts": {"cost": 3, "uncertain": 0}
  },
  "sources": {
    "tasks_resources": "cost",
    "xer_assignments": "cost",
    "currency_header": "USD"
  }
}
```

## Test plan

### Pure module (~22 test)
- `test_xer_assignments_all_target_cost_equals_qty_returns_hours`
- `test_xer_assignments_real_cost_returns_cost`
- `test_xer_assignments_mixed_returns_mixed`
- `test_xer_assignments_empty_returns_uncertain`
- `test_xer_assignments_zero_qty_skipped` + with cost
- `test_tasks_resources_*` (5 cases)
- `test_extract_currency_code_*` (4 cases)
- `test_cross_validate_*` (7 cases — all-agree, majority, split, uncertain-filtered, empty, warnings)

### Integration (~6 test)
- `test_msp_evm_validate_currency_mode_xml` — sample MSPDI sanity
- `test_msp_evm_validate_currency_mode_xer` — XER routing
- `test_msp_evm_validate_currency_mode_currency_code_extracted`
- `test_msp_evm_validate_currency_mode_cross_validation_present`
- `test_msp_evm_dispatcher_validate_currency_mode_action`
- `test_msp_evm_dispatcher_detect_currency_mode_backward_compat` (assert "cost"|"hours")

Hedef: ~28 yeni PASS. Baseline 413 → ~441.

## Tasks (T115a-d)

- **T115a**: pure module + 22 unit tests, all PASS isolated
- **T115b**: adapter `_msp_evm_validate_currency_mode` + 6 integration tests
- **T115c**: dispatcher `validate_currency_mode` action + 1 backward-compat test
- **T115d**: full suite verification + commit + push

## DOKUNULMAZ

- `_evm_detect_currency_mode` signature
- `_msp_evm_detect_currency_mode` adapter return shape
- Dispatcher `detect_currency_mode` action behavior
- Phase 4 file MCP read helpers
- Phase 5d xer_parser.py
