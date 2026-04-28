# MS Project MCP — Phase 3a Baseline Design

**Versiyon:** 1.0
**Tarih:** 28 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → writing-plans next
**Phase 2b HEAD:** `6060e09` (origin/main, in sync, 156 PASS + 1 xfail)

---

## 1. Hedef

`msproject_mcp` server'ına `msproject_baseline` tool'u (6. tool) ekle. Geniş kapsamlı baseline desteği — MSP'nin 11 baseline slot'unun tamamı (Baseline + Baseline1..Baseline10), variance reporting, baseline-to-baseline karşılaştırma, RAG summary.

Phase 3a sonunda kullanıcı tek mesajla:

> "Bu schedule'ı 'Original' olarak baseline 0'a kaydet. Hafta 1 progress sonrası variance raporu çıkar. Change order sonrası revize schedule'ı baseline 1'e kaydet. Original ile revize arasındaki delta'yı göster."

talimatını verebilmeli.

## 2. Karar Geçmişi

### Q1 — Phase 3 yönü: C
Phase 3a Baseline only önce, Phase 3b Progress ayrı paket. Phase 2a/2b disiplini gibi (küçük ısırık + onay).

### Q2 — Baseline scope: Geniş kapsamlı
Original 4 action (save/clear/compare/list) yetersiz. 9 action: 4 original + clear_all + get_task_baseline + compare_two + summary + set_active. Multi-baseline (0-10) tam destek.

## 3. Tool Surface — `msproject_baseline` 9 Action

| # | Action | Parametreler | Çıktı |
|---|---|---|---|
| 1 | `save` | `baseline_number=0` (0-10), [`name`, `scope="all"\|"selected"`, `roll_up_to_summary=True`] | `{status, baseline_number, saved_date, task_count, total_duration_days, total_work_hours, total_cost}` |
| 2 | `clear` | `baseline_number=0` | `{status, baseline_number, was_saved_date}` |
| 3 | `clear_all` | (yok) | `{status, cleared: [int], count}` |
| 4 | `list` | (yok) | `{status, count_saved, baselines: [{number, name, saved_date, task_count, total_duration_days, total_work_hours, total_cost}]}` |
| 5 | `get_task_baseline` | `task_id`, `baseline_number=0` | `{status, task_id, baseline_number, baseline: {start, finish, duration_h, work_h, cost}}` |
| 6 | `compare` | `baseline_number=0`, [`include_unchanged=False`, `variance_threshold_days=0`] | `{status, summary: {slipped_count, ahead_count, on_time_count, total_start_drift_days, total_finish_drift_days, total_duration_var_h, total_work_var_h, total_cost_var}, tasks: [...]}` |
| 7 | `compare_two` | `baseline_a` (0-10), `baseline_b` (0-10), [filtreler] | Aynı compare çıktısı; B_a → B_b delta |
| 8 | `summary` | `baseline_number=0` | `{status, baseline_number, project: {start_drift_days, finish_drift_days, slipped_pct, schedule_health: "green"\|"amber"\|"red"}}` |
| 9 | `set_active` | `baseline_number` (0-10) | `{status, active_baseline}` |

## 4. Variance Hesabı

**Per task:**
- `start_var_days` = (current_start - baseline_start).days
- `finish_var_days` = (current_finish - baseline_finish).days
- `duration_var_h` = current_duration_minutes/60 - baseline_duration_minutes/60
- `work_var_h` = current_work_minutes/60 - baseline_work_minutes/60
- `cost_var` = current_cost - baseline_cost (locale-aware via `_parse_rate`)
- `status` =
  - `"slipped"` if finish_var_days > variance_threshold_days
  - `"ahead"` if finish_var_days < -variance_threshold_days
  - `"on_time"` otherwise

**RAG status logic (summary action):**
- `green`: slipped_pct ≤ 5%
- `amber`: 5% < slipped_pct ≤ 20%
- `red`: slipped_pct > 20%

## 5. Implementation Architecture

```
msproject_mcp_core.py
├── (existing Phase 1+2a+2b)
└── (NEW Phase 3a — BASELINE section)
    ├── BASELINE_NUMBERS = list(range(11))  # 0..10
    ├── _baseline_property_name(field, baseline_number) → "BaselineStart" or "Baseline3Start"
    ├── _read_task_baseline(task, baseline_number) → dict with start/finish/duration/work/cost
    ├── _baseline_saved_date(proj, baseline_number) → datetime or None
    ├── _msp_baseline_save / clear / clear_all / list
    ├── _msp_baseline_get_task_baseline
    ├── _msp_baseline_compare / compare_two / summary
    ├── _msp_baseline_set_active
    └── @mcp.tool msproject_baseline dispatcher (9 action)
```

**Reuse from previous phases:**
- `_validate_active_project` (Phase 1)
- `_find_task_by_id` (Phase 1)
- `_format_com_error` (T29)
- `_parse_rate` (T32 — for cost variance locale handling)
- `_build_task_id_map` (T37 — for compare 200+ task perf)
- `clean_test_project` fixture (Phase 1 SAFETY)

**Dynamic baseline property access** (DRY pattern):
```python
def _baseline_property_name(field: str, baseline_number: int) -> str:
    """Map (field='Start', baseline_number=0) → 'BaselineStart';
    (field='Start', baseline_number=3) → 'Baseline3Start'."""
    suffix = "" if baseline_number == 0 else str(baseline_number)
    return f"Baseline{suffix}{field}"
```

Then `getattr(task, _baseline_property_name("Start", N))` works for all 11 baselines uniformly.

## 6. COM API Reference

**Save / Clear:**
```python
proj.SaveBaseline(BaselineNumber=0)            # default: all tasks, summary roll-up
proj.SaveBaseline(BaselineNumber=3, FromTaskScope=2, FromAllOrSelected=2)  # selected only
proj.ClearBaseline(BaselineNumber=0)
```

**Saved date check (presence):**
```python
saved = proj.BaselineSavedDate(BaselineNumber=0)  # → datetime or 0/None if not saved
```

**Per-task baseline data (read-only):**
```python
task.BaselineStart, task.BaselineFinish, task.BaselineDuration, task.BaselineWork, task.BaselineCost
task.Baseline1Start, ..., task.Baseline10Cost  # (44 properties × 5 fields × 11 baselines)
```

**Set active (UI/views):**
- Investigation needed during impl — likely `app.OptionsCalculation` or view-level setting; may require `View.Baseline` per view. Worst case: skip set_active or return "not yet supported" for Phase 3a tail.

## 7. Test Stratejisi

**~14 yeni test:**
- `test_msproject_baseline_helpers.py` — `_baseline_property_name`, `BASELINE_NUMBERS` constant
- `test_msproject_baseline_save.py` — default save + numbered + scope=selected (3 tests)
- `test_msproject_baseline_clear.py` — clear single + clear_all (2 tests)
- `test_msproject_baseline_list.py` — empty / 1 saved / 3 saved with metadata (3 tests)
- `test_msproject_baseline_get_task.py` — happy + missing task + missing baseline (3 tests)
- `test_msproject_baseline_compare.py` — no progress 0 variance, with progress drift, threshold filter (3 tests)
- `test_msproject_baseline_compare_two.py` — B0 vs B1 delta + missing baseline (2 tests)
- `test_msproject_baseline_summary.py` — RAG green / amber / red (3 tests)
- `test_msproject_baseline_set_active.py` — happy + invalid number (or skip if API not exposed) (1-2 tests)
- `test_msproject_baseline_dispatcher.py` — 9 action routing + invalid action (4-5 tests)

**Total target:** ~28 new tests. Final regression: **184 PASS + 1 xfail** (156 baseline + 28 new).

**Performance hedefleri:**
- `compare` 200 task projede **<2s** (read-only O(N) iteration)
- `save` herhangi proje **<1s** (single COM call)
- `list` 11 baseline scan **<500ms**

## 8. Acceptance Script

`samples/build_baseline_lifecycle.py` — CAU/Akfa Medline benzeri full scenario:

1. 50 villa task oluştur (`msproject_task bulk_add`)
2. Resource ekleme + assign (mini Phase 2b chain)
3. **save baseline_number=0 "Original"**
4. Bazı task'lara progress gir (raw COM, T36 helper kullanılarak)
5. **compare(0)** → variance raporu yazdır
6. Bazı task duration'larını değiştir (revize plan simulation)
7. **save baseline_number=1 "Rev1-AfterChangeOrder"**
8. **compare_two(0, 1)** → revize delta yazdır
9. **summary(0)** → RAG status

Hedef: end-to-end **<10s**, MS Project UI'da Tools → Tracking → Baselines'da 2 baseline görünür olmalı.

## 9. Out of Scope (Phase 3a'da YOK — Phase 3b veya sonra)

- `msproject_progress` tool (Phase 3b)
- Time-phased baseline (per-period stored values) — Phase 4+
- Baseline cost rate tables — Phase 4+
- Custom baseline metadata fields — Phase 4+
- True MSPDI baseline merge — bulk save senaryolarında bile single COM call yeterli

## 10. Acceptance Kriterleri (Phase 3a Tamam)

1. ✅ `msproject_baseline` tool 9 action ile çalışıyor (T39-T47 + T48 dispatcher)
2. ✅ Acceptance script `samples/build_baseline_lifecycle.py` end-to-end <10s
3. ✅ Phase 3a yeni testleri (~28) PASS
4. ✅ Phase 1 + Phase 2a + Phase 2b mevcut 156+1xfail regression PASS — total **~184 PASS + 1 xfail**
5. ✅ Phase 1 SAFETY: kullanıcının aktif projesi DOKUNULMAZ
6. ✅ Commit + push GitHub'a
7. ⏸ Kullanıcı manuel onayı → Phase 3b (Progress) başlar

## 11. Plan Paketi

- T39: Baseline foundations (helpers + constant)
- T40: save action
- T41: clear + clear_all actions (paired)
- T42: list action
- T43: get_task_baseline action
- T44: compare action (variance calc + threshold filter — büyük task)
- T45: compare_two action
- T46: summary action
- T47: set_active action (investigate API)
- T48: FastMCP dispatcher + acceptance + README + push

10 task TDD plan, ~6-8 saat tahmin.

---

*Approved by user: 28 Nisan 2026*
*Next: writing-plans skill → Phase 3a Baseline implementation plan*
