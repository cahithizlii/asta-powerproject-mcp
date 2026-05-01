# MS Project MCP — Phase 5a EVM Design

**Versiyon:** 1.0
**Tarih:** 30 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → writing-plans next
**Phase 4 HEAD:** `44eb816` (origin/main, in sync, 80/80 file tests + Phase 1-3 regression)

---

## 1. Hedef

`msproject_mcp` server'ına `msproject_evm` tool'u (9. tool) ekle. Hibrit (file+COM) Earned Value Management implementation — CLAUDE.md RULE 4-9 + RULE 12 RAG + RULE 3 currency mode auto-detection. PMI PMBOK 8th § 7.4.2 forecasting + Lipke 2003 Earned Schedule + JSON snapshot history.

Phase 5a sonunda kullanıcı tek mesajla:

> "CAU project'in son durumu — week 17 itibariyle SPI ve CPI nedir? EAC₃ ne çıkıyor? Geçen haftaya göre period delta nasıl? Schedule health AMBER mi GREEN mi? Snapshot olarak kaydet."

talimatını verebilmeli. Hibrit: file_path verilmişse dosyadan okur, yoksa MSP COM aktif projeden alır.

## 2. Karar Geçmişi (5 Brainstorming Q&A + Mimari)

### Q1 — Phase 5 bölme: B (önerim B sırası)
Phase 5'i 5a/5b/5c'ye böl, sırayla **EVM → DCMA → Excel**. Phase 2a/2b, 3a/3b disiplini. EVM önce çünkü Phase 3b summary (BAC/ACWP/StatusDate) doğal devamı.

### Q2 — EVM action surface: B (Full PMBOK + history, 13 action)
Compute (4) + Time-phased (2) + Data quality (1) + Baseline integration (2) + History (3) + Mode detect (1) = **13 action**.

### Q3 — Snapshot storage: A (JSON file)
Default `~/msproject_evm_snapshots.json`, kullanıcı override edebilir. Persistent, version-controllable, debuggable.

### Q4 — Tool mimarisi: C (Hybrid)
Yeni `msproject_evm` 9. tool, `file_path` opsiyonel. Verilmezse Phase 1 COM path, verilirse Phase 4 file path. Math source-agnostic.

### Q5 — Acceptance senaryosu: C (Hero CAU-benzeri)
200 task CAU-style + 3 baseline + multi-week time-phased + 4 JSON snapshots + DCMA-ready summary. Hedef <30s wall clock.

### Mimari — Yaklaşım C (Math util ayrı, I/O+dispatcher core'da)
- Yeni dosya `evm_math.py` — saf RULE 4-9 algoritmaları, MSP/COM/file independent
- `msproject_mcp_core.py` Phase 5a section — adapters + helpers + dispatcher
- Phase 1-4 helpers DOKUNULMAZ

## 3. Tool Surface — `msproject_evm` 13 Action

### Compute / Snapshot (4)
| # | Action | Ne yapar |
|---|---|---|
| 1 | `compute_metrics` | RULE 4 — `{spi, cpi, sv, cv, bac, ev, ac, pv}` |
| 2 | `forecast` | RULE 9 — `{eac_t1, eac_t2, eac_t3, etc, vac, tcpi_bac, tcpi_eac}` |
| 3 | `earned_schedule` | RULE 8 Lipke 2003 — `{at, es, sv_t, spi_t}` |
| 4 | `summary` | RULE 12 — `{rag, completion_pct, schedule_health, executive_text}` |

### Time-Phased (2)
| # | Action | Ne yapar |
|---|---|---|
| 5 | `time_phased_evm` | bucket day/week/month per period PV/EV/AC |
| 6 | `period_delta` | RULE 6 — period delta vs prev snapshot |

### Data Quality (1)
| # | Action | Ne yapar |
|---|---|---|
| 7 | `progress_data_quality` | RULE 7 — SPI(h) vs SPI(t), missing actuals warnings |

### Baseline Integration (2)
| # | Action | Ne yapar |
|---|---|---|
| 8 | `variance_to_baseline` | EVM vs Baseline N (Phase 3a) |
| 9 | `compare_baselines_evm` | B0 → B1 EVM delta (Phase 3a compare_two pattern) |

### History (3 — JSON-backed)
| # | Action | Ne yapar |
|---|---|---|
| 10 | `save_period_snapshot` | snapshot_path'e current EVM dump |
| 11 | `get_period_history` | saved snapshots list (filter date range) |
| 12 | `trend` | period-over-period SPI/CPI/EAC trajectory series |

### Setup / Mode (1)
| # | Action | Ne yapar |
|---|---|---|
| 13 | `detect_currency_mode` | RULE 3 — hours vs cost loading |

### Tüm action'ların ortak parametreleri
- `file_path` (opsiyonel) — verilirse Phase 4 path; yoksa Phase 1 COM
- `baseline_number` (default 0) — Phase 3a 11-slot uyumluluk
- `bucket` (time-phased için: `day`/`week`/`month`)
- `snapshot_path` (history için, default `~/msproject_evm_snapshots.json`)

**Tool count: 8 → 9 (yeni `msproject_evm`).**

## 4. Architecture (Yaklaşım C)

```
evm_math.py (NEW — pure Python, MSP/COM/file independent)
├── compute_metrics(bac, pv, ev, ac) → SPI/CPI/SV/CV
├── forecast(bac, ev, ac, cpi, spi) → EAC₁/₂/₃, ETC, VAC, TCPI(BAC), TCPI(EAC)
├── earned_schedule(pv_curve, ev_now, project_start, data_date) → AT/ES/SV(t)/SPI(t)
├── time_phased_pv(tasks, buckets) → linear distribution per RULE 5
├── time_phased_ev(tasks, buckets, data_date) → cumulative EV at bucket end
├── period_delta(snap_now, snap_prev) → period_pv/ev/ac (RULE 6)
├── progress_data_quality(spi_h, spi_t, completion_pct, has_resources) → warnings list
└── rag_status(spi, completion_pct) → RED/AMBER/GREEN per RULE 12

msproject_mcp_core.py PHASE 5A SECTION
├── from evm_math import compute_metrics, forecast, ...
├── _evm_load_task_data(file_path: Optional[str]) → List[task_dict]
│      file_path → Phase 4 _msp_file_read_tasks
│      None     → Phase 1 COM iter
├── _evm_load_progress_data(file_path) → {tasks, status_date}
├── _evm_load_baseline_data(file_path, baseline_number)
├── _evm_detect_currency_mode(tasks, resources) → "hours"|"cost" (RULE 3)
├── _evm_build_pv_curve(tasks, baseline_number, bucket) → curve points
├── _evm_snapshot_save(snapshot_path, dict) — JSON append
├── _evm_snapshot_load(snapshot_path, filter) — JSON read+filter
├── _msp_evm_compute_metrics / forecast / earned_schedule / summary (4)
├── _msp_evm_time_phased_evm / period_delta (2)
├── _msp_evm_progress_data_quality (1)
├── _msp_evm_variance_to_baseline / compare_baselines_evm (2)
├── _msp_evm_save_period_snapshot / get_period_history / trend (3)
├── _msp_evm_detect_currency_mode (1)
└── @mcp.tool msproject_evm dispatcher (13 action routing)
```

**Reuse from previous phases:**
- Phase 1: `_validate_active_project`, `_format_com_error`, `_parse_rate`
- Phase 3a: `BASELINE_NUMBERS`, `_baseline_property_name`, `_read_task_baseline`
- Phase 3b: `_PROGRESS_*` constants, `_read_task_progress_dict`, `_minutes_to_hours`
- Phase 4: `_msp_file_read_tasks/progress/baselines/resources/assignments`, `_get_msp_file_manager`

**Phase 1+2a+2b+3a+3b+4 kodu DOKUNULMAZ.** Sadece read-only çağrılar.

## 5. Hybrid Data Source

```
_evm_load_task_data(file_path):
    file_path verilmiş?
    ├── EVET → _msp_file_read_tasks(file_path) + read_baselines + read_progress
    │           + read_resources + read_assignments
    └── HAYIR → _validate_active_project() → proj
                proj.Tasks iterate → task dict
                proj.Resources → resource dict
                proj.StatusDate → data_date
```

Her iki path **aynı şema** döner. EVM math source-agnostic.

**Auto-sync:** EVM tool **read-only** (sadece save_period_snapshot JSON yazar). Phase 4 auto-sync kuralı dışında.

## 6. EVM Math (RULE 4-9) — Implementation

### compute_metrics (RULE 4)
```python
spi = ev / pv if pv > 0 else None
cpi = ev / ac if ac > 0 else None
sv  = ev - pv      # negative = behind schedule
cv  = ev - ac      # negative = over budget
```

### forecast (RULE 9 — PMI PMBOK 8th § 7.4.2)
```python
eac_t1 = ac + (bac - ev)
eac_t2 = bac / cpi      if cpi > 0 else None
eac_t3 = ac + (bac - ev) / (cpi * spi) if cpi > 0 and spi > 0 else None
etc    = eac - ac
vac    = bac - eac
tcpi_bac = (bac - ev) / (bac - ac) if (bac - ac) > 0 else None
tcpi_eac = (bac - ev) / (eac - ac) if eac and (eac - ac) > 0 else None
```

### time_phased_pv (RULE 5 — linear distribution)
```python
def task_pv_at_date(task, eval_date):
    bs, bf = task['baseline_start'], task['baseline_finish']
    bw = task['baseline_work']
    if bf <= eval_date: return bw
    if bs >= eval_date: return 0.0
    return bw * (eval_date - bs).days / max((bf - bs).days, 1)
```

### earned_schedule (RULE 8 — Lipke 2003)
PV curve üzerinde linear interpolation: t where cumulative PV(t) = current EV.
```python
es_weeks = ...  # interpolated date in weeks since project_start
at_weeks = (data_date - project_start).days / 7.0
sv_t  = es_weeks - at_weeks
spi_t = es_weeks / at_weeks
```

### period_delta (RULE 6)
```python
period_pv = cum_pv_now - cum_pv_prev
period_ev = cum_ev_now - cum_ev_prev
period_ac = cum_ac_now - cum_ac_prev
period_bac = 0   # BAC sabit
```

### rag_status (RULE 12)
```python
if spi is None or completion_pct == 0: return "RED"
if spi < 0.3:  return "RED"
if spi < 0.7:  return "AMBER"
return "GREEN"
```

### progress_data_quality (RULE 7) — warning rules
- SPI(h) vs SPI(t) divergence: `abs(spi_h - spi_t) > 0.15` → "EV input quality concern"
- Tasks with %complete > 0 but actual_work = 0 → Phase 3b silent EV pattern
- Tasks past data_date without actuals → late progress entry
- Resource-less progress → missing assignments

## 7. JSON Snapshot Schema

```json
{
  "snapshots": [
    {
      "id": "20260430-1330",
      "saved_at": "2026-04-30T13:30:00",
      "project_name": "CAU Construction",
      "project_file": "C:/.../cau.xml",
      "data_date": "2026-04-30",
      "baseline_number": 0,
      "currency_mode": "hours",
      "metrics": {
        "bac": 5058787, "pv": 1200000, "ev": 980000, "ac": 1100000,
        "spi": 0.817, "cpi": 0.891, "sv": -220000, "cv": -120000,
        "completion_pct": 19.4
      },
      "forecast": {
        "eac_t1": 5178787, "eac_t2": 5677654, "eac_t3": 5945120,
        "etc": 4078787, "vac": -119000, "tcpi_bac": 1.04, "tcpi_eac": 0.95
      },
      "earned_schedule": {
        "at": 42.0, "es": 38.5, "sv_t": -3.5, "spi_t": 0.917
      },
      "rag": "AMBER",
      "tag": "week-17"
    }
  ]
}
```

**Storage convention:** default `~/msproject_evm_snapshots.json`, override via `snapshot_path`. Append-only — `save_period_snapshot` reads, appends, writes back. `get_period_history` filters (date range, baseline_number, project_name).

## 8. Test Stratejisi

**~50 yeni test:**
- `test_evm_math.py` — saf math (RULE 4-9), no fixtures (~25 test):
  - SPI/CPI/SV/CV edge cases (PV=0, AC=0)
  - Forecast formulas (EAC₁/₂/₃, TCPI signs)
  - earned_schedule interpolation
  - time_phased_pv / time_phased_ev linear distribution
  - period_delta deltas
  - rag_status threshold boundaries
  - progress_data_quality warning rules
- `test_msproject_evm_loader.py` — `_evm_load_task_data` both paths (~6 test)
- `test_msproject_evm_compute.py` — compute_metrics dispatcher (~3 test)
- `test_msproject_evm_forecast.py` — forecast dispatcher (~3 test)
- `test_msproject_evm_earned_schedule.py` — ES dispatcher (~2 test)
- `test_msproject_evm_time_phased.py` — bucket day/week/month (~4 test)
- `test_msproject_evm_period_delta.py` — period delta (~2 test)
- `test_msproject_evm_baseline.py` — variance_to_baseline + compare (~3 test)
- `test_msproject_evm_snapshot.py` — JSON save/load/filter (~4 test)
- `test_msproject_evm_dispatcher.py` — 13 action routing (~5 test)
- `test_msproject_evm_dataquality.py` — RULE 7 warnings (~3 test)

**Total target:** **~330+ PASS + 0 xfail** (Phase 4 80/80 + Phase 1-3 regression + ~50 new T75-T84).

**Performance hedefleri:**
- compute_metrics on 1000 tasks → **<200ms**
- time_phased_evm 12 monthly buckets, 1000 tasks → **<500ms**
- save_period_snapshot 4-entry JSON → **<50ms**

## 9. Acceptance Script — `samples/build_evm_lifecycle.py`

```
1. FileNew + 200 task CAU-style + 14 CAU resources + assignments (Phase 1+2b)
2. Save Baseline 0 (Phase 3a — Original)
3. Slip + revize, save Baseline 1 (Rev1)
4. Phase 3b: %30, %60 progress (week 1-2 simulation)
5. set_status_date "week 1"
6. msproject_evm: compute_metrics + forecast + ES + summary
7. save_period_snapshot (week-1)
8. Daha progress + set_status_date "week 2"
9. msproject_evm: compute_metrics + period_delta vs week-1
10. save_period_snapshot (week-2)
11. Devam (week 3 + 4) → 4 snapshot total
12. trend → SPI/CPI trajectory
13. variance_to_baseline 0 vs 1 (revize impact)
14. progress_data_quality → warnings list
15. compare_baselines_evm B0 vs B1
16. detect_currency_mode → "hours" (CAU pattern, RULE 3)
```

**Hedef:** <30s wall clock, 4 snapshot history, JSON file ~5KB. Phase 1 SAFETY pattern (FileNew+FileClose 0).

## 10. Out of Scope (Phase 5a'da YOK)

- DCMA 14-Point validate → Phase 5b `msproject_health`
- Excel export → Phase 5c `msproject_excel`
- Non-linear PV distribution (Beta curve, S-curve) → Phase 6+ polish
- Time-phased baseline (per-period stored values) — Phase 6+
- Concurrent-write file lock → Phase 6+ polish
- Live SPI/CPI dashboard websocket → out of scope (MCP stdio only)

## 11. Acceptance Kriterleri (Phase 5a Tamam)

1. ✅ `msproject_evm` tool 13 action ile çalışır (T75-T84)
2. ✅ Acceptance script `samples/build_evm_lifecycle.py` <30s
3. ✅ `evm_math` saf math test'leri ~25 PASS (no fixtures, no COM)
4. ✅ Phase 4 file path + Phase 1 COM path her ikisi end-to-end test edilir
5. ✅ Phase 1+2a+2b+3a+3b+4 mevcut regression PASS — DOKUNULMAZ
6. ✅ Total ~330+ PASS + 0 xfail
7. ✅ Snapshot JSON CAU-scale 4 entry, <5KB
8. ✅ Currency mode auto-detect (RULE 3 hours/cost)
9. ✅ RAG status (RULE 12) ve forecasting (RULE 9) executive output
10. ✅ Commit + push GitHub'a
11. ⏸ Kullanıcı manuel onayı → Phase 5b (DCMA) başlar

## 12. Plan Paketi (T75-T84, ~12 task TDD chain)

| Task | İçerik | ~Süre |
|---|---|---|
| **T75** | `evm_math.py` foundations: compute_metrics + forecast + rag_status | 2h |
| **T76** | `evm_math.py` time-phased: time_phased_pv + time_phased_ev + period_delta | 2h |
| **T77** | `evm_math.py` earned_schedule (Lipke linear interp) + progress_data_quality | 2h |
| **T78** | `_evm_load_*` adapters (hybrid file + COM data source) | 2h |
| **T79** | `_msp_evm_compute_metrics` + `forecast` + `summary` action helpers | 1h |
| **T80** | `_msp_evm_earned_schedule` + `data_quality` + `detect_currency_mode` | 1h |
| **T81** | `_msp_evm_time_phased_evm` + `period_delta` (BIG ONE — bucket logic) | 3h |
| **T82** | `_msp_evm_variance_to_baseline` + `compare_baselines_evm` (Phase 3a integration) | 2h |
| **T83** | Snapshot helpers + 3 history actions (save/get_history/trend) | 2h |
| **T84** | FastMCP dispatcher + acceptance script + README + push | 3h |

**Toplam:** ~20 saat impl, ~12-15 commit (T75-T84 + olası fix commits).

**Pattern:**
- T75-T77 saf math → manuel write + self-verify (no probe, test-driven)
- T81 BIG ONE (time-phased + bucket edge cases), T83 (JSON history) → subagent dispatch
- T82 Phase 3a integration → subagent (compare_two pattern probe)
- Phase 1+2+3+4 helpers DOKUNULMAZ; sadece `evm_math.py` (yeni) + Phase 5a section

---

*Approved by user: 30 Nisan 2026*
*Next: writing-plans skill → Phase 5a EVM implementation plan*
