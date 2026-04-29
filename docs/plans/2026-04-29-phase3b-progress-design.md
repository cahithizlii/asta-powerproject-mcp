# MS Project MCP — Phase 3b Progress Management Design

**Versiyon:** 1.0
**Tarih:** 29 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → writing-plans next
**Phase 3a HEAD:** `3567ed1` (origin/main, in sync, 207 PASS + 1 xfail)

---

## 1. Hedef

`msproject_mcp` server'ına `msproject_progress` tool'u (7. tool) ekle. Geniş kapsamlı progress yönetimi — task-level + assignment-level (per-resource man-hour) çift-yollu ilerleme, time-phased (per-period) actual_work, status date, hibrit bulk path ve EVM-ready summary.

Phase 3b sonunda kullanıcı tek mesajla:

> "21 Nisan haftasının fiili saatlerini yükle: T101 için COW=24h, STL=18h, MSN=10h. Kalan tasklara `set_progress_by_date` ile haftaya kadar plan=actual progress yaz. Status date'i 27 Nisan'a al. Sonra `summary` çıkar — total ACWP, project % complete, BAC vs work-completed."

talimatını verebilmeli. Phase 3b'nin `summary` action'ı, Phase 5'te gelecek `msproject_evm` tool'unun temelini oluşturur (CLAUDE.md RULE 4-9).

## 2. Karar Geçmişi (4 Soru — Hepsi YES)

### Q1 — Dual-track (task + assignment): ✅ YES
Phase 1 `task.PercentComplete` minimal; hakediş için per-resource man-hour gerekiyor. **Hem task-level (`set_task_progress`) hem assignment-level (`set_assignment_progress`)** ayrı action'lar olarak. Roll-up: assignment.ActualWork → MSP otomatik task.ActualWork'a roll up eder; tersi de geçerli. Çakışma riski (aynı task'a ikisi yazılırsa MSP'nin "last-write-wins" davranışı — dokümante edilecek).

### Q2 — Time-phased dahil mi (Phase 3c'ye ertelemek yerine): ✅ YES
`time_phased_actual_write` ve `_read` Phase 3b'de. Granülarite: default `day` (`pjTimescaleDays=8`), opt-in `week` (`pjTimescaleWeeks=6`). Yazım/okuma `assignment.TimeScaleData(StartDate, EndDate, pjAssignmentTimescaledActualWork, PjTimescaleUnit)` koleksiyonu üstünden. Bu, hakediş raporları (CLAUDE.md RULE 6 — period EV/AC delta) için kritik.

### Q3 — Hybrid bulk path (Phase 2b T37 paterni): ✅ YES
`bulk_progress_update` → 1-5 com_direct, 6-19 com_batch, 20+ mspdi_bulk (Phase 2b'deki gibi com_batch_fallback). Pre-built ID maps (`_build_task_id_map`) + `_route_operation()`. T37 dersleri (cache once, reuse, no O(N×M)) korunur.

### Q4 — DCMA `PhysicalPercentComplete` Phase 3b'de mi (Phase 5 EVM'e ertelemek yerine): ✅ YES
`task.PhysicalPercentComplete` `set_task_progress` + `get_task_progress`'in alanı. Otomatik `% complete`'ten türetilmez — kullanıcı DCMA semantiği gereği (örn. structural earthwork %50 yapıldı ama duration %30 geçti) açıkça verir. Phase 5 EVM tool'u bu alanı EV hesabında kullanır.

## 3. Tool Surface — `msproject_progress` 12 Action

| # | Action | Parametreler | Çıktı |
|---|---|---|---|
| 1 | `set_task_progress` | `task_id`, [`percent_complete`, `percent_work_complete`, `actual_start`, `actual_finish`, `actual_duration_h`, `actual_work_h`, `remaining_work_h`, `remaining_duration_h`, `physical_pct`, `stop`, `resume`] | `{status, task_id, changes: [field], readback: {...}}` |
| 2 | `get_task_progress` | `task_id` | `{status, task_id, progress: {percent_complete, percent_work_complete, actual_start, actual_finish, actual_duration_h, actual_work_h, remaining_work_h, remaining_duration_h, physical_pct, stop, resume}}` |
| 3 | `set_assignment_progress` | `task_id`, `resource_id`, [`actual_work_h`, `actual_start`, `actual_finish`, `percent_work_complete`, `remaining_work_h`, `units`] | `{status, task_id, resource_id, changes: [field]}` |
| 4 | `get_assignment_progress` | `task_id` | `{status, task_id, assignments: [{resource_id, resource_name, actual_work_h, percent_work_complete, remaining_work_h, units}]}` |
| 5 | `set_progress_by_date` | `progress_date` (ISO), [`scope="all"\|"selected"`, `as_scheduled=True`] | `{status, progress_date, mode, task_count_affected}` |
| 6 | `set_status_date` | `status_date` (ISO) | `{status, status_date, previous}` |
| 7 | `clear_progress` | `task_id` | `{status, task_id, cleared_fields: [field]}` |
| 8 | `clear_all_progress` | (yok) | `{status, cleared_count}` |
| 9 | `time_phased_actual_write` | `task_id`, `resource_id`, `periods: [{start, end, actual_work_h}]`, [`unit="day"\|"week"`] | `{status, written_count, failures: []}` |
| 10 | `time_phased_actual_read` | `task_id`, `resource_id`, `start_date`, `end_date`, [`unit="day"\|"week"`] | `{status, periods: [{start, end, actual_work_h}]}` |
| 11 | `bulk_progress_update` | `items: [{task_id, percent_complete?, actual_work_h?, ...}]` | `{status, path, count, updated, failures}` |
| 12 | `summary` | (yok) | `{status, project: {bac_h, acwp_h, total_actual_work_h, total_remaining_work_h, project_percent_complete, status_date, task_count, in_progress_count, completed_count, not_started_count}}` |

## 4. KEY COM API REFERENCE

Verified from `msproject_typelib.txt`. Implementer T52'de probe ile teyit eder.

### Task-level progress fields
| Property | Read/Write | Notes |
|---|---|---|
| `task.PercentComplete` | RW | duration-based 0-100 |
| `task.PercentWorkComplete` | RW | work-based 0-100 |
| `task.ActualStart` | RW | datetime |
| `task.ActualFinish` | RW | datetime; setting marks task 100% |
| `task.ActualDuration` | RW | minutes |
| `task.ActualWork` | RW | minutes |
| `task.RemainingWork` | RW | minutes |
| `task.RemainingDuration` | RW | minutes |
| `task.PhysicalPercentComplete` | RW (probe in T52) | DCMA EV input — independent of % complete |
| `task.Stop` | RW | last "stop" date for in-progress task |
| `task.Resume` | RW | next resume date |

### Assignment-level fields
| Property | Read/Write | Notes |
|---|---|---|
| `task.Assignments(i)` | R | 1-indexed collection |
| `assignment.ResourceID` | R | maps back to resource |
| `assignment.Resource` | R | full resource object |
| `assignment.ActualWork` | RW | minutes |
| `assignment.ActualStart` | RW | datetime |
| `assignment.ActualFinish` | RW | datetime |
| `assignment.PercentWorkComplete` | RW | 0-100 |
| `assignment.RemainingWork` | RW | minutes |
| `assignment.Units` | RW | float (e.g. 1.0 = 100%) |

### Time-phased
```python
# READ
tsv = assignment.TimeScaleData(
    StartDate=dt_start, EndDate=dt_end,
    Type=pjAssignmentTimescaledActualWork,  # 24 (verify in probe)
    TimescaleUnit=pjTimescaleDays           # 8 for day, 6 for week
)
# tsv is a TimeScaleValues collection; iterate tsv.Count, tsv(i).Value (minutes)

# WRITE
tsv(i).Value = minutes_int
```

### Project-level
```python
proj.StatusDate = dt              # RW
app.UpdateProject(ProgressDate=dt, UpdatePercentCompleteOnly=False, AllTasks=True)
proj.PercentComplete              # R, derived
```

### Critical enums (probe in T60-T61)
- `pjTimescaleDays = 8`
- `pjTimescaleWeeks = 6`
- `pjAssignmentTimescaledActualWork = 24` (or close — confirm via `msproject_typelib.txt` → `PjAssignmentTimescaledData`)
- `pjAssignmentTimescaledWork = 23`

## 5. Implementation Architecture

```
msproject_mcp_core.py
├── (existing Phase 1+2a+2b)
├── (existing Phase 3a — BASELINE section, lines ~1080-1612)
└── (NEW Phase 3b — PROGRESS section, INSERT @ line 1614)
    ├── Constants
    │   ├── _PROGRESS_PCT_FIELDS = {"percent_complete", "percent_work_complete", "physical_pct"}
    │   ├── _PROGRESS_WORK_FIELDS = {"actual_work_h", "remaining_work_h"}
    │   ├── _PROGRESS_DURATION_FIELDS = {"actual_duration_h", "remaining_duration_h"}
    │   ├── _PROGRESS_DATE_FIELDS = {"actual_start", "actual_finish", "stop", "resume"}
    │   ├── _TIMESCALE_UNIT_MAP = {"day": 8, "week": 6}
    │   └── _PJ_TIMESCALED_ACTUAL_WORK = 24  (probe-confirmed)
    ├── Helpers
    │   ├── _normalize_progress_pct(v) -> float (0-100, accepts int/float/str)
    │   ├── _hours_to_minutes(h) / _minutes_to_hours(m)
    │   ├── _validate_actual_dates(start, finish) -> Optional[error_msg]
    │   ├── _msp_task_set_progress_field(task, field, value) -> bool
    │   ├── _get_assignment_by_resource_id(task, resource_id) -> Optional[Assignment]
    │   └── _read_task_progress_dict(task) -> Dict
    ├── Action functions (12)
    │   ├── _msp_progress_set_task / _get_task
    │   ├── _msp_progress_set_assignment / _get_assignments
    │   ├── _msp_progress_set_by_date / _set_status_date
    │   ├── _msp_progress_clear / _clear_all
    │   ├── _msp_progress_time_phased_write / _read
    │   ├── _msp_progress_bulk_update / _bulk_update_loop
    │   └── _msp_progress_summary
    └── @mcp.tool msproject_progress dispatcher (12 actions)
```

**Insertion point (verified):**
- Phase 3a's last function: `_msp_baseline_set_active` ends @ line ~1612
- Phase 1's first task function after Phase 3a: `_msp_task_update` @ line 1615
- Phase 3b inserts between these — code ordering: progress helpers → set/get → bulk → time-phased → summary

**Reuse from previous phases (BLACK BOXES — do not refactor):**
- `_validate_active_project` (Phase 1)
- `_format_com_error` (T29)
- `_parse_rate` (T32 — for cost numerics)
- `_find_task_by_id` (Phase 1)
- `_find_resource_by_id` (Phase 2b)
- `_msp_dt_or_none` (Phase 3a T44 fix — sentinel-aware datetime str)
- `_route_operation` + `_enter_batch_mode` / `_exit_batch_mode` (Phase 1)
- `_build_task_id_map` (T37 — bulk perf)
- `clean_test_project` fixture (Phase 1 SAFETY)

## 6. KEY DESIGN DECISIONS

### Q1 — Dual-track relationship (task ↔ assignment)
- **MSP behavior:** `task.ActualWork` is **derived** from sum(`assignment.ActualWork`). When user writes `task.ActualWork = X` directly, MSP back-distributes proportionally to assignments (per Units). When user writes `assignment.ActualWork = Y`, MSP rolls up to `task.ActualWork`.
- **Implication:** If user writes BOTH `set_task_progress(actual_work_h=40)` AND `set_assignment_progress(resource_id=R, actual_work_h=10)` on same task, the second write wins (MSP rebalances). Phase 3b documents this in the action docstrings + acceptance script demonstrates per-resource preferred path.
- **Test coverage:** T53 + T55 each verifies its own write path; T55 has an additional test "task.ActualWork rolls up after assignment write".

### Q2 — Time-phased granularity
- **Default:** `day`. Reason: hakediş periodu typically weekly but daily granularity supports both filtering up + Saturday-only entries.
- **`unit` parameter:** `"day"` (= `pjTimescaleDays = 8`) or `"week"` (= `pjTimescaleWeeks = 6`). Probe T60.
- **Read return shape:** `[{period_start: ISO, period_end: ISO, actual_work_hours: float}, ...]` — period_end is exclusive end (start of next period).
- **Write contract:** caller specifies start/end per period; we map them to `TimeScaleValues` collection indices via date range covered by `assignment.TimeScaleData(start, end, ...)` call. Mismatch → graceful per-period failure with reason.
- **Edge case:** if assignment has no work in a given period, MSP returns `Value = 0` (not None). Read normalizes to `0.0` hours.

### Q3 — Hybrid bulk path
- **Same routing as Phase 2b T37:** `_route_operation(N)` → `com_direct` (1-5), `com_batch` (6-19), `mspdi_bulk` (20+, falls back to com_batch in Phase 3b — true MSPDI progress merge is Phase 4+).
- **Pre-built map:** `task_map = _build_task_id_map(proj)` once before the loop.
- **Per-item failure:** collect into `failures: []`, never raise — match Phase 2b semantics.
- **T37 lessons applied:** map built once per call (NEVER inside loop), batch mode entered ONLY for com_batch / mspdi_bulk paths, no `_msp_task_update` call (it re-validates active project per call — way too slow at 200 items).

### Q4 — `PhysicalPercentComplete` semantics
- **NOT auto-derived** from `PercentComplete` or `PercentWorkComplete` — user explicitly sets it.
- **DCMA RULE:** EV = sum(BAC × physical_pct). If user doesn't supply, EV defaults to PercentComplete-based (Phase 5 EVM tool decision).
- **Probe T52:** confirm `task.PhysicalPercentComplete = 50` round-trips on MSP 16.0. If write fails (older MSP), `set_task_progress` returns explicit "field not supported on this MSP version" instead of silent no-op.
- **Get always reads:** `get_task_progress` returns `physical_pct` even if 0 (so caller can detect "not set" vs "set to 0").

## 7. Variance & Validation Rules

**Pct field validation** (`_normalize_progress_pct`):
- Accepts int, float, str ("50", "50.5", "50%")
- Rejects: <0, >100, non-numeric → `ValueError`
- Returns `float` rounded to 2 decimals

**Date validation** (`_validate_actual_dates`):
- If both supplied: `actual_start <= actual_finish` else error
- ISO 8601 (or pywintypes datetime); use `dateutil.parser.parse` for tolerant input

**Work/Duration unit conversion:**
- Public API uses **hours** (e.g., `actual_work_h`)
- MSP COM uses **minutes**; `_hours_to_minutes(h) = round(h * 60)`, `_minutes_to_hours(m) = m / 60`

**Stop/Resume relationship** (advanced — partial progress):
- `task.Stop = D` means task is paused as-of D; future actual work picks up at `task.Resume`
- We expose `stop` and `resume` as optional setters in `set_task_progress`
- If user provides both, `Stop < Resume` else error

## 8. Test Strategy

~50 yeni test:

| File | Tests | Notes |
|---|---|---|
| `test_msproject_progress_helpers.py` | 8-10 | constants, conversions, validators |
| `test_msproject_progress_set_task.py` | 5 | each field setter + dual-mode % vs actuals |
| `test_msproject_progress_get_task.py` | 3 | full read shape + zero-state + after writes |
| `test_msproject_progress_set_assignment.py` | 5 | per-resource write + roll-up verification |
| `test_msproject_progress_get_assignment.py` | 3 | empty list + 1 + 3 assignments |
| `test_msproject_progress_set_by_date.py` | 4 | retroactive update + date-too-early + date-too-late + scope=selected |
| `test_msproject_progress_status_date.py` | 2 | set + readback + clear |
| `test_msproject_progress_clear.py` | 4 | clear single + clear_all + idempotent |
| `test_msproject_progress_time_phased.py` | 6-8 | read empty/populated, write 3 days, write 2 weeks, edge: no overlap |
| `test_msproject_progress_bulk.py` | 5 | 3 com_direct, 10 com_batch, 25 mspdi_bulk + 1 invalid task_id, perf |
| `test_msproject_progress_summary.py` | 4 | empty / partial / fully complete / mixed |
| `test_msproject_progress_dispatcher.py` | 5-6 | 12 action routing + invalid action |

**Total:** ~52 yeni test. Final regression: **~260 PASS + 1 xfail** (207 baseline + ~53 new).

**Performance hedefleri:**
- `bulk_progress_update` 50 task projede **<3s** (com_batch path)
- `time_phased_actual_read` 1 task / 30-day window **<200ms**
- `summary` 200 task projede **<2s** (read-only iteration)

## 9. Acceptance Script

`samples/build_progress_lifecycle.py` — full hakediş workflow:

1. `FileNew` → 50 villa task + 3 resources + assignments (mini Phase 2b chain)
2. **save baseline 0** ("Original" — Phase 3a integration)
3. **set_task_progress** on first 10 tasks with `percent_complete=50`
4. **set_assignment_progress** on next 10 tasks with per-resource man-hours (COW=24h, STL=18h, MSN=10h)
5. **time_phased_actual_write** for 1 task across 5 weekdays (varying daily hours)
6. **time_phased_actual_read** verification — readback matches written periods
7. **set_status_date** to `data_date = today`
8. **set_progress_by_date** for remaining tasks ("plan = actual" assumption to data_date)
9. **bulk_progress_update** with 25 items (mspdi_bulk path)
10. **summary** → BAC, ACWP, project_percent_complete printed
11. **clear_all_progress** → verify reset
12. Cleanup: close test project without saving

Hedef: end-to-end **<15s**. Slowest action expected: bulk_progress_update at ~25 items + time_phased_actual_write/read.

## 10. CLAUDE.md Alignment

| Rule | Phase 3b implication |
|---|---|
| **RULE 0 — DATA INTEGRITY** | Tüm değerler aktif MSP projesinden okunur, asla halüsinasyon ya da varsayım. `summary` action MSP'nin gerçek state'ini yansıtır; cache YOK. |
| **RULE 4 — EVM** | `summary.bac_h = sum(task.Work_h)` (project-level), `summary.acwp_h = sum(task.ActualCost)` (Phase 5 EVM tool BAC/AC kullanır). Phase 3b foundation, EVM matematiği Phase 5'te. |
| **RULE 5 — Time-Phased PV** | `time_phased_actual_read` → period-by-period AC; PV hesabı Phase 5'te ama veri yapısı uyumlu (ISO dates, hours). |
| **RULE 6 — Period Delta** | Period AC = read_actuals(this_week) − read_actuals(prev_week) — Phase 5 hesabı için zemin. |
| **RULE 7 — EV vs SPI(h) data quality** | `set_task_progress` ile `physical_pct` ayrı tutulur; user EV elle yazabilir. SPI(h) çelişki tespiti Phase 5 EVM tool'unda. |
| **RULE 11 — DCMA 14-Point** | Madde 12 (Missed Tasks) tespiti `summary.task_count` + actual_finish bilgisi gerektirir → Phase 3b sağlar. Madde 14 (BEI) için completed_count Phase 3b'den okunur. |
| **RULE 12 — RAG** | Phase 3a `summary` (baseline-relative RAG) + Phase 3b `summary` (progress-relative project_pct_complete) → Phase 5 EVM RAG'i ikisinden gelir. |
| **RULE 13 — STANDART** | "DCMA 14-Point" doğru terim, "PMI PMBOK 8th § 7.4 EVM" referansı uygun. McKinsey/PwC/Deloitte gibi firma ASLA. |

## 11. Open Questions / Risks

1. **`task.PhysicalPercentComplete` setattr** — yazılabilir mi (MSP 16.0)? **T52 PROBE ZORUNLU.** Yazılamazsa Phase 3b sadece getter implementasyonu yapar, setter "not supported" döner.
2. **`UpdateProject(ProgressDate=...)`** — already-progressed tasks için davranış? **T57 PROBE.** Hipotez: MSP "Update tasks: 0% to 100%" mantığı uygular, mevcut % korunabilir / ezilebilir. Test edilecek.
3. **`assignment.TimeScaleData`** — timezone-aware datetime kabul ediyor mu? **T60 PROBE.** Naive datetime kullanılması daha güvenli; UTC offset taşımıyoruz.
4. **`pjAssignmentTimescaledActualWork` enum kodu** — typelib'de `=24`? **T60 PROBE.** Yanlışsa T60 implementation enum'ı düzeltir.
5. **`time_phased_actual_write` yeni period yaratır mı?** Hipotez: TimeScaleData koleksiyonu schedule'a dayalı hücreleri döndürür — sadece zaten ay/saat planlanmış hücrelere yazılabilir. Tasks dışında TimeScaleData yok. Bu durumda T60 implementation, period start/end → TSV index mapping yapar; "no slot" → per-period failure.
6. **`set_progress_by_date` "selected" scope** — programatik selection olmadan nasıl test edilir? T57 muhtemelen sadece `scope="all"` test eder, `scope="selected"` MSP UI gerektirir → manual smoke.

## 12. Out of Scope (Phase 3b'de YOK — Phase 3c veya sonra)

- Tam EVM matematiği (PV, EV, SPI, CPI, EAC, ETC) → **Phase 5 `msproject_evm`**
- Earned Schedule (Lipke 2003 — ES, SPI(t)) → **Phase 5**
- DCMA 14-Point full assessment → **Phase 5**
- Productivity learning curve (Wright 1936) → **Phase 5**
- Time-phased PV/EV write (sadece AC bu Phase'de) → Phase 5
- Cost actuals (`task.ActualCost`, `task.ACWP`) — sadece work-hours bu Phase'de; cost Phase 5 EVM'in işi
- True MSPDI progress merge (offline file write) → **Phase 4 `msproject_file`**

## 13. Acceptance Kriterleri (Phase 3b Tamam)

1. ✅ `msproject_progress` tool 12 action ile çalışıyor (T52-T63 + T64 dispatcher)
2. ✅ Acceptance script `samples/build_progress_lifecycle.py` end-to-end <15s
3. ✅ Phase 3b yeni testleri (~50) PASS
4. ✅ Phase 1+2a+2b+3a mevcut 207+1xfail regression PASS — total **~260 PASS + 1 xfail**
5. ✅ Phase 1 SAFETY: kullanıcının aktif projesi DOKUNULMAZ
6. ✅ Tool count 6 → 7, action count ~40 → ~52
7. ✅ Commit + push GitHub'a
8. ⏸ Kullanıcı manuel onayı → Phase 4 (File MCP) başlar

## 14. Plan Paketi

| Task | Action | Notes |
|---|---|---|
| **T52** | Foundations: helpers + constants + `PhysicalPercentComplete` probe | 8-10 unit tests |
| **T53** | `set_task_progress` action | Dual-mode (% complete / actuals); 5 tests |
| **T54** | `get_task_progress` action | All 9 fields read; 3 tests |
| **T55** | `set_assignment_progress` action | Per-resource man-hour write; 5 tests |
| **T56** | `get_assignment_progress` action | List per-task; 3 tests |
| **T57** | `set_progress_by_date` action | `app.UpdateProject(ProgressDate=)` BIG ONE; 4 tests |
| **T58** | `set_status_date` action | `proj.StatusDate` setter; 2 tests |
| **T59** | `clear_progress` + `clear_all_progress` | Paired (Phase 3a T41 mirror); 4 tests |
| **T60** | `time_phased_actual_write` action | TimeScaleData per-period write + probe; 4 tests |
| **T61** | `time_phased_actual_read` action | TimeScaleData read + week bucket; 4 tests |
| **T62** | `bulk_progress_update` action | Hybrid 1-5 / 6-19 / 20+ (Phase 2b T37 mirror); 5 tests |
| **T63** | `summary` action | EVM-ready aggregate; 4 tests |
| **T64** | FastMCP dispatcher + acceptance + README + push (FINAL) | 5-6 dispatcher tests |

13 task TDD plan, ~8-10 saat tahmin.

---

*Approved by user: 29 Nisan 2026*
*Next: writing-plans skill → Phase 3b Progress implementation plan*
