# MS Project MCP — Phase 5b DCMA Design

**Versiyon:** 1.0
**Tarih:** 1 Mayıs 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → writing-plans next
**Phase 5a HEAD:** `83277e6` (origin/main, in sync, 138/138 cumulative + Phase 5a TAIL fixed)

---

## 1. Hedef

`msproject_mcp` server'ına `msproject_health` tool'u (10. tool) ekle. Hibrit (file+COM) DCMA 14-Point Schedule Health Assessment per CLAUDE.md RULE 10. Phase 5a EVM helper'larını reuse + DCMA-specific link/float/constraint extension. RULE 12 RAG entegrasyonu (>12 pass GREEN / 8-11 AMBER / <8 RED).

Phase 5b sonunda kullanıcı tek mesajla:

> "CAU project'in DCMA durumunu çıkar. RULE 7 (high float) ve RULE 9 (high duration) başarısız mı? Hangi 17 task floatlı? Geçen haftaya göre health iyileşti mi?"

talimatını verebilmeli. Hibrit: file_path verilirse dosyadan, yoksa MSP COM aktif projeden.

## 2. Karar Geçmişi (4 Brainstorming Q&A + Mimari)

### Q1 — Action surface granularity: A
All-in-one + drill_down (3-4 action). DCMA 14 rule doğal grup; tek `assess_all` call'la full report.

### Q2 — Threshold customization: A
Hardcoded DCMA standard. CLAUDE.md RULE 13 rakip firma yasağı + RULE 10 numerik değerleri sabit. Override → Phase 6 polish.

### Q3 — Phase 5a EVM helper reuse: C
Hybrid — base data (`_evm_load_*`) Phase 5a'dan, DCMA-specific (links, floats, constraints, predecessor counts) Phase 5b ek helpers. DRY + hibrit file+COM bedava.

### Q4 — Acceptance scope: C
Hero CAU-style — 200 task + intentional issues for failure detection + 1 baseline + partial progress + assess_all + drill_down + summary + compare. ~50-70s wall clock.

### Mimari — Yaklaşım A (Phase 5a Yaklaşım C aynısı)
- Yeni dosya `dcma_checks.py` — saf 14 rule check fonksiyonu, MSP/COM/file independent
- `msproject_mcp_core.py` Phase 5b section — I/O adapters + 4 action helpers + dispatcher
- Phase 1+2+3+4+5a helpers DOKUNULMAZ

## 3. Tool Surface — `msproject_health` 4 Action

| # | Action | Parametreler | Çıktı |
|---|---|---|---|
| 1 | `assess_all` | `file_path`, `baseline_number=0` | `{rules: [{id, name, threshold, actual, status, failed_count, total_count}], summary: {pass_count, fail_count, overall_rag, executive_text}}` |
| 2 | `summary` | `file_path`, `baseline_number=0` | `{overall_rag, pass_count, fail_count, executive_text}` |
| 3 | `drill_down` | `file_path`, `rule_id` (1-14), `baseline_number=0` | `{rule, threshold, actual_pct, failed_count, total_count, failed_tasks: [{id, name, ...metric...}]}` |
| 4 | `compare` | `file_path`, `snapshot_path` (opsiyonel — Phase 5a snapshot file reuse) | `{current, prev, delta: {rules_improved: [], rules_degraded: []}}` |

**Tüm action'lar `file_path` opsiyonel** — Phase 5a hibrit pattern aynısı.

**Tool count: 9 → 10 (yeni `msproject_health`).**

## 4. Architecture (Yaklaşım A)

```
dcma_checks.py (NEW — pure Python, MSP-independent)
├── DCMA_RULES — list of 14 rule metadata dicts
├── _DCMA_THRESHOLDS — hardcoded values per RULE 10
├── 14 check_* functions:
│   check_no_predecessor / check_no_successor / check_leads / check_lags
│   check_fs_link_pct / check_hard_constraints / check_high_float
│   check_negative_float / check_high_duration / check_invalid_dates
│   check_resources_missing / check_missed_tasks / check_critical_path
│   check_bei
├── assess_all(tasks, links, assignments, baseline, status_date) — aggregate
└── compute_overall_rag(rules) — RAG per pass count

msproject_mcp_core.py PHASE 5B SECTION (after Phase 5a dispatcher)
├── from dcma_checks import check_*, assess_all, DCMA_RULES, compute_overall_rag
├── _dcma_load_links(file_path) — Phase 4 _msp_file_read_links reuse + COM iter
├── _dcma_extract_floats(tasks_iter) — task.TotalSlack / FreeSlack
├── _dcma_extract_constraints(tasks_iter) — task.ConstraintType
├── _dcma_validate_dates(task) — start>finish detection
├── _dcma_collect_full_data(file_path, baseline_number) — wraps Phase 5a _evm_load_*
├── 4 _msp_dcma_* action helpers
└── @mcp.tool msproject_health dispatcher (4 action routing)
```

**Reuse:**
- Phase 5a: `_evm_load_task_data/baseline_data/progress_data` (read-only)
- Phase 4: `_msp_file_read_links/tasks/assignments` (read-only)
- Phase 1: `_validate_active_project`, `_format_com_error`
- Phase 3a: `BASELINE_NUMBERS`

## 5. Hybrid Data Source

```
_dcma_collect_full_data(file_path, baseline_number=0):
    base = _evm_load_task_data(file_path)        # Phase 5a reuse
    baseline = _evm_load_baseline_data(file_path, baseline_number)
    links = _dcma_load_links(file_path)          # NEW Phase 5b
    floats = _dcma_extract_floats(...)           # NEW (COM only — XML has total_float)
    constraints = _dcma_extract_constraints(...) # NEW (COM only — XML has constraint_type)
    return {tasks, links, assignments, resources,
            baseline, status_date, floats, constraints}
```

**File path:** Phase 4 already exposes `total_float`, `critical`, `milestone`, `constraint_type`, `predecessors` per task (T66 probe). DCMA reads these directly.
**COM path:** New iter helpers — `task.TotalSlack`, `task.ConstraintType`, `task.Critical`, predecessor walk.

**Auto-sync:** DCMA tool **read-only**, no auto-sync needed.

## 6. DCMA 14-Point Rules (CLAUDE.md RULE 10)

| # | Rule | Threshold | Math (rough pseudocode) |
|---|---|---|---|
| 1 | No Predecessor | <5% | `count(real_tasks where predecessors==[] AND id!=root) / total_real * 100` |
| 2 | No Successor | <5% | `count(real_tasks where successors==[] AND not last_milestone) / total_real * 100` |
| 3 | Leads | =0 | `count(links where lag<0)` |
| 4 | Lags | <5% | `count(links where lag>0) / total_links * 100` |
| 5 | FS Link % | >90% | `count(links where type=='FS') / total_links * 100` |
| 6 | Hard Constraints | <5% | `count(real_tasks where constraint_type IN [MSO,MFO,SNLT,FNLT]) / total_real * 100` |
| 7 | High Float | <5% | `count(real_tasks where total_slack > 44 working days) / total_real * 100` |
| 8 | Negative Float | =0 | `count(real_tasks where total_slack < 0)` |
| 9 | High Duration | <5% | `count(real_tasks where duration > 44 working days AND not summary) / total_real * 100` |
| 10 | Invalid Dates | =0 | `count(tasks where start>finish OR actual_start>now OR ...)` |
| 11 | Resources Missing | <20% | `count(real_tasks where assignments==[] AND duration>0) / total_real * 100` |
| 12 | Missed Tasks | <5% | `count(real_tasks where baseline_finish < status_date AND not_completed) / total_real * 100` |
| 13 | Critical Path | >0 | `count(real_tasks where critical==True) > 0` (binary) |
| 14 | BEI | >95% | `count(actually_completed_through_status_date) / count(should_have_been_completed_per_baseline) * 100` |

**Overall RAG (industry convention):**
- `pass_count >= 12` → **GREEN**
- `8 <= pass_count <= 11` → **AMBER**
- `pass_count < 8` → **RED**

**Status per rule:**
- `pass` if actual meets threshold
- `fail` if violates
- `warning` (reserved for Phase 6 — currently unused)

## 7. Output Schema

### `assess_all` response
```json
{
  "status": "ok",
  "baseline_number": 0,
  "rules": [
    {
      "id": 1,
      "name": "No Predecessor",
      "threshold": "<5%",
      "actual": 8.5,
      "actual_unit": "%",
      "status": "fail",
      "failed_count": 17,
      "total_count": 200
    }
    // 14 rules total
  ],
  "summary": {
    "pass_count": 11,
    "fail_count": 3,
    "overall_rag": "amber",
    "executive_text": "11/14 DCMA rules pass. 3 issues: ..."
  }
}
```

### `drill_down(rule_id=7)` response
```json
{
  "status": "ok",
  "rule": {"id": 7, "name": "High Float (>44d)", "threshold": "<5%"},
  "actual_pct": 8.5,
  "failed_count": 17,
  "total_count": 200,
  "failed_tasks": [
    {"id": 23, "name": "Foundation Pour", "total_slack_days": 67.5},
    {"id": 45, "name": "MEP Rough-in", "total_slack_days": 52.0}
    // ...17 items
  ]
}
```

### `compare` response
```json
{
  "status": "ok",
  "current": {"pass_count": 11, "rag": "amber", "rules": [...]},
  "prev": {"pass_count": 8, "rag": "red", "rules": [...]},
  "delta": {
    "rules_improved": [{"id": 7, "from_actual": 12.5, "to_actual": 8.5}],
    "rules_degraded": []
  }
}
```

## 8. Test Stratejisi

**~40 yeni test:**
- `test_dcma_checks.py` — 14 rule × 2-3 edge case = ~30 test (pure math, fixture-free, no MSP)
- `test_msproject_dcma_loader.py` — links + floats + constraints extraction (file + COM mock, ~5 test)
- `test_msproject_dcma_assess.py` — assess_all integration on sample fixture (~3 test)
- `test_msproject_dcma_summary.py` — RAG threshold (>12/8-11/<8) (~2 test)
- `test_msproject_dcma_drill_down.py` — per-rule failed_tasks list (~2 test)
- `test_msproject_dcma_compare.py` — snapshot delta (~1 test)
- `test_msproject_dcma_dispatcher.py` — 4 action routing (~3 test)

**Total target:** ~178 PASS (138 Phase 4+5a + ~40 Phase 5b).

**Performance:**
- `assess_all` 1000 tasks → **<2s**
- `drill_down` per-rule fetch → **<200ms**

## 9. Acceptance Script — `samples/build_dcma_lifecycle.py`

```
1. FileNew + 200 task CAU-style + 14 CAU resources + assignments
2. Intentional issues for failure detection:
   - First 12 tasks WITHOUT predecessor (RULE 1 fail: 6% > 5%)
   - 15 tasks duration 60d (RULE 9 fail: 7.5% > 5%)
   - 8 tasks with SNLT constraint (RULE 6 borderline)
   - 1 task with manipulated start > finish (RULE 10 fail)
   - Some tasks without resource assignment (RULE 11 borderline)
3. Save Baseline 0
4. Phase 3b: progress for ~30 tasks (RULE 14 BEI computation)
5. set_status_date "week 2"
6. msproject_health assess_all → display 14 rule results
7. drill_down rule_id=1 → 12 failed task IDs
8. drill_down rule_id=9 → 15 failed task IDs
9. summary → RAG (expected AMBER, ~10/14 pass)
10. (Optional) compare with Phase 5a snapshot — DCMA delta
```

**Hedef:** <60s wall clock. Phase 1 SAFETY pattern (FileNew + FileClose 0).

## 10. Out of Scope (Phase 5b'de YOK — Phase 6+)

- Threshold customization (override per-call) → Phase 6 polish
- Per-rule weight scoring (not all rules equal severity) → Phase 6
- Time-phased DCMA tracking (rule by week) → Phase 6
- DCMA report PDF/Excel export → Phase 5c (Excel) integration
- Calendar-aware working days for high_float (RULE 1: 6d×9h=54h/wk) → MSP'nin native calendar bunu zaten yansıtıyor; eşik standart 44 working days bağlı kalır
- Custom DCMA dialects (NDIA, EVMS Gold Card) → Phase 6+

## 11. Acceptance Kriterleri (Phase 5b Tamam)

1. ✅ `msproject_health` tool 4 action ile çalışır (T85-T93)
2. ✅ Acceptance script `samples/build_dcma_lifecycle.py` <60s
3. ✅ `dcma_checks` saf math test'leri ~30 PASS (no fixtures, no COM)
4. ✅ Phase 4 file path + Phase 1 COM path her ikisi end-to-end test edilir
5. ✅ Phase 1+2+3+4+5a mevcut regression PASS — DOKUNULMAZ
6. ✅ Total ~178 PASS + 0 xfail
7. ✅ All 14 DCMA rules per CLAUDE.md RULE 10 doğru implement edilir
8. ✅ RAG status (>12/8-11/<8 pass) executive output
9. ✅ Commit + push GitHub'a
10. ⏸ Kullanıcı manuel onayı → Phase 5c (Excel) başlar

## 12. Plan Paketi (T85-T93, ~10 task TDD chain)

| Task | İçerik | ~Süre |
|---|---|---|
| **T85** | `dcma_checks.py` foundations + check_no_predecessor + check_no_successor (RULE 1-2) | 1.5h |
| **T86** | `dcma_checks.py` link rules (RULE 3-5: leads, lags, fs_link_pct) | 1.5h |
| **T87** | `dcma_checks.py` task quality (RULE 6, 10, 11: hard_constraints, invalid_dates, resources_missing) | 1.5h |
| **T88** | `dcma_checks.py` duration/float (RULE 7-9: high_float, negative_float, high_duration) | 2h |
| **T89** | `dcma_checks.py` schedule health (RULE 12-14: missed_tasks, critical_path, bei) | 2h |
| **T90** | `dcma_checks.py` `assess_all` aggregator + `compute_overall_rag` | 1h |
| **T91** | `_dcma_load_links` + `_dcma_extract_floats/constraints` (Phase 4 reuse + COM extension) | 2h |
| **T92** | `_dcma_collect_full_data` + 4 action helpers (assess_all/summary/drill_down/compare) | 2.5h |
| **T93** | FastMCP dispatcher + `samples/build_dcma_lifecycle.py` + README + push (FINAL) | 3h |

**Toplam:** ~17 saat impl, ~10-12 commit (T85-T93 + olası fix commit'ler).

**Pattern:**
- T85-T90 saf math → manuel write + self-verify (no probe gerek, test-driven)
- T91/T92 BIG ONE'lar → subagent dispatch (link extraction + integration heavy)
- T93 standard finalize

**Phase 1+2+3+4+5a helpers DOKUNULMAZ.** Sadece `dcma_checks.py` (yeni) + Phase 5b section.

---

*Approved by user: 1 Mayıs 2026*
*Next: writing-plans skill → Phase 5b DCMA implementation plan*
