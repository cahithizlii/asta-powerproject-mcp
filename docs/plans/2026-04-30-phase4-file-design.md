# MS Project MCP — Phase 4 File MCP Design

**Versiyon:** 1.0
**Tarih:** 30 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → writing-plans next
**Phase 3b/2b TAIL HEAD:** `2af190c` (origin/main, in sync, 282 PASS + 1 xfail)

---

## 1. Hedef

`msproject_mcp` server'ına `msproject_file` tool'u (8. tool) ekle. File-based MS Project file manipulation — `.xml`/`.mspdi` (native Python parser, zero Java dependency) ve `.mpp` (MPXJ + Java fallback) için read + write capability. Phase 2b hero gate (`test_bulk_assign_hero_2800_under_5s`) Phase 4'te FLIP eder.

Phase 4 sonunda kullanıcı tek mesajla:

> "Bu kapalı `.mpp` dosyasındaki tüm task'leri oku ve finish > Haziran 2026 olanları listele. Sonra açık MSP projesine 200 villa task + 14 resource + 2800 atama ekle, <5s'de bitmeli, MSP otomatik güncellensin."

talimatını verebilmeli.

## 2. Karar Geçmişi (5 Brainstorming Q&A)

### Q1 — Format desteği: B
Her ikisi: `.xml`/`.mspdi` (native parser) **+** `.mpp` (MPXJ/Java fallback). Asta `asta_mcp_file.py` factory pattern'ı reuse. Plus: **XML schema detection** — bir XML geldiğinde içeriğine bakarak Asta vs MS Project schema ayrımı yapılır (`<Project xmlns="http://schemas.microsoft.com/project">` MS Project; farklı schema → Asta MCP'ye yönlendir).

### Q2 — Read-only mu Read+Write mı: B
Read + Write tam destek. **CRITICAL RULE:** Write action'larından herhangi biri çalıştığında ve MSP açıksa, **otomatik** olarak XML write + MSP COM import + `proj.Reschedule()` zinciri tetiklenir. Opsiyonel parametre yok, default davranış. Memory: `feedback_file_mcp_auto_sync.md`.

`.mpp` write yok (Microsoft proprietary format, hiçbir Python kütüphanesi yazamıyor; `.mpp` üretmek MSP'nin tekelinde — `app.SaveAs(path, ...)` üzerinden).

### Q3 — Phase 2b hero `bulk_add_assignments` MSPDI merge: A
Phase 4'te dahil. `test_bulk_assign_hero_2800_under_5s` strict=True xfail Phase 4'ün success gate'i — Phase 4 implementation gerçek MSPDI `<Assignment>` bulk merge yazınca otomatik flip eder.

### Q4 — DCMA validate action: B
Phase 5'e bırak. `msproject_health` ayrı tool olarak Phase 5'te (separation of concerns: Phase 4 = file I/O, Phase 5 = analiz/raporlama).

### Q5 — Acceptance script senaryosu: A
Tam hero + lifecycle. `samples/build_file_lifecycle.py`:
- 200 villa task + 14 CAU resource + 2800 atama hero (<5s strict)
- File-based read demo (temp.xml export + parse + verify counts)
- Write demo (task duration update + auto-sync verify)
- `.mpp` read demo (MPXJ path)
- Hedef: <30s wall clock

## 3. Tool Surface — `msproject_file` 14 Action

### Read (8) — dosya kapalıyken bile

| # | Action | Parametreler | Çıktı |
|---|---|---|---|
| 1 | `read_tasks` | `file_path`, [`filters`, `limit`] | `{status, count, tasks: [{id, name, duration_h, start, finish, percent_complete, work_h, cost, summary}]}` |
| 2 | `read_links` | `file_path` | `{status, count, links: [{from_id, to_id, type, lag_days}]}` |
| 3 | `read_resources` | `file_path` | `{status, count, resources: [{id, name, type, max_units, std_rate, ovt_rate}]}` |
| 4 | `read_assignments` | `file_path`, [`task_id`] | `{status, count, assignments: [{task_id, resource_id, units, work_h, cost}]}` |
| 5 | `read_calendars` | `file_path` | `{status, calendars: [{name, week_days, exceptions}]}` |
| 6 | `read_baselines` | `file_path`, [`baseline_number=0`] | `{status, baseline_number, saved_date, tasks: [...]}` (Phase 3a entegrasyon) |
| 7 | `read_progress` | `file_path`, [`include_assignments=False`] | `{status, status_date, tasks: [...]}` (Phase 3b entegrasyon) |
| 8 | `query` | `file_path`, `expression`, [`limit`] | `{status, count, results: [...]}` (ad-hoc filter, örn. `"finish > '2026-06-01' AND duration_h > 40"`) |

### Write (6) — XML write + MSP açıksa otomatik COM import + Reschedule

| # | Action | Parametreler | Çıktı |
|---|---|---|---|
| 9 | `add_tasks` | `file_path`, `items: [{name, duration, ...}]` | `{status, count, task_ids, auto_imported: bool, reschedule_ok: bool}` |
| 10 | `add_links` | `file_path`, `items: [{from_id, to_id, type, lag}]` | `{status, count, link_ids, auto_imported, reschedule_ok}` |
| 11 | `add_resources` | `file_path`, `items: [{name, type, max_units}]` | `{status, count, resource_ids, auto_imported, reschedule_ok}` |
| 12 | `bulk_add_assignments` | `file_path`, `items: [{task_id, resource_id, units}]` | `{status, count, auto_imported, reschedule_ok, elapsed_s}` 🚀 HERO |
| 13 | `update_task` | `file_path`, `task_id`, `fields: {duration, start, ...}` | `{status, task_id, auto_imported, reschedule_ok}` |
| 14 | `save_as` | `file_path`, `output_path` (`.xml`) | `{status, output_path, size_bytes}` |

## 4. Architecture (Yaklaşım B — Factory + Manager Classes)

```
msproject_mcp_core.py
│
├── Phase 1-3 (DOKUNULMAZ)
│   └── _msp_task_*, _msp_link_*, _msp_resource_*, _msp_baseline_*, _msp_progress_*
│
└── Phase 4 — YENİ SECTION (en altta, after Phase 3b helpers)
    │
    ├── from mspdi_parser import MspdiProject  (Asta'dan reuse — XML read+write)
    │
    ├── class MspMppFileManager:           (NEW — MPP read-only via MPXJ)
    │   __init__(file_path), read_tasks(), read_links(), read_resources(),
    │   read_assignments(), read_calendars(), read_baselines(), read_progress()
    │   (Asta'nın AstaFileManager'ından adapt; .pp uzantısı yok, sadece .mpp)
    │
    ├── _detect_msp_xml_schema(file_path) -> bool
    │   Reads first 512 bytes; checks for "schemas.microsoft.com/project" namespace
    │   Returns False (and raises informative error) if Asta schema detected
    │
    ├── _get_msp_file_manager(file_path) -> Union[MspdiProject, MspMppFileManager]
    │   .xml/.mspdi → MspdiProject (after schema check)
    │   .mpp        → MspMppFileManager
    │   else        → ValueError("Unsupported extension")
    │
    ├── _msp_file_*  helper functions (14 adet)
    │   _msp_file_read_tasks(file_path, ...)
    │   _msp_file_add_links(file_path, items)
    │   _msp_file_bulk_add_assignments(file_path, items)
    │   _msp_file_query(file_path, expression, ...)
    │   ... (Phase 1-3 helper convention preserved)
    │
    ├── _auto_sync_to_open_msp(modified_xml_path) -> dict
    │   {auto_imported: bool, reschedule_ok: bool, error: Optional[str]}
    │   Logic:
    │     try GetActiveObject('MSProject.Application')
    │       if fail → return {auto_imported: False}
    │     app.FileOpen(temp.xml) → temp_proj
    │     EditCopy on temp_proj.Tasks (or batch operations)
    │     active_proj activate + EditPaste merge
    │     active_proj.Reschedule()
    │     close temp_proj without save (FileClose 0)
    │     return {auto_imported: True, reschedule_ok: True}
    │
    └── @mcp.tool msproject_file dispatcher  (14 action routing)
        (Phase 1-3 dispatcher pattern, e.g. msproject_baseline)
```

**Reuse from previous phases:**
- `_validate_active_project` (Phase 1) — auto-sync için MSP COM gate
- `_format_com_error` (T29) — error formatting
- `_parse_rate` (T32) — locale-aware float (cost fields)
- `clean_test_project` fixture (Phase 1 SAFETY)
- `mspdi_parser.MspdiProject` (Asta) — XML read+write base
- `asta_mcp_file.AstaFileManager` (Asta) — MPP MPXJ pattern reference (adapt to MspMppFileManager)

**Pattern uyarıları (Phase 3b'den ders):**
- Probe-first: MspdiProject'in `<Assignment>` write capability'sini implementation'dan ÖNCE probe et (mevcut Asta usage'ı assignment write yapıyor mu?)
- MSP MSPDI FileOpen Duration drop bug'ı (Phase 2b TAIL): assignment merge sonrası Duration verify et, gerekirse post-paste re-set
- `app.UpdateProject` positional only (Phase 3b T57): aynı pattern Phase 4 auto-sync için de geçerli olabilir
- TimeScaleData enum constants empirically corrected (Phase 3b T60): MSP COM constants'a kör güvenme

## 5. Auto-Sync Behavior (Q2 cevabı detayı)

**Kural:** Write action → XML'e yazıldıktan sonra `_auto_sync_to_open_msp(temp.xml)` ZORUNLU çağrılır. Opsiyonel parametre YOK.

**Akış:**
```
write action
  │
  ├─ XML write (mspdi_parser)
  ├─ _auto_sync_to_open_msp(temp.xml)
  │   ├─ MSP COM available?
  │   │   NO  → {auto_imported: False, msg: "MSP closed; XML at <path>"}
  │   │   YES ↓
  │   ├─ app.FileOpen(temp.xml) → temp_proj
  │   ├─ EditCopy/EditPaste merge (existing pattern from _msp_task_bulk_add mspdi path)
  │   ├─ Phase 2b TAIL pattern: post-paste field re-establish if Duration/dates dropped
  │   ├─ active_proj.Reschedule()
  │   └─ {auto_imported: True, reschedule_ok: True}
  │
  └─ return {action result + auto_imported + reschedule_ok}
```

**Hata durumunda:** XML diskte kalır (manuel recovery için), `{auto_imported: False, error: <COM error>}` döndürülür. User aware.

**MSP closed senaryosu (CI/headless):** Sadece XML kaydedilir, no merge attempt. Bu meşru bir use case — kullanıcı sonra MSP'de açar.

## 6. Hero Path — `bulk_add_assignments` <5s

**Phase 2b mevcut (yavaş):**
- `_msp_resource_bulk_assign` `mspdi_bulk` path → `com_batch_fallback` (gerçek MSPDI assignment merge yapmıyor, tek tek COM `Assignments.Add`)
- 2800 × 11.30ms = **31s+** (bugün villa run'da 700 × 11.30ms = 7.91s ölçüldü)
- `test_bulk_assign_hero_2800_under_5s` strict xfail

**Phase 4 hero (hızlı):**
1. `mspdi_parser.MspdiProject` ile geçerli proje state'i temp XML'e export et (~1s)
2. 2800 `<Assignment>` elemanını bulk olarak XML'e yaz — pure Python, no COM (~0.5-1s)
3. `_auto_sync_to_open_msp(temp.xml)` → FileOpen + EditPaste merge (~1-2s)
4. `proj.Reschedule()` (~0.5s)
- **Toplam ~3-4s** (target <5s, ~25% headroom)

**Strict xfail flip:** Phase 4 implementation çalışınca `test_bulk_assign_hero_2800_under_5s` automatic FLIP eder (xpass), test suite fail eder eğer flip etmezse → Phase 4 success gate.

## 7. XML Schema Detection (Asta vs MS Project routing)

Bir `.xml` dosyası geldiğinde, Phase 4 file MCP **mı** Asta file MCP'ye **mi** yönlendirilmesi gerektiğine Claude (controller) karar verir. `_detect_msp_xml_schema(file_path)` helper:

```python
def _detect_msp_xml_schema(file_path: str) -> bool:
    """Read first 512 bytes; check for MS Project MSPDI namespace.
    Returns True if MS Project XML, False if Asta or unknown.
    """
    with open(file_path, 'rb') as f:
        head = f.read(512).decode('utf-8', errors='replace')
    return 'schemas.microsoft.com/project' in head
```

`_get_msp_file_manager` bunu çağırır; MS Project değilse `ValueError("Not a MS Project XML — appears to be Asta or unknown schema")` raise eder. Kullanıcı veya Claude doğru MCP'ye yönlendirir.

## 8. Test Stratejisi

**~35 yeni test:**
- `test_msproject_file_factory.py` — _get_msp_file_manager, _detect_msp_xml_schema, ext routing (5 test)
- `test_msproject_file_read_xml.py` — 8 read action × XML path (8-10 test)
- `test_msproject_file_read_mpp.py` — 8 read action × MPP path (5-6 test, smaller fixture)
- `test_msproject_file_write.py` — 6 write action × XML (6-8 test)
- `test_msproject_file_auto_sync.py` — MSP açık vs kapalı, hata recovery (4-5 test)
- `test_msproject_file_hero.py` — bulk_add_assignments 2800 <5s (1 test, strict perf)
- `test_msproject_file_query.py` — filter expression parser (3-4 test)
- `test_msproject_file_dispatcher.py` — 14 action routing + invalid action (3-4 test)

**Total target:** **~317 PASS + 0 xfail** (282 baseline + ~35 new + 1 xfail flipped to passing).

**Performance hedefleri:**
- `read_tasks` 1000 task XML → **<1s**
- `bulk_add_assignments` 2800 assignment → **<5s** (strict)
- `query` 1000 task with filter → **<500ms**
- Acceptance script end-to-end → **<30s**

## 9. Acceptance Script — `samples/build_file_lifecycle.py`

```
1. Empty MSP project (FileNew, isolated)
2. Build base: 200 villa task (mspdi_bulk path) + 14 CAU resources
3. 🚀 HERO: bulk_add_assignments 200×14=2800 → <5s strict + auto-sync verify
4. Read demo: temp.xml export → msproject_file read_tasks/links/resources/assignments → counts verify
5. Write demo: update_task duration → auto-sync verify (proj.Reschedule reflected in COM)
6. .mpp read demo: temp.mpp export (MSP COM) → msproject_file read_tasks (MPXJ path) → count match
7. Query demo: msproject_file query "finish > '...' AND cost > 50000"
8. Cleanup: FileClose 0 (Phase 1 SAFETY)
```

**Hedef:** <30s wall clock, MSP açık 1 user project korumalı.

## 10. Out of Scope (Phase 4'te YOK — Phase 5+)

- DCMA 14-Point validate (Phase 5 `msproject_health`)
- EVM math (Phase 5 `msproject_evm`)
- Excel import/export (Phase 5 `msproject_excel`)
- `.mpp` write (Microsoft proprietary, technically impossible)
- Asta file format (`.pp`) — Asta MCP'nin sorumluluğu
- Time-phased baseline write (Phase 4+ ek capability)
- Custom fields write (Phase 4+ ek capability)

## 11. Acceptance Kriterleri (Phase 4 Tamam)

1. ✅ `msproject_file` tool 14 action ile çalışıyor (T65-T74)
2. ✅ Acceptance script `samples/build_file_lifecycle.py` end-to-end <30s
3. ✅ Phase 4 yeni testleri (~35) PASS
4. ✅ Phase 1+2+3 mevcut 282+1xfail regression PASS — total **~317 PASS + 0 xfail**
5. ✅ `test_bulk_assign_hero_2800_under_5s` xfail FLIP eder (automatic via strict=True)
6. ✅ Phase 1 SAFETY: kullanıcının aktif projesi DOKUNULMAZ
7. ✅ XML schema detection: MS Project ≠ Asta ayrımı net hata mesajıyla
8. ✅ Auto-sync: write action → MSP açık ise otomatik import + Reschedule (memory rule)
9. ✅ Commit + push GitHub'a
10. ⏸ Kullanıcı manuel onayı → Phase 5 (Power Tools — health/evm/excel) başlar

## 12. Plan Paketi (T65-T74, ~10 task TDD chain)

| Task | İçerik | ~Süre |
|---|---|---|
| **T65** | Foundations: factory + MspMppFileManager class iskeleti + JVM lifecycle + _detect_msp_xml_schema | 2h |
| **T66** | `read_tasks` + `read_links` (XML + MPP path) | 1h |
| **T67** | `read_resources` + `read_assignments` + `read_calendars` | 1h |
| **T68** | `read_baselines` + `read_progress` (Phase 3a/3b entegrasyon) | 1h |
| **T69** | `query` action (filter expression parser) | 2h |
| **T70** | `add_tasks` + `add_links` + `add_resources` (XML write) | 2h |
| **T71** | `update_task` + `save_as` (XML write) | 1h |
| **T72** | `_auto_sync_to_open_msp` helper + tests | 2h |
| **T73** | 🚀 `bulk_add_assignments` HERO + xfail flip | 3h |
| **T74** | FastMCP dispatcher + acceptance script + README + push | 2h |

**Toplam:** ~17 saat impl, ~12-15 commit (T65-T74 + olası fix commit'ler).

**Pattern (Phase 3a/3b'de kanıtlanmış):**
- BIG ONEs (T65, T69, T72, T73) → subagent-driven-development (implementer + spec reviewer + quality reviewer)
- Trivial actions (T66-T68, T70-T71) → manuel write + self-verify
- Probe-first (T57/T60 lessons): MspdiProject'in assignment write API'sini implementation'dan ÖNCE probe et
- Phase 1+2+3 kodu DOKUNULMAZ; Phase 4 kendi section'ında

---

*Approved by user: 30 Nisan 2026*
*Next: writing-plans skill → Phase 4 File MCP implementation plan*
