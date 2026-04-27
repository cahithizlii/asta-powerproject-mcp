# MS Project MCP Server — Design Document

**Versiyon:** 1.0
**Tarih:** 24 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → Implementation (Phase 1 starting)

---

## Executive Summary

Asta Powerproject MCP'ye paralel olarak MS Project MCP Server eklenecek. **Hedef:** Asta MCP'den daha güçlü ve daha hızlı — built-in DCMA health check, EVM analytics ve Excel ↔ MSP çift yönlü akış ile.

- **Repo:** Mevcut `asta-powerproject-mcp` (aynı repo, ayrı MCP server'lar)
- **Mimari:** 3 MCP server — `msproject_mcp` (COM, 8 tool) + `msproject_file` (.mpp/.xml, 4 tool) + `msproject_power` (3 tool)
- **Toplam:** 15 tool (Asta MCP'nin 12'sine karşılık + 3 power tool)
- **Hız:** Hibrit — 1-5 item COM doğrudan, 6-19 COM batch (Calculation manual), 20+ MSPDI XML bulk path (~2-5 sn / 200 task)
- **Timeline:** 6 phase × 1-3 gün = ~10 gün toplam
- **Test discipline:** Her phase için build → test → bug fix → onay → next phase

## 1. Karar Geçmişi (Brainstorming Sonuçları)

### Q1 — Kapsam Önceliği
**A+B+C+D first wave** (Task/Schedule + Resource + Cost/Budget + Baseline/Tracking).
E (Custom Fields), F (View/Reporting), G (Master/Sub + Advanced) ikinci dalga.

### Q2 — Repo Yapısı
**Aynı repo, ayrı MCP server'lar** (Yaklaşım A). `mspdi_parser.py` paylaşımlı.

### Q3 — Hız & Bulk Pattern
**Hibrit (Yaklaşım C)**: COM küçük op + MSPDI XML bulk büyük op. Excel import first-class tool olarak.

### Q4 — File MCP Kapsamı
**Tam simetrik (Yaklaşım A+B)**: `.mpp` (MPXJ Java) + `.xml/.mspdi` (mspdi_parser).

### Q5 — Excel Import Şeması
**Otomatik kolon tespiti (Yaklaşım C)**: Microsoft default + MCS extended + Türkçe header'lar.

### Strateji
**Yaklaşım 3: Power-Tier** — 12 tool Asta paralel + 3 power tool (`health`, `evm`, `excel`).

## 2. Architecture & Layers

```
asta-powerproject-mcp/
├── asta_mcp_core.py          # mevcut — Asta COM
├── asta_mcp_file.py          # mevcut — Asta dosya
├── mspdi_parser.py           # ortak — Asta + MSP XML parser
│
├── msproject_mcp_core.py     # YENİ — MS Project COM (8 tool)
├── msproject_mcp_file.py     # YENİ — MSP file (4 tool)
├── msproject_power.py        # YENİ — Power (3 tool)
├── msproject_bulk.py         # YENİ — MSPDI XML bulk-write engine
├── msproject_typelib.txt     # YENİ — COM type library dump
│
├── docs/plans/               # YENİ — design docs
├── tests/                    # YENİ — test suite
└── samples/                  # YENİ — örnek build script'ler
```

3 MCP server config'i Claude Code'a kaydedilecek:
- `msproject_mcp` (COM-based, 8 tool)
- `msproject_file` (file-based, 4 tool)
- `msproject_power` (analytics, 3 tool)

## 3. Tool Surface (15 Tool)

### A. COM MCP — `msproject_mcp` (8 tool)

| Tool | Action'lar |
|---|---|
| `msproject_task` | add, update, delete, add_summary, add_milestone, get, list, bulk_add |
| `msproject_link` | add, delete, update, bulk_add, chain |
| `msproject_resource` | add, update, delete, list, assign, unassign, bulk_assign |
| `msproject_schedule` | reschedule, level, set_data_date, protect_actuals |
| `msproject_progress` | update, bulk_update, set_status_date, update_actual, update_remaining |
| `msproject_baseline` | save, clear, compare, list_baselines |
| `msproject_calendar` | create, update, add_exception, set_working_hours, assign_to_task, assign_to_resource, list, holidays_uzbek |
| `msproject_export` | xml, mpp, pdf, xlsx, csv, report |

### B. File MCP — `msproject_file` (4 tool)

| Tool | Format |
|---|---|
| `msproject_query` | analyze, list_tasks, critical_path, wbs, float, delay, get_task, search, latest_finishing, missing_links, link_chain — `.mpp` (MPXJ) + `.xml/.mspdi` |
| `msproject_file_resource` | list, assignments, loading |
| `msproject_file_calendar` | get |
| `msproject_file_edit` | add_task, update_task, delete_task, add_link, remove_link, update_link, update_progress, assign_resource, save (sadece XML) |

### C. Power MCP — `msproject_power` (3 tool)

| Tool | İşlev |
|---|---|
| `msproject_health` | DCMA 14 + zero-float check + open ends + naming consistency + full_audit |
| `msproject_evm` | summary (BAC/PV/EV/AC + KPI), time_phased (aylık), forecast (EAC + TCPI), s_curve, manhour_distribution |
| `msproject_excel` | import (auto-detect), export, bulk_progress_update, template_generate |

## 4. Hibrit Speed Strategy

**Otomatik routing — `_route_operation(op_count)`:**

| op_count | Path | Hız |
|---|---|---|
| 1-5 | COM doğrudan (real-time) | ~50-200ms/item |
| 6-19 | COM batch (Calculation manual + ScreenUpdating off + EventsEnabled off) | ~10-30ms/item |
| 20+ | MSPDI XML bulk (write XML + FileOpen import) | ~3-5 sn / 200 task |

**Performance hedefi (250 task + 300 link):**
- Asta MCP: ~5 dakika
- MSP MCP: **~3-5 saniye** ⚡

**MS Project specific tricks:**
- `app.Calculation = pjManual` — bulk sırasında auto-recalc kapatma
- `proj.EventsEnabled = False` — VBA event tetiklenmesin
- `app.ScreenUpdating = False` — UI freeze
- `atexit` ile mode restore garantisi (crash sonrası bile)

## 5. Power Tools Detayı

### `msproject_health` — Schedule Quality
- DCMA 14-Point (1-14 kontroller)
- Zero-float saturation check (Asta CAU raporundan öğrenildi)
- Open ends (no predecessor/successor)
- Naming consistency (Forth/Fourth, BT/B karışıklığı)
- Full audit: tek call → markdown rapor + JSON skor

### `msproject_evm` — Earned Value Management
- BAC/PV(t)/EV(t)/AC(t) + CPI/SPI
- Earned Schedule (Lipke 2003): ES, SPI(t), SV(t)
- Forecast (PMI PMBOK 8th § 7.4.2): EAC₁ = AC + (BAC-EV); EAC₂ = BAC/CPI; EAC₃ = AC + (BAC-EV)/(CPI×SPI); ETC; VAC; TCPI
- Time-phased monthly breakdown (CAU raporundakine eş)
- Manhour distribution per resource/ekip

### `msproject_excel` — Excel ↔ MSP
- **Import auto-detect:** kolon başlıklarına bakarak Microsoft default + MCS extended + Türkçe alanları map'ler
- **Export:** template veya populated
- **Bulk progress update:** Excel/CSV'den 200+ task progress'i tek call'da
- **Template generate:** boş Excel template oluşturucu

## 6. Error Handling

10 hata kategorisi:
1. COM bağlantı hatası → file MCP fallback önerisi
2. Active project yok → açık olmasını iste
3. Calculation mode kalmış → atexit restore
4. MSPDI bulk import hatası → COM Path 2 fallback
5. Concurrent op çakışması → module-level lock (timeout 60s)
6. Türkçe karakter encoding → UTF-8 disiplini
7. Reschedule loop → link rollback + kullanıcıya hangi task'lar göster
8. Excel auto-detect failure → partial map + actionable error
9. MPXJ Java not found → "Save As XML" yönlendirmesi
10. File lock → "MS Project'te kapalı tutun veya COM kullanın"

**User-facing hata mesajı ilkesi:** Her zaman eylem önerisi içersin.

## 7. Testing Strategy

```
        ┌─────────────────────┐
        │  E2E (5 senaryo)    │
        └─────────────────────┘
      ┌───────────────────────────┐
      │  Integration (~25 test)   │
      └───────────────────────────┘
    ┌─────────────────────────────────┐
    │  Unit (~60 test)                │
    └─────────────────────────────────┘
```

- **Unit:** mspdi parser symmetry, Excel auto-detect, EVM math, DCMA scoring (CI'de çalışır)
- **Integration:** COM tool'ları gerçek ActiveProject ile (sadece local Windows)
- **E2E:** 5 senaryo (Excel→200 task villa, CAU subset, master+sub, .mpp analiz, error recovery)
- **Performance benchmarks:** timing assertions (`<5sn / 200 task`)
- **Coverage hedefi:** Unit %85+, parser %95+, power tools %90+

## 8. Build Phasing & Timeline

| Faz | Süre | Deliverable | Demo |
|---|---|---|---|
| **Phase 1** | 2 gün | Foundation + Task/Schedule + Bulk engine | "200-task villa 5sn'de yüklenir" |
| **Phase 2** | 1 gün | Resource + Calendar (Uzbek bayramları) | "9 tatil tek call'da, 6 ekip atanmış" |
| **Phase 3** | 1 gün | Progress + Baseline (multi-baseline) | "B1 kaydet, progress gir, B2 ile karşılaştır" |
| **Phase 4** | 1 gün | File MCP (.mpp + XML) | "MSP kapalı, XML'den DCMA çıkarımı" |
| **Phase 5** | 3 gün | Power tools (health, evm, excel) | "Excel import → DCMA + EVM → 30sn'de PDF" |
| **Phase 6** | 2 gün | Polish + diğer PC kurulum + push | Final commit + readme update |

**Toplam: 10 iş günü** — 15 tool, %85+ coverage, repo'da hazır.

## 9. Test/Onay Disiplini (Kullanıcı Talebi)

> "Her phase'de böyle yapacağız, emin olmadan herşeyin full çalıştığından diğer phase'e geçmek yok."

**Phase tamamlama kriterleri (her phase için zorunlu):**
1. Tüm planlanmış tool/action'lar implemented
2. Unit test'ler passing
3. Integration test'ler passing (kullanıcının açık olduğu Project1'e karşı)
4. Manuel demo başarılı
5. Bilinen bug'lar 0
6. Kullanıcı görüş verip onaylar
7. Commit + push
8. **ONAY ALINMADAN** sonraki phase BAŞLAMAZ

## 10. Bu Oturumda Plan

**Phase 1 başlıyor:** Foundation + Task/Schedule Core + Bulk engine.

Sonraki adım: `writing-plans` skill'i invoke edilecek → Phase 1 detaylı implementation plan oluşturulacak (taskname-by-taskname, dosya-by-dosya).

---

*Approved by user: 24 Nisan 2026*
*Next: writing-plans skill → implementation plan for Phase 1 only*
