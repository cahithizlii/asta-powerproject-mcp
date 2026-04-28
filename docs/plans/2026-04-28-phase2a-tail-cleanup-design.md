# Phase 2a TAIL Cleanup Design

**Versiyon:** 1.0
**Tarih:** 28 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → writing-plans next
**Phase 2a HEAD:** `42574f2` (origin/main, in sync)

---

## 1. Hedef

Phase 2a per-task code review'lerinde "non-blocker, defer to TAIL" olarak işaretlenen 9 hijyen item'i + 1 untracked file kararını 4 ardışık commit ile halletmek. Phase 2b başlamadan önce Phase 2a'yı tertemiz GitHub'a bırakmak.

## 2. Karar Geçmişi (Brainstorming Çıktıları)

### Q1 — Scope ve commit yapısı: A
9 item, **3 grup commit** (Contract / DX / Test+doc) + ayrı 4. commit untracked file için. Her grup kendi başına revertable, git history granular kalır.

### Q2 — Untracked file: A
`tools/export_empty_msp_fixture.py` **commit** edilecek. Yanındaki kardeş `tools/dump_msproject_typelib.py` zaten commit'li, simetri için tutarlı + dokümante edilmiş kullanımı var (Phase 1 T3'ten kalma helper).

## 3. Architecture

4 ardışık commit, hepsi `main` üzerinde, Phase 1 SAFETY pattern korunur (testler `clean_test_project` fixture kullanmaya devam). Tüm değişiklikler:
- `msproject_mcp_core.py` (helper'lar + dispatcher)
- `tests/test_msproject_calendar_*.py` (mevcut test dosyalarına ek + 1-2 yeni dosya gerekirse)
- `tools/export_empty_msp_fixture.py` (sadece git add)

Yeni production file YOK.

## 4. Commit 1 — Contract/Behavior (3 item)

### 4.1 `uid` → `calendar_uid` standardize
**Sorun:** `_msp_calendar_list` çıktısı `uid` field, `_msp_calendar_create` `calendar_uid` field. Tutarsızlık.
**Çözüm:** `_msp_calendar_list` çıktısında `uid` → `calendar_uid`. T26 dispatcher passthrough — değişiklik gerekmez.
**Test güncellemesi:** `test_msproject_calendar_list.py`'da varsa `uid` referanslarını güncelle (mevcut testlerde uid'ye assertion yok ama eklenecek count assertion `calendar_uid` ile uyumlu olmalı).

### 4.2 Dead shift loop temizle
**Sorun:** `_msp_calendar_add_exception`'daki `Shift1Start`/`Shift1Finish`/`Shift2Start`/... loop. MSP 16.0'da Calendar Exception COM object'inde bu flat property'ler YOK (sub-object `Shift1.Start` var). Loop her iterasyonda `AttributeError` raise ediyor, inner try/except silently swallow ediyor. Sonuç: dead code.
**Çözüm:** Loop'u tamamen kaldır. Comment ekle: `# Type=PJ_EXCEPTION_DAILY=7 already implies non-working in MSP semantics — no shift mutation needed`.
**Test güncellemesi:** Mevcut `test_add_exception_actually_non_working` (T21 fix) zaten `cal.Period(date).Working is False` ile non-working contract'ını lock'ladı. Loop silindiğinde test hâlâ PASS olmalı (kontrat aynı).

### 4.3 Vestigial `working=working` field düşür
**Sorun:** `_msp_calendar_add_exception` success payload'ında `"working": working` field var. Ama T21 fix'inde `working=True` artık erken error veriyor — fonksiyon success path'e ulaştığında `working` ALWAYS False. Field misleading.
**Çözüm:** Success payload'dan `"working"` field kaldır.
**Test güncellemesi:** Mevcut testlerde `working` field'a assertion var mı kontrol — varsa kaldır.

## 5. Commit 2 — DX (2 item)

### 5.1 `_format_com_error(e)` helper
**Sorun:** `pywintypes.com_error` exception'ları `str(e)` ile dump edilince ugly tuple: `(-2147352567, 'Exception occurred.', (0, 'Microsoft Project', 'Some message', None, 0, -2147352567), None)`. End-user bu tuple'ı okuyamaz.
**Çözüm:** Top-level helper:
```python
def _format_com_error(e: Exception) -> str:
    """Extract human-readable message from pywintypes.com_error or other exceptions."""
    if hasattr(e, "args") and len(e.args) >= 3 and isinstance(e.args[2], tuple):
        # pywintypes.com_error: args = (hresult, msg, excepinfo, argerr)
        # excepinfo = (wCode, source, description, helpFile, helpContext, scode)
        excepinfo = e.args[2]
        if len(excepinfo) >= 3 and excepinfo[2]:
            return str(excepinfo[2]).strip()
        if len(e.args) >= 2 and e.args[1]:
            return str(e.args[1]).strip()
    return str(e)
```
Tüm calendar `_msp_*` fonksiyonlarındaki `return {"status": "error", "error": str(e)}` → `_format_com_error(e)` ile değiştir. Phase 1 fonksiyonları DEĞİŞTİRİLMEZ (out of scope).
**Test:** Yeni `tests/test_msproject_format_com_error.py` — pure unit test, MS Project gerek yok. ~3 test (com_error tuple parse, plain Exception fallback, empty args edge case).

### 5.2 Dispatcher `name`/`calendar_name` alias
**Sorun:** `_msp_calendar_create`/`_update` `name=` bekler ama `_add_exception`/`_assign_*`/`_list`/`_holidays_uzbek` `calendar_name=` bekler. User-facing inconsistency.
**Çözüm:** `msproject_calendar` dispatcher'da pre-process — eğer action `calendar_name` bekliyor ama params'ta `name` varsa, `calendar_name`'e çevir (ve tersi). YAGNI: sadece `name` ↔ `calendar_name` swap, başka alias yok.
```python
# Inside dispatcher, after action extracted, before kwargs splat:
NAME_ALIAS_ACTIONS = {"add_exception", "assign_to_task", "assign_to_resource",
                      "list", "holidays_uzbek"}
if action in NAME_ALIAS_ACTIONS and "name" in p and "calendar_name" not in p:
    p["calendar_name"] = p.pop("name")
elif action in {"create", "update"} and "calendar_name" in p and "name" not in p:
    p["name"] = p.pop("calendar_name")
```
**Test:** `test_msproject_calendar_dispatcher.py`'a 1 test ekle — `msproject_calendar({"action": "add_exception", "name": "MyCal", ...})` → ok (alias çalıştı).

## 6. Commit 3 — Test/Doc Polish (4 item)

### 6.1 T19 monkeypatch test — "succeeded but not found" guard
**Hedef:** `_msp_calendar_create`'in dead branch'i (`BaseCalendarCreate succeeded but '{name}' not found`) için test coverage.
**Yöntem:** Monkeypatch `app.BaseCalendarCreate` to a no-op (does nothing). Then `_msp_calendar_create("Y", "Standard")` should return error with "succeeded but not found".
**Dosya:** `test_msproject_calendar_create.py`'a 1 test ekle.

### 6.2 T24 docstring — list ordering note
**Hedef:** `_msp_calendar_list` ordering davranışını belgele.
**Çözüm:** Docstring'e ekle: `"Order: matches proj.BaseCalendars enumeration (typically insertion order, not sorted). If callers need lexicographic ordering, sort client-side."`

### 6.3 T24 count assertion in test
**Hedef:** `test_list_includes_standard`'a `count` field assertion ekle.
**Çözüm:** Test'e ekle: `assert r["count"] == len(r["calendars"]) >= 1`.

### 6.4 T25 debug log — pre-scan exception swallow
**Hedef:** `_msp_calendar_holidays_uzbek` pre-scan loop'unda `cal.Exceptions` okunamadığında debug log düş.
**Çözüm:** `except Exception as e: pass` → `except Exception as e: logger.debug(f"holidays_uzbek pre-scan failed (treating as empty): {e}")`.

## 7. Commit 4 — Add untracked tool file

```bash
git add tools/export_empty_msp_fixture.py
git commit -m "Add tools/export_empty_msp_fixture.py (Phase 1 T3 fixture regenerator)"
```

Message açıklayıcı: bu dosya Phase 1 T3'te yazılmış, `tests/fixtures/empty_msp.xml` regenerate eder. Yanındaki `tools/dump_msproject_typelib.py` ile simetri.

## 8. Testing Strategy

Her commit sonunda **full regression** (`python -m pytest tests/ -v --tb=short -q`):
- Commit 1 sonrası: 83/83 PASS (mevcut testler kontrat değiştirmediği için PASS kalmalı)
- Commit 2 sonrası: 83 + 3-4 yeni unit test = ~87
- Commit 3 sonrası: ~87 + 1 monkeypatch test = ~88
- Commit 4 sonrası: ~88 (tool helper, test gerekmez)

Her commit'in kendi yeni testleri PASS olmalı + mevcut testler regression yok. Tek bir test bile FAIL olursa commit yapılmaz.

## 9. Acceptance (TAIL Cleanup Tamam)

1. ✅ 4 commit landed on main
2. ✅ `python -m pytest tests/ -v` → ~88 PASS, 0 FAIL
3. ✅ Phase 2a calendar functionality unchanged (acceptance script `samples/build_uzbek_calendar.py` hâlâ <5s'de çalışır — opsiyonel manual smoke)
4. ✅ `git push origin main` (4 yeni commit GitHub'a)
5. ⏸ Sonraki adım: Phase 2b brainstorm

## 10. Out of Scope

- `set_working_hours` action (Phase 3+)
- Recurring exception support (Phase 3+)
- Phase 1 fonksiyonlarına `_format_com_error` migration (sadece calendar fonksiyonları — Phase 1 dokunulmaz)
- Resource tool refactoring (Phase 2b)

## 11. Sonraki Adım

`writing-plans` skill → bu design'ı baz alarak T28-T31 (4 commit, ~10-12 implementation task'i) bite-sized plan oluştur. Plan dosyası: `docs/plans/2026-04-28-phase2a-tail-cleanup-impl.md`.

---

*Approved by user: 28 Nisan 2026*
*Next: writing-plans skill → Phase 2a TAIL cleanup implementation plan*
