# MS Project MCP — Phase 2a Calendar Design

**Versiyon:** 1.0
**Tarih:** 27 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude
**Status:** Approved → writing-plans next
**Phase:** 2a (Calendar). Phase 2b (Resource) onay sonrası başlayacak.

---

## 1. Hedef

`msproject_mcp` server'ına `msproject_calendar` tool'unu ekle. Phase 2a sonunda kullanıcı tek mesajla:

> "Yeni bir 'Uzbekistan-2026' takvimi yarat, 9 resmi bayramı ekle, bütün proje task'larına ata."

talimatını verebilmeli ve MS Project UI'da takvim + exception'lar + task assignment görünür olmalı.

## 2. Karar Geçmişi (Brainstorming Çıktıları)

### Q1 — Scope ve Sıra: A
Calendar önce, Resource sonra. Calendar foundational (resource'ların `BaseCalendar`'ı var, task duration calendar'a göre genişler). Phase 2a → onay → Phase 2b.

### Q2 — Action Surface: D
6 core action + `holidays_uzbek` bonus = **7 action**. `set_working_hours` Phase 2a'da YOK (WeekDays + WorkWeeks + Periods zinciri bug-prone, default Standard'tan kopya çoğu use case'i karşılıyor). Phase 3'e ertelendi.

## 3. Tool Surface

**Tool:** `msproject_calendar`

| # | Action | Parametreler | Çıktı |
|---|---|---|---|
| 1 | `create` | `name` (str), `base_calendar` (str = "Standard") | `{status, calendar_uid, name}` |
| 2 | `update` | `name` (str), `new_name` (opt), `weekday_off` (opt: int 0-6) | `{status, changes}` |
| 3 | `add_exception` | `calendar_name` (str), `exception_name` (str), `start` (date), `finish` (opt date), `working` (bool=False) | `{status, exception_name}` |
| 4 | `assign_to_task` | `task_id` (int), `calendar_name` (str) | `{status, task_id, calendar_name}` |
| 5 | `assign_to_resource` | `resource_id` (int), `calendar_name` (str) — Phase 2b sonrası anlamlı | `{status, resource_id, calendar_name}` |
| 6 | `list` | (yok) | `{status, calendars: [{name, base_calendar, exceptions: [...]}]}` |
| 7 | `holidays_uzbek` | `calendar_name` (str), `year` (int=2026) | `{status, count, holidays: [{name, date}]}` |

## 4. Built-in: Özbekistan 2026 Resmi Tatilleri

`holidays_uzbek` action'unun içinde sabit liste:

| # | Tarih | Tatil Adı |
|---|---|---|
| 1 | 1 Ocak | Yılbaşı |
| 2 | 14 Ocak | Vatan Müdafaası Günü |
| 3 | 8 Mart | Kadınlar Günü |
| 4 | 21 Mart | Navruz |
| 5 | 1 Mayıs | İşçi Bayramı |
| 6 | 9 Mayıs | Hatıra ve Şeref Günü |
| 7 | 1 Eylül | Bağımsızlık Günü |
| 8 | 1 Ekim | Öğretmenler Günü |
| 9 | 8 Aralık | Anayasa Günü |

**Notlar:**
- Ramazan/Kurban Bayramı hicri takvim — Phase 2a'ya dahil DEĞİL. İleride parametre eklenebilir (`include_islamic=True` + hicri-Gregoryen converter).
- Tatil bir hafta sonuna denk gelirse "öbür güne taşıma" mantığı YOK (Asta'da da yok, kullanıcı bilerek).
- Yıl parametresi default 2026, ama 2027/2028 için aynı liste tarihleri çoğunlukla aynı (Navruz 21 Mart sabit, Bağımsızlık 1 Eylül sabit).

## 5. Implementation Yaklaşımı

### COM-first, MSPDI bulk YOK

Calendar volume düşük (proje başına 1-3 takvim, ≤30 exception). Hibrit routing burada gereksiz overhead. Phase 1'de yazılan `_route_operation()` Calendar tool'da kullanılmaz.

### COM API kullanımı

**Calendar create:**
```python
# Yüksek seviye Application metodu — tercih edilen
app.BaseCalendarCreate(Name="Uzbekistan-2026", FromName="Standard")
# Sonra Calendars collection'dan al
cal = proj.BaseCalendars("Uzbekistan-2026")
```

**Exception ekleme:**
```python
# Tek günlük tatil için
ex = cal.Exceptions.Add(
    Type=7,           # pjDaily — daily, single occurrence
    Start=pywintypes.Time(date(2026, 1, 1)),
    Finish=pywintypes.Time(date(2026, 1, 1)),
)
ex.Name = "Yılbaşı"
ex.Shift1Start = None  # working=False → çalışılmıyor
```

Type kodları (typelib'den):
- 1 = Daily (günlük periyodik)
- 7 = Daily, fixed range (en yaygın — tek tarih veya tarih aralığı)

Phase 2a'da sadece Type=7 kullanılacak (recurring exceptions Phase 3+).

**Update — weekday off:**
```python
# Pazar (1=Pazar, 7=Cumartesi MSP'de)
weekday = cal.WeekDays(1)  # Sunday
weekday.Working = False
```

**Assign to task:**
```python
task.Calendar = "Uzbekistan-2026"
```

**Assign to resource:**
```python
resource.BaseCalendar = "Uzbekistan-2026"
```

### Helper fonksiyonlar (yeni)

`msproject_mcp_core.py`'a eklenecek:
- `_find_calendar_by_name(proj, name) -> Calendar | None`
- `_msp_calendar_create(name, base_calendar) -> dict`
- `_msp_calendar_update(name, new_name, weekday_off) -> dict`
- `_msp_calendar_add_exception(calendar_name, exception_name, start, finish, working) -> dict`
- `_msp_calendar_assign_to_task(task_id, calendar_name) -> dict`
- `_msp_calendar_assign_to_resource(resource_id, calendar_name) -> dict`
- `_msp_calendar_list() -> dict`
- `_msp_calendar_holidays_uzbek(calendar_name, year) -> dict`

Built-in tatil listesi sabit constant olarak:
```python
UZBEK_HOLIDAYS_2026 = [
    ("Yılbaşı", 1, 1),
    ("Vatan Müdafaası Günü", 1, 14),
    ("Kadınlar Günü", 3, 8),
    ("Navruz", 3, 21),
    ("İşçi Bayramı", 5, 1),
    ("Hatıra ve Şeref Günü", 5, 9),
    ("Bağımsızlık Günü", 9, 1),
    ("Öğretmenler Günü", 10, 1),
    ("Anayasa Günü", 12, 8),
]
```

## 6. Test Stratejisi

### Phase 1 SAFETY Pattern Korunur
Tüm calendar testleri `clean_test_project` fixture'ından FileNew ile izole proje açar. **Kullanıcının aktif projesi DOKUNULMAZ.** Phase 1 SAFETY FIX disiplini Phase 2a'da da geçerli.

### Yeni fixture
`clean_test_calendar(clean_test_project)` — test sonunda oluşturulan custom calendar'ları siler:
```python
@pytest.fixture
def clean_test_calendar(clean_test_project):
    proj = clean_test_project.ActiveProject
    yield proj
    # Teardown: silinmesi gereken custom calendar isimlerini topla
    # (test başlangıcında snapshot al, sonunda farkı sil)
```

### Test dosyaları (yeni)

| Dosya | Test sayısı | Coverage |
|---|---|---|
| `tests/test_msproject_calendar_create.py` | 2 | create from Standard, name conflict error |
| `tests/test_msproject_calendar_update.py` | 2 | rename, weekday_off |
| `tests/test_msproject_calendar_exception.py` | 3 | single date, date range, exception not found error |
| `tests/test_msproject_calendar_assign.py` | 2 | assign to task, calendar_name not found error |
| `tests/test_msproject_calendar_list.py` | 1 | base + custom listed correctly |
| `tests/test_msproject_calendar_uzbek.py` | 2 | 9 holidays added, dates verified |

**Toplam:** ~12 yeni test.

### Performans Hedefi
- Tüm Phase 2a suite: **<15 saniye**
- `holidays_uzbek` tek call: **<2 saniye** (9 exception)
- Mevcut 43 Phase 1 test: **regression check — PASS kalmalı**

## 7. Hata Yönetimi

| Senaryo | Davranış |
|---|---|
| Calendar adı zaten var | `{status:"error", error:"Calendar 'X' already exists"}` |
| Base calendar yok | `{status:"error", error:"Base calendar 'X' not found"}` |
| Calendar adı assign sırasında yok | `{status:"error", error:"Calendar 'X' not found in project"}` |
| Task ID assign sırasında yok | `{status:"error", error:"Task ID 99 not found"}` |
| Exception start > finish | `{status:"error", error:"Start date must be <= finish date"}` |
| Active project yok | Phase 1 `_validate_active_project` reuse |

Tüm hata mesajları **eylem önerisi** içermeli (Phase 1 disiplini).

## 8. Acceptance Kriterleri (Phase 2a Tamam)

1. ✅ `msproject_calendar` tool tüm 7 action ile çalışıyor
2. ✅ Acceptance script `samples/build_uzbek_calendar.py` başarılı:
   - "Uzbekistan-2026" calendar yarat
   - 9 Uzbek bayramı bulk ekle
   - 1 task'a ata
   - MS Project UI'da Tools → Change Working Time → 9 exception görünür
3. ✅ Phase 2a yeni testleri (~12) PASS
4. ✅ Phase 1 mevcut 43 test regression PASS
5. ✅ Kullanıcının aktif projesi hiç dokunulmadı (clean_test_project disiplini)
6. ✅ Commit + push GitHub'a
7. ⏸ Kullanıcı manuel onayı → Phase 2b (Resource) başlar

## 9. Out of Scope (Phase 2a'da YOK)

- `set_working_hours` action (WeekDays/WorkWeeks/Periods zinciri) → Phase 3+
- Recurring exceptions (haftalık/aylık periyodik) → Phase 3+
- İslami bayramlar (Ramazan/Kurban) hicri → Phase 3+
- Multiple year holiday packs → şu an default 2026, ileride
- Resource calendar assignment kullanımı (sadece API hazır, gerçek demo Phase 2b'de)
- MSPDI bulk path (volume düşük, gereksiz)

## 10. Sonraki Adım

`writing-plans` skill'i invoke edilecek → bu design doc baz alınarak T18-T27 (~8 task) bite-sized TDD implementation plan oluşturulacak. Plan ayrı dosya: `docs/plans/2026-04-27-phase2a-calendar-impl.md`.

---

*Approved by user: 27 Nisan 2026*
*Next: writing-plans skill → Phase 2a implementation plan*
