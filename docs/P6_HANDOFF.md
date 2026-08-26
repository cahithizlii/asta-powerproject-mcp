# P6 Professional MCP — Devir Notu

Son güncelleme: **26.08.2026** · Dal: `feat/p6-mcp` · P6 Professional **24.12.0.51267**

Bu belge repo-göreli yollar kullanır. Makineye özgü kurulum yolları ve yedekler
yerel çalışma klasöründedir (bkz. memory: `project_p6_mcp_server.md`).

---

## 1. Neden bu mimari — üç kanıtlı kısıt

P6, Asta ve MS Project'ten temelde farklı; onlardaki "canlı uygulamaya COM ile bağlan,
yaz, reschedule et" modeli **P6'da mümkün değil**:

1. **`PM.exe` otomasyona kapalı.** Dışa açık COM/ActiveX arayüzü yok.
2. **Komut satırı schedule edemez.** `Primavera.CacheService.exe`'nin desteklediği tüm
   eylemler `PM.exe` ikilisinden çıkarıldı: **`import` · `export` · `batchrpt`**.
   `import` üstelik yalnız `CREATE` — mevcut projeyi güncelleyemez.
3. **Launcher↔PM iç API'sinde schedule opcode'u yok**
   (`API_LOGIN_WITH_APP`, `API_OPEN_PROJECTS`, `API_Process_ImportExport`, … — hepsi tarandı).

**Çözüm: P6'nın kendi Job Service'i.** `JOBSVC` kuyruk tablosuna `RT_ASAP` satırı bırakılır,
`PrmJobSv.exe` saniyeler içinde alır ve **gerçek P6 CPM motorunu** çalıştırır.

### 🔴 Job Service SQLite'ta ÇALIŞMAZ

`prmjob.exe` ve `PrmJobSv.exe` içinde, string tablosunda `SQLite` sabitinin hemen ardından:

> `Job Services are not supported for P6 Professional Standalone.`

Bu yüzden veritabanı **SQL Server**'a taşındı. Standalone SQLite alias'ı bozulmadan duruyor.

---

## 2. Kod haritası (repo-göreli)

| Dosya | Rol |
|---|---|
| `p6_mcp_core.py` | MCP sunucusu (ince dispatcher). Tool'lar: `p6_query`, `p6_job` |
| `mcp_common.py` | Paylaşılan katman: redaksiyon, JSON zarfı + **veri-seviyesi kısaltma**, dispatch, kimlik-parametresi reddi |
| `p6/db.py` | Alias çözümleme (bootstrap XML), SQLite/SQL Server salt-okuma backend'leri, snapshot, `connect_rw` (yalnız JOBSVC), `parse_schedule_options` |
| `p6/jobs.py` | **F9 motoru**: `build_job_data`, `submit`, `wait`, `list_jobs`, `cancel`, `purge`, `preflight`, `translate_error` |
| `p6/source.py` | `source` parametresi → tablo torbası (db / xer), `day_hr_cnt` çözümleme |
| `tests/live/` | Canlı kabul testleri (P6 + SQL Server gerektirir) |

**Yeniden kullanılan, DEĞİŞTİRİLMEYEN modüller:** `xer_parser.py`, `xer_compare.py`,
`xer_drivers.py`, `dcma_checks.py`, `evm_math.py`, `currency_validator.py`,
`excel_io.py`, `report_builder.py`.

Anahtar tasarım: `p6/db.py` SQL satırlarını kolon-`lower()` + değer-`str()` yapıp
`{"TASK": {"headers": [...], "rows": [...]}}` torbası kurar ve **`xer_parser`'ın
okuyucularını unbound çağırır**. Böylece XER alan eşlemesi — özellikle
`forecast_finish = reend_date` (RULE 16.B) — tek yerde kalır.

---

## 3. Kanıtlanmış durum

### Faz 0 — headless F9 (26.08.2026 09:29)

```
JS_Pending → JS_Running → JS_Complete   last_error_descr = "OK"   ~3 sn
```
Data date 90 gün ileri alındığında **950 aktivitenin tamamı yeniden hesaplandı**:
`ES_min` 2026-05-20 → 2026-08-18 · `EF_max` 2027-07-30 → 2027-10-23 ·
`SUM(TOTAL_FLOAT)` 961.936 → 1.417.736 · negatif float'lı aktivite 7 → 9.

### `p6_job` — 13 action

`schedule` · `level` · `summarize` · `apply_actuals` · `update_baseline` ·
`status` · `wait` · `list` · `cancel` · `purge` · `preflight` · `service_health` · `job_data`

Tool üzerinden `schedule`: **6,1 sn**, `JS_Complete`, tarihler değişti.
Yalnız **`JT_Sched` gerçek veriyle kanıtlandı**; diğer iş tipleri aynı reçeteyi
kullanıyor ama **denenmedi**.

### Faz 1 — `p6_query` — 14 action

`list_projects` · `list_eps` · `read_tasks` · `read_links` · `read_resources` ·
`read_assignments` · `read_calendars` · `read_wbs` · `read_project` · `read_progress` ·
`finish_drivers` · `schedule_options` · `sql` · `db_info`

- `read_tasks`/`read_links` → 950 / 1701, bağımsız SQL sayımıyla **birebir**
- **db ve XER kaynağı aynı sayıyı veriyor** (950 = 950)
- `day_hr_cnt` takvimden: `CALENDAR:CLNDR_ID=638 (Akfa HQ Project 7x8)` — asla varsayılmıyor
- `sql` yazma denemeleri (DELETE/UPDATE/EXEC/çoklu ifade/DDL) → **5/5 reddedildi**

### MCP stdio testi
`initialize → p6_mcp 1.26.0` · `tools/list → ['p6_query','p6_job']` · `tools/call` ✅

---

## 4. Kurulum ön koşulları (yeni makinede tekrarlanacaksa)

| Adım | Kritik ayrıntı |
|---|---|
| Job Service kurulumu | MSI feature adı **`PrmJob`** (`PrmJobProFeatures` bir dialog adıdır — 2711 hatası verir). `msiexec /i p6pro.msi ADDLOCAL=PrmJob JOBSVCALIAS=<alias> /qb!` — **`/qn` ile servis kaydı ATLANIR** (koşul: `not UILevel=2`), yükseltilmiş çalıştır |
| SQL Server | 2019/2022 test edilmiş; **2025 de çalıştı**. Collation **case-insensitive** zorunlu (`SQL_Latin1_General_CP1_CI_AS`) |
| P6 şeması | `dbsetup.bat` → `installppm`. JDK 21'de `--add-opens java.base/java.lang=ALL-UNNAMED` (+`java.util`, `java.lang.reflect`) şart. Log yolu **göreli** olmalı. **Veritabanını dbsetup kendisi yaratır** — önceden `CREATE DATABASE` yapma |
| P6 alias | `Primavera.Launcher.DBconfig.exe /runsilent=Yes /dbtype=SQLServer /alias=<ad> /connectionString=<host,port>/<db> /pusername=<pub> /puserpwd=<parola> /groupid=1`. **`puserpwd` verilmezse parola boş kalır** → `Login failed`. `/bootstrapFile` ve `/runtest` KULLANMA. Aynı isim varsa sessizce `_1`, `_2` ekler |
| 🔴 Bootstrap yayılımı | Servis **LocalSystem** olarak çalışır. `prmbootstrapV2.xml` LocalSystem profilinde de olmalı: `%WINDIR%\System32\config\systemprofile\AppData\Roaming\Oracle\Primavera P6\P6 Professional\[<sürüm>\]` — yoksa **sessizce Oracle sürücüsüne düşer**: `Cannot find OCI DLL` |
| 🔴 `USEROBS` | P6 kurulumu bu tabloyu **boş bırakır**. Kullanıcı global superuser olsa bile proje erişimi OBS'den gelir; yoksa `No projects to schedule`. Çözüm: P6'da *Enterprise ▸ OBS ▸ Users* → `<Project Superuser>` (veya `INSERT INTO USEROBS(user_id,obs_id,prof_id)`) |

---

## 5. Bilinen tuzaklar

- **MPXJ P6 verisini yanlış okur.** `PrimaveraDatabaseFileReader` 82 projeden 1'ini döndürüp
  `NullPointerException` verdi; XER'de `PROJWBS` satırlarını aktivite sayıyor (1735 vs 950).
  → DB/XER okuması için **kendi `xer_parser`'ımız** kullanılır. MPXJ yalnız XER/PMXML **yazma**
  ve CPM ön hesabı için.
- **XER charset = UTF-16LE.** MPXJ varsayılanı Windows-1252 ve Kiril'i yok ediyor
  (`Кирилл` → `??????`), UTF-8'i de kendi parser'ımız okuyamıyor. Kanıtlı karar: UTF-16LE.
- **PMXML kaynağı bilerek kapalı** — ham parser yazılana kadar sessizce yanlış sayı
  döndürmemek için `p6/source.py` açık hata veriyor.
- **`/encryptpassonly` bu derlemede desteklenmiyor** (-1 döner ve süreç asılır) → `/password` kullanılıyor.
- CLI import sonrası `PM.EXE` access violation verir — çıkış kodundan SONRA, kozmetik.
- `p6_job` veritabanına yalnız **`JOBSVC` + `NEXTKEY`** yazar; program verisine dokunmaz.
  Yazan tek şey P6'nın kendi scheduler'ıdır.

---

## 6. Sıradaki işler

1. **Faz 2 — `p6_health`** (DCMA 14-Point): mevcut `dcma_checks.py`'ye bağla, yeni hesap kodu yazma.
2. **Faz 2 — `p6_evm`**: `evm_math.py` + `currency_validator.py`'ye bağla; `msproject_evm` ile aynı action sözlüğü.
3. **`p6_compare`**: `xer_compare.py` (revizyon delta).
4. **`p6_write` / `p6_cli` / `p6_revision`**: MPXJ ile XER/PMXML yaz → CLI ile revizyon projesi olarak import → F9 → parity.
5. **`mcp_common.py`'yi diğer 3 sunucuya taşı** — özellikle JSON kısaltma düzeltmesi orada da gerekli.
6. `JT_Level` / `JT_Sum` / `JT_ApplyActuals` / `JT_UpdateBaseline` gerçek veriyle doğrula.
7. `PrmJob.Job` COM `Execute`'u `comtypes` ile yeniden dene (pywin32 `VT_BYREF` OUT parametrelerini kabul etmiyor; servis kuyruğu çalıştığı için bloklayıcı değil).

## 7. Doğrulanmamış / riskli

- `JT_Level`, `JT_Sum`, `JT_ApplyActuals`, `JT_UpdateBaseline`, `JT_XERExport` — kod yolu hazır, **çalıştırılmadı**.
- `p6/db.py`'nin **SQLite backend'i** P6 24.12 SQLite şemasında test edildi; SQL Server yolu asıl kullanılan.
- `snapshot()` fallback dalı (3 dosya kopyası) hiç tetiklenmedi — `VACUUM INTO` her seferinde çalıştı.
- Repo'nun mevcut 1168 testi bu dalda **yeniden çalıştırılmadı**; yeni dosyalar mevcut testlere dokunmuyor
  ama `mcp_common.py` diğer sunuculara taşınırsa tam suite koşulmalı.
