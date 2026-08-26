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
| `p6_mcp_core.py` | MCP sunucusu (ince dispatcher). Tool'lar: `p6_query`, `p6_job`, `p6_health`, `p6_evm`, `p6_progress`, `p6_baseline` |
| `mcp_common.py` | Paylaşılan katman: redaksiyon, JSON zarfı + **veri-seviyesi kısaltma**, dispatch, kimlik-parametresi reddi |
| `p6/db.py` | Alias çözümleme (bootstrap XML), SQLite/SQL Server salt-okuma backend'leri, snapshot, `connect_rw` (yalnız JOBSVC), `parse_schedule_options` |
| `p6/jobs.py` | **F9 motoru**: `build_job_data`, `submit`, `wait`, `list_jobs`, `cancel`, `purge`, `preflight`, `translate_error` |
| `p6/source.py` | `source` parametresi → tablo torbası (db / xer), `day_hr_cnt` çözümleme |
| `p6/analysis.py` | **Faz 2 ortak yükleyici**: DCMA/EVM şekli, birim seçimi (`resolve_units`), baseline çözümleme, `aggregate` (BAC/PV/EV/AC), S-eğrisi kovaları |
| `p6/health.py` | `p6_health` aksiyonları — `dcma_checks`'e bağlar, kendi hesabı yoktur |
| `p6/evm.py` | `p6_evm` aksiyonları — `evm_math` + `currency_validator`'a bağlar, snapshot deposu |
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

### Faz 2 — `p6_health` (4 action) + `p6_evm` (15 action)

`p6_health`: `assess_all` · `summary` · `drill_down` · `compare`
`p6_evm`: `compute_metrics` · `forecast` · `earned_schedule` · `summary` ·
`time_phased_evm` · `period_delta` · `progress_data_quality` ·
`variance_to_baseline` · `compare_baselines_evm` · `save_period_snapshot` ·
`get_period_history` · `trend` · `detect_currency_mode` ·
`validate_currency_mode` · `verify`

Yeni hesap kodu **yazılmadı**; `dcma_checks` / `evm_math` / `currency_validator`
olduğu gibi kullanılıyor, `msproject_evm` ile aynı action sözlüğü.

Kabul testi `tests/live/test_p6_health_evm.py` — **35 kontrol, 0 hata**.
İddiaya değil kanıta bakıyor:

| Kontrol | Tool | Bağımsız kaynak |
|---|---|---|
| R3 leads | 338 | ham SQL 338 |
| R4 lags | 308 | ham SQL 308 |
| R5 FS % | 74,43 | 1266/1701 = 74,43 |
| R8 negatif float | 9 | ham SQL 9 |
| aktivite | 950 | ham SQL 950 |
| BAC | 70.632 | Σ`target_qty` 70.632 |
| S-eğrisi son kümülatif PV | 70.632 | = BAC (monoton artan, 15 aylık kova) |

db ↔ XER yapısal parite: R1/R2/R3/R4/R5/R9/R11 **birebir**. Float ve kritik
kuralları (7/8/13) veri tarihine bağlıdır — XER 25.08 (F9 öncesi, negatif float 7),
db 26.08 (F9 sonrası, negatif float 9); fark beklenen ve Faz 0 ölçümüyle tutarlı.

Repo test paketi bu dalda **yeniden koşuldu: 890 passed, 279 skipped, 0 fail**
(skip'ler MS Project/Asta COM gerektiren testler).

### Faz 3 -- yazma: `p6_progress` (5 action) + `p6_baseline` (4 action)

`p6_progress`: `read` * `set_progress` * `set_assignment_actuals` * `clear` * `set_data_date`
`p6_baseline`: `list` * `create` * `assign` * `delete`

P6'da otomasyon arayuzu yok; yazma dogrudan veritabanina yapilir, tarihleri
**yalniz P6'nin kendi CPM motoru** hesaplar (`schedule=true` -> Job Service).
Her yazma `confirm=true` ister, hepsinde `dry_run` var ve dry_run alan alan
before/after gosterir.

**Faz 2'de gozden kacmis, Faz 3'te bulunup duzeltilen uc hata:**

1. 🔴 **`complete_pct_type` yok sayiliyordu.** `xer_parser` dogrudan
   `phys_complete_pct` okuyor; bukhtourcity'nin **950 aktivitesinin tamami
   CP_Drtn**, yani yuzde tamamlanma kalan sureden turer. Ilerleme girildikten
   sonra bile EV = 0 cikiyordu. `p6/analysis.resolve_percent_complete` artik
   aktivitenin kendi tabanini kullaniyor (CP_Drtn / CP_Phys / CP_Units) ve
   hangi tabanin kac aktivitede kullanildigini `percent_complete_basis` ile
   raporluyor.
2. 🔴 **Biten is "veri tarihinde bitiyor" gorunuyordu.** `forecast_finish`
   zinciri `reend_date -> early_end_date -> act_end_date` idi; P6 biten
   aktivitede `reend_date`'i bosaltip `early_end_date`'i veri tarihine
   kaydirdigi icin 24.09'da biten is 01.11 olarak donuyordu -- gecikme
   analizinde 38 gun sessizce siliniyordu. Zincir `act_end_date` ile basliyor:
   biten isin bitisi tahmin degil, olgudur.
3. 🔴 **Kaynak yuklu aktivitede kalan sure F9'da geri yaziliyordu.** Sure ile
   birimi baglayan duration_type'ta (burada `DT_FixedDUR2`) P6 kalan sureyi
   **atamanin kalan biriminden** yeniden hesaplar. Olcum: bukhtourcity85'e
   72 saat yazildi, F9 sonrasi 240 saate dondu; atamasi olmayan ayni yapidaki
   bukhtourcity1346 ise degerini korudu. `set_progress` artik atama
   defterlerini (act_reg_qty / remain_qty / maliyet) aktiviteyle birlikte
   tasiyor. `update_assignments=false` verilirse bu davranis uyariyla bildirilir.

**Baseline.** P6'nin Job Service'inde baseline yaratma is tipi YOK; kopya tek
transaction'da SQL ile yapiliyor (PROJECT/PROJPROP/PROJWBS/TASK/TASKPRED/
TASKRSRC, NEXTKEY'den taze id). Kalici test baseline'i: **proj_id 369
"bukhtourcity BL01 Initial"** (Initial Plan, 26.08). Sadakat: 950/950 aktivite
tarih+ad birebir, 0 kopuk bag, 0 kopuk WBS, `project_flag='N'` (EPS'te
gorunmez), canli projeyle ortak task_id yok.

### Testler

| Paket | Kapsam | Sonuc |
|---|---|---|
| `pytest tests/` (cevrimdisi) | 969 test | **969 passed, 279 skipped, 0 fail** |
| `tests/test_xer_encoding_detect.py` | XER kod sayfasi saptama, 13 test | gecti |
| `tests/test_p6_progress_rules.py` | P6 ilerleme semantigi, 33 test | gecti |
| `tests/test_p6_analysis_rules.py` | yuzde tabani / birim / WBS yolu, 33 test | gecti |
| `tests/live/test_p6_health_evm.py` | DCMA + EVM, ham SQL capraz kontrol | **35/35** |
| `tests/live/test_p6_full_acceptance.py` | 6 aracin tamami, uctan uca | **187/187** |

Tam kabul testi veriyi degistirir ve sonunda baslangic durumuna geri alir
(ilerleme temizlenir, test baseline'i silinir, veri tarihi geri konur);
tekrar tekrar calistirilabilir.

### MCP stdio testi
`initialize → p6_mcp 1.26.0` ·
`tools/list → ['p6_query','p6_job','p6_health','p6_evm','p6_progress','p6_baseline']` · `tools/call` ✅

---

## 4. Kurulum ön koşulları (yeni makinede tekrarlanacaksa)

| Adım | Kritik ayrıntı |
|---|---|
| Job Service kurulumu | MSI feature adı **`PrmJob`** (`PrmJobProFeatures` bir dialog adıdır — 2711 hatası verir). `msiexec /i p6pro.msi ADDLOCAL=PrmJob JOBSVCALIAS=<alias> /qb!` — **`/qn` ile servis kaydı ATLANIR** (koşul: `not UILevel=2`), yükseltilmiş çalıştır |
| SQL Server | 2019/2022 test edilmiş; **2025 de çalıştı**. Collation **case-insensitive** zorunlu. Bu makinede `SQL_Latin1_General_CP1_CI_AS` kullanıldı — 🔴 **Kiril metni bozuyor, bkz. §5.1**; Kiril programlar için `Cyrillic_General_CI_AS` seçin (o da CI'dir, P6'nın şartını sağlar) |
| P6 şeması | `dbsetup.bat` → `installppm`. JDK 21'de `--add-opens java.base/java.lang=ALL-UNNAMED` (+`java.util`, `java.lang.reflect`) şart. Log yolu **göreli** olmalı. **Veritabanını dbsetup kendisi yaratır** — önceden `CREATE DATABASE` yapma |
| P6 alias | `Primavera.Launcher.DBconfig.exe /runsilent=Yes /dbtype=SQLServer /alias=<ad> /connectionString=<host,port>/<db> /pusername=<pub> /puserpwd=<parola> /groupid=1`. **`puserpwd` verilmezse parola boş kalır** → `Login failed`. `/bootstrapFile` ve `/runtest` KULLANMA. Aynı isim varsa sessizce `_1`, `_2` ekler |
| 🔴 Bootstrap yayılımı | Servis **LocalSystem** olarak çalışır. `prmbootstrapV2.xml` LocalSystem profilinde de olmalı: `%WINDIR%\System32\config\systemprofile\AppData\Roaming\Oracle\Primavera P6\P6 Professional\[<sürüm>\]` — yoksa **sessizce Oracle sürücüsüne düşer**: `Cannot find OCI DLL` |
| 🔴 `USEROBS` | P6 kurulumu bu tabloyu **boş bırakır**. Kullanıcı global superuser olsa bile proje erişimi OBS'den gelir; yoksa `No projects to schedule`. Çözüm: P6'da *Enterprise ▸ OBS ▸ Users* → `<Project Superuser>` (veya `INSERT INTO USEROBS(user_id,obs_id,prof_id)`) |

---

## 5. Veri bulguları (Faz 2'de bulundu, Faz 3'te ikisi çözüldü)

Üçü de kodda değil **veride**; hiçbiri sessizce düzeltilmedi. 5.1 ve 5.3
Faz 3'te kapatıldı, 5.2 açık.

### ✅ 5.1 Latin1 collation Kiril'i bozuyordu — ÇÖZÜLDÜ (26.08)

`TASK.task_name` = `varchar` + `SQL_Latin1_General_CP1_CI_AS`. Kiril metin bu
kolona sığmıyor ve SQL Server "best-fit" ile en yakın Latin harfi yazıyor:

```
XER ham bayt : C3 F0 E0 ED E8 F2   -> cp1251 -> "Гранит"   (doğru)
DB'de duran  : C3 67 E0 ED E8 F2   -> cp1251 -> "Гgанит"   (bozuk)
                  ^^ 0xF0 ('ğ', cp1254) -> 0x67 ('g')
```

Bayt kaybolduğu için **veritabanından kurtarılamazdı**; tek doğru kaynak
projeyi üreten XER dosyasıdır. Sayısal analiz (DCMA/EVM/float) hiç
etkilenmemişti — yalnız metin.

**Çözüm uygulandı** (yedek: `backup_20260826\PMDB_faz3_oncesi.bak`, 43,8 MB):

1. `_P6_MCP\collation_migrate.py` — veritabanının **tamamını**
   `Cyrillic_General_CI_AS`'e taşır (931 varchar kolon + 32 indeks + 62 check
   constraint drop/recreate + `ALTER DATABASE COLLATE`). Yalnız birkaç kolonu
   çevirmek karışık collation bırakır ve join'lerde "collation conflict"
   üretir — 495 trigger / 210 view'lı bir şemada tek tek denetlemek yerine
   tamamı tek collation'a alındı. Önce yedekten kurulan `PMDB_COLLTEST`
   üzerinde denendi, satır sayıları birebir doğrulandı, sonra PMDB'ye uygulandı.
2. `_P6_MCP
epair_cyrillic.py` — kaybolan baytları XER'den geri yazar.
   TASK `task_code` ile, PROJWBS **kök yolu** ile eşleşir (`wbs_short_name`
   benzersiz DEĞİL: '1','2','3' farklı üst düğümler altında tekrar eder; ilk
   denemede bununla eşleşince "СНАБЖЕНИЕ / PROCUREMENT" düğümüne
   "НАРУЖНЫЕ СЕТИ" adı yazılacaktı).

Sonuç: **950/950 görev adı XER ile bayt bayt aynı**, 529 görev + 785 WBS
Kiril, U+FFFD kalıntısı yok, F9 çalışmaya devam ediyor (8 sn, JS_Complete).

🔴 **Bilinen takas:** `varchar` tek bir kod sayfası taşır. cp1251'de Türkçe'ye
özgü harfler (Ş Ğ İ) YOKTUR; SQL Server en yakın harfi yazar (Ş→S). Kiril
programlar için doğru seçim budur, ama aynı veritabanında Türkçe metin
tutulamaz. Kabul testi bunu açıkça sabitler.

**Ayrıca `xer_parser` düzeltildi:** BOM'suz ANSI XER `utf-8 errors='replace'`
ile okunuyordu, yani cp1251 Kiril'in TAMAMI sessizce U+FFFD oluyordu. Artık
kod sayfası kelime bazlı skorlamayla saptanıyor (gerçek dosyada cp1251 66.022,
cp1252 −44.260), `encoding` parametresiyle ezilebiliyor, eşitlikte düşük güven
işaretleniyor. 13 birim testi: `tests/test_xer_encoding_detect.py`.

### 🔴 5.2 CLI import kaynak ücretlerini düşürüyor (AÇIK)

XER'de `TASKRSRC.cost_per_qty = 5,00` ve `target_cost = target_qty × 5`.
Aynı XER CLI ile import edildikten sonra DB'de `cost_per_qty = 0`,
`target_cost = 0`, `RSRCRATE.cost_per_qty = 0` (2/2 kaynak).

Yani **aynı program iki kaynakta iki farklı BAC veriyor**: XER 353.160 (maliyet),
DB 70.632 (saat) — tam 5× fark. `p6_evm` bunu gizlemiyor: her yanıtta `units`,
`units_reason` ve üç adayın BAC'ı (`candidate_bac`) birlikte dönüyor, `verify`
ise BAC'ı ham `target_qty` toplamıyla çapraz doğruluyor (RULE 16.A).
Faz 4 (`p6_cli` parity) bu kaybı kapatmadan XER→import→F9 döngüsü maliyet
tarafında güvenilir değildir.

### ✅ 5.3 P6'nın "planned" tarihleri baseline DEĞİLDİR — gerçek baseline üretildi

bukhtourcity'de **950 aktivitenin 950'sinde** `target_start_date = early_start_date`
ve `target_end_date = early_end_date`. P6 başlamamış aktivitelerin planned
tarihlerini mevcut programa senkronlar; canlı bir veritabanında "baseline
verilmezse target tarihleri kullan" yaklaşımı programı kendisiyle kıyaslar →
SPI ≈ 1, gecikme ≈ 0, sonsuza kadar.

`p6/analysis.py` bu oranı **ölçüyor** (varsaymıyor) ve %95 üstündeyse
`baseline_warnings` ile açıkça uyarıyor.

Faz 3'te `p6_baseline action='create'` eklendi ve **kalıcı bir baseline
üretildi: proj_id 369 "bukhtourcity BL01 Initial"** (Initial Plan, 26.08).
`p6_evm action='variance_to_baseline' baseline_proj_id=369` artık gerçek bir
karşılaştırma yapıyor: 950 aktivite eşleşti, 0 eşleşmeyen, uyarı yok.

---

## 6. Bilinen tuzaklar

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

- 🔴 **`mcp_common.json_response` en uzun listeyi kırpar.** DCMA `assess_all`
  ham hâlde 61 KB (kural 3/4/5 içindeki `failed_links` 18-22 KB) → kırpılan liste
  `rules`'un kendisi oluyordu ve çağıran **14 kural yerine 1 kural** alıyordu,
  hata da vermeden. `p6/health.py:_slim` id/link listelerini 10'luk örneğe
  indiriyor, tam liste `drill_down`'da. Aynı tuzak diğer sunucularda da vardır —
  büyük iç liste taşıyan her yanıt için kontrol edin.

---

## 7. Sıradaki işler

1. **`p6_compare`**: `xer_compare.py` (revizyon delta).
2. **`p6_write` / `p6_cli` / `p6_revision`**: MPXJ ile XER/PMXML yaz → CLI ile
   revizyon projesi olarak import → F9 → parity.
   **§5.2'deki maliyet kaybı bu adımda çözülmeli.**
3. **`mcp_common.py`'yi diğer 3 sunucuya taşı** — JSON kısaltma düzeltmesi
   ve §6'daki kırpma tuzağı orada da geçerli.
4. `JT_Level` / `JT_Sum` / `JT_ApplyActuals` / `JT_UpdateBaseline` gerçek veriyle doğrula.
5. `PrmJob.Job` COM `Execute`'u `comtypes` ile yeniden dene (pywin32 `VT_BYREF`
   OUT parametrelerini kabul etmiyor; servis kuyruğu çalıştığı için bloklayıcı değil).

## 8. Doğrulanmamış / riskli

- `JT_Level`, `JT_Sum`, `JT_ApplyActuals`, `JT_UpdateBaseline`, `JT_XERExport` —
  kod yolu hazır, **çalıştırılmadı**.
- `compare_baselines_evm` **iki baseline ile denenmedi** — veritabanında tek
  baseline var (369). `variance_to_baseline` gerçek baseline ile doğrulandı.
- **Türkçe karakter taşıyan bir P6 programı denenmedi** — collation cp1251;
  Türkçe'ye özgü harfler bu veritabanında tutulamaz (§5.1 takası).
- P6 Professional arayüzü **açılıp bakılmadı**: yazılan ilerlemenin ve
  baseline'ın P6 GUI'sinde nasıl göründüğü doğrulanmadı; tüm doğrulama
  veritabanı + Job Service üzerinden yapıldı.
- `p6/db.py`'nin **SQLite backend'i** P6 24.12 SQLite şemasında test edildi;
  SQL Server yolu asıl kullanılan.
- `snapshot()` fallback dalı (3 dosya kopyası) hiç tetiklenmedi — `VACUUM INTO`
  her seferinde çalıştı.
