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
| `p6_mcp_core.py` | MCP sunucusu (ince dispatcher). 10 tool: `p6_query`, `p6_job`, `p6_health`, `p6_evm`, `p6_progress`, `p6_baseline`, `p6_compare`, `p6_write`, `p6_cli`, `p6_task` |
| `mcp_common.py` | Paylaşılan katman: redaksiyon, JSON zarfı + **veri-seviyesi kısaltma**, dispatch, kimlik-parametresi reddi |
| `p6/db.py` | Alias çözümleme (bootstrap XML), SQLite/SQL Server salt-okuma backend'leri, snapshot, `connect_rw` (yalnız JOBSVC), `parse_schedule_options` |
| `p6/jobs.py` | **F9 motoru**: `build_job_data`, `submit`, `wait`, `list_jobs`, `cancel`, `purge`, `preflight`, `translate_error` |
| `p6/source.py` | `source` parametresi → tablo torbası (db / xer), `day_hr_cnt` çözümleme |
| `p6/analysis.py` | **Faz 2 ortak yükleyici**: DCMA/EVM şekli, birim seçimi (`resolve_units`), baseline çözümleme, `aggregate` (BAC/PV/EV/AC), S-eğrisi kovaları |
| `p6/health.py` | `p6_health` aksiyonları — `dcma_checks`'e bağlar, kendi hesabı yoktur |
| `p6/evm.py` | `p6_evm` aksiyonları — `evm_math` + `currency_validator`'a bağlar, snapshot deposu |
| `p6/write.py` | **Yazma oturumu**: tek işlem = tek transaction, NEXTKEY'den blok id, denetim kolonu damgası, `require_confirm` |
| `p6/progress.py` | `p6_progress` — ilerleme/fiili giriş, P6 alan tutarlılığı, atama defteri senkronu |
| `p6/baseline.py` | `p6_baseline` — baseline kopyala/ata/sil **+ `revision`** (kopyayı gerçek proje bırakır) |
| `p6/compare.py` | `p6_compare` — iki programı **`task_code`** üzerinden karşılaştırır |
| `p6/writer.py` | `p6_write` — veritabanından XER yazar, yazdığını geri okuyup doğrular |
| `p6/cli.py` | `p6_cli` — P6 komut satırıyla XER import + import'un düşürdüğü ücretlerin onarımı |
| `p6/tasks.py` | `p6_task` — **Faz 6 CRUD**: sıfırdan proje, WBS, aktivite, bağ, atama; yapısal varsayılanlar projenin modal değerinden; yeni aktivite tarihsiz yazılır, tarihleri F9 hesaplar |
| `tests/live/` | Canlı kabul testleri (P6 + SQL Server gerektirir) |

⚠️ **`p6/write.py` ≠ `p6/writer.py`.** Birincisi yazma *oturumu* (transaction,
anahtar ayırma), ikincisi *XER yazıcısı*. Adları benzer, işleri ayrı — import
ederken karıştırmayın.

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

### Faz 4 -- `p6_compare` (5 action)

`summary` * `tasks` * `links` * `progress` * `evm`

Aritmetik yine paylasilan `xer_compare`; eklenen tek sey P6 icin
karsilastirmayi gecerli kilan sey:

🔴 **Eslesme anahtari `task_code`, `task_id` DEGIL.** P6 bir program her sinir
gectiginde id'leri yeniden numaralandirir -- ayni aktivite XER'de 3274452,
CLI import sonrasi veritabaninda 35847, o projenin baseline'inda ucuncu bir
sayidir. id ile eslesen bir diff 950 aktivitenin 950'sini birden "silinmis" ve
"eklenmis" gosterir: kesin gorunen ama hicbir sey soylemeyen bir rapor.
Kabul testi bunu dogrudan olcuyor (baseline kopyasiyla canli projenin ortak
task_id'si 0, buna ragmen eklenen/silinen 0).

Iki taraf farkli birimde, farkli takvimde ya da farkli veri tarihindeyse
**uyari veriliyor, sessizce cikarilmiyor** -- XER (maliyet, 353.160) ile
veritabani (saat, 70.632) karsilastirilinca §5.2 kaybi tam da boyle yuzeye
cikiyor.

### P6'nin kendi motoruyla dogrulama (arayuz degil)

"P6'da da kontrol et" sorusunun dogru cevabi ekran goruntusu degil, P6'nin
**kendi hesaplama motoru**: Job Service (`prmjob.exe`) veriyi okuyup CPM
calistiriyor, sonuc olculebiliyor.

| Olcum | Sonuc |
|---|---|
| bukhtourcity85'e yazilan kalan sure | 72 saat |
| P6'nin planladigi kalan is penceresi | 2026-11-01 -> 2026-11-09 = **9 is gunu** = 72s ÷ 8s/gun |
| bukhtourcity1346'ya yazilan | 30 saat -> P6 4 gun planladi (30 ÷ 8 = 3,75) |
| Tamamlananlarda acik kalan-is penceresi | 0 |
| Baslamamis isin veri tarihinden once planlanmasi | 0 |

Ayrica P6'nin biten aktivitede `early_end_date`'i veri tarihine kaydirdigi
**canli veride tekrar goruldu** -- Faz 3'teki `forecast_finish` duzeltmesinin
gerekcesi hala gecerli, ve duzeltme sayesinde tahmini bitis fiili bitise esit
raporlaniyor.

**`JT_XERExport` denendi ve KISMEN cozuldu:** is tipi kuyruktan aliniyor,
`JS_Pending -> JS_Running -> JS_Failed` ve uygulama seviyesinde anlamli bir
hata donuyor: `File name not specified.` Yani P6'nin export kodu calisiyor
ama dosya adinin nereden okundugu belgesiz. Denenen ve **calismayan**
yollar: JOB_DATA bolum parametresi, kok anahtar (`File Name` / `FileName` /
`Filename`), proje dugumu parametresi, `JOBSVC.audit_file_path`,
`JOBSVC.recur_data`. CLI action script'i yalniz **import** ogeleri tasiyor
(`importFormat`/`importType`/`importAction`/`importTo`/`importFile`); PM.EXE
icinde `exportFile`/`exportFormat` karsiliklari YOK. Referans blob'u
yakalamanin bilinen tek yolu, P6 arayuzunde Job Services penceresinden bir
kez export isi olusturup JOBSVC satirini okumak.

### Faz 5 -- `p6_write` (XER export) ve uc gizli hata

`JT_XERExport` cikmaza girdigi icin export dogrudan veritabanindan yaziliyor:
`p6/writer.py`, UTF-16LE + BOM, P6'nin kendi tablo sirasinda. `verify=true`
dosyayi **kendi parser'imizla geri okur** ve satir sayilarini, baslik
tekilligini, U+FFFD kalintisini denetler.

Parite kaniti: yazilan XER ile kaynak veritabani `p6_compare` ile
karsilastirildiginda **950/950 aktivite ayni, 0 eklenen / 0 silinen /
0 degisen, bitis kaymasi 0, uyari yok**.

**Bu is uc hata ortaya cikardi; ucu de sessiz sinifindandi:**

1. 🔴 **`SqlServerBackend.columns()` her kolonu IKI KEZ donduruyordu.** P6
   semasi her `dbo` tablosunun uzerine bir `privuser` VIEW'i kuruyor (burada
   164 adet) ve sorgu `INFORMATION_SCHEMA.COLUMNS`'u yalnizca tablo ADIYLA
   suzuyordu. Okumalar sagkalmisti (satirlar kolon adiyla anahtarlanan bir
   sozluge giriyor, tekrarlar birbirini eziyor) ama **Faz 1'den beri her
   sorgu her kolonu iki kez cekiyordu** ve XER yazici bu listeyi dogrudan
   bozuk bir `%F` basligina cevirdi. Duzeltme: `OBJECT_ID` ile tek nesneye
   cozumleme. Dosya boyutu 3,3 MB -> 1,7 MB.
2. 🔴 **ERMHDR'de para birimi yanlis alandan okunuyordu.** Gercek bir P6
   export'u sekiz alan tasir ve para birimi SONDADIR; parser bes alan varsayip
   5. alani aliyordu. bukhtourcity.xer'de bu **"Izzat Islomov"** demekti --
   `currency_validator.extract_currency_code` bir kisi adini para birimi kodu
   olarak donduruyordu, ve cost/hours karari bu koda bakiyor.
3. 🔴 **`Decimal("0E-8")` XER'e bilimsel gosterimle yaziliyordu.** SQL Server
   sifirlanmis bir numeric kolonu tam olarak boyle veriyor (RSRCRATE
   .cost_per_qty bunlardan biri) ve `str()` bunu `0E-8` yapiyor -- bir XER
   importer'inin anlamasi icin sebep yok. `format(v, "f")` ile duzeltildi.

### §5.2 tehisi ilerledi (hala ACIK)

CLI import'un kaynak ucretlerini dusurmesi bir P6 hatasi degil: import
**hicbir import konfigurasyonu olmadan** kosmustu. `importConfiguration`
ogesi `VIEWPROP` tablosundaki `view_type='VP_IMP_OPT'` satirlarina cozuluyor
ve bu veritabaninda **sifir tane VP_IMP_OPT satiri var** -- P6 varsayilanlari
uyguladi: kaynagi ekle, ucretini sifirla. Kanit: DB'deki RSRCRATE satirlari
dogru `rsrc_id`, `max_qty_per_hr` ve `start_date` tasiyor, yalnizca
`cost_per_qty` sifir. (`admin` global superuser, yani yetki sorunu degil.)

Cozum icin iki yol: (a) VP_IMP_OPT satirini uretmek -- `view_data` kodlamasi
belgesiz, referans P6 arayuzunden yakalanmali; (b) import sonrasi ucretleri
kaynak XER'den geri yazmak -- Kiril onariminda kullanilan, kanitlanmis desen.

### Faz 5b -- `p6_cli`: import, ucret onarimi ve iki kesin kodlama bulgusu

**§5.2 KAPANDI.** CLI import'un kaynak ucretlerini dusurmesi yeniden uretildi
ve onarildi:

| Adim | Sonuc |
|---|---|
| Orijinal XER'i CLI ile import | yeni proje, 442 atama, `target_qty` 70.632 dogru |
| Import sonrasi `target_cost` | **0** (ucreti olan atama: 0/442) -- §5.2 birebir tekrarlandi |
| `p6_cli action='repair_costs'` | 2 kaynak ucreti + 442 atama duzeltildi |
| Onarim sonrasi proje toplami | **353.160** = kaynak XER toplami **353.160** ✅ |

Onarim id ile degil **is anahtariyla** eslesir: kaynaklar `rsrc_short_name`,
atamalar (aktivite kodu, kaynak kisa adi) ikilisiyle. P6 import'ta her id'yi
yeniden numaralandirdigi icin id eslesmesi sessizce yanlis satiri tutardi --
baseline kopyasinin ve `p6_compare`'in kacindigi ayni tuzak.

**🔴 P6'nin import'u UTF-16LE XER'i REDDEDER.** Devir notunun onceki
"XER yazimi UTF-16LE" karari MPXJ + kendi parser'imiz icindi; P6'nin kendi
importer'i icin gecerli DEGIL. Kanit, kodlamayi icerikten ayiran bir deney:
orijinal ANSI dosya sorunsuz girer, **ayni icerigin** UTF-16LE kopyasi
`The import file is invalid.` (cikis kodu 6) ile reddedilir. `p6_write` artik
`encoding` parametresi aliyor ve ne kaybedildigini sayiyor
(cp1251 → 0 karakter, cp1254 → **22.359 karakter '?' oluyor**).

**🔴 Kiril bir program CLI import'undan saglam cikamaz.** P6 ANSI XER'i
*makinenin* kod sayfasiyla okur (burada cp1254). Olcum, ayni kaynaktan iki
kopya:

| Yol | Kirilli gorev | `bukhtourcity437` |
|---|---|---|
| Canli proje (onarilmis) | 529 | `Гранит` |
| **CLI import** | **0** | `Agaieo` |
| **Veritabani ici revizyon kopyasi** | **529** | `Гранит` |

Bu yuzden **revizyon veritabani icinde kopyalanir**, XER'den gecirilmez:
`p6_baseline action='revision'` baseline kopyalama makinesini kullanir ama
kopyayi gercek bir proje olarak birakir (`project_flag='Y'`,
`orig_proj_id` bos, EPS'te gorunur). Sadakat: 950/950 ad+tarih birebir,
ortak `task_id` 0, `p6_compare` ile sifir fark.

### Testler

| Paket | Kapsam | Sonuc |
|---|---|---|
| `pytest tests/` (cevrimdisi) | 1013 test | **1013 passed, 279 skipped, 0 fail** |
| `tests/test_xer_encoding_detect.py` | XER kod sayfasi saptama, 13 test | gecti |
| `tests/test_p6_progress_rules.py` | P6 ilerleme semantigi, 33 test | gecti |
| `tests/test_p6_analysis_rules.py` | yuzde tabani / birim / WBS yolu, 33 test | gecti |
| `tests/live/test_p6_health_evm.py` | DCMA + EVM, ham SQL capraz kontrol | **35/35** |
| `tests/test_p6_compare_rules.py` | task_code eslesmesi, 19 test | gecti |
| `tests/test_p6_writer_rules.py` | ERMHDR + XER bicimlendirme, 25 test | gecti |
| `tests/live/test_p6_full_acceptance.py` | 8 aracin tamami + P6 motoru dogrulamasi | **233/233** |

Tam kabul testi veriyi degistirir ve sonunda baslangic durumuna geri alir
(ilerleme temizlenir, test baseline'i silinir, veri tarihi geri konur);
tekrar tekrar calistirilabilir.

### MCP stdio testi
`initialize → p6_mcp 1.26.0` ·
`tools/list → ['p6_query','p6_job','p6_health','p6_evm','p6_progress','p6_baseline','p6_compare','p6_write','p6_cli']` · `tools/call` ✅

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

### 4.1 Yedek nerede — ve neden `ls` ile göremezsiniz

🔴 **Yedek çalışma klasöründe DEĞİL.** İlk deneme oraya yazmak istedi ve
başarısız oldu: SQL Server servis hesabı `Downloads` altına yazamıyor
(`Cannot open backup device ... Operating system error 5`). Yedek bu yüzden
SQL Server'ın **kendi** Backup dizinine alındı:

```
C:\Program Files\Microsoft SQL Server\MSSQL17.P6EXPRESS\MSSQL\Backup\PMDB_faz3_oncesi.bak
PMDB · 2026-08-26 14:09:34 · 43,8 MB · Database (full)
```

Çalışma klasöründeki `backup_20260826\` dizini **boştur ve öyle olmalıdır** —
orada bir yedek yok, hiç olmadı.

⚠️ **Bu dizin normal kullanıcıyla listelenemez** (`ls` → *Permission denied*),
dosya sistemi taraması yedeği "yok" gibi gösterir. Varlığı **SQL Server'a
sordurarak** doğrulayın — dosyayı gerçekten açıp okur:

```sql
EXEC master.dbo.xp_fileexist
  'C:\Program Files\Microsoft SQL Server\MSSQL17.P6EXPRESS\MSSQL\Backup\PMDB_faz3_oncesi.bak';

RESTORE HEADERONLY FROM DISK =
  'C:\Program Files\Microsoft SQL Server\MSSQL17.P6EXPRESS\MSSQL\Backup\PMDB_faz3_oncesi.bak';

SELECT TOP 5 bs.backup_finish_date, bs.backup_size/1024/1024 AS mb,
       bmf.physical_device_name
  FROM msdb.dbo.backupset bs
  JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
 WHERE bs.database_name = 'PMDB'
 ORDER BY bs.backup_finish_date DESC;
```

Yeni bir yedek alırken de hedefi `SERVERPROPERTY('InstanceDefaultBackupPath')`
ile sorun; Express sürümde `WITH COMPRESSION` desteklenmez.

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

**Çözüm uygulandı** (yedek için bkz. §4.1 — çalışma klasöründe **değil**, SQL
Server'ın kendi Backup dizininde):

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

### ✅ 5.2 CLI import kaynak ücretlerini düşürüyordu — ÇÖZÜLDÜ (26.08, bkz. §3 Faz 5b)

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

## 7. Sıradaki işler (26.08 akşam güncellemesi)

1. ✅ **`mcp_common.py` diğer 3 sunucuya taşındı** (857fd79) — asta_mcp_core /
   asta_mcp_file yerel ham-kesicilerini `shrink_json_text`'e devretti,
   msproject_mcp_core'un 14 tool çıkışı (hiç guard'sızdı) `json_response`'a
   bağlandı, `dispatch`'in str dalı da düzeltildi. 28 birim testi
   (`tests/test_mcp_common.py`) + sunucu bağlantısını sabitleyen wiring
   testleri eklendi.
2. **`JT_XERExport` dosya adı parametresi** — P6 arayüzünde bir kez export işi
   oluşturup JOBSVC satırını okumak yeterli (§3, Faz 4 notu). `p6_write`
   ihtiyacı karşıladığı için acil değil. **KULLANICI EYLEMİ GEREKLİ (GUI).**
3. **VP_IMP_OPT import konfigürasyonu** — `repair_costs` sorunu çözüyor ama
   asıl temiz yol, P6 arayüzünde bir import şablonu kaydedip `view_data`
   kodlamasını oradan öğrenmek. **KULLANICI EYLEMİ GEREKLİ (GUI).**
4. ✅ **İş tipleri gerçek veriyle ölçüldü** (e5aba31) — bkz. §3 "Faz 5c".
5. `PrmJob.Job` COM `Execute`'u `comtypes` ile yeniden dene (pywin32 `VT_BYREF`
   OUT parametrelerini kabul etmiyor; servis kuyruğu çalıştığı için bloklayıcı değil).
6. ✅ **Faz 6 — `p6_task` (10. tool, 3bb3670)**: create_project / add_wbs /
   add_task / update_task / delete_task / add_link / delete_link /
   assign_resource / remove_assignment. Sıfırdan proje kurma + mevcut
   programı düzenleme hedefi KAPANDI — kabul testi N bölümü (21 kontrol):
   boş veritabanından proje kur → WBS → aktiviteler → bağlar → atama → F9
   → P6 CPM tarihleri doğru ve zincir sıralı → süre değişimi bitişi tam
   +10 iş günü kaydırdı → silmeler → korumalı silme, iz yok. Toplam 273/273.
   Ölçülen iki F9 tuzağı koda gömüldü:
   🔴 (a) **Atamada units/time 0 olursa F9 aktivitenin kalan VE planlanan
   süresini SIFIRLAR** (80h DT_FixedDrtn aktivite 0h döndü) —
   `assign_resource` artık RSRCRATE.max_qty_per_hr'den (yoksa 1/saat)
   dolduruyor. (b) **Süre değişince atama defteri taşınmalı** — yoksa F9
   kalan süreyi atamanın kalan biriminden geri yazar (160h yazıldı, bayat
   80h atama geri çekti); `update_task` defterleri p6_progress gibi taşıyor
   (`update_assignments=false` ile kapatılabilir, uyarı verir).
   Yapısal varsayılanlar projenin kendi geleneğinden (modal değer; boş
   projede PROJECT satırı). Yeni aktivite TARİHSİZ yazılır — ilk F9'a kadar
   P6 istemcisindeki gibi. `create_project` OBSPROJ'u `TR_PROJECT_OBSPROJ`
   trigger'ından alır (doğrulandı). `add_link` çift bağı ve döngü kapatacak
   bağı BFS ile baştan reddeder.

### Faz 5c — iş tipleri ölçüldü, iki tanesi P6 Pro'da YOK (26.08 akşam)

- ✅ **JT_Sum hiç "sessiz başarısız" değildi — yanlış tabloya bakılmıştı.**
  Özet `SUMTASK` / `SUMTASKSPREAD` / `SUMTRSRC` tablolarına,
  `PROJECT.wbs_max_sum_level` (=2 → 7 WBS düğümü) derinliğine yazılır;
  `TASKSUM`/`TRSRCSUM` bu derlemede hep boş kalır. Kanıt: yeniden koşumda kök
  SUMTASK 3/2/945 → 0/0/950 = canlı TASK birebir, `last_tasksum_date` ilerledi.
- 🔴 **JT_Level ve JT_UpdateBaseline P6 Professional kuyruğunda ÇALIŞMAZ.**
  Ölçüm: ikisi de `Invalid Job type` ile JS_Failed. İkili kanıt: prmjob.exe
  UTF-16LE string tablosunda "Invalid Job type:" hemen öncesindeki dispatch
  listesi yalnız şu yediyi taşıyor: `JT_Sched · JT_ApplyActuals · JT_XERExport
  · JT_Sum · JT_Enterprise_Sum · JT_Batch · JT_Report`. (JT_Level /
  JT_UpdateBaseline / JT_CreateBaseline sabitleri ikilide BAŞKA yerlerde var
  ama dispatcher koşmuyor — leveling ve baseline güncelleme GUI-only.)
  `jobs.submit` artık bu tipleri kuyruğa hiç bırakmadan açıklayıcı hata verir
  (`jobs.DISPATCHABLE`).
- **JT_ApplyActuals** P6 uygulama koduna ulaşıyor; auto-compute-actuals
  işaretli öğe yoksa `No projects to apply actual to.` ile anlamlı biçimde reddediyor
  (bukhtourcity'de 950/950 `auto_compute_act_flag='N'`). Pozitif yol için
  işaretli veri gerekir — tool'larla bu flag henüz yazılamıyor.
- ✅ **`compare_baselines_evm` iki gerçek baseline ile doğrulandı**
  (379 vs 380, sandbox): iki tarafta da 950/950 eşleşme, 0 eşleşmeyen,
  BAC birebir, delta 0 (kopyalar arasında yalnız ilerleme alanı farkı vardı).
- ✅ **`p6_baseline delete` korumalı gerçek-proje silme kazandı**: revizyon
  kopyası gerçek projedir (orig_proj_id boş) ve düz delete onu reddediyordu —
  revizyonlar silinemiyordu. Artık `delete_project=true` +
  `expected_short_name` (birebir ad) ister; projeye bakan canlı baseline
  varken reddeder. İki guard da canlı ölçüldü.
- Not: revizyon/baseline kopyasında `not_copied` listesi OBSPROJ'u sayar ama
  378 için OBSPROJ satırı vardı (muhtemelen şema trigger'ı üretiyor) ve
  JT_Sched kopya üzerinde sorunsuz koştu.
- **Kabul testi 233 → 252 kontrol** (L: revizyon sadakati + korumalı silme;
  M: p6_cli import+repair_costs — `P6_CLI_PASSWORD` yoksa temiz atlanır;
  G: summarize/level/update_baseline/apply_actuals ölçümleri). 252/252, 42 sn.
- Yedek: `PMDB_kalanisler_oncesi_20260826.bak` (SQL Backup dizini, 17:11, 87,8 MB).

## 8. Doğrulanmamış / riskli

- ✅(26.08 akşam) ~~JT_Level / JT_ApplyActuals / JT_UpdateBaseline çalıştırılmadı~~ —
  ölçüldü, bkz. §3 Faz 5c: JT_Level ve JT_UpdateBaseline **P6 Pro kuyruğunda YOK**
  (dispatch whitelist 7 tip), JT_ApplyActuals çalışıyor ama auto-compute işaretli
  öğe ister. Pozitif ApplyActuals yolu hâlâ denenmedi (flag yazma yolu yok).
- ✅(26.08 akşam) ~~JT_Sum sessiz başarısızlık~~ — yanlış tabloya bakılmıştı;
  özet SUMTASK/SUMTASKSPREAD/SUMTRSRC'de, kanıtla doğrulandı (§3 Faz 5c).
- `JT_XERExport` **çalıştırıldı**: kuyruktan alınıyor, P6'nın export koduna
  ulaşıyor, `File name not specified.` ile düşüyor (§3, Faz 4 notu).
- ✅(26.08 akşam) ~~compare_baselines_evm iki baseline ile denenmedi~~ —
  iki gerçek baseline ile doğrulandı (§3 Faz 5c).
- ✅(26.08 akşam) ~~p6_cli ve revision kabul testinde yok~~ — kabul testine
  L (revizyon, koşuldu) ve M (p6_cli, `P6_CLI_PASSWORD` yoksa atlanır) bölümleri
  eklendi; 252/252. M bölümü parola tanımlı bir oturumda henüz KOŞULMADI.
- **`p6_write`'ın ANSI çıktısı P6'ya geri import edilerek denenmedi.** UTF-16LE
  reddedildiği ölçüldü; `encoding='cp1251'` ile yazılan dosyanın P6 tarafından
  kabul edilip edilmediği sınanmadı.
- **Türkçe karakter taşıyan bir P6 programı denenmedi** — collation cp1251;
  Türkçe'ye özgü harfler bu veritabanında tutulamaz (§5.1 takası).
- **P6 arayüzünde göze bakılmadı.** Veri doğruluğu P6'nın kendi motoruyla
  kanıtlandı (§3, "P6'nin kendi motoruyla dogrulama") — yazılan kalan süreden
  P6'nın türettiği iş penceresi birebir tutuyor. Doğrulanmayan tek şey GUI
  *render*'ı: grid'de Kiril adların ve yüzdelerin görünümü.
  🔴 **computer-use bu iş için çalışmıyor:** Start menüsündeki
  "P6 Professional 24 (x64)" izni `primavera.cacheservice.exe`'ye çözülüyor,
  arayüz ise `PM.exe` — pencere maskeleniyor ve tıklama reddediliyor.
- `p6/db.py`'nin **SQLite backend'i** P6 24.12 SQLite şemasında test edildi;
  SQL Server yolu asıl kullanılan.
- `snapshot()` fallback dalı (3 dosya kopyası) hiç tetiklenmedi — `VACUUM INTO`
  her seferinde çalıştı.
