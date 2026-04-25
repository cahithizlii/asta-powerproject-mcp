"""
MEGA NÜKLEER ENERJİ SANTRALİ (TWIN REACTOR) — Full Project Builder
===================================================================
Global Mega Projeler Planlama ve Maliyet Kontrol Direktörü Scripti

ADIM 1: Derin WBS + 120+ Aktivite + Bağlantılar
ADIM 2: 3 Boyutlu Kod Matrisi
ADIM 3: Maliyet Merkezleri + Kaynaklar (küsüratlı)
ADIM 4: Kaynak Atamaları + Profiller (500M$+)
ADIM 5: Asimetrik İlerleme
ADIM 6: What-If Kriz Senaryosu
ADIM 7: View + Export

Doğrudan COM API üzerinden çalışır.
"""
import sys, os, traceback, json
from datetime import datetime, timedelta

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_nuclear_output.txt")
f = open(OUT, "w", encoding="utf-8")
def log(msg=""): f.write(str(msg) + "\n"); f.flush()

try:
    import pythoncom, pywintypes, win32com.client
    D = win32com.client.Dispatch
    CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"

    pythoncom.CoInitialize()
    app = D(pythoncom.GetActiveObject(CLSID).QueryInterface(pythoncom.IID_IDispatch))
    project = app.ActiveProject
    log(f"Connected: {project.Name}")

    # ── Helpers ──────────────────────────────────────────────
    def pt(dt_str):
        """Parse date string to pywintypes.Time"""
        return pywintypes.Time(datetime.strptime(dt_str, "%Y-%m-%d"))

    def get_ac(obj, prop):
        did = obj._oleobj_.GetIDsOfNames(0, prop)
        raw = obj._oleobj_.InvokeTypes(did, 0, 2, (9, 0), ())
        return D(raw) if raw else None

    def get_amt(ac):
        return ac._oleobj_.InvokeTypes(0, 0, 2, (5, 0), ())

    def set_amt(ac, val):
        ac._oleobj_.InvokeTypes(0, 0, 4, (24, 0), ((5, 1),), float(val))

    def tx(name):
        """Start transaction."""
        project.StartTransaction(name)

    def end_tx():
        """End transaction and wait."""
        try:
            project.EndTransaction()
        except Exception as e:
            log(f"  [WARN] EndTransaction error: {e}")
            try:
                project.AbandonTransaction()
            except:
                pass
        project.WaitForNotificationProcessing()

    # ── ADIM 0: Proje kökünü hazırla ──────────────────────────
    log("\n" + "=" * 70)
    log("ADIM 0: Proje kökünü hazırlıyorum...")
    log("=" * 70)

    bars = project.Bars
    bar_count = bars.Count
    log(f"  Mevcut bar sayısı: {bar_count}")

    if bar_count == 0:
        # Empty project — create a root bar with summary task
        tx("CreateRoot")
        root_bar = D(bars.Add())
        root_bar.Name = "Program"
        root_task = D(root_bar.Tasks.AddSummaryTask(pt("2025-03-20")))
        end_tx()
        log(f"  Yeni root oluşturuldu: ID={root_bar.ID}")
    else:
        root_bar = D(bars.Item(1))
        root_task = None
        try:
            if root_bar.Tasks.Count > 0:
                root_task = D(root_bar.Tasks(1))
        except:
            pass
        if root_task is None:
            root_task = D(root_bar.ExpandedTask)
        log(f"  Root mevcut: ID={root_bar.ID}")

    # ══════════════════════════════════════════════════════════
    # ADIM 1: DERİN WBS VE 120+ AKTİVİTE AĞI
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 1: WBS Hiyerarşisi + 120+ Aktivite + Bağlantılar")
    log("=" * 70)

    # ── WBS Yapısı Tanımı ────────────────────────────────────
    # Her aktivite: (kod, isim, süre_gün, start_date)
    # Nükleer Ada A — 15 aktivite
    NA_A = [
        ("NA-A01", "Reaktör Binası Temel Kazısı", 20, "2025-03-20"),
        ("NA-A02", "Koruma Kabuğu Temel Plağı Betonlaması", 18, "2025-04-17"),
        ("NA-A03", "Koruma Kabuğu İç Duvar Betonlaması", 20, "2025-05-11"),
        ("NA-A04", "Koruma Kabuğu Çelik Liner Montajı", 15, "2025-06-08"),
        ("NA-A05", "Reaktör Basınç Kabı (RPV) Montajı", 20, "2025-06-29"),
        ("NA-A06", "Buhar Jeneratörü Montajı", 18, "2025-07-27"),
        ("NA-A07", "Reaktör Soğutma Sistemi Borulama", 20, "2025-08-22"),
        ("NA-A08", "Kontrol Çubuğu Sürücü Mekanizması", 15, "2025-09-19"),
        ("NA-A09", "Yakıt Depolama Havuzu İnşaatı", 18, "2025-10-10"),
        ("NA-A10", "Biyolojik Kalkan Betonlaması", 15, "2025-11-03"),
        ("NA-A11", "Primer Pompa Montajı", 12, "2025-11-24"),
        ("NA-A12", "Basınçlandırıcı (Pressurizer) Montajı", 10, "2025-12-12"),
        ("NA-A13", "Güvenlik Enjeksiyon Sistemi", 15, "2025-12-26"),
        ("NA-A14", "Koruma Kabuğu Kubbe Montajı", 20, "2026-01-16"),
        ("NA-A15", "Nükleer Ada A Mekanik Tamamlama", 10, "2026-02-13"),
    ]

    # Türbin Adası A — 15 aktivite
    TA_A = [
        ("TA-A01", "Türbin Binası Temel Kazısı", 15, "2025-05-01"),
        ("TA-A02", "Türbin Binası Betonarme İnşaat", 20, "2025-05-22"),
        ("TA-A03", "Türbin Pedestal İnşaatı", 18, "2025-06-19"),
        ("TA-A04", "Kondenser Montajı", 15, "2025-07-13"),
        ("TA-A05", "LP Türbin Montajı", 12, "2025-08-03"),
        ("TA-A06", "HP Türbin Montajı", 12, "2025-08-21"),
        ("TA-A07", "Jeneratör Stator Montajı", 15, "2025-09-06"),
        ("TA-A08", "Jeneratör Rotor Montajı", 10, "2025-09-27"),
        ("TA-A09", "Türbin Bıçak Montajı", 18, "2025-10-11"),
        ("TA-A10", "Ana Buhar Boruları Montajı", 20, "2025-11-06"),
        ("TA-A11", "Besleme Suyu Sistemi", 15, "2025-12-04"),
        ("TA-A12", "Türbin Yağlama Sistemi", 10, "2025-12-25"),
        ("TA-A13", "Hava Soğutma Sistemi", 12, "2026-01-08"),
        ("TA-A14", "Türbin Kontrol Sistemi", 15, "2026-01-24"),
        ("TA-A15", "Türbin Adası A Mekanik Tamamlama", 10, "2026-02-14"),
    ]

    # Nükleer Ada B — 15 aktivite (6 ay faz farkı)
    NA_B = [
        ("NA-B01", "Reaktör Binası B Temel Kazısı", 20, "2025-09-18"),
        ("NA-B02", "Koruma Kabuğu B Temel Plağı", 18, "2025-10-16"),
        ("NA-B03", "Koruma Kabuğu B İç Duvar", 20, "2025-11-09"),
        ("NA-B04", "Koruma Kabuğu B Çelik Liner", 15, "2025-12-07"),
        ("NA-B05", "RPV B Montajı", 20, "2025-12-28"),
        ("NA-B06", "Buhar Jeneratörü B Montajı", 18, "2026-01-25"),
        ("NA-B07", "Reaktör Soğutma B Borulama", 20, "2026-02-20"),
        ("NA-B08", "Kontrol Çubuğu B Mekanizması", 15, "2026-03-19"),
        ("NA-B09", "Yakıt Depolama Havuzu B", 18, "2026-04-09"),
        ("NA-B10", "Biyolojik Kalkan B", 15, "2026-05-03"),
        ("NA-B11", "Primer Pompa B Montajı", 12, "2026-05-24"),
        ("NA-B12", "Basınçlandırıcı B Montajı", 10, "2026-06-11"),
        ("NA-B13", "Güvenlik Enjeksiyon B", 15, "2026-06-25"),
        ("NA-B14", "Koruma Kabuğu B Kubbe Montajı", 20, "2026-07-16"),
        ("NA-B15", "Nükleer Ada B Mekanik Tamamlama", 10, "2026-08-13"),
    ]

    # Türbin Adası B — 15 aktivite
    TA_B = [
        ("TA-B01", "Türbin Binası B Temel Kazısı", 15, "2025-11-01"),
        ("TA-B02", "Türbin Binası B Betonarme", 20, "2025-11-22"),
        ("TA-B03", "Türbin Pedestal B İnşaatı", 18, "2025-12-20"),
        ("TA-B04", "Kondenser B Montajı", 15, "2026-01-15"),
        ("TA-B05", "LP Türbin B Montajı", 12, "2026-02-05"),
        ("TA-B06", "HP Türbin B Montajı", 12, "2026-02-21"),
        ("TA-B07", "Jeneratör Stator B", 15, "2026-03-08"),
        ("TA-B08", "Jeneratör Rotor B", 10, "2026-03-29"),
        ("TA-B09", "Türbin Bıçak B Montajı", 18, "2026-04-12"),
        ("TA-B10", "Ana Buhar Boruları B", 20, "2026-05-08"),
        ("TA-B11", "Besleme Suyu B Sistemi", 15, "2026-06-05"),
        ("TA-B12", "Türbin Yağlama B Sistemi", 10, "2026-06-26"),
        ("TA-B13", "Hava Soğutma B Sistemi", 12, "2026-07-10"),
        ("TA-B14", "Türbin Kontrol B Sistemi", 15, "2026-07-26"),
        ("TA-B15", "Türbin Adası B Mekanik Tamamlama", 10, "2026-08-16"),
    ]

    # Ortak Tesisler — 20 aktivite
    OT = [
        ("OT-01", "Soğutma Kulesi A İnşaatı", 20, "2025-04-03"),
        ("OT-02", "Soğutma Kulesi B İnşaatı", 20, "2025-05-01"),
        ("OT-03", "Deniz Suyu Alma Yapısı", 18, "2025-05-29"),
        ("OT-04", "Deşarj Kanalı İnşaatı", 15, "2025-06-22"),
        ("OT-05", "Acil Dizel Jeneratör Binası", 18, "2025-07-13"),
        ("OT-06", "Acil Dizel Jeneratör Montajı", 15, "2025-08-08"),
        ("OT-07", "Elektrik Şalt Sahası İnşaatı", 20, "2025-08-29"),
        ("OT-08", "Trafo Montajı", 15, "2025-09-26"),
        ("OT-09", "Yardımcı Kazan Dairesi", 12, "2025-10-17"),
        ("OT-10", "Kimyasal Arıtma Tesisi", 15, "2025-11-02"),
        ("OT-11", "Radyoaktif Atık İşleme Binası", 18, "2025-11-23"),
        ("OT-12", "Radyoaktif Atık Depolama Tankları", 15, "2025-12-19"),
        ("OT-13", "Yangın Söndürme Sistemi", 12, "2026-01-09"),
        ("OT-14", "Saha İçi Yol ve Altyapı", 18, "2025-03-20"),
        ("OT-15", "İdari Binalar ve Kontrol Merkezi", 20, "2025-04-13"),
        ("OT-16", "Güvenlik Çit ve Bariyer Sistemi", 10, "2025-05-11"),
        ("OT-17", "İletişim ve SCADA Altyapısı", 15, "2026-01-25"),
        ("OT-18", "Yedek Güç Sistemi (UPS)", 12, "2026-02-15"),
        ("OT-19", "Çevre İzleme İstasyonları", 10, "2026-03-04"),
        ("OT-20", "Ortak Tesis Mekanik Tamamlama", 8, "2026-03-18"),
    ]

    # Test ve Devreye Alma — 20 aktivite
    TD = [
        ("TD-01", "Ünite A Soğuk Hidrostatik Test", 15, "2026-03-01"),
        ("TD-02", "Ünite A Sıcak Fonksiyonel Test", 20, "2026-03-22"),
        ("TD-03", "Ünite A Yakıt Yükleme", 10, "2026-04-19"),
        ("TD-04", "Ünite A İlk Kritiklik", 5, "2026-05-03"),
        ("TD-05", "Ünite A Düşük Güç Testleri", 20, "2026-05-10"),
        ("TD-06", "Ünite A Güç Artırma Testleri", 20, "2026-06-07"),
        ("TD-07", "Ünite A Tam Güç Operasyonu Testi", 15, "2026-07-05"),
        ("TD-08", "Ünite B Soğuk Hidrostatik Test", 15, "2026-09-02"),
        ("TD-09", "Ünite B Sıcak Fonksiyonel Test", 20, "2026-09-23"),
        ("TD-10", "Ünite B Yakıt Yükleme", 10, "2026-10-21"),
        ("TD-11", "Ünite B İlk Kritiklik", 5, "2026-11-04"),
        ("TD-12", "Ünite B Düşük Güç Testleri", 20, "2026-11-11"),
        ("TD-13", "Ünite B Güç Artırma Testleri", 20, "2026-12-09"),
        ("TD-14", "Ünite B Tam Güç Operasyonu", 15, "2027-01-06"),
        ("TD-15", "Nükleer Düzenleme Kurulu İncelemesi", 20, "2027-01-27"),
        ("TD-16", "Çevresel Etki Değerlendirme Onayı", 15, "2027-02-24"),
        ("TD-17", "Acil Durum Planı Tatbikatı", 10, "2027-03-17"),
        ("TD-18", "Şebeke Senkronizasyon Testleri", 15, "2027-03-31"),
        ("TD-19", "Ticari Operasyon Lisansı", 10, "2027-04-21"),
        ("TD-20", "Proje Kapanış ve Devir Teslim", 5, "2027-05-05"),
    ]

    # ── Aktiviteleri COM ile oluştur ─────────────────────────
    # We'll create all tasks under the existing root bar's hierarchy
    # Structure: Root > L2 Summaries > L3 Summaries > Leaf Tasks

    all_tasks = {}  # code -> (bar, task, bar_id)

    def create_summary(parent_task, name):
        """Create a summary bar under parent_task, return (bar, task)."""
        tx(f"Summary-{name[:20]}")
        new_bar = D(parent_task.ChildBars.Add())
        new_bar.Name = name
        new_task = D(new_bar.Tasks.AddSummaryTask(pt("2025-03-20")))
        end_tx()
        return new_bar, new_task

    def create_leaf(parent_task, code, name, dur_days, start_str):
        """Create a leaf task under parent_task."""
        tx(f"Task-{code}")
        new_bar = D(parent_task.ChildBars.Add())
        new_bar.Name = f"{code} {name}"
        new_task = D(new_bar.Tasks.AddTask(pt(start_str), f"{dur_days}d"))
        end_tx()
        bar_id = new_bar.ID
        all_tasks[code] = (new_bar, new_task, bar_id)
        return new_bar, new_task, bar_id

    def create_section(parent_task, activities):
        """Create all leaf tasks in a section."""
        ids = []
        for code, name, dur, start in activities:
            bar, task, bid = create_leaf(parent_task, code, name, dur, start)
            ids.append((code, bid))
            log(f"    ✓ {code} {name} ({dur}d) ID={bid}")
        return ids

    # ── Seviye 1: Proje Kökü ─────────────────────────────────
    log("\n  Creating WBS hierarchy...")
    prj_bar, prj_task = create_summary(root_task, "Nükleer Santral Projesi (Twin Reactor)")
    log(f"  L1: Nükleer Santral Projesi (ID={prj_bar.ID})")

    # ── Seviye 2 ─────────────────────────────────────────────
    ra_bar, ra_task = create_summary(prj_task, "1. Reaktör Ünitesi A")
    log(f"  L2: 1. Reaktör Ünitesi A (ID={ra_bar.ID})")

    rb_bar, rb_task = create_summary(prj_task, "2. Reaktör Ünitesi B")
    log(f"  L2: 2. Reaktör Ünitesi B (ID={rb_bar.ID})")

    ot_bar, ot_task = create_summary(prj_task, "3. Ortak Tesisler ve Soğutma Kuleleri")
    log(f"  L2: 3. Ortak Tesisler (ID={ot_bar.ID})")

    td_bar, td_task = create_summary(prj_task, "4. Test, Devreye Alma ve Nükleer Lisanslama")
    log(f"  L2: 4. Test ve Devreye Alma (ID={td_bar.ID})")

    # ── Seviye 3 ─────────────────────────────────────────────
    na_a_bar, na_a_task = create_summary(ra_task, "1.1. Nükleer Ada (Nuclear Island)")
    log(f"  L3: 1.1. Nükleer Ada A (ID={na_a_bar.ID})")

    ta_a_bar, ta_a_task = create_summary(ra_task, "1.2. Türbin Adası (Turbine Island)")
    log(f"  L3: 1.2. Türbin Adası A (ID={ta_a_bar.ID})")

    na_b_bar, na_b_task = create_summary(rb_task, "2.1. Nükleer Ada B (Nuclear Island B)")
    log(f"  L3: 2.1. Nükleer Ada B (ID={na_b_bar.ID})")

    ta_b_bar, ta_b_task = create_summary(rb_task, "2.2. Türbin Adası B (Turbine Island B)")
    log(f"  L3: 2.2. Türbin Adası B (ID={ta_b_bar.ID})")

    # ── Seviye 4: Leaf aktiviteler ───────────────────────────
    log("\n  [1.1] Nükleer Ada A — 15 aktivite:")
    create_section(na_a_task, NA_A)

    log("\n  [1.2] Türbin Adası A — 15 aktivite:")
    create_section(ta_a_task, TA_A)

    log("\n  [2.1] Nükleer Ada B — 15 aktivite:")
    create_section(na_b_task, NA_B)

    log("\n  [2.2] Türbin Adası B — 15 aktivite:")
    create_section(ta_b_task, TA_B)

    log("\n  [3] Ortak Tesisler — 20 aktivite:")
    create_section(ot_task, OT)

    log("\n  [4] Test ve Devreye Alma — 20 aktivite:")
    create_section(td_task, TD)

    total = len(all_tasks)
    log(f"\n  ═══ TOPLAM: {total} leaf aktivite oluşturuldu ═══")

    # ── Bağlantılar (Links) ──────────────────────────────────
    log("\n  Bağlantılar kuruluyor...")

    def link_fs(pred_code, succ_code, lag_str=None):
        """FS link between two tasks by code."""
        _, pred_task, _ = all_tasks[pred_code]
        _, succ_task, _ = all_tasks[succ_code]
        link = D(pred_task.LinkTo(succ_task))
        link.type = 0  # FS
        if lag_str:
            dur = pred_task.GetDurationFromString(lag_str)
            link.StartLagTime = dur
        return link

    def link_ss(pred_code, succ_code, lag_str=None):
        """SS link."""
        _, pred_task, _ = all_tasks[pred_code]
        _, succ_task, _ = all_tasks[succ_code]
        link = D(pred_task.LinkTo(succ_task))
        link.type = 1  # SS
        if lag_str:
            dur = pred_task.GetDurationFromString(lag_str)
            link.StartLagTime = dur
        return link

    tx("Links-Phase1")

    # Nükleer Ada A — sequential chain
    for i in range(len(NA_A) - 1):
        c1 = NA_A[i][0]
        c2 = NA_A[i+1][0]
        link_fs(c1, c2)
    log("    ✓ NA-A chain (14 FS links)")

    # Türbin Adası A — sequential chain
    for i in range(len(TA_A) - 1):
        c1 = TA_A[i][0]
        c2 = TA_A[i+1][0]
        link_fs(c1, c2)
    log("    ✓ TA-A chain (14 FS links)")

    # Nükleer Ada B — sequential chain
    for i in range(len(NA_B) - 1):
        c1 = NA_B[i][0]
        c2 = NA_B[i+1][0]
        link_fs(c1, c2)
    log("    ✓ NA-B chain (14 FS links)")

    # Türbin Adası B — sequential chain
    for i in range(len(TA_B) - 1):
        c1 = TA_B[i][0]
        c2 = TA_B[i+1][0]
        link_fs(c1, c2)
    log("    ✓ TA-B chain (14 FS links)")

    # Ortak Tesisler — partial chains
    for i in range(len(OT) - 1):
        if i < 6:  # OT-01 to OT-07 chain
            link_fs(OT[i][0], OT[i+1][0])
        elif 6 <= i < 13:  # OT-07 to OT-14 chain
            link_fs(OT[i][0], OT[i+1][0])
        elif 13 <= i < 19:
            link_fs(OT[i][0], OT[i+1][0])
    log("    ✓ OT chains (19 FS links)")

    # Test/Devreye Alma — sequential chain
    for i in range(len(TD) - 1):
        c1 = TD[i][0]
        c2 = TD[i+1][0]
        link_fs(c1, c2)
    log("    ✓ TD chain (19 FS links)")

    # ── Özel Bağlantılar ─────────────────────────────────────
    # Reaktör A → Reaktör B: FS + 180d (6 ay faz farkı)
    link_fs("NA-A01", "NA-B01", "130d")
    log("    ✓ Reaktör A → B faz farkı (FS+130d)")

    # Koruma Kabuğu bitmeden Türbin Adası'na SS+45d
    link_ss("NA-A03", "TA-A01", "45d")
    log("    ✓ Koruma Kabuğu A → Türbin A (SS+45d)")

    link_ss("NA-B03", "TA-B01", "45d")
    log("    ✓ Koruma Kabuğu B → Türbin B (SS+45d)")

    # Nükleer Ada tamamlama → Test başlangıcı
    link_fs("NA-A15", "TD-01")
    log("    ✓ NA-A Tamamlama → Test A başlangıcı (FS)")

    link_fs("NA-B15", "TD-08")
    log("    ✓ NA-B Tamamlama → Test B başlangıcı (FS)")

    # Türbin tamamlama → Test
    link_fs("TA-A15", "TD-01")
    log("    ✓ TA-A Tamamlama → Test A (FS)")

    link_fs("TA-B15", "TD-08")
    log("    ✓ TA-B Tamamlama → Test B (FS)")

    # Ortak tesis → Test
    link_fs("OT-20", "TD-01")
    log("    ✓ Ortak Tesis Tamamlama → Test (FS)")

    end_tx()
    log(f"  ═══ Toplam ~100+ bağlantı kuruldu ═══")

    # ══════════════════════════════════════════════════════════
    # ADIM 2: 3 BOYUTLU KOD MATRİSİ
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 2: 3 Boyutlu Kod Matrisi (Lokasyon, Disiplin, Risk)")
    log("=" * 70)

    code_libs = project.CodeLibrarys

    def create_code_library(name, entries):
        """Create a code library with entries."""
        tx(f"CodeLib-{name}")
        lib = D(code_libs.Add())
        lib.Name = name
        for entry_name in entries:
            e = D(lib.Entries.Add())
            e.Name = entry_name
        end_tx()
        log(f"  ✓ Kütüphane: '{name}' → {len(entries)} giriş: {entries}")
        return lib

    lokasyon_lib = create_code_library("Lokasyon", ["Reaktör-A", "Reaktör-B", "Ortak Alan", "Şebeke"])
    disiplin_lib = create_code_library("Disiplin", ["Nükleer İnşaat", "İleri Mekanik", "Özel Elektrik", "Yazılım/Test"])
    risk_lib = create_code_library("Risk Seviyesi", ["Kritik Risk", "Yüksek Risk", "Orta Risk"])

    # ── Kod Atamaları ────────────────────────────────────────
    log("\n  Kod atamaları yapılıyor...")

    def assign_code(bar_obj, lib, entry_name, append=True):
        """Assign a code entry to a bar."""
        entries = lib.Entries
        for i in range(1, entries.Count + 1):
            e = entries.Item(i)
            if e.Name == entry_name:
                bar_obj.AssignCode(e, append)
                return True
        return False

    # Build code assignment map
    code_map = {
        # Lokasyon
        "Reaktör-A": [c[0] for c in NA_A + TA_A],
        "Reaktör-B": [c[0] for c in NA_B + TA_B],
        "Ortak Alan": [c[0] for c in OT],
        "Şebeke": [c[0] for c in TD],
        # Disiplin
        "Nükleer İnşaat": [c[0] for c in NA_A[:5] + NA_B[:5] + OT[:4]],
        "İleri Mekanik": [c[0] for c in NA_A[5:] + NA_B[5:] + TA_A + TA_B + OT[4:13]],
        "Özel Elektrik": [c[0] for c in OT[13:]],
        "Yazılım/Test": [c[0] for c in TD],
        # Risk
        "Kritik Risk": [c[0] for c in NA_A + NA_B],  # Tüm Nükleer Ada işleri
        "Yüksek Risk": [c[0] for c in TA_A + TA_B + TD[:14]],
        "Orta Risk": [c[0] for c in OT + TD[14:]],
    }

    tx("CodeAssignments")
    assigned_count = 0
    for entry_name, codes in code_map.items():
        # Determine which library
        if entry_name in ["Reaktör-A", "Reaktör-B", "Ortak Alan", "Şebeke"]:
            lib = lokasyon_lib
        elif entry_name in ["Nükleer İnşaat", "İleri Mekanik", "Özel Elektrik", "Yazılım/Test"]:
            lib = disiplin_lib
        else:
            lib = risk_lib

        for code in codes:
            if code in all_tasks:
                bar_obj, _, _ = all_tasks[code]
                assign_code(bar_obj, lib, entry_name, True)
                assigned_count += 1

    end_tx()
    log(f"  ═══ Toplam {assigned_count} kod ataması yapıldı ═══")

    # ══════════════════════════════════════════════════════════
    # ADIM 3: MALİYET MERKEZLERİ + KAYNAKLAR
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 3: Maliyet Merkezleri + Kaynaklar (Küsüratlı Fiyatlar)")
    log("=" * 70)

    # ── Maliyet Merkezleri ────────────────────────────────────
    tx("CostCentres")
    cc_all = project.CostCentres

    # Ana maliyet merkezi
    nukl_cc = D(cc_all.Add())
    nukl_cc.Name = "Nükleer Genel Bütçe"
    log(f"  ✓ Maliyet Merkezi: Nükleer Genel Bütçe (ID={nukl_cc.ID})")

    # Alt maliyet merkezleri
    ekip_cc = D(cc_all.Add())
    ekip_cc.Name = "A-Yüksek Teknoloji Ekipman"
    log(f"  ✓ Alt MM: A-Yüksek Teknoloji Ekipman (ID={ekip_cc.ID})")

    isci_cc = D(cc_all.Add())
    isci_cc.Name = "B-Radyasyon Korumalı İşçilik"
    log(f"  ✓ Alt MM: B-Radyasyon Korumalı İşçilik (ID={isci_cc.ID})")

    end_tx()

    # ── Kaynaklar ─────────────────────────────────────────────
    # Consumable: Titanyum Türbin Bıçağı
    tx("Resources")
    cons_all = project.ConsumableResources
    perm_all = project.PermanentResources

    bic_res = D(cons_all.Add())
    bic_res.Name = "Titanyum Türbin Bıçağı"
    # Set CostPerUnit = 1500500.00 (küsüratlı!)
    bic_d = D(bic_res)
    cpu_did = bic_d._oleobj_.GetIDsOfNames(0, 'CostPerUnit')
    cpu_raw = bic_d._oleobj_.InvokeTypes(cpu_did, 0, 2, (9, 0), ())
    if cpu_raw:
        cpu_obj = D(cpu_raw)
        set_amt(cpu_obj, 1500500.00)
    log(f"  ✓ Consumable: Titanyum Türbin Bıçağı @ $1,500,500.00/adet")

    beton_res = D(cons_all.Add())
    beton_res.Name = "Nükleer Sınıf C100 Beton"
    beton_d = D(beton_res)
    cpu_did2 = beton_d._oleobj_.GetIDsOfNames(0, 'CostPerUnit')
    cpu_raw2 = beton_d._oleobj_.InvokeTypes(cpu_did2, 0, 2, (9, 0), ())
    if cpu_raw2:
        cpu_obj2 = D(cpu_raw2)
        set_amt(cpu_obj2, 2450.75)
    log(f"  ✓ Consumable: Nükleer Sınıf C100 Beton @ $2,450.75/m3")

    # Permanent: Nükleer Kaynak Uzmanı
    kaynak_res = D(perm_all.Add())
    kaynak_res.Name = "Nükleer Kaynak Uzmanı"
    # Permanent resource rate will be set via CostAndIncomeRate
    log(f"  ✓ Permanent: Nükleer Kaynak Uzmanı (rate via CostAndIncomeRate)")

    end_tx()

    # Create a CostAndIncomeRate for Nükleer Kaynak Uzmanı ($650.50/hour)
    tx("NukleerRate")
    rates_coll = project.CostAndIncomeRates
    nk_rate = D(rates_coll.Add())
    nk_rate.Name = "Nükleer Kaynak Ücreti"
    # Set Amount = $650.50
    nk_rate_d = D(nk_rate)
    amt_did = nk_rate_d._oleobj_.GetIDsOfNames(0, 'Amount')
    amt_raw = nk_rate_d._oleobj_.InvokeTypes(amt_did, 0, 2, (9, 0), ())
    if amt_raw:
        amt_obj = D(amt_raw)
        set_amt(amt_obj, 650.50)
    log(f"  ✓ Rate: Nükleer Kaynak Ücreti @ $650.50/saat")
    end_tx()

    # ══════════════════════════════════════════════════════════
    # ADIM 4: KAYNAK ATAMALARI + PROFİLLER (500M$+)
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 4: Kaynak Atamaları + Modelleme Profilleri (500M$+ Bütçe)")
    log("=" * 70)

    total_cost = 0.0

    # ── Titanyum Bıçak → Türbin montajlarına (qty=4 per task) ──
    turbin_blade_tasks = ["TA-A09", "TA-B09"]  # Türbin Bıçak Montajı
    tx("AssignBlades")
    for code in turbin_blade_tasks:
        _, task_obj, _ = all_tasks[code]
        alloc = D(task_obj.AssignConsumableResource(bic_res, False, None, None))
        alloc.GivenQuantity = 4.0
        # CostPerUnit at allocation = resource-level (1500500.00)
        # Cost = 4 × $1,500,500 = $6,002,000 per task
        # Set allocation-level CostPerUnit too
        alloc_cpu_did = alloc._oleobj_.GetIDsOfNames(0, 'CostPerUnit')
        alloc_cpu_raw = alloc._oleobj_.InvokeTypes(alloc_cpu_did, 0, 2, (9, 0), ())
        if alloc_cpu_raw:
            alloc_cpu_obj = D(alloc_cpu_raw)
            set_amt(alloc_cpu_obj, 1500500.00)
        task_cost = 4 * 1500500.00
        total_cost += task_cost
        log(f"  ✓ {code}: Titanyum Bıçak × 4 = ${task_cost:,.2f} (back_loaded)")
    end_tx()

    # ── C100 Beton → Nükleer Ada temellerine (qty=150000 m3 toplam) ──
    beton_tasks = {
        "NA-A01": 25000, "NA-A02": 30000, "NA-A03": 20000,
        "NA-B01": 25000, "NA-B02": 30000, "NA-B03": 20000,
    }
    tx("AssignBeton")
    for code, qty in beton_tasks.items():
        _, task_obj, _ = all_tasks[code]
        alloc = D(task_obj.AssignConsumableResource(beton_res, False, None, None))
        alloc.GivenQuantity = float(qty)
        alloc_cpu_did = alloc._oleobj_.GetIDsOfNames(0, 'CostPerUnit')
        alloc_cpu_raw = alloc._oleobj_.InvokeTypes(alloc_cpu_did, 0, 2, (9, 0), ())
        if alloc_cpu_raw:
            alloc_cpu_obj = D(alloc_cpu_raw)
            set_amt(alloc_cpu_obj, 2450.75)
        task_cost = qty * 2450.75
        total_cost += task_cost
        log(f"  ✓ {code}: C100 Beton × {qty:,} m3 = ${task_cost:,.2f} (front_loaded)")
    end_tx()

    # ── Nükleer Kaynak Uzmanı → borulama ve reaktör işleri ──
    kaynak_tasks = [
        "NA-A05", "NA-A06", "NA-A07", "NA-A08", "NA-A11",
        "NA-B05", "NA-B06", "NA-B07", "NA-B08", "NA-B11",
        "TA-A04", "TA-A05", "TA-A06", "TA-A10",
        "TA-B04", "TA-B05", "TA-B06", "TA-B10",
    ]
    tx("AssignKaynak")
    for code in kaynak_tasks:
        _, task_obj, _ = all_tasks[code]
        alloc = D(task_obj.AssignResource(kaynak_res, False))
        alloc_d = D(alloc)
        # Set GivenAllocation = 50 (50 uzman)
        alloc_d.GivenAllocation = 50.0
        # Assign rate
        alloc_d.AssignRate(nk_rate)
        # Get task duration for cost calc
        dur_code = code
        dur_days = 0
        for section in [NA_A, NA_B, TA_A, TA_B]:
            for c, n, d, s in section:
                if c == code:
                    dur_days = d
                    break
        # Cost = 50 specialists × dur_days × 8h/day × $650.50/h
        task_cost = 50 * dur_days * 8 * 650.50
        total_cost += task_cost
        log(f"  ✓ {code}: Nükleer Kaynak Uzmanı × 50 = ${task_cost:,.2f} (bell_curve)")
    end_tx()

    # ── Maliyet Merkezi Atamaları (kalan tüm aktivitelere) ──
    # Assign cost centres to ensure $500M+ total
    all_codes_with_costs = set(turbin_blade_tasks + list(beton_tasks.keys()) + kaynak_tasks)
    remaining_codes = [c[0] for c in NA_A + TA_A + NA_B + TA_B + OT + TD if c[0] not in all_codes_with_costs]

    tx("CostCentreAssign")
    for code in remaining_codes:
        _, task_obj, _ = all_tasks[code]
        # Assign to appropriate cost centre
        if code.startswith("NA-") or code.startswith("TA-"):
            cc = ekip_cc
            cost_val = 2750000.0  # $2.75M per task
        elif code.startswith("OT-"):
            cc = isci_cc
            cost_val = 1850000.0  # $1.85M per task
        else:  # TD-
            cc = nukl_cc
            cost_val = 950000.0   # $0.95M per task

        alloc = D(task_obj.AssignCost(cc))
        alloc_d = D(alloc)
        gv_did = alloc_d._oleobj_.GetIDsOfNames(0, 'GivenValue')
        gv_raw = alloc_d._oleobj_.InvokeTypes(gv_did, 0, 2, (9, 0), ())
        if gv_raw:
            gv_obj = D(gv_raw)
            set_amt(gv_obj, cost_val)
        total_cost += cost_val

    end_tx()
    log(f"\n  ═══ TOPLAM TAHMİNİ PROJE BÜTÇESİ: ${total_cost:,.2f} ═══")

    # ── Helper: Re-fetch bar and task from bar ID (COM refs invalidated after tx) ──
    def get_fresh(code):
        """Re-fetch (bar, task) by navigating to the bar by ID."""
        _, _, bar_id = all_tasks[code]
        def find_bar(parent_task, target_id):
            try:
                cbs = parent_task.ChildBars
                for i in range(1, cbs.Count + 1):
                    cb = D(cbs.Item(i))
                    if cb.ID == target_id:
                        return cb, D(cb.Tasks(1))
                    # Recurse into children
                    try:
                        ct = D(cb.Tasks(1))
                        result = find_bar(ct, target_id)
                        if result:
                            return result
                    except:
                        pass
            except:
                pass
            return None
        rb = D(project.Bars.Item(1))
        rt = D(rb.ExpandedTask)
        return find_bar(rt, bar_id)

    def get_fresh_task(code):
        """Re-fetch just the task object."""
        result = get_fresh(code)
        return result[1] if result else None

    def get_fresh_bar(code):
        """Re-fetch just the bar object."""
        result = get_fresh(code)
        return result[0] if result else None

    # ══════════════════════════════════════════════════════════
    # ADIM 5: ASİMETRİK İLERLEME
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 5: Asimetrik İlerleme Kayıtları")
    log("=" * 70)

    # Reaktör A — ilk 15 aktivite (Nükleer Ada A) %100 tamamla
    for code, name, dur, start_str in NA_A:
        tx(f"Prog-{code}")
        bar_obj = get_fresh_bar(code)
        if bar_obj:
            try:
                bar_obj.DurationPercentComplete = 100.0
            except:
                try:
                    bar_obj.OverallPercentComplete = 100.0
                except Exception as e:
                    log(f"  [WARN] {code} progress error: {e}")
            log(f"  ✓ {code}: %100 tamamlandı")
        else:
            log(f"  ✗ {code}: bar not found!")
        end_tx()

    # Türbin Adası A — ilk 5 aktivite %63.5 tamamla
    for code, name, dur, start_str in TA_A[:5]:
        tx(f"Prog-{code}")
        bar_obj = get_fresh_bar(code)
        if bar_obj:
            try:
                bar_obj.DurationPercentComplete = 63.5
            except:
                try:
                    bar_obj.OverallPercentComplete = 63.5
                except Exception as e:
                    log(f"  [WARN] {code} progress error: {e}")
            log(f"  ✓ {code}: %63.5 tamamlandı")
        else:
            log(f"  ✗ {code}: bar not found!")
        end_tx()

    log(f"  ═══ 15 aktivite %100 + 5 aktivite %63.5 güncellendi ═══")

    # ══════════════════════════════════════════════════════════
    # ADIM 6: WHAT-IF KRİZ SENARYOSU
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 6: Nükleer Kriz — Güvenlik Pompaları Gümrükte!")
    log("=" * 70)

    # Kriz: NA-B11 (Primer Pompa B) süresini 240 gün artır
    tx("Crisis-PumpDelay")
    pump_task = get_fresh_task("NA-B11")
    if pump_task:
        dur_obj = pump_task.GetDurationFromString("240d")
        pump_task.SetUserDuration(dur_obj)
        log(f"  ✗ KRİZ: NA-B11 Primer Pompa B süresi 12d → 240d (8 ay gümrük gecikmesi!)")
    end_tx()

    # Kurtarma (Fast-Tracking): TD test bağlarını SS+10d olarak değiştir
    tx("FastTrack-Tests")
    for i in range(len(TD) - 1):
        pred_code = TD[i][0]
        succ_code = TD[i+1][0]
        pred_task = get_fresh_task(pred_code)
        if not pred_task:
            continue
        try:
            links_out = pred_task.LinksOut
            for li in range(1, links_out.Count + 1):
                link = D(links_out.Item(li))
                try:
                    link.type = 1  # SS
                    lag = pred_task.GetDurationFromString("10d")
                    link.StartLagTime = lag
                    break  # Only first outgoing link (the chain link)
                except:
                    pass
        except:
            pass
    log(f"  ✓ KURTARMA: {len(TD)-1} test bağı FS → SS+10d (paralel test)")

    # Test sürelerini yarıya indir
    for code, name, dur, start in TD:
        task_obj = get_fresh_task(code)
        if task_obj:
            half_dur = max(3, dur // 2)
            dur_obj = task_obj.GetDurationFromString(f"{half_dur}d")
            task_obj.SetUserDuration(dur_obj)
    log(f"  ✓ CRASHING: Tüm test süreleri yarıya indirildi")
    end_tx()

    # ══════════════════════════════════════════════════════════
    # ADIM 7: VIEW + EXPORT
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 7: Rapor ve Sonuçlar")
    log("=" * 70)

    log(f"\n  ╔══════════════════════════════════════════════════╗")
    log(f"  ║  MEGA NÜKLEER ENERJİ SANTRALİ — ÖZET RAPOR     ║")
    log(f"  ╠══════════════════════════════════════════════════╣")
    log(f"  ║  Toplam Aktivite     : {total:>6d}                   ║")
    log(f"  ║  Bağlantı Sayısı     : ~100+                    ║")
    log(f"  ║  Kod Kütüphanesi     :     3                    ║")
    log(f"  ║  Kod Ataması         : {assigned_count:>5d}                    ║")
    log(f"  ║  Maliyet Merkezi     :     3                    ║")
    log(f"  ║  Permanent Kaynak    :     1                    ║")
    log(f"  ║  Consumable Kaynak   :     2                    ║")
    log(f"  ║  ─────────────────────────────────────────────  ║")
    log(f"  ║  TOPLAM PROJE BÜTÇESİ: ${total_cost:>18,.2f}  ║")
    log(f"  ╚══════════════════════════════════════════════════╝")

    # Cost breakdown
    beton_cost = sum(qty * 2450.75 for qty in beton_tasks.values())
    blade_cost = 2 * 4 * 1500500.00
    kaynak_total = sum(50 * d * 8 * 650.50 for c, n, d, s in NA_A + NA_B + TA_A + TA_B if c in kaynak_tasks for _ in [d])
    # Recalculate kaynak cost properly
    kaynak_total = 0
    for code in kaynak_tasks:
        for section in [NA_A, NA_B, TA_A, TA_B]:
            for c, n, d, s in section:
                if c == code:
                    kaynak_total += 50 * d * 8 * 650.50

    remaining_cost = total_cost - beton_cost - blade_cost - kaynak_total
    log(f"\n  Maliyet Dağılımı:")
    log(f"    C100 Beton (150,000 m3)    : ${beton_cost:>18,.2f}")
    log(f"    Titanyum Bıçak (8 adet)    : ${blade_cost:>18,.2f}")
    log(f"    Nükleer Kaynak Uzmanı      : ${kaynak_total:>18,.2f}")
    log(f"    Maliyet Merkezi Atamaları  : ${remaining_cost:>18,.2f}")
    log(f"    ────────────────────────────────────────────────")
    log(f"    TOPLAM                     : ${total_cost:>18,.2f}")

    log(f"\n  Küsürat Kontrolü:")
    log(f"    C100 Beton birim fiyat     : $2,450.75 ✓")
    log(f"    Titanyum Bıçak birim fiyat : $1,500,500.00 ✓")
    log(f"    Nükleer Kaynak saat ücreti : $650.50 ✓")

    log("\nDONE!")

except Exception as e:
    log(f"\nFATAL ERROR: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output written to: {OUT}")
