"""
160.000 m2 SEHIR HASTANESI KOMPLEKSI — Part 1
===============================================
WBS Hierarchy + 250 Activities + 270+ Links
Global Mega Saglik Projeleri Planlama Direktoru Script

Start: 23 March 2026, Target End: September 2028 (30 months)
Budget: $250,000,000 USD
"""
import sys, os, traceback, json
sys.stdout.reconfigure(encoding='utf-8')
from datetime import datetime, timedelta

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_hospital_p1_output.txt")
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

    # ── Helpers ──
    START = "2026-03-23"
    def pt(dt_str):
        return pywintypes.Time(datetime.strptime(dt_str, "%Y-%m-%d"))

    def tx(name):
        project.StartTransaction(name)

    def end_tx():
        try:
            project.EndTransaction()
        except Exception as e:
            log(f"  [WARN] EndTx: {e}")
            try: project.AbandonTransaction()
            except: pass
        project.WaitForNotificationProcessing()

    all_bars = {}  # code -> bar_id

    def find_bar_by_id(target_id):
        """Walk tree to find bar by ID, return (bar, task)"""
        def search(parent_t):
            try:
                cbs = parent_t.ChildBars
                for i in range(1, cbs.Count + 1):
                    cb = D(cbs.Item(i))
                    if cb.ID == target_id:
                        t = D(cb.Tasks(1)) if cb.Tasks.Count > 0 else None
                        return cb, t
                    try:
                        ct = D(cb.Tasks(1))
                        r = search(ct)
                        if r: return r
                    except: pass
            except: pass
            return None
        rb = D(project.Bars.Item(1))
        rt = D(rb.ExpandedTask)
        return search(rt)

    def get_task(code):
        r = find_bar_by_id(all_bars[code])
        return r[1] if r else None

    def get_bar(code):
        r = find_bar_by_id(all_bars[code])
        return r[0] if r else None

    # ══════════════════════════════════════════════════
    # PHASE 0: ROOT SETUP
    # ══════════════════════════════════════════════════
    log("=" * 60)
    log("PHASE 0: Root Setup")
    log("=" * 60)

    bars_col = project.Bars
    if bars_col.Count == 0:
        tx("Root")
        root_bar = D(bars_col.Add())
        root_bar.Name = "Program"
        D(root_bar.Tasks.AddSummaryTask(pt(START)))
        end_tx()
    root_bar = D(bars_col.Item(1))
    root_task = D(root_bar.ExpandedTask)
    log(f"  Root: ID={root_bar.ID}")

    # ══════════════════════════════════════════════════
    # PHASE 1: CREATE WBS HIERARCHY
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 1: WBS Hierarchy")
    log("=" * 60)

    WBS_NAMES = [
        ("PROJ", "160K M2 SEHIR HASTANESI KOMPLEKSI"),
        ("WBS1", "1. Dizayn, Ruhsat ve Medikal Planlama"),
        ("WBS2", "2. Iksa, Hafriyat ve Zemin Iyilestirme"),
        ("WBS3", "3. Alt Yapi ve Radyoaktif Yalitimli Temeller"),
        ("WBS4", "4. Ust Yapi Karkas (Blok A-Yatakli, Blok B-Poliklinik, Blok C-DTC)"),
        ("WBS5", "5. Dis Cephe ve Yalitim"),
        ("WBS6", "6. MEP (Mekanik, Elektrik, Medikal Gaz ve Otomasyon)"),
        ("WBS7", "7. Ince Isler ve Temiz Oda (Ameliyathane) Imalatlari"),
        ("WBS8", "8. Medikal Cihaz Montaji (MRI, CT, Rontgen, vb.)"),
        ("WBS9", "9. Test, Devreye Alma (Commissioning) ve Saglik Bakanligi Kabulu"),
    ]

    # Create project root summary
    tx("WBS-Root")
    proj_bar = D(root_task.ChildBars.Add())
    proj_bar.Name = WBS_NAMES[0][1]
    D(proj_bar.Tasks.AddSummaryTask(pt(START)))
    proj_bar_id = proj_bar.ID
    end_tx()
    all_bars["PROJ"] = proj_bar_id
    log(f"  Created: PROJ ID={proj_bar_id}")

    # Create 9 WBS summaries under project
    for code, name in WBS_NAMES[1:]:
        tx(f"WBS-{code}")
        proj_bar = D(project.Bars.Item(1))
        proj_task = D(proj_bar.ExpandedTask)
        # Find PROJ bar
        pb = None
        for i in range(1, proj_task.ChildBars.Count + 1):
            cb = D(proj_task.ChildBars.Item(i))
            if cb.ID == proj_bar_id:
                pb = cb
                break
        pt_obj = D(pb.Tasks(1))
        wbs_bar = D(pt_obj.ChildBars.Add())
        wbs_bar.Name = name
        D(wbs_bar.Tasks.AddSummaryTask(pt(START)))
        all_bars[code] = wbs_bar.ID
        end_tx()
        log(f"  Created: {code} ID={all_bars[code]} - {name}")

    log(f"  WBS Summary bars created: {len(all_bars)}")

    # ══════════════════════════════════════════════════
    # PHASE 2: CREATE 250 ACTIVITIES
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 2: Creating 250 Activities")
    log("=" * 60)

    # Activity data: (code, name, duration_days)
    # Grouped by WBS parent
    ACTIVITIES = {
        "WBS1": [
            ("D01", "Hastane Master Plan Hazirligi", 40),
            ("D02", "Mimari Konsept Tasarim", 30),
            ("D03", "Statik ve Geoteknik Etud", 20),
            ("D04", "MEP Konsept Tasarim", 20),
            ("D05", "Medikal Planlama ve Ekipman Listesi", 20),
            ("D06", "Radyoloji Koruma Hesaplari", 15),
            ("D07", "Cevre ve CED Raporu", 20),
            ("D08", "Imar ve Yapi Ruhsati Basvurusu", 15),
            ("D09", "Saglik Bakanligi On Izin", 20),
            ("D10", "Ihale ve Taseron Secimi", 20),
            ("D11", "Uygulama Projesi - Mimari", 20),
            ("D12", "Uygulama Projesi - Statik", 20),
            ("D13", "Uygulama Projesi - MEP", 20),
            ("D14", "Medikal Gaz ve Radyoloji Detay Projesi", 20),
            ("D15", "Proje Onay ve Mobilizasyon", 10),
        ],
        "WBS2": [
            ("H01", "Santiye Kurulumu ve Guvenlik", 15),
            ("H02", "Topografik Olcum ve Aplikasyon", 10),
            ("H03", "Agac Sokumu ve Saha Temizligi", 10),
            ("H04", "Fore Kazik Imalati - Blok A", 20),
            ("H05", "Fore Kazik Imalati - Blok B", 18),
            ("H06", "Fore Kazik Imalati - Blok C", 15),
            ("H07", "Ankraj ve Iksa Sistemi - Kuzey Cephe", 20),
            ("H08", "Ankraj ve Iksa Sistemi - Guney Cephe", 18),
            ("H09", "Toprak Kazisi - Bodrum Kat 2 Seviyesi", 20),
            ("H10", "Toprak Kazisi - Bodrum Kat 1 Seviyesi", 15),
            ("H11", "Zemin Iyilestirme - Jet Grout Blok A", 15),
            ("H12", "Zemin Iyilestirme - Jet Grout Blok B", 12),
            ("H13", "Zemin Iyilestirme - Jet Grout Blok C", 10),
            ("H14", "Drenaj Sistemi Dosemesi", 12),
            ("H15", "Gecici Yol ve Rampa Yapimi", 8),
            ("H16", "Perde Duvar Imalati - Kuzey", 15),
            ("H17", "Perde Duvar Imalati - Guney", 15),
            ("H18", "Su Yalitimi - Bodrum Dis Perde", 12),
            ("H19", "Geri Dolgu ve Sikistirma", 10),
            ("H20", "Hafriyat Tamamlama ve Kabul", 5),
        ],
        "WBS3": [
            ("T01", "Grobeton Dokum - Blok A", 8),
            ("T02", "Grobeton Dokum - Blok B", 6),
            ("T03", "Grobeton Dokum - Blok C", 5),
            ("T04", "Temel Yalitim - Blok A", 10),
            ("T05", "Temel Yalitim - Blok B", 8),
            ("T06", "Temel Yalitim - Blok C", 7),
            ("T07", "Radye Temel Kalip Donati - Blok A", 15),
            ("T08", "Radye Temel Beton - Blok A", 8),
            ("T09", "Radye Temel Kalip Donati - Blok B", 12),
            ("T10", "Radye Temel Beton - Blok B", 6),
            ("T11", "Radye Temel Kalip Donati - Blok C DTC", 10),
            ("T12", "Radye Temel Beton - Blok C DTC", 5),
            ("T13", "Radyoloji Bunker Temeli - Kursun Levha Doseme", 15),
            ("T14", "Radyoloji Bunker Temeli - Ozel Beton Dokum", 12),
            ("T15", "Nukleer Tip Odasi Zemin Yalitimi", 10),
            ("T16", "Bodrum Kat Kolon ve Perde - Blok A", 18),
            ("T17", "Bodrum Kat Kolon ve Perde - Blok B", 15),
            ("T18", "Bodrum Kat Kolon ve Perde - Blok C", 12),
            ("T19", "Bodrum Kat Tabliye - Blok A", 12),
            ("T20", "Bodrum Kat Tabliye - Blok B", 10),
            ("T21", "Bodrum Kat Tabliye - Blok C", 8),
            ("T22", "Alt Yapi Kanalizasyon Hatti", 15),
            ("T23", "Alt Yapi Icme Suyu Hatti", 12),
            ("T24", "Alt Yapi Elektrik Altyapi (Trafo Temeli)", 10),
            ("T25", "Temel Tamamlama ve Kabul", 5),
        ],
        "WBS4": [
            # Blok A - Yatakli Servis (10 kat)
            ("K01", "Blok A Zemin Kat Kolon", 10),
            ("K02", "Blok A Zemin Kat Tabliye", 8),
            ("K03", "Blok A 1. Kat Kolon", 10),
            ("K04", "Blok A 1. Kat Tabliye", 8),
            ("K05", "Blok A 2. Kat Kolon", 10),
            ("K06", "Blok A 2. Kat Tabliye", 8),
            ("K07", "Blok A 3. Kat Kolon", 10),
            ("K08", "Blok A 3. Kat Tabliye", 8),
            ("K09", "Blok A 4. Kat Kolon", 10),
            ("K10", "Blok A 4. Kat Tabliye", 8),
            ("K11", "Blok A 5-7. Kat Kolon ve Tabliye", 20),
            ("K12", "Blok A 8-10. Kat Kolon ve Tabliye", 20),
            ("K13", "Blok A Cati Tabliyesi", 10),
            # Blok B - Poliklinik (5 kat)
            ("K14", "Blok B Zemin Kat Kolon", 8),
            ("K15", "Blok B Zemin Kat Tabliye", 7),
            ("K16", "Blok B 1. Kat Kolon", 8),
            ("K17", "Blok B 1. Kat Tabliye", 7),
            ("K18", "Blok B 2. Kat Kolon", 8),
            ("K19", "Blok B 2. Kat Tabliye", 7),
            ("K20", "Blok B 3-4. Kat Kolon ve Tabliye", 15),
            ("K21", "Blok B 5. Kat ve Cati Tabliye", 12),
            # Blok C - DTC (3 kat)
            ("K22", "Blok C DTC Zemin Kat Kolon", 7),
            ("K23", "Blok C DTC Zemin Kat Tabliye", 6),
            ("K24", "Blok C DTC 1. Kat Kolon", 7),
            ("K25", "Blok C DTC 1. Kat Tabliye", 6),
            ("K26", "Blok C DTC 2. Kat Kolon", 7),
            ("K27", "Blok C DTC 2. Kat Tabliye", 6),
            ("K28", "Blok C DTC 3. Kat ve Cati", 10),
            # Ortak
            ("K29", "Bloklar Arasi Baglanti Koprusu", 15),
            ("K30", "Helipad Platformu Yapimi", 12),
            ("K31", "Merdiven ve Asansor Kuyusu - Blok A", 18),
            ("K32", "Merdiven ve Asansor Kuyusu - Blok B", 14),
            ("K33", "Merdiven ve Asansor Kuyusu - Blok C", 10),
            ("K34", "Celik Cati Konstrüksiyon - Blok A", 15),
            ("K35", "Celik Cati Konstrüksiyon - Blok B", 12),
            ("K36", "Celik Cati Konstrüksiyon - Blok C", 10),
            ("K37", "Ust Yapi Karkas Tamamlama", 5),
        ],
        "WBS5": [
            ("C01", "Dis Cephe Iskelesi Kurulumu - Blok A", 12),
            ("C02", "Dis Cephe Iskelesi Kurulumu - Blok B", 10),
            ("C03", "Dis Cephe Iskelesi Kurulumu - Blok C", 8),
            ("C04", "Mantolama - Blok A", 20),
            ("C05", "Mantolama - Blok B", 15),
            ("C06", "Mantolama - Blok C", 12),
            ("C07", "Aluminyum Dograma Montaji - Blok A", 18),
            ("C08", "Aluminyum Dograma Montaji - Blok B", 14),
            ("C09", "Aluminyum Dograma Montaji - Blok C", 10),
            ("C10", "Giydirme Cephe (Curtain Wall) - Ana Giris", 15),
            ("C11", "Cam Cephe - Poliklinik Blok B", 12),
            ("C12", "Cati Su Yalitimi - Blok A", 10),
            ("C13", "Cati Su Yalitimi - Blok B", 8),
            ("C14", "Cati Su Yalitimi - Blok C", 7),
            ("C15", "Cati Isi Yalitimi ve Membran", 12),
            ("C16", "Dis Cephe Boyasi - Blok A", 10),
            ("C17", "Dis Cephe Boyasi - Blok B", 8),
            ("C18", "Dis Cephe Boyasi - Blok C", 7),
            ("C19", "Radyoloji Bolumu Kursun Cephe Kaplamasi", 12),
            ("C20", "Yangin Kacis Merdiveni Cephesi", 8),
            ("C21", "Dis Cephe Iskele Sokumu", 10),
            ("C22", "Dis Cephe Tamamlama", 5),
        ],
        "WBS6": [
            ("M01", "Ana Trafo ve Enerji Odasi", 15),
            ("M02", "Jenerator Montaji ve Baglantisi", 12),
            ("M03", "UPS Sistemleri Montaji", 10),
            ("M04", "Elektrik Ana Dagitim Panosu", 10),
            ("M05", "Kablo Tavasi ve Kablolama - Blok A", 15),
            ("M06", "Kablo Tavasi ve Kablolama - Blok B", 12),
            ("M07", "Kablo Tavasi ve Kablolama - Blok C", 10),
            ("M08", "Aydinlatma Tesisati - Blok A", 14),
            ("M09", "Aydinlatma Tesisati - Blok B", 12),
            ("M10", "Aydinlatma Tesisati - Blok C", 10),
            ("M11", "Yangin Algilama ve Alarm Sistemi", 15),
            ("M12", "Sihhi Tesisat Ana Hatlar", 12),
            ("M13", "Sihhi Tesisat Dagitim - Blok A", 14),
            ("M14", "Sihhi Tesisat Dagitim - Blok B", 12),
            ("M15", "Sihhi Tesisat Dagitim - Blok C", 10),
            ("M16", "Kazan Dairesi ve Sicak Su Sistemi", 15),
            ("M17", "Sogutma Grubu (Chiller) Montaji", 12),
            ("M18", "Klima Santrali (AHU) Montaji - Blok A", 14),
            ("M19", "Klima Santrali (AHU) Montaji - Blok B", 12),
            ("M20", "Klima Santrali (AHU) Montaji - Blok C", 10),
            ("M21", "Havalandirma Kanali (Duct) - Blok A", 15),
            ("M22", "Havalandirma Kanali (Duct) - Blok B", 12),
            ("M23", "Havalandirma Kanali (Duct) - Blok C", 10),
            ("M24", "Medikal Gaz Santrali Montaji", 12),
            ("M25", "Medikal Gaz Boru Tesisati - Blok A", 14),
            ("M26", "Medikal Gaz Boru Tesisati - Blok B", 12),
            ("M27", "Medikal Gaz Boru Tesisati - Blok C", 10),
            ("M28", "Pnomatik Tup Sistemi", 15),
            ("M29", "Asansor Montaji - Blok A (6 Adet)", 20),
            ("M30", "Asansor Montaji - Blok B (4 Adet)", 15),
            ("M31", "Asansor Montaji - Blok C (2 Adet)", 10),
            ("M32", "BMS (Bina Otomasyon) Altyapisi", 15),
            ("M33", "CCTV ve Guvenlik Sistemi", 12),
            ("M34", "Hemsire Cagri Sistemi", 10),
            ("M35", "Data ve Telekomunikasyon Altyapisi", 12),
            ("M36", "Yangin Sondurme (Sprinkler) Sistemi", 15),
            ("M37", "Mutfak Havalandirma ve Davlumbaz", 8),
            ("M38", "Camasirhane MEP Baglantilari", 8),
            ("M39", "Otopark Havalandirma ve CO Algilama", 10),
            ("M40", "MEP Tamamlama ve Koordinasyon", 5),
        ],
        "WBS7": [
            ("I01", "Duvar Orme - Blok A", 18),
            ("I02", "Duvar Orme - Blok B", 14),
            ("I03", "Duvar Orme - Blok C", 10),
            ("I04", "Siva - Blok A", 15),
            ("I05", "Siva - Blok B", 12),
            ("I06", "Siva - Blok C", 10),
            ("I07", "Seramik Kaplama - Blok A", 15),
            ("I08", "Seramik Kaplama - Blok B", 12),
            ("I09", "Seramik Kaplama - Blok C", 10),
            ("I10", "Asma Tavan Altyapisi - Blok A", 12),
            ("I11", "Asma Tavan Altyapisi - Blok B", 10),
            ("I12", "Asma Tavan Altyapisi - Blok C", 8),
            ("I13", "Asma Tavan Kapama - Blok A", 10),
            ("I14", "Asma Tavan Kapama - Blok B", 8),
            ("I15", "Asma Tavan Kapama - Blok C", 7),
            ("I16", "Epoksi Zemin - Ameliyathane Blok A", 10),
            ("I17", "Epoksi Zemin - Ameliyathane Blok C", 8),
            ("I18", "Antibakteriyel Duvar Kaplamasi - Ameliyathane", 12),
            ("I19", "Ameliyathane HEPA Filtre Montaji", 15),
            ("I20", "Ameliyathane Laminar Flow Tavan Sistemi", 12),
            ("I21", "Temiz Oda (Clean Room) Paneli Montaji", 15),
            ("I22", "Ic Boyama - Blok A", 12),
            ("I23", "Ic Boyama - Blok B", 10),
            ("I24", "Ic Boyama - Blok C", 8),
            ("I25", "Kapi ve Dograma Montaji - Blok A", 12),
            ("I26", "Kapi ve Dograma Montaji - Blok B", 10),
            ("I27", "Kapi ve Dograma Montaji - Blok C", 8),
            ("I28", "Radyoloji Odasi Kursun Kaplama", 15),
            ("I29", "Radyoloji Gozlem Cami Montaji", 8),
            ("I30", "Ince Isler Tamamlama", 5),
        ],
        "WBS8": [
            ("MC01", "MRI Cihazi Oda Hazirligi (Faraday Kafesi)", 15),
            ("MC02", "MRI Cihazi Montaji (3 Tesla) - 1 Adet", 12),
            ("MC03", "MRI Cihazi Montaji (1.5 Tesla) - 2 Adet", 15),
            ("MC04", "CT Cihazi Montaji - 3 Adet", 15),
            ("MC05", "Anjiyografi Cihazi Montaji - 2 Adet", 12),
            ("MC06", "Dijital Rontgen Montaji - 5 Adet", 10),
            ("MC07", "Ultrason Cihazi Montaji - 8 Adet", 8),
            ("MC08", "Mammografi Cihazi Montaji - 2 Adet", 5),
            ("MC09", "Ameliyathane Masa ve Lamba Montaji", 12),
            ("MC10", "Ameliyathane Pendanlar ve Kollar", 10),
            ("MC11", "Sterilizasyon Unitesi Montaji", 10),
            ("MC12", "Yogun Bakim Yatak Basi Uniteleri", 12),
            ("MC13", "Laboratuvar Cihazlari Montaji", 15),
            ("MC14", "Eczane Otomasyon Sistemi", 10),
            ("MC15", "PACS (Goruntu Arsiv) Sistemi Kurulumu", 8),
            ("MC16", "HIS (Hastane Bilgi Sistemi) Kurulumu", 12),
            ("MC17", "Hasta Yatak ve Mobilya Montaji - Blok A", 15),
            ("MC18", "Hasta Yatak ve Mobilya Montaji - Blok B", 10),
            ("MC19", "Poliklinik Muayene Odasi Donatimi", 10),
            ("MC20", "Acil Servis Ekipman Montaji", 8),
            ("MC21", "Liner Akselerator (LINAC) Montaji", 15),
            ("MC22", "Nukleer Tip Gamma Kamera Montaji", 10),
            ("MC23", "Endoskopi Unitesi Donatimi", 8),
            ("MC24", "Diyaliz Unitesi Cihaz Montaji", 8),
            ("MC25", "Medikal Cihaz Tamamlama", 5),
        ],
        "WBS9": [
            ("TC01", "Elektrik Sistemleri Test", 10),
            ("TC02", "Jenerator Yuk Testi", 5),
            ("TC03", "UPS Yuk ve Gecis Testi", 5),
            ("TC04", "Yangin Algilama Sistemi Testi", 8),
            ("TC05", "Sprinkler Sistemi Basinc Testi", 7),
            ("TC06", "Sihhi Tesisat Basinc Testi", 7),
            ("TC07", "Kazan ve Sicak Su Sistemi Testi", 5),
            ("TC08", "Chiller ve Sogutma Sistemi Testi", 7),
            ("TC09", "Klima ve Havalandirma Dengeleme", 10),
            ("TC10", "Medikal Gaz Test ve Sertifikasyon", 8),
            ("TC11", "Asansor Test ve Sertifikasyon", 10),
            ("TC12", "BMS Entegrasyon Testi", 10),
            ("TC13", "CCTV ve Guvenlik Test", 5),
            ("TC14", "HIS ve PACS Entegrasyon Testi", 8),
            ("TC15", "Pnomatik Tup Sistemi Testi", 5),
            ("TC16", "Hemsire Cagri Sistemi Testi", 5),
            ("TC17", "Ameliyathane Temiz Oda Validasyonu", 10),
            ("TC18", "Radyoloji Radyasyon Kacak Testi", 8),
            ("TC19", "MRI Manyetik Alan Testi", 5),
            ("TC20", "LINAC Radyasyon Guvenlik Testi", 7),
            ("TC21", "Ic Ortam Hava Kalitesi Testi", 5),
            ("TC22", "Akustik Test - Ameliyathane ve YBU", 5),
            ("TC23", "Su Kalitesi ve Legionella Testi", 5),
            ("TC24", "Engelli Erisim Denetimi", 5),
            ("TC25", "Itfaiye Onayi ve Yangin Tatbikati", 5),
            ("TC26", "Cevre ve Atik Yonetim Onayi", 5),
            ("TC27", "Saglik Bakanligi Teknik Inceleme", 10),
            ("TC28", "Saglik Bakanligi Nihai Kabul", 10),
            ("TC29", "Belediye ve Iskan Onayi", 8),
            ("TC30", "Peyzaj ve Cevre Duzenleme", 15),
            ("TC31", "Otopark Cizgileme ve Yonlendirme", 5),
            ("TC32", "Tabela ve Yonlendirme Sistemi", 8),
            ("TC33", "Gecici Santiye Sokumu", 8),
            ("TC34", "Hasta Kabul Simulasyonu", 5),
            ("TC35", "Personel Egitimi", 10),
            ("TC36", "Gecici Kabul Tutanagi", 3),
        ],
    }

    # Count total
    total_count = sum(len(v) for v in ACTIVITIES.values())
    log(f"  Total activities to create: {total_count}")

    # Create activities per WBS group
    for wbs_code, act_list in ACTIVITIES.items():
        wbs_bar_id = all_bars[wbs_code]
        log(f"\n  Creating {len(act_list)} activities under {wbs_code}...")

        for code, name, dur_days in act_list:
            tx(f"Add-{code}")
            # Re-find WBS bar
            r = find_bar_by_id(wbs_bar_id)
            if not r:
                log(f"    ERROR: Cannot find WBS bar {wbs_code}")
                end_tx()
                continue
            wbs_bar_obj, wbs_task_obj = r

            # Add child bar + task
            new_bar = D(wbs_task_obj.ChildBars.Add())
            new_bar.Name = name
            new_task = D(new_bar.Tasks.AddTask(pt(START), f"{dur_days}d"))
            bar_id = new_bar.ID
            all_bars[code] = bar_id
            end_tx()

        log(f"    Done: {len(act_list)} activities, last ID={all_bars[act_list[-1][0]]}")

    log(f"\n  Total bars created: {len(all_bars)} (including WBS summaries)")

    # ══════════════════════════════════════════════════
    # PHASE 3: CREATE LINKS (270+)
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 3: Creating Links")
    log("=" * 60)

    # Link data: (from_code, to_code, type, lag_days)
    # type: 0=FS, 1=SS, 2=FF, 3=SF
    LINKS = [
        # ── WBS1: Design ──
        ("D01", "D02", 0, 0),
        ("D01", "D03", 1, 10),
        ("D02", "D04", 1, 10),
        ("D02", "D05", 1, 5),
        ("D05", "D06", 0, 0),
        ("D03", "D07", 1, 10),
        ("D07", "D08", 0, 0),
        ("D08", "D09", 1, 5),
        ("D09", "D10", 0, 0),
        ("D02", "D11", 0, 0),
        ("D03", "D12", 1, 10),
        ("D04", "D13", 1, 10),
        ("D06", "D14", 0, 0),
        ("D10", "D15", 0, 0),
        ("D11", "D15", 0, 0),
        ("D12", "D15", 0, 0),
        ("D13", "D15", 0, 0),
        ("D14", "D15", 0, 0),

        # ── WBS1 → WBS2 ──
        ("D15", "H01", 0, 0),

        # ── WBS2: Excavation ──
        ("H01", "H02", 0, 0),
        ("H02", "H03", 0, 0),
        ("H03", "H04", 1, 5),
        ("H03", "H05", 1, 8),
        ("H03", "H06", 1, 10),
        ("H04", "H07", 1, 10),
        ("H05", "H08", 1, 10),
        ("H04", "H09", 1, 10),
        ("H09", "H10", 0, 0),
        ("H04", "H11", 0, 0),
        ("H05", "H12", 0, 0),
        ("H06", "H13", 0, 0),
        ("H10", "H14", 1, 5),
        ("H09", "H15", 1, 5),
        ("H07", "H16", 0, 0),
        ("H08", "H17", 0, 0),
        ("H16", "H18", 0, 0),
        ("H17", "H18", 1, 5),
        ("H18", "H19", 0, 0),
        ("H19", "H20", 0, 0),
        ("H11", "H20", 0, 0),
        ("H12", "H20", 0, 0),
        ("H13", "H20", 0, 0),
        ("H14", "H20", 0, 0),

        # ── WBS2 → WBS3 ──
        ("H20", "T01", 0, 0),
        ("H20", "T02", 1, 3),
        ("H20", "T03", 1, 5),

        # ── WBS3: Foundation ──
        ("T01", "T04", 0, 0),
        ("T02", "T05", 0, 0),
        ("T03", "T06", 0, 0),
        ("T04", "T07", 0, 0),
        ("T05", "T09", 0, 0),
        ("T06", "T11", 0, 0),
        ("T07", "T08", 0, 0),
        ("T09", "T10", 0, 0),
        ("T11", "T12", 0, 0),
        ("T08", "T13", 0, 0),
        ("T13", "T14", 0, 0),
        ("T14", "T15", 0, 0),
        ("T08", "T16", 0, 0),
        ("T10", "T17", 0, 0),
        ("T12", "T18", 0, 0),
        ("T16", "T19", 0, 0),
        ("T17", "T20", 0, 0),
        ("T18", "T21", 0, 0),
        ("T19", "T22", 1, 5),
        ("T19", "T23", 1, 5),
        ("T19", "T24", 1, 5),
        ("T19", "T25", 0, 0),
        ("T20", "T25", 0, 0),
        ("T21", "T25", 0, 0),
        ("T22", "T25", 0, 0),
        ("T23", "T25", 0, 0),
        ("T24", "T25", 0, 0),
        ("T15", "T25", 0, 0),

        # ── WBS3 → WBS4 ──
        ("T19", "K01", 0, 0),
        ("T20", "K14", 0, 0),
        ("T21", "K22", 0, 0),

        # ── WBS4: Karkas ──
        # Blok A floor-by-floor
        ("K01", "K02", 0, 0),
        ("K02", "K03", 0, 0),
        ("K03", "K04", 0, 0),
        ("K04", "K05", 0, 0),
        ("K05", "K06", 0, 0),
        ("K06", "K07", 0, 0),
        ("K07", "K08", 0, 0),
        ("K08", "K09", 0, 0),
        ("K09", "K10", 0, 0),
        ("K10", "K11", 0, 0),
        ("K11", "K12", 0, 0),
        ("K12", "K13", 0, 0),
        # Blok B
        ("K14", "K15", 0, 0),
        ("K15", "K16", 0, 0),
        ("K16", "K17", 0, 0),
        ("K17", "K18", 0, 0),
        ("K18", "K19", 0, 0),
        ("K19", "K20", 0, 0),
        ("K20", "K21", 0, 0),
        # Blok C
        ("K22", "K23", 0, 0),
        ("K23", "K24", 0, 0),
        ("K24", "K25", 0, 0),
        ("K25", "K26", 0, 0),
        ("K26", "K27", 0, 0),
        ("K27", "K28", 0, 0),
        # Ortak
        ("K13", "K29", 1, 5),
        ("K21", "K29", 1, 5),
        ("K13", "K30", 0, 0),
        ("K02", "K31", 1, 5),
        ("K15", "K32", 1, 5),
        ("K23", "K33", 1, 5),
        ("K13", "K34", 0, 0),
        ("K21", "K35", 0, 0),
        ("K28", "K36", 0, 0),
        ("K34", "K37", 0, 0),
        ("K35", "K37", 0, 0),
        ("K36", "K37", 0, 0),
        ("K29", "K37", 0, 0),
        ("K30", "K37", 0, 0),

        # ── WBS4 → WBS5 ──
        ("K13", "C01", 0, 0),
        ("K21", "C02", 0, 0),
        ("K28", "C03", 0, 0),

        # ── WBS5: Cephe ──
        ("C01", "C04", 0, 0),
        ("C02", "C05", 0, 0),
        ("C03", "C06", 0, 0),
        ("C04", "C07", 0, 0),
        ("C05", "C08", 0, 0),
        ("C06", "C09", 0, 0),
        ("C04", "C10", 1, 10),
        ("C05", "C11", 1, 5),
        ("K34", "C12", 0, 0),
        ("K35", "C13", 0, 0),
        ("K36", "C14", 0, 0),
        ("C12", "C15", 0, 0),
        ("C13", "C15", 1, 3),
        ("C14", "C15", 1, 5),
        ("C07", "C16", 0, 0),
        ("C08", "C17", 0, 0),
        ("C09", "C18", 0, 0),
        ("C04", "C19", 1, 10),
        ("C07", "C20", 1, 5),
        ("C16", "C21", 0, 0),
        ("C17", "C21", 1, 3),
        ("C18", "C21", 1, 5),
        ("C21", "C22", 0, 0),
        ("C19", "C22", 0, 0),

        # ── Cross: Karkas → MEP (early start for heavy MEP) ──
        ("K10", "M05", 1, 5),   # cable tray Blok A after 4th floor
        ("K19", "M06", 1, 5),   # cable tray Blok B after 2nd floor
        ("K27", "M07", 1, 5),   # cable tray Blok C after 2nd floor
        ("K34", "M01", 0, 0),   # trafo after steel roof Blok A
        ("K31", "M29", 0, 0),   # elevator after shaft Blok A
        ("K32", "M30", 0, 0),
        ("K33", "M31", 0, 0),

        # ── Cross: Cephe → MEP (after windows) ──
        ("C07", "M08", 0, 0),   # lighting Blok A after windows
        ("C08", "M09", 0, 0),
        ("C09", "M10", 0, 0),

        # ── WBS6: MEP ──
        ("M01", "M02", 0, 0),
        ("M02", "M03", 0, 0),
        ("M01", "M04", 1, 5),
        ("M04", "M05", 0, 0),
        ("M05", "M08", 1, 10),
        ("M06", "M09", 1, 10),
        ("M07", "M10", 1, 10),
        ("M01", "M11", 1, 5),
        ("M04", "M12", 0, 0),
        ("M12", "M13", 0, 0),
        ("M12", "M14", 1, 5),
        ("M12", "M15", 1, 8),
        ("M01", "M16", 1, 5),
        ("M16", "M17", 0, 0),
        ("M17", "M18", 0, 0),
        ("M17", "M19", 1, 5),
        ("M17", "M20", 1, 8),
        ("M18", "M21", 0, 0),
        ("M19", "M22", 0, 0),
        ("M20", "M23", 0, 0),
        ("M16", "M24", 1, 5),
        ("M24", "M25", 0, 0),
        ("M24", "M26", 1, 5),
        ("M24", "M27", 1, 8),
        ("M25", "M28", 1, 10),
        ("M04", "M32", 1, 10),
        ("M11", "M33", 1, 5),
        ("M32", "M34", 1, 5),
        ("M04", "M35", 1, 5),
        ("M11", "M36", 0, 0),
        ("M18", "M37", 1, 5),
        ("M16", "M38", 1, 10),
        ("M21", "M39", 1, 5),
        ("M36", "M40", 0, 0),
        ("M28", "M40", 0, 0),
        ("M39", "M40", 0, 0),

        # ── Cross: MEP → Ince Is (ductwork → ceiling) ──
        ("M21", "I10", 0, 0),   # duct Blok A → ceiling grid Blok A
        ("M22", "I11", 0, 0),
        ("M23", "I12", 0, 0),
        # CRITICAL: Medikal Gaz → Asma Tavan Kapama
        ("M25", "I13", 0, 0),   # medikal gaz Blok A → ceiling close Blok A
        ("M26", "I14", 0, 0),
        ("M27", "I15", 0, 0),

        # ── Cross: Karkas → Ince Is (walls start after partial karkas) ──
        ("K10", "I01", 1, 20),  # walls Blok A after 4th floor slab
        ("K19", "I02", 1, 10),
        ("K27", "I03", 1, 5),

        # ── WBS7: Ince Is ──
        ("I01", "I04", 0, 0),
        ("I02", "I05", 0, 0),
        ("I03", "I06", 0, 0),
        ("I04", "I07", 0, 0),
        ("I05", "I08", 0, 0),
        ("I06", "I09", 0, 0),
        ("I10", "I13", 0, 0),   # ceiling grid → ceiling close (also depends on M25)
        ("I11", "I14", 0, 0),
        ("I12", "I15", 0, 0),
        ("I13", "I16", 0, 0),   # ceiling close → epoksi ameliyathane
        ("I15", "I17", 0, 0),
        ("I16", "I18", 0, 0),   # epoksi → antibacterial wall
        ("I18", "I19", 0, 0),   # antibacterial → HEPA
        ("I19", "I20", 0, 0),   # HEPA → laminar flow
        ("I20", "I21", 0, 0),   # laminar → clean room
        ("I07", "I22", 0, 0),
        ("I08", "I23", 0, 0),
        ("I09", "I24", 0, 0),
        ("I22", "I25", 0, 0),
        ("I23", "I26", 0, 0),
        ("I24", "I27", 0, 0),
        ("I18", "I28", 0, 0),   # antibacterial → kursun kaplama
        ("I28", "I29", 0, 0),
        ("I25", "I30", 0, 0),
        ("I26", "I30", 0, 0),
        ("I27", "I30", 0, 0),
        ("I21", "I30", 0, 0),
        ("I29", "I30", 0, 0),

        # ── WBS7 → WBS8 ──
        ("I19", "MC01", 0, 0),   # HEPA → MRI room prep
        ("I28", "MC02", 1, 5),   # kursun kaplama → MRI install
        ("I21", "MC09", 0, 0),   # clean room → surgery tables
        ("I30", "MC17", 0, 0),   # finishing → furniture Blok A

        # ── WBS8: Medikal Cihaz ──
        ("MC01", "MC02", 0, 0),
        ("MC01", "MC03", 1, 5),
        ("MC02", "MC04", 1, 5),
        ("MC03", "MC04", 1, 3),
        ("MC04", "MC05", 1, 5),
        ("MC04", "MC06", 1, 3),
        ("MC06", "MC07", 1, 3),
        ("MC07", "MC08", 1, 3),
        ("MC09", "MC10", 0, 0),
        ("MC10", "MC11", 1, 5),
        ("MC11", "MC12", 1, 5),
        ("MC12", "MC13", 1, 3),
        ("MC13", "MC14", 1, 5),
        ("MC14", "MC15", 0, 0),
        ("MC15", "MC16", 0, 0),
        ("MC17", "MC18", 1, 5),
        ("MC18", "MC19", 1, 3),
        ("MC19", "MC20", 1, 3),
        ("I28", "MC21", 0, 0),   # LINAC needs lead shielding
        ("MC21", "MC22", 1, 5),
        ("MC22", "MC23", 1, 3),
        ("MC23", "MC24", 1, 3),
        ("MC16", "MC25", 0, 0),
        ("MC24", "MC25", 0, 0),
        ("MC20", "MC25", 0, 0),

        # ── WBS8 → WBS9 ──
        ("MC25", "TC01", 0, 0),

        # ── WBS9: Commissioning ──
        ("TC01", "TC02", 0, 0),
        ("TC02", "TC03", 0, 0),
        ("TC01", "TC04", 1, 3),
        ("TC01", "TC05", 1, 3),
        ("TC01", "TC06", 1, 3),
        ("TC06", "TC07", 0, 0),
        ("TC07", "TC08", 0, 0),
        ("TC08", "TC09", 0, 0),
        ("TC09", "TC10", 0, 0),
        ("TC10", "TC11", 1, 3),
        ("TC11", "TC12", 0, 0),
        ("TC12", "TC13", 1, 3),
        ("TC13", "TC14", 0, 0),
        ("TC14", "TC15", 1, 3),
        ("TC15", "TC16", 1, 3),
        ("TC16", "TC17", 0, 0),
        ("TC17", "TC18", 0, 0),
        ("TC18", "TC19", 0, 0),
        ("TC19", "TC20", 0, 0),
        ("TC09", "TC21", 0, 0),
        ("TC17", "TC22", 0, 0),
        ("TC21", "TC23", 0, 0),
        ("TC22", "TC24", 1, 3),
        ("TC24", "TC25", 0, 0),
        ("TC23", "TC26", 1, 3),
        ("TC25", "TC27", 0, 0),
        ("TC26", "TC27", 1, 3),
        ("TC27", "TC28", 0, 0),
        ("TC28", "TC29", 1, 5),
        ("C22", "TC30", 1, 10),  # landscaping after facade
        ("TC30", "TC31", 0, 0),
        ("TC31", "TC32", 0, 0),
        ("TC28", "TC33", 0, 0),
        ("TC33", "TC34", 0, 0),
        ("TC34", "TC35", 0, 0),
        ("TC35", "TC36", 0, 0),
        ("TC29", "TC36", 0, 0),
        ("TC32", "TC36", 0, 0),
    ]

    log(f"  Total links to create: {len(LINKS)}")
    link_ok = 0
    link_fail = 0

    for from_code, to_code, link_type, lag_days in LINKS:
        if from_code not in all_bars or to_code not in all_bars:
            log(f"    SKIP: {from_code}->{to_code} (bar not found)")
            link_fail += 1
            continue

        tx(f"L-{from_code}-{to_code}")
        try:
            t1 = get_task(from_code)
            t2 = get_task(to_code)
            if not t1 or not t2:
                log(f"    SKIP: {from_code}->{to_code} (task not found)")
                link_fail += 1
                end_tx()
                continue
            link = D(t1.LinkTo(t2))
            if link_type != 0:
                link.Type = link_type
            if lag_days > 0:
                lag_dur = t1.GetDurationFromString(f"{lag_days}d")
                link.StartLagTime = lag_dur
            link_ok += 1
        except Exception as e:
            log(f"    ERROR: {from_code}->{to_code}: {e}")
            link_fail += 1
        end_tx()

        if link_ok % 50 == 0 and link_ok > 0:
            log(f"    Links created: {link_ok}...")

    log(f"  Links: OK={link_ok}, Failed={link_fail}")

    # ══════════════════════════════════════════════════
    # PHASE 4: RESCHEDULE
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 4: Reschedule")
    log("=" * 60)

    project.Reschedule()
    log("  Reschedule complete!")

    project.Save()
    log("  Project saved!")

    # Save bar ID mapping for Part 2
    mapping_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hospital_bar_ids.json")
    with open(mapping_file, "w") as mf:
        json.dump(all_bars, mf, indent=2)
    log(f"  Bar ID mapping saved to: {mapping_file}")
    log(f"  Total entries: {len(all_bars)}")

    # Summary
    log("\n" + "=" * 60)
    log("SUMMARY")
    log("=" * 60)
    log(f"  Activities created: {total_count}")
    log(f"  Links created: {link_ok}")
    log(f"  WBS groups: 9")
    log(f"  Bar IDs saved for Part 2")
    log("\nPart 1 DONE!")

except Exception as e:
    log(f"FATAL ERROR: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output: {OUT}")
