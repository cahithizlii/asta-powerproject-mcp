"""
FAST link completion + bar_ids export.
Key optimization: Traverse tree ONCE per batch, cache all refs.
Batch 20 links per transaction instead of 1.
"""
import sys, os, traceback, json
sys.stdout.reconfigure(encoding='utf-8')
from datetime import datetime

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_hospital_fast_links_output.txt")
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

    # ── ACTIVITY NAME → CODE reverse mapping ──
    ACTIVITIES_FLAT = [
        ("D01","Hastane Master Plan Hazirligi"),("D02","Mimari Konsept Tasarim"),
        ("D03","Statik ve Geoteknik Etud"),("D04","MEP Konsept Tasarim"),
        ("D05","Medikal Planlama ve Ekipman Listesi"),("D06","Radyoloji Koruma Hesaplari"),
        ("D07","Cevre ve CED Raporu"),("D08","Imar ve Yapi Ruhsati Basvurusu"),
        ("D09","Saglik Bakanligi On Izin"),("D10","Ihale ve Taseron Secimi"),
        ("D11","Uygulama Projesi - Mimari"),("D12","Uygulama Projesi - Statik"),
        ("D13","Uygulama Projesi - MEP"),("D14","Medikal Gaz ve Radyoloji Detay Projesi"),
        ("D15","Proje Onay ve Mobilizasyon"),
        ("H01","Santiye Kurulumu ve Guvenlik"),("H02","Topografik Olcum ve Aplikasyon"),
        ("H03","Agac Sokumu ve Saha Temizligi"),("H04","Fore Kazik Imalati - Blok A"),
        ("H05","Fore Kazik Imalati - Blok B"),("H06","Fore Kazik Imalati - Blok C"),
        ("H07","Ankraj ve Iksa Sistemi - Kuzey Cephe"),("H08","Ankraj ve Iksa Sistemi - Guney Cephe"),
        ("H09","Toprak Kazisi - Bodrum Kat 2 Seviyesi"),("H10","Toprak Kazisi - Bodrum Kat 1 Seviyesi"),
        ("H11","Zemin Iyilestirme - Jet Grout Blok A"),("H12","Zemin Iyilestirme - Jet Grout Blok B"),
        ("H13","Zemin Iyilestirme - Jet Grout Blok C"),("H14","Drenaj Sistemi Dosemesi"),
        ("H15","Gecici Yol ve Rampa Yapimi"),("H16","Perde Duvar Imalati - Kuzey"),
        ("H17","Perde Duvar Imalati - Guney"),("H18","Su Yalitimi - Bodrum Dis Perde"),
        ("H19","Geri Dolgu ve Sikistirma"),("H20","Hafriyat Tamamlama ve Kabul"),
        ("T01","Grobeton Dokum - Blok A"),("T02","Grobeton Dokum - Blok B"),
        ("T03","Grobeton Dokum - Blok C"),("T04","Temel Yalitim - Blok A"),
        ("T05","Temel Yalitim - Blok B"),("T06","Temel Yalitim - Blok C"),
        ("T07","Radye Temel Kalip Donati - Blok A"),("T08","Radye Temel Beton - Blok A"),
        ("T09","Radye Temel Kalip Donati - Blok B"),("T10","Radye Temel Beton - Blok B"),
        ("T11","Radye Temel Kalip Donati - Blok C DTC"),("T12","Radye Temel Beton - Blok C DTC"),
        ("T13","Radyoloji Bunker Temeli - Kursun Levha Doseme"),
        ("T14","Radyoloji Bunker Temeli - Ozel Beton Dokum"),
        ("T15","Nukleer Tip Odasi Zemin Yalitimi"),
        ("T16","Bodrum Kat Kolon ve Perde - Blok A"),("T17","Bodrum Kat Kolon ve Perde - Blok B"),
        ("T18","Bodrum Kat Kolon ve Perde - Blok C"),
        ("T19","Bodrum Kat Tabliye - Blok A"),("T20","Bodrum Kat Tabliye - Blok B"),
        ("T21","Bodrum Kat Tabliye - Blok C"),
        ("T22","Alt Yapi Kanalizasyon Hatti"),("T23","Alt Yapi Icme Suyu Hatti"),
        ("T24","Alt Yapi Elektrik Altyapi (Trafo Temeli)"),("T25","Temel Tamamlama ve Kabul"),
        ("K01","Blok A Zemin Kat Kolon"),("K02","Blok A Zemin Kat Tabliye"),
        ("K03","Blok A 1. Kat Kolon"),("K04","Blok A 1. Kat Tabliye"),
        ("K05","Blok A 2. Kat Kolon"),("K06","Blok A 2. Kat Tabliye"),
        ("K07","Blok A 3. Kat Kolon"),("K08","Blok A 3. Kat Tabliye"),
        ("K09","Blok A 4. Kat Kolon"),("K10","Blok A 4. Kat Tabliye"),
        ("K11","Blok A 5-7. Kat Kolon ve Tabliye"),("K12","Blok A 8-10. Kat Kolon ve Tabliye"),
        ("K13","Blok A Cati Tabliyesi"),
        ("K14","Blok B Zemin Kat Kolon"),("K15","Blok B Zemin Kat Tabliye"),
        ("K16","Blok B 1. Kat Kolon"),("K17","Blok B 1. Kat Tabliye"),
        ("K18","Blok B 2. Kat Kolon"),("K19","Blok B 2. Kat Tabliye"),
        ("K20","Blok B 3-4. Kat Kolon ve Tabliye"),("K21","Blok B 5. Kat ve Cati Tabliye"),
        ("K22","Blok C DTC Zemin Kat Kolon"),("K23","Blok C DTC Zemin Kat Tabliye"),
        ("K24","Blok C DTC 1. Kat Kolon"),("K25","Blok C DTC 1. Kat Tabliye"),
        ("K26","Blok C DTC 2. Kat Kolon"),("K27","Blok C DTC 2. Kat Tabliye"),
        ("K28","Blok C DTC 3. Kat ve Cati"),
        ("K29","Bloklar Arasi Baglanti Koprusu"),("K30","Helipad Platformu Yapimi"),
        ("K31","Merdiven ve Asansor Kuyusu - Blok A"),("K32","Merdiven ve Asansor Kuyusu - Blok B"),
        ("K33","Merdiven ve Asansor Kuyusu - Blok C"),
        ("K34","Celik Cati Konstrüksiyon - Blok A"),("K35","Celik Cati Konstrüksiyon - Blok B"),
        ("K36","Celik Cati Konstrüksiyon - Blok C"),("K37","Ust Yapi Karkas Tamamlama"),
        ("C01","Dis Cephe Iskelesi Kurulumu - Blok A"),("C02","Dis Cephe Iskelesi Kurulumu - Blok B"),
        ("C03","Dis Cephe Iskelesi Kurulumu - Blok C"),
        ("C04","Mantolama - Blok A"),("C05","Mantolama - Blok B"),("C06","Mantolama - Blok C"),
        ("C07","Aluminyum Dograma Montaji - Blok A"),("C08","Aluminyum Dograma Montaji - Blok B"),
        ("C09","Aluminyum Dograma Montaji - Blok C"),
        ("C10","Giydirme Cephe (Curtain Wall) - Ana Giris"),("C11","Cam Cephe - Poliklinik Blok B"),
        ("C12","Cati Su Yalitimi - Blok A"),("C13","Cati Su Yalitimi - Blok B"),
        ("C14","Cati Su Yalitimi - Blok C"),("C15","Cati Isi Yalitimi ve Membran"),
        ("C16","Dis Cephe Boyasi - Blok A"),("C17","Dis Cephe Boyasi - Blok B"),
        ("C18","Dis Cephe Boyasi - Blok C"),
        ("C19","Radyoloji Bolumu Kursun Cephe Kaplamasi"),("C20","Yangin Kacis Merdiveni Cephesi"),
        ("C21","Dis Cephe Iskele Sokumu"),("C22","Dis Cephe Tamamlama"),
        ("M01","Ana Trafo ve Enerji Odasi"),("M02","Jenerator Montaji ve Baglantisi"),
        ("M03","UPS Sistemleri Montaji"),("M04","Elektrik Ana Dagitim Panosu"),
        ("M05","Kablo Tavasi ve Kablolama - Blok A"),("M06","Kablo Tavasi ve Kablolama - Blok B"),
        ("M07","Kablo Tavasi ve Kablolama - Blok C"),
        ("M08","Aydinlatma Tesisati - Blok A"),("M09","Aydinlatma Tesisati - Blok B"),
        ("M10","Aydinlatma Tesisati - Blok C"),
        ("M11","Yangin Algilama ve Alarm Sistemi"),("M12","Sihhi Tesisat Ana Hatlar"),
        ("M13","Sihhi Tesisat Dagitim - Blok A"),("M14","Sihhi Tesisat Dagitim - Blok B"),
        ("M15","Sihhi Tesisat Dagitim - Blok C"),
        ("M16","Kazan Dairesi ve Sicak Su Sistemi"),("M17","Sogutma Grubu (Chiller) Montaji"),
        ("M18","Klima Santrali (AHU) Montaji - Blok A"),("M19","Klima Santrali (AHU) Montaji - Blok B"),
        ("M20","Klima Santrali (AHU) Montaji - Blok C"),
        ("M21","Havalandirma Kanali (Duct) - Blok A"),("M22","Havalandirma Kanali (Duct) - Blok B"),
        ("M23","Havalandirma Kanali (Duct) - Blok C"),
        ("M24","Medikal Gaz Santrali Montaji"),
        ("M25","Medikal Gaz Boru Tesisati - Blok A"),("M26","Medikal Gaz Boru Tesisati - Blok B"),
        ("M27","Medikal Gaz Boru Tesisati - Blok C"),
        ("M28","Pnomatik Tup Sistemi"),
        ("M29","Asansor Montaji - Blok A (6 Adet)"),("M30","Asansor Montaji - Blok B (4 Adet)"),
        ("M31","Asansor Montaji - Blok C (2 Adet)"),
        ("M32","BMS (Bina Otomasyon) Altyapisi"),("M33","CCTV ve Guvenlik Sistemi"),
        ("M34","Hemsire Cagri Sistemi"),("M35","Data ve Telekomunikasyon Altyapisi"),
        ("M36","Yangin Sondurme (Sprinkler) Sistemi"),("M37","Mutfak Havalandirma ve Davlumbaz"),
        ("M38","Camasirhane MEP Baglantilari"),("M39","Otopark Havalandirma ve CO Algilama"),
        ("M40","MEP Tamamlama ve Koordinasyon"),
        ("I01","Duvar Orme - Blok A"),("I02","Duvar Orme - Blok B"),("I03","Duvar Orme - Blok C"),
        ("I04","Siva - Blok A"),("I05","Siva - Blok B"),("I06","Siva - Blok C"),
        ("I07","Seramik Kaplama - Blok A"),("I08","Seramik Kaplama - Blok B"),
        ("I09","Seramik Kaplama - Blok C"),
        ("I10","Asma Tavan Altyapisi - Blok A"),("I11","Asma Tavan Altyapisi - Blok B"),
        ("I12","Asma Tavan Altyapisi - Blok C"),
        ("I13","Asma Tavan Kapama - Blok A"),("I14","Asma Tavan Kapama - Blok B"),
        ("I15","Asma Tavan Kapama - Blok C"),
        ("I16","Epoksi Zemin - Ameliyathane Blok A"),("I17","Epoksi Zemin - Ameliyathane Blok C"),
        ("I18","Antibakteriyel Duvar Kaplamasi - Ameliyathane"),
        ("I19","Ameliyathane HEPA Filtre Montaji"),("I20","Ameliyathane Laminar Flow Tavan Sistemi"),
        ("I21","Temiz Oda (Clean Room) Paneli Montaji"),
        ("I22","Ic Boyama - Blok A"),("I23","Ic Boyama - Blok B"),("I24","Ic Boyama - Blok C"),
        ("I25","Kapi ve Dograma Montaji - Blok A"),("I26","Kapi ve Dograma Montaji - Blok B"),
        ("I27","Kapi ve Dograma Montaji - Blok C"),
        ("I28","Radyoloji Odasi Kursun Kaplama"),("I29","Radyoloji Gozlem Cami Montaji"),
        ("I30","Ince Isler Tamamlama"),
        ("MC01","MRI Cihazi Oda Hazirligi (Faraday Kafesi)"),
        ("MC02","MRI Cihazi Montaji (3 Tesla) - 1 Adet"),
        ("MC03","MRI Cihazi Montaji (1.5 Tesla) - 2 Adet"),
        ("MC04","CT Cihazi Montaji - 3 Adet"),("MC05","Anjiyografi Cihazi Montaji - 2 Adet"),
        ("MC06","Dijital Rontgen Montaji - 5 Adet"),("MC07","Ultrason Cihazi Montaji - 8 Adet"),
        ("MC08","Mammografi Cihazi Montaji - 2 Adet"),
        ("MC09","Ameliyathane Masa ve Lamba Montaji"),("MC10","Ameliyathane Pendanlar ve Kollar"),
        ("MC11","Sterilizasyon Unitesi Montaji"),("MC12","Yogun Bakim Yatak Basi Uniteleri"),
        ("MC13","Laboratuvar Cihazlari Montaji"),("MC14","Eczane Otomasyon Sistemi"),
        ("MC15","PACS (Goruntu Arsiv) Sistemi Kurulumu"),
        ("MC16","HIS (Hastane Bilgi Sistemi) Kurulumu"),
        ("MC17","Hasta Yatak ve Mobilya Montaji - Blok A"),
        ("MC18","Hasta Yatak ve Mobilya Montaji - Blok B"),
        ("MC19","Poliklinik Muayene Odasi Donatimi"),("MC20","Acil Servis Ekipman Montaji"),
        ("MC21","Liner Akselerator (LINAC) Montaji"),
        ("MC22","Nukleer Tip Gamma Kamera Montaji"),("MC23","Endoskopi Unitesi Donatimi"),
        ("MC24","Diyaliz Unitesi Cihaz Montaji"),("MC25","Medikal Cihaz Tamamlama"),
        ("TC01","Elektrik Sistemleri Test"),("TC02","Jenerator Yuk Testi"),
        ("TC03","UPS Yuk ve Gecis Testi"),("TC04","Yangin Algilama Sistemi Testi"),
        ("TC05","Sprinkler Sistemi Basinc Testi"),("TC06","Sihhi Tesisat Basinc Testi"),
        ("TC07","Kazan ve Sicak Su Sistemi Testi"),("TC08","Chiller ve Sogutma Sistemi Testi"),
        ("TC09","Klima ve Havalandirma Dengeleme"),("TC10","Medikal Gaz Test ve Sertifikasyon"),
        ("TC11","Asansor Test ve Sertifikasyon"),("TC12","BMS Entegrasyon Testi"),
        ("TC13","CCTV ve Guvenlik Test"),("TC14","HIS ve PACS Entegrasyon Testi"),
        ("TC15","Pnomatik Tup Sistemi Testi"),("TC16","Hemsire Cagri Sistemi Testi"),
        ("TC17","Ameliyathane Temiz Oda Validasyonu"),("TC18","Radyoloji Radyasyon Kacak Testi"),
        ("TC19","MRI Manyetik Alan Testi"),("TC20","LINAC Radyasyon Guvenlik Testi"),
        ("TC21","Ic Ortam Hava Kalitesi Testi"),("TC22","Akustik Test - Ameliyathane ve YBU"),
        ("TC23","Su Kalitesi ve Legionella Testi"),("TC24","Engelli Erisim Denetimi"),
        ("TC25","Itfaiye Onayi ve Yangin Tatbikati"),("TC26","Cevre ve Atik Yonetim Onayi"),
        ("TC27","Saglik Bakanligi Teknik Inceleme"),("TC28","Saglik Bakanligi Nihai Kabul"),
        ("TC29","Belediye ve Iskan Onayi"),("TC30","Peyzaj ve Cevre Duzenleme"),
        ("TC31","Otopark Cizgileme ve Yonlendirme"),("TC32","Tabela ve Yonlendirme Sistemi"),
        ("TC33","Gecici Santiye Sokumu"),("TC34","Hasta Kabul Simulasyonu"),
        ("TC35","Personel Egitimi"),("TC36","Gecici Kabul Tutanagi"),
    ]
    # WBS summaries
    WBS_NAMES = [
        ("PROJ","160K M2 SEHIR HASTANESI KOMPLEKSI"),
        ("WBS1","1. Dizayn, Ruhsat ve Medikal Planlama"),
        ("WBS2","2. Iksa, Hafriyat ve Zemin Iyilestirme"),
        ("WBS3","3. Alt Yapi ve Radyoaktif Yalitimli Temeller"),
        ("WBS4","4. Ust Yapi Karkas (Blok A-Yatakli, Blok B-Poliklinik, Blok C-DTC)"),
        ("WBS5","5. Dis Cephe ve Yalitim"),
        ("WBS6","6. MEP (Mekanik, Elektrik, Medikal Gaz ve Otomasyon)"),
        ("WBS7","7. Ince Isler ve Temiz Oda (Ameliyathane) Imalatlari"),
        ("WBS8","8. Medikal Cihaz Montaji (MRI, CT, Rontgen, vb.)"),
        ("WBS9","9. Test, Devreye Alma (Commissioning) ve Saglik Bakanligi Kabulu"),
    ]

    name_to_code = {}
    for code, name in ACTIVITIES_FLAT:
        name_to_code[name] = code
    for code, name in WBS_NAMES:
        name_to_code[name] = code

    # ── STEP 1: Single traversal — build code→bar_id map ──
    log("Step 1: Building bar_id map (single traversal)...")
    all_bars = {}

    def traverse(parent_task, depth=0):
        try:
            cbs = parent_task.ChildBars
            for i in range(1, cbs.Count + 1):
                cb = D(cbs.Item(i))
                name = cb.Name
                bid = cb.ID
                code = name_to_code.get(name)
                if code:
                    all_bars[code] = bid
                try:
                    ct = D(cb.Tasks(1))
                    traverse(ct, depth + 1)
                except: pass
        except: pass

    rb = D(project.Bars.Item(1))
    rt = D(rb.ExpandedTask)
    traverse(rt)
    log(f"  Found {len(all_bars)} bars")

    # Save bar_ids JSON
    mapping_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hospital_bar_ids.json")
    with open(mapping_file, "w") as mf:
        json.dump(all_bars, mf, indent=2)
    log(f"  Saved to {mapping_file}")

    # ── STEP 2: Check which links already exist ──
    log("\nStep 2: Checking existing links...")

    # Build existing link set by traversing once more
    existing_links = set()  # (from_bar_id, to_bar_id)

    def collect_links(parent_task):
        try:
            cbs = parent_task.ChildBars
            for i in range(1, cbs.Count + 1):
                cb = D(cbs.Item(i))
                try:
                    ct = D(cb.Tasks(1))
                    lo = ct.LinksOut
                    for li in range(1, lo.Count + 1):
                        link = D(lo.Item(li))
                        end_task = D(link.EndTask)
                        # Get bar IDs for start and end tasks
                        # The link connects tasks, find their parent bars
                        existing_links.add((cb.ID, end_task.ID))
                    collect_links(ct)
                except: pass
        except: pass

    rb = D(project.Bars.Item(1))
    rt = D(rb.ExpandedTask)
    collect_links(rt)
    log(f"  Found {len(existing_links)} existing links")

    # ── STEP 3: Define ALL links ──
    LINKS = [
        ("D01","D02",0,0),("D01","D03",1,10),("D02","D04",1,10),("D02","D05",1,5),
        ("D05","D06",0,0),("D03","D07",1,10),("D07","D08",0,0),("D08","D09",1,5),
        ("D09","D10",0,0),("D02","D11",0,0),("D03","D12",1,10),("D04","D13",1,10),
        ("D06","D14",0,0),("D10","D15",0,0),("D11","D15",0,0),("D12","D15",0,0),
        ("D13","D15",0,0),("D14","D15",0,0),
        ("D15","H01",0,0),
        ("H01","H02",0,0),("H02","H03",0,0),("H03","H04",1,5),("H03","H05",1,8),
        ("H03","H06",1,10),("H04","H07",1,10),("H05","H08",1,10),("H04","H09",1,10),
        ("H09","H10",0,0),("H04","H11",0,0),("H05","H12",0,0),("H06","H13",0,0),
        ("H10","H14",1,5),("H09","H15",1,5),("H07","H16",0,0),("H08","H17",0,0),
        ("H16","H18",0,0),("H17","H18",1,5),("H18","H19",0,0),("H19","H20",0,0),
        ("H11","H20",0,0),("H12","H20",0,0),("H13","H20",0,0),("H14","H20",0,0),
        ("H20","T01",0,0),("H20","T02",1,3),("H20","T03",1,5),
        ("T01","T04",0,0),("T02","T05",0,0),("T03","T06",0,0),("T04","T07",0,0),
        ("T05","T09",0,0),("T06","T11",0,0),("T07","T08",0,0),("T09","T10",0,0),
        ("T11","T12",0,0),("T08","T13",0,0),("T13","T14",0,0),("T14","T15",0,0),
        ("T08","T16",0,0),("T10","T17",0,0),("T12","T18",0,0),("T16","T19",0,0),
        ("T17","T20",0,0),("T18","T21",0,0),("T19","T22",1,5),("T19","T23",1,5),
        ("T19","T24",1,5),("T19","T25",0,0),("T20","T25",0,0),("T21","T25",0,0),
        ("T22","T25",0,0),("T23","T25",0,0),("T24","T25",0,0),("T15","T25",0,0),
        ("T19","K01",0,0),("T20","K14",0,0),("T21","K22",0,0),
        ("K01","K02",0,0),("K02","K03",0,0),("K03","K04",0,0),("K04","K05",0,0),
        ("K05","K06",0,0),("K06","K07",0,0),("K07","K08",0,0),("K08","K09",0,0),
        ("K09","K10",0,0),("K10","K11",0,0),("K11","K12",0,0),("K12","K13",0,0),
        ("K14","K15",0,0),("K15","K16",0,0),("K16","K17",0,0),("K17","K18",0,0),
        ("K18","K19",0,0),("K19","K20",0,0),("K20","K21",0,0),
        ("K22","K23",0,0),("K23","K24",0,0),("K24","K25",0,0),("K25","K26",0,0),
        ("K26","K27",0,0),("K27","K28",0,0),
        ("K13","K29",1,5),("K21","K29",1,5),("K13","K30",0,0),
        ("K02","K31",1,5),("K15","K32",1,5),("K23","K33",1,5),
        ("K13","K34",0,0),("K21","K35",0,0),("K28","K36",0,0),
        ("K34","K37",0,0),("K35","K37",0,0),("K36","K37",0,0),("K29","K37",0,0),("K30","K37",0,0),
        ("K13","C01",0,0),("K21","C02",0,0),("K28","C03",0,0),
        ("C01","C04",0,0),("C02","C05",0,0),("C03","C06",0,0),
        ("C04","C07",0,0),("C05","C08",0,0),("C06","C09",0,0),
        ("C04","C10",1,10),("C05","C11",1,5),
        ("K34","C12",0,0),("K35","C13",0,0),("K36","C14",0,0),
        ("C12","C15",0,0),("C13","C15",1,3),("C14","C15",1,5),
        ("C07","C16",0,0),("C08","C17",0,0),("C09","C18",0,0),
        ("C04","C19",1,10),("C07","C20",1,5),
        ("C16","C21",0,0),("C17","C21",1,3),("C18","C21",1,5),
        ("C21","C22",0,0),("C19","C22",0,0),
        ("K10","M05",1,5),("K19","M06",1,5),("K27","M07",1,5),
        ("K34","M01",0,0),("K31","M29",0,0),("K32","M30",0,0),("K33","M31",0,0),
        ("C07","M08",0,0),("C08","M09",0,0),("C09","M10",0,0),
        ("M01","M02",0,0),("M02","M03",0,0),("M01","M04",1,5),("M04","M05",0,0),
        ("M05","M08",1,10),("M06","M09",1,10),("M07","M10",1,10),
        ("M01","M11",1,5),("M04","M12",0,0),("M12","M13",0,0),("M12","M14",1,5),
        ("M12","M15",1,8),("M01","M16",1,5),("M16","M17",0,0),("M17","M18",0,0),
        ("M17","M19",1,5),("M17","M20",1,8),("M18","M21",0,0),("M19","M22",0,0),
        ("M20","M23",0,0),("M16","M24",1,5),("M24","M25",0,0),("M24","M26",1,5),
        ("M24","M27",1,8),("M25","M28",1,10),("M04","M32",1,10),("M11","M33",1,5),
        ("M32","M34",1,5),("M04","M35",1,5),("M11","M36",0,0),("M18","M37",1,5),
        ("M16","M38",1,10),("M21","M39",1,5),("M36","M40",0,0),("M28","M40",0,0),
        ("M39","M40",0,0),
        ("M21","I10",0,0),("M22","I11",0,0),("M23","I12",0,0),
        ("M25","I13",0,0),("M26","I14",0,0),("M27","I15",0,0),
        ("K10","I01",1,20),("K19","I02",1,10),("K27","I03",1,5),
        ("I01","I04",0,0),("I02","I05",0,0),("I03","I06",0,0),
        ("I04","I07",0,0),("I05","I08",0,0),("I06","I09",0,0),
        ("I10","I13",0,0),("I11","I14",0,0),("I12","I15",0,0),
        ("I13","I16",0,0),("I15","I17",0,0),("I16","I18",0,0),
        ("I18","I19",0,0),("I19","I20",0,0),("I20","I21",0,0),
        ("I07","I22",0,0),("I08","I23",0,0),("I09","I24",0,0),
        ("I22","I25",0,0),("I23","I26",0,0),("I24","I27",0,0),
        ("I18","I28",0,0),("I28","I29",0,0),
        ("I25","I30",0,0),("I26","I30",0,0),("I27","I30",0,0),("I21","I30",0,0),("I29","I30",0,0),
        ("I19","MC01",0,0),("I28","MC02",1,5),("I21","MC09",0,0),("I30","MC17",0,0),
        ("MC01","MC02",0,0),("MC01","MC03",1,5),("MC02","MC04",1,5),("MC03","MC04",1,3),
        ("MC04","MC05",1,5),("MC04","MC06",1,3),("MC06","MC07",1,3),("MC07","MC08",1,3),
        ("MC09","MC10",0,0),("MC10","MC11",1,5),("MC11","MC12",1,5),("MC12","MC13",1,3),
        ("MC13","MC14",1,5),("MC14","MC15",0,0),("MC15","MC16",0,0),
        ("MC17","MC18",1,5),("MC18","MC19",1,3),("MC19","MC20",1,3),
        ("I28","MC21",0,0),("MC21","MC22",1,5),("MC22","MC23",1,3),("MC23","MC24",1,3),
        ("MC16","MC25",0,0),("MC24","MC25",0,0),("MC20","MC25",0,0),
        ("MC25","TC01",0,0),
        ("TC01","TC02",0,0),("TC02","TC03",0,0),("TC01","TC04",1,3),("TC01","TC05",1,3),
        ("TC01","TC06",1,3),("TC06","TC07",0,0),("TC07","TC08",0,0),("TC08","TC09",0,0),
        ("TC09","TC10",0,0),("TC10","TC11",1,3),("TC11","TC12",0,0),("TC12","TC13",1,3),
        ("TC13","TC14",0,0),("TC14","TC15",1,3),("TC15","TC16",1,3),("TC16","TC17",0,0),
        ("TC17","TC18",0,0),("TC18","TC19",0,0),("TC19","TC20",0,0),
        ("TC09","TC21",0,0),("TC17","TC22",0,0),("TC21","TC23",0,0),("TC22","TC24",1,3),
        ("TC24","TC25",0,0),("TC23","TC26",1,3),("TC25","TC27",0,0),("TC26","TC27",1,3),
        ("TC27","TC28",0,0),("TC28","TC29",1,5),
        ("C22","TC30",1,10),("TC30","TC31",0,0),("TC31","TC32",0,0),
        ("TC28","TC33",0,0),("TC33","TC34",0,0),("TC34","TC35",0,0),("TC35","TC36",0,0),
        ("TC29","TC36",0,0),("TC32","TC36",0,0),
    ]

    # ── STEP 4: Create missing links (batch 10 per transaction) ──
    log(f"\nStep 3: Creating missing links (total defined: {len(LINKS)})...")

    # Filter to missing links only
    missing = []
    for fc, tc, lt, lag in LINKS:
        if fc not in all_bars or tc not in all_bars:
            continue
        # Can't efficiently check existing (task ID != bar ID), just try all
        missing.append((fc, tc, lt, lag))

    log(f"  Links to attempt: {len(missing)}")

    BATCH = 10
    link_ok = 0
    link_dup = 0
    link_fail = 0

    for batch_start in range(0, len(missing), BATCH):
        batch = missing[batch_start:batch_start + BATCH]

        # Fresh refs for this batch
        project.StartTransaction(f"Links-{batch_start}")

        # Build fresh task cache for this batch
        task_cache = {}
        rb = D(project.Bars.Item(1))
        rt = D(rb.ExpandedTask)

        def cache_tasks(parent_t):
            try:
                cbs = parent_t.ChildBars
                for i in range(1, cbs.Count + 1):
                    cb = D(cbs.Item(i))
                    name = cb.Name
                    code = name_to_code.get(name)
                    if code:
                        try:
                            task_cache[code] = D(cb.Tasks(1))
                        except: pass
                    try:
                        ct = D(cb.Tasks(1))
                        cache_tasks(ct)
                    except: pass
            except: pass

        cache_tasks(rt)

        for fc, tc, lt, lag in batch:
            t1 = task_cache.get(fc)
            t2 = task_cache.get(tc)
            if not t1 or not t2:
                link_fail += 1
                continue
            try:
                link = D(t1.LinkTo(t2))
                if lt != 0:
                    link.Type = lt
                if lag > 0:
                    lag_dur = t1.GetDurationFromString(f"{lag}d")
                    link.StartLagTime = lag_dur
                link_ok += 1
            except Exception as e:
                err_str = str(e)
                if "already" in err_str.lower() or "exists" in err_str.lower():
                    link_dup += 1
                else:
                    link_fail += 1

        try:
            project.EndTransaction()
        except:
            try: project.AbandonTransaction()
            except: pass
        project.WaitForNotificationProcessing()

        if (batch_start + BATCH) % 50 < BATCH:
            log(f"    Processed {min(batch_start + BATCH, len(missing))}/{len(missing)} (ok={link_ok}, dup={link_dup}, fail={link_fail})")

    log(f"  DONE: OK={link_ok}, Duplicates={link_dup}, Failed={link_fail}")

    # ── STEP 5: Reschedule and save ──
    log("\nStep 4: Reschedule and save...")
    project.Reschedule()
    project.Save()
    log("  DONE!")

    log(f"\n{'='*60}")
    log(f"SUMMARY: {len(all_bars)} bars mapped, {link_ok} new links, {link_dup} duplicates")
    log(f"{'='*60}")

except Exception as e:
    log(f"FATAL: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output: {OUT}")
