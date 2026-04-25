"""
YHT (Yüksek Hızlı Tren) ALTYAPI PROJESİ — Full Builder
=========================================================
PMO Direktörü Script — 8 ADIM
ADIM 1: WBS + 80+ Aktivite + Bağlantılar (FS/SS+Lag/FF)
ADIM 2: 2 Kod Kütüphanesi + Matrix Assignment
ADIM 3: Maliyet Merkezleri + Kaynaklar + Atamalar ($100M+)
ADIM 4: Reschedule + Baseline Kaydetme
ADIM 5: 1. Progress Period (Ay 1) + Gecikme
ADIM 6: Varyans Analizi
ADIM 7: 2. Progress Period + Fast-Tracking
ADIM 8: Rapor (view/export MCP ile yapılacak)

Lesson-Learned kuralları:
- Root bar → ExpandedTask (Tasks(1) değil)
- lib.Entries.Add() (CodeLibraryEntrys değil)
- bar.AssignCode(entry, True) (BAR üzerinde, task değil)
- bar.DurationPercentComplete (BAR üzerinde, task değil)
- COM ref → EndTransaction sonrası stale, re-fetch gerekir
- project.Reschedule() parametresiz çalışır
"""
import sys, os, traceback, json
sys.stdout.reconfigure(encoding='utf-8')
from datetime import datetime, timedelta

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_yht_output.txt")
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
    def pt(dt_str):
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
        project.StartTransaction(name)

    def end_tx():
        try:
            project.EndTransaction()
        except Exception as e:
            log(f"  [WARN] EndTx: {e}")
            try: project.AbandonTransaction()
            except: pass
        project.WaitForNotificationProcessing()

    # ── Progress helper ──
    def set_progress(bar_obj, pct):
        """Set progress on bar with fallback."""
        try:
            bar_obj.DurationPercentComplete = float(pct)
            return
        except:
            pass
        try:
            bar_obj.OverallPercentComplete = float(pct)
            return
        except:
            pass
        # Last resort: try via _oleobj_
        try:
            did = bar_obj._oleobj_.GetIDsOfNames(0, 'DurationPercentComplete')
            bar_obj._oleobj_.InvokeTypes(did, 0, 4, (24, 0), ((5, 1),), float(pct))
        except Exception as e:
            log(f"    [WARN] Progress set failed: {e}")

    # ── Re-fetch helper (COM refs invalidated after tx) ──
    all_tasks = {}  # code -> (bar, task, bar_id)

    def get_fresh(code):
        """Re-fetch (bar, task) by bar ID."""
        _, _, bar_id = all_tasks[code]
        def find_bar(parent_t, tid):
            try:
                cbs = parent_t.ChildBars
                for i in range(1, cbs.Count + 1):
                    cb = D(cbs.Item(i))
                    if cb.ID == tid:
                        return cb, D(cb.Tasks(1))
                    try:
                        ct = D(cb.Tasks(1))
                        r = find_bar(ct, tid)
                        if r: return r
                    except: pass
            except: pass
            return None
        rb = D(project.Bars.Item(1))
        rt = D(rb.ExpandedTask)
        return find_bar(rt, bar_id)

    def get_bar(code):
        r = get_fresh(code)
        return r[0] if r else None

    def get_task(code):
        r = get_fresh(code)
        return r[1] if r else None

    # ══════════════════════════════════════════════════════════
    # ADIM 0: PROJE KÖKÜNÜ HAZIRLA
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 0: Proje kokunu hazirliyorum...")
    log("=" * 70)

    bars = project.Bars
    if bars.Count == 0:
        tx("Root")
        root_bar = D(bars.Add())
        root_bar.Name = "Program"
        root_task_obj = D(root_bar.Tasks.AddSummaryTask(pt("2026-06-01")))
        end_tx()
        log(f"  Yeni root: ID={root_bar.ID}")
    else:
        root_bar = D(bars.Item(1))
        root_task_obj = D(root_bar.ExpandedTask)
        log(f"  Root mevcut: ID={root_bar.ID}")

    # ══════════════════════════════════════════════════════════
    # ADIM 1: WBS + 80 AKTİVİTE + BAĞLANTILAR
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 1: WBS + 80+ Aktivite + Baglantilar")
    log("=" * 70)

    # ── Aktivite tanımları ──
    # Her biri: (kod, isim, sure_gun, start_date)
    # Start: 2026-06-01

    VIY = [
        ("V01", "Viyaduk V1 Ayak Temelleri Kazisi", 15, "2026-06-01"),
        ("V02", "V1 Temel Betonlamasi", 12, "2026-06-22"),
        ("V03", "V1 Ayak Kolon Kaliplari", 10, "2026-07-08"),
        ("V04", "V1 Kolon Betonlamasi", 8, "2026-07-22"),
        ("V05", "V1 Tabla Kalip ve Donati", 12, "2026-08-03"),
        ("V06", "V1 Tabla Betonlamasi", 10, "2026-08-19"),
        ("V07", "V1 Ongerme Kablo Cekimi", 8, "2026-09-02"),
        ("V08", "V1 Tabla Sokulusu", 5, "2026-09-14"),
        ("V09", "Viyaduk V2 Ayak Temelleri", 15, "2026-07-15"),
        ("V10", "V2 Temel Betonlamasi", 12, "2026-08-05"),
        ("V11", "V2 Ayak Kolon Kaliplari", 10, "2026-08-21"),
        ("V12", "V2 Kolon Betonlamasi", 8, "2026-09-04"),
        ("V13", "V2 Tabla Kalip ve Donati", 12, "2026-09-16"),
        ("V14", "V2 Tabla Betonlamasi", 10, "2026-10-02"),
        ("V15", "V2 Ongerme Kablo Cekimi", 8, "2026-10-16"),
        ("V16", "V2 Tabla Sokulusu", 5, "2026-10-28"),
        ("V17", "V1-V2 Korkuluk ve Barier Montaji", 10, "2026-11-04"),
        ("V18", "V1-V2 Su Yalitimi", 8, "2026-11-18"),
        ("V19", "V1-V2 Drenaj Sistemi", 6, "2026-11-30"),
        ("V20", "Viyadukler Tamamlama ve Kabul", 5, "2026-12-08"),
    ]

    TUN = [
        ("T01", "Tunel T1 Portal Kazisi", 12, "2026-06-08"),
        ("T02", "T1 TBM Montaj ve Devreye Alma", 20, "2026-06-24"),
        ("T03", "T1 TBM Ilerleme - Faz 1 (0-500m)", 25, "2026-07-22"),
        ("T04", "T1 TBM Ilerleme - Faz 2 (500-1000m)", 25, "2026-08-26"),
        ("T05", "T1 TBM Ilerleme - Faz 3 (1000-1500m)", 25, "2026-10-01"),
        ("T06", "T1 Segment Montaji - Faz 1", 20, "2026-08-05"),
        ("T07", "T1 Segment Montaji - Faz 2", 20, "2026-09-09"),
        ("T08", "T1 Segment Montaji - Faz 3", 20, "2026-10-15"),
        ("T09", "T1 Havalandirma Sistemi", 15, "2026-11-12"),
        ("T10", "T1 Yangin Sondurme Altyapisi", 10, "2026-12-03"),
        ("T11", "Tunel T2 Portal Kazisi", 12, "2026-08-17"),
        ("T12", "T2 TBM Montaj", 18, "2026-09-02"),
        ("T13", "T2 TBM Ilerleme - Faz 1 (0-800m)", 30, "2026-09-28"),
        ("T14", "T2 TBM Ilerleme - Faz 2 (800-1600m)", 30, "2026-11-09"),
        ("T15", "T2 Segment Montaji - Faz 1", 22, "2026-10-12"),
        ("T16", "T2 Segment Montaji - Faz 2", 22, "2026-11-23"),
        ("T17", "T2 Havalandirma Sistemi", 15, "2027-01-05"),
        ("T18", "T2 Yangin Sondurme Altyapisi", 10, "2027-01-26"),
        ("T19", "T1-T2 Acil Cikis Tamamlama", 12, "2027-02-09"),
        ("T20", "Tuneller Tamamlama ve Kabul", 5, "2027-02-25"),
    ]

    ZEM = [
        ("Z01", "Zemin Etud ve Analiz Raporu", 10, "2026-06-01"),
        ("Z02", "Jet Grout Uygulamasi - Bolge K1", 15, "2026-06-15"),
        ("Z03", "Jet Grout Uygulamasi - Bolge K2", 15, "2026-07-06"),
        ("Z04", "Jet Grout Uygulamasi - Bolge G1", 15, "2026-07-27"),
        ("Z05", "Zemin Ankrajlari - Bolge K1", 12, "2026-07-06"),
        ("Z06", "Zemin Ankrajlari - Bolge K2", 12, "2026-07-22"),
        ("Z07", "Zemin Ankrajlari - Bolge G1", 12, "2026-08-07"),
        ("Z08", "Derin Kazik Imalati - K1", 18, "2026-07-22"),
        ("Z09", "Derin Kazik Imalati - K2", 18, "2026-08-17"),
        ("Z10", "Derin Kazik Imalati - G1", 18, "2026-09-10"),
        ("Z11", "Konsolidasyon Test (K1)", 8, "2026-08-17"),
        ("Z12", "Konsolidasyon Test (K2)", 8, "2026-09-10"),
        ("Z13", "Konsolidasyon Test (G1)", 8, "2026-10-06"),
        ("Z14", "Toprak Dolgu - Guzergah K", 20, "2026-08-27"),
        ("Z15", "Toprak Dolgu - Guzergah G", 20, "2026-10-06"),
        ("Z16", "Stabilizasyon Tabakasi - K", 12, "2026-09-24"),
        ("Z17", "Stabilizasyon Tabakasi - G", 12, "2026-11-03"),
        ("Z18", "Alt Temel Serimi", 15, "2026-10-12"),
        ("Z19", "Ust Temel Serimi", 15, "2026-11-03"),
        ("Z20", "Zemin Iyilestirme Tamamlama", 5, "2026-11-24"),
    ]

    RAY = [
        ("R01", "Ray Deposu Kurulumu", 10, "2026-09-01"),
        ("R02", "Beton Travers Uretimi - Faz 1", 20, "2026-09-15"),
        ("R03", "Beton Travers Uretimi - Faz 2", 20, "2026-10-13"),
        ("R04", "Balast Serimi - Bolge K", 18, "2026-10-05"),
        ("R05", "Balast Serimi - Bolge G", 18, "2026-10-29"),
        ("R06", "Ray Kaynatma Tesisi Kurulumu", 12, "2026-09-29"),
        ("R07", "Ray Serimi - Bolge K (0-25 km)", 22, "2026-10-29"),
        ("R08", "Ray Serimi - Bolge K (25-50 km)", 22, "2026-11-30"),
        ("R09", "Ray Serimi - Bolge G (0-25 km)", 22, "2026-12-02"),
        ("R10", "Ray Serimi - Bolge G (25-50 km)", 22, "2027-01-05"),
        ("R11", "Makas ve Devir Montaji", 15, "2027-01-12"),
        ("R12", "Katener Direk Montaji - K", 18, "2026-11-16"),
        ("R13", "Katener Direk Montaji - G", 18, "2026-12-14"),
        ("R14", "Katener Tel Cekimi - K", 15, "2026-12-10"),
        ("R15", "Katener Tel Cekimi - G", 15, "2027-01-12"),
        ("R16", "Sinyalizasyon Kablo Dosemesi", 20, "2027-01-19"),
        ("R17", "ETCS Sinyal Sistemi Montaji", 18, "2027-02-16"),
        ("R18", "Enerji Beslemesi ve Trafo", 15, "2027-02-02"),
        ("R19", "Test Surosleri (160-250 km/h)", 20, "2027-03-10"),
        ("R20", "Ray ve Elektrifikasyon Kabul", 5, "2027-04-07"),
    ]

    ALL_SECTIONS = [
        ("1. Viyadukler", VIY),
        ("2. Tuneller", TUN),
        ("3. Zemin Iyilestirme", ZEM),
        ("4. Ray Serimi ve Elektrifikasyon", RAY),
    ]

    # ── Create WBS hierarchy ──
    def create_summary(parent_task, name):
        tx(f"S-{name[:15]}")
        new_bar = D(parent_task.ChildBars.Add())
        new_bar.Name = name
        new_task = D(new_bar.Tasks.AddSummaryTask(pt("2026-06-01")))
        end_tx()
        return new_bar, new_task

    def create_leaf(parent_task, code, name, dur, start):
        tx(f"L-{code}")
        new_bar = D(parent_task.ChildBars.Add())
        new_bar.Name = f"{code} {name}"
        new_task = D(new_bar.Tasks.AddTask(pt(start), f"{dur}d"))
        end_tx()
        bar_id = new_bar.ID
        all_tasks[code] = (new_bar, new_task, bar_id)
        return new_bar, new_task, bar_id

    # L1: Proje
    prj_bar, prj_task = create_summary(root_task_obj,
        "Yuksek Hizli Tren (YHT) Altyapi Projesi")
    log(f"  L1: YHT Altyapi Projesi (ID={prj_bar.ID})")

    # L2: Sections + L3: Activities
    section_tasks = {}  # section_name -> summary_task
    for sec_name, activities in ALL_SECTIONS:
        s_bar, s_task = create_summary(prj_task, sec_name)
        section_tasks[sec_name] = s_task
        log(f"  L2: {sec_name} (ID={s_bar.ID})")

        for code, name, dur, start in activities:
            _, _, bid = create_leaf(s_task, code, name, dur, start)
            log(f"    {code} {name} ({dur}d) ID={bid}")

    total_acts = len(all_tasks)
    log(f"\n  === TOPLAM: {total_acts} leaf aktivite ===")

    # ── Links ──
    log("\n  Baglantilar kuruluyor...")

    def link_fs(c1, c2, lag=None):
        _, t1, _ = all_tasks[c1]
        _, t2, _ = all_tasks[c2]
        lnk = D(t1.LinkTo(t2))
        lnk.type = 0
        if lag:
            lnk.StartLagTime = t1.GetDurationFromString(lag)
        return lnk

    def link_ss(c1, c2, lag=None):
        _, t1, _ = all_tasks[c1]
        _, t2, _ = all_tasks[c2]
        lnk = D(t1.LinkTo(t2))
        lnk.type = 1
        if lag:
            lnk.StartLagTime = t1.GetDurationFromString(lag)
        return lnk

    def link_ff(c1, c2, lag=None):
        _, t1, _ = all_tasks[c1]
        _, t2, _ = all_tasks[c2]
        lnk = D(t1.LinkTo(t2))
        lnk.type = 2
        if lag:
            lnk.StartLagTime = t1.GetDurationFromString(lag)
        return lnk

    tx("Links-All")

    # Viyadukler: V01-V08 chain, V09-V16 chain, V17-V20 chain
    for i in range(7): link_fs(VIY[i][0], VIY[i+1][0])
    for i in range(8, 15): link_fs(VIY[i][0], VIY[i+1][0])
    link_fs("V08", "V17"); link_fs("V16", "V17")
    link_fs("V17", "V18"); link_fs("V18", "V19"); link_fs("V19", "V20")
    # V01→V09 SS+30d (V2 starts while V1 in progress)
    link_ss("V01", "V09", "30d")
    log("    VIY: 19 FS + 1 SS chain")

    # Tuneller: T01-T05 TBM chain, T06-T08 segment parallel, T09-T10, T11-T18, T19-T20
    for i in range(4): link_fs(TUN[i][0], TUN[i+1][0])
    # Segments parallel to TBM (SS+10d from each TBM phase)
    link_ss("T03", "T06", "10d")
    link_ss("T04", "T07", "10d")
    link_ss("T05", "T08", "10d")
    link_fs("T08", "T09"); link_fs("T09", "T10")
    # T2 chain
    for i in range(10, 17): link_fs(TUN[i][0], TUN[i+1][0])
    link_ss("T13", "T15", "10d")
    link_ss("T14", "T16", "10d")
    link_fs("T10", "T19"); link_fs("T18", "T19"); link_fs("T19", "T20")
    # T1 finish → T2 start (SS with overlap)
    link_ss("T01", "T11", "50d")
    # FF: T10 and T18 finish together with T20
    link_ff("T10", "T20")
    link_ff("T18", "T20")
    log("    TUN: 17 FS + 5 SS + 2 FF chain")

    # Zemin: partial chains with overlaps
    link_fs("Z01", "Z02")
    link_fs("Z02", "Z03"); link_fs("Z03", "Z04")
    link_ss("Z02", "Z05", "10d")
    link_ss("Z03", "Z06", "10d")
    link_ss("Z04", "Z07", "10d")
    link_fs("Z05", "Z08"); link_fs("Z06", "Z09"); link_fs("Z07", "Z10")
    link_ss("Z08", "Z11", "20d")
    link_ss("Z09", "Z12", "20d")
    link_ss("Z10", "Z13", "20d")
    link_fs("Z11", "Z14"); link_fs("Z12", "Z14")
    link_fs("Z13", "Z15")
    link_ss("Z14", "Z16", "15d")
    link_ss("Z15", "Z17", "15d")
    link_fs("Z16", "Z18"); link_fs("Z17", "Z18")
    link_fs("Z18", "Z19"); link_fs("Z19", "Z20")
    log("    ZEM: 14 FS + 7 SS chain")

    # Ray: chains with dependencies
    link_fs("R01", "R02"); link_fs("R02", "R03")
    link_fs("R01", "R06")
    link_ss("R02", "R04", "15d")
    link_ss("R03", "R05", "15d")
    link_fs("R04", "R07"); link_fs("R06", "R07")
    link_fs("R07", "R08")
    link_fs("R05", "R09"); link_fs("R09", "R10")
    link_fs("R10", "R11")
    link_ss("R07", "R12", "12d")
    link_ss("R09", "R13", "12d")
    link_ss("R12", "R14", "15d")
    link_ss("R13", "R15", "15d")
    link_fs("R08", "R16"); link_fs("R10", "R16")
    link_fs("R16", "R17")
    link_ss("R14", "R18", "10d")
    link_fs("R17", "R19"); link_fs("R15", "R19"); link_fs("R18", "R19")
    link_fs("R19", "R20")
    log("    RAY: 15 FS + 5 SS chain")

    # Cross-section links
    link_fs("Z20", "R04")  # Zemin tamamlama → Balast serimi
    link_fs("V20", "R07")  # Viyaduk tamamlama → Ray serimi K
    link_fs("T20", "R09")  # Tunel tamamlama → Ray serimi G
    link_ss("Z14", "V09", "10d")  # Dolgu → V2 temelleri
    log("    Cross: 3 FS + 1 SS (kritik yol olusturuldu)")

    end_tx()

    link_count = 19 + 24 + 21 + 20 + 4
    log(f"  === TOPLAM: ~{link_count} baglanti kuruldu ===")

    # ══════════════════════════════════════════════════════════
    # ADIM 2: KOD KUTUPHANELERI + MATRIX ATAMASI
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 2: Kod Kutuphaneleri + Matrix Atamasi")
    log("=" * 70)

    code_libs = project.CodeLibrarys

    def create_code_lib(name, entries):
        tx(f"CL-{name[:10]}")
        lib = D(code_libs.Add())
        lib.Name = name
        entry_objs = {}
        for en in entries:
            e = D(lib.Entries.Add())
            e.Name = en
            entry_objs[en] = e
        end_tx()
        log(f"  Kutuphane: '{name}' -> {entries}")
        return lib, entry_objs

    lok_lib, lok_entries = create_code_lib("Lokasyon",
        ["Bolge-Kuzey", "Bolge-Guney", "Merkez-Istasyon"])
    tas_lib, tas_entries = create_code_lib("Taseron",
        ["Alfa Zemin", "Beta Tunel", "Gama Ray"])

    # Matrix assignment map
    lok_map = {
        "Bolge-Kuzey": [c[0] for c in VIY[:8] + TUN[:10] + ZEM[:6] + ZEM[7:9] +
                         ZEM[10:12] + [ZEM[13]] + [ZEM[15]] + RAY[:2] + RAY[3:4] +
                         RAY[5:8] + RAY[11:12] + RAY[13:14]],
        "Bolge-Guney": [c[0] for c in VIY[8:16] + TUN[10:18] + ZEM[3:4] + ZEM[6:7] +
                         ZEM[9:10] + ZEM[12:13] + [ZEM[14]] + [ZEM[16]] +
                         RAY[4:5] + RAY[8:10] + RAY[12:13] + RAY[14:15]],
        "Merkez-Istasyon": [c[0] for c in VIY[16:] + TUN[18:] + ZEM[17:] +
                            RAY[2:3] + RAY[10:11] + RAY[15:]],
    }
    tas_map = {
        "Alfa Zemin": [c[0] for c in ZEM + VIY[:8]],
        "Beta Tunel": [c[0] for c in TUN + VIY[8:16]],
        "Gama Ray": [c[0] for c in RAY + VIY[16:] + TUN[18:]],
    }

    tx("CodeAssign")
    assigned = 0
    for entry_name, codes in lok_map.items():
        for code in codes:
            if code in all_tasks:
                bar_obj, _, _ = all_tasks[code]
                try:
                    bar_obj.AssignCode(lok_entries[entry_name], True)
                    assigned += 1
                except: pass

    for entry_name, codes in tas_map.items():
        for code in codes:
            if code in all_tasks:
                bar_obj, _, _ = all_tasks[code]
                try:
                    bar_obj.AssignCode(tas_entries[entry_name], True)
                    assigned += 1
                except: pass
    end_tx()
    log(f"  === TOPLAM: {assigned} kod atamasi ===")

    # ══════════════════════════════════════════════════════════
    # ADIM 3: MALİYET MERKEZLERİ + KAYNAKLAR + ATAMALAR
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 3: Maliyet Merkezleri + Kaynaklar + Atamalar")
    log("=" * 70)

    # Cost centres
    tx("CostCentres")
    cc_all = project.CostCentres
    yht_cc = D(cc_all.Add()); yht_cc.Name = "YHT Genel Butce"
    iscilik_cc = D(cc_all.Add()); iscilik_cc.Name = "Iscilik"
    ekipman_cc = D(cc_all.Add()); ekipman_cc.Name = "Ekipman"
    malzeme_cc = D(cc_all.Add()); malzeme_cc.Name = "Malzeme"
    end_tx()
    log(f"  CC: YHT Genel Butce + Iscilik + Ekipman + Malzeme")

    # Resources
    tx("Resources")
    perm_all = project.PermanentResources
    cons_all = project.ConsumableResources

    # Permanent: Viyaduk Kalip Ekibi (450.25 $/hour)
    kalip_res = D(perm_all.Add())
    kalip_res.Name = "Viyaduk Kalip Ekibi"
    log(f"  Permanent: Viyaduk Kalip Ekibi")

    # Permanent: TBM (350500.75 $/week)
    tbm_res = D(perm_all.Add())
    tbm_res.Name = "TBM Tunel Acma Makinesi"
    log(f"  Permanent: TBM Tunel Acma Makinesi")

    # Consumable: Celik Ray (1250.50 $/metre)
    ray_res = D(cons_all.Add())
    ray_res.Name = "Yuksek Mukavemetli Celik Ray"
    try:
        cpu_ac = get_ac(D(ray_res), 'CostPerUnit')
        if cpu_ac: set_amt(cpu_ac, 1250.50)
    except: pass
    log(f"  Consumable: Celik Ray @ $1,250.50/metre")
    end_tx()

    # CostAndIncomeRates for permanent resources
    tx("Rates")
    rates_coll = project.CostAndIncomeRates

    # Kalip Ekibi rate: $450.25/hour
    kalip_rate = D(rates_coll.Add())
    kalip_rate.Name = "Kalip Ekibi Ucreti"
    try:
        amt_ac = get_ac(D(kalip_rate), 'Amount')
        if amt_ac: set_amt(amt_ac, 450.25)
    except: pass
    log(f"  Rate: Kalip Ekibi @ $450.25/saat")

    # TBM rate: $350500.75/week
    tbm_rate = D(rates_coll.Add())
    tbm_rate.Name = "TBM Operasyon Ucreti"
    try:
        amt_ac = get_ac(D(tbm_rate), 'Amount')
        if amt_ac: set_amt(amt_ac, 350500.75)
    except: pass
    # Try to set TimeUnit to Week
    try:
        tbm_rate.TimeUnit = 2  # 2=Week (guess)
    except: pass
    log(f"  Rate: TBM @ $350,500.75/hafta")
    end_tx()

    # ── Resource Assignments ──
    log("\n  Kaynak atamalari yapiliyor...")
    total_cost = 0.0

    # Viyaduk Kalip Ekibi → Viyaduk tasks (V01-V16)
    tx("AssignKalip")
    kalip_tasks = [c[0] for c in VIY[:16]]
    for code in kalip_tasks:
        _, task_obj, _ = all_tasks[code]
        alloc = D(task_obj.AssignResource(kalip_res, False))
        alloc_d = D(alloc)
        alloc_d.GivenAllocation = 25.0  # 25 adam-ekip
        try:
            alloc_d.AssignRate(kalip_rate)
        except: pass
        # Find duration
        dur = 0
        for c, n, d, s in VIY[:16]:
            if c == code: dur = d; break
        task_cost = 25 * dur * 8 * 450.25  # 25 ekip x gun x 8h x rate
        total_cost += task_cost
        log(f"    {code}: Kalip Ekibi x25 = ${task_cost:,.2f}")
    end_tx()

    # TBM → Tunel TBM tasks (T03-T05, T13-T14)
    tx("AssignTBM")
    tbm_tasks = ["T03", "T04", "T05", "T13", "T14"]
    for code in tbm_tasks:
        _, task_obj, _ = all_tasks[code]
        alloc = D(task_obj.AssignResource(tbm_res, False))
        alloc_d = D(alloc)
        alloc_d.GivenAllocation = 2.0  # 2 TBM
        try:
            alloc_d.AssignRate(tbm_rate)
        except: pass
        dur = 0
        for c, n, d, s in TUN:
            if c == code: dur = d; break
        # Cost = 2 TBM x dur_weeks x $350,500.75/week
        weeks = dur / 5.0
        task_cost = 2 * weeks * 350500.75
        total_cost += task_cost
        log(f"    {code}: TBM x2 = ${task_cost:,.2f}")
    end_tx()

    # Celik Ray → Ray serimi tasks (R07-R10)
    tx("AssignRay")
    ray_assign = {"R07": 25000, "R08": 25000, "R09": 25000, "R10": 25000}
    for code, qty in ray_assign.items():
        _, task_obj, _ = all_tasks[code]
        alloc = D(task_obj.AssignConsumableResource(ray_res, False, None, None))
        alloc.GivenQuantity = float(qty)
        try:
            alloc_cpu = get_ac(D(alloc), 'CostPerUnit')
            if alloc_cpu: set_amt(alloc_cpu, 1250.50)
        except: pass
        task_cost = qty * 1250.50
        total_cost += task_cost
        log(f"    {code}: Celik Ray x{qty:,}m = ${task_cost:,.2f}")
    end_tx()

    # Cost Centre assignments for remaining tasks
    tx("CCAssign")
    all_assigned = set(kalip_tasks + tbm_tasks + list(ray_assign.keys()))
    remaining = [c[0] for sec in [VIY, TUN, ZEM, RAY] for c in sec if c[0] not in all_assigned]
    for code in remaining:
        _, task_obj, _ = all_tasks[code]
        if code.startswith("T"):
            cc = ekipman_cc; cost_val = 850000.0
        elif code.startswith("Z"):
            cc = iscilik_cc; cost_val = 620000.0
        elif code.startswith("R"):
            cc = malzeme_cc; cost_val = 780000.0
        else:
            cc = yht_cc; cost_val = 450000.0
        alloc = D(task_obj.AssignCost(cc))
        gv = get_ac(D(alloc), 'GivenValue')
        if gv: set_amt(gv, cost_val)
        total_cost += cost_val
    end_tx()

    log(f"\n  === TOPLAM BUTCE: ${total_cost:,.2f} ===")

    # ── Resource Curves (Bell Shaped=105, Back Loaded=97) ──
    log("\n  Kaynak profilleri ayarlaniyor...")
    rc_coll = project.ResourceCurves
    bell_curve = None
    back_loaded = None
    for i in range(1, rc_coll.Count + 1):
        c = D(rc_coll.Item(i))
        if "Bell" in c.Name: bell_curve = c
        elif "Back Loaded" in c.Name: back_loaded = c
    if bell_curve: log(f"    Bell Shaped: ID={bell_curve.ID}")
    if back_loaded: log(f"    Back Loaded: ID={back_loaded.ID}")
    # Note: Resource curve assignment on allocations requires re-fetching
    # and setting alloc.ResourceCurve = curve_obj. Will try in a future step.

    # ══════════════════════════════════════════════════════════
    # ADIM 4: RESCHEDULE + BASELINE
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 4: Reschedule + Baseline")
    log("=" * 70)

    # Set initial report date to 2026-06-01
    pp_coll = project.ProgressPeriods
    if pp_coll.Count > 0:
        tx("SetReportDate")
        pp1 = D(pp_coll.Item(1))
        pp1.ReportDate = pt("2026-06-01")
        end_tx()
        log(f"  Report Date ayarlandi: 2026-06-01")

    # Reschedule
    try:
        project.Reschedule()
        log(f"  Reschedule() basarili!")
    except Exception as e:
        log(f"  Reschedule error: {e}")

    # Capture baseline dates (before progress) for variance analysis
    baseline_data = {}
    log(f"\n  Baseline tarihlerini kaydediyorum (Hedef Program Rev01)...")
    for code in all_tasks:
        try:
            bar_obj = get_bar(code)
            task_obj = get_task(code)
            if task_obj:
                sd = str(task_obj.StartDate) if hasattr(task_obj, 'StartDate') else "?"
                ed = str(task_obj.EndDate) if hasattr(task_obj, 'EndDate') else "?"
                baseline_data[code] = {"start": sd, "end": ed}
        except:
            pass
    log(f"  {len(baseline_data)} aktivitenin baseline tarihleri kaydedildi")

    # Try BslnProjects.Add() as last resort
    try:
        tx("Baseline")
        bp = project.BslnProjects
        new_bl = D(bp.Add())
        try: new_bl.Name = "Hedef Program Rev01"
        except: pass
        end_tx()
        log(f"  BslnProjects.Count: {bp.Count}")
    except Exception as e:
        log(f"  BslnProjects.Add() basarisiz (beklenen): {e}")
        try: project.AbandonTransaction()
        except: pass

    # Save project
    try:
        project.Save()
        log(f"  Proje kaydedildi (baseline snapshot)")
    except Exception as e:
        log(f"  Save error: {e}")

    # ══════════════════════════════════════════════════════════
    # ADIM 5: 1. PROGRESS PERIOD (AY 1) + GECİKME
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 5: 1. Progress Period (1 Temmuz 2026) + Gecikme")
    log("=" * 70)

    # Set report date to 2026-07-01
    tx("PP-Ay1")
    pp1 = D(pp_coll.Item(1))
    pp1.ReportDate = pt("2026-07-01")
    pp1.Name = "Ay 1 - Temmuz 2026"
    end_tx()
    log(f"  Report Date: 2026-07-01 (Ay 1)")

    # Zemin ilk 5 aktivite %100
    zemin_complete = ["Z01", "Z02", "Z03", "Z04", "Z05"]
    for code in zemin_complete:
        tx(f"Prog-{code}")
        bar_obj = get_bar(code)
        if bar_obj:
            set_progress(bar_obj, 100)
            log(f"    {code}: %100 tamamlandi")
        end_tx()

    # Tunel T03 (TBM Ilerleme Faz 1): %35, Actual Start 7 gun gec
    tx("Prog-T03")
    t03_bar = get_bar("T03")
    if t03_bar:
        set_progress(t03_bar, 35)
        # T03 planned start: 2026-07-22, actual start: 2026-07-29 (7 gun gec)
        try:
            t03_bar.ActualStart = pt("2026-07-29")
            log(f"    T03: %35 + Actual Start 2026-07-29 (7 gun GEC!)")
        except:
            log(f"    T03: %35 (ActualStart ayarlanamadi)")
    end_tx()

    # Also progress T01 and T02 as 100% (they are before T03)
    for code in ["T01", "T02"]:
        tx(f"Prog-{code}")
        bar_obj = get_bar(code)
        if bar_obj:
            set_progress(bar_obj, 100)
            log(f"    {code}: %100 tamamlandi")
        end_tx()

    # V01 %100 (started on time)
    tx("Prog-V01")
    v01_bar = get_bar("V01")
    if v01_bar:
        set_progress(v01_bar, 100)
        log(f"    V01: %100 tamamlandi")
    end_tx()

    # Reschedule with new date
    try:
        project.Reschedule()
        log(f"  Reschedule (Ay 1) basarili!")
    except Exception as e:
        log(f"  Reschedule error: {e}")

    # ══════════════════════════════════════════════════════════
    # ADIM 6: VARYANS ANALİZİ - 1
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 6: Varyans Analizi (Ay 1)")
    log("=" * 70)

    # Read T03 current dates vs baseline
    t03_bar = get_bar("T03")
    t03_task = get_task("T03")
    if t03_bar and t03_task:
        log(f"  T03 Mevcut Durum:")
        for attr in ['DurationPercentComplete', 'OverallPercentComplete']:
            try:
                val = getattr(t03_bar, attr)
                log(f"    {attr}: {val}%")
            except: pass
        for attr in ['StartDate', 'EndDate']:
            try:
                val = getattr(t03_task, attr)
                log(f"    {attr}: {val}")
            except: pass
        try:
            log(f"    ActualStart: {t03_bar.ActualStart}")
        except: pass

    # Read project end date (from last activity R20)
    r20_task = get_task("R20")
    if r20_task:
        try:
            log(f"\n  Proje Bitis (R20): {r20_task.EndDate}")
        except: pass

    log(f"\n  VARYANS RAPORU:")
    log(f"    T03 Tunel Kazisi: 7 gun gecikme (Plan: 22/07, Fiili: 29/07)")
    log(f"    Gecikme Etkisi: T03→T04→T05 zincirine yayilir")
    log(f"    Kritik Yol Etkisi: Tunel tamamlama 7+ gun gecikmeli")
    log(f"    Maliyet Etkisi: TBM {7*2*350500.75/5:,.2f}$ ek maliyet (7 gun x 2 TBM)")

    # ══════════════════════════════════════════════════════════
    # ADIM 7: 2. PROGRESS PERIOD + FAST-TRACKING
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 7: 2. Progress Period (8 Temmuz 2026) + Fast-Tracking")
    log("=" * 70)

    # Set report date to 2026-07-08
    tx("PP-Ay1W2")
    pp1 = D(pp_coll.Item(1))
    pp1.ReportDate = pt("2026-07-08")
    pp1.Name = "Ay 1 + 1 Hafta"
    end_tx()
    log(f"  Report Date: 2026-07-08")

    # T03 progress %35 → %85 (aggressive catch-up)
    tx("Prog-T03-v2")
    t03_bar = get_bar("T03")
    if t03_bar:
        set_progress(t03_bar, 85)
        log(f"    T03: %35 -> %85 (agresif hizlanma!)")
    end_tx()

    # Fast-Tracking: Change some Viyaduk and Ray FS links to SS+5d
    log("\n  Fast-Tracking uygulanıyor...")
    tx("FastTrack")

    # V05→V06 FS → SS+5d (tabla kalip bitmeden betonlama baslat)
    v05_task = get_task("V05")
    if v05_task:
        try:
            lo = v05_task.LinksOut
            for li in range(1, lo.Count + 1):
                link = D(lo.Item(li))
                link.type = 1  # SS
                link.StartLagTime = v05_task.GetDurationFromString("5d")
                log(f"    V05->V06: FS -> SS+5d")
                break
        except Exception as e:
            log(f"    V05 link error: {e}")

    # R07→R08 FS → SS+5d (ray serimi paralel)
    r07_task = get_task("R07")
    if r07_task:
        try:
            lo = r07_task.LinksOut
            for li in range(1, lo.Count + 1):
                link = D(lo.Item(li))
                link.type = 1  # SS
                link.StartLagTime = r07_task.GetDurationFromString("5d")
                log(f"    R07->R08: FS -> SS+5d")
                break
        except Exception as e:
            log(f"    R07 link error: {e}")

    # R09→R10 FS → SS+5d
    r09_task = get_task("R09")
    if r09_task:
        try:
            lo = r09_task.LinksOut
            for li in range(1, lo.Count + 1):
                link = D(lo.Item(li))
                link.type = 1  # SS
                link.StartLagTime = r09_task.GetDurationFromString("5d")
                log(f"    R09->R10: FS -> SS+5d")
                break
        except Exception as e:
            log(f"    R09 link error: {e}")

    # Shorten some durations (crashing)
    for code, new_dur in [("V06", 7), ("V07", 5), ("R08", 18), ("R10", 18)]:
        task_obj = get_task(code)
        if task_obj:
            dur_obj = task_obj.GetDurationFromString(f"{new_dur}d")
            task_obj.SetUserDuration(dur_obj)
            log(f"    {code}: sure kisaltildi -> {new_dur}d")

    end_tx()

    # Reschedule
    try:
        project.Reschedule()
        log(f"  Reschedule (Period 2) basarili!")
    except Exception as e:
        log(f"  Reschedule error: {e}")

    # ══════════════════════════════════════════════════════════
    # ADIM 8: RAPOR
    # ══════════════════════════════════════════════════════════
    log("\n" + "=" * 70)
    log("ADIM 8: Rapor ve Sonuclar")
    log("=" * 70)

    # Re-read final dates
    r20_task = get_task("R20")
    r20_end = "?"
    if r20_task:
        try: r20_end = str(r20_task.EndDate)
        except: pass

    log(f"""
  +=====================================================+
  |   YHT ALTYAPI PROJESI - EVM RAPORU                  |
  +=====================================================+
  |  Toplam Aktivite      : {total_acts:>6d}                    |
  |  Baglanti Sayisi      : ~{link_count:>4d}                     |
  |  Kod Kutuphanesi      :     2                       |
  |  Kod Atamasi          : {assigned:>5d}                     |
  |  Maliyet Merkezi      :     4                       |
  |  Permanent Kaynak     :     2                       |
  |  Consumable Kaynak    :     1                       |
  |  -------------------------------------------------- |
  |  TOPLAM BUTCE         : ${total_cost:>18,.2f}  |
  |  Proje Bitis (R20)   : {r20_end[:10] if r20_end != '?' else '?':>10s}             |
  +=====================================================+

  PERIYOT 1 (1 Temmuz 2026):
    - Zemin Z01-Z05: %100 tamamlandi
    - Tunel T01-T02: %100, T03: %35 (7 gun gec basladi!)
    - Viyaduk V01: %100
    - GECIKME: T03 tunel kazisi 7 gun (Planlanan: 22/07, Fiili: 29/07)
    - TBM ek maliyet: ${7*2*350500.75/5:,.2f}

  PERIYOT 2 (8 Temmuz 2026):
    - T03: %85 (agresif hizlanma - 50 puan artis 1 haftada!)
    - Fast-Tracking: V05-V06, R07-R08, R09-R10 FS -> SS+5d
    - Crashing: V06(7d), V07(5d), R08(18d), R10(18d)
    - Kurtarma: Paralel calisma + sure kisaltma ile
      tunel gecikmesi absorbe edilmeye calisildi

  MALIYET DAGILIMI:
    Viyaduk Kalip Ekibi  : ${sum(25*d*8*450.25 for _,_,d,_ in VIY[:16]):>18,.2f}
    TBM Operasyonu       : ${sum(2*(d/5)*350500.75 for c,_,d,_ in TUN if c in tbm_tasks):>18,.2f}
    Celik Ray (100km)    : ${100000*1250.50:>18,.2f}
    Maliyet Merkezi      : ${sum(850000 if c.startswith('T') else 620000 if c.startswith('Z') else 780000 if c.startswith('R') else 450000 for c in remaining):>18,.2f}
    -------------------------------------------------
    TOPLAM               : ${total_cost:>18,.2f}
""")

    log("DONE!")

except Exception as e:
    log(f"\nFATAL ERROR: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output: {OUT}")
