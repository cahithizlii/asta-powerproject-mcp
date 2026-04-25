"""Build International Airport Terminal Project - OPTIMIZED
Target: 102+ activities, 4-level WBS, cross-links, $50M+ budget
All batched transactions. Target: < 3 minutes.
"""
import sys, io, datetime, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stdout.reconfigure(line_buffering=True)
import pythoncom, pywintypes, win32com.client

APP_CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
D = win32com.client.Dispatch
TODAY = pywintypes.Time(datetime.datetime(2026, 4, 1, 8, 0))
t0 = time.time()
def elapsed(): return f"[{time.time()-t0:.1f}s]"
def ole(dt): return pywintypes.Time(dt)

def connect():
    pythoncom.CoInitialize()
    obj = pythoncom.GetActiveObject(APP_CLSID)
    return D(obj.QueryInterface(pythoncom.IID_IDispatch))

# =====================================================================
# HIERARCHY DEFINITION (4-level WBS)
# =====================================================================
# tree[name] = (children_dict_or_task_list, is_summary)
# Leaf groups contain task tuples: (name, duration)
# Summaries contain sub-dicts

ALTYAPI = [
    ("Santiye Kurulumu ve Mobilizasyon","15d"),("Topografik Olcum ve Arazi Etud","8d"),
    ("Geoteknik Sondaj Raporlari","12d"),("Kazik Cakma Makinesi Montaj","10d"),
    ("Fore Kazik Imalati (Cep A)","25d"),("Fore Kazik Imalati (Cep B)","25d"),
    ("Fore Kazik Imalati (Cep C)","25d"),("Kazik Basligi Kirimi","15d"),
    ("Radye Temel Demir Baglamasi","20d"),("Radye Temel Beton Dokumu","12d"),
    ("Su Yalitimi ve Drenaj","15d"),("Bodrum Kat Perde Duvar","18d"),
    ("Bodrum Kat Kolon Imalati","16d"),("Bodrum Zemin Beton Dokumu","10d"),
    ("Mekanik Oda Altyapi","12d"),("Elektrik Oda Altyapi","10d"),
    ("Apron Baglanti Altyapisi","20d"),("Altyapi Teslim ve Kontrol","5d"),
]

BETONARME = [
    ("Kolon Kalip Imalati Kat 1","12d"),("Kolon Demir Baglama Kat 1","10d"),
    ("Kolon Beton Dokumu Kat 1","5d"),("Tabliye Kalip Kat 1","15d"),
    ("Tabliye Demir Kat 1","12d"),("Tabliye Beton Kat 1","6d"),
    ("Kolon Kalip Imalati Kat 2","12d"),("Kolon Demir Baglama Kat 2","10d"),
    ("Kolon Beton Dokumu Kat 2","5d"),("Tabliye Kalip Kat 2","15d"),
    ("Tabliye Demir Kat 2","12d"),("Tabliye Beton Kat 2","6d"),
    ("Perde Duvar Imalati","20d"),("Deprem Izolator Montaji","15d"),
    ("Cati Kirisleri Imalati","18d"),("Konsol Kirisleri Montaj","12d"),
    ("Karkas Kontrol ve Olcum","5d"),("Karkas Teslim Tutanagi","3d"),
]

UZAY_KAFES = [
    ("Atolye Celik Kesimi","20d"),("Celik Boyama ve Korozyon","15d"),
    ("Bilesenlerin Santiyeye Nakli","10d"),("Gecici Destek Kuleleri Montaji","12d"),
    ("Ana Makaslarin Kaldirma","8d"),("Ara Makaslar Montaji","15d"),
    ("Diagonal Eleman Baglantilari","12d"),("Guse Plakasi Kaynagi","10d"),
    ("Torklama ve Sikilastirma","8d"),("Jeodetik Olcum Kontrolu","5d"),
    ("Gecici Desteklerin Sokulumu","6d"),("Yuk Testi (Dead Load)","5d"),
    ("Uzay Kafes NDT Testi","8d"),("Anti-Korozyon Son Kat","10d"),
    ("Uzay Kafes Kabul Tutanagi","3d"),
]

TITANYUM = [
    ("Titanyum Panel Uretim Siparisi","5d"),("Titanyum Panel Fabrika Uretimi","45d"),
    ("Nakliye ve Gumruk Islemleri","20d"),("Alt Konstruksiyon Profilleri","15d"),
    ("Isi Yalitim Katmani","12d"),("Su Yalitim Membrani","10d"),
    ("Panel Montaj Baslangic (Bati)","18d"),("Panel Montaj Devam (Merkez)","18d"),
    ("Panel Montaj Bitis (Dogu)","18d"),("Derz ve Conta Uygulamalari","12d"),
    ("Yildirim Topraklama Baglantisi","8d"),("Su Sizdirmazlik Testi","5d"),
    ("Estetik Kontrol ve Rotus","6d"),("Dis Cephe Aydinlatma Baglantisi","10d"),
    ("Kaplama Kabul Tutanagi","3d"),
]

HVAC = [
    ("HVAC Muhendislik Hesaplari","15d"),("Chiller Grubu Siparis","10d"),
    ("Chiller Grubu Teslim","60d"),("AHU Uniteleri Siparis","8d"),
    ("AHU Uniteleri Teslim","45d"),("Kanal Imalati (Galvaniz)","25d"),
    ("Ana Kanal Montaji (Bodrum)","20d"),("Saft Yukseltme Kanallari","15d"),
    ("Terminal Salon Kanal Dagitim","18d"),("Difuzor ve Menfez Montaji","12d"),
    ("Chiller Makine Dairesi Montaj","15d"),("AHU Montaji ve Baglanti","12d"),
    ("Boru Tesisati (Sogutma)","20d"),("Izolasyon Isleri","15d"),
    ("Otomasyon ve BMS Baglanti","12d"),("Devreye Alma ve Balans","10d"),
    ("Hava Kalitesi Testi","5d"),("HVAC Kabul Tutanagi","3d"),
]

BAGAJ = [
    ("Bagaj Sistemi Muhendislik","15d"),("Konveyor Hat Siparis","10d"),
    ("Konveyor Hat Uretim/Teslim","75d"),("X-Ray Cihazlari Siparis","8d"),
    ("X-Ray Cihazlari Teslim","50d"),("Ana Hat Montaj (Varis)","20d"),
    ("Ana Hat Montaj (Gidis)","20d"),("Carosuel Montaji (4 Unite)","15d"),
    ("Check-in Konveyor Hatti","18d"),("Sorting Robot Montaji","12d"),
    ("Tracking ve RFID Sistemi","10d"),("EDS Guvenlik Tarama Hatti","15d"),
    ("Over-Size Bagaj Hatti","10d"),("Early Bag Storage Sistemi","12d"),
    ("PLC ve SCADA Programlama","18d"),("Sistem Entegrasyon Testi","12d"),
    ("72 Saat Performans Testi","5d"),("Bagaj Sistemi Kabul","3d"),
]

# Groups indexed 0-5 for link/assign references
GROUPS = [ALTYAPI, BETONARME, UZAY_KAFES, TITANYUM, HVAC, BAGAJ]
GROUP_NAMES = ["Altyapi","Betonarme","Uzay Kafes","Titanyum","HVAC","Bagaj"]

total_activities = sum(len(g) for g in GROUPS)
print(f"Defined: {total_activities} activities across 6 groups")

# =====================================================================
# LINK DEFINITIONS: (grp_from, idx_from, grp_to, idx_to, type, lag)
# =====================================================================
LINKS = []

# --- Internal FS chains per group ---
for gi, grp in enumerate(GROUPS):
    for i in range(len(grp) - 1):
        LINKS.append((gi, i, gi, i+1, "FS", None))

# --- Cross-WBS links ---
# Altyapi last -> Betonarme first (Altyapi bitmeden karkas baslamaz)
LINKS.append((0, 17, 1, 0, "FS", None))
# Altyapi bodrum -> HVAC kanal (bodrum tamamlaninca HVAC baslayabilir)
LINKS.append((0, 13, 4, 6, "FS", None))
# Betonarme cati kirisleri -> Uzay Kafes destek kuleleri
LINKS.append((1, 14, 2, 3, "FS", None))
# Betonarme teslim -> Titanyum alt konstruksiyon
LINKS.append((1, 17, 3, 3, "FS", None))
# Uzay Kafes kabul -> Titanyum panel montaj
LINKS.append((2, 14, 3, 6, "FS", None))
# HVAC AHU teslim -> HVAC AHU montaj (procurement -> install)
# (already internal)
# Bagaj konveyor teslim -> Ana hat montaj (already internal)
# Betonarme Kat 1 tabliye -> HVAC saft kanallari
LINKS.append((1, 5, 4, 7, "FS", None))
# Titanyum kaplama kabul -> Bagaj sistem entegrasyon
LINKS.append((3, 14, 5, 15, "FS", None))
# Altyapi mekanik oda -> HVAC chiller montaj
LINKS.append((0, 14, 4, 10, "FS", None))
# Altyapi elektrik oda -> Bagaj PLC programlama
LINKS.append((0, 15, 5, 14, "FS", None))

# --- SS + 5d lag links (at least 5) ---
# Altyapi Fore Kazik A SS+5d -> Fore Kazik B (parallel drilling)
LINKS.append((0, 4, 0, 5, "SS", "5d"))
# Altyapi Fore Kazik B SS+5d -> Fore Kazik C
LINKS.append((0, 5, 0, 6, "SS", "5d"))
# Betonarme Kolon Kat1 SS+5d -> Tabliye Kat1 (overlap)
LINKS.append((1, 0, 1, 3, "SS", "5d"))
# Betonarme Kolon Kat2 SS+5d -> Tabliye Kat2
LINKS.append((1, 6, 1, 9, "SS", "5d"))
# Uzay Kafes kesim SS+5d -> boyama (overlap in workshop)
LINKS.append((2, 0, 2, 1, "SS", "5d"))
# HVAC kanal imalat SS+5d -> ana kanal montaj
LINKS.append((4, 5, 4, 6, "SS", "5d"))
# Titanyum panel montaj Bati SS+5d -> Merkez
LINKS.append((3, 6, 3, 7, "SS", "5d"))
# Titanyum Merkez SS+5d -> Dogu
LINKS.append((3, 7, 3, 8, "SS", "5d"))

# --- FF + 2d lag links (at least 3) ---
# Betonarme demir Kat1 FF+2d -> beton Kat1 (finish together)
LINKS.append((1, 1, 1, 2, "FF", "2d"))
# Betonarme demir Kat2 FF+2d -> beton Kat2
LINKS.append((1, 7, 1, 8, "FF", "2d"))
# HVAC izolasyon FF+2d -> otomasyon (finish overlap)
LINKS.append((4, 13, 4, 14, "FF", "2d"))
# Bagaj RFID FF+2d -> EDS guvenlik
LINKS.append((5, 10, 5, 11, "FF", "2d"))

print(f"Defined: {len(LINKS)} links")

# =====================================================================
# RESOURCE ASSIGNMENTS
# =====================================================================
# (group_idx, task_idx, resource_name, units_or_none)
RES_ASSIGNS = []

# Agir Celik Montaj Ekibi -> all Uzay Kafes + Betonarme heavy tasks
for i in range(15): RES_ASSIGNS.append((2, i, "Agir Celik Montaj Ekibi", None))
for i in [12, 13, 14, 15]: RES_ASSIGNS.append((1, i, "Agir Celik Montaj Ekibi", None))

# Ozel Vinc Filosu -> steel tasks (Uzay Kafes lifting + Betonarme cranes)
# 1000 hours effort spread across tasks
for i in [3, 4, 5, 6, 10]: RES_ASSIGNS.append((2, i, "Ozel Vinc Filosu", None))  # ~200h each
for i in [14, 15]: RES_ASSIGNS.append((1, i, "Ozel Vinc Filosu", None))

# Titanyum Cati Paneli -> kaplama tasks, units=5000 total
# Spread across 3 montaj tasks + siparisler
for i in [6, 7, 8]: RES_ASSIGNS.append((3, i, "Titanyum Cati Paneli", None))

# Bagaj Ayristirici Robot -> bagaj tasks, units=40
for i in [9]: RES_ASSIGNS.append((5, i, "Bagaj Ayristirici Robot", None))

# Also assign Agir Celik to Altyapi heavy tasks
for i in [4, 5, 6, 7, 8, 9, 11, 12]: RES_ASSIGNS.append((0, i, "Agir Celik Montaj Ekibi", None))

# HVAC tasks get generic assignment
for i in [6, 7, 8, 9, 10, 11, 12]: RES_ASSIGNS.append((4, i, "Agir Celik Montaj Ekibi", None))

print(f"Defined: {len(RES_ASSIGNS)} resource assignments")

# =====================================================================
# CODE ASSIGNMENTS: (group_idx, task_idx, zone, discipline)
# =====================================================================
# Zone mapping: Altyapi->Merkez, Betonarme->varied, Uzay/Titanyum->varied, HVAC->Merkez, Bagaj->Dogu
ZONE_MAP = {
    0: ["Zone-Merkez"]*18,
    1: ["Zone-Bati"]*6 + ["Zone-Dogu"]*6 + ["Zone-Merkez"]*6,
    2: ["Zone-Bati"]*5 + ["Zone-Merkez"]*5 + ["Zone-Dogu"]*5,
    3: ["Zone-Bati"]*5 + ["Zone-Merkez"]*5 + ["Zone-Dogu"]*5,
    4: ["Zone-Merkez"]*18,
    5: ["Zone-Dogu"]*18,
}
DISC_MAP = {
    0: ["Statik"]*18,
    1: ["Statik"]*18,
    2: ["Statik"]*8 + ["Mimari"]*7,
    3: ["Mimari"]*15,
    4: ["Mekanik"]*18,
    5: ["Elektrik"]*18,
}

# =====================================================================
# WHAT-IF: Fast-tracking between Betonarme teslim and Titanyum alt konstr.
# Original: FS link (1,17) -> (3,3)
# New: SS + 10d lag (start together with overlap)
# Also crash some durations
# =====================================================================
WHATIF_REMOVE = (1, 17, 3, 3)  # Remove FS
WHATIF_ADD = (1, 14, 3, 3, "SS", "10d")  # Add SS+10d (cati kirisleri -> titanyum alt konstr)
WHATIF_CRASH = {
    (2, 0): "15d",  # Atolye celik kesimi 20->15d
    (2, 5): "10d",  # Ara makaslar 15->10d
    (3, 1): "35d",  # Titanyum fabrika uretimi 45->35d
    (3, 6): "14d",  # Panel montaj Bati 18->14d
    (3, 7): "14d",  # Panel montaj Merkez 18->14d
    (3, 8): "14d",  # Panel montaj Dogu 18->14d
}


# =====================================================================
# MAIN EXECUTION
# =====================================================================
def main():
    app = connect()
    project = app.ActiveProject
    print(f"{elapsed()} Project: {project.Name} (Bars={project.Bars.Count})", flush=True)

    # Store IDs
    root_id = None
    # summary_ids[path] = bar_id  (path like "root", "L2_0", "L3_2.1", "L4_2.2.1")
    summary_ids = {}
    # task_ids[(group_idx, task_idx)] = bar_id
    task_ids = {}

    def find_root_pos():
        for i in range(1, project.Bars.Count + 1):
            if D(project.Bars.Item(i)).ID == root_id:
                return i
        return project.Bars.Count

    # =================================================================
    # STEP 1: CREATE HIERARCHY (root + L2 + L3 + L4 summaries)
    # =================================================================
    print(f"\n{elapsed()} === STEP 1: HIERARCHY ===", flush=True)

    project.StartTransaction("Hierarchy")
    try:
        # Root (L1)
        root_bar = project.Bars.Add()
        root_bar.Name = "Havalimani Terminal Projesi"
        root_bar.Tasks.AddSummaryTask(TODAY)
        rt = D(root_bar.Tasks(1))
        rt.type = 1

        # L2: 1. Altyapi ve Temel
        b = rt.ChildBars.Add(); b.Name = "1. Altyapi ve Temel"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1

        # L2: 2. Ana Terminal Binasi
        b = rt.ChildBars.Add(); b.Name = "2. Ana Terminal Binasi"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1

        # L2: 3. MEP ve Bagaj Sistemleri
        b = rt.ChildBars.Add(); b.Name = "3. MEP ve Bagaj Sistemleri"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1

        # L3 under "2. Ana Terminal Binasi"
        l2_terminal = D(rt.ChildBars.Item(2))
        l2t = D(l2_terminal.Tasks(1))
        b = l2t.ChildBars.Add(); b.Name = "2.1. Betonarme Karkas"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1
        b = l2t.ChildBars.Add(); b.Name = "2.2. Celik Cati Sistemleri"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1

        # L4 under "2.2. Celik Cati Sistemleri"
        l3_celik = D(l2t.ChildBars.Item(2))
        l3c = D(l3_celik.Tasks(1))
        b = l3c.ChildBars.Add(); b.Name = "2.2.1. Uzay Kafes Montaji"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1
        b = l3c.ChildBars.Add(); b.Name = "2.2.2. Titanyum Kaplama"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1

        # L3 under "3. MEP ve Bagaj Sistemleri"
        l2_mep = D(rt.ChildBars.Item(3))
        l2m = D(l2_mep.Tasks(1))
        b = l2m.ChildBars.Add(); b.Name = "3.1. Havalandirma (HVAC)"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1
        b = l2m.ChildBars.Add(); b.Name = "3.2. Tam Otomatik Bagaj Bantlari"
        b.Tasks.AddSummaryTask(TODAY); D(b.Tasks(1)).type = 1

        project.EndTransaction()
        project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  FATAL hierarchy error: {e}")
        try: project.AbandonTransaction()
        except: pass
        return

    # Fetch all IDs
    root_bar = D(project.Bars.Item(project.Bars.Count))
    root_id = root_bar.ID
    rt = D(root_bar.Tasks(1))
    print(f"  Root: [{root_id}] {root_bar.Name}")

    # Map: group_idx -> path to parent summary for adding tasks
    # 0: Altyapi -> rt.ChildBars(1)
    # 1: Betonarme -> rt.ChildBars(2).Tasks(1).ChildBars(1)
    # 2: Uzay Kafes -> rt.ChildBars(2).Tasks(1).ChildBars(2).Tasks(1).ChildBars(1)
    # 3: Titanyum -> rt.ChildBars(2).Tasks(1).ChildBars(2).Tasks(1).ChildBars(2)
    # 4: HVAC -> rt.ChildBars(3).Tasks(1).ChildBars(1)
    # 5: Bagaj -> rt.ChildBars(3).Tasks(1).ChildBars(2)

    def get_parent_task(group_idx):
        """Navigate to the parent summary task for a given group."""
        rt = D(D(project.Bars.Item(find_root_pos())).Tasks(1))
        if group_idx == 0:
            return D(D(rt.ChildBars.Item(1)).Tasks(1))
        elif group_idx == 1:
            l2 = D(D(rt.ChildBars.Item(2)).Tasks(1))
            return D(D(l2.ChildBars.Item(1)).Tasks(1))
        elif group_idx == 2:
            l2 = D(D(rt.ChildBars.Item(2)).Tasks(1))
            l3 = D(D(l2.ChildBars.Item(2)).Tasks(1))
            return D(D(l3.ChildBars.Item(1)).Tasks(1))
        elif group_idx == 3:
            l2 = D(D(rt.ChildBars.Item(2)).Tasks(1))
            l3 = D(D(l2.ChildBars.Item(2)).Tasks(1))
            return D(D(l3.ChildBars.Item(2)).Tasks(1))
        elif group_idx == 4:
            l2 = D(D(rt.ChildBars.Item(3)).Tasks(1))
            return D(D(l2.ChildBars.Item(1)).Tasks(1))
        elif group_idx == 5:
            l2 = D(D(rt.ChildBars.Item(3)).Tasks(1))
            return D(D(l2.ChildBars.Item(2)).Tasks(1))

    # Print hierarchy
    for gi in range(6):
        pt = get_parent_task(gi)
        # Get parent bar
        print(f"  Group {gi} ({GROUP_NAMES[gi]}): parent found", flush=True)

    # =================================================================
    # STEP 2: CREATE ALL ACTIVITIES (one tx per group)
    # =================================================================
    print(f"\n{elapsed()} === STEP 2: ACTIVITIES ===", flush=True)

    for gi, grp in enumerate(GROUPS):
        project.StartTransaction(f"G{gi}")
        try:
            pt = get_parent_task(gi)
            for tname, dur in grp:
                nb = pt.ChildBars.Add()
                nb.Name = tname
                nb.Tasks.AddTask(TODAY, dur)
            project.EndTransaction()
            project.WaitForNotificationProcessing()
        except Exception as e:
            print(f"  ERROR group {gi}: {e}")
            try: project.AbandonTransaction()
            except: pass
            continue

        # Fetch IDs
        pt = get_parent_task(gi)
        for ti in range(1, pt.ChildBars.Count + 1):
            tb = D(pt.ChildBars.Item(ti))
            task_ids[(gi, ti - 1)] = tb.ID
        print(f"  {elapsed()} Group {gi} ({GROUP_NAMES[gi]}): {len(grp)} tasks", flush=True)

    print(f"  Total: {len(task_ids)} activities", flush=True)

    # =================================================================
    # STEP 3: CREATE ALL LINKS (one transaction)
    # =================================================================
    print(f"\n{elapsed()} === STEP 3: LINKS ===", flush=True)

    def build_task_cache():
        """Build bar_id -> task cache from all groups."""
        cache = {}
        for gi in range(6):
            pt = get_parent_task(gi)
            try:
                for ti in range(1, pt.ChildBars.Count + 1):
                    tb = D(pt.ChildBars.Item(ti))
                    tt = D(tb.Tasks(1))
                    cache[tb.ID] = tt
            except:
                pass
        return cache

    tmap = {"FS": 0, "SS": 1, "FF": 2, "SF": 3}

    project.StartTransaction("Links")
    try:
        cache = build_task_cache()
        link_ok = 0
        link_fail = 0

        for gi1, ti1, gi2, ti2, ltype, lag in LINKS:
            bid1 = task_ids.get((gi1, ti1))
            bid2 = task_ids.get((gi2, ti2))
            if not bid1 or not bid2 or bid1 not in cache or bid2 not in cache:
                link_fail += 1
                continue
            try:
                lnk = D(cache[bid1].LinkTo(cache[bid2]))
                lnk.type = tmap[ltype]
                if lag:
                    lnk.StartLagTime = cache[bid1].GetDurationFromString(lag)
                link_ok += 1
            except:
                link_fail += 1

        project.EndTransaction()
        project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  FATAL: {e}")
        try: project.AbandonTransaction()
        except: pass
        link_ok = 0; link_fail = len(LINKS)

    print(f"  Links: {link_ok}/{len(LINKS)} ok ({link_fail} failed)", flush=True)

    # Reschedule
    project.Reschedule(); project.WaitForNotificationProcessing()
    original_end = str(project.ProjectEnd)
    print(f"  {elapsed()} Rescheduled: {project.ProjectStart} -> {project.ProjectEnd}", flush=True)

    # =================================================================
    # STEP 4: COST CENTRES
    # =================================================================
    print(f"\n{elapsed()} === STEP 4: COST CENTRES ===", flush=True)

    project.StartTransaction("CC")
    try:
        cc1 = project.CostCentres.Add(); cc1.Name = "Terminal Genel Butcesi"
        cc2 = project.CostCentres.Add(); cc2.Name = "A-Iscilik Butcesi"
        cc3 = project.CostCentres.Add(); cc3.Name = "B-Malzeme ve Ekipman Butcesi"
        cc4 = project.CostCentres.Add(); cc4.Name = "C-Ozel Sistemler"
        print(f"  Created 4 cost centres")
        project.EndTransaction(); project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  ERROR: {e}")
        try: project.AbandonTransaction()
        except: pass

    # =================================================================
    # STEP 5: RESOURCES
    # =================================================================
    print(f"\n{elapsed()} === STEP 5: RESOURCES ===", flush=True)

    project.StartTransaction("Res")
    try:
        r1 = project.PermanentResources.Add()
        r1.Name = "Agir Celik Montaj Ekibi"
        try: r1.StandardRate = 1500.0
        except: pass
        print(f"  Perm: Agir Celik Montaj Ekibi ($1500/hr)")

        r2 = project.PermanentResources.Add()
        r2.Name = "Ozel Vinc Filosu"
        try: r2.StandardRate = 5000.0
        except: pass
        print(f"  Perm: Ozel Vinc Filosu ($5000/hr)")

        r3 = project.ConsumableResources.Add()
        r3.Name = "Titanyum Cati Paneli"
        try: r3.CostPerUnit = 12500.0
        except: pass
        print(f"  Cons: Titanyum Cati Paneli ($12500/adet)")

        r4 = project.ConsumableResources.Add()
        r4.Name = "Bagaj Ayristirici Robot"
        try: r4.CostPerUnit = 250000.0
        except: pass
        print(f"  Cons: Bagaj Ayristirici Robot ($250000/adet)")

        project.EndTransaction(); project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  ERROR: {e}")
        try: project.AbandonTransaction()
        except: pass

    # =================================================================
    # STEP 6: RESOURCE ASSIGNMENTS
    # =================================================================
    print(f"\n{elapsed()} === STEP 6: RESOURCE ASSIGNMENTS ===", flush=True)

    project.StartTransaction("Assign")
    try:
        cache = build_task_cache()

        # Build resource lookup
        res_map = {}
        for ri in range(1, project.PermanentResources.Count + 1):
            r_raw = project.PermanentResources.Item(ri)
            res_map[D(r_raw).Name] = r_raw
        for ri in range(1, project.ConsumableResources.Count + 1):
            r_raw = project.ConsumableResources.Item(ri)
            res_map[D(r_raw).Name] = r_raw

        assign_ok = 0
        for gi, ti, rname, units in RES_ASSIGNS:
            bid = task_ids.get((gi, ti))
            if not bid or bid not in cache or rname not in res_map:
                continue
            try:
                cache[bid].AssignResource(res_map[rname], True)
                assign_ok += 1
            except:
                pass

        project.EndTransaction(); project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  ERROR: {e}")
        try: project.AbandonTransaction()
        except: pass
        assign_ok = 0

    print(f"  Assignments: {assign_ok}/{len(RES_ASSIGNS)}", flush=True)

    # =================================================================
    # STEP 7: COST ASSIGNMENTS (assign cost centres to all tasks)
    # =================================================================
    print(f"\n{elapsed()} === STEP 7: COST ASSIGNMENTS ===", flush=True)

    # Map: group -> cost centre name
    COST_MAP = {
        0: "A-Iscilik Butcesi",       # Altyapi = labor
        1: "A-Iscilik Butcesi",       # Betonarme = labor
        2: "B-Malzeme ve Ekipman Butcesi",  # Uzay Kafes = material/equip
        3: "B-Malzeme ve Ekipman Butcesi",  # Titanyum = material
        4: "B-Malzeme ve Ekipman Butcesi",  # HVAC = equipment
        5: "C-Ozel Sistemler",        # Bagaj = special systems
    }

    project.StartTransaction("Costs")
    try:
        cache = build_task_cache()
        cc_map = {}
        for ci in range(1, project.CostCentres.Count + 1):
            cc_raw = project.CostCentres.Item(ci)
            cc_map[D(cc_raw).Name] = cc_raw

        cost_ok = 0
        for gi in range(6):
            ccname = COST_MAP[gi]
            if ccname not in cc_map: continue
            grp = GROUPS[gi]
            for ti in range(len(grp)):
                bid = task_ids.get((gi, ti))
                if not bid or bid not in cache: continue
                try:
                    cache[bid].AssignCost(cc_map[ccname])
                    cost_ok += 1
                except:
                    pass

        project.EndTransaction(); project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  ERROR: {e}")
        try: project.AbandonTransaction()
        except: pass
        cost_ok = 0

    print(f"  Cost assignments: {cost_ok}/{total_activities}", flush=True)

    # =================================================================
    # STEP 8: CODE LIBRARIES (2 libraries, matrix coding)
    # =================================================================
    print(f"\n{elapsed()} === STEP 8: CODE LIBRARIES ===", flush=True)

    # Create libraries
    project.StartTransaction("Libs")
    try:
        lib1 = project.CodeLibrarys.Add()
        lib1.Name = "Zonlar"
        for ename in ["Zone-Bati", "Zone-Dogu", "Zone-Merkez"]:
            e = lib1.Entries.Add(); e.Name = ename
        print(f"  Library: Zonlar (3 entries)")

        lib2 = project.CodeLibrarys.Add()
        lib2.Name = "Disiplin"
        for ename in ["Mimari", "Statik", "Mekanik", "Elektrik"]:
            e = lib2.Entries.Add(); e.Name = ename
        print(f"  Library: Disiplin (4 entries)")

        project.EndTransaction(); project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  ERROR: {e}")
        try: project.AbandonTransaction()
        except: pass

    # Assign codes (both zone AND discipline per task)
    project.StartTransaction("Codes")
    try:
        cache = build_task_cache()

        # Build entry lookup
        zone_entries = {}
        disc_entries = {}
        libs = project.CodeLibrarys
        for ci in range(1, libs.Count + 1):
            l = D(libs.Item(ci))
            if l.Name == "Zonlar":
                ents = l.Entries
                for ei in range(1, ents.Count + 1):
                    ent_raw = ents.Item(ei)
                    zone_entries[D(ent_raw).Name] = ent_raw
            elif l.Name == "Disiplin":
                ents = l.Entries
                for ei in range(1, ents.Count + 1):
                    ent_raw = ents.Item(ei)
                    disc_entries[D(ent_raw).Name] = ent_raw

        zone_ok = 0; disc_ok = 0
        for gi in range(6):
            grp = GROUPS[gi]
            zones = ZONE_MAP[gi]
            discs = DISC_MAP[gi]
            for ti in range(len(grp)):
                bid = task_ids.get((gi, ti))
                if not bid or bid not in cache: continue
                task = cache[bid]
                # Assign zone
                zname = zones[ti] if ti < len(zones) else zones[-1]
                if zname in zone_entries:
                    try:
                        task.AssignCode(zone_entries[zname], True)
                        zone_ok += 1
                    except: pass
                # Assign discipline
                dname = discs[ti] if ti < len(discs) else discs[-1]
                if dname in disc_entries:
                    try:
                        task.AssignCode(disc_entries[dname], True)
                        disc_ok += 1
                    except: pass

        project.EndTransaction(); project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  ERROR: {e}")
        try: project.AbandonTransaction()
        except: pass
        zone_ok = 0; disc_ok = 0

    print(f"  Zone codes: {zone_ok}/{total_activities}", flush=True)
    print(f"  Discipline codes: {disc_ok}/{total_activities}", flush=True)

    # =================================================================
    # STEP 9: WHAT-IF (Fast-Tracking)
    # =================================================================
    print(f"\n{elapsed()} === STEP 9: WHAT-IF (Fast-Tracking) ===", flush=True)

    project.Reschedule(); project.WaitForNotificationProcessing()
    pre_whatif_end = str(project.ProjectEnd)
    print(f"  Pre-WhatIf End: {pre_whatif_end}", flush=True)

    # 9a: Remove FS link and add SS+10d (fast-tracking)
    project.StartTransaction("FastTrack")
    try:
        cache = build_task_cache()

        # Remove FS link: Betonarme teslim -> Titanyum alt konstruksiyon
        gi1, ti1, gi2, ti2 = WHATIF_REMOVE
        bid1 = task_ids.get((gi1, ti1))
        bid2 = task_ids.get((gi2, ti2))
        if bid1 in cache and bid2 in cache:
            t1 = cache[bid1]
            # Find and remove the link
            try:
                lo = t1.LinksOut
                for li in range(lo.Count, 0, -1):
                    lnk = D(lo.Item(li))
                    end_task = D(lnk.EndTask)
                    # Check if this links to bid2
                    # Compare by checking if the end task matches
                    try:
                        lo.Remove(li)
                        print(f"  Removed FS link: {GROUP_NAMES[gi1]}[{ti1}] -> {GROUP_NAMES[gi2]}[{ti2}]")
                        break
                    except:
                        pass
            except Exception as e:
                print(f"  Remove link error: {e}")

        # Add SS+10d: Betonarme cati kirisleri -> Titanyum alt konstr
        gi1b, ti1b, gi2b, ti2b, ltype, lag = WHATIF_ADD
        bid1b = task_ids.get((gi1b, ti1b))
        bid2b = task_ids.get((gi2b, ti2b))
        if bid1b in cache and bid2b in cache:
            try:
                lnk = D(cache[bid1b].LinkTo(cache[bid2b]))
                lnk.type = tmap[ltype]
                if lag:
                    lnk.StartLagTime = cache[bid1b].GetDurationFromString(lag)
                print(f"  Added {ltype}+{lag}: {GROUP_NAMES[gi1b]}[{ti1b}] -> {GROUP_NAMES[gi2b]}[{ti2b}]")
            except Exception as e:
                print(f"  Add link error: {e}")

        # Crash durations
        crash_ok = 0
        for (cgi, cti), ndur in WHATIF_CRASH.items():
            bid = task_ids.get((cgi, cti))
            if bid and bid in cache:
                try:
                    t = cache[bid]
                    t.SetUserDuration(t.GetDurationFromString(ndur))
                    crash_ok += 1
                except:
                    pass
        print(f"  Crashed: {crash_ok}/{len(WHATIF_CRASH)} tasks")

        project.EndTransaction(); project.WaitForNotificationProcessing()
    except Exception as e:
        print(f"  FATAL: {e}")
        try: project.AbandonTransaction()
        except: pass

    project.Reschedule(); project.WaitForNotificationProcessing()
    post_whatif_end = str(project.ProjectEnd)
    print(f"  Post-WhatIf End: {post_whatif_end}", flush=True)
    print(f"  Original End:    {pre_whatif_end}", flush=True)

    # =================================================================
    # STEP 10: VIEW SETTINGS
    # =================================================================
    print(f"\n{elapsed()} === STEP 10: VIEW SETTINGS ===", flush=True)

    view = D(project.Views.Item(1))
    try: view.DisplayCriticalPath = True; print(f"  Critical Path ON")
    except: pass
    try: view.DisplayTotalFloat = True; print(f"  Total Float ON")
    except: pass
    try: view.DisplayProgressLines = True; print(f"  Progress Lines ON")
    except: pass
    try: view.ShowHierarchy(2); print(f"  Hierarchy Level 2")
    except Exception as e: print(f"  ShowHierarchy error: {e}")

    # Configure histogram
    try:
        hpane = D(view.Histogram())
        if hpane.HistogramCount == 0:
            project.StartTransaction("Hist")
            hpane2 = D(D(project.Views.Item(1)).Histogram())
            hpane2.AddHistogram()
            project.EndTransaction(); project.WaitForNotificationProcessing()
            hpane = D(D(project.Views.Item(1)).Histogram())

        report = project.HistogramReports.Item(1)
        hpane.SetHistogramReport(0, report)
        print(f"  Histogram: Allocation report set")
    except Exception as e:
        print(f"  Histogram error: {e}")

    # =================================================================
    # STEP 11: EXPORTS
    # =================================================================
    print(f"\n{elapsed()} === STEP 11: EXPORTS ===", flush=True)

    # XML
    xml_path = r"C:\Users\CahAsus\Desktop\Airport_Terminal.xml"
    try:
        project.SaveAsXMLFile(xml_path, None, None)
        print(f"  XML: {xml_path}")
    except Exception as e:
        print(f"  XML error: {e}")

    # HTML
    html_path = r"C:\Users\CahAsus\Desktop\Airport_Terminal.html"
    try:
        project.SaveAsHTMLFile(html_path)
        print(f"  HTML: {html_path}")
    except Exception as e:
        print(f"  HTML error: {e}")

    # PDF attempt
    try:
        pp = project.PrintProfiles
        project.StartTransaction("PDFProf")
        profile = pp.Add()
        pd = D(profile)
        pd.Name = "Airport PDF"
        pd.SetPrinterName("Microsoft Print to PDF")
        pd.PrintToFile = True
        pd.OutputFile = r"C:\Users\CahAsus\Desktop\Terminal_Executive_Summary.pdf"
        pd.Landscape = True
        pd.HistogramLegend = True
        pd.ProgressLegend = True
        pd.CodeLibraryLegend = True
        project.EndTransaction(); project.WaitForNotificationProcessing()
        print(f"  PDF profile 'Airport PDF' created (use Ctrl+P in Asta to print)")
    except Exception as e:
        print(f"  PDF profile error: {e}")
        try: project.AbandonTransaction()
        except: pass

    # =================================================================
    # FINAL SUMMARY
    # =================================================================
    print(f"\n{elapsed()} === FINAL SUMMARY ===", flush=True)

    # Hierarchy check
    rt = D(D(project.Bars.Item(find_root_pos())).Tasks(1))
    print(f"  Root: [{root_id}] type={rt.type}")
    for i in range(1, rt.ChildBars.Count + 1):
        l2 = D(rt.ChildBars.Item(i))
        l2t = D(l2.Tasks(1))
        try:
            l2c = l2t.ChildBars.Count
        except:
            l2c = 0
        print(f"    L2: {l2.Name} (type={l2t.type}, children={l2c})")
        if l2c > 0:
            for j in range(1, l2c + 1):
                l3 = D(l2t.ChildBars.Item(j))
                l3t = D(l3.Tasks(1))
                try:
                    l3c = l3t.ChildBars.Count
                except:
                    l3c = 0
                print(f"      L3: {l3.Name} (type={l3t.type}, children={l3c})")
                if l3c > 0:
                    for k in range(1, l3c + 1):
                        l4 = D(l3t.ChildBars.Item(k))
                        l4t = D(l4.Tasks(1))
                        try:
                            l4c = l4t.ChildBars.Count
                        except:
                            l4c = 0
                        print(f"        L4: {l4.Name} (type={l4t.type}, children={l4c})")

    print(f"\n  Activities: {len(task_ids)}")
    print(f"  Links: {link_ok}")
    print(f"  Resource Assignments: {assign_ok}")
    print(f"  Cost Assignments: {cost_ok}")
    print(f"  Zone Codes: {zone_ok}")
    print(f"  Discipline Codes: {disc_ok}")
    print(f"  Start: {project.ProjectStart}")
    print(f"  End: {project.ProjectEnd}")
    print(f"  Pre-WhatIf End: {pre_whatif_end}")
    print(f"  Post-WhatIf End: {post_whatif_end}")
    print(f"  Perm Res: {project.PermanentResources.Count}")
    print(f"  Cons Res: {project.ConsumableResources.Count}")
    print(f"  Cost Centres: {project.CostCentres.Count}")

    # Save IDs
    with open(r"C:\Users\CahAsus\asta-powerproject-mcp\airport_ids.txt", "w", encoding="utf-8") as f:
        for gi, gname in enumerate(GROUP_NAMES):
            f.write(f"=== {gname} ===\n")
            for ti, (tname, dur) in enumerate(GROUPS[gi]):
                bid = task_ids.get((gi, ti), "?")
                f.write(f"{bid}\t{tname}\n")

    print(f"\n{elapsed()} DONE!", flush=True)

if __name__ == "__main__":
    main()
