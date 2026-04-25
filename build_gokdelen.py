"""
Gokdelen Kompleksi - 80+ Aktivite toplu COM scripti
Mevcut projedeki tum barlari temizler, sifirdan olusturur
"""
import sys, io, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pythoncom, pywintypes, win32com.client

APP_CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
D = win32com.client.Dispatch

def connect():
    pythoncom.CoInitialize()
    obj = pythoncom.GetActiveObject(APP_CLSID)
    return D(obj.QueryInterface(pythoncom.IID_IDispatch))

def ole(dt):
    return pywintypes.Time(dt)

TODAY = ole(datetime.datetime(2025, 9, 15))

def add_task_under(project, parent_task, name, dur_str):
    """Add child task under parent. Returns (bar, task) or (None, None)."""
    project.StartTransaction(f"Add {name[:30]}")
    try:
        new_bar = parent_task.ChildBars.Add()
        new_bar.Name = name
        task = new_bar.Tasks.AddTask(TODAY, dur_str)
        project.EndTransaction()
        project.WaitForNotificationProcessing()
        # Re-fetch
        cb = parent_task.ChildBars
        bar = D(cb.Item(cb.Count))
        t = D(bar.Tasks(1))
        return bar, t
    except Exception as e:
        print(f"  ERR {name}: {e}")
        try: project.AbandonTransaction()
        except: pass
        return None, None

def add_summary_under(project, parent_task, name):
    """Add summary child under parent. Returns (bar, etask)."""
    project.StartTransaction(f"AddSum {name[:25]}")
    try:
        new_bar = parent_task.ChildBars.Add()
        new_bar.Name = name
        etask = new_bar.Tasks.AddExpandedTask(TODAY)
        project.EndTransaction()
        project.WaitForNotificationProcessing()
        cb = parent_task.ChildBars
        bar = D(cb.Item(cb.Count))
        t = D(bar.Tasks(1))
        return bar, t
    except Exception as e:
        print(f"  ERR sum {name}: {e}")
        try: project.AbandonTransaction()
        except: pass
        return None, None

def link_fs(project, pred, succ, lag=None):
    project.StartTransaction("LnkFS")
    try:
        lnk = pred.LinkTo(succ)
        if lag:
            lnk.StartLagTime = pred.GetDurationFromString(lag)
        project.EndTransaction()
        project.WaitForNotificationProcessing()
        return True
    except:
        try: project.AbandonTransaction()
        except: pass
        return False

def link_typed(project, pred, succ, ltype, lag=None):
    tmap = {"FS":0, "SS":1, "FF":2, "SF":3}
    project.StartTransaction(f"Lnk{ltype}")
    try:
        lnk = pred.LinkTo(succ)
        lnk.type = tmap[ltype]
        if lag:
            lnk.StartLagTime = pred.GetDurationFromString(lag)
        project.EndTransaction()
        project.WaitForNotificationProcessing()
        return True
    except:
        try: project.AbandonTransaction()
        except: pass
        return False

def clean_project(project):
    """Remove all existing bars."""
    bars = project.Bars
    while bars.Count > 0:
        project.StartTransaction("Clean")
        try:
            bars.Remove(1)
            project.EndTransaction()
            project.WaitForNotificationProcessing()
        except:
            try: project.AbandonTransaction()
            except: pass
            break

def main():
    app = connect()
    project = app.ActiveProject
    print(f"Project: {project.Name}")

    # Clean existing
    print("Cleaning existing bars...")
    clean_project(project)
    print(f"  Bars after clean: {project.Bars.Count}")

    # Create root summary
    print("\nCreating root summary...")
    project.StartTransaction("Root")
    root_bar = project.Bars.Add()
    root_bar.Name = "Gokdelen Kompleksi Insaati"
    root_etask = root_bar.Tasks.AddExpandedTask(TODAY)
    project.EndTransaction()
    project.WaitForNotificationProcessing()
    root_bar = D(project.Bars.Item(1))
    root_task = D(root_bar.Tasks(1))
    print(f"  Root: [{root_bar.ID}] {root_bar.Name}")

    # 6 sub-summaries
    print("\nCreating 6 WBS summaries...")
    phase_names = [
        "1. Muhendislik ve Tasarim",
        "2. Tedarik ve Satin Alma",
        "3. Alt Yapi ve Temel",
        "4. Ust Yapi - 20 Kat",
        "5. MEP ve Ince Isler",
        "6. Test ve Devreye Alma",
    ]
    phases = {}
    phase_keys = ["eng", "proc", "sub", "super", "mep", "comm"]
    for i, name in enumerate(phase_names):
        bar, task = add_summary_under(project, root_task, name)
        if task:
            phases[phase_keys[i]] = (bar, task)
            print(f"  [{bar.ID}] {name}")

    # =========================================================================
    # PHASE 1: ENGINEERING (10 activities)
    # =========================================================================
    print("\n=== PHASE 1: ENGINEERING (10) ===")
    eng = []
    eng_def = [
        ("Mimari On Tasarim", "15d"),
        ("Statik Hesap ve Tasarim", "20d"),
        ("Mekanik Tesisat Projesi", "15d"),
        ("Elektrik ve Otomasyon Projesi", "12d"),
        ("Yangin Guvenligi Projesi", "10d"),
        ("Asansor Teknik Sartnamesi", "8d"),
        ("Cephe Muhendisligi Projesi", "12d"),
        ("Peyzaj ve Cevre Duzenleme Proj", "8d"),
        ("Proje Koordinasyonu ve Uyumlama", "10d"),
        ("Belediye Onay Sureci", "20d"),
    ]
    p = phases["eng"][1]
    for nm, dr in eng_def:
        b, t = add_task_under(project, p, nm, dr)
        if t: eng.append((b, t, nm)); print(f"  [{b.ID}] {nm}")

    # Link chain
    for i in range(len(eng)-1):
        if i in [2, 3, 5, 6]: link_typed(project, eng[i][1], eng[i+1][1], "SS", "5d")
        else: link_fs(project, eng[i][1], eng[i+1][1])
    print(f"  Linked {len(eng)} engineering tasks")

    # =========================================================================
    # PHASE 2: PROCUREMENT (10 activities)
    # =========================================================================
    print("\n=== PHASE 2: PROCUREMENT (10) ===")
    proc = []
    proc_def = [
        ("Yapisal Celik Siparisi", "15d"),
        ("Yapisal Celik Teslimati", "30d"),
        ("C50 Beton Sozlesmesi", "10d"),
        ("Asansor Sistemleri Siparisi", "12d"),
        ("Asansor Uretim ve Sevkiyati", "60d"),
        ("Cephe Panelleri Siparisi", "10d"),
        ("Cephe Panelleri Uretimi", "45d"),
        ("MEP Ekipman Tedarigi", "20d"),
        ("Akilli Bina Otomasyon Sistemi", "15d"),
        ("Is Guvenligi Ekipmanlari", "10d"),
    ]
    p = phases["proc"][1]
    for nm, dr in proc_def:
        b, t = add_task_under(project, p, nm, dr)
        if t: proc.append((b, t, nm)); print(f"  [{b.ID}] {nm}")

    # Links
    if eng: link_fs(project, eng[-1][1], proc[0][1])  # Onay -> Celik Siparis
    link_fs(project, proc[0][1], proc[1][1])           # Siparis -> Teslimat
    link_typed(project, proc[0][1], proc[2][1], "SS", "5d")  # Siparis -> C50
    if len(eng) >= 6: link_fs(project, eng[5][1], proc[3][1])  # Asansor Sart -> Siparis
    link_fs(project, proc[3][1], proc[4][1])  # Siparis -> Uretim
    if len(eng) >= 7: link_fs(project, eng[6][1], proc[5][1])  # Cephe Proj -> Cephe Sip
    link_fs(project, proc[5][1], proc[6][1])  # Sip -> Uretim
    link_typed(project, proc[2][1], proc[7][1], "SS", "10d")
    link_typed(project, proc[7][1], proc[8][1], "SS", "5d")
    link_typed(project, proc[7][1], proc[9][1], "SS", "3d")
    print(f"  Linked {len(proc)} procurement tasks")

    # =========================================================================
    # PHASE 3: SUBSTRUCTURE (15 activities)
    # =========================================================================
    print("\n=== PHASE 3: SUBSTRUCTURE (15) ===")
    sub = []
    sub_def = [
        ("Santiye Kurulumu", "15d"),
        ("Topografik Olcum", "5d"),
        ("Genel Hafriyat", "20d"),
        ("Iksa Sistemi Tasarimi", "10d"),
        ("Iksa Kazik Cakma", "25d"),
        ("Dewatering Kurulumu", "10d"),
        ("Derin Hafriyat B3-B1", "30d"),
        ("Fore Kazik Imalati", "35d"),
        ("Kazik Basligi Betonu", "15d"),
        ("Radye Temel Demir Donatisi", "20d"),
        ("Radye Temel Beton Dokumu", "10d"),
        ("Bodrum B3 Perde Kolon", "15d"),
        ("Bodrum B2 Perde Kolon", "15d"),
        ("Bodrum B1 Perde Kolon", "15d"),
        ("Zemin Kat Tabliye Betonu", "12d"),
    ]
    p = phases["sub"][1]
    for nm, dr in sub_def:
        b, t = add_task_under(project, p, nm, dr)
        if t: sub.append((b, t, nm)); print(f"  [{b.ID}] {nm}")

    # Links
    if proc: link_fs(project, proc[1][1], sub[0][1])  # Celik Teslimat -> Santiye
    link_fs(project, sub[0][1], sub[1][1])
    link_fs(project, sub[1][1], sub[2][1])
    link_typed(project, sub[2][1], sub[3][1], "SS", "10d")
    link_fs(project, sub[3][1], sub[4][1])
    link_typed(project, sub[4][1], sub[5][1], "SS", "5d")
    link_fs(project, sub[4][1], sub[6][1])
    link_typed(project, sub[6][1], sub[7][1], "SS", "10d")
    link_fs(project, sub[7][1], sub[8][1])
    link_fs(project, sub[8][1], sub[9][1])
    link_fs(project, sub[9][1], sub[10][1])
    link_fs(project, sub[10][1], sub[11][1])
    link_fs(project, sub[11][1], sub[12][1])
    link_fs(project, sub[12][1], sub[13][1])
    link_fs(project, sub[13][1], sub[14][1])
    print(f"  Linked {len(sub)} substructure tasks")

    # =========================================================================
    # PHASE 4: SUPERSTRUCTURE 20 Kat (40 activities)
    # =========================================================================
    print("\n=== PHASE 4: SUPERSTRUCTURE 20 KAT (40) ===")
    sup = []
    p = phases["super"][1]
    for kat in range(1, 21):
        nm1 = f"Kat {kat:02d} Kolon Perde Kalip"
        dr1 = "5d" if kat <= 5 else ("4d" if kat <= 15 else "3d")
        b1, t1 = add_task_under(project, p, nm1, dr1)
        if t1: sup.append((b1, t1, nm1)); print(f"  [{b1.ID}] {nm1}")

        nm2 = f"Kat {kat:02d} Tabliye Beton"
        dr2 = "3d" if kat <= 10 else "2d"
        b2, t2 = add_task_under(project, p, nm2, dr2)
        if t2: sup.append((b2, t2, nm2)); print(f"  [{b2.ID}] {nm2}")

    # Links
    if sub and sup:
        link_fs(project, sub[-1][1], sup[0][1])  # Zemin Tabliye -> Kat1 Kolon
    for i in range(len(sup)-1):
        if i % 2 == 0:
            link_fs(project, sup[i][1], sup[i+1][1])  # Kolon -> Tabliye
        else:
            link_fs(project, sup[i][1], sup[i+1][1], "1d")  # Tabliye -> next Kolon +1d cure
    print(f"  Linked {len(sup)} superstructure tasks")

    # =========================================================================
    # PHASE 5: MEP & INCE ISLER (15 activities)
    # =========================================================================
    print("\n=== PHASE 5: MEP & INCE ISLER (15) ===")
    mep = []
    mep_def = [
        ("Ic Duvar Orgu Isleri", "25d"),
        ("Sihhi Tesisat Kaba Montaj", "20d"),
        ("Elektrik Kablo Cekimi", "20d"),
        ("HVAC Kanal Montaji", "18d"),
        ("Yangin Algilama Sondurme", "15d"),
        ("Asansor Kuyusu Hazirligi", "10d"),
        ("Asansor Montaji", "40d"),
        ("Ic Cephe Siva ve Boya", "30d"),
        ("Dis Cephe Kaplama Montaji", "35d"),
        ("Zemin Kaplama Seramik", "25d"),
        ("Asma Tavan Montaji", "18d"),
        ("Ic Dograma ve Kapilar", "15d"),
        ("Mutfak Banyo Armaturleri", "12d"),
        ("Peyzaj ve Dis Mekan", "20d"),
        ("Genel Temizlik Punch List", "10d"),
    ]
    p = phases["mep"][1]
    for nm, dr in mep_def:
        b, t = add_task_under(project, p, nm, dr)
        if t: mep.append((b, t, nm)); print(f"  [{b.ID}] {nm}")

    # Links
    if sup and mep:
        link_fs(project, sup[-1][1], mep[0][1])  # Kat20 Tabliye -> Ic Duvar
    link_typed(project, mep[0][1], mep[1][1], "SS", "5d")
    link_typed(project, mep[0][1], mep[2][1], "SS", "5d")
    link_typed(project, mep[2][1], mep[3][1], "SS", "5d")
    link_typed(project, mep[3][1], mep[4][1], "SS", "5d")
    if proc and len(proc) >= 5:
        link_fs(project, proc[4][1], mep[5][1])  # Asansor Uretim -> Kuyu
    link_fs(project, mep[5][1], mep[6][1])  # Kuyu -> Montaj
    link_typed(project, mep[1][1], mep[7][1], "FF", "5d")  # Tesisat FF Siva
    if proc and len(proc) >= 7:
        link_fs(project, proc[6][1], mep[8][1])  # Cephe Uretim -> Dis Cephe
    if len(sup) >= 20:
        link_fs(project, sup[19][1], mep[8][1])  # Kat10 Tabliye -> Dis Cephe
    link_fs(project, mep[7][1], mep[9][1])
    link_typed(project, mep[9][1], mep[10][1], "SS", "5d")
    link_fs(project, mep[10][1], mep[11][1])
    link_fs(project, mep[11][1], mep[12][1])
    link_fs(project, mep[8][1], mep[13][1])
    link_fs(project, mep[12][1], mep[14][1])
    link_typed(project, mep[13][1], mep[14][1], "FF")
    print(f"  Linked {len(mep)} MEP tasks")

    # =========================================================================
    # PHASE 6: COMMISSIONING (5 activities)
    # =========================================================================
    print("\n=== PHASE 6: COMMISSIONING (5) ===")
    comm = []
    comm_def = [
        ("MEP Sistem Testleri", "15d"),
        ("Asansor Test Sertifikasyon", "10d"),
        ("Yangin Tatbikati Onay", "8d"),
        ("Iskan Ruhsati Sureci", "20d"),
        ("Proje Teslimi ve Kapanisi", "5d"),
    ]
    p = phases["comm"][1]
    for nm, dr in comm_def:
        b, t = add_task_under(project, p, nm, dr)
        if t: comm.append((b, t, nm)); print(f"  [{b.ID}] {nm}")

    # Links
    if mep and comm:
        link_fs(project, mep[-1][1], comm[0][1])
    if len(mep) >= 7:
        link_fs(project, mep[6][1], comm[1][1])  # Asansor Montaj -> Test
    if len(mep) >= 5:
        link_fs(project, mep[4][1], comm[2][1])  # Yangin -> Tatbikat
    if len(comm) >= 4:
        link_fs(project, comm[0][1], comm[3][1])
        link_fs(project, comm[1][1], comm[3][1])
        link_fs(project, comm[2][1], comm[3][1])
    if len(comm) >= 5:
        link_fs(project, comm[3][1], comm[4][1])
    print(f"  Linked {len(comm)} commissioning tasks")

    # =========================================================================
    # RESCHEDULE
    # =========================================================================
    print("\n=== RESCHEDULE ===")
    project.Reschedule()
    project.WaitForNotificationProcessing()
    print(f"  Start: {project.ProjectStart}")
    print(f"  End:   {project.ProjectEnd}")

    total = len(eng) + len(proc) + len(sub) + len(sup) + len(mep) + len(comm)
    print(f"\n=== TOTAL: {total} activities created ===")
    print(f"  Eng={len(eng)} Proc={len(proc)} Sub={len(sub)} Super={len(sup)} MEP={len(mep)} Comm={len(comm)}")

    # Save bar IDs
    with open("gokdelen_ids.txt", "w", encoding="utf-8") as f:
        for phase, tasks in [("eng",eng),("proc",proc),("sub",sub),("super",sup),("mep",mep),("comm",comm)]:
            f.write(f"=== {phase} ===\n")
            for b, t, nm in tasks:
                f.write(f"{b.ID}\t{nm}\n")
    print("IDs saved to gokdelen_ids.txt")
    print("Done!")

if __name__ == "__main__":
    main()
