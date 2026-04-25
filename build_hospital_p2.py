"""
160.000 m2 SEHIR HASTANESI KOMPLEKSI — Part 2
===============================================
Code Libraries (3x) + Resources + Cost Centres + Assignments ($250M)
Uses bar_ids from Part 1 (hospital_bar_ids.json)
"""
import sys, os, traceback, json
sys.stdout.reconfigure(encoding='utf-8')
from datetime import datetime

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_hospital_p2_output.txt")
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

    # Load bar IDs from Part 1
    mapping_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hospital_bar_ids.json")
    with open(mapping_file, "r") as mf:
        all_bars = json.load(mf)
    log(f"Loaded {len(all_bars)} bar IDs from Part 1")

    def tx(name): project.StartTransaction(name)
    def end_tx():
        try:
            project.EndTransaction()
        except Exception as e:
            log(f"  [WARN] EndTx: {e}")
            try: project.AbandonTransaction()
            except: pass
        project.WaitForNotificationProcessing()

    def find_bar_by_id(target_id):
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
        bid = all_bars.get(code)
        if not bid: return None
        r = find_bar_by_id(int(bid))
        return r[1] if r else None

    def get_bar(code):
        bid = all_bars.get(code)
        if not bid: return None
        r = find_bar_by_id(int(bid))
        return r[0] if r else None

    def get_ac(obj, prop):
        did = obj._oleobj_.GetIDsOfNames(0, prop)
        raw = obj._oleobj_.InvokeTypes(did, 0, 2, (9, 0), ())
        return D(raw) if raw else None

    def set_amt(ac, val):
        ac._oleobj_.InvokeTypes(0, 0, 4, (24, 0), ((5, 1),), float(val))

    # ══════════════════════════════════════════════════
    # PHASE 5: CODE LIBRARIES (3 libraries)
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 5: Code Libraries")
    log("=" * 60)

    # Code Library definitions
    CODE_LIBS = {
        "Lokasyon/Zon": ["Blok-A", "Blok-B", "Blok-C", "Ortak Alan"],
        "Disiplin": ["Kaba Yapi", "Ince Yapi", "MEP", "Medikal Sistemler"],
        "Taseron": ["Ana Yuklenici", "Elektrik-Taseron", "Mekanik-Taseron", "Medikal-Vendor"],
    }

    code_entries = {}  # "LibName:EntryName" -> entry_id (for assignment later)

    for lib_name, entries in CODE_LIBS.items():
        tx(f"CL-{lib_name}")
        code_libs = project.CodeLibrarys
        cl = D(code_libs.Add())
        cl.Name = lib_name
        cl_id = cl.ID
        log(f"  Created Code Library: '{lib_name}' ID={cl_id}")

        for entry_name in entries:
            e = D(cl.Entries.Add())
            e.Name = entry_name
            code_entries[f"{lib_name}:{entry_name}"] = e.ID
            log(f"    Entry: '{entry_name}' ID={e.ID}")

        end_tx()

    log(f"  Total code entries: {len(code_entries)}")

    # ── Code Assignments ──
    log("\n  Assigning codes to activities...")

    # Lokasyon mapping
    LOK_MAP = {
        "Blok-A": ["K01","K02","K03","K04","K05","K06","K07","K08","K09","K10","K11","K12","K13",
                    "K31","K34","C01","C04","C07","C12","C16","M05","M08","M13","M18","M21","M25",
                    "M29","I01","I04","I07","I10","I13","I16","I22","I25","MC01","MC02","MC03",
                    "MC17","MC09","MC10","MC12"],
        "Blok-B": ["K14","K15","K16","K17","K18","K19","K20","K21","K32","K35","C02","C05","C08",
                    "C11","C13","C17","M06","M09","M14","M19","M22","M26","M30","I02","I05","I08",
                    "I11","I14","I23","I26","MC18","MC19"],
        "Blok-C": ["K22","K23","K24","K25","K26","K27","K28","K33","K36","C03","C06","C09","C14",
                    "C18","M07","M10","M15","M20","M23","M27","M31","I03","I06","I09","I12","I15",
                    "I17","I24","I27","MC04","MC21","MC22"],
        "Ortak Alan": ["D01","D02","D03","D04","D05","D06","D07","D08","D09","D10","D11","D12",
                       "D13","D14","D15","H01","H02","H03","H04","H05","H06","H07","H08","H09",
                       "H10","H11","H12","H13","H14","H15","H16","H17","H18","H19","H20",
                       "T01","T02","T03","T04","T05","T06","T07","T08","T09","T10","T11","T12",
                       "T13","T14","T15","T16","T17","T18","T19","T20","T21","T22","T23","T24","T25",
                       "K29","K30","K37","C10","C15","C19","C20","C21","C22",
                       "M01","M02","M03","M04","M11","M12","M16","M17","M24","M28","M32","M33",
                       "M34","M35","M36","M37","M38","M39","M40",
                       "I18","I19","I20","I21","I28","I29","I30",
                       "MC05","MC06","MC07","MC08","MC11","MC13","MC14","MC15","MC16","MC20",
                       "MC23","MC24","MC25",
                       "TC01","TC02","TC03","TC04","TC05","TC06","TC07","TC08","TC09","TC10",
                       "TC11","TC12","TC13","TC14","TC15","TC16","TC17","TC18","TC19","TC20",
                       "TC21","TC22","TC23","TC24","TC25","TC26","TC27","TC28","TC29","TC30",
                       "TC31","TC32","TC33","TC34","TC35","TC36"],
    }

    # Disiplin mapping
    DIS_MAP = {
        "Kaba Yapi": ["H01","H02","H03","H04","H05","H06","H07","H08","H09","H10","H11","H12",
                       "H13","H14","H15","H16","H17","H18","H19","H20",
                       "T01","T02","T03","T04","T05","T06","T07","T08","T09","T10","T11","T12",
                       "T13","T14","T15","T16","T17","T18","T19","T20","T21","T22","T23","T24","T25",
                       "K01","K02","K03","K04","K05","K06","K07","K08","K09","K10","K11","K12","K13",
                       "K14","K15","K16","K17","K18","K19","K20","K21",
                       "K22","K23","K24","K25","K26","K27","K28","K29","K30","K31","K32","K33",
                       "K34","K35","K36","K37",
                       "C01","C02","C03","C04","C05","C06","C12","C13","C14","C15"],
        "Ince Yapi": ["C07","C08","C09","C10","C11","C16","C17","C18","C19","C20","C21","C22",
                       "I01","I02","I03","I04","I05","I06","I07","I08","I09","I10","I11","I12",
                       "I13","I14","I15","I16","I17","I18","I19","I20","I21","I22","I23","I24",
                       "I25","I26","I27","I28","I29","I30"],
        "MEP": ["M01","M02","M03","M04","M05","M06","M07","M08","M09","M10","M11","M12","M13",
                "M14","M15","M16","M17","M18","M19","M20","M21","M22","M23","M24","M25","M26",
                "M27","M28","M29","M30","M31","M32","M33","M34","M35","M36","M37","M38","M39","M40",
                "TC01","TC02","TC03","TC04","TC05","TC06","TC07","TC08","TC09","TC10","TC11","TC12",
                "TC13","TC14","TC15","TC16"],
        "Medikal Sistemler": ["D05","D06","D14",
                               "MC01","MC02","MC03","MC04","MC05","MC06","MC07","MC08","MC09",
                               "MC10","MC11","MC12","MC13","MC14","MC15","MC16","MC17","MC18",
                               "MC19","MC20","MC21","MC22","MC23","MC24","MC25",
                               "TC17","TC18","TC19","TC20","TC21","TC22","TC23","TC24","TC25",
                               "TC26","TC27","TC28","TC29","TC30","TC31","TC32","TC33","TC34","TC35","TC36"],
    }

    # Taseron mapping
    TAS_MAP = {
        "Ana Yuklenici": ["D01","D02","D03","D04","D05","D06","D07","D08","D09","D10","D11","D12",
                          "D13","D14","D15",
                          "H01","H02","H03","H04","H05","H06","H07","H08","H09","H10","H11","H12",
                          "H13","H14","H15","H16","H17","H18","H19","H20",
                          "T01","T02","T03","T04","T05","T06","T07","T08","T09","T10","T11","T12",
                          "T13","T14","T15","T16","T17","T18","T19","T20","T21","T22","T23","T24","T25",
                          "K01","K02","K03","K04","K05","K06","K07","K08","K09","K10","K11","K12","K13",
                          "K14","K15","K16","K17","K18","K19","K20","K21",
                          "K22","K23","K24","K25","K26","K27","K28","K29","K30","K31","K32","K33",
                          "K34","K35","K36","K37",
                          "C01","C02","C03","C04","C05","C06","C07","C08","C09","C10","C11",
                          "C12","C13","C14","C15","C16","C17","C18","C20","C21","C22",
                          "I01","I02","I03","I04","I05","I06","I07","I08","I09",
                          "I22","I23","I24","I25","I26","I27","I30",
                          "TC24","TC25","TC26","TC27","TC28","TC29","TC30","TC31","TC32","TC33",
                          "TC34","TC35","TC36"],
        "Elektrik-Taseron": ["M01","M02","M03","M04","M05","M06","M07","M08","M09","M10","M11",
                             "M32","M33","M34","M35",
                             "TC01","TC02","TC03","TC04","TC12","TC13","TC14"],
        "Mekanik-Taseron": ["M12","M13","M14","M15","M16","M17","M18","M19","M20","M21","M22",
                            "M23","M24","M25","M26","M27","M28","M29","M30","M31","M36","M37",
                            "M38","M39","M40",
                            "I10","I11","I12","I13","I14","I15","I16","I17","I18","I19","I20","I21",
                            "I28","I29","C19",
                            "TC05","TC06","TC07","TC08","TC09","TC10","TC11","TC15","TC16"],
        "Medikal-Vendor": ["MC01","MC02","MC03","MC04","MC05","MC06","MC07","MC08","MC09","MC10",
                           "MC11","MC12","MC13","MC14","MC15","MC16","MC17","MC18","MC19","MC20",
                           "MC21","MC22","MC23","MC24","MC25",
                           "TC17","TC18","TC19","TC20","TC21","TC22","TC23"],
    }

    def assign_codes(lib_name, mapping):
        """Assign codes from a library to activities"""
        ok = 0
        fail = 0
        for entry_name, codes in mapping.items():
            key = f"{lib_name}:{entry_name}"
            entry_id = code_entries.get(key)
            if not entry_id:
                log(f"    Entry not found: {key}")
                continue

            for code in codes:
                bid = all_bars.get(code)
                if not bid:
                    continue
                tx(f"AC-{code}")
                try:
                    bar_obj = get_bar(code)
                    if bar_obj:
                        # Re-fetch entry
                        code_libs = project.CodeLibrarys
                        found_entry = None
                        for cli in range(1, code_libs.Count + 1):
                            cl = D(code_libs.Item(cli))
                            if cl.Name == lib_name:
                                entries = cl.Entries
                                for ei in range(1, entries.Count + 1):
                                    e = D(entries.Item(ei))
                                    if e.ID == entry_id:
                                        found_entry = e
                                        break
                                break
                        if found_entry:
                            bar_obj.AssignCode(found_entry, True)
                            ok += 1
                        else:
                            fail += 1
                    else:
                        fail += 1
                except Exception as e:
                    fail += 1
                end_tx()
        return ok, fail

    # Assign all 3 code libraries
    for lib_name, mapping in [("Lokasyon/Zon", LOK_MAP), ("Disiplin", DIS_MAP), ("Taseron", TAS_MAP)]:
        log(f"\n  Assigning '{lib_name}'...")
        ok, fail = assign_codes(lib_name, mapping)
        log(f"    OK={ok}, Failed={fail}")

    # ══════════════════════════════════════════════════
    # PHASE 6: RESOURCES & COST CENTRES ($250M)
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 6: Resources & Cost Centres ($250M Budget)")
    log("=" * 60)

    # Create Cost Centres
    tx("CC-Main")
    ccs = project.CostCentres
    main_cc = D(ccs.Add())
    main_cc.Name = "Hastane Genel Butcesi"
    main_cc_id = main_cc.ID
    log(f"  Main CC: '{main_cc.Name}' ID={main_cc_id}")
    end_tx()

    # Sub cost centres
    sub_cc_names = ["Iscilik", "Ekipman", "Medikal Sistemler"]
    sub_cc_ids = {}
    for ccn in sub_cc_names:
        tx(f"CC-{ccn}")
        ccs = project.CostCentres
        cc = D(ccs.Add())
        cc.Name = ccn
        sub_cc_ids[ccn] = cc.ID
        log(f"  Sub CC: '{ccn}' ID={cc.ID}")
        end_tx()

    # Create Resources
    # 1. Consumable: MRI/CT Cihazlari
    tx("Res-MRI")
    cons_res = project.ConsumableResources
    mri_res = D(cons_res.Add())
    mri_res.Name = "Ileri Teknoloji MRI ve CT Cihazlari"
    mri_res_id = mri_res.ID
    log(f"  Consumable: '{mri_res.Name}' ID={mri_res_id}")
    end_tx()

    # 2. Consumable: Ameliyathane Moduler Sistemleri
    tx("Res-Amel")
    cons_res = project.ConsumableResources
    amel_res = D(cons_res.Add())
    amel_res.Name = "Ameliyathane Moduler Sistemleri"
    amel_res_id = amel_res.ID
    log(f"  Consumable: '{amel_res.Name}' ID={amel_res_id}")
    end_tx()

    # 3. Consumable: Kursun Zirhlama Malzemesi
    tx("Res-Kursun")
    cons_res = project.ConsumableResources
    kursun_res = D(cons_res.Add())
    kursun_res.Name = "Kursun Zirhlama Malzemesi"
    kursun_res_id = kursun_res.ID
    log(f"  Consumable: '{kursun_res.Name}' ID={kursun_res_id}")
    end_tx()

    # 4. Permanent: Mekanik Tesisat Ekibi
    tx("Res-Mek")
    perm_res = project.PermanentResources
    mek_res = D(perm_res.Add())
    mek_res.Name = "Mekanik Tesisat Ekibi"
    mek_res_id = mek_res.ID
    log(f"  Permanent: '{mek_res.Name}' ID={mek_res_id}")
    end_tx()

    # Create Cost/Income Rate
    tx("Rate-Main")
    rates = project.CostAndIncomeRates
    rate = D(rates.Add())
    rate.Name = "Hastane Proje Orani"
    rate_id = rate.ID
    log(f"  Rate: '{rate.Name}' ID={rate_id}")
    end_tx()

    # ── Resource Curves ──
    log("\n  Finding resource curves...")
    res_curves = project.ResourceCurves
    bell_curve = None
    front_curve = None
    for i in range(1, res_curves.Count + 1):
        rc = D(res_curves.Item(i))
        rn = rc.Name.lower()
        if "bell" in rn:
            bell_curve = rc
            log(f"    Bell curve found: '{rc.Name}'")
        if "front" in rn:
            front_curve = rc
            log(f"    Front loaded found: '{rc.Name}'")

    # ── Assign resources and costs to activities ──
    log("\n  Assigning resources and costs...")

    # Strategy: Distribute $250M across all 250 activities
    # Major cost drivers:
    # - MRI/CT devices: 15 units × $1,550,000.50 = $23.25M (assigned to MC02-MC08)
    # - Ameliyathane modular: 30 units × $250,000 = $7.5M (assigned to MC09-MC11, I19-I21)
    # - Kursun zirhlama: 5000 m² × $450.50 = $2.25M (assigned to T13-T15, I28, C19)
    # - Mekanik tesisat: assigned to MEP activities
    # - Remaining via cost allocations

    # Cost per activity category (to reach $250M total)
    COST_MAP = {
        # WBS1: Design ~$8M
        "D01": 1200000, "D02": 900000, "D03": 600000, "D04": 500000, "D05": 800000,
        "D06": 400000, "D07": 500000, "D08": 300000, "D09": 200000, "D10": 500000,
        "D11": 800000, "D12": 500000, "D13": 500000, "D14": 400000, "D15": 300000,
        # WBS2: Excavation ~$15M
        "H01": 500000, "H02": 200000, "H03": 300000, "H04": 2500000, "H05": 2000000,
        "H06": 1500000, "H07": 1200000, "H08": 1000000, "H09": 1500000, "H10": 800000,
        "H11": 800000, "H12": 600000, "H13": 500000, "H14": 400000, "H15": 200000,
        "H16": 500000, "H17": 500000, "H18": 400000, "H19": 300000, "H20": 100000,
        # WBS3: Foundation ~$25M
        "T01": 400000, "T02": 300000, "T03": 250000, "T04": 800000, "T05": 600000,
        "T06": 500000, "T07": 3000000, "T08": 2500000, "T09": 2200000, "T10": 1800000,
        "T11": 1500000, "T12": 1200000, "T13": 1500000, "T14": 2000000, "T15": 1200000,
        "T16": 1500000, "T17": 1200000, "T18": 900000, "T19": 800000, "T20": 600000,
        "T21": 500000, "T22": 600000, "T23": 400000, "T24": 500000, "T25": 100000,
        # WBS4: Karkas ~$62M
        "K01": 2000000, "K02": 2500000, "K03": 2000000, "K04": 2500000,
        "K05": 2000000, "K06": 2500000, "K07": 2000000, "K08": 2500000,
        "K09": 2000000, "K10": 2500000, "K11": 5000000, "K12": 5000000,
        "K13": 3000000,
        "K14": 1500000, "K15": 1800000, "K16": 1500000, "K17": 1800000,
        "K18": 1500000, "K19": 1800000, "K20": 3000000, "K21": 2500000,
        "K22": 1200000, "K23": 1500000, "K24": 1200000, "K25": 1500000,
        "K26": 1200000, "K27": 1500000, "K28": 2000000,
        "K29": 2500000, "K30": 1500000, "K31": 1200000, "K32": 900000,
        "K33": 700000, "K34": 2000000, "K35": 1500000, "K36": 1200000, "K37": 100000,
        # WBS5: Cephe ~$18M
        "C01": 400000, "C02": 350000, "C03": 300000, "C04": 2500000, "C05": 1800000,
        "C06": 1200000, "C07": 2200000, "C08": 1600000, "C09": 1000000,
        "C10": 1800000, "C11": 1200000, "C12": 500000, "C13": 400000, "C14": 350000,
        "C15": 600000, "C16": 400000, "C17": 350000, "C18": 300000,
        "C19": 800000, "C20": 300000, "C21": 200000, "C22": 100000,
        # WBS6: MEP ~$55M
        "M01": 3500000, "M02": 2500000, "M03": 1800000, "M04": 1500000,
        "M05": 2000000, "M06": 1500000, "M07": 1200000,
        "M08": 1500000, "M09": 1200000, "M10": 900000,
        "M11": 1800000, "M12": 1200000, "M13": 1500000, "M14": 1200000, "M15": 900000,
        "M16": 2500000, "M17": 3000000, "M18": 2000000, "M19": 1500000, "M20": 1200000,
        "M21": 1800000, "M22": 1400000, "M23": 1000000,
        "M24": 2000000, "M25": 1500000, "M26": 1200000, "M27": 900000,
        "M28": 1800000, "M29": 3000000, "M30": 2000000, "M31": 1000000,
        "M32": 1500000, "M33": 800000, "M34": 600000, "M35": 900000,
        "M36": 2000000, "M37": 400000, "M38": 350000, "M39": 500000, "M40": 100000,
        # WBS7: Ince Is ~$22M
        "I01": 1500000, "I02": 1200000, "I03": 800000,
        "I04": 1200000, "I05": 900000, "I06": 700000,
        "I07": 1500000, "I08": 1200000, "I09": 800000,
        "I10": 600000, "I11": 500000, "I12": 400000,
        "I13": 500000, "I14": 400000, "I15": 350000,
        "I16": 800000, "I17": 600000,
        "I18": 1000000, "I19": 1500000, "I20": 1200000, "I21": 1800000,
        "I22": 400000, "I23": 350000, "I24": 300000,
        "I25": 500000, "I26": 400000, "I27": 350000,
        "I28": 1200000, "I29": 400000, "I30": 100000,
        # WBS8: Medikal Cihaz ~$40M
        "MC01": 1500000, "MC02": 5000000, "MC03": 6200000, "MC04": 7500000,
        "MC05": 3500000, "MC06": 2500000, "MC07": 1200000, "MC08": 600000,
        "MC09": 1800000, "MC10": 1200000, "MC11": 1500000, "MC12": 1800000,
        "MC13": 2000000, "MC14": 800000, "MC15": 500000, "MC16": 1200000,
        "MC17": 2000000, "MC18": 1200000, "MC19": 800000, "MC20": 600000,
        "MC21": 4500000, "MC22": 2500000, "MC23": 500000, "MC24": 600000, "MC25": 100000,
        # WBS9: Test & Commissioning ~$5M
        "TC01": 200000, "TC02": 100000, "TC03": 100000, "TC04": 150000, "TC05": 150000,
        "TC06": 100000, "TC07": 100000, "TC08": 150000, "TC09": 200000, "TC10": 200000,
        "TC11": 250000, "TC12": 200000, "TC13": 100000, "TC14": 150000, "TC15": 100000,
        "TC16": 100000, "TC17": 300000, "TC18": 200000, "TC19": 150000, "TC20": 200000,
        "TC21": 100000, "TC22": 100000, "TC23": 100000, "TC24": 100000, "TC25": 150000,
        "TC26": 100000, "TC27": 200000, "TC28": 150000, "TC29": 100000,
        "TC30": 500000, "TC31": 100000, "TC32": 200000, "TC33": 150000,
        "TC34": 100000, "TC35": 200000, "TC36": 50000,
    }

    total_budget = sum(COST_MAP.values())
    log(f"  Planned budget: ${total_budget:,.0f}")

    # Assign costs via CostAllocation
    cost_ok = 0
    cost_fail = 0
    for code, amount in COST_MAP.items():
        bid = all_bars.get(code)
        if not bid:
            cost_fail += 1
            continue
        tx(f"Cost-{code}")
        try:
            task = get_task(code)
            if task:
                # Find cost centre
                ccs = project.CostCentres
                cc_obj = None
                for ci in range(1, ccs.Count + 1):
                    cc = D(ccs.Item(ci))
                    if cc.ID == main_cc_id:
                        cc_obj = cc
                        break
                if cc_obj:
                    ca = D(task.AssignCost(cc_obj))
                    # Set the cost amount
                    ac = get_ac(ca, "GivenCost")
                    if ac:
                        set_amt(ac, amount)
                    cost_ok += 1
                else:
                    cost_fail += 1
            else:
                cost_fail += 1
        except Exception as e:
            cost_fail += 1
            if cost_ok < 5:  # only log first few errors
                log(f"    Cost error {code}: {e}")
        end_tx()

        if cost_ok % 50 == 0 and cost_ok > 0:
            log(f"    Costs assigned: {cost_ok}...")

    log(f"  Cost assignments: OK={cost_ok}, Failed={cost_fail}")

    # ── Assign specific consumable resources ──
    log("\n  Assigning consumable resources...")

    # MRI/CT to MC02-MC08 (15 units total, rate $1,550,000.50)
    MRI_ASSIGNMENTS = [
        ("MC02", 1), ("MC03", 2), ("MC04", 3), ("MC05", 2),
        ("MC06", 5), ("MC07", 1), ("MC08", 1),
    ]
    for code, units in MRI_ASSIGNMENTS:
        tx(f"Res-MRI-{code}")
        try:
            task = get_task(code)
            cons_res = project.ConsumableResources
            mri = None
            for ri in range(1, cons_res.Count + 1):
                r = D(cons_res.Item(ri))
                if r.ID == mri_res_id:
                    mri = r
                    break
            if task and mri:
                alloc = D(task.AssignConsumableResource(mri, False, None, None))
                alloc.GivenAllocation = float(units)
                log(f"    MRI/CT: {code} x{units}")
        except Exception as e:
            log(f"    MRI error {code}: {e}")
        end_tx()

    # Ameliyathane Moduler to MC09-MC11, I19-I21 (30 units, $250,000)
    AMEL_ASSIGNMENTS = [
        ("MC09", 6), ("MC10", 5), ("MC11", 5), ("I19", 5), ("I20", 4), ("I21", 5),
    ]
    for code, units in AMEL_ASSIGNMENTS:
        tx(f"Res-Amel-{code}")
        try:
            task = get_task(code)
            cons_res = project.ConsumableResources
            amel = None
            for ri in range(1, cons_res.Count + 1):
                r = D(cons_res.Item(ri))
                if r.ID == amel_res_id:
                    amel = r
                    break
            if task and amel:
                alloc = D(task.AssignConsumableResource(amel, False, None, None))
                alloc.GivenAllocation = float(units)
                log(f"    Ameliyathane: {code} x{units}")
        except Exception as e:
            log(f"    Amel error {code}: {e}")
        end_tx()

    # Kursun Zirhlama to T13, T14, T15, I28, C19 (5000 m², $450.50/m²)
    KURSUN_ASSIGNMENTS = [
        ("T13", 1200), ("T14", 1000), ("T15", 800), ("I28", 1200), ("C19", 800),
    ]
    for code, sqm in KURSUN_ASSIGNMENTS:
        tx(f"Res-Kur-{code}")
        try:
            task = get_task(code)
            cons_res = project.ConsumableResources
            kur = None
            for ri in range(1, cons_res.Count + 1):
                r = D(cons_res.Item(ri))
                if r.ID == kursun_res_id:
                    kur = r
                    break
            if task and kur:
                alloc = D(task.AssignConsumableResource(kur, False, None, None))
                alloc.GivenAllocation = float(sqm)
                log(f"    Kursun: {code} x{sqm}m2")
        except Exception as e:
            log(f"    Kursun error {code}: {e}")
        end_tx()

    # Mekanik Tesisat Ekibi (permanent) to MEP activities with bell/front curves
    MEK_ASSIGNMENTS = [
        ("M12", 5, "bell"), ("M13", 8, "bell"), ("M14", 6, "bell"), ("M15", 5, "bell"),
        ("M16", 10, "front"), ("M17", 8, "front"), ("M18", 12, "bell"), ("M19", 10, "bell"),
        ("M20", 8, "bell"), ("M21", 10, "bell"), ("M22", 8, "bell"), ("M23", 6, "bell"),
        ("M24", 6, "front"), ("M25", 8, "front"), ("M26", 6, "front"), ("M27", 5, "front"),
        ("M36", 10, "bell"), ("M37", 4, "front"), ("M38", 4, "front"), ("M39", 6, "bell"),
    ]
    for code, alloc_count, curve_type in MEK_ASSIGNMENTS:
        tx(f"Res-Mek-{code}")
        try:
            task = get_task(code)
            perm_res = project.PermanentResources
            mek = None
            for ri in range(1, perm_res.Count + 1):
                r = D(perm_res.Item(ri))
                if r.ID == mek_res_id:
                    mek = r
                    break
            if task and mek:
                alloc = D(task.AssignResource(mek, False))
                alloc.GivenAllocation = float(alloc_count)
                # Assign resource curve
                curve = bell_curve if curve_type == "bell" else front_curve
                if curve:
                    alloc.ResourceCurve = curve
                log(f"    Mekanik: {code} x{alloc_count} ({curve_type})")
        except Exception as e:
            log(f"    Mek error {code}: {e}")
        end_tx()

    # ── Final reschedule and save ──
    log("\n  Rescheduling...")
    project.Reschedule()
    log("  Reschedule complete!")

    project.Save()
    log("  Project saved!")

    log("\n" + "=" * 60)
    log("SUMMARY")
    log("=" * 60)
    log(f"  Code Libraries: 3 ({len(code_entries)} entries)")
    log(f"  Cost Centres: 4 (1 parent + 3 sub)")
    log(f"  Resources: 4 (1 permanent + 3 consumable)")
    log(f"  Cost Assignments: {cost_ok}")
    log(f"  Total Budget: ${total_budget:,.0f}")
    log("\nPart 2 DONE!")

except Exception as e:
    log(f"FATAL ERROR: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output: {OUT}")
