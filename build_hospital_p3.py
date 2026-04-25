"""
160.000 m2 SEHIR HASTANESI KOMPLEKSI — Part 3
===============================================
Baseline + Progress (6 months) + Variance Analysis + EVM
"""
import sys, os, traceback, json
sys.stdout.reconfigure(encoding='utf-8')
from datetime import datetime, timedelta

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_hospital_p3_output.txt")
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

    mapping_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hospital_bar_ids.json")
    with open(mapping_file, "r") as mf:
        all_bars = json.load(mf)
    log(f"Loaded {len(all_bars)} bar IDs")

    def pt(dt_str):
        return pywintypes.Time(datetime.strptime(dt_str, "%Y-%m-%d"))

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

    def set_progress(bar_obj, pct):
        try:
            bar_obj.DurationPercentComplete = float(pct)
            return True
        except: pass
        try:
            bar_obj.OverallPercentComplete = float(pct)
            return True
        except: pass
        try:
            did = bar_obj._oleobj_.GetIDsOfNames(0, 'DurationPercentComplete')
            bar_obj._oleobj_.InvokeTypes(did, 0, 4, (24, 0), ((5, 1),), float(pct))
            return True
        except:
            return False

    def safe_date(obj, attr):
        try:
            v = getattr(obj, attr)
            if v: return str(v)[:10]
        except: pass
        return "N/A"

    def get_ac(obj, prop):
        did = obj._oleobj_.GetIDsOfNames(0, prop)
        raw = obj._oleobj_.InvokeTypes(did, 0, 2, (9, 0), ())
        return D(raw) if raw else None

    def get_amt(ac):
        try:
            return ac._oleobj_.InvokeTypes(0, 0, 2, (5, 0), ())
        except:
            return 0.0

    # ══════════════════════════════════════════════════
    # PHASE 7: BASELINE
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 7: Create Baseline")
    log("=" * 60)

    # Step 1: Save project as .pp file
    baseline_path = r"C:\Users\CahAsus\Documents\Powerproject\Projects\Hastane_Hedef_Program_Rev00.pp"
    try:
        project.SaveAs(baseline_path)
        log(f"  SaveAs: {baseline_path}")
    except Exception as e:
        log(f"  SaveAs error: {e}")

    # Step 2: Import as baseline
    try:
        result = project.ImportBaseline(0, "", baseline_path, 0, "Hastane_Hedef_Program_Rev00")
        log(f"  ImportBaseline result: {result}")
    except Exception as e:
        log(f"  ImportBaseline error: {e}")

    # Step 3: Open baseline
    try:
        bsumm = project.BaselineSummaries
        log(f"  BaselineSummaries count: {bsumm.Count}")
        if bsumm.Count > 0:
            bs = D(bsumm.Item(1))
            bid = bs.ID
            log(f"  Baseline: '{bs.Name}' ID={bid}")
            project.OpenBaseline(bid, 0)
            log(f"  OpenBaseline OK!")
    except Exception as e:
        log(f"  Baseline open error: {e}")

    # ══════════════════════════════════════════════════
    # PHASE 8: PROGRESS (Data Date: September 2026 = 6 months)
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 8: Progress Update (6 months - Sept 2026)")
    log("=" * 60)

    # Activities to mark 100% complete (Hafriyat + Zemin + Temel)
    COMPLETE_100 = [
        # WBS1: All design complete
        "D01", "D02", "D03", "D04", "D05", "D06", "D07", "D08", "D09", "D10",
        "D11", "D12", "D13", "D14", "D15",
        # WBS2: All excavation complete
        "H01", "H02", "H03", "H04", "H05", "H06", "H07", "H08", "H09", "H10",
        "H11", "H12", "H13", "H14", "H15", "H16", "H17", "H18", "H19", "H20",
        # WBS3: Foundation mostly complete
        "T01", "T02", "T03", "T04", "T05", "T06", "T07", "T08", "T09", "T10",
        "T11", "T12", "T13", "T14", "T15", "T16", "T17", "T18", "T19", "T20",
        "T21", "T22", "T23", "T24", "T25",
    ]

    # Partial progress on karkas (started)
    PARTIAL_PROGRESS = {
        "K01": 100, "K02": 100, "K03": 100, "K04": 100,
        "K05": 80, "K06": 50,
        "K14": 100, "K15": 100, "K16": 80,
        "K22": 100, "K23": 80,
    }

    # CRITICAL: K07 (Blok A 3. Kat Kolon) - 15 gun gec basladi, %20'de
    # This is on the critical path!
    DELAYED_TASK = "K07"  # Blok A 3. Kat Kolon - should be further along but delayed

    prog_ok = 0
    prog_fail = 0

    # Set 100% complete activities
    for code in COMPLETE_100:
        tx(f"Prog-{code}")
        try:
            bar = get_bar(code)
            if bar:
                if set_progress(bar, 100):
                    prog_ok += 1
                else:
                    prog_fail += 1
            else:
                prog_fail += 1
        except:
            prog_fail += 1
        end_tx()

    log(f"  100% complete: OK={prog_ok}")

    # Set partial progress
    for code, pct in PARTIAL_PROGRESS.items():
        tx(f"PProg-{code}")
        try:
            bar = get_bar(code)
            if bar:
                set_progress(bar, pct)
                prog_ok += 1
        except:
            prog_fail += 1
        end_tx()

    log(f"  Partial progress set")

    # CRITICAL: Delayed task K07 - set to 20%
    tx(f"Delay-{DELAYED_TASK}")
    try:
        bar = get_bar(DELAYED_TASK)
        if bar:
            # Set low progress to simulate 15-day delay
            set_progress(bar, 20)
            log(f"  DELAYED: {DELAYED_TASK} set to 20% (15-day delay on Critical Path)")
    except Exception as e:
        log(f"  Delay error: {e}")
    end_tx()

    # Reschedule with updated progress
    project.Reschedule()
    log("  Reschedule after progress OK!")

    project.Save()
    log("  Saved!")

    # ══════════════════════════════════════════════════
    # PHASE 9: VARIANCE ANALYSIS
    # ══════════════════════════════════════════════════
    log("\n" + "=" * 60)
    log("PHASE 9: Variance & EVM Analysis")
    log("=" * 60)

    # Collect data for analysis
    total_activities = 0
    completed_activities = 0
    total_cost = 0.0
    actual_cost = 0.0

    # Check all activities
    all_activity_codes = []
    for wbs_code in ["WBS1","WBS2","WBS3","WBS4","WBS5","WBS6","WBS7","WBS8","WBS9"]:
        for key in all_bars:
            if key.startswith(("D","H","T","K","C","M","I","MC","TC")) and key not in ["PROJ"] + [f"WBS{i}" for i in range(1,10)]:
                if key not in all_activity_codes:
                    all_activity_codes.append(key)

    log(f"\n  Analyzing {len(all_activity_codes)} activities...")

    for code in all_activity_codes:
        r = find_bar_by_id(int(all_bars[code]))
        if not r: continue
        bar, task = r
        total_activities += 1

        try:
            pct = bar.DurationPercentComplete
            if pct >= 100:
                completed_activities += 1
        except:
            pass

        # Try to get cost info
        try:
            allocs = task.Allocations
            for ai in range(1, allocs.Count + 1):
                a = D(allocs.Item(ai))
                try:
                    gc = get_ac(a, "GivenCost")
                    if gc:
                        c = get_amt(gc)
                        total_cost += c if c else 0
                except: pass
                try:
                    ac_obj = get_ac(a, "ActualCost")
                    if ac_obj:
                        ac_val = get_amt(ac_obj)
                        actual_cost += ac_val if ac_val else 0
                except: pass
        except: pass

    # Get project dates
    rb = D(project.Bars.Item(1))
    rt = D(rb.ExpandedTask)
    proj_start = safe_date(rt, "Start")
    proj_finish = safe_date(rt, "Finish")

    # Check delayed task
    delayed_bar = get_bar(DELAYED_TASK)
    delayed_task = get_task(DELAYED_TASK)
    delayed_start = safe_date(delayed_task, "Start") if delayed_task else "N/A"
    delayed_finish = safe_date(delayed_task, "Finish") if delayed_task else "N/A"
    delayed_pct = 0
    try:
        delayed_pct = delayed_bar.DurationPercentComplete
    except: pass

    # EVM Calculations (simplified)
    # BAC = total planned cost = $250M
    BAC = 250000000.0
    # At 6 months into 30 months: planned progress = 20%
    # But we completed 60+ activities out of 250 (Hafriyat, Zemin, Temel)
    planned_pct = (completed_activities / total_activities * 100) if total_activities > 0 else 0
    # Weighted by activities completed (60 design + excavation + foundation out of 250)
    # That's about 24% of activities
    # EVM metrics
    PV = BAC * 0.20  # 20% of budget should be spent at month 6
    EV = BAC * (completed_activities / total_activities) if total_activities > 0 else 0
    AC = actual_cost if actual_cost > 0 else EV * 1.05  # estimate 5% over if no actual data

    SPI = EV / PV if PV > 0 else 0
    CPI = EV / AC if AC > 0 else 0
    SV = EV - PV
    CV = EV - AC
    EAC = BAC / CPI if CPI > 0 else BAC
    ETC = EAC - AC
    VAC = BAC - EAC
    TCPI = (BAC - EV) / (BAC - AC) if (BAC - AC) > 0 else 0

    # 15-day delay impact on project finish
    # Critical path delay: Blok A 3. Kat Kolon delayed by 15 days
    # This cascades through all subsequent Blok A activities

    log(f"\n  === PROJECT SUMMARY ===")
    log(f"  Project Start: {proj_start}")
    log(f"  Project Finish: {proj_finish}")
    log(f"  Total Activities: {total_activities}")
    log(f"  Completed (100%): {completed_activities}")
    log(f"  Completion Rate: {planned_pct:.1f}%")
    log(f"  Data Date: September 2026 (Month 6)")

    log(f"\n  === DELAY ANALYSIS ===")
    log(f"  Delayed Activity: K07 - Blok A 3. Kat Kolon")
    log(f"  Delay: 15 gun (Critical Path)")
    log(f"  Current Progress: {delayed_pct}%")
    log(f"  Task Start: {delayed_start}")
    log(f"  Task Finish: {delayed_finish}")
    log(f"  Impact: 15+ days delay on project finish")
    log(f"  Affected Chain: K07->K08->K09->K10->K11->K12->K13->Cephe->MEP->Ince Is->Medikal->Commissioning")

    log(f"\n  === EVM ANALYSIS ===")
    log(f"  BAC (Budget at Completion): ${BAC:,.0f}")
    log(f"  PV  (Planned Value):        ${PV:,.0f}")
    log(f"  EV  (Earned Value):         ${EV:,.0f}")
    log(f"  AC  (Actual Cost):          ${AC:,.0f}")
    log(f"  SV  (Schedule Variance):    ${SV:,.0f} ({'AHEAD' if SV >= 0 else 'BEHIND'})")
    log(f"  CV  (Cost Variance):        ${CV:,.0f} ({'UNDER' if CV >= 0 else 'OVER'})")
    log(f"  SPI (Schedule Perf Index):  {SPI:.3f} ({'Good' if SPI >= 1 else 'Behind Schedule'})")
    log(f"  CPI (Cost Perf Index):      {CPI:.3f} ({'Good' if CPI >= 1 else 'Over Budget'})")
    log(f"  EAC (Estimate at Compl):    ${EAC:,.0f}")
    log(f"  ETC (Estimate to Compl):    ${ETC:,.0f}")
    log(f"  VAC (Variance at Compl):    ${VAC:,.0f}")
    log(f"  TCPI (To-Complete Perf):    {TCPI:.3f}")

    log(f"\n  === COST BREAKDOWN ===")
    log(f"  Planned Total (GivenCost):  ${total_cost:,.0f}")
    log(f"  Actual Cost:                ${actual_cost:,.0f}")

    project.Save()
    log("\n  Project saved!")
    log("\nPart 3 DONE!")

except Exception as e:
    log(f"FATAL ERROR: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output: {OUT}")
