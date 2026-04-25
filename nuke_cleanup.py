"""Aggressive cleanup: remove ALL project bars including stubborn ones."""
import sys, os, traceback
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nuke_cleanup_output.txt")
f = open(OUT, "w", encoding="utf-8")
def log(msg=""): f.write(str(msg) + "\n"); f.flush()

try:
    import pythoncom, pywintypes, win32com.client
    D = win32com.client.Dispatch
    CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
    pythoncom.CoInitialize()
    app = D(pythoncom.GetActiveObject(CLSID).QueryInterface(pythoncom.IID_IDispatch))
    project = app.ActiveProject
    log(f"Project: {project.Name}")

    def tx(n): project.StartTransaction(n)
    def end_tx():
        try:
            project.EndTransaction()
        except Exception as e:
            log(f"  [tx-warn]: {e}")
            try: project.AbandonTransaction()
            except: pass
        project.WaitForNotificationProcessing()

    # Get root
    bars = project.Bars
    root_bar = D(bars.Item(1))
    root_task = D(root_bar.ExpandedTask)
    log(f"Root: {root_task.Name}, ChildBars: {root_task.ChildBars.Count}")

    # List all children
    for i in range(1, root_task.ChildBars.Count + 1):
        cb = D(root_task.ChildBars.Item(i))
        log(f"  [{i}] ID={cb.ID} '{cb.Name}'")

    # Deep recursive cleanup: bottom-up
    def deep_remove_all(parent_task, depth=0):
        """Remove all children of parent_task, deepest first."""
        try:
            cbs = parent_task.ChildBars
            count = cbs.Count
        except:
            return

        for i in range(count, 0, -1):
            try:
                cb = D(cbs.Item(i))
                ct = None
                try:
                    if cb.Tasks.Count > 0:
                        ct = D(cb.Tasks(1))
                except:
                    pass

                # First recurse into children
                if ct:
                    try:
                        if ct.ChildBars.Count > 0:
                            deep_remove_all(ct, depth + 1)
                    except:
                        pass

                # Now remove links from this task
                if ct:
                    try:
                        lo = ct.LinksOut
                        while lo.Count > 0:
                            tx("rl")
                            lo.Remove(1)
                            end_tx()
                    except:
                        try: project.EndTransaction()
                        except: pass
                    try:
                        li = ct.LinksIn
                        while li.Count > 0:
                            tx("rl")
                            li.Remove(1)
                            end_tx()
                    except:
                        try: project.EndTransaction()
                        except: pass

                    # Clear progress
                    try:
                        tx("cp")
                        ct.OverallPercentComplete = 0
                        end_tx()
                    except:
                        try: project.EndTransaction()
                        except: pass

                    # Remove allocations
                    try:
                        allocs = ct.Allocations
                        while allocs.Count > 0:
                            tx("ra")
                            allocs.Remove(1)
                            end_tx()
                    except:
                        try: project.EndTransaction()
                        except: pass

                    # Remove task
                    try:
                        tx("rt")
                        cb.Tasks.Remove(1)
                        end_tx()
                    except:
                        try: project.EndTransaction()
                        except: pass

                # Remove bar
                bid = cb.ID
                name = cb.Name
                tx(f"rb-{bid}")
                cbs.Remove(i)
                end_tx()
                prefix = "  " * depth
                log(f"{prefix}  Removed ID={bid} '{name}'")

            except Exception as e:
                log(f"  Error at [{i}]: {e}")
                try: project.EndTransaction()
                except: pass

    log("\nStarting deep cleanup...")
    deep_remove_all(root_task)

    # Also remove any code libraries, resources, cost centres we created
    log("\nCleaning resources/codes...")

    # Code libraries
    code_libs = project.CodeLibrarys
    for i in range(code_libs.Count, 0, -1):
        try:
            cl = D(code_libs.Item(i))
            name = cl.Name
            if name in ("Lokasyon", "Disiplin", "Risk Seviyesi"):
                tx("rcl")
                code_libs.Remove(i)
                end_tx()
                log(f"  Removed code lib: {name}")
        except:
            try: project.EndTransaction()
            except: pass

    # Consumable resources
    cons = project.ConsumableResources
    for i in range(cons.Count, 0, -1):
        try:
            r = D(cons.Item(i))
            name = r.Name
            if "Titanyum" in name or "C100" in name or "Nükleer" in name:
                tx("rc")
                cons.Remove(i)
                end_tx()
                log(f"  Removed consumable: {name}")
        except:
            try: project.EndTransaction()
            except: pass

    # Permanent resources
    perm = project.PermanentResources
    for i in range(perm.Count, 0, -1):
        try:
            r = D(perm.Item(i))
            name = r.Name
            if "Nükleer" in name:
                tx("rp")
                perm.Remove(i)
                end_tx()
                log(f"  Removed permanent: {name}")
        except:
            try: project.EndTransaction()
            except: pass

    # Cost centres
    ccs = project.CostCentres
    for i in range(ccs.Count, 0, -1):
        try:
            cc = D(ccs.Item(i))
            name = cc.Name
            if name in ("Nükleer Genel Bütçe", "A-Yüksek Teknoloji Ekipman", "B-Radyasyon Korumalı İşçilik"):
                tx("rcc")
                ccs.Remove(i)
                end_tx()
                log(f"  Removed cost centre: {name}")
        except:
            try: project.EndTransaction()
            except: pass

    # CostAndIncomeRates
    rates = project.CostAndIncomeRates
    for i in range(rates.Count, 0, -1):
        try:
            r = D(rates.Item(i))
            name = r.Name
            if "Nükleer" in name:
                tx("rr")
                rates.Remove(i)
                end_tx()
                log(f"  Removed rate: {name}")
        except:
            try: project.EndTransaction()
            except: pass

    # Final check
    remaining = root_task.ChildBars.Count
    log(f"\nFinal: {remaining} child bars remaining")
    for i in range(1, remaining + 1):
        cb = D(root_task.ChildBars.Item(i))
        log(f"  [{i}] ID={cb.ID} '{cb.Name}'")

    log("\nNUKE CLEANUP DONE!")

except Exception as e:
    log(f"FATAL: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output: {OUT}")
