"""
COM Explorer v17 — CORRECTED hierarchy navigation
KEY INSIGHT:
  - bar.Tasks(1) for SUMMARY bars → returns IExpandedTask with correct ChildBars
  - bar.Tasks(1) for LEAF bars → returns ITask (no ChildBars)
  - bar.ExpandedTask → returns IExpandedTask but ChildBars may differ
  - For NEW bars: bar.Tasks(1) fails ("Item does not exist")
                  bar.ExpandedTask works

STRATEGY:
  - Navigate hierarchy via bar.Tasks(1) on existing bars
  - Access new bars via bar.ExpandedTask
  - Use ExpandedTask for ImposedStart/ImposedEnd (only on IExpandedTask)
  - SetUserDuration, LinkTo, LinksIn/LinksOut work on both ITask and IExpandedTask
"""
import pythoncom
import win32com.client
import pywintypes
from datetime import datetime
import traceback


def connect():
    pythoncom.CoInitialize()
    clsid = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
    app = win32com.client.GetActiveObject(clsid)
    project = app.ActiveProject
    print(f"Connected to: {project.Name}")
    return app, project


def wait(project):
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass


def get_root_task(project):
    """Get root IExpandedTask via bar.Tasks(1)."""
    bar = project.Bars.Item(1)
    return win32com.client.Dispatch(bar.Tasks(1))


def find_bar_by_id(project, bar_id, max_depth=5):
    """Find bar by ID traversing hierarchy via Tasks(1).ChildBars."""
    root_bar = project.Bars.Item(1)
    if root_bar.ID == bar_id:
        return root_bar
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    return _search_bar_via_tasks(root_task, bar_id, 0, max_depth)


def _search_bar_via_tasks(parent_task, target_id, depth, max_depth):
    """Recursive bar search using Tasks(1).ChildBars."""
    try:
        child_bars = parent_task.ChildBars
        for i in range(1, child_bars.Count + 1):
            cb = win32com.client.Dispatch(child_bars.Item(i))
            if cb.ID == target_id:
                return cb
            if depth < max_depth:
                try:
                    ct = win32com.client.Dispatch(cb.Tasks(1))
                    result = _search_bar_via_tasks(ct, target_id, depth + 1, max_depth)
                    if result:
                        return result
                except Exception:
                    # Leaf bar — Tasks(1) may fail, skip
                    pass
    except Exception:
        pass
    return None


def test_childbars_comparison(project):
    """Compare ChildBars from Tasks(1) vs ExpandedTask."""
    print("\n" + "=" * 80)
    print("TEST 0: ChildBars — Tasks(1) vs ExpandedTask")
    print("=" * 80)

    root_bar = project.Bars.Item(1)

    # Via Tasks(1)
    task = win32com.client.Dispatch(root_bar.Tasks(1))
    print(f"bar.Tasks(1) type: {type(task).__name__}")
    try:
        cb1 = task.ChildBars
        print(f"  ChildBars.Count = {cb1.Count}")
        for i in range(1, cb1.Count + 1):
            b = win32com.client.Dispatch(cb1.Item(i))
            print(f"    [{i}] ID={b.ID}, Name={b.Name[:30]}")
    except Exception as e:
        print(f"  ChildBars error: {e}")

    # Via ExpandedTask
    et = root_bar.ExpandedTask
    print(f"\nbar.ExpandedTask type: {type(et).__name__}")
    try:
        cb2 = et.ChildBars
        print(f"  ChildBars.Count = {cb2.Count}")
        for i in range(1, cb2.Count + 1):
            b = win32com.client.Dispatch(cb2.Item(i))
            print(f"    [{i}] ID={b.ID}, Name={b.Name[:30]}")
    except Exception as e:
        print(f"  ChildBars error: {e}")


def test_full_workflow(project):
    """Create 2 bars, set duration/dates, link them."""
    print("\n" + "=" * 80)
    print("TEST 1: FULL WORKFLOW")
    print("=" * 80)

    root_task = get_root_task(project)
    child_bars = root_task.ChildBars

    print(f"Top-level summaries ({child_bars.Count}):")
    parent_bar = None
    for i in range(1, child_bars.Count + 1):
        cb = win32com.client.Dispatch(child_bars.Item(i))
        ct = win32com.client.Dispatch(cb.Tasks(1))
        cc = 0
        try:
            cc = ct.ChildBars.Count
        except Exception:
            pass
        print(f"  [{i}] ID={cb.ID}, Name={cb.Name[:40]}, Children={cc}")
        if cc > 2 and not parent_bar:
            parent_bar = cb

    if not parent_bar:
        print("  No suitable parent!")
        return []

    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    print(f"\nUsing parent: ID={parent_bar.ID}, Name={parent_bar.Name[:40]}")
    print(f"  ChildBars: {parent_task.ChildBars.Count}")

    # === Create Bar A ===
    print(f"\n--- Step 1: Create Bar A ---")
    project.StartTransaction("Create A")
    try:
        new_a = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_a.Name = "V17_TEST_A"
        bar_a_id = new_a.ID
        # Use ExpandedTask for newly created bar
        et_a = new_a.ExpandedTask
        print(f"  BarID={bar_a_id}, ExpandedTask.ID={et_a.ID}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return []

    # === Create Bar B ===
    print(f"\n--- Step 2: Create Bar B ---")
    project.StartTransaction("Create B")
    try:
        parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
        new_b = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_b.Name = "V17_TEST_B"
        bar_b_id = new_b.ID
        print(f"  BarID={bar_b_id}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return [bar_a_id]

    # === Set Durations ===
    print(f"\n--- Step 3: Set Durations ---")
    project.StartTransaction("Durations")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        et_a = bar_a.ExpandedTask
        dur_a = et_a.GetDurationFromString("10d")
        et_a.SetUserDuration(dur_a)
        print(f"  A: SetUserDuration(10d) => OK")

        bar_b = find_bar_by_id(project, bar_b_id)
        et_b = bar_b.ExpandedTask
        dur_b = et_b.GetDurationFromString("5d")
        et_b.SetUserDuration(dur_b)
        print(f"  B: SetUserDuration(5d) => OK")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    bar_a = find_bar_by_id(project, bar_a_id)
    bar_b = find_bar_by_id(project, bar_b_id)
    et_a = bar_a.ExpandedTask
    et_b = bar_b.ExpandedTask
    print(f"  A: Dur={et_a.GetUserDuration().Hours}h, Start={et_a.Start}, End={et_a.End}")
    print(f"  B: Dur={et_b.GetUserDuration().Hours}h, Start={et_b.Start}, End={et_b.End}")

    # === Set Date on A (ImposedStart) ===
    print(f"\n--- Step 4: ImposedStart on A ---")
    project.StartTransaction("ImposedStart")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        et_a = bar_a.ExpandedTask
        et_a.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
        print(f"  ImposedStart = 2026-07-01 => OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    bar_a = find_bar_by_id(project, bar_a_id)
    et_a = bar_a.ExpandedTask
    print(f"  A: Start={et_a.Start}, End={et_a.End}, Constraint={et_a.Constraint}")

    # === Link A -> B ===
    print(f"\n--- Step 5: LinkTo A -> B ---")
    project.StartTransaction("LinkTo")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        bar_b = find_bar_by_id(project, bar_b_id)
        et_a = bar_a.ExpandedTask
        et_b = bar_b.ExpandedTask
        link = et_a.LinkTo(et_b)
        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  LINK CREATED! ID={ld.ID}, type={ld.type}")
        else:
            print(f"  LinkTo returned None!")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # === Verify ===
    print(f"\n--- Step 6: VERIFY ---")
    bar_a = find_bar_by_id(project, bar_a_id)
    bar_b = find_bar_by_id(project, bar_b_id)
    et_a = bar_a.ExpandedTask
    et_b = bar_b.ExpandedTask
    print(f"  A: Start={et_a.Start}, End={et_a.End}")
    print(f"     Dur={et_a.GetUserDuration().Hours}h, LinksOut={et_a.LinksOut.Count}, Constraint={et_a.Constraint}")
    print(f"  B: Start={et_b.Start}, End={et_b.End}")
    print(f"     Dur={et_b.GetUserDuration().Hours}h, LinksIn={et_b.LinksIn.Count}, Constraint={et_b.Constraint}")

    if et_b.Start >= et_a.End:
        print(f"\n  *** SUCCESS: B starts at/after A ends (FS link working!) ***")
    else:
        print(f"\n  WARNING: B starts before A ends!")

    return [bar_a_id, bar_b_id]


def test_link_operations(project, bar_ids):
    """Test link type/lag changes."""
    print("\n" + "=" * 80)
    print("TEST 2: LINK OPERATIONS")
    print("=" * 80)

    if len(bar_ids) < 2:
        return

    bar_a_id, bar_b_id = bar_ids

    # === Link type changes ===
    for type_val, label in [(1, "SS"), (2, "FF"), (0, "FS")]:
        print(f"\n--- Set type={type_val} ({label}) ---")
        project.StartTransaction(f"Type {label}")
        try:
            bar_a = find_bar_by_id(project, bar_a_id)
            et_a = bar_a.ExpandedTask
            link = win32com.client.Dispatch(et_a.LinksOut.Item(1))
            link.type = type_val
            print(f"  type={type_val} => readback: {link.type}")
            project.EndTransaction()
            wait(project)
            project.Reschedule(pywintypes.Time(datetime.now()))
            wait(project)
            bar_a = find_bar_by_id(project, bar_a_id)
            bar_b = find_bar_by_id(project, bar_b_id)
            et_a = bar_a.ExpandedTask
            et_b = bar_b.ExpandedTask
            print(f"  A: Start={et_a.Start}, End={et_a.End}")
            print(f"  B: Start={et_b.Start}, End={et_b.End}")
        except Exception as e:
            print(f"  Error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

    # === Lag ===
    print(f"\n--- Set lag = 3d on FS link ---")
    project.StartTransaction("Lag")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        et_a = bar_a.ExpandedTask
        link = win32com.client.Dispatch(et_a.LinksOut.Item(1))
        lag_dur = et_a.GetDurationFromString("3d")
        link.StartLagTime = lag_dur
        print(f"  StartLagTime = 3d ({lag_dur.Hours}h) => readback: {link.StartLagTime.Hours}h")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        bar_a = find_bar_by_id(project, bar_a_id)
        bar_b = find_bar_by_id(project, bar_b_id)
        et_a = bar_a.ExpandedTask
        et_b = bar_b.ExpandedTask
        print(f"  A End={et_a.End}, B Start={et_b.Start}")
        link = win32com.client.Dispatch(et_a.LinksOut.Item(1))
        print(f"  Link: type={link.type}, lag={link.StartLagTime.Hours}h")
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # === Remove and re-add ===
    print(f"\n--- Remove link ---")
    project.StartTransaction("Remove")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        et_a = bar_a.ExpandedTask
        et_a.LinksOut.Remove(1)
        print(f"  Removed. LinksOut={et_a.LinksOut.Count}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    print(f"\n--- Re-add link ---")
    project.StartTransaction("Re-link")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        bar_b = find_bar_by_id(project, bar_b_id)
        et_a = bar_a.ExpandedTask
        et_b = bar_b.ExpandedTask
        link = et_a.LinkTo(et_b)
        if link:
            print(f"  Re-linked. type={win32com.client.Dispatch(link).type}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def cleanup(project, bar_ids):
    """Remove test bars by finding them in parent's ChildBars."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup ---")
    for bid in reversed(bar_ids):
        bar = find_bar_by_id(project, bid)
        if not bar:
            print(f"  Bar {bid} not found, skipping")
            continue

        et = bar.ExpandedTask
        try:
            parent_bar = win32com.client.Dispatch(et.GetActualParentBar())
            parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
        except Exception:
            print(f"  Can't find parent for {bid}")
            continue

        project.StartTransaction(f"Del {bid}")
        try:
            cb = parent_task.ChildBars
            for i in range(cb.Count, 0, -1):
                b = win32com.client.Dispatch(cb.Item(i))
                if b.ID == bid:
                    cb.Remove(i)
                    print(f"  Removed bar ID={bid}")
                    break
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  Remove error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass


if __name__ == "__main__":
    print("COM Explorer v17 — Corrected Workflow")
    print("=" * 80)
    created_ids = []
    try:
        app, project = connect()

        # Pre-cleanup
        print("\n--- Pre-cleanup ---")
        root_task = get_root_task(project)
        stale = []
        for i in range(1, root_task.ChildBars.Count + 1):
            cb = win32com.client.Dispatch(root_task.ChildBars.Item(i))
            ct = win32com.client.Dispatch(cb.Tasks(1))
            try:
                for j in range(1, ct.ChildBars.Count + 1):
                    sb = win32com.client.Dispatch(ct.ChildBars.Item(j))
                    if sb.Name.startswith(("V14_", "V15_", "V16_", "V17_", "TEST_")):
                        stale.append(sb.ID)
                        print(f"  Stale: ID={sb.ID}, Name={sb.Name}")
            except Exception:
                pass
        if stale:
            cleanup(project, stale)

        # Test 0: Compare ChildBars
        test_childbars_comparison(project)

        # Test 1: Full workflow
        created_ids = test_full_workflow(project)

        # Test 2: Link operations
        if len(created_ids) >= 2:
            test_link_operations(project, created_ids)

        # Cleanup
        cleanup(project, created_ids)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
        if created_ids:
            try:
                cleanup(project, created_ids)
            except Exception:
                pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
