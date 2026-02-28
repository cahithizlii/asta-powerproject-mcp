"""
COM Explorer v16 — FIXED Full Workflow
KEY FIXES:
  - Use bar.ExpandedTask (not bar.Tasks(1)) for newly created bars
  - Use non-milestone parent summary (Satinalma/Insaat)
  - IExpandedTask has ALL methods we need (common 338 + unique 101)

CONFIRMED METHODS:
  - bar.ExpandedTask.SetUserDuration(dur) - set duration
  - bar.ExpandedTask.GetDurationFromString("10d") - parse duration
  - bar.ExpandedTask.ImposedStart = datetime - set start date
  - bar.ExpandedTask.LinkTo(other_etask) - create FS link
  - bar.ExpandedTask.LinksOut.Remove(index) - remove link
  - link.type = 0/1/2/3 - set FS/SS/FF/SF
  - link.StartLagTime = dur - set lag
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


def find_bar_by_id(project, bar_id, max_depth=5):
    """Find bar by ID traversing hierarchy."""
    root_bar = project.Bars.Item(1)
    if root_bar.ID == bar_id:
        return root_bar
    root_et = root_bar.ExpandedTask
    return _search_bar(root_et, bar_id, 0, max_depth)


def _search_bar(parent_et, target_id, depth, max_depth):
    """Recursive bar search using ExpandedTask.ChildBars."""
    try:
        child_bars = parent_et.ChildBars
        for i in range(1, child_bars.Count + 1):
            cb = win32com.client.Dispatch(child_bars.Item(i))
            if cb.ID == target_id:
                return cb
            if depth < max_depth:
                cb_et = cb.ExpandedTask
                result = _search_bar(cb_et, target_id, depth + 1, max_depth)
                if result:
                    return result
    except Exception:
        pass
    return None


def get_etask(project, bar_id):
    """Get IExpandedTask for a bar."""
    bar = find_bar_by_id(project, bar_id)
    if bar:
        return bar.ExpandedTask
    return None


def test_full_workflow(project):
    """Create 2 bars, set duration + dates, link them."""
    print("\n" + "=" * 80)
    print("TEST 1: FULL WORKFLOW")
    print("=" * 80)

    # Find a proper summary (not Milestones, not root)
    root_bar = project.Bars.Item(1)
    root_et = root_bar.ExpandedTask

    print(f"Top-level summaries:")
    parent_bar = None
    for i in range(1, root_et.ChildBars.Count + 1):
        cb = win32com.client.Dispatch(root_et.ChildBars.Item(i))
        cb_et = cb.ExpandedTask
        child_count = 0
        try:
            child_count = cb_et.ChildBars.Count
        except Exception:
            pass
        print(f"  [{i}] ID={cb.ID}, Name={cb.Name[:40]}, Children={child_count}")
        # Pick second summary (skip Milestones at index 1)
        if child_count > 0 and i >= 2 and not parent_bar:
            parent_bar = cb

    if not parent_bar:
        print("  No suitable parent!")
        return []

    parent_et = parent_bar.ExpandedTask
    print(f"\nUsing parent: ID={parent_bar.ID}, Name={parent_bar.Name[:40]}")
    print(f"  ChildBars before: {parent_et.ChildBars.Count}")

    # === STEP 1: Create Bar A ===
    print(f"\n--- Step 1: Create Bar A ---")
    project.StartTransaction("Create A")
    try:
        new_a = win32com.client.Dispatch(parent_et.ChildBars.Add())
        new_a.Name = "V16_TEST_A"
        bar_a_id = new_a.ID
        # Use ExpandedTask, NOT Tasks(1)
        et_a = new_a.ExpandedTask
        print(f"  Created: BarID={bar_a_id}, ET type={type(et_a).__name__}")
        print(f"  ET.ID={et_a.ID}, ET.Name={et_a.Name[:30]}")
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

    # === STEP 2: Create Bar B ===
    print(f"\n--- Step 2: Create Bar B ---")
    project.StartTransaction("Create B")
    try:
        parent_et = parent_bar.ExpandedTask  # re-fetch
        new_b = win32com.client.Dispatch(parent_et.ChildBars.Add())
        new_b.Name = "V16_TEST_B"
        bar_b_id = new_b.ID
        print(f"  Created: BarID={bar_b_id}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return [bar_a_id]

    # === STEP 3: Set Durations ===
    print(f"\n--- Step 3: Set Durations ---")
    project.StartTransaction("Set durations")
    try:
        et_a = get_etask(project, bar_a_id)
        print(f"  A found: {et_a is not None}")
        dur_a = et_a.GetDurationFromString("10d")
        et_a.SetUserDuration(dur_a)
        print(f"  A: SetUserDuration(10d) => OK")

        et_b = get_etask(project, bar_b_id)
        dur_b = et_b.GetDurationFromString("5d")
        et_b.SetUserDuration(dur_b)
        print(f"  B: SetUserDuration(5d) => OK")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Duration error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    et_a = get_etask(project, bar_a_id)
    et_b = get_etask(project, bar_b_id)
    print(f"  A: Dur={et_a.GetUserDuration().Hours}h, Start={et_a.Start}, End={et_a.End}")
    print(f"  B: Dur={et_b.GetUserDuration().Hours}h, Start={et_b.Start}, End={et_b.End}")

    # === STEP 4: Set Date on A ===
    print(f"\n--- Step 4: ImposedStart on A ---")
    project.StartTransaction("ImposedStart")
    try:
        et_a = get_etask(project, bar_a_id)
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

    et_a = get_etask(project, bar_a_id)
    print(f"  A: Start={et_a.Start}, End={et_a.End}, Constraint={et_a.Constraint}")

    # === STEP 5: Link A -> B ===
    print(f"\n--- Step 5: LinkTo A -> B ---")
    project.StartTransaction("LinkTo")
    try:
        et_a = get_etask(project, bar_a_id)
        et_b = get_etask(project, bar_b_id)
        link = et_a.LinkTo(et_b)
        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  LINK CREATED! ID={ld.ID}, type={ld.type}")
        else:
            print(f"  LinkTo returned None!")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  LinkTo error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # === STEP 6: Verify ===
    print(f"\n--- Step 6: VERIFY ---")
    et_a = get_etask(project, bar_a_id)
    et_b = get_etask(project, bar_b_id)
    print(f"  A: Start={et_a.Start}, End={et_a.End}")
    print(f"     Dur={et_a.GetUserDuration().Hours}h, LinksOut={et_a.LinksOut.Count}")
    print(f"  B: Start={et_b.Start}, End={et_b.End}")
    print(f"     Dur={et_b.GetUserDuration().Hours}h, LinksIn={et_b.LinksIn.Count}")

    # Check B starts after A
    a_end = et_a.End
    b_start = et_b.Start
    if b_start >= a_end:
        print(f"\n  *** SUCCESS: B starts at/after A ends (FS link working!) ***")
    else:
        print(f"\n  WARNING: B starts before A ends!")

    return [bar_a_id, bar_b_id]


def test_link_modifications(project, bar_ids):
    """Test link type/lag changes."""
    print("\n" + "=" * 80)
    print("TEST 2: LINK TYPE + LAG")
    print("=" * 80)

    bar_a_id, bar_b_id = bar_ids[0], bar_ids[1]

    # Change to SS
    print(f"\n--- Change to SS (type=1) ---")
    project.StartTransaction("SS")
    try:
        et_a = get_etask(project, bar_a_id)
        link = win32com.client.Dispatch(et_a.LinksOut.Item(1))
        link.type = 1  # SS
        print(f"  Set type=1 => readback: {link.type}")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        et_a = get_etask(project, bar_a_id)
        et_b = get_etask(project, bar_b_id)
        print(f"  A Start={et_a.Start}, B Start={et_b.Start}")
        print(f"  (SS: B should start when A starts)")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Change to FF
    print(f"\n--- Change to FF (type=2) ---")
    project.StartTransaction("FF")
    try:
        et_a = get_etask(project, bar_a_id)
        link = win32com.client.Dispatch(et_a.LinksOut.Item(1))
        link.type = 2  # FF
        print(f"  Set type=2 => readback: {link.type}")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        et_a = get_etask(project, bar_a_id)
        et_b = get_etask(project, bar_b_id)
        print(f"  A End={et_a.End}, B End={et_b.End}")
        print(f"  (FF: B should end when A ends)")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Back to FS with lag
    print(f"\n--- FS with 3d lag ---")
    project.StartTransaction("FS+lag")
    try:
        et_a = get_etask(project, bar_a_id)
        link = win32com.client.Dispatch(et_a.LinksOut.Item(1))
        link.type = 0  # FS
        lag = et_a.GetDurationFromString("3d")
        link.StartLagTime = lag
        print(f"  type=0 (FS), StartLagTime=3d ({lag.Hours}h)")
        print(f"  Readback: type={link.type}, lag={link.StartLagTime.Hours}h")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        et_a = get_etask(project, bar_a_id)
        et_b = get_etask(project, bar_b_id)
        print(f"  A End={et_a.End}")
        print(f"  B Start={et_b.Start}")
        print(f"  (FS+3d: B should start 3d after A ends)")
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Remove and re-add link
    print(f"\n--- Remove link ---")
    project.StartTransaction("Remove link")
    try:
        et_a = get_etask(project, bar_a_id)
        et_a.LinksOut.Remove(1)
        print(f"  Removed. LinksOut={et_a.LinksOut.Count}")
        project.EndTransaction()
        wait(project)

        et_b = get_etask(project, bar_b_id)
        print(f"  B LinksIn={et_b.LinksIn.Count}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    print(f"\n--- Re-add link (FS, no lag) ---")
    project.StartTransaction("Re-link")
    try:
        et_a = get_etask(project, bar_a_id)
        et_b = get_etask(project, bar_b_id)
        link = et_a.LinkTo(et_b)
        if link:
            print(f"  Re-linked! type={win32com.client.Dispatch(link).type}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def test_constraint_types(project, bar_ids):
    """Test different constraint types on a bar."""
    print("\n" + "=" * 80)
    print("TEST 3: CONSTRAINT TYPES")
    print("=" * 80)

    bar_a_id = bar_ids[0]
    et_a = get_etask(project, bar_a_id)
    print(f"Current: Constraint={et_a.Constraint}, Start={et_a.Start}")

    # Test RemoveConstraint
    print(f"\n--- RemoveConstraint ---")
    project.StartTransaction("RemoveConstraint")
    try:
        et_a = get_etask(project, bar_a_id)
        et_a.RemoveConstraint()
        print(f"  RemoveConstraint => Constraint={et_a.Constraint}")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        et_a = get_etask(project, bar_a_id)
        print(f"  After reschedule: Constraint={et_a.Constraint}, Start={et_a.Start}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test ImposedEnd
    print(f"\n--- ImposedEnd ---")
    project.StartTransaction("ImposedEnd")
    try:
        et_a = get_etask(project, bar_a_id)
        et_a.ImposedEnd = pywintypes.Time(datetime(2026, 8, 15))
        print(f"  ImposedEnd = 2026-08-15 => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        et_a = get_etask(project, bar_a_id)
        print(f"  After: Start={et_a.Start}, End={et_a.End}, Constraint={et_a.Constraint}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test StartConstraintDate
    print(f"\n--- StartConstraintDate ---")
    project.StartTransaction("StartConstraint")
    try:
        et_a = get_etask(project, bar_a_id)
        et_a.RemoveConstraint()
        et_a.StartConstraintDate = pywintypes.Time(datetime(2026, 9, 1))
        print(f"  StartConstraintDate = 2026-09-01 => OK")
        print(f"  Readback: {et_a.StartConstraintDate}")
        print(f"  Constraint: {et_a.Constraint}")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        et_a = get_etask(project, bar_a_id)
        print(f"  After: Start={et_a.Start}, End={et_a.End}, Constraint={et_a.Constraint}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Restore ImposedStart
    print(f"\n--- Restore ImposedStart ---")
    project.StartTransaction("Restore")
    try:
        et_a = get_etask(project, bar_a_id)
        et_a.RemoveConstraint()
        et_a.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        et_a = get_etask(project, bar_a_id)
        print(f"  Restored: Start={et_a.Start}, End={et_a.End}, Constraint={et_a.Constraint}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def cleanup(project, bar_ids):
    """Remove test bars."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup ---")
    for bid in reversed(bar_ids):
        bar = find_bar_by_id(project, bid)
        if not bar:
            print(f"  Bar {bid} not found")
            continue

        # Find in parent's ChildBars
        bar_et = bar.ExpandedTask
        parent_bar = win32com.client.Dispatch(bar_et.GetActualParentBar())
        parent_et = parent_bar.ExpandedTask

        project.StartTransaction(f"Del {bid}")
        try:
            cb = parent_et.ChildBars
            for i in range(cb.Count, 0, -1):
                b = win32com.client.Dispatch(cb.Item(i))
                if b.ID == bid:
                    cb.Remove(i)
                    print(f"  Removed bar ID={bid}")
                    break
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  Remove error for {bid}: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass


if __name__ == "__main__":
    print("COM Explorer v16 — Full Working Workflow (FIXED)")
    print("=" * 80)
    created_ids = []
    try:
        app, project = connect()

        # Pre-cleanup
        print("\n--- Pre-cleanup ---")
        root_bar = project.Bars.Item(1)
        root_et = root_bar.ExpandedTask
        stale = []
        for i in range(1, root_et.ChildBars.Count + 1):
            cb = win32com.client.Dispatch(root_et.ChildBars.Item(i))
            cb_et = cb.ExpandedTask
            try:
                for j in range(1, cb_et.ChildBars.Count + 1):
                    sb = win32com.client.Dispatch(cb_et.ChildBars.Item(j))
                    if sb.Name.startswith(("V14_", "V15_", "V16_", "TEST_")):
                        stale.append(sb.ID)
                        print(f"  Stale: ID={sb.ID}, Name={sb.Name}")
            except Exception:
                pass
        if stale:
            cleanup(project, stale)

        # Test 1: Full workflow
        created_ids = test_full_workflow(project)

        # Test 2: Link modifications
        if len(created_ids) >= 2:
            test_link_modifications(project, created_ids)

        # Test 3: Constraint types
        if created_ids:
            test_constraint_types(project, created_ids)

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
