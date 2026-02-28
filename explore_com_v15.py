"""
COM Explorer v15 — Full working workflow
CONFIRMED:
  - SetUserDuration WORKS on real leaf tasks
  - LinkTo WORKS on leaf tasks
  - LinksOut.Remove() works
  - Link types: 0=FS, 1=SS, 2=FF

KEY FIXES:
  - Use bar.ExpandedTask for ImposedStart (not ITask)
  - Use proper summary bar for ChildBars.Add() (not Milestones)

TESTS:
  1. Full workflow: create 2 bars under summary, set duration, set dates, link, reschedule
  2. Link type/lag modification
  3. Verify ITask vs IExpandedTask property differences
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
    """Find a bar by ID by traversing hierarchy."""
    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    if root_bar.ID == bar_id:
        return root_bar
    return _search_bar(root_task, bar_id, 0, max_depth)


def _search_bar(parent_task, target_id, depth, max_depth):
    """Recursive bar search."""
    try:
        child_bars = parent_task.ChildBars
        for i in range(1, child_bars.Count + 1):
            cb = win32com.client.Dispatch(child_bars.Item(i))
            if cb.ID == target_id:
                return cb
            if depth < max_depth:
                ct = win32com.client.Dispatch(cb.Tasks(1))
                result = _search_bar(ct, target_id, depth + 1, max_depth)
                if result:
                    return result
    except Exception:
        pass
    return None


def get_etask(project, bar_id):
    """Get IExpandedTask for a bar (has ImposedStart, more properties)."""
    bar = find_bar_by_id(project, bar_id)
    if bar:
        return bar.ExpandedTask
    return None


def get_task(project, bar_id):
    """Get ITask for a bar (via bar.Tasks(1))."""
    bar = find_bar_by_id(project, bar_id)
    if bar:
        return win32com.client.Dispatch(bar.Tasks(1))
    return None


def test_itask_vs_iexpandedtask(project):
    """Compare ITask vs IExpandedTask properties."""
    print("\n" + "=" * 80)
    print("TEST 0: ITask vs IExpandedTask COMPARISON")
    print("=" * 80)

    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))

    # Find a leaf task with duration
    child_bars = root_task.ChildBars
    for i in range(1, child_bars.Count + 1):
        cb = win32com.client.Dispatch(child_bars.Item(i))
        ct = win32com.client.Dispatch(cb.Tasks(1))
        try:
            sub_bars = ct.ChildBars
            if sub_bars.Count > 0:
                for j in range(1, min(sub_bars.Count + 1, 10)):
                    sb = win32com.client.Dispatch(sub_bars.Item(j))
                    st = win32com.client.Dispatch(sb.Tasks(1))
                    try:
                        dur = st.GetUserDuration().Hours
                        if dur > 0:
                            # Found a leaf task with duration
                            bar_id = sb.ID
                            print(f"Found leaf: BarID={bar_id}, Name={sb.Name[:30]}, Dur={dur}h")

                            # Get as ITask (via Tasks(1))
                            itask = win32com.client.Dispatch(sb.Tasks(1))
                            print(f"\n  ITask (bar.Tasks(1)) type: {type(itask)}")

                            # Get as IExpandedTask (via ExpandedTask)
                            ietask = sb.ExpandedTask
                            print(f"  IExpandedTask (bar.ExpandedTask) type: {type(ietask)}")

                            # Compare attributes
                            itask_attrs = set(a for a in dir(itask) if not a.startswith('_'))
                            ietask_attrs = set(a for a in dir(ietask) if not a.startswith('_'))

                            only_itask = itask_attrs - ietask_attrs
                            only_ietask = ietask_attrs - itask_attrs
                            common = itask_attrs & ietask_attrs

                            print(f"\n  Common attributes: {len(common)}")
                            print(f"  Only in ITask: {len(only_itask)}")
                            if only_itask:
                                print(f"    {sorted(only_itask)[:20]}")
                            print(f"  Only in IExpandedTask: {len(only_ietask)}")
                            if only_ietask:
                                imp_attrs = [a for a in sorted(only_ietask)
                                             if any(kw in a.lower() for kw in
                                                    ['imposed', 'constraint', 'start', 'end',
                                                     'date', 'dur', 'link', 'move', 'type'])]
                                print(f"    Important: {imp_attrs}")
                                remaining = sorted(only_ietask - set(imp_attrs))[:20]
                                print(f"    Others (first 20): {remaining}")

                            # Test ImposedStart on both
                            print(f"\n  Testing ImposedStart:")
                            for label, obj in [("ITask", itask), ("IExpandedTask", ietask)]:
                                try:
                                    val = obj.ImposedStart
                                    print(f"    {label}.ImposedStart = {val}")
                                except AttributeError:
                                    print(f"    {label}.ImposedStart => AttributeError (not available)")
                                except Exception as e:
                                    print(f"    {label}.ImposedStart => {str(e)[:50]}")

                            return bar_id
                    except Exception:
                        pass
        except Exception:
            pass

    print("  No suitable leaf task found!")
    return None


def test_full_workflow(project):
    """Create bars, set duration/dates, link them, verify schedule."""
    print("\n" + "=" * 80)
    print("TEST 1: FULL WORKFLOW — CREATE + DURATION + DATES + LINK")
    print("=" * 80)

    # Find a proper summary bar (not Milestones)
    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    child_bars = root_task.ChildBars

    print(f"Top-level summaries:")
    parent_bar = None
    parent_task = None
    for i in range(1, child_bars.Count + 1):
        cb = win32com.client.Dispatch(child_bars.Item(i))
        ct = win32com.client.Dispatch(cb.Tasks(1))
        child_count = 0
        try:
            child_count = ct.ChildBars.Count
        except Exception:
            pass
        print(f"  [{i}] ID={cb.ID}, Name={cb.Name[:30]}, Children={child_count}")
        # Pick a summary with children (not Milestones)
        if child_count > 0 and not parent_bar:
            parent_bar = cb
            parent_task = ct

    if not parent_bar:
        print("  No suitable parent summary found!")
        return []

    print(f"\nUsing parent: ID={parent_bar.ID}, Name={parent_bar.Name[:40]}")
    orig_child_count = parent_task.ChildBars.Count

    # CREATE BAR A
    print(f"\n--- Step 1: Create Bar A ---")
    project.StartTransaction("Create A")
    try:
        new_a = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_a.Name = "V15_TEST_A"
        bar_a_id = new_a.ID
        print(f"  Created: BarID={bar_a_id}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return []

    # CREATE BAR B
    print(f"\n--- Step 2: Create Bar B ---")
    project.StartTransaction("Create B")
    try:
        # Re-fetch parent
        parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
        new_b = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_b.Name = "V15_TEST_B"
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

    # SET DURATIONS
    print(f"\n--- Step 3: Set Durations ---")
    project.StartTransaction("Set durations")
    try:
        ta = get_task(project, bar_a_id)
        dur_a = ta.GetDurationFromString("10d")
        ta.SetUserDuration(dur_a)
        print(f"  A: SetUserDuration(10d) => OK")

        tb = get_task(project, bar_b_id)
        dur_b = tb.GetDurationFromString("5d")
        tb.SetUserDuration(dur_b)
        print(f"  B: SetUserDuration(5d) => OK")

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

    # Check durations
    ta = get_task(project, bar_a_id)
    tb = get_task(project, bar_b_id)
    print(f"  A: Dur={ta.GetUserDuration().Hours}h, Start={ta.Start}, End={ta.End}")
    print(f"  B: Dur={tb.GetUserDuration().Hours}h, Start={tb.Start}, End={tb.End}")

    # SET DATES via IExpandedTask.ImposedStart
    print(f"\n--- Step 4: Set Start Date (ImposedStart via IExpandedTask) ---")
    project.StartTransaction("ImposedStart A")
    try:
        eta = get_etask(project, bar_a_id)
        print(f"  IExpandedTask type: {type(eta)}")
        eta.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
        print(f"  ImposedStart = 2026-07-01 => OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  ImposedStart error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    eta = get_etask(project, bar_a_id)
    print(f"  A after reschedule: Start={eta.Start}, End={eta.End}, Constraint={eta.Constraint}")

    # LINK A -> B (FS)
    print(f"\n--- Step 5: Link A -> B (FS) ---")
    project.StartTransaction("LinkTo")
    try:
        ta = get_task(project, bar_a_id)
        tb = get_task(project, bar_b_id)
        link = ta.LinkTo(tb)
        if link:
            ld = win32com.client.Dispatch(link)
            link_id = ld.ID
            print(f"  LINK CREATED! ID={link_id}, type={ld.type} (FS)")
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

    # VERIFY FINAL STATE
    print(f"\n--- Step 6: Verify Final State ---")
    eta = get_etask(project, bar_a_id)
    etb = get_etask(project, bar_b_id)
    ta = get_task(project, bar_a_id)
    tb = get_task(project, bar_b_id)

    print(f"  A: Start={eta.Start}, End={eta.End}, Dur={ta.GetUserDuration().Hours}h")
    print(f"     LinksOut={ta.LinksOut.Count}, Constraint={eta.Constraint}")
    print(f"  B: Start={etb.Start}, End={etb.End}, Dur={tb.GetUserDuration().Hours}h")
    print(f"     LinksIn={tb.LinksIn.Count}, Constraint={etb.Constraint}")
    print(f"  B should start AFTER A ends (FS link)")

    # Check: B.Start >= A.End?
    if etb.Start >= eta.End:
        print(f"  *** SCHEDULE LOGIC CONFIRMED: B starts after A ***")
    else:
        print(f"  WARNING: B starts before A ends!")

    return [bar_a_id, bar_b_id]


def test_link_modifications(project, bar_ids):
    """Test link type changes and lag."""
    print("\n" + "=" * 80)
    print("TEST 2: LINK TYPE AND LAG MODIFICATION")
    print("=" * 80)

    if len(bar_ids) < 2:
        print("  Need 2 bar IDs")
        return

    bar_a_id, bar_b_id = bar_ids[0], bar_ids[1]

    # Get current link
    ta = get_task(project, bar_a_id)
    if ta.LinksOut.Count == 0:
        print("  No links to modify")
        return

    link = win32com.client.Dispatch(ta.LinksOut.Item(1))
    print(f"Current link: type={link.type}, lag={link.StartLagTime.Hours}h")

    # Test type changes
    for type_val, label in [(1, "SS"), (2, "FF"), (3, "SF"), (0, "FS")]:
        print(f"\n--- Set type={type_val} ({label}) ---")
        project.StartTransaction(f"Type {label}")
        try:
            ta = get_task(project, bar_a_id)
            link = win32com.client.Dispatch(ta.LinksOut.Item(1))
            link.type = type_val
            print(f"  Set type={type_val} => readback: {link.type}")
            project.EndTransaction()
            wait(project)
            project.Reschedule(pywintypes.Time(datetime.now()))
            wait(project)

            eta = get_etask(project, bar_a_id)
            etb = get_etask(project, bar_b_id)
            print(f"  A: Start={eta.Start}, End={eta.End}")
            print(f"  B: Start={etb.Start}, End={etb.End}")
        except Exception as e:
            print(f"  Error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

    # Test lag
    print(f"\n--- Set lag = 3d (type FS) ---")
    project.StartTransaction("Lag 3d")
    try:
        ta = get_task(project, bar_a_id)
        link = win32com.client.Dispatch(ta.LinksOut.Item(1))
        link.type = 0  # FS

        lag_dur = ta.GetDurationFromString("3d")
        link.StartLagTime = lag_dur
        print(f"  Set StartLagTime = 3d ({lag_dur.Hours}h) => OK")
        print(f"  Readback: {link.StartLagTime.Hours}h")

        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        eta = get_etask(project, bar_a_id)
        etb = get_etask(project, bar_b_id)
        print(f"  A: Start={eta.Start}, End={eta.End}")
        print(f"  B: Start={etb.Start}, End={etb.End}")

        ta = get_task(project, bar_a_id)
        link = win32com.client.Dispatch(ta.LinksOut.Item(1))
        print(f"  Link: type={link.type}, lag={link.StartLagTime.Hours}h")
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test negative lag (lead)
    print(f"\n--- Set lag = -2d (lead) ---")
    project.StartTransaction("Lead -2d")
    try:
        ta = get_task(project, bar_a_id)
        link = win32com.client.Dispatch(ta.LinksOut.Item(1))

        lag_dur = ta.GetDurationFromString("-2d")
        print(f"  GetDurationFromString('-2d') => {lag_dur.Hours}h")
        link.StartLagTime = lag_dur
        print(f"  Set StartLagTime = -2d => OK")

        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        eta = get_etask(project, bar_a_id)
        etb = get_etask(project, bar_b_id)
        print(f"  A: Start={eta.Start}, End={eta.End}")
        print(f"  B: Start={etb.Start}, End={etb.End}")
    except Exception as e:
        print(f"  Lead error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def test_link_removal_methods(project, bar_ids):
    """Test different ways to remove links."""
    print("\n" + "=" * 80)
    print("TEST 3: LINK REMOVAL METHODS")
    print("=" * 80)

    if len(bar_ids) < 2:
        print("  Need 2 bar IDs")
        return

    bar_a_id, bar_b_id = bar_ids[0], bar_ids[1]
    ta = get_task(project, bar_a_id)
    tb = get_task(project, bar_b_id)

    print(f"  A LinksOut: {ta.LinksOut.Count}")
    print(f"  B LinksIn: {tb.LinksIn.Count}")

    # Remove via LinksOut.Remove(index)
    if ta.LinksOut.Count > 0:
        project.StartTransaction("Remove link")
        try:
            ta.LinksOut.Remove(1)
            print(f"  LinksOut.Remove(1) => OK")
            project.EndTransaction()
            wait(project)

            ta = get_task(project, bar_a_id)
            tb = get_task(project, bar_b_id)
            print(f"  After: A LinksOut={ta.LinksOut.Count}, B LinksIn={tb.LinksIn.Count}")
        except Exception as e:
            print(f"  Remove error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

    # Re-create link
    print(f"\n  Re-creating link...")
    project.StartTransaction("Re-link")
    try:
        ta = get_task(project, bar_a_id)
        tb = get_task(project, bar_b_id)
        link = ta.LinkTo(tb)
        if link:
            print(f"  Re-linked OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Re-link error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Remove via LinksIn.Remove(index)
    tb = get_task(project, bar_b_id)
    if tb.LinksIn.Count > 0:
        project.StartTransaction("Remove via LinksIn")
        try:
            tb.LinksIn.Remove(1)
            print(f"  LinksIn.Remove(1) => OK")
            project.EndTransaction()
            wait(project)

            ta = get_task(project, bar_a_id)
            tb = get_task(project, bar_b_id)
            print(f"  After: A LinksOut={ta.LinksOut.Count}, B LinksIn={tb.LinksIn.Count}")
        except Exception as e:
            print(f"  LinksIn Remove error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass


def cleanup(project, bar_ids):
    """Remove test bars."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup: {len(bar_ids)} bars ---")

    for bid in reversed(bar_ids):
        bar = find_bar_by_id(project, bid)
        if not bar:
            print(f"  Bar {bid} not found")
            continue

        # Get parent
        task = win32com.client.Dispatch(bar.Tasks(1))
        parent_bar_obj = win32com.client.Dispatch(task.GetActualParentBar())
        parent_task = win32com.client.Dispatch(parent_bar_obj.Tasks(1))

        project.StartTransaction(f"Remove {bid}")
        try:
            child_bars = parent_task.ChildBars
            for i in range(child_bars.Count, 0, -1):
                b = win32com.client.Dispatch(child_bars.Item(i))
                if b.ID == bid:
                    child_bars.Remove(i)
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
    print("COM Explorer v15 — Full Working Workflow")
    print("=" * 80)
    created_ids = []
    try:
        app, project = connect()

        # Pre-cleanup
        print("\n--- Pre-cleanup ---")
        root_bar = project.Bars.Item(1)
        root_task = win32com.client.Dispatch(root_bar.Tasks(1))
        stale = []
        for i in range(1, root_task.ChildBars.Count + 1):
            cb = win32com.client.Dispatch(root_task.ChildBars.Item(i))
            ct = win32com.client.Dispatch(cb.Tasks(1))
            try:
                for j in range(1, ct.ChildBars.Count + 1):
                    sb = win32com.client.Dispatch(ct.ChildBars.Item(j))
                    if sb.Name.startswith(("V14_", "V15_", "TEST_")):
                        stale.append(sb.ID)
                        print(f"  Stale: ID={sb.ID}, Name={sb.Name}")
            except Exception:
                pass
        if stale:
            cleanup(project, stale)

        # Test 0: ITask vs IExpandedTask
        test_itask_vs_iexpandedtask(project)

        # Test 1: Full workflow
        created_ids = test_full_workflow(project)

        # Test 2: Link modifications
        if len(created_ids) >= 2:
            test_link_modifications(project, created_ids)

        # Test 3: Link removal
        if len(created_ids) >= 2:
            test_link_removal_methods(project, created_ids)

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
