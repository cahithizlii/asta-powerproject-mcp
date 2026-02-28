"""
COM Explorer v11 — Access bars INSIDE the hierarchy
1. Dynamic dispatch on ChildBars
2. Navigate via NextTask/PreviousTask
3. Test IBarChartView bar access methods
4. Find linkable bars and test LinkTo
5. Test duration on hierarchy bars
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


def test_childbars_dispatch(project):
    """Test accessing child bars via dynamic dispatch."""
    print("\n" + "=" * 80)
    print("TEST 1: DYNAMIC DISPATCH ON CHILDBARS")
    print("=" * 80)

    bars = project.Bars
    summary = bars.Item(1)
    et = summary.ExpandedTask

    print(f"Summary: ID={summary.ID}, Name={summary.Name[:40]}")

    # Static ChildBars
    cb_static = et.ChildBars
    print(f"\nStatic ChildBars:")
    print(f"  type: {type(cb_static)}")
    print(f"  Count: {cb_static.Count}")
    for i in range(1, min(cb_static.Count + 1, 6)):
        try:
            child = cb_static.Item(i)
            print(f"  Item({i}): ID={child.ID}, Name={child.Name[:30]}")
        except Exception as e:
            print(f"  Item({i}): {str(e)[:40]}")

    # Dynamic dispatch ChildBars
    cb_dyn = win32com.client.Dispatch(cb_static)
    print(f"\nDynamic ChildBars:")
    print(f"  type: {type(cb_dyn)}")
    print(f"  Count: {cb_dyn.Count}")
    for i in range(1, min(cb_dyn.Count + 1, 6)):
        try:
            child = cb_dyn.Item(i)
            print(f"  Item({i}): ID={child.ID}, Name={child.Name[:30]}")
        except Exception as e:
            print(f"  Item({i}): {str(e)[:40]}")

    # Try Dispatch on individual items
    print(f"\nDispatch on individual items:")
    for i in range(1, min(cb_static.Count + 1, 6)):
        try:
            child_raw = cb_static.Item(i)
            child_dyn = win32com.client.Dispatch(child_raw)
            print(f"  Dispatch(Item({i})): ID={child_dyn.ID}, Name={child_dyn.Name[:30]}")
        except Exception as e:
            print(f"  Dispatch(Item({i})): {str(e)[:40]}")

    # Try _oleobj_ dispatch
    print(f"\nRaw _oleobj_ dispatch:")
    try:
        cb_ole = et._oleobj_.Invoke(
            et._oleobj_.GetIDsOfNames('ChildBars', 0)[0],
            0, 2, 1)
        print(f"  ChildBars via Invoke: {cb_ole}")
        print(f"  type: {type(cb_ole)}")
        if hasattr(cb_ole, 'Count'):
            print(f"  Count: {cb_ole.Count}")
    except Exception as e:
        print(f"  Error: {e}")

    # Try bars.All() and check IDs
    print(f"\nBars.All():")
    try:
        all_bars = bars.All()
        if all_bars:
            for idx, b in enumerate(all_bars):
                print(f"  All[{idx}]: ID={b.ID}, Name={b.Name[:30]}")
                if idx > 5:
                    print(f"  ... (total {len(all_bars)})")
                    break
    except Exception as e:
        print(f"  Error: {e}")


def test_next_task(project):
    """Navigate through tasks using NextTask/PreviousTask."""
    print("\n" + "=" * 80)
    print("TEST 2: NEXTTASK TRAVERSAL")
    print("=" * 80)

    bars = project.Bars
    summary = bars.Item(1)
    et = summary.ExpandedTask

    print(f"Summary: ID={summary.ID}")

    # Try NextTask from summary
    print(f"\nFrom summary:")
    current = et
    for i in range(10):
        try:
            next_et = current.NextTask()
            if next_et is None:
                print(f"  step {i}: NextTask() => None")
                break
            next_dyn = win32com.client.Dispatch(next_et)
            print(f"  step {i}: NextTask() => ID={next_dyn.ID}, Name={next_dyn.Name[:30]}, "
                  f"type={next_dyn.type}, Level={next_dyn.HierarchyLevel}, "
                  f"LinksIn={next_dyn.LinksIn.Count}, LinksOut={next_dyn.LinksOut.Count}")
            current = next_dyn
        except Exception as e:
            print(f"  step {i}: NextTask error: {str(e)[:50]}")
            break

    # If we found tasks, try getting their bar
    if current != et:
        print(f"\n  Current task's Bar:")
        try:
            bar = current.Bar
            print(f"    Bar ID={bar.ID}, Name={bar.Name[:30]}")
            bar_et = bar.ExpandedTask
            print(f"    Bar.ExpandedTask ID={bar_et.ID}")
        except Exception as e:
            print(f"    Error: {e}")


def test_barchartview_bar_access(project):
    """Test IBarChartView methods for accessing bars by ID."""
    print("\n" + "=" * 80)
    print("TEST 3: BARCHARTVIEW BAR ACCESS")
    print("=" * 80)

    view = project.CurrentView
    bcv = win32com.client.Dispatch(view)

    all_ids = bcv.AllBarIds()
    print(f"AllBarIds: {len(all_ids)} IDs")
    print(f"First 10: {all_ids[:10]}")

    # Try every possible method to get a bar from the view
    test_id = all_ids[1] if len(all_ids) > 1 else all_ids[0]
    print(f"\nTrying to get bar ID={test_id} from IBarChartView:")

    for method_name in ['GetBar', 'FindBar', 'Bar', 'GetBarById',
                        'BarById', 'GetBarObject', 'GetObjectById',
                        'GetItem', 'Item', 'FindObject']:
        try:
            fn = getattr(bcv, method_name)
            if callable(fn):
                try:
                    result = fn(test_id)
                    print(f"  {method_name}({test_id}) => {result}")
                    if result:
                        try:
                            r_dyn = win32com.client.Dispatch(result)
                            print(f"    ID={r_dyn.ID}, Name={r_dyn.Name[:30]}")
                        except Exception:
                            pass
                except Exception as e:
                    print(f"  {method_name}({test_id}) => {str(e)[:50]}")
            else:
                print(f"  {method_name} = {fn}")
        except AttributeError:
            pass

    # Try AllTaskBaseIds
    print(f"\nAllTaskBaseIds:")
    try:
        tb_ids = bcv.AllTaskBaseIds()
        print(f"  Count: {len(tb_ids)}")
        print(f"  First 10: {tb_ids[:10]}")
    except Exception as e:
        print(f"  Error: {e}")

    # Try selection-related methods
    print(f"\nSelection methods:")
    for method_name in ['SelectedBars', 'selection', 'GetSelection',
                        'SelectedBarIds', 'SelectBar', 'SelectAll']:
        try:
            fn = getattr(bcv, method_name)
            if callable(fn):
                try:
                    result = fn()
                    print(f"  {method_name}() => {result}")
                except Exception as e:
                    print(f"  {method_name}() => {str(e)[:50]}")
            else:
                print(f"  {method_name} = {fn}")
        except AttributeError:
            pass


def test_link_hierarchy_bars(project):
    """Find bars inside hierarchy that can be linked."""
    print("\n" + "=" * 80)
    print("TEST 4: FIND AND LINK HIERARCHY BARS")
    print("=" * 80)

    bars = project.Bars
    summary = bars.Item(1)
    et_summary = summary.ExpandedTask

    # Navigate using NextTask to find real task bars
    current = et_summary
    tasks_found = []
    for i in range(50):
        try:
            next_et = current.NextTask()
            if next_et is None:
                break
            next_dyn = win32com.client.Dispatch(next_et)
            level = next_dyn.HierarchyLevel
            li = next_dyn.LinksIn.Count
            lo = next_dyn.LinksOut.Count
            name = next_dyn.Name[:30]
            tid = next_dyn.ID

            if level > 0:
                tasks_found.append(next_dyn)
                if len(tasks_found) <= 10:
                    print(f"  [{i}] ID={tid}, Name={name}, Level={level}, "
                          f"LinksIn={li}, LinksOut={lo}")
            current = next_dyn
        except Exception as e:
            print(f"  step {i}: {str(e)[:50]}")
            break

    print(f"\n  Found {len(tasks_found)} hierarchy tasks")

    if len(tasks_found) < 2:
        print("  Need at least 2 hierarchy tasks")
        return

    # Find two unlinked tasks to test
    t1 = tasks_found[0]
    t2 = tasks_found[1]
    print(f"\n  Will try to link:")
    print(f"    T1: ID={t1.ID}, Name={t1.Name[:30]}, Level={t1.HierarchyLevel}")
    print(f"    T2: ID={t2.ID}, Name={t2.Name[:30]}, Level={t2.HierarchyLevel}")
    print(f"    T1.LinksOut before: {t1.LinksOut.Count}")
    print(f"    T2.LinksIn before: {t2.LinksIn.Count}")

    # Try LinkTo
    project.StartTransaction("Link hierarchy bars")
    try:
        link = t1.LinkTo(t2)
        print(f"\n  t1.LinkTo(t2) => {link}")
        if link:
            print(f"  SUCCESS! Link created!")
            link_dyn = win32com.client.Dispatch(link)
            for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
                try:
                    val = getattr(link_dyn, attr)
                    if not callable(val):
                        print(f"    {attr} = {str(val)[:50]}")
                    else:
                        print(f"    {attr}() [callable]")
                except Exception:
                    pass
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"\n  LinkTo error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Check
    print(f"\n  After LinkTo attempt:")
    print(f"    T1.LinksOut: {t1.LinksOut.Count}")
    print(f"    T2.LinksIn: {t2.LinksIn.Count}")

    # If that didn't work, try linking bars that are already proven linkable
    # (they already have links)
    if t1.LinksOut.Count == 0 or not link:
        print(f"\n  Trying to find already-linked bars...")
        for t in tasks_found:
            if t.LinksOut.Count > 0:
                print(f"    Found bar with LinksOut: ID={t.ID}, Name={t.Name[:30]}")
                # Get existing link to understand the pattern
                existing_link = t.LinksOut.Item(1)
                el_dyn = win32com.client.Dispatch(existing_link)
                print(f"    Existing link details:")
                for attr in sorted([a for a in dir(el_dyn) if not a.startswith('_')]):
                    try:
                        val = getattr(el_dyn, attr)
                        if not callable(val):
                            print(f"      {attr} = {str(val)[:50]}")
                        else:
                            print(f"      {attr}() [callable]")
                    except Exception:
                        pass
                break


def test_duration_hierarchy_bar(project):
    """Test duration on hierarchy bars."""
    print("\n" + "=" * 80)
    print("TEST 5: DURATION ON HIERARCHY BARS")
    print("=" * 80)

    bars = project.Bars
    summary = bars.Item(1)
    et_summary = summary.ExpandedTask

    # Find a hierarchy bar
    current = et_summary
    leaf_task = None
    for i in range(50):
        try:
            next_et = current.NextTask()
            if next_et is None:
                break
            next_dyn = win32com.client.Dispatch(next_et)
            if next_dyn.HierarchyLevel >= 2:  # Leaf level
                leaf_task = next_dyn
                break
            current = next_dyn
        except Exception:
            break

    if not leaf_task:
        print("  No leaf task found")
        return

    print(f"  Leaf task: ID={leaf_task.ID}, Name={leaf_task.Name[:30]}, Level={leaf_task.HierarchyLevel}")
    print(f"  Start: {leaf_task.Start}")
    print(f"  End: {leaf_task.End}")
    try:
        print(f"  Duration: {leaf_task.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"  Constraint: {leaf_task.Constraint}")
    print(f"  type: {leaf_task.type}")

    # Check if SetUserDuration works on a leaf bar
    original_dur = leaf_task.GetUserDuration().Hours if leaf_task.GetUserDuration() else 0
    print(f"\n  Attempting SetUserDuration(5d=40h) on leaf bar...")
    project.StartTransaction("Set dur leaf")
    try:
        dur_obj = leaf_task.GetDurationFromString("5d")
        leaf_task.SetUserDuration(dur_obj)
        print(f"  SetUserDuration(5d) => OK")
    except Exception as e:
        print(f"  SetUserDuration error: {e}")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    try:
        new_dur = leaf_task.GetUserDuration().Hours
        print(f"  After SetUserDuration + Reschedule:")
        print(f"    Duration: {new_dur}h (was {original_dur}h)")
        print(f"    Start: {leaf_task.Start}")
        print(f"    End: {leaf_task.End}")
        if new_dur != original_dur:
            print(f"    *** DURATION CHANGED! ***")
    except Exception as e:
        print(f"  Error reading: {e}")

    # Restore original duration
    print(f"\n  Restoring original duration ({original_dur}h)...")
    project.StartTransaction("Restore dur")
    try:
        dur_obj = leaf_task.GetDurationFromString(f"{int(original_dur/8)}d")
        leaf_task.SetUserDuration(dur_obj)
    except Exception:
        pass
    project.EndTransaction()
    wait(project)
    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)


def test_create_inside_hierarchy(project):
    """Create new bars INSIDE the hierarchy and test them."""
    print("\n" + "=" * 80)
    print("TEST 6: CREATE BARS INSIDE HIERARCHY")
    print("=" * 80)

    bars = project.Bars
    summary = bars.Item(1)
    et_summary = summary.ExpandedTask

    # Find a summary bar (level 1) to create children under
    current = et_summary
    parent_bar = None
    for i in range(50):
        try:
            next_et = current.NextTask()
            if next_et is None:
                break
            next_dyn = win32com.client.Dispatch(next_et)
            if next_dyn.HierarchyLevel == 1:
                # Check if this is a summary (has children)
                try:
                    cb = next_dyn.ChildBars
                    if cb and cb.Count > 0:
                        parent_bar = next_dyn
                        print(f"  Found summary bar: ID={parent_bar.ID}, "
                              f"Name={parent_bar.Name[:30]}, Level={parent_bar.HierarchyLevel}")
                        print(f"  ChildBars.Count: {cb.Count}")
                        break
                except Exception:
                    pass
            current = next_dyn
        except Exception:
            break

    if not parent_bar:
        print("  No suitable parent bar found. Using project summary.")
        parent_bar = et_summary

    # Try creating a bar via ChildBars.Add() on the parent
    print(f"\n--- ChildBars.Add() on parent ---")
    try:
        parent_bar_obj = parent_bar.Bar
    except Exception:
        parent_bar_obj = summary

    cb = parent_bar.ChildBars
    project.StartTransaction("Add child bar")
    try:
        new_child = cb.Add()
        new_child.Name = "NEW_CHILD_TEST"
        child_id = new_child.ID
        child_et = new_child.ExpandedTask
        print(f"  Created: ID={child_id}")
        print(f"  Level: {child_et.HierarchyLevel}")
        print(f"  Parent: {child_et.Parentname}")
        print(f"  type: {child_et.type}")
        project.EndTransaction()
        wait(project)

        # Check if duration works on this child
        print(f"\n  Testing duration on new child...")
        project.StartTransaction("Set child dur")
        dur_obj = child_et.GetDurationFromString("10d")
        child_et.SetUserDuration(dur_obj)
        print(f"  SetUserDuration(10d) => OK")
        project.EndTransaction()
        wait(project)

        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        try:
            print(f"  Duration: {child_et.GetUserDuration().Hours}h")
        except Exception:
            pass
        print(f"  Start: {child_et.Start}")
        print(f"  End: {child_et.End}")

        # Try to link this child with another
        current = et_summary
        other_task = None
        for i in range(20):
            try:
                next_et = current.NextTask()
                if next_et is None:
                    break
                next_dyn = win32com.client.Dispatch(next_et)
                if next_dyn.ID != child_id and next_dyn.HierarchyLevel >= 1:
                    other_task = next_dyn
                    break
                current = next_dyn
            except Exception:
                break

        if other_task:
            print(f"\n  Linking new child to: ID={other_task.ID}, Name={other_task.Name[:30]}")
            project.StartTransaction("Link new child")
            try:
                link = child_et.LinkTo(other_task)
                print(f"  LinkTo => {link}")
                if link:
                    print(f"  SUCCESS! Link created!")
                project.EndTransaction()
                wait(project)
            except Exception as e:
                print(f"  LinkTo error: {e}")
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass

        # Cleanup - delete the child
        print(f"\n  Deleting child bar...")
        project.StartTransaction("Delete child")
        found = False
        for i in range(bars.Count, 0, -1):
            try:
                b = bars.Item(i)
                if b.ID == child_id:
                    bars.Remove(i)
                    found = True
                    break
            except Exception:
                pass
        if not found:
            # Try cb.Remove
            try:
                for i in range(cb.Count, 0, -1):
                    try:
                        c = cb.Item(i)
                        if c.ID == child_id:
                            cb.Remove(i)
                            found = True
                            break
                    except Exception:
                        pass
            except Exception:
                pass
        project.EndTransaction()
        wait(project)
        print(f"  Deleted: {found}")

    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass


if __name__ == "__main__":
    print("COM Explorer v11 — Hierarchy Bar Access")
    print("=" * 80)
    try:
        app, project = connect()

        # Clean stale bars first
        bars = project.Bars
        stale = []
        for i in range(1, bars.Count + 1):
            b = bars.Item(i)
            if b.Name.startswith(("WF_TEST", "LF_", "DUR_", "DATE_", "TOKEN_",
                                  "LINK_", "TYPE_", "TEST_", "CHILD_", "CONV_",
                                  "WORKFLOW", "DIRECT_", "NEW_CHILD")):
                stale.append(b.ID)
        if stale:
            print(f"Pre-cleanup: {len(stale)} stale bars")
            project.StartTransaction("Cleanup stale")
            for tid in reversed(stale):
                for i in range(bars.Count, 0, -1):
                    try:
                        b = bars.Item(i)
                        if b.ID == tid:
                            bars.Remove(i)
                            break
                    except Exception:
                        pass
            project.EndTransaction()
            wait(project)

        test_childbars_dispatch(project)
        test_next_task(project)
        test_barchartview_bar_access(project)
        test_link_hierarchy_bars(project)
        test_duration_hierarchy_bar(project)
        test_create_inside_hierarchy(project)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
