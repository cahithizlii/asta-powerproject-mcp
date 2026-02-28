"""
COM Explorer v18 — Fix new bar task access
PROBLEM: new_bar.ExpandedTask returns PARENT's task, not new bar's task
HYPOTHESIS: After EndTransaction, bar.Tasks(1) should work on committed bars

TESTS:
  1. Create bar under Satinalma (not Milestones), verify Tasks(1) vs ExpandedTask
  2. If Tasks(1) works after commit: full workflow with duration + dates + link
  3. If not: try other approaches (re-fetch, dynamic dispatch, etc.)
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
    """Root IExpandedTask via bar.Tasks(1) — correct hierarchy."""
    bar = project.Bars.Item(1)
    return win32com.client.Dispatch(bar.Tasks(1))


def find_bar_by_id(project, bar_id, max_depth=5):
    """Find bar by ID via Tasks(1).ChildBars hierarchy."""
    root_bar = project.Bars.Item(1)
    if root_bar.ID == bar_id:
        return root_bar
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    return _search_bar(root_task, bar_id, 0, max_depth)


def _search_bar(parent_task, target_id, depth, max_depth):
    try:
        child_bars = parent_task.ChildBars
        for i in range(1, child_bars.Count + 1):
            cb = win32com.client.Dispatch(child_bars.Item(i))
            if cb.ID == target_id:
                return cb
            if depth < max_depth:
                try:
                    ct = win32com.client.Dispatch(cb.Tasks(1))
                    result = _search_bar(ct, target_id, depth + 1, max_depth)
                    if result:
                        return result
                except Exception:
                    pass
    except Exception:
        pass
    return None


def test_new_bar_task_access(project):
    """Create a bar and test all ways to access its task."""
    print("\n" + "=" * 80)
    print("TEST 1: NEW BAR TASK ACCESS METHODS")
    print("=" * 80)

    root_task = get_root_task(project)

    # Use Satinalma (index 2) as parent, not Milestones
    parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(2))
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    print(f"Parent: ID={parent_bar.ID}, Name={parent_bar.Name[:40]}")
    print(f"  Parent TaskID via Tasks(1): {parent_task.ID}")
    print(f"  Parent TaskID via ExpandedTask: {parent_bar.ExpandedTask.ID}")
    print(f"  ChildBars: {parent_task.ChildBars.Count}")

    # Create bar
    print(f"\n--- Creating bar ---")
    project.StartTransaction("Create test")
    try:
        new_bar = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_bar.Name = "V18_TEST_BAR"
        bar_id = new_bar.ID
        print(f"  BarID={bar_id}")

        # Check ExpandedTask INSIDE transaction
        try:
            et_in = new_bar.ExpandedTask
            print(f"  ExpandedTask (in-txn): ID={et_in.ID}, Name={et_in.Name[:30]}")
        except Exception as e:
            print(f"  ExpandedTask (in-txn): {str(e)[:50]}")

        # Check Tasks(1) INSIDE transaction
        try:
            t1_in = win32com.client.Dispatch(new_bar.Tasks(1))
            print(f"  Tasks(1) (in-txn): ID={t1_in.ID}, Name={t1_in.Name[:30]}")
        except Exception as e:
            print(f"  Tasks(1) (in-txn): {str(e)[:50]}")

        # Check Tasks.Count INSIDE transaction
        try:
            tc_in = new_bar.Tasks.Count
            print(f"  Tasks.Count (in-txn): {tc_in}")
        except Exception as e:
            print(f"  Tasks.Count (in-txn): {str(e)[:50]}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Create error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return None

    # Now test AFTER transaction
    print(f"\n--- After EndTransaction ---")

    # Method 1: Direct bar reference
    print(f"  Direct bar ref (new_bar):")
    try:
        et_direct = new_bar.ExpandedTask
        print(f"    ExpandedTask: ID={et_direct.ID}, Name={et_direct.Name[:30]}")
    except Exception as e:
        print(f"    ExpandedTask: {str(e)[:50]}")

    try:
        t1_direct = win32com.client.Dispatch(new_bar.Tasks(1))
        print(f"    Tasks(1): ID={t1_direct.ID}, Name={t1_direct.Name[:30]}")
    except Exception as e:
        print(f"    Tasks(1): {str(e)[:50]}")

    # Method 2: Find via hierarchy
    print(f"\n  Find via hierarchy (find_bar_by_id):")
    found_bar = find_bar_by_id(project, bar_id)
    if found_bar:
        print(f"    Found bar: ID={found_bar.ID}, Name={found_bar.Name[:30]}")
        try:
            et_found = found_bar.ExpandedTask
            print(f"    ExpandedTask: ID={et_found.ID}, Name={et_found.Name[:30]}")
        except Exception as e:
            print(f"    ExpandedTask: {str(e)[:50]}")

        try:
            t1_found = win32com.client.Dispatch(found_bar.Tasks(1))
            print(f"    Tasks(1): ID={t1_found.ID}, Name={t1_found.Name[:30]}")
        except Exception as e:
            print(f"    Tasks(1): {str(e)[:50]}")
    else:
        print(f"    NOT FOUND!")

    # Method 3: Find in parent's ChildBars directly
    print(f"\n  Find in parent's ChildBars:")
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    for i in range(1, parent_task.ChildBars.Count + 1):
        cb = win32com.client.Dispatch(parent_task.ChildBars.Item(i))
        if cb.ID == bar_id:
            print(f"    Found at index {i}: ID={cb.ID}, Name={cb.Name[:30]}")
            try:
                et_cb = cb.ExpandedTask
                print(f"    ExpandedTask: ID={et_cb.ID}, Name={et_cb.Name[:30]}")
            except Exception as e:
                print(f"    ExpandedTask: {str(e)[:50]}")
            try:
                t1_cb = win32com.client.Dispatch(cb.Tasks(1))
                print(f"    Tasks(1): ID={t1_cb.ID}, Name={t1_cb.Name[:30]}")
            except Exception as e:
                print(f"    Tasks(1): {str(e)[:50]}")
            break

    # Method 4: Dynamic dispatch (bypass gen_py cache)
    print(f"\n  Dynamic dispatch:")
    try:
        dyn_bar = win32com.client.dynamic.Dispatch(found_bar._oleobj_) if found_bar else None
        if dyn_bar:
            try:
                dyn_et = dyn_bar.ExpandedTask
                print(f"    ExpandedTask: {dyn_et}, ID={dyn_et.ID}")
            except Exception as e:
                print(f"    ExpandedTask: {str(e)[:50]}")
            try:
                dyn_t1 = dyn_bar.Tasks(1)
                print(f"    Tasks(1): {dyn_t1}, ID={dyn_t1.ID}")
            except Exception as e:
                print(f"    Tasks(1): {str(e)[:50]}")
    except Exception as e:
        print(f"    Dynamic dispatch error: {str(e)[:50]}")

    # Method 5: Use IBarChartView to get bar
    print(f"\n  IBarChartView access:")
    try:
        bcv = win32com.client.Dispatch(project.CurrentView)
        all_ids = bcv.AllBarIds()
        if bar_id in all_ids:
            print(f"    Bar {bar_id} IS in AllBarIds ({len(all_ids)} total)")
        else:
            print(f"    Bar {bar_id} NOT in AllBarIds!")
    except Exception as e:
        print(f"    Error: {str(e)[:50]}")

    return bar_id


def test_full_workflow_fixed(project, test_bar_id):
    """Full workflow using correct task access method."""
    print("\n" + "=" * 80)
    print("TEST 2: FULL WORKFLOW (using whichever method works)")
    print("=" * 80)

    if not test_bar_id:
        print("  No test bar available")
        return []

    # First, determine which method works for task access
    found_bar = find_bar_by_id(project, test_bar_id)
    if not found_bar:
        print(f"  Can't find bar {test_bar_id}")
        return []

    # Try Tasks(1)
    task = None
    try:
        task = win32com.client.Dispatch(found_bar.Tasks(1))
        print(f"  Tasks(1) works! TaskID={task.ID}")
    except Exception:
        pass

    if not task:
        # Try ExpandedTask
        try:
            task = found_bar.ExpandedTask
            print(f"  ExpandedTask: ID={task.ID}")
            # Verify it's not the parent
            root_task = get_root_task(project)
            parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(2))
            parent_et = parent_bar.ExpandedTask
            if task.ID == parent_et.ID:
                print(f"  WARNING: ExpandedTask is PARENT ({parent_et.ID})!")
                task = None
        except Exception:
            pass

    if not task:
        print(f"  No valid task access method!")
        return []

    # Set duration
    print(f"\n--- Set Duration (10d) ---")
    project.StartTransaction("Duration")
    try:
        task = win32com.client.Dispatch(found_bar.Tasks(1))
        dur = task.GetDurationFromString("10d")
        task.SetUserDuration(dur)
        print(f"  SetUserDuration(10d) => OK")
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

    found_bar = find_bar_by_id(project, test_bar_id)
    try:
        task = win32com.client.Dispatch(found_bar.Tasks(1))
        dur_h = task.GetUserDuration().Hours
        print(f"  Duration after reschedule: {dur_h}h")
        print(f"  Start: {task.Start}")
        print(f"  End: {task.End}")
        if dur_h == 80.0:
            print(f"  *** DURATION CORRECT (10d = 80h) ***")
    except Exception as e:
        print(f"  Check error: {e}")

    # Set date via ImposedStart (on ExpandedTask)
    print(f"\n--- ImposedStart (2026-07-01) ---")
    project.StartTransaction("ImposedStart")
    try:
        found_bar = find_bar_by_id(project, test_bar_id)
        et = found_bar.ExpandedTask
        print(f"  ExpandedTask ID: {et.ID}")
        et.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
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

    found_bar = find_bar_by_id(project, test_bar_id)
    et = found_bar.ExpandedTask
    print(f"  After: Start={et.Start}, End={et.End}, Constraint={et.Constraint}")

    # Create second bar + link
    print(f"\n--- Create second bar ---")
    root_task = get_root_task(project)
    parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(2))
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))

    project.StartTransaction("Create B")
    try:
        new_b = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_b.Name = "V18_TEST_B"
        bar_b_id = new_b.ID
        print(f"  Created B: BarID={bar_b_id}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return [test_bar_id]

    # Set duration on B
    project.StartTransaction("Dur B")
    try:
        bar_b = find_bar_by_id(project, bar_b_id)
        task_b = win32com.client.Dispatch(bar_b.Tasks(1))
        dur_b = task_b.GetDurationFromString("5d")
        task_b.SetUserDuration(dur_b)
        print(f"  B SetUserDuration(5d) => OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  B dur error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Link A -> B
    print(f"\n--- Link A -> B ---")
    project.StartTransaction("Link")
    try:
        bar_a = find_bar_by_id(project, test_bar_id)
        bar_b = find_bar_by_id(project, bar_b_id)
        et_a = bar_a.ExpandedTask
        et_b = bar_b.ExpandedTask
        print(f"  A: ExpandedTask ID={et_a.ID}")
        print(f"  B: ExpandedTask ID={et_b.ID}")

        # Verify they're different tasks
        if et_a.ID == et_b.ID:
            print(f"  ERROR: Same ExpandedTask ID! Using Tasks(1) instead...")
            task_a = win32com.client.Dispatch(bar_a.Tasks(1))
            task_b = win32com.client.Dispatch(bar_b.Tasks(1))
            print(f"  A Tasks(1) ID={task_a.ID}, B Tasks(1) ID={task_b.ID}")
            link = task_a.LinkTo(task_b)
        else:
            link = et_a.LinkTo(et_b)

        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  LINK CREATED! ID={ld.ID}, type={ld.type}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Link error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Verify
    print(f"\n--- VERIFY ---")
    bar_a = find_bar_by_id(project, test_bar_id)
    bar_b = find_bar_by_id(project, bar_b_id)
    et_a = bar_a.ExpandedTask
    et_b = bar_b.ExpandedTask

    try:
        task_a = win32com.client.Dispatch(bar_a.Tasks(1))
        task_b = win32com.client.Dispatch(bar_b.Tasks(1))
        print(f"  A: TaskID={task_a.ID}, Start={task_a.Start}, End={task_a.End}")
        print(f"     Dur={task_a.GetUserDuration().Hours}h, LinksOut={task_a.LinksOut.Count}")
        print(f"  B: TaskID={task_b.ID}, Start={task_b.Start}, End={task_b.End}")
        print(f"     Dur={task_b.GetUserDuration().Hours}h, LinksIn={task_b.LinksIn.Count}")

        if task_b.Start >= task_a.End:
            print(f"\n  *** SUCCESS: B starts after A ends! ***")
        else:
            print(f"\n  NOTE: B.Start < A.End (may need constraint)")
    except Exception as e:
        print(f"  Verify error: {e}")

    return [test_bar_id, bar_b_id]


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
        try:
            et = bar.ExpandedTask
            parent_bar = win32com.client.Dispatch(et.GetActualParentBar())
            parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
        except Exception:
            # Try hierarchy approach
            root_task = get_root_task(project)
            parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(2))
            parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))

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
    print("COM Explorer v18 — Fix New Bar Task Access")
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
            try:
                ct = win32com.client.Dispatch(cb.Tasks(1))
                for j in range(1, ct.ChildBars.Count + 1):
                    sb = win32com.client.Dispatch(ct.ChildBars.Item(j))
                    if sb.Name.startswith(("V14_", "V15_", "V16_", "V17_", "V18_", "TEST_")):
                        stale.append(sb.ID)
                        print(f"  Stale: ID={sb.ID}, Name={sb.Name}")
            except Exception:
                pass
        if stale:
            cleanup(project, stale)

        # Test 1: Task access methods
        test_bar_id = test_new_bar_task_access(project)

        # Test 2: Full workflow
        if test_bar_id:
            created_ids = test_full_workflow_fixed(project, test_bar_id)

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
