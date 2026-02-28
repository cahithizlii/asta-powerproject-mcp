"""
COM Explorer v18b — Fix new bar task access (all refs re-fetched after txn)
KEY RULE: After any EndTransaction/AbandonTransaction, ALL COM object references
          must be re-fetched. They may become invalid.
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


def get_child_bar(project, parent_index, child_index=None):
    """Safely get a child bar. Always re-fetches from project.
    parent_index: 1-based index in root's ChildBars
    child_index: if set, gets this child of the parent (1-based)
    """
    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(parent_index))
    if child_index is None:
        return parent_bar
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    return win32com.client.Dispatch(parent_task.ChildBars.Item(child_index))


def find_bar_by_id(project, bar_id, max_depth=5):
    """Find bar by ID via Tasks(1).ChildBars hierarchy. Always fresh."""
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


def test_new_bar_access(project):
    """Create a bar and test task access after commit."""
    print("\n" + "=" * 80)
    print("TEST 1: NEW BAR TASK ACCESS")
    print("=" * 80)

    # Fresh fetch: use Satinalma (index 2) as parent
    parent_bar = get_child_bar(project, 2)
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    parent_id = parent_bar.ID
    print(f"Parent: ID={parent_id}, Name={parent_bar.Name[:40]}")
    print(f"  Parent TaskID (Tasks(1)): {parent_task.ID}")
    print(f"  Parent TaskID (ExpandedTask): {parent_bar.ExpandedTask.ID}")
    orig_child_count = parent_task.ChildBars.Count
    print(f"  ChildBars: {orig_child_count}")

    # Create bar
    print(f"\n--- Create bar ---")
    project.StartTransaction("Create")
    bar_id = None
    try:
        new_bar = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_bar.Name = "V18B_TEST"
        bar_id = new_bar.ID
        print(f"  BarID={bar_id}")

        # In-transaction checks
        try:
            print(f"  In-txn ExpandedTask.ID: {new_bar.ExpandedTask.ID}")
        except Exception as e:
            print(f"  In-txn ExpandedTask: {str(e)[:60]}")
        try:
            t = win32com.client.Dispatch(new_bar.Tasks(1))
            print(f"  In-txn Tasks(1).ID: {t.ID}")
        except Exception as e:
            print(f"  In-txn Tasks(1): {str(e)[:60]}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return None

    # After-transaction checks (ALL fresh fetches)
    print(f"\n--- After EndTransaction (fresh fetches) ---")

    # Re-fetch parent
    parent_bar = get_child_bar(project, 2)
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    new_count = parent_task.ChildBars.Count
    print(f"  Parent ChildBars: {orig_child_count} => {new_count}")

    # Find new bar in parent's children
    found = False
    for i in range(1, new_count + 1):
        cb = win32com.client.Dispatch(parent_task.ChildBars.Item(i))
        if cb.ID == bar_id:
            print(f"  Found at index {i}: ID={cb.ID}, Name={cb.Name[:30]}")
            found = True

            # Test ExpandedTask
            try:
                et = cb.ExpandedTask
                print(f"    ExpandedTask: ID={et.ID}, Name={et.Name[:30]}")
                print(f"    ExpandedTask == parent? {et.ID == parent_task.ID}")
            except Exception as e:
                print(f"    ExpandedTask error: {str(e)[:50]}")

            # Test Tasks(1)
            try:
                t1 = win32com.client.Dispatch(cb.Tasks(1))
                print(f"    Tasks(1): ID={t1.ID}, Name={t1.Name[:30]}")
                print(f"    Tasks(1) == parent? {t1.ID == parent_task.ID}")
            except Exception as e:
                print(f"    Tasks(1) error: {str(e)[:50]}")

            # Test Tasks.Count
            try:
                print(f"    Tasks.Count: {cb.Tasks.Count}")
            except Exception as e:
                print(f"    Tasks.Count error: {str(e)[:50]}")

            break

    if not found:
        print(f"  Bar {bar_id} NOT found in parent's children!")

    # Find via hierarchy search
    print(f"\n  find_bar_by_id({bar_id}):")
    fb = find_bar_by_id(project, bar_id)
    if fb:
        print(f"    Found: ID={fb.ID}")
        try:
            et = fb.ExpandedTask
            print(f"    ExpandedTask: ID={et.ID}")
        except Exception as e:
            print(f"    ExpandedTask: {str(e)[:50]}")
        try:
            t1 = win32com.client.Dispatch(fb.Tasks(1))
            print(f"    Tasks(1): ID={t1.ID}")
        except Exception as e:
            print(f"    Tasks(1): {str(e)[:50]}")
    else:
        print(f"    NOT FOUND")

    return bar_id


def test_duration_and_dates(project, bar_id):
    """Test duration and date setting on the new bar."""
    print("\n" + "=" * 80)
    print("TEST 2: DURATION + DATES ON NEW BAR")
    print("=" * 80)

    bar = find_bar_by_id(project, bar_id)
    if not bar:
        print(f"  Bar {bar_id} not found!")
        return

    # Determine which interface to use
    task = None
    task_type = None

    try:
        task = win32com.client.Dispatch(bar.Tasks(1))
        task_type = "Tasks(1)"
    except Exception:
        try:
            task = bar.ExpandedTask
            task_type = "ExpandedTask"
        except Exception:
            print("  No task access method works!")
            return

    print(f"  Using: {task_type}, ID={task.ID}")
    print(f"  Current: Start={task.Start}, End={task.End}")
    try:
        print(f"  Duration: {task.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"  Duration: {str(e)[:50]}")

    # Set duration
    print(f"\n--- SetUserDuration(10d) ---")
    project.StartTransaction("Dur")
    try:
        bar = find_bar_by_id(project, bar_id)
        task = win32com.client.Dispatch(bar.Tasks(1)) if task_type == "Tasks(1)" else bar.ExpandedTask
        dur = task.GetDurationFromString("10d")
        task.SetUserDuration(dur)
        print(f"  OK")
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

    bar = find_bar_by_id(project, bar_id)
    task = win32com.client.Dispatch(bar.Tasks(1)) if task_type == "Tasks(1)" else bar.ExpandedTask
    print(f"  After reschedule: Dur={task.GetUserDuration().Hours}h, Start={task.Start}, End={task.End}")

    # Set ImposedStart
    print(f"\n--- ImposedStart(2026-07-01) ---")
    project.StartTransaction("Start")
    try:
        bar = find_bar_by_id(project, bar_id)
        et = bar.ExpandedTask
        et.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
        print(f"  OK (ExpandedTask.ID={et.ID})")
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

    bar = find_bar_by_id(project, bar_id)
    et = bar.ExpandedTask
    print(f"  After: Start={et.Start}, End={et.End}, Constraint={et.Constraint}")


def test_link_workflow(project, bar_a_id):
    """Create second bar and link them."""
    print("\n" + "=" * 80)
    print("TEST 3: LINK WORKFLOW")
    print("=" * 80)

    # Create bar B under same parent
    parent_bar = get_child_bar(project, 2)
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))

    project.StartTransaction("Create B")
    bar_b_id = None
    try:
        new_b = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_b.Name = "V18B_TEST_B"
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
        return None

    # Set duration on B
    project.StartTransaction("Dur B")
    try:
        bar_b = find_bar_by_id(project, bar_b_id)
        task_b = win32com.client.Dispatch(bar_b.Tasks(1))
        dur = task_b.GetDurationFromString("5d")
        task_b.SetUserDuration(dur)
        print(f"  B dur(5d) OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  B dur error: {e}")
        # Try with ExpandedTask
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        project.StartTransaction("Dur B v2")
        try:
            bar_b = find_bar_by_id(project, bar_b_id)
            et_b = bar_b.ExpandedTask
            dur = et_b.GetDurationFromString("5d")
            et_b.SetUserDuration(dur)
            print(f"  B dur(5d) via ExpandedTask OK")
            project.EndTransaction()
            wait(project)
        except Exception as e2:
            print(f"  B dur ExpandedTask error: {e2}")
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
        bar_a = find_bar_by_id(project, bar_a_id)
        bar_b = find_bar_by_id(project, bar_b_id)

        # Try Tasks(1) first for linking
        try:
            task_a = win32com.client.Dispatch(bar_a.Tasks(1))
            task_b = win32com.client.Dispatch(bar_b.Tasks(1))
            print(f"  Using Tasks(1): A.ID={task_a.ID}, B.ID={task_b.ID}")
            link = task_a.LinkTo(task_b)
        except Exception:
            et_a = bar_a.ExpandedTask
            et_b = bar_b.ExpandedTask
            print(f"  Using ExpandedTask: A.ID={et_a.ID}, B.ID={et_b.ID}")
            if et_a.ID == et_b.ID:
                print(f"  ERROR: Same ExpandedTask ID! Can't link!")
                project.AbandonTransaction()
                return bar_b_id
            link = et_a.LinkTo(et_b)

        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  LINK CREATED! ID={ld.ID}, type={ld.type}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Link error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Verify
    print(f"\n--- VERIFY ---")
    bar_a = find_bar_by_id(project, bar_a_id)
    bar_b = find_bar_by_id(project, bar_b_id)
    try:
        task_a = win32com.client.Dispatch(bar_a.Tasks(1))
        task_b = win32com.client.Dispatch(bar_b.Tasks(1))
        print(f"  A: Start={task_a.Start}, End={task_a.End}, LinksOut={task_a.LinksOut.Count}")
        print(f"  B: Start={task_b.Start}, End={task_b.End}, LinksIn={task_b.LinksIn.Count}")
    except Exception:
        et_a = bar_a.ExpandedTask
        et_b = bar_b.ExpandedTask
        print(f"  A: Start={et_a.Start}, End={et_a.End}, LinksOut={et_a.LinksOut.Count}")
        print(f"  B: Start={et_b.Start}, End={et_b.End}, LinksIn={et_b.LinksIn.Count}")

    return bar_b_id


def cleanup(project, bar_ids):
    """Remove test bars (fresh fetches each time)."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup ---")
    for bid in reversed(bar_ids):
        bar = find_bar_by_id(project, bid)
        if not bar:
            print(f"  Bar {bid} not found")
            continue
        # Find in parent
        parent_bar = get_child_bar(project, 2)
        parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
        project.StartTransaction(f"Del {bid}")
        try:
            cb = parent_task.ChildBars
            removed = False
            for i in range(cb.Count, 0, -1):
                b = win32com.client.Dispatch(cb.Item(i))
                if b.ID == bid:
                    cb.Remove(i)
                    print(f"  Removed bar ID={bid}")
                    removed = True
                    break
            if not removed:
                print(f"  Bar {bid} not in parent's ChildBars")
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  Error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass


if __name__ == "__main__":
    print("COM Explorer v18b — New Bar Task Access (Fixed)")
    print("=" * 80)
    created_ids = []
    try:
        app, project = connect()

        # Pre-cleanup
        print("\n--- Pre-cleanup ---")
        root_task = win32com.client.Dispatch(project.Bars.Item(1).Tasks(1))
        stale = []
        for i in range(1, root_task.ChildBars.Count + 1):
            cb = win32com.client.Dispatch(root_task.ChildBars.Item(i))
            try:
                ct = win32com.client.Dispatch(cb.Tasks(1))
                for j in range(1, ct.ChildBars.Count + 1):
                    sb = win32com.client.Dispatch(ct.ChildBars.Item(j))
                    if sb.Name.startswith(("V14_", "V15_", "V16_", "V17_", "V18")):
                        stale.append(sb.ID)
                        print(f"  Stale: ID={sb.ID}, Name={sb.Name}")
            except Exception:
                pass
        if stale:
            cleanup(project, stale)

        # Test 1
        bar_a_id = test_new_bar_access(project)

        # Test 2
        if bar_a_id:
            created_ids.append(bar_a_id)
            test_duration_and_dates(project, bar_a_id)

        # Test 3
        if bar_a_id:
            bar_b_id = test_link_workflow(project, bar_a_id)
            if bar_b_id:
                created_ids.append(bar_b_id)

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
