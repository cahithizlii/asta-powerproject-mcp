"""
COM Explorer v14 — Confirmed working methods + focused tests
CONFIRMED from v13:
  - ChildBars navigation works (deep hierarchy)
  - task.LinkTo(task2) WORKS on leaf tasks
  - LinksOut.Remove(index) works
  - Link type=0 is FS

NOW TESTING:
  1. SetUserDuration on non-milestone leaf task (dur > 0)
  2. Date setting (SetUserStart, MoveToDate, ImposedStart) on leaf tasks
  3. ChildBars.Add() under a summary + LinkTo between new bars
  4. Link type setting (SS, FF, SF)
  5. Lag setting on links
  6. Re-fetch objects after each transaction to avoid "Object is no longer valid"
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


def find_task_by_id(project, task_id):
    """Find a task by ID by traversing hierarchy."""
    bar = project.Bars.Item(1)
    task = win32com.client.Dispatch(bar.Tasks(1))
    return _search_task(task, task_id, max_depth=5)


def _search_task(parent_task, target_id, depth=0, max_depth=5):
    """Recursive search for task by ID."""
    if parent_task.ID == target_id:
        return parent_task
    if depth >= max_depth:
        return None
    try:
        child_bars = parent_task.ChildBars
        for i in range(1, child_bars.Count + 1):
            child_bar = win32com.client.Dispatch(child_bars.Item(i))
            child_task = win32com.client.Dispatch(child_bar.Tasks(1))
            result = _search_task(child_task, target_id, depth + 1, max_depth)
            if result:
                return result
    except Exception:
        pass
    return None


def find_leaf_tasks_with_duration(project, min_count=5):
    """Find leaf tasks with duration > 0 (actual tasks, not milestones)."""
    bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(bar.Tasks(1))
    results = []
    _collect_leaves(root_task, results, max_depth=4, want_duration=True)
    return results[:min_count]


def _collect_leaves(task, results, depth=0, max_depth=4, want_duration=False):
    """Collect leaf tasks recursively."""
    try:
        child_bars = task.ChildBars
        child_count = child_bars.Count if child_bars else 0
    except Exception:
        child_count = 0

    if child_count == 0:
        # Leaf task
        try:
            dur = task.GetUserDuration().Hours
        except Exception:
            dur = 0
        if want_duration and dur <= 0:
            return  # Skip milestones
        results.append({
            'task_id': task.ID,
            'bar_id': task.Bar.ID,
            'name': task.Name[:40],
            'duration': dur,
            'links_in': task.LinksIn.Count,
            'links_out': task.LinksOut.Count,
        })
        return

    if depth >= max_depth:
        return

    for i in range(1, min(child_count + 1, 20)):
        if len(results) >= 20:
            break
        try:
            child_bar = win32com.client.Dispatch(child_bars.Item(i))
            child_task = win32com.client.Dispatch(child_bar.Tasks(1))
            _collect_leaves(child_task, results, depth + 1, max_depth, want_duration)
        except Exception:
            pass


def test_duration_real_task(project):
    """Test SetUserDuration on a real task (non-milestone)."""
    print("\n" + "=" * 80)
    print("TEST 1: SetUserDuration ON REAL TASK (dur > 0)")
    print("=" * 80)

    leaves = find_leaf_tasks_with_duration(project, 5)
    if not leaves:
        print("  No leaf tasks with duration found!")
        return

    print(f"Found {len(leaves)} leaf tasks with duration:")
    for r in leaves:
        print(f"  TaskID={r['task_id']}, Name={r['name']}, Dur={r['duration']}h")

    # Pick first one
    r = leaves[0]
    task = find_task_by_id(project, r['task_id'])
    if not task:
        print(f"  Can't re-find task {r['task_id']}")
        return

    task_id = task.ID
    print(f"\nUsing: TaskID={task_id}, Name={task.Name[:40]}")
    print(f"  Start: {task.Start}")
    print(f"  End: {task.End}")
    print(f"  Duration: {task.GetUserDuration().Hours}h")
    print(f"  Constraint: {task.Constraint}")

    orig_dur = task.GetUserDuration().Hours

    # Set new duration
    print(f"\n--- SetUserDuration('15d') ---")
    project.StartTransaction("SetDur real")
    try:
        dur_obj = task.GetDurationFromString("15d")
        print(f"  GetDurationFromString('15d') => {dur_obj.Hours}h")
        task.SetUserDuration(dur_obj)
        print(f"  SetUserDuration => OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return

    # Re-fetch task (may be invalidated after transaction)
    task = find_task_by_id(project, task_id)
    if not task:
        print(f"  Can't re-find task after EndTransaction!")
        return

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Re-fetch again after reschedule
    task = find_task_by_id(project, task_id)
    if not task:
        print(f"  Can't re-find task after reschedule!")
        return

    new_dur = task.GetUserDuration().Hours
    print(f"  After reschedule: Duration={new_dur}h")
    print(f"  Start: {task.Start}")
    print(f"  End: {task.End}")

    if new_dur != orig_dur:
        print(f"  *** DURATION CHANGED! {orig_dur}h => {new_dur}h ***")
    else:
        print(f"  Duration UNCHANGED: {orig_dur}h")

    # Restore
    print(f"\n--- Restore original duration ({orig_dur}h) ---")
    project.StartTransaction("Restore dur")
    try:
        task = find_task_by_id(project, task_id)
        dur_str = f"{int(orig_dur)}h"
        dur_obj = task.GetDurationFromString(dur_str)
        task.SetUserDuration(dur_obj)
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        task = find_task_by_id(project, task_id)
        print(f"  Restored: Duration={task.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"  Restore error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def test_dates_real_task(project):
    """Test date setting on a real task."""
    print("\n" + "=" * 80)
    print("TEST 2: DATE SETTING ON REAL TASK")
    print("=" * 80)

    leaves = find_leaf_tasks_with_duration(project, 5)
    if not leaves:
        print("  No leaf tasks found!")
        return

    # Pick a task that has no predecessor (so date change won't be overridden)
    unlinked = [r for r in leaves if r['links_in'] == 0]
    if unlinked:
        r = unlinked[0]
    else:
        r = leaves[0]

    task = find_task_by_id(project, r['task_id'])
    if not task:
        print(f"  Can't find task {r['task_id']}")
        return

    task_id = task.ID
    print(f"Using: TaskID={task_id}, Name={task.Name[:40]}")
    print(f"  Start: {task.Start}")
    print(f"  End: {task.End}")
    print(f"  Duration: {task.GetUserDuration().Hours}h")
    print(f"  Constraint: {task.Constraint}")
    print(f"  LinksIn: {task.LinksIn.Count}")

    orig_start = task.Start
    orig_constraint = task.Constraint

    # Test A: ImposedStart
    print(f"\n--- Test A: ImposedStart ---")
    project.StartTransaction("ImposedStart")
    try:
        task.ImposedStart = pywintypes.Time(datetime(2026, 8, 3))
        print(f"  ImposedStart = 2026-08-03 => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        task = find_task_by_id(project, task_id)
        print(f"  After reschedule: Start={task.Start}, End={task.End}, Constraint={task.Constraint}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test B: MoveToDate
    print(f"\n--- Test B: MoveToDate ---")
    project.StartTransaction("MoveToDate")
    try:
        task = find_task_by_id(project, task_id)
        task.RemoveConstraint()
        task.MoveToDate(pywintypes.Time(datetime(2026, 9, 1)))
        print(f"  MoveToDate(2026-09-01) => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        task = find_task_by_id(project, task_id)
        print(f"  After reschedule: Start={task.Start}, End={task.End}, Constraint={task.Constraint}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test C: SetUserStart
    print(f"\n--- Test C: SetUserStart ---")
    project.StartTransaction("SetUserStart")
    try:
        task = find_task_by_id(project, task_id)
        task.RemoveConstraint()
        task.SetUserStart(pywintypes.Time(datetime(2026, 10, 1)))
        print(f"  SetUserStart(2026-10-01) => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        task = find_task_by_id(project, task_id)
        print(f"  After reschedule: Start={task.Start}, End={task.End}, Constraint={task.Constraint}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Restore
    print(f"\n--- Restoring original ---")
    project.StartTransaction("Restore")
    try:
        task = find_task_by_id(project, task_id)
        task.RemoveConstraint()
        if orig_constraint > 0:
            task.ImposedStart = orig_start
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        task = find_task_by_id(project, task_id)
        print(f"  Restored: Start={task.Start}, Constraint={task.Constraint}")
    except Exception as e:
        print(f"  Restore error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def test_create_and_link(project):
    """Create two bars under a summary and link them."""
    print("\n" + "=" * 80)
    print("TEST 3: CREATE BARS + LINK + DURATION")
    print("=" * 80)

    # Navigate to a summary bar
    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(1))
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    print(f"Parent: ID={parent_bar.ID}, Name={parent_bar.Name[:40]}")
    print(f"  ChildBars before: {parent_task.ChildBars.Count}")

    # Create bar A
    project.StartTransaction("Create A")
    try:
        new_a = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_a.Name = "V14_TEST_A"
        bar_a_id = new_a.ID
        task_a = win32com.client.Dispatch(new_a.Tasks(1))
        task_a_id = task_a.ID
        print(f"  Created A: BarID={bar_a_id}, TaskID={task_a_id}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Create A error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return []

    # Create bar B
    project.StartTransaction("Create B")
    try:
        parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
        new_b = win32com.client.Dispatch(parent_task.ChildBars.Add())
        new_b.Name = "V14_TEST_B"
        bar_b_id = new_b.ID
        task_b = win32com.client.Dispatch(new_b.Tasks(1))
        task_b_id = task_b.ID
        print(f"  Created B: BarID={bar_b_id}, TaskID={task_b_id}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Create B error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return [bar_a_id]

    # Set duration on A (10d) and B (5d)
    print(f"\n--- Set durations ---")
    project.StartTransaction("Set durs")
    try:
        ta = find_task_by_id(project, task_a_id)
        dur_a = ta.GetDurationFromString("10d")
        ta.SetUserDuration(dur_a)
        print(f"  A SetUserDuration(10d) => OK")

        tb = find_task_by_id(project, task_b_id)
        dur_b = tb.GetDurationFromString("5d")
        tb.SetUserDuration(dur_b)
        print(f"  B SetUserDuration(5d) => OK")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Duration error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Check durations
    ta = find_task_by_id(project, task_a_id)
    tb = find_task_by_id(project, task_b_id)
    if ta:
        print(f"  A: Duration={ta.GetUserDuration().Hours}h, Start={ta.Start}, End={ta.End}")
    if tb:
        print(f"  B: Duration={tb.GetUserDuration().Hours}h, Start={tb.Start}, End={tb.End}")

    # Set ImposedStart on A
    print(f"\n--- ImposedStart on A ---")
    project.StartTransaction("ImposedStart A")
    try:
        ta = find_task_by_id(project, task_a_id)
        ta.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
        print(f"  ImposedStart = 2026-07-01 => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        ta = find_task_by_id(project, task_a_id)
        print(f"  A: Start={ta.Start}, End={ta.End}")
    except Exception as e:
        print(f"  ImposedStart error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # LinkTo: A -> B (FS)
    print(f"\n--- LinkTo A -> B ---")
    project.StartTransaction("LinkTo A->B")
    try:
        ta = find_task_by_id(project, task_a_id)
        tb = find_task_by_id(project, task_b_id)
        link = ta.LinkTo(tb)
        print(f"  LinkTo result: {link}")
        if link:
            ld = win32com.client.Dispatch(link)
            link_id = ld.ID
            print(f"  SUCCESS! Link ID={link_id}, type={ld.type}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  LinkTo error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return [bar_a_id, bar_b_id]

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Check result
    ta = find_task_by_id(project, task_a_id)
    tb = find_task_by_id(project, task_b_id)
    print(f"\n  After Link + Reschedule:")
    print(f"    A: Start={ta.Start}, End={ta.End}, LinksOut={ta.LinksOut.Count}")
    print(f"    B: Start={tb.Start}, End={tb.End}, LinksIn={tb.LinksIn.Count}")
    print(f"    B should start after A ends!")

    # Test link type change: set to SS (type=1)
    print(f"\n--- Change link type to SS ---")
    project.StartTransaction("Link type SS")
    try:
        ta = find_task_by_id(project, task_a_id)
        link_obj = win32com.client.Dispatch(ta.LinksOut.Item(1))
        print(f"  Current type: {link_obj.type}")
        link_obj.type = 1  # SS
        print(f"  Set type=1 (SS) => OK, readback: {link_obj.type}")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        ta = find_task_by_id(project, task_a_id)
        tb = find_task_by_id(project, task_b_id)
        print(f"  After SS: A Start={ta.Start}, B Start={tb.Start}")
    except Exception as e:
        print(f"  Link type error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test lag setting
    print(f"\n--- Set lag on link ---")
    project.StartTransaction("Link lag")
    try:
        ta = find_task_by_id(project, task_a_id)
        link_obj = win32com.client.Dispatch(ta.LinksOut.Item(1))
        # Set back to FS first
        link_obj.type = 0  # FS

        # Try setting lag
        lag_dur = ta.GetDurationFromString("2d")
        print(f"  lag_dur = {lag_dur.Hours}h")
        link_obj.StartLagTime = lag_dur
        print(f"  StartLagTime = 2d => OK")
        print(f"  Readback: StartLagTime={link_obj.StartLagTime.Hours}h")

        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        ta = find_task_by_id(project, task_a_id)
        tb = find_task_by_id(project, task_b_id)
        print(f"  After lag: A End={ta.End}, B Start={tb.Start}")
        link_obj = win32com.client.Dispatch(ta.LinksOut.Item(1))
        print(f"  Link: type={link_obj.type}, lag={link_obj.StartLagTime.Hours}h")
    except Exception as e:
        print(f"  Lag error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    return [bar_a_id, bar_b_id]


def test_link_type_enum(project):
    """Test all link type values."""
    print("\n" + "=" * 80)
    print("TEST 4: LINK TYPE ENUM VALUES")
    print("=" * 80)

    # Check existing links to see what type values are used
    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))

    type_counts = {}
    count = 0
    _count_link_types(root_task, type_counts, max_depth=3, counter=[0])

    print(f"Link type distribution:")
    for t, c in sorted(type_counts.items()):
        label = {0: "FS", 1: "SS", 2: "FF", 3: "SF"}.get(t, f"?{t}")
        print(f"  type={t} ({label}): {c} links")


def _count_link_types(task, type_counts, depth=0, max_depth=3, counter=None):
    """Count link types across the hierarchy."""
    if counter[0] > 500:
        return
    counter[0] += 1

    # Check this task's links
    try:
        for i in range(1, task.LinksOut.Count + 1):
            link = win32com.client.Dispatch(task.LinksOut.Item(i))
            t = link.type
            type_counts[t] = type_counts.get(t, 0) + 1
    except Exception:
        pass

    if depth >= max_depth:
        return

    try:
        child_bars = task.ChildBars
        for i in range(1, min(child_bars.Count + 1, 30)):
            try:
                cb = win32com.client.Dispatch(child_bars.Item(i))
                ct = win32com.client.Dispatch(cb.Tasks(1))
                _count_link_types(ct, type_counts, depth + 1, max_depth, counter)
            except Exception:
                pass
    except Exception:
        pass


def cleanup(project, bar_ids):
    """Remove test bars."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup: {len(bar_ids)} bars ---")

    # Find parent
    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(1))
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))

    project.StartTransaction("Cleanup")
    child_bars = parent_task.ChildBars
    for tid in reversed(bar_ids):
        for i in range(child_bars.Count, 0, -1):
            try:
                b = win32com.client.Dispatch(child_bars.Item(i))
                if b.ID == tid:
                    child_bars.Remove(i)
                    print(f"  Removed bar ID={tid}")
                    break
            except Exception:
                pass
    project.EndTransaction()
    wait(project)


if __name__ == "__main__":
    print("COM Explorer v14 — Duration + Dates + Create + Link + Types")
    print("=" * 80)
    created_ids = []
    try:
        app, project = connect()

        # Pre-cleanup
        print("\n--- Pre-cleanup stale test bars ---")
        root_bar = project.Bars.Item(1)
        root_task = win32com.client.Dispatch(root_bar.Tasks(1))
        parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(1))
        parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
        stale = []
        for i in range(1, parent_task.ChildBars.Count + 1):
            try:
                b = win32com.client.Dispatch(parent_task.ChildBars.Item(i))
                if b.Name.startswith(("V14_TEST", "TEST_CHILD", "V13_")):
                    stale.append(b.ID)
                    print(f"  Stale: ID={b.ID}, Name={b.Name}")
            except Exception:
                pass
        if stale:
            cleanup(project, stale)

        # Test 1: Duration on real task
        test_duration_real_task(project)

        # Test 2: Dates on real task
        test_dates_real_task(project)

        # Test 3: Create + Link + Duration
        created_ids = test_create_and_link(project)

        # Test 4: Link type enum
        test_link_type_enum(project)

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
