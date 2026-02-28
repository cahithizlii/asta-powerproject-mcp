"""
COM Explorer v13 — Deep hierarchy traversal + LinkTo on leaf tasks
1. Recursively navigate ChildBars to find leaf tasks (no children)
2. Check which tasks have links (LinksIn/LinksOut > 0)
3. Test LinkTo between leaf-level ITaskBase objects
4. Test SetUserDuration on leaf tasks
5. Test AddTask with various parameter patterns
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


def traverse_hierarchy(task, depth=0, max_depth=4, results=None):
    """Recursively traverse ChildBars to find all tasks."""
    if results is None:
        results = []

    prefix = "  " * depth
    bar = win32com.client.Dispatch(task.Bar)
    bar_id = bar.ID
    task_id = task.ID

    try:
        child_bars = task.ChildBars
        child_count = child_bars.Count if child_bars else 0
    except Exception:
        child_count = 0

    links_in = task.LinksIn.Count
    links_out = task.LinksOut.Count

    try:
        dur = task.GetUserDuration().Hours
    except Exception:
        dur = "?"

    is_leaf = (child_count == 0)

    info = {
        'task': task,
        'bar': bar,
        'bar_id': bar_id,
        'task_id': task_id,
        'name': task.Name[:40],
        'depth': depth,
        'is_leaf': is_leaf,
        'links_in': links_in,
        'links_out': links_out,
        'duration': dur,
        'child_count': child_count,
    }
    results.append(info)

    if depth < max_depth and child_count > 0:
        for i in range(1, min(child_count + 1, 10)):  # limit to 10 children per level
            try:
                child_bar = win32com.client.Dispatch(child_bars.Item(i))
                child_task = win32com.client.Dispatch(child_bar.Tasks(1))
                traverse_hierarchy(child_task, depth + 1, max_depth, results)
            except Exception as e:
                results.append({
                    'depth': depth + 1,
                    'error': str(e)[:60],
                    'is_leaf': True,
                })

    return results


def test_deep_hierarchy(project):
    """Navigate deep into hierarchy to find leaf tasks."""
    print("\n" + "=" * 80)
    print("TEST 1: DEEP HIERARCHY TRAVERSAL")
    print("=" * 80)

    bar = project.Bars.Item(1)
    task = win32com.client.Dispatch(bar.Tasks(1))
    print(f"Root: Bar ID={bar.ID}, Task ID={task.ID}, Name={task.Name[:40]}")

    results = traverse_hierarchy(task, depth=0, max_depth=4)

    # Print summary
    print(f"\nTotal nodes found: {len(results)}")

    leaf_tasks = [r for r in results if r.get('is_leaf') and 'task' in r]
    linked_tasks = [r for r in results if r.get('links_in', 0) > 0 or r.get('links_out', 0) > 0]
    summary_tasks = [r for r in results if not r.get('is_leaf') and 'task' in r]

    print(f"Leaf tasks: {len(leaf_tasks)}")
    print(f"Summary tasks: {len(summary_tasks)}")
    print(f"Tasks with links: {len(linked_tasks)}")

    # Print first 20 leaf tasks
    print(f"\nFirst 20 leaf tasks:")
    for i, r in enumerate(leaf_tasks[:20]):
        prefix = "  " * r['depth']
        print(f"  {prefix}[{i}] BarID={r['bar_id']}, TaskID={r['task_id']}, "
              f"Name={r['name']}, LinksIn={r['links_in']}, LinksOut={r['links_out']}, "
              f"Dur={r['duration']}h")

    # Print first 10 linked tasks
    if linked_tasks:
        print(f"\nFirst 10 tasks WITH links:")
        for i, r in enumerate(linked_tasks[:10]):
            prefix = "  " * r['depth']
            print(f"  {prefix}[{i}] BarID={r['bar_id']}, TaskID={r['task_id']}, "
                  f"Name={r['name']}, LinksIn={r['links_in']}, LinksOut={r['links_out']}")

            # Dump link details
            t = r['task']
            if r['links_out'] > 0:
                for j in range(1, min(r['links_out'] + 1, 4)):
                    try:
                        link = win32com.client.Dispatch(t.LinksOut.Item(j))
                        end_task = win32com.client.Dispatch(link.EndTask)
                        print(f"      LinksOut[{j}]: type={link.type}, "
                              f"EndTask={end_task.Name[:30]}, ID={end_task.ID}")
                    except Exception as e:
                        print(f"      LinksOut[{j}]: {str(e)[:50]}")

            if r['links_in'] > 0:
                for j in range(1, min(r['links_in'] + 1, 4)):
                    try:
                        link = win32com.client.Dispatch(t.LinksIn.Item(j))
                        start_task = win32com.client.Dispatch(link.StartTask)
                        print(f"      LinksIn[{j}]: type={link.type}, "
                              f"StartTask={start_task.Name[:30]}, ID={start_task.ID}")
                    except Exception as e:
                        print(f"      LinksIn[{j}]: {str(e)[:50]}")

    return leaf_tasks, linked_tasks


def test_linkto_leaf_tasks(project, leaf_tasks):
    """Test LinkTo between two leaf-level tasks."""
    print("\n" + "=" * 80)
    print("TEST 2: LINKTO BETWEEN LEAF TASKS")
    print("=" * 80)

    if len(leaf_tasks) < 2:
        print("  Need 2+ leaf tasks")
        return

    # Find two unlinked leaf tasks to test with
    unlinked = [r for r in leaf_tasks if r['links_in'] == 0 and r['links_out'] == 0]
    if len(unlinked) >= 2:
        t1_info = unlinked[0]
        t2_info = unlinked[1]
    else:
        t1_info = leaf_tasks[0]
        t2_info = leaf_tasks[1]

    t1 = t1_info['task']
    t2 = t2_info['task']

    print(f"Task 1: ID={t1_info['task_id']}, Name={t1_info['name']}")
    print(f"  LinksIn={t1.LinksIn.Count}, LinksOut={t1.LinksOut.Count}")
    print(f"Task 2: ID={t2_info['task_id']}, Name={t2_info['name']}")
    print(f"  LinksIn={t2.LinksIn.Count}, LinksOut={t2.LinksOut.Count}")

    # Test LinkTo
    print(f"\n--- t1.LinkTo(t2) ---")
    project.StartTransaction("LinkTo leaf")
    try:
        link = t1.LinkTo(t2)
        print(f"  Result: {link}")
        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  SUCCESS! Link created!")
            print(f"    type = {ld.type}")
            try:
                print(f"    StartTask = {win32com.client.Dispatch(ld.StartTask).Name[:30]}")
                print(f"    EndTask = {win32com.client.Dispatch(ld.EndTask).Name[:30]}")
            except Exception:
                pass
            try:
                lag = ld.StartLagTime
                if lag:
                    print(f"    StartLagTime = {lag.Hours}h")
            except Exception:
                pass

            # Dump all link properties
            for attr in sorted([a for a in dir(ld) if not a.startswith('_')]):
                try:
                    val = getattr(ld, attr)
                    if not callable(val):
                        print(f"    {attr} = {str(val)[:60]}")
                except Exception:
                    pass

        project.EndTransaction()
        wait(project)

        # Verify
        print(f"\n  After LinkTo:")
        print(f"    t1.LinksOut: {t1.LinksOut.Count}")
        print(f"    t2.LinksIn: {t2.LinksIn.Count}")

        # Remove the test link
        print(f"\n  Removing test link...")
        project.StartTransaction("Remove test link")
        try:
            # Remove via LinksOut
            t1.LinksOut.Remove(t1.LinksOut.Count)  # Remove last link
            project.EndTransaction()
            wait(project)
            print(f"  Removed. t1.LinksOut: {t1.LinksOut.Count}, t2.LinksIn: {t2.LinksIn.Count}")
        except Exception as e:
            print(f"  Remove error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

    except Exception as e:
        print(f"  LinkTo ERROR: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def test_duration_leaf(project, leaf_tasks):
    """Test SetUserDuration on a leaf task."""
    print("\n" + "=" * 80)
    print("TEST 3: DURATION ON LEAF TASK")
    print("=" * 80)

    if not leaf_tasks:
        print("  No leaf tasks")
        return

    r = leaf_tasks[0]
    t = r['task']
    print(f"Task: ID={r['task_id']}, Name={r['name']}")
    print(f"  Start: {t.Start}")
    print(f"  End: {t.End}")
    try:
        print(f"  Duration: {t.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"  Constraint: {t.Constraint}")

    # Save original duration
    try:
        orig_dur = t.GetUserDuration().Hours
    except Exception:
        orig_dur = None

    # Test SetUserDuration
    print(f"\n--- SetUserDuration(5d) ---")
    project.StartTransaction("Set dur leaf")
    try:
        dur_obj = t.GetDurationFromString("5d")
        print(f"  GetDurationFromString('5d') => {dur_obj.Hours}h")
        t.SetUserDuration(dur_obj)
        print(f"  SetUserDuration => OK")
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

    try:
        new_dur = t.GetUserDuration().Hours
        print(f"  After reschedule: Duration={new_dur}h")
        if orig_dur and new_dur != orig_dur:
            print(f"  DURATION CHANGED! {orig_dur}h => {new_dur}h")
        elif orig_dur:
            print(f"  Duration UNCHANGED: {orig_dur}h")
    except Exception:
        pass
    print(f"  Start: {t.Start}")
    print(f"  End: {t.End}")

    # Restore original duration
    if orig_dur:
        print(f"\n--- Restoring original duration ({orig_dur}h) ---")
        project.StartTransaction("Restore dur")
        try:
            dur_str = f"{int(orig_dur / 8)}d" if orig_dur % 8 == 0 else f"{orig_dur}h"
            dur_obj = t.GetDurationFromString(dur_str)
            t.SetUserDuration(dur_obj)
            project.EndTransaction()
            wait(project)
            project.Reschedule(pywintypes.Time(datetime.now()))
            wait(project)
            print(f"  Restored: Duration={t.GetUserDuration().Hours}h")
        except Exception as e:
            print(f"  Restore error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass


def test_dates_leaf(project, leaf_tasks):
    """Test date setting methods on a leaf task."""
    print("\n" + "=" * 80)
    print("TEST 4: DATE SETTING ON LEAF TASK")
    print("=" * 80)

    if not leaf_tasks:
        print("  No leaf tasks")
        return

    r = leaf_tasks[0]
    t = r['task']
    print(f"Task: ID={r['task_id']}, Name={r['name']}")
    print(f"  Start: {t.Start}")
    print(f"  End: {t.End}")
    print(f"  Constraint: {t.Constraint}")

    # Save originals
    orig_start = t.Start
    orig_end = t.End
    orig_constraint = t.Constraint

    # Test SetUserStart
    print(f"\n--- Test A: SetUserStart ---")
    project.StartTransaction("SetUserStart")
    try:
        t.SetUserStart(pywintypes.Time(datetime(2026, 8, 3)))
        print(f"  SetUserStart(2026-08-03) => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        print(f"  After reschedule: Start={t.Start}, End={t.End}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test MoveToDate
    print(f"\n--- Test B: MoveToDate ---")
    project.StartTransaction("MoveToDate")
    try:
        t.MoveToDate(pywintypes.Time(datetime(2026, 9, 1)))
        print(f"  MoveToDate(2026-09-01) => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        print(f"  After reschedule: Start={t.Start}, End={t.End}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Test MoveStart
    print(f"\n--- Test C: MoveStart ---")
    project.StartTransaction("MoveStart")
    try:
        t.MoveStart(pywintypes.Time(datetime(2026, 10, 1)))
        print(f"  MoveStart(2026-10-01) => OK")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        print(f"  After reschedule: Start={t.Start}, End={t.End}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Restore original - use ImposedStart
    print(f"\n--- Restoring original dates ---")
    project.StartTransaction("Restore dates")
    try:
        t.RemoveConstraint()
        if orig_constraint > 0:
            t.ImposedStart = orig_start
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        print(f"  Restored: Start={t.Start}, End={t.End}, Constraint={t.Constraint}")
    except Exception as e:
        print(f"  Restore error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def test_addtask_patterns(project):
    """Test AddTask with various parameter patterns on ITaskBases."""
    print("\n" + "=" * 80)
    print("TEST 5: ADDTASK PARAMETER PATTERNS")
    print("=" * 80)

    bar = project.Bars.Item(1)

    # Get a child bar (summary level) to try AddTask on
    task = win32com.client.Dispatch(bar.Tasks(1))
    child_bars = task.ChildBars

    if child_bars.Count == 0:
        print("  No child bars")
        return []

    # Use first child bar
    child_bar = win32com.client.Dispatch(child_bars.Item(1))
    child_task = win32com.client.Dispatch(child_bar.Tasks(1))
    print(f"Parent bar: ID={child_bar.ID}, Name={child_bar.Name[:40]}")
    print(f"  Tasks.Count = {child_bar.Tasks.Count}")

    # Get the ITaskBases collection
    # In gen_py, bar.Tasks is an indexer. We need the underlying collection.
    # Try getting it via _oleobj_
    print(f"\n--- ITaskBases collection access ---")

    # Try direct method calls on the Tasks property
    tasks_col = None

    # Method 1: Dynamic dispatch on the bar
    try:
        bar_dyn = win32com.client.Dispatch(child_bar)
        # Try to get Tasks as a property returning an object
        print(f"  bar_dyn type: {type(bar_dyn)}")
        tc = bar_dyn.Tasks
        print(f"  bar_dyn.Tasks = {tc}, type={type(tc)}")
        if hasattr(tc, 'Count'):
            print(f"  Count = {tc.Count}")
        if hasattr(tc, 'AddTask'):
            print(f"  Has AddTask!")
            tasks_col = tc
    except Exception as e:
        print(f"  bar_dyn.Tasks error: {e}")

    # Method 2: Use _oleobj_ to invoke Tasks property ID
    if not tasks_col:
        try:
            import pythoncom as pc
            # Get dispatch ID for "Tasks"
            oleobj = child_bar._oleobj_
            disp_id = oleobj.GetIDsOfNames(0, "Tasks")
            print(f"  Tasks DISPID = {disp_id}")
            # Invoke as property get (DISPATCH_PROPERTYGET=2)
            result = oleobj.Invoke(disp_id, 0, 2, True)
            print(f"  Invoke result = {result}, type={type(result)}")
            if result:
                tc = win32com.client.Dispatch(result)
                print(f"  Dispatched: {tc}, type={type(tc)}")
                if hasattr(tc, 'Count'):
                    print(f"  Count = {tc.Count}")
                if hasattr(tc, 'AddTask'):
                    print(f"  Has AddTask!")
                    tasks_col = tc
                # List all methods
                attrs = sorted([a for a in dir(tc) if not a.startswith('_')])
                print(f"  Attrs: {attrs[:30]}")
        except Exception as e:
            print(f"  _oleobj_ error: {e}")

    # Method 3: Call Tasks without parameter via _oleobj_
    if not tasks_col:
        try:
            oleobj = child_bar._oleobj_
            disp_id = oleobj.GetIDsOfNames(0, "Tasks")
            # Try DISPATCH_METHOD (1)
            result = oleobj.Invoke(disp_id, 0, 1, True)
            print(f"  Invoke(METHOD) result = {result}")
            if result:
                tc = win32com.client.Dispatch(result)
                print(f"  Dispatched: {tc}")
                tasks_col = tc
        except Exception as e:
            print(f"  Invoke METHOD error: {e}")

    if not tasks_col:
        print("  Could not get ITaskBases collection")
        # Still try direct call patterns
        print(f"\n--- Direct AddTask patterns ---")
        created_ids = []

        project.StartTransaction("AddTask patterns")
        # Pattern A: child_bar.Tasks.AddTask()
        for method_name in ['AddTask', 'AddExpandedTask', 'AddMilestone',
                            'AddSummaryTask', 'AddHammockTask']:
            try:
                method = getattr(child_bar.Tasks, method_name)
                result = method()
                print(f"  child_bar.Tasks.{method_name}() => {result}")
                if result:
                    rd = win32com.client.Dispatch(result)
                    print(f"    Created: ID={rd.ID}, Name={rd.Name[:30]}")
                    created_ids.append(rd.ID)
            except Exception as e:
                print(f"  child_bar.Tasks.{method_name}() => {str(e)[:60]}")

        project.EndTransaction()
        wait(project)
        return created_ids

    # We have the collection!
    print(f"\n--- AddTask on ITaskBases collection ---")
    created_ids = []

    project.StartTransaction("AddTask col")
    for method_name in ['AddTask', 'AddExpandedTask', 'AddMilestone',
                        'AddSummaryTask', 'AddHammockTask']:
        try:
            method = getattr(tasks_col, method_name)
            result = method()
            print(f"  tasks_col.{method_name}() => {result}")
            if result:
                rd = win32com.client.Dispatch(result)
                rd.Name = f"TEST_{method_name}"
                print(f"    Created: ID={rd.ID}, Name={rd.Name}")
                created_ids.append(rd.ID)
        except Exception as e:
            print(f"  tasks_col.{method_name}() => {str(e)[:60]}")

    project.EndTransaction()
    wait(project)

    return created_ids


def test_bars_add_under_parent(project):
    """Test creating a bar under a parent bar using ChildBars.Add()."""
    print("\n" + "=" * 80)
    print("TEST 6: CREATE BAR UNDER PARENT")
    print("=" * 80)

    bar = project.Bars.Item(1)
    task = win32com.client.Dispatch(bar.Tasks(1))
    child_bars = task.ChildBars

    # Use first child (e.g. Milestones)
    parent_bar = win32com.client.Dispatch(child_bars.Item(1))
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))
    print(f"Parent: ID={parent_bar.ID}, Name={parent_bar.Name[:40]}")
    print(f"  ChildBars.Count = {parent_task.ChildBars.Count}")

    # Create a bar under this parent
    project.StartTransaction("ChildBars.Add")
    try:
        new_bar = parent_task.ChildBars.Add()
        print(f"  ChildBars.Add() => {new_bar}")
        if new_bar:
            nd = win32com.client.Dispatch(new_bar)
            nd.Name = "TEST_CHILD_BAR"
            bar_id = nd.ID
            print(f"  Created: ID={bar_id}, Name={nd.Name}")

            # Check its task
            new_task = win32com.client.Dispatch(nd.Tasks(1))
            print(f"  Task: ID={new_task.ID}")
            print(f"  HierarchyLevel: {new_task.HierarchyLevel}")
            print(f"  Parent: {new_task.Parentname}")

            # Check if it appears in parent's ChildBars
            project.EndTransaction()
            wait(project)

            print(f"  Parent ChildBars.Count now: {parent_task.ChildBars.Count}")

            # Test SetUserDuration on this new bar's task
            print(f"\n--- Duration on new child task ---")
            project.StartTransaction("Set dur child")
            try:
                dur_obj = new_task.GetDurationFromString("5d")
                new_task.SetUserDuration(dur_obj)
                print(f"  SetUserDuration(5d) => OK")
                project.EndTransaction()
                wait(project)
                project.Reschedule(pywintypes.Time(datetime.now()))
                wait(project)
                print(f"  Duration: {new_task.GetUserDuration().Hours}h")
                print(f"  Start: {new_task.Start}")
                print(f"  End: {new_task.End}")
            except Exception as e:
                print(f"  Duration error: {e}")
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass

            # Test ImposedStart
            print(f"\n--- ImposedStart on new child task ---")
            project.StartTransaction("ImposedStart child")
            try:
                new_task.ImposedStart = pywintypes.Time(datetime(2026, 6, 1))
                print(f"  ImposedStart = 2026-06-01 => OK")
                project.EndTransaction()
                wait(project)
                project.Reschedule(pywintypes.Time(datetime.now()))
                wait(project)
                print(f"  Start: {new_task.Start}")
                print(f"  End: {new_task.End}")
                print(f"  Constraint: {new_task.Constraint}")
            except Exception as e:
                print(f"  ImposedStart error: {e}")
                try:
                    project.AbandonTransaction()
                except Exception:
                    pass

            return bar_id
        else:
            project.EndTransaction()
            wait(project)
    except Exception as e:
        print(f"  ChildBars.Add() error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    return None


def test_linkto_new_bars(project, bar_id_1):
    """Create a second bar and test LinkTo between them."""
    print("\n" + "=" * 80)
    print("TEST 7: LINKTO BETWEEN NEW CHILD BARS")
    print("=" * 80)

    if not bar_id_1:
        print("  No bar_id_1")
        return None

    # Find bar_id_1
    bar1 = None
    root_bar = project.Bars.Item(1)
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    parent_bar = win32com.client.Dispatch(root_task.ChildBars.Item(1))
    parent_task = win32com.client.Dispatch(parent_bar.Tasks(1))

    for i in range(1, parent_task.ChildBars.Count + 1):
        b = win32com.client.Dispatch(parent_task.ChildBars.Item(i))
        if b.ID == bar_id_1:
            bar1 = b
            break

    if not bar1:
        print(f"  Can't find bar {bar_id_1}")
        return None

    task1 = win32com.client.Dispatch(bar1.Tasks(1))
    print(f"Bar 1: ID={bar1.ID}, Name={bar1.Name[:30]}, TaskID={task1.ID}")

    # Create second bar under same parent
    project.StartTransaction("Create bar2")
    try:
        new_bar = parent_task.ChildBars.Add()
        nd = win32com.client.Dispatch(new_bar)
        nd.Name = "TEST_CHILD_BAR_2"
        bar_id_2 = nd.ID
        task2 = win32com.client.Dispatch(nd.Tasks(1))
        print(f"Bar 2: ID={bar_id_2}, Name={nd.Name}, TaskID={task2.ID}")

        # Set duration
        dur_obj = task2.GetDurationFromString("3d")
        task2.SetUserDuration(dur_obj)

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Create bar2 error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return None

    # NOW test LinkTo
    print(f"\n--- task1.LinkTo(task2) ---")
    project.StartTransaction("LinkTo children")
    try:
        link = task1.LinkTo(task2)
        print(f"  Result: {link}")
        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  SUCCESS! LINK CREATED!")
            print(f"    type = {ld.type}")
            try:
                st = win32com.client.Dispatch(ld.StartTask)
                et = win32com.client.Dispatch(ld.EndTask)
                print(f"    StartTask = {st.Name[:30]} (ID={st.ID})")
                print(f"    EndTask = {et.Name[:30]} (ID={et.ID})")
            except Exception:
                pass

            # Dump all link properties
            for attr in sorted([a for a in dir(ld) if not a.startswith('_')]):
                try:
                    val = getattr(ld, attr)
                    if not callable(val):
                        print(f"    {attr} = {str(val)[:60]}")
                except Exception:
                    pass

        project.EndTransaction()
        wait(project)

        # Reschedule
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        # Verify
        print(f"\n  After LinkTo + Reschedule:")
        print(f"    task1.LinksOut: {task1.LinksOut.Count}")
        print(f"    task2.LinksIn: {task2.LinksIn.Count}")
        print(f"    task1 Start={task1.Start}, End={task1.End}")
        print(f"    task2 Start={task2.Start}, End={task2.End}")

    except Exception as e:
        print(f"  LinkTo error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    return bar_id_2


def cleanup_bars(project, bar_ids, parent_task):
    """Remove created bars."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup: {len(bar_ids)} bars ---")
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
    print("COM Explorer v13 — Deep Hierarchy + Leaf LinkTo")
    print("=" * 80)
    created_bar_ids = []
    parent_task_for_cleanup = None
    try:
        app, project = connect()

        # Test 1: Deep hierarchy
        leaf_tasks, linked_tasks = test_deep_hierarchy(project)

        # Test 2: LinkTo between existing leaf tasks
        if len(leaf_tasks) >= 2:
            test_linkto_leaf_tasks(project, leaf_tasks)

        # Test 3: Duration on leaf task
        if leaf_tasks:
            test_duration_leaf(project, leaf_tasks)

        # Test 4: Date setting on leaf task
        if leaf_tasks:
            test_dates_leaf(project, leaf_tasks)

        # Test 5: AddTask patterns
        addtask_ids = test_addtask_patterns(project)

        # Test 6: Create bar under parent
        bar_id_1 = test_bars_add_under_parent(project)
        if bar_id_1:
            created_bar_ids.append(bar_id_1)
            # Get parent task for cleanup
            root_bar = project.Bars.Item(1)
            root_task = win32com.client.Dispatch(root_bar.Tasks(1))
            p_bar = win32com.client.Dispatch(root_task.ChildBars.Item(1))
            parent_task_for_cleanup = win32com.client.Dispatch(p_bar.Tasks(1))

        # Test 7: LinkTo between new child bars
        bar_id_2 = test_linkto_new_bars(project, bar_id_1)
        if bar_id_2:
            created_bar_ids.append(bar_id_2)

        # Cleanup
        if created_bar_ids and parent_task_for_cleanup:
            cleanup_bars(project, created_bar_ids, parent_task_for_cleanup)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
        if created_bar_ids and parent_task_for_cleanup:
            try:
                cleanup_bars(project, created_bar_ids, parent_task_for_cleanup)
            except Exception:
                pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
