"""
COM Explorer v12b — Navigate ITaskBase hierarchy
1. Get ITaskBase from bar.Tasks(1)
2. Find child tasks/task bases
3. Navigate hierarchy
4. Test AddTask on ITaskBases
5. Link and set duration on proper TaskBase objects
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


def test_taskbase_hierarchy(project):
    """Navigate the ITaskBase hierarchy."""
    print("\n" + "=" * 80)
    print("TEST 1: ITaskBase HIERARCHY NAVIGATION")
    print("=" * 80)

    bars = project.Bars
    bar = bars.Item(1)
    print(f"Bar: ID={bar.ID}, Name={bar.Name[:40]}")

    # Get the task from the bar
    task = win32com.client.Dispatch(bar.Tasks(1))
    print(f"Task: ID={task.ID}, Name={task.Name[:40]}")
    print(f"  type: {type(task)}")

    # Check all attributes of the task
    print(f"\nTask attributes (non-underscore):")
    attrs = sorted([a for a in dir(task) if not a.startswith('_')])
    child_related = [a for a in attrs if any(kw in a.lower() for kw in
                     ['child', 'sub', 'member', 'task', 'bar', 'next', 'prev',
                      'parent', 'summary', 'hierarchy'])]
    print(f"  Child/hierarchy-related: {child_related}")

    # Try each access pattern
    for method_name in child_related:
        try:
            val = getattr(task, method_name)
            if callable(val):
                try:
                    result = val()
                    print(f"  {method_name}() => {result}")
                    if hasattr(result, 'Count'):
                        print(f"    Count: {result.Count}")
                    elif hasattr(result, 'ID'):
                        print(f"    ID: {result.ID}, Name: {result.Name[:30]}")
                except Exception as e:
                    # Try with param
                    try:
                        result = val(1)
                        print(f"  {method_name}(1) => {result}")
                    except Exception:
                        print(f"  {method_name}() => {str(e)[:50]}")
            else:
                vs = str(val)[:50]
                print(f"  {method_name} = {vs}")
                if hasattr(val, 'Count'):
                    print(f"    Count: {val.Count}")
        except Exception as e:
            print(f"  {method_name} => {str(e)[:40]}")

    # Try to get Bar property from the TaskBase
    print(f"\n  Task.Bar:")
    try:
        tb = task.Bar
        print(f"    Bar ID: {tb.ID}")
    except Exception as e:
        print(f"    Error: {e}")

    # Try to get ChildBars from the TaskBase
    print(f"\n  Trying ChildBars on TaskBase:")
    try:
        cb = task.ChildBars
        print(f"    ChildBars type: {type(cb)}")
        if hasattr(cb, 'Count'):
            print(f"    Count: {cb.Count}")
            for i in range(1, min(cb.Count + 1, 6)):
                try:
                    child_bar = cb.Item(i)
                    child_bar_dyn = win32com.client.Dispatch(child_bar)
                    print(f"    [{i}] Bar ID={child_bar_dyn.ID}, Name={child_bar_dyn.Name[:30]}")
                    # Get task from this child bar
                    try:
                        child_task = win32com.client.Dispatch(child_bar_dyn.Tasks(1))
                        print(f"        Task ID={child_task.ID}, Name={child_task.Name[:30]}")
                        print(f"        LinksIn={child_task.LinksIn.Count}, LinksOut={child_task.LinksOut.Count}")
                        try:
                            dur = child_task.GetUserDuration()
                            print(f"        Duration={dur.Hours}h")
                        except Exception:
                            pass
                    except Exception as e3:
                        print(f"        Tasks(1) error: {str(e3)[:40]}")
                except Exception as e2:
                    print(f"    [{i}] Error: {str(e2)[:40]}")
    except Exception as e:
        print(f"    Error: {e}")

    # Try NextTask from TaskBase
    print(f"\n  Trying NextTask from TaskBase:")
    current = task
    for i in range(10):
        try:
            next_t = current.NextTask()
            if next_t is None:
                print(f"    step {i}: None")
                break
            nd = win32com.client.Dispatch(next_t)
            print(f"    step {i}: ID={nd.ID}, Name={nd.Name[:30]}, "
                  f"LinksIn={nd.LinksIn.Count}, LinksOut={nd.LinksOut.Count}")
            current = nd
        except Exception as e:
            print(f"    step {i}: {str(e)[:40]}")
            break


def test_add_task_to_summary(project):
    """Create tasks via ITaskBases.AddTask on a summary bar."""
    print("\n" + "=" * 80)
    print("TEST 2: ADD TASK VIA ITaskBases")
    print("=" * 80)

    bars = project.Bars
    bar = bars.Item(1)

    # Get the tasks collection via _oleobj_
    # bar.Tasks is treated as a default indexer in gen_py
    # We need the collection object itself
    print(f"bar.Tasks.Count = {bar.Tasks.Count}")

    # Try to get ITaskBases as a collection
    # Approach 1: Use dynamic dispatch to get the Tasks collection
    print(f"\nTrying to get ITaskBases collection:")

    # The bar.Tasks property in the gen_py returns a callable
    # bar.Tasks.AddTask() should work if we can chain it
    try:
        tasks_col = bar.Tasks
        print(f"  bar.Tasks = {tasks_col}")
        print(f"  type = {type(tasks_col)}")
    except Exception as e:
        print(f"  bar.Tasks error: {e}")

    # Try AddTask through the collection
    print(f"\nTrying AddTask methods:")

    # Method 1: bar.Tasks.AddTask()
    project.StartTransaction("AddTask test")
    created_ids = []
    try:
        try:
            result = bar.Tasks.AddTask()
            print(f"  bar.Tasks.AddTask() => {result}")
            rd = win32com.client.Dispatch(result)
            rd.Name = "NEW_TASK_A"
            created_ids.append(rd.ID)
            print(f"    Created: ID={rd.ID}")
        except Exception as e:
            print(f"  bar.Tasks.AddTask() => {str(e)[:60]}")

        # Method 2: bar.Tasks.AddExpandedTask()
        try:
            result = bar.Tasks.AddExpandedTask()
            print(f"  bar.Tasks.AddExpandedTask() => {result}")
            rd = win32com.client.Dispatch(result)
            rd.Name = "NEW_TASK_B"
            created_ids.append(rd.ID)
            print(f"    Created: ID={rd.ID}")
        except Exception as e:
            print(f"  bar.Tasks.AddExpandedTask() => {str(e)[:60]}")

        # Method 3: bar.Tasks.Add()
        try:
            result = bar.Tasks.Add()
            print(f"  bar.Tasks.Add() => {result}")
            rd = win32com.client.Dispatch(result)
            rd.Name = "NEW_TASK_C"
            created_ids.append(rd.ID)
            print(f"    Created: ID={rd.ID}")
        except Exception as e:
            print(f"  bar.Tasks.Add() => {str(e)[:60]}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Verify
    if created_ids:
        print(f"\nCreated task IDs: {created_ids}")
        print(f"bar.Tasks.Count now: {bar.Tasks.Count}")

        # Check if we can find them
        for tid in created_ids:
            try:
                found = False
                for i in range(1, bar.Tasks.Count + 1):
                    t = win32com.client.Dispatch(bar.Tasks(i))
                    if t.ID == tid:
                        print(f"  Found task {tid}: Name={t.Name}, "
                              f"Start={t.Start}, Duration=?")
                        found = True
                        break
                if not found:
                    print(f"  Task {tid} NOT found in bar.Tasks")
            except Exception as e:
                print(f"  Error finding {tid}: {e}")

    return created_ids


def test_duration_on_taskbase(project, task_ids):
    """Set duration on ITaskBase objects."""
    print("\n" + "=" * 80)
    print("TEST 3: DURATION ON ITASKBASE")
    print("=" * 80)

    if not task_ids:
        # Use existing task from bar
        bar = project.Bars.Item(1)
        task = win32com.client.Dispatch(bar.Tasks(1))
        # Try to find a child task
        try:
            cb = task.ChildBars
            if cb and cb.Count > 0:
                child_bar = win32com.client.Dispatch(cb.Item(1))
                child_task = win32com.client.Dispatch(child_bar.Tasks(1))
                print(f"Using child task: ID={child_task.ID}")
                task_ids = [child_task.ID]
        except Exception:
            pass

    if not task_ids:
        print("  No task IDs available")
        return

    bar = project.Bars.Item(1)

    # Find the task
    task = None
    for i in range(1, bar.Tasks.Count + 1):
        t = win32com.client.Dispatch(bar.Tasks(i))
        if t.ID == task_ids[0]:
            task = t
            break

    if not task:
        print(f"  Can't find task {task_ids[0]}")
        return

    print(f"Task: ID={task.ID}, Name={task.Name}")
    print(f"  Start: {task.Start}")
    print(f"  End: {task.End}")
    try:
        print(f"  Duration: {task.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"  Constraint: {task.Constraint}")

    # Set duration
    print(f"\n--- SetUserDuration ---")
    project.StartTransaction("Set dur")
    try:
        dur_obj = task.GetDurationFromString("10d")
        print(f"  GetDurationFromString('10d') => {dur_obj.Hours}h")
        task.SetUserDuration(dur_obj)
        print(f"  SetUserDuration => OK")
    except Exception as e:
        print(f"  Error: {e}")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    try:
        print(f"  After reschedule: Duration={task.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"  Start: {task.Start}")
    print(f"  End: {task.End}")


def test_link_taskbases(project, task_ids):
    """Link ITaskBase objects."""
    print("\n" + "=" * 80)
    print("TEST 4: LINK ITASKBASE OBJECTS")
    print("=" * 80)

    if not task_ids or len(task_ids) < 2:
        print("  Need 2+ task IDs")
        return

    bar = project.Bars.Item(1)
    t1 = t2 = None
    for i in range(1, bar.Tasks.Count + 1):
        t = win32com.client.Dispatch(bar.Tasks(i))
        if t.ID == task_ids[0]:
            t1 = t
        elif t.ID == task_ids[1]:
            t2 = t

    if not (t1 and t2):
        print(f"  Can't find tasks")
        return

    print(f"Task 1: ID={t1.ID}, Name={t1.Name}")
    print(f"Task 2: ID={t2.ID}, Name={t2.Name}")

    # LinkTo
    project.StartTransaction("LinkTo")
    try:
        link = t1.LinkTo(t2)
        print(f"\n  t1.LinkTo(t2) => {link}")
        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  SUCCESS! Link ID={ld.ID}")
            print(f"    type: {ld.type}")
            try:
                print(f"    StartTask: {win32com.client.Dispatch(ld.StartTask).Name[:30]}")
                print(f"    EndTask: {win32com.client.Dispatch(ld.EndTask).Name[:30]}")
            except Exception:
                pass
            try:
                lag = ld.StartLagTime
                if lag:
                    print(f"    StartLagTime: {lag.Hours}h")
            except Exception:
                pass

            # Dump all link properties
            for attr in sorted([a for a in dir(ld) if not a.startswith('_')]):
                try:
                    val = getattr(ld, attr)
                    if not callable(val):
                        print(f"    {attr} = {str(val)[:50]}")
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

    # Verify
    print(f"\n  After LinkTo:")
    print(f"    T1 LinksOut: {t1.LinksOut.Count}")
    print(f"    T2 LinksIn: {t2.LinksIn.Count}")

    # Reschedule
    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)


def cleanup_tasks(project, task_ids):
    """Remove created tasks."""
    if not task_ids:
        return
    print(f"\n--- Cleanup: {len(task_ids)} tasks ---")
    bar = project.Bars.Item(1)
    project.StartTransaction("Cleanup")
    for tid in reversed(task_ids):
        for i in range(bar.Tasks.Count, 0, -1):
            try:
                t = win32com.client.Dispatch(bar.Tasks(i))
                if t.ID == tid:
                    bar.Tasks.Remove(i)
                    print(f"  Removed task ID={tid}")
                    break
            except Exception:
                pass
    project.EndTransaction()
    wait(project)


if __name__ == "__main__":
    print("COM Explorer v12b — ITaskBase Hierarchy + AddTask + Link")
    print("=" * 80)
    task_ids = []
    try:
        app, project = connect()

        # Test 1: Explore hierarchy
        test_taskbase_hierarchy(project)

        # Test 2: Create tasks
        task_ids = test_add_task_to_summary(project)

        if task_ids:
            # Test 3: Duration
            test_duration_on_taskbase(project, task_ids)

            # Test 4: Link
            if len(task_ids) >= 2:
                test_link_taskbases(project, task_ids)

            # Cleanup
            cleanup_tasks(project, task_ids)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
        if task_ids:
            try:
                cleanup_tasks(project, task_ids)
            except Exception:
                pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
