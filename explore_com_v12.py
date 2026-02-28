"""
COM Explorer v12 — Use ITaskBases (bar.Tasks()) for proper task creation
KEY INSIGHT: bars.Add() creates IBars (chart-level), but tasks live in ITaskBases.
Each IBar has bar.Tasks() -> ITaskBases which has AddTask/AddSummaryTask etc.

1. Access bar.Tasks() and navigate ITaskBases
2. Create tasks via ITaskBases.AddTask()
3. Set dates/duration on ITaskBase
4. Link tasks via ITaskBase.LinkTo()
5. Test ILink properties (type, lag)
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


def test_itaskbases(project):
    """Explore bar.Tasks() and ITaskBases interface."""
    print("\n" + "=" * 80)
    print("TEST 1: bar.Tasks() — ITaskBases")
    print("=" * 80)

    bars = project.Bars
    bar = bars.Item(1)
    print(f"Bar: ID={bar.ID}, Name={bar.Name[:40]}")

    # Get Tasks from bar - try different access patterns
    print(f"\nAccessing Tasks:")
    bar_dyn = win32com.client.Dispatch(bar)

    # Try as property
    for access in ['Tasks', 'TaskBases', 'AllTasks']:
        try:
            val = getattr(bar_dyn, access)
            print(f"  bar.{access}: {val} (type={type(val)})")
            if hasattr(val, 'Count'):
                print(f"    Count: {val.Count}")
            elif callable(val):
                # Try calling with different args
                for arg in [None, 1, 0]:
                    try:
                        if arg is None:
                            result = val()
                        else:
                            result = val(arg)
                        print(f"    {access}({arg}) => {result}")
                        if hasattr(result, 'Count'):
                            print(f"      Count: {result.Count}")
                        break
                    except Exception as e2:
                        print(f"    {access}({arg}) => {str(e2)[:40]}")
        except AttributeError:
            pass
        except Exception as e:
            print(f"  bar.{access} => {str(e)[:50]}")

    # Try bar._oleobj_ direct dispatch
    try:
        # Tasks is PROP_GET which returns ptr - should be a property
        tasks_val = bar_dyn.Tasks
        print(f"\n  bar_dyn.Tasks (property) = {tasks_val}")
        print(f"  type = {type(tasks_val)}")
    except Exception as e:
        print(f"  bar_dyn.Tasks property error: {str(e)[:60]}")

    # Use etask instead - IExpandedTask also has access to the hierarchy
    et = bar.ExpandedTask
    et_dyn = win32com.client.Dispatch(et)
    print(f"\nExpanded task methods:")
    for method_name in ['Tasks', 'TaskBases', 'ChildTasks', 'Children',
                        'SubTasks', 'Members', 'AllTasks', 'ChildBars']:
        try:
            val = getattr(et_dyn, method_name)
            print(f"  et.{method_name}: {val} (type={type(val)})")
            if hasattr(val, 'Count'):
                print(f"    Count: {val.Count}")
        except AttributeError:
            pass
        except Exception as e:
            print(f"  et.{method_name} => {str(e)[:50]}")

    # Access the ITaskBases from the FULL type library
    # bar.Tasks is a collection that maps to ITaskBases
    # Try getting it as a default property
    try:
        # Count should be available if Tasks returns ITaskBases
        tasks_count = bar.Tasks.Count
        print(f"\n  bar.Tasks.Count = {tasks_count}")
    except Exception as e:
        print(f"\n  bar.Tasks.Count error: {str(e)[:50]}")

    # Try: The IBar has Tasks as ptr — maybe we access via oleobj
    print(f"\n  Trying _oleobj_ access:")
    try:
        # Get the DISPID for Tasks
        dispid = bar._oleobj_.GetIDsOfNames('Tasks', 0)
        print(f"  DISPID for Tasks: {dispid}")
        # Invoke as PROP_GET (flag=2)
        result = bar._oleobj_.Invoke(dispid[0], 0, 2, 1)  # DISPATCH_PROPERTYGET
        print(f"  Invoke result: {result}")
        print(f"  type: {type(result)}")
        if result:
            tasks_dyn = win32com.client.Dispatch(result)
            print(f"  Dispatch type: {type(tasks_dyn)}")
            print(f"  Count: {tasks_dyn.Count}")
            attrs = sorted([a for a in dir(tasks_dyn) if not a.startswith('_')])
            print(f"  Attributes: {attrs}")
    except Exception as e:
        print(f"  _oleobj_ error: {str(e)[:60]}")
        traceback.print_exc()

    tasks_dyn = None
    # Last try: use gen_py generated accessor differently
    try:
        # The gen_py makes Tasks a callable that requires an index
        # This means Tasks(1) returns the first task in the collection
        t1 = bar.Tasks(1)
        print(f"\n  bar.Tasks(1) => {t1}")
        t1_dyn = win32com.client.Dispatch(t1)
        print(f"    ID: {t1_dyn.ID}, Name: {t1_dyn.Name[:30]}")

        # Count items by iterating
        count = 0
        for i in range(1, 10000):
            try:
                t = bar.Tasks(i)
                count = i
            except Exception:
                break
        print(f"  Total Tasks: {count}")

        if count > 0:
            # Print first 10
            for i in range(1, min(count + 1, 11)):
                t = win32com.client.Dispatch(bar.Tasks(i))
                print(f"    [{i}] ID={t.ID}, Name={t.Name[:30]}")
    except Exception as e:
        print(f"\n  bar.Tasks(1) error: {str(e)[:50]}")

    return tasks_dyn

    # List attributes
    attrs = sorted([a for a in dir(tasks_dyn) if not a.startswith('_')])
    print(f"Attributes: {attrs}")

    # Dump all tasks
    print(f"\nFirst 10 tasks:")
    for i in range(1, min(tasks_dyn.Count + 1, 11)):
        try:
            task = tasks_dyn.Item(i)
            task_dyn = win32com.client.Dispatch(task)
            li = task_dyn.LinksIn.Count if hasattr(task_dyn, 'LinksIn') else '?'
            lo = task_dyn.LinksOut.Count if hasattr(task_dyn, 'LinksOut') else '?'
            name = task_dyn.Name[:30] if hasattr(task_dyn, 'Name') else '?'
            tid = task_dyn.ID
            try:
                dur = task_dyn.GetUserDuration()
                dur_h = dur.Hours if dur else 0
            except Exception:
                dur_h = '?'
            print(f"  [{i}] ID={tid}, Name={name}, "
                  f"Duration={dur_h}h, LinksIn={li}, LinksOut={lo}")
        except Exception as e:
            print(f"  [{i}] Error: {str(e)[:50]}")

    # Check if we can create tasks
    print(f"\nITaskBases methods:")
    for method_name in ['Add', 'AddTask', 'AddSummaryTask', 'AddExpandedTask',
                        'AddHammockTask', 'AddMilestone', 'Remove', 'All', 'Count', 'Item']:
        try:
            fn = getattr(tasks_dyn, method_name)
            if callable(fn):
                print(f"  {method_name}() [callable]")
            else:
                print(f"  {method_name} = {fn}")
        except AttributeError:
            pass
        except Exception as e:
            print(f"  {method_name} => {str(e)[:40]}")

    return tasks_dyn


def test_create_task(project):
    """Create a task via ITaskBases.AddTask()."""
    print("\n" + "=" * 80)
    print("TEST 2: CREATE TASK VIA AddTask()")
    print("=" * 80)

    bars = project.Bars
    bar = bars.Item(1)
    tasks = win32com.client.Dispatch(bar.Tasks())
    initial_count = tasks.Count
    print(f"Using bar: ID={bar.ID}, Name={bar.Name[:40]}")
    print(f"Tasks.Count before: {initial_count}")

    # Try AddTask
    project.StartTransaction("AddTask")
    created_ids = []
    try:
        # AddTask might take no args, or take a name, or take a position
        try:
            new_task = tasks.AddTask()
            print(f"  AddTask() => {new_task}")
            nt_dyn = win32com.client.Dispatch(new_task)
            print(f"    type: {type(nt_dyn)}")
            print(f"    ID: {nt_dyn.ID}")
            nt_dyn.Name = "TASK_A"
            print(f"    Name set to: TASK_A")
            created_ids.append(nt_dyn.ID)
        except Exception as e:
            print(f"  AddTask() error: {e}")
            # Try Add() as fallback
            try:
                new_task = tasks.Add()
                print(f"  Add() => {new_task}")
                nt_dyn = win32com.client.Dispatch(new_task)
                nt_dyn.Name = "TASK_A"
                created_ids.append(nt_dyn.ID)
            except Exception as e2:
                print(f"  Add() error: {e2}")

        # Create second task
        try:
            new_task2 = tasks.AddTask()
            nt2_dyn = win32com.client.Dispatch(new_task2)
            nt2_dyn.Name = "TASK_B"
            created_ids.append(nt2_dyn.ID)
            print(f"  Created TASK_B: ID={nt2_dyn.ID}")
        except Exception as e:
            print(f"  Second AddTask error: {e}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Verify
    tasks2 = win32com.client.Dispatch(bar.Tasks())
    print(f"\nTasks.Count after: {tasks2.Count} (was {initial_count})")

    return created_ids


def test_task_dates_duration(project, task_ids):
    """Set dates and duration on ITaskBase tasks."""
    print("\n" + "=" * 80)
    print("TEST 3: SET DATES AND DURATION ON TASKBASE")
    print("=" * 80)

    if not task_ids or len(task_ids) < 1:
        print("  No task IDs")
        return

    bars = project.Bars
    bar = bars.Item(1)
    tasks = win32com.client.Dispatch(bar.Tasks())

    # Find our tasks
    task_a = task_b = None
    for i in range(1, tasks.Count + 1):
        try:
            t = win32com.client.Dispatch(tasks.Item(i))
            if t.ID == task_ids[0]:
                task_a = t
            elif len(task_ids) > 1 and t.ID == task_ids[1]:
                task_b = t
        except Exception:
            pass

    if not task_a:
        print(f"  Can't find task {task_ids[0]}")
        return

    print(f"  Task A: ID={task_a.ID}, Name={task_a.Name}")
    print(f"    Start: {task_a.Start}")
    print(f"    End: {task_a.End}")
    try:
        print(f"    Duration: {task_a.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"    Constraint: {task_a.Constraint}")

    # Set duration
    print(f"\n--- Set Duration ---")
    project.StartTransaction("Set duration")
    try:
        dur_obj = task_a.GetDurationFromString("10d")
        print(f"  GetDurationFromString('10d') => {dur_obj.Hours}h")
        task_a.SetUserDuration(dur_obj)
        print(f"  SetUserDuration(10d) => OK")
    except Exception as e:
        print(f"  Duration error: {e}")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"  After duration + reschedule:")
    try:
        print(f"    Duration: {task_a.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"    Start: {task_a.Start}")
    print(f"    End: {task_a.End}")

    # Set start date
    print(f"\n--- Set Start Date ---")
    project.StartTransaction("Set start")
    try:
        task_a.MoveToDate(pywintypes.Time(datetime(2026, 6, 1)))
        print(f"  MoveToDate(2026-06-01) => OK")
    except Exception as e:
        print(f"  MoveToDate error: {e}")
        # Try ImposedStart (it's on IExpandedTask but ITaskBase inherits from it)
        try:
            task_a.ImposedStart = pywintypes.Time(datetime(2026, 6, 1))
            print(f"  ImposedStart = 2026-06-01 => OK")
        except Exception as e2:
            print(f"  ImposedStart error: {e2}")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"  After start date + reschedule:")
    try:
        print(f"    Duration: {task_a.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"    Start: {task_a.Start}")
    print(f"    End: {task_a.End}")

    # Set duration on task B too
    if task_b:
        project.StartTransaction("Set B duration")
        try:
            dur_obj = task_b.GetDurationFromString("5d")
            task_b.SetUserDuration(dur_obj)
            task_b.MoveToDate(pywintypes.Time(datetime(2026, 6, 15)))
            print(f"  Task B: 5d from 2026-06-15 => OK")
        except Exception as e:
            print(f"  Task B error: {e}")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)


def test_link_tasks(project, task_ids):
    """Link ITaskBase tasks via LinkTo."""
    print("\n" + "=" * 80)
    print("TEST 4: LINK TASKS VIA LinkTo")
    print("=" * 80)

    if not task_ids or len(task_ids) < 2:
        print("  Need at least 2 task IDs")
        return

    bars = project.Bars
    bar = bars.Item(1)
    tasks = win32com.client.Dispatch(bar.Tasks())

    # Find our tasks
    task_a = task_b = None
    for i in range(1, tasks.Count + 1):
        try:
            t = win32com.client.Dispatch(tasks.Item(i))
            if t.ID == task_ids[0]:
                task_a = t
            elif t.ID == task_ids[1]:
                task_b = t
        except Exception:
            pass

    if not (task_a and task_b):
        print(f"  Can't find tasks")
        return

    print(f"  Task A: ID={task_a.ID}, LinksOut={task_a.LinksOut.Count}")
    print(f"  Task B: ID={task_b.ID}, LinksIn={task_b.LinksIn.Count}")

    # LinkTo: A -> B
    project.StartTransaction("Link A->B")
    try:
        link = task_a.LinkTo(task_b)
        print(f"\n  task_a.LinkTo(task_b) => {link}")
        if link:
            link_dyn = win32com.client.Dispatch(link)
            print(f"  SUCCESS! Link created!")
            print(f"    Link ID: {link_dyn.ID}")
            print(f"    Link type: {link_dyn.type}")
            print(f"    StartTask: {link_dyn.StartTask}")
            print(f"    EndTask: {link_dyn.EndTask}")

            # Try to read link properties
            for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
                try:
                    val = getattr(link_dyn, attr)
                    if not callable(val):
                        vs = str(val)[:50]
                        print(f"    {attr} = {vs}")
                except Exception:
                    pass

            # Try to set link type (FS=0, SS=1, FF=2, SF=3?)
            try:
                current_type = link_dyn.type
                print(f"\n    Current link type: {current_type}")
            except Exception as e:
                print(f"    Link type error: {e}")

            # Try to set lag
            try:
                lag = link_dyn.StartLagTime
                print(f"    StartLagTime: {lag}")
                if lag:
                    print(f"      Hours: {lag.Hours}")
            except Exception as e:
                print(f"    StartLagTime error: {e}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"\n  LinkTo error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Verify
    print(f"\n  After LinkTo:")
    print(f"    Task A LinksOut: {task_a.LinksOut.Count}")
    print(f"    Task B LinksIn: {task_b.LinksIn.Count}")
    print(f"    Task A HasSuccessor: {task_a.HasSuccessor}")
    print(f"    Task B HasPredecessor: {task_b.HasPredecessor}")

    # Reschedule
    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"\n  After reschedule:")
    print(f"    Task A: Start={task_a.Start}, End={task_a.End}")
    print(f"    Task B: Start={task_b.Start}, End={task_b.End}")


def test_link_set_type_lag(project, task_ids):
    """Test modifying link type and lag."""
    print("\n" + "=" * 80)
    print("TEST 5: SET LINK TYPE AND LAG")
    print("=" * 80)

    if not task_ids or len(task_ids) < 2:
        return

    bars = project.Bars
    bar = bars.Item(1)
    tasks = win32com.client.Dispatch(bar.Tasks())

    task_a = None
    for i in range(1, tasks.Count + 1):
        try:
            t = win32com.client.Dispatch(tasks.Item(i))
            if t.ID == task_ids[0]:
                task_a = t
                break
        except Exception:
            pass

    if not task_a or task_a.LinksOut.Count == 0:
        print("  No link to modify")
        return

    link = win32com.client.Dispatch(task_a.LinksOut.Item(1))
    print(f"  Link ID: {link.ID}")
    print(f"  Current type: {link.type}")

    # Try setting link type
    project.StartTransaction("Set link type")
    for link_type in range(4):
        try:
            link.type = link_type
            readback = link.type
            print(f"  type = {link_type} => readback: {readback}")
        except Exception as e:
            print(f"  type = {link_type} => {str(e)[:50]}")

    # Try setting lag
    try:
        # Create a duration object for lag
        dur_obj = task_a.GetDurationFromString("2d")
        link.StartLagTime = dur_obj
        print(f"  StartLagTime = 2d => OK")
        readback = link.StartLagTime
        if readback:
            print(f"    StartLagTime readback: {readback.Hours}h")
    except Exception as e:
        print(f"  StartLagTime error: {e}")

    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"\n  After type+lag set + reschedule:")
    print(f"    Link type: {link.type}")
    try:
        print(f"    StartLagTime: {link.StartLagTime.Hours}h")
    except Exception:
        pass


def cleanup_tasks(project, task_ids):
    """Remove created tasks."""
    if not task_ids:
        return
    print(f"\n--- Cleanup: {len(task_ids)} tasks ---")
    bars = project.Bars
    bar = bars.Item(1)
    tasks = win32com.client.Dispatch(bar.Tasks())

    project.StartTransaction("Cleanup tasks")
    for tid in reversed(task_ids):
        for i in range(tasks.Count, 0, -1):
            try:
                t = win32com.client.Dispatch(tasks.Item(i))
                if t.ID == tid:
                    tasks.Remove(i)
                    print(f"  Removed task ID={tid}")
                    break
            except Exception:
                pass
    project.EndTransaction()
    wait(project)
    print(f"  Tasks.Count now: {win32com.client.Dispatch(bar.Tasks()).Count}")


if __name__ == "__main__":
    print("COM Explorer v12 — ITaskBases: The REAL Task Interface")
    print("=" * 80)
    task_ids = []
    try:
        app, project = connect()

        # Clean stale top-level bars
        bars = project.Bars
        stale = []
        for i in range(1, bars.Count + 1):
            b = bars.Item(i)
            if b.Name.startswith(("WF_TEST", "LF_", "DUR_", "DATE_", "TOKEN_",
                                  "LINK_", "TYPE_", "TEST_", "CHILD_", "CONV_",
                                  "WORKFLOW", "DIRECT_", "NEW_CHILD")):
                stale.append(b.ID)
        if stale:
            print(f"\nPre-cleanup: {len(stale)} stale bars")
            project.StartTransaction("Cleanup")
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

        # Test 1: Explore ITaskBases
        test_itaskbases(project)

        # Test 2: Create tasks
        task_ids = test_create_task(project)

        if task_ids:
            # Test 3: Set dates/duration
            test_task_dates_duration(project, task_ids)

            # Test 4: Link tasks
            if len(task_ids) >= 2:
                test_link_tasks(project, task_ids)

                # Test 5: Modify link type/lag
                test_link_set_type_lag(project, task_ids)

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
