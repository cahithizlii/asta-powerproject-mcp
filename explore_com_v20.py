"""
COM Explorer v20 — Task creation strategies
PROBLEM: ChildBars.Add() creates bars WITHOUT Tasks(1)
STRATEGY: Find working task creation method

TESTS:
  1. Dump ITaskBases interface (AddTask params)
  2. project.Bars.Add() — check if it has Tasks(1)
  3. Bars.Add() then ExpandedTask operations
  4. Try AddTask/AddExpandedTask with various params
  5. Explore ITaskBases._oleobj_ for method signatures
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
    bar = project.Bars.Item(1)
    return win32com.client.Dispatch(bar.Tasks(1))


def test_bars_add(project):
    """Test project.Bars.Add() — does it have Tasks(1)?"""
    print("\n" + "=" * 80)
    print("TEST 1: project.Bars.Add()")
    print("=" * 80)

    project.StartTransaction("Bars.Add")
    try:
        new_bar = win32com.client.Dispatch(project.Bars.Add())
        new_bar.Name = "V20_TOP_LEVEL"
        bar_id = new_bar.ID
        print(f"  Created: BarID={bar_id}, Name={new_bar.Name}")

        # In-txn checks
        try:
            t = win32com.client.Dispatch(new_bar.Tasks(1))
            print(f"  In-txn Tasks(1): ID={t.ID}, type={type(t).__name__}")
        except Exception as e:
            print(f"  In-txn Tasks(1): FAILED - {str(e)[:60]}")

        try:
            et = new_bar.ExpandedTask
            print(f"  In-txn ExpandedTask: ID={et.ID}, type={type(et).__name__}")
        except Exception as e:
            print(f"  In-txn ExpandedTask: FAILED - {str(e)[:60]}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return None

    # After commit
    print(f"\n--- After commit ---")
    # Find via project.Bars
    bars = project.Bars
    print(f"  Bars.Count: {bars.Count}")
    for i in range(1, bars.Count + 1):
        b = win32com.client.Dispatch(bars.Item(i))
        if b.ID == bar_id:
            print(f"  Found at Bars[{i}]: ID={b.ID}")
            try:
                t = win32com.client.Dispatch(b.Tasks(1))
                print(f"    Tasks(1): ID={t.ID}, type={type(t).__name__}")
                print(f"    Start={t.Start}, End={t.End}")
                try:
                    print(f"    Duration={t.GetUserDuration().Hours}h")
                except Exception:
                    pass
                # Check ImposedStart
                if hasattr(t, 'ImposedStart'):
                    print(f"    HAS ImposedStart")
                else:
                    print(f"    NO ImposedStart")
                # Check LinkTo
                if hasattr(t, 'LinkTo'):
                    print(f"    HAS LinkTo")
                else:
                    print(f"    NO LinkTo")
            except Exception as e:
                print(f"    Tasks(1) FAILED: {str(e)[:60]}")
            try:
                et = b.ExpandedTask
                print(f"    ExpandedTask: ID={et.ID}, type={type(et).__name__}")
                if hasattr(et, 'ImposedStart'):
                    print(f"    ET HAS ImposedStart")
                if hasattr(et, 'LinkTo'):
                    print(f"    ET HAS LinkTo")
            except Exception as e:
                print(f"    ExpandedTask FAILED: {str(e)[:60]}")
            break

    return bar_id


def test_itaskbases_interface(project):
    """Explore ITaskBases interface for AddTask parameters."""
    print("\n" + "=" * 80)
    print("TEST 2: ITaskBases INTERFACE EXPLORATION")
    print("=" * 80)

    root_bar = project.Bars.Item(1)

    # Get Tasks collection
    print(f"--- root_bar.Tasks ---")
    tasks = root_bar.Tasks
    print(f"  type: {type(tasks).__name__}")
    print(f"  Count: {tasks.Count}")

    # List all attributes
    attrs = sorted([a for a in dir(tasks) if not a.startswith('_')])
    print(f"  Attributes ({len(attrs)}): {attrs}")

    # Try to get type info
    try:
        oleobj = tasks._oleobj_
        type_info = oleobj.GetTypeInfo()
        type_attr = type_info.GetTypeAttr()
        print(f"\n  TypeInfo: {type_attr}")
        print(f"  cFuncs: {type_attr.cFuncs}")
        print(f"  cVars: {type_attr.cVars}")

        for i in range(type_attr.cFuncs):
            try:
                func_desc = type_info.GetFuncDesc(i)
                names = type_info.GetNames(func_desc.memid)
                func_name = names[0] if names else f"func_{i}"
                param_names = names[1:] if len(names) > 1 else []

                # Get parameter types
                param_types = []
                for j, elem_desc in enumerate(func_desc.args):
                    pname = param_names[j] if j < len(param_names) else f"p{j}"
                    try:
                        vt = elem_desc[0]
                        if isinstance(vt, tuple):
                            vt_val = vt[0]
                        elif isinstance(vt, int):
                            vt_val = vt
                        else:
                            vt_val = str(vt)

                        vt_map = {0: 'void', 2: 'int16', 3: 'int32', 4: 'float',
                                  5: 'double', 6: 'currency', 7: 'date', 8: 'string',
                                  9: 'IDispatch', 11: 'bool', 12: 'variant',
                                  13: 'IUnknown', 16: 'int8', 17: 'uint8',
                                  18: 'uint16', 19: 'uint32', 22: 'int', 23: 'uint',
                                  26: 'ptr', 29: 'carray'}
                        vt_str = vt_map.get(vt_val, f"vt{vt_val}")
                        param_types.append(f"{vt_str} {pname}")
                    except Exception:
                        param_types.append(f"? {pname}")

                ret_type = func_desc.rettype
                try:
                    if isinstance(ret_type[0], tuple):
                        ret_vt = ret_type[0][0]
                    elif isinstance(ret_type[0], int):
                        ret_vt = ret_type[0]
                    else:
                        ret_vt = str(ret_type[0])
                    vt_map = {0: 'void', 2: 'int16', 3: 'int32', 8: 'string',
                              9: 'IDispatch', 11: 'bool', 12: 'variant', 24: 'void'}
                    ret_str = vt_map.get(ret_vt, f"vt{ret_vt}")
                except Exception:
                    ret_str = "?"

                invoke_kind = {1: 'FUNC', 2: 'GET', 4: 'PUT', 8: 'PUTREF'}
                kind = invoke_kind.get(func_desc.invkind, f"k{func_desc.invkind}")

                print(f"  [{i}] {kind} {ret_str} {func_name}({', '.join(param_types)})")
            except Exception as e:
                print(f"  [{i}] Error: {str(e)[:50]}")
    except Exception as e:
        print(f"  TypeInfo error: {e}")

    # Try AddTask with various params
    print(f"\n--- AddTask attempts ---")

    # First try getting a child bar's Tasks
    root_task = get_root_task(project)
    insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
    insaat_task = win32com.client.Dispatch(insaat.Tasks(1))

    # Try on insaat's Tasks
    print(f"  insaat.Tasks.Count: {insaat.Tasks.Count}")

    # Try via _oleobj_ directly
    try:
        oleobj = insaat.Tasks._oleobj_
        disp_id = oleobj.GetIDsOfNames(0, "AddTask")
        print(f"  AddTask DISPID: {disp_id}")

        # Try invoke with no params
        try:
            result = oleobj.Invoke(disp_id, 0, 1, True)
            print(f"  AddTask() => {result}")
        except Exception as e:
            print(f"  AddTask(): {str(e)[:80]}")

        # Try with empty string
        try:
            result = oleobj.Invoke(disp_id, 0, 1, True, "")
            print(f"  AddTask(''): {result}")
        except Exception as e:
            print(f"  AddTask(''): {str(e)[:80]}")

    except Exception as e:
        print(f"  _oleobj_ error: {e}")


def test_childbars_add_then_addtask(project):
    """Try: ChildBars.Add() creates bar, then add a task to it."""
    print("\n" + "=" * 80)
    print("TEST 3: ChildBars.Add() + AddTask on new bar")
    print("=" * 80)

    root_task = get_root_task(project)
    insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
    insaat_task = win32com.client.Dispatch(insaat.Tasks(1))

    project.StartTransaction("Add + Task")
    try:
        new_bar = win32com.client.Dispatch(insaat_task.ChildBars.Add())
        new_bar.Name = "V20_ADD_TASK_TEST"
        bar_id = new_bar.ID
        print(f"  Created bar: ID={bar_id}")

        # Try to add a task to this bar
        try:
            tasks_col = new_bar.Tasks
            print(f"  Tasks type: {type(tasks_col).__name__}")
            print(f"  Tasks.Count: {tasks_col.Count}")

            # Try AddTask
            try:
                result = tasks_col.AddTask()
                print(f"  AddTask(): {result}")
            except Exception as e:
                print(f"  AddTask(): {str(e)[:60]}")

            # Try AddExpandedTask
            try:
                result = tasks_col.AddExpandedTask()
                print(f"  AddExpandedTask(): {result}")
            except Exception as e:
                print(f"  AddExpandedTask(): {str(e)[:60]}")

            # Try Add
            try:
                result = tasks_col.Add()
                print(f"  Add(): {result}")
            except Exception as e:
                print(f"  Add(): {str(e)[:60]}")
        except Exception as e:
            print(f"  Tasks access error: {e}")

        # Also check ExpandedTask
        try:
            et = new_bar.ExpandedTask
            print(f"\n  ExpandedTask: ID={et.ID}, Name={et.Name[:30]}")
            # Is it the same as insaat's task?
            print(f"  Same as insaat_task? {et.ID == insaat_task.ID}")

            # What about SetUserDuration directly on ExpandedTask?
            try:
                dur = et.GetDurationFromString("5d")
                et.SetUserDuration(dur)
                print(f"  SetUserDuration(5d) on ET => OK")
            except Exception as e:
                print(f"  SetUserDuration on ET: {str(e)[:60]}")
        except Exception as e:
            print(f"  ExpandedTask error: {e}")

        project.EndTransaction()
        wait(project)
        return bar_id
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return None


def test_top_bar_operations(project, bar_id):
    """Test operations on top-level bar (from Bars.Add())."""
    print("\n" + "=" * 80)
    print("TEST 4: OPERATIONS ON TOP-LEVEL BAR")
    print("=" * 80)

    if not bar_id:
        print("  No bar to test")
        return

    bars = project.Bars
    bar = None
    for i in range(1, bars.Count + 1):
        b = win32com.client.Dispatch(bars.Item(i))
        if b.ID == bar_id:
            bar = b
            break

    if not bar:
        print(f"  Bar {bar_id} not in project.Bars")
        return

    et = bar.ExpandedTask
    print(f"  Bar: ID={bar.ID}, Name={bar.Name}")
    print(f"  ExpandedTask: ID={et.ID}")
    print(f"  Start={et.Start}, End={et.End}")
    print(f"  Constraint={et.Constraint}")

    # Set duration
    print(f"\n--- SetUserDuration(10d) ---")
    project.StartTransaction("Dur top")
    try:
        bars = project.Bars
        for i in range(1, bars.Count + 1):
            b = win32com.client.Dispatch(bars.Item(i))
            if b.ID == bar_id:
                et = b.ExpandedTask
                dur = et.GetDurationFromString("10d")
                et.SetUserDuration(dur)
                print(f"  OK (ET.ID={et.ID})")
                break
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

    bars = project.Bars
    for i in range(1, bars.Count + 1):
        b = win32com.client.Dispatch(bars.Item(i))
        if b.ID == bar_id:
            et = b.ExpandedTask
            print(f"  After: Dur={et.GetUserDuration().Hours}h, Start={et.Start}, End={et.End}")
            break

    # Set ImposedStart
    print(f"\n--- ImposedStart(2026-07-01) ---")
    project.StartTransaction("Start top")
    try:
        bars = project.Bars
        for i in range(1, bars.Count + 1):
            b = win32com.client.Dispatch(bars.Item(i))
            if b.ID == bar_id:
                et = b.ExpandedTask
                et.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
                print(f"  OK")
                break
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

    bars = project.Bars
    for i in range(1, bars.Count + 1):
        b = win32com.client.Dispatch(bars.Item(i))
        if b.ID == bar_id:
            et = b.ExpandedTask
            print(f"  After: Start={et.Start}, End={et.End}, Constraint={et.Constraint}")
            break

    # ChangeParentBar
    print(f"\n--- ChangeParentBar (move under Insaat) ---")
    project.StartTransaction("Move")
    try:
        bars = project.Bars
        for i in range(1, bars.Count + 1):
            b = win32com.client.Dispatch(bars.Item(i))
            if b.ID == bar_id:
                et = b.ExpandedTask
                root_task = get_root_task(project)
                insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
                et.ChangeParentBar(insaat, False)
                print(f"  ChangeParentBar => OK!")
                break
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  ChangeParentBar error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def cleanup_top(project, bar_id):
    """Remove top-level bar."""
    if not bar_id:
        return
    print(f"\n--- Cleanup top-level bar {bar_id} ---")
    project.StartTransaction("Del top")
    try:
        bars = project.Bars
        for i in range(bars.Count, 0, -1):
            b = win32com.client.Dispatch(bars.Item(i))
            if b.ID == bar_id:
                bars.Remove(i)
                print(f"  Removed bar ID={bar_id}")
                break
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def cleanup_child(project, bar_id, parent_index=3):
    """Remove child bar under parent."""
    if not bar_id:
        return
    print(f"\n--- Cleanup child bar {bar_id} ---")
    try:
        root_task = get_root_task(project)
        parent = win32com.client.Dispatch(root_task.ChildBars.Item(parent_index))
        parent_task = win32com.client.Dispatch(parent.Tasks(1))
        project.StartTransaction(f"Del {bar_id}")
        cb = parent_task.ChildBars
        for i in range(cb.Count, 0, -1):
            b = win32com.client.Dispatch(cb.Item(i))
            if b.ID == bar_id:
                cb.Remove(i)
                print(f"  Removed bar ID={bar_id}")
                break
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


if __name__ == "__main__":
    print("COM Explorer v20 — Task Creation Strategies")
    print("=" * 80)
    top_bar_id = None
    child_bar_id = None
    try:
        app, project = connect()

        # Test 1: Bars.Add() top-level
        top_bar_id = test_bars_add(project)

        # Test 2: ITaskBases interface dump
        test_itaskbases_interface(project)

        # Test 3: ChildBars.Add() + AddTask
        child_bar_id = test_childbars_add_then_addtask(project)

        # Test 4: Operations on top-level bar
        test_top_bar_operations(project, top_bar_id)

        # Cleanup
        cleanup_top(project, top_bar_id)
        cleanup_child(project, child_bar_id, 3)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
        if top_bar_id:
            cleanup_top(project, top_bar_id)
        if child_bar_id:
            cleanup_child(project, child_bar_id, 3)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
