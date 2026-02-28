"""
COM Explorer v9 — Complete typelib dump + link creation + date persistence
1. Fixed param parsing to reveal ALL hidden [ERROR] methods
2. Dump ILinks/ILink interface from LinksIn/LinksOut
3. Test ALL link creation approaches
4. Test date/duration persistence (verify values after reschedule)
5. Test AddConstraint with all parameter combos
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


def _vt_to_str(vt_code):
    """Convert a VARIANT type code to human-readable string."""
    VT_NAMES = {
        0: "void", 1: "null", 2: "int16", 3: "int32", 4: "float",
        5: "double", 6: "currency", 7: "date", 8: "string", 9: "IDispatch",
        10: "error", 11: "bool", 12: "variant", 13: "IUnknown", 14: "decimal",
        16: "int8", 17: "uint8", 18: "uint16", 19: "uint32", 20: "int64",
        21: "uint64", 22: "int", 23: "uint", 24: "void", 25: "HRESULT",
        26: "ptr", 27: "safearray", 28: "carray", 29: "userdefined",
        30: "lpstr", 31: "lpwstr",
    }
    # Handle various input types
    if vt_code is None:
        return "void"
    if isinstance(vt_code, tuple):
        vt_code = vt_code[0] if len(vt_code) > 0 else 0
    if not isinstance(vt_code, int):
        return f"unknown({vt_code})"
    base = vt_code & 0xFFF
    is_array = bool(vt_code & 0x2000)
    is_byref = bool(vt_code & 0x4000)
    name = VT_NAMES.get(base, f"VT_{base}")
    if is_array:
        name = f"{name}[]"
    if is_byref:
        name = f"{name}&"
    return name


def _safe_param_type(param):
    """Safely extract parameter type from a func_desc parameter entry."""
    if param is None:
        return "variant"
    try:
        p0 = param[0]
        if p0 is None:
            return "variant"
        if isinstance(p0, int):
            return _vt_to_str(p0)
        if isinstance(p0, tuple) and len(p0) > 0:
            return _vt_to_str(p0[0])
        return _vt_to_str(p0)
    except Exception:
        return f"?({param})"


def dump_interface_full(obj, label="Object"):
    """Dump ALL methods/properties of a COM object, handling all param types."""
    print(f"\n{'='*80}")
    print(f"INTERFACE DUMP: {label}")
    print(f"{'='*80}")

    try:
        type_info = obj._oleobj_.GetTypeInfo()
        type_attr = type_info.GetTypeAttr()
        type_name = type_info.GetDocumentation(-1)[0]
        func_count = type_attr[6]
        print(f"Interface: {type_name}, Functions: {func_count}")
        print()

        for i in range(func_count):
            try:
                fd = type_info.GetFuncDesc(i)
                fn = type_info.GetNames(fd[0])
                func_name = fn[0] if fn else f"func_{i}"

                invoke_kind = fd[4]
                kind_str = {1: "METHOD", 2: "PROP_GET", 4: "PROP_PUT",
                           8: "PROP_PUTREF"}.get(invoke_kind, f"KIND_{invoke_kind}")

                # Return type - handle safely
                try:
                    ret_type = fd[8]
                    if ret_type is None:
                        type_str = "void"
                    elif isinstance(ret_type, tuple) and len(ret_type) > 0:
                        type_str = _vt_to_str(ret_type[0])
                    elif isinstance(ret_type, int):
                        type_str = _vt_to_str(ret_type)
                    else:
                        type_str = f"ret?({ret_type})"
                except Exception:
                    type_str = "void"

                # Parameters - handle safely
                param_strs = []
                try:
                    if fd[2]:
                        for pi, param in enumerate(fd[2]):
                            p_name = fn[pi + 1] if (pi + 1) < len(fn) else f"p{pi}"
                            p_type = _safe_param_type(param)
                            param_strs.append(f"{p_type} {p_name}")
                except Exception as pe:
                    param_strs.append(f"<param_error: {pe}>")

                params_s = ", ".join(param_strs)
                print(f"  [{kind_str}] {type_str} {func_name}({params_s})")

            except Exception as e:
                # Even if GetFuncDesc fails, try to get the name
                try:
                    fd2 = type_info.GetFuncDesc(i)
                    fn2 = type_info.GetNames(fd2[0])
                    print(f"  [ERROR] func_{i} name={fn2[0] if fn2 else '?'}: {e}")
                except Exception:
                    print(f"  [ERROR] func_{i}: {e}")

    except Exception as e:
        print(f"  Cannot get type info: {e}")

    # Also list dir() attributes for comparison
    print(f"\n  --- dir() attributes (non-underscore) ---")
    try:
        dyn = win32com.client.Dispatch(obj)
        attrs = sorted([a for a in dir(dyn) if not a.startswith('_')])
        print(f"  Total: {len(attrs)}")
        for attr in attrs:
            try:
                val = getattr(dyn, attr)
                if callable(val):
                    print(f"    {attr}() [callable]")
                else:
                    vs = str(val)[:50]
                    print(f"    {attr} = {vs}")
            except Exception as e:
                print(f"    {attr} => ERR: {str(e)[:40]}")
    except Exception:
        pass


def test_links_interface(project):
    """Dump the ILinks interface from LinksIn/LinksOut and test Add()."""
    print("\n" + "#" * 80)
    print("# LINKS INTERFACE ANALYSIS")
    print("#" * 80)

    bars = project.Bars
    bar = bars.Item(1)
    etask = bar.ExpandedTask
    print(f"Using bar: ID={bar.ID}, Name={bar.Name}")

    # Dump LinksIn interface
    links_in = etask.LinksIn
    print(f"\nLinksIn type: {type(links_in)}")
    print(f"LinksIn Count: {links_in.Count}")
    dump_interface_full(links_in, "LinksIn (ILinks)")

    # Dump LinksOut interface
    links_out = etask.LinksOut
    print(f"\nLinksOut type: {type(links_out)}")
    print(f"LinksOut Count: {links_out.Count}")

    # If there's at least one link, dump the ILink interface
    if links_in.Count > 0:
        link = links_in.Item(1)
        print(f"\nLink Item(1) type: {type(link)}")
        dump_interface_full(link, "Single Link (ILink)")
    elif links_out.Count > 0:
        link = links_out.Item(1)
        print(f"\nLink Item(1) from LinksOut type: {type(link)}")
        dump_interface_full(link, "Single Link (ILink)")
    else:
        # Try the second bar
        if bars.Count >= 2:
            bar2 = bars.Item(2)
            et2 = bar2.ExpandedTask
            if et2.LinksIn.Count > 0:
                link = et2.LinksIn.Item(1)
                dump_interface_full(link, "Single Link from bar2 (ILink)")
            elif et2.LinksOut.Count > 0:
                link = et2.LinksOut.Item(1)
                dump_interface_full(link, "Single Link from bar2 (ILink)")

    # Try dynamic dispatch on LinksIn
    print("\n--- Dynamic dispatch Add() attempts on LinksIn ---")
    links_dyn = win32com.client.Dispatch(links_in)
    print(f"  LinksIn dynamic dir: {sorted([a for a in dir(links_dyn) if not a.startswith('_')])}")

    # Try Add with various signatures
    for method_name in ['Add', 'AddLink', 'Create', 'CreateLink', 'Insert', 'New']:
        try:
            fn = getattr(links_dyn, method_name)
            print(f"  LinksIn.{method_name} EXISTS: {fn}")
        except AttributeError:
            pass

    # Try LinksOut too
    links_out_dyn = win32com.client.Dispatch(links_out)
    print(f"  LinksOut dynamic dir: {sorted([a for a in dir(links_out_dyn) if not a.startswith('_')])}")

    for method_name in ['Add', 'AddLink', 'Create', 'CreateLink', 'Insert', 'New']:
        try:
            fn = getattr(links_out_dyn, method_name)
            print(f"  LinksOut.{method_name} EXISTS: {fn}")
        except AttributeError:
            pass

    # Dump LinkCategorys
    print("\n--- LinkCategorys ---")
    try:
        lcs = project.LinkCategorys
        print(f"  Count: {lcs.Count}")
        for i in range(1, lcs.Count + 1):
            lc = lcs.Item(i)
            print(f"  [{i}] ID={lc.ID}, Name={lc.Name}")
        if lcs.Count > 0:
            dump_interface_full(lcs.Item(1), "Single LinkCategory")
    except Exception as e:
        print(f"  Error: {e}")

    # AllLinkCategorys
    print("\n--- AllLinkCategorys ---")
    try:
        alcs = project.AllLinkCategorys
        print(f"  Count: {alcs.Count}")
    except Exception as e:
        print(f"  Error: {e}")


def test_expanded_task_full(project):
    """Dump ALL IExpandedTask methods (fixed parser)."""
    print("\n" + "#" * 80)
    print("# IEXPANDEDTASK FULL DUMP (FIXED)")
    print("#" * 80)

    bar = project.Bars.Item(1)
    etask = bar.ExpandedTask
    dump_interface_full(etask, f"IExpandedTask (bar {bar.ID}: {bar.Name})")


def test_link_creation(project):
    """Test ALL possible link creation approaches."""
    print("\n" + "#" * 80)
    print("# LINK CREATION TESTS")
    print("#" * 80)

    bars = project.Bars

    # Get existing bar IDs
    view = project.CurrentView
    bcv = win32com.client.Dispatch(view)
    all_ids = bcv.AllBarIds()
    print(f"All bar IDs: {all_ids}")

    if len(all_ids) < 2:
        print("Need at least 2 bars for link testing")
        return

    # Find two bars
    bar1_id = all_ids[0]
    bar2_id = all_ids[1]

    bar1 = None
    bar2 = None
    for i in range(1, bars.Count + 1):
        b = bars.Item(i)
        if b.ID == bar1_id:
            bar1 = b
        elif b.ID == bar2_id:
            bar2 = b

    if not bar1 or not bar2:
        print(f"Could not find bars {bar1_id}, {bar2_id}")
        return

    etask1 = bar1.ExpandedTask
    etask2 = bar2.ExpandedTask

    print(f"Bar1: ID={bar1.ID}, Name={bar1.Name}")
    print(f"Bar2: ID={bar2.ID}, Name={bar2.Name}")
    print(f"Bar1 LinksOut: {etask1.LinksOut.Count}")
    print(f"Bar2 LinksIn: {etask2.LinksIn.Count}")

    # Strategy 1: LinksOut.Add() with no params
    print("\n--- Strategy 1: LinksOut.Add() ---")
    project.StartTransaction("Link test 1")
    try:
        lo = etask1.LinksOut
        lo_dyn = win32com.client.Dispatch(lo)
        try:
            new_link = lo_dyn.Add()
            print(f"  Add() => {new_link}")
        except Exception as e:
            print(f"  Add() => {str(e)[:80]}")
    except Exception as e:
        print(f"  Error: {e}")
    project.AbandonTransaction()

    # Strategy 2: LinksIn.Add() with no params
    print("\n--- Strategy 2: LinksIn.Add() ---")
    project.StartTransaction("Link test 2")
    try:
        li = etask2.LinksIn
        li_dyn = win32com.client.Dispatch(li)
        try:
            new_link = li_dyn.Add()
            print(f"  Add() => {new_link}")
        except Exception as e:
            print(f"  Add() => {str(e)[:80]}")
    except Exception as e:
        print(f"  Error: {e}")
    project.AbandonTransaction()

    # Strategy 3: LinksOut.Add(task2)
    print("\n--- Strategy 3: LinksOut.Add(etask2) ---")
    project.StartTransaction("Link test 3")
    try:
        lo = etask1.LinksOut
        lo_dyn = win32com.client.Dispatch(lo)
        try:
            new_link = lo_dyn.Add(etask2)
            print(f"  Add(etask2) => {new_link}")
        except Exception as e:
            print(f"  Add(etask2) => {str(e)[:80]}")
        try:
            new_link = lo_dyn.Add(bar2)
            print(f"  Add(bar2) => {new_link}")
        except Exception as e:
            print(f"  Add(bar2) => {str(e)[:80]}")
        try:
            new_link = lo_dyn.Add(bar2_id)
            print(f"  Add(bar2_id={bar2_id}) => {new_link}")
        except Exception as e:
            print(f"  Add(bar2_id={bar2_id}) => {str(e)[:80]}")
    except Exception as e:
        print(f"  Error: {e}")
    project.AbandonTransaction()

    # Strategy 4: Try IExpandedTask methods for link creation
    print("\n--- Strategy 4: IExpandedTask link methods ---")
    et1_dyn = win32com.client.Dispatch(etask1)
    for method_name in ['AddLink', 'CreateLink', 'AddLinkTo', 'LinkTo',
                        'AddSuccessor', 'AddPredecessor', 'ConnectTo',
                        'MakeLink', 'AddDependency', 'AddLinkOut', 'AddLinkIn',
                        'CreateLinkTo', 'CreateLinkFrom']:
        try:
            fn = getattr(et1_dyn, method_name)
            print(f"  etask.{method_name} EXISTS: {fn}")
            if callable(fn):
                project.StartTransaction(f"Link test {method_name}")
                try:
                    # Try with task
                    result = fn(etask2)
                    print(f"    {method_name}(etask2) => {result}")
                    project.EndTransaction()
                except Exception as e1:
                    try:
                        result = fn(bar2)
                        print(f"    {method_name}(bar2) => {result}")
                        project.EndTransaction()
                    except Exception as e2:
                        try:
                            result = fn(bar2_id)
                            print(f"    {method_name}(bar2_id) => {result}")
                            project.EndTransaction()
                        except Exception as e3:
                            print(f"    {method_name}() all args fail: {str(e1)[:50]}")
                            project.AbandonTransaction()
        except AttributeError:
            pass

    # Strategy 5: IBar methods for links
    print("\n--- Strategy 5: IBar link methods ---")
    bar1_dyn = win32com.client.Dispatch(bar1)
    for method_name in ['AddLink', 'CreateLink', 'LinkTo', 'AddSuccessor',
                        'AddPredecessor', 'AddDependency', 'MakeLink',
                        'AddLinkOut', 'AddLinkIn']:
        try:
            fn = getattr(bar1_dyn, method_name)
            print(f"  bar.{method_name} EXISTS: {fn}")
        except AttributeError:
            pass

    # Strategy 6: IProject methods for links
    print("\n--- Strategy 6: IProject link methods ---")
    proj_dyn = win32com.client.Dispatch(project)
    for method_name in ['AddLink', 'CreateLink', 'MakeLink', 'LinkBars',
                        'LinkTasks', 'AddDependency', 'CreateDependency']:
        try:
            fn = getattr(proj_dyn, method_name)
            print(f"  project.{method_name} EXISTS: {fn}")
        except AttributeError:
            pass

    # Strategy 7: IBarChartView methods for links
    print("\n--- Strategy 7: IBarChartView link methods ---")
    for method_name in ['AddLink', 'CreateLink', 'LinkBars', 'LinkTasks',
                        'MakeLink', 'AddDependency', 'SelectBarsAndLink',
                        'CreateLinkBetween']:
        try:
            fn = getattr(bcv, method_name)
            print(f"  bcv.{method_name} EXISTS: {fn}")
        except AttributeError:
            pass

    # Strategy 8: Application CreateRelativeTimeObject — maybe there's CreateLink?
    print("\n--- Strategy 8: IApplication methods ---")
    app = project.Application
    app_dyn = win32com.client.Dispatch(app)
    for method_name in ['CreateLink', 'AddLink', 'MakeLink', 'CreateDependency']:
        try:
            fn = getattr(app_dyn, method_name)
            print(f"  app.{method_name} EXISTS: {fn}")
        except AttributeError:
            pass

    # Strategy 9: Use EditToken/EditTokenV to set link properties
    print("\n--- Strategy 9: EditToken link-related tokens ---")
    for token in ['Predecessor', 'Successor', 'LinkIn', 'LinkOut',
                  'Predecessors', 'Successors', 'Link', 'Links',
                  'Dependency', 'Dependencies']:
        try:
            result = etask1.GetToken(token)
            print(f"  GetToken('{token}') = {result}")
        except Exception as e:
            print(f"  GetToken('{token}') => {str(e)[:50]}")

    # Strategy 10: Try type library for ALL interface methods containing 'Link'
    print("\n--- Strategy 10: Type library link search ---")
    try:
        type_info = etask1._oleobj_.GetTypeInfo()
        type_attr = type_info.GetTypeAttr()
        func_count = type_attr[6]
        for i in range(func_count):
            try:
                fd = type_info.GetFuncDesc(i)
                fn = type_info.GetNames(fd[0])
                func_name = fn[0] if fn else ""
                if any(kw in func_name.lower() for kw in ['link', 'pred', 'succ', 'dep', 'connect', 'attach']):
                    invoke_kind = fd[4]
                    kind_str = {1: "METHOD", 2: "GET", 4: "PUT", 8: "PUTREF"}.get(invoke_kind, str(invoke_kind))
                    print(f"  [{kind_str}] {func_name} (func_{i})")
            except Exception:
                pass
    except Exception:
        pass


def test_date_duration_persistence(project):
    """Test if dates and duration actually persist after creation + reschedule."""
    print("\n" + "#" * 80)
    print("# DATE/DURATION PERSISTENCE TEST")
    print("#" * 80)

    bars = project.Bars
    initial_count = bars.Count
    print(f"Initial bars count: {initial_count}")

    # Create a test bar
    project.StartTransaction("Date persistence test")
    new_bar = bars.Add()
    new_bar.Name = "DATE_PERSIST_TEST"
    bar_id = new_bar.ID
    etask = new_bar.ExpandedTask
    print(f"Created bar: ID={bar_id}")

    # Read defaults BEFORE setting anything
    print(f"\n--- DEFAULTS (before any settings) ---")
    print(f"  Start: {etask.Start}")
    print(f"  End: {etask.End}")
    print(f"  GetUserStart: {etask.GetUserStart()}")
    print(f"  GetUserEnd: {etask.GetUserEnd()}")
    try:
        print(f"  GetUserDuration: {etask.GetUserDuration()}")
        dur = etask.GetUserDuration()
        if dur:
            print(f"    Hours: {dur.Hours}")
    except Exception as e:
        print(f"  GetUserDuration: {e}")
    print(f"  Constraint: {etask.Constraint}")
    print(f"  ImposedStart: {etask.ImposedStart}")
    print(f"  ImposedEnd: {etask.ImposedEnd}")

    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    # Test A: Set duration first, then dates
    print(f"\n--- TEST A: Set Duration first, then Start ---")
    project.StartTransaction("Set duration")
    try:
        dur_obj = etask.GetDurationFromString("10d")
        print(f"  GetDurationFromString('10d'): Hours={dur_obj.Hours}")
        etask.SetUserDuration(dur_obj)
        print(f"  SetUserDuration(10d) => OK")
    except Exception as e:
        print(f"  SetUserDuration error: {e}")
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    print(f"  After SetUserDuration:")
    print(f"    GetUserDuration: {etask.GetUserDuration().Hours}h")
    print(f"    Start: {etask.Start}")
    print(f"    End: {etask.End}")

    # Now set start date via MoveToDate
    project.StartTransaction("Set start date")
    try:
        target = pywintypes.Time(datetime(2026, 6, 1))
        etask.MoveToDate(target)
        print(f"  MoveToDate(2026-06-01) => OK")
    except Exception as e:
        print(f"  MoveToDate error: {e}")
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    # Reschedule
    project.Reschedule(pywintypes.Time(datetime.now()))
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    print(f"\n  After MoveToDate + Reschedule:")
    print(f"    Start: {etask.Start}")
    print(f"    End: {etask.End}")
    print(f"    GetUserStart: {etask.GetUserStart()}")
    print(f"    GetUserEnd: {etask.GetUserEnd()}")
    print(f"    GetUserDuration: {etask.GetUserDuration().Hours}h")
    print(f"    Constraint: {etask.Constraint}")

    # Test B: ImposedStart only, then duration
    print(f"\n--- TEST B: ImposedStart + Duration ---")
    project.StartTransaction("Set imposed start")
    try:
        target = pywintypes.Time(datetime(2026, 7, 1))
        etask.ImposedStart = target
        print(f"  ImposedStart = 2026-07-01 => OK")
        readback = etask.ImposedStart
        print(f"  ImposedStart readback: {readback}")
    except Exception as e:
        print(f"  ImposedStart error: {e}")
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    print(f"  After ImposedStart + Reschedule:")
    print(f"    Start: {etask.Start}")
    print(f"    End: {etask.End}")
    print(f"    ImposedStart: {etask.ImposedStart}")
    print(f"    Constraint: {etask.Constraint}")

    # Test C: EditToken for dates
    print(f"\n--- TEST C: EditToken/EditTokenV ---")
    project.StartTransaction("EditToken dates")
    try:
        etask.EditTokenV('Start', '01/08/2026')
        print(f"  EditTokenV('Start', '01/08/2026') => OK")
    except Exception as e:
        print(f"  EditTokenV Start error: {e}")
    try:
        etask.EditTokenV('Finish', '15/08/2026')
        print(f"  EditTokenV('Finish', '15/08/2026') => OK")
    except Exception as e:
        print(f"  EditTokenV Finish error: {e}")
    try:
        etask.EditTokenV('Duration', '10d')
        print(f"  EditTokenV('Duration', '10d') => OK")
    except Exception as e:
        print(f"  EditTokenV Duration error: {e}")
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    print(f"  After EditTokenV dates + Reschedule:")
    print(f"    Start: {etask.Start}")
    print(f"    End: {etask.End}")
    print(f"    GetUserStart: {etask.GetUserStart()}")
    print(f"    GetUserEnd: {etask.GetUserEnd()}")
    print(f"    GetUserDuration: {etask.GetUserDuration().Hours}h")

    # Test D: SetUserStart/SetUserEnd (these might be the hidden funcs)
    print(f"\n--- TEST D: SetUserStart/SetUserEnd ---")
    project.StartTransaction("SetUserStart/End")
    et_dyn = win32com.client.Dispatch(etask)
    for method_name in ['SetUserStart', 'SetUserEnd', 'SetStart', 'SetEnd',
                        'SetStartDate', 'SetEndDate', 'SetFinish',
                        'SetImposedStart', 'SetImposedEnd']:
        try:
            fn = getattr(et_dyn, method_name)
            print(f"  {method_name} EXISTS and callable={callable(fn)}")
            if callable(fn):
                try:
                    target = pywintypes.Time(datetime(2026, 9, 1))
                    fn(target)
                    print(f"    {method_name}(2026-09-01) => OK")
                except Exception as e:
                    print(f"    {method_name}(date) => {str(e)[:60]}")
        except AttributeError:
            pass
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    print(f"  After SetUser* + Reschedule:")
    print(f"    Start: {etask.Start}")
    print(f"    End: {etask.End}")

    # Test E: AddConstraint
    print(f"\n--- TEST E: AddConstraint ---")
    project.StartTransaction("AddConstraint")
    try:
        # Try various constraint param combos
        target = pywintypes.Time(datetime(2026, 10, 1))
        # 0=ASAP, 1=ALAP, 2=StartNoEarlierThan, 3=StartNoLaterThan,
        # 4=FinishNoEarlierThan, 5=FinishNoLaterThan, 6=MustStartOn, 7=MustFinishOn
        for ctype in range(8):
            try:
                et_dyn.AddConstraint(ctype, target)
                print(f"  AddConstraint({ctype}, date) => OK")
                break
            except Exception as e:
                if ctype == 0:
                    print(f"  AddConstraint(int, date) => {str(e)[:50]}")
                    break

        # Try with just a date
        try:
            et_dyn.AddConstraint(target)
            print(f"  AddConstraint(date) => OK")
        except Exception as e:
            print(f"  AddConstraint(date) => {str(e)[:50]}")

        # Try with just an int
        try:
            et_dyn.AddConstraint(6)
            print(f"  AddConstraint(6) => OK")
        except Exception as e:
            print(f"  AddConstraint(int) => {str(e)[:50]}")

    except Exception as e:
        print(f"  AddConstraint error: {e}")
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    print(f"  After AddConstraint + Reschedule:")
    print(f"    Start: {etask.Start}")
    print(f"    End: {etask.End}")
    print(f"    Constraint: {etask.Constraint}")
    print(f"    StartConstraintDate: {etask.StartConstraintDate}")
    print(f"    EndConstraintDate: {etask.EndConstraintDate}")

    # Test F: Constraint property (PROP_PUT via type)
    print(f"\n--- TEST F: Constraint property set ---")
    project.StartTransaction("Constraint prop")
    try:
        etask.type = 0  # Task type
        print(f"  type = 0 => OK, type: {etask.type}")
    except Exception as e:
        print(f"  type set error: {e}")
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    # Test G: StartConstraintDate property set
    print(f"\n--- TEST G: StartConstraintDate set ---")
    project.StartTransaction("Constraint date")
    try:
        et_dyn.StartConstraintDate = pywintypes.Time(datetime(2026, 10, 15))
        print(f"  StartConstraintDate set => OK")
    except Exception as e:
        print(f"  StartConstraintDate set error: {str(e)[:60]}")
    try:
        et_dyn.EndConstraintDate = pywintypes.Time(datetime(2026, 11, 1))
        print(f"  EndConstraintDate set => OK")
    except Exception as e:
        print(f"  EndConstraintDate set error: {str(e)[:60]}")
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass

    print(f"  After constraint dates + Reschedule:")
    print(f"    Start: {etask.Start}")
    print(f"    End: {etask.End}")
    print(f"    Constraint: {etask.Constraint}")

    # FINAL: Check what Asta actually shows
    print(f"\n--- FINAL STATE OF TEST BAR ---")
    print(f"  ID: {bar_id}")
    print(f"  Name: {new_bar.Name}")
    print(f"  Start: {etask.Start}")
    print(f"  End: {etask.End}")
    print(f"  GetUserStart: {etask.GetUserStart()}")
    print(f"  GetUserEnd: {etask.GetUserEnd()}")
    try:
        print(f"  GetUserDuration: {etask.GetUserDuration().Hours}h")
    except Exception:
        print(f"  GetUserDuration: N/A")
    print(f"  Constraint: {etask.Constraint}")
    print(f"  ImposedStart: {etask.ImposedStart}")
    print(f"  ImposedEnd: {etask.ImposedEnd}")

    # Clean up - delete test bar
    print(f"\n--- Cleanup ---")
    project.StartTransaction("Delete test bar")
    for i in range(bars.Count, 0, -1):
        try:
            b = bars.Item(i)
            if b.ID == bar_id:
                bars.Remove(i)
                print(f"  Deleted bar ID={bar_id}")
                break
        except Exception:
            pass
    project.EndTransaction()
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass


def test_type_library_full(project):
    """Dump ALL types from the type library file."""
    print("\n" + "#" * 80)
    print("# FULL TYPE LIBRARY — ALL INTERFACES")
    print("#" * 80)

    import winreg
    CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"

    # Find the type library
    tlb = None

    # Try from registry
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{CLSID}\\InprocServer32")
        dll_path = winreg.QueryValue(key, None)
        print(f"DLL from registry: {dll_path}")
        # Try loading the DLL as typelib
        try:
            tlb = pythoncom.LoadTypeLib(dll_path)
            print(f"Loaded typelib from DLL: {tlb.GetTypeInfoCount()} types")
        except Exception:
            pass

        # Try .tlb in same dir
        if not tlb:
            import os
            tlb_dir = os.path.dirname(dll_path)
            for fname in ['teamplan.tlb', 'TeamPlan.tlb', 'astadkit.ocx', 'AstaDkit.ocx']:
                path = os.path.join(tlb_dir, fname)
                if os.path.exists(path):
                    try:
                        tlb = pythoncom.LoadTypeLib(path)
                        print(f"Loaded typelib from {path}: {tlb.GetTypeInfoCount()} types")
                        break
                    except Exception:
                        pass
    except Exception as e:
        print(f"Registry lookup: {e}")

    # Try from running object
    if not tlb:
        try:
            obj = pythoncom.GetActiveObject(CLSID)
            disp = obj.QueryInterface(pythoncom.IID_IDispatch)
            ti = disp.GetTypeInfo()
            tl = ti.GetContainingTypeLib()
            tlb = tl[0]
            print(f"Loaded typelib from running Asta: {tlb.GetTypeInfoCount()} types")
        except Exception as e:
            print(f"Running object typelib: {e}")

    if not tlb:
        print("COULD NOT LOAD TYPE LIBRARY")
        return

    type_count = tlb.GetTypeInfoCount()
    print(f"\nTotal types in library: {type_count}")

    # List ALL types
    for i in range(type_count):
        try:
            ti = tlb.GetTypeInfo(i)
            ta = ti.GetTypeAttr()
            doc = tlb.GetDocumentation(i)
            type_name = doc[0]

            type_kind = ta[5]
            kind_names = {0: "ENUM", 1: "RECORD", 2: "MODULE", 3: "INTERFACE",
                          4: "DISPATCH", 5: "COCLASS", 6: "ALIAS", 7: "UNION"}
            kind_str = kind_names.get(type_kind, f"KIND_{type_kind}")
            func_count = ta[6]
            var_count = ta[7]

            # Only print details for interfaces/dispatches with "Link" or all interfaces
            is_link_related = any(kw in type_name.lower() for kw in
                                  ['link', 'pred', 'succ', 'dep', 'connect'])

            if kind_str in ('INTERFACE', 'DISPATCH') or is_link_related:
                print(f"\n--- [{kind_str}] {type_name} (funcs={func_count}, vars={var_count}) ---")

                if is_link_related or type_name.startswith('ILink') or 'Link' in type_name:
                    # Full dump for link-related types
                    for fi in range(func_count):
                        try:
                            fd = ti.GetFuncDesc(fi)
                            fn = ti.GetNames(fd[0])
                            func_name = fn[0] if fn else f"func_{fi}"
                            invoke_kind = fd[4]
                            kind_s = {1: "METHOD", 2: "GET", 4: "PUT", 8: "PUTREF"}.get(invoke_kind, str(invoke_kind))

                            # Return type
                            try:
                                ret = fd[8]
                                if ret is None:
                                    ret_type = "void"
                                elif isinstance(ret, tuple):
                                    ret_type = _vt_to_str(ret[0])
                                elif isinstance(ret, int):
                                    ret_type = _vt_to_str(ret)
                                else:
                                    ret_type = str(ret)
                            except Exception:
                                ret_type = "?"

                            # Params
                            param_strs = []
                            try:
                                if fd[2]:
                                    for pi, param in enumerate(fd[2]):
                                        p_name = fn[pi + 1] if (pi + 1) < len(fn) else f"p{pi}"
                                        p_type = _safe_param_type(param)
                                        param_strs.append(f"{p_type} {p_name}")
                            except Exception:
                                pass
                            params_s = ", ".join(param_strs)
                            print(f"    [{kind_s}] {ret_type} {func_name}({params_s})")
                        except Exception:
                            pass

                    # Variables (enums)
                    for vi in range(var_count):
                        try:
                            vd = ti.GetVarDesc(vi)
                            vn = ti.GetNames(vd[0])
                            var_name = vn[0] if vn else f"var_{vi}"
                            print(f"    [CONST] {var_name} = {vd[1]}")
                        except Exception:
                            pass
                else:
                    # Just list method names for non-link interfaces
                    method_names = []
                    for fi in range(func_count):
                        try:
                            fd = ti.GetFuncDesc(fi)
                            fn = ti.GetNames(fd[0])
                            if fn:
                                method_names.append(fn[0])
                        except Exception:
                            pass
                    # Filter interesting ones
                    interesting = [m for m in method_names
                                   if any(kw in m.lower() for kw in
                                          ['link', 'pred', 'succ', 'add', 'create', 'connect',
                                           'start', 'end', 'date', 'duration', 'constraint',
                                           'move', 'set', 'imposed'])]
                    if interesting:
                        print(f"    Interesting methods: {interesting}")

            elif kind_str == 'ENUM':
                # Dump enum values if link-related
                if is_link_related:
                    print(f"\n--- [{kind_str}] {type_name} ---")
                    for vi in range(var_count):
                        try:
                            vd = ti.GetVarDesc(vi)
                            vn = ti.GetNames(vd[0])
                            print(f"    {vn[0]} = {vd[1]}")
                        except Exception:
                            pass
                else:
                    print(f"  [{kind_str}] {type_name} ({var_count} values)")

            elif kind_str == 'COCLASS':
                print(f"  [{kind_str}] {type_name}")

        except Exception as e:
            print(f"  Error reading type {i}: {e}")


if __name__ == "__main__":
    print("Asta COM Explorer v9 — Complete Analysis")
    print("=" * 80)
    try:
        app, project = connect()

        # 1. Full IExpandedTask dump with fixed param parser
        test_expanded_task_full(project)

        # 2. Links interface analysis
        test_links_interface(project)

        # 3. Link creation tests
        test_link_creation(project)

        # 4. Date/duration persistence
        test_date_duration_persistence(project)

        # 5. Full type library dump (ALL interfaces)
        test_type_library_full(project)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
