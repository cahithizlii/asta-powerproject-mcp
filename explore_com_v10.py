"""
COM Explorer v10 — Focused tests for:
1. LinkTo/LinkFrom between bars (creating links)
2. RemoveConstraint + SetUserDuration (duration setting)
3. Proper date workflow: ImposedStart with ASAP constraint
4. Full create + configure + link + delete workflow

Key fix: Re-fetch bars collection after transactions.
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


def find_bar(project, bar_id):
    """Find a bar by ID across all bars."""
    bars = project.Bars
    for i in range(1, bars.Count + 1):
        try:
            b = bars.Item(i)
            if b.ID == bar_id:
                return b
        except Exception:
            pass
    return None


def wait(project):
    """Wait for notification processing."""
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass


def test_duration_workflow(project):
    """Test: Create bar -> RemoveConstraint -> SetUserDuration -> Set dates."""
    print("\n" + "=" * 80)
    print("TEST 1: DURATION WORKFLOW")
    print("=" * 80)

    # Create bar
    project.StartTransaction("Create bar")
    new_bar = project.Bars.Add()
    new_bar.Name = "DUR_WORKFLOW"
    bar_id = new_bar.ID
    et = new_bar.ExpandedTask
    print(f"Created bar ID={bar_id}")
    project.EndTransaction()
    wait(project)

    # Re-fetch
    bar = find_bar(project, bar_id)
    if not bar:
        print(f"ERROR: Can't find bar {bar_id}!")
        # Try counting
        bars = project.Bars
        print(f"  bars.Count = {bars.Count}")
        for i in range(1, bars.Count + 1):
            b = bars.Item(i)
            print(f"  [{i}] ID={b.ID}, Name={b.Name}")
        return None

    et = bar.ExpandedTask
    print(f"\nDefaults:")
    print(f"  Start={et.Start}, End={et.End}")
    print(f"  Constraint={et.Constraint}")
    print(f"  ImposedStart={et.ImposedStart}, ImposedEnd={et.ImposedEnd}")
    try:
        print(f"  Duration={et.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"  Duration error: {e}")

    # Step 1: RemoveConstraint
    print(f"\n--- Step 1: RemoveConstraint ---")
    project.StartTransaction("RemoveConstraint")
    result = et.RemoveConstraint()
    print(f"  RemoveConstraint() => {result}")
    project.EndTransaction()
    wait(project)

    print(f"  Constraint now: {et.Constraint}")
    print(f"  ImposedStart={et.ImposedStart}, ImposedEnd={et.ImposedEnd}")

    # Step 2: SetUserDuration
    print(f"\n--- Step 2: SetUserDuration ---")
    project.StartTransaction("SetDuration")
    dur_obj = et.GetDurationFromString("10d")
    print(f"  GetDurationFromString('10d') => {dur_obj.Hours}h")
    et.SetUserDuration(dur_obj)
    print(f"  SetUserDuration OK")
    project.EndTransaction()
    wait(project)

    print(f"  After SetUserDuration:")
    try:
        print(f"    Duration={et.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"    Duration error: {e}")
    print(f"    Start={et.Start}, End={et.End}")

    # Reschedule
    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)
    print(f"  After Reschedule:")
    try:
        print(f"    Duration={et.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"    Duration error: {e}")
    print(f"    Start={et.Start}, End={et.End}")

    # Step 3: Set start date via ImposedStart
    print(f"\n--- Step 3: ImposedStart ---")
    project.StartTransaction("ImposedStart")
    et.ImposedStart = pywintypes.Time(datetime(2026, 6, 1))
    print(f"  ImposedStart = 2026-06-01 => OK")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)
    print(f"  After ImposedStart + Reschedule:")
    try:
        print(f"    Duration={et.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"    Duration error: {e}")
    print(f"    Start={et.Start}")
    print(f"    End={et.End}")
    print(f"    Constraint={et.Constraint}")

    # Step 4: Try setting ImposedEnd to limit duration
    print(f"\n--- Step 4: ImposedEnd for end date ---")
    project.StartTransaction("ImposedEnd")
    et.ImposedEnd = pywintypes.Time(datetime(2026, 6, 12))
    print(f"  ImposedEnd = 2026-06-12 => OK")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)
    print(f"  After ImposedEnd + Reschedule:")
    try:
        print(f"    Duration={et.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"    Duration error: {e}")
    print(f"    Start={et.Start}")
    print(f"    End={et.End}")
    print(f"    Constraint={et.Constraint}")

    # Step 5: Clear both and set via AddConstraint + date properties
    print(f"\n--- Step 5: Clear + AddConstraint(SNET) + dates ---")
    project.StartTransaction("Clear constraints")
    et.RemoveConstraint()
    print(f"  RemoveConstraint => Constraint={et.Constraint}")

    # Set constraint type 2 = Start No Earlier Than
    et.AddConstraint(2)
    print(f"  AddConstraint(2=SNET) => Constraint={et.Constraint}")

    # Set constraint date
    et.StartConstraintDate = pywintypes.Time(datetime(2026, 8, 1))
    print(f"  StartConstraintDate = 2026-08-01 => OK")

    # Set duration
    dur_obj = et.GetDurationFromString("15d")
    et.SetUserDuration(dur_obj)
    print(f"  SetUserDuration(15d) => OK")

    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"  After AddConstraint(SNET) + StartConstraintDate + SetUserDuration + Reschedule:")
    try:
        print(f"    Duration={et.GetUserDuration().Hours}h")
    except Exception as e:
        print(f"    Duration error: {e}")
    print(f"    Start={et.Start}")
    print(f"    End={et.End}")
    print(f"    Constraint={et.Constraint}")
    print(f"    StartConstraintDate={et.StartConstraintDate}")

    return bar_id


def test_link_creation(project):
    """Test LinkTo and LinkFrom."""
    print("\n" + "=" * 80)
    print("TEST 2: LINK CREATION")
    print("=" * 80)

    bar_ids = []

    # Create 2 bars
    project.StartTransaction("Create link test bars")
    b1 = project.Bars.Add()
    b1.Name = "LINK_PRED"
    b2 = project.Bars.Add()
    b2.Name = "LINK_SUCC"
    bar_ids = [b1.ID, b2.ID]
    print(f"Created bars: {bar_ids}")
    project.EndTransaction()
    wait(project)

    # Re-fetch
    bar1 = find_bar(project, bar_ids[0])
    bar2 = find_bar(project, bar_ids[1])
    if not bar1:
        print(f"  Can't find bar1 ({bar_ids[0]}), listing all:")
        bars = project.Bars
        for i in range(1, bars.Count + 1):
            b = bars.Item(i)
            print(f"    [{i}] ID={b.ID}, Name={b.Name}")
        return bar_ids
    if not bar2:
        print(f"  Can't find bar2 ({bar_ids[1]})")
        return bar_ids

    et1 = bar1.ExpandedTask
    et2 = bar2.ExpandedTask
    print(f"  Bar1: ID={bar1.ID}, Name={bar1.Name}")
    print(f"  Bar2: ID={bar2.ID}, Name={bar2.Name}")
    print(f"  et1.LinksOut: {et1.LinksOut.Count}, et1.LinksIn: {et1.LinksIn.Count}")
    print(f"  et2.LinksOut: {et2.LinksOut.Count}, et2.LinksIn: {et2.LinksIn.Count}")

    # Test LinkTo: et1 -> et2 (FS)
    print(f"\n--- LinkTo ---")
    project.StartTransaction("LinkTo")
    try:
        link = et1.LinkTo(et2)
        print(f"  et1.LinkTo(et2) => {link}")
        if link:
            print(f"    type: {type(link)}")
            try:
                print(f"    ID: {link.ID}")
            except Exception:
                pass
            # Dump attributes
            link_dyn = win32com.client.Dispatch(link)
            for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
                try:
                    val = getattr(link_dyn, attr)
                    if not callable(val):
                        vs = str(val)[:50]
                        print(f"    {attr} = {vs}")
                except Exception:
                    pass
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  LinkTo error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Verify
    print(f"\n  After LinkTo:")
    print(f"    et1.LinksOut: {et1.LinksOut.Count}")
    print(f"    et2.LinksIn: {et2.LinksIn.Count}")
    print(f"    et1.HasSuccessor: {et1.HasSuccessor}")
    print(f"    et2.HasPredecessor: {et2.HasPredecessor}")

    # If link exists, dump the ILink interface
    if et1.LinksOut.Count > 0:
        link_obj = et1.LinksOut.Item(1)
        print(f"\n  ILink object dump:")
        link_dyn = win32com.client.Dispatch(link_obj)
        for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
            try:
                val = getattr(link_dyn, attr)
                if not callable(val):
                    vs = str(val)[:60]
                    print(f"    {attr} = {vs}")
                else:
                    print(f"    {attr}() [callable]")
            except Exception as e:
                print(f"    {attr} => ERR: {str(e)[:40]}")

        # Full typelib dump of ILink
        print(f"\n  ILink typelib dump:")
        try:
            type_info = link_obj._oleobj_.GetTypeInfo()
            type_attr = type_info.GetTypeAttr()
            func_count = type_attr[6]
            type_name = type_info.GetDocumentation(-1)[0]
            print(f"    Interface: {type_name}, Functions: {func_count}")
            for i in range(func_count):
                try:
                    fd = type_info.GetFuncDesc(i)
                    fn = type_info.GetNames(fd[0])
                    func_name = fn[0] if fn else f"func_{i}"
                    invoke_kind = fd[4]
                    kind_str = {1: "METHOD", 2: "GET", 4: "PUT", 8: "PUTREF"}.get(invoke_kind, str(invoke_kind))

                    params = []
                    if fd[2]:
                        for pi, p in enumerate(fd[2]):
                            pn = fn[pi+1] if pi+1 < len(fn) else f"p{pi}"
                            try:
                                pt = p[0] if isinstance(p[0], int) else (p[0][0] if isinstance(p[0], tuple) else '?')
                            except Exception:
                                pt = '?'
                            params.append(f"{pt}:{pn}")
                    params_s = ", ".join(params)
                    print(f"    [{kind_str}] {func_name}({params_s})")
                except Exception:
                    pass
        except Exception as e:
            print(f"    Type info error: {e}")

    # Reschedule
    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"\n  After Reschedule:")
    print(f"    Bar1 Start={et1.Start}, End={et1.End}")
    print(f"    Bar2 Start={et2.Start}, End={et2.End}")

    return bar_ids


def test_link_with_type(project, bar_ids_to_use=None):
    """Test creating links with specific type (SS, FF, SF)."""
    print("\n" + "=" * 80)
    print("TEST 3: LINK TYPE MODIFICATION")
    print("=" * 80)

    if not bar_ids_to_use or len(bar_ids_to_use) < 2:
        print("  Need bar_ids from test 2")
        return

    bar1 = find_bar(project, bar_ids_to_use[0])
    bar2 = find_bar(project, bar_ids_to_use[1])
    if not (bar1 and bar2):
        print("  Can't find bars")
        return

    et1 = bar1.ExpandedTask
    et2 = bar2.ExpandedTask

    if et1.LinksOut.Count > 0:
        link = et1.LinksOut.Item(1)
        link_dyn = win32com.client.Dispatch(link)

        # Try to find and set link type
        print("  Attempting to modify link type...")
        for prop_name in ['LinkType', 'Type', 'type', 'Kind', 'Category',
                          'Lag', 'LeadLag', 'float_type', 'link_type',
                          'StartEnd', 'link_kind']:
            try:
                val = getattr(link_dyn, prop_name)
                print(f"    {prop_name} = {val}")
            except AttributeError:
                pass
            except Exception as e:
                print(f"    {prop_name} => {str(e)[:40]}")

        # Check the link through GetToken
        for token in ['Type', 'LinkType', 'Lag', 'LeadLag', 'Category']:
            try:
                val = link_dyn.GetToken(token)
                print(f"    GetToken('{token}') = {val}")
            except Exception:
                pass

        # Try EditTokenV on the link
        project.StartTransaction("Modify link")
        for token, value in [('Lag', '2d'), ('Type', 'SS'), ('LinkType', '1')]:
            try:
                link_dyn.EditTokenV(token, value)
                print(f"    EditTokenV('{token}', '{value}') => OK")
            except Exception as e:
                print(f"    EditTokenV('{token}', '{value}') => {str(e)[:50]}")
        project.EndTransaction()
        wait(project)

        # Check if it changed
        for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
            try:
                val = getattr(link_dyn, attr)
                if not callable(val):
                    vs = str(val)[:50]
                    print(f"    {attr} = {vs}")
            except Exception:
                pass


def cleanup_bars(project, bar_ids):
    """Clean up test bars."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup: deleting {len(bar_ids)} bars ---")
    project.StartTransaction("Cleanup")
    bars = project.Bars
    deleted = 0
    for target_id in reversed(bar_ids):
        for i in range(bars.Count, 0, -1):
            try:
                b = bars.Item(i)
                if b.ID == target_id:
                    bars.Remove(i)
                    print(f"  Deleted ID={target_id}")
                    deleted += 1
                    break
            except Exception:
                pass
    project.EndTransaction()
    wait(project)
    print(f"  Deleted {deleted}/{len(bar_ids)} bars")


if __name__ == "__main__":
    print("Asta COM Explorer v10 — Focused Link + Duration Tests")
    print("=" * 80)
    all_created = []
    try:
        app, project = connect()

        # Test 1: Duration workflow
        bar_id = test_duration_workflow(project)
        if bar_id:
            all_created.append(bar_id)

        # Test 2: Link creation
        link_bar_ids = test_link_creation(project)
        if link_bar_ids:
            all_created.extend(link_bar_ids)

            # Test 3: Link type modification
            test_link_with_type(project, link_bar_ids)

        # Cleanup all
        cleanup_bars(project, all_created)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        # Try cleanup
        if all_created:
            try:
                cleanup_bars(project, all_created)
            except Exception:
                pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
