"""
COM Explorer v7 — Test new discoveries from typelib:
1. bars.All() — get ALL bars as variant
2. bar.Tasks() — sub-tasks from bar
3. etask.NextTask()/PreviousTask() — task navigation
4. IBarChartView vs IView — AllBarIds
5. project.OrphanBars, ProjectSummary
6. Duration object from GetUserDuration -> use for SetUserDuration
7. Links Add() with transaction
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


def test_bars_all(project):
    """TEST: bars.All() to get ALL bars."""
    print("\n" + "="*80)
    print("TEST 1: bars.All() — ALL BARS")
    print("="*80)

    bars = project.Bars
    print(f"bars.Count = {bars.Count}")

    # Try All()
    print("\n--- bars.All() ---")
    try:
        all_bars = bars.All()
        print(f"  Type: {type(all_bars)}")
        if all_bars is not None:
            if hasattr(all_bars, '__len__'):
                print(f"  Length: {len(all_bars)}")
            if hasattr(all_bars, '__iter__'):
                bar_list = list(all_bars)
                print(f"  List length: {len(bar_list)}")
                for i, b in enumerate(bar_list[:20]):
                    try:
                        print(f"  [{i}] ID={b.ID}, Name={b.Name[:50]}")
                    except:
                        print(f"  [{i}] {b} (type={type(b)})")
            else:
                print(f"  Value: {all_bars}")
    except Exception as e:
        print(f"  Error: {e}")

    # Try _NewEnum
    print("\n--- bars._NewEnum() ---")
    try:
        enum = bars._NewEnum()
        print(f"  Type: {type(enum)}")
        count = 0
        for b in bars:
            try:
                print(f"  [{count}] ID={b.ID}, Name={b.Name[:50]}")
                count += 1
                if count >= 20:
                    print("  ... (stopped at 20)")
                    break
            except:
                print(f"  [{count}] {b}")
                count += 1
        print(f"  Total iterated: {count}")
    except Exception as e:
        print(f"  Error: {e}")


def test_bar_tasks(project):
    """TEST: bar.Tasks() — sub-tasks from bar."""
    print("\n" + "="*80)
    print("TEST 2: bar.Tasks() and etask navigation")
    print("="*80)

    bars = project.Bars
    bar1 = bars.Item(1)
    etask1 = bar1.ExpandedTask
    print(f"Bar 1: ID={bar1.ID}, Name={bar1.Name}")

    # bar.Tasks()
    print("\n--- bar.Tasks() ---")
    try:
        tasks = bar1.Tasks
        print(f"  Type: {type(tasks)}")
        if tasks:
            try:
                count = tasks.Count
                print(f"  Count: {count}")
                for i in range(1, min(count + 1, 10)):
                    try:
                        t = tasks.Item(i)
                        print(f"  [{i}] {t.Name[:50] if hasattr(t, 'Name') else t}")
                    except Exception as e:
                        print(f"  [{i}] {str(e)[:60]}")
            except Exception as e:
                print(f"  Count error: {e}")
    except Exception as e:
        print(f"  Error: {e}")

    # bar.Etask
    print("\n--- bar.Etask ---")
    try:
        et = bar1.Etask
        print(f"  Type: {type(et)}")
        if et:
            print(f"  Name: {et.Name[:50] if hasattr(et, 'Name') else 'N/A'}")
    except Exception as e:
        print(f"  Error: {e}")

    # bar.Summary
    print("\n--- bar.Summary ---")
    try:
        summary = bar1.Summary
        print(f"  Type: {type(summary)}")
        if summary:
            try:
                print(f"  ID: {summary.ID}, Name: {summary.Name[:50]}")
            except:
                print(f"  Value: {summary}")
    except Exception as e:
        print(f"  Error: {e}")

    # etask.NextTask() / PreviousTask()
    print("\n--- etask.NextTask() traversal ---")
    try:
        current = etask1.NextTask()
        if current:
            print(f"  First NextTask: {type(current)}")
            try:
                print(f"  Name: {current.Name[:50]}")
            except:
                pass

            count = 1
            visited = set()
            while current and count < 50:
                try:
                    # Get identifying info
                    try:
                        name = current.Name[:45] if hasattr(current, 'Name') else "?"
                    except:
                        name = "?"
                    try:
                        bar = current.Bar
                        bid = bar.ID if bar else "no_bar"
                    except:
                        bid = "no_bar"

                    if str(bid) in visited:
                        print(f"  Loop at bar ID={bid}")
                        break
                    visited.add(str(bid))
                    print(f"  [{count}] BarID={bid}, {name}")
                    count += 1
                    current = current.NextTask()
                except Exception as e:
                    print(f"  Error at {count}: {str(e)[:60]}")
                    break
            print(f"  Total via NextTask: {count}")
        else:
            print("  NextTask returned None")
    except Exception as e:
        print(f"  Error: {e}")

    # etask.Activities()
    print("\n--- etask.Activities() ---")
    try:
        acts = etask1.Activities
        print(f"  Type: {type(acts)}")
        if acts:
            try:
                count = acts.Count
                print(f"  Count: {count}")
            except:
                pass
    except Exception as e:
        print(f"  Error: {e}")


def test_project_special(project):
    """TEST: project.OrphanBars, ProjectSummary."""
    print("\n" + "="*80)
    print("TEST 3: PROJECT SPECIAL PROPERTIES")
    print("="*80)

    # ProjectSummary
    print("\n--- project.ProjectSummary ---")
    try:
        ps = project.ProjectSummary
        print(f"  Type: {type(ps)}")
        if ps:
            try:
                print(f"  ID: {ps.ID}, Name: {ps.Name[:50]}")
            except:
                print(f"  Value: {ps}")
    except Exception as e:
        print(f"  Error: {e}")

    # OrphanBars
    print("\n--- project.OrphanBars ---")
    try:
        ob = project.OrphanBars
        print(f"  Type: {type(ob)}")
        if ob:
            try:
                count = ob.Count
                print(f"  Count: {count}")
                for i in range(1, min(count + 1, 5)):
                    b = ob.Item(i)
                    print(f"  [{i}] ID={b.ID}, Name={b.Name[:50]}")
            except Exception as e:
                print(f"  Count/Item error: {e}")
    except Exception as e:
        print(f"  Error: {e}")

    # BaselineSummaries
    print("\n--- project.BaselineSummaries ---")
    try:
        bs = project.BaselineSummaries
        print(f"  Type: {type(bs)}")
    except Exception as e:
        print(f"  Error: {e}")


def test_barchartview(project, app):
    """TEST: IBarChartView (different from IView)."""
    print("\n" + "="*80)
    print("TEST 4: IBarChartView — AllBarIds")
    print("="*80)

    view = project.CurrentView
    print(f"IView type: {type(view)}")

    # Try to get IBarChartView
    # IBarChartView CLSID might be different
    print("\n--- Trying IBarChartView methods on IView ---")
    for method in ['AllBarIds', 'AllLinkIds', 'AllTaskBaseIds', 'AllAllocIds',
                   'GetDisplayedNonWorkingExceptions']:
        try:
            fn = getattr(view, method)
            result = fn()
            if result:
                if hasattr(result, '__iter__'):
                    items = list(result)
                    print(f"  {method}() => {len(items)} items")
                    if items:
                        print(f"    First 10: {items[:10]}")
                else:
                    print(f"  {method}() => {result}")
            else:
                print(f"  {method}() => None")
        except Exception as e:
            print(f"  {method}() => {str(e)[:60]}")

    # Try to cast/query IBarChartView interface
    print("\n--- Trying to get IBarChartView ---")
    try:
        # IBarChartView might be accessible via QueryInterface
        bcv_clsid = "{A57A0000-0200-0003-B2C5-00C0DF438041}"  # coclass_clsid from IView
        print(f"  View coclass: {view.coclass_clsid}")
        print(f"  View CLSID: {view.CLSID}")

        # Try dynamic dispatch
        view_disp = win32com.client.Dispatch(view)
        print(f"  Dynamic dispatch type: {type(view_disp)}")
        for method in ['AllBarIds', 'AllLinkIds']:
            try:
                result = getattr(view_disp, method)()
                print(f"  dynamic.{method}() => {result}")
            except Exception as e:
                print(f"  dynamic.{method}() => {str(e)[:60]}")
    except Exception as e:
        print(f"  Error: {e}")

    # View.selection()
    print("\n--- view.selection() ---")
    try:
        sel = view.selection()
        print(f"  Type: {type(sel)}")
        if sel:
            attrs = [a for a in dir(sel) if not a.startswith('_')]
            print(f"  Attrs: {attrs}")
            try:
                count = sel.Count
                print(f"  Count: {count}")
            except:
                pass
    except Exception as e:
        print(f"  Error: {e}")

    # View.Definition()
    print("\n--- view.Definition() ---")
    try:
        defn = view.Definition()
        print(f"  Type: {type(defn)}")
        if defn:
            attrs = [a for a in dir(defn) if not a.startswith('_')]
            print(f"  Attrs ({len(attrs)}): {attrs[:30]}")
            # Check for bar-related methods
            bar_attrs = [a for a in attrs if 'bar' in a.lower() or 'task' in a.lower() or 'all' in a.lower()]
            print(f"  Bar/Task/All attrs: {bar_attrs}")
    except Exception as e:
        print(f"  Error: {e}")


def test_duration_object(project):
    """TEST: Get duration object from existing task, use for SetUserDuration."""
    print("\n" + "="*80)
    print("TEST 5: DURATION OBJECT")
    print("="*80)

    bars = project.Bars
    bar1 = bars.Item(1)
    etask1 = bar1.ExpandedTask

    # Get duration from existing task
    print("\n--- GetUserDuration() from existing task ---")
    try:
        dur = etask1.GetUserDuration()
        print(f"  Type: {type(dur)}")
        print(f"  Value: {dur}")
        print(f"  Attrs: {[a for a in dir(dur) if not a.startswith('_')]}")
    except Exception as e:
        print(f"  Error: {e}")

    # Get Duration property
    print("\n--- Duration() from existing task ---")
    try:
        dur = etask1.Duration()
        print(f"  Type: {type(dur)}")
        print(f"  Value: {dur}")
    except Exception as e:
        print(f"  Error: {e}")

    # Get PlannedDuration
    print("\n--- PlannedDuration() ---")
    try:
        dur = etask1.PlannedDuration()
        print(f"  Type: {type(dur)}")
        print(f"  Value: {dur}")
    except Exception as e:
        print(f"  Error: {e}")

    # GetDurationFromString
    print("\n--- GetDurationFromString() ---")
    for val in ["10d", "10", "10.0", "80h", "2w", "10 days", "80 hours"]:
        try:
            dur = etask1.GetDurationFromString(val)
            print(f"  GetDurationFromString('{val}') => {dur} (type={type(dur)})")
        except Exception as e:
            print(f"  GetDurationFromString('{val}') => {str(e)[:60]}")

    # Create a new bar and try SetUserDuration with duration object
    print("\n--- SetUserDuration with duration OBJECT ---")
    project.StartTransaction("Dur Object Test")
    try:
        new_bar = bars.Add()
        new_bar.Name = "TEST_DUR_OBJ"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask

        # First get a duration object from the existing task
        try:
            dur_obj = etask1.GetUserDuration()
            print(f"  Got duration object: {dur_obj} (type={type(dur_obj)})")

            # Try SetUserDuration with this object
            try:
                etask.SetUserDuration(dur_obj)
                print(f"  SetUserDuration(dur_obj) => SUCCESS!")
                print(f"  New duration: {etask.GetUserDuration()}")
            except Exception as e:
                print(f"  SetUserDuration(dur_obj) => {str(e)[:80]}")
        except Exception as e:
            print(f"  GetUserDuration error: {e}")

        # Try GetDurationFromString then SetUserDuration
        try:
            dur_str = etask.GetDurationFromString("10d")
            print(f"\n  GetDurationFromString('10d') => {dur_str}")
            if dur_str is not None:
                try:
                    etask.SetUserDuration(dur_str)
                    print(f"  SetUserDuration(from_string) => SUCCESS!")
                    print(f"  Duration now: {etask.GetUserDuration()}")
                except Exception as e:
                    print(f"  SetUserDuration(from_string) => {str(e)[:80]}")
        except Exception as e:
            print(f"  GetDurationFromString error: {e}")

        # Cleanup
        for i in range(1, bars.Count + 1):
            try:
                b = bars.Item(i)
                if b.ID == bar_id:
                    bars.Remove(i)
                    break
            except:
                pass
        project.EndTransaction()
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except:
            pass


def test_link_in_transaction(project):
    """TEST: Link creation within a transaction."""
    print("\n" + "="*80)
    print("TEST 6: LINK CREATION IN TRANSACTION")
    print("="*80)

    bars = project.Bars

    # Create two bars and try to link them
    project.StartTransaction("Link Test")
    try:
        bar_a = bars.Add()
        bar_a.Name = "LINK_TEST_A"
        bar_b = bars.Add()
        bar_b.Name = "LINK_TEST_B"
        etask_a = bar_a.ExpandedTask
        etask_b = bar_b.ExpandedTask
        print(f"Created: A={bar_a.ID}, B={bar_b.ID}")

        # Set some dates
        etask_a.ImposedStart = pywintypes.Time(datetime(2026, 6, 1))
        etask_b.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))

        # Try link creation methods
        print("\n--- LinksOut.Add() on A ---")
        lo = etask_a.LinksOut
        print(f"  LinksOut type: {type(lo)}")
        print(f"  LinksOut count: {lo.Count}")
        print(f"  LinksOut attrs: {[a for a in dir(lo) if not a.startswith('_')]}")

        # The typelib shows Add() with no params
        attempts = [
            ("lo.Add()", lambda: lo.Add()),
            ("lo.Add(bar_b)", lambda: lo.Add(bar_b)),
            ("lo.Add(etask_b)", lambda: lo.Add(etask_b)),
            ("lo.Add(bar_b.ID)", lambda: lo.Add(bar_b.ID)),
        ]
        for desc, fn in attempts:
            try:
                link = fn()
                print(f"  {desc} => SUCCESS! {link}")
                if link:
                    # Try to set link properties
                    try:
                        print(f"    Link type: {type(link)}")
                        print(f"    Link attrs: {[a for a in dir(link) if not a.startswith('_')]}")
                    except:
                        pass
                break
            except Exception as e:
                print(f"  {desc} => {str(e)[:80]}")

        # Try LinksIn.Add() on B
        print("\n--- LinksIn.Add() on B ---")
        li = etask_b.LinksIn
        print(f"  LinksIn attrs: {[a for a in dir(li) if not a.startswith('_')]}")

        attempts = [
            ("li.Add()", lambda: li.Add()),
            ("li.Add(bar_a)", lambda: li.Add(bar_a)),
            ("li.Add(etask_a)", lambda: li.Add(etask_a)),
            ("li.Add(bar_a.ID)", lambda: li.Add(bar_a.ID)),
        ]
        for desc, fn in attempts:
            try:
                link = fn()
                print(f"  {desc} => SUCCESS! {link}")
                break
            except Exception as e:
                print(f"  {desc} => {str(e)[:80]}")

        # Try project-level link creation
        print("\n--- Project link methods ---")
        for method_name in ['AddLink', 'CreateLink', 'MakeLink', 'Link']:
            try:
                fn = getattr(project, method_name)
                result = fn(bar_a, bar_b)
                print(f"  project.{method_name}(a, b) => {result}")
            except AttributeError:
                pass
            except Exception as e:
                print(f"  project.{method_name}(a, b) => {str(e)[:60]}")

        # Try EditTokenV with link-related tokens
        print("\n--- EditTokenV link tokens ---")
        for token in ['Predecessor', 'Successor', 'Link', 'Dependency']:
            try:
                etask_b.EditTokenV(token, str(bar_a.ID))
                print(f"  EditTokenV('{token}', id) => SUCCESS!")
            except Exception as e:
                err = str(e)[:60]
                if 'Unknown token' in err:
                    pass
                else:
                    print(f"  EditTokenV('{token}') => {err}")

        # Cleanup
        for i in range(bars.Count, 0, -1):
            try:
                b = bars.Item(i)
                if 'LINK_TEST' in (b.Name or ''):
                    bars.Remove(i)
            except:
                pass
        project.EndTransaction()

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


if __name__ == "__main__":
    print("Asta COM Explorer v7 — New Discoveries")
    print("="*80)
    try:
        app, project = connect()
        test_bars_all(project)
        test_bar_tasks(project)
        test_project_special(project)
        test_barchartview(project, app)
        test_duration_object(project)
        test_link_in_transaction(project)
        print("\n" + "="*80)
        print("ALL TESTS COMPLETE")
        print("="*80)
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
