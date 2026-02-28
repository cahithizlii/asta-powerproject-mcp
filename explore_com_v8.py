"""
COM Explorer v8 — Access bars by ID and expand view:
1. Dynamic dispatch bars.Item(id) with known AllBarIds IDs
2. Expand view to show all bars
3. Confirm duration + date + deletion workflow
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


def test_bar_access_by_id(project):
    """Test accessing bars by ID using dynamic dispatch."""
    print("\n" + "="*80)
    print("TEST 1: ACCESS BARS BY ID")
    print("="*80)

    bars = project.Bars

    # Get IDs from AllBarIds
    view = project.CurrentView
    bcv = win32com.client.Dispatch(view)
    bar_ids = bcv.AllBarIds()
    link_ids = bcv.AllLinkIds()
    print(f"AllBarIds: {bar_ids}")
    print(f"AllLinkIds: {link_ids}")

    # Try dynamic dispatch on bars collection
    print("\n--- Dynamic dispatch bars.Item(id) ---")
    bars_dyn = win32com.client.Dispatch(bars)
    print(f"  bars type: {type(bars)}")
    print(f"  bars_dyn type: {type(bars_dyn)}")

    for bid in bar_ids[:5]:
        # Try with static
        try:
            bar = bars.Item(bid)
            print(f"  static bars.Item({bid}) => ID={bar.ID}, Name={bar.Name[:50]}")
        except Exception as e:
            print(f"  static bars.Item({bid}) => {str(e)[:60]}")

        # Try with dynamic
        try:
            bar = bars_dyn.Item(bid)
            print(f"  dynamic bars.Item({bid}) => ID={bar.ID}, Name={bar.Name[:50]}")
        except Exception as e:
            print(f"  dynamic bars.Item({bid}) => {str(e)[:60]}")

    # Also try project.Bars with different interface
    print("\n--- project-level access by ID ---")
    for bid in bar_ids[:3]:
        try:
            # Access directly via project's Bars with dynamic dispatch
            bar = win32com.client.Dispatch(project).Bars.Item(bid)
            print(f"  Dispatch(project).Bars.Item({bid}) => ID={bar.ID}, Name={bar.Name[:50]}")
        except Exception as e:
            print(f"  Dispatch(project).Bars.Item({bid}) => {str(e)[:60]}")

    # Try to get IBarChartView.Bars
    print("\n--- IBarChartView direct bar access ---")
    for method_name in ['GetBar', 'Bar', 'FindBar', 'Bars']:
        try:
            fn = getattr(bcv, method_name)
            if callable(fn):
                for bid in bar_ids[:2]:
                    try:
                        result = fn(bid)
                        print(f"  bcv.{method_name}({bid}) => {result}")
                    except Exception as e:
                        print(f"  bcv.{method_name}({bid}) => {str(e)[:60]}")
            else:
                print(f"  bcv.{method_name} = {fn}")
        except AttributeError:
            pass

    # IBarChartView full attribute dump
    print("\n--- IBarChartView attributes ---")
    attrs = [a for a in dir(bcv) if not a.startswith('_')]
    print(f"  Total: {len(attrs)}")
    for attr in sorted(attrs):
        try:
            val = getattr(bcv, attr)
            if callable(val):
                print(f"  {attr}() [callable]")
            else:
                val_str = str(val)[:60]
                print(f"  {attr} = {val_str}")
        except Exception as e:
            print(f"  {attr} => {str(e)[:50]}")

    return bar_ids, link_ids


def test_view_expand(project, app):
    """Test expanding the view to show all bars."""
    print("\n" + "="*80)
    print("TEST 2: VIEW EXPANSION")
    print("="*80)

    view = project.CurrentView
    bcv = win32com.client.Dispatch(view)

    # Check current filter state
    print("\n--- Current view state ---")
    for prop in ['Filter', 'FilterExpression', 'ActiveFilter', 'IsFiltered',
                 'ExpandAll', 'CollapseAll', 'ShowAllTasks', 'ShowAll',
                 'SuppressSummaryTaskOption', 'SuppressHammockTaskOption',
                 'SuppressExpandedTaskOption',
                 'DisplaySummariesAsTasks', 'DisplayExpandedsAsTasks', 'DisplayHammocksAsTasks',
                 'HierarchyExpansionLevel', 'MaxHierarchyLevel',
                 'SummaryRowAppearanceOnSpreadsheet', 'SummaryRowAppearanceOnBarchart',
                 'ShowHierarchy', 'ExpandLevel', 'Level']:
        try:
            val = getattr(bcv, prop)
            if callable(val):
                try:
                    result = val()
                    print(f"  {prop}() = {result}")
                except Exception as e:
                    print(f"  {prop}() => {str(e)[:40]}")
            else:
                print(f"  {prop} = {val}")
        except:
            pass

    # Try to expand all summaries
    print("\n--- Try to expand/show all ---")
    for method_name in ['ExpandAll', 'ShowAll', 'ShowAllTasks', 'Expand',
                        'ExpandToLevel', 'SetHierarchyLevel', 'ShowHierarchy']:
        try:
            fn = getattr(bcv, method_name)
            if callable(fn):
                try:
                    fn()
                    print(f"  {method_name}() => SUCCESS")
                    ids = bcv.AllBarIds()
                    print(f"    AllBarIds now: {len(ids)} IDs")
                except Exception as e:
                    # Try with parameter
                    try:
                        fn(99)
                        print(f"  {method_name}(99) => SUCCESS")
                        ids = bcv.AllBarIds()
                        print(f"    AllBarIds now: {len(ids)} IDs")
                    except:
                        print(f"  {method_name}() => {str(e)[:60]}")
        except AttributeError:
            pass

    # Try removing filter
    print("\n--- Try removing filter ---")
    for method_name in ['RemoveFilter', 'ClearFilter', 'ResetFilter', 'NoFilter',
                        'RemoveAllFilters', 'UnapplyFilter']:
        try:
            fn = getattr(bcv, method_name)
            fn()
            print(f"  {method_name}() => SUCCESS")
            ids = bcv.AllBarIds()
            print(f"    AllBarIds now: {len(ids)} IDs")
        except AttributeError:
            pass
        except Exception as e:
            print(f"  {method_name}() => {str(e)[:60]}")

    # Check final state
    print("\n--- Final AllBarIds count ---")
    try:
        ids = bcv.AllBarIds()
        print(f"  {len(ids)} bar IDs")
        if len(ids) > 9:
            print(f"  First 20: {ids[:20]}")
            print(f"  Last 5: {ids[-5:]}")
    except Exception as e:
        print(f"  Error: {e}")


def test_link_access(project, link_ids):
    """Test accessing link objects by ID."""
    print("\n" + "="*80)
    print("TEST 3: LINK ACCESS BY ID")
    print("="*80)

    view = project.CurrentView
    bcv = win32com.client.Dispatch(view)

    # Try to get link objects
    print(f"Link IDs: {link_ids}")

    # Try via project
    for method_name in ['Links', 'AllLinks', 'LinkCategorys']:
        try:
            coll = getattr(project, method_name)
            if coll:
                print(f"\n  project.{method_name} = {type(coll)}")
                try:
                    count = coll.Count
                    print(f"    Count: {count}")
                    if count > 0:
                        link = coll.Item(1)
                        print(f"    Item(1): {link}")
                except:
                    pass
        except Exception as e:
            print(f"  project.{method_name} => {str(e)[:60]}")

    # Try via IBarChartView
    for method_name in ['GetLink', 'Link', 'FindLink', 'Links']:
        try:
            fn = getattr(bcv, method_name)
            if callable(fn):
                result = fn(link_ids[0])
                print(f"  bcv.{method_name}({link_ids[0]}) => {result}")
            else:
                print(f"  bcv.{method_name} = {fn}")
        except AttributeError:
            pass
        except Exception as e:
            print(f"  bcv.{method_name}({link_ids[0]}) => {str(e)[:60]}")

    # Get links from the bars we found
    print("\n--- Links from visible bars ---")
    bar_ids = bcv.AllBarIds()
    bars = project.Bars
    for bid in bar_ids[:3]:
        try:
            # Try to find this bar
            for i in range(1, bars.Count + 1):
                b = bars.Item(i)
                if b.ID == bid:
                    etask = b.ExpandedTask
                    li = etask.LinksIn
                    lo = etask.LinksOut
                    print(f"  Bar {bid} ({b.Name[:30]}): LinksIn={li.Count}, LinksOut={lo.Count}")
                    break
        except:
            pass


def test_full_workflow(project):
    """Test the complete working workflow: create, set dates, set duration, delete."""
    print("\n" + "="*80)
    print("TEST 4: FULL WORKFLOW (Create+Dates+Duration+Delete)")
    print("="*80)

    bars = project.Bars

    project.StartTransaction("Full Workflow Test")
    try:
        # Create bar
        new_bar = bars.Add()
        new_bar.Name = "WORKFLOW_TEST"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask
        print(f"1. Created bar: ID={bar_id}")

        # Set dates via ImposedStart/End
        start_date = pywintypes.Time(datetime(2026, 6, 1))
        end_date = pywintypes.Time(datetime(2026, 6, 20))
        etask.ImposedStart = start_date
        etask.ImposedEnd = end_date
        print(f"2. Set ImposedStart=2026-06-01, ImposedEnd=2026-06-20")

        # Set duration via GetDurationFromString + SetUserDuration
        try:
            dur = etask.GetDurationFromString("10d")
            etask.SetUserDuration(dur)
            print(f"3. Set duration=10d (80 hours)")
        except Exception as e:
            print(f"3. Duration error: {e}")

        # End transaction
        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # Reschedule
        project.Reschedule(pywintypes.Time(datetime(2026, 2, 28)))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # Verify
        print(f"\n--- Verification after reschedule ---")
        print(f"  Start: {etask.Start}")
        print(f"  End: {etask.End}")
        print(f"  GetUserStart: {etask.GetUserStart()}")
        print(f"  GetUserEnd: {etask.GetUserEnd()}")
        print(f"  Duration: {etask.Duration()}")
        print(f"  GetUserDuration: {etask.GetUserDuration()}")
        print(f"  ImposedStart: {etask.ImposedStart}")
        print(f"  ImposedEnd: {etask.ImposedEnd}")

        # MoveToDate
        project.StartTransaction("Move test")
        etask.MoveToDate(pywintypes.Time(datetime(2026, 8, 1)))
        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        project.Reschedule(pywintypes.Time(datetime(2026, 2, 28)))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        print(f"\n--- After MoveToDate(2026-08-01) + reschedule ---")
        print(f"  Start: {etask.Start}")
        print(f"  End: {etask.End}")
        print(f"  Duration: {etask.GetUserDuration()}")

        # Delete
        project.StartTransaction("Delete workflow test")
        for i in range(bars.Count, 0, -1):
            try:
                b = bars.Item(i)
                if b.ID == bar_id:
                    bars.Remove(i)
                    print(f"\n4. Deleted bar ID={bar_id} at index {i}")
                    break
            except:
                pass
        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        print(f"5. Final bars.Count: {bars.Count}")

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


if __name__ == "__main__":
    print("Asta COM Explorer v8 — Bar Access + Full Workflow")
    print("="*80)
    try:
        app, project = connect()
        bar_ids, link_ids = test_bar_access_by_id(project)
        test_view_expand(project, app)
        test_link_access(project, link_ids)
        test_full_workflow(project)
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
