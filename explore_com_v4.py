"""
COM Explorer v4 — Test direct ID access, CurrentView, and verified date setting.
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

def test_direct_id_access(project):
    """Test bars.Item(id) with known nested task IDs from MPXJ."""
    print("\n" + "="*80)
    print("TEST 1: DIRECT ID ACCESS (bars.Item(id) for nested bars)")
    print("="*80)

    bars = project.Bars
    print(f"bars.Count = {bars.Count}")

    # IDs from MPXJ critical path: Temel Betonu=845, +0.00 Kotu Dosemesi=846, etc.
    # Also first few from MPXJ: Milestones=56(uid), Yer Teslimi=57(uid)
    # MPXJ uses different IDs (id=1...2377) vs COM IDs (unique_id)
    # The MPXJ unique_ids were: 55=Akfa, 56=Milestones, 57=Yer Teslimi...
    # COM ID for top-level is 1082
    # Let's try various ID ranges

    test_ids = [
        # Known COM IDs
        1082,   # Akfa Medline Polyclinic (top-level, confirmed working)
        # Try nearby IDs
        1083, 1084, 1085, 1086, 1087, 1088, 1089, 1090,
        # Try IDs in various ranges
        1100, 1200, 1300, 1398, 1399,  # From MPXJ critical path
        2000, 2279, 2280, 2369,
        # Try low IDs
        1, 2, 3, 4, 5, 10, 50, 100, 500, 1000,
        # Test bars that were just created (7620, 7621, 7622)
        7620, 7621, 7622,
    ]

    found_ids = []
    for tid in test_ids:
        try:
            bar = bars.Item(tid)
            if bar is not None:
                bid = bar.ID
                name = bar.Name[:50] if bar.Name else "?"
                print(f"  bars.Item({tid}) => ID={bid}, Name={name}")
                found_ids.append(bid)
        except Exception as e:
            err = str(e)[:60]
            # Only show "item does not exist" once
            if "Item does not exist" in err or "does not exist" in err:
                pass  # Skip common errors
            else:
                print(f"  bars.Item({tid}) => {err}")

    print(f"\nFound {len(found_ids)} valid bars out of {len(test_ids)} tested")
    return found_ids


def test_view_access(project, app):
    """Test CurrentView and AllBarIds."""
    print("\n" + "="*80)
    print("TEST 2: VIEW ACCESS AND AllBarIds")
    print("="*80)

    # Try various ways to get the view
    print("\n--- Finding the view ---")
    view = None

    # app.ActiveView
    try:
        view = app.ActiveView
        print(f"  app.ActiveView => {view}")
    except Exception as e:
        print(f"  app.ActiveView => {str(e)[:60]}")

    # project.CurrentView
    try:
        view = project.CurrentView
        print(f"  project.CurrentView => {view}")
        if view:
            print(f"  View type: {type(view)}")
            print(f"  View name: {view.Name if hasattr(view, 'Name') else 'N/A'}")
    except Exception as e:
        print(f"  project.CurrentView => {str(e)[:60]}")

    # project.Views
    try:
        views = project.Views
        print(f"  project.Views => {views}")
        if views:
            print(f"  Views.Count: {views.Count}")
            for i in range(1, min(views.Count + 1, 4)):
                v = views.Item(i)
                print(f"    View {i}: {v.Name if hasattr(v, 'Name') else v}")
    except Exception as e:
        print(f"  project.Views => {str(e)[:60]}")

    if view:
        print(f"\n--- View attributes ---")
        for attr in sorted(dir(view)):
            if not attr.startswith('_') and 'Bar' in attr or 'Link' in attr or 'All' in attr or 'Id' in attr:
                print(f"  {attr}")

        # AllBarIds
        print(f"\n--- AllBarIds ---")
        try:
            bar_ids = view.AllBarIds()
            if bar_ids is not None:
                id_list = list(bar_ids) if hasattr(bar_ids, '__iter__') else [bar_ids]
                print(f"  AllBarIds returned {len(id_list)} IDs")
                print(f"  First 20: {id_list[:20]}")
                print(f"  Last 5: {id_list[-5:]}")

                # Try to access these IDs via bars.Item()
                bars = project.Bars
                success = 0
                for bid in id_list[:10]:
                    try:
                        bar = bars.Item(int(bid))
                        if bar:
                            success += 1
                    except:
                        pass
                print(f"  Can access via bars.Item(): {success}/10")
            else:
                print("  AllBarIds returned None")
        except Exception as e:
            print(f"  AllBarIds error: {e}")

        # AllLinkIds
        print(f"\n--- AllLinkIds ---")
        try:
            link_ids = view.AllLinkIds()
            if link_ids is not None:
                id_list = list(link_ids) if hasattr(link_ids, '__iter__') else [link_ids]
                print(f"  AllLinkIds returned {len(id_list)} IDs")
                print(f"  First 20: {id_list[:20]}")
            else:
                print("  AllLinkIds returned None")
        except Exception as e:
            print(f"  AllLinkIds error: {e}")


def test_date_setting_verified(project):
    """Test date setting with verification after reschedule."""
    print("\n" + "="*80)
    print("TEST 3: DATE SETTING WITH VERIFICATION")
    print("="*80)

    bars = project.Bars
    project.StartTransaction("Test Dates Verified")
    try:
        new_bar = bars.Add()
        new_bar.Name = "TEST_DATES_V4"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask
        print(f"Created test bar: ID={bar_id}")

        start = pywintypes.Time(datetime(2026, 6, 1))
        end = pywintypes.Time(datetime(2026, 6, 15))

        # Set dates using all known working methods
        print("\n--- Setting Start via EditTokenV ---")
        try:
            etask.EditTokenV("Start", start)
            print("  EditTokenV('Start') => SUCCESS")
        except Exception as e:
            print(f"  EditTokenV('Start') => {e}")

        print("\n--- Setting End via EditTokenV ---")
        try:
            etask.EditTokenV("End", end)
            print("  EditTokenV('End') => SUCCESS")
        except Exception as e:
            print(f"  EditTokenV('End') => {e}")

        # End transaction and reschedule
        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # Read back after transaction
        print("\n--- After EndTransaction ---")
        try:
            print(f"  etask.Start = {etask.Start}")
            print(f"  etask.End = {etask.End}")
            print(f"  etask.GetUserStart() = {etask.GetUserStart()}")
            print(f"  etask.GetUserEnd() = {etask.GetUserEnd()}")
            print(f"  bar.Start = {new_bar.Start}")
            print(f"  bar.End = {new_bar.End}")
        except Exception as e:
            print(f"  Read error: {e}")

        # Try reschedule to see if dates stick
        print("\n--- After Reschedule ---")
        try:
            project.Reschedule(pywintypes.Time(datetime(2026, 2, 28)))
            try:
                project.WaitForNotificationProcessing()
            except:
                pass
            print(f"  etask.Start = {etask.Start}")
            print(f"  etask.End = {etask.End}")
            print(f"  bar.Start = {new_bar.Start}")
            print(f"  bar.End = {new_bar.End}")
        except Exception as e:
            print(f"  Reschedule error: {e}")

        # Now try ImposedStart/End
        print("\n--- Setting via ImposedStart/End ---")
        project.StartTransaction("Test Imposed")
        try:
            etask.ImposedStart = pywintypes.Time(datetime(2026, 8, 1))
            etask.ImposedEnd = pywintypes.Time(datetime(2026, 8, 20))
            print("  ImposedStart/End set => SUCCESS")
            project.EndTransaction()
            try:
                project.WaitForNotificationProcessing()
            except:
                pass
            print(f"  After commit: Start={etask.Start}, End={etask.End}")

            # Reschedule again
            project.Reschedule(pywintypes.Time(datetime(2026, 2, 28)))
            try:
                project.WaitForNotificationProcessing()
            except:
                pass
            print(f"  After reschedule: Start={etask.Start}, End={etask.End}")
        except Exception as e:
            print(f"  Error: {e}")
            try:
                project.AbandonTransaction()
            except:
                pass

        # Clean up - delete the test bar
        print("\n--- Cleanup ---")
        project.StartTransaction("Delete test")
        try:
            bars.Remove(bar_id)
            print(f"  bars.Remove({bar_id}) => SUCCESS!")
            project.EndTransaction()
        except Exception as e1:
            print(f"  bars.Remove({bar_id}) => {e1}")
            # Try other approaches
            try:
                # Try Remove with index
                for i in range(1, bars.Count + 1):
                    b = bars.Item(i)
                    if b.ID == bar_id:
                        bars.Remove(i)
                        print(f"  bars.Remove(index={i}) => SUCCESS!")
                        break
            except Exception as e2:
                print(f"  Remove by index => {e2}")
            try:
                project.EndTransaction()
            except:
                try:
                    project.AbandonTransaction()
                except:
                    pass

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


if __name__ == "__main__":
    print("Asta COM Explorer v4")
    print("="*80)
    try:
        app, project = connect()
        test_direct_id_access(project)
        test_view_access(project, app)
        test_date_setting_verified(project)
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
