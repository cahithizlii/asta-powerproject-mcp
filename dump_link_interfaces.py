"""
Explore LinksIn/LinksOut COM interfaces at runtime.
Run while Asta Powerproject is open with a project that has tasks.
"""
import pythoncom
import win32com.client

def explore_dispatch(obj, name, max_depth=1):
    """Explore a COM dispatch object's methods and properties."""
    print(f"\n{'='*60}")
    print(f"Exploring: {name}")
    print(f"Type: {type(obj)}")
    print(f"{'='*60}")

    # Try common collection methods
    for attr in ["Count", "Item", "Add", "Remove", "Delete", "_NewEnum", "All",
                 "Application", "Parent", "ObjectType", "ID", "Name",
                 "PredecessorTask", "SuccessorTask", "Predecessor", "Successor",
                 "FromTask", "ToTask", "Type", "LinkType", "Lag", "LagDuration",
                 "Bar", "Task", "Start", "End", "Duration",
                 "LinkCategory", "Category", "Critical",
                 "TotalFloat", "FreeFloat", "DrivingLink"]:
        try:
            val = getattr(obj, attr)
            if callable(val) and attr not in ["Count"]:
                print(f"  [METHOD] {attr}() -> callable")
            else:
                print(f"  [PROP]   {attr} = {val}")
        except AttributeError:
            pass  # Not available
        except Exception as e:
            err_str = str(e)[:80]
            print(f"  [ERROR]  {attr} -> {err_str}")

    # Try type info
    try:
        ti = obj._oleobj_.GetTypeInfo()
        ta = ti.GetTypeAttr()
        print(f"\n  TypeInfo: {ta[0]}, funcs={ta[6]}, vars={ta[7]}")
        for i in range(ta[6]):
            try:
                fd = ti.GetFuncDesc(i)
                names = ti.GetNames(fd[0])
                invkind = fd[3]
                kind_str = {1: "METHOD", 2: "PROP_GET", 4: "PROP_PUT", 8: "PROP_PUTREF"}.get(invkind, f"UNK({invkind})")
                if names:
                    print(f"  [{kind_str}] {names[0]}({', '.join(names[1:])})")
            except Exception as e:
                print(f"  [func_{i}] ERROR: {e}")
    except Exception as e:
        print(f"  TypeInfo not available: {e}")


def main():
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("PowerProject.Application")
        project = app.ActiveProject
        print(f"Connected to: {app.Name}")
        print(f"Project: {project.Name}")

        # Get first bar with tasks
        bars = project.Bars
        print(f"\nTotal bars: {bars.Count}")

        test_bar = None
        test_task = None
        for i in range(1, min(bars.Count + 1, 20)):
            try:
                bar = bars.Item(i)
                task = bar.ExpandedTask
                if task is not None:
                    test_bar = bar
                    test_task = task
                    print(f"\nUsing bar #{i}: ID={bar.ID}, Name='{bar.Name}'")
                    break
            except Exception:
                continue

        if test_task is None:
            print("No bar with ExpandedTask found!")
            return

        # Explore LinksIn
        try:
            links_in = test_task.LinksIn
            explore_dispatch(links_in, f"task.LinksIn (task ID={test_task.ID})")

            # If there are links, explore first one
            try:
                count = links_in.Count
                print(f"\n  LinksIn count: {count}")
                if count > 0:
                    first_link = links_in.Item(1)
                    explore_dispatch(first_link, "LinksIn.Item(1) — single link object")
            except Exception as e:
                print(f"  LinksIn.Count error: {e}")
        except Exception as e:
            print(f"task.LinksIn error: {e}")

        # Explore LinksOut
        try:
            links_out = test_task.LinksOut
            explore_dispatch(links_out, f"task.LinksOut (task ID={test_task.ID})")

            try:
                count = links_out.Count
                print(f"\n  LinksOut count: {count}")
                if count > 0:
                    first_link = links_out.Item(1)
                    explore_dispatch(first_link, "LinksOut.Item(1) — single link object")
            except Exception as e:
                print(f"  LinksOut.Count error: {e}")
        except Exception as e:
            print(f"task.LinksOut error: {e}")

        # Also explore LinkCategorys
        try:
            link_cats = project.LinkCategorys
            explore_dispatch(link_cats, "project.LinkCategorys")
            try:
                count = link_cats.Count
                print(f"\n  LinkCategorys count: {count}")
                if count > 0:
                    first_cat = link_cats.Item(1)
                    explore_dispatch(first_cat, "LinkCategorys.Item(1)")
            except Exception as e:
                print(f"  Count error: {e}")
        except Exception as e:
            print(f"project.LinkCategorys error: {e}")

        # Explore AllLinkCategorys
        try:
            all_link_cats = project.AllLinkCategorys
            explore_dispatch(all_link_cats, "project.AllLinkCategorys")
        except Exception as e:
            print(f"project.AllLinkCategorys error: {e}")

    except Exception as e:
        print(f"Connection error: {e}")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
