"""Phase 2a acceptance: build a Uzbekistan-2026 calendar end-to-end.

SAFETY: opens an isolated MS Project via FileNew, never touches the user's
active project. Closes without saving on completion.

Steps:
  1. Create 'Uzbekistan-2026' calendar from Standard
  2. Bulk-add 9 official Ozbek holidays
  3. Mark Sunday as non-working
  4. Add 1 task and assign the calendar to it
  5. List calendars and print summary
  6. Close the test project without saving
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_update, _msp_calendar_holidays_uzbek,
    _msp_calendar_assign_to_task, _msp_calendar_list,
    _msp_task_add_single,
)


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")

    # Remember user's active project
    original_name = None
    if app.ActiveProject is not None:
        original_name = app.ActiveProject.Name

    # Open isolated test project
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] Using isolated test project: {test_name}")
    print(f"[SAFE] User's project preserved: {original_name}")

    try:
        t0 = time.time()

        # 1. Create calendar
        print("\n1. Creating 'Uzbekistan-2026' calendar...")
        r = _msp_calendar_create(name="Uzbekistan-2026", base_calendar="Standard")
        assert r["status"] == "ok", r
        print(f"   OK uid={r['calendar_uid']}")

        # 2. Add 9 Uzbek holidays
        print("2. Adding 9 official Uzbek holidays...")
        r = _msp_calendar_holidays_uzbek(calendar_name="Uzbekistan-2026", year=2026)
        assert r["status"] in ("ok", "partial"), r
        print(f"   OK {r['count']} holidays added")
        for h in r["holidays"]:
            print(f"      - {h['date']}  {h['name']}")

        # 3. Sunday off (weekday=1)
        print("3. Marking Sunday as non-working...")
        r = _msp_calendar_update(name="Uzbekistan-2026", weekday_off=1)
        assert r["status"] == "ok", r
        print(f"   OK changes={r['changes']}")

        # 4. Add a task and assign the calendar
        print("4. Adding 'Hafriyat' task and assigning Uzbekistan-2026...")
        add_r = _msp_task_add_single(name="Hafriyat", duration="10d")
        assert add_r["status"] == "ok"
        r = _msp_calendar_assign_to_task(task_id=add_r["task_id"],
                                         calendar_name="Uzbekistan-2026")
        assert r["status"] == "ok", r
        print(f"   OK task_id={add_r['task_id']} -> Uzbekistan-2026")

        # 5. List
        print("5. Listing calendars...")
        r = _msp_calendar_list()
        assert r["status"] == "ok"
        for c in r["calendars"]:
            marker = "*" if c["name"] == "Uzbekistan-2026" else " "
            print(f"   {marker} {c['name']}  exceptions={c['exception_count']}")

        elapsed = time.time() - t0
        print(f"\nOK ACCEPTANCE: end-to-end in {elapsed:.2f}s (target: <5s total)")
        assert elapsed < 5.0, f"Too slow: {elapsed}s"

    finally:
        # 6. Always restore user's project (close test without saving)
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)  # 0 = pjDoNotSave
                    break
            if original_name:
                for i in range(1, app.Projects.Count + 1):
                    if app.Projects(i).Name == original_name:
                        app.WindowActivate(app.Projects(i).Windows(1).Caption)
                        break
        except Exception as e:
            print(f"[WARN] cleanup error: {e}")


if __name__ == "__main__":
    main()
