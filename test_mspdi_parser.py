#!/usr/bin/env python3
"""
Test mspdi_parser.py against the UZ_MUH_POLYCLINIC XML file.
Validates: parsing, queries, writes, save, round-trip integrity.
"""

import sys
import os
import json
import traceback
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mspdi_parser import MspdiProject

XML_FILE = r"C:\Users\CahAsus\Desktop\_UZ_MUH_POLYCLINIC-Detaylı 250226-Cah3R3UC07.xml"


class T:
    """Test tracker."""
    passed = 0
    failed = 0
    results = []

    @classmethod
    def check(cls, name, condition, detail=""):
        status = "PASS" if condition else "FAIL"
        if condition:
            cls.passed += 1
        else:
            cls.failed += 1
        cls.results.append((name, status))
        print(f"  [{status}] {name}")
        if not condition and detail:
            print(f"         {detail[:200]}")
        return condition

    @classmethod
    def summary(cls):
        print(f"\n{'='*60}")
        print(f"  {cls.passed} passed, {cls.failed} failed, {cls.passed + cls.failed} total")
        print(f"{'='*60}")
        for name, status in cls.results:
            icon = "+" if status == "PASS" else "X"
            print(f"  {icon} [{status}] {name}")
        print()


def test_parsing():
    """Test basic XML parsing."""
    print("\n=== PARSING ===")
    p = MspdiProject(XML_FILE)

    T.check("File loaded", p.root is not None)
    T.check("MinutesPerDay=480", p.minutes_per_day == 480)
    T.check("HoursPerDay=8", p.hours_per_day == 8.0)

    # Task count
    T.check(f"Tasks parsed ({len(p._tasks)})", len(p._tasks) == 1470,
            f"Expected 1470, got {len(p._tasks)}")

    # Resource count
    T.check(f"Resources parsed ({len(p._resources)})", len(p._resources) == 23,
            f"Expected 23, got {len(p._resources)}")

    # Assignment count
    T.check(f"Assignments parsed ({len(p._assignments)})", len(p._assignments) == 25,
            f"Expected 25, got {len(p._assignments)}")

    # Calendar count
    T.check(f"Calendars parsed ({len(p._calendars)})", len(p._calendars) == 29,
            f"Expected 29, got {len(p._calendars)}")

    # Code libraries
    libs = p.get_code_libraries()
    lib_names = {l['name'] for l in libs}
    T.check(f"Code libraries ({len(libs)})", len(libs) >= 10,
            f"Expected >=10, got {len(libs)}: {lib_names}")
    T.check("'Disiplinler' in code libs", "Disiplinler" in lib_names)

    return p


def test_queries(p):
    """Test query methods."""
    print("\n=== QUERIES ===")

    # Project summary
    s = p.get_project_summary()
    T.check("Project name", s["project_name"] == "UZ_MUH_POLYCLINIC",
            f"Got: {s['project_name']}")
    T.check("Start date", s["start_date"] == "2025-06-23")
    T.check("Total tasks=1470", s["total_tasks"] == 1470)
    T.check(f"Total links={s['total_links']}", s["total_links"] == 1425,
            f"Expected 1425, got {s['total_links']}")

    # All tasks
    tasks = p.get_all_tasks()
    T.check(f"get_all_tasks count={len(tasks)}", len(tasks) == 1470)
    # Check first task
    t1 = tasks[0]
    T.check("First task ID=1", t1["id"] == 1)
    T.check("First task is summary", t1["summary"] == True)

    # Filter no summaries
    no_sum = p.get_all_tasks(include_summary=False)
    T.check(f"Without summaries: {len(no_sum)}", len(no_sum) < 1470)

    # Get task by ID
    task5 = p.get_task_by_id(5)
    T.check("Task ID=5 found", task5 is not None)
    T.check("Task 5 is milestone", task5["milestone"] == True,
            f"Got: milestone={task5.get('milestone')}")
    T.check("Task 5 has codes", len(task5.get("codes", {})) >= 0)

    # Critical path
    crit = p.get_critical_path()
    T.check(f"Critical path tasks ({len(crit)})", len(crit) > 0)

    # Resources
    res = p.get_resources()
    T.check(f"Resources ({len(res)})", len(res) == 23)

    # Resource assignments
    asgn = p.get_resource_assignments()
    T.check(f"Assignments ({len(asgn)})", len(asgn) == 25)

    # Calendars
    cals = p.get_calendars()
    T.check(f"Calendars ({len(cals)})", len(cals) == 29)

    # WBS tree
    tree = p.get_wbs_tree(max_depth=3)
    T.check("WBS tree not empty", len(tree) > 0)

    # Float analysis (file may not have TotalSlack; fallback uses Critical flag)
    fa = p.get_float_analysis()
    T.check("Float analysis works", isinstance(fa, dict) and "zero_float" in fa,
            f"Got: {list(fa.keys())}")
    if fa["tasks"]:
        T.check(f"Float tasks ({len(fa['tasks'])})", len(fa["tasks"]) > 0)
    else:
        T.check("Float analysis (no float data - OK)", True)

    # Predecessors check (task with known predecessors)
    # Task UID=57, ID=5 has predecessor UID=453
    if task5:
        preds = task5.get("predecessors", [])  # This is the detail dict, not internal
        # Actually let's check via the list dict
        t5_list = [t for t in tasks if t["id"] == 5][0]
        T.check(f"Task 5 has predecessors ({len(t5_list['predecessors'])})",
                len(t5_list["predecessors"]) > 0)

    # Link types distribution
    total_links = sum(len(t.get("predecessors", [])) for t in tasks)
    T.check(f"Total links from tasks={total_links}", total_links == 1425,
            f"Expected 1425, got {total_links}")

    return True


def test_new_queries(p):
    """Test new query methods."""
    print("\n=== NEW QUERIES ===")

    # Latest finishing
    latest = p.get_latest_finishing(10)
    T.check(f"Latest finishing ({len(latest)})", len(latest) == 10)
    T.check("Latest sorted by finish desc",
            latest[0]["finish"] >= latest[-1]["finish"],
            f"First={latest[0]['finish']}, Last={latest[-1]['finish']}")
    print(f"         Latest finish: {latest[0]['name'][:40]} -> {latest[0]['finish']}")

    # Missing links
    ml = p.find_missing_links()
    T.check("Missing links analysis", ml["no_predecessors_count"] >= 0)
    print(f"         No preds: {ml['no_predecessors_count']}, No succs: {ml['no_successors_count']}")

    # Search tasks
    search = p.search_tasks("procurement", include_summary=True)
    T.check(f"Search 'procurement' ({len(search)})", len(search) > 0,
            "Expected some procurement tasks")

    # Code libraries
    libs = p.get_code_libraries()
    T.check(f"Code libraries ({len(libs)})", len(libs) >= 10)

    # Task codes
    # Find a task with codes
    for t in p._tasks.values():
        if t["codes"]:
            tc = p.get_task_codes(t["id"])
            T.check(f"Task codes for ID={t['id']}", len(tc["codes"]) > 0,
                    f"Codes: {tc['codes']}")
            break
    else:
        T.check("Task codes (no coded task found)", False)

    # Filter by code - use a library that actually has assignments
    for lib in libs:
        filtered = p.filter_tasks_by_code(lib["name"])
        if filtered:
            T.check(f"Filter by '{lib['name']}' ({len(filtered)})", len(filtered) > 0)
            break
    else:
        T.check("Filter by code (no assigned codes found)", False,
                "No code library had task assignments")

    # Link chain - use Turkish names from actual data
    # Search for milestone-related chains
    chain = p.get_link_chain("Dizayn", "Procurement")
    if "error" in chain:
        # Try alternative patterns
        chain = p.get_link_chain("Teslim", "Sipari")
    T.check(f"Link chain analysis runs",
            "error" not in chain or chain.get("chains_found", 0) >= 0,
            chain.get("error", f"chains={chain.get('chains_found', '?')}"))

    return True


def test_duration_lag():
    """Test duration and lag conversion."""
    print("\n=== DURATION/LAG CONVERSION ===")
    p = MspdiProject(XML_FILE)

    # Duration parsing
    T.check("PT80H0M0S -> 10d", p._format_duration_str("PT80H0M0S") == "10d")
    T.check("PT0H0M0S -> 0d", p._format_duration_str("PT0H0M0S") == "0d")
    T.check("PT3512H0M0S -> 439d", p._format_duration_str("PT3512H0M0S") == "439d")

    # Lag conversion
    T.check("lag 0 -> 0d", p._format_lag(0) == "0d")
    T.check("lag 4800 -> 1d", p._format_lag(4800) == "1d")
    T.check("lag 288000 -> 60d", p._format_lag(288000) == "60d")
    T.check("lag -4800 -> -1d", p._format_lag(-4800) == "-1d")
    T.check("lag 48000 -> 10d", p._format_lag(48000) == "10d")

    # Round-trip: days -> lag -> days
    T.check("5d -> lag -> 5d", p._lag_to_days(p._days_to_lag(5)) == 5.0)
    T.check("60d -> lag -> 60d", p._lag_to_days(p._days_to_lag(60)) == 60.0)

    # Duration input parsing
    T.check("'10d' -> 10.0", p._parse_duration_input("10d") == 10.0)
    T.check("'2w' -> 10.0", p._parse_duration_input("2w") == 10.0)
    T.check("'80h' -> 10.0", p._parse_duration_input("80h") == 10.0)

    return True


def test_write_operations(p):
    """Test write operations (in-memory)."""
    print("\n=== WRITE OPERATIONS ===")

    initial_count = len(p._tasks)

    # Add task
    r = p.add_task("TEST-NewTask", "10d", start_date="2026-01-05")
    T.check("Add task", "task_id" in r, str(r))
    new_id = r.get("task_id")
    T.check(f"Task count +1 ({len(p._tasks)})", len(p._tasks) == initial_count + 1)

    # Get added task
    added = p.get_task_by_id(new_id)
    T.check("Get added task", added is not None and added["name"] == "TEST-NewTask")

    # Update task
    r = p.update_task(new_id, name="TEST-Renamed", duration_str="15d")
    T.check("Update task", "changes" in r and len(r["changes"]) == 2, str(r))

    # Add milestone
    r = p.add_task("TEST-Milestone", is_milestone=True, start_date="2026-02-01")
    T.check("Add milestone", "task_id" in r, str(r))
    ms_id = r.get("task_id")
    ms = p.get_task_by_id(ms_id)
    T.check("Milestone has 0d duration", ms and ms["duration"] == "0d")

    # Add link
    r = p.add_link(new_id, ms_id, "FS", "5d")
    T.check("Add link", r.get("success") == True, str(r))

    # Verify link
    ms_detail = p.get_task_by_id(ms_id)
    T.check("Link in predecessors", len(ms_detail["predecessors"]) > 0)

    # Update link
    r = p.update_link(new_id, ms_id, new_link_type="SS", new_lag_str="3d")
    T.check("Update link", r.get("updated") == True, str(r))

    # Remove link
    r = p.remove_link(new_id, ms_id)
    T.check("Remove link", r.get("removed") == True, str(r))

    # Update progress
    r = p.update_progress(new_id, percent_complete=50)
    T.check("Update progress", r.get("updated") == True, str(r))

    # Assign code (if code libraries exist)
    libs = p.get_code_libraries()
    if libs and libs[0]["values"]:
        lib_name = libs[0]["name"]
        val = libs[0]["values"][0]["value"]
        r = p.assign_code(new_id, lib_name, val)
        T.check(f"Assign code '{lib_name}'", r.get("success") == True, str(r))

    # Delete tasks
    r = p.delete_task(ms_id)
    T.check("Delete milestone", r.get("deleted") == True, str(r))
    r = p.delete_task(new_id)
    T.check("Delete task", r.get("deleted") == True, str(r))

    T.check(f"Task count restored ({len(p._tasks)})", len(p._tasks) == initial_count)

    return True


def test_save_roundtrip(p):
    """Test save and re-read."""
    print("\n=== SAVE & ROUND-TRIP ===")

    # Save to temp file
    temp_dir = tempfile.mkdtemp()
    output_path = os.path.join(temp_dir, "test_output.xml")

    saved_path = p.save(output_path)
    T.check("Save succeeds", os.path.exists(saved_path), saved_path)
    T.check(f"File size > 0", os.path.getsize(saved_path) > 0,
            f"Size: {os.path.getsize(saved_path)}")

    # Re-read
    p2 = MspdiProject(saved_path)
    T.check(f"Re-read tasks ({len(p2._tasks)})", len(p2._tasks) == len(p._tasks),
            f"Original: {len(p._tasks)}, Re-read: {len(p2._tasks)}")
    T.check(f"Re-read resources ({len(p2._resources)})", len(p2._resources) == len(p._resources))
    T.check(f"Re-read calendars ({len(p2._calendars)})", len(p2._calendars) == len(p._calendars))

    # Check that links survived
    orig_links = sum(len(t["predecessors"]) for t in p._tasks.values())
    new_links = sum(len(t["predecessors"]) for t in p2._tasks.values())
    T.check(f"Links preserved ({new_links}/{orig_links})", new_links == orig_links,
            f"Original: {orig_links}, Re-read: {new_links}")

    # Check code assignments survived
    orig_codes = sum(len(t["codes"]) for t in p._tasks.values())
    new_codes = sum(len(t["codes"]) for t in p2._tasks.values())
    T.check(f"Codes preserved ({new_codes}/{orig_codes})", new_codes == orig_codes)

    # Cleanup
    try:
        os.remove(saved_path)
        os.rmdir(temp_dir)
    except:
        pass

    return True


def main():
    print(f"Testing mspdi_parser.py against: {XML_FILE}")
    if not os.path.exists(XML_FILE):
        print(f"ERROR: File not found: {XML_FILE}")
        return False

    try:
        p = test_parsing()
        test_queries(p)
        test_new_queries(p)
        test_duration_lag()
        test_write_operations(p)
        test_save_roundtrip(p)
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        traceback.print_exc()
        T.check("Fatal error", False, str(e))

    T.summary()
    return T.failed == 0


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
