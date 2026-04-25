"""Test the new resource assignment features via COM:
1. Rate assignment to permanent allocation
2. Effort/work/allocation setting
3. Consumable CostPerUnit/quantity on allocation
4. Task work properties
5. list_rates action
"""
import sys, os, json, traceback

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_new_features_output.txt")
f = open(OUT, "w", encoding="utf-8")
def log(msg=""): f.write(str(msg) + "\n"); f.flush()

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    # Test 1: list_rates action
    log("=== TEST 1: list_rates action ===")
    from asta_mcp_core import asta_manage_resources, ManageResourcesInput
    params = ManageResourcesInput(action="list_rates")
    result_str = asta_manage_resources(params)
    result = json.loads(result_str)
    log(f"  Full result: {json.dumps(result, indent=2, default=str)}")
    rates = result.get("cost_and_income_rates", [])
    log(f"  rates count: {len(rates)}")
    for r in rates:
        log(f"    {r['name']}: ${r['amount']}/{r['time_unit']} ({r['type']}) -> {r['cost_centre']}")
    log(f"  TEST 1: {'PASS' if result.get('success') and len(rates) > 0 else 'FAIL'}")

    # Test 2: list action (regression check)
    log("\n=== TEST 2: list action (regression) ===")
    params2 = ManageResourcesInput(action="list")
    result_str2 = asta_manage_resources(params2)
    result2 = json.loads(result_str2)
    log(f"  success: {result2.get('success')}")
    log(f"  permanent: {len(result2.get('permanent_resources', []))}")
    log(f"  consumable: {len(result2.get('consumable_resources', []))}")
    log(f"  cost_centres: {len(result2.get('cost_centres', []))}")
    log(f"  TEST 2: {'PASS' if result2.get('success') else 'FAIL'}")

    # Test 3: Assign permanent resource with rate
    log("\n=== TEST 3: Assign permanent resource with rate ===")
    from asta_mcp_core import asta_assign_resource_model, ResourceAssignmentInput, _com_get_all_bars

    import pythoncom, win32com.client
    pythoncom.CoInitialize()

    # Find a task without resource assignments
    from asta_mcp_core import _connect_asta_com
    app, project, method = _connect_asta_com()
    all_bars = _com_get_all_bars(project, max_bars=200)

    # Use the first leaf task
    test_bar = None
    for bar in all_bars:
        try:
            task_obj = bar.Tasks(1)
            try:
                if task_obj.ChildBars.Count > 0:
                    continue
            except:
                pass
            test_bar = bar
            break
        except:
            continue

    if test_bar:
        log(f"  Using task: {test_bar.Name} (ID={test_bar.ID})")

        # Check permanent resources
        perm_res = project.PermanentResources
        perm_name = perm_res.Item(1).Name if perm_res.Count > 0 else None
        log(f"  Using resource: {perm_name}")

        # Check rates
        rates_coll = project.CostAndIncomeRates
        rate_name = rates_coll.Item(1).Name if rates_coll.Count > 0 else None
        log(f"  Using rate: {rate_name}")

        if perm_name and rate_name:
            params3 = ResourceAssignmentInput(assignments=[{
                "task_id": test_bar.ID,
                "resource_name": perm_name,
                "resource_type": "permanent",
                "rate_name": rate_name,
                "effort_hours": 50.0,
                "given_allocation": 2.0,
            }])
            result_str3 = asta_assign_resource_model(params3)
            result3 = json.loads(result_str3)
            log(f"  success: {result3.get('success')}")
            if result3.get("assigned"):
                for a in result3["assigned"]:
                    log(f"    assigned: {a.get('assigned')}")
                    log(f"    rate_assigned: {a.get('rate_assigned')}")
                    log(f"    effort_hours_set: {a.get('effort_hours_set')}")
                    log(f"    given_allocation_set: {a.get('given_allocation_set')}")
                    log(f"    warnings: {[k for k in a if 'warning' in k]}")
            if result3.get("errors"):
                for e in result3["errors"]:
                    log(f"    ERROR: {e}")
            log(f"  TEST 3: {'PASS' if result3.get('success') and result3.get('assigned_count', 0) > 0 else 'FAIL'}")

    # Test 4: Assign consumable with quantity and cost_per_unit
    log("\n=== TEST 4: Assign consumable with quantity + CostPerUnit ===")
    cons_res = project.ConsumableResources
    cons_name = cons_res.Item(1).Name if cons_res.Count > 0 else None

    if cons_name and test_bar:
        log(f"  Using consumable: {cons_name}")
        params4 = ResourceAssignmentInput(assignments=[{
            "task_id": test_bar.ID,
            "resource_name": cons_name,
            "resource_type": "consumable",
            "quantity": 25.0,
            "cost_per_unit": 500.0,
        }])
        result_str4 = asta_assign_resource_model(params4)
        result4 = json.loads(result_str4)
        log(f"  success: {result4.get('success')}")
        if result4.get("assigned"):
            for a in result4["assigned"]:
                log(f"    assigned: {a.get('assigned')}")
                log(f"    quantity_set: {a.get('quantity_set')}")
                log(f"    cost_per_unit_set: {a.get('cost_per_unit_set')}")
        log(f"  TEST 4: {'PASS' if result4.get('success') and result4.get('assigned_count', 0) > 0 else 'FAIL'}")

    # Test 5: Cost centre with cost_value (bug fix test)
    log("\n=== TEST 5: Cost centre assignment with cost_value (bug fix) ===")
    ccs = project.CostCentres
    cc_name = None
    for i in range(1, ccs.Count + 1):
        cc = ccs.Item(i)
        if cc.Name not in ("Project Costs", "Terminal Genel Butcesi"):
            cc_name = cc.Name
            break

    if cc_name and test_bar:
        log(f"  Using cost centre: {cc_name}")
        params5 = ResourceAssignmentInput(assignments=[{
            "task_id": test_bar.ID,
            "resource_name": cc_name,
            "resource_type": "cost_centre",
            "cost_value": 75000.0,
        }])
        result_str5 = asta_assign_resource_model(params5)
        result5 = json.loads(result_str5)
        log(f"  success: {result5.get('success')}")
        if result5.get("assigned"):
            for a in result5["assigned"]:
                log(f"    assigned: {a.get('assigned')}")
                log(f"    cost_value_set: {a.get('cost_value_set')}")
                log(f"    warnings: {[k for k in a if 'warning' in k]}")
        if result5.get("errors"):
            for e in result5["errors"]:
                log(f"    ERROR: {e}")
        log(f"  TEST 5: {'PASS' if result5.get('success') and result5.get('assigned_count', 0) > 0 else 'FAIL'}")

    # Test 6: Task work
    log("\n=== TEST 6: Task work rate + work quantity ===")
    if test_bar:
        params6 = ResourceAssignmentInput(assignments=[{
            "task_id": test_bar.ID,
            "resource_name": cc_name or "A-Iscilik Butcesi",
            "resource_type": "cost_centre",
            "cost_value": 1000.0,
            "task_work_rate": 3.5,
            "task_work": 50.0,
        }])
        result_str6 = asta_assign_resource_model(params6)
        result6 = json.loads(result_str6)
        log(f"  success: {result6.get('success')}")
        if result6.get("assigned"):
            for a in result6["assigned"]:
                log(f"    task_work_rate_set: {a.get('task_work_rate_set')}")
                log(f"    task_work_set: {a.get('task_work_set')}")
        log(f"  TEST 6: {'PASS' if result6.get('success') else 'FAIL'}")

    log("\n=== ALL TESTS COMPLETE ===")

except Exception as e:
    log(f"FATAL: {e}")
    log(traceback.format_exc())

f.close()
print(f"Output: {OUT}")
