"""Phase 6.1 acceptance — currency mode cross-validation lifecycle.

Builds a CAU-style XER fixture (target_cost == target_qty pattern, RULE 3
'hours' mode, USD currency code) and a 'cost' variant, then exercises:
1. _msp_evm_validate_currency_mode (4-mode + cross-validation)
2. msproject_evm dispatcher action 'validate_currency_mode'
3. Backward-compat 'detect_currency_mode' (2-mode return)

Demonstrates that:
- CAU pattern (target_cost == target_qty) -> 'hours' (cost not loaded)
- Real cost loading (cost > qty) -> 'cost'
- ERMHDR.currency code extracted (USD)
- Cross-validation produces high/medium/low confidence
- Legacy dispatcher action unchanged

Run:
    python -m samples.build_currency_validation_lifecycle
"""
import json
import os

from samples._lifecycle_common import (
    write_synthetic_xer as _write_xer,
    call_async_dispatcher,
)
from msproject_mcp_core import (
    _msp_evm_validate_currency_mode,
    msproject_evm,
)


# RULE 3 'hours' pattern — target_cost == target_qty in every TASKRSRC row
HOURS_XER = """ERMHDR\t18.8\t2026-05-01\tcahit\tProject Management\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tCAU\t2024-07-08 08:00\t2028-06-20 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tCAU 6x9\t9.0\t54.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tConcrete Workers\tCOW\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA1010\tFoundation\tTT_Task\t180.0\t2024-07-08 08:00\t2024-07-29 17:00\t100.0
%R\t1002\t1\t1\t1\tA1020\tFrame\tTT_Task\t360.0\t2024-07-30 08:00\t2024-09-09 17:00\t75.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t180.0\t180.0\t180.0\t180.0
%R\t2\t1002\t101\t360.0\t270.0\t360.0\t270.0
%E
"""

# Real cost loading — target_cost > target_qty, varies independently
COST_XER = """ERMHDR\t18.8\t2026-05-01\tcahit\tProject Management\tEUR
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tEU\t2024-07-08 08:00\t2028-06-20 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tEU 5x8\t8.0\t40.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tWorkers\tWRK\tRT_Labor\t10.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tphys_complete_pct
%R\t1001\t1\t1\t1\tA1010\tDesign\tTT_Task\t160.0\t2024-07-08 08:00\t2024-07-29 17:00\t100.0
%R\t1002\t1\t1\t1\tA1020\tBuild\tTT_Task\t320.0\t2024-07-30 08:00\t2024-09-09 17:00\t50.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t160.0\t160.0\t8000.0\t8000.0
%R\t2\t1002\t101\t320.0\t160.0\t16000.0\t8000.0
%E
"""


def _print_validate(label: str, xer_path: str):
    print(f"\n{'=' * 70}")
    print(f"  {label}")
    print(f"  {xer_path}")
    print("=" * 70)
    r = _msp_evm_validate_currency_mode(file_path=xer_path)
    print(f"primary_mode    : {r['primary_mode']}")
    print(f"currency_code   : {r['currency_code']}")
    print(f"sources         :")
    for k, v in r["sources"].items():
        print(f"  {k:20s} = {v}")
    cv = r["cross_validation"]
    print(f"cross-validation:")
    print(f"  consensus_mode = {cv['consensus_mode']}")
    print(f"  confidence     = {cv['confidence']}")
    print(f"  conflicts      = {cv['conflicts']}")
    print(f"  warnings       = {cv['warnings']}")
    print(f"  source_counts  = {cv['source_counts']}")
    return r


def _print_dispatcher(label: str, xer_path: str, action: str):
    print(f"\n--- dispatcher: {label} ({action}) ---")
    r = call_async_dispatcher(msproject_evm, action, file_path=xer_path)
    print(json.dumps(r, indent=2, ensure_ascii=False, default=str)[:600])


def main():
    hours_path = _write_xer(HOURS_XER, "phase6_1_hours_sample.xer")
    cost_path = _write_xer(COST_XER, "phase6_1_cost_sample.xer")

    # Adapter direct calls
    hours = _print_validate("CAU 'hours' pattern (RULE 3: cost == qty)", hours_path)
    cost = _print_validate("EU 'cost' pattern (cost > qty)", cost_path)

    # Dispatcher action wiring
    _print_dispatcher("CAU hours XER", hours_path, "validate_currency_mode")
    _print_dispatcher("CAU hours XER (legacy)", hours_path, "detect_currency_mode")

    # Acceptance assertions
    print(f"\n{'=' * 70}")
    print("  ACCEPTANCE ASSERTIONS")
    print("=" * 70)
    assert hours["primary_mode"] == "hours", \
        f"CAU pattern expected 'hours', got {hours['primary_mode']}"
    assert hours["currency_code"] == "USD"
    assert hours["sources"]["xer_assignments"] == "hours"
    assert cost["primary_mode"] == "cost", \
        f"Real cost loading expected 'cost', got {cost['primary_mode']}"
    assert cost["currency_code"] == "EUR"
    assert cost["sources"]["xer_assignments"] == "cost"
    print("[PASS] CAU 'hours' pattern: primary_mode=hours, currency=USD")
    print("[PASS] EU 'cost' pattern : primary_mode=cost,  currency=EUR")
    print("[PASS] All assertions passed.")

    # Cleanup
    for p in (hours_path, cost_path):
        try:
            os.remove(p)
        except OSError:
            pass


if __name__ == "__main__":
    main()
