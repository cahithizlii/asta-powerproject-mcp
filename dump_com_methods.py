#!/usr/bin/env python3
"""
Dump ALL COM methods, properties, and interfaces from a running Asta Powerproject instance.
Saves the full type library to a text file for analysis.

Usage: python dump_com_methods.py
Output: asta_com_typelib.txt (in same directory)
"""

import os
import sys
import json

def dump_dispatch_object(obj, name="object", depth=0, visited=None, output_lines=None):
    """Recursively enumerate all properties and methods of a COM dispatch object."""
    if visited is None:
        visited = set()
    if output_lines is None:
        output_lines = []

    indent = "  " * depth
    obj_id = id(obj)
    if obj_id in visited or depth > 3:
        return output_lines
    visited.add(obj_id)

    # Try to get type info
    try:
        type_info = obj._oleobj_.GetTypeInfo()
        type_attr = type_info.GetTypeAttr()

        # Get the type name
        try:
            type_name = type_info.GetDocumentation(-1)[0]
        except Exception:
            type_name = name

        output_lines.append(f"\n{'='*80}")
        output_lines.append(f"{indent}INTERFACE: {type_name}")
        output_lines.append(f"{'='*80}")

        # Enumerate all functions (methods + properties)
        func_count = type_attr[6]  # cFuncs
        var_count = type_attr[7]   # cVars

        output_lines.append(f"{indent}  Functions: {func_count}, Variables: {var_count}")
        output_lines.append("")

        # Dump functions
        for i in range(func_count):
            try:
                func_desc = type_info.GetFuncDesc(i)
                func_names = type_info.GetNames(func_desc[0])
                func_name = func_names[0] if func_names else f"func_{i}"

                # Function kind
                invoke_kind = func_desc[4]
                kind_str = {1: "METHOD", 2: "PROP_GET", 4: "PROP_PUT", 8: "PROP_PUTREF"}.get(invoke_kind, f"KIND_{invoke_kind}")

                # Return type
                return_type = func_desc[8]
                type_str = _vt_to_str(return_type[0]) if return_type else "void"

                # Parameters
                param_count = func_desc[6]  # cParams
                param_names = list(func_names[1:]) if len(func_names) > 1 else []

                # Parameter types from func_desc[2] (lprgelemdescParam)
                param_types = []
                if func_desc[2]:
                    for p_idx, param in enumerate(func_desc[2]):
                        p_name = param_names[p_idx] if p_idx < len(param_names) else f"p{p_idx}"
                        p_type = _vt_to_str(param[0][0]) if param and param[0] else "variant"
                        param_types.append(f"{p_type} {p_name}")

                params_str = ", ".join(param_types) if param_types else ""

                output_lines.append(f"{indent}  [{kind_str}] {type_str} {func_name}({params_str})")

            except Exception as e:
                output_lines.append(f"{indent}  [ERROR] func_{i}: {e}")

        # Dump variables
        if var_count > 0:
            output_lines.append(f"\n{indent}  --- Variables ---")
            for i in range(var_count):
                try:
                    var_desc = type_info.GetVarDesc(i)
                    var_names = type_info.GetNames(var_desc[0])
                    var_name = var_names[0] if var_names else f"var_{i}"
                    var_type = _vt_to_str(var_desc[1][0]) if var_desc[1] else "variant"
                    output_lines.append(f"{indent}  [VAR] {var_type} {var_name}")
                except Exception as e:
                    output_lines.append(f"{indent}  [ERROR] var_{i}: {e}")

        # Dump implemented interfaces
        impl_count = type_attr[8]  # cImplTypes
        if impl_count > 0:
            output_lines.append(f"\n{indent}  --- Implemented Interfaces ({impl_count}) ---")
            for i in range(impl_count):
                try:
                    impl_ref = type_info.GetImplTypeFlags(i)
                    ref_type = type_info.GetRefTypeOfImplType(i)
                    ref_info = type_info.GetRefTypeInfo(ref_type)
                    ref_name = ref_info.GetDocumentation(-1)[0]
                    output_lines.append(f"{indent}  [IMPL] {ref_name} (flags={impl_ref})")
                except Exception as e:
                    pass

    except Exception as e:
        output_lines.append(f"{indent}[Could not get type info for {name}: {e}]")

    return output_lines


def dump_type_library(tlb_path=None):
    """Dump the entire Asta type library."""
    import pythoncom

    output_lines = []
    output_lines.append("ASTA POWERPROJECT COM TYPE LIBRARY DUMP")
    output_lines.append("=" * 80)

    # Try to load from the TeamPlan type library
    CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"

    try:
        # Method 1: Load from registered type library
        import win32com.client

        # Get the type library from the CLSID
        try:
            obj = pythoncom.GetActiveObject(CLSID)
            app = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            output_lines.append(f"Connected to running Asta instance")
        except Exception:
            app = win32com.client.dynamic.Dispatch(CLSID)
            output_lines.append(f"Created Asta dispatch object")

        # Dump the Application object
        output_lines.append(f"\n{'#'*80}")
        output_lines.append("# APPLICATION OBJECT")
        output_lines.append(f"{'#'*80}")
        dump_dispatch_object(app, "Application", 0, set(), output_lines)

        # Dump the Project object
        try:
            project = app.ActiveProject
            if project:
                output_lines.append(f"\n{'#'*80}")
                output_lines.append("# PROJECT OBJECT (ActiveProject)")
                output_lines.append(f"{'#'*80}")
                dump_dispatch_object(project, "Project", 0, set(), output_lines)

                # Dump Bars collection
                try:
                    bars = project.Bars
                    if bars:
                        output_lines.append(f"\n{'#'*80}")
                        output_lines.append("# BARS COLLECTION")
                        output_lines.append(f"{'#'*80}")
                        dump_dispatch_object(bars, "Bars", 0, set(), output_lines)

                        # Dump a single Bar
                        try:
                            bar = bars.Item(1)
                            if bar:
                                output_lines.append(f"\n{'#'*80}")
                                output_lines.append("# BAR OBJECT (single bar)")
                                output_lines.append(f"{'#'*80}")
                                dump_dispatch_object(bar, "Bar", 0, set(), output_lines)

                                # Dump ExpandedTask from bar
                                try:
                                    task = bar.ExpandedTask
                                    if task:
                                        output_lines.append(f"\n{'#'*80}")
                                        output_lines.append("# TASK OBJECT (from bar.ExpandedTask)")
                                        output_lines.append(f"{'#'*80}")
                                        dump_dispatch_object(task, "Task", 0, set(), output_lines)
                                except Exception as e:
                                    output_lines.append(f"Could not get ExpandedTask: {e}")
                        except Exception as e:
                            output_lines.append(f"Could not get bar: {e}")
                except Exception as e:
                    output_lines.append(f"Could not get Bars: {e}")

                # Dump Links if available
                try:
                    links = project.Links
                    if links:
                        output_lines.append(f"\n{'#'*80}")
                        output_lines.append("# LINKS COLLECTION")
                        output_lines.append(f"{'#'*80}")
                        dump_dispatch_object(links, "Links", 0, set(), output_lines)
                except Exception as e:
                    output_lines.append(f"project.Links not available: {e}")

                # Dump Dependencies from first bar
                try:
                    bar = project.Bars.Item(1)
                    deps = bar.Dependencies
                    if deps:
                        output_lines.append(f"\n{'#'*80}")
                        output_lines.append("# DEPENDENCIES COLLECTION (from bar)")
                        output_lines.append(f"{'#'*80}")
                        dump_dispatch_object(deps, "Dependencies", 0, set(), output_lines)
                except Exception as e:
                    output_lines.append(f"bar.Dependencies not available: {e}")

                # Dump Resources
                for res_type in ["PermanentResources", "ConsumableResources", "CostCentres"]:
                    try:
                        res_col = getattr(project, res_type)
                        if res_col:
                            output_lines.append(f"\n{'#'*80}")
                            output_lines.append(f"# {res_type.upper()} COLLECTION")
                            output_lines.append(f"{'#'*80}")
                            dump_dispatch_object(res_col, res_type, 0, set(), output_lines)

                            # Single resource
                            try:
                                if res_col.Count > 0:
                                    res = res_col.Item(1)
                                    output_lines.append(f"\n--- Single {res_type[:-1]} ---")
                                    dump_dispatch_object(res, f"{res_type}_Item", 0, set(), output_lines)
                            except Exception:
                                pass
                    except Exception as e:
                        output_lines.append(f"{res_type} not available: {e}")

                # Dump ProgressPeriods
                try:
                    pp = project.AllProgressPeriods
                    if pp:
                        output_lines.append(f"\n{'#'*80}")
                        output_lines.append("# PROGRESS PERIODS")
                        output_lines.append(f"{'#'*80}")
                        dump_dispatch_object(pp, "AllProgressPeriods", 0, set(), output_lines)
                except Exception as e:
                    output_lines.append(f"AllProgressPeriods not available: {e}")

                # Dump CurrentView
                try:
                    view = project.CurrentView
                    if view:
                        output_lines.append(f"\n{'#'*80}")
                        output_lines.append("# CURRENT VIEW")
                        output_lines.append(f"{'#'*80}")
                        dump_dispatch_object(view, "CurrentView", 0, set(), output_lines)
                except Exception as e:
                    output_lines.append(f"CurrentView not available: {e}")

        except Exception as e:
            output_lines.append(f"Could not get ActiveProject: {e}")

    except Exception as e:
        output_lines.append(f"FATAL: Could not connect to Asta: {e}")

    # Also try to dump from the type library file directly
    output_lines.append(f"\n\n{'#'*80}")
    output_lines.append("# TYPE LIBRARY FILE ANALYSIS")
    output_lines.append(f"{'#'*80}")

    try:
        import win32com.client
        # Try to find and load the .tlb file
        tlb_paths = [
            r"C:\Program Files\Elecosoft\Powerproject\teamplan.tlb",
            r"C:\Program Files (x86)\Elecosoft\Powerproject\teamplan.tlb",
            r"C:\Program Files\Asta Development\Powerproject\teamplan.tlb",
        ]

        # Also search in registry
        import winreg
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{CLSID}\\InprocServer32")
            dll_path = winreg.QueryValue(key, None)
            tlb_dir = os.path.dirname(dll_path)
            tlb_paths.insert(0, os.path.join(tlb_dir, "teamplan.tlb"))
            output_lines.append(f"DLL path from registry: {dll_path}")
        except Exception:
            pass

        for tlb_path in tlb_paths:
            if os.path.exists(tlb_path):
                output_lines.append(f"Found TLB: {tlb_path}")
                try:
                    tlb = pythoncom.LoadTypeLib(tlb_path)
                    type_count = tlb.GetTypeInfoCount()
                    output_lines.append(f"Type library has {type_count} types")

                    for i in range(type_count):
                        try:
                            ti = tlb.GetTypeInfo(i)
                            ta = ti.GetTypeAttr()
                            doc = tlb.GetDocumentation(i)
                            type_name = doc[0]
                            type_doc = doc[1] if doc[1] else ""

                            # Type kind
                            type_kind = ta[5]
                            kind_names = {0: "ENUM", 1: "RECORD", 2: "MODULE", 3: "INTERFACE",
                                        4: "DISPATCH", 5: "COCLASS", 6: "ALIAS", 7: "UNION"}
                            kind_str = kind_names.get(type_kind, f"KIND_{type_kind}")

                            output_lines.append(f"\n--- [{kind_str}] {type_name} ---")
                            if type_doc:
                                output_lines.append(f"    Doc: {type_doc}")

                            # Functions
                            func_count = ta[6]
                            for fi in range(func_count):
                                try:
                                    fd = ti.GetFuncDesc(fi)
                                    fn = ti.GetNames(fd[0])
                                    func_name = fn[0] if fn else f"func_{fi}"
                                    invoke_kind = fd[4]
                                    kind_s = {1: "METHOD", 2: "GET", 4: "PUT", 8: "PUTREF"}.get(invoke_kind, str(invoke_kind))

                                    ret_type = _vt_to_str(fd[8][0]) if fd[8] else "void"

                                    # Params
                                    param_strs = []
                                    if fd[2]:
                                        for pi, param in enumerate(fd[2]):
                                            p_name = fn[pi+1] if pi+1 < len(fn) else f"p{pi}"
                                            p_type = _vt_to_str(param[0][0]) if param and param[0] else "variant"
                                            param_strs.append(f"{p_type} {p_name}")

                                    params_s = ", ".join(param_strs)
                                    output_lines.append(f"    [{kind_s}] {ret_type} {func_name}({params_s})")
                                except Exception:
                                    pass

                            # Variables
                            var_count = ta[7]
                            for vi in range(var_count):
                                try:
                                    vd = ti.GetVarDesc(vi)
                                    vn = ti.GetNames(vd[0])
                                    var_name = vn[0] if vn else f"var_{vi}"
                                    output_lines.append(f"    [CONST] {var_name} = {vd[1]}")
                                except Exception:
                                    pass

                        except Exception as e:
                            output_lines.append(f"  Error reading type {i}: {e}")

                    break  # Found and processed TLB
                except Exception as e:
                    output_lines.append(f"Error loading TLB {tlb_path}: {e}")
        else:
            output_lines.append("No teamplan.tlb file found. Checking Asta install paths...")
            # List Asta-related files
            for search_dir in [r"C:\Program Files", r"C:\Program Files (x86)"]:
                for root, dirs, files in os.walk(search_dir):
                    for f in files:
                        if f.lower() in ("teamplan.tlb", "astadkit.ocx", "powerproject.exe"):
                            output_lines.append(f"  Found: {os.path.join(root, f)}")
    except Exception as e:
        output_lines.append(f"TLB analysis error: {e}")

    return output_lines


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
    if isinstance(vt_code, tuple):
        vt_code = vt_code[0] if vt_code else 0
    base = vt_code & 0xFFF
    is_array = bool(vt_code & 0x2000)
    is_byref = bool(vt_code & 0x4000)

    name = VT_NAMES.get(base, f"VT_{base}")
    if is_array:
        name = f"{name}[]"
    if is_byref:
        name = f"{name}&"
    return name


if __name__ == "__main__":
    import pythoncom
    pythoncom.CoInitialize()

    try:
        lines = dump_type_library()

        output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "asta_com_typelib.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

        print(f"Type library dumped to: {output_path}")
        print(f"Total lines: {len(lines)}")

        # Also print a summary
        method_count = sum(1 for l in lines if "[METHOD]" in l)
        get_count = sum(1 for l in lines if "[GET]" in l or "[PROP_GET]" in l)
        put_count = sum(1 for l in lines if "[PUT]" in l or "[PROP_PUT]" in l)
        print(f"Methods: {method_count}, Properties (get): {get_count}, Properties (put): {put_count}")

    finally:
        pythoncom.CoUninitialize()
