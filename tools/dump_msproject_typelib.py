"""Dump MS Project COM type library to a text reference document.

One-time generation. Output: ``msproject_typelib.txt`` at the repo root.

The dump is a Python module produced by ``win32com.client.makepy`` (the same
generator used by ``gencache.EnsureModule``). It contains every COM class,
interface, dispatch ID, enum, and constant exposed by MS Project's type library
and is used as a reference (not as an importable module) for subsequent MCP
implementation work.

Run from the repo root with MS Project either open or installed locally:

    python tools/dump_msproject_typelib.py

Re-runnable: overwrites the existing ``msproject_typelib.txt`` each time.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client
from win32com.client import gencache, makepy

PROGID = "MSProject.Application"
OUTPUT = Path(__file__).resolve().parent.parent / "msproject_typelib.txt"


def _connect_or_warn() -> tuple[str, str] | None:
    """Try to read MS Project version/build from a running instance.

    Not strictly required — the type library can be read even if MS Project is
    not running, as long as it is installed and registered. We just print the
    version when available so the dump can be tied to a specific build.
    """
    try:
        app = win32com.client.GetActiveObject(PROGID)
    except pythoncom.com_error:
        return None
    try:
        version = str(app.Version)
    except Exception:
        version = "unknown"
    try:
        build = str(app.Build)
    except Exception:
        build = "unknown"
    return version, build


def main() -> int:
    pythoncom.CoInitialize()

    info = _connect_or_warn()
    if info is not None:
        version, build = info
        print(f"MS Project (running): version {version}, build {build}")
    else:
        print(
            "MS Project is not currently running — reading registered type "
            "library directly. (Open MS Project once if registration fails.)"
        )

    # Make sure the gencache module is generated/up to date. This also
    # registers the type library spec so makepy can find it by ProgID.
    try:
        gencache.EnsureDispatch(PROGID)
    except pythoncom.com_error as exc:
        print(
            f"ERROR: could not dispatch {PROGID!r}. Is MS Project installed "
            f"and registered? Underlying error: {exc}",
            file=sys.stderr,
        )
        return 2

    # Resolve the registered type library spec(s) for the ProgID, then have
    # makepy generate the Python interface stub into an in-memory buffer.
    # ``GetTypeLibsForSpec`` returns a list of ``(PyITypeLib, TypelibSpec)``
    # pairs — pass the ``TypelibSpec`` to ``GenerateFromTypeLibSpec``.
    matches = makepy.GetTypeLibsForSpec(PROGID)
    if not matches:
        print(
            f"ERROR: no type library spec found for {PROGID!r}.",
            file=sys.stderr,
        )
        return 3

    _, spec = matches[0]
    # genpy asserts ``file.encoding`` is set, so ``io.StringIO`` doesn't work.
    # Write straight to the output path; makepy expects a regular text file.
    with OUTPUT.open("w", encoding="utf-8") as fh:
        makepy.GenerateFromTypeLibSpec(spec, file=fh, verboseLevel=0)

    size = OUTPUT.stat().st_size
    line_count = sum(1 for _ in OUTPUT.open("r", encoding="utf-8"))
    if size < 5_000:
        print(
            f"WARNING: generated dump is unexpectedly small ({size} bytes). "
            "The type library may not have been read correctly.",
            file=sys.stderr,
        )

    print(f"Wrote {OUTPUT} ({size:,} bytes, {line_count:,} lines)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
