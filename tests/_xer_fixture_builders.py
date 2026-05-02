"""Phase 10.5 — shared synthetic XER fixture helper for integration tests.

Extracted from copy-pasted `_write_xer` blocks in:
- tests/test_msproject_evm_time_phased_ac_integration.py
- tests/test_msproject_compare_dispatcher.py

Static XER content strings remain in the test files that own them
(SYNTH_EARLY_FINISH, SYNTH_STAGGERED, SNAPSHOT_A, SNAPSHOT_B) — only
the file-write boilerplate moves here.

Usage:
    from tests._xer_fixture_builders import write_synthetic_xer
    path = write_synthetic_xer(SYNTH_EARLY_FINISH, "p62_early_finish.xer")
"""
import os
import tempfile


def write_synthetic_xer(content: str, name: str) -> str:
    """Write a UTF-16-LE BOM XER fixture to the system tempdir.

    Args:
        content: full XER text (ERMHDR + table sections + %E).
        name: bare filename (no path); placed in tempfile.gettempdir().

    Returns:
        absolute path to the written file.

    The caller is responsible for deletion (typically via try/finally).
    """
    path = os.path.join(tempfile.gettempdir(), name)
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(content.encode("utf-16-le"))
    return path
