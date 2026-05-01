"""Phase 5d - Pure-Python Primavera P6 XER reader.

XER format (text, typically UTF-16-LE with BOM, fallback UTF-8):
- ERMHDR <version>\\t<date>\\t<user>\\t<app>\\t<currency>
- %T <table_name>            : table marker
- %F <header1>\\t<header2>... : field names (column headers)
- %R <val1>\\t<val2>...       : data row (position-mapped to %F)
- %E                         : end of file

NO mpxj dependency. Tractable in ~400 lines pure Python.
"""
import logging
import os

logger = logging.getLogger(__name__)


class XerFile:
    """Parse a P6 XER file into structured table dicts.

    Public attributes:
        file_path: original file path string.
        header_fields: dict of ERMHDR positional fields (version/exported/user/app/currency).
        tables: dict {table_name: {"headers": [str], "rows": [{col: str}]}}.

    Public read methods (added below class body):
        read_tasks() -> [task dicts in MSP shape]
        read_links() -> [link dicts {from_id, to_id, type, lag_days}]
        read_resources() -> [resource dicts]
        read_assignments() -> [assignment dicts]
        read_calendars() -> [calendar dicts]
        read_progress() -> {status_date, tasks: [...]}
        read_project() -> {proj_id, plan_start_date, plan_end_date, ...}
    """

    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"XER file not found: {file_path}")
        self.file_path = file_path
        self.header_fields = {}
        self.tables = {}
        self._parse()

    def _read_text(self):
        """Read file with encoding auto-detect (UTF-16-LE BOM or UTF-8)."""
        with open(self.file_path, "rb") as f:
            raw = f.read()
        if raw[:2] == b"\xff\xfe":
            return raw[2:].decode("utf-16-le", errors="replace")
        if raw[:3] == b"\xef\xbb\xbf":
            return raw[3:].decode("utf-8", errors="replace")
        # No BOM - try UTF-16-LE first (P6 default), fallback UTF-8
        try:
            return raw.decode("utf-16-le")
        except UnicodeDecodeError:
            return raw.decode("utf-8", errors="replace")

    def _parse(self):
        text = self._read_text()
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        current_table = None
        for line in lines:
            if not line:
                continue
            if line.startswith("ERMHDR"):
                parts = line.split("\t")
                self.header_fields = {
                    "version": parts[1] if len(parts) > 1 else "",
                    "exported": parts[2] if len(parts) > 2 else "",
                    "user": parts[3] if len(parts) > 3 else "",
                    "app": parts[4] if len(parts) > 4 else "",
                    "currency": parts[5] if len(parts) > 5 else "",
                }
                continue
            if line.startswith("%T"):
                parts = line.split("\t", 1)
                if len(parts) == 2:
                    current_table = parts[1].strip()
                    self.tables[current_table] = {"headers": [], "rows": []}
                continue
            if line.startswith("%F"):
                if current_table is None:
                    continue
                parts = line.split("\t")
                self.tables[current_table]["headers"] = [p.strip() for p in parts[1:]]
                continue
            if line.startswith("%R"):
                if current_table is None:
                    continue
                headers = self.tables[current_table]["headers"]
                if not headers:
                    continue
                parts = line.split("\t")
                values = parts[1:]
                # Pad/truncate to match header count
                if len(values) < len(headers):
                    values = values + [""] * (len(headers) - len(values))
                row = {h: values[i] for i, h in enumerate(headers)}
                self.tables[current_table]["rows"].append(row)
                continue
            if line.startswith("%E"):
                break
            # Unknown marker - skip silently (forward-compat with new P6 markers)
