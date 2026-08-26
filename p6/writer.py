"""Write an XER out of a P6 project -- the export P6 will not give us headlessly.

P6's own headless export is `JT_XERExport`, and it is blocked: the job reaches
P6's export code and fails with "File name not specified.", with no documented
place to put the name (see docs/P6_HANDOFF.md). The CLI action script carries
only `import*` elements. So the export path is written here instead.

Two decisions matter and both are already-paid-for lessons:

* **UTF-16-LE.** cp1252 destroys Cyrillic and our own reader could not take
  UTF-8 before; UTF-16-LE with a BOM is the one encoding proven to survive the
  round trip in both directions.
* **The rows go out exactly as the database holds them.** No re-derivation, no
  re-formatting of numbers or dates beyond what the XER grammar demands. An
  exporter that "helpfully" recomputes a value is an exporter that can be
  silently wrong, and the whole point of this file is to produce something a
  parity check can trust.

Table order follows P6's own export: lookup tables first, then the project,
then WBS, resources and finally activities -- an importer needs the parents
before the children.
"""
from __future__ import annotations

import datetime as _dt
import os
from typing import Any, Iterable, Mapping, Sequence

from . import db as p6db

# P6's export order. Tables absent from the database are skipped silently;
# tables the project simply has no rows for are written with a header only,
# exactly as P6 does.
TABLE_ORDER = (
    "CURRTYPE", "NONWORK", "OBS", "RCATTYPE", "RISKTYPE", "UDFTYPE", "RCATVAL",
    "PROJECT", "CALENDAR", "SCHEDOPTIONS", "PROJWBS", "RSRC", "RSRCRATE",
    "RSRCRCAT", "TASK", "TASKPRED", "TASKRSRC", "UDFVALUE",
)

# Project-scoped tables: filtered to the exported project. Everything else is
# global and goes out whole.
PROJECT_SCOPED = {"PROJECT", "PROJWBS", "TASK", "TASKPRED", "TASKRSRC",
                  "PROJPROP", "SCHEDOPTIONS", "UDFVALUE"}

# The XER section name and the database table name are not always the same:
# scheduling options live in PROJPROP in the schema but travel as
# SCHEDOPTIONS in the file.
XER_TO_DB_TABLE = {"SCHEDOPTIONS": "PROJPROP"}

# Columns P6 does not put in an XER: they are local bookkeeping and an
# importer would reject or misread them.
SKIP_COLUMNS = {"delete_session_id", "delete_date"}

HEADER_VERSION = "19.12"
HEADER_APP = "P6 MCP"


class WriterError(RuntimeError):
    """Export refused."""


def _fmt(value: Any) -> str:
    """One database value as XER text.

    Tabs and newlines would break the row grammar, so they are collapsed --
    silently dropping a field would be worse than the space it becomes.
    """
    if value is None:
        return ""
    if isinstance(value, _dt.datetime):
        # P6 writes midnight without the time part, everything else with it.
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.strftime("%Y-%m-%d %H:%M")
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, _dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, bool):
        return "Y" if value else "N"
    if isinstance(value, float) or type(value).__name__ == "Decimal":
        # str(Decimal("0E-8")) is "0E-8" -- scientific notation, which an XER
        # importer has no reason to understand. SQL Server hands back exactly
        # that for a zeroed numeric column (RSRCRATE.cost_per_qty is one), so
        # format explicitly and only then trim the padding zeros.
        text = format(value, "f")
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        if text in ("", "-", "-0"):
            text = "0"
        return text
    text = str(value)
    return text.replace("\t", " ").replace("\r", " ").replace("\n", " ")


def _header(currency: str, user: str, database: str) -> str:
    today = _dt.date.today().strftime("%Y-%m-%d")
    # Eight fields, currency last -- see xer_parser._parse_ermhdr.
    return "\t".join(("ERMHDR", HEADER_VERSION, today, "Project", user, user,
                      database, HEADER_APP, currency or "USD"))


def _table_columns(backend, table: str) -> list[str]:
    return [c for c in backend.columns(table) if c.lower() not in SKIP_COLUMNS]


def _rows(backend, table: str, cols: Sequence[str], proj_id: int | None
          ) -> list[Sequence[Any]]:
    qo, qc = backend.quote_open, backend.quote_close
    sql = ("SELECT " + ", ".join(qo + c + qc for c in cols)
           + " FROM " + qo + table + qc)
    where, args = [], []
    upper = {c.upper() for c in backend.columns(table)}
    if proj_id is not None and table in PROJECT_SCOPED and "PROJ_ID" in upper:
        where.append(qo + "PROJ_ID" + qc + " = " + backend.param)
        args.append(proj_id)
    if "DELETE_SESSION_ID" in upper:
        where.append(qo + "DELETE_SESSION_ID" + qc + " IS NULL")
    if where:
        sql += " WHERE " + " AND ".join(where)
    return backend.select(sql, args)


def write_xer(params: Mapping[str, Any]) -> dict[str, Any]:
    """Export one project to an XER file.

    Returns the counts actually written, per table, so the caller can check
    them against the source instead of trusting the writer.
    """
    path = params.get("path") or params.get("out")
    if not path:
        raise WriterError("Cikti yolu zorunlu ('path').")
    proj_id = params.get("proj_id")
    if proj_id is None:
        raise WriterError("proj_id zorunlu.")
    proj_id = int(proj_id)
    if os.path.exists(path) and not params.get("overwrite"):
        raise WriterError(
            "Dosya zaten var: %s. Uzerine yazmak icin overwrite=true verin."
            % path)

    alias = p6db.resolve_alias(params.get("alias"))
    backend, info = p6db.open_backend(alias, use_snapshot=False)
    try:
        name = backend.select(
            "SELECT " + backend.quote_open + "proj_short_name"
            + backend.quote_close + " FROM " + backend.quote_open + "PROJECT"
            + backend.quote_close + " WHERE " + backend.quote_open + "proj_id"
            + backend.quote_close + " = " + backend.param, [proj_id])
        if not name:
            raise WriterError("Proje bulunamadi: %s" % proj_id)
        short_name = name[0][0]

        currency = ""
        try:
            cur_rows = backend.select(
                "SELECT TOP 1 " + backend.quote_open + "curr_short_name"
                + backend.quote_close + " FROM " + backend.quote_open
                + "CURRTYPE" + backend.quote_close + " ORDER BY "
                + backend.quote_open + "curr_id" + backend.quote_close, [])
            currency = cur_rows[0][0] if cur_rows else ""
        except Exception:  # noqa: BLE001
            currency = ""

        tables = params.get("tables") or TABLE_ORDER
        lines: list[str] = [_header(currency, params.get("audit_user", "MCP"),
                                    alias.database or "")]
        counts: dict[str, int] = {}
        skipped: list[str] = []
        for table in tables:
            db_table = XER_TO_DB_TABLE.get(table, table)
            cols = _table_columns(backend, db_table)
            if not cols:
                skipped.append(table)
                continue
            rows = _rows(backend, db_table, cols, proj_id)
            lines.append("%T\t" + table)
            lines.append("%F\t" + "\t".join(cols))
            for row in rows:
                lines.append("%R\t" + "\t".join(_fmt(v) for v in row))
            counts[table] = len(rows)
        lines.append("%E")
    finally:
        backend.close()

    text = "\r\n".join(lines) + "\r\n"
    tmp = path + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(b"\xff\xfe")                      # UTF-16-LE BOM
        fh.write(text.encode("utf-16-le"))
    os.replace(tmp, path)

    return {
        "path": os.path.abspath(path),
        "bytes": os.path.getsize(path),
        "encoding": "utf-16-le (BOM)",
        "proj_id": proj_id, "proj_short_name": short_name,
        "currency": currency,
        "tables": counts,
        "tables_not_in_database": skipped,
        "source": {"alias": alias.name, "database": alias.database,
                   "driver": info.get("driver")},
    }


def verify_roundtrip(path: str, expected: Mapping[str, int],
                     encoding: str | None = None) -> dict[str, Any]:
    """Read the file back with our own parser and compare the row counts.

    An exporter that reports success without reading its own output back is
    just a hopeful writer.
    """
    import xer_parser

    x = xer_parser.XerFile(path, encoding=encoding)
    got = {t: len(x.tables[t]["rows"]) for t in x.tables}
    mismatch = {t: {"written": expected[t], "read_back": got.get(t, 0)}
                for t in expected if got.get(t, 0) != expected[t]}

    # Row counts alone are not enough: a duplicated column list produced a
    # malformed %F header that still parsed, because the reader keys rows by
    # column name and the duplicates collapsed. Check the headers too.
    dup_headers = {t: sorted({c for c in x.tables[t]["headers"]
                              if x.tables[t]["headers"].count(c) > 1})
                   for t in x.tables
                   if len(set(x.tables[t]["headers"])) != len(x.tables[t]["headers"])}
    bad_text = sum(
        1 for t in x.tables for r in x.tables[t]["rows"]
        for v in r.values() if isinstance(v, str) and "�" in v)
    return {
        "path": path,
        "encoding_detected": x.encoding,
        "encoding_source": x.encoding_source,
        "tables_read": got,
        "row_count_match": not mismatch,
        "mismatch": mismatch,
        "duplicate_headers": dup_headers,
        "headers_unique": not dup_headers,
        "replacement_chars": bad_text,
        "ok": not mismatch and not dup_headers and not bad_text,
    }
