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

    def __init__(self, file_path, encoding=None):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"XER file not found: {file_path}")
        self.file_path = file_path
        self.requested_encoding = encoding
        self.encoding = None
        self.encoding_source = None
        self.encoding_scores = None
        self.encoding_confidence = None
        self.header_fields = {}
        self.tables = {}
        self._parse()

    def _read_text(self):
        """Decode the file, and say which encoding was used and why.

        P6 writes XER either as UTF-16-LE or in the *exporting machine's ANSI
        code page*, and the ERMHDR header does not record which. The old
        fallback here was ``utf-8, errors="replace"``, which turned every
        Cyrillic byte of a cp1251 export into U+FFFD without a word of
        warning -- silently wrong data, the failure class RULE 0 exists to
        prevent. So: try the unambiguous decodings first, then score the ANSI
        candidates, and refuse to guess when the top two are close.

        Pass ``encoding=`` to settle it explicitly.
        """
        with open(self.file_path, "rb") as f:
            raw = f.read()

        if self.requested_encoding:
            self.encoding = self.requested_encoding
            self.encoding_source = "parametre"
            return raw.decode(self.requested_encoding, errors="replace")

        if raw[:2] == b"\xff\xfe":
            self.encoding, self.encoding_source = "utf-16-le", "BOM"
            return raw[2:].decode("utf-16-le", errors="replace")
        if raw[:3] == b"\xef\xbb\xbf":
            self.encoding, self.encoding_source = "utf-8", "BOM"
            return raw[3:].decode("utf-8", errors="replace")

        for enc in ("utf-16-le", "utf-8"):
            try:
                text = raw.decode(enc)
            except UnicodeDecodeError:
                continue
            # A UTF-16-LE guess on ASCII-ish bytes decodes into CJK noise;
            # require the tab/newline skeleton an XER must have.
            if enc == "utf-16-le" and text.count("\t") < 10:
                continue
            self.encoding, self.encoding_source = enc, "gecerli cozumleme"
            return text

        return self._decode_ansi(raw)

    # ANSI code pages P6 actually exports in, in the regions this repo covers.
    ANSI_CANDIDATES = ("cp1251", "cp1254", "cp1252", "cp1250")

    @staticmethod
    def _score_ansi(text):
        """How plausible is this decoding as human text? Scored per word.

        Per-character scoring does not work here. Cyrillic read through
        cp1252 becomes a run of accented Latin vowels ("Гранит" -> "Ãðàíèò")
        which still look like letters, and Turkish read through cp1251
        becomes Cyrillic ("Şantiye" -> "Юantiye") which a Cyrillic bonus
        would happily reward. What separates right from wrong is *inside the
        word*:

        * a word mixing scripts ("Юantiye" = Cyrillic Ю + Latin antiye) can
          only come from the wrong code page -- heavily penalised;
        * a word that is entirely accented Latin with no plain ASCII letters
          ("Ãðàíèò") is the signature of mojibake -- penalised;
        * a word entirely in one script, or ASCII with a few accents
          ("Şantiye"), is what real text looks like -- rewarded.

        Bilingual names ("ЗЕМЛЯНЫЕ РАБОТЫ / EARTHWORKS") work because the
        scripts live in separate words.
        """
        score = 0
        for word in text.split():
            if not any(ch >= "\x80" for ch in word):
                continue
            letters = [c for c in word if c.isalpha()]
            high_nonletter = sum(1 for c in word if c >= "\x80" and not c.isalpha())
            score -= 2 * high_nonletter
            if not letters:
                continue
            cyr = sum(1 for c in letters if 0x0400 <= ord(c) <= 0x04FF)
            lat = len(letters) - cyr
            ascii_lat = sum(1 for c in letters if c.isascii())
            high_lat = lat - ascii_lat
            if cyr and lat:
                score -= 5 * min(cyr, lat)
            elif cyr:
                score += 3 * cyr
            elif high_lat:
                score += 2 * high_lat if high_lat <= ascii_lat else -2 * high_lat
        return score

    @staticmethod
    def _system_ansi_codepage():
        """The machine's ANSI code page, e.g. 'cp1254' on a Turkish Windows."""
        try:
            import locale
            enc = locale.getpreferredencoding(False)
            return enc.lower().replace("windows-", "cp").replace("-", "")
        except Exception:  # noqa: BLE001
            return None

    def _decode_ansi(self, raw):
        scores = {}
        for enc in self.ANSI_CANDIDATES:
            try:
                text = raw.decode(enc)
            except UnicodeDecodeError:
                continue
            scores[enc] = self._score_ansi(text)
        if not scores:
            raise ValueError(
                "XER dosyasi cozumlenemedi (UTF-16LE/UTF-8/ANSI hicbiri "
                "uymadi): " + self.file_path)

        ranked = sorted(scores.items(), key=lambda kv: (-kv[1],
                                                        self.ANSI_CANDIDATES.index(kv[0])))
        self.encoding_scores = dict(ranked)
        best, best_score = ranked[0]
        tied = [e for e, s in ranked if s == best_score]

        if len(tied) == 1:
            self.encoding_confidence = "yuksek"
            self.encoding_source = "sezgisel"
        else:
            # cp1252/cp1254/cp1250 are indistinguishable for Latin text --
            # the same bytes are a valid word in each. P6 writes an ANSI XER
            # in the exporting machine's code page, so the local one is the
            # best available tie-break; say that it was a tie-break rather
            # than present it as a finding.
            system = self._system_ansi_codepage()
            if system in tied:
                best = system
                self.encoding_source = "sezgisel (esitlik sistem kod sayfasiyla bozuldu)"
            else:
                self.encoding_source = "sezgisel (esitlik, aday sirasi kullanildi)"
            self.encoding_confidence = "dusuk"
            logger.warning(
                "XER %s: %s kod sayfalari esit skorda (%d); %s secildi. "
                "Metin yanlis gorunuyorsa encoding parametresini acikca verin.",
                self.file_path, tied, best_score, best)

        self.encoding = best
        logger.info("XER %s -> ANSI kod sayfasi %s (skorlar=%s, guven=%s)",
                    self.file_path, best, self.encoding_scores,
                    self.encoding_confidence)
        return raw.decode(best, errors="replace")

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


# ---------- Field mapping helpers ----------

CONSTRAINT_TYPE_MAP = {
    "CS_ASAP": 0, "CS_ALAP": 1,
    "CS_MSO": 2, "CS_MFO": 3,
    "CS_MSOA": 4, "CS_MSOB": 5,
    "CS_MEOA": 6, "CS_MEOB": 7,
}

LINK_TYPE_MAP = {
    "PR_FS": "FS", "PR_SS": "SS", "PR_FF": "FF", "PR_SF": "SF",
}

# DCMA/MSP convention: Summary = WBS rollup or LOE (Level of Effort).
# Milestones (TT_Mile, TT_FinMile) are leaf tasks, NOT summaries.
SUMMARY_TASK_TYPES = {"TT_LOE", "TT_WBS"}


def _to_int(s, default=None):
    try:
        return int(float(s)) if s else default
    except (ValueError, TypeError):
        return default


def _to_float(s, default=0.0):
    try:
        return float(s) if s else default
    except (ValueError, TypeError):
        return default


def _to_iso_date(s):
    """XER dates are 'YYYY-MM-DD HH:MM' or empty. Return ISO date or None."""
    if not s or not s.strip():
        return None
    return s[:10]  # 'YYYY-MM-DD' prefix


# ---------- T103: read_tasks + read_links ----------

def _read_tasks(self, day_hr_cnt=8.0):
    """TASK section -> list of MSP-shape task dicts.

    day_hr_cnt: hours per working day (CAU = 9.0; default 8.0). Used to
    convert total_float_hr_cnt to days.

    DATE FIELDS (CLAUDE.md RULE 16.B — critical distinction):
    - `finish`          = target_end_date (BASELINE target; backward-compat,
                          unchanged). Do NOT use this for forecast finish.
    - `target_finish`   = explicit alias of `finish` (baseline target).
    - `early_finish`    = early_end_date (CPM early finish).
    - `late_finish`     = late_end_date (CPM late finish).
    - `forecast_finish` = when the activity actually finishes or is expected
                          to. Chain: act_end_date -> reend_date ->
                          early_end_date -> target_end_date. THIS is the field
                          a forecast-finish / slip / driver consumer must use.
    The ALFB1 9x/345-day error came from reading target_end_date as forecast.

    act_end_date comes FIRST because a completed activity's finish is a fact,
    not a forecast. P6 leaves reend_date empty once an activity is complete
    and lets early_end_date drift to the data date, so the old chain reported
    finished work as finishing on the data date -- e.g. an activity actually
    completed 2026-09-24 came back as 2026-11-01, silently erasing 38 days of
    float in any slip analysis.
    """
    tbl = self.tables.get("TASK", {"rows": []})
    out = []
    for row in tbl["rows"]:
        ttype = row.get("task_type", "")
        target_finish = _to_iso_date(row.get("target_end_date"))
        reend = _to_iso_date(row.get("reend_date"))
        early_finish = _to_iso_date(row.get("early_end_date"))
        late_finish = _to_iso_date(row.get("late_end_date"))
        actual_finish = _to_iso_date(row.get("act_end_date"))
        out.append({
            "id": _to_int(row.get("task_id")),
            "wbs_id": _to_int(row.get("wbs_id")),
            "name": row.get("task_name", ""),
            "code": row.get("task_code", ""),
            "duration_h": _to_float(row.get("target_drtn_hr_cnt")),
            "start": _to_iso_date(row.get("target_start_date")),
            "finish": target_finish,
            "target_finish": target_finish,
            "early_finish": early_finish,
            "late_finish": late_finish,
            "forecast_finish": (actual_finish or reend or early_finish
                                or target_finish),
            "actual_start": _to_iso_date(row.get("act_start_date")),
            "actual_finish": actual_finish,
            "percent_complete": _to_float(row.get("phys_complete_pct")),
            "total_float": _to_float(row.get("total_float_hr_cnt")) / day_hr_cnt
                           if day_hr_cnt > 0 else 0.0,
            "summary": ttype in SUMMARY_TASK_TYPES,
            "task_type": ttype,
            "constraint_type": CONSTRAINT_TYPE_MAP.get(row.get("cstr_type", ""), 0),
            "status": row.get("status_code", ""),
        })
    return out


def _read_links(self, day_hr_cnt=8.0):
    """TASKPRED section -> list of {from_id, to_id, type, lag_days}.

    XER `task_id` = successor; `pred_task_id` = predecessor. Map to MSP shape:
    from_id = predecessor, to_id = successor.

    day_hr_cnt: hours per working day for lag hr->day conversion (CLAUDE.md
    RULE 1). Default 8.0; pass the calendar's day_hr_cnt (CAU = 9.0) so lag
    days reflect the real working calendar.
    """
    div = float(day_hr_cnt) if day_hr_cnt else 8.0
    tbl = self.tables.get("TASKPRED", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "from_id": _to_int(row.get("pred_task_id")),
            "to_id": _to_int(row.get("task_id")),
            "type": LINK_TYPE_MAP.get(row.get("pred_type", ""), "FS"),
            "lag_days": _to_float(row.get("lag_hr_cnt")) / div,
        })
    return out


XerFile.read_tasks = _read_tasks
XerFile.read_links = _read_links


# ---------- T104: read_resources + read_assignments + read_calendars ----------

def _read_resources(self):
    """RSRC section -> list of {id, name, code, type, max_units}.

    P6 RT_Labor/RT_Equip → MSP "Work"; RT_Mat → "Material".
    """
    tbl = self.tables.get("RSRC", {"rows": []})
    out = []
    for row in tbl["rows"]:
        rtype = row.get("rsrc_type", "")
        msp_type = "Material" if rtype == "RT_Mat" else "Work"
        out.append({
            "id": _to_int(row.get("rsrc_id")),
            "name": row.get("rsrc_name", ""),
            "code": row.get("rsrc_short_name", ""),
            "type": msp_type,
            "max_units": _to_float(row.get("max_qty_per_hr"), default=1.0),
        })
    return out


def _read_assignments(self):
    """TASKRSRC section -> list of {task_id, resource_id, target_qty,
    actual_qty, target_cost, actual_cost}."""
    tbl = self.tables.get("TASKRSRC", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "task_id": _to_int(row.get("task_id")),
            "resource_id": _to_int(row.get("rsrc_id")),
            "target_qty": _to_float(row.get("target_qty")),
            "actual_qty": _to_float(row.get("act_reg_qty")),
            "target_cost": _to_float(row.get("target_cost")),
            "actual_cost": _to_float(row.get("act_reg_cost")),
        })
    return out


import datetime as _dt
import re as _re

# Phase 6.4 — Primavera P6 clndr_data exception/holiday BLOB pattern.
# P6 stores exception dates as `(d|<excel_serial>|f|<bit>)` inside the
# clndr_data text BLOB. <bit>=0 means non-working (holiday), <bit>=1 means
# working-day exception (override). Excel serial 1 = 1900-01-01 with the
# off-by-2 quirk (epoch base 1899-12-30).
_CLNDR_EXCEPT_RE = _re.compile(r"\(d\|(\d+)\|f\|(\d)")
_CLNDR_EXCEL_EPOCH = _dt.date(1899, 12, 30)


def _parse_clndr_data(clndr_data):
    """Best-effort extract of exception/holiday dates from a P6 clndr_data
    BLOB. Returns list of {date: 'YYYY-MM-DD', working: bool}.

    Tolerant: empty/None input -> []. Unparseable serials skipped silently.
    No exceptions block in BLOB -> [].
    """
    if not clndr_data:
        return []
    out = []
    for serial_str, bit_str in _CLNDR_EXCEPT_RE.findall(clndr_data):
        try:
            j = int(serial_str)
            d = _CLNDR_EXCEL_EPOCH + _dt.timedelta(days=j)
            out.append({"date": d.isoformat(), "working": bit_str == "1"})
        except (ValueError, OverflowError):
            continue
    return out


def _read_calendars(self):
    """CALENDAR section -> list of {id, name, day_hr_cnt, week_hr_cnt,
    exceptions: [{date, working}]}.

    Phase 6.4 — clndr_data BLOB parsed for exception/holiday dates via
    best-effort regex (Primavera proprietary format, not publicly
    specified). Empty BLOB or no exception block -> exceptions=[].
    """
    tbl = self.tables.get("CALENDAR", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "id": _to_int(row.get("clndr_id")),
            "name": row.get("clndr_name", ""),
            "day_hr_cnt": _to_float(row.get("day_hr_cnt"), default=8.0),
            "week_hr_cnt": _to_float(row.get("week_hr_cnt"), default=40.0),
            "exceptions": _parse_clndr_data(row.get("clndr_data", "")),
        })
    return out


XerFile.read_resources = _read_resources
XerFile.read_assignments = _read_assignments
XerFile.read_calendars = _read_calendars


# ---------- T105: read_progress + status_date + project metadata ----------

def _read_project(self):
    """PROJECT section -> first row dict (typically only 1 project per XER)."""
    tbl = self.tables.get("PROJECT", {"rows": []})
    if not tbl["rows"]:
        return {}
    row = tbl["rows"][0]
    return {
        "proj_id": _to_int(row.get("proj_id")),
        "proj_short_name": row.get("proj_short_name", ""),
        "plan_start_date": _to_iso_date(row.get("plan_start_date")),
        "plan_end_date": _to_iso_date(row.get("plan_end_date")),
        "last_recalc_date": _to_iso_date(row.get("last_recalc_date")),
    }


def _read_progress(self):
    """Return {status_date, tasks: [{id, percent_complete, actual_work_h}]}.

    status_date = PROJECT.last_recalc_date (P6 convention — XER files do
    not track a separate 'data date' field, last recalc is the de-facto
    status date).

    actual_work_h: aggregate sum of TASKRSRC.act_reg_qty per task (XER
    stores actuals at assignment level, not task level).
    """
    proj = _read_project(self)
    # Pre-aggregate actual work per task from TASKRSRC
    actual_by_task = {}
    for asgn in self.tables.get("TASKRSRC", {"rows": []})["rows"]:
        tid = _to_int(asgn.get("task_id"))
        if tid is None:
            continue
        actual_by_task[tid] = actual_by_task.get(tid, 0.0) + _to_float(
            asgn.get("act_reg_qty"))
    progress_tasks = []
    for row in self.tables.get("TASK", {"rows": []})["rows"]:
        tid = _to_int(row.get("task_id"))
        progress_tasks.append({
            "id": tid,
            "percent_complete": _to_float(row.get("phys_complete_pct")),
            "actual_work_h": actual_by_task.get(tid, 0.0),
        })
    return {
        "status_date": (proj or {}).get("last_recalc_date"),
        "tasks": progress_tasks,
    }


def _read_wbs(self):
    """PROJWBS section -> list of {id, parent_id, name, code, proj_id}.

    Used by forecast-driver analysis to group tasks under top-level WBS
    nodes and compare per-branch forecast finish (CLAUDE.md RULE 16.C).
    Top-level WBS = node whose parent_id is the project root (a WBS id not
    itself present as a child) or None.
    """
    tbl = self.tables.get("PROJWBS", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "id": _to_int(row.get("wbs_id")),
            "parent_id": _to_int(row.get("parent_wbs_id")),
            "name": row.get("wbs_name", ""),
            "code": row.get("wbs_short_name", ""),
            "proj_id": _to_int(row.get("proj_id")),
        })
    return out


XerFile.read_project = _read_project
XerFile.read_progress = _read_progress
XerFile.read_wbs = _read_wbs
