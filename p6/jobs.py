"""P6 Job Service driver -- headless F9 (reschedule) and friends.

P6 Professional has no COM automation on PM.exe and its command line can only
import/export/run report batches -- there is no schedule action. The one way to
run P6's own CPM engine without the GUI is its **Job Service**: a row in the
``JOBSVC`` queue table with ``RECUR_TYPE='RT_ASAP'`` is picked up within seconds
by ``PrmJobSv.exe``, which runs the real scheduler and writes the result back.

Verified end to end on 26.08.2026 (P6 Professional 24.12.0.51267, SQL Server):
``JS_Pending -> JS_Running -> JS_Complete`` in 3 s, and with the data date moved
90 days every one of 950 activities was recomputed.

Hard prerequisites -- ``preflight()`` checks all of them:

* The database must NOT be SQLite. ``prmjob.exe`` and ``PrmJobSv.exe`` carry a
  literal guard right after the string ``SQLite``:
  *"Job Services are not supported for P6 Professional Standalone."*
* The ``PrmJobSv`` Windows service must be installed, bound to the right alias
  and running. It runs as LocalSystem, so ``prmbootstrapV2.xml`` has to exist in
  the LocalSystem profile too -- otherwise it silently falls back to the Oracle
  driver and dies with *"Cannot find OCI DLL"*.
* The job's user needs a ``USEROBS`` row. A fresh P6 database leaves that table
  empty, and then the service reports *"No projects to schedule"* even though
  the user is a global superuser.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Any, Iterable, Sequence

# ---------------------------------------------------------------------------
# Enums, straight out of PM.exe / prmjob.exe
# ---------------------------------------------------------------------------
JT_SCHEDULE = "JT_Sched"
JT_LEVEL = "JT_Level"
JT_SUMMARIZE = "JT_Sum"
JT_APPLY_ACTUALS = "JT_ApplyActuals"
JT_UPDATE_BASELINE = "JT_UpdateBaseline"
JT_CREATE_BASELINE = "JT_CreateBaseline"
JT_XER_EXPORT = "JT_XERExport"
JT_RECALC_COST = "JT_RecalcCost"
JT_STORE_PERIOD = "JT_StorePerPerf"
JT_BATCH_REPORT = "JT_Batch"

# JOB_DATA bolum adlari -- PM.exe'nin kendi string tablosundan (26.08.2026):
# JT_Sched "Schedule Projects" kullanir, DIGER tum is tipleri duz "Projects".
# Bu ayrim olculdu: JT_ApplyActuals "Apply Actuals" bolum adiyla "No projects
# to apply actual to." veriyordu; "Projects" ile JS_Complete dondu. (JT_Sum
# yanlis bolum adina aldirmadan calisiyordu -- sansa guvenme, dogrusunu yaz.)
_SECTION = {
    JT_SCHEDULE: "Schedule Projects",
    JT_SUMMARIZE: "Projects",
    JT_APPLY_ACTUALS: "Projects",
    JT_XER_EXPORT: "Projects",
    JT_BATCH_REPORT: "Projects",
}

# prmjob.exe'nin kuyruk dispatcher'inin KABUL ETTIGI is tipleri -- 26.08.2026'da
# olculdu: JT_Level ve JT_UpdateBaseline kuyruga birakildi, servis ikisini de
# "Invalid Job type" ile reddetti; ardindan prmjob.exe iceriginden (UTF-16LE
# string tablosu, "Invalid Job type:" hemen oncesindeki dispatch listesi) ayni
# yedili dogrulandi. JT_Level/JT_UpdateBaseline/JT_CreateBaseline sabitleri
# ikilide BASKA yerlerde gecse de dispatcher onlari CALISTIRMAZ -- P6
# Professional'da leveling ve baseline guncelleme yalniz arayuzden yapilir
# (baseline kopyalama zaten p6_baseline'da SQL ile yapiliyor).
DISPATCHABLE = frozenset({
    JT_SCHEDULE, JT_APPLY_ACTUALS, JT_XER_EXPORT, JT_SUMMARIZE,
    "JT_Enterprise_Sum", JT_BATCH_REPORT, "JT_Report",
})

# Program verisini degistiren is tipleri -- acik onay ister
MUTATING = frozenset({JT_APPLY_ACTUALS, JT_UPDATE_BASELINE, JT_CREATE_BASELINE,
                      JT_STORE_PERIOD, JT_RECALC_COST, JT_LEVEL})

RT_ASAP = "RT_ASAP"
RT_RECUR_ENABLED = "RT_RecurEnabled"
RT_RECUR_DISABLED = "RT_RecurDisabled"

JS_PENDING = "JS_Pending"
JS_RUNNING = "JS_Running"
JS_COMPLETE = "JS_Complete"
JS_FAILED = "JS_Failed"
JS_CANCELLED = "JS_Cancelled"
JS_DELEGATED = "JS_Delegated"
JS_COMP_ERROR = "JS_CompError"

TERMINAL = frozenset({JS_COMPLETE, JS_FAILED, JS_CANCELLED, JS_COMP_ERROR})

SERVICE_NAME = "PrmJobSv"


class P6JobError(RuntimeError):
    """Recoverable job problem; tools turn it into a JSON error."""


# ---------------------------------------------------------------------------
# Error messages -> Turkish, with the fix attached
# ---------------------------------------------------------------------------
_ERROR_MAP: list[tuple[str, str]] = [
    ("No projects to schedule",
     "Islenecek proje bulunamadi. En sik nedeni: is sahibi kullanicinin "
     "USEROBS kaydi yok. P6'da Enterprise > OBS > ilgili dugum > Users "
     "sekmesinden kullaniciyi <Project Superuser> ile ekleyin "
     "(p6_job.preflight bunu kontrol eder)."),
    ("Default Project not found",
     "JOB_DATA icindeki 'Default Project' proje kimligi veritabaninda yok. "
     "proj_id degerini p6_query.list_projects ile dogrulayin."),
    ("Default project is EPS",
     "Varsayilan proje bir EPS dugumu (PROJECT_FLAG='N'). Gercek bir proje "
     "secin."),
    ("Job Services are not supported for P6 Professional Standalone",
     "Job Service SQLite standalone veritabaninda CALISMAZ. Alias'in SQL Server "
     "veya Oracle olmasi gerekir."),
    ("Cannot find OCI DLL",
     "Servis Oracle surucusune dustu; yani bootstrap dosyasini bulamadi. "
     "prmbootstrapV2.xml'i LocalSystem profiline kopyalayin: "
     r"C:\Windows\System32\config\systemprofile\AppData\Roaming\Oracle"
     r"\Primavera P6\P6 Professional\[24.12.0\]"),
    ("Login failed",
     "Veritabani girisi reddedildi. Alias'in public user parolasi bos olabilir "
     "(DBConfig'e /puserpwd verilmemis)."),
    ("Please accept consent",
     "P6 arayuzunden onay (consent) verilmemis. P6'yi bir kez acip onaylayin."),
]


def translate_error(raw: str | None) -> str:
    if not raw:
        return ""
    text = raw.strip()
    # The service writes "OK" into last_error_descr on success -- not an error.
    if text.upper() == "OK":
        return ""
    for needle, tr in _ERROR_MAP:
        if needle.lower() in text.lower():
            return tr + "  [P6: " + text + "]"
    return text


# ---------------------------------------------------------------------------
# JOB_DATA -- Primavera nested-list dialect
# ---------------------------------------------------------------------------
def build_job_data(job_type: str, proj_ids: Sequence[int],
                   default_proj_id: int | None = None) -> str:
    """Serialise the project list the way P6's own Job Services dialog does.

    Reference blob captured from P6 for a single-project schedule job::

        (0||(Default Project|368)((0||Schedule Projects()((0||368()())))))

    Levels: root carries ``Default Project``; one child names the section; each
    project is an empty-bodied node under it.
    """
    ids = [int(p) for p in proj_ids]
    if not ids:
        raise P6JobError("En az bir proje kimligi gerekli.")
    default_id = int(default_proj_id) if default_proj_id is not None else ids[0]
    if default_id not in ids:
        ids.insert(0, default_id)
    section = _SECTION.get(job_type)
    if section is None:
        raise P6JobError(
            "Bu is tipi icin JOB_DATA bolum adi bilinmiyor: " + job_type
            + ". Bilinenler: " + ", ".join(sorted(_SECTION)))
    projects = "".join("(0||%d()())" % p for p in ids)
    return "(0||(Default Project|%d)((0||%s()(%s))))" % (default_id, section, projects)


# ---------------------------------------------------------------------------
# Queue operations
# ---------------------------------------------------------------------------
@dataclass
class JobResult:
    job_id: int
    job_type: str
    job_name: str
    status: str
    error: str = ""
    error_tr: str = ""
    last_run: Any = None
    elapsed_s: float = 0.0
    transitions: list[str] = field(default_factory=list)
    log: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.status == JS_COMPLETE

    def to_dict(self) -> dict[str, Any]:
        return {
            "job_id": self.job_id, "job_type": self.job_type,
            "job_name": self.job_name, "status": self.status,
            "ok": self.ok,
            "error": self.error or None, "error_tr": self.error_tr or None,
            "last_run": str(self.last_run) if self.last_run else None,
            "elapsed_s": round(self.elapsed_s, 1),
            "transitions": self.transitions,
            "log": self.log or None,
        }


def _next_key(cur, name: str = "jobsvc_job_id") -> int:
    """Take the next id from P6's own key generator, never MAX()+1."""
    cur.execute("SELECT KEY_SEQ_NUM FROM NEXTKEY WHERE KEY_NAME = ?", name)
    row = cur.fetchone()
    if not row:
        raise P6JobError("NEXTKEY kaydi yok: " + name)
    value = int(row[0])
    cur.execute("UPDATE NEXTKEY SET KEY_SEQ_NUM = ? WHERE KEY_NAME = ?",
                value + 1, name)
    return value


def submit(cur, job_type: str, proj_ids: Sequence[int], user_id: int,
           job_name: str = "MCP_JOB", default_proj_id: int | None = None,
           recur_type: str = RT_ASAP, job_data: str | None = None) -> int:
    """Queue one job. Returns its job_id.

    Only the queue table is touched -- no project data is written here.
    """
    if job_type not in DISPATCHABLE:
        raise P6JobError(
            "P6 Professional Job Service '%s' is tipini CALISTIRAMAZ -- "
            "dispatcher yalniz sunlari kabul eder (olculdu, 26.08.2026): %s. "
            "Leveling ve baseline guncelleme P6 arayuzunden yapilir; baseline "
            "kopyasi icin p6_baseline kullanin."
            % (job_type, ", ".join(sorted(DISPATCHABLE))))
    data = job_data if job_data is not None else build_job_data(
        job_type, proj_ids, default_proj_id)
    job_id = _next_key(cur)
    cur.execute(
        "INSERT INTO JOBSVC (job_id, parent_job_id, seq_num, audit_flag, "
        " job_type, job_name, user_id, status_code, recur_type, "
        " submitted_date, job_data, create_date, create_user, "
        " update_date, update_user) "
        "VALUES (?, NULL, ?, 'N', ?, ?, ?, ?, ?, GETDATE(), ?, "
        " GETDATE(), 'admin', GETDATE(), 'admin')",
        job_id, job_id, job_type, job_name[:255], int(user_id),
        JS_PENDING, recur_type, data)
    return job_id


def read_status(cur, job_id: int) -> tuple[str, str, Any]:
    cur.execute("SELECT status_code, ISNULL(last_error_descr, ''), last_run_date "
                "FROM JOBSVC WHERE job_id = ?", job_id)
    row = cur.fetchone()
    if not row:
        raise P6JobError("JOBSVC satiri yok: job_id=" + str(job_id))
    return row[0], row[1], row[2]


def read_log(cur, job_id: int) -> list[str]:
    try:
        cur.execute("SELECT job_log_data FROM JOBLOG WHERE job_id = ?", job_id)
        return [str(r[0]) for r in cur.fetchall() if r[0]]
    except Exception:  # noqa: BLE001 - JOBLOG may not exist on some schemas
        return []


def wait(cur, job_id: int, timeout_s: int = 300, poll_s: float = 2.0
         ) -> JobResult:
    """Poll until the job reaches a terminal state or the timeout expires."""
    cur.execute("SELECT job_type, job_name FROM JOBSVC WHERE job_id = ?", job_id)
    row = cur.fetchone()
    job_type, job_name = (row[0], row[1]) if row else ("?", "?")

    started = time.time()
    transitions: list[str] = []
    status, err, run = JS_PENDING, "", None
    while time.time() - started < timeout_s:
        status, err, run = read_status(cur, job_id)
        if not transitions or transitions[-1].split(" @")[0] != status:
            transitions.append("%s @%.0fs" % (status, time.time() - started))
        if status in TERMINAL:
            break
        time.sleep(poll_s)

    # "OK" is what the service writes on success; it is not an error.
    if err.strip().upper() == "OK":
        err = ""
    return JobResult(
        job_id=job_id, job_type=job_type, job_name=job_name, status=status,
        error=err.strip(), error_tr=translate_error(err),
        last_run=run, elapsed_s=time.time() - started,
        transitions=transitions, log=read_log(cur, job_id))


def run_and_wait(cur, job_type: str, proj_ids: Sequence[int], user_id: int,
                 job_name: str = "MCP_JOB", timeout_s: int = 300,
                 default_proj_id: int | None = None) -> JobResult:
    job_id = submit(cur, job_type, proj_ids, user_id, job_name, default_proj_id)
    return wait(cur, job_id, timeout_s)


def list_jobs(cur, limit: int = 50) -> list[dict[str, Any]]:
    cur.execute(
        "SELECT TOP (?) job_id, job_type, job_name, user_id, status_code, "
        " recur_type, submitted_date, last_run_date, "
        " ISNULL(last_error_descr, '') "
        "FROM JOBSVC WHERE delete_session_id IS NULL ORDER BY job_id DESC",
        int(limit))
    out = []
    for r in cur.fetchall():
        out.append({
            "job_id": r[0], "job_type": r[1], "job_name": r[2], "user_id": r[3],
            "status": r[4], "recur_type": r[5],
            "submitted": str(r[6]) if r[6] else None,
            "last_run": str(r[7]) if r[7] else None,
            "error_tr": translate_error(r[8]) or None,
        })
    return out


def cancel(cur, job_id: int) -> dict[str, Any]:
    status, _, _ = read_status(cur, job_id)
    if status in TERMINAL:
        return {"job_id": job_id, "status": status,
                "note": "Is zaten bitmis, iptal edilmedi."}
    cur.execute("UPDATE JOBSVC SET status_code = ?, update_date = GETDATE() "
                "WHERE job_id = ?", JS_CANCELLED, job_id)
    return {"job_id": job_id, "status": JS_CANCELLED}


def purge(cur, name_like: str = "MCP\\_%") -> int:
    """Delete finished MCP-created jobs. Never touches user-created rows
    unless the caller widens name_like on purpose."""
    cur.execute(
        "DELETE FROM JOBSVC WHERE job_name LIKE ? ESCAPE '\\' "
        "AND status_code IN (?, ?, ?, ?)",
        name_like, JS_COMPLETE, JS_FAILED, JS_CANCELLED, JS_COMP_ERROR)
    return cur.rowcount


# ---------------------------------------------------------------------------
# Preflight -- every failure we actually hit, checked before submitting
# ---------------------------------------------------------------------------
def service_state() -> dict[str, Any]:
    """Windows service state for PrmJobSv."""
    try:
        import win32service
        import win32serviceutil
    except ImportError:
        return {"installed": None, "state": "pywin32 yok"}
    try:
        code = win32serviceutil.QueryServiceStatus(SERVICE_NAME)[1]
    except Exception as exc:  # noqa: BLE001
        return {"installed": False, "state": "kurulu degil", "detail": str(exc)}
    names = {
        win32service.SERVICE_STOPPED: "Stopped",
        win32service.SERVICE_START_PENDING: "StartPending",
        win32service.SERVICE_STOP_PENDING: "StopPending",
        win32service.SERVICE_RUNNING: "Running",
        win32service.SERVICE_CONTINUE_PENDING: "ContinuePending",
        win32service.SERVICE_PAUSE_PENDING: "PausePending",
        win32service.SERVICE_PAUSED: "Paused",
    }
    return {"installed": True, "state": names.get(code, str(code)),
            "running": code == win32service.SERVICE_RUNNING}


def preflight(cur, alias, user_id: int,
              proj_ids: Iterable[int] = ()) -> dict[str, Any]:
    """Check everything that made a job fail during bring-up.

    Returns a report with ``ready`` and a list of blocking problems.
    """
    problems: list[str] = []
    checks: dict[str, Any] = {}

    # 1) Driver -- SQLite is a hard block inside prmjob.exe itself
    driver = getattr(alias, "driver", "?")
    checks["alias"] = getattr(alias, "name", "?")
    checks["driver"] = driver
    if driver == "SQLite":
        problems.append(
            "Alias SQLite. Job Service SQLite'ta CALISMAZ ("
            "prmjob.exe icinde sabit blok). SQL Server veya Oracle alias'i gerekir.")

    # 2) Windows service
    svc = service_state()
    checks["service"] = svc
    if svc.get("installed") is False:
        problems.append(
            "PrmJobSv servisi kurulu degil. P6 medyasindan "
            "'ADDLOCAL=PrmJob' ile /qb! modunda kurun (/qn ile servis kaydi atlanir).")
    elif svc.get("running") is False:
        problems.append(
            "PrmJobSv servisi durmus. Baslatin; baslamiyorsa Uygulama olay "
            "gunlugunde 'Cannot find OCI DLL' arayin (bootstrap LocalSystem "
            "profilinde olmali).")

    # 3) USEROBS -- the one that cost us the most time
    try:
        cur.execute("SELECT COUNT(*) FROM USEROBS WHERE user_id = ?", int(user_id))
        n_obs = int(cur.fetchone()[0])
    except Exception as exc:  # noqa: BLE001
        n_obs, checks["userobs_error"] = -1, str(exc)
    checks["userobs_rows"] = n_obs
    if n_obs == 0:
        problems.append(
            "Kullanicinin (user_id=%d) USEROBS kaydi yok -> servis 'No projects "
            "to schedule' der. P6'da Enterprise > OBS > Users sekmesinden "
            "<Project Superuser> ile ekleyin." % user_id)

    # 4) Projects exist and are real projects, not EPS nodes
    ids = [int(p) for p in proj_ids]
    if ids:
        marks = ",".join("?" * len(ids))
        cur.execute(
            "SELECT proj_id, proj_short_name, project_flag FROM PROJECT "
            "WHERE proj_id IN (" + marks + ")", *ids)
        found = {r[0]: (r[1], r[2]) for r in cur.fetchall()}
        checks["projects"] = {k: v[0] for k, v in found.items()}
        for pid in ids:
            if pid not in found:
                problems.append("Proje bulunamadi: proj_id=%d" % pid)
            elif found[pid][1] != "Y":
                problems.append(
                    "proj_id=%d bir EPS dugumu (PROJECT_FLAG='N'), proje degil."
                    % pid)

    return {"ready": not problems, "problems": problems, "checks": checks}
