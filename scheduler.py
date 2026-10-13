"""
IFRS Report Scheduler
---------------------
Runs ERP bots on daily / weekly / monthly schedules.
Logs every run to scheduler_log.xlsx (created once; appended forever).

Start alongside ifrs.py:
    python scheduler.py

Or import and call start_scheduler() from your Flask bootstrap.
"""

import shutil
import sys
import threading
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Bot modules are imported locally inside each job function (after sys.path is
# updated) to avoid top-level ImportError when bots/ is not on sys.path yet.

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).resolve().parent   # project root (c:\CBSL\Script)
BOTS_DIR      = BASE_DIR / "bots"                 # Selenium bot scripts
LOG_FILE      = BASE_DIR / "scheduler_log.xlsx"   # persistent run log
SCHED_LOG_DIR = BASE_DIR / "logs" / "scheduler"   # per-job text logs

DOWNLOAD_FOLDER_DIR = Path.home() / "Downloads"   # where Chrome saves files


# ── Per-job log file ──────────────────────────────────────────────────────────
@contextmanager
def _job_log(job_name: str):
    """
    Context manager that tees stdout/stderr to a timestamped log file in
    logs/scheduler/<YYYYMMDD_HHMMSS>_<job_name>.log for the duration of a job.
    Restores original stdout/stderr on exit even if an exception is raised.
    """
    SCHED_LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Sanitise job_name so it is safe as a filename
    safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in job_name)
    log_path = SCHED_LOG_DIR / f"{ts}_{safe_name}.log"

    orig_stdout, orig_stderr = sys.stdout, sys.stderr
    try:
        with open(log_path, "w", encoding="utf-8", errors="replace") as fh:
            # _Tee mirrors every write to both the original stream and the log file
            class _Tee:
                def __init__(self, orig):
                    self._orig = orig
                def write(self, data):
                    self._orig.write(data)
                    fh.write(data)
                    fh.flush()
                def flush(self):
                    self._orig.flush()
                    fh.flush()
                def fileno(self):
                    return self._orig.fileno()

            sys.stdout = _Tee(orig_stdout)
            sys.stderr = _Tee(orig_stderr)
            print(f"Log file: {log_path}")
            yield log_path
    finally:
        # Always restore original streams, even on exception
        sys.stdout = orig_stdout
        sys.stderr = orig_stderr

# ── Excel log ────────────────────────────────────────────────────────────────
_LOG_HEADERS = [
    "Timestamp", "Job Name", "Frequency",
    "Downloaded File", "Transferred To", "Status", "Notes",
]
# Lock prevents concurrent writes from two job threads corrupting the workbook
_log_lock = threading.Lock()


def _init_log_file() -> None:
    """Create scheduler_log.xlsx with a header row ONLY if it does not exist yet."""
    if LOG_FILE.exists():
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Scheduler Log"
    hdr_fill = PatternFill("solid", fgColor="1F4E79")   # dark blue header
    hdr_font = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(_LOG_HEADERS, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal="center")
    # Set column widths to make the log readable without manual resizing
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 38
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 48
    wb.save(LOG_FILE)


def log_entry(
    job_name: str,
    frequency: str,
    downloaded_file: str = "",
    transferred_to: str = "",
    status: str = "Success",
    notes: str = "",
) -> None:
    """Append one row to the persistent scheduler_log.xlsx (thread-safe)."""
    with _log_lock:
        wb = openpyxl.load_workbook(LOG_FILE)
        ws = wb.active
        ws.append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            job_name, frequency, downloaded_file, transferred_to, status, notes,
        ])
        # Colour-code the Status cell: green for Success, red for Failed
        status_cell = ws.cell(row=ws.max_row, column=6)
        if status == "Success":
            status_cell.fill = PatternFill("solid", fgColor="C6EFCE")
            status_cell.font = Font(color="276221")
        else:
            status_cell.fill = PatternFill("solid", fgColor="FFC7CE")
            status_cell.font = Font(color="9C0006")
        wb.save(LOG_FILE)


def read_log(limit: int = 50) -> list[dict]:
    """Return the last `limit` log rows as a list of dicts, newest first."""
    with _log_lock:
        wb = openpyxl.load_workbook(LOG_FILE, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        wb.close()
    # Slice the tail then reverse so the newest entry is index 0
    rows = rows[-limit:][::-1]
    return [dict(zip(_LOG_HEADERS, r)) for r in rows]


# ── Helpers ──────────────────────────────────────────────────────────────────

def _newest_file(folder: Path, *patterns: str) -> Path | None:
    """Return the most-recently-modified file matching any of the given glob patterns."""
    for pat in patterns:
        matches = list(folder.glob(pat))
        if matches:
            return max(matches, key=lambda f: f.stat().st_mtime)
    return None


def _latest_files(folder: Path, count: int, *patterns: str) -> list[Path]:
    """Return the `count` most-recently-modified files matching any pattern."""
    files = []
    for pat in patterns:
        files.extend(folder.glob(pat))
    files = list(set(files))  # remove duplicates if patterns overlap
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files[:count]


def _find_lgd_outputs_lgd_subfolder() -> Path | None:
    """
    Resolve outputs/LGD/<latest-date-folder>/LGD/ using the last date in
    outputs/LGD/completed_dates.txt and the max versioned folder for that date.
    Returns None if the folder cannot be resolved.
    """
    import re as _re
    outputs_lgd = BASE_DIR / "outputs" / "LGD"
    dates_file  = outputs_lgd / "completed_dates.txt"
    if not dates_file.exists():
        return None
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    if not lines:
        return None
    last_date = lines[-1]  # most recently completed LGD run date (DD-MM-YYYY)

    # Build alternative date formats so we match both "31-03-2025" and "31/03/2025"
    date_variants = {last_date}
    for fmt_in, fmt_out in [("%d-%m-%Y", "%d/%m/%Y"), ("%d/%m/%Y", "%d-%m-%Y")]:
        try:
            date_variants.add(datetime.strptime(last_date, fmt_in).strftime(fmt_out))
        except ValueError:
            pass

    # Collect candidate folders and their version suffixes (e.g. 31-03-2025(2) → version 2)
    candidates: list[tuple[int, Path]] = []
    for child in outputs_lgd.iterdir():
        if not child.is_dir():
            continue
        if any(child.name.startswith(v) for v in date_variants):
            m = _re.search(r"\((\d+)\)$", child.name)
            candidates.append((int(m.group(1)) if m else 0, child))

    if not candidates:
        return None
    # Pick the highest-version folder (most recent run for that date)
    candidates.sort(key=lambda x: x[0], reverse=True)
    lgd_sub = candidates[0][1] / "LGD"
    lgd_sub.mkdir(parents=True, exist_ok=True)
    return lgd_sub

def _find_lgd_outputs_firstDefault_subfolder() -> Path | None:
    """
    Resolve outputs/LGD/<latest-date-folder>/firstDefault/ using the last date in
    outputs/LGD/completed_dates.txt and the max versioned folder for that date.
    Returns None if the folder cannot be resolved.
    """
    import re as _re
    outputs_lgd = BASE_DIR / "outputs" / "LGD"
    dates_file  = outputs_lgd / "completed_dates.txt"
    if not dates_file.exists():
        return None
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    if not lines:
        return None
    last_date = lines[-1]  # most recently completed LGD run date (DD-MM-YYYY)

    # Build alternative date formats so we match both "31-03-2025" and "31/03/2025"
    date_variants = {last_date}
    for fmt_in, fmt_out in [("%d-%m-%Y", "%d/%m/%Y"), ("%d/%m/%Y", "%d-%m-%Y")]:
        try:
            date_variants.add(datetime.strptime(last_date, fmt_in).strftime(fmt_out))
        except ValueError:
            pass

    # Collect candidate folders and their version suffixes
    candidates: list[tuple[int, Path]] = []
    for child in outputs_lgd.iterdir():
        if not child.is_dir():
            continue
        if any(child.name.startswith(v) for v in date_variants):
            m = _re.search(r"\((\d+)\)$", child.name)
            candidates.append((int(m.group(1)) if m else 0, child))

    if not candidates:
        return None
    # Pick the highest-version folder (most recent run for that date)
    candidates.sort(key=lambda x: x[0], reverse=True)
    lgd_sub = candidates[0][1] / "firstDefault"
    lgd_sub.mkdir(parents=True, exist_ok=True)
    return lgd_sub


def _find_summary_input_subfolder() -> Path | None:
    """
    Resolve outputs/Summary/<latest-date-folder>/Input/ using the last date in
    outputs/Summary/completed_dates.txt and the max versioned folder for that date.
    Returns None if the folder cannot be resolved.
    """
    import re as _re
    outputs_summary = BASE_DIR / "outputs" / "Summary"
    dates_file      = outputs_summary / "completed_dates.txt"
    if not dates_file.exists():
        return None
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    if not lines:
        return None
    last_date = lines[-1]

    date_variants = {last_date}
    for fmt_in, fmt_out in [("%d-%m-%Y", "%d/%m/%Y"), ("%d/%m/%Y", "%d-%m-%Y")]:
        try:
            date_variants.add(datetime.strptime(last_date, fmt_in).strftime(fmt_out))
        except ValueError:
            pass

    candidates: list[tuple[int, Path]] = []
    for child in outputs_summary.iterdir():
        if not child.is_dir():
            continue
        if any(child.name.startswith(v) for v in date_variants):
            m = _re.search(r"\((\d+)\)$", child.name)
            candidates.append((int(m.group(1)) if m else 0, child))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    input_sub = candidates[0][1] / "Input"
    input_sub.mkdir(parents=True, exist_ok=True)
    return input_sub


# ── Job definitions ───────────────────────────────────────────────────────────
# Each job function:
#   1. Inserts BOTS_DIR onto sys.path so the bot can be imported by short name.
#   2. Calls the bot's entry-point function; raises if it returns False.
#   3. Locates the downloaded file(s) in the Downloads folder.
#   4. Moves the file(s) to the correct outputs subfolder.
#   5. Appends a row to scheduler_log.xlsx via log_entry().
# The docstring first line is used as the display name in the dashboard.

def job_ledger_Accounts():
    """DM Ledger Accounts"""
    job_name  = "Ledger Details"
    frequency = "Daily"
    with _job_log(job_name):
        try:
            sys.path.insert(0, str(BOTS_DIR))  # allow bare `from DM_Ledger_Accounts_main import …`
            from DM_Ledger_Accounts_main import Download_Ledger_Accounts_report
            ok = Download_Ledger_Accounts_report()
            if not ok:
                raise RuntimeError("Bot returned False — download did not complete")
            newest = _newest_file(DOWNLOAD_FOLDER_DIR, "Ledger Details Report - Internal*.xlsx")
            if not newest:
                raise RuntimeError("Downloaded file not found in Downloads folder")
            lgd_sub = _find_lgd_outputs_lgd_subfolder()
            if not lgd_sub:
                raise RuntimeError("Could not resolve outputs/LGD/<date>/LGD/ folder")
            # Remove any stale copy before moving the fresh download
            for old in lgd_sub.glob("Ledger Details Report - Internal*.xlsx"):
                old.unlink()
            dest = lgd_sub / newest.name
            shutil.move(str(newest), str(dest))
            print(f"Moved {newest.name} → {dest}")
            log_entry(job_name, frequency,
                      downloaded_file=newest.name,
                      transferred_to=str(dest),
                      status="Success")
        except Exception as exc:
            print(f"ERROR: {exc}")
            log_entry(job_name, frequency, status="Failed", notes=str(exc))


def job_refinance():
    """LGD Refinance DPD"""
    job_name  = "Refinance DPD"
    frequency = "Daily"
    with _job_log(job_name):
        try:
            sys.path.insert(0, str(BOTS_DIR))
            from LGD_refinance_dpd import run
            # Cumulative window (REFINANCE_START_DATE .. REFINANCE_END_DATE in the bot)
            ok = run()
            if not ok:
                raise RuntimeError("Bot returned False — download did not complete")
            newest = _newest_file(DOWNLOAD_FOLDER_DIR, "Refinance DPD Report*.xlsx")
            if not newest:
                raise RuntimeError("Downloaded file not found in Downloads folder")
            lgd_sub = _find_lgd_outputs_lgd_subfolder()
            if not lgd_sub:
                raise RuntimeError("Could not resolve outputs/LGD/<date>/LGD/ folder")
            for old in lgd_sub.glob("Refinance DPD Report*.xlsx"):
                old.unlink()
            dest = lgd_sub / newest.name
            shutil.move(str(newest), str(dest))
            print(f"Moved {newest.name} → {dest}")
            log_entry(job_name, frequency,
                      downloaded_file=newest.name,
                      transferred_to=str(dest),
                      status="Success")
        except Exception as exc:
            print(f"ERROR: {exc}")
            log_entry(job_name, frequency, status="Failed", notes=str(exc))


def job_summary_reports():
    """DM Daily Deposit Balance"""
    job_name  = "Summary"
    frequency = "Daily"
    with _job_log(job_name):
        try:
            sys.path.insert(0, str(BOTS_DIR))
            from IFRS9_Download_Summary_Requirements import run_main
            ok = run_main(datetime.today().strftime("%d/%m/%Y"))
            if not ok:
                raise RuntimeError("Bot returned False")
            # This bot downloads up to 4 files (mix of CSV and XLSX)
            latest_files = _latest_files(DOWNLOAD_FOLDER_DIR, 4, "*.csv", "*.xlsx")
            if not latest_files:
                raise RuntimeError("No downloaded files found in Downloads folder")
            summary_input = _find_summary_input_subfolder()
            if not summary_input:
                raise RuntimeError("Could not resolve outputs/Summary/<date>/Input/ folder")
            downloaded_list = []
            for f in latest_files:
                dest = summary_input / f.name
                shutil.move(str(f), str(dest))
                downloaded_list.append(f.name)
                print(f"Moved {f.name} → {dest}")
            log_entry(job_name, frequency,
                      downloaded_file=", ".join(downloaded_list),
                      transferred_to=str(summary_input),
                      status="Success")
        except Exception as exc:
            print(f"ERROR: {exc}")
            log_entry(job_name, frequency, status="Failed", notes=str(exc))


def job_legal_bal_new():
    """Legal Bal New Report"""
    job_name  = "Legal Bal New"
    frequency = "Daily"
    with _job_log(job_name):
        try:
            sys.path.insert(0, str(BOTS_DIR))
            from Legal_Bal_New import erp_login_closure_report
            driver = erp_login_closure_report()
            try:
                driver.quit()
            except Exception:
                pass
            log_entry(job_name, frequency,
                      transferred_to=str(BASE_DIR / "outputs" / "LGD"),
                      status="Success")
        except Exception as exc:
            print(f"ERROR: {exc}")
            log_entry(job_name, frequency, status="Failed", notes=str(exc))


def job_all_contracts_sector():
    """All Contracts Sector Report"""
    job_name  = "All Contracts Sector Report"
    frequency = "Daily"
    with _job_log(job_name):
        try:
            sys.path.insert(0, str(BOTS_DIR))
            from All_Contract_Secter_Report import erp_all_contracts_sector_report
            ok = erp_all_contracts_sector_report()
            if not ok:
                raise RuntimeError("Bot returned False — download did not complete")
            log_entry(job_name, frequency,
                      transferred_to=str(BASE_DIR / "outputs" / "LGD"),
                      status="Success")
        except Exception as exc:
            print(f"ERROR: {exc}")
            log_entry(job_name, frequency, status="Failed", notes=str(exc))


def BackEndDueType_ContractWise():
    """Backend Due Type - Contract Wise"""
    job_name  = "Backend Due Type - Contract Wise"
    frequency = "Daily"
    with _job_log(job_name):
        try:
            sys.path.insert(0, str(BOTS_DIR))
            from BackEndDueType_ContractWise import erp_all_contracts_sector_report
            # Cumulative window (BACKEND_START_DATE .. BACKEND_END_DATE in the bot):
            # write_off.py rebuilds the whole 0130 working from this export.
            _drv = erp_all_contracts_sector_report()   # returns the open browser
            ok = bool(_drv)
            try:
                _drv.quit()            # close Chrome once the file has arrived
            except Exception:
                pass
            if not ok:
                raise RuntimeError("Bot returned False — download did not complete")
            time.sleep(15)   # let Chrome finish writing the file before looking in Downloads
            newest = _newest_file(DOWNLOAD_FOLDER_DIR, "Backend Due Types*.xlsx")
            if not newest:
                raise RuntimeError("Downloaded file not found in Downloads folder")
            # This report goes into firstDefault/, not LGD/
            lgd_sub = _find_lgd_outputs_firstDefault_subfolder()
            if not lgd_sub:
                raise RuntimeError("Could not resolve outputs/LGD/<date>/firstDefault/ folder")
            for old in lgd_sub.glob("Backend Due Types*.xlsx"):
                old.unlink()
            dest = lgd_sub / newest.name
            shutil.move(str(newest), str(dest))
            print(f"Moved {newest.name} → {dest}")
            log_entry(job_name, frequency,
                      downloaded_file=newest.name,
                      transferred_to=str(dest),
                      status="Success")
        except Exception as exc:
            print(f"ERROR: {exc}")
            log_entry(job_name, frequency, status="Failed", notes=str(exc))


def job_cbsl_provision_comparison():
    """CBSL Provision Comparison"""
    job_name  = "CBSL Provision Comparison"
    frequency = "Monthly"
    with _job_log(job_name):
        try:
            sys.path.insert(0, str(BASE_DIR))  # allow bare `from cbsl_provision_comparison import …`
            from cbsl_provision_comparison import ExcelAutomation, resolve_file_path
            file_path = resolve_file_path()  # re-resolved every run — the correct folder changes monthly
            automation = ExcelAutomation(file_path)
            ok = automation.run_all_steps(reporting_date=datetime.now())
            if not ok:
                raise RuntimeError("One or more steps failed — see job log for details")
            log_entry(job_name, frequency,
                      transferred_to=str(file_path),
                      status="Success")
        except Exception as exc:
            print(f"ERROR: {exc}")
            log_entry(job_name, frequency, status="Failed", notes=str(exc))


# ── Job registry ──────────────────────────────────────────────────────────────
# Each tuple: (job_function, CronTrigger kwargs, unique job id).
# Uncomment / add rows to enable more jobs. Adjust hour/minute to your schedule.

_JOBS: list[tuple] = [
    #(job_ledger_Accounts, {"hour": 9, "minute": 34}, "Ledger Details"),
    #(job_refinance,       {"hour": 11, "minute": 57}, "Refinance DPD"),
    # (job_summary_reports, {"hour": 9,  "minute": 48}, "Summary Reports"),
    #(job_legal_bal_new,           {"hour": 9,  "minute": 34},  "Legal Bal New"),
    #(job_all_contracts_sector,    {"hour": 9, "minute": 34},  "All Contracts Sector Report"),
    #(BackEndDueType_ContractWise, {"hour": 11, "minute": 35}, "Backend Due Type - Contract Wise"),
    #(job_cbsl_provision_comparison, {"hour": 9, "minute": 34}, "CBSL Provision Comparison"),

    # ── Weekly template (every Monday 08:00) ─────────────────────────────
    # (my_weekly_job, {"day_of_week": "mon", "hour": 8, "minute": 0}, "my_weekly"),

    # ── Monthly template (1st of month 06:30) ───────────────────────────
    # (my_monthly_job, {"day": 1, "hour": 6, "minute": 30}, "my_monthly"),
]


# ── Scheduler lifecycle ───────────────────────────────────────────────────────
_scheduler: BackgroundScheduler | None = None


def start_scheduler() -> BackgroundScheduler:
    """Initialise and start the APScheduler background scheduler."""
    global _scheduler
    _init_log_file()   # create scheduler_log.xlsx once if it doesn't exist
    _scheduler = BackgroundScheduler(timezone="Asia/Colombo")
    for func, trigger_kw, job_id in _JOBS:
        _scheduler.add_job(
            func,
            trigger=CronTrigger(**trigger_kw),
            id=job_id,
            replace_existing=True,
            misfire_grace_time=300,  # allow up to 5 min late start before skipping
        )
    _scheduler.start()
    return _scheduler


def stop_scheduler() -> None:
    """Gracefully stop the scheduler without waiting for running jobs to finish."""
    if _scheduler and _scheduler.running:
        _scheduler.shutdown(wait=False)


def get_scheduled_jobs() -> list[dict]:
    """Return current job metadata for the dashboard."""
    if _scheduler is None:
        return []
    result = []
    for job in _scheduler.get_jobs():
        next_run = job.next_run_time
        # Use the first line of the docstring as a human-readable name
        doc = (job.func.__doc__ or job.id).strip().splitlines()[0]
        result.append({
            "id":       job.id,
            "name":     doc,
            "next_run": next_run.strftime("%Y-%m-%d %H:%M") if next_run else "—",
        })
    return result


# ── Standalone entry point ────────────────────────────────────────────────────
if __name__ == "__main__":
    import time
    print("Scheduler starting (Ctrl-C to stop)…")
    start_scheduler()
    try:
        while True:
            time.sleep(60)
    except (KeyboardInterrupt, SystemExit):
        stop_scheduler()
        print("Scheduler stopped.")
