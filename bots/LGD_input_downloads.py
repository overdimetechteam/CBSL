"""
LGD input downloads - one entry point for the ERP (Selenium) bots the LGD run needs.

Triggered by ifrs.py as the first step of the IFRS-03 (LGD) run, using the run date
the user picked:

    python LGD_input_downloads.py DD/MM/YYYY [working run folder]

Bots run one after another (never in parallel, so they cannot pick up each
other's downloads) and every file is delivered into the LGD working run folder
(LGD_BASE_FOLDER, set by ifrs.py), replacing the copy carried forward from the
previous run:

    step                       window / date               delivered to
    -------------------------  --------------------------  ------------------------
    Ledger Details Report      01/04/2025 .. run date      <run folder>/LGD
    Backend Due Types          31/03/2025 .. run date      <run folder>/firstDefault
    Refinance DPD Report       01/04/2025 .. run date      <run folder>/LGD
    Legal Bal New              snapshot (no dates)         <run folder>/firstDefault
    All Contracts Sector       snapshot (no dates)         <run folder>/firstDefault

(Window starts are constants in each bot, overridable by environment variable.)

The first three are date-bound exports. After download, the Ledger and Backend
files are checked: their latest SLT_TRNDATE must fall in the run month, otherwise
the export is treated as stale. A failure in any of the first three stops the LGD run
(exit code 1) so it never silently runs on old data. Legal Bal New and All
Contracts Sector are snapshots: a failure there only prints a warning and the copy
already in the run folder is used.

Environment:
    LGD_BASE_FOLDER          working run folder (set by ifrs.py; or pass as 2nd arg)
    IFRS_LGD_LOG_DIR         if set, a LGD_input_downloads.log is written there
    LGD_INPUTS_STRICT=0      continue the LGD run even if a date-bound step fails
"""
import os
import sys
import time
import shutil
import traceback
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE_DIR = HERE.parent
for _p in (str(BASE_DIR), str(HERE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

DOWNLOADS = Path.home() / "Downloads"
DOWNLOAD_WAIT_S = 180          # how long to wait for a new file after a bot returns


# --------------------------------------------------------------------------
# logging: mirror stdout to <IFRS_LGD_LOG_DIR>/LGD_input_downloads.log
# --------------------------------------------------------------------------
def _setup_log() -> None:
    log_dir = os.environ.get("IFRS_LGD_LOG_DIR", "").strip()
    if not log_dir:
        return
    try:
        Path(log_dir).mkdir(parents=True, exist_ok=True)
        fh = open(Path(log_dir) / "LGD_input_downloads.log", "a", encoding="utf-8", errors="replace")
    except OSError:
        return

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

    sys.stdout = _Tee(sys.stdout)
    sys.stderr = _Tee(sys.stderr)


def section(title: str) -> None:
    print("\n" + "=" * 65 + f"\n  {title}\n" + "=" * 65, flush=True)


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def parse_run_date(text: str) -> datetime:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text.strip(), fmt)
        except ValueError:
            continue
    raise ValueError(f"Unsupported run date: {text!r} (expected DD/MM/YYYY)")


def _matches(folder: Path, pattern: str) -> list:
    return [f for f in folder.glob(pattern)
            if not f.name.startswith("~$") and f.suffix.lower() != ".crdownload"]


def stash_stale(pattern: str, stash_dir: Path) -> None:
    """
    Move any file already in Downloads that matches the pattern out of the way.
    Several bots accept "any matching file in Downloads" as the finished download,
    so a leftover from an earlier run would be mistaken for the new one.
    Moved (not deleted) so nothing is lost.
    """
    for f in _matches(DOWNLOADS, pattern):
        stash_dir.mkdir(parents=True, exist_ok=True)
        target = stash_dir / f.name
        n = 1
        while target.exists():
            target = stash_dir / f"{f.stem}_{n}{f.suffix}"
            n += 1
        shutil.move(str(f), str(target))
        print(f"  Moved stale Downloads file aside: {f.name} -> {target}")


def wait_for_new(folder: Path, pattern: str, since: float, timeout: int = DOWNLOAD_WAIT_S) -> Path:
    """Wait for a finished file matching pattern, modified since `since`."""
    end = time.time() + timeout
    while time.time() < end:
        fresh = [f for f in _matches(folder, pattern) if f.stat().st_mtime >= since - 2]
        if fresh and not list(folder.glob(pattern + ".crdownload")):
            return max(fresh, key=lambda f: f.stat().st_mtime)
        time.sleep(2)
    raise RuntimeError(f"No new file matching '{pattern}' appeared in {folder} within {timeout}s")


def place(file: Path, dest_dir: Path, pattern: str) -> Path:
    """Move the downloaded file into dest_dir and remove older copies of it there."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / file.name
    for old in _matches(dest_dir, pattern):
        if old.resolve() != dest.resolve():
            old.unlink()
            print(f"  Removed previous copy: {old.name}")
    if dest.exists():
        dest.unlink()
    shutil.move(str(file), str(dest))
    print(f"  Delivered: {dest}")
    return dest


def check_latest_date(path: Path, column: str, run_dt: datetime, label: str, strict: bool) -> None:
    """The export must contain transactions in the run month, otherwise it is stale."""
    import pandas as pd
    try:
        df = pd.read_excel(path, usecols=[column])
    except Exception as exc:
        msg = f"{label}: could not read '{column}' to check dates ({exc})"
        if strict:
            raise RuntimeError(msg)
        print(f"  [WARN] {msg}")
        return
    latest = pd.to_datetime(df[column], errors="coerce").max()
    month_start = run_dt.replace(day=1)
    print(f"  {label}: latest {column} = {latest.date() if pd.notna(latest) else 'n/a'} "
          f"(run date {run_dt.date()}, month starts {month_start.date()})")
    if pd.isna(latest) or latest < month_start:
        msg = (f"{label}: latest {column} is {latest.date() if pd.notna(latest) else 'blank'}, "
               f"before the run month ({month_start:%B %Y}) - the export looks stale")
        if strict:
            raise RuntimeError(msg)
        print(f"  [WARN] {msg}")


def _quit(driver) -> None:
    try:
        if driver is not None and hasattr(driver, "quit"):
            driver.quit()
    except Exception:
        pass


# --------------------------------------------------------------------------
# steps. Each takes (run_dt, base) and raises on failure.
# --------------------------------------------------------------------------
def step_ledger(run_dt: datetime, base: Path, stash: Path) -> None:
    import DM_Ledger_Accounts_main as bot
    pattern = "Ledger Details Report - Internal*.xlsx"
    stash_stale(pattern, stash)
    t0 = time.time()
    ok = bot.Download_Ledger_Accounts_report(end_date=run_dt.strftime("%d/%m/%Y"))
    if not ok:
        raise RuntimeError("Ledger Details bot reported that the download did not complete")
    got = wait_for_new(DOWNLOADS, pattern, t0)
    dest = place(got, base / "LGD", pattern)
    check_latest_date(dest, "SLT_TRNDATE", run_dt, "Ledger Details Report", strict=True)


def step_backend(run_dt: datetime, base: Path, stash: Path) -> None:
    import BackEndDueType_ContractWise as bot
    pattern = "Backend Due Types - Contract wise Report - Internal*.xlsx"
    stash_stale(pattern, stash)
    t0 = time.time()
    driver = None
    try:
        driver = bot.erp_all_contracts_sector_report(to_date=run_dt.strftime("%Y-%m-%d"))
        got = wait_for_new(DOWNLOADS, pattern, t0)
    finally:
        _quit(driver)
    dest = place(got, base / "firstDefault", pattern)
    check_latest_date(dest, "SLT_TRNDATE", run_dt, "Backend Due Types", strict=True)


def step_refinance(run_dt: datetime, base: Path, stash: Path) -> None:
    import LGD_refinance_dpd as bot
    pattern = "Refinance DPD Report *.xlsx"
    stash_stale(pattern, stash)
    t0 = time.time()
    ok = bot.run(run_dt.strftime("%d/%m/%Y"))
    if not ok:
        raise RuntimeError("Refinance DPD bot reported that the download did not complete")
    got = wait_for_new(DOWNLOADS, pattern, t0)
    dest = place(got, base / "LGD", pattern)
    # Informational only: there may legitimately be no refinances late in the month.
    check_latest_date(dest, "NEW_CON_NO_ACTIVATION_DATE", run_dt, "Refinance DPD Report", strict=False)


def _bot_places_itself(module_name: str, func_name: str, pattern: str, dest_dir: Path) -> None:
    """For bots that move their own download into LGD_BASE_FOLDER/firstDefault."""
    module = __import__(module_name)
    t0 = time.time()
    driver = None
    try:
        driver = getattr(module, func_name)()
    finally:
        _quit(driver if not isinstance(driver, bool) else None)
    got = wait_for_new(dest_dir, pattern, t0, timeout=30)
    print(f"  Delivered: {got}")


def step_legal_bal(run_dt: datetime, base: Path, stash: Path) -> None:
    _bot_places_itself("Legal_Bal_New", "erp_login_closure_report",
                       "Legal Bal New  - Internal *.xlsx", base / "firstDefault")


def step_all_contracts(run_dt: datetime, base: Path, stash: Path) -> None:
    _bot_places_itself("All_Contract_Secter_Report", "erp_all_contracts_sector_report",
                       "All Contracts Sector Report - Internal *.xlsx", base / "firstDefault")


# (label, function, stops the LGD run on failure?)
STEPS = [
    ("Ledger Details Report", step_ledger, True),
    ("Backend Due Types", step_backend, True),
    ("Refinance DPD Report", step_refinance, True),
    ("Legal Bal New", step_legal_bal, False),
    ("All Contracts Sector Report", step_all_contracts, False),
]


def main(argv: list) -> int:
    _setup_log()
    if len(argv) < 2:
        print("Usage: python LGD_input_downloads.py DD/MM/YYYY [working run folder]")
        return 1
    run_dt = parse_run_date(argv[1])
    base_text = (argv[2] if len(argv) > 2 else os.environ.get("LGD_BASE_FOLDER", "")).strip()
    if not base_text:
        print("ERROR: no working run folder (LGD_BASE_FOLDER not set and none passed).")
        return 1
    base = Path(base_text)
    os.environ["LGD_BASE_FOLDER"] = str(base)       # the bots read this too
    strict = os.environ.get("LGD_INPUTS_STRICT", "1").strip() != "0"

    section(f"LGD input downloads - run date {run_dt:%d/%m/%Y}")
    print(f"  Run folder : {base}")
    print(f"  Strict     : {strict} (a failed date-bound download stops the LGD run)")
    stash = (Path(os.environ["IFRS_LGD_LOG_DIR"]) / "stale_downloads"
             if os.environ.get("IFRS_LGD_LOG_DIR") else BASE_DIR / "logs" / "stale_downloads")

    results = []
    for label, fn, stops_run in STEPS:
        section(f"STEP : {label}")
        t0 = time.time()
        try:
            fn(run_dt, base, stash)
            results.append((label, "OK", stops_run, ""))
            print(f"  -> OK ({time.time() - t0:.0f}s)")
        except Exception as exc:
            results.append((label, "FAILED", stops_run, str(exc)))
            print(f"  -> FAILED ({time.time() - t0:.0f}s): {exc}")
            print(traceback.format_exc())

    section("SUMMARY")
    for label, status, stops_run, msg in results:
        kind = "date-bound" if stops_run else "snapshot  "
        print(f"  [{status:6}] {kind} {label}" + (f"  - {msg}" if msg else ""))
    blocking = [r for r in results if r[1] == "FAILED" and r[2]]
    if blocking and strict:
        print("\n  ERROR: a date-bound input could not be refreshed - stopping the LGD run so it "
              "does not run on old data. (Set LGD_INPUTS_STRICT=0 to continue anyway.)")
        return 1
    if any(r[1] == "FAILED" for r in results):
        print("\n  [WARN] Some inputs were not refreshed; the LGD run will use the copies already "
              "in the run folder.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
