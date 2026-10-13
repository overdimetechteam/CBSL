"""
supporting_reports_main.py
--------------------------
Handles all Supporting Reports UI routes and script execution.
Registered into app.py via register_supporting_routes(app).

Reports (in priority order):
  1. Portfolio 4 ECL + ECL Model + Summary Sheet  (one script, one step)
  2. Deferred Promotion Calculation
  3. CBSL Provision Comparison
  4. Supporting Schedules
  5. Unutilized Credit Limits
  6. Disbursement Report
  7. Collection Efficiency Report
"""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from collections import deque
from datetime import datetime, timedelta
from pathlib import Path
from threading import Lock

from flask import (
    Blueprint,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

# ---------------------------------------------------------------------------
# Blueprint
# ---------------------------------------------------------------------------
sup_bp = Blueprint("sup", __name__, url_prefix="/supporting")

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve().parent
OUTPUTS_DIR     = _HERE / "outputs"
WORKING_DIR     = _HERE / "working"
SCRIPTS_DIR     = _HERE / "report_automations"
LOGS_DIR        = _HERE / "logs"

# ---------------------------------------------------------------------------
# Report definitions
# ---------------------------------------------------------------------------
# Each entry:
#   id          – unique key
#   label       – display name shown in UI
#   description – short help text
#   script      – filename in report_automations/ (None = not yet automated)
#   date_input  – "date" | "month_year" | "none"
#   group       – used to visually group items  (None = standalone)
#   group_label – heading for the group
SUPPORTING_REPORTS = [
    {
        "id": "ECL_GROUP",
        "label": "Portfolio 4 ECL / ECL Model / Summary Sheet",
        "description": "Generates Portfolio 4 ECL, ECL Model and Summary Sheet in a single run.",
        "script": "supporting_ecl_summary.py",
        "date_input": "month_year",
        "group": "ecl",
        "group_label": "ECL & Summary",
        "priority": 1,
    },
    {
        "id": "DEFERRED_PROMO",
        "label": "Deferred Promotion Calculation",
        "description": "Calculates deferred promotion figures.",
        "script": "supporting_deferred_promo.py",
        "date_input": "month_year",
        "group": None,
        "group_label": None,
        "priority": 2,
    },
    {
        "id": "CBSL_PROVISION",
        "label": "CBSL Provision Comparison",
        "description": "Compares CBSL provision figures.",
        "script": "cbsl_provision_comparison.py",
        "date_input": "month_year",
        "group": None,
        "group_label": None,
        "priority": 3,
    },
    {
        "id": "SUPPORTING_SCHEDULES",
        "label": "Supporting Schedules",
        "description": "Generates supporting schedule workbooks.",
        "script": "Supporting_Schedules.py",
        "date_input": "month_year",
        "group": None,
        "group_label": None,
        "priority": 4,
    },
    {
        "id": "UNUTILIZED_CREDIT",
        "label": "Unutilized Credit Limits",
        "description": "Reports on unutilized credit limit balances.",
        "script": "unutilized_credit_limits.py",
        "date_input": "month_year",
        "group": None,
        "group_label": None,
        "priority": 5,
    },
    {
        "id": "DISBURSEMENT",
        "label": "Disbursement Report",
        "description": "Generates the disbursement report.",
        "script": "supporting_disbursement.py",
        "date_input": "month_year",
        "group": None,
        "group_label": None,
        "priority": 6,
    },
    {
        "id": "COLLECTION_EFFICIENCY",
        "label": "Collection Efficiency Report",
        "description": "Calculates collection efficiency metrics.",
        "script": "supporting_collection_efficiency.py",
        "date_input": "month_year",
        "group": None,
        "group_label": None,
        "priority": 7,
    },
]

# ---------------------------------------------------------------------------
# Per-module state  (isolated from main app.py state)
# ---------------------------------------------------------------------------
_state_lock = Lock()
_state: dict = {
    "running": False,
    "stage": "idle",        # idle | running | done | error | stopping
    "current_report": "",
    "message": "",
    "ss_download_ready": False,
    "cbsl_download_ready": False,
    "ucl_download_ready": False,
}

_log_lock = Lock()
_log_buffer: deque[str] = deque(maxlen=1000)

_checklist_lock = Lock()
_checklist: list[dict] = []       # [{id, label, status, elapsed}]

_stop_flag = False
_running_proc: subprocess.Popen | None = None
_download_file: Path | None = None         # set after SUPPORTING_SCHEDULES completes
_cbsl_download_file: Path | None = None    # set after CBSL_PROVISION completes
_ucl_download_file: Path | None = None     # set after UNUTILIZED_CREDIT completes
# Pending run info for CBSL_PROVISION (set during upload, consumed by run thread)
_pending_cbsl_run: dict | None = None      # {"reports": [...], "date_str": "YYYY-MM-DD"}

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _emit(msg: str) -> None:
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
    with _log_lock:
        _log_buffer.append(line)


def _set_state(**kwargs) -> None:
    with _state_lock:
        _state.update(kwargs)


def _set_checklist_status(report_id: str, status: str) -> None:
    with _checklist_lock:
        for item in _checklist:
            if item["id"] == report_id:
                item["status"] = status
                if status == "running":
                    item["start_time"] = time.time()
                elif status in ("done", "failed", "stopped"):
                    if item.get("start_time"):
                        item["elapsed"] = int(time.time() - item["start_time"])
                break


def _reset() -> None:
    global _stop_flag, _running_proc, _download_file, _cbsl_download_file, _ucl_download_file
    _stop_flag = False
    _running_proc = None
    _download_file = None
    _cbsl_download_file = None
    _ucl_download_file = None
    with _log_lock:
        _log_buffer.clear()
    with _checklist_lock:
        _checklist.clear()
    _set_state(running=False, stage="idle", current_report="", message="",
               ss_download_ready=False, cbsl_download_ready=False, ucl_download_ready=False)


def _get_report(report_id: str) -> dict | None:
    for r in SUPPORTING_REPORTS:
        if r["id"] == report_id:
            return r
    return None


def _last_run_info(report_id: str) -> str:
    """Return a human-readable last-run date string from outputs, if available."""
    out_folder = OUTPUTS_DIR / report_id.lower()
    dates_file = out_folder / "completed_dates.txt"
    try:
        if dates_file.exists():
            lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
            if lines:
                return lines[-1]
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------------------
# Supporting Schedules — preparation helpers
# ---------------------------------------------------------------------------

def _find_date_folder_with_variations(parent: Path, date_str: str) -> Path | None:
    """
    Return the highest-version folder whose name starts with date_str (DD-MM-YYYY).
    Handles plain date folder and versioned ones: DD-MM-YYYY, DD-MM-YYYY(1), ...
    """
    if not parent.exists():
        return None

    date_variants: set[str] = {date_str}
    for fmt_in, fmt_out in [("%d-%m-%Y", "%d/%m/%Y"), ("%d/%m/%Y", "%d-%m-%Y")]:
        try:
            date_variants.add(datetime.strptime(date_str, fmt_in).strftime(fmt_out))
        except ValueError:
            pass

    candidates: list[tuple[int, Path]] = []
    for child in parent.iterdir():
        if not child.is_dir():
            continue
        if any(child.name.startswith(v) for v in date_variants):
            m = re.search(r'\((\d+)\)$', child.name)
            candidates.append((int(m.group(1)) if m else 0, child))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def _find_latest_outputs_folder(outputs_subdir: Path) -> Path | None:
    """
    Read the last line of completed_dates.txt in outputs_subdir and return
    the highest-version matching date folder, or None if not found.
    """
    dates_file = outputs_subdir / "completed_dates.txt"
    if not dates_file.exists():
        _emit(f"[WARN] No completed_dates.txt in {outputs_subdir}")
        return None

    lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
    if not lines:
        _emit(f"[WARN] completed_dates.txt is empty in {outputs_subdir}")
        return None

    last_date_str = lines[-1]
    _emit(f"Latest completed date: {last_date_str}")

    folder = _find_date_folder_with_variations(outputs_subdir, last_date_str)
    if folder:
        _emit(f"Source folder: {folder}")
    else:
        _emit(f"[WARN] Date folder not found for {last_date_str} in {outputs_subdir}")
    return folder


def _prepare_supporting_schedules(run_date_iso: str) -> bool:
    """
    Prepare the working directory for the Supporting Schedules report:
      1. Run GL_CodeBase_Report (GL 1247) from start_date (run_date - 30 days) to run_date,
         download and move the file to the latest Supporting Schedules outputs folder.
      2. Find the latest outputs folder (outputs/Supporting Schedules/<latest_date>)
      3. Clean working/Supporting Schedules  (delete all contents)
      4. Copy the outputs folder into working/Supporting Schedules/<folder_name>
    Returns True on success, False on failure.
    """
    outputs_subdir = OUTPUTS_DIR / "Supporting Schedules"
    working_subdir = WORKING_DIR / "Supporting Schedules"

    # ── Step 1: Run GL_CodeBase_Report for GL 1247 ───────────────────────────
    try:
        run_dt = datetime.strptime(run_date_iso, "%Y-%m-%d")
    except ValueError:
        _emit(f"[ERROR] Invalid run date: {run_date_iso}")
        return False

    start_dt = run_dt - timedelta(days=30)
    start_date_str = start_dt.strftime("%d/%m/%Y")
    end_date_str   = run_dt.strftime("%d/%m/%Y")

    _emit(f"Running GL CodeBase Report — GL 1247, {start_date_str} → {end_date_str} …")
    try:
        import sys as _sys
        _bots_dir = str(_HERE / "bots")
        if _bots_dir not in _sys.path:
            _sys.path.insert(0, _bots_dir)
        import GL_CodeBase_Report as _gl_bot
        _gl_bot.run(
            start_date_str=start_date_str,
            end_date_str=end_date_str,
            gl_code="1247",
            dest_outputs_name="Supporting Schedules",
        )
        _emit("[OK] GL CodeBase Report completed — file moved to Supporting Schedules outputs.")
    except Exception as e:
        _emit(f"[ERROR] GL CodeBase Report failed: {e}")
        return False

    # ── Step 2: Prepare working directory ────────────────────────────────────
    _emit("Preparing Supporting Schedules working directory…")

    src_folder = _find_latest_outputs_folder(outputs_subdir)
    if not src_folder:
        _emit("[ERROR] Cannot find latest outputs folder for Supporting Schedules.")
        return False

    # Step 3: clean working/Supporting Schedules
    if working_subdir.exists():
        _emit(f"Cleaning {working_subdir} …")
        try:
            shutil.rmtree(working_subdir)
        except Exception as e:
            _emit(f"[ERROR] Failed to clean working folder: {e}")
            return False
    working_subdir.mkdir(parents=True, exist_ok=True)

    # Step 4: copy src_folder into working/Supporting Schedules/<same name>
    dest_folder = working_subdir / src_folder.name
    _emit(f"Copying {src_folder.name} → {dest_folder} …")
    try:
        shutil.copytree(str(src_folder), str(dest_folder))
    except Exception as e:
        _emit(f"[ERROR] Failed to copy folder: {e}")
        return False

    _emit("Working directory prepared successfully.")
    return True


def _copy_ss_to_sofp(ss_file: Path, sofp_name: str) -> None:
    """Replace Supporting Schedules - *.xlsx in the latest date folder of a SOFP outputs dir."""
    sofp_outputs = OUTPUTS_DIR / sofp_name
    sofp_folder = _find_latest_outputs_folder(sofp_outputs)
    if not sofp_folder:
        _emit(f"[WARN] No latest folder found in outputs/{sofp_name} — skipping copy.")
        return
    # Remove existing Supporting Schedules - *.xlsx files
    for old in sofp_folder.glob("Supporting Schedules - *.xlsx"):
        try:
            old.unlink()
            _emit(f"Removed existing: outputs/{sofp_name}/{sofp_folder.name}/{old.name}")
        except Exception as e:
            _emit(f"[WARN] Could not remove {old.name}: {e}")
    try:
        shutil.copy2(str(ss_file), str(sofp_folder / ss_file.name))
        _emit(f"[OK] Copied {ss_file.name} → outputs/{sofp_name}/{sofp_folder.name}/")
    except Exception as e:
        _emit(f"[WARN] Failed to copy to outputs/{sofp_name}: {e}")


def _post_supporting_schedules(run_date_iso: str) -> bool:
    """
    After Supporting_Schedules.py completes successfully:
      1. Find the single date subfolder inside working/Supporting Schedules/
      2. Locate 'Supporting Schedules - *.xlsx' in that folder (for download)
      3. Rename that subfolder to DD-MM-YYYY (the selected run date)
      4. Copy renamed folder to outputs/Supporting Schedules/<DD-MM-YYYY>/
         (versioned if folder already exists: DD-MM-YYYY(1), (2), …)
      5. Append DD-MM-YYYY to outputs/Supporting Schedules/completed_dates.txt
      6. Replace 'Supporting Schedules - *.xlsx' in latest NBD_MF_01_SOFP_SOCI outputs folder
      7. Replace 'Supporting Schedules - *.xlsx' in latest NBD_AF_01_SOFP_SOCI outputs folder
      8. Clean working/Supporting Schedules/
    Returns True on success, False on failure.
    """
    global _download_file

    working_subdir = WORKING_DIR / "Supporting Schedules"
    outputs_subdir = OUTPUTS_DIR / "Supporting Schedules"

    # Convert run date to DD-MM-YYYY folder name
    try:
        run_dt = datetime.strptime(run_date_iso, "%Y-%m-%d")
    except ValueError:
        _emit(f"[ERROR] Invalid run date for post-processing: {run_date_iso}")
        return False
    date_folder_name = run_dt.strftime("%d-%m-%Y")

    # ── Step 1: find the single date subfolder in working ────────────────────
    subdirs = [d for d in working_subdir.iterdir() if d.is_dir()] if working_subdir.exists() else []
    if len(subdirs) != 1:
        _emit(f"[ERROR] Expected 1 subfolder in {working_subdir}, found {len(subdirs)}.")
        return False
    work_date_folder = subdirs[0]

    # ── Step 2: locate the SS xlsx in working (for download) ─────────────────
    ss_matches = [f for f in work_date_folder.glob("Supporting Schedules - *.xlsx")
                  if not f.name.startswith("~$")]
    if not ss_matches:
        _emit("[WARN] No 'Supporting Schedules - *.xlsx' found in working folder.")
        ss_file_working = None
    else:
        ss_file_working = ss_matches[0]
        _download_file = ss_file_working
        _emit(f"Download file set: {ss_file_working.name}")

    # ── Step 3: rename working subfolder to DD-MM-YYYY ───────────────────────
    renamed_folder = working_subdir / date_folder_name
    if work_date_folder != renamed_folder:
        _emit(f"Renaming {work_date_folder.name} → {date_folder_name} …")
        try:
            work_date_folder.rename(renamed_folder)
            work_date_folder = renamed_folder
            # Update download path after rename
            if ss_file_working:
                _download_file = work_date_folder / ss_file_working.name
        except Exception as e:
            _emit(f"[ERROR] Failed to rename working folder: {e}")
            return False

    # ── Step 4: copy to outputs/Supporting Schedules/<DD-MM-YYYY[(/n)]> ─────
    outputs_subdir.mkdir(parents=True, exist_ok=True)
    dest_name = date_folder_name
    version = 0
    while (outputs_subdir / dest_name).exists():
        version += 1
        dest_name = f"{date_folder_name}({version})"
    dest_folder = outputs_subdir / dest_name

    _emit(f"Copying {work_date_folder.name} → outputs/Supporting Schedules/{dest_name} …")
    try:
        shutil.copytree(str(work_date_folder), str(dest_folder))
    except Exception as e:
        _emit(f"[ERROR] Failed to copy to Supporting Schedules outputs: {e}")
        return False

    # ── Step 5: append to completed_dates.txt ────────────────────────────────
    dates_file = outputs_subdir / "completed_dates.txt"
    try:
        with open(dates_file, "a", encoding="utf-8") as fh:
            fh.write(date_folder_name + "\n")
        _emit(f"Appended {date_folder_name} to completed_dates.txt")
    except Exception as e:
        _emit(f"[WARN] Could not update completed_dates.txt: {e}")

    # ── Steps 6 & 7: copy to NBD_MF_01_SOFP_SOCI and NBD_AF_01_SOFP_SOCI ───
    out_matches = [f for f in dest_folder.glob("Supporting Schedules - *.xlsx")
                   if not f.name.startswith("~$")]
    if out_matches:
        ss_file = out_matches[0]
        # Point download file to the outputs copy (working folder will be cleaned)
        _download_file = ss_file
        _emit(f"Download file updated to outputs copy: {ss_file.name}")
        _copy_ss_to_sofp(ss_file, "NBD_MF_01_SOFP_SOCI")
        _copy_ss_to_sofp(ss_file, "NBD_AF_01_SOFP_SOCI")
    else:
        _emit("[WARN] No 'Supporting Schedules - *.xlsx' in outputs folder — skipping SOFP copies.")

    # ── Step 8: clean working/Supporting Schedules/ ──────────────────────────
    _emit(f"Cleaning working folder {working_subdir} …")
    try:
        shutil.rmtree(working_subdir)
        working_subdir.mkdir(parents=True, exist_ok=True)
        _emit("Working folder cleaned.")
    except Exception as e:
        _emit(f"[WARN] Failed to clean working folder: {e}")

    _emit(f"[OK] Supporting Schedules post-processing complete → outputs/Supporting Schedules/{dest_name}/")
    return True


# ---------------------------------------------------------------------------
# CBSL Provision Comparison — helpers
# ---------------------------------------------------------------------------

MASTER_INPUTS_DIR = _HERE / "Master Inputs"


def _get_latest_property_mortgages_name() -> str:
    """Return the filename of the latest Property Mortgages file, or empty string."""
    mi_dir = MASTER_INPUTS_DIR / "Property mortgages"
    if not mi_dir.exists():
        return ""
    files = list(mi_dir.glob("Property mortgages - *.xlsx")) + list(mi_dir.glob("Property mortgages - *.xls"))
    if not files:
        return ""
    return max(files, key=lambda f: f.stat().st_mtime).name


def _find_cbsl_latest_outputs_folder() -> Path | None:
    """
    Find the latest date folder in outputs/CBSL Provision Comparison.
    First tries completed_dates.txt; if missing, scans the folder directly for
    the highest-version DD-MM-YYYY[(/n)] subfolder and repairs completed_dates.txt.
    """
    outputs_subdir = OUTPUTS_DIR / "CBSL Provision Comparison"
    dates_file = outputs_subdir / "completed_dates.txt"

    # ── Primary: use completed_dates.txt ────────────────────────────────────
    if dates_file.exists():
        lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
        if lines:
            last_date_str = lines[-1]
            folder = _find_date_folder_with_variations(outputs_subdir, last_date_str)
            if folder:
                _emit(f"Latest CBSL outputs folder (from completed_dates.txt): {folder.name}")
                return folder

    # ── Fallback: scan folder for DD-MM-YYYY[(/n)] subfolders ───────────────
    _emit("[WARN] completed_dates.txt missing or stale — scanning outputs folder directly…")
    date_pattern = re.compile(r'^(\d{2}-\d{2}-\d{4})(\(\d+\))?$')
    candidates: list[tuple[str, int, Path]] = []
    if outputs_subdir.exists():
        for child in outputs_subdir.iterdir():
            if not child.is_dir():
                continue
            m = date_pattern.match(child.name)
            if m:
                date_part = m.group(1)
                version = int(m.group(2)[1:-1]) if m.group(2) else 0
                candidates.append((date_part, version, child))

    if not candidates:
        _emit("[WARN] No date folders found in outputs/CBSL Provision Comparison.")
        return None

    # Sort by date then version, pick the latest
    candidates.sort(key=lambda x: (datetime.strptime(x[0], "%d-%m-%Y"), x[1]))
    latest_date_str, _, latest_folder = candidates[-1]
    _emit(f"Latest CBSL outputs folder (scanned): {latest_folder.name}")

    # Repair completed_dates.txt so future runs use the primary path
    try:
        outputs_subdir.mkdir(parents=True, exist_ok=True)
        with open(dates_file, "w", encoding="utf-8") as fh:
            fh.write(latest_date_str + "\n")
        _emit(f"[OK] Repaired completed_dates.txt → {latest_date_str}")
    except Exception as e:
        _emit(f"[WARN] Could not write completed_dates.txt: {e}")

    return latest_folder


def _prepare_cbsl_provision(run_date_iso: str) -> bool:
    """
    Prepare the working directory for the CBSL Provision Comparison report:
      1. Find the latest outputs/CBSL Provision Comparison date folder.
      2. Clean working/CBSL Provision Comparison (remove all contents).
      3. Copy the *contents* of the outputs date folder directly into
         working/CBSL Provision Comparison/ (no nested date subfolder).
         The script receives working/CBSL Provision Comparison/ as its base_path.
    Returns True on success, False on failure.
    """
    working_subdir = WORKING_DIR / "CBSL Provision Comparison"

    _emit("Preparing CBSL Provision Comparison working directory…")

    src_folder = _find_cbsl_latest_outputs_folder()
    if not src_folder:
        _emit("[ERROR] Cannot find latest outputs folder for CBSL Provision Comparison.")
        return False

    # Clean working folder entirely
    if working_subdir.exists():
        _emit(f"Cleaning {working_subdir} …")
        try:
            shutil.rmtree(working_subdir)
        except Exception as e:
            _emit(f"[ERROR] Failed to clean working folder: {e}")
            return False
    working_subdir.mkdir(parents=True, exist_ok=True)

    # Copy the entire outputs date folder into working (same pattern as Supporting Schedules)
    dest_folder = working_subdir / src_folder.name
    _emit(f"Copying {src_folder.name} → working/CBSL Provision Comparison/{src_folder.name} …")
    try:
        shutil.copytree(str(src_folder), str(dest_folder))
    except Exception as e:
        _emit(f"[ERROR] Failed to copy folder: {e}")
        return False

    _emit("CBSL Provision Comparison working directory prepared successfully.")
    return True


def _copy_cbsl_to_outputs(cbsl_file: Path, dest_outputs_name: str,
                           dest_subdir: str | None = None) -> None:
    """
    Replace 'CBSL Provision Comparison - *.xlsb' in the latest date folder of
    outputs/<dest_outputs_name>[/<dest_subdir>] with cbsl_file.
    dest_subdir is used for paths like outputs/NBD-QF-23-C3-C10-SF-FI/<date>/Input
    """
    base_outputs = OUTPUTS_DIR / dest_outputs_name
    latest = _find_latest_outputs_folder(base_outputs)
    if not latest:
        _emit(f"[WARN] No latest folder found in outputs/{dest_outputs_name} — skipping copy.")
        return
    target_dir = latest / dest_subdir if dest_subdir else latest
    target_dir.mkdir(parents=True, exist_ok=True)
    for old in target_dir.glob("CBSL Provision Comparison - *.xlsb"):
        try:
            old.unlink()
            _emit(f"Removed existing: {dest_outputs_name}/{latest.name}/{dest_subdir or ''}/{old.name}")
        except Exception as e:
            _emit(f"[WARN] Could not remove {old.name}: {e}")
    try:
        shutil.copy2(str(cbsl_file), str(target_dir / cbsl_file.name))
        _emit(f"[OK] Copied {cbsl_file.name} → outputs/{dest_outputs_name}/{latest.name}/{dest_subdir or ''}/")
    except Exception as e:
        _emit(f"[WARN] Failed to copy to outputs/{dest_outputs_name}: {e}")


def _post_cbsl_provision(run_date_iso: str) -> bool:
    """
    After cbsl_provision_comparison.py completes:
      1. Find the single date subfolder in working/CBSL Provision Comparison/.
      2. Locate 'CBSL Provision Comparison - *.xlsb' inside that subfolder.
      3. Rename subfolder to DD-MM-YYYY.
      4. Set download file + distribute to 5 output folders from working copy.
      5. Copy renamed subfolder → outputs/CBSL Provision Comparison/<DD-MM-YYYY[(/n)]>.
      6. Point download file to outputs copy.
      7. Clean working/CBSL Provision Comparison/.
    """
    global _cbsl_download_file

    working_subdir = WORKING_DIR / "CBSL Provision Comparison"
    outputs_subdir = OUTPUTS_DIR / "CBSL Provision Comparison"

    try:
        run_dt = datetime.strptime(run_date_iso, "%Y-%m-%d")
    except ValueError:
        _emit(f"[ERROR] Invalid run date for post-processing: {run_date_iso}")
        return False
    date_folder_name = run_dt.strftime("%d-%m-%Y")

    # Step 1: find the single date subfolder in working
    subdirs = [d for d in working_subdir.iterdir() if d.is_dir()] if working_subdir.exists() else []
    if len(subdirs) != 1:
        _emit(f"[ERROR] Expected 1 subfolder in {working_subdir}, found {len(subdirs)}.")
        return False
    work_date_folder = subdirs[0]

    # Step 2: locate the CBSL xlsb inside the subfolder
    cbsl_matches = [f for f in work_date_folder.glob("CBSL Provision Comparison - *.xlsb")
                    if not f.name.startswith("~$")]
    if not cbsl_matches:
        _emit("[WARN] No 'CBSL Provision Comparison - *.xlsb' found in working subfolder.")
        cbsl_file_working = None
    else:
        cbsl_file_working = cbsl_matches[0]
        _cbsl_download_file = cbsl_file_working
        _emit(f"Download file set: {cbsl_file_working.name}")

    # Step 3: rename working subfolder to DD-MM-YYYY
    renamed_folder = working_subdir / date_folder_name
    if work_date_folder != renamed_folder:
        _emit(f"Renaming {work_date_folder.name} → {date_folder_name} …")
        try:
            work_date_folder.rename(renamed_folder)
            work_date_folder = renamed_folder
            if cbsl_file_working:
                cbsl_file_working = work_date_folder / cbsl_file_working.name
        except Exception as e:
            _emit(f"[ERROR] Failed to rename working subfolder: {e}")
            return False

    # Step 4: set download file and distribute to 5 output folders from working copy
    if cbsl_file_working and cbsl_file_working.exists():
        _cbsl_download_file = cbsl_file_working
        _emit(f"Download file: {cbsl_file_working.name}")
        _copy_cbsl_to_outputs(cbsl_file_working, "NBD_AF_20_C1_C6")
        _copy_cbsl_to_outputs(cbsl_file_working, "NBD_MF_20_C1_C6")
        _copy_cbsl_to_outputs(cbsl_file_working, "NBD_MF_23_IA")
        _copy_cbsl_to_outputs(cbsl_file_working, "NBD-MF-10-GA & NBD-MF-11-IS")
        _copy_cbsl_to_outputs(cbsl_file_working, "NBD-QF-23-C3-C10-SF-FI", dest_subdir="Input")
        _copy_cbsl_to_outputs(cbsl_file_working, "Supporting Schedules")
    else:
        _emit("[WARN] No 'CBSL Provision Comparison - *.xlsb' found — skipping distribution.")

    # Step 5: copy working subfolder → outputs/CBSL Provision Comparison/<DD-MM-YYYY[(/n)]>
    outputs_subdir.mkdir(parents=True, exist_ok=True)
    dest_name = date_folder_name
    version = 0
    while (outputs_subdir / dest_name).exists():
        version += 1
        dest_name = f"{date_folder_name}({version})"
    dest_folder = outputs_subdir / dest_name

    _emit(f"Copying {work_date_folder.name} → outputs/CBSL Provision Comparison/{dest_name} …")
    try:
        shutil.copytree(str(work_date_folder), str(dest_folder))
    except Exception as e:
        _emit(f"[ERROR] Failed to copy to CBSL Provision Comparison outputs: {e}")
        return False

    # Step 6: point download file to outputs copy (working is about to be cleaned)
    out_matches = [f for f in dest_folder.glob("CBSL Provision Comparison - *.xlsb")
                   if not f.name.startswith("~$")]
    if out_matches:
        _cbsl_download_file = out_matches[0]
        _emit(f"Download file updated to outputs copy: {_cbsl_download_file.name}")

    # Step 7: clean working folder
    _emit(f"Cleaning working folder {working_subdir} …")
    try:
        shutil.rmtree(working_subdir)
        working_subdir.mkdir(parents=True, exist_ok=True)
        _emit("Working folder cleaned.")
    except Exception as e:
        _emit(f"[WARN] Failed to clean working folder: {e}")

    _emit(f"[OK] CBSL Provision Comparison post-processing complete → outputs/CBSL Provision Comparison/{dest_name}/")
    return True


# ---------------------------------------------------------------------------
# Unutilized Credit Limits — helpers
# ---------------------------------------------------------------------------

def _prepare_unutilized_credit(run_date_iso: str) -> bool:
    """
    Prepare the working directory for the Unutilized Credit Limits report:
      1. Find the latest outputs/Unutilized Credit Limits date folder.
      2. Clean working/Unutilized Credit Limits.
      3. Copy the entire outputs date folder into working/Unutilized Credit Limits/<same name>.
    Returns True on success, False on failure.
    """
    outputs_subdir = OUTPUTS_DIR / "Unutilized Credit Limits"
    working_subdir = WORKING_DIR / "Unutilized Credit Limits"

    _emit("Preparing Unutilized Credit Limits working directory…")

    src_folder = _find_latest_outputs_folder(outputs_subdir)
    if not src_folder:
        _emit("[ERROR] Cannot find latest outputs folder for Unutilized Credit Limits.")
        return False

    if working_subdir.exists():
        _emit(f"Cleaning {working_subdir} …")
        try:
            shutil.rmtree(working_subdir)
        except Exception as e:
            _emit(f"[ERROR] Failed to clean working folder: {e}")
            return False
    working_subdir.mkdir(parents=True, exist_ok=True)

    dest_folder = working_subdir / src_folder.name
    _emit(f"Copying {src_folder.name} → working/Unutilized Credit Limits/{src_folder.name} …")
    try:
        shutil.copytree(str(src_folder), str(dest_folder))
    except Exception as e:
        _emit(f"[ERROR] Failed to copy folder: {e}")
        return False

    _emit("Unutilized Credit Limits working directory prepared successfully.")
    return True


def _copy_ucl_to_outputs(ucl_file: Path, dest_outputs_name: str) -> None:
    """Replace 'Unutilized Credit Limits - *.xlsx' in the latest date folder of
    outputs/<dest_outputs_name> with ucl_file."""
    base_outputs = OUTPUTS_DIR / dest_outputs_name
    latest = _find_latest_outputs_folder(base_outputs)
    if not latest:
        _emit(f"[WARN] No latest folder found in outputs/{dest_outputs_name} — skipping copy.")
        return
    for old in latest.glob("Unutilized Credit Limits - *.xlsx"):
        try:
            old.unlink()
            _emit(f"Removed existing: {dest_outputs_name}/{latest.name}/{old.name}")
        except Exception as e:
            _emit(f"[WARN] Could not remove {old.name}: {e}")
    try:
        shutil.copy2(str(ucl_file), str(latest / ucl_file.name))
        _emit(f"[OK] Copied {ucl_file.name} → outputs/{dest_outputs_name}/{latest.name}/")
    except Exception as e:
        _emit(f"[WARN] Failed to copy to outputs/{dest_outputs_name}: {e}")


def _post_unutilized_credit(run_date_iso: str) -> bool:
    """
    After unutilized_credit_limits.py completes:
      1. Find the single date subfolder in working/Unutilized Credit Limits/.
      2. Locate 'Unutilized Credit Limits - *.xlsx' inside that subfolder.
      3. Rename subfolder to DD-MM-YYYY.
      4. Set download file + distribute to output folders from working copy.
      5. Copy renamed subfolder → outputs/Unutilized Credit Limits/<DD-MM-YYYY[(/n)]>.
      6. Append DD-MM-YYYY to completed_dates.txt.
      7. Point download file to outputs copy.
      8. Clean working/Unutilized Credit Limits/.
    """
    global _ucl_download_file

    working_subdir = WORKING_DIR / "Unutilized Credit Limits"
    outputs_subdir = OUTPUTS_DIR / "Unutilized Credit Limits"

    try:
        run_dt = datetime.strptime(run_date_iso, "%Y-%m-%d")
    except ValueError:
        _emit(f"[ERROR] Invalid run date for post-processing: {run_date_iso}")
        return False
    date_folder_name = run_dt.strftime("%d-%m-%Y")

    # Step 1: find the single date subfolder in working
    subdirs = [d for d in working_subdir.iterdir() if d.is_dir()] if working_subdir.exists() else []
    if len(subdirs) != 1:
        _emit(f"[ERROR] Expected 1 subfolder in {working_subdir}, found {len(subdirs)}.")
        return False
    work_date_folder = subdirs[0]

    # Step 2: locate the UCL xlsx inside the subfolder
    ucl_matches = [f for f in work_date_folder.glob("Unutilized Credit Limits - *.xlsx")
                   if not f.name.startswith("~$")]
    if not ucl_matches:
        _emit("[WARN] No 'Unutilized Credit Limits - *.xlsx' found in working subfolder.")
        ucl_file_working = None
    else:
        ucl_file_working = ucl_matches[0]
        _ucl_download_file = ucl_file_working
        _emit(f"Download file set: {ucl_file_working.name}")

    # Step 3: rename working subfolder to DD-MM-YYYY
    renamed_folder = working_subdir / date_folder_name
    if work_date_folder != renamed_folder:
        _emit(f"Renaming {work_date_folder.name} → {date_folder_name} …")
        try:
            work_date_folder.rename(renamed_folder)
            work_date_folder = renamed_folder
            if ucl_file_working:
                ucl_file_working = work_date_folder / ucl_file_working.name
        except Exception as e:
            _emit(f"[ERROR] Failed to rename working subfolder: {e}")
            return False

    # Step 4: distribute to output folders from working copy
    if ucl_file_working and ucl_file_working.exists():
        _ucl_download_file = ucl_file_working
        _copy_ucl_to_outputs(ucl_file_working, "NBD_MF_20_C1_C6")
        _copy_ucl_to_outputs(ucl_file_working, "NBD_AF_20_C1_C6")
    else:
        _emit("[WARN] No 'Unutilized Credit Limits - *.xlsx' found — skipping distribution.")

    # Step 5: copy working subfolder → outputs/Unutilized Credit Limits/<DD-MM-YYYY[(/n)]>
    outputs_subdir.mkdir(parents=True, exist_ok=True)
    dest_name = date_folder_name
    version = 0
    while (outputs_subdir / dest_name).exists():
        version += 1
        dest_name = f"{date_folder_name}({version})"
    dest_folder = outputs_subdir / dest_name

    _emit(f"Copying {work_date_folder.name} → outputs/Unutilized Credit Limits/{dest_name} …")
    try:
        shutil.copytree(str(work_date_folder), str(dest_folder))
    except Exception as e:
        _emit(f"[ERROR] Failed to copy to Unutilized Credit Limits outputs: {e}")
        return False

    # Step 6: append to completed_dates.txt
    dates_file = outputs_subdir / "completed_dates.txt"
    try:
        with open(dates_file, "a", encoding="utf-8") as fh:
            fh.write(date_folder_name + "\n")
        _emit(f"Appended {date_folder_name} to completed_dates.txt")
    except Exception as e:
        _emit(f"[WARN] Could not update completed_dates.txt: {e}")

    # Step 7: point download file to outputs copy
    out_matches = [f for f in dest_folder.glob("Unutilized Credit Limits - *.xlsx")
                   if not f.name.startswith("~$")]
    if out_matches:
        _ucl_download_file = out_matches[0]
        _emit(f"Download file updated to outputs copy: {_ucl_download_file.name}")

    # Step 8: clean working folder
    _emit(f"Cleaning working folder {working_subdir} …")
    try:
        shutil.rmtree(working_subdir)
        working_subdir.mkdir(parents=True, exist_ok=True)
        _emit("Working folder cleaned.")
    except Exception as e:
        _emit(f"[WARN] Failed to clean working folder: {e}")

    _emit(f"[OK] Unutilized Credit Limits post-processing complete → outputs/Unutilized Credit Limits/{dest_name}/")
    return True


# ---------------------------------------------------------------------------
# Background runner
# ---------------------------------------------------------------------------
def _run_reports_thread(reports_to_run: list[dict], date_str: str) -> None:
    """
    date_str: ISO date string YYYY-MM-DD from the date picker.
    Converted to DD/MM/YYYY before passing to each script as argv[1].
    """
    global _stop_flag, _running_proc

    # Convert YYYY-MM-DD → DD/MM/YYYY for script argv
    try:
        _dt = datetime.strptime(date_str, "%Y-%m-%d")
        script_date = _dt.strftime("%d/%m/%Y")
    except ValueError:
        script_date = date_str  # pass as-is if conversion fails

    _set_state(running=True, stage="running", message="Starting…")
    _emit("=" * 50)
    _emit(f"Supporting Reports run started  (date: {script_date})")
    _emit("=" * 50)

    overall_ok = True

    for report in reports_to_run:
        if _stop_flag:
            _emit("Stop requested — aborting remaining reports.")
            _set_checklist_status(report["id"], "stopped")
            continue

        script_path = SCRIPTS_DIR / report["script"]
        if not script_path.exists():
            _emit(f"[SKIP] Script not found: {report['script']}")
            _set_checklist_status(report["id"], "failed")
            overall_ok = False
            continue

        _set_state(current_report=report["label"], message=f"Running {report['label']}…")
        _set_checklist_status(report["id"], "running")
        _emit(f"--- Running: {report['label']} ---")

        # Report-specific preparation before running the script
        if report["id"] == "SUPPORTING_SCHEDULES":
            _emit("Preparing working directory for Supporting Schedules…")
            if not _prepare_supporting_schedules(date_str):
                _emit("[FAIL] Preparation failed — aborting Supporting Schedules.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False
                continue
        elif report["id"] == "UNUTILIZED_CREDIT":
            _emit("Preparing working directory for Unutilized Credit Limits…")
            if not _prepare_unutilized_credit(date_str):
                _emit("[FAIL] Preparation failed — aborting Unutilized Credit Limits.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False
                continue
        elif report["id"] == "CBSL_PROVISION":
            _emit("Preparing working directory for CBSL Provision Comparison…")
            if not _prepare_cbsl_provision(date_str):
                _emit("[FAIL] Preparation failed — aborting CBSL Provision Comparison.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False
                continue

        # Unutilized Credit Limits runs via direct import to override BASE_PATH.
        if report["id"] == "UNUTILIZED_CREDIT":
            wf_base = WORKING_DIR / "Unutilized Credit Limits"
            wf_subdirs = [d for d in wf_base.iterdir() if d.is_dir()] \
                         if wf_base.exists() else []
            if not wf_subdirs:
                _emit("[ERROR] No date subfolder found in working/Unutilized Credit Limits after preparation.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False
                continue

            working_folder = str(wf_subdirs[0])
            _emit(f"Running Unutilized Credit Limits with base path: {working_folder}")

            _ucl_log_file = None
            try:
                import importlib, contextlib as _ctx
                _scripts_str = str(SCRIPTS_DIR)
                if _scripts_str not in sys.path:
                    sys.path.insert(0, _scripts_str)

                import unutilized_credit_limits as _ucl_mod
                importlib.reload(_ucl_mod)
                _ucl_mod.BASE_PATH = working_folder

                _ucl_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                _ucl_log_dir = LOGS_DIR / f"run_{_ucl_ts}"
                _ucl_log_dir.mkdir(parents=True, exist_ok=True)
                _ucl_log_file = _ucl_log_dir / "Unutilized_Credit_Limits.log"

                class _UCLLogWriter:
                    def __init__(self, fh):
                        self._fh = fh
                    def write(self, s):
                        for line in s.splitlines():
                            if line.strip():
                                _emit(line)
                                self._fh.write(line + "\n")
                                self._fh.flush()
                    def flush(self): pass

                with open(_ucl_log_file, "w", encoding="utf-8") as _ucl_fh:
                    _ucl_fh.write(f"Unutilized Credit Limits — started {datetime.now().isoformat()}\n")
                    _ucl_fh.write(f"Working folder: {working_folder}\n")
                    _ucl_fh.write("-" * 60 + "\n")
                    _writer = _UCLLogWriter(_ucl_fh)
                    with _ctx.redirect_stdout(_writer), _ctx.redirect_stderr(_writer):
                        processor = _ucl_mod.ExcelProcessor(working_folder)
                        processor.run_all_steps(date_str=script_date)

                _emit(f"Log written: logs/run_{_ucl_ts}/Unutilized_Credit_Limits.log")
                ucl_returncode = 0
            except Exception as _e:
                _emit(f"[ERROR] Unutilized Credit Limits failed: {_e}")
                import traceback as _tb
                _tb_str = _tb.format_exc()
                _emit(_tb_str)
                if _ucl_log_file:
                    try:
                        with open(_ucl_log_file, "a", encoding="utf-8") as _ucl_fh:
                            _ucl_fh.write(f"\n[ERROR] {_e}\n{_tb_str}\n")
                    except Exception:
                        pass
                ucl_returncode = 1

            if _stop_flag:
                _set_checklist_status(report["id"], "stopped")
            elif ucl_returncode == 0:
                _emit(f"[OK] {report['label']} completed.")
                if _post_unutilized_credit(date_str):
                    if _ucl_download_file and _ucl_download_file.exists():
                        _set_state(ucl_download_ready=True)
                else:
                    _emit("[WARN] Post-processing failed — Unutilized Credit Limits output may not have been copied.")
                _set_checklist_status(report["id"], "done")
            else:
                _emit(f"[FAIL] {report['label']} failed.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False
            continue  # skip generic subprocess block below

        # CBSL Provision runs via direct import so we can redirect base_path
        # without modifying the script itself.
        if report["id"] == "CBSL_PROVISION":
            wf_base = WORKING_DIR / "CBSL Provision Comparison"
            wf_subdirs = [d for d in wf_base.iterdir() if d.is_dir()] \
                         if wf_base.exists() else []
            if not wf_subdirs:
                _emit("[ERROR] No date subfolder found in working/CBSL Provision Comparison after preparation.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False
                continue

            working_folder = str(wf_subdirs[0])
            _emit(f"Running CBSL Provision Comparison with base path: {working_folder}")

            _cbsl_log_file = None
            try:
                import importlib, io as _io, contextlib as _ctx
                _scripts_str = str(SCRIPTS_DIR)
                if _scripts_str not in sys.path:
                    sys.path.insert(0, _scripts_str)

                # Force fresh import (may have been imported before with old FILE_PATH)
                import cbsl_provision_comparison as _cbsl_mod
                importlib.reload(_cbsl_mod)

                # Patch FILE_PATH so ExcelAutomation uses the working folder
                _cbsl_mod.FILE_PATH = working_folder

                reporting_date = datetime.strptime(script_date, "%d/%m/%Y")

                # Set up log file in logs/run_YYYYMMDD_HHMMSS/
                _cbsl_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                _cbsl_log_dir = LOGS_DIR / f"run_{_cbsl_ts}"
                _cbsl_log_dir.mkdir(parents=True, exist_ok=True)
                _cbsl_log_file = _cbsl_log_dir / "CBSL_Provision_Comparison.log"

                # Capture stdout/stderr — emit to UI buffer and write to log file
                class _LogWriter:
                    def __init__(self, fh):
                        self._fh = fh
                    def write(self, s):
                        for line in s.splitlines():
                            if line.strip():
                                _emit(line)
                                self._fh.write(line + "\n")
                                self._fh.flush()
                    def flush(self): pass

                with open(_cbsl_log_file, "w", encoding="utf-8") as _cbsl_fh:
                    _cbsl_fh.write(f"CBSL Provision Comparison — started {datetime.now().isoformat()}\n")
                    _cbsl_fh.write(f"Working folder: {working_folder}\n")
                    _cbsl_fh.write("-" * 60 + "\n")
                    _writer = _LogWriter(_cbsl_fh)
                    with _ctx.redirect_stdout(_writer), _ctx.redirect_stderr(_writer):
                        automation = _cbsl_mod.ExcelAutomation(working_folder)
                        automation.run_all_steps(reporting_date=reporting_date)

                _emit(f"Log written: logs/run_{_cbsl_ts}/CBSL_Provision_Comparison.log")
                cbsl_returncode = 0
            except Exception as _e:
                _emit(f"[ERROR] CBSL Provision Comparison failed: {_e}")
                import traceback as _tb
                _tb_str = _tb.format_exc()
                _emit(_tb_str)
                # Append error to log file if it was created
                if _cbsl_log_file:
                    try:
                        with open(_cbsl_log_file, "a", encoding="utf-8") as _cbsl_fh:
                            _cbsl_fh.write(f"\n[ERROR] {_e}\n{_tb_str}\n")
                    except Exception:
                        pass
                cbsl_returncode = 1

            if _stop_flag:
                _set_checklist_status(report["id"], "stopped")
            elif cbsl_returncode == 0:
                _emit(f"[OK] {report['label']} completed.")
                if _post_cbsl_provision(date_str):
                    if _cbsl_download_file and _cbsl_download_file.exists():
                        _set_state(cbsl_download_ready=True)
                else:
                    _emit("[WARN] Post-processing failed — CBSL Provision output may not have been copied.")
                _set_checklist_status(report["id"], "done")
            else:
                _emit(f"[FAIL] {report['label']} failed.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False
            continue  # skip generic subprocess block below

        _gen_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        _gen_log_dir = LOGS_DIR / f"run_{_gen_ts}"
        _gen_log_dir.mkdir(parents=True, exist_ok=True)
        _gen_log_file = _gen_log_dir / f"{re.sub(r'[^A-Za-z0-9_-]', '_', report['id'])}.log"
        cmd = [sys.executable, str(script_path), script_date, "--log-dir", str(_gen_log_dir)]
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                cwd=str(SCRIPTS_DIR),
            )
            with _state_lock:
                pass  # just a fence
            _running_proc = proc

            for line in proc.stdout:
                _emit(line.rstrip())
                if _stop_flag:
                    proc.terminate()
                    break

            proc.wait()
            _running_proc = None

            if _stop_flag:
                _set_checklist_status(report["id"], "stopped")
            elif proc.returncode == 0:
                _emit(f"[OK] {report['label']} completed.")
                if report["id"] == "SUPPORTING_SCHEDULES":
                    if _post_supporting_schedules(date_str):
                        if _download_file and _download_file.exists():
                            _set_state(ss_download_ready=True)
                    else:
                        _emit("[WARN] Post-processing step failed — report output may not have been copied.")
                _set_checklist_status(report["id"], "done")
            else:
                _emit(f"[FAIL] {report['label']} exited with code {proc.returncode}.")
                _set_checklist_status(report["id"], "failed")
                overall_ok = False

        except Exception as e:
            _running_proc = None
            import traceback as _tb
            _emit(f"[ERROR] {report['label']}: {e}")
            _emit(_tb.format_exc())
            try:
                with open(_gen_log_file, "a", encoding="utf-8") as _gen_fh:
                    _gen_fh.write(f"\n[ERROR] {e}\n{_tb.format_exc()}\n")
            except Exception:
                pass
            _set_checklist_status(report["id"], "failed")
            overall_ok = False

    stage = "done" if overall_ok else "error"
    msg = "All selected reports completed." if overall_ok else "One or more reports failed — check logs."
    _emit("=" * 50)
    _emit(msg)
    _set_state(running=False, stage=stage, message=msg, current_report="")


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@sup_bp.route("/", methods=["GET"])
def index():
    """Supporting Reports landing page."""
    reports = sorted(SUPPORTING_REPORTS, key=lambda r: r["priority"])
    last_runs = {r["id"]: _last_run_info(r["id"]) for r in reports}
    today_iso = datetime.now().strftime("%Y-%m-%d")
    return render_template(
        "supporting_reports.html",
        reports=reports,
        last_runs=last_runs,
        today_iso=today_iso,
    )


@sup_bp.route("/run", methods=["POST"])
def run():
    """Start selected supporting reports."""
    global _stop_flag, _pending_cbsl_run

    with _state_lock:
        if _state["running"]:
            flash("A run is already in progress.", "warning")
            return redirect(url_for("sup.index"))

    selected_ids = request.form.getlist("report_ids")
    date_str     = request.form.get("date", "").strip()   # YYYY-MM-DD from date picker

    if not selected_ids:
        flash("Please select at least one report.", "warning")
        return redirect(url_for("sup.index"))

    if not date_str:
        flash("Please select a date.", "warning")
        return redirect(url_for("sup.index"))

    # Basic ISO date validation
    try:
        datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        flash("Invalid date format. Please use the date picker.", "warning")
        return redirect(url_for("sup.index"))

    reports_to_run = [r for r in SUPPORTING_REPORTS if r["id"] in selected_ids]
    if not reports_to_run:
        flash("No valid reports selected.", "warning")
        return redirect(url_for("sup.index"))

    # Sort by priority
    reports_to_run.sort(key=lambda r: r["priority"])

    # If CBSL_PROVISION is in the list, redirect to master-data-form first
    if any(r["id"] == "CBSL_PROVISION" for r in reports_to_run):
        _pending_cbsl_run = {
            "reports": [r["id"] for r in reports_to_run],
            "date_str": date_str,
        }
        return redirect(url_for("sup.master_data_form_view", date=date_str))

    return _start_run(reports_to_run, date_str)


def _start_run(reports_to_run: list[dict], date_str: str):
    """Initialise checklist and start background thread."""
    _reset()
    with _checklist_lock:
        for r in reports_to_run:
            _checklist.append({
                "id": r["id"],
                "label": r["label"],
                "status": "pending",
                "elapsed": None,
                "start_time": None,
            })

    t = threading.Thread(
        target=_run_reports_thread,
        args=(reports_to_run, date_str),
        daemon=True,
    )
    t.start()
    return redirect(url_for("sup.index"))


@sup_bp.route("/stop", methods=["POST"])
def stop():
    """Request stop of the current run."""
    global _stop_flag, _running_proc
    _stop_flag = True
    if _running_proc:
        try:
            _running_proc.terminate()
        except Exception:
            pass
    _set_state(stage="stopping", message="Stop requested…")
    return jsonify({"ok": True})


@sup_bp.route("/status")
def status():
    """Poll endpoint: returns current run state."""
    with _state_lock:
        data = dict(_state)
    return jsonify(data)


@sup_bp.route("/logs")
def logs():
    """Poll endpoint: returns log buffer as plain text."""
    with _log_lock:
        content = "\n".join(_log_buffer)
    return current_app.response_class(content, mimetype="text/plain")


@sup_bp.route("/checklist")
def checklist():
    """Poll endpoint: returns report checklist JSON."""
    with _checklist_lock:
        data = list(_checklist)
    return jsonify(data)


@sup_bp.route("/download")
def download():
    """Serve the Supporting Schedules output file for download."""
    if not _download_file or not _download_file.exists():
        return current_app.response_class("No file available for download.", status=404, mimetype="text/plain")
    return send_file(str(_download_file), as_attachment=True, download_name=_download_file.name)


@sup_bp.route("/download-ack", methods=["POST"])
def download_ack():
    """Clear ss_download_ready so page refreshes don't re-trigger the download."""
    _set_state(ss_download_ready=False)
    return jsonify({"ok": True})


@sup_bp.route("/download-cbsl")
def download_cbsl():
    """Serve the CBSL Provision Comparison output file for download."""
    if not _cbsl_download_file or not _cbsl_download_file.exists():
        return current_app.response_class("No file available for download.", status=404, mimetype="text/plain")
    return send_file(str(_cbsl_download_file), as_attachment=True, download_name=_cbsl_download_file.name)


@sup_bp.route("/download-cbsl-ack", methods=["POST"])
def download_cbsl_ack():
    """Clear cbsl_download_ready so page refreshes don't re-trigger the download."""
    _set_state(cbsl_download_ready=False)
    return jsonify({"ok": True})


@sup_bp.route("/download-ucl")
def download_ucl():
    """Serve the Unutilized Credit Limits output file for download."""
    if not _ucl_download_file or not _ucl_download_file.exists():
        return current_app.response_class("No file available for download.", status=404, mimetype="text/plain")
    return send_file(str(_ucl_download_file), as_attachment=True, download_name=_ucl_download_file.name)


@sup_bp.route("/download-ucl-ack", methods=["POST"])
def download_ucl_ack():
    """Clear ucl_download_ready so page refreshes don't re-trigger the download."""
    _set_state(ucl_download_ready=False)
    return jsonify({"ok": True})


@sup_bp.route("/master-data-form", methods=["GET"])
def master_data_form_view():
    """Show the CBSL Provision master-data form (Property Mortgages upload)."""
    date_str = request.args.get("date", "")
    latest_pm = _get_latest_property_mortgages_name()
    return render_template(
        "master_data_form.html",
        # Section visibility flags
        show_sofp=False, show_ia=False, show_c1c2=False, show_c1c6=False,
        show_ga_is=False, show_c4=False, show_la=False, show_ecs=False,
        show_qf=False, show_cbsl_provision=True,
        run_date=date_str,
        # Data placeholders expected by the template
        sofp_data={}, form_data={},
        uploaded_files={"Property Mortgages": latest_pm or "No file uploaded"},
        # Back button should go to supporting reports
        back_url=url_for("sup.index"),
    )


@sup_bp.route("/upload-property-mortgages", methods=["POST"])
def upload_property_mortgages():
    """Handle Property Mortgages file upload and start CBSL Provision run."""
    global _pending_cbsl_run

    run_date = request.form.get("_run_date", "").strip()
    month    = request.form.get("property_morgage_month", "").strip()
    year     = request.form.get("property_mrgage_year", "").strip()
    file_obj = request.files.get("property_morgage_file")

    if not run_date:
        flash("Run date missing — please go back and try again.", "danger")
        return redirect(url_for("sup.master_data_form_view"))

    # Handle optional file upload
    if file_obj and file_obj.filename:
        if not month or not year:
            flash("Month and year are required when uploading a Property Mortgages file.", "warning")
            return redirect(url_for("sup.master_data_form_view", date=run_date))

        if not year.isdigit() or len(year) != 4:
            flash("Year must be a 4-digit number.", "warning")
            return redirect(url_for("sup.master_data_form_view", date=run_date))

        # Save to Master Inputs/Property mortgages/
        mi_dir = MASTER_INPUTS_DIR / "Property mortgages"
        mi_dir.mkdir(parents=True, exist_ok=True)

        new_filename = f"Property mortgages - {month} {year}.xlsx"
        dest_path = mi_dir / new_filename

        # Remove existing Property mortgages - *.xlsx files
        for old in mi_dir.glob("Property mortgages - *.xlsx"):
            try:
                old.unlink()
            except Exception:
                pass
        for old in mi_dir.glob("Property mortgages - *.xls"):
            try:
                old.unlink()
            except Exception:
                pass

        file_obj.save(str(dest_path))
        _emit(f"Property Mortgages file saved: {new_filename}")

        # Also copy to latest CBSL Provision Comparison outputs folder
        latest_cbsl_folder = _find_cbsl_latest_outputs_folder()
        if latest_cbsl_folder:
            for old in latest_cbsl_folder.glob("Property mortgages *.xlsx"):
                try:
                    old.unlink()
                except Exception:
                    pass
            for old in latest_cbsl_folder.glob("Property mortgages - *.xlsx"):
                try:
                    old.unlink()
                except Exception:
                    pass
            try:
                shutil.copy2(str(dest_path), str(latest_cbsl_folder / new_filename))
                _emit(f"Copied {new_filename} → outputs/CBSL Provision Comparison/{latest_cbsl_folder.name}/")
            except Exception as e:
                _emit(f"[WARN] Could not copy to CBSL outputs: {e}")

    # Determine which reports to run from pending state or form
    if _pending_cbsl_run:
        selected_ids = _pending_cbsl_run["reports"]
        date_str     = _pending_cbsl_run["date_str"] or run_date
        _pending_cbsl_run = None
    else:
        selected_ids = ["CBSL_PROVISION"]
        date_str = run_date

    reports_to_run = [r for r in SUPPORTING_REPORTS if r["id"] in selected_ids]
    reports_to_run.sort(key=lambda r: r["priority"])

    if not reports_to_run:
        flash("No valid reports to run.", "warning")
        return redirect(url_for("sup.index"))

    return _start_run(reports_to_run, date_str)


@sup_bp.route("/reset", methods=["POST"])
def reset():
    """Reset state so a new run can be started."""
    with _state_lock:
        if _state["running"]:
            return jsonify({"ok": False, "error": "Still running"})
    _reset()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# Registration helper (called from app.py)
# ---------------------------------------------------------------------------
def register_supporting_routes(app) -> None:
    """Import this function in app.py and call it after creating the Flask app."""
    from flask import current_app  # noqa: F401 — needed inside blueprint functions
    app.register_blueprint(sup_bp)


# Fix: current_app reference inside the logs() view needs the import at module level
from flask import current_app  # noqa: E402
