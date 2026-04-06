import openpyxl
from pathlib import Path
import os
import shutil
import subprocess
import sys
import threading
from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Tuple, List
import time
import re
import glob
import zipfile
import json

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from threading import Lock
from collections import deque
import openpyxl

app = Flask(__name__)

# Register supporting reports blueprint
from supporting_reports_main import register_supporting_routes
register_supporting_routes(app)

# Suppress Waitress task queue warnings
import logging
logging.getLogger('waitress.queue').setLevel(logging.ERROR)

def kill_excel_instances():
    """Kill any running Excel instances to prevent COM conflicts."""
    try:
        print("Checking for running Excel instances...")
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE'],
            capture_output=True,
            text=True,
            timeout=5
        )

        if 'EXCEL.EXE' in result.stdout:
            print("Found running Excel instances. Killing them...")
            subprocess.run(
                ['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5
            )
            import time
            time.sleep(2)
            print("Excel instances killed successfully")
        else:
            print("No Excel instances found")
    except Exception as e:
        print(f"Error killing Excel instances: {e}")

def _append_run_history(
    frequency: str,
    picked_date: str,
    scripts: list[str],
    checklist: list[dict],
    output_dir: Path | None
) -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    
    if HISTORY_XLSX.exists():
        wb = openpyxl.load_workbook(HISTORY_XLSX)
    else:
        wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "History"

    # Add header if file is new
    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        ws.append(["timestamp", "frequency", "picked_date", "reports", "statuses", "output_folder"])

    from datetime import datetime as _dt
    ts = _dt.now().strftime("%Y-%m-%d %H:%M:%S")

    reports = ", ".join(Path(s).stem for s in scripts)
    statuses = ", ".join(f"{it['name']}={it['status']}" for it in checklist)
    out_name = output_dir.name if output_dir else ""

    ws.append([ts, frequency, picked_date, reports, statuses, out_name])

    # Keep last 8 records + header
    while ws.max_row > 9:
        ws.delete_rows(2)

    wb.save(HISTORY_XLSX)
    wb.close()

# ==================== GLOBAL CONFIGURATION ====================
BASE_DIR = Path(__file__).resolve().parent
OUTPUTS_DIR = BASE_DIR / "outputs"
WORKING_DIR = BASE_DIR / "working"
REPORT_AUTOMATIONS_DIR = BASE_DIR / "report_automations"
LOGS_DIR = BASE_DIR / "logs"
MASTER_INPUTS_DIR = BASE_DIR / "Master Inputs"
OLD_OUTPUTS_DIR = BASE_DIR / "old_outputs"
HISTORY_XLSX = LOGS_DIR / "run_history.xlsx"
MASTER_DATA_XLSX = BASE_DIR / "Master_Data.xlsx"

# Status tracking
_status_lock = Lock()
_status: dict[str, str | bool] = {
    "running": False,
    "stage": "idle",
    "message": "",
}

_log_lock = Lock()
_log_buffer: deque[str] = deque(maxlen=2000)
_proc_lock = Lock()
_stop_requested: bool = False
_running_processes: dict[str, subprocess.Popen] = {}

# Create timestamped log file for each run
_timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
_run_log_dir = LOGS_DIR / f"run_{_timestamp_str}"
_run_log_dir.mkdir(parents=True, exist_ok=True)
RUN_LOG_FILE = _run_log_dir / "web_run.log"
_status_feed: deque[str] = deque(maxlen=200)
_report_checklist: list[dict] = []
_report_messages: dict[str, deque[str]] = {}
_last_output_dirs: dict[str, Path] = {}  # Track output dirs per report
_download_files: list[Path] = []  # Track files to download
_sofp_download_file: Path | None = None  # Track SOFP file for download after save
_ia_rate_updates: list[dict] = []  # Track rate updates for IA report (passed from form)
_ia_ltv_updates: list[dict] = []  # Track LTV updates for IA report (passed from form)
_la_month_end: bool = False  # Track if LA report is month-end version
_la_time_deposits: str = ""  # Time deposits value for month-end LA report

# ==================== REPORT DEPENDENCY STRUCTURE ====================
# Mapping of report sets to their output folder paths for completion date tracking
SET_COMPLETION_MAPPING = {
    1: "NBD_MF_01_SOFP_SOCI",
    2: {
        1: "NBD-WF-18-DM",
        2: "NBD_WF_15_LA",
        3: "NBD_WF_18_ECS",
        4: "NBD_MF_04_LA"
    },
    3: "NBD-MF-10-GA & NBD-MF-11-IS",
    4: "NBD_MF_23_IA",
    5: "NBD_MF_20_C1_C6",
    6: "NBD_MF_23_C4",
    7: "NBD-QF-23-C3-C10-SF-FI",
    8: "NBD_QF_23_C7",
    9: "NBD_QF_23_C8",
    10: "NBD_AF_01_SOFP_SOCI",
    11: "NBD_AF_20_C1_C6"
}

REPORT_STRUCTURE = {
    # LEVEL 1 Reports
    "NBD-MF-01-SOFP-SOCI": {
        "level": 1,
        "set": 1,
        "step": 1,
        "name": "Statement of Financial position",
        "script": "NBD_MF_01_SOFP_SOCI.py",
        "folder": "NBD_MF_01_SOFP_SOCI",
        "dependencies": [],
        "estimated_minutes": 5
    },
    "NBD-WF-18-DM": {
        "level": 1,
        "set": 2,
        "step": 1,
        "name": "Deposits Movement in Previous Week",
        "script": "NBD_WF_18_DM.py",
        "folder": "NBD-WF-18-DM",
        "dependencies": [],
        "estimated_minutes": 30
    },
    "NBD-WF-15-LA": {
        "level": 1,
        "set": 2,
        "step": 2,
        "name": "Liquid Assets",
        "script": "NBD_WF_15_LA.py",
        "folder": "NBD_WF_15_LA",
        "dependencies": [],
        "estimated_minutes": 3
    },
    "NBD-WF-18-ECS": {
        "level": 1,
        "set": 2,
        "step": 3,
        "name": "Expected Cashflow Summary",
        "script": "NBD_WF_18_ECS.py",
        "folder": "NBD_WF_18_ECS",
        "dependencies": [],
        "estimated_minutes": 4
    },
    "NBD-MF-04-LA": {
        "level": 1,
        "set": 2,
        "step": 4,
        "name": "Monthly Liquidity Report",
        "script": "NBD_MF_04_LA.py",
        "folder": "NBD_MF_04_LA",
        "dependencies": [],
        "estimated_minutes": 5
    },

    # LEVEL 2 Reports - Set 3 (GA & IS combined pipeline)
    "NBD-MF-10-GA-11-IS": {
        "level": 2,
        "set": 3,
        "step": 1,
        "name": "GA & IS Pipeline",
        "script": "NBD_MF_10_GA_NBD_MF_11_IS.py",
        "folder": "NBD-MF-10-GA & NBD-MF-11-IS",
        "dependencies": [],
        "estimated_minutes": 25
    },

    # LEVEL 2 Reports - Set 4 (IA and C1-C2)
    "NBD-MF-23-IA": {
        "level": 2,
        "set": 4,
        "step": 1,
        "name": "Product wise Classification of Loans",
        "script": "NBD_MF_23_IA.py",
        "folder": "NBD_MF_23_IA",
        "dependencies": [],
        "estimated_minutes": 15
    },
    "NBD-MF-23-C1-C2": {
        "level": 2,
        "set": 4,
        "step": 2,
        "name": "Product wise Classification of Loans (Amounts)",
        "script": "NBD_MF_23_C1C2.py",
        "folder": "NBD_MF_23_C1C2",
        "dependencies": [],
        "estimated_minutes": 10
    },

    # LEVEL 2 Reports - Set 5 (C1-C6 Forms)
    "NBD-MF-20-C1": {
        "level": 2,
        "set": 5,
        "step": 1,
        "name": "Form 1 : Computation of Capital Ratios",
        "script": "NBD_MF_20_C1.py",
        "folder": "NBD_MF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 10
    },
    "NBD-MF-20-C2": {
        "level": 2,
        "set": 5,
        "step": 2,
        "name": "Form 2 : Computation of Total Capital",
        "script": "NBD_MF_20_C2.py",
        "folder": "NBD_MF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 12
    },
    "NBD-MF-20-C3": {
        "level": 2,
        "set": 5,
        "step": 3,
        "name": "Form 3: Computation of Risk Weighted Amount for Credit Risk",
        "script": "NBD_MF_20_C3.py",
        "folder": "NBD_MF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 15
    },
    "NBD-MF-20-C4": {
        "level": 2,
        "set": 5,
        "step": 4,
        "name": "Form 4: Credit equivalent of Off-Balance Sheet Item",
        "script": "NBD_MF_20_C4.py",
        "folder": "NBD_MF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 10
    },
    "NBD-MF-20-C5": {
        "level": 2,
        "set": 5,
        "step": 5,
        "name": "Form 5: Exposures recognized under Credit Risk Mitigation CRM",
        "script": "NBD_MF_20_C5.py",
        "folder": "NBD_MF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 12
    },
    "NBD-MF-20-C6": {
        "level": 2,
        "set": 5,
        "step": 6,
        "name": "Form 6: Computation of Risk Weighted Amount for Operational Risk",
        "script": "NBD_MF_20_C6.py",
        "folder": "NBD_MF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 15
    },

    # LEVEL 2 Reports - Set 6 (C4 standalone)
    "NBD-MF-23-C4": {
        "level": 2,
        "set": 6,
        "step": 1,
        "name": "Loan CF Analysis",
        "script": "NBD-MF-23-C4.py",
        "dependencies": [],
        "folder": "NBD_MF_23_C4",
        "estimated_minutes": 30
    },

    # LEVEL 3 Reports - Set 7 (All QF Reports - Select one = Select all)
    "NBD-QF-23-C9": {
        "level": 3,
        "set": 7,
        "step": 1,
        "name": "Sector Wise Credit Exposures",
        "script": "NBD_QF_23_C9.py",
        "folder": "NBD_QF_23_C9",
        "dependencies": [],
        "estimated_minutes": 3
    },
    "NBD-QF-23-C10": {
        "level": 3,
        "set": 7,
        "step": 2,
        "name": "District Wise Credit Exposures",
        "script": "NBD_QF_23_C10.py",
        "folder": "NBD_QF_23_C10",
        "dependencies": [],
        "estimated_minutes": 3
    },
    "NBD-QF-23-C6": {
        "level": 3,
        "set": 7,
        "step": 3,
        "name": "Large Exposures (Top 50 Borrowers)",
        "script": "NBD_QF_23_C6.py",
        "folder": "NBD_QF_23_C6",
        "dependencies": [],
        "estimated_minutes": 3
    },
    "NBD-QF-23-C4": {
        "level": 3,
        "set": 7,
        "step": 4,
        "name": "Product wise Probability of Default and Loss Given Default of Loans",
        "script": "NBD_QF_23_C4.py",
        "folder": "NBD_MF_23_C4",
        "dependencies": [],
        "estimated_minutes": 3
    },
    "NBD-QF-23-C3": {
        "level": 3,
        "set": 7,
        "step": 5,
        "name": "Measurement of Credit Risk in Loans Based on Regulatory Classification",
        "script": "NBD_QF_23_C3.py",
        "folder": "NBD_QF_23_C3",
        "dependencies": [],
        "estimated_minutes": 3
    },
    "NBD-QF-16-SF": {
        "level": 3,
        "set": 7,
        "step": 6,
        "name": "Sustainable Finance Activities (SFAs)",
        "script": "NBD_QF_16_SF.py",
        "folder": "NBD_QF_16_SF",
        "dependencies": [],
        "estimated_minutes": 3
    },
    "NBD-QF-16-FI": {
        "level": 3,
        "set": 7,
        "step": 7,
        "name": "National Financial Inclusion Council Survey",
        "script": "NBD_QF_16_FI.py",
        "folder": "NBD_QF_16_FI",
        "dependencies": [],
        "estimated_minutes": 3
    },

    # LEVEL 3 Reports - Other Sets
    "NBD-QF-23-C7": {
        "level": 3,
        "set": 8,
        "step": 1,
        "name": "Detailed breakdown of the stage-wise movement of loans",
        "script": "NBD_QF_23_C7.py",
        "folder": "NBD_QF_23_C7",
        "dependencies": [],
        "estimated_minutes": 10
    },
    "NBD-QF-23-C8": {
        "level": 3,
        "set": 9,
        "step": 1,
        "name": "Classification of Loans based on Internal Credit Risk Rating",
        "script": "NBD_QF_23_C8.py",
        "folder": "NBD_QF_23_C8",
        "dependencies": [],
        "estimated_minutes": 12
    },

    # LEVEL 4 Reports (Annual Reports)
    "NBD-AF-01-SOFP-SOCI": {
        "level": 4,
        "set": 10,
        "step": 1,
        "name": "Statement of Financial position",
        "script": "NBD_AF_01_SOFP_SOCI.py",
        "folder": "NBD_AF_01_SOFP_SOCI",
        "dependencies": [],
        "estimated_minutes": 5
    },
    "NBD-AF-20-C2": {
        "level": 4,
        "set": 11,
        "step": 1,
        "name": "Form 2 : Computation of Total Capital",
        "script": "NBD_AF_20_C2.py",
        "folder": "NBD_AF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 12
    },
    "NBD-AF-20-C3": {
        "level": 4,
        "set": 11,
        "step": 2,
        "name": "Form 3: Computation of Risk Weighted Amount for Credit Risk",
        "script": "NBD_AF_20_C3.py",
        "folder": "NBD_AF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 15
    },
    "NBD-AF-20-C4": {
        "level": 4,
        "set": 11,
        "step": 3,
        "name": "Form 4: Credit equivalent of Off-Balance Sheet Item",
        "script": "NBD_AF_20_C4.py",
        "folder": "NBD_AF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 10
    },
    "NBD-AF-20-C5": {
        "level": 4,
        "set": 11,
        "step": 4,
        "name": "Form 5: Exposures recognized under Credit Risk Mitigation CRM",
        "script": "NBD_AF_20_C5.py",
        "folder": "NBD_AF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 12
    },
    "NBD-AF-20-C6": {
        "level": 4,
        "set": 11,
        "step": 5,
        "name": "Form 6: Computation of Risk Weighted Amount for Operational Risk",
        "script": "NBD_AF_20_C6.py",
        "folder": "NBD_AF_20_C1_C6",
        "dependencies": [],
        "estimated_minutes": 15
    },
}

# ==================== HELPER FUNCTIONS ====================

def _emit_log(message: str, *, important: bool = False) -> None:
    """Emit log message."""
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
    with _log_lock:
        _log_buffer.append(line)
    try:
        with RUN_LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def _set_status(**kwargs) -> None:
    with _status_lock:
        _status.update(kwargs)


def _reset_status() -> None:
    _set_status(running=False, stage="idle", message="")


def _initialize_new_run() -> None:
    """Initialize a new run with fresh timestamp and log directory."""
    global _timestamp_str, _run_log_dir, RUN_LOG_FILE
    _timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    _run_log_dir = LOGS_DIR / f"run_{_timestamp_str}"
    _run_log_dir.mkdir(parents=True, exist_ok=True)
    RUN_LOG_FILE = _run_log_dir / "web_run.log"
    _emit_log(f"Starting new run: {_timestamp_str}", important=True)


def _calculate_estimated_time_from_logs(report_id: str) -> int:
    """Calculate estimated execution time based on historical log files.

    Args:
        report_id: The report ID (e.g., 'NBD-MF-23-IA')

    Returns:
        Estimated time in minutes (rounded up), or default from REPORT_STRUCTURE if no logs found
    """
    import re
    from datetime import datetime as dt

    try:
        # Log file name is the report_id with .log extension (e.g., NBD-MF-23-IA.log)
        log_file_name = f"{report_id}.log"

        # Find recent run folders (look at last 10 runs)
        run_folders = []
        if LOGS_DIR.exists():
            all_folders = [f for f in LOGS_DIR.iterdir() if f.is_dir() and f.name.startswith("run_")]
            # Sort by folder name (timestamp) descending
            all_folders.sort(reverse=True)
            run_folders = all_folders[:10]  # Last 10 runs

        execution_times = []

        for run_folder in run_folders:
            log_file = run_folder / log_file_name
            if not log_file.exists():
                continue

            try:
                with open(log_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()

                # Find first timestamp
                start_time = None
                for line in lines[:50]:  # Check first 50 lines
                    match = re.match(r'^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', line)
                    if match:
                        start_time = dt.strptime(match.group(1), '%Y-%m-%d %H:%M:%S')
                        break

                # Find last timestamp (completion)
                end_time = None
                for line in reversed(lines[-100:]):  # Check last 100 lines
                    match = re.match(r'^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', line)
                    if match:
                        end_time = dt.strptime(match.group(1), '%Y-%m-%d %H:%M:%S')
                        break

                # Calculate duration
                if start_time and end_time:
                    duration_seconds = (end_time - start_time).total_seconds()
                    if duration_seconds > 0:  # Valid duration
                        execution_times.append(duration_seconds)

            except Exception:
                continue  # Skip this log file if error

        # Calculate average from recent executions
        if execution_times:
            avg_seconds = sum(execution_times) / len(execution_times)
            # Round up to nearest minute, minimum 1 minute
            est_minutes = max(1, int((avg_seconds + 59) / 60))
            return est_minutes
        else:
            # No historical data, use default
            return REPORT_STRUCTURE.get(report_id, {}).get("estimated_minutes", 10)

    except Exception:
        # Fallback to hardcoded value on any error
        return REPORT_STRUCTURE.get(report_id, {}).get("estimated_minutes", 10)


def _init_report_checklist(report_ids: list[str]) -> None:
    """Initialize report checklist for the UI."""
    global _report_checklist, _report_messages
    _report_checklist = []
    _report_messages = {}
    for report_id in report_ids:
        if report_id in REPORT_STRUCTURE:
            name = REPORT_STRUCTURE[report_id]["name"]
            # Calculate estimated time from historical logs
            est_min = _calculate_estimated_time_from_logs(report_id)
            default_min = REPORT_STRUCTURE[report_id].get("estimated_minutes", 10)

            # Log the estimation for debugging
            if est_min != default_min:
                _emit_log(f"[{report_id}] Estimated time from logs: {est_min} min (default was: {default_min} min)", important=False)
            else:
                _emit_log(f"[{report_id}] Using default estimated time: {est_min} min (no historical data found)", important=False)

            _report_checklist.append({
                "name": report_id,
                "display_name": name,
                "status": "pending",
                "estimated_minutes": est_min,
                "start_time": None,
                "elapsed_seconds": 0
            })
            _report_messages[report_id] = deque(maxlen=200)


def _set_report_status(report_id: str, status: str) -> None:
    """Set status for a report by ID."""
    for item in _report_checklist:
        if item["name"] == report_id:
            item["status"] = status
            if status == "running":
                item["start_time"] = time.time()
            elif status in ["completed", "failed", "stopped"]:
                if item["start_time"]:
                    item["elapsed_seconds"] = int(time.time() - item["start_time"])
            break


def _update_report_elapsed_time(report_id: str) -> None:
    """Update elapsed time for running report."""
    for item in _report_checklist:
        if item["name"] == report_id and item["status"] == "running" and item["start_time"]:
            item["elapsed_seconds"] = int(time.time() - item["start_time"])
            break


def parse_date_folder_name(folder_name: str) -> datetime:
    """Parse date from folder name format: 30/09/2025 or 31/10/2025(1)"""
    base_name = re.sub(r'\(\d+\)$', '', folder_name)
    try:
        return datetime.strptime(base_name, "%d/%m/%Y")
    except ValueError:
        try:
            return datetime.strptime(base_name, "%d-%m-%Y")
        except ValueError:
            try:
                return datetime.strptime(base_name, "%m/%d/%Y")
            except ValueError:
                raise ValueError(f"Cannot parse date from folder name: {folder_name}")


def find_latest_date_folder(report_folder: Path) -> Path | None:
    """Find the latest versioned date folder in a report folder."""
    if not report_folder.exists():
        return None
    
    candidates: list[tuple[datetime, int, Path]] = []
    
    for child in report_folder.iterdir():
        if not child.is_dir():
            continue
        
        folder_name = child.name
        try:
            date_obj = parse_date_folder_name(folder_name)
            version_match = re.search(r'\((\d+)\)$', folder_name)
            version = int(version_match.group(1)) if version_match else 0
            candidates.append((date_obj, version, child))
        except ValueError:
            continue
    
    if not candidates:
        return None
    
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2]


def find_latest_completed_date(source_folder: Path) -> tuple[datetime | None, Path | None]:
    """
    Find the latest completed date from completed_dates.txt in source folder.
    Returns tuple of (datetime, date_folder_path)
    """
    completed_dates_file = source_folder / "completed_dates.txt"
    
    if not completed_dates_file.exists():
        _emit_log(f"No completed_dates.txt found in {source_folder}", important=True)
        return None, None
    
    try:
        with open(completed_dates_file, 'r') as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]
        
        if not lines:
            _emit_log(f"completed_dates.txt is empty in {source_folder}", important=True)
            return None, None
        
        # Get the last date from file
        last_date_str = lines[-1]
        _emit_log(f"Found last completed date: {last_date_str}", important=True)
        
        # Parse the date (supports DD-MM-YYYY and DD/MM/YYYY)
        try:
            last_date = datetime.strptime(last_date_str, "%d-%m-%Y")
        except ValueError:
            last_date = datetime.strptime(last_date_str, "%d/%m/%Y")
        
        # Find the folder with this date (including variations)
        date_folder = find_date_folder_with_variations(source_folder, last_date_str)
        
        if date_folder:
            _emit_log(f"Found date folder: {date_folder}", important=True)
            return last_date, date_folder
        else:
            _emit_log(f"Date folder not found for: {last_date_str}", important=True)
            return last_date, None
            
    except Exception as e:
        _emit_log(f"Error reading completed_dates.txt: {e}", important=True)
        return None, None


def find_date_folder_with_variations(parent_folder: Path, date_str: str) -> Path | None:
    """
    Find the latest variation of a date folder (e.g., 30/09/2025, 30/09/2025(1), etc.)
    """
    if not parent_folder.exists():
        return None
    
    candidates: list[tuple[int, Path]] = []
    
    date_variants = {date_str}
    # Accept both DD/MM/YYYY and DD-MM-YYYY
    try:
        _dt = datetime.strptime(date_str, "%d/%m/%Y")
        date_variants.add(_dt.strftime("%d-%m-%Y"))
    except Exception:
        try:
            _dt = datetime.strptime(date_str, "%d-%m-%Y")
            date_variants.add(_dt.strftime("%d/%m/%Y"))
        except Exception:
            pass
    
    for child in parent_folder.iterdir():
        if not child.is_dir():
            continue
        
        folder_name = child.name
        
        # Check if folder name starts with the date
        if any(folder_name.startswith(v) for v in date_variants):
            # Extract version number
            version_match = re.search(r'\((\d+)\)$', folder_name)
            version = int(version_match.group(1)) if version_match else 0
            candidates.append((version, child))
    
    if not candidates:
        return None
    
    # Sort by version descending and return the latest
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def format_date_folder_name(date_str: str) -> str:
    """Convert date from HTML format (YYYY-MM-DD) to folder format (DD-MM-YYYY)."""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return dt.strftime("%d-%m-%Y")


def copy_ia_to_additional_outputs(ia_file_path: Path, gui_date_str: str) -> None:
    """
    Copy the IA 'Prod. wise Class. of Loans' file to three additional output folders:
    1. outputs/NBD_MF_20_C1_C6/{latest-date-folder}/
    2. outputs/NBD_QF_23_C8/{latest-date-folder}/
    3. outputs/NBD-QF-23-C3-C10-SF-FI/{latest-date-folder}/Input/

    Before copying, deletes old 'Prod. wise Class. of Loans*.xlsb' files.
    Uses completed_dates.txt to find the latest date folder.
    """
    import shutil

    # Define the three target output folders
    target_configs = [
        {
            "name": "C1-C6",
            "base_folder": OUTPUTS_DIR / "NBD_MF_20_C1_C6",
            "subfolder": None  # Copy directly to date folder
        },
        {
            "name": "C8",
            "base_folder": OUTPUTS_DIR / "NBD_QF_23_C8",
            "subfolder": None  # Copy directly to date folder
        },
        {
            "name": "C3-C10",
            "base_folder": OUTPUTS_DIR / "NBD-QF-23-C3-C10-SF-FI",
            "subfolder": "Input"  # Copy to Input subfolder
        }
    ]

    for config in target_configs:
        try:
            base_folder = config["base_folder"]

            # Check if base folder exists
            if not base_folder.exists():
                _emit_log(f"{config['name']}: Base folder does not exist: {base_folder}", important=True)
                continue

            # Find latest completed date folder
            _, date_folder = find_latest_completed_date(base_folder)

            if not date_folder or not date_folder.exists():
                _emit_log(f"{config['name']}: Could not find latest date folder in {base_folder}", important=True)
                continue

            # Determine the actual target directory
            if config["subfolder"]:
                target_dir = date_folder / config["subfolder"]
                target_dir.mkdir(parents=True, exist_ok=True)
            else:
                target_dir = date_folder

            # Remove old "Prod. wise Class. of Loans*.xlsb" files
            old_files = list(target_dir.glob("Prod. wise Class. of Loans*.xlsb"))
            for old_file in old_files:
                try:
                    old_file.unlink()
                    _emit_log(f"{config['name']}: Removed old file: {old_file.name}", important=True)
                except Exception as e:
                    _emit_log(f"{config['name']}: Failed to remove old file {old_file.name}: {e}", important=True)

            # Copy the new file
            dest_file = target_dir / ia_file_path.name
            shutil.copy2(ia_file_path, dest_file)
            _emit_log(f"{config['name']}: Copied to {dest_file}", important=True)

        except Exception as e:
            _emit_log(f"{config['name']}: Error copying file: {e}", important=True)


def _append_report_message(report_id: str, message: str) -> None:
    """Append a single log line to the per-report buffer for UI display."""
    buf = _report_messages.get(report_id)
    if buf is None:
        _report_messages[report_id] = deque(maxlen=200)
        buf = _report_messages[report_id]
    ts = datetime.now().strftime('%H:%M:%S')
    buf.append(f"[{ts}] {message.strip()}")


def save_master_data(sofp_data: dict | None, ia_data: dict | None, c1c6_data: dict | None = None, sofp_sheet: str | None = None, c1c6_sheet: str | None = None, ga_is_data: dict | None = None, ecs_m2m_data: dict | None = None, ecs_treasury_data: dict | None = None, ecs_manual_payments_data: dict | None = None) -> tuple[bool, str]:
    """Save master data to Master_Data.xlsx"""
    try:
        wb = openpyxl.load_workbook(MASTER_DATA_XLSX)

        # Generate current timestamp
        current_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Save SOFP-SOCI data only if provided and non-empty
        if sofp_data:
            has_sofp_values = any((sofp_data.get("bank_loans"), sofp_data.get("securitization_loans"), sofp_data.get("foreign_funding"),
                                  sofp_data.get("interest_on_lease_liabilities"), sofp_data.get("depreciation_of_right_of_use_assets"),
                                  sofp_data.get("bs_1_1_2_2_0_0"), sofp_data.get("bs_1_2_2_2_0_0"), sofp_data.get("bs_1_1_12_2_0_0")))
            if has_sofp_values:
                # Write strictly to the provided sheet; do not fallback between monthly/annual
                if sofp_sheet and sofp_sheet in wb.sheetnames:
                    ws = wb[sofp_sheet]
                    next_row = ws.max_row + 1
                    ws[f"A{next_row}"] = current_timestamp
                    ws[f"B{next_row}"] = sofp_data.get("bank_loans", "")
                    ws[f"C{next_row}"] = sofp_data.get("securitization_loans", "")
                    ws[f"D{next_row}"] = sofp_data.get("foreign_funding", "")
                    ws[f"E{next_row}"] = sofp_data.get("interest_on_lease_liabilities", "")
                    ws[f"F{next_row}"] = sofp_data.get("depreciation_of_right_of_use_assets", "")
                    ws[f"G{next_row}"] = sofp_data.get("bs_1_1_2_2_0_0", "")
                    ws[f"H{next_row}"] = sofp_data.get("bs_1_2_2_2_0_0", "")
                    ws[f"I{next_row}"] = sofp_data.get("bs_1_1_12_2_0_0", "")

        # Save IA data only if provided and non-empty
        if ia_data:
            has_ia_values = bool(ia_data.get("annual_interest_rate"))
            if has_ia_values and "NBD-MF-23-IA-C1&C2-C5-IP" in wb.sheetnames:
                ws = wb["NBD-MF-23-IA-C1&C2-C5-IP"]
                next_row = ws.max_row + 1
                ws[f"A{next_row}"] = current_timestamp
                ws[f"B{next_row}"] = ia_data.get("annual_interest_rate", "")

        # Save C1-C6 data (Guarantees and Tier 2 capital values)
        if c1c6_data is not None:
            has_c1c6_values = any((c1c6_data.get("guarantees"), c1c6_data.get("tier2_instruments"),
                                  c1c6_data.get("tier2_revaluation"), c1c6_data.get("tier2_provisions"),
                                  c1c6_data.get("new_audit_completed")))
            if has_c1c6_values:
                # Write strictly to the provided sheet; do not fallback between monthly/annual
                if c1c6_sheet and c1c6_sheet in wb.sheetnames:
                    ws = wb[c1c6_sheet]
                    next_row = ws.max_row + 1
                    ws[f"A{next_row}"] = current_timestamp
                    ws[f"B{next_row}"] = c1c6_data.get("guarantees", "")
                    # Column mapping for Tier 2 capital: C=tier2_instruments, D=tier2_revaluation, E=tier2_provisions
                    ws[f"C{next_row}"] = c1c6_data.get("tier2_instruments", "")
                    ws[f"D{next_row}"] = c1c6_data.get("tier2_revaluation", "")
                    ws[f"E{next_row}"] = c1c6_data.get("tier2_provisions", "")
                    # Column F: audited? (Yes/No)
                    audit_completed = c1c6_data.get("new_audit_completed", "")
                    ws[f"F{next_row}"] = "Yes" if audit_completed == "yes" else "No"
                    # Columns G-K: Capital values (only saved when audit is completed)
                    ws[f"G{next_row}"] = c1c6_data.get("stated_capital", "")
                    ws[f"H{next_row}"] = c1c6_data.get("statutory_reserve_fund", "")
                    ws[f"I{next_row}"] = c1c6_data.get("fair_value_through_other_comprehensive_income_reserve", "")
                    ws[f"J{next_row}"] = c1c6_data.get("revaluation_reserve", "")
                    ws[f"K{next_row}"] = c1c6_data.get("retained_earnings", "")

        # Save GA & IS data (DEFERRED TAX PAYABLE)
        if ga_is_data is not None:
            has_ga_is_values = any((ga_is_data.get("deferred_tax_0_3m"), ga_is_data.get("deferred_tax_3_6m"),
                                   ga_is_data.get("deferred_tax_6_12m"), ga_is_data.get("deferred_tax_1_3y"),
                                   ga_is_data.get("deferred_tax_3_5y"), ga_is_data.get("deferred_tax_over_5y")))
            if has_ga_is_values and "NBD-MF-10-GA_IS" in wb.sheetnames:
                ws = wb["NBD-MF-10-GA_IS"]
                next_row = ws.max_row + 1
                ws[f"A{next_row}"] = current_timestamp
                ws[f"B{next_row}"] = ga_is_data.get("deferred_tax_0_3m", "")
                ws[f"C{next_row}"] = ga_is_data.get("deferred_tax_3_6m", "")
                ws[f"D{next_row}"] = ga_is_data.get("deferred_tax_6_12m", "")
                ws[f"E{next_row}"] = ga_is_data.get("deferred_tax_1_3y", "")
                ws[f"F{next_row}"] = ga_is_data.get("deferred_tax_3_5y", "")
                ws[f"G{next_row}"] = ga_is_data.get("deferred_tax_over_5y", "")

        # Save ECS M2M data (7 daily Net M2M Gain/Loss values) and Treasury Input values
        if (ecs_m2m_data is not None or ecs_treasury_data is not None) and "NBD_WF_18_ECS" in wb.sheetnames:
            ws = wb["NBD_WF_18_ECS"]
            if ecs_m2m_data is not None:
                has_m2m_values = any(ecs_m2m_data.get(f"day{i}") not in (None, "") for i in range(1, 8))
                if has_m2m_values:
                    next_row = ws.max_row + 1
                    ws[f"A{next_row}"] = ecs_m2m_data.get("day1", "")
                    ws[f"B{next_row}"] = ecs_m2m_data.get("day2", "")
                    ws[f"C{next_row}"] = ecs_m2m_data.get("day3", "")
                    ws[f"D{next_row}"] = ecs_m2m_data.get("day4", "")
                    ws[f"E{next_row}"] = ecs_m2m_data.get("day5", "")
                    ws[f"F{next_row}"] = ecs_m2m_data.get("day6", "")
                    ws[f"G{next_row}"] = ecs_m2m_data.get("day7", "")
            if ecs_treasury_data is not None:
                _treasury_cols = {
                    "com_paper_cap":     "H",
                    "com_paper_int_pay": "I",
                    "sec_loan_cap":      "J",
                    "sec_loan_int_pay":  "K",
                    "repo_cap":          "L",
                    "repo_int_rec":      "M",
                    "cost_of_inv":       "N",
                    "int_rec":           "O",
                    "m2m_gain":          "P",
                }
                has_treasury_values = any(ecs_treasury_data.get(k) not in (None, "") for k in _treasury_cols)
                if has_treasury_values:
                    # Write to the same last row (or a new row if no M2M was written)
                    target_row = ws.max_row if ws.max_row > 1 else ws.max_row + 1
                    for field, col in _treasury_cols.items():
                        ws[f"{col}{target_row}"] = ecs_treasury_data.get(field, "")

        # Save ECS Manual Payments data (cols Q-Y)
        _manual_payments_cols = {
            "current_week":      "Q",
            "cbsl":              "R",
            "epf_etf_ot":        "S",
            "outsource_staff":   "T",
            "apiit":             "U",
            "fd":                "V",
            "taxes":             "W",
            "dividend_payment":  "X",
            "sscl":              "Y",
        }
        if ecs_manual_payments_data is not None and "NBD_WF_18_ECS" in wb.sheetnames:
            has_values = any(ecs_manual_payments_data.get(k) not in (None, "") for k in _manual_payments_cols)
            if has_values:
                ws = wb["NBD_WF_18_ECS"]
                next_row = ws.max_row + 1
                for field, col in _manual_payments_cols.items():
                    ws[f"{col}{next_row}"] = ecs_manual_payments_data.get(field, "")

        wb.save(MASTER_DATA_XLSX)
        wb.close()
        return True, "Master data saved successfully"
    except Exception as e:
        return False, f"Failed to save master data: {e}"


def get_latest_master_data() -> tuple[dict, dict, dict, dict, dict]:
    """Get latest master data from Excel sheets for pre-filling form"""
    sofp_data = {"timestamp": "", "bank_loans": "", "securitization_loans": "", "foreign_funding": "",
                 "interest_on_lease_liabilities": "", "depreciation_of_right_of_use_assets": "",
                 "bs_1_1_2_2_0_0": "", "bs_1_2_2_2_0_0": "", "bs_1_1_12_2_0_0": ""}
    ia_data = {"timestamp": "", "annual_interest_rate": ""}
    c1c6_data = {"timestamp": "", "guarantees": "", "tier2_instruments": "", "tier2_revaluation": "", "tier2_provisions": "",
                 "new_audit_completed": "", "stated_capital": "", "statutory_reserve_fund": "",
                 "fair_value_through_other_comprehensive_income_reserve": "", "revaluation_reserve": "", "retained_earnings": ""}
    ga_is_data = {"timestamp": "", "deferred_tax_0_3m": "", "deferred_tax_3_6m": "", "deferred_tax_6_12m": "",
                  "deferred_tax_1_3y": "", "deferred_tax_3_5y": "", "deferred_tax_over_5y": ""}
    ecs_treasury_data = {"com_paper_cap": "", "com_paper_int_pay": "", "sec_loan_cap": "", "sec_loan_int_pay": "",
                         "repo_cap": "", "repo_int_rec": "", "cost_of_inv": "", "int_rec": "", "m2m_gain": ""}

    try:
        wb = openpyxl.load_workbook(MASTER_DATA_XLSX)

        # Get latest SOFP-SOCI data
        sofp_sheet = None
        if "NBD-MF-01-SOFP-SOCI" in wb.sheetnames:
            sofp_sheet = "NBD-MF-01-SOFP-SOCI"
        elif "NBD-AF-01-SOFP-SOCI" in wb.sheetnames:
            sofp_sheet = "NBD-AF-01-SOFP-SOCI"

        if sofp_sheet:
            ws = wb[sofp_sheet]
            if ws.max_row > 1:  # Has data beyond header
                last_row = ws.max_row
                sofp_data = {
                    "timestamp": str(ws[f"A{last_row}"].value or ""),
                    "bank_loans": str(ws[f"B{last_row}"].value or ""),
                    "securitization_loans": str(ws[f"C{last_row}"].value or ""),
                    "foreign_funding": str(ws[f"D{last_row}"].value or ""),
                    "interest_on_lease_liabilities": str(ws[f"E{last_row}"].value or ""),
                    "depreciation_of_right_of_use_assets": str(ws[f"F{last_row}"].value or ""),
                    "bs_1_1_2_2_0_0": str(ws[f"G{last_row}"].value or ""),
                    "bs_1_2_2_2_0_0": str(ws[f"H{last_row}"].value or ""),
                    "bs_1_1_12_2_0_0": str(ws[f"I{last_row}"].value or "")
                }

        # Get latest IA data
        if "NBD-MF-23-IA-C1&C2-C5-IP" in wb.sheetnames:
            ws = wb["NBD-MF-23-IA-C1&C2-C5-IP"]
            if ws.max_row > 1:  # Has data beyond header
                last_row = ws.max_row
                ia_data = {
                    "timestamp": str(ws[f"A{last_row}"].value or ""),
                    "annual_interest_rate": str(ws[f"B{last_row}"].value or "")
                }

        # Get latest C1-C6 data (Guarantees)
        c1c6_sheet = None
        if "NBD-MF-20-C1-C6" in wb.sheetnames:
            c1c6_sheet = "NBD-MF-20-C1-C6"
        elif "NBD-AF-20-C1-C6" in wb.sheetnames:
            c1c6_sheet = "NBD-AF-20-C1-C6"

        if c1c6_sheet:
            ws = wb[c1c6_sheet]
            if ws.max_row > 1:
                last_row = ws.max_row
                c1c6_data["timestamp"] = str(ws[f"A{last_row}"].value or "")
                c1c6_data["guarantees"] = str(ws[f"B{last_row}"].value or "")

        # Get Tier 2 capital values and audit completion data from NBD-MF-20-C1-C6 sheet only
        if "NBD-MF-20-C1-C6" in wb.sheetnames:
            ws_tier2 = wb["NBD-MF-20-C1-C6"]
            if ws_tier2.max_row > 1:
                last_row_tier2 = ws_tier2.max_row
                c1c6_data["tier2_instruments"] = str(ws_tier2[f"C{last_row_tier2}"].value or "")
                c1c6_data["tier2_revaluation"] = str(ws_tier2[f"D{last_row_tier2}"].value or "")
                c1c6_data["tier2_provisions"] = str(ws_tier2[f"E{last_row_tier2}"].value or "")
                # Read audit completion status (column F)
                audit_status = str(ws_tier2[f"F{last_row_tier2}"].value or "")
                c1c6_data["new_audit_completed"] = "yes" if audit_status == "Yes" else "no"
                # Read capital values (columns G-K)
                c1c6_data["stated_capital"] = str(ws_tier2[f"G{last_row_tier2}"].value or "")
                c1c6_data["statutory_reserve_fund"] = str(ws_tier2[f"H{last_row_tier2}"].value or "")
                c1c6_data["fair_value_through_other_comprehensive_income_reserve"] = str(ws_tier2[f"I{last_row_tier2}"].value or "")
                c1c6_data["revaluation_reserve"] = str(ws_tier2[f"J{last_row_tier2}"].value or "")
                c1c6_data["retained_earnings"] = str(ws_tier2[f"K{last_row_tier2}"].value or "")

        # Get latest GA & IS data (DEFERRED TAX PAYABLE)
        if "NBD-MF-10-GA_IS" in wb.sheetnames:
            ws = wb["NBD-MF-10-GA_IS"]
            if ws.max_row > 1:  # Has data beyond header
                last_row = ws.max_row
                ga_is_data = {
                    "timestamp": str(ws[f"A{last_row}"].value or ""),
                    "deferred_tax_0_3m": str(ws[f"B{last_row}"].value or ""),
                    "deferred_tax_3_6m": str(ws[f"C{last_row}"].value or ""),
                    "deferred_tax_6_12m": str(ws[f"D{last_row}"].value or ""),
                    "deferred_tax_1_3y": str(ws[f"E{last_row}"].value or ""),
                    "deferred_tax_3_5y": str(ws[f"F{last_row}"].value or ""),
                    "deferred_tax_over_5y": str(ws[f"G{last_row}"].value or "")
                }

        # Get latest ECS Treasury Input values
        if "NBD_WF_18_ECS" in wb.sheetnames:
            ws = wb["NBD_WF_18_ECS"]
            if ws.max_row > 1:
                last_row = ws.max_row
                ecs_treasury_data = {
                    "com_paper_cap":     str(ws[f"H{last_row}"].value or ""),
                    "com_paper_int_pay": str(ws[f"I{last_row}"].value or ""),
                    "sec_loan_cap":      str(ws[f"J{last_row}"].value or ""),
                    "sec_loan_int_pay":  str(ws[f"K{last_row}"].value or ""),
                    "repo_cap":          str(ws[f"L{last_row}"].value or ""),
                    "repo_int_rec":      str(ws[f"M{last_row}"].value or ""),
                    "cost_of_inv":       str(ws[f"N{last_row}"].value or ""),
                    "int_rec":           str(ws[f"O{last_row}"].value or ""),
                    "m2m_gain":          str(ws[f"P{last_row}"].value or ""),
                }

        wb.close()
    except Exception as e:
        # Return empty data if file doesn't exist or can't be read
        pass

    return sofp_data, ia_data, c1c6_data, ga_is_data, ecs_treasury_data


def get_la_treasury_amounts() -> dict:
    """Get LA treasury amounts from Master_Data.xlsx"""
    amounts = {}
    try:
        if MASTER_DATA_XLSX.exists():
            wb = openpyxl.load_workbook(MASTER_DATA_XLSX, data_only=True)
            if "NBD-WF-15-LA" in wb.sheetnames:
                ws = wb["NBD-WF-15-LA"]
                # Read amounts from row 2 (row 1 has headers)
                amounts = {
                    'com_papers_cap': ws.cell(row=2, column=1).value or "",
                    'com_papers_init_payable': ws.cell(row=2, column=2).value or "",
                    'sec_laon_cap': ws.cell(row=2, column=3).value or "",
                    'sec_laon_int_payable': ws.cell(row=2, column=4).value or "",
                    'rel_party_loans_cap': ws.cell(row=2, column=5).value or "",
                    'rel_party_loans_int_pay': ws.cell(row=2, column=6).value or "",
                    'repo_cap': ws.cell(row=2, column=7).value or "",
                    'repo_int_recei': ws.cell(row=2, column=8).value or "",
                    'cost_invest': ws.cell(row=2, column=9).value or "",
                    'int_recei': ws.cell(row=2, column=10).value or "",
                    'M2M_gain': ws.cell(row=2, column=11).value or "",
                }
            wb.close()
    except Exception as e:
        pass  # Return empty dict on error
    return amounts


def get_latest_uploaded_files() -> dict:
    """Get latest uploaded files from Master Inputs directories"""
    file_info = {}

    categories = [
        "ALCL Management Accounts",
        "Investment Schedule",
        "Loan Schedule",
        "Supporting Schedules",
        "IFRS",
        "Borrowing report",
        "Disbursement with Budget",
        "Information Request from Credit",
        "Net Portfolio",
        "YARD STOCK AS AT",
        "Cadre",
        "Unutilized Amount",
        "Rec Target",
        "Daily Bank Balances",
        "M2M",
        "BankBalanceGLCodes",
        "Product Category",
        "Prod_Cat C4",
        "CBSL Sector Request",
        "JOBHD",
        "Sector Report",
        "Business Gender",
        "Information Request from IT",
        "Client Available Limits 1",
        "Client Available Limits 2",
        "Client Available Limits 3"
    ]

    for category in categories:
        # Map "Product Category" to "Prod Cat" folder
        if category == "Product Category":
            category_dir = MASTER_INPUTS_DIR / "Prod Cat"
        # Map "Prod_Cat C4" to "Prod Cat C4" folder
        elif category == "Prod_Cat C4":
            category_dir = MASTER_INPUTS_DIR / "Prod Cat C4"
        # Map all Client Available Limits to single folder
        elif category in ["Client Available Limits 1", "Client Available Limits 2", "Client Available Limits 3"]:
            category_dir = MASTER_INPUTS_DIR / "Client Available Limits"
        else:
            category_dir = MASTER_INPUTS_DIR / category
        if category_dir.exists():
            # Find the most recent file (.xlsx for most categories, .xlsm for Borrowing report, .xlsb for Sector Report)
            if category == "Borrowing report":
                files = list(category_dir.glob("*.xlsm"))
            elif category == "Sector Report":
                # Only check for .xlsb files
                files = list(category_dir.glob("*.xlsb"))
            elif category in ["Client Available Limits 1", "Client Available Limits 2", "Client Available Limits 3"]:
                # For Client Available Limits, get all files and sort by modification time
                files = list(category_dir.glob("*.xlsx")) + list(category_dir.glob("*.xls"))
                files = sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)
                # Display up to 3 files based on category number
                idx = int(category.split()[-1]) - 1  # Extract index from "Client Available Limits 1" -> 0
                if idx < len(files):
                    file_info[category] = files[idx].name
                else:
                    file_info[category] = "No files uploaded"
                continue  # Skip the normal file handling below
            else:
                files = list(category_dir.glob("*.xlsx"))

            if files:
                latest_file = max(files, key=lambda x: x.stat().st_mtime)
                # For BankBalanceGLCodes, include timestamp
                if category == "BankBalanceGLCodes":
                    from datetime import datetime
                    timestamp = latest_file.stat().st_mtime
                    dt = datetime.fromtimestamp(timestamp)
                    formatted_time = dt.strftime("%Y-%m-%d %H:%M:%S")
                    file_info[category] = f"{latest_file.name} (Updated: {formatted_time})"
                else:
                    file_info[category] = latest_file.name
            else:
                file_info[category] = "No files uploaded"
        else:
            file_info[category] = "No files uploaded"

    return file_info


def get_ecs_master_input_info() -> dict:
    """Get latest file info for ECS master input upload panel."""
    def _latest_name(subfolder):
        d = MASTER_INPUTS_DIR / subfolder
        if d.exists():
            files = list(d.glob("*.xlsx"))
            if files:
                return max(files, key=lambda f: f.stat().st_mtime).name
        return None

    def _latest_name_with_mtime(subfolder):
        d = MASTER_INPUTS_DIR / subfolder
        if d.exists():
            files = list(d.glob("*.xlsx"))
            if files:
                f = max(files, key=lambda f: f.stat().st_mtime)
                mtime = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d-%m-%Y %H:%M")
                return f"{f.name} (Updated: {mtime})"
        return None

    return {
        "afl_borrowing": _latest_name("AFL Borrowing schedule"),
        "cash_flow": _latest_name("Cash Flow"),
        "payments_receipts": _latest_name_with_mtime("Payments and Receipts"),
        "daily_bank_balances": _latest_name_with_mtime("Daily Bank Balances"),
        "payments_latest": _latest_name_with_mtime("Payments"),
        "types_of_receipt_latest": _latest_name_with_mtime("Types of Receipt"),
        "sasianet_latest": _latest_name_with_mtime("SasiaNet"),
        "tax_payments_latest": _latest_name_with_mtime("Tax Payments"),
        "cashflow_movement_latest": _latest_name_with_mtime("Cashflow Movement"),
    }


def copy_uploaded_file_to_sofp_outputs(category: str, file_path: str, month: str) -> None:
    """Copy a single uploaded file to SOFP outputs folder immediately"""
    try:
        # Find the latest SOFP outputs date folder
        sofp_outputs_dir = OUTPUTS_DIR / "NBD_MF_01_SOFP_SOCI"
        if not sofp_outputs_dir.exists():
            _emit_log("SOFP outputs directory not found", important=True)
            return
            
        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(sofp_outputs_dir)
        if not date_folder or not date_folder.exists():
            _emit_log("No completed date folder found for SOFP", important=True)
            return
            
        # Use the original filename (already renamed when stored)
        new_name = Path(file_path).name
            
        # Remove existing files with same prefix
        removed = 0
        _emit_log(f"Scanning {date_folder} for existing files to remove for category: {category}", important=True)
        
        for f in date_folder.iterdir():
            if f.is_file() and f.suffix.lower() == ".xlsx":
                _emit_log(f"Found file: {f.name}", important=True)
                should_remove = False
                
                if category == "ALCL Management Accounts" and f.name.startswith("ALCL Management Accounts"):
                    should_remove = True
                    _emit_log(f"Marking for removal (ALCL): {f.name}", important=True)
                elif category == "Investment Schedule" and f.name.startswith("Investment Schedule"):
                    should_remove = True
                    _emit_log(f"Marking for removal (Investment): {f.name}", important=True)
                elif category == "Loan Schedule" and f.name.startswith("Loan Schedule"):
                    should_remove = True
                    _emit_log(f"Marking for removal (Loan): {f.name}", important=True)
                elif category == "Supporting Schedules" and f.name.startswith("Supporting Schedules"):
                    should_remove = True
                    _emit_log(f"Marking for removal (Supporting): {f.name}", important=True)
                
                if should_remove:
                    try:
                        f.unlink()
                        removed += 1
                        _emit_log(f"Successfully removed existing file: {f.name}", important=True)
                    except Exception as e:
                        _emit_log(f"ERROR: Failed to remove existing file {f.name}: {e}", important=True)
        
        if removed:
            _emit_log(f"Removed {removed} existing {category} file(s) in {date_folder}", important=True)
        
        # Copy the file
        dest_path = date_folder / new_name
        shutil.copy2(file_path, dest_path)
        _emit_log(f"Copied {category} file to {dest_path}", important=True)
        
    except Exception as e:
        _emit_log(f"Failed to copy {category} file to SOFP outputs: {e}", important=True)


def validate_sofp_file_for_ia(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Validate that SOFP file exists for the selected month before running IA"""
    try:
        # Get the selected date from the first report (they should all be the same)
        if not report_ids_with_dates:
            return False, "No report data found"
            
        first_report_id = list(report_ids_with_dates.keys())[0]
        selected_date_str = report_ids_with_dates[first_report_id][0]  # Get the date string
        
        # Parse the selected date to get month and year
        from datetime import datetime
        try:
            # Try different date formats
            try:
                selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d")  # YYYY-MM-DD format
            except ValueError:
                try:
                    selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")  # DD/MM/YYYY format
                except ValueError:
                    try:
                        selected_date = datetime.strptime(selected_date_str, "%m/%d/%Y")  # MM/DD/YYYY format
                    except ValueError:
                        return False, f"Invalid date format: {selected_date_str}"
            
            month_name = selected_date.strftime("%B")  # Full month name like "October"
            year = selected_date.year
        except Exception as e:
            return False, f"Error parsing date: {e}"
        
        # Check if SOFP file exists in IA outputs using selected month and year
        ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
        if not ia_outputs_dir.exists():
            return False, f"IA outputs directory not found"
            
        # Find latest completed date folder
        last_date, date_folder = find_latest_completed_date(ia_outputs_dir)
        if not date_folder or not date_folder.exists():
            return False, f"No completed IA date folder found"
        
        # Look for the SOFP file with the expected name using selected month and year
        expected_filename = f"NBD-MF-01-SOFP & SOCI AFL Monthly FS {month_name} {year}.xlsx"
        sofp_file_path = date_folder / expected_filename

        if not sofp_file_path.exists():
            # Try to find any SOFP file for debugging
            import glob
            sofp_pattern = str(date_folder / "NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx")
            found_files = glob.glob(sofp_pattern)
            if found_files:
                _emit_log(f"SOFP file exists but name doesn't match. Expected: '{expected_filename}', Found: {[Path(f).name for f in found_files]}", important=True)
                return False, f"'{expected_filename}' file not available. Please run NBD-MF-01-SOFP-SOCI - Statement of Financial position first for {month_name} {year}."
            else:
                return False, f"'{expected_filename}' file not available. Please run NBD-MF-01-SOFP-SOCI - Statement of Financial position first for {month_name} {year}."

        _emit_log(f"SOFP file validation passed for IA: {expected_filename} found", important=True)
        return True, "SOFP file validation passed"
        
    except Exception as e:
        return False, f"Error validating SOFP file for IA: {e}"


def validate_ia_file_for_c1c2(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Validate that IA file exists for the selected month before running C1-C2"""
    try:
        # Get the selected date from the first report (they should all be the same)
        if not report_ids_with_dates:
            return False, "No report data found"
            
        first_report_id = list(report_ids_with_dates.keys())[0]
        selected_date_str = report_ids_with_dates[first_report_id][0]  # Get the date string
        
        # Parse the selected date to get month and year
        from datetime import datetime
        try:
            # Try different date formats
            try:
                selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d")  # YYYY-MM-DD format
            except ValueError:
                try:
                    selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")  # DD/MM/YYYY format
                except ValueError:
                    try:
                        selected_date = datetime.strptime(selected_date_str, "%m/%d/%Y")  # MM/DD/YYYY format
                    except ValueError:
                        return False, f"Invalid date format: {selected_date_str}"
            
            month_name = selected_date.strftime("%B")  # Full month name like "October"
            year = selected_date.year
        except Exception as e:
            return False, f"Error parsing date: {e}"
        
        # Check if IA file exists in IA outputs using selected month and year
        ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
        if not ia_outputs_dir.exists():
            return False, f"IA outputs directory not found"
            
        # Find latest completed date folder
        last_date, date_folder = find_latest_completed_date(ia_outputs_dir)
        if not date_folder or not date_folder.exists():
            return False, f"No completed IA date folder found"
        
        # Look for the IA file with the expected name using selected month and year
        expected_filename = f"Prod. wise Class. of Loans - {month_name} {year}.xlsb"
        ia_file_path = date_folder / expected_filename
        
        if not ia_file_path.exists():
            return False, f"Prod. wise Class. of Loans - {month_name} {year}.xlsb file not available, run NBD-MF-23-IA - Product wise Classification of Loans first for the {month_name} {year}."
        
        _emit_log(f"IA file validation passed for C1-C2: {expected_filename} found", important=True)
        return True, "IA file validation passed"
        
    except Exception as e:
        return False, f"Error validating IA file for C1-C2: {e}"


def validate_set7_required_files(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Validate if required files exist for Set 7 reports (C2-C6)"""
    try:
        # Get the first report's date to determine month/year
        first_report_id = next(iter(report_ids_with_dates.keys()))
        date_data = report_ids_with_dates[first_report_id]
        
        # Handle both tuple format (date_str, month, year) and string format
        if isinstance(date_data, tuple):
            date_str = date_data[0]  # Extract date string from tuple
        else:
            date_str = date_data  # Use string directly
        
        # Parse the date to get month and year
        try:
            # Try YYYY-MM-DD format first
            parsed_date = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            try:
                # Try DD/MM/YYYY format
                parsed_date = datetime.strptime(date_str, "%d/%m/%Y")
            except ValueError:
                # Try MM/DD/YYYY format
                parsed_date = datetime.strptime(date_str, "%m/%d/%Y")
        
        month_name = parsed_date.strftime("%B")
        year = parsed_date.year
        
        # Check if C1-C6 outputs directory exists
        c1c6_outputs_dir = OUTPUTS_DIR / "NBD_MF_20_C1_C6"
        if not c1c6_outputs_dir.exists():
            return False, "C1-C6 outputs directory not found. Please run C1-C6 reports first."
            
        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(c1c6_outputs_dir)
        if not date_folder or not date_folder.exists():
            return False, f"No completed C1-C6 date folder found. Please run C1-C6 reports first for {month_name} {year}."
            
        # Check for the IA file (Prod. wise Class. of Loans)
        ia_filename = f"Prod. wise Class. of Loans - {month_name} {year}.xlsb"
        ia_file_path = date_folder / ia_filename
        
        if not ia_file_path.exists():
            return False, f"'{ia_filename}' file not available. Please run NBD-MF-23-IA - Product wise Classification of Loans first for {month_name} {year}."
            
        # Check for the SOFP file
        sofp_filename = f"NBD-MF-01-SOFP & SOCI AFL Monthly FS {month_name} {year}.xlsx"
        sofp_file_path = date_folder / sofp_filename
        
        if not sofp_file_path.exists():
            return False, f"'{sofp_filename}' file not available. Please run NBD-MF-01-SOFP-SOCI - Statement of Financial position first for {month_name} {year}."
            
        return True, f"Set 7 validation passed: '{ia_filename}' and '{sofp_filename}' found"
        
    except Exception as e:
        return False, f"Error validating Set 7 required files: {str(e)}"


def validate_prod_wise_file_for_c8(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Validate that Prod. wise Class. of Loans file exists for the selected month before running C8"""
    try:
        # Get the selected date from the first report
        if not report_ids_with_dates:
            return False, "No report data found"

        first_report_id = list(report_ids_with_dates.keys())[0]
        selected_date_str = report_ids_with_dates[first_report_id][0]  # Get the date string

        # Parse the selected date to get month and year
        try:
            # Try different date formats
            try:
                selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d")  # YYYY-MM-DD format
            except ValueError:
                try:
                    selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")  # DD/MM/YYYY format
                except ValueError:
                    try:
                        selected_date = datetime.strptime(selected_date_str, "%m/%d/%Y")  # MM/DD/YYYY format
                    except ValueError:
                        return False, f"Invalid date format: {selected_date_str}"

            month_name = selected_date.strftime("%B")  # Full month name like "October"
            year = selected_date.year
        except Exception as e:
            return False, f"Error parsing date: {e}"

        # Check if Prod. wise Class. of Loans file exists in C8 outputs using selected month and year
        c8_outputs_dir = OUTPUTS_DIR / "NBD_QF_23_C8"
        if not c8_outputs_dir.exists():
            return False, f"C8 outputs directory not found"

        # Find latest completed date folder
        last_date, date_folder = find_latest_completed_date(c8_outputs_dir)
        if not date_folder or not date_folder.exists():
            return False, f"No completed C8 date folder found"

        # Look for the Prod. wise Class. of Loans file with the expected name
        expected_filename = f"Prod. wise Class. of Loans - {month_name} {year}.xlsb"
        prod_wise_file_path = date_folder / expected_filename

        if not prod_wise_file_path.exists():
            # Try to find any Prod. wise file for debugging
            import glob
            prod_wise_pattern = str(date_folder / "Prod. wise Class. of Loans*.xlsb")
            found_files = glob.glob(prod_wise_pattern)
            if found_files:
                _emit_log(f"Prod. wise file exists but name doesn't match. Expected: '{expected_filename}', Found: {[Path(f).name for f in found_files]}", important=True)
                return False, f"'{expected_filename}' file not available. Please run NBD-MF-23-IA - Product wise Classification of Loans first for {month_name} {year}."
            else:
                return False, f"'{expected_filename}' file not available. Please run NBD-MF-23-IA - Product wise Classification of Loans first for {month_name} {year}."

        _emit_log(f"Prod. wise file validation passed for C8: {expected_filename} found", important=True)
        return True, "Prod. wise file validation passed"

    except Exception as e:
        return False, f"Error validating Prod. wise file for C8: {e}"


def validate_deposit_liability_file_for_ecs(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Validate that Deposit Liability file (NBD-WF-18-DM) exists with matching date before running ECS."""
    try:
        import re

        if "NBD-WF-18-ECS" not in report_ids_with_dates:
            return True, "ECS report not selected, skipping validation"

        selected_date_str = report_ids_with_dates["NBD-WF-18-ECS"][0]

        try:
            selected_date = datetime.strptime(selected_date_str, "%m/%d/%Y")
        except ValueError:
            try:
                selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d")
            except ValueError:
                try:
                    selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")
                except ValueError:
                    return False, f"Invalid date format: {selected_date_str}"

        expected_date_str = selected_date.strftime("%d-%m-%Y")
        expected_filename = f"NBD-WF-18-DM Deposit Liability ({expected_date_str}).xlsx"

        # Check ECS outputs folder for the DM file
        ecs_outputs_dir = OUTPUTS_DIR / "NBD_WF_18_ECS"
        search_folder = None

        if ecs_outputs_dir.exists():
            _, date_folder = find_latest_completed_date(ecs_outputs_dir)
            if date_folder and date_folder.exists():
                search_folder = date_folder

        if search_folder is None:
            return False, f"No ECS outputs folder found. Please run the DM report for {expected_date_str} first to generate: {expected_filename}"

        dm_files = list(search_folder.glob("NBD-WF-18-DM Deposit Liability*.xlsx"))

        if not dm_files:
            return False, f"Deposit Liability file not found in ECS outputs. Please run the NBD-WF-18-DM report for {expected_date_str} first to generate: {expected_filename}"

        for dm_file in dm_files:
            date_match = re.search(r'\((\d{2}-\d{2}-\d{4})\)', dm_file.name)
            if date_match:
                file_date_str = date_match.group(1)
                if file_date_str == expected_date_str:
                    _emit_log(f"✓ ECS Deposit Liability file validated: {dm_file.name}", important=True)
                    return True, "Deposit Liability file validation passed"
                else:
                    return False, (
                        f"Deposit Liability file date mismatch: selected date is {expected_date_str} "
                        f"but file has date {file_date_str}. "
                        f"Please run the NBD-WF-18-DM report for {expected_date_str} first to generate: {expected_filename}"
                    )

        _emit_log(f"Found DM file but could not extract date: {[f.name for f in dm_files]}", important=True)
        return True, "Deposit Liability file found (date could not be verified)"

    except Exception as e:
        return False, f"Error validating Deposit Liability file for ECS: {e}"


def validate_deposit_liability_file_for_la(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Validate that Deposit Liability file (NBD-WF-18-DM) exists with matching date before running LA"""
    try:
        import re
        from datetime import datetime

        # Get the selected date from NBD-WF-15-LA report
        if "NBD-WF-15-LA" not in report_ids_with_dates:
            return True, "LA report not selected, skipping validation"

        selected_date_str = report_ids_with_dates["NBD-WF-15-LA"][0]  # Get the date string

        # Parse the selected date
        try:
            # Try MM/DD/YYYY format first (UI format)
            selected_date = datetime.strptime(selected_date_str, "%m/%d/%Y")
        except ValueError:
            try:
                selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d")
            except ValueError:
                try:
                    selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")
                except ValueError:
                    return False, f"Invalid date format: {selected_date_str}"

        # Expected date string format in filename: dd-mm-yyyy
        expected_date_str = selected_date.strftime("%d-%m-%Y")

        # Check LA working folder for NBD-WF-18-DM file
        # If working folder is empty, check outputs folder instead (will be copied during run)
        la_working_dir = WORKING_DIR / "NBD_WF_15_LA"
        la_outputs_dir = OUTPUTS_DIR / "NBD_WF_15_LA"

        working_folder = None

        # First try working folder
        if la_working_dir.exists():
            date_folders = [p for p in la_working_dir.iterdir() if p.is_dir()]
            if date_folders:
                working_folder = date_folders[0]

        # If working folder is empty, check outputs folder
        if working_folder is None:
            _emit_log("LA working folder is empty, checking outputs folder for validation...", important=True)
            if la_outputs_dir.exists():
                last_date, date_folder = find_latest_completed_date(la_outputs_dir)
                if date_folder and date_folder.exists():
                    working_folder = date_folder
                    _emit_log(f"Using LA outputs folder for validation: {date_folder.name}", important=True)

        if working_folder is None:
            return False, "No date folder found in LA working or outputs directory. Please ensure LA has been run at least once."

        # Find NBD-WF-18-DM files
        dm_files = list(working_folder.glob("NBD-WF-18-DM*.xlsx"))

        if not dm_files:
            return False, f"Deposit Liability file (NBD-WF-18-DM*.xlsx) not found in {working_folder}. Please run NBD-WF-18-DM automation first."

        # Check if any file has matching date
        for dm_file in dm_files:
            filename = dm_file.name
            # Extract date from filename pattern like "NBD-WF-18-DM Deposit Liability (27-01-2026).xlsx"
            date_match = re.search(r'\((\d{2}-\d{2}-\d{4})\)', filename)
            if date_match:
                file_date_str = date_match.group(1)
                if file_date_str == expected_date_str:
                    _emit_log(f"✓ Deposit Liability file date validated: {filename}", important=True)
                    return True, "Deposit Liability file validation passed"
                else:
                    # Date mismatch
                    return False, f"Date mismatch: Selected date is {expected_date_str} but file has date {file_date_str}. Please run NBD-WF-18-DM automation for {expected_date_str} first, or select {file_date_str} for this automation."

        # File found but no date in filename
        _emit_log(f"Found Deposit Liability file but could not extract date from filename: {[f.name for f in dm_files]}", important=True)
        return True, "Deposit Liability file found (date could not be verified)"

    except Exception as e:
        return False, f"Error validating Deposit Liability file for LA: {e}"


def validate_sofp_file_for_ga_is(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Validate that SOFP file exists for the selected month before running GA IS"""
    try:
        # Get the selected date from the first report (they should all be the same)
        if not report_ids_with_dates:
            return False, "No report data found"
            
        first_report_id = list(report_ids_with_dates.keys())[0]
        selected_date_str = report_ids_with_dates[first_report_id][0]  # Get the date string
        
        # Parse the selected date to get month and year
        from datetime import datetime
        try:
            # Try different date formats
            try:
                selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d")  # YYYY-MM-DD format
            except ValueError:
                try:
                    selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")  # DD/MM/YYYY format
                except ValueError:
                    try:
                        selected_date = datetime.strptime(selected_date_str, "%m/%d/%Y")  # MM/DD/YYYY format
                    except ValueError:
                        return False, f"Invalid date format: {selected_date_str}"
            
            month_name = selected_date.strftime("%B")  # July, August, etc.
            year = selected_date.year
        except Exception as e:
            return False, f"Error parsing date: {e}"
        
        # Check if SOFP file exists in outputs using selected month and year
        sofp_outputs_dir = OUTPUTS_DIR / "NBD-MF-10-GA & NBD-MF-11-IS"
        if not sofp_outputs_dir.exists():
            return False, f"SOFP outputs directory not found"
            
        # Find latest completed date folder
        last_date, date_folder = find_latest_completed_date(sofp_outputs_dir)
        if not date_folder or not date_folder.exists():
            return False, f"No completed SOFP date folder found"
        
        # Look for the SOFP file with the expected name using selected month and year
        expected_filename = f"NBD-MF-01-SOFP & SOCI AFL Monthly FS {month_name} {year}.xlsx"
        sofp_file_path = date_folder / expected_filename

        if not sofp_file_path.exists():
            # Try to find any SOFP file for debugging
            import glob
            sofp_pattern = str(date_folder / "NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx")
            found_files = glob.glob(sofp_pattern)
            if found_files:
                _emit_log(f"SOFP file exists but name doesn't match. Expected: '{expected_filename}', Found: {[Path(f).name for f in found_files]}", important=True)
                return False, f"'{expected_filename}' file not available. Please run NBD-MF-01-SOFP-SOCI - Statement of Financial position first for {month_name} {year}."
            else:
                return False, f"'{expected_filename}' file not available. Please run NBD-MF-01-SOFP-SOCI - Statement of Financial position first for {month_name} {year}."

        _emit_log(f"SOFP file validation passed: {expected_filename} found", important=True)
        return True, "SOFP file validation passed"
        
    except Exception as e:
        return False, f"Error validating SOFP file: {e}"


def cleanup_duplicate_ia_files() -> None:
    """Clean up duplicate files in IA outputs folder, keeping only the latest ones"""
    try:
        # Find the latest IA outputs date folder
        ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
        if not ia_outputs_dir.exists():
            return
            
        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(ia_outputs_dir)
        if not date_folder or not date_folder.exists():
            return
            
        _emit_log(f"Cleaning up duplicate files in {date_folder}", important=True)
        
        # Use the new general cleanup function
        removed_count = cleanup_duplicate_files_in_folder(date_folder)
        if removed_count > 0:
            _emit_log(f"Removed {removed_count} duplicate files from IA folder", important=True)
        else:
            _emit_log(f"No duplicate files found in IA folder", important=True)
                
    except Exception as e:
        _emit_log(f"Error cleaning up duplicate IA files: {e}", important=True)


def get_full_month_name(month_abbr: str) -> str:
    """Convert month abbreviation to full month name"""
    month_mapping = {
        'Jan': 'January', 'Feb': 'February', 'Mar': 'March', 'Apr': 'April',
        'May': 'May', 'Jun': 'June', 'Jul': 'July', 'Aug': 'August',
        'Sep': 'September', 'Oct': 'October', 'Nov': 'November', 'Dec': 'December',
        'January': 'January', 'February': 'February', 'March': 'March', 'April': 'April',
        'May': 'May', 'June': 'June', 'July': 'July', 'August': 'August',
        'September': 'September', 'October': 'October', 'November': 'November', 'December': 'December'
    }
    return month_mapping.get(month_abbr, month_abbr)


def copy_uploaded_files_to_ia_outputs(uploaded_files: dict, months: dict, years: dict) -> None:
    """Copy uploaded files to IA outputs folder after renaming"""
    try:
        # Find the latest IA outputs date folder
        ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
        if not ia_outputs_dir.exists():
            _emit_log("IA outputs directory not found", important=True)
            return
            
        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(ia_outputs_dir)
        if not date_folder or not date_folder.exists():
            _emit_log("No completed date folder found for IA", important=True)
            return
            
        # Process each uploaded file
        for category, file_path in uploaded_files.items():
            if not file_path or not Path(file_path).exists():
                continue
                
            # Get the month and year for this category
            month_field = f"{category.lower().replace(' ', '_')}_month"
            year_field = f"{category.lower().replace(' ', '_')}_year"
            month = months.get(month_field, "")
            year = years.get(year_field, "")
            
            # For file removal, we don't need month/year - just remove files with matching prefix
            # But we still need month/year for the actual file processing
            if not month or not year:
                _emit_log(f"No month/year specified for {category}, skipping file processing", important=True)
                # Don't continue - we still want to remove old files even if month/year are missing
                
            # Use the original filename (already renamed when stored)
            new_name = Path(file_path).name
                
            # Only copy the file if month and year are available
            if month and year:
                dest_path = date_folder / new_name
                shutil.copy2(file_path, dest_path)
                _emit_log(f"Copied {category} file to {dest_path}", important=True)
                
                # AFTER copying, remove ALL OLD files with same prefix (excluding the one we just copied)
                removed = 0
                _emit_log(f"Scanning {date_folder} for OLD files to remove for category: {category}", important=True)

                # Determine the prefix to match based on category
                prefix_to_remove = None
                if category == "Disbursement with Budget":
                    prefix_to_remove = "Disbursement with Budget"
                elif category == "Information Request from Credit":
                    prefix_to_remove = "Information Request from Credit"
                elif category == "Net Portfolio":
                    prefix_to_remove = "Net Portfolio"

                if prefix_to_remove:
                    for f in date_folder.iterdir():
                        if f.is_file() and f.suffix.lower() == ".xlsx":
                            _emit_log(f"Found xlsx file: {f.name}", important=True)
                            if f.name.startswith(prefix_to_remove) and f.name != new_name:
                                try:
                                    _emit_log(f"Removing OLD file: {f.name} (keeping latest: {new_name})", important=True)
                                    f.unlink()
                                    removed += 1
                                except Exception as e:
                                    _emit_log(f"ERROR: Failed to remove OLD file {f.name}: {e}", important=True)
                            elif f.name.startswith(prefix_to_remove) and f.name == new_name:
                                _emit_log(f"Keeping latest file: {f.name}", important=True)
                            else:
                                _emit_log(f"File does not match prefix '{prefix_to_remove}': {f.name}", important=True)

                    if removed:
                        _emit_log(f"Removed {removed} OLD {category} file(s) in {date_folder}, kept latest: {new_name}", important=True)
                    else:
                        _emit_log(f"No OLD {category} files found to remove, latest file: {new_name}", important=True)

                # SHARED FILES: Also copy to QF outputs folder that uses these files
                if category in ["Information Request from Credit", "Net Portfolio"]:
                    _emit_log(f"'{category}' is shared - copying to QF outputs folder as well", important=True)

                    # Copy to QF outputs folder
                    qf_outputs_dir = OUTPUTS_DIR / "NBD-QF-23-C3-C10-SF-FI"
                    if qf_outputs_dir.exists():
                        last_date_qf, qf_date_folder = find_latest_completed_date(qf_outputs_dir)
                        if qf_date_folder and qf_date_folder.exists():
                            qf_input_folder = qf_date_folder / "Input"
                            qf_input_folder.mkdir(parents=True, exist_ok=True)
                            qf_dest_path = qf_input_folder / new_name

                            # Remove old files in QF Input folder with same prefix
                            removed_qf = 0
                            for f in qf_input_folder.iterdir():
                                if f.is_file() and f.name.startswith(category) and f.suffix.lower() == ".xlsx" and f.name != new_name:
                                    try:
                                        _emit_log(f"Removing OLD {category} from QF: {f.name}", important=True)
                                        f.unlink()
                                        removed_qf += 1
                                    except Exception as e:
                                        _emit_log(f"Failed to remove OLD QF file {f.name}: {e}", important=True)

                            # Copy to QF outputs
                            shutil.copy2(file_path, qf_dest_path)
                            _emit_log(f"Copied shared {category} file to QF outputs: {qf_dest_path}", important=True)

                            # Remove "Copy of " prefix if it exists in the copied file
                            if qf_dest_path.exists() and qf_dest_path.name.startswith("Copy of "):
                                new_clean_name = qf_dest_path.name.replace("Copy of ", "", 1)
                                clean_path = qf_dest_path.parent / new_clean_name
                                qf_dest_path.rename(clean_path)
                                _emit_log(f"Removed 'Copy of ' prefix from file: {new_clean_name}", important=True)

                            if removed_qf:
                                _emit_log(f"Removed {removed_qf} OLD {category} file(s) from QF folder", important=True)
                        else:
                            _emit_log(f"No completed date folder found for QF - skipping shared file copy", important=False)
                    else:
                        _emit_log(f"QF outputs directory not found - skipping shared file copy", important=False)

            else:
                _emit_log(f"Skipped copying {category} file due to missing month/year", important=True)

    except Exception as e:
        _emit_log(f"Failed to copy uploaded files to IA outputs: {e}", important=True)


def cleanup_duplicate_c1c2_files() -> None:
    """Clean up duplicate files in C1-C2 outputs folder, keeping only the latest ones"""
    try:
        # Find the latest IA outputs date folder (C1-C2 uses IA outputs)
        ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
        if not ia_outputs_dir.exists():
            return
            
        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(ia_outputs_dir)
        if not date_folder or not date_folder.exists():
            return
            
        _emit_log(f"Cleaning up duplicate C1-C2 files in {date_folder}", important=True)
        
        # Use the new general cleanup function
        removed_count = cleanup_duplicate_files_in_folder(date_folder)
        if removed_count > 0:
            _emit_log(f"Removed {removed_count} duplicate files from C1-C2 folder", important=True)
        else:
            _emit_log(f"No duplicate files found in C1-C2 folder", important=True)
                
    except Exception as e:
        _emit_log(f"Error cleaning up duplicate C1-C2 files: {e}", important=True)


def copy_uploaded_files_to_c1c2_outputs(uploaded_files: dict, dates: dict, months: dict, years: dict) -> None:
    """Copy uploaded files to C1-C2 outputs folder after renaming"""
    try:
        # Find the latest IA outputs date folder (C1-C2 uses IA outputs)
        ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
        if not ia_outputs_dir.exists():
            _emit_log("IA outputs directory not found", important=True)
            return
            
        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(ia_outputs_dir)
        if not date_folder or not date_folder.exists():
            _emit_log("No completed date folder found for IA", important=True)
            return
            
        # Process each uploaded file
        for category, file_path in uploaded_files.items():
            if not file_path or not Path(file_path).exists():
                continue
                
            # Get the date/month/year for this category
            if category == "Cadre":
                date_field = f"{category.lower().replace(' ', '_')}_date"
                date_value = dates.get(date_field, "")
                if not date_value:
                    _emit_log(f"No date specified for {category}, skipping", important=True)
                    continue
            else:
                month_field = f"{category.lower().replace(' ', '_')}_month"
                year_field = f"{category.lower().replace(' ', '_')}_year"
                month = months.get(month_field, "")
                year = years.get(year_field, "")
                if not month or not year:
                    _emit_log(f"No month/year specified for {category}, skipping", important=True)
                    continue
                
            # Use the original filename (already renamed when stored)
            new_name = Path(file_path).name
                
            # Remove ALL existing files with same prefix (regardless of date/month/year)
            removed = 0
            _emit_log(f"Scanning {date_folder} for existing files to remove for category: {category}", important=True)

            # Determine the prefix to match based on category
            prefix_to_remove = None
            if category == "Cadre":
                prefix_to_remove = "Cadre"
            elif category == "Unutilized Amount":
                prefix_to_remove = "Unutilized Amount"

            if prefix_to_remove:
                for f in date_folder.iterdir():
                    if f.is_file() and f.suffix.lower() == ".xlsx":
                        _emit_log(f"Found xlsx file: {f.name}", important=True)
                        if f.name.startswith(prefix_to_remove):
                            try:
                                _emit_log(f"Removing old file: {f.name}", important=True)
                                f.unlink()
                                removed += 1
                            except Exception as e:
                                _emit_log(f"ERROR: Failed to remove existing file {f.name}: {e}", important=True)
                        else:
                            _emit_log(f"File does not match prefix '{prefix_to_remove}': {f.name}", important=True)

                if removed:
                    _emit_log(f"Removed {removed} existing {category} file(s) in {date_folder}", important=True)
                else:
                    _emit_log(f"No old {category} files found to remove", important=True)
            
            # Copy the file
            dest_path = date_folder / new_name
            shutil.copy2(file_path, dest_path)
            _emit_log(f"Copied {category} file to {dest_path}", important=True)
            
    except Exception as e:
        _emit_log(f"Failed to copy uploaded files to C1-C2 outputs: {e}", important=True)


def copy_uploaded_files_to_ga_is_outputs(uploaded_files: dict, dates: dict) -> None:
    """Copy uploaded files to GA IS outputs folder after renaming"""
    try:
        # Find the latest GA IS outputs date folder
        ga_is_outputs_dir = OUTPUTS_DIR / "NBD-MF-10-GA & NBD-MF-11-IS"
        if not ga_is_outputs_dir.exists():
            _emit_log("GA IS outputs directory not found", important=True)
            return
            
        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(ga_is_outputs_dir)
        if not date_folder or not date_folder.exists():
            _emit_log("No completed date folder found for GA IS", important=True)
            return
            
        # Process each uploaded file
        for category, file_path in uploaded_files.items():
            if not file_path or not Path(file_path).exists():
                continue
                
            # Use the original filename (already renamed when stored)
            new_name = Path(file_path).name
                
            # Remove existing files with same prefix
            removed = 0
            _emit_log(f"Scanning {date_folder} for existing files to remove for category: {category}", important=True)
            
            for f in date_folder.iterdir():
                if f.is_file() and f.suffix.lower() in [".xlsm", ".xlsx"]:
                    _emit_log(f"Found file: {f.name}", important=True)
                    should_remove = False
                    
                    if category == "Borrowing report" and f.name.startswith("Borrowing report"):
                        should_remove = True
                        _emit_log(f"Marking for removal (Borrowing report): {f.name}", important=True)
                    
                    if should_remove:
                        try:
                            f.unlink()
                            removed += 1
                            _emit_log(f"Successfully removed existing file: {f.name}", important=True)
                        except Exception as e:
                            _emit_log(f"ERROR: Failed to remove existing file {f.name}: {e}", important=True)
            
            if removed:
                _emit_log(f"Removed {removed} existing {category} file(s) in {date_folder}", important=True)
            
            # Copy the file
            dest_path = date_folder / new_name
            shutil.copy2(file_path, dest_path)
            _emit_log(f"Copied {category} file to {dest_path}", important=True)
            
    except Exception as e:
        _emit_log(f"Failed to copy uploaded files to GA IS outputs: {e}", important=True)


def copy_uploaded_files_to_la_outputs(uploaded_files: dict, months: dict, dates: dict) -> None:
    """Copy uploaded files to LA outputs folder after renaming."""
    try:
        # Get the LA outputs directory
        la_outputs_dir = OUTPUTS_DIR / "NBD_WF_15_LA"
        if not la_outputs_dir.exists():
            _emit_log(f"LA outputs directory not found: {la_outputs_dir}", important=True)
            return
        
        # Find the latest completed date folder
        _, date_folder = find_latest_completed_date(la_outputs_dir)
        if not date_folder or not date_folder.exists():
            _emit_log(f"No completed date folder found for LA in {la_outputs_dir}", important=True)
            return
        
        _emit_log(f"Found LA date folder: {date_folder}", important=True)
        
        # Process each uploaded file
        for category, file_path in uploaded_files.items():
            if not file_path or not Path(file_path).exists():
                continue
                
            # Use the original filename (already renamed when stored)
            new_name = Path(file_path).name
            
            # LA only handles specific files - skip others
            # Note: M2M is no longer a file - amounts are entered via form and saved to Master_Data.xlsx
            if category not in ["Daily Bank Balances", "BankBalanceGLCodes"]:
                _emit_log(f"Skipping {category} - not an LA file", important=True)
                continue

            # BankBalanceGLCodes goes to working folder (date folder itself), others go to Input subfolder
            if category == "BankBalanceGLCodes":
                target_dir = date_folder
            else:
                target_dir = date_folder / "Input"
            target_dir.mkdir(parents=True, exist_ok=True)

            # Remove existing files with same prefix (with force removal for locked files)
            removed = 0
            _emit_log(f"Scanning {target_dir} for existing files to remove for category: {category}", important=True)

            for f in target_dir.iterdir():
                if f.is_file() and f.suffix.lower() in [".xlsx", ".xlsm"]:
                    _emit_log(f"Found file: {f.name}", important=True)
                    should_remove = False

                    if category == "Daily Bank Balances" and f.name.startswith("Daily Bank Balances"):
                        should_remove = True
                    elif category == "BankBalanceGLCodes" and f.name == "BankBalanceGLCodes.xlsx":
                        should_remove = True

                    if should_remove:
                        # Try to remove with retries and force kill Excel if needed
                        max_attempts = 5
                        for attempt in range(1, max_attempts + 1):
                            try:
                                f.unlink()
                                removed += 1
                                _emit_log(f"Successfully removed existing file: {f.name}", important=True)
                                break
                            except PermissionError as e:
                                if attempt < max_attempts:
                                    _emit_log(f"File locked, retrying in 2 seconds... (attempt {attempt}/{max_attempts})", important=True)
                                    time.sleep(2)

                                    # On 3rd attempt, try killing Excel processes
                                    if attempt == 3:
                                        _emit_log(f"Attempting to kill Excel processes to unlock file: {f.name}", important=True)
                                        try:
                                            subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"],
                                                         capture_output=True, timeout=5)
                                            time.sleep(1)
                                        except Exception:
                                            pass
                                else:
                                    # Last attempt - try changing permissions and force delete
                                    try:
                                        import stat
                                        os.chmod(str(f), stat.S_IWRITE)
                                        time.sleep(1)
                                        f.unlink()
                                        removed += 1
                                        _emit_log(f"Force removed file after permission change: {f.name}", important=True)
                                    except Exception as final_error:
                                        _emit_log(f"ERROR: Failed to remove old {category} file {f.name} after {max_attempts} attempts: {final_error}", important=True)
                            except Exception as e:
                                _emit_log(f"ERROR: Failed to remove existing file {f.name}: {e}", important=True)
                                break
            
            if removed:
                _emit_log(f"Removed {removed} existing {category} file(s) in {target_dir}", important=True)
            
            # Copy the file
            dest_path = target_dir / new_name
            shutil.copy2(file_path, dest_path)
            _emit_log(f"Copied {category} file to {dest_path}", important=True)
            
    except Exception as e:
        _emit_log(f"Failed to copy uploaded files to LA outputs: {e}", important=True)


def copy_uploaded_files_to_qf_outputs(uploaded_files: dict, months: dict, years: dict, dates: dict) -> None:
    """Copy uploaded QF files to QF outputs folder with specific subfolder routing"""
    try:
        # Find the latest QF outputs date folder
        qf_outputs_dir = OUTPUTS_DIR / "NBD-QF-23-C3-C10-SF-FI"
        if not qf_outputs_dir.exists():
            _emit_log("QF outputs directory not found", important=True)
            return

        # Find latest completed date
        last_date, date_folder = find_latest_completed_date(qf_outputs_dir)
        if not date_folder or not date_folder.exists():
            _emit_log("No completed date folder found for QF", important=True)
            return

        # Input folder is the base
        input_folder = date_folder / "Input"
        input_folder.mkdir(parents=True, exist_ok=True)

        # Process each uploaded file
        for category, file_path in uploaded_files.items():
            if not file_path or not Path(file_path).exists():
                continue

            # Use the original filename (already renamed when stored)
            new_name = Path(file_path).name

            # Determine target folder and cleanup pattern
            if "Client Available Limits" in category:
                # Client Available Limits go to MT subfolder
                target_folder = input_folder / "MT"
                target_folder.mkdir(parents=True, exist_ok=True)
                dest_path = target_folder / new_name

                # Remove old files: starts with "Client Available Limits Report" and ends with ".xlsx"
                removed = 0
                for f in target_folder.iterdir():
                    if f.is_file() and f.name.startswith("Client Available Limits Report") and f.suffix.lower() == ".xlsx" and f.name != new_name:
                        try:
                            _emit_log(f"Removing OLD file: {f.name} (keeping latest: {new_name})", important=True)
                            f.unlink()
                            removed += 1
                        except Exception as e:
                            _emit_log(f"Failed to remove OLD file {f.name}: {e}", important=True)
                if removed:
                    _emit_log(f"Removed {removed} OLD Client Available Limits file(s) from MT folder", important=True)

            else:
                # All other files go to Input folder
                target_folder = input_folder
                dest_path = target_folder / new_name

                # Determine cleanup pattern based on category
                removed = 0
                if category == "Sector Report":
                    # Remove files: starts with "Sector Report" and ends with ".xlsb"
                    for f in target_folder.iterdir():
                        if f.is_file() and f.name.startswith("Sector Report") and f.suffix.lower() == ".xlsb" and f.name != new_name:
                            try:
                                _emit_log(f"Removing OLD Sector Report: {f.name}", important=True)
                                f.unlink()
                                removed += 1
                            except Exception as e:
                                _emit_log(f"Failed to remove OLD file {f.name}: {e}", important=True)

                elif category == "JOBHD":
                    # Remove files: contains "JOBHD" and ends with ".xlsx"
                    for f in target_folder.iterdir():
                        if f.is_file() and "JOBHD" in f.name and f.suffix.lower() == ".xlsx" and f.name != new_name:
                            try:
                                _emit_log(f"Removing OLD JOBHD: {f.name}", important=True)
                                f.unlink()
                                removed += 1
                            except Exception as e:
                                _emit_log(f"Failed to remove OLD file {f.name}: {e}", important=True)

                elif category == "Business Gender":
                    # Remove files: starts with "Business Gender" and ends with ".xlsx"
                    for f in target_folder.iterdir():
                        if f.is_file() and f.name.startswith("Business Gender") and f.suffix.lower() == ".xlsx" and f.name != new_name:
                            try:
                                _emit_log(f"Removing OLD Business Gender: {f.name}", important=True)
                                f.unlink()
                                removed += 1
                            except Exception as e:
                                _emit_log(f"Failed to remove OLD file {f.name}: {e}", important=True)

                elif category == "Information Request from Credit":
                    # Remove files: starts with "Information Request from Credit" and ends with ".xlsx"
                    for f in target_folder.iterdir():
                        if f.is_file() and f.name.startswith("Information Request from Credit") and f.suffix.lower() == ".xlsx" and f.name != new_name:
                            try:
                                _emit_log(f"Removing OLD Info Request from Credit: {f.name}", important=True)
                                f.unlink()
                                removed += 1
                            except Exception as e:
                                _emit_log(f"Failed to remove OLD file {f.name}: {e}", important=True)

                elif category == "Information Request from IT":
                    # Remove files: starts with "Information Request from IT" and ends with ".xlsx"
                    for f in target_folder.iterdir():
                        if f.is_file() and f.name.startswith("Information Request from IT") and f.suffix.lower() == ".xlsx" and f.name != new_name:
                            try:
                                _emit_log(f"Removing OLD Info Request from IT: {f.name}", important=True)
                                f.unlink()
                                removed += 1
                            except Exception as e:
                                _emit_log(f"Failed to remove OLD file {f.name}: {e}", important=True)

                elif category == "Net Portfolio":
                    # Remove files: starts with "Net Portfolio" and ends with ".xlsx"
                    for f in target_folder.iterdir():
                        if f.is_file() and f.name.startswith("Net Portfolio") and f.suffix.lower() == ".xlsx" and f.name != new_name:
                            try:
                                _emit_log(f"Removing OLD Net Portfolio: {f.name}", important=True)
                                f.unlink()
                                removed += 1
                            except Exception as e:
                                _emit_log(f"Failed to remove OLD file {f.name}: {e}", important=True)

                if removed:
                    _emit_log(f"Removed {removed} OLD {category} file(s)", important=True)

            # Copy the new file to target folder
            shutil.copy2(file_path, dest_path)
            _emit_log(f"Copied {category} file to {dest_path}", important=True)

            # Remove "Copy of " prefix if it exists in the copied file
            if dest_path.exists() and dest_path.name.startswith("Copy of "):
                new_clean_name = dest_path.name.replace("Copy of ", "", 1)
                clean_path = dest_path.parent / new_clean_name
                dest_path.rename(clean_path)
                _emit_log(f"Removed 'Copy of ' prefix from file: {new_clean_name}", important=True)

            # SHARED FILES: Also copy to other report folders that use these files
            if category in ["Information Request from Credit", "Net Portfolio"]:
                _emit_log(f"'{category}' is shared - copying to IA outputs folder as well", important=True)

                # Copy to IA outputs folder
                ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
                if ia_outputs_dir.exists():
                    last_date_ia, ia_date_folder = find_latest_completed_date(ia_outputs_dir)
                    if ia_date_folder and ia_date_folder.exists():
                        ia_dest_path = ia_date_folder / new_name

                        # Remove old files in IA folder with same prefix
                        removed_ia = 0
                        for f in ia_date_folder.iterdir():
                            if f.is_file() and f.name.startswith(category) and f.suffix.lower() == ".xlsx" and f.name != new_name:
                                try:
                                    _emit_log(f"Removing OLD {category} from IA: {f.name}", important=True)
                                    f.unlink()
                                    removed_ia += 1
                                except Exception as e:
                                    _emit_log(f"Failed to remove OLD IA file {f.name}: {e}", important=True)

                        # Copy to IA outputs
                        shutil.copy2(file_path, ia_dest_path)
                        _emit_log(f"Copied shared {category} file to IA outputs: {ia_dest_path}", important=True)

                        # Remove "Copy of " prefix if it exists in the copied file
                        if ia_dest_path.exists() and ia_dest_path.name.startswith("Copy of "):
                            new_clean_name = ia_dest_path.name.replace("Copy of ", "", 1)
                            clean_path = ia_dest_path.parent / new_clean_name
                            ia_dest_path.rename(clean_path)
                            _emit_log(f"Removed 'Copy of ' prefix from file: {new_clean_name}", important=True)

                        if removed_ia:
                            _emit_log(f"Removed {removed_ia} OLD {category} file(s) from IA folder", important=True)
                    else:
                        _emit_log(f"No completed date folder found for IA - skipping shared file copy", important=False)
                else:
                    _emit_log(f"IA outputs directory not found - skipping shared file copy", important=False)

    except Exception as e:
        _emit_log(f"Failed to copy uploaded files to QF outputs: {e}", important=True)


def copy_uploaded_files_to_sofp_outputs(uploaded_files: dict, months: dict) -> None:
    """Copy uploaded files to SOFP outputs folder after renaming."""
    try:
        # Get the SOFP outputs directory
        sofp_outputs_dir = OUTPUTS_DIR / "NBD_MF_01_SOFP_SOCI"
        if not sofp_outputs_dir.exists():
            _emit_log(f"SOFP outputs directory not found: {sofp_outputs_dir}", important=True)
            return
        
        # Find the latest completed date folder
        _, date_folder = find_latest_completed_date(sofp_outputs_dir)
        if not date_folder or not date_folder.exists():
            _emit_log(f"No completed date folder found for SOFP in {sofp_outputs_dir}", important=True)
            return
        
        _emit_log(f"Found SOFP date folder: {date_folder}", important=True)
        
        # Process each uploaded file
        for category, file_path in uploaded_files.items():
            if not file_path or not Path(file_path).exists():
                continue
                
            # Use the original filename (already renamed when stored)
            new_name = Path(file_path).name
            
            # Remove existing files with same prefix
            removed = 0
            _emit_log(f"Scanning {date_folder} for existing files to remove for category: {category}", important=True)
            
            for f in date_folder.iterdir():
                if f.is_file() and f.suffix.lower() in [".xlsx", ".xlsm"]:
                    _emit_log(f"Found file: {f.name}", important=True)
                    should_remove = False
                    
                    if category == "ALCL Management Accounts" and f.name.startswith("ALCL Management Accounts"):
                        should_remove = True
                    elif category == "Investment Schedule" and f.name.startswith("Investment Schedule"):
                        should_remove = True
                    elif category == "Loan Schedule" and f.name.startswith("Loan Schedule"):
                        should_remove = True
                    elif category == "Supporting Schedules" and f.name.startswith("Supporting Schedules"):
                        should_remove = True
                    
                    if should_remove:
                        try:
                            f.unlink()
                            removed += 1
                            _emit_log(f"Successfully removed existing file: {f.name}", important=True)
                        except Exception as e:
                            _emit_log(f"ERROR: Failed to remove existing file {f.name}: {e}", important=True)
            
            if removed:
                _emit_log(f"Removed {removed} existing {category} file(s) in {date_folder}", important=True)
            
            # Copy the file
            dest_path = date_folder / new_name
            shutil.copy2(file_path, dest_path)
            _emit_log(f"Copied {category} file to {dest_path}", important=True)
                    
    except Exception as e:
        _emit_log(f"Failed to copy uploaded files to SOFP outputs: {e}", important=True)


def copy_working_to_outputs_with_versioning(report_ids_with_dates: dict) -> None:
    """Copy working folders to outputs with proper versioning after completion."""
    for report_id in report_ids_with_dates:
        if report_id not in ("NBD-MF-01-SOFP-SOCI", "NBD-AF-01-SOFP-SOCI", "NBD-MF-23-IA"):
            continue
            
        date_str, _, _ = report_ids_with_dates[report_id]
        folder_name = REPORT_STRUCTURE[report_id]["folder"]
        
        # Get working folder
        working_report_dir = WORKING_DIR / folder_name
        if not working_report_dir.exists():
            _emit_log(f"Working folder not found: {working_report_dir}", important=True)
            continue
            
        # Find the date folder in working
        date_folders = [p for p in working_report_dir.iterdir() if p.is_dir()]
        if len(date_folders) != 1:
            _emit_log(f"Expected one date folder in {working_report_dir}, found {len(date_folders)}", important=True)
            continue
            
        source_folder = date_folders[0]
        target_date_name = format_date_folder_name(date_str)
        
        # Create output report directory
        output_report_dir = OUTPUTS_DIR / folder_name
        output_report_dir.mkdir(parents=True, exist_ok=True)
        
        # Check if target date folder exists and find next version
        target_path = output_report_dir / target_date_name
        if target_path.exists():
            version = 1
            while (output_report_dir / f"{target_date_name}({version})").exists():
                version += 1
            target_path = output_report_dir / f"{target_date_name}({version})"
        
        try:
            shutil.copytree(source_folder, target_path)
            _emit_log(f"Copied {source_folder} to {target_path}", important=True)
        except Exception as e:
            _emit_log(f"Failed to copy {source_folder} to {target_path}: {e}", important=True)


def auto_download_files() -> None:
    """Automatically prepare and download files after successful completion."""
    global _download_files

    if not _download_files:
        _emit_log("No files in _download_files to prepare for download", important=True)
        return
    
    # Filter out non-existent files
    existing_files = [f for f in _download_files if f.exists() and f.is_file()]
    
    if not existing_files:
        _emit_log("No existing files found in _download_files", important=True)
        return
    
    # If only one file, keep it as-is for direct download
    if len(existing_files) == 1:
        _download_files = existing_files
        _emit_log(f"Single file ready for download: {existing_files[0].name}", important=True)
        return
    
    # Multiple files - create zip
    temp_zip_dir = BASE_DIR / "temp_zips"
    temp_zip_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = temp_zip_dir / f"reports_{timestamp}.zip"

    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            files_added = 0
            
            for file_path in existing_files:
                arcname = f"output_files/{file_path.name}"
                zipf.write(file_path, arcname=arcname)
                files_added += 1
                _emit_log(f"Added to zip: {arcname}", important=False)

            _emit_log(f"Created zip with {files_added} files", important=True)

        if files_added > 0:
            # Store the zip path for potential download
            _download_files = [zip_path]
            _emit_log(f"Files prepared for download: {zip_path}", important=True)
        else:
            _emit_log("No files found to download", important=True)

    except Exception as e:
        _emit_log(f"Failed to prepare files for download: {e}", important=True)

    # DO NOT clean up working directories here - files are still needed for zip
    # Cleanup will happen later after download is complete or on next run
    _emit_log("Zip file created successfully. Working directories will be cleaned up on next run.", important=True)


def copy_folder_from_outputs_to_working(source_report_id: str, dest_report_id: str) -> tuple[bool, str, Path | None]:
    """
    Copy the latest completed date folder from outputs/source to working/dest.
    For NBD-MF-23-IA, source is NBD_MF_23_C4.
    """
    if source_report_id not in REPORT_STRUCTURE or dest_report_id not in REPORT_STRUCTURE:
        return False, f"Unknown report: {source_report_id} or {dest_report_id}", None
    
    source_folder_name = REPORT_STRUCTURE[source_report_id]["folder"]
    dest_folder_name = REPORT_STRUCTURE[dest_report_id]["folder"]
    
    source_dir = OUTPUTS_DIR / source_folder_name
    dest_report_dir = WORKING_DIR / dest_folder_name
    
    if not source_dir.exists():
        return False, f"Source folder not found: {source_dir}", None
    
    # Find latest completed date
    last_date, date_folder = find_latest_completed_date(source_dir)
    
    if not date_folder or not date_folder.exists():
        return False, f"No completed date folder found in: {source_dir}", None
    
    # STEP 1: Clean up duplicate files in the source folder BEFORE copying
    _emit_log(f"*** CLEANING DUPLICATES IN SOURCE FOLDER *** {date_folder}", important=True)
    removed_count = cleanup_duplicate_files_in_folder(date_folder)
    if removed_count > 0:
        _emit_log(f"Removed {removed_count} duplicate files from source folder", important=True)
    
    # Create working report directory
    dest_report_dir.mkdir(parents=True, exist_ok=True)
    
    # Remove existing date folders in working
    for child in dest_report_dir.iterdir():
        if child.is_dir():
            shutil.rmtree(child, ignore_errors=True)
    
    # Copy latest folder to working
    dest_path = dest_report_dir / date_folder.name
    try:
        shutil.copytree(date_folder, dest_path)
        _emit_log(f"Copied {date_folder} to {dest_path}", important=True)
        return True, f"Copied {date_folder.name}", dest_path
    except Exception as e:
        return False, f"Failed to copy: {e}", None


def cleanup_duplicate_files_in_folder(folder_path: Path) -> int:
    """
    Clean up duplicate files in a folder by keeping only the latest file.
    Returns the number of files removed.
    """
    if not folder_path.exists():
        return 0
    
    _emit_log(f"*** CLEANING UP DUPLICATE FILES *** in {folder_path}", important=True)
    
    # Group files by base name (without extension)
    file_groups = {}
    removed_count = 0
    
    for file_path in folder_path.iterdir():
        if file_path.is_file():
            # Get base name without extension
            base_name = file_path.stem
            extension = file_path.suffix
            
            if base_name not in file_groups:
                file_groups[base_name] = []
            file_groups[base_name].append(file_path)
    
    # Process each group of files with the same base name
    for base_name, files in file_groups.items():
        if len(files) > 1:
            _emit_log(f"Found {len(files)} duplicate files for '{base_name}':", important=True)
            
            # Sort by modification time (newest first)
            files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
            
            # Keep the newest file, remove the rest
            latest_file = files[0]
            old_files = files[1:]
            
            from datetime import datetime
            _emit_log(f"  KEEPING (latest): {latest_file.name} (modified: {datetime.fromtimestamp(latest_file.stat().st_mtime)})", important=True)
            
            for old_file in old_files:
                try:
                    _emit_log(f"  REMOVING (older): {old_file.name} (modified: {datetime.fromtimestamp(old_file.stat().st_mtime)})", important=True)
                    
                    # Check if file is locked first
                    if is_file_locked(old_file):
                        _emit_log(f"  FILE IS LOCKED - RENAMING: {old_file.name}", important=True)
                        if rename_locked_file(old_file):
                            removed_count += 1
                            _emit_log(f"  SUCCESSFULLY RENAMED LOCKED FILE: {old_file.name}", important=True)
                        else:
                            _emit_log(f"  FAILED TO RENAME LOCKED FILE: {old_file.name}", important=True)
                    else:
                        # File is not locked, remove it normally
                        try:
                            old_file.unlink()
                            removed_count += 1
                            _emit_log(f"  SUCCESSFULLY REMOVED: {old_file.name}", important=True)
                        except Exception as e:
                            _emit_log(f"  FAILED TO REMOVE: {old_file.name} - {e}", important=True)
                except Exception as e:
                    _emit_log(f"  FAILED TO REMOVE: {old_file.name} - {e}", important=True)
        else:
            _emit_log(f"No duplicates for '{base_name}' - keeping: {files[0].name}", important=True)
    
    _emit_log(f"*** DUPLICATE CLEANUP COMPLETE *** Removed {removed_count} duplicate files", important=True)
    return removed_count


def is_file_locked(filepath: Path) -> bool:
    """Check if a file is locked by trying to open it exclusively"""
    if not filepath.exists():
        return False
    
    try:
        # Try to open the file in exclusive mode
        with open(filepath, 'r+b') as f:
            # Try to acquire an exclusive lock (Windows specific)
            import msvcrt
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
            msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
        _emit_log(f"File is NOT locked: {filepath.name}")
        return False
    except (IOError, OSError, PermissionError) as e:
        _emit_log(f"File IS locked: {filepath.name} - {e}")
        return True
    except ImportError:
        # If not on Windows, try a simpler approach
        try:
            with open(filepath, 'r+b'):
                pass
            _emit_log(f"File is NOT locked (non-Windows): {filepath.name}")
            return False
        except (IOError, OSError, PermissionError) as e:
            _emit_log(f"File IS locked (non-Windows): {filepath.name} - {e}")
            return True


def force_remove_file(filepath: Path, max_retries: int = 3, retry_delay: float = 1.0) -> bool:
    """
    Attempt to forcefully remove a file with retries.
    Returns True if successful, False otherwise.
    """
    for attempt in range(max_retries):
        try:
            if filepath.exists():
                # Try normal deletion first
                filepath.unlink()
                _emit_log(f"Successfully removed file: {filepath.name}")
                return True
        except PermissionError:
            if attempt < max_retries - 1:
                _emit_log(f"File locked, attempt {attempt + 1}/{max_retries}. Waiting {retry_delay}s...")
                time.sleep(retry_delay)
                
                # Try to close any file handles (Windows specific)
                try:
                    import subprocess
                    # Use handle.exe if available to close file handles
                    result = subprocess.run(
                        ['handle.exe', str(filepath), '-c', '-nobanner'],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                except:
                    pass
            else:
                _emit_log(f"Failed to remove locked file after {max_retries} attempts: {filepath.name}")
                return False
    
    return not filepath.exists()


def rename_locked_file(filepath: Path) -> bool:
    """
    If a file can't be deleted, try renaming it with a timestamp suffix.
    This allows the new file to be copied without creating a duplicate name.
    """
    try:
        if not filepath.exists():
            _emit_log(f"File does not exist, nothing to rename: {filepath.name}")
            return True
            
        # Generate a unique backup name with timestamp
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = filepath.stem
        suffix = filepath.suffix
        backup_name = f"{stem}_OLD_{timestamp}{suffix}"
        backup_path = filepath.parent / backup_name
        
        _emit_log(f"Attempting to rename locked file: {filepath.name} -> {backup_name}")
        
        # Try to rename the file
        filepath.rename(backup_path)
        _emit_log(f"SUCCESSFULLY renamed locked file to: {backup_name}", important=True)
        return True
    except Exception as e:
        _emit_log(f"FAILED to rename locked file {filepath.name}: {e}", important=True)
        return False


def cleanup_report_working_folder(report_id: str) -> None:
    """
    Clean up the working folder for a specific report.
    Removes all date folders from the report's working directory.

    Args:
        report_id: The report ID (e.g., "NBD-WF-15-LA")
    """
    try:
        if report_id not in REPORT_STRUCTURE:
            return

        folder_name = REPORT_STRUCTURE[report_id]["folder"]
        report_working_dir = WORKING_DIR / folder_name

        if not report_working_dir.exists():
            return

        # Remove all date folders
        date_folders = [d for d in report_working_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]

        if date_folders:
            _emit_log(f"Cleaning {folder_name} working folder: Found {len(date_folders)} date folders", important=True)

            for date_folder in date_folders:
                try:
                    shutil.rmtree(date_folder)
                    _emit_log(f"  Removed: {date_folder.name}")
                except Exception as e:
                    _emit_log(f"  Failed to remove {date_folder.name}: {e}", important=True)

    except Exception as e:
        _emit_log(f"Error cleaning working folder for {report_id}: {e}", important=True)


def cleanup_all_working_directories() -> None:
    """
    Clean up all date folders from working directories to prevent conflicts.
    This should be called:
    1. Before starting new automation
    2. When user stops a process
    3. After downloads complete
    """
    try:
        _emit_log("*** CLEANING UP ALL WORKING DIRECTORIES ***", important=True)

        if not WORKING_DIR.exists():
            _emit_log("Working directory does not exist, nothing to clean")
            return

        total_removed = 0

        # Iterate through all subdirectories in working/
        for report_dir in WORKING_DIR.iterdir():
            if not report_dir.is_dir():
                continue

            # For each report folder, remove all date folders
            date_folders = [d for d in report_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]

            if date_folders:
                _emit_log(f"Cleaning {report_dir.name}: Found {len(date_folders)} date folders", important=True)

                for date_folder in date_folders:
                    try:
                        shutil.rmtree(date_folder)
                        _emit_log(f"  Removed: {date_folder.name}")
                        total_removed += 1
                    except Exception as e:
                        _emit_log(f"  Failed to remove {date_folder.name}: {e}", important=True)

        _emit_log(f"*** WORKING DIRECTORY CLEANUP COMPLETE *** Removed {total_removed} date folders", important=True)

    except Exception as e:
        _emit_log(f"Error cleaning working directories: {e}", important=True)


def get_files_with_prefix(directory: Path, prefix: str, extensions: Tuple[str, ...] = ('.xlsx', '.xlsm', '.xlsb')) -> List[Path]:
    """Get all files in directory that start with the given prefix and have valid extensions"""
    matching_files = []
    if directory.exists():
        for file_path in directory.iterdir():
            if file_path.is_file():
                file_name = file_path.name
                if file_name.startswith(prefix) and file_name.lower().endswith(extensions):
                    matching_files.append(file_path)
    return matching_files


def copy_master_data_file_with_cleanup_improved(
    file_category: str, 
    source_file_path: Path, 
    target_dir: Path, 
    target_filename: str,
    force_overwrite: bool = True,
    handle_locked_files: str = "rename"  # Options: "rename", "skip", "retry"
) -> Tuple[bool, str]:
    """
    Improved version of copy_master_data_file_with_cleanup that handles locked files better.
    
    Parameters:
    -----------
    file_category : str
        The category of master data file being copied
    source_file_path : Path
        Source file to copy
    target_dir : Path
        Target directory to copy to
    target_filename : str
        Name for the target file
    force_overwrite : bool
        If True, will attempt to overwrite locked files
    handle_locked_files : str
        Strategy for handling locked files:
        - "rename": Rename locked files with timestamp suffix
        - "skip": Skip copying if any old file is locked
        - "retry": Retry deletion with delays
    
    Returns:
    --------
    Tuple[bool, str]: (Success status, Message describing the result)
    """
    try:
        _emit_log(f"COPYING MASTER DATA FILE WITH IMPROVED CLEANUP: {file_category}", important=True)
        _emit_log(f"Source: {source_file_path}")
        _emit_log(f"Target: {target_dir / target_filename}")
        _emit_log(f"Strategy for locked files: {handle_locked_files}")
        
        # Validate source file
        if not source_file_path.exists():
            error_msg = f"Source file does not exist: {source_file_path}"
            _emit_log(error_msg, important=True)
            return False, error_msg
        
        # Create target directory if it doesn't exist
        target_dir.mkdir(parents=True, exist_ok=True)
        
        # Define prefix mapping (same as original)
        prefix_mapping = {
            "Disbursement with Budget": "Disbursement with Budget",
            "Information Request from Credit": "Information Request from Credit",
            "Net Portfolio": "Net Portfolio",
            "YARD STOCK AS AT": "YARD STOCK AS AT",
            "ALCL Management Accounts": "ALCL Management Accounts",
            "Investment Schedule": "Investment Schedule",
            "Loan Schedule": "Loan Schedule",
            "Supporting Schedules": "Supporting Schedules",
            "Borrowing report": "Borrowing report",
            "Cadre": "Cadre",
            "Unutilized Amount": "Unutilized Amount",
            "Daily Bank Balances": "Daily Bank Balances",
            "M2M": "M2M",
            "BankBalanceGLCodes": "BankBalanceGLCodes",
            "Rec Target": "Rec Target"
        }
        
        prefix = prefix_mapping.get(file_category, file_category)
        
        # STEP 1: Find and handle old files with the same prefix
        _emit_log(f"STEP 1: CLEANING UP OLD FILES", important=True)
        old_files = get_files_with_prefix(target_dir, prefix)
        _emit_log(f"Found {len(old_files)} old file(s) with prefix '{prefix}' in {target_dir}")
        
        for old_file in old_files:
            _emit_log(f"  - {old_file.name}")
        
        removed_count = 0
        renamed_count = 0
        locked_files = []
        
        for old_file in old_files:
            _emit_log(f"Processing old file: {old_file.name}")
            
            # Check if file is locked
            if is_file_locked(old_file):
                locked_files.append(old_file)
                _emit_log(f"File is locked: {old_file.name}")
                
                # Handle locked file based on strategy
                if handle_locked_files == "rename":
                    if rename_locked_file(old_file):
                        renamed_count += 1
                    else:
                        # If rename fails and we're trying to copy a file with the same name, this is a problem
                        if old_file.name == target_filename:
                            error_msg = f"Cannot overwrite locked file with same name: {old_file.name}"
                            _emit_log(error_msg, important=True)
                            if not force_overwrite:
                                return False, error_msg
                                
                elif handle_locked_files == "retry":
                    if force_remove_file(old_file, max_retries=3, retry_delay=1.0):
                        removed_count += 1
                    elif old_file.name == target_filename:
                        error_msg = f"Cannot remove locked file after retries: {old_file.name}"
                        _emit_log(error_msg, important=True)
                        if not force_overwrite:
                            return False, error_msg
                            
                elif handle_locked_files == "skip":
                    error_msg = f"Skipping copy due to locked file: {old_file.name}"
                    _emit_log(error_msg, important=True)
                    return False, error_msg
            else:
                # File is not locked, remove it normally
                try:
                    old_file.unlink()
                    removed_count += 1
                    _emit_log(f"Removed old file: {old_file.name}")
                except Exception as e:
                    _emit_log(f"Failed to remove file {old_file.name}: {e}")
        
        _emit_log(f"Cleanup complete: Removed {removed_count} file(s), Renamed {renamed_count} file(s)")
        
        # STEP 2: Check if target file already exists and is locked
        target_path = target_dir / target_filename
        if target_path.exists() and is_file_locked(target_path):
            _emit_log(f"Target file already exists and is locked: {target_filename}", important=True)
            
            if handle_locked_files == "rename":
                if not rename_locked_file(target_path):
                    if not force_overwrite:
                        return False, f"Cannot rename existing locked target file: {target_filename}"
            elif handle_locked_files == "retry":
                if not force_remove_file(target_path):
                    if not force_overwrite:
                        return False, f"Cannot remove existing locked target file: {target_filename}"
            else:
                return False, f"Target file is locked: {target_filename}"
        
        # STEP 3: Copy the new file
        _emit_log(f"STEP 2: COPYING NEW FILE", important=True)
        
        try:
            # Use shutil.copy2 to preserve metadata
            shutil.copy2(source_file_path, target_path)
            _emit_log(f"Successfully copied {file_category} to {target_path}")
            
            # Verify the copy
            if target_path.exists():
                source_size = source_file_path.stat().st_size
                target_size = target_path.stat().st_size
                
                if source_size == target_size:
                    success_msg = f"Successfully copied and verified {file_category} ({source_size} bytes)"
                    _emit_log(success_msg, important=True)
                    return True, success_msg
                else:
                    warning_msg = f"File copied but sizes don't match. Source: {source_size}, Target: {target_size}"
                    _emit_log(warning_msg, important=True)
                    return True, warning_msg
            else:
                error_msg = f"File not found after copy operation"
                _emit_log(error_msg, important=True)
                return False, error_msg
                
        except Exception as e:
            error_msg = f"Failed to copy file: {e}"
            _emit_log(error_msg, important=True)
            return False, error_msg
            
    except Exception as e:
        error_msg = f"Unexpected error in copy_master_data_file_with_cleanup_improved: {e}"
        _emit_log(error_msg, important=True)
        return False, error_msg


def copy_folder_to_working(report_id: str) -> tuple[bool, str, Path | None]:
    """Copy latest completed date folder from outputs/<folder> to working/<folder>.

    For this app, the authoritative latest date comes from completed_dates.txt in the
    corresponding outputs/<folder>. We find the last line, locate the matching date
    folder (including version suffix), clear existing subfolders in working/<folder>,
    then copy that folder in full.
    """
    if report_id not in REPORT_STRUCTURE:
        return False, f"Unknown report: {report_id}", None

    folder_name = REPORT_STRUCTURE[report_id]["folder"]
    # Special-case GA & IS SET8: outputs folder uses ampersand name; working uses underscore name
    if report_id == "NBD-MF-10-GA-11-IS":
        source_report_dir = OUTPUTS_DIR / "NBD-MF-10-GA & NBD-MF-11-IS"
        dest_report_dir = WORKING_DIR / "NBD_MF_10_GA_NBD_MF_11_IS"
    elif report_id in ("NBD-QF-23-C3", "NBD-QF-23-C4", "NBD-QF-23-C6",
                       "NBD-QF-23-C9", "NBD-QF-23-C10",
                       "NBD-QF-16-SF", "NBD-QF-16-FI"):
        # QF Level 3 reports all use shared folder
        source_report_dir = OUTPUTS_DIR / "NBD-QF-23-C3-C10-SF-FI"
        dest_report_dir = WORKING_DIR / "NBD-QF-23-C3-C10-SF-FI"
    else:
        source_report_dir = OUTPUTS_DIR / folder_name
        dest_report_dir = WORKING_DIR / folder_name
    
    if not source_report_dir.exists():
        return False, f"Report folder not found: {source_report_dir}", None
    
    # Find latest completed date from completed_dates.txt
    last_date, date_folder = find_latest_completed_date(source_report_dir)
    if last_date is None:
        return False, f"No completed_dates.txt or last date in: {source_report_dir}", None
    if not date_folder or not date_folder.exists():
        return False, f"Date folder not found for: {last_date.strftime('%d-%m-%Y')} under {source_report_dir}", None
    
    # STEP 1: Clean up duplicate files in the source folder BEFORE copying
    _emit_log(f"*** CLEANING DUPLICATES IN SOURCE FOLDER *** {date_folder}", important=True)
    removed_count = cleanup_duplicate_files_in_folder(date_folder)
    if removed_count > 0:
        _emit_log(f"Removed {removed_count} duplicate files from source folder", important=True)
    
    dest_report_dir.mkdir(parents=True, exist_ok=True)

    # Preserve CBSL Provision Comparison files ONLY for IA (not for GA/IS or other reports)
    # These files are created by GA/IS and should persist in IA folder across C1-C2 runs
    # But GA/IS process should manage its own CBSL files independently
    cbsl_files_to_preserve = []
    if report_id == "NBD-MF-23-IA":
        for child in list(dest_report_dir.iterdir()):
            if child.is_dir():
                try:
                    for f in child.rglob("CBSL Provision Comparison*.xlsb"):
                        # Save file to temp location
                        temp_path = dest_report_dir / f.name
                        shutil.copy2(f, temp_path)
                        cbsl_files_to_preserve.append(temp_path)
                        _emit_log(f"Preserving CBSL file for IA: {f.name}", important=True)
                except Exception as e:
                    _emit_log(f"Warning: Failed to preserve CBSL files: {e}", important=True)

    # Remove existing subfolders before copying
    for child in list(dest_report_dir.iterdir()):
        if child.is_dir():
            try:
                _emit_log(f"Removing existing folder: {child}", important=True)
                shutil.rmtree(child)
                _emit_log(f"Successfully removed: {child.name}", important=True)
            except PermissionError as e:
                _emit_log(f"Permission error removing {child}: {e}. Trying to handle locked files...", important=True)
                # Try to handle locked files by using onerror callback
                def handle_remove_error(func, path, exc_info):
                    """Error handler for shutil.rmtree - tries to unlock and remove files"""
                    import stat
                    import time
                    try:
                        # Try to change permissions and remove
                        os.chmod(path, stat.S_IWRITE)
                        time.sleep(0.1)  # Brief pause
                        func(path)
                    except Exception:
                        # If still fails, skip this file
                        _emit_log(f"Cannot remove locked file: {path}", important=True)

                try:
                    shutil.rmtree(child, onerror=handle_remove_error)
                    _emit_log(f"Successfully removed {child.name} with error handling", important=True)
                except Exception as e2:
                    _emit_log(f"Still cannot remove {child.name}: {e2}. Attempting force delete...", important=True)
                    # Last resort - try Windows robocopy to mirror an empty directory
                    import subprocess
                    import tempfile
                    try:
                        # Create a temporary empty directory
                        with tempfile.TemporaryDirectory() as empty_dir:
                            # Use robocopy to mirror empty dir over the target (effectively deleting it)
                            subprocess.run(
                                ["robocopy", empty_dir, str(child), "/MIR", "/R:0", "/W:0"],
                                capture_output=True,
                                timeout=30
                            )
                            # robocopy returns 0-7 for success, > 7 for failure
                            # Now try to remove the directory
                            time.sleep(0.5)
                            if child.exists():
                                shutil.rmtree(child, ignore_errors=True)

                        if not child.exists():
                            _emit_log(f"Force deleted {child.name} using robocopy", important=True)
                        else:
                            _emit_log(f"WARNING: Could not fully remove {child.name}, but continuing...", important=True)
                    except Exception as e3:
                        _emit_log(f"Force delete also failed: {e3}. Continuing anyway...", important=True)
            except Exception as e:
                _emit_log(f"Unexpected error removing {child}: {e}", important=True)
                return False, f"Cannot remove existing folder {child.name}: {e}", None

    dest_path = dest_report_dir / date_folder.name

    # Final check - if destination exists, try one more time to remove it
    if dest_path.exists():
        _emit_log(f"Destination still exists after cleanup, attempting final removal: {dest_path}", important=True)
        try:
            import time
            time.sleep(1)  # Wait for any file handles to close
            shutil.rmtree(dest_path, ignore_errors=True)
            if dest_path.exists():
                _emit_log(f"WARNING: Could not remove destination, will try dirs_exist_ok=True", important=True)
        except Exception as e:
            _emit_log(f"Final removal attempt failed: {e}", important=True)

    try:
        # Try to copy with dirs_exist_ok=True to handle any remaining files
        if dest_path.exists():
            shutil.copytree(date_folder, dest_path, dirs_exist_ok=True)
            _emit_log(f"Copied {date_folder} to {dest_path} (with dirs_exist_ok=True)", important=True)
        else:
            shutil.copytree(date_folder, dest_path)
            _emit_log(f"Copied {date_folder} to {dest_path}", important=True)

        # Restore preserved CBSL files to the new date folder
        for cbsl_file in cbsl_files_to_preserve:
            try:
                restored_path = dest_path / cbsl_file.name
                shutil.move(str(cbsl_file), str(restored_path))
                _emit_log(f"Restored CBSL file: {cbsl_file.name}", important=True)
            except Exception as e:
                _emit_log(f"Warning: Failed to restore CBSL file {cbsl_file.name}: {e}", important=True)
                # Clean up temp file if move failed
                try:
                    cbsl_file.unlink(missing_ok=True)
                except Exception:
                    pass

        return True, f"Copied {date_folder.name}", dest_path
    except Exception as e:
        # Clean up preserved files if copy failed
        for cbsl_file in cbsl_files_to_preserve:
            try:
                cbsl_file.unlink(missing_ok=True)
            except Exception:
                pass
        return False, f"Failed to copy: {e}", None


def run_report_script(report_id: str, date_str: str, report_month: str, report_year: str) -> tuple[bool, str]:
    """Run a single report script."""
    if report_id not in REPORT_STRUCTURE:
        return False, f"Unknown report: {report_id}"

    script_name = REPORT_STRUCTURE[report_id]["script"]
    folder_name = REPORT_STRUCTURE[report_id]["folder"]
    script_path = REPORT_AUTOMATIONS_DIR / script_name

    if not script_path.exists():
        msg = f"Script not found: {script_path}"
        _emit_log(msg, important=True)
        return False, msg

    # Determine working directory. For C1/C2, reuse IA's date folder.
    if report_id in ("NBD-MF-23-C1", "NBD-MF-23-C2", "NBD-MF-23-C1-C2"):
        ia_folder_name = REPORT_STRUCTURE["NBD-MF-23-IA"]["folder"]
        ia_working_dir = WORKING_DIR / ia_folder_name

        # For C1-C2 only mode: Use IA outputs directly if working folder doesn't exist
        if not ia_working_dir.exists():
            _emit_log(f"Working directory not found for {report_id}, checking IA outputs", important=True)
            ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
            if not ia_outputs_dir.exists():
                return False, f"Neither working nor outputs directory found for IA"
            # Find latest completed date folder in outputs
            from pathlib import Path
            date_folders = [p for p in ia_outputs_dir.iterdir() if p.is_dir() and not p.name.startswith('.')]
            if not date_folders:
                return False, f"No date folders found in IA outputs"
            # Get the most recent folder
            latest_folder = max(date_folders, key=lambda x: x.stat().st_mtime)
            working_dir = latest_folder
            _emit_log(f"Using IA outputs folder: {working_dir}", important=True)
        else:
            ia_date_folders = [p for p in ia_working_dir.iterdir() if p.is_dir()]
            if len(ia_date_folders) != 1:
                return False, f"Expected exactly one date folder in {ia_working_dir}, found {len(ia_date_folders)}"
            working_dir = ia_date_folders[0]
    elif report_id == "NBD-MF-10-GA-11-IS":
        # Special case for GA & IS SET8: working folder uses underscore name
        working_report_dir = WORKING_DIR / "NBD_MF_10_GA_NBD_MF_11_IS"
        if not working_report_dir.exists():
            return False, f"Working directory not found: {working_report_dir}"
        date_folders = [p for p in working_report_dir.iterdir() if p.is_dir()]
        if len(date_folders) != 1:
            return False, f"Expected exactly one date folder in {working_report_dir}, found {len(date_folders)}"
        working_dir = date_folders[0]
    elif report_id in ("NBD-QF-23-C3", "NBD-QF-23-C4", "NBD-QF-23-C6",
                       "NBD-QF-23-C9", "NBD-QF-23-C10",
                       "NBD-QF-16-SF", "NBD-QF-16-FI"):
        # QF Level 3 reports use shared working folder
        # All QF reports work in working/NBD-QF-23-C3-C10-SF-FI/{date}/
        working_report_dir = WORKING_DIR / "NBD-QF-23-C3-C10-SF-FI"
        if not working_report_dir.exists():
            return False, f"QF working directory not found: {working_report_dir}"

        date_folders = [p for p in working_report_dir.iterdir() if p.is_dir()]
        if len(date_folders) != 1:
            return False, f"Expected exactly one date folder in {working_report_dir}, found {len(date_folders)}"

        working_dir = date_folders[0]
        _emit_log(f"Using QF shared working folder: {working_dir}", important=True)
    elif report_id == "NBD-QF-23-C7":
        # C7 report uses its own working folder
        working_report_dir = WORKING_DIR / "NBD_QF_23_C7"
        if not working_report_dir.exists():
            return False, f"C7 working directory not found: {working_report_dir}"
        date_folders = [p for p in working_report_dir.iterdir() if p.is_dir()]
        if len(date_folders) != 1:
            return False, f"Expected exactly one date folder in {working_report_dir}, found {len(date_folders)}"
        working_dir = date_folders[0]
    elif report_id == "NBD-WF-18-ECS":
        # ECS manages its own working folder internally (copies outputs→working after bots)
        # No pre-existing working folder required here
        working_dir = WORKING_DIR / folder_name
    else:
        working_report_dir = WORKING_DIR / folder_name
        if not working_report_dir.exists():
            return False, f"Working directory not found: {working_report_dir}"
        date_folders = [p for p in working_report_dir.iterdir() if p.is_dir()]
        if len(date_folders) != 1:
            # Provide detailed error message showing what folders were found
            folder_names = [f.name for f in date_folders] if date_folders else []
            return False, f"Expected exactly one date folder in {working_report_dir}, found {len(date_folders)}: {folder_names}"
        working_dir = date_folders[0]

    _emit_log(f"Starting {report_id}: {REPORT_STRUCTURE[report_id]['name']}", important=True)
    _set_report_status(report_id, "running")

    # Build command according to script-specific CLI expectations
    ui_date = None
    # Convert UI date to the formats scripts expect
    try:
        ui_dt = datetime.strptime(date_str, "%Y-%m-%d")
    except Exception:
        try:
            ui_dt = datetime.strptime(date_str, "%d-%m-%Y")
        except Exception:
            ui_dt = None
    ui_date_mmddyyyy = ui_dt.strftime("%m/%d/%Y") if ui_dt else date_str

    if script_name == "NBD_MF_01_SOFP_SOCI.py":
        cmd = [
            sys.executable,
            str(script_path),
            "--date", ui_date_mmddyyyy,
        ]
    elif script_name == "NBD_AF_01_SOFP_SOCI.py":
        cmd = [
            sys.executable,
            str(script_path),
            "--date", ui_date_mmddyyyy,
        ]
    elif script_name == "NBD_MF_23_IA.py":
        cmd = [
            sys.executable,
            str(script_path),
            ui_date_mmddyyyy,
        ]
        # Pass rate_updates as command line argument if available from global variable
        global _ia_rate_updates, _ia_ltv_updates
        if _ia_rate_updates:
            cmd.extend(['--rate-updates', json.dumps(_ia_rate_updates)])
            _emit_log(f"Passing {len(_ia_rate_updates)} rate updates to IA script", important=True)
        # Pass ltv_updates as command line argument if available from global variable
        if _ia_ltv_updates:
            cmd.extend(['--ltv-updates', json.dumps(_ia_ltv_updates)])
            _emit_log(f"Passing {len(_ia_ltv_updates)} LTV updates to IA script", important=True)
    elif script_name == "NBD_MF_23_C1C2.py":
        cmd = [
            sys.executable,
            str(script_path),
        ]
    elif script_name == "NBD_MF_20_C3.py":
        # C3 now works like C4 - auto-discovers files from working directory
        # Pass the parent NBD_MF_20_C1_C6 folder, not the dated folder
        base_working_dir = working_dir.parent if working_dir.name != "NBD_MF_20_C1_C6" else working_dir
        cmd = [
            sys.executable,
            str(script_path),
            "--working-dir", str(base_working_dir),
        ]
        if report_month and report_year:
            cmd.extend(["--month", report_month, "--year", report_year])
    elif script_name == "NBD_MF_20_C6.py":
        # C6 script expects --gui-date, --month, and --year parameters for SOFP copy operation
        base_working_dir = working_dir.parent if working_dir.name != "NBD_MF_20_C1_C6" else working_dir
        cmd = [
            sys.executable,
            str(script_path),
            "--working-dir", str(base_working_dir),
        ]
        # Pass exact date for SOFP copy (convert from YYYY-MM-DD to DD-MM-YYYY)
        if date_str:
            try:
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                date_ddmmyyyy = dt.strftime("%d-%m-%Y")
                cmd.extend(["--gui-date", date_ddmmyyyy])
            except:
                cmd.extend(["--gui-date", date_str])
        # Pass month and year for file parsing
        if report_month and report_year:
            cmd.extend(["--month", report_month, "--year", report_year])
    elif script_name == "NBD_MF_10_GA_NBD_MF_11_IS.py":
        # GA & IS script expects --working-dir, --gui-date, --month, and --year parameters
        cmd = [
            sys.executable,
            str(script_path),
            "--working-dir", str(working_dir),
        ]
        # Pass exact date for borrowing filtering
        if date_str:
            cmd.extend(["--gui-date", date_str])
        # Pass month and year for other operations (file naming, etc.)
        if report_month and report_year:
            cmd.extend(["--month", report_month, "--year", report_year])
    elif script_name == "NBD_QF_23_C8.py":
        # C8 script expects date in MM/DD/YYYY format
        cmd = [
            sys.executable,
            str(script_path),
            ui_date_mmddyyyy,
        ]
    elif script_name == "NBD_QF_23_C7.py":
        # C7 script expects --gui-date and --working-dir arguments
        cmd = [
            sys.executable,
            str(script_path),
            "--gui-date", date_str,
            "--working-dir", str(working_dir),
        ]
    elif script_name == "NBD_WF_15_LA.py":
        # LA script expects date in MM/DD/YYYY format
        cmd = [
            sys.executable,
            str(script_path),
            ui_date_mmddyyyy,
        ]
        if _la_month_end:
            cmd.append("--month-end")
            if _la_time_deposits:
                cmd.extend(["--time-deposits", str(_la_time_deposits)])
    elif script_name == "NBD_WF_18_ECS.py":
        # ECS script expects date in MM/DD/YYYY format
        cmd = [
            sys.executable,
            str(script_path),
            ui_date_mmddyyyy,
        ]
    elif script_name == "NBD_MF_04_LA.py":
        # MF-04-LA script expects date in MM/DD/YYYY format
        cmd = [
            sys.executable,
            str(script_path),
            ui_date_mmddyyyy,
        ]
    elif script_name == "NBD_WF_18_DM.py":
        # DM script expects date in MM/DD/YYYY format
        cmd = [
            sys.executable,
            str(script_path),
            ui_date_mmddyyyy,
        ]
    elif script_name in ("NBD_QF_23_C3.py", "NBD_QF_23_C4.py", "NBD_QF_23_C6.py",
                          "NBD_QF_23_C9.py", "NBD_QF_23_C10.py",
                          "NBD_QF_16_SF.py", "NBD_QF_16_FI.py"):
        # QF Level 3 reports - these scripts run on existing data
        # Pass working directory and date parameters
        cmd = [
            sys.executable,
            str(script_path),
            "--working-dir", str(working_dir),
        ]
        if report_month and report_year:
            cmd.extend(["--month", report_month, "--year", report_year])
    else:
        cmd = [
            sys.executable,
            str(script_path),
            "--month", report_month,
            "--year", report_year,
        ]

    log_file = _run_log_dir / f"{report_id}.log"

    try:
        _emit_log(f"Running: {' '.join(cmd)} (cwd={working_dir})", important=True)

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        with subprocess.Popen(
            cmd,
            cwd=str(working_dir),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            bufsize=1,
            env=env,
        ) as proc:
            with _proc_lock:
                _running_processes[report_id] = proc

            with open(log_file, 'w', encoding='utf-8') as lf:
                for line in proc.stdout:
                    if _stop_requested:
                        proc.terminate()
                        _set_report_status(report_id, "stopped")
                        return False, "Stopped by user"

                    raw = line.rstrip()
                    lf.write(raw + '\n')
                    lf.flush()
                    _emit_log(raw, important=False)
                    _append_report_message(report_id, raw)

            ret = proc.wait()

            with _proc_lock:
                if report_id in _running_processes:
                    del _running_processes[report_id]

            if ret != 0:
                _emit_log(f"{report_id} failed with exit code {ret}", important=True)
                _set_report_status(report_id, "failed")
                return False, f"Failed with exit code {ret}"

            _emit_log(f"{report_id} completed successfully", important=True)
            _set_report_status(report_id, "completed")
            return True, "Completed"

    except Exception as exc:
        _emit_log(f"Failed to run {report_id}: {exc}", important=True)
        _set_report_status(report_id, "failed")
        return False, str(exc)


def find_output_files(report_id: str, file_pattern: str) -> list[Path]:
    """
    Find output files in working directory matching the pattern.
    Returns list of matching file paths.
    """
    if report_id not in REPORT_STRUCTURE:
        return []
    
    folder_name = REPORT_STRUCTURE[report_id]["folder"]
    working_report_dir = WORKING_DIR / folder_name
    
    if not working_report_dir.exists():
        return []
    
    date_folders = [p for p in working_report_dir.iterdir() if p.is_dir()]
    if len(date_folders) != 1:
        return []
    
    date_folder = date_folders[0]
    
    # Find files matching pattern
    matches = glob.glob(str(date_folder / file_pattern))
    return [Path(p) for p in matches]


def _copy_af_sofp_fs_to_outputs_targets() -> None:
    """Copy latest AF SOFP Annual FS file from working AF SOFP date folder to
    latest versioned date folder of outputs/NBD_AF_20_C1_C6.
    If a file with same name exists at destination, remove it before copying.
    """
    try:
        # Locate source file in working/AF SOFP date folder
        sofp_dir = WORKING_DIR / REPORT_STRUCTURE["NBD-AF-01-SOFP-SOCI"]["folder"]
        sofp_dates = [p for p in sofp_dir.iterdir() if p.is_dir()]
        if len(sofp_dates) != 1:
            _emit_log("AF SOFP working folder not in expected single-date structure; skipping extra copies", important=True)
            return
        sofp_date_dir = sofp_dates[0]
        matches = glob.glob(str(sofp_date_dir / "NBD-AF-01-SOFP & SOCI AFL Annual FS*.xlsx"))
        if not matches:
            _emit_log("No AF SOFP Annual FS file found to copy", important=True)
            return
        src_file = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))

        # Copy to C1_C6 outputs latest date folder
        out_dir = OUTPUTS_DIR / REPORT_STRUCTURE["NBD-AF-20-C2"]["folder"]
        last_date, date_folder = find_latest_completed_date(out_dir)
        if not last_date or not date_folder or not date_folder.exists():
            _emit_log(f"Skip copy to {out_dir} (no completed date folder found)", important=True)
            return
        dest_path = date_folder / src_file.name
        try:
            # Remove any existing .xlsx file with the same prefix in the target folder
            prefix = "NBD-AF-01-SOFP & SOCI AFL Annual FS"
            removed = 0
            for f in date_folder.iterdir():
                if f.is_file() and f.suffix.lower() == ".xlsx" and f.name.startswith(prefix):
                    try:
                        f.unlink()
                        removed += 1
                    except Exception as _e:
                        _emit_log(f"Warning: failed removing existing file {f}: {_e}", important=True)
            if removed:
                _emit_log(f"Removed {removed} existing AF SOFP Annual FS .xlsx file(s) in {date_folder}", important=True)
            shutil.copy2(src_file, dest_path)
            _emit_log(f"Copied {src_file.name} to {dest_path}", important=True)
        except Exception as e:
            _emit_log(f"Failed to copy {src_file.name} to {dest_path}: {e}", important=True)

    except Exception as e:
        _emit_log(f"Extra copy of AF SOFP FS to outputs targets failed: {e}", important=True)


def _copy_sofp_fs_to_outputs_targets() -> None:
    """Copy latest SOFP Monthly FS file from working SOFP date folder to
    latest versioned date folders of outputs/NBD_MF_23_IA and outputs/NBD_MF_20_C1_C6.
    If a file with same name exists at destination, remove it before copying.
    """
    try:
        # Locate source file in working/SOFP date folder
        sofp_dir = WORKING_DIR / REPORT_STRUCTURE["NBD-MF-01-SOFP-SOCI"]["folder"]
        sofp_dates = [p for p in sofp_dir.iterdir() if p.is_dir()]
        if len(sofp_dates) != 1:
            _emit_log("SOFP working folder not in expected single-date structure; skipping extra copies", important=True)
            return
        sofp_date_dir = sofp_dates[0]
        matches = glob.glob(str(sofp_date_dir / "NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx"))
        if not matches:
            _emit_log("No SOFP Monthly FS file found to copy", important=True)
            return
        src_file = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))

        # Helper to copy into a target outputs/<folder> latest date version
        def copy_into_outputs_folder(folder_key: str) -> None:
            out_dir = OUTPUTS_DIR / REPORT_STRUCTURE[folder_key]["folder"]
            last_date, date_folder = find_latest_completed_date(out_dir)
            if not last_date or not date_folder or not date_folder.exists():
                _emit_log(f"Skip copy to {out_dir} (no completed date folder found)", important=True)
                return
            dest_path = date_folder / src_file.name
            try:
                # Remove any existing .xlsx file with the same prefix in the target folder
                prefix = "NBD-MF-01-SOFP & SOCI AFL Monthly FS"
                removed = 0
                for f in date_folder.iterdir():
                    if f.is_file() and f.suffix.lower() == ".xlsx" and f.name.startswith(prefix):
                        try:
                            f.unlink()
                            removed += 1
                        except Exception as _e:
                            _emit_log(f"Warning: failed removing existing file {f}: {_e}", important=True)
                if removed:
                    _emit_log(f"Removed {removed} existing SOFP Monthly FS .xlsx file(s) in {date_folder}", important=True)
                shutil.copy2(src_file, dest_path)
                _emit_log(f"Copied {src_file.name} to {dest_path}", important=True)
            except Exception as e:
                _emit_log(f"Failed to copy {src_file.name} to {dest_path}: {e}", important=True)

        # Copy to IA outputs latest date folder
        copy_into_outputs_folder("NBD-MF-23-IA")
        # Copy to C1_C6 outputs latest date folder
        copy_into_outputs_folder("NBD-MF-20-C1")
        # Copy to GA & IS outputs latest date folder
        try:
            copy_into_outputs_folder("NBD-MF-10-GA-11-IS")
        except Exception as _e:
            _emit_log(f"Skip copy to GA & IS outputs: {_e}", important=True)

    except Exception as e:
        _emit_log(f"Extra copy of SOFP FS to outputs targets failed: {e}", important=True)


def save_folder_to_outputs(report_id: str, selected_date: str) -> tuple[bool, str, Path | None]:
    """
    Copy report folder from working back to outputs with selected date.
    Returns (success, message, output_path)
    """
    if report_id not in REPORT_STRUCTURE:
        return False, f"Unknown report: {report_id}", None

    folder_name = REPORT_STRUCTURE[report_id]["folder"]
    # Special-case GA & IS SET8: working folder uses underscore name; outputs uses ampersand name
    if report_id == "NBD-MF-10-GA-11-IS":
        working_report_dir = WORKING_DIR / "NBD_MF_10_GA_NBD_MF_11_IS"
        output_report_dir = OUTPUTS_DIR / "NBD-MF-10-GA & NBD-MF-11-IS"
    elif report_id in ("NBD-QF-23-C3", "NBD-QF-23-C4", "NBD-QF-23-C6",
                       "NBD-QF-23-C9", "NBD-QF-23-C10",
                       "NBD-QF-16-SF", "NBD-QF-16-FI"):
        # QF Level 3 reports all use shared folder
        working_report_dir = WORKING_DIR / "NBD-QF-23-C3-C10-SF-FI"
        output_report_dir = OUTPUTS_DIR / "NBD-QF-23-C3-C10-SF-FI"
    else:
        working_report_dir = WORKING_DIR / folder_name
        output_report_dir = OUTPUTS_DIR / folder_name
    
    # Find the date folder in working
    date_folders = [p for p in working_report_dir.iterdir() if p.is_dir()]
    if len(date_folders) != 1:
        return False, f"Expected one date folder in {working_report_dir}", None
    
    source_folder = date_folders[0]
    target_date_name = format_date_folder_name(selected_date)
    
    # Create output report directory
    output_report_dir.mkdir(parents=True, exist_ok=True)
    
    # Check if target date folder exists
    target_path = output_report_dir / target_date_name
    if target_path.exists():
        # Find next version number
        version = 1
        while (output_report_dir / f"{target_date_name}({version})").exists():
            version += 1
        target_path = output_report_dir / f"{target_date_name}({version})"
    
    try:
        shutil.copytree(source_folder, target_path)
        _emit_log(f"Saved {report_id} to {target_path}", important=True)
        return True, f"Saved to {target_path.name}", target_path
    except Exception as e:
        return False, f"Failed to save: {e}", None


def get_report_dependencies(report_id: str) -> list[str]:
    """Get all dependencies for a report (recursive)."""
    if report_id not in REPORT_STRUCTURE:
        return []

    deps = []
    for dep in REPORT_STRUCTURE[report_id]["dependencies"]:
        deps.append(dep)
        deps.extend(get_report_dependencies(dep))

    seen = set()
    result = []
    for item in deps:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result


def get_execution_order(report_ids: list[str]) -> list[str]:
    """Get execution order including all dependencies."""
    all_reports = set()

    # IA and C1C2 always run together - no standalone C1C2 anymore
    for report_id in report_ids:
        all_reports.add(report_id)
        all_reports.update(get_report_dependencies(report_id))

    sorted_reports = sorted(
        all_reports,
        key=lambda x: (REPORT_STRUCTURE[x]["level"], REPORT_STRUCTURE[x]["set"], REPORT_STRUCTURE[x]["step"])
    )

    return sorted_reports


def run_selected_reports(report_ids_with_dates: dict) -> tuple[bool, str]:
    """Run selected reports in correct order with modular approach."""
    global _download_files, _sofp_download_file, _ia_rate_updates, _ia_ltv_updates
    _download_files = []
    _sofp_download_file = None

    try:
        report_ids = list(report_ids_with_dates.keys())
        execution_order = get_execution_order(report_ids)
    except Exception as e:
        _emit_log(f"Error in execution order: {e}", important=True)
        import traceback
        _emit_log(traceback.format_exc(), important=True)
        return False, f"Error in execution order: {e}"
    # Ensure SOFP appears first when IA is selected WITHOUT C1-C2
    # When both IA and C1-C2 are selected, don't include SOFP
    if "NBD-MF-23-IA" in report_ids and "NBD-MF-23-C1-C2" not in report_ids:
        if "NBD-MF-01-SOFP-SOCI" in execution_order:
            execution_order = ["NBD-MF-01-SOFP-SOCI"] + [r for r in execution_order if r != "NBD-MF-01-SOFP-SOCI"]

    _emit_log(f"Execution order: {', '.join(execution_order)}", important=True)
    only_sofp_flow = set(report_ids) == {"NBD-MF-01-SOFP-SOCI"} or set(report_ids) == {"NBD-AF-01-SOFP-SOCI"}

    # Build checklist: Only show reports that will actually run
    checklist_ids = list(execution_order)

    # Determine if we should show SOFP in the checklist
    # SOFP is shown only when IA is selected WITHOUT C1-C2 (but now they always run together)
    # Since IA and C1C2 always run together, we don't show SOFP in the checklist
    # The validation will ensure SOFP file exists before running

    _init_report_checklist(checklist_ids)
    
    # Detect if this is the C1–C6 set flow
    c1c6_ids = {"NBD-MF-20-C1","NBD-MF-20-C2","NBD-MF-20-C3","NBD-MF-20-C4","NBD-MF-20-C5","NBD-MF-20-C6"}
    annual_c_ids = {"NBD-AF-20-C2","NBD-AF-20-C3","NBD-AF-20-C4","NBD-AF-20-C5","NBD-AF-20-C6"}
    is_c1c6_flow = any(r in c1c6_ids for r in execution_order)
    is_annual_c_flow = any(r in annual_c_ids for r in execution_order)
    is_ga_is_set8 = any(r == "NBD-MF-10-GA-11-IS" for r in execution_order)
    is_c7_flow = any(r == "NBD-QF-23-C7" for r in execution_order)
    is_c8_flow = any(r == "NBD-QF-23-C8" for r in execution_order)
    is_c4_flow = any(r == "NBD-MF-23-C4" for r in execution_order)
    is_dm_flow = any(r == "NBD-WF-18-DM" for r in execution_order)
    is_la_flow = any(r == "NBD-WF-15-LA" for r in execution_order)
    is_ecs_flow = any(r == "NBD-WF-18-ECS" for r in execution_order)
    is_mf04_flow = any(r == "NBD-MF-04-LA" for r in execution_order)

    # Detect QF reports group flow (C3, C4, C6, C9, C10, SF, FI only - excludes MF-C4 which is monthly)
    qf_reports_group = {"NBD-QF-16-FI", "NBD-QF-16-SF", "NBD-QF-23-C3", "NBD-QF-23-C4", "NBD-QF-23-C6", "NBD-QF-23-C10", "NBD-QF-23-C9"}
    is_qf_reports_flow = any(r in qf_reports_group for r in execution_order)

    # Detect IA flow (IA and C1C2 always run together)
    is_ia_flow = "NBD-MF-23-IA" in report_ids or "NBD-MF-23-C1-C2" in report_ids
    
    # Copy all required reports to working
    _set_status(stage="copying", message="Copying reports to working directory")
    try:
        if is_c1c6_flow:
            # Copy only NBD_MF_20_C1_C6 based on completed_dates.txt
            _emit_log("Preparing working folder for C1–C6 set", important=True)
            ok, msg, _ = copy_folder_to_working("NBD-MF-20-C1")
            if not ok:
                return False, f"Failed to copy NBD-MF-20-C1_C6: {msg}"
        elif is_annual_c_flow:
            # Copy only NBD_AF_20_C1_C6 based on completed_dates.txt
            _emit_log("Preparing working folder for Annual C2–C6 set", important=True)
            ok, msg, _ = copy_folder_to_working("NBD-AF-20-C2")
            if not ok:
                return False, f"Failed to copy NBD-AF-20-C1_C6: {msg}"
        elif is_ga_is_set8:
            _emit_log("Preparing working folder for GA & IS Set 8", important=True)
            ok, msg, _ = copy_folder_to_working("NBD-MF-10-GA-11-IS")
            if not ok:
                return False, f"Failed to copy GA & IS: {msg}"
        elif is_c7_flow:
            _emit_log("Preparing working folder for C7", important=True)
            ok, msg, _ = copy_folder_to_working("NBD-QF-23-C7")
            if not ok:
                return False, f"Failed to copy C7: {msg}"
        elif is_c8_flow:
            _emit_log("Preparing working folder for C8", important=True)
            ok, msg, _ = copy_folder_to_working("NBD-QF-23-C8")
            if not ok:
                return False, f"Failed to copy C8: {msg}"
        elif is_c4_flow:
            _emit_log("Preparing working folder for C4", important=True)
            ok, msg, _ = copy_folder_to_working("NBD-MF-23-C4")
            if not ok:
                return False, f"Failed to copy C4: {msg}"
        # Set 2 reports (DM, LA, ECS, MF04) use separate if blocks so each
        # working folder is prepared independently when running together
        if is_dm_flow:
            _emit_log("Preparing working folder for DM", important=True)
            # Copy the latest date folder from outputs to working (remove existing folder first)
            try:
                dm_outputs_dir = OUTPUTS_DIR / "NBD-WF-18-DM"
                if dm_outputs_dir.exists():
                    last_date, date_folder = find_latest_completed_date(dm_outputs_dir)
                    if date_folder and date_folder.exists():
                        dm_working_dir = WORKING_DIR / "NBD-WF-18-DM"

                        # Remove existing working folder if it exists
                        if dm_working_dir.exists():
                            import shutil
                            shutil.rmtree(dm_working_dir)
                            _emit_log("Removed existing DM working folder", important=True)

                        # Create new working directory
                        dm_working_dir.mkdir(parents=True, exist_ok=True)

                        # Copy the latest date folder to working
                        target_working_folder = dm_working_dir / date_folder.name
                        shutil.copytree(date_folder, target_working_folder)
                        _emit_log(f"Copied latest DM date folder from outputs to working: {date_folder.name}", important=True)
                    else:
                        _emit_log("No completed DM date folder found in outputs, proceeding with normal copy", important=True)
                        # Do the normal copy if no completed folder found
                        ok, msg, _ = copy_folder_to_working("NBD-WF-18-DM")
                        if not ok:
                            return False, f"Failed to copy DM: {msg}"
                else:
                    _emit_log("DM outputs directory not found, proceeding with normal copy", important=True)
                    # Do the normal copy if outputs directory doesn't exist
                    ok, msg, _ = copy_folder_to_working("NBD-WF-18-DM")
                    if not ok:
                        return False, f"Failed to copy DM: {msg}"
            except Exception as e:
                _emit_log(f"Failed to copy latest DM date folder from outputs: {e}, proceeding with normal copy", important=True)
                # Do the normal copy if copying fails
                ok, msg, _ = copy_folder_to_working("NBD-WF-18-DM")
                if not ok:
                    return False, f"Failed to copy DM: {msg}"
        if is_la_flow:
            _emit_log("Preparing working folder for LA", important=True)

            la_working_dir = WORKING_DIR / "NBD_WF_15_LA"
            la_outputs_dir = OUTPUTS_DIR / "NBD_WF_15_LA"

            # Always clear working folder and copy fresh from outputs
            try:
                import shutil
                if la_working_dir.exists():
                    shutil.rmtree(la_working_dir)
                    _emit_log("Cleared existing LA working folder", important=True)

                if la_outputs_dir.exists():
                    last_date, date_folder = find_latest_completed_date(la_outputs_dir)
                    if date_folder and date_folder.exists():
                        la_working_dir.mkdir(parents=True, exist_ok=True)

                        target_working_folder = la_working_dir / date_folder.name
                        shutil.copytree(date_folder, target_working_folder)
                        _emit_log(f"Copied latest LA date folder from outputs to working: {date_folder.name}", important=True)
                    else:
                        _emit_log("No completed LA date folder found in outputs", important=True)
                        return False, "No LA date folder found in outputs. Please ensure LA outputs folder has completed runs."
                else:
                    _emit_log("LA outputs directory not found", important=True)
                    return False, "LA outputs directory not found"
            except Exception as e:
                _emit_log(f"Failed to copy LA folder from outputs: {e}", important=True)
                return False, f"Failed to copy LA folder from outputs: {e}"

            # Copy Daily Bank Balances from Master Inputs to LA working folder
            _emit_log("Copying Daily Bank Balances to LA working folder...", important=True)
            try:
                # Find the Daily Bank Balances file in Master Inputs
                daily_bank_dir = MASTER_INPUTS_DIR / "Daily Bank Balances"
                if daily_bank_dir.exists():
                    daily_bank_files = list(daily_bank_dir.glob("Daily Bank Balances*.xlsx"))
                    if daily_bank_files:
                        # Get the most recent file (by modification time)
                        latest_file = max(daily_bank_files, key=lambda f: f.stat().st_mtime)

                        # Find LA working folder (refresh after copy)
                        if la_working_dir.exists():
                            # Find the date subfolder in working
                            date_folders = [p for p in la_working_dir.iterdir() if p.is_dir()]
                            if date_folders:
                                working_date_folder = date_folders[0]

                                # Remove existing Daily Bank Balances files
                                for existing in working_date_folder.glob("Daily Bank Balances*.xlsx"):
                                    try:
                                        existing.unlink()
                                        _emit_log(f"Removed existing: {existing.name}", important=True)
                                    except Exception as e:
                                        _emit_log(f"Could not remove {existing.name}: {e}", important=True)

                                # Copy the file
                                import shutil
                                dest_path = working_date_folder / latest_file.name
                                shutil.copy2(latest_file, dest_path)
                                _emit_log(f"✓ Copied Daily Bank Balances to LA working: {latest_file.name}", important=True)
                            else:
                                _emit_log("No date folder found in LA working directory after copy", important=True)
                        else:
                            _emit_log("LA working directory not found", important=True)
                    else:
                        _emit_log("No Daily Bank Balances file found in Master Inputs", important=True)
                else:
                    _emit_log("Daily Bank Balances Master Inputs directory not found", important=True)
            except Exception as e:
                _emit_log(f"Warning: Failed to copy Daily Bank Balances to LA working: {e}", important=True)
                # Continue anyway - the file might already be in working folder
        if is_mf04_flow:
            _emit_log("Preparing working folder for MF-04-LA (Monthly Liquidity)", important=True)
            # Copy the latest date folder from outputs to working
            try:
                mf04_outputs_dir = OUTPUTS_DIR / "NBD_MF_04_LA"
                if mf04_outputs_dir.exists():
                    last_date, date_folder = find_latest_completed_date(mf04_outputs_dir)
                    if date_folder and date_folder.exists():
                        mf04_working_dir = WORKING_DIR / "NBD_MF_04_LA"

                        # Remove existing working folder if it exists
                        if mf04_working_dir.exists():
                            import shutil
                            shutil.rmtree(mf04_working_dir)
                            _emit_log("Removed existing MF-04-LA working folder", important=True)

                        # Create new working directory
                        mf04_working_dir.mkdir(parents=True, exist_ok=True)

                        # Copy the latest date folder to working
                        target_working_folder = mf04_working_dir / date_folder.name
                        shutil.copytree(date_folder, target_working_folder)
                        _emit_log(f"Copied latest MF-04-LA date folder from outputs to working: {date_folder.name}", important=True)
                    else:
                        _emit_log("No completed MF-04-LA date folder found in outputs", important=True)
                        return False, "No MF-04-LA date folder found in outputs. Please ensure MF-04-LA outputs folder has completed runs."
                else:
                    _emit_log("MF-04-LA outputs directory not found", important=True)
                    return False, "MF-04-LA outputs directory not found"
            except Exception as e:
                _emit_log(f"Failed to copy MF-04-LA folder from outputs: {e}", important=True)
                return False, f"Failed to copy MF-04-LA folder from outputs: {e}"
        if is_qf_reports_flow:
            _emit_log("Preparing working folder for QF reports", important=True)
            # Copy NBD-QF-23-C3-C10-SF-FI from outputs to working
            # All QF reports (C3, C4, C6, C9, C10, SF, FI) use this shared folder
            ok, msg, _ = copy_folder_to_working("NBD-QF-23-C3")
            if not ok:
                return False, f"Failed to copy QF reports folder: {msg}"
        elif is_ia_flow:
            _emit_log("Preparing working folder for IA flow (IA and C1C2 run together)", important=True)
            # Copy IA folder first
            ok, msg, _ = copy_folder_to_working("NBD-MF-23-IA")
            if not ok:
                return False, f"Failed to copy IA: {msg}"
            # C1-C2 will use the same IA working folder
    except Exception as e:
        _emit_log(f"Error during copy phase: {e}", important=True)
        import traceback
        _emit_log(traceback.format_exc(), important=True)
        return False, f"Copy phase error: {e}"
    else:
        allowed_copy = {"NBD-MF-01-SOFP-SOCI", "NBD-AF-01-SOFP-SOCI"}
        for report_id in execution_order:
            # Copy only SOFP-SOCI (monthly and annual) and IA; skip others
            if report_id not in allowed_copy:
                _emit_log(f"Skipping copy for {report_id} (copy only SOFP-SOCI and IA)", important=True)
                continue
            success, msg, _ = copy_folder_to_working(report_id)
            if not success:
                return False, f"Failed to copy {report_id}: {msg}"
    
    # Track which scripts have been run to avoid duplicates
    executed_scripts = {}
    all_success = True
    failed_reports = []

    # For QF reports, run process_summary.py instead of individual scripts
    if is_qf_reports_flow:
        _emit_log("QF Reports flow detected - will run process_summary.py for all selected reports", important=True)
        _set_status(stage="running", message="Running QF reports process_summary")

        # Run process_summary.py
        try:
            script_path = REPORT_AUTOMATIONS_DIR / "process_summary.py"
            _emit_log(f"Running process_summary.py for QF reports", important=True)

            # Get the report date from the first QF report in the execution order
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date

            result = subprocess.run(
                [sys.executable, str(script_path), "--report-date", gui_date_str],
                capture_output=True,
                text=True,
                cwd=str(REPORT_AUTOMATIONS_DIR)
            )

            # Log output
            if result.stdout:
                _emit_log(f"process_summary.py output:\n{result.stdout}", important=False)

            if result.stderr:
                _emit_log(f"process_summary.py stderr:\n{result.stderr}", important=True)

            if result.returncode == 0:
                _emit_log("✓ process_summary.py completed successfully", important=True)
                # Mark all QF reports as completed
                for report_id in execution_order:
                    if report_id in qf_reports_group:
                        _set_report_status(report_id, "completed")
            else:
                _emit_log(f"✗ process_summary.py failed with return code {result.returncode}", important=True)
                # Mark all QF reports as failed
                for report_id in execution_order:
                    if report_id in qf_reports_group:
                        _set_report_status(report_id, "failed")
                        failed_reports.append(report_id)
                all_success = False
        except Exception as e:
            _emit_log(f"Error running process_summary.py: {e}", important=True)
            import traceback
            _emit_log(traceback.format_exc(), important=True)
            # Mark all QF reports as failed
            for report_id in execution_order:
                if report_id in qf_reports_group:
                    _set_report_status(report_id, "failed")
                    failed_reports.append(report_id)
            all_success = False

        # Skip the normal execution loop for QF reports
        sequence = []
    elif is_c1c6_flow:
        # Specific enforced order for C1–C6: C6 → C5 → C4 → C3 → C2
        enforced_order = ["NBD-MF-20-C6","NBD-MF-20-C5","NBD-MF-20-C4","NBD-MF-20-C3","NBD-MF-20-C2"]
        sequence = [rid for rid in enforced_order if rid in execution_order]
    elif is_annual_c_flow:
        # Specific enforced order for Annual C2–C6: C6 → C5 → C4 → C3 → C2
        enforced_order = ["NBD-AF-20-C6","NBD-AF-20-C5","NBD-AF-20-C4","NBD-AF-20-C3","NBD-AF-20-C2"]
        sequence = [rid for rid in enforced_order if rid in execution_order]
    else:
        sequence = execution_order

    for report_id in sequence:
        if _stop_requested:
            _emit_log("Stop requested, halting execution", important=True)
            break
            
        # Get date for this report
        if report_id in report_ids_with_dates:
            date_str, report_month, report_year = report_ids_with_dates[report_id]
        else:
            date_str, report_month, report_year = list(report_ids_with_dates.values())[0]
        
        script_name = REPORT_STRUCTURE[report_id]["script"]

        # Determine if this report should run based on the flow
        # Normal flows (IA and C1C2 always run together)
        run_allowed = (report_id in ("NBD-MF-01-SOFP-SOCI", "NBD-AF-01-SOFP-SOCI", "NBD-MF-23-IA", "NBD-MF-23-C1-C2")) or (is_c1c6_flow and report_id in c1c6_ids) or (is_annual_c_flow and report_id in annual_c_ids) or (is_ga_is_set8 and report_id == "NBD-MF-10-GA-11-IS") or (is_c8_flow and report_id == "NBD-QF-23-C8") or (is_c7_flow and report_id == "NBD-QF-23-C7") or (is_c4_flow and report_id == "NBD-MF-23-C4") or (is_dm_flow and report_id == "NBD-WF-18-DM") or (is_la_flow and report_id == "NBD-WF-15-LA") or (is_ecs_flow and report_id == "NBD-WF-18-ECS") or (is_mf04_flow and report_id == "NBD-MF-04-LA") or (is_qf_reports_flow and report_id in qf_reports_group) or (is_ia_flow and report_id in ("NBD-MF-23-IA", "NBD-MF-23-C1-C2"))
        
        # Check if this script was already executed
        if script_name in executed_scripts:
            prev_report_id = executed_scripts[script_name]
            _emit_log(f"Skipping {report_id} - script {script_name} already executed by {prev_report_id}", important=True)
            _set_report_status(report_id, "completed")
            continue
        
        # Skip running non-required reports while keeping them visible
        if not run_allowed:
            _emit_log(f"Skipping run for {report_id} (only SOFP-SOCI → IA → C1C2 run in this set)", important=True)
            _set_report_status(report_id, "skipped")
            continue
        
        # Emit clear sequence status line for the UI
        try:
            seq_name = REPORT_STRUCTURE[report_id]["folder"]
        except Exception:
            seq_name = report_id
        _set_status(stage="running", message=f"running {seq_name}")

        # For IA, copy its latest working folder only after SOFP completes
        if not is_c1c6_flow and report_id == "NBD-MF-23-IA":
            _emit_log("Preparing IA working folder after SOFP completion", important=True)
            ok_copy, copy_msg, _copy_path = copy_folder_to_working("NBD-MF-23-IA")
            if not ok_copy:
                _set_report_status(report_id, "failed")
                return False, f"Failed to copy NBD-MF-23-IA: {copy_msg}"
            # Remove old SOFP monthly FS file in IA date folder, if exists
            try:
                ia_working_dir = WORKING_DIR / REPORT_STRUCTURE["NBD-MF-23-IA"]["folder"]
                date_folders = [p for p in ia_working_dir.iterdir() if p.is_dir()]
                if len(date_folders) == 1:
                    ia_date_dir = date_folders[0]
                    pattern = ia_date_dir / "NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx"
                    match_paths = [Path(p) for p in glob.glob(str(pattern))]
                    if len(match_paths) > 1:
                        # Keep the latest by modified time; remove older ones
                        match_paths.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                        keep = match_paths[0]
                        to_remove = match_paths[1:]
                        removed = 0
                        for fp in to_remove:
                            try:
                                fp.unlink(missing_ok=True)
                                removed += 1
                            except Exception:
                                pass
                        if removed > 0:
                            _emit_log(f"Kept latest '{keep.name}', removed {removed} older SOFP Monthly FS file(s) in {ia_date_dir}", important=True)
                else:
                    _emit_log("IA working folder structure unexpected; skip old file cleanup", important=True)
            except Exception as _e:
                _emit_log(f"IA old file cleanup skipped: {_e}", important=True)
        
        # Run the report script
        success, msg = run_report_script(report_id, date_str, report_month, report_year)
        
        if success:
            executed_scripts[script_name] = report_id
            _set_status(stage="running", message=f"completed {seq_name}")
            # After SOFP completes: rename output file with GUI month/year (don't queue for download yet)
            if not is_c1c6_flow and not is_annual_c_flow and report_id in ("NBD-MF-01-SOFP-SOCI", "NBD-AF-01-SOFP-SOCI"):
                try:
                    # Determine GUI month/year from provided date
                    any_date = report_ids_with_dates.get(report_id) or list(report_ids_with_dates.values())[0]
                    _gui_date_str, gui_month, gui_year = any_date
                    sofp_work_dir = WORKING_DIR / REPORT_STRUCTURE[report_id]["folder"]
                    date_folders = [p for p in sofp_work_dir.iterdir() if p.is_dir()]
                    if len(date_folders) == 1:
                        date_folder = date_folders[0]
                        # Determine file pattern based on report type
                        if report_id == "NBD-MF-01-SOFP-SOCI":
                            file_pattern = "NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx"
                            base = "NBD-MF-01-SOFP & SOCI AFL Monthly FS"
                        else:  # NBD-AF-01-SOFP-SOCI
                            file_pattern = "NBD-AF-01-SOFP & SOCI AFL Annual FS*.xlsx"
                            base = "NBD-AF-01-SOFP & SOCI AFL Annual FS"
                        matches = glob.glob(str(date_folder / file_pattern))
                        if matches:
                            src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                            new_name = f"{base} {gui_month} {gui_year}.xlsx"
                            dst = src.with_name(new_name)
                            if src.name != new_name:
                                try:
                                    src.rename(dst)
                                    src = dst
                                    _emit_log(f"Renamed SOFP file to {dst.name}", important=True)
                                except Exception as e:
                                    _emit_log(f"SOFP rename failed (continuing with original): {e}", important=True)
                            # Store the renamed file info for later download queueing after save
                            _sofp_download_file = src

                        else:
                            _emit_log(f"SOFP file not found for renaming with pattern {file_pattern}", important=True)
                    else:
                        _emit_log("Unexpected SOFP working folder structure; skipping rename", important=True)
                except Exception as e:
                    _emit_log(f"SOFP post-processing failed: {e}", important=True)
            
            # Special handling for NBD-MF-23-C1 (download output file)
            if report_id == "NBD-MF-23-C1":
                _emit_log("Looking for output files to download...", important=True)
                output_files = find_output_files("NBD-MF-23-IA", "Prod. wise Class. of Loans*.xlsb")
                if output_files:
                    _download_files.extend(output_files)
                    _emit_log(f"Added {len(output_files)} file(s) for download", important=True)
                else:
                    _emit_log("Warning: No output files found matching pattern", important=True)
            # After SOFP completes in IA flow, copy its Monthly FS file to outputs/IA and outputs/C1_C6 latest date folders
            if not is_c1c6_flow and report_id == "NBD-MF-01-SOFP-SOCI":
                try:
                    _copy_sofp_fs_to_outputs_targets()
                except Exception as e:
                    _emit_log(f"Warning: extra copy of SOFP FS failed: {e}", important=True)
            # After AF SOFP completes, copy its Annual FS file to outputs/AF C1_C6 latest date folder
            if not is_annual_c_flow and report_id == "NBD-AF-01-SOFP-SOCI":
                try:
                    _copy_af_sofp_fs_to_outputs_targets()
                except Exception as e:
                    _emit_log(f"Warning: extra copy of AF SOFP FS failed: {e}", important=True)
            # Small delay before starting C1C2 after IA completes
            if not is_c1c6_flow and report_id == "NBD-MF-23-IA":
                try:
                    time.sleep(3)
                except Exception:
                    pass
        else:
            all_success = False
            failed_reports.append(f"{report_id}: {msg}")
    
    # Post-processing: rename/download/archive
    if is_c1c6_flow:
        # Rename the C1–C6 Excel file in working to GUI month/year
        try:
            c1c6_dir = WORKING_DIR / REPORT_STRUCTURE["NBD-MF-20-C1"]["folder"]
            date_folders = [p for p in c1c6_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                # Find target file
                matches = glob.glob(str(date_folder / "NBD-MF-20-C1 to C6*.xlsx"))
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    # Build new name from GUI month/year (use one of selected dates)
                    any_date = list(report_ids_with_dates.values())[0]
                    _, gui_month, gui_year = any_date
                    new_name = f"NBD-MF-20-C1 to C6 {gui_month} {gui_year}.xlsx"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            _emit_log(f"Renamed file to {dst.name}", important=True)
                            src = dst
                        except Exception as e:
                            _emit_log(f"Rename failed (continuing with original): {e}", important=True)
                    # Queue for download
                    _download_files.append(src)
                else:
                    _emit_log("No C1–C6 output file found to rename/download", important=True)
            else:
                _emit_log("Unexpected working folder structure for C1–C6; skipping rename/download", important=True)
        except Exception as e:
            _emit_log(f"C1–C6 post-processing failed: {e}", important=True)

        # Save working folder back to outputs with versioning using GUI date
        _set_status(stage="saving", message="Saving C1–C6 to outputs")
        date_str, _, _ = list(report_ids_with_dates.values())[0]
        ok, msg, output_path = save_folder_to_outputs("NBD-MF-20-C1", date_str)
        if ok and output_path:
            _emit_log(f"Saved C1–C6: {msg}", important=True)
            _last_output_dirs["NBD-MF-20-C1"] = output_path
    elif is_annual_c_flow:
        # Rename the Annual C2–C6 Excel file in working to GUI month/year
        try:
            annual_c_dir = WORKING_DIR / REPORT_STRUCTURE["NBD-AF-20-C2"]["folder"]
            date_folders = [p for p in annual_c_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                # Find target file
                matches = glob.glob(str(date_folder / "NBD-AF-20-C1 to C6*.xlsx"))
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    # Build new name from GUI month/year (use one of selected dates)
                    any_date = list(report_ids_with_dates.values())[0]
                    _, gui_month, gui_year = any_date
                    new_name = f"NBD-AF-20-C1 to C6 {gui_month} {gui_year}.xlsx"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            _emit_log(f"Renamed file to {dst.name}", important=True)
                            src = dst
                        except Exception as e:
                            _emit_log(f"Rename failed (continuing with original): {e}", important=True)
                    # Queue for download
                    _download_files.append(src)
                else:
                    _emit_log("No Annual C2–C6 output file found to rename/download", important=True)
            else:
                _emit_log("Unexpected working folder structure for Annual C2–C6; skipping rename/download", important=True)
        except Exception as e:
            _emit_log(f"Annual C2–C6 post-processing failed: {e}", important=True)

        # Save working folder back to outputs with versioning using GUI date
        _set_status(stage="saving", message="Saving Annual C2–C6 to outputs")
        date_str, _, _ = list(report_ids_with_dates.values())[0]
        ok, msg, output_path = save_folder_to_outputs("NBD-AF-20-C2", date_str)
        if ok and output_path:
            _emit_log(f"Saved Annual C2–C6: {msg}", important=True)
            _last_output_dirs["NBD-AF-20-C2"] = output_path
    elif is_ga_is_set8:
        # Post-processing for GA & IS: rename files, queue downloads, archive working folder
        try:
            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date
            # working date folder - use underscore name for working directory
            ga_work_dir = WORKING_DIR / "NBD_MF_10_GA_NBD_MF_11_IS"
            date_folders = [p for p in ga_work_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                # Rename GA/IS file
                matches = glob.glob(str(date_folder / "NBD-MF-10-GA & NBD-MF-11-IS*.xlsx"))
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    new_name = f"NBD-MF-10-GA & NBD-MF-11-IS {gui_month} {gui_year}.xlsx"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            src = dst
                            _emit_log(f"Renamed GA/IS file to {dst.name}", important=True)
                        except Exception as e:
                            _emit_log(f"GA/IS rename failed (continuing): {e}", important=True)
                    _download_files.append(src)
                    _emit_log(f"Added GA/IS file to download queue: {src.name}", important=True)
                else:
                    _emit_log("GA/IS .xlsx file not found for renaming", important=True)

                # Rename CBSL converted file to GUI month/year with .xlsb extension
                cbsl_matches = glob.glob(str(date_folder / "CBSL Provision Comparison*converted.xlsx"))
                cbsl_path_for_delete = None
                if cbsl_matches:
                    cbsl_src = Path(max(cbsl_matches, key=lambda x: Path(x).stat().st_mtime))
                    new_cbsl = cbsl_src.with_name(f"CBSL Provision Comparison - {gui_month} {gui_year}.xlsb")
                    try:
                        cbsl_src.rename(new_cbsl)
                        _emit_log(f"Renamed CBSL converted to {new_cbsl.name}", important=True)
                        _download_files.append(new_cbsl)
                        _emit_log(f"Added CBSL file to download queue: {new_cbsl.name}", important=True)
                        cbsl_path_for_delete = new_cbsl
                    except Exception as e:
                        _emit_log(f"CBSL rename failed: {e}", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                try:
                    dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                except Exception:
                    try:
                        dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                    except Exception:
                        dt = None
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-MF-10-GA-11-IS", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved GA & IS: {msg}", important=True)
                    _last_output_dirs["NBD-MF-10-GA-11-IS"] = output_path
                    
                    # Update file paths in _download_files to point to the new output location
                    updated_download_files = []
                    for file_path in _download_files:
                        if file_path.exists():
                            # File still exists at original location, keep it
                            updated_download_files.append(file_path)
                        else:
                            # File was moved, try to find it in the new output location
                            new_path = output_path / file_path.name
                            if new_path.exists():
                                updated_download_files.append(new_path)
                                _emit_log(f"Updated file path: {file_path.name} -> {new_path}", important=True)
                            else:
                                _emit_log(f"Could not find moved file: {file_path.name}", important=True)
                    
                    _download_files = updated_download_files
                    _emit_log(f"Updated download files after move: {len(_download_files)} files", important=True)

                # Prepare downloads
                _set_status(stage="saving", message="Preparing files for download")
                try:
                    auto_download_files()
                except Exception as e:
                    _emit_log(f"Warning: GA&IS download prep failed: {e}", important=True)

                # Note: CBSL Provision Comparison file is kept for download
                _emit_log("CBSL Provision Comparison file kept for download", important=True)
            else:
                _emit_log("Unexpected working folder structure for GA & IS; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"GA & IS post-processing failed: {e}", important=True)
    elif is_c8_flow:
        # Post-processing for C8: rename file, queue download, archive working folder
        try:
            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date
            # working date folder
            c8_work_dir = WORKING_DIR / "NBD_QF_23_C8"
            date_folders = [p for p in c8_work_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                # Rename C8 file: C8-Client Rating-Jun-2025.xlsx
                matches = glob.glob(str(date_folder / "C8-Client Rating*.xlsx"))
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    new_name = f"C8-Client Rating-{gui_month}-{gui_year}.xlsx"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            src = dst
                            _emit_log(f"Renamed C8 file to {dst.name}", important=True)
                        except Exception as e:
                            _emit_log(f"C8 rename failed (continuing): {e}", important=True)
                    _download_files.append(src)
                else:
                    _emit_log("C8 .xlsx file not found for renaming", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                try:
                    dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                except Exception:
                    try:
                        dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                    except Exception:
                        dt = None
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-QF-23-C8", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved C8: {msg}", important=True)
                    _last_output_dirs["NBD-QF-23-C8"] = output_path

                    # Update file paths in _download_files to point to the new output location
                    updated_download_files = []
                    for file_path in _download_files:
                        if file_path.exists():
                            # File still exists at original location, keep it
                            updated_download_files.append(file_path)
                        else:
                            # File was moved, try to find it in the new output location
                            new_path = output_path / file_path.name
                            if new_path.exists():
                                updated_download_files.append(new_path)
                                _emit_log(f"Updated file path: {file_path.name} -> {new_path}", important=True)
                            else:
                                _emit_log(f"Could not find moved file: {file_path.name}", important=True)

                    _download_files = updated_download_files
                    _emit_log(f"Updated download files after move: {len(_download_files)} files", important=True)

                # Prepare downloads
                _set_status(stage="saving", message="Preparing files for download")
                try:
                    auto_download_files()
                except Exception as e:
                    _emit_log(f"Warning: C8 download prep failed: {e}", important=True)
            else:
                _emit_log("Unexpected working folder structure for C8; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"C8 post-processing failed: {e}", important=True)
    elif is_c7_flow:
        # Post-processing for C7: calculate quarter, rename file, queue download, archive working folder
        try:
            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date

            # Calculate quarter and financial year from GUI date
            # Financial Year: April - March
            # Q1 (Apr-Jun), Q2 (Jul-Sep), Q3 (Oct-Dec), Q4 (Jan-Mar)
            try:
                dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
            except Exception:
                try:
                    dt = datetime.strptime(gui_date_str, "%d-%m-%Y")
                except Exception:
                    dt = None

            if dt is not None:
                month = dt.month
                year = dt.year

                # Determine quarter and financial year
                if month >= 4:  # Apr-Dec of same year
                    fy_year = year
                    if month in [4, 5, 6]:
                        quarter = 1  # Q1: Apr-Jun
                    elif month in [7, 8, 9]:
                        quarter = 2  # Q2: Jul-Sep
                    else:  # month in [10, 11, 12]
                        quarter = 3  # Q3: Oct-Dec
                else:  # Jan-Mar of next year
                    fy_year = year - 1
                    quarter = 4  # Q4: Jan-Mar

                _emit_log(f"Calculated Quarter {quarter}, Financial Year {fy_year} from date {gui_date_str}", important=True)
            else:
                _emit_log("Could not parse date for quarter calculation", important=True)
                quarter = 1
                fy_year = year

            # working date folder
            c7_work_dir = WORKING_DIR / "NBD_QF_23_C7"
            date_folders = [p for p in c7_work_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                # Rename C7 file: C7- Report for Stage - Movement Q{quarter} FY{fy_year}.xlsb
                matches = glob.glob(str(date_folder / "*.xlsb"))
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    new_name = f"C7- Report for Stage - Movement Q{quarter} FY{fy_year}.xlsb"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            src = dst
                            _emit_log(f"Renamed C7 file to {dst.name}", important=True)
                        except Exception as e:
                            _emit_log(f"C7 rename failed (continuing): {e}", important=True)
                    _download_files.append(src)
                else:
                    _emit_log("C7 .xlsb file not found for renaming", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-QF-23-C7", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved C7: {msg}", important=True)
                    _last_output_dirs["NBD-QF-23-C7"] = output_path

                    # Update file paths in _download_files to point to the new output location
                    updated_download_files = []
                    for file_path in _download_files:
                        if file_path.exists():
                            # File still exists at original location, keep it
                            updated_download_files.append(file_path)
                        else:
                            # File was moved, try to find it in the new output location
                            new_path = output_path / file_path.name
                            if new_path.exists():
                                updated_download_files.append(new_path)
                                _emit_log(f"Updated file path: {file_path.name} -> {new_path}", important=True)
                            else:
                                _emit_log(f"Could not find moved file: {file_path.name}", important=True)

                    _download_files = updated_download_files
                    _emit_log(f"Updated download files after move: {len(_download_files)} files", important=True)

                # Prepare downloads
                _set_status(stage="saving", message="Preparing files for download")
                try:
                    auto_download_files()
                except Exception as e:
                    _emit_log(f"Warning: C7 download prep failed: {e}", important=True)
            else:
                _emit_log("Unexpected working folder structure for C7; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"C7 post-processing failed: {e}", important=True)
    elif is_c4_flow:
        # Post-processing for C4: rename file, queue download, archive working folder
        try:
            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date
            # working date folder
            c4_work_dir = WORKING_DIR / "NBD_MF_23_C4"
            date_folders = [p for p in c4_work_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                # Rename C4 file: Loan CF Analysis NBD-MF-23-C4-Jun-2025.xlsb
                matches = glob.glob(str(date_folder / "Loan CF Analysis*.xlsb"))
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    new_name = f"Loan CF Analysis NBD-MF-23-C4-{gui_month}-{gui_year}.xlsb"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            src = dst
                            _emit_log(f"Renamed C4 file to {dst.name}", important=True)
                        except Exception as e:
                            _emit_log(f"C4 rename failed (continuing): {e}", important=True)
                    _download_files.append(src)
                else:
                    _emit_log("C4 .xlsb file not found for renaming", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                try:
                    dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                except Exception:
                    try:
                        dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                    except Exception:
                        dt = None
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-MF-23-C4", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved C4: {msg}", important=True)
                    _last_output_dirs["NBD-MF-23-C4"] = output_path

                    # Update file paths in _download_files to point to the new output location
                    updated_download_files = []
                    for file_path in _download_files:
                        if file_path.exists():
                            # File still exists at original location, keep it
                            updated_download_files.append(file_path)
                        else:
                            # File was moved, try to find it in the new output location
                            new_path = output_path / file_path.name
                            if new_path.exists():
                                updated_download_files.append(new_path)
                                _emit_log(f"Updated file path: {file_path.name} -> {new_path}", important=True)
                            else:
                                _emit_log(f"Could not find moved file: {file_path.name}", important=True)

                    _download_files = updated_download_files
                    _emit_log(f"Updated download files after move: {len(_download_files)} files", important=True)
            else:
                _emit_log("Unexpected working folder structure for C4; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"C4 post-processing failed: {e}", important=True)
    elif is_dm_flow:
        # Post-processing for DM: rename file with week number, queue download, archive working folder
        try:
            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date

            # Calculate week number from the date
            try:
                dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
            except Exception:
                try:
                    dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                except Exception:
                    dt = datetime.now()

            # Get ISO week number
            week_number = dt.isocalendar()[1]

            # working date folder
            dm_work_dir = WORKING_DIR / "NBD-WF-18-DM"
            date_folders = [p for p in dm_work_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                _emit_log(f"Looking for DM files in: {date_folder}", important=True)
                
                # List all files in the directory for debugging
                all_files = list(date_folder.glob("*"))
                _emit_log(f"Files found in directory: {[f.name for f in all_files]}", important=True)
                
                # Find DM file: NBD-WF-18-DM Deposit Liability Week*.xlsx
                matches = glob.glob(str(date_folder / "NBD-WF-18-DM Deposit Liability Week*.xlsx")) + \
                         glob.glob(str(date_folder / "NBD-WF-18-DM*.xlsx")) + \
                         glob.glob(str(date_folder / "NBD*WF*18*DM*.xlsx"))
                _emit_log(f"Pattern matches found: {matches}", important=True)
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    new_name = f"NBD-WF-18-DM Deposit Liability Week {week_number}.xlsx"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            src = dst
                            _emit_log(f"Renamed DM file to {dst.name}", important=True)
                        except Exception as e:
                            _emit_log(f"DM rename failed (continuing): {e}", important=True)
                    _download_files.append(src)
                    _emit_log(f"Added main DM file to download queue: {src.name}", important=True)
                else:
                    _emit_log("DM .xlsx file not found for renaming", important=True)
                    # Try to find any Excel file that might be the DM file
                    excel_files = list(date_folder.glob("*.xlsx"))
                    if excel_files:
                        _emit_log(f"Found Excel files but none matched pattern: {[f.name for f in excel_files]}", important=True)
                        # Use the most recent Excel file as fallback
                        src = max(excel_files, key=lambda x: x.stat().st_mtime)
                        _download_files.append(src)
                        _emit_log(f"Added fallback Excel file to download queue: {src.name}", important=True)

                # Add exceptions file if it exists
                exceptions_file = date_folder / "NBD_WF_18_DM_Detailed_Exceptions.txt"
                if exceptions_file.exists():
                    _download_files.append(exceptions_file)
                    _emit_log(f"Added exceptions file to download: {exceptions_file.name}", important=True)
                else:
                    _emit_log("No exceptions file found for DM", important=True)
                
                _emit_log(f"Total files in download queue: {len(_download_files)}", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                try:
                    dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                except Exception:
                    try:
                        dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                    except Exception:
                        dt = None
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-WF-18-DM", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved DM: {msg}", important=True)
                    _last_output_dirs["NBD-WF-18-DM"] = output_path
                    
                    # Update file paths in _download_files to point to the new output location
                    updated_download_files = []
                    for file_path in _download_files:
                        if file_path.exists():
                            # File still exists at original location, keep it
                            updated_download_files.append(file_path)
                        else:
                            # File was moved, try to find it in the new output location
                            new_path = output_path / file_path.name
                            if new_path.exists():
                                updated_download_files.append(new_path)
                                _emit_log(f"Updated file path: {file_path.name} -> {new_path}", important=True)
                            else:
                                _emit_log(f"Could not find moved file: {file_path.name}", important=True)
                    
                    _download_files = updated_download_files
                    _emit_log(f"Updated download files after move: {len(_download_files)} files", important=True)

                # Copy ECS folder's Deposit Liability file into the DM output folder
                try:
                    ecs_outputs_dir = OUTPUTS_DIR / "NBD_WF_18_ECS"
                    _, ecs_date_folder = find_latest_completed_date(ecs_outputs_dir)
                    if ecs_date_folder and ecs_date_folder.exists():
                        _emit_log(f"ECS latest folder: {ecs_date_folder}", important=True)
                        # Find the Deposit Liability file in ECS folder
                        dep_matches = glob.glob(str(ecs_date_folder / "NBD-WF-18-DM Deposit Liability*.xlsx"))
                        if dep_matches:
                            ecs_dep_src = Path(dep_matches[0])
                            # Determine DM output folder
                            dm_output_folder = _last_output_dirs.get("NBD-WF-18-DM")
                            if dm_output_folder and dm_output_folder.exists():
                                # Remove any existing Deposit Liability file in DM output
                                for old_file in dm_output_folder.glob("NBD-WF-18-DM Deposit Liability*.xlsx"):
                                    old_file.unlink()
                                    _emit_log(f"Removed existing: {old_file.name}", important=True)
                                # Build new name using picked date DD-MM-YYYY
                                try:
                                    _dt_picked = datetime.strptime(gui_date_str, "%Y-%m-%d")
                                except Exception:
                                    try:
                                        _dt_picked = datetime.strptime(gui_date_str, "%m/%d/%Y")
                                    except Exception:
                                        _dt_picked = datetime.now()
                                picked_date_str = _dt_picked.strftime("%d-%m-%Y")
                                new_dep_name = f"NBD-WF-18-DM Deposit Liability ({picked_date_str}).xlsx"
                                dst_dep = dm_output_folder / new_dep_name
                                shutil.copy2(str(ecs_dep_src), str(dst_dep))
                                _emit_log(f"Copied ECS Deposit Liability as: {new_dep_name}", important=True)
                                _download_files.append(dst_dep)
                            else:
                                _emit_log("DM output folder not available for ECS copy", important=True)
                        else:
                            _emit_log(f"No 'NBD-WF-18-DM Deposit Liability*.xlsx' found in ECS folder: {ecs_date_folder}", important=True)
                    else:
                        _emit_log("ECS latest completed date folder not found; skipping ECS copy", important=True)
                except Exception as e:
                    _emit_log(f"ECS Deposit Liability copy failed (continuing): {e}", important=True)

                # Prepare downloads
                _set_status(stage="saving", message="Preparing files for download")
                try:
                    auto_download_files()
                except Exception as e:
                    _emit_log(f"Warning: DM download prep failed: {e}", important=True)
            else:
                _emit_log("Unexpected working folder structure for DM; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"DM post-processing failed: {e}", important=True)
    elif is_la_flow:
        # Post-processing for LA: rename AFL file with week number, queue download, archive working folder
        try:
            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date

            # Calculate week number from the date
            try:
                dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
            except Exception:
                try:
                    dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                except Exception:
                    dt = datetime.now()

            # Get ISO week number
            week_number = dt.isocalendar()[1]

            # working date folder
            la_work_dir = WORKING_DIR / "NBD_WF_15_LA"
            date_folders = [p for p in la_work_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                _emit_log(f"Looking for LA files in: {date_folder}", important=True)

                # List all files in the directory for debugging
                all_files = list(date_folder.glob("*"))
                _emit_log(f"Files found in directory: {[f.name for f in all_files]}", important=True)

                # Find AFL file: AFL Liquidity - Week_*.xlsx
                matches = list(date_folder.glob("AFL Liquidity - Week_*.xlsx"))
                _emit_log(f"AFL files found: {[f.name for f in matches]}", important=True)
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))

                    if _la_month_end:
                        # Month-end: rename to "AFL Liquidity - Mon-YYYY.xlsx"
                        new_name = f"AFL Liquidity - {dt.strftime('%b')}-{dt.year}.xlsx"
                    else:
                        new_name = f"AFL Liquidity - Week_{week_number}.xlsx"

                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            src = dst
                            _emit_log(f"Renamed AFL file to {dst.name}", important=True)
                        except Exception as e:
                            _emit_log(f"AFL rename failed (continuing): {e}", important=True)
                    _download_files.append(src)
                    _emit_log(f"Added AFL file to download queue: {src.name}", important=True)

                    # Remove all other AFL files, keep only the latest one
                    for old_afl in date_folder.glob("AFL Liquidity - Week_*.xlsx"):
                        if old_afl != src:
                            try:
                                old_afl.unlink()
                                _emit_log(f"Removed old AFL file: {old_afl.name}", important=True)
                            except Exception as e:
                                _emit_log(f"Failed to remove old AFL file {old_afl.name}: {e}", important=True)
                else:
                    _emit_log("AFL .xlsx file not found", important=True)

                # Check for ALCL Management Accounts file and add to download queue
                alcl_files = list(date_folder.glob("ALCL Management Accounts*.xlsx"))
                if alcl_files:
                    # Get the most recently modified ALCL file
                    latest_alcl = max(alcl_files, key=lambda x: x.stat().st_mtime)
                    _download_files.append(latest_alcl)
                    _emit_log(f"Added ALCL Management Accounts file to download queue: {latest_alcl.name}", important=True)
                else:
                    _emit_log("ALCL Management Accounts file not found in working folder", important=True)

                # Check for BankBalanceGLCodes file and add to download queue
                bank_balance_file = date_folder / "BankBalanceGLCodes.xlsx"
                if bank_balance_file.exists():
                    _download_files.append(bank_balance_file)
                    _emit_log(f"Added BankBalanceGLCodes file to download queue: {bank_balance_file.name}", important=True)
                else:
                    _emit_log("BankBalanceGLCodes.xlsx file not found in working folder", important=True)

                _emit_log(f"Total files in download queue: {len(_download_files)}", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                try:
                    dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                except Exception:
                    try:
                        dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                    except Exception:
                        dt = None
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-WF-15-LA", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved LA: {msg}", important=True)
                    _last_output_dirs["NBD-WF-15-LA"] = output_path

                    # NOTE: AFL file copy to NBD_MF_04_LA is handled by mark_completed() function
                    # when user manually marks the LA report as completed, not automatically here

                    # Update file paths in _download_files to point to the new output location
                    updated_download_files = []
                    for file_path in _download_files:
                        if file_path.exists():
                            # File still exists at original location, keep it
                            updated_download_files.append(file_path)
                        else:
                            # File was moved, try to find it in the new output location
                            new_path = output_path / file_path.name
                            if new_path.exists():
                                updated_download_files.append(new_path)
                                _emit_log(f"Updated file path: {file_path.name} -> {new_path}", important=True)
                            else:
                                _emit_log(f"Could not find moved file: {file_path.name}", important=True)

                    _download_files = updated_download_files
                    _emit_log(f"Updated download files after move: {len(_download_files)} files", important=True)

                # Prepare downloads
                _set_status(stage="saving", message="Preparing files for download")
                try:
                    auto_download_files()
                except Exception as e:
                    _emit_log(f"Warning: LA download prep failed: {e}", important=True)
            else:
                _emit_log("Unexpected working folder structure for LA; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"LA post-processing failed: {e}", important=True)

        # Clean up working directories to prevent "multiple date folders" errors
        _emit_log("Cleaning up working directories before starting automation", important=True)
        cleanup_all_working_directories()
    elif is_ecs_flow:
        # Post-processing for ECS: rename Cash-Flow-Forecast file with week number, queue download, archive working folder
        try:
            # Clear download files to ensure only ECS file is downloaded
            _download_files = []
            _emit_log("Cleared download files for ECS flow", important=True)

            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date

            # Calculate ISO week number of the year (1-52/53) for file renaming
            try:
                dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
            except Exception:
                try:
                    dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                except Exception:
                    dt = datetime.now()

            # Get ISO week number of the year (separate calculation for renaming only)
            week_number = dt.isocalendar()[1]
            _emit_log(f"Calculated ISO week number for renaming: {week_number} (date: {dt.strftime('%Y-%m-%d')})", important=True)

            # working date folder (created by ECS script after bots complete)
            ecs_work_dir = WORKING_DIR / "NBD_WF_18_ECS"
            date_folders = [p for p in ecs_work_dir.iterdir() if p.is_dir()] if ecs_work_dir.exists() else []
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                _emit_log(f"Looking for ECS files in: {date_folder}", important=True)

                # List all files in the directory for debugging
                all_files = list(date_folder.glob("*"))
                _emit_log(f"Files found in directory: {[f.name for f in all_files]}", important=True)

                # Find Cash-Flow-Forecast file
                matches = list(date_folder.glob("Cash-Flow-Forecast*.xlsm"))
                _emit_log(f"Cash-Flow-Forecast files found: {[f.name for f in matches]}", important=True)
                if matches:
                    src = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                    new_name = f"Cash-Flow-Forecast-Week {week_number}.xlsm"
                    dst = src.with_name(new_name)
                    if src.name != new_name:
                        try:
                            src.rename(dst)
                            src = dst
                            _emit_log(f"Renamed ECS file to {dst.name}", important=True)
                        except Exception as e:
                            _emit_log(f"ECS rename failed (continuing): {e}", important=True)
                    _download_files.append(src)
                    _emit_log(f"Added ECS file to download queue: {src.name}", important=True)
                else:
                    _emit_log("Cash-Flow-Forecast .xlsm file not found", important=True)

                _emit_log(f"Total files in download queue: {len(_download_files)}", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                try:
                    dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                except Exception:
                    try:
                        dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                    except Exception:
                        dt = None
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)

                            # Update the file path in _download_files after folder rename
                            if _download_files:
                                old_file_path = _download_files[0]
                                new_file_path = target / old_file_path.name
                                _download_files[0] = new_file_path
                                _emit_log(f"Updated download file path after folder rename: {new_file_path}", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-WF-18-ECS", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved ECS: {msg}", important=True)
                    _last_output_dirs["NBD-WF-18-ECS"] = output_path

                    # Keep the working folder file path for download (don't update to outputs)
                    # The file remains in working folder until next run cleanup
                    _emit_log(f"Download will use working folder file: {_download_files[0].name if _download_files else 'none'}", important=True)

                # Export "Types of Receipt" sheet as a separate xlsx file
                types_of_receipt_path = None
                if _download_files:
                    ecs_xlsm = _download_files[0]
                    if ecs_xlsm.exists():
                        try:
                            import openpyxl as _opx
                            wb_src = _opx.load_workbook(str(ecs_xlsm), keep_vba=False, data_only=True)
                            if "Types of Receipt" in wb_src.sheetnames:
                                wb_tor = _opx.Workbook()
                                ws_tor_dst = wb_tor.active
                                ws_tor_dst.title = "Types of Receipt"
                                ws_tor_src = wb_src["Types of Receipt"]
                                for row in ws_tor_src.iter_rows(values_only=True):
                                    ws_tor_dst.append(list(row))
                                tor_path = ecs_xlsm.parent / "Types_of_Receipt.xlsx"
                                wb_tor.save(str(tor_path))
                                wb_tor.close()
                                wb_src.close()
                                types_of_receipt_path = tor_path
                                _emit_log(f"Exported Types of Receipt sheet: {tor_path.name}", important=True)
                            else:
                                wb_src.close()
                                _emit_log("'Types of Receipt' sheet not found in ECS file — skipping export", important=True)
                        except Exception as e:
                            _emit_log(f"Types of Receipt export failed: {e}", important=True)

                # Export "TAX Payments" sheet A1:D15 as a separate xlsx file
                tax_payments_export_path = None
                if _download_files:
                    ecs_xlsm = _download_files[0]
                    if ecs_xlsm.exists():
                        try:
                            import openpyxl as _opx2
                            wb_tp_src = _opx2.load_workbook(str(ecs_xlsm), keep_vba=False, data_only=True)
                            tp_sheet_name = None
                            for _cand in ("TAX Payments", "Tax Payments"):
                                if _cand in wb_tp_src.sheetnames:
                                    tp_sheet_name = _cand
                                    break
                            if tp_sheet_name:
                                wb_tp = _opx2.Workbook()
                                ws_tp_dst = wb_tp.active
                                ws_tp_dst.title = "Tax Payments"
                                ws_tp_src = wb_tp_src[tp_sheet_name]
                                for row in ws_tp_src.iter_rows(min_row=1, max_row=15, min_col=1, max_col=4, values_only=True):
                                    ws_tp_dst.append(list(row))
                                today_str = datetime.today().strftime("%d-%m-%Y")
                                tp_export_name = f"Tax Payments - {today_str}.xlsx"
                                tp_export_path = ecs_xlsm.parent / tp_export_name
                                wb_tp.save(str(tp_export_path))
                                wb_tp.close()
                                wb_tp_src.close()
                                tax_payments_export_path = tp_export_path
                                _emit_log(f"Exported Tax Payments sheet: {tp_export_path.name}", important=True)
                                # Also copy to ECS outputs folder
                                try:
                                    ecs_out_dir = OUTPUTS_DIR / "NBD_WF_18_ECS"
                                    _, out_folder = find_latest_completed_date(ecs_out_dir)
                                    if out_folder and out_folder.exists():
                                        for _old_tp in out_folder.glob("Tax Payments - *.xlsx"):
                                            _old_tp.unlink(missing_ok=True)
                                        shutil.copy2(str(tp_export_path), str(out_folder / tp_export_name))
                                        _emit_log(f"Copied Tax Payments to ECS outputs: {tp_export_name}", important=True)
                                except Exception as _ce:
                                    _emit_log(f"Tax Payments copy to outputs failed: {_ce}", important=True)
                            else:
                                wb_tp_src.close()
                                _emit_log("'TAX Payments' / 'Tax Payments' sheet not found in ECS file — skipping export", important=True)
                        except Exception as e:
                            _emit_log(f"Tax Payments export failed: {e}", important=True)

                # Export "Cashflow Movement" sheet A:B as a separate xlsx file
                cashflow_movement_export_path = None
                if _download_files:
                    ecs_xlsm = _download_files[0]
                    if ecs_xlsm.exists():
                        try:
                            import openpyxl as _opx3
                            wb_cm_src = _opx3.load_workbook(str(ecs_xlsm), keep_vba=False, data_only=True)
                            if "Cashflow Movement" in wb_cm_src.sheetnames:
                                ws_cm_src = wb_cm_src["Cashflow Movement"]
                                wb_cm_out = _opx3.Workbook()
                                ws_cm_out = wb_cm_out.active
                                ws_cm_out.title = "Cashflow Movement"
                                for row in ws_cm_src.iter_rows(values_only=True):
                                    ws_cm_out.append(list(row))
                                now_str = datetime.now().strftime("%d-%m-%Y %H-%M-%S")
                                cm_export_name = f"Cashflow Movement - {now_str}.xlsx"
                                cm_export_path = ecs_xlsm.parent / cm_export_name
                                wb_cm_out.save(str(cm_export_path))
                                wb_cm_out.close()
                                wb_cm_src.close()
                                cashflow_movement_export_path = cm_export_path
                                _emit_log(f"Exported Cashflow Movement sheet: {cm_export_name}", important=True)
                                # Copy to ECS outputs folder
                                try:
                                    ecs_out_dir = OUTPUTS_DIR / "NBD_WF_18_ECS"
                                    _, out_folder = find_latest_completed_date(ecs_out_dir)
                                    if out_folder and out_folder.exists():
                                        for _old_cm in out_folder.glob("Cashflow Movement - *.xlsx"):
                                            _old_cm.unlink(missing_ok=True)
                                        shutil.copy2(str(cm_export_path), str(out_folder / cm_export_name))
                                        _emit_log(f"Copied Cashflow Movement to ECS outputs: {cm_export_name}", important=True)
                                except Exception as _ce:
                                    _emit_log(f"Cashflow Movement copy to outputs failed: {_ce}", important=True)
                            else:
                                wb_cm_src.close()
                                _emit_log("'Cashflow Movement' sheet not found in ECS file — skipping export", important=True)
                        except Exception as e:
                            _emit_log(f"Cashflow Movement export failed: {e}", important=True)

                # Zip ECS xlsm + Types_of_Receipt.xlsx + Tax Payments + Cashflow Movement together
                zip_extras = [p for p in [types_of_receipt_path, tax_payments_export_path, cashflow_movement_export_path] if p and p.exists()]
                if zip_extras:
                    try:
                        temp_zip_dir = BASE_DIR / "temp_zips"
                        temp_zip_dir.mkdir(parents=True, exist_ok=True)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        zip_path = temp_zip_dir / f"ECS_Week{week_number}_{timestamp}.zip"
                        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                            zipf.write(str(_download_files[0]), arcname=_download_files[0].name)
                            for _extra in zip_extras:
                                zipf.write(str(_extra), arcname=_extra.name)
                        _download_files = [zip_path]
                        _emit_log(f"Created ECS zip: {zip_path.name}", important=True)
                    except Exception as e:
                        _emit_log(f"ECS zip creation failed: {e}", important=True)

                _set_status(stage="saving", message="ECS file ready for download")
                _emit_log(f"ECS download ready: {_download_files[0].name if _download_files else 'none'}", important=True)
            else:
                _emit_log("Unexpected working folder structure for ECS; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"ECS post-processing failed: {e}", important=True)
    elif is_mf04_flow:
        # Post-processing for MF-04-LA: rename output, save to outputs, queue download
        try:
            _download_files = []
            _emit_log("Starting MF-04-LA post-processing", important=True)

            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date

            # Working date folder
            mf04_work_dir = WORKING_DIR / "NBD_MF_04_LA"
            date_folders = [p for p in mf04_work_dir.iterdir() if p.is_dir()]
            if len(date_folders) == 1:
                date_folder = date_folders[0]
                _emit_log(f"Looking for MF-04-LA files in: {date_folder}", important=True)

                # Find the monthly report output file
                monthly_files = list(date_folder.glob("NBD-MF-04-LA Liquid Assets*.xlsm"))
                if monthly_files:
                    latest_monthly = max(monthly_files, key=lambda x: x.stat().st_mtime)
                    _download_files.append(latest_monthly)
                    _emit_log(f"Added monthly report file to download queue: {latest_monthly.name}", important=True)
                else:
                    _emit_log("MF-04-LA output file not found in working folder", important=True)

                _emit_log(f"Total files in download queue: {len(_download_files)}", important=True)

                # Rename working folder to picked date format DD-MM-YYYY
                try:
                    dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                except Exception:
                    try:
                        dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                    except Exception:
                        dt = None
                if dt is not None:
                    picked_name = dt.strftime("%d-%m-%Y")
                    if date_folder.name != picked_name:
                        target = date_folder.parent / picked_name
                        if target.exists():
                            shutil.rmtree(target)
                        try:
                            date_folder.rename(target)
                            date_folder = target
                            _emit_log(f"Renamed working folder to {picked_name}", important=True)

                            # Update download file paths after folder rename
                            if _download_files:
                                updated = []
                                for fp in _download_files:
                                    new_fp = target / fp.name
                                    updated.append(new_fp)
                                _download_files = updated
                                _emit_log(f"Updated download file paths after folder rename", important=True)
                        except Exception as e:
                            _emit_log(f"Failed to rename working folder: {e}", important=True)

                # Save to outputs with versioning using selected date
                ok, msg, output_path = save_folder_to_outputs("NBD-MF-04-LA", gui_date_str)
                if ok and output_path:
                    _emit_log(f"Saved MF-04-LA: {msg}", important=True)
                    _last_output_dirs["NBD-MF-04-LA"] = output_path

                # Prepare downloads
                _set_status(stage="saving", message="Preparing MF-04-LA files for download")
                try:
                    auto_download_files()
                except Exception as e:
                    _emit_log(f"Warning: MF-04-LA download prep failed: {e}", important=True)
            else:
                _emit_log("Unexpected working folder structure for MF-04-LA; skipping post-processing", important=True)
        except Exception as e:
            _emit_log(f"MF-04-LA post-processing failed: {e}", important=True)
    elif is_ia_flow:
        # Post-processing for IA flow: handle both IA and C1-C2 reports
        try:
            # Determine GUI date
            any_date = list(report_ids_with_dates.values())[0]
            gui_date_str, gui_month, gui_year = any_date
            
            # Handle IA report
            # IA and C1C2 always run together, so IA will save to outputs and C1C2 will also save
            # Both will save their respective outputs

            if "NBD-MF-23-IA" in report_ids_with_dates:
                ia_work_dir = WORKING_DIR / "NBD_MF_23_IA"
                date_folders = [p for p in ia_work_dir.iterdir() if p.is_dir()]
                if len(date_folders) == 1:
                    date_folder = date_folders[0]

                    # Find and rename IA output file
                    ia_matches = glob.glob(str(date_folder / "Prod. wise Class. of Loans*.xlsb"))
                    if ia_matches:
                        src = Path(max(ia_matches, key=lambda x: Path(x).stat().st_mtime))
                        new_name = f"Prod. wise Class. of Loans - {gui_month} {gui_year}.xlsb"
                        dst = src.with_name(new_name)
                        if src.name != new_name:
                            try:
                                src.rename(dst)
                                src = dst
                                _emit_log(f"Renamed IA file to {dst.name}", important=True)
                            except Exception as e:
                                _emit_log(f"IA rename failed (continuing): {e}", important=True)

                    # Rename working folder to picked date format DD-MM-YYYY
                    try:
                        dt = datetime.strptime(gui_date_str, "%Y-%m-%d")
                    except Exception:
                        try:
                            dt = datetime.strptime(gui_date_str, "%m/%d/%Y")
                        except Exception:
                            dt = None
                    if dt is not None:
                        picked_name = dt.strftime("%d-%m-%Y")
                        if date_folder.name != picked_name:
                            target = date_folder.parent / picked_name
                            if target.exists():
                                shutil.rmtree(target)
                            try:
                                date_folder.rename(target)
                                date_folder = target
                                _emit_log(f"Renamed working folder to {picked_name}", important=True)
                            except Exception as e:
                                _emit_log(f"Failed to rename working folder: {e}", important=True)

                    # IA files stay in working folder for now
                    # They will be saved to outputs after C1-C2 completes
                    _emit_log("IA completed - files ready in working folder", important=True)

            # Handle C1-C2 report (if selected)
            if "NBD-MF-23-C1-C2" in report_ids_with_dates:
                # C1-C2 uses the same working folder as IA
                ia_work_dir = WORKING_DIR / "NBD_MF_23_IA"
                date_folders = [p for p in ia_work_dir.iterdir() if p.is_dir()]
                if len(date_folders) == 1:
                    date_folder = date_folders[0]

                    # Find the updated C1-C2 file (already renamed by IA section)
                    c1c2_file_path = date_folder / f"Prod. wise Class. of Loans - {gui_month} {gui_year}.xlsb"
                    if c1c2_file_path.exists():
                        # Add to download queue
                        _download_files.append(c1c2_file_path)
                        _emit_log(f"Added C1-C2 file to download queue: {c1c2_file_path.name}", important=True)

                        # Copy updated C1-C2 file from working folder to C1-C6, C8, and C3-C10/Input
                        _emit_log(f"Copying updated C1-C2 file to additional output folders...", important=True)
                        copy_ia_to_additional_outputs(c1c2_file_path, gui_date_str)
                    else:
                        _emit_log(f"Warning: C1-C2 file not found in working folder", important=True)

                    # Find and add EXCEPTIONS.xlsx if it exists
                    exceptions_file = date_folder / "EXCEPTIONS.xlsx"
                    if exceptions_file.exists():
                        _download_files.append(exceptions_file)
                        _emit_log(f"Added EXCEPTIONS.xlsx to download queue", important=True)

                    # Find and add Minimum_Rate_Exceptions.xlsx if it exists
                    min_rate_exceptions = date_folder / "Minimum_Rate_Exceptions.xlsx"
                    if min_rate_exceptions.exists():
                        _download_files.append(min_rate_exceptions)
                        _emit_log(f"Added Minimum_Rate_Exceptions.xlsx to download queue", important=True)

                    # Find and add initial_valuation.xlsx if it exists
                    iv_file = date_folder / "initial_valuation.xlsx"
                    if iv_file.exists():
                        _download_files.append(iv_file)
                        _emit_log(f"Added initial_valuation.xlsx to download queue", important=True)

                    # Find and add Product_Cat.xlsx if it exists
                    product_cat_file = date_folder / "Product_Cat.xlsx"
                    if product_cat_file.exists():
                        _download_files.append(product_cat_file)
                        _emit_log(f"Added Product_Cat.xlsx to download queue", important=True)

                    # Save working folder to IA outputs (not C1-C2 outputs)
                    ok, msg, output_path = save_folder_to_outputs("NBD-MF-23-IA", gui_date_str)
                    if ok and output_path:
                        _emit_log(f"Saved IA/C1-C2 working folder to IA outputs: {msg}", important=True)
                        _last_output_dirs["NBD-MF-23-IA"] = output_path
                        _last_output_dirs["NBD-MF-23-C1-C2"] = output_path
                    else:
                        _emit_log(f"Warning: Could not save working folder to outputs: {msg}", important=True)

                    # C1C2 uses Ctrl+S method - files stay in working location for download
                    _emit_log("C1C2 completed - files ready for download from working location", important=True)

            # Prepare downloads
            _set_status(stage="saving", message="Preparing files for download")
            try:
                auto_download_files()
            except Exception as e:
                _emit_log(f"Warning: IA flow download prep failed: {e}", important=True)
                
        except Exception as e:
            _emit_log(f"IA flow post-processing failed: {e}", important=True)
    else:
        # Save completed reports back to outputs with selected dates (IA/SOFP flow)
        _set_status(stage="saving", message="Saving reports to outputs")
        for report_id in report_ids:
            # Skip saving for C1/C2 as they don't have their own working folder
            if report_id in ("NBD-MF-23-C1", "NBD-MF-23-C2"):
                continue
            # QF reports are saved together after process_summary.py completes
            if report_id in ("NBD-QF-23-C3", "NBD-QF-23-C4", "NBD-QF-23-C6",
                           "NBD-QF-23-C9", "NBD-QF-23-C10",
                           "NBD-QF-16-SF", "NBD-QF-16-FI"):
                _emit_log(f"Skipping individual save for {report_id} (will be saved after process_summary.py)", important=True)
                continue
            if _stop_requested:
                break

            date_str, _, _ = report_ids_with_dates[report_id]
            success, msg, output_path = save_folder_to_outputs(report_id, date_str)
            if success and output_path:
                _emit_log(f"Saved {report_id}: {msg}", important=True)
                _last_output_dirs[report_id] = output_path
                # For SOFP-only flow, queue the file from working directory for download
                if only_sofp_flow and report_id in ("NBD-MF-01-SOFP-SOCI", "NBD-AF-01-SOFP-SOCI") and _sofp_download_file:
                    # Download from working directory where the file was renamed
                    if _sofp_download_file.exists():
                        _download_files = [_sofp_download_file]
                        _emit_log(f"Queued SOFP file from working directory for download: {_sofp_download_file}", important=True)

                        # Also include ifrs_code.xlsx if it exists in the same folder
                        ifrs_path = _sofp_download_file.parent / "ifrs_code.xlsx"
                        if ifrs_path.exists():
                            _download_files.append(ifrs_path)
                            _emit_log(f"Added IFRS code file for download: {ifrs_path.name}", important=True)
                        else:
                            _emit_log("IFRS code file not found in SOFP working directory", important=True)
    
    if _stop_requested:
        return False, "Process stopped by user"
    
    if all_success:
        # Outputs have already been saved explicitly per-flow; avoid duplicating version folders
        # Note: This section runs AFTER:
        #   1. Master data collection is completed
        #   2. Copy outputs to working is done
        #   3. All QF report scripts have completed successfully

        # Run process_summary.py for QF reports group
        # This runs AFTER all QF reports complete and BEFORE downloads
        if is_qf_reports_flow:
            try:
                # Get the working directory for QF reports
                qf_working_dir = WORKING_DIR / "NBD-QF-23-C3-C10-SF-FI"
                if not qf_working_dir.exists():
                    _emit_log(f"QF working directory not found: {qf_working_dir}", important=True)
                else:
                    date_folders = [p for p in qf_working_dir.iterdir() if p.is_dir()]
                    if len(date_folders) == 1:
                        qf_date_folder = date_folders[0]
                        _emit_log(f"QF working date folder: {qf_date_folder}", important=True)

                        # Locate QF files for download
                        _emit_log("Locating Quarter Classification report for download", important=True)

                        # Calculate quarter month name
                        if report_ids_with_dates:
                            any_date = list(report_ids_with_dates.values())[0]
                            _, gui_month, gui_year = any_date
                            month_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                                       "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
                            if gui_month in month_map:
                                current_month_date = datetime(int(gui_year), month_map[gui_month], 1)
                                current_q_monthname = current_month_date.strftime("%b %Y")
                            else:
                                current_q_monthname = f"{gui_month} {gui_year}"
                        else:
                            current_month_date = datetime.now().replace(day=1)
                            current_q_monthname = current_month_date.strftime("%b %Y")

                        # Find Quarter Classification file in working directory
                        pattern = f"Quarter Classification Loans & Advance-{current_q_monthname}*.xls*"
                        matches = glob.glob(str(qf_date_folder / pattern))
                        if not matches:
                            matches = glob.glob(str(qf_date_folder / "Quarter Classification Loans & Advance*.xls*"))

                        if matches:
                            qc_file = Path(max(matches, key=lambda x: Path(x).stat().st_mtime))
                            _download_files.append(qc_file)
                            _emit_log(f"Added Quarter Classification report to download: {qc_file.name}", important=True)
                        else:
                            _emit_log(f"Quarter Classification report not found. Pattern: {pattern}", important=True)

                        # Find Missing_Records.xlsx file in working directory
                        _emit_log("Locating Missing_Records file for download", important=True)
                        missing_pattern = "Missing_Records*.xlsx"
                        missing_matches = glob.glob(str(qf_date_folder / missing_pattern))

                        if missing_matches:
                            missing_file = Path(max(missing_matches, key=lambda x: Path(x).stat().st_mtime))
                            _download_files.append(missing_file)
                            _emit_log(f"Added Missing_Records file to download: {missing_file.name}", important=True)
                        else:
                            _emit_log(f"Missing_Records file not found. Pattern: {missing_pattern}", important=True)
                    else:
                        _emit_log(f"Expected exactly one date folder in {qf_working_dir}, found {len(date_folders)}", important=True)

            except Exception as e:
                _emit_log(f"Failed to process QF reports: {e}", important=True)

        # Prepare downloads: create zip if there are multiple files
        # Skip for ECS - it already has the single .xlsm file ready in _download_files
        if _download_files and len(_download_files) > 0 and not is_ecs_flow:
            _set_status(stage="saving", message="Preparing files for download")
            try:
                auto_download_files()
            except Exception as e:
                _emit_log(f"Warning: Failed to prepare files for download: {e}", important=True)

        # Clear rate_updates and ltv_updates global variables after successful completion
        _ia_rate_updates = []
        _ia_ltv_updates = []
        return True, "All reports completed successfully"
    else:
        # Clear rate_updates and ltv_updates global variables even if some failed
        _ia_rate_updates = []
        _ia_ltv_updates = []
        return False, f"Some reports failed: {'; '.join(failed_reports)}"


# Master data file mapping - defines which files are used by which reports
MASTER_DATA_FILE_MAPPING = {
    "ALCL Management Accounts": {
        "reports": ["NBD-MF-01-SOFP-SOCI", "NBD-WF-15-LA"],
        "filename_pattern": "ALCL Management Accounts {month}_{year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Investment Schedule": {
        "reports": ["NBD-MF-01-SOFP-SOCI", "NBD-MF-10-GA-11-IS"],
        "filename_pattern": "Investment Schedule - {month} {year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Loan Schedule": {
        "reports": ["NBD-MF-01-SOFP-SOCI"],
        "filename_pattern": "Loan Schedule - {month} {year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Supporting Schedules": {
        "reports": ["NBD-MF-01-SOFP-SOCI"],
        "filename_pattern": "Supporting Schedules - {month} {year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Borrowing report": {
        "reports": ["NBD-MF-10-GA-11-IS"],
        "filename_pattern": "Borrowing report {date}.xlsm",
        "target_subfolder": None  # Goes to main date folder
    },
    "Disbursement with Budget": {
        "reports": ["NBD-MF-23-IA"],
        "filename_pattern": "Disbursement with Budget - {month} {year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Information Request from Credit": {
        "reports": ["NBD-MF-23-IA"],
        "filename_pattern": "Information Request from Credit-{month} {year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Net Portfolio": {
        "reports": ["NBD-MF-23-IA"],
        "filename_pattern": "Net Portfolio-{month}-{year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "YARD STOCK AS AT": {
        "reports": ["NBD-MF-23-IA"],
        "filename_pattern": "YARD STOCK AS AT {date} FINAL.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Cadre": {
        "reports": ["NBD-MF-23-C1-C2"],
        "filename_pattern": "Cadre{date}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Unutilized Amount": {
        "reports": ["NBD-MF-23-C1-C2"],
        "filename_pattern": "Unutilized Amount {month} {year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    },
    "Daily Bank Balances": {
        "reports": ["NBD-WF-15-LA"],
        "filename_pattern": "Daily Bank Balances - {month} {year}.xlsx",
        "target_subfolder": "Input"  # Goes to Input subfolder
    },
    # M2M is no longer a file - amounts are entered via form and saved to Master_Data.xlsx NBD-WF-15-LA sheet
    # "M2M": {
    #     "reports": ["NBD-WF-15-LA"],
    #     "filename_pattern": "M2M {mmdd}.xlsx",
    #     "target_subfolder": "Input"  # Goes to Input subfolder
    # },
    "BankBalanceGLCodes": {
        "reports": ["NBD-WF-15-LA"],
        "filename_pattern": "BankBalanceGLCodes.xlsx",
        "target_subfolder": None  # Special handling - goes to working folder
    },
    "Rec Target": {
        "reports": ["NBD-MF-23-C4"],
        "filename_pattern": "Rec Target {month} {year}.xlsx",
        "target_subfolder": None  # Goes to main date folder
    }
}

# Global counter for debugging
_centralized_call_count = 0
_processed_files = set()  # Track processed files to prevent duplicates

def update_master_data_across_reports(file_category: str, source_file_path: Path, month: str = None, year: str = None, date: str = None) -> None:
    global _centralized_call_count, _processed_files
    
    # Create a unique key for this file processing
    file_key = f"{file_category}_{source_file_path.name}_{month}_{year}_{date}"
    
    # Check if this file has already been processed
    if file_key in _processed_files:
        _emit_log(f"*** SKIPPING DUPLICATE PROCESSING *** {file_category} from {source_file_path} (already processed)", important=True)
        return
    
    # Mark this file as being processed
    _processed_files.add(file_key)
    
    # Additional check: if the source file doesn't exist, skip processing
    if not source_file_path.exists():
        _emit_log(f"*** SKIPPING PROCESSING - SOURCE FILE NOT FOUND *** {file_category} from {source_file_path}", important=True)
        return
    
    _centralized_call_count += 1
    _emit_log(f"*** CENTRALIZED FUNCTION CALLED #{_centralized_call_count} *** {file_category} from {source_file_path}", important=True)
    """
    Update a master data file across all reports that use it.
    
    Args:
        file_category: The category name from MASTER_DATA_FILE_MAPPING
        source_file_path: Path to the source file in Master Inputs
        month: Month name (for files that use month/year)
        year: Year (for files that use month/year)
        date: Date string (for files that use date)
    """
    if file_category not in MASTER_DATA_FILE_MAPPING:
        _emit_log(f"Unknown file category: {file_category}", important=True)
        return
    
    file_config = MASTER_DATA_FILE_MAPPING[file_category]
    reports = file_config["reports"]
    filename_pattern = file_config["filename_pattern"]
    target_subfolder = file_config["target_subfolder"]
    
    _emit_log(f"=== CENTRALIZED UPDATE START === {file_category} across {len(reports)} reports: {reports}", important=True)
    _emit_log(f"Source file: {source_file_path}", important=True)
    _emit_log(f"Parameters - month: {month}, year: {year}, date: {date}", important=True)
    
    # Generate the target filename
    if month and year:
        target_filename = filename_pattern.format(month=month, year=year)
    elif date:
        if file_category == "Borrowing report":
            # Convert YYYY-MM-DD to DD-MM-YYYY for Borrowing report
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%d-%m-%Y")
            target_filename = filename_pattern.format(date=formatted_date)
        elif file_category == "M2M":
            # Convert YYYY-MM-DD to MMDD for M2M
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            mmdd = date_obj.strftime("%m%d")
            target_filename = filename_pattern.format(mmdd=mmdd)
        elif file_category == "Cadre":
            # Convert YYYY-MM-DD to DD MM YYYY for Cadre
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%d %m %Y")
            target_filename = filename_pattern.format(date=formatted_date)
        elif file_category == "YARD STOCK AS AT":
            # Convert YYYY-MM-DD to DD-MM-YYYY for YARD STOCK AS AT
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%d-%m-%Y")
            target_filename = filename_pattern.format(date=formatted_date)
        else:
            target_filename = filename_pattern.format(date=date)
    else:
        _emit_log(f"Missing required parameters for {file_category}", important=True)
        return
    
    # Copy to each report's outputs folder
    _emit_log(f"Processing {file_category} for {len(reports)} reports: {reports}", important=True)
    for report_id in reports:
        try:
            _emit_log(f"*** PROCESSING {file_category} FOR REPORT: {report_id} ***", important=True)
            # Get the outputs folder for this report
            if report_id == "NBD-MF-10-GA-11-IS":
                # Special case for GA IS
                report_outputs_dir = OUTPUTS_DIR / "NBD-MF-10-GA & NBD-MF-11-IS"
            else:
                report_outputs_dir = OUTPUTS_DIR / REPORT_STRUCTURE[report_id]["folder"]
            
            _emit_log(f"Processing {file_category} for {report_id} -> {report_outputs_dir}", important=True)
            
            if not report_outputs_dir.exists():
                _emit_log(f"Outputs directory not found for {report_id}: {report_outputs_dir}, creating it", important=True)
                report_outputs_dir.mkdir(parents=True, exist_ok=True)
            
            # Find the latest completed date folder
            _, date_folder = find_latest_completed_date(report_outputs_dir)
            if not date_folder or not date_folder.exists():
                _emit_log(f"No completed date folder found for {report_id} in {report_outputs_dir}, creating placeholder", important=True)
                # Create a placeholder date folder using today's date
                today = datetime.now()
                date_folder_name = today.strftime("%d-%m-%Y")
                date_folder = report_outputs_dir / date_folder_name
                date_folder.mkdir(parents=True, exist_ok=True)
                _emit_log(f"Created placeholder date folder: {date_folder}", important=True)
            
            _emit_log(f"Found date folder for {report_id}: {date_folder}", important=True)
            
            # Determine target path
            # Check for report-specific subfolder first, then fall back to default
            report_specific_subfolders = file_config.get("report_specific_subfolders", {})
            if report_id in report_specific_subfolders:
                target_subfolder = report_specific_subfolders[report_id]
            else:
                target_subfolder = file_config.get("target_subfolder")
            
            if target_subfolder:
                target_dir = date_folder / target_subfolder
                target_dir.mkdir(parents=True, exist_ok=True)
                _emit_log(f"Using subfolder for {report_id}: {target_subfolder}", important=True)
            else:
                target_dir = date_folder
                _emit_log(f"Using main date folder for {report_id}", important=True)
            
            # Copy the new file FIRST using win32
            target_path = target_dir / target_filename
            _emit_log(f"Copying {file_category} from {source_file_path} to {target_path}", important=True)
            _emit_log(f"Source file exists: {source_file_path.exists()}", important=True)
            _emit_log(f"Target directory exists: {target_dir.exists()}", important=True)
            _emit_log(f"Target path: {target_path}", important=True)
            # Skip the old win32 copy logic - we'll use the new approach below
            
            # Use the new master data file copying approach
            success, message = copy_master_data_file_with_cleanup_improved(
                file_category, 
                source_file_path, 
                target_dir, 
                target_filename,
                force_overwrite=True,
                handle_locked_files="rename"
            )
            if not success:
                _emit_log(f"Failed to copy {file_category} to {report_id}: {message}", important=True)
                return
            else:
                _emit_log(f"Successfully copied {file_category} to {report_id}: {message}", important=True)
            
            # Simple approach: No complex file removal needed
            # The shutil.copy2 will overwrite existing files automatically
            _emit_log(f"*** SIMPLE APPROACH COMPLETE *** {file_category} -> {report_id}", important=True)
            
            # Simple validation: just check if the file exists
            if target_path.exists():
                _emit_log(f"*** SUCCESS: FILE COPIED AND VERIFIED *** {file_category} -> {report_id}: {target_path.name}", important=True)
            else:
                _emit_log(f"*** ERROR: FILE NOT FOUND AFTER COPY *** {file_category} -> {report_id}", important=True)
            
        except Exception as e:
            _emit_log(f"Failed to update {file_category} for {report_id}: {e}", important=True)
    
    _emit_log(f"=== CENTRALIZED UPDATE COMPLETE === {file_category}", important=True)


def test_file_removal_logic(target_dir: Path, prefix: str) -> None:
    """Test function to verify file removal logic works correctly."""
    _emit_log(f"*** TESTING FILE REMOVAL LOGIC ***", important=True)
    _emit_log(f"Target directory: {target_dir}", important=True)
    _emit_log(f"Prefix: '{prefix}'", important=True)
    
    if not target_dir.exists():
        _emit_log(f"ERROR: Target directory does not exist: {target_dir}", important=True)
        return
    
    # Get all files in the directory
    all_files = [f for f in target_dir.iterdir() if f.is_file()]
    _emit_log(f"Found {len(all_files)} files in target directory", important=True)
    
    for f in all_files:
        file_name = f.name
        file_suffix = f.suffix.lower()
        starts_with_prefix = file_name.startswith(prefix)
        has_valid_suffix = file_suffix in [".xlsx", ".xlsm"]
        
        _emit_log(f"File: '{file_name}' (prefix: '{prefix}', starts_with: {starts_with_prefix}, suffix: '{file_suffix}', valid_suffix: {has_valid_suffix})", important=True)
        
        if starts_with_prefix and has_valid_suffix:
            _emit_log(f"WOULD REMOVE: {file_name}", important=True)
        else:
            _emit_log(f"WOULD KEEP: {file_name}", important=True)

def ensure_directories_exist() -> None:
    """Ensure all necessary directories exist."""
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    WORKING_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    MASTER_INPUTS_DIR.mkdir(parents=True, exist_ok=True)
    
    # Create Master Inputs subdirectories
    (MASTER_INPUTS_DIR / "ALCL Management Accounts").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Supporting Schedules").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Borrowing report").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Disbursement with Budget").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Information Request from Credit").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Net Portfolio").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "YARD STOCK AS AT").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Cadre").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Unutilized Amount").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Daily Bank Balances").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "M2M").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "BankBalanceGLCodes").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "AFL Borrowing schedule").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Cash Flow").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Payments and Receipts").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Payments").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Types of Receipt").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "SasiaNet").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Tax Payments").mkdir(parents=True, exist_ok=True)
    (MASTER_INPUTS_DIR / "Cashflow Movement").mkdir(parents=True, exist_ok=True)

    # Ensure Master_Data.xlsx exists with required sheets
    if not MASTER_DATA_XLSX.exists():
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default sheet
        wb.create_sheet("NBD-MF-01-SOFP-SOCI")
        wb.create_sheet("NBD-MF-23-IA-C1&C2-C5-IP")
        wb.create_sheet("NBD-MF-20-C1-C6")
        wb.save(MASTER_DATA_XLSX)
        wb.close()


@app.route("/upload-ecs-master-inputs", methods=["POST"])
def upload_ecs_master_inputs():
    """Standalone endpoint for uploading ECS Master Input files.
    Saves each uploaded file to Master Inputs/<subfolder>/ and copies
    it to the latest ECS outputs date folder. All 4 slots are optional;
    only uploaded files are processed.
    """
    ecs_outputs_dir = OUTPUTS_DIR / "NBD_WF_18_ECS"
    _, ecs_date_folder = find_latest_completed_date(ecs_outputs_dir)

    any_uploaded = False

    # ── 1. AFL Borrowing Schedule ─────────────────────────────────────────────
    afl_file = request.files.get("afl_borrowing_file")
    if afl_file and afl_file.filename:
        borrowing_date = request.form.get("borrowing_schedule_date", "").strip()
        if not borrowing_date:
            flash("Date is required when uploading AFL Borrowing schedule.", "error")
            return redirect(url_for("index"))
        try:
            date_obj = datetime.strptime(borrowing_date, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%d-%m-%Y")
        except ValueError:
            flash("Invalid date format for AFL Borrowing schedule.", "error")
            return redirect(url_for("index"))

        try:
            afl_dir = MASTER_INPUTS_DIR / "AFL Borrowing schedule"
            afl_dir.mkdir(parents=True, exist_ok=True)
            filename = f"AFL Borrowing schedule {formatted_date}.xlsx"
            dest_master = afl_dir / filename
            afl_file.save(str(dest_master))
            any_uploaded = True
            _emit_log(f"Saved AFL Borrowing schedule: {filename}", important=True)
            if ecs_date_folder and ecs_date_folder.exists():
                for old in ecs_date_folder.glob("AFL Borrowing schedule *.xlsx"):
                    old.unlink()
                    _emit_log(f"Removed old AFL Borrowing schedule: {old.name}", important=True)
                shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                _emit_log(f"Copied AFL Borrowing schedule to ECS outputs: {filename}", important=True)
            else:
                _emit_log("No ECS date folder found; AFL Borrowing schedule saved to Master Inputs only", important=True)
        except Exception as e:
            _emit_log(f"AFL Borrowing schedule upload error: {e}", important=True)
            flash(f"AFL Borrowing schedule upload error: {e}", "error")
            return redirect(url_for("index"))

    # ── 2. Cash Flow ──────────────────────────────────────────────────────────
    cf_file = request.files.get("cash_flow_file")
    if cf_file and cf_file.filename:
        cf_month = request.form.get("cash_flow_month", "").strip()
        cf_year = request.form.get("cash_flow_year", "").strip()
        if not cf_month or not cf_year:
            flash("Month and Year are required when uploading Cash Flow file.", "error")
            return redirect(url_for("index"))
        if not cf_year.isdigit() or len(cf_year) != 4:
            flash(f"Invalid year '{cf_year}' for Cash Flow. Enter a 4-digit year.", "error")
            return redirect(url_for("index"))
        year_int = int(cf_year)
        if year_int < 2000 or year_int > 2100:
            flash(f"Year {cf_year} is out of range (2000–2100).", "error")
            return redirect(url_for("index"))
        try:
            cf_dir = MASTER_INPUTS_DIR / "Cash Flow"
            cf_dir.mkdir(parents=True, exist_ok=True)
            filename = f"Cash Flow - {cf_month} {cf_year}.xlsx"
            dest_master = cf_dir / filename
            cf_file.save(str(dest_master))
            any_uploaded = True
            _emit_log(f"Saved Cash Flow: {filename}", important=True)
            if ecs_date_folder and ecs_date_folder.exists():
                for old in ecs_date_folder.glob("Cash Flow *.xlsx"):
                    old.unlink()
                    _emit_log(f"Removed old Cash Flow: {old.name}", important=True)
                shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                _emit_log(f"Copied Cash Flow to ECS outputs: {filename}", important=True)
            else:
                _emit_log("No ECS date folder found; Cash Flow saved to Master Inputs only", important=True)
        except Exception as e:
            _emit_log(f"Cash Flow upload error: {e}", important=True)
            flash(f"Cash Flow upload error: {e}", "error")
            return redirect(url_for("index"))

    # ── 3. Payments and Receipts ──────────────────────────────────────────────
    pr_file = request.files.get("payments_receipts_file")
    if pr_file and pr_file.filename:
        try:
            pr_dir = MASTER_INPUTS_DIR / "Payments and Receipts"
            pr_dir.mkdir(parents=True, exist_ok=True)
            filename = "Payments and Receipts.xlsx"
            dest_master = pr_dir / filename
            if dest_master.exists():
                dest_master.unlink()
            pr_file.save(str(dest_master))
            any_uploaded = True
            _emit_log("Saved Payments and Receipts.xlsx", important=True)
            if ecs_date_folder and ecs_date_folder.exists():
                old_pr = ecs_date_folder / filename
                if old_pr.exists():
                    old_pr.unlink()
                    _emit_log("Removed old Payments and Receipts from ECS outputs", important=True)
                shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                _emit_log("Copied Payments and Receipts to ECS outputs", important=True)
            else:
                _emit_log("No ECS date folder found; Payments and Receipts saved to Master Inputs only", important=True)
        except Exception as e:
            _emit_log(f"Payments and Receipts upload error: {e}", important=True)
            flash(f"Payments and Receipts upload error: {e}", "error")
            return redirect(url_for("index"))

    # ── 4. Daily Bank Balances ────────────────────────────────────────────────
    dbb_file = request.files.get("daily_bank_balances_file")
    if dbb_file and dbb_file.filename:
        dbb_month = request.form.get("dbb_month", "").strip()
        dbb_year = request.form.get("dbb_year", "").strip()
        if not dbb_month or not dbb_year:
            flash("Month and Year are required when uploading Daily Bank Balances file.", "error")
            return redirect(url_for("index"))
        if not dbb_year.isdigit() or len(dbb_year) != 4:
            flash(f"Invalid year '{dbb_year}' for Daily Bank Balances. Enter a 4-digit year.", "error")
            return redirect(url_for("index"))
        try:
            dbb_dir = MASTER_INPUTS_DIR / "Daily Bank Balances"
            dbb_dir.mkdir(parents=True, exist_ok=True)
            filename = f"Daily Bank Balances - {dbb_month} {dbb_year}.xlsx"
            dest_master = dbb_dir / filename
            dbb_file.save(str(dest_master))
            any_uploaded = True
            _emit_log(f"Saved Daily Bank Balances: {filename}", important=True)
            if ecs_date_folder and ecs_date_folder.exists():
                for old in ecs_date_folder.glob("Daily Bank Balances*.xlsx"):
                    old.unlink()
                    _emit_log(f"Removed old Daily Bank Balances: {old.name}", important=True)
                shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                _emit_log(f"Copied Daily Bank Balances to ECS outputs: {filename}", important=True)
            else:
                _emit_log("No ECS date folder found; Daily Bank Balances saved to Master Inputs only", important=True)
        except Exception as e:
            _emit_log(f"Daily Bank Balances upload error: {e}", important=True)
            flash(f"Daily Bank Balances upload error: {e}", "error")
            return redirect(url_for("index"))

    # ── Final feedback ────────────────────────────────────────────────────────
    if any_uploaded:
        flash("ECS master input files uploaded successfully.", "success")
    else:
        flash("No files were uploaded.", "warning")
    return redirect(url_for("index"))


# ==================== FLASK ROUTES ====================

app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-prod")


@app.route("/", methods=["GET"])
def index():
    from flask import session

    ensure_directories_exist()
    today_iso = datetime.today().strftime("%Y-%m-%d")

    # Clear session if requested (after user clicks "Thanks" in modal)
    if request.args.get('clear') == '1':
        session.pop('download_completed', None)
        session.pop('download_filename', None)

    # Group reports by level and set
    reports_by_level = {}
    sets_info = {}

    for report_id, info in REPORT_STRUCTURE.items():
        level = info["level"]
        if level not in reports_by_level:
            reports_by_level[level] = {}

        set_num = info["set"]
        if set_num not in reports_by_level[level]:
            reports_by_level[level][set_num] = []
            sets_info[set_num] = {"estimated_minutes": 0}

        # Filter out specific reports from Level 2 Set 7 UI (handled by C1–C6 flow)
        if level == 2 and set_num == 7 and report_id in ("NBD-MF-20-C1", "NBD-MF-01-SOFP-SOCI"):
            continue

        # Filter out old GA & IS reports from Set 5 UI (replaced by NBD-MF-10-GA-11-IS)
        if level == 2 and set_num == 5 and report_id in ("NBD-MF-10-GA", "NBD-MF-11-IS"):
            continue

        # Filter out SOFP-SOCI from Set 6 UI (not required to be listed)
        if level == 2 and set_num == 6 and report_id == "NBD-MF-01-SOFP-SOCI":
            continue

        report_entry = {
            "id": report_id,
            "name": info["name"],
            "step": info["step"],
            "estimated_minutes": info.get("estimated_minutes", 10)
        }
        reports_by_level[level][set_num].append(report_entry)
        sets_info[set_num]["estimated_minutes"] += report_entry["estimated_minutes"]

    # Sort reports within each set by step
    for level in reports_by_level:
        for set_num in reports_by_level[level]:
            reports_by_level[level][set_num].sort(key=lambda x: x["step"])


    # Read last completed dates for each report output folder
    last_completed_dates = {}
    for set_num, folder_mapping in SET_COMPLETION_MAPPING.items():
        if isinstance(folder_mapping, dict):
            # Multi-step set (e.g., Set 2)
            for step_num, folder_name in folder_mapping.items():
                dates_file = OUTPUTS_DIR / folder_name / "completed_dates.txt"
                try:
                    if dates_file.exists():
                        lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
                        if lines:
                            last_completed_dates[f"{set_num}_{step_num}"] = lines[-1]
                except Exception:
                    pass
        else:
            dates_file = OUTPUTS_DIR / folder_mapping / "completed_dates.txt"
            try:
                if dates_file.exists():
                    lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
                    if lines:
                        last_completed_dates[str(set_num)] = lines[-1]
            except Exception:
                pass

    # Get AFL files from MF_04_LA last completed date folder
    mf04_afl_files = []
    try:
        mf04_outputs = OUTPUTS_DIR / "NBD_MF_04_LA"
        _, mf04_date_folder = find_latest_completed_date(mf04_outputs)
        if mf04_date_folder and mf04_date_folder.exists():
            afl_files = sorted(
                mf04_date_folder.glob("AFL Liquidity*.xlsx"),
                key=lambda x: x.stat().st_mtime,
                reverse=True  # Newest first
            )
            mf04_afl_files = [f.name for f in afl_files[:5]]  # Max 5 files
    except Exception:
        pass

    return render_template("index_new.html",
                         today_iso=today_iso,
                         reports_by_level=reports_by_level,
                         sets_info=sets_info,
                         report_structure=REPORT_STRUCTURE,
                         last_completed_dates=last_completed_dates,
                         mf04_afl_files=mf04_afl_files)


@app.route("/run-reports", methods=["POST"])
def run_reports():
    """Run selected reports with individual dates per set."""
    ensure_directories_exist()

    form_data = request.form
    report_ids_with_dates = {}

    for key in form_data:
        if key.startswith("date_"):
            set_num = key.replace("date_", "")
            date_str = form_data[key]

            set_reports = form_data.getlist(f"reports_{set_num}")

            if set_reports and date_str:
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                report_month = dt.strftime("%B")
                report_year = str(dt.year)

                for report_id in set_reports:
                    report_ids_with_dates[report_id] = (date_str, report_month, report_year)

    if not report_ids_with_dates:
        flash("Please select at least one report with a date.", "error")
        return redirect(url_for("index"))

    # Check if this set needs master data input
    c1c6_ids = {"NBD-MF-20-C1","NBD-MF-20-C2","NBD-MF-20-C3","NBD-MF-20-C4","NBD-MF-20-C5","NBD-MF-20-C6"}
    annual_c_ids = {"NBD-AF-20-C2","NBD-AF-20-C3","NBD-AF-20-C4","NBD-AF-20-C5","NBD-AF-20-C6"}
    qf_set8_ids = {"NBD-QF-23-C9", "NBD-QF-23-C10", "NBD-QF-23-C6", "NBD-QF-23-C4", "NBD-QF-23-C3", "NBD-QF-16-SF", "NBD-QF-16-FI"}
    needs_master_data = any(report_id in ("NBD-MF-01-SOFP-SOCI", "NBD-AF-01-SOFP-SOCI", "NBD-MF-23-IA", "NBD-MF-23-C1-C2", "NBD-MF-10-GA-11-IS", "NBD-MF-23-C4", "NBD-WF-15-LA", "NBD-WF-18-ECS") or report_id in c1c6_ids or report_id in annual_c_ids or report_id in qf_set8_ids for report_id in report_ids_with_dates)
    _emit_log(f"*** MASTER DATA CHECK *** needs_master_data: {needs_master_data}, report_ids: {list(report_ids_with_dates.keys())}", important=True)
    
    # Store LA month-end flag and time deposits value if provided
    la_month_end = form_data.get("la_month_end") == "1"
    la_time_deposits = form_data.get("la_time_deposits", "")

    if needs_master_data:
        _emit_log(f"*** REDIRECTING TO MASTER DATA FORM ***", important=True)
        # Store report data in session and redirect to master data form
        from flask import session
        session['report_ids_with_dates'] = report_ids_with_dates
        session['la_month_end'] = la_month_end
        session['la_time_deposits'] = la_time_deposits
        return redirect(url_for("master_data_form"))
    
    # Proceed directly if no master data needed
    return start_automation_with_data(report_ids_with_dates, {}, {})


@app.route("/master-data-form", methods=["GET"])
def master_data_form():
    """Display master data input form with latest data pre-filled."""
    _emit_log(f"*** MASTER DATA FORM ROUTE REACHED ***", important=True)
    from flask import session
    sofp_data, ia_data, c1c6_data, ga_is_data, ecs_treasury_data = get_latest_master_data()
    uploaded_files = get_latest_uploaded_files()
    la_treasury_amounts = get_la_treasury_amounts()
    report_ids_with_dates = session.get('report_ids_with_dates', {})
    selected_ids = set(report_ids_with_dates.keys())
    c1c6_ids = {"NBD-MF-20-C1","NBD-MF-20-C2","NBD-MF-20-C3","NBD-MF-20-C4","NBD-MF-20-C5","NBD-MF-20-C6"}
    annual_c_ids = {"NBD-AF-20-C2","NBD-AF-20-C3","NBD-AF-20-C4","NBD-AF-20-C5","NBD-AF-20-C6"}
    qf_set8_ids = {"NBD-QF-23-C9", "NBD-QF-23-C10", "NBD-QF-23-C6", "NBD-QF-23-C4", "NBD-QF-23-C3", "NBD-QF-16-SF", "NBD-QF-16-FI"}
    show_sofp = any(x in ("NBD-MF-01-SOFP-SOCI", "NBD-AF-01-SOFP-SOCI") for x in selected_ids)
    show_ia = any(x in ("NBD-MF-23-IA",) for x in selected_ids)
    show_c1c2 = any(x in ("NBD-MF-23-C1-C2",) for x in selected_ids)
    show_c1c6 = any(x in c1c6_ids or x in annual_c_ids for x in selected_ids)
    show_ga_is = any(x in ("NBD-MF-10-GA-11-IS",) for x in selected_ids)
    show_c4 = any(x in ("NBD-MF-23-C4",) for x in selected_ids)
    show_la = any(x in ("NBD-WF-15-LA",) for x in selected_ids)
    show_ecs = any(x in ("NBD-WF-18-ECS",) for x in selected_ids)
    show_qf = any(x in qf_set8_ids for x in selected_ids)

    # Compute M2M day labels (Mon–Sun) from the ECS input date
    ecs_m2m_dates = []
    if show_ecs:
        ecs_date_str = None
        if "NBD-WF-18-ECS" in report_ids_with_dates:
            ecs_date_str = report_ids_with_dates["NBD-WF-18-ECS"][0]
        elif report_ids_with_dates:
            ecs_date_str = list(report_ids_with_dates.values())[0][0]
        if ecs_date_str:
            ecs_dt = None
            for _fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"):
                try:
                    ecs_dt = datetime.strptime(ecs_date_str, _fmt)
                    break
                except ValueError:
                    continue
            if ecs_dt:
                week_start = ecs_dt - timedelta(days=ecs_dt.weekday())
                ecs_m2m_dates = [(week_start + timedelta(days=i)).strftime("%d/%m/%Y") for i in range(7)]

    # Get form data from request parameters (for validation error preservation)
    form_data = request.args.to_dict()

    # Merge LA treasury amounts into form_data if not already present (for pre-filling)
    if not form_data:
        form_data = la_treasury_amounts.copy()
    else:
        # Only add LA amounts if they're not already in form_data (form_data takes precedence)
        for key, value in la_treasury_amounts.items():
            if key not in form_data:
                form_data[key] = value

    # Merge ECS treasury data into form_data for pre-filling
    for key, value in ecs_treasury_data.items():
        if key not in form_data:
            form_data[key] = value

    return render_template("master_data_form.html",
                         sofp_data=sofp_data,
                         ia_data=ia_data,
                         c1c6_data=c1c6_data,
                         ga_is_data=ga_is_data,
                         show_sofp=show_sofp,
                         show_ia=show_ia,
                         show_c1c2=show_c1c2,
                         show_c1c6=show_c1c6,
                         show_ga_is=show_ga_is,
                         show_c4=show_c4,
                         show_la=show_la,
                         show_ecs=show_ecs,
                         show_qf=show_qf,
                         uploaded_files=uploaded_files,
                         ecs_master_input_info=get_ecs_master_input_info(),
                         ecs_m2m_dates=ecs_m2m_dates,
                         ecs_treasury_data=ecs_treasury_data,
                         form_data=form_data)


@app.route("/submit-master-data", methods=["POST"])
def submit_master_data():
    """Process master data form submission."""
    _emit_log(f"*** SUBMIT MASTER DATA ROUTE REACHED ***", important=True)
    from flask import session
    
    # Clear processed files set for this submission
    global _processed_files
    _processed_files.clear()
    _emit_log(f"*** CLEARED PROCESSED FILES SET ***", important=True)

    # Clear download files from any previous run to prevent unwanted downloads on validation failure
    global _download_files, _sofp_download_file
    _download_files = []
    _sofp_download_file = None

    # Get master data from form (exclude timestamp as it's auto-generated)
    sofp_data = None
    ia_data = None
    c1c6_data = None

    # Determine which sections were shown
    show_sofp = request.form.get("_show_sofp") == "1"
    show_ia = request.form.get("_show_ia") == "1"
    show_c1c2 = request.form.get("_show_c1c2") == "1"
    show_c1c6 = request.form.get("_show_c1c6") == "1"
    show_ga_is = request.form.get("_show_ga_is") == "1"
    show_c4 = request.form.get("_show_c4") == "1"

    if show_sofp:
        sofp_data = {
            "bank_loans": request.form.get("bank_loans", ""),
            "securitization_loans": request.form.get("securitization_loans", ""),
            "foreign_funding": request.form.get("foreign_funding", ""),
            "interest_on_lease_liabilities": request.form.get("interest_on_lease_liabilities", ""),
            "depreciation_of_right_of_use_assets": request.form.get("depreciation_of_right_of_use_assets", ""),
            "bs_1_1_2_2_0_0": request.form.get("bs_1_1_2_2_0_0", ""),
            "bs_1_2_2_2_0_0": request.form.get("bs_1_2_2_2_0_0", ""),
            "bs_1_1_12_2_0_0": request.form.get("bs_1_1_12_2_0_0", "")
        }
    if show_ia:
        ia_data = {
            "annual_interest_rate": request.form.get("annual_interest_rate", "")
        }
        # Get rate updates (optional - only if user provides data)
        rate_updates_data = request.form.get("rate_updates_data", "[]")
        try:
            rate_updates = json.loads(rate_updates_data)
            if rate_updates:
                ia_data["rate_updates"] = rate_updates
                _emit_log(f"Received {len(rate_updates)} contract rate updates from user", important=True)
            else:
                ia_data["rate_updates"] = []
        except json.JSONDecodeError as e:
            ia_data["rate_updates"] = []
            _emit_log(f"Error parsing rate updates JSON: {e}", important=False)
        except Exception as e:
            ia_data["rate_updates"] = []
            _emit_log(f"Unexpected error processing rate updates: {e}", important=False)

        # LTV updates are handled via ltv.xlsx file upload (not form data)
        ia_data["ltv_updates"] = []  # Keep for backward compatibility but not used
    if show_c1c6:
        # Checkbox: if checked, value is "yes"; if unchecked, value is empty (default to "no")
        audit_completed = request.form.get("new_audit_completed", "")
        if audit_completed != "yes":
            audit_completed = "no"

        c1c6_data = {
            "guarantees": request.form.get("c1c6_guarantees", ""),
            "tier2_instruments": request.form.get("tier2_instruments", ""),
            "tier2_revaluation": request.form.get("tier2_revaluation", ""),
            "tier2_provisions": request.form.get("tier2_provisions", ""),
            "new_audit_completed": audit_completed,
            "stated_capital": request.form.get("stated_capital", ""),
            "statutory_reserve_fund": request.form.get("statutory_reserve_fund", ""),
            "fair_value_through_other_comprehensive_income_reserve": request.form.get("fair_value_through_other_comprehensive_income_reserve", ""),
            "revaluation_reserve": request.form.get("revaluation_reserve", ""),
            "retained_earnings": request.form.get("retained_earnings", "")
        }

    # Extract GA & IS data (DEFERRED TAX PAYABLE)
    ga_is_data = None
    if show_ga_is:
        ga_is_data = {
            "deferred_tax_0_3m": request.form.get("deferred_tax_0_3m", ""),
            "deferred_tax_3_6m": request.form.get("deferred_tax_3_6m", ""),
            "deferred_tax_6_12m": request.form.get("deferred_tax_6_12m", ""),
            "deferred_tax_1_3y": request.form.get("deferred_tax_1_3y", ""),
            "deferred_tax_3_5y": request.form.get("deferred_tax_3_5y", ""),
            "deferred_tax_over_5y": request.form.get("deferred_tax_over_5y", "")
        }

    # Determine exact target sheets based on selected report IDs (monthly vs annual)
    report_ids_with_dates = session.get('report_ids_with_dates', {})
    selected_ids = set(report_ids_with_dates.keys())

    sofp_sheet = None
    if "NBD-AF-01-SOFP-SOCI" in selected_ids:
        sofp_sheet = "NBD-AF-01-SOFP-SOCI"
    elif "NBD-MF-01-SOFP-SOCI" in selected_ids:
        sofp_sheet = "NBD-MF-01-SOFP-SOCI"

    c1c6_sheet = None
    monthly_c_ids = {"NBD-MF-20-C1","NBD-MF-20-C2","NBD-MF-20-C3","NBD-MF-20-C4","NBD-MF-20-C5","NBD-MF-20-C6"}
    annual_c_ids = {"NBD-AF-20-C2","NBD-AF-20-C3","NBD-AF-20-C4","NBD-AF-20-C5","NBD-AF-20-C6"}
    if any(r in selected_ids for r in annual_c_ids):
        c1c6_sheet = "NBD-AF-20-C1-C6"
    elif any(r in selected_ids for r in monthly_c_ids):
        c1c6_sheet = "NBD-MF-20-C1-C6"

    # Handle file uploads for SOFP
    _emit_log("*** SOFP FILE UPLOAD SECTION REACHED ***", important=True)
    uploaded_files = {}
    months = {}
    if show_sofp:
        categories = [
            "ALCL Management Accounts",
            "Investment Schedule", 
            "Loan Schedule",
            "Supporting Schedules"
        ]
        
        for category in categories:
            field_name = category.lower().replace(' ', '_')
            month_field = f"{field_name}_month"

            # Get month selection
            month = request.form.get(month_field, "").strip()

            # Handle file upload
            if f"{field_name}_file" in request.files:
                file = request.files[f"{field_name}_file"]
                if file and file.filename:
                    # Month is REQUIRED when file is uploaded
                    if not month:
                        flash(f"Month is required when uploading {category} file.", "error")
                        return redirect(url_for("master_data_form", **request.form.to_dict()))

                    # Save month for later use
                    months[month_field] = month

                    # Save file to Master Inputs directory
                    category_dir = MASTER_INPUTS_DIR / category
                    category_dir.mkdir(parents=True, exist_ok=True)

                    # Generate filename with month and year (exact format as specified)
                    current_year = datetime.now().year
                    if category == "ALCL Management Accounts":
                        filename = f"ALCL Management Accounts {month}_{current_year}.xlsx"
                    elif category == "Investment Schedule":
                        filename = f"Investment Schedule - {month} {current_year}.xlsx"
                    elif category == "Loan Schedule":
                        filename = f"Loan Schedule {month} {current_year}.xlsx"
                    elif category == "Supporting Schedules":
                        filename = f"Supporting Schedules - {month} {current_year}.xlsx"
                    else:
                        filename = file.filename

                    file_path = category_dir / filename
                    file.save(str(file_path))
                    uploaded_files[category] = str(file_path)
                    _emit_log(f"Uploaded {category} file: {filename}", important=True)

                    # Use centralized master data update system
                    _emit_log(f"*** CALLING CENTRALIZED SYSTEM FOR {category} FROM SOFP ***", important=True)
                    update_master_data_across_reports(category, file_path, month=month, year=current_year)
            elif month:
                # Month provided but no file - still save the month
                months[month_field] = month

        # Handle IFRS Codes file upload (no month required)
        if "ifrs_file" in request.files:
            file = request.files["ifrs_file"]
            if file and file.filename:
                # Save file to Master Inputs/IFRS directory
                ifrs_dir = MASTER_INPUTS_DIR / "IFRS"
                ifrs_dir.mkdir(parents=True, exist_ok=True)

                # Always rename to ifrs_code.xlsx
                filename = "ifrs_code.xlsx"
                file_path = ifrs_dir / filename

                # Remove existing file if present
                if file_path.exists():
                    file_path.unlink()
                    _emit_log(f"Removed existing IFRS file: {filename}", important=True)

                # Save the new file
                file.save(str(file_path))
                uploaded_files["IFRS"] = str(file_path)
                _emit_log(f"Uploaded IFRS file: {filename}", important=True)

                # Copy to NBD_MF_01_SOFP_SOCI outputs folder
                try:
                    sofp_outputs_dir = OUTPUTS_DIR / "NBD_MF_01_SOFP_SOCI"
                    if sofp_outputs_dir.exists():
                        # Read completed_dates.txt to find the last date
                        completed_dates_file = sofp_outputs_dir / "completed_dates.txt"
                        if completed_dates_file.exists():
                            with open(completed_dates_file, 'r') as f:
                                lines = [line.strip() for line in f if line.strip()]
                                if lines:
                                    last_date = lines[-1]
                                    _emit_log(f"Last completed date for IFRS copy: {last_date}", important=True)

                                    # Find the date folder with variations (handles versioning like 31-08-2025(30))
                                    date_folder = find_date_folder_with_variations(sofp_outputs_dir, last_date)
                                    if date_folder and date_folder.exists():
                                        _emit_log(f"Found date folder with versioning: {date_folder}", important=True)

                                        # Remove existing ifrs_code.xlsx
                                        existing_ifrs = date_folder / "ifrs_code.xlsx"
                                        if existing_ifrs.exists():
                                            existing_ifrs.unlink()
                                            _emit_log(f"Removed existing IFRS file from outputs: {existing_ifrs}", important=True)

                                        # Copy the new file
                                        dest_path = date_folder / "ifrs_code.xlsx"
                                        shutil.copy2(file_path, dest_path)
                                        _emit_log(f"Copied IFRS file to outputs: {dest_path}", important=True)
                                    else:
                                        _emit_log(f"Date folder not found for: {last_date}", important=True)
                                else:
                                    _emit_log("completed_dates.txt is empty", important=True)
                        else:
                            _emit_log("completed_dates.txt not found", important=True)
                    else:
                        _emit_log("SOFP outputs directory not found", important=True)
                except Exception as e:
                    _emit_log(f"Error copying IFRS file to outputs: {e}", important=True)

    # Handle file uploads for GA IS
    ga_is_uploaded_files = {}
    ga_is_dates = {}
    if show_ga_is:
        # Handle Borrowing report upload
        if "borrowing_report_file" in request.files:
            file = request.files["borrowing_report_file"]
            if file and file.filename:
                # Get date selection
                borrowing_date = request.form.get("borrowing_date", "").strip()
                if borrowing_date:
                    ga_is_dates["borrowing_date"] = borrowing_date
                    
                    # Save file to Master Inputs directory
                    category_dir = MASTER_INPUTS_DIR / "Borrowing report"
                    category_dir.mkdir(parents=True, exist_ok=True)
                    
                    # Generate filename with date (format: Borrowing report 31-08-2025.xlsm)
                    # Convert date from YYYY-MM-DD to DD-MM-YYYY format
                    date_obj = datetime.strptime(borrowing_date, "%Y-%m-%d")
                    formatted_date = date_obj.strftime("%d-%m-%Y")
                    filename = f"Borrowing report {formatted_date}.xlsm"
                    
                    file_path = category_dir / filename
                    file.save(str(file_path))
                    ga_is_uploaded_files["Borrowing report"] = str(file_path)
                    _emit_log(f"Uploaded Borrowing report file: {filename}", important=True)

                    # Use centralized master data update system
                    update_master_data_across_reports("Borrowing report", file_path, date=borrowing_date)

    # Handle file uploads for C4 (Rec Target)
    c4_uploaded_files = {}
    c4_months = {}
    c4_years = {}
    if show_c4:
        # Handle Rec Target upload
        if "rectarget_file" in request.files:
            file = request.files["rectarget_file"]
            if file and file.filename:
                # Get month and year selection
                month = request.form.get("rectarget_month", "").strip()
                year = request.form.get("rectarget_year", "").strip()

                # Month and Year are REQUIRED when file is uploaded
                if not month or not year:
                    flash(f"Month and Year are required when uploading Rec Target file.", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

                # Validate year - must be exactly 4 digits, no symbols or characters
                if not year.isdigit() or len(year) != 4:
                    flash(f"Invalid year for Rec Target: {year}. Please enter a valid 4-digit year (e.g., 2025).", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))
                year_int = int(year)
                if year_int < 2000 or year_int > 2100:
                    flash(f"Invalid year for Rec Target: {year}. Please enter a year between 2000 and 2100.", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

                # Save month/year for later use
                c4_months["rectarget_month"] = month
                c4_years["rectarget_year"] = year

                # Save file to Master Inputs directory
                category_dir = MASTER_INPUTS_DIR / "Rec Target"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: Rec Target - July 2025.xlsx
                filename = f"Rec Target - {month} {year}.xlsx"

                file_path = category_dir / filename
                file.save(str(file_path))
                c4_uploaded_files["Rec Target"] = str(file_path)
                _emit_log(f"Uploaded Rec Target file: {filename}", important=True)

                # Use centralized master data update system
                update_master_data_across_reports("Rec Target", file_path, month=month, year=year)

                # Copy to latest C4 outputs folder
                try:
                    c4_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_C4"
                    _emit_log(f"C4 outputs directory: {c4_outputs_dir}", important=True)

                    if c4_outputs_dir.exists():
                        _emit_log(f"C4 outputs directory exists, finding latest completed date...", important=True)
                        # Find latest completed date in C4
                        _, c4_date_folder = find_latest_completed_date(c4_outputs_dir)

                        if c4_date_folder and c4_date_folder.exists():
                            _emit_log(f"Found C4 latest date folder: {c4_date_folder}", important=True)

                            # Remove old Rec Target files (files starting with "Rec Target" and ending with ".xlsx")
                            removed_count = 0
                            for f in c4_date_folder.iterdir():
                                if f.is_file() and f.suffix.lower() == ".xlsx":
                                    if f.name.startswith("Rec Target"):
                                        try:
                                            _emit_log(f"Removing old Rec Target file: {f.name}", important=True)
                                            f.unlink()
                                            removed_count += 1
                                        except Exception as e:
                                            _emit_log(f"Failed to remove old Rec Target file {f.name}: {e}", important=True)

                            if removed_count > 0:
                                _emit_log(f"Removed {removed_count} old Rec Target file(s) from C4 outputs", important=True)
                            else:
                                _emit_log(f"No old Rec Target files found to remove", important=True)

                            # Copy the new file
                            dest_path = c4_date_folder / filename
                            shutil.copy2(file_path, dest_path)
                            _emit_log(f"Copied Rec Target file to C4 outputs: {dest_path}", important=True)
                        else:
                            _emit_log("No completed date folder found for C4", important=True)
                    else:
                        _emit_log(f"C4 outputs directory does not exist: {c4_outputs_dir}", important=True)
                except Exception as e:
                    _emit_log(f"Failed to copy Rec Target file to C4 outputs: {e}", important=True)
                    import traceback
                    _emit_log(f"Traceback: {traceback.format_exc()}", important=True)

        # Handle Prod_Cat C4 file upload (special case - no month/year, just file)
        if "prodcat_c4_file" in request.files:
            file = request.files["prodcat_c4_file"]
            if file and file.filename:
                _emit_log("Processing Prod_Cat C4 file upload", important=True)

                # Save file to Master Inputs/Prod Cat C4 directory
                prodcat_c4_dir = MASTER_INPUTS_DIR / "Prod Cat C4"
                prodcat_c4_dir.mkdir(parents=True, exist_ok=True)

                # Rename file to Prod_Cat.xlsx (change entire name)
                filename = "Prod_Cat.xlsx"
                file_path = prodcat_c4_dir / filename

                # Remove existing file if it exists
                if file_path.exists():
                    file_path.unlink()
                    _emit_log(f"Removed existing Prod_Cat C4 file: {filename}", important=True)

                # Save new file
                file.save(str(file_path))
                c4_uploaded_files["Prod_Cat C4"] = str(file_path)
                _emit_log(f"Uploaded Prod_Cat C4 file: {filename}", important=True)

                # Copy to outputs/NBD_MF_23_C4/{date}/ folder
                # Find the last completed date from completed_dates.txt
                c4_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_C4"
                if c4_outputs_dir.exists():
                    try:
                        # Read completed_dates.txt to find last date
                        completed_dates_file = c4_outputs_dir / "completed_dates.txt"
                        if completed_dates_file.exists():
                            with open(completed_dates_file, 'r') as f:
                                dates = [line.strip() for line in f if line.strip()]
                            if dates:
                                last_date = dates[-1]  # Get last line
                                _emit_log(f"Last completed C4 date: {last_date}", important=True)

                                # Find the max version date folder for that date
                                date_folders = [d for d in c4_outputs_dir.iterdir() if d.is_dir() and d.name.startswith(last_date)]
                                if date_folders:
                                    # Sort by modification time to get the latest
                                    target_folder = max(date_folders, key=lambda x: x.stat().st_mtime)
                                    _emit_log(f"Copying Prod_Cat to C4 folder: {target_folder}", important=True)

                                    # Remove existing Prod_Cat*.xlsx files in target folder
                                    for existing_file in target_folder.glob("Prod_Cat*.xlsx"):
                                        existing_file.unlink()
                                        _emit_log(f"Removed existing {existing_file.name} from {target_folder.name}", important=True)

                                    # Copy the new file
                                    target_file = target_folder / "Prod_Cat.xlsx"
                                    shutil.copy2(file_path, target_file)
                                    _emit_log(f"Copied Prod_Cat.xlsx to {target_folder.name}", important=True)
                                else:
                                    _emit_log(f"No date folder found for {last_date} in C4 outputs", important=True)
                        else:
                            _emit_log("completed_dates.txt not found in C4 outputs", important=True)
                    except Exception as e:
                        _emit_log(f"Error copying Prod_Cat to C4 outputs: {e}", important=True)
                else:
                    _emit_log("C4 outputs directory not found", important=True)

    # Handle file uploads for LA
    la_uploaded_files = {}
    la_months = {}
    la_years = {}
    la_dates = {}
    show_la = request.form.get("_show_la") == "1"
    if show_la:
        # Handle Daily Bank Balances upload
        if "daily_bank_balances_file" in request.files:
            file = request.files["daily_bank_balances_file"]
            if file and file.filename:
                # Get month and year
                month = request.form.get("daily_bank_balances_month", "").strip()
                year = request.form.get("daily_bank_balances_year", "").strip()

                # Month and Year are REQUIRED when file is uploaded
                if not month or not year:
                    flash(f"Month and Year are required when uploading Daily Bank Balances file.", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

                # Validate year
                if not year.isdigit() or len(year) != 4:
                    flash(f"Invalid year for Daily Bank Balances: {year}. Please enter a valid 4-digit year (e.g., 2025).", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))
                year_int = int(year)
                if year_int < 2000 or year_int > 2100:
                    flash(f"Invalid year for Daily Bank Balances: {year}. Please enter a year between 2000 and 2100.", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

                la_months["daily_bank_balances_month"] = month
                la_years["daily_bank_balances_year"] = year

                # Save file to Master Inputs
                category_dir = MASTER_INPUTS_DIR / "Daily Bank Balances"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: Daily Bank Balances - Jul 2025.xlsx
                filename = f"Daily Bank Balances - {month} {year}.xlsx"

                file_path = category_dir / filename
                file.save(str(file_path))
                la_uploaded_files["Daily Bank Balances"] = str(file_path)
                _emit_log(f"Uploaded Daily Bank Balances file: {filename}", important=True)

                # Use centralized master data update system
                update_master_data_across_reports("Daily Bank Balances", file_path, month=month, year=year)

        # Handle LA Treasury Amounts input (replacing M2M file upload)
        # Collect all 11 amounts from form
        la_treasury_amounts = {
            'com_papers_cap': request.form.get("com_papers_cap", "").strip(),
            'com_papers_init_payable': request.form.get("com_papers_init_payable", "").strip(),
            'sec_laon_cap': request.form.get("sec_laon_cap", "").strip(),
            'sec_laon_int_payable': request.form.get("sec_laon_int_payable", "").strip(),
            'rel_party_loans_cap': request.form.get("rel_party_loans_cap", "").strip(),
            'rel_party_loans_int_pay': request.form.get("rel_party_loans_int_pay", "").strip(),
            'repo_cap': request.form.get("repo_cap", "").strip(),
            'repo_int_recei': request.form.get("repo_int_recei", "").strip(),
            'cost_invest': request.form.get("cost_invest", "").strip(),
            'int_recei': request.form.get("int_recei", "").strip(),
            'M2M_gain': request.form.get("M2M_gain", "").strip(),
        }

        # Check if any LA treasury amounts were provided
        has_any_amount = any(val for val in la_treasury_amounts.values())

        if has_any_amount:
            # Save amounts to Master_Data.xlsx NBD-WF-15-LA sheet
            try:
                master_data_file = Path(__file__).parent / "Master_Data.xlsx"
                wb = openpyxl.load_workbook(master_data_file)

                if "NBD-WF-15-LA" in wb.sheetnames:
                    ws = wb["NBD-WF-15-LA"]

                    # Write amounts to row 2 (row 1 has headers)
                    col_mapping = {
                        'com_papers_cap': 1,  # Column A
                        'com_papers_init_payable': 2,  # Column B
                        'sec_laon_cap': 3,  # Column C
                        'sec_laon_int_payable': 4,  # Column D
                        'rel_party_loans_cap': 5,  # Column E
                        'rel_party_loans_int_pay': 6,  # Column F
                        'repo_cap': 7,  # Column G
                        'repo_int_recei': 8,  # Column H
                        'cost_invest': 9,  # Column I
                        'int_recei': 10,  # Column J
                        'M2M_gain': 11,  # Column K
                    }

                    for key, col in col_mapping.items():
                        value_str = la_treasury_amounts[key]
                        if value_str:
                            try:
                                # Convert to float
                                value = float(value_str)
                                ws.cell(row=2, column=col).value = value
                            except ValueError:
                                flash(f"Invalid number format for {key}: {value_str}", "error")
                                return redirect(url_for("master_data_form", **request.form.to_dict()))

                    wb.save(master_data_file)
                    wb.close()
                    _emit_log("Updated LA treasury amounts in Master_Data.xlsx", important=True)
                    la_uploaded_files["Treasury Amounts"] = "Saved to Master_Data.xlsx"
                else:
                    _emit_log("NBD-WF-15-LA sheet not found in Master_Data.xlsx", important=True)

            except Exception as e:
                flash(f"Error saving LA treasury amounts: {e}", "error")
                _emit_log(f"Error saving LA treasury amounts: {e}", important=True)
                return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Handle BankBalanceGLCodes upload
        if "bank_balance_gl_codes_file" in request.files:
            file = request.files["bank_balance_gl_codes_file"]
            if file and file.filename:
                # Save file to Master Inputs with fixed name
                category_dir = MASTER_INPUTS_DIR / "BankBalanceGLCodes"
                category_dir.mkdir(parents=True, exist_ok=True)

                filename = "BankBalanceGLCodes.xlsx"
                file_path = category_dir / filename

                # Remove old file if exists
                if file_path.exists():
                    try:
                        file_path.unlink()
                        _emit_log("Removed old BankBalanceGLCodes.xlsx file", important=True)
                    except Exception as e:
                        _emit_log(f"Warning: Could not remove old BankBalanceGLCodes.xlsx: {e}", important=True)

                file.save(str(file_path))
                la_uploaded_files["BankBalanceGLCodes"] = str(file_path)
                _emit_log(f"Uploaded BankBalanceGLCodes file: {filename}", important=True)

                # Copy to latest LA output folder
                try:
                    # Read completed_dates.txt to find last completed date
                    completed_dates_file = OUTPUTS_DIR / "NBD_WF_15_LA" / "completed_dates.txt"
                    if completed_dates_file.exists():
                        with open(completed_dates_file, 'r') as f:
                            lines = [line.strip() for line in f.readlines() if line.strip()]
                            if lines:
                                last_date = lines[-1]
                                _emit_log(f"Last completed LA date: {last_date}", important=True)

                                # Find the highest version folder for this date
                                la_output_base = OUTPUTS_DIR / "NBD_WF_15_LA"
                                target_folder = None

                                # Check for base folder and versioned folders
                                base_folder = la_output_base / last_date
                                if base_folder.exists():
                                    target_folder = base_folder
                                    max_version = 0

                                    # Look for versioned folders like date(1), date(2), etc.
                                    for folder in la_output_base.iterdir():
                                        if folder.is_dir() and folder.name.startswith(f"{last_date}("):
                                            try:
                                                version = int(folder.name.split('(')[1].rstrip(')'))
                                                if version > max_version:
                                                    max_version = version
                                                    target_folder = folder
                                            except:
                                                pass

                                if target_folder:
                                    # Remove old BankBalanceGLCodes.xlsx in target folder
                                    old_file = target_folder / "BankBalanceGLCodes.xlsx"
                                    if old_file.exists():
                                        try:
                                            old_file.unlink()
                                            _emit_log(f"Removed old BankBalanceGLCodes.xlsx from {target_folder.name}", important=True)
                                        except Exception as e:
                                            _emit_log(f"Warning: Could not remove old file in output: {e}", important=True)

                                    # Copy new file to target folder
                                    target_file = target_folder / filename
                                    shutil.copy2(file_path, target_file)
                                    _emit_log(f"Copied BankBalanceGLCodes.xlsx to {target_folder.name}", important=True)
                                else:
                                    _emit_log(f"Could not find output folder for date {last_date}", important=True)
                            else:
                                _emit_log("No completed dates found in completed_dates.txt", important=True)
                    else:
                        _emit_log("completed_dates.txt not found", important=True)
                except Exception as e:
                    _emit_log(f"Warning: Could not copy to output folder: {e}", important=True)

    # ── ECS Section ──────────────────────────────────────────────────────────
    show_ecs = request.form.get("_show_ecs") == "1"
    if show_ecs:
        try:
            ecs_outputs_dir = OUTPUTS_DIR / "NBD_WF_18_ECS"
            _, ecs_date_folder = find_latest_completed_date(ecs_outputs_dir)
        except Exception as e:
            _emit_log(f"ECS: could not determine outputs date folder: {e}", important=True)
            ecs_date_folder = None

        # AFL Borrowing Schedule
        ecs_afl_file = request.files.get("ecs_afl_borrowing_file")
        if ecs_afl_file and ecs_afl_file.filename:
            borrowing_date = request.form.get("ecs_borrowing_schedule_date", "").strip()
            if not borrowing_date:
                flash("Date is required when uploading AFL Borrowing schedule.", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))
            try:
                date_obj = datetime.strptime(borrowing_date, "%Y-%m-%d")
                formatted_date = date_obj.strftime("%d-%m-%Y")
            except ValueError:
                flash("Invalid date format for AFL Borrowing schedule.", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))
            try:
                afl_dir = MASTER_INPUTS_DIR / "AFL Borrowing schedule"
                afl_dir.mkdir(parents=True, exist_ok=True)
                filename = f"AFL Borrowing schedule {formatted_date}.xlsx"
                dest_master = afl_dir / filename
                ecs_afl_file.save(str(dest_master))
                _emit_log(f"Saved AFL Borrowing schedule: {filename}", important=True)
                if ecs_date_folder and ecs_date_folder.exists():
                    for old in ecs_date_folder.glob("AFL Borrowing schedule *.xlsx"):
                        old.unlink()
                    shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                    _emit_log(f"Copied AFL Borrowing schedule to ECS outputs: {filename}", important=True)
            except Exception as e:
                _emit_log(f"ECS AFL Borrowing schedule upload failed: {e}", important=True)
                flash(f"AFL Borrowing schedule upload error: {e}", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Cash Flow
        ecs_cf_file = request.files.get("ecs_cash_flow_file")
        if ecs_cf_file and ecs_cf_file.filename:
            cf_month = request.form.get("ecs_cash_flow_month", "").strip()
            cf_year = request.form.get("ecs_cash_flow_year", "").strip()
            if not cf_month or not cf_year:
                flash("Month and Year are required when uploading Cash Flow file.", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))
            if not cf_year.isdigit() or len(cf_year) != 4:
                flash(f"Invalid year '{cf_year}' for Cash Flow. Enter a 4-digit year.", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))
            year_int = int(cf_year)
            if year_int < 2000 or year_int > 2100:
                flash(f"Year {cf_year} is out of range (2000–2100).", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))
            try:
                cf_dir = MASTER_INPUTS_DIR / "Cash Flow"
                cf_dir.mkdir(parents=True, exist_ok=True)
                filename = f"Cash Flow - {cf_month} {cf_year}.xlsx"
                dest_master = cf_dir / filename
                ecs_cf_file.save(str(dest_master))
                _emit_log(f"Saved Cash Flow: {filename}", important=True)
                if ecs_date_folder and ecs_date_folder.exists():
                    for old in ecs_date_folder.glob("Cash Flow *.xlsx"):
                        old.unlink()
                    shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                    _emit_log(f"Copied Cash Flow to ECS outputs: {filename}", important=True)
            except Exception as e:
                _emit_log(f"ECS Cash Flow upload failed: {e}", important=True)
                flash(f"Cash Flow upload error: {e}", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Payments and Receipts
        ecs_pr_file = request.files.get("ecs_payments_receipts_file")
        if ecs_pr_file and ecs_pr_file.filename:
            try:
                pr_dir = MASTER_INPUTS_DIR / "Payments and Receipts"
                pr_dir.mkdir(parents=True, exist_ok=True)
                filename = "Payments and Receipts.xlsx"
                dest_master = pr_dir / filename
                if dest_master.exists():
                    dest_master.unlink()
                ecs_pr_file.save(str(dest_master))
                _emit_log("Saved Payments and Receipts.xlsx", important=True)
                if ecs_date_folder and ecs_date_folder.exists():
                    old_pr = ecs_date_folder / filename
                    if old_pr.exists():
                        old_pr.unlink()
                    shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                    _emit_log("Copied Payments and Receipts to ECS outputs", important=True)
            except Exception as e:
                _emit_log(f"ECS Payments and Receipts upload failed: {e}", important=True)
                flash(f"Payments and Receipts upload error: {e}", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Daily Bank Balances
        ecs_dbb_file = request.files.get("ecs_daily_bank_balances_file")
        if ecs_dbb_file and ecs_dbb_file.filename:
            dbb_month = request.form.get("ecs_dbb_month", "").strip()
            dbb_year = request.form.get("ecs_dbb_year", "").strip()
            if not dbb_month or not dbb_year:
                flash("Month and Year are required when uploading Daily Bank Balances file.", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))
            if not dbb_year.isdigit() or len(dbb_year) != 4:
                flash(f"Invalid year '{dbb_year}' for Daily Bank Balances. Enter a 4-digit year.", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))
            try:
                dbb_dir = MASTER_INPUTS_DIR / "Daily Bank Balances"
                dbb_dir.mkdir(parents=True, exist_ok=True)
                filename = f"Daily Bank Balances - {dbb_month} {dbb_year}.xlsx"
                dest_master = dbb_dir / filename
                ecs_dbb_file.save(str(dest_master))
                _emit_log(f"Saved Daily Bank Balances: {filename}", important=True)
                if ecs_date_folder and ecs_date_folder.exists():
                    for old in ecs_date_folder.glob("Daily Bank Balances*.xlsx"):
                        old.unlink()
                    shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                    _emit_log(f"Copied Daily Bank Balances to ECS outputs: {filename}", important=True)
            except Exception as e:
                _emit_log(f"ECS Daily Bank Balances upload failed: {e}", important=True)
                flash(f"Daily Bank Balances upload error: {e}", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Payments (separate from "Payments and Receipts")
        ecs_pay_enabled = request.form.get("ecs_payments_upload_enabled") == "1"
        if ecs_pay_enabled:
            ecs_pay_file = request.files.get("ecs_payments_file")
            if ecs_pay_file and ecs_pay_file.filename:
                try:
                    pay_dir = MASTER_INPUTS_DIR / "Payments"
                    pay_dir.mkdir(parents=True, exist_ok=True)
                    today_str = datetime.today().strftime("%d-%m-%Y")
                    filename = f"Payments - {today_str}.xlsx"
                    dest_master = pay_dir / filename
                    ecs_pay_file.save(str(dest_master))
                    _emit_log(f"Saved Payments: {filename}", important=True)
                    if ecs_date_folder and ecs_date_folder.exists():
                        for old in ecs_date_folder.glob("Payments -*.xlsx"):
                            old.unlink()
                            _emit_log(f"Removed old Payments: {old.name}", important=True)
                        shutil.copy2(str(dest_master), str(ecs_date_folder / filename))
                        _emit_log(f"Copied Payments to ECS outputs: {filename}", important=True)
                    else:
                        _emit_log("No ECS date folder found; Payments saved to Master Inputs only", important=True)
                except Exception as e:
                    _emit_log(f"ECS Payments upload failed: {e}", important=True)
                    flash(f"Payments upload error: {e}", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Types of Receipt
        ecs_tor_enabled = request.form.get("ecs_types_of_receipt_upload_enabled") == "1"
        if ecs_tor_enabled:
            ecs_tor_file = request.files.get("ecs_types_of_receipt_file")
            if ecs_tor_file and ecs_tor_file.filename:
                try:
                    tor_dir = MASTER_INPUTS_DIR / "Types of Receipt"
                    tor_dir.mkdir(parents=True, exist_ok=True)
                    dest_master = tor_dir / "Types_of_Receipt.xlsx"
                    if dest_master.exists():
                        dest_master.unlink()
                    ecs_tor_file.save(str(dest_master))
                    _emit_log("Saved Types_of_Receipt.xlsx to Master Inputs", important=True)
                    if ecs_date_folder and ecs_date_folder.exists():
                        old_tor = ecs_date_folder / "Types_of_Receipt.xlsx"
                        if old_tor.exists():
                            old_tor.unlink()
                            _emit_log("Removed old Types_of_Receipt.xlsx from ECS outputs", important=True)
                        shutil.copy2(str(dest_master), str(ecs_date_folder / "Types_of_Receipt.xlsx"))
                        _emit_log("Copied Types_of_Receipt.xlsx to ECS outputs", important=True)
                    else:
                        _emit_log("No ECS date folder found; Types of Receipt saved to Master Inputs only", important=True)
                except Exception as e:
                    _emit_log(f"ECS Types of Receipt upload failed: {e}", important=True)
                    flash(f"Types of Receipt upload error: {e}", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

        # SasiaNet (optional)
        ecs_sn_file = request.files.get("ecs_sasianet_file")
        if ecs_sn_file and ecs_sn_file.filename:
            try:
                sn_dir = MASTER_INPUTS_DIR / "SasiaNet"
                sn_dir.mkdir(parents=True, exist_ok=True)
                today_str = datetime.today().strftime("%d-%m-%Y")
                sn_filename = f"SasiaNet - {today_str}.xlsx"
                dest_master = sn_dir / sn_filename
                ecs_sn_file.save(str(dest_master))
                _emit_log(f"Saved SasiaNet: {sn_filename}", important=True)
                if ecs_date_folder and ecs_date_folder.exists():
                    for old in ecs_date_folder.glob("SasiaNet - *.xlsx"):
                        old.unlink()
                        _emit_log(f"Removed old SasiaNet: {old.name}", important=True)
                    shutil.copy2(str(dest_master), str(ecs_date_folder / sn_filename))
                    _emit_log(f"Copied SasiaNet to ECS outputs: {sn_filename}", important=True)
                else:
                    _emit_log("No ECS date folder found; SasiaNet saved to Master Inputs only", important=True)
            except Exception as e:
                _emit_log(f"ECS SasiaNet upload failed: {e}", important=True)
                flash(f"SasiaNet upload error: {e}", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Handle Tax Payments upload (optional, checkbox-gated)
        ecs_tp_file = request.files.get("ecs_tax_payments_file")
        if ecs_tp_file and ecs_tp_file.filename:
            try:
                tp_dir = MASTER_INPUTS_DIR / "Tax Payments"
                tp_dir.mkdir(parents=True, exist_ok=True)
                today_str = datetime.today().strftime("%d-%m-%Y")
                tp_filename = f"Tax Payments - {today_str}.xlsx"
                dest_master_tp = tp_dir / tp_filename
                # Remove old Tax Payments files in Master Inputs
                for old_tp in tp_dir.glob("Tax Payments - *.xlsx"):
                    old_tp.unlink(missing_ok=True)
                ecs_tp_file.save(str(dest_master_tp))
                _emit_log(f"Saved Tax Payments: {tp_filename}", important=True)
                if ecs_date_folder and ecs_date_folder.exists():
                    for old_tp_out in ecs_date_folder.glob("Tax Payments - *.xlsx"):
                        old_tp_out.unlink(missing_ok=True)
                        _emit_log(f"Removed old Tax Payments from outputs: {old_tp_out.name}", important=True)
                    shutil.copy2(str(dest_master_tp), str(ecs_date_folder / tp_filename))
                    _emit_log(f"Copied Tax Payments to ECS outputs: {tp_filename}", important=True)
                else:
                    _emit_log("No ECS date folder found; Tax Payments saved to Master Inputs only", important=True)
            except Exception as e:
                _emit_log(f"ECS Tax Payments upload failed: {e}", important=True)
                flash(f"Tax Payments upload error: {e}", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))

        # Cashflow Movement (optional)
        ecs_cm_file = request.files.get("ecs_cashflow_movement_file")
        if ecs_cm_file and ecs_cm_file.filename:
            try:
                cm_dir = MASTER_INPUTS_DIR / "Cashflow Movement"
                cm_dir.mkdir(parents=True, exist_ok=True)
                now_str = datetime.now().strftime("%d-%m-%Y %H-%M-%S")
                cm_filename = f"Cashflow Movement - {now_str}.xlsx"
                dest_master_cm = cm_dir / cm_filename
                ecs_cm_file.save(str(dest_master_cm))
                _emit_log(f"Saved Cashflow Movement: {cm_filename}", important=True)
                if ecs_date_folder and ecs_date_folder.exists():
                    for old_cm in ecs_date_folder.glob("Cashflow Movement - *.xlsx"):
                        old_cm.unlink(missing_ok=True)
                        _emit_log(f"Removed old Cashflow Movement: {old_cm.name}", important=True)
                    shutil.copy2(str(dest_master_cm), str(ecs_date_folder / cm_filename))
                    _emit_log(f"Copied Cashflow Movement to ECS outputs: {cm_filename}", important=True)
                else:
                    _emit_log("No ECS date folder found; Cashflow Movement saved to Master Inputs only", important=True)
            except Exception as e:
                _emit_log(f"ECS Cashflow Movement upload failed: {e}", important=True)
                flash(f"Cashflow Movement upload error: {e}", "error")
                return redirect(url_for("master_data_form", **request.form.to_dict()))

    # Handle ALCL Management Accounts upload
    if "alclW_mg_accounts_file" in request.files:
            file = request.files["alclW_mg_accounts_file"]
            if file and file.filename:
                # Get month and year
                month = request.form.get("alclW_mg_accounts_month", "").strip()
                year = request.form.get("alclW_mg_accounts_year", "").strip()

                # Month and Year are REQUIRED when file is uploaded
                if not month or not year:
                    flash(f"Month and Year are required when uploading ALCL Management Accounts file.", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

                # Validate year
                if not year.isdigit() or len(year) != 4:
                    flash(f"Invalid year for ALCL Management Accounts: {year}. Please enter a valid 4-digit year (e.g., 2025).", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))
                year_int = int(year)
                if year_int < 2000 or year_int > 2100:
                    flash(f"Invalid year for ALCL Management Accounts: {year}. Please enter a year between 2000 and 2100.", "error")
                    return redirect(url_for("master_data_form", **request.form.to_dict()))

                la_months["alclW_mg_accounts_month"] = month
                la_years["alclW_mg_accounts_year"] = year

                # Save file to Master Inputs
                category_dir = MASTER_INPUTS_DIR / "ALCL Management Accounts"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: ALCL Management Accounts August_2025.xlsx
                filename = f"ALCL Management Accounts {month}_{year}.xlsx"

                file_path = category_dir / filename
                file.save(str(file_path))
                la_uploaded_files["ALCL Management Accounts"] = str(file_path)
                _emit_log(f"Uploaded ALCL Management Accounts file: {filename}", important=True)

                # Use centralized master data update system
                update_master_data_across_reports("ALCL Management Accounts", file_path, month=month, year=year)

    # Handle file uploads for IA
    _emit_log("*** IA FILE UPLOAD SECTION REACHED ***", important=True)
    ia_uploaded_files = {}
    ia_months = {}
    ia_years = {}
    if show_ia:
        categories = [
            "Disbursement with Budget",
            "Information Request from Credit",
            "Net Portfolio",
            "YARD STOCK AS AT"
        ]
        
        for category in categories:
            field_name = category.lower().replace(' ', '_')
            
            # Special handling for YARD STOCK AS AT (uses date instead of month/year)
            if category == "YARD STOCK AS AT":
                date_field = "yard_stock_date"  # Use the actual form field name
                date = request.form.get(date_field, "").strip()
                _emit_log(f"YARD STOCK AS AT - Date field: {date_field}, Date value: {date}", important=True)
                
                # Handle file upload
                if "yard_stock_file" in request.files:  # Use the actual form field name
                    file = request.files["yard_stock_file"]
                    _emit_log(f"YARD STOCK AS AT - File found: {file.filename if file else 'None'}", important=True)
                    if file and file.filename:
                        # Date is REQUIRED when file is uploaded
                        if not date:
                            flash(f"Date is required when uploading {category} file.", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))

                        # Save file to Master Inputs directory
                        category_dir = MASTER_INPUTS_DIR / category
                        category_dir.mkdir(parents=True, exist_ok=True)

                        # Generate filename with date (format: YARD STOCK AS AT 31-07-2025 FINAL.xlsx)
                        date_obj = datetime.strptime(date, "%Y-%m-%d")
                        formatted_date = date_obj.strftime("%d-%m-%Y")
                        filename = f"YARD STOCK AS AT {formatted_date} FINAL.xlsx"

                        file_path = category_dir / filename
                        _emit_log(f"YARD STOCK AS AT - Saving file to: {file_path}", important=True)
                        file.save(str(file_path))
                        ia_uploaded_files[category] = str(file_path)
                        _emit_log(f"Uploaded {category} file: {filename}", important=True)

                        # Use centralized master data update system
                        update_master_data_across_reports(category, file_path, date=date)
            else:
                # Regular handling for other categories (month/year based)
                month_field = f"{field_name}_month"
                year_field = f"{field_name}_year"

                # Get month and year selection
                month = request.form.get(month_field, "").strip()
                year = request.form.get(year_field, "").strip()

                # Handle file upload
                if f"{field_name}_file" in request.files:
                    file = request.files[f"{field_name}_file"]
                    if file and file.filename:
                        # Month and Year are REQUIRED when file is uploaded
                        if not month or not year:
                            flash(f"Month and Year are required when uploading {category} file.", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))

                        # Validate year - must be exactly 4 digits, no symbols or characters
                        if not year.isdigit() or len(year) != 4:
                            flash(f"Invalid year for {category}: {year}. Please enter a valid 4-digit year (e.g., 2025).", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))
                        year_int = int(year)
                        if year_int < 2000 or year_int > 2100:
                            flash(f"Invalid year for {category}: {year}. Please enter a year between 2000 and 2100.", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))

                        # Save month/year for later use
                        ia_months[month_field] = month
                        ia_years[year_field] = year

                        # Save file to Master Inputs directory
                        category_dir = MASTER_INPUTS_DIR / category
                        category_dir.mkdir(parents=True, exist_ok=True)

                        # Generate filename with month and year (exact format as specified)
                        if category == "Disbursement with Budget":
                            full_month = get_full_month_name(month)
                            filename = f"Disbursement with Budget - {full_month} {year}.xlsx"
                        elif category == "Information Request from Credit":
                            full_month = get_full_month_name(month)
                            filename = f"Information Request from Credit-{full_month} {year}.xlsx"
                        elif category == "Net Portfolio":
                            full_month = get_full_month_name(month)
                            filename = f"Net Portfolio-{full_month}-{year}.xlsx"
                        else:
                            filename = file.filename

                        file_path = category_dir / filename
                        file.save(str(file_path))
                        ia_uploaded_files[category] = str(file_path)
                        _emit_log(f"Uploaded {category} file: {filename}", important=True)

                        # Use centralized master data update system
                        update_master_data_across_reports(category, file_path, month=month, year=year)
                    elif month and year:
                        # Month/year provided but no file - validate and save
                        # Validate year - must be exactly 4 digits, no symbols or characters
                        if not year.isdigit() or len(year) != 4:
                            flash(f"Invalid year for {category}: {year}. Please enter a valid 4-digit year (e.g., 2025).", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))
                        year_int = int(year)
                        if year_int < 2000 or year_int > 2100:
                            flash(f"Invalid year for {category}: {year}. Please enter a year between 2000 and 2100.", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))

                        ia_months[month_field] = month
                        ia_years[year_field] = year

        # Handle Product Category file upload (special case - no month/year, just file)
        if "product_category_file" in request.files:
            file = request.files["product_category_file"]
            if file and file.filename:
                _emit_log("Processing Product Category file upload", important=True)

                # Save file to Master Inputs/Prod Cat directory
                prod_cat_dir = MASTER_INPUTS_DIR / "Prod Cat"
                prod_cat_dir.mkdir(parents=True, exist_ok=True)

                # Rename file to Product_Cat.xlsx (change entire name)
                filename = "Product_Cat.xlsx"
                file_path = prod_cat_dir / filename

                # Remove existing file if it exists
                if file_path.exists():
                    file_path.unlink()
                    _emit_log(f"Removed existing Product Category file: {filename}", important=True)

                # Save new file
                file.save(str(file_path))
                ia_uploaded_files["Product Category"] = str(file_path)
                _emit_log(f"Uploaded Product Category file: {filename}", important=True)

                # Copy to outputs/NBD_MF_23_IA/{date}/ folder
                # Find the last completed date from completed_dates.txt
                ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
                if ia_outputs_dir.exists():
                    try:
                        # Read completed_dates.txt to find last date
                        completed_dates_file = ia_outputs_dir / "completed_dates.txt"
                        if completed_dates_file.exists():
                            with open(completed_dates_file, 'r') as f:
                                dates = [line.strip() for line in f if line.strip()]
                            if dates:
                                last_date = dates[-1]  # Get last line
                                _emit_log(f"Last completed IA date: {last_date}", important=True)

                                # Find the max version date folder for that date
                                date_folders = [d for d in ia_outputs_dir.iterdir() if d.is_dir() and d.name.startswith(last_date)]
                                if date_folders:
                                    # Sort by modification time to get the latest
                                    target_folder = max(date_folders, key=lambda x: x.stat().st_mtime)
                                    _emit_log(f"Copying Product Category to: {target_folder}", important=True)

                                    # Remove existing Product_Cat.xlsx in target folder
                                    target_file = target_folder / "Product_Cat.xlsx"
                                    if target_file.exists():
                                        target_file.unlink()
                                        _emit_log(f"Removed existing Product_Cat.xlsx from {target_folder.name}", important=True)

                                    # Copy the new file
                                    shutil.copy2(file_path, target_file)
                                    _emit_log(f"Copied Product_Cat.xlsx to {target_folder.name}", important=True)
                                else:
                                    _emit_log(f"No date folder found for {last_date} in IA outputs", important=True)
                        else:
                            _emit_log("completed_dates.txt not found in IA outputs", important=True)
                    except Exception as e:
                        _emit_log(f"Error copying Product Category to outputs: {e}", important=True)
                else:
                    _emit_log("IA outputs directory not found", important=True)

        # Handle Initial Valuation file upload (special case - no month/year, just file)
        if "initial_valuation_file" in request.files:
            file = request.files["initial_valuation_file"]
            if file and file.filename:
                _emit_log("Processing Initial Valuation file upload", important=True)

                # Save file to IA working folder (same location where script expects it)
                ia_outputs_dir = OUTPUTS_DIR / "NBD_MF_23_IA"
                if ia_outputs_dir.exists():
                    try:
                        # Read completed_dates.txt to find last date
                        completed_dates_file = ia_outputs_dir / "completed_dates.txt"
                        if completed_dates_file.exists():
                            with open(completed_dates_file, 'r') as f:
                                dates = [line.strip() for line in f if line.strip()]
                            if dates:
                                last_date = dates[-1]
                                _emit_log(f"Last completed IA date: {last_date}", important=True)

                                # Find the max version date folder for that date
                                date_folders = [d for d in ia_outputs_dir.iterdir() if d.is_dir() and d.name.startswith(last_date)]
                                if date_folders:
                                    target_folder = max(date_folders, key=lambda x: x.stat().st_mtime)
                                    _emit_log(f"Saving Initial Valuation file to: {target_folder}", important=True)

                                    # Save as initial_valuation.xlsx
                                    filename = "initial_valuation.xlsx"
                                    file_path = target_folder / filename

                                    # Remove existing file if it exists
                                    if file_path.exists():
                                        file_path.unlink()
                                        _emit_log(f"Removed existing Initial Valuation file: {filename}", important=True)

                                    # Save new file
                                    file.save(str(file_path))
                                    ia_uploaded_files["Initial Valuation"] = str(file_path)
                                    _emit_log(f"Uploaded Initial Valuation file: {filename}", important=True)
                                else:
                                    _emit_log(f"No date folder found for {last_date} in IA outputs", important=True)
                                    flash(f"No date folder found for {last_date}. Run IA report first.", "warning")
                        else:
                            _emit_log("completed_dates.txt not found in IA outputs", important=True)
                            flash("IA report must be run at least once before uploading Initial Valuation file.", "warning")
                    except Exception as e:
                        _emit_log(f"Error saving Initial Valuation file: {e}", important=True)
                        flash(f"Error uploading Initial Valuation file: {e}", "error")
                else:
                    _emit_log("IA outputs directory not found", important=True)
                    flash("IA outputs directory not found. Run IA report first.", "warning")

    # Handle file uploads for C1-C2
    c1c2_uploaded_files = {}
    c1c2_dates = {}
    c1c2_months = {}
    c1c2_years = {}
    if show_c1c2:
        categories = [
            ("Cadre", "date"),
            ("Unutilized Amount", "month_year")
        ]
        
        for category, type_info in categories:
            field_name = category.lower().replace(' ', '_')
            
            if type_info == "date":
                # Handle Cadre with date picker
                date_field = f"{field_name}_date"
                date_value = request.form.get(date_field, "").strip()

                # Handle file upload
                if f"{field_name}_file" in request.files:
                    file = request.files[f"{field_name}_file"]
                    if file and file.filename:
                        # Date is REQUIRED when file is uploaded
                        if not date_value:
                            flash(f"Date is required when uploading {category} file.", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))

                        # Save date for later use
                        c1c2_dates[date_field] = date_value

                        # Save file to Master Inputs directory
                        category_dir = MASTER_INPUTS_DIR / category
                        category_dir.mkdir(parents=True, exist_ok=True)

                        # Generate filename with date (format: Cadre 31 07 2025.xlsx)
                        # Convert date from YYYY-MM-DD to DD MM YYYY format
                        date_obj = datetime.strptime(date_value, "%Y-%m-%d")
                        formatted_date = date_obj.strftime("%d %m %Y")
                        filename = f"Cadre {formatted_date}.xlsx"

                        file_path = category_dir / filename
                        file.save(str(file_path))
                        c1c2_uploaded_files[category] = str(file_path)
                        _emit_log(f"Uploaded {category} file: {filename}", important=True)

                        # Use centralized master data update system
                        update_master_data_across_reports(category, file_path, date=date_value)
                elif date_value:
                    # Date provided but no file - still save the date
                    c1c2_dates[date_field] = date_value
            
            elif type_info == "month_year":
                # Handle Unutilized Amount with month/year
                month_field = f"{field_name}_month"
                year_field = f"{field_name}_year"

                month = request.form.get(month_field, "").strip()
                year = request.form.get(year_field, "").strip()

                # Handle file upload
                if f"{field_name}_file" in request.files:
                    file = request.files[f"{field_name}_file"]
                    if file and file.filename:
                        # Month and Year are REQUIRED when file is uploaded
                        if not month or not year:
                            flash(f"Month and Year are required when uploading {category} file.", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))

                        # Validate year - must be exactly 4 digits, no symbols or characters
                        if not year.isdigit() or len(year) != 4:
                            flash(f"Invalid year for {category}: {year}. Please enter a valid 4-digit year (e.g., 2025).", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))
                        year_int = int(year)
                        if year_int < 2000 or year_int > 2100:
                            flash(f"Invalid year for {category}: {year}. Please enter a year between 2000 and 2100.", "error")
                            return redirect(url_for("master_data_form", **request.form.to_dict()))

                        # Save month/year for later use
                        c1c2_months[month_field] = month
                        c1c2_years[year_field] = year

                        # Save file to Master Inputs directory
                        category_dir = MASTER_INPUTS_DIR / category
                        category_dir.mkdir(parents=True, exist_ok=True)

                        # Generate filename with month and year (format: Unutilized Amount JULY 2025.xlsx)
                        full_month = get_full_month_name(month)
                        filename = f"Unutilized Amount {full_month.upper()} {year}.xlsx"

                        file_path = category_dir / filename
                        file.save(str(file_path))
                        c1c2_uploaded_files[category] = str(file_path)
                        _emit_log(f"Uploaded {category} file: {filename}", important=True)

                        # Use centralized master data update system
                        update_master_data_across_reports(category, file_path, month=month, year=year)
                elif month and year:
                    # Month/year provided but no file - validate and save
                    # Validate year - must be exactly 4 digits, no symbols or characters
                    if not year.isdigit() or len(year) != 4:
                        flash(f"Invalid year for {category}: {year}. Please enter a valid 4-digit year (e.g., 2025).", "error")
                        return redirect(url_for("master_data_form", **request.form.to_dict()))
                    year_int = int(year)
                    if year_int < 2000 or year_int > 2100:
                        flash(f"Invalid year for {category}: {year}. Please enter a year between 2000 and 2100.", "error")
                        return redirect(url_for("master_data_form", **request.form.to_dict()))

                    c1c2_months[month_field] = month
                    c1c2_years[year_field] = year

    # Collect ECS M2M, treasury and manual payments data
    ecs_m2m_data = None
    ecs_treasury_data = None
    ecs_manual_payments_data = None
    if show_ecs:
        def _parse_m2m(val):
            try:
                return float(val) if val.strip() else None
            except Exception:
                return None
        ecs_m2m_data = {f"day{i}": _parse_m2m(request.form.get(f"ecs_m2m_day{i}", "")) for i in range(1, 8)}

        ecs_treasury_data = {
            "com_paper_cap": request.form.get("com_paper_cap", ""),
            "com_paper_int_pay": request.form.get("com_paper_int_pay", ""),
            "sec_loan_cap": request.form.get("sec_loan_cap", ""),
            "sec_loan_int_pay": request.form.get("sec_loan_int_pay", ""),
            "repo_cap": request.form.get("repo_cap", ""),
            "repo_int_rec": request.form.get("repo_int_rec", ""),
            "cost_of_inv": request.form.get("cost_of_inv", ""),
            "int_rec": request.form.get("int_rec", ""),
            "m2m_gain": request.form.get("m2m_gain", ""),
        }

        def _parse_num(val):
            try:
                return float(val) if val.strip() else None
            except Exception:
                return None

        ecs_manual_payments_data = {
            "current_week":     _parse_num(request.form.get("ecs_mp_current_week", "")),
            "cbsl":             _parse_num(request.form.get("ecs_mp_cbsl", "")),
            "epf_etf_ot":       _parse_num(request.form.get("ecs_mp_epf_etf_ot", "")),
            "outsource_staff":  _parse_num(request.form.get("ecs_mp_outsource_staff", "")),
            "apiit":            _parse_num(request.form.get("ecs_mp_apiit", "")),
            "fd":               _parse_num(request.form.get("ecs_mp_fd", "")),
            "taxes":            _parse_num(request.form.get("ecs_mp_taxes", "")),
            "dividend_payment": _parse_num(request.form.get("ecs_mp_dividend_payment", "")),
            "sscl":             _parse_num(request.form.get("ecs_mp_sscl", "")),
        }

    # Save to Master_Data.xlsx
    success, msg = save_master_data(sofp_data, ia_data, c1c6_data, sofp_sheet=sofp_sheet, c1c6_sheet=c1c6_sheet, ga_is_data=ga_is_data, ecs_m2m_data=ecs_m2m_data, ecs_treasury_data=ecs_treasury_data, ecs_manual_payments_data=ecs_manual_payments_data)
    if not success:
        flash(f"Failed to save master data: {msg}", "error")
        return redirect(url_for("master_data_form"))
    
    # Get report data from session
    report_ids_with_dates = session.get('report_ids_with_dates', {})
    if not report_ids_with_dates:
        flash("No report data found. Please start again.", "error")
        return redirect(url_for("index"))

    # Clear report selection from session
    session.pop('report_ids_with_dates', None)
    
    # Note: File copying is now handled by the centralized update_master_data_across_reports() function
    # which is called during the individual file upload processing above
    
    # Validate SOFP file exists for IA automation (IA and C1C2 run together)
    if show_ia or show_c1c2:
        try:
            _emit_log("Validating SOFP file exists for IA automation (IA and C1C2 run together)", important=True)
            is_valid, validation_msg = validate_sofp_file_for_ia(report_ids_with_dates)
            if not is_valid:
                flash(f"Validation failed: {validation_msg}", "error")
                return redirect(url_for("index"))
        except Exception as e:
            _emit_log(f"Failed to validate SOFP file for IA: {e}", important=True)
            flash(f"Validation error: {e}", "error")
            return redirect(url_for("index"))
    
    # Validate SOFP file exists for GA IS automation
    if show_ga_is:
        try:
            _emit_log("Validating SOFP file exists for GA IS automation", important=True)
            is_valid, validation_msg = validate_sofp_file_for_ga_is(report_ids_with_dates)
            if not is_valid:
                flash(f"Validation failed: {validation_msg}", "error")
                return redirect(url_for("index"))
        except Exception as e:
            _emit_log(f"Failed to validate SOFP file for GA IS: {e}", important=True)
            flash(f"Validation error: {e}", "error")
            return redirect(url_for("index"))

    # Validate Prod. wise file exists for C8 automation
    if "NBD-QF-23-C8" in report_ids_with_dates:
        try:
            _emit_log("Validating Prod. wise Class. of Loans file exists for C8 automation", important=True)
            is_valid, validation_msg = validate_prod_wise_file_for_c8(report_ids_with_dates)
            if not is_valid:
                flash(f"Validation failed: {validation_msg}", "error")
                return redirect(url_for("index"))
        except Exception as e:
            _emit_log(f"Failed to validate Prod. wise file for C8: {e}", important=True)
            flash(f"Validation error: {e}", "error")
            return redirect(url_for("index"))

    # Validate Set 7 reports (C2-C6) - must select all together and check required files
    set7_reports = {"NBD-MF-20-C2", "NBD-MF-20-C3", "NBD-MF-20-C4", "NBD-MF-20-C5", "NBD-MF-20-C6"}
    selected_set7_reports = set(report_ids_with_dates.keys()) & set7_reports
    
    if selected_set7_reports:
        # Check if all Set 7 reports are selected
        if len(selected_set7_reports) != len(set7_reports):
            missing_reports = set7_reports - selected_set7_reports
            flash(f"Set 7 reports must be selected together. Missing: {', '.join(missing_reports)}", "error")
            return redirect(url_for("index"))
        
        # Validate required files exist
        try:
            _emit_log("Validating required files for Set 7 automation", important=True)
            is_valid, validation_msg = validate_set7_required_files(report_ids_with_dates)
            if not is_valid:
                flash(f"Validation failed: {validation_msg}", "error")
                return redirect(url_for("index"))
        except Exception as e:
            _emit_log(f"Failed to validate Set 7 required files: {e}", important=True)
            flash(f"Validation error: {e}", "error")
            return redirect(url_for("index"))

    # Validate Deposit Liability file exists for ECS automation
    if "NBD-WF-18-ECS" in report_ids_with_dates:
        try:
            _emit_log("Validating Deposit Liability file exists for ECS automation", important=True)
            is_valid, validation_msg = validate_deposit_liability_file_for_ecs(report_ids_with_dates)
            if not is_valid:
                flash(f"Validation failed: {validation_msg}", "error")
                return redirect(url_for("index"))
        except Exception as e:
            _emit_log(f"Failed to validate Deposit Liability file for ECS: {e}", important=True)
            flash(f"Validation error: {e}", "error")
            return redirect(url_for("index"))

    # Validate Deposit Liability file exists for LA automation
    if "NBD-WF-15-LA" in report_ids_with_dates:
        try:
            _emit_log("Validating Deposit Liability file exists for LA automation", important=True)
            is_valid, validation_msg = validate_deposit_liability_file_for_la(report_ids_with_dates)
            if not is_valid:
                flash(f"Validation failed: {validation_msg}", "error")
                return redirect(url_for("index"))
        except Exception as e:
            _emit_log(f"Failed to validate Deposit Liability file for LA: {e}", important=True)
            flash(f"Validation error: {e}", "error")
            return redirect(url_for("index"))

        # Copy Daily Bank Balances from Master Inputs to LA outputs folder
        try:
            _emit_log("Copying Daily Bank Balances to LA outputs folder...", important=True)
            daily_bank_dir = MASTER_INPUTS_DIR / "Daily Bank Balances"
            if daily_bank_dir.exists():
                daily_bank_files = list(daily_bank_dir.glob("Daily Bank Balances*.xlsx"))
                if daily_bank_files:
                    # Get the most recent file (by modification time)
                    latest_file = max(daily_bank_files, key=lambda f: f.stat().st_mtime)

                    # Find LA outputs folder using completed_dates.txt
                    la_outputs_dir = OUTPUTS_DIR / "NBD_WF_15_LA"
                    if la_outputs_dir.exists():
                        completed_dates_file = la_outputs_dir / "completed_dates.txt"
                        if completed_dates_file.exists():
                            with open(completed_dates_file, 'r') as f:
                                lines = [line.strip() for line in f.readlines() if line.strip()]
                            if lines:
                                target_date = lines[-1]  # Last date

                                # Find the max version folder for this date
                                matching_folders = []
                                for folder in la_outputs_dir.iterdir():
                                    if folder.is_dir():
                                        folder_name = folder.name
                                        if folder_name == target_date or folder_name.startswith(f"{target_date}("):
                                            matching_folders.append(folder)

                                if matching_folders:
                                    # Get max version folder
                                    def get_version(folder):
                                        name = folder.name
                                        if name == target_date:
                                            return 0
                                        if "(" in name and name.endswith(")"):
                                            try:
                                                return int(name.split("(")[-1].rstrip(")"))
                                            except ValueError:
                                                return 0
                                        return 0

                                    target_folder = max(matching_folders, key=get_version)

                                    # Remove existing Daily Bank Balances files
                                    for existing in target_folder.glob("Daily Bank Balances*.xlsx"):
                                        try:
                                            existing.unlink()
                                            _emit_log(f"Removed existing: {existing.name}", important=True)
                                        except Exception as e:
                                            _emit_log(f"Could not remove {existing.name}: {e}", important=True)

                                    # Copy the file
                                    dest_path = target_folder / latest_file.name
                                    shutil.copy2(latest_file, dest_path)
                                    _emit_log(f"✓ Copied Daily Bank Balances to LA outputs: {dest_path}", important=True)
                                else:
                                    _emit_log(f"No folder found for date {target_date} in LA outputs", important=True)
                            else:
                                _emit_log("completed_dates.txt is empty", important=True)
                        else:
                            _emit_log("completed_dates.txt not found in LA outputs", important=True)
                    else:
                        _emit_log("LA outputs directory not found", important=True)
                else:
                    _emit_log("No Daily Bank Balances file found in Master Inputs", important=True)
            else:
                _emit_log("Daily Bank Balances Master Inputs directory not found", important=True)
        except Exception as e:
            _emit_log(f"Warning: Failed to copy Daily Bank Balances to LA outputs: {e}", important=True)
            # Continue anyway - not a critical error

    # Handle QF Reports file uploads
    _emit_log("*** QF REPORTS FILE UPLOAD SECTION REACHED ***", important=True)
    qf_uploaded_files = {}
    qf_months = {}
    qf_years = {}
    qf_dates = {}
    show_qf = request.form.get("_show_qf") == "1"

    if show_qf:
        # 1. CBSL Sector Request (file only, no date)
        if "cbsl_sector_request_file" in request.files:
            file = request.files["cbsl_sector_request_file"]
            if file and file.filename:
                category_dir = MASTER_INPUTS_DIR / "CBSL Sector Request"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Keep original filename
                filename = file.filename
                file_path = category_dir / filename
                file.save(str(file_path))
                qf_uploaded_files["CBSL Sector Request"] = str(file_path)
                _emit_log(f"Uploaded CBSL Sector Request file: {filename}", important=True)

        # 2. JOBHD (file + date with timestamp format YYYYMMDD_HHMMSS)
        if "jobhd_file" in request.files:
            file = request.files["jobhd_file"]
            if file and file.filename:
                jobhd_date = request.form.get("jobhd_date", "").strip()

                category_dir = MASTER_INPUTS_DIR / "JOBHD"
                category_dir.mkdir(parents=True, exist_ok=True)

                # If date provided, format as YYYYMMDD_HHMMSS, else use current timestamp
                if jobhd_date:
                    # from datetime import datetime
                    date_obj = datetime.strptime(jobhd_date, "%Y-%m-%d")
                    timestamp = date_obj.strftime("%Y%m%d_000000")
                    qf_dates["jobhd_date"] = jobhd_date
                else:
                    # from datetime import datetime
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                # Generate filename: JOBHD_YYYYMMDD_HHMMSS.xlsx
                filename = f"JOBHD_{timestamp}.xlsx"
                file_path = category_dir / filename
                file.save(str(file_path))
                qf_uploaded_files["JOBHD"] = str(file_path)
                _emit_log(f"Uploaded JOBHD file: {filename}", important=True)

        # 3. Sector Report (file + month + year)
        if "sector_report_file" in request.files:
            file = request.files["sector_report_file"]
            if file and file.filename:
                month = request.form.get("sector_report_month", "").strip()
                year = request.form.get("sector_report_year", "").strip()

                category_dir = MASTER_INPUTS_DIR / "Sector Report"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: Sector Report - Month Year.xlsb (always use .xlsb)
                if month and year:
                    filename = f"Sector Report - {month} {year}.xlsb"
                    qf_months["sector_report_month"] = month
                    qf_years["sector_report_year"] = year
                else:
                    filename = file.filename

                file_path = category_dir / filename
                file.save(str(file_path))
                qf_uploaded_files["Sector Report"] = str(file_path)
                _emit_log(f"Uploaded Sector Report file: {filename}", important=True)

        # 4. Business Gender (file + month + year)
        if "business_gender_file" in request.files:
            file = request.files["business_gender_file"]
            if file and file.filename:
                month = request.form.get("business_gender_month", "").strip()
                year = request.form.get("business_gender_year", "").strip()

                category_dir = MASTER_INPUTS_DIR / "Business Gender"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: Business Gender - Month Year.xlsx
                if month and year:
                    filename = f"Business Gender - {month} {year}.xlsx"
                    qf_months["business_gender_month"] = month
                    qf_years["business_gender_year"] = year
                else:
                    filename = f"Business Gender - {file.filename}"

                file_path = category_dir / filename
                file.save(str(file_path))
                qf_uploaded_files["Business Gender"] = str(file_path)
                _emit_log(f"Uploaded Business Gender file: {filename}", important=True)

        # 5. Information Request from Credit (SHARED - file + month + year)
        if "info_request_credit_file" in request.files:
            file = request.files["info_request_credit_file"]
            if file and file.filename:
                month = request.form.get("info_request_credit_month", "").strip()
                year = request.form.get("info_request_credit_year", "").strip()

                # This is shared with IA, use existing "Information Request from Credit" folder
                category_dir = MASTER_INPUTS_DIR / "Information Request from Credit"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: Information Request from Credit Month Year.xlsx
                if month and year:
                    filename = f"Information Request from Credit {month} {year}.xlsx"
                    qf_months["info_request_credit_month"] = month
                    qf_years["info_request_credit_year"] = year
                else:
                    filename = file.filename

                file_path = category_dir / filename
                file.save(str(file_path))
                qf_uploaded_files["Information Request from Credit"] = str(file_path)
                _emit_log(f"Uploaded Information Request from Credit file: {filename}", important=True)

        # 6. Information Request from IT (file + month + year)
        if "info_request_it_file" in request.files:
            file = request.files["info_request_it_file"]
            if file and file.filename:
                month = request.form.get("info_request_it_month", "").strip()
                year = request.form.get("info_request_it_year", "").strip()

                category_dir = MASTER_INPUTS_DIR / "Information Request from IT"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: Information Request from IT Month Year.xlsx
                if month and year:
                    filename = f"Information Request from IT {month} {year}.xlsx"
                    qf_months["info_request_it_month"] = month
                    qf_years["info_request_it_year"] = year
                else:
                    filename = file.filename

                file_path = category_dir / filename
                file.save(str(file_path))
                qf_uploaded_files["Information Request from IT"] = str(file_path)
                _emit_log(f"Uploaded Information Request from IT file: {filename}", important=True)

        # 7. Net Portfolio (SHARED - file + month + year)
        if "net_portfolio_file" in request.files:
            file = request.files["net_portfolio_file"]
            if file and file.filename:
                month = request.form.get("net_portfolio_month", "").strip()
                year = request.form.get("net_portfolio_year", "").strip()

                # This is shared with IA, use existing "Net Portfolio" folder
                category_dir = MASTER_INPUTS_DIR / "Net Portfolio"
                category_dir.mkdir(parents=True, exist_ok=True)

                # Generate filename: Net Portfolio Month Year.xlsx
                if month and year:
                    filename = f"Net Portfolio {month} {year}.xlsx"
                    qf_months["net_portfolio_month"] = month
                    qf_years["net_portfolio_year"] = year
                else:
                    filename = file.filename

                file_path = category_dir / filename
                file.save(str(file_path))
                qf_uploaded_files["Net Portfolio"] = str(file_path)
                _emit_log(f"Uploaded Net Portfolio file: {filename}", important=True)

        # 8. Client Available Limits (3 files in single folder with different dates)
        # First, remove all old files from the Client Available Limits folder
        category_dir = MASTER_INPUTS_DIR / "Client Available Limits"
        category_dir.mkdir(parents=True, exist_ok=True)

        # Collect all new files to upload
        new_files_uploaded = False
        for i in range(1, 4):
            file_field = f"client_avail_limits{i}_file"
            if file_field in request.files:
                file = request.files[file_field]
                if file and file.filename:
                    new_files_uploaded = True
                    break

        # If any new files are being uploaded, clear the folder first
        if new_files_uploaded:
            existing_files = list(category_dir.glob("*.xlsx")) + list(category_dir.glob("*.xls"))
            for existing_file in existing_files:
                try:
                    existing_file.unlink()
                    _emit_log(f"Removed old Client Available Limits file: {existing_file.name}", important=True)
                except Exception as e:
                    _emit_log(f"Failed to remove old file {existing_file.name}: {e}", important=True)

        # Now upload the new files
        for i in range(1, 4):
            file_field = f"client_avail_limits{i}_file"
            date_field = f"client_avail_limits{i}_date"

            if file_field in request.files:
                file = request.files[file_field]
                if file and file.filename:
                    date_value = request.form.get(date_field, "").strip()

                    # Generate filename: DD-MM-YYYY  Client Available Limits Report.xlsx
                    if date_value:
                        # Convert YYYY-MM-DD to DD-MM-YYYY format
                        # from datetime import datetime
                        date_obj = datetime.strptime(date_value, "%Y-%m-%d")
                        formatted_date = date_obj.strftime("%d-%m-%Y")
                        filename = f"{formatted_date}  Client Available Limits Report.xlsx"
                        qf_dates[date_field] = date_value
                    else:
                        filename = file.filename

                    file_path = category_dir / filename
                    file.save(str(file_path))
                    qf_uploaded_files[f"Client Available Limits {i}"] = str(file_path)
                    _emit_log(f"Uploaded Client Available Limits file {i}: {filename}", important=True)

        # Copy all QF uploaded files to outputs folder
        if qf_uploaded_files:
            _emit_log(f"Copying {len(qf_uploaded_files)} QF files to outputs folder", important=True)
            copy_uploaded_files_to_qf_outputs(qf_uploaded_files, qf_months, qf_years, qf_dates)

    # Start automation (ensure data is dict, not None)
    return start_automation_with_data(
        report_ids_with_dates,
        sofp_data or {},
        ia_data or {},
        uploaded_files,
        months
    )


@app.route("/mark-completed", methods=["POST"])
def mark_completed():
    """Mark a report set as completed by appending date to completed_dates.txt"""
    set_num = request.form.get("set_num", "").strip()
    step_num = request.form.get("step_num", "").strip()
    date_str = request.form.get("date", "").strip()

    if not set_num or not date_str:
        flash("Please select a set and date.", "error")
        return redirect(url_for("index"))

    try:
        set_num = int(set_num)

        # Convert date from YYYY-MM-DD to DD-MM-YYYY format
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        formatted_date = dt.strftime("%d-%m-%Y")

        # Get folder path(s) based on set number
        if set_num not in SET_COMPLETION_MAPPING:
            flash("Invalid set selected.", "error")
            return redirect(url_for("index"))

        folder_mapping = SET_COMPLETION_MAPPING[set_num]
        folders_to_update = []

        # Handle Set 2 which has multiple steps
        if isinstance(folder_mapping, dict):
            if step_num:
                step_num = int(step_num)
                if step_num in folder_mapping:
                    folders_to_update = [folder_mapping[step_num]]
                else:
                    flash("Invalid step selected.", "error")
                    return redirect(url_for("index"))
            else:
                # Step number is required for Set 2
                flash("Please select a specific step for Set 2.", "error")
                return redirect(url_for("index"))
        else:
            folders_to_update = [folder_mapping]

        # Update completed_dates.txt for each folder
        for folder_name in folders_to_update:
            completed_file = OUTPUTS_DIR / folder_name / "completed_dates.txt"

            # Ensure the directory exists
            completed_file.parent.mkdir(parents=True, exist_ok=True)

            # Append the date to completed_dates.txt
            with open(completed_file, "a", encoding="utf-8") as f:
                f.write(formatted_date + "\n")

        # When LA (Set 2 Step 2) is marked complete, copy AFL file to MF_04_LA outputs
        if set_num == 2 and step_num == 2:
            # Get the week selection from the form
            la_week = request.form.get("la_week", "").strip()

            if not la_week:
                flash("Warning: Please select which week you are completing (Week 1-4 or Month End).", "warning")
            else:
                try:
                    import shutil
                    # Find the AFL file in WF_LA outputs for the SPECIFIC date being marked
                    wf_la_outputs = OUTPUTS_DIR / "NBD_WF_15_LA"

                    # Look for the exact date folder (formatted_date = DD-MM-YYYY)
                    wf_la_date_folder = find_date_folder_with_variations(wf_la_outputs, formatted_date)

                    if wf_la_date_folder and wf_la_date_folder.exists():
                        afl_files = list(wf_la_date_folder.glob("AFL Liquidity*.xlsx"))
                        if afl_files:
                            afl_file = max(afl_files, key=lambda x: x.stat().st_mtime)

                            # Find or create MF_04_LA date folder - use LATEST completed date folder
                            mf04_outputs = OUTPUTS_DIR / "NBD_MF_04_LA"

                            # First try to find existing completed date folder
                            _, mf04_date_folder = find_latest_completed_date(mf04_outputs)

                            # If no folder exists, create one with the latest MF04 completion date
                            if not mf04_date_folder or not mf04_date_folder.exists():
                                # Use the latest date from MF04's completed_dates.txt, or create first folder
                                mf04_completed_file = mf04_outputs / "completed_dates.txt"
                                if mf04_completed_file.exists():
                                    lines = [l.strip() for l in mf04_completed_file.read_text().splitlines() if l.strip()]
                                    if lines:
                                        mf04_latest_date = lines[-1]
                                        mf04_date_folder = find_date_folder_with_variations(mf04_outputs, mf04_latest_date)

                                # If still no folder, create with the marked date
                                if not mf04_date_folder or not mf04_date_folder.exists():
                                    mf04_date_folder = mf04_outputs / formatted_date
                                    mf04_date_folder.mkdir(parents=True, exist_ok=True)

                            if mf04_date_folder and mf04_date_folder.exists():
                                # Determine the target filename based on week selection
                                if la_week == "month_end":
                                    # Month end: "AFL Liquidity - Dec-2025.xlsx"
                                    # First, remove any existing month-end file (they follow pattern: AFL Liquidity - Mon-YYYY.xlsx)
                                    existing_month_end_files = list(mf04_date_folder.glob("AFL Liquidity - *-*.xlsx"))
                                    # Filter to only month-end files (exclude Week_ files)
                                    month_end_files = [f for f in existing_month_end_files if "Week_" not in f.name]

                                    removed_old_month_end = None
                                    for old_file in month_end_files:
                                        removed_old_month_end = old_file.name
                                        old_file.unlink()

                                    # Generate new month-end filename from selected date
                                    dt = datetime.strptime(date_str, "%Y-%m-%d")
                                    target_filename = f"AFL Liquidity - {dt.strftime('%b')}-{dt.year}.xlsx"
                                else:
                                    # Week 1-4: "AFL Liquidity - Week_1.xlsx"
                                    target_filename = f"AFL Liquidity - Week_{la_week}.xlsx"

                                # Copy and rename the AFL file
                                dest = mf04_date_folder / target_filename

                                # If file exists, remove it first (replace) - mainly for weekly files
                                if dest.exists():
                                    dest.unlink()

                                shutil.copy2(afl_file, dest)

                                week_label = "Month End" if la_week == "month_end" else f"Week {la_week}"
                                if la_week == "month_end" and removed_old_month_end:
                                    flash(f"✓ Replaced old month-end file '{removed_old_month_end}' with '{target_filename}' ({week_label}) from {formatted_date}.", "success")
                                else:
                                    flash(f"✓ Copied AFL file from {formatted_date} as '{target_filename}' ({week_label}) to MF-04-LA outputs.", "success")
                            else:
                                flash("Warning: Could not create/find MF-04-LA date folder for AFL copy.", "warning")
                        else:
                            flash(f"Warning: No AFL Liquidity file found in WF-LA outputs for {formatted_date}.", "warning")
                    else:
                        flash(f"Warning: No date folder found in WF-LA outputs for {formatted_date}.", "warning")
                except Exception as e:
                    flash(f"Warning: Failed to copy AFL file to MF-04-LA: {e}", "warning")

        set_name = f"Set {set_num}" + (f" Step {step_num}" if step_num else "")
        flash(f"Marked {set_name} as completed for {formatted_date}.", "success")

    except ValueError as e:
        flash(f"Invalid input: {e}", "error")
    except Exception as e:
        flash(f"Failed to mark as completed: {e}", "error")

    return redirect(url_for("index"))


def start_automation_with_data(report_ids_with_dates: dict, sofp_data: dict, ia_data: dict, uploaded_files: dict = None, months: dict = None):
    """Start automation with master data."""
    # Initialize new run
    _initialize_new_run()
    global _report_checklist, _report_messages, _last_output_dirs, _stop_requested, _download_files, _sofp_download_file, _ia_rate_updates, _ia_ltv_updates, _la_month_end, _la_time_deposits
    _report_checklist = []
    _report_messages = {}
    _last_output_dirs = {}
    _download_files = []
    _sofp_download_file = None
    _stop_requested = False

    # Store LA month-end data globally for background thread access
    from flask import session
    _la_month_end = session.pop('la_month_end', False)
    _la_time_deposits = session.pop('la_time_deposits', '')
    if _la_month_end:
        _emit_log(f"LA month-end mode enabled, time deposits: {_la_time_deposits}", important=True)

    # Store rate_updates globally for background thread access
    _ia_rate_updates = ia_data.get("rate_updates", []) if ia_data else []
    if _ia_rate_updates:
        _emit_log(f"Stored {len(_ia_rate_updates)} rate updates for IA report", important=True)

    # Store ltv_updates globally for background thread access
    _ia_ltv_updates = ia_data.get("ltv_updates", []) if ia_data else []
    if _ia_ltv_updates:
        _emit_log(f"Stored {len(_ia_ltv_updates)} LTV updates for IA report", important=True)
    
    # PROCESS ISOLATION: Wait for master data form to complete
    _emit_log("*** PROCESS ISOLATION: WAITING FOR MASTER DATA TO COMPLETE ***", important=True)
    import time
    time.sleep(3)  # Wait 3 seconds for master data form to complete
    
    # Check for any locked files and wait if necessary
    _emit_log("*** CHECKING FOR FILE LOCKS ***", important=True)
    max_wait_time = 30  # Maximum wait time in seconds
    wait_interval = 1   # Check every 1 second
    waited = 0
    
    while waited < max_wait_time:
        # Check if any output files are locked
        locked_files = []
        for report_id in report_ids_with_dates:
            if report_id in REPORT_STRUCTURE:
                folder_name = REPORT_STRUCTURE[report_id]["folder"]
                if report_id == "NBD-MF-10-GA-11-IS":
                    output_dir = OUTPUTS_DIR / "NBD-MF-10-GA & NBD-MF-11-IS"
                else:
                    output_dir = OUTPUTS_DIR / folder_name
                
                if output_dir.exists():
                    # Find latest date folder
                    last_date, date_folder = find_latest_completed_date(output_dir)
                    if date_folder and date_folder.exists():
                        # Check for locked files
                        for f in date_folder.iterdir():
                            if f.is_file() and f.suffix.lower() in ['.xlsx', '.xlsm']:
                                try:
                                    # Try to open file in exclusive mode to check if it's locked
                                    with open(f, 'r+b') as test_file:
                                        pass
                                except (PermissionError, OSError):
                                    locked_files.append(str(f))
        
        if locked_files:
            _emit_log(f"Found {len(locked_files)} locked files, waiting... ({waited}s/{max_wait_time}s)", important=True)
            time.sleep(wait_interval)
            waited += wait_interval
        else:
            _emit_log("No locked files found, proceeding with automation", important=True)
            break
    
    if waited >= max_wait_time:
        _emit_log("WARNING: Maximum wait time reached, proceeding anyway", important=True)
    
    # Clean up any existing duplicate files before starting automation
    _emit_log("Cleaning up duplicate files before starting automation", important=True)
    cleanup_duplicate_ia_files()
    cleanup_duplicate_c1c2_files()

    # Clean up working directories to prevent "multiple date folders" errors
    _emit_log("Cleaning up working directories before starting automation", important=True)
    cleanup_all_working_directories()

    _set_status(running=True, stage="starting", message=f"Starting {len(report_ids_with_dates)} reports")

    # Run reports in background thread
    def run_in_background():
        _set_status(stage="running", message="Running selected reports")
        success, msg = run_selected_reports(report_ids_with_dates)
        _set_status(running=False, stage="done" if success else "error", message=msg)

    thread = threading.Thread(target=run_in_background, daemon=True)
    thread.start()

    flash(f"Started running {len(report_ids_with_dates)} reports", "success")
    return redirect(url_for("index"))


@app.route("/status", methods=["GET"])
def status():
    with _status_lock:
        return jsonify(_status)


@app.route("/logs", methods=["GET"])
def logs():
    with _log_lock:
        content = "\n".join(_log_buffer)
    return app.response_class(content, mimetype="text/plain")


@app.route("/status-feed", methods=["GET"])
def status_feed():
    content = "\n".join(_status_feed)
    return app.response_class(content, mimetype="text/plain")


@app.route("/report-messages/<report_id>", methods=["GET"])
def report_messages(report_id: str):
    buf = _report_messages.get(report_id)
    if not buf:
        return app.response_class("", mimetype="text/plain")
    return app.response_class("\n".join(buf), mimetype="text/plain")


@app.route("/report-checklist", methods=["GET"])
def report_checklist():
    # Update elapsed times for running reports
    for item in _report_checklist:
        if item["status"] == "running":
            _update_report_elapsed_time(item["name"])
    return jsonify(_report_checklist)


@app.route("/report-set-info", methods=["GET"])
def report_set_info():
    """Return set information for each report to display in accordion"""
    set_info = {}
    for report_id, config in REPORT_STRUCTURE.items():
        set_num = config.get("set")
        step_num = config.get("step")
        if set_num:
            set_info[report_id] = {
                "set": set_num,
                "step": step_num
            }
    return jsonify(set_info)


@app.route("/stop", methods=["POST"])
def stop():
    global _stop_requested
    _stop_requested = True
    _set_status(stage="stopping", message="Stopping requested")

    with _proc_lock:
        for report_id, proc in _running_processes.items():
            try:
                # Kill entire process tree (parent + all child processes)
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    capture_output=True, timeout=10
                )
                _emit_log(f"Stopped {report_id} (PID {proc.pid} and child processes)", important=True)
            except Exception as e:
                # Fallback to proc.kill() if taskkill fails
                try:
                    proc.kill()
                    _emit_log(f"Killed {report_id}", important=True)
                except Exception:
                    pass
        _running_processes.clear()

    # Kill any Excel instances that may have been spawned by report scripts
    try:
        subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True, timeout=5)
        _emit_log("Killed remaining Excel instances", important=True)
    except Exception:
        pass

    # Kill any cscript/wscript instances (VBScript automation)
    try:
        subprocess.run(["taskkill", "/F", "/IM", "cscript.exe"], capture_output=True, timeout=5)
    except Exception:
        pass

    # Clean up working directories to prevent conflicts on next run
    cleanup_all_working_directories()

    # Reset status and initialize new session
    _reset_status()
    _initialize_new_run()

    return jsonify({"ok": True, "refresh": True})


@app.route("/download-latest", methods=["GET"])
def download_latest():
    """Download the prepared zip file with completed report files."""
    global _download_files, _sofp_download_file

    _emit_log(f"Download requested. Current _download_files: {len(_download_files) if _download_files else 0} files", important=True)

    # If no files available (download already completed), reset status and redirect to main page
    if not _download_files:
        _emit_log("No files in _download_files, resetting status and redirecting to main page", important=True)
        _reset_status()  # Reset status to idle to stop frontend polling
        return redirect(url_for("index"))

    if not _download_files[0].exists():
        _emit_log("No valid files in _download_files, attempting to prepare files", important=True)
        # If no prepared zip exists, create one on-demand
        try:
            auto_download_files()
        except Exception as e:
            _emit_log(f"Failed to prepare files for download: {e}", important=True)
            flash("Failed to prepare files for download.", "error")
            return redirect(url_for("index"))

    if not _download_files or not _download_files[0].exists():
        _emit_log("Still no files available after auto_download_files(), redirecting to main page", important=True)
        flash("No files available for download.", "error")
        return redirect(url_for("index"))

    path = _download_files[0]

    # Extract the download filename for display
    if path.suffix.lower() in [".xlsx", ".xlsb", ".xlsm"]:
        download_name = path.name
    else:
        timestamp = path.stem.replace("reports_", "")
        download_name = f"reports_{timestamp}.zip"

    # Clear download files to prevent re-download on page refresh
    _download_files = []
    _sofp_download_file = None
    _emit_log(f"Download initiated: {download_name}. Cleared download session.", important=True)

    # Reset status to idle to stop frontend polling after download
    _reset_status()
    _emit_log("Status reset to idle after download initiated", important=True)

    # Store download completion info in session for display after redirect
    from flask import session
    session['download_completed'] = True
    session['download_filename'] = download_name

    # If the queued item is a single Excel file, send as-is; otherwise send the zip
    try:
        if path.suffix.lower() in [".xlsx", ".xlsb", ".xlsm"]:
            return send_file(str(path), as_attachment=True, download_name=download_name)
        else:
            return send_file(str(path), as_attachment=True, download_name=download_name)
    except Exception as e:
        _emit_log(f"Failed to send file: {e}", important=True)
        flash(f"Failed to send file: {e}", "error")
        _reset_status()  # Reset status on error as well
        return redirect(url_for("index"))


if __name__ == "__main__":
     # Kill any running Excel instances to prevent COM errors
    kill_excel_instances()

    ensure_directories_exist()

    # Production-safe defaults
    # host = os.environ.get("FLASK_HOST", "0.0.0.0")
    # port = int(os.environ.get("PORT", 5000))

    app.run(host="127.0.0.1", port=5000, debug=True)

    # print(f"Starting Waitress server on http://{host}:{port}")
    # serve(app, host=host, port=port)