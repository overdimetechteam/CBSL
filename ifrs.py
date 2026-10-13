import calendar
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import zipfile
from datetime import date, datetime
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from openpyxl import load_workbook
import scheduler as _sched

app = Flask(__name__, template_folder="templates")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-prod")

# Start the background job scheduler (creates log file once if missing)
_sched.start_scheduler()

BASE_DIR = Path(__file__).resolve().parent
OUTPUTS_DIR = BASE_DIR / "outputs"
WORKING_DIR = BASE_DIR / "working"
MASTER_PD_CATEGORY_DIR        = BASE_DIR / "Master Inputs" / "PD Category"
MASTER_PD_ECONOMIC_FACTOR_DIR = BASE_DIR / "Master Inputs" / "Economic Factor"
MASTER_DATA_FILE = BASE_DIR / "Master_Data.xlsx"

# IFRS reports definition
IFRS_REPORTS = [
    {
        "id": "IFRS-01",
        "num": "01",
        "name": "Portfolio 4 ECL / ECL Model / Summary",
        "desc": "Expected Credit Loss - Portfolio 4, ECL model computation, and IFRS 9 ECL summary.",
        "script": "portfolio_automation_xlwings.py",
        "prepare": "summary_working",
    },
    {
        "id": "IFRS-02",
        "num": "02",
        "name": "PD Data Working / Historic PD Calculation / PD Loan / PD Lease",
        "desc": "Consolidated Probability of Default processing for data working, historic PD, PD Loan and PD Lease.",
        "script": "PD.py",
        "prepare": "pd_working",
    },
    {
        "id": "IFRS-03",
        "num": "03",
        "name": "First Default Date / Write Backs / LGD Working",
        "desc": "First default date derivation, write-back adjustments, and Loss Given Default working.",
        "script": "FDDW.py",
        "prepare": "lgd_working",
    },
]

_REPORT_BY_ID = {r["id"]: r for r in IFRS_REPORTS}

# Shared run state
_lock = threading.Lock()
_state = {
    "stage": "idle",  # idle | running | done | error | stopping
    "running": False,
    "current_report": "",
    "message": "",
    "log_lines": [],
    "checklist": [],  # [{"id", "label", "status", "elapsed"}]
    "stop_requested": False,
    "current_proc": None,
}
_pd_master_cache: dict[str, str] = {}
_download_files: list[Path] = []

_PD_MASTER_FIELD_TO_COLUMN = {
    "nsw": "nsw",
    "bcs": "bcs",
    "wcs": "wcs",
}


def _empty_pd_master_values() -> dict[str, str]:
    return {k: "" for k in _PD_MASTER_FIELD_TO_COLUMN}


def _parse_run_date_to_ddmmyyyy(run_date: str) -> str:
    """Convert supported input date formats to DD.MM.YYYY for PD Loan/Lease scripts."""
    value = (run_date or "").strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    raise ValueError(f"Unsupported run date format: {run_date}")


def _resolve_pd_runtime_values() -> dict[str, str]:
    """
    Resolve runtime PD parameters from in-memory values, with fallback to latest Master_Data values.
    Returns all values as strings.
    """
    latest = _read_latest_pd_master_values()
    out = _empty_pd_master_values()

    for key in out:
        mem_val = (_pd_master_cache.get(key) or "").strip()
        latest_val = (latest.get(key) or "").strip()
        out[key] = mem_val if mem_val else latest_val

    return out


def _build_ifrs02_command_plan(run_date: str) -> list[tuple[list[str], str]]:
    """
    Build ordered IFRS-02 subprocess commands:
    1) PD.py -> 2) PD_Loan.py -> 3) PD_Lease.py -> 4) ECL.py
    """
    pd_params = _resolve_pd_runtime_values()
    required_fields = ("nsw", "bcs", "wcs")
    missing = [k for k in required_fields if not pd_params.get(k)]
    if missing:
        raise ValueError(
            "Missing PD master data values: "
            + ", ".join(missing)
            + ". Please fill PD Lease fields in Master Data."
        )

    run_date_ddmmyyyy = _parse_run_date_to_ddmmyyyy(run_date)

    report_dir = BASE_DIR / "report_automations"
    pd_script = report_dir / "PD.py"
    loan_script = report_dir / "PD_Loan.py"
    lease_script = report_dir / "PD_Lease.py"
    ecl_script = report_dir / "ECL.py"

    for p in (pd_script, loan_script, lease_script, ecl_script):
        if not p.exists():
            raise FileNotFoundError(f"Required script not found: {p.name}")

    common_args = [
        "--running_date", run_date_ddmmyyyy,
        "--normal_scenario", pd_params["nsw"],
        "--best_case_scenario", pd_params["bcs"],
        "--worst_case_scenario", pd_params["wcs"],
    ]

    return [
        ([sys.executable, str(pd_script), run_date], "PD.py"),
        ([sys.executable, str(loan_script), *common_args], "PD_Loan.py"),
        ([sys.executable, str(lease_script), *common_args], "PD_Lease.py"),
        ([sys.executable, str(ecl_script)], "ECL.py"),
    ]


def _build_ifrs03_command_plan(lgd_folder: Path, log_dir: Path, run_date: str) -> list[tuple[list[str], str, str | None]]:
    """
    IFRS-03 plan:
      1) bots/Closure_Category_Wise_Report.py  (runs from project root)
      2) report_automations/FDDW.py <YYYY-MM-DD> (runs from lgd_folder)
    Each tuple: (cmd, display_name, cwd_override)  — None means use default run_cwd.
    log_dir is passed to both scripts via IFRS_LGD_LOG_DIR env var (set by caller).
    """
    bot_script  = BASE_DIR / "bots" / "Closure_Category_Wise_Report.py"
    fddw_script = BASE_DIR / "report_automations" / "FDDW.py"
    for p in (bot_script, fddw_script):
        if not p.exists():
            raise FileNotFoundError(f"Required script not found: {p.name}")

    # FDDW.py and write_off.py expect YYYY-MM-DD
    fddw_date = run_date.strip()
    for _fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            fddw_date = datetime.strptime(fddw_date, _fmt).strftime("%Y-%m-%d")
            break
        except ValueError:
            continue

    # refinance_dpd.py and LGD_computation.py expect DD/MM/YYYY
    lgd_date = datetime.strptime(fddw_date, "%Y-%m-%d").strftime("%d/%m/%Y")

    write_off_script       = BASE_DIR / "report_automations" / "write_off.py"
    lgd_computation_script = BASE_DIR / "report_automations" / "LGD_computation.py"
    for p in (write_off_script, lgd_computation_script):
        if not p.exists():
            raise FileNotFoundError(f"Required script not found: {p.name}")

    first_default_folder = str(lgd_folder / "firstDefault")

    # ==================================================================
    # TEMPORARY DEV FIX — REMOVE BEFORE PRODUCTION
    # Set LGD_DEV_ALLOW_RERUN=1 to pass --force to FDDW.py so the same
    # date can be re-run repeatedly during development instead of being
    # rejected by its "To_Date must be after From_Date" guard.
    # Once development is done: delete this if-block and restore the
    # plain fddw_cmd = [sys.executable, str(fddw_script), fddw_date] line.
    # ==================================================================
    fddw_cmd = [sys.executable, str(fddw_script), fddw_date]
    if os.environ.get("LGD_DEV_ALLOW_RERUN") == "1":
        fddw_cmd.append("--force")
    # ==================================================================
    # END TEMPORARY DEV FIX
    # ==================================================================

    return [
        ([sys.executable, str(bot_script)],                           "Closure_Category_Wise_Report.py", str(BASE_DIR)),
        (fddw_cmd,                                                    "FDDW.py",                         first_default_folder),
        ([sys.executable, str(write_off_script), fddw_date],          "write_off.py",                    first_default_folder),
        ([sys.executable, str(lgd_computation_script), lgd_date],     "LGD_computation.py",              str(lgd_folder)),
    ]


def _find_latest_pd_run_log_dir() -> Path | None:
    """Find latest logs/run_* folder containing PD.log."""
    logs_root = BASE_DIR / "logs"
    if not logs_root.exists():
        return None

    candidates: list[tuple[float, Path]] = []
    for run_dir in logs_root.glob("run_*"):
        if not run_dir.is_dir():
            continue
        pd_log = run_dir / "PD.log"
        if pd_log.exists():
            candidates.append((pd_log.stat().st_mtime, run_dir))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def _normalize_header(value: object) -> str:
    return str(value).strip().lower() if value is not None else ""


def _resolve_pd_columns(ws) -> dict[str, int]:
    """Return mapping of required PD headers to worksheet column indexes, creating missing headers."""
    header_to_col: dict[str, int] = {}
    max_col = max(ws.max_column, 1)

    for col_idx in range(1, max_col + 1):
        header = _normalize_header(ws.cell(row=1, column=col_idx).value)
        if header:
            header_to_col[header] = col_idx

    for excel_header in _PD_MASTER_FIELD_TO_COLUMN.values():
        key = excel_header.lower()
        if key not in header_to_col:
            max_col += 1
            ws.cell(row=1, column=max_col, value=excel_header)
            header_to_col[key] = max_col

    return header_to_col


def _read_latest_pd_master_values() -> dict[str, str]:
    """
    Read latest non-empty row values for PD master inputs from Master_Data.xlsx -> PD sheet.
    Returns empty values when file/sheet is unavailable.
    """
    result = _empty_pd_master_values()
    if not MASTER_DATA_FILE.exists():
        return result

    try:
        wb = load_workbook(MASTER_DATA_FILE)
        if "PD" not in wb.sheetnames:
            wb.close()
            return result

        ws = wb["PD"]
        header_to_col = _resolve_pd_columns(ws)
        col_indexes = [header_to_col[h.lower()] for h in _PD_MASTER_FIELD_TO_COLUMN.values()]

        target_row = None
        for row_idx in range(ws.max_row, 1, -1):
            if any(
                ws.cell(row=row_idx, column=col_idx).value not in (None, "")
                for col_idx in col_indexes
            ):
                target_row = row_idx
                break

        if target_row:
            for field, excel_header in _PD_MASTER_FIELD_TO_COLUMN.items():
                value = ws.cell(row=target_row, column=header_to_col[excel_header.lower()]).value
                result[field] = "" if value is None else str(value)

        wb.close()
    except Exception:
        # Keep form available even if master workbook has issues.
        return _empty_pd_master_values()

    return result


def _write_pd_master_values(values: dict[str, str]) -> None:
    """Append PD master values to Master_Data.xlsx -> PD sheet using configured columns."""
    if MASTER_DATA_FILE.exists():
        wb = load_workbook(MASTER_DATA_FILE)
    else:
        from openpyxl import Workbook

        wb = Workbook()
        # Remove default sheet to keep naming clean when creating new workbook.
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    if "PD" not in wb.sheetnames:
        ws = wb.create_sheet("PD")
    else:
        ws = wb["PD"]

    header_to_col = _resolve_pd_columns(ws)
    new_row = ws.max_row + 1 if ws.max_row >= 1 else 2

    for field, excel_header in _PD_MASTER_FIELD_TO_COLUMN.items():
        ws.cell(row=new_row, column=header_to_col[excel_header.lower()], value=values.get(field, ""))

    wb.save(MASTER_DATA_FILE)
    wb.close()


def _reset_state():
    with _lock:
        _state.update(
            {
                "stage": "idle",
                "running": False,
                "current_report": "",
                "message": "",
                "log_lines": [],
                "checklist": [],
                "stop_requested": False,
                "current_proc": None,
            }
        )


def _log(line: str):
    with _lock:
        _state["log_lines"].append(line)


def _set_checklist_status(report_id: str, status: str, elapsed: float | None = None):
    with _lock:
        for item in _state["checklist"]:
            if item["id"] == report_id:
                item["status"] = status
                if elapsed is not None:
                    item["elapsed"] = round(elapsed, 1)
                break


def _find_date_folder_with_variations(parent: Path, date_str: str) -> Path | None:
    date_variants = {date_str.strip()}
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
            match = re.search(r"\((\d+)\)$", child.name)
            candidates.append((int(match.group(1)) if match else 0, child))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def _prepare_ifrs01_summary_working() -> Path:
    outputs_subdir = OUTPUTS_DIR / "Summary"
    working_subdir = WORKING_DIR / "Summary"
    dates_file = outputs_subdir / "completed_dates.txt"

    if not dates_file.exists():
        raise FileNotFoundError(f"Missing file: {dates_file}")

    lines = [line.strip() for line in dates_file.read_text(encoding="utf-8").splitlines() if line.strip()]
    if not lines:
        raise ValueError(f"completed_dates.txt is empty: {dates_file}")

    last_date = lines[-1]
    src_folder = _find_date_folder_with_variations(outputs_subdir, last_date)
    if src_folder is None:
        # Fall back to the most recently modified sub-folder (mark-complete date may
        # differ from the actual folder name created by the post-run copy)
        subdirs = [d for d in outputs_subdir.iterdir() if d.is_dir()]
        if not subdirs:
            raise FileNotFoundError(f"No date folders found in {outputs_subdir}")
        src_folder = max(subdirs, key=lambda d: d.stat().st_mtime)

    working_subdir.mkdir(parents=True, exist_ok=True)
    dest_folder = working_subdir / src_folder.name
    if dest_folder.exists():
        shutil.rmtree(dest_folder)
    shutil.copytree(src_folder, dest_folder)
    return dest_folder


def _prepare_ifrs02_pd_working() -> Path:
    outputs_subdir = OUTPUTS_DIR / "PD"
    working_subdir = WORKING_DIR / "PD"
    dates_file = outputs_subdir / "completed_dates.txt"

    if not dates_file.exists():
        raise FileNotFoundError(f"Missing file: {dates_file}")

    lines = [line.strip() for line in dates_file.read_text(encoding="utf-8").splitlines() if line.strip()]
    if not lines:
        raise ValueError(f"completed_dates.txt is empty: {dates_file}")

    last_date = lines[-1]
    src_folder = _find_date_folder_with_variations(outputs_subdir, last_date)
    if src_folder is None:
        raise FileNotFoundError(f"No folder found for last date '{last_date}' in {outputs_subdir}")

    if working_subdir.exists():
        shutil.rmtree(working_subdir)
    working_subdir.mkdir(parents=True, exist_ok=True)

    dest_folder = working_subdir / src_folder.name
    shutil.copytree(src_folder, dest_folder)
    return dest_folder


def _prepare_lgd_working() -> Path:
    outputs_subdir = OUTPUTS_DIR / "LGD"
    working_subdir = WORKING_DIR / "LGD"
    dates_file = outputs_subdir / "completed_dates.txt"

    if not dates_file.exists():
        raise FileNotFoundError(f"Missing file: {dates_file}")

    lines = [line.strip() for line in dates_file.read_text(encoding="utf-8").splitlines() if line.strip()]
    if not lines:
        raise ValueError(f"completed_dates.txt is empty: {dates_file}")

    last_date = lines[-1]
    src_folder = _find_date_folder_with_variations(outputs_subdir, last_date)
    if src_folder is None:
        raise FileNotFoundError(f"No folder found for last date '{last_date}' in {outputs_subdir}")

    if working_subdir.exists():
        shutil.rmtree(working_subdir)
    working_subdir.mkdir(parents=True, exist_ok=True)

    dest_folder = working_subdir / src_folder.name
    shutil.copytree(src_folder, dest_folder)
    return dest_folder


def _next_versioned_folder(parent: Path, base_name: str) -> Path:
    """Return parent/base_name if it doesn't exist, else parent/base_name(1), (2), …"""
    candidate = parent / base_name
    if not candidate.exists():
        return candidate
    n = 1
    while True:
        candidate = parent / f"{base_name}({n})"
        if not candidate.exists():
            return candidate
        n += 1


def auto_download_files() -> None:
    """Zip multiple files or pass a single file through for browser download."""
    global _download_files
    existing = [f for f in _download_files if f.exists() and f.is_file()]
    if not existing:
        return
    if len(existing) == 1:
        _download_files = existing
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_dir = BASE_DIR / "working" / "_download_tmp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    zip_path = temp_dir / f"ifrs_reports_{ts}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in existing:
            zf.write(f, f.name)
    _download_files = [zip_path]


def _write_download_manifest(dest_folder: Path, queued_files: list[Path], run_root: Path) -> None:
    """
    Record, inside the archived outputs folder, exactly which files were
    queued for download by the auto-download system for this run.

    This lets the standalone "download last run" button (download_report())
    reuse the *same* file-selection the auto-download flow just used,
    instead of re-deriving it independently (which could drift out of sync)
    or falling back to grabbing every file in the folder.
    """
    try:
        rel_paths = []
        for f in queued_files:
            if not f.exists():
                continue
            try:
                rel_paths.append(str(f.relative_to(run_root)))
            except ValueError:
                continue
        if not rel_paths:
            return
        manifest_path = dest_folder / ".download_manifest.json"
        manifest_path.write_text(json.dumps({"files": rel_paths}), encoding="utf-8")
    except Exception:
        pass  # best-effort only — download_report() falls back to rglob("*") if this is missing


# Background worker

def _run_worker(report: dict, run_date: str):
    """Run a single IFRS report item in a background thread."""
    script_names = report.get("scripts") or [report["script"]]

    with _lock:
        _state["stage"] = "running"
        _state["running"] = True
        _state["current_report"] = report["name"]
        _state["message"] = f"Running {report['name']}..."
        _state["checklist"] = [
            {
                "id": report["id"],
                "label": f"{report['num']}. {report['name']}",
                "status": "running",
                "elapsed": None,
            }
        ]

    def emit(line: str):
        _log(line)

    emit(f"{'=' * 60}")
    emit(f"[IFRS] Starting: {report['name']}")
    emit(f"[IFRS] Date: {run_date}")
    emit(f"{'=' * 60}")

    start = time.time()
    success = True
    run_cwd = str(BASE_DIR)
    script_env = os.environ.copy()

    try:
        if report.get("prepare") == "summary_working":
            emit("[IFRS] Preparing working/Summary from outputs/Summary...")
            prepared_folder = _prepare_ifrs01_summary_working()
            emit(f"[IFRS] Prepared folder: {prepared_folder}")
            run_cwd = str(prepared_folder)
            script_env["IFRS_DAILY_FOLDER"] = str(prepared_folder)
            script_env["IFRS_SKIP_ROOT_XLSB_CLEANUP"] = "1"
            script_env["IFRS_SKIP_TRANSFER"] = "1"
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            summary_log_dir = BASE_DIR / "logs" / f"run_{ts}"
            summary_log_dir.mkdir(parents=True, exist_ok=True)
            script_env["IFRS_LOG_DIR"] = str(summary_log_dir)
            emit(f"[IFRS] Log folder: {summary_log_dir}")
        elif report.get("prepare") == "pd_working":
            emit("[IFRS] Preparing working/PD from outputs/PD...")
            prepared_folder = _prepare_ifrs02_pd_working()
            emit(f"[IFRS] Prepared folder: {prepared_folder}")
            run_cwd = str(prepared_folder)
            script_env["PD_BASE_FOLDER"] = str(prepared_folder)
            script_env["PD_RUN_DATE"] = run_date
            # Keep submitted PD master inputs available to the report process.
            script_env["PD_GDP"] = _pd_master_cache.get("gdp", "")
            script_env["PD_CPI"] = _pd_master_cache.get("cpi", "")
            script_env["PD_RF"] = _pd_master_cache.get("rf", "")
            script_env["PD_UR"] = _pd_master_cache.get("ur", "")
            script_env["PD_NSW"] = _pd_master_cache.get("nsw", "")
            script_env["PD_BCS"] = _pd_master_cache.get("bcs", "")
            script_env["PD_WCS"] = _pd_master_cache.get("wcs", "")

            # Copy latest PD Category override file into the prepared working folder
            pd_cat_matches = sorted(MASTER_PD_CATEGORY_DIR.glob("Test_PD_Cate - *.xlsx"))
            if pd_cat_matches:
                src = pd_cat_matches[-1]
                for old in prepared_folder.glob("Test_PD_Cate*.xlsx"):
                    old.unlink()
                shutil.copy2(str(src), str(prepared_folder / src.name))
                emit(f"[IFRS] PD Category file copied: {src.name}")
        elif report.get("prepare") == "lgd_working":
            emit("[IFRS] Preparing working/LGD from outputs/LGD...")
            prepared_folder = _prepare_lgd_working()
            emit(f"[IFRS] Prepared folder: {prepared_folder}")
            run_cwd = str(prepared_folder)
            script_env["LGD_BASE_FOLDER"] = str(prepared_folder)
            script_env["LGD_RUN_DATE"] = run_date
    except Exception as exc:
        _set_checklist_status(report["id"], "failed")
        emit(f"[ERROR] Preparation failed: {exc}")
        success = False

    script_plan: list[tuple[list[str], str, str | None]] = []
    if success:
        try:
            if report.get("id") == "IFRS-02":
                script_plan = [(*t, None) for t in _build_ifrs02_command_plan(run_date)]
            elif report.get("id") == "IFRS-03":
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                lgd_log_dir = BASE_DIR / "logs" / f"run_{ts}"
                lgd_log_dir.mkdir(parents=True, exist_ok=True)
                script_env["IFRS_LGD_LOG_DIR"] = str(lgd_log_dir)
                emit(f"[IFRS] LGD log folder: {lgd_log_dir}")
                script_plan = _build_ifrs03_command_plan(Path(run_cwd), lgd_log_dir, run_date)
            else:
                for script_name in script_names:
                    script_path = BASE_DIR / "report_automations" / script_name
                    if not script_path.exists():
                        raise FileNotFoundError(f"Script not found: {script_path.name}")
                    script_plan.append(([sys.executable, str(script_path), run_date], script_path.name, None))
        except Exception as exc:
            _set_checklist_status(report["id"], "failed")
            emit(f"[ERROR] Script plan creation failed: {exc}")
            success = False

    for cmd, display_name, cwd_override in script_plan if success else []:
        try:
            emit(f"[IFRS] Running script: {display_name}")
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=cwd_override if cwd_override else run_cwd,
                env=script_env,
            )
            with _lock:
                _state["current_proc"] = proc

            _collected_lines: list[str] = []
            for line in proc.stdout:
                stripped = line.rstrip()
                _collected_lines.append(stripped)
                with _lock:
                    if _state["stop_requested"]:
                        proc.terminate()
                        break
                emit(stripped)

            proc.wait()

            with _lock:
                _state["current_proc"] = None
                stopped = _state["stop_requested"]

            elapsed = time.time() - start
            if stopped:
                _set_checklist_status(report["id"], "stopped", elapsed)
                emit(f"[WARN] Run stopped by user after {elapsed:.1f}s.")
                success = False
                break
            if proc.returncode == 2:
                _set_checklist_status(report["id"], "failed", elapsed)
                emit("[ERROR] Invalid date selected for PD report.")

                # Parse sentinel lines emitted by PD.py
                _suggest_date: str | None = None
                _missing_files: list[str] = []
                for _ln in _collected_lines:
                    if _ln.startswith("SUGGEST_DATE:"):
                        _suggest_date = _ln[len("SUGGEST_DATE:"):].strip()
                    elif _ln.startswith("MISSING_FILES:"):
                        raw = _ln[len("MISSING_FILES:"):].strip()
                        _missing_files = [f for f in raw.split("|") if f]

                # Build a rich UI message
                _msg_parts: list[str] = ["Cannot run PD report for the selected date."]
                if _missing_files:
                    _file_list = "".join(f"<li>{f}</li>" for f in _missing_files)
                    _msg_parts.append(
                        f"The following summary files are missing from the working folder:"
                        f"<ul>{_file_list}</ul>"
                    )
                    if _suggest_date:
                        _msg_parts.append(
                            f"Add the missing files and run again with date "
                            f"<strong>{_suggest_date}</strong>."
                        )
                elif _suggest_date:
                    # Fiscal year exceeded — no missing files, just wrong date
                    _msg_parts.append(
                        f"The selected date exceeds the fiscal year boundary. "
                        f"Maximum allowed date for this cycle: <strong>{_suggest_date}</strong>."
                    )
                else:
                    _msg_parts.append(
                        "No summary files (3. Summary_*.xlsb) found in the working folder."
                    )

                with _lock:
                    _state["message"] = "<br>".join(_msg_parts)
                success = False
                break
            if proc.returncode != 0:
                _set_checklist_status(report["id"], "failed", elapsed)
                emit(f"[ERROR] Script exited with code {proc.returncode}.")
                success = False
                break

            emit(f"[DONE] Script finished: {display_name}")
            if report.get("id") == "IFRS-02" and display_name == "PD.py":
                latest_pd_log_dir = _find_latest_pd_run_log_dir()
                if latest_pd_log_dir is not None:
                    script_env["IFRS_PD_LOG_DIR"] = str(latest_pd_log_dir)
                    emit(f"[IFRS] Shared PD log folder set: {latest_pd_log_dir}")
                else:
                    emit("[WARN] Could not resolve PD run log folder; using script default log paths.")

            # Bot moves the file directly to LGD_BASE_FOLDER — no copy needed here

        except Exception as exc:
            _set_checklist_status(report["id"], "failed")
            emit(f"[ERROR] {exc}")
            success = False
            break

    total_elapsed = time.time() - start
    global _download_files
    if success:
        _set_checklist_status(report["id"], "done", total_elapsed)
        emit(f"[DONE] {report['name']} completed in {total_elapsed:.1f}s.")

        # ── IFRS-02 post-completion: download PD outputs + move folder to outputs/PD ──
        if report.get("id") == "IFRS-02":
            _download_files = []

            # Find the single date subfolder inside working/PD
            working_pd = WORKING_DIR / "PD"
            pd_run_folder: Path | None = None
            if working_pd.is_dir():
                subdirs = [d for d in working_pd.iterdir() if d.is_dir()]
                if subdirs:
                    pd_run_folder = max(subdirs, key=lambda d: d.stat().st_mtime)

            if pd_run_folder:
                # Queue the 4 output files for download
                _PD_DOWNLOAD_GLOBS = [
                    "04. PD Lease *.xlsx",
                    "03. PD Loan *.xlsx",
                    "02. Historic PD Calculation_*.xlsb",
                    "01. PD_data_*.xlsb",
                ]
                for _pat in _PD_DOWNLOAD_GLOBS:
                    _matches = sorted(pd_run_folder.glob(_pat))
                    if _matches:
                        _download_files.append(_matches[-1])
                        emit(f"[IFRS] Queued for download: {_matches[-1].name}")
                    else:
                        emit(f"[WARN] No file found matching: {_pat}")

                # Snapshot before auto_download_files() collapses multi-file lists into
                # a single temp zip path — the manifest needs the individual filenames.
                _queued_for_manifest = list(_download_files)
                auto_download_files()

                # Rename folder to selected run date and move to outputs/PD with versioning
                try:
                    _pd_dt: date | None = None
                    for _fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                        try:
                            _pd_dt = datetime.strptime(run_date.strip(), _fmt).date()
                            break
                        except ValueError:
                            continue
                    if _pd_dt is not None:
                        dest_folder_name = _pd_dt.strftime("%d-%m-%Y")
                        outputs_pd = OUTPUTS_DIR / "PD"
                        outputs_pd.mkdir(parents=True, exist_ok=True)
                        dest_folder = _next_versioned_folder(outputs_pd, dest_folder_name)
                        shutil.copytree(str(pd_run_folder), str(dest_folder))
                        _write_download_manifest(dest_folder, _queued_for_manifest, pd_run_folder)
                        emit(f"[IFRS] PD working folder saved to outputs/PD: {dest_folder.name}")
                        # completed_dates.txt is written only when the user marks PD as complete
                        # Clean out the working/PD folder (remove the date subfolder)
                        shutil.rmtree(str(pd_run_folder))
                        emit(f"[IFRS] Removed working/PD/{pd_run_folder.name}")
                    else:
                        emit("[WARN] Could not parse run date — working/PD folder not moved.")
                except Exception as exc:
                    emit(f"[WARN] PD folder move to outputs failed: {exc}")
            else:
                emit("[WARN] Could not locate working/PD run folder for downloads.")

        # ── IFRS-01 post-completion: queue downloads + month-end PD copy ──
        if report.get("id") == "IFRS-01":
            _download_files = []

            # Find the working/Summary/<date>/ root folder
            working_summary = WORKING_DIR / "Summary"
            summary_run_folder: Path | None = None
            if working_summary.is_dir():
                subdirs = [d for d in working_summary.iterdir() if d.is_dir()]
                if subdirs:
                    summary_run_folder = max(subdirs, key=lambda d: d.stat().st_mtime)

            if summary_run_folder:
                # Parse run date for renaming
                _sum_dt: date | None = None
                for _fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                    try:
                        _sum_dt = datetime.strptime(run_date.strip(), _fmt).date()
                        break
                    except ValueError:
                        continue

                # Remove old versioned files and keep only the newest produced file
                # for each of the three output patterns
                for _glob, _label in [
                    ("1. Portfolio_4_ECL*.xlsb", "Portfolio ECL"),
                    ("2. ECL Model*.xlsb",        "ECL Model"),
                    ("3. Summary*.xlsb",           "Summary"),
                ]:
                    _all = sorted(summary_run_folder.glob(_glob), key=lambda f: f.stat().st_mtime)
                    for _old in _all[:-1]:   # delete all but the most recently modified
                        try:
                            _old.unlink()
                            emit(f"[IFRS] Removed old {_label}: {_old.name}")
                        except Exception:
                            pass

                # Rename Portfolio ECL to canonical name 1. Portfolio_4_ECL-DD.MM.YYYY.xlsb
                if _sum_dt:
                    _date_dot = _sum_dt.strftime("%d.%m.%Y")
                    _port_files = sorted(summary_run_folder.glob("1. Portfolio_4_ECL*.xlsb"),
                                         key=lambda f: f.stat().st_mtime)
                    if _port_files:
                        _port_latest = _port_files[-1]
                        _port_canonical = summary_run_folder / f"1. Portfolio_4_ECL-{_date_dot}.xlsb"
                        if _port_latest != _port_canonical:
                            _port_latest.rename(_port_canonical)
                            emit(f"[IFRS] Renamed Portfolio ECL → {_port_canonical.name}")

                # Queue newest of each output file for download
                for _glob, _label in [
                    ("3. Summary*.xlsb",           "Summary"),
                    ("1. Portfolio_4_ECL*.xlsb",   "Portfolio ECL"),
                    ("2. ECL Model*.xlsb",          "ECL Model"),
                ]:
                    _matches = sorted(summary_run_folder.glob(_glob), key=lambda f: f.stat().st_mtime)
                    if _matches:
                        _download_files.append(_matches[-1])
                        emit(f"[IFRS] Queued for download: {_matches[-1].name}")

                input_folder = summary_run_folder / "Input"
                if input_folder.is_dir():
                    overlay_matches = sorted(input_folder.glob("Overlay-*.xlsx"))
                    if overlay_matches:
                        _download_files.append(overlay_matches[-1])
                        emit(f"[IFRS] Queued for download: {overlay_matches[-1].name}")

                # Snapshot before auto_download_files() collapses multi-file lists into
                # a single temp zip path — the manifest needs the individual filenames.
                _queued_for_manifest = list(_download_files)
                auto_download_files()

                # ── Copy working/Summary/<folder> → outputs/Summary/<run-date>[(n)] ──
                try:
                    if _sum_dt is not None:
                        dest_folder_name = _sum_dt.strftime("%d-%m-%Y")
                        outputs_summary = OUTPUTS_DIR / "Summary"
                        dest_folder = _next_versioned_folder(outputs_summary, dest_folder_name)
                        shutil.copytree(str(summary_run_folder), str(dest_folder))
                        _write_download_manifest(dest_folder, _queued_for_manifest, summary_run_folder)
                        emit(f"[IFRS] Working folder saved to outputs/Summary: {dest_folder.name}")
                        # Clean up working/Summary after successful copy
                        for item in (WORKING_DIR / "Summary").iterdir():
                            if item.is_dir():
                                shutil.rmtree(item)
                            else:
                                item.unlink()
                        emit("[IFRS] Cleaned working/Summary folder.")

                        # Copy 1. Portfolio_4_ECL*.xlsb from root → Input/ subfolder
                        try:
                            _port_matches = sorted(dest_folder.glob("1. Portfolio_4_ECL*.xlsb"),
                                                   key=lambda f: f.stat().st_mtime)
                            if _port_matches:
                                _port_src = _port_matches[-1]
                                _input_sub = dest_folder / "Input"
                                _input_sub.mkdir(parents=True, exist_ok=True)
                                for _old in _input_sub.glob("1. Portfolio_4_ECL*.xlsb"):
                                    try:
                                        _old.unlink()
                                    except Exception:
                                        pass
                                shutil.copy2(str(_port_src), str(_input_sub / _port_src.name))
                                emit(f"[IFRS] Portfolio ECL copied to outputs/Summary Input: {_port_src.name}")
                        except Exception as _exc:
                            emit(f"[WARN] Portfolio ECL → Summary Input copy failed: {_exc}")
                except Exception as exc:
                    emit(f"[WARN] Copy working folder to outputs/Summary failed: {exc}")
            else:
                emit("[WARN] Could not locate working/Summary run folder for downloads.")

        # ── IFRS-03 post-completion: download LGD outputs + move folder to outputs/LGD ──
        if report.get("id") == "IFRS-03":
            _download_files = []

            # Find the single date subfolder inside working/LGD
            working_lgd = WORKING_DIR / "LGD"
            lgd_run_folder: Path | None = None
            if working_lgd.is_dir():
                subdirs = [d for d in working_lgd.iterdir() if d.is_dir()]
                if subdirs:
                    lgd_run_folder = max(subdirs, key=lambda d: d.stat().st_mtime)

            _queued_for_manifest: list[Path] = []
            if lgd_run_folder:
                lgd_sub = lgd_run_folder / "LGD"
                if lgd_sub.is_dir():
                    # Month-named inputs (Receipts Discount, DC List) must match the
                    # run month, otherwise an older month's file with a newer mtime
                    # (e.g. "Receipts Discount March 2025.xlsb") would be picked.
                    _run_month = ""
                    for _fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                        try:
                            _run_month = datetime.strptime(run_date.strip(), _fmt).strftime("%B")
                            break
                        except ValueError:
                            continue
                    _LGD_DOWNLOAD_GLOBS = [
                        "LGD*.xlsb",
                        "receipt details.xlsx",
                        "writebacks * to *.xlsx",
                        "writeoff *.xlsx",
                        "Write Off Working - *.xlsx",
                        f"Receipts Discount {_run_month}*.xlsb" if _run_month else "Receipts Discount *.xlsb",
                        f"DC List {_run_month}*.xlsx" if _run_month else "DC List *.xlsx",
                        "DCF_Receipts_Pivot.xlsx",
                    ]
                    for _pat in _LGD_DOWNLOAD_GLOBS:
                        _matches = sorted(lgd_sub.glob(_pat), key=lambda f: f.stat().st_mtime)
                        if _matches:
                            _download_files.append(_matches[-1])
                            emit(f"[IFRS] Queued for download: {_matches[-1].name}")
                        else:
                            emit(f"[WARN] No file found matching: {_pat} in LGD subfolder")

                    # First Default Date Working (FDDW's output) lives in the
                    # firstDefault subfolder, not the LGD subfolder.
                    firstdefault_sub = lgd_run_folder / "firstDefault"
                    if firstdefault_sub.is_dir():
                        _fdd_matches = sorted(
                            firstdefault_sub.glob("First Default Date Working DPD*.xlsx"),
                            key=lambda f: f.stat().st_mtime,
                        )
                        if _fdd_matches:
                            _download_files.append(_fdd_matches[-1])
                            emit(f"[IFRS] Queued for download: {_fdd_matches[-1].name}")
                        else:
                            emit("[WARN] No file found matching: First Default Date Working DPD*.xlsx in firstDefault subfolder")

                    # Snapshot before auto_download_files() collapses multi-file lists
                    # into a single temp zip path — the manifest needs the individual
                    # filenames, not the zip.
                    _queued_for_manifest = list(_download_files)
                    auto_download_files()
                else:
                    emit(f"[WARN] LGD subfolder not found in {lgd_run_folder.name}")

                # Rename folder to selected run date and move to outputs/LGD with versioning
                try:
                    _lgd_dt: date | None = None
                    for _fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                        try:
                            _lgd_dt = datetime.strptime(run_date.strip(), _fmt).date()
                            break
                        except ValueError:
                            continue
                    if _lgd_dt is not None:
                        dest_folder_name = _lgd_dt.strftime("%d-%m-%Y")
                        outputs_lgd = OUTPUTS_DIR / "LGD"
                        outputs_lgd.mkdir(parents=True, exist_ok=True)
                        dest_folder = _next_versioned_folder(outputs_lgd, dest_folder_name)
                        shutil.copytree(str(lgd_run_folder), str(dest_folder))
                        _write_download_manifest(dest_folder, _queued_for_manifest, lgd_run_folder)
                        emit(f"[IFRS] LGD working folder saved to outputs/LGD: {dest_folder.name}")
                        # Update completed_dates.txt
                        _lgd_dates_file = outputs_lgd / "completed_dates.txt"
                        _lgd_existing = _lgd_dates_file.read_text(encoding="utf-8") if _lgd_dates_file.exists() else ""
                        _lgd_prefix = "\n" if _lgd_existing and not _lgd_existing.endswith("\n") else ""
                        with _lgd_dates_file.open("a", encoding="utf-8") as _f:
                            _f.write(_lgd_prefix + dest_folder.name + "\n")
                        emit(f"[IFRS] Updated outputs/LGD/completed_dates.txt: {dest_folder.name}")
                        # Clean out the working/LGD folder
                        shutil.rmtree(str(lgd_run_folder))
                        emit(f"[IFRS] Removed working/LGD/{lgd_run_folder.name}")
                    else:
                        emit("[WARN] Could not parse run date — working/LGD folder not moved.")
                except Exception as exc:
                    emit(f"[WARN] LGD folder move to outputs failed: {exc}")
            else:
                emit("[WARN] Could not locate working/LGD run folder for downloads.")

    emit(f"{'=' * 60}")

    with _lock:
        _state["running"] = False
        _state["current_report"] = ""
        _state["stage"] = "done" if success else "error"
        _state["message"] = "Completed successfully." if success else "Finished with errors."


# Routes


@app.route("/")
def ifrs_dashboard():
    today_iso = date.today().isoformat()
    last_runs: dict = {}  # TODO: persist last-run dates per report if needed
    return render_template(
        "ifrs_dashboard.html", reports=IFRS_REPORTS, today_iso=today_iso, last_runs=last_runs
    )


@app.route("/mark-complete", methods=["POST"])
def mark_complete():
    """
    Mark Summary report as complete for the selected date.
    Appends date to outputs/Summary/completed_dates.txt, then transfers files
    from the max version folder of that date in outputs/Summary/:
      - 2. ECL Model_*.xlsb      → outputs/PD/<latest-folder>/
      - 3. Summary_*.xlsb        → outputs/PD/<latest-folder>/  (month-ends only)
      - 1. Portfolio_4_ECL*.xlsb → outputs/LGD/<latest-folder>/ (month-ends only)
    """
    import re as _re

    date_str = request.form.get("date", "").strip()
    if not date_str:
        flash("Please select a date.", "warning")
        return redirect(url_for("ifrs_dashboard"))

    try:
        mark_dt = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        flash("Invalid date format.", "error")
        return redirect(url_for("ifrs_dashboard"))

    date_ddmmyyyy = mark_dt.strftime("%d-%m-%Y")
    date_iso      = mark_dt.strftime("%Y-%m-%d")
    m_val, d_val, y_val = mark_dt.month, mark_dt.day, mark_dt.year
    is_month_end  = calendar.monthrange(y_val, m_val)[1] == d_val

    # ── 1. Find the max version folder for the marked date ───────────────
    summary_outputs = OUTPUTS_DIR / "Summary"
    sum_folder = _find_date_folder_with_variations(summary_outputs, date_ddmmyyyy)
    if not sum_folder:
        flash(f"No matching folder found in outputs/Summary for {date_ddmmyyyy}. Date not marked.", "error")
        return redirect(url_for("ifrs_dashboard"))

    # ── 2. Append to completed_dates.txt ──────────────────────────────────
    summary_dates_file = summary_outputs / "completed_dates.txt"
    try:
        existing = summary_dates_file.read_text(encoding="utf-8") if summary_dates_file.exists() else ""
        prefix = "\n" if existing and not existing.endswith("\n") else ""
        with summary_dates_file.open("a", encoding="utf-8") as f:
            f.write(prefix + date_ddmmyyyy + "\n")
    except Exception as exc:
        flash(f"Failed to update completed_dates.txt: {exc}", "error")
        return redirect(url_for("ifrs_dashboard"))

    messages: list[str] = [f"Marked complete for {date_ddmmyyyy}."]

    # ── 3. ECL Model → outputs/PD/<latest-folder>/ and outputs/LGD/<latest-folder>/LGD/ ──
    try:
        ecl_matches = list(sum_folder.glob("2. ECL Model_*.xlsb"))
        if ecl_matches:
            ecl_src = max(ecl_matches, key=lambda f: f.stat().st_mtime)

            # → outputs/PD/<latest-folder>/
            pd_outputs    = OUTPUTS_DIR / "PD"
            pd_dates_file = pd_outputs / "completed_dates.txt"
            if pd_dates_file.exists():
                pd_lines = [l.strip() for l in pd_dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
                if pd_lines:
                    pd_folder = _find_date_folder_with_variations(pd_outputs, pd_lines[-1])
                    if pd_folder:
                        ecl_dest_name = f"ECL Model_{date_iso}.xlsb"
                        for old in pd_folder.glob("ECL Model_*.xlsb"):
                            try:
                                old.unlink()
                            except Exception:
                                pass
                        shutil.copy2(str(ecl_src), str(pd_folder / ecl_dest_name))
                        messages.append(f"ECL Model copied to outputs/PD as {ecl_dest_name}.")

            # → outputs/LGD/<latest-folder>/LGD/
            lgd_outputs    = OUTPUTS_DIR / "LGD"
            lgd_dates_file = lgd_outputs / "completed_dates.txt"
            if lgd_dates_file.exists():
                lgd_lines = [l.strip() for l in lgd_dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
                if lgd_lines:
                    lgd_folder = _find_date_folder_with_variations(lgd_outputs, lgd_lines[-1])
                    if lgd_folder:
                        lgd_sub = lgd_folder / "LGD"
                        lgd_sub.mkdir(parents=True, exist_ok=True)
                        for old in lgd_sub.glob("2. ECL Model_*.xlsb"):
                            try:
                                old.unlink()
                            except Exception:
                                pass
                        shutil.copy2(str(ecl_src), str(lgd_sub / ecl_src.name))
                        messages.append(f"ECL Model copied to outputs/LGD/LGD as {ecl_src.name}.")
    except Exception as exc:
        messages.append(f"ECL Model copy failed: {exc}")

    # ── 4. Summary file → outputs/PD/<latest-folder>/ (month-ends only) ──
    if is_month_end:
        try:
            sum_files = sorted(sum_folder.glob("3. Summary_*.xlsb"))
            if sum_files:
                new_sum   = max(sum_files, key=lambda f: f.stat().st_mtime)
                pd_outputs    = OUTPUTS_DIR / "PD"
                pd_dates_file = pd_outputs / "completed_dates.txt"
                if pd_dates_file.exists():
                    pd_lines = [l.strip() for l in pd_dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
                    if pd_lines:
                        pd_folder = _find_date_folder_with_variations(pd_outputs, pd_lines[-1])
                        if pd_folder:
                            dest_name = f"3. Summary_{date_iso}.xlsb"
                            if m_val == 4:
                                # April: keep March 31 file, remove all other Summary files
                                keep_file: Path | None = None
                                march31_files = [
                                    f for f in pd_folder.glob("3. Summary_????-??-??*.xlsb")
                                    if _re.search(r"\d{4}-03-31", f.stem)
                                ]
                                if march31_files:
                                    keep_file = max(march31_files, key=lambda f: f.stat().st_mtime)
                                for old in pd_folder.glob("3. Summary_*.xlsb"):
                                    if keep_file and old == keep_file:
                                        continue
                                    try:
                                        old.unlink()
                                    except Exception:
                                        pass
                            shutil.copy2(str(new_sum), str(pd_folder / dest_name))
                            messages.append(f"Summary copied to outputs/PD as {dest_name}.")
        except Exception as exc:
            messages.append(f"Summary→PD copy failed: {exc}")

        # ── 5. Portfolio ECL → outputs/LGD/<latest-folder>/firstDefault/ and /LGD/ ──
        try:
            ecl_port = sorted(sum_folder.glob("1. Portfolio_4_ECL*.xlsb"))
            if ecl_port:
                new_ecl        = max(ecl_port, key=lambda f: f.stat().st_mtime)
                lgd_outputs    = OUTPUTS_DIR / "LGD"
                lgd_dates_file = lgd_outputs / "completed_dates.txt"
                if lgd_dates_file.exists():
                    lgd_lines = [l.strip() for l in lgd_dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
                    if lgd_lines:
                        lgd_folder = _find_date_folder_with_variations(lgd_outputs, lgd_lines[-1])
                        if lgd_folder:
                            # firstDefault: April clears all existing files first
                            fd_sub = lgd_folder / "firstDefault"
                            fd_sub.mkdir(parents=True, exist_ok=True)
                            if m_val == 4:
                                for old in fd_sub.glob("1. Portfolio_4_ECL*.xlsb"):
                                    try:
                                        old.unlink()
                                    except Exception:
                                        pass
                            shutil.copy2(str(new_ecl), str(fd_sub / new_ecl.name))

                            # LGD: always remove existing and paste new
                            lgd_sub = lgd_folder / "LGD"
                            lgd_sub.mkdir(parents=True, exist_ok=True)
                            for old in lgd_sub.glob("1. Portfolio_4_ECL*.xlsb"):
                                try:
                                    old.unlink()
                                except Exception:
                                    pass
                            shutil.copy2(str(new_ecl), str(lgd_sub / new_ecl.name))

                            messages.append(f"Portfolio ECL copied to outputs/LGD firstDefault and LGD as {new_ecl.name}.")
        except Exception as exc:
            messages.append(f"Portfolio ECL→LGD copy failed: {exc}")

    flash(" ".join(messages), "success")
    return redirect(url_for("ifrs_dashboard"))


@app.route("/mark-complete-pd", methods=["POST"])
def mark_complete_pd():
    """
    Mark PD report as complete for the selected date.
    Checks outputs/PD/<date-folder> exists (and is the only match), appends date to
    outputs/PD/completed_dates.txt, then copies ECL Model*.xlsb from that PD folder
    to outputs/Summary/<latest-folder>/ (last line of Summary/completed_dates.txt, max version).
    """
    date_str = request.form.get("date", "").strip()
    if not date_str:
        flash("Please select a date.", "warning")
        return redirect(url_for("ifrs_dashboard"))

    try:
        mark_dt = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        flash("Invalid date format.", "error")
        return redirect(url_for("ifrs_dashboard"))

    date_ddmmyyyy = mark_dt.strftime("%d-%m-%Y")

    pd_outputs = OUTPUTS_DIR / "PD"
    pd_folder  = _find_date_folder_with_variations(pd_outputs, date_ddmmyyyy)
    if not pd_folder:
        flash(f"No matching folder found in outputs/PD for {date_ddmmyyyy}. Date not marked.", "error")
        return redirect(url_for("ifrs_dashboard"))

    # Append to completed_dates.txt
    pd_dates_file = pd_outputs / "completed_dates.txt"
    try:
        existing = pd_dates_file.read_text(encoding="utf-8") if pd_dates_file.exists() else ""
        prefix = "\n" if existing and not existing.endswith("\n") else ""
        with pd_dates_file.open("a", encoding="utf-8") as f:
            f.write(prefix + date_ddmmyyyy + "\n")
    except Exception as exc:
        flash(f"Failed to update completed_dates.txt: {exc}", "error")
        return redirect(url_for("ifrs_dashboard"))

    messages: list[str] = [f"PD marked complete for {date_ddmmyyyy}."]

    # Copy ECL Model*.xlsb → outputs/Summary/<latest-folder>/
    try:
        ecl_matches = list(pd_folder.glob("ECL Model*.xlsb"))
        if ecl_matches:
            ecl_src = max(ecl_matches, key=lambda f: f.stat().st_mtime)
            summary_outputs    = OUTPUTS_DIR / "Summary"
            summary_dates_file = summary_outputs / "completed_dates.txt"
            if summary_dates_file.exists():
                sum_lines = [l.strip() for l in summary_dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
                if sum_lines:
                    sum_folder = _find_date_folder_with_variations(summary_outputs, sum_lines[-1])
                    if sum_folder:
                        import re as _re
                        _dm = _re.search(r'(\d{4})-(\d{2})-(\d{2})|(\d{2})[.\-](\d{2})[.\-](\d{4})', ecl_src.stem)
                        if _dm and _dm.group(1):
                            _date_part = f"{_dm.group(3)}.{_dm.group(2)}.{_dm.group(1)}"
                        elif _dm:
                            _date_part = f"{_dm.group(4)}.{_dm.group(5)}.{_dm.group(6)}"
                        else:
                            _date_part = mark_dt.strftime("%d.%m.%Y")
                        dest_name = f"2. ECL Model_{_date_part}.xlsb"
                        for old in sum_folder.glob("2. ECL Model*.xlsb"):
                            try:
                                old.unlink()
                            except Exception:
                                pass
                        shutil.copy2(str(ecl_src), str(sum_folder / dest_name))
                        messages.append(f"ECL Model copied to outputs/Summary as {dest_name}.")
                    else:
                        messages.append("Warning: Could not find outputs/Summary date folder.")
                else:
                    messages.append("Warning: outputs/Summary/completed_dates.txt is empty.")
            else:
                messages.append("Warning: outputs/Summary/completed_dates.txt not found.")
        else:
            messages.append(f"Warning: No ECL Model*.xlsb found in outputs/PD/{pd_folder.name}.")
    except Exception as exc:
        messages.append(f"ECL Model copy failed: {exc}")

    # Copy ECL Model*.xlsb → outputs/LGD/<latest-folder>/LGD/
    try:
        ecl_matches = list(pd_folder.glob("ECL Model*.xlsb"))
        messages.append(f"[LGD] PD folder: {pd_folder.name} | ECL files found: {[f.name for f in ecl_matches]}")
        if ecl_matches:
            ecl_src = max(ecl_matches, key=lambda f: f.stat().st_mtime)
            lgd_outputs    = OUTPUTS_DIR / "LGD"
            lgd_dates_file = lgd_outputs / "completed_dates.txt"
            if lgd_dates_file.exists():
                lgd_lines = [l.strip() for l in lgd_dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
                messages.append(f"[LGD] Last completed date: {lgd_lines[-1] if lgd_lines else 'none'}")
                if lgd_lines:
                    lgd_folder = _find_date_folder_with_variations(lgd_outputs, lgd_lines[-1])
                    messages.append(f"[LGD] Resolved folder: {lgd_folder}")
                    if lgd_folder:
                        lgd_sub = lgd_folder / "LGD"
                        lgd_sub.mkdir(parents=True, exist_ok=True)
                        removed = []
                        for old in lgd_sub.glob("2. ECL Model*.xlsb"):
                            try:
                                old.unlink()
                                removed.append(old.name)
                            except Exception:
                                pass
                        if removed:
                            messages.append(f"[LGD] Removed: {removed}")
                        shutil.copy2(str(ecl_src), str(lgd_sub / ecl_src.name))
                        messages.append(f"ECL Model copied to outputs/LGD/LGD as {ecl_src.name}.")
                    else:
                        messages.append(f"Warning: Could not find outputs/LGD date folder for '{lgd_lines[-1]}'.")
                else:
                    messages.append("Warning: outputs/LGD/completed_dates.txt is empty.")
            else:
                messages.append(f"Warning: outputs/LGD/completed_dates.txt not found at {lgd_dates_file}.")
        else:
            messages.append(f"Warning: No ECL Model*.xlsb found in outputs/PD/{pd_folder.name}.")
    except Exception as exc:
        messages.append(f"ECL Model LGD copy failed: {exc}")

    flash(" ".join(messages), "success")
    return redirect(url_for("ifrs_dashboard"))


@app.route("/mark-complete-lgd", methods=["POST"])
def mark_complete_lgd():
    """
    Mark LGD report as complete for the selected date.
    Checks outputs/LGD/<date-folder> exists, appends date to
    outputs/LGD/completed_dates.txt, then copies ECL Model*.xlsb from
    outputs/LGD/<date-folder>/LGD/ to outputs/Summary/<latest-folder>/
    (last line of Summary/completed_dates.txt, max version).
    """
    date_str = request.form.get("date", "").strip()
    if not date_str:
        flash("Please select a date.", "warning")
        return redirect(url_for("ifrs_dashboard"))

    try:
        mark_dt = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        flash("Invalid date format.", "error")
        return redirect(url_for("ifrs_dashboard"))

    date_ddmmyyyy = mark_dt.strftime("%d-%m-%Y")

    lgd_outputs = OUTPUTS_DIR / "LGD"
    lgd_folder  = _find_date_folder_with_variations(lgd_outputs, date_ddmmyyyy)
    if not lgd_folder:
        flash(f"No matching folder found in outputs/LGD for {date_ddmmyyyy}. Date not marked.", "error")
        return redirect(url_for("ifrs_dashboard"))

    # Append to completed_dates.txt
    lgd_dates_file = lgd_outputs / "completed_dates.txt"
    try:
        existing = lgd_dates_file.read_text(encoding="utf-8") if lgd_dates_file.exists() else ""
        prefix = "\n" if existing and not existing.endswith("\n") else ""
        with lgd_dates_file.open("a", encoding="utf-8") as f:
            f.write(prefix + date_ddmmyyyy + "\n")
    except Exception as exc:
        flash(f"Failed to update completed_dates.txt: {exc}", "error")
        return redirect(url_for("ifrs_dashboard"))

    messages: list[str] = [f"LGD marked complete for {date_ddmmyyyy}."]

    # Copy ECL Model*.xlsb from outputs/LGD/<date>/LGD/ → outputs/Summary/<latest-folder>/
    try:
        lgd_sub     = lgd_folder / "LGD"
        ecl_matches = list(lgd_sub.glob("2. ECL Model*.xlsb")) if lgd_sub.exists() else []
        if ecl_matches:
            ecl_src = max(ecl_matches, key=lambda f: f.stat().st_mtime)
            summary_outputs    = OUTPUTS_DIR / "Summary"
            summary_dates_file = summary_outputs / "completed_dates.txt"
            if summary_dates_file.exists():
                sum_lines = [l.strip() for l in summary_dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
                if sum_lines:
                    sum_folder = _find_date_folder_with_variations(summary_outputs, sum_lines[-1])
                    if sum_folder:
                        import re as _re
                        _dm = _re.search(r'(\d{4})-(\d{2})-(\d{2})|(\d{2})[.\-](\d{2})[.\-](\d{4})', ecl_src.stem)
                        if _dm and _dm.group(1):
                            _date_part = f"{_dm.group(3)}.{_dm.group(2)}.{_dm.group(1)}"
                        elif _dm:
                            _date_part = f"{_dm.group(4)}.{_dm.group(5)}.{_dm.group(6)}"
                        else:
                            _date_part = mark_dt.strftime("%d.%m.%Y")
                        dest_name = f"2. ECL Model_{_date_part}.xlsb"
                        for old in sum_folder.glob("2. ECL Model*.xlsb"):
                            try:
                                old.unlink()
                            except Exception:
                                pass
                        shutil.copy2(str(ecl_src), str(sum_folder / dest_name))
                        messages.append(f"ECL Model copied to outputs/Summary as {dest_name}.")
                    else:
                        messages.append("Warning: Could not find outputs/Summary date folder.")
                else:
                    messages.append("Warning: outputs/Summary/completed_dates.txt is empty.")
            else:
                messages.append("Warning: outputs/Summary/completed_dates.txt not found.")
        else:
            messages.append(f"Warning: No ECL Model*.xlsb found in outputs/LGD/{lgd_folder.name}/LGD/.")
    except Exception as exc:
        messages.append(f"ECL Model copy failed: {exc}")

    flash(" ".join(messages), "success")
    return redirect(url_for("ifrs_dashboard"))


@app.route("/run-report", methods=["POST"])
def run_report():
    report_id = request.form.get("report_id", "").strip()
    run_date = request.form.get("date", "").strip()

    if not report_id or report_id not in _REPORT_BY_ID:
        flash("Invalid report selected.", "error")
        return redirect(url_for("ifrs_dashboard"))

    if not run_date:
        flash("Please select a date.", "warning")
        return redirect(url_for("ifrs_dashboard"))

    # Convert yyyy-mm-dd -> M/D/YYYY for scripts
    try:
        y, m, d = run_date.split("-")
        script_date = f"{int(m)}/{int(d)}/{y}"
    except ValueError:
        flash("Invalid date format.", "error")
        return redirect(url_for("ifrs_dashboard"))

    with _lock:
        if _state["running"]:
            flash("A report is already running. Please wait.", "warning")
            return redirect(url_for("ifrs_dashboard"))

    _reset_state()
    report = _REPORT_BY_ID[report_id]
    t = threading.Thread(target=_run_worker, args=(report, script_date), daemon=True)
    t.start()

    return redirect(url_for("ifrs_dashboard"))


@app.route("/stop", methods=["POST"])
def stop():
    with _lock:
        _state["stop_requested"] = True
        _state["stage"] = "stopping"
        proc = _state.get("current_proc")
    if proc:
        try:
            proc.terminate()
        except Exception:
            pass
    return jsonify({"ok": True})


@app.route("/reset", methods=["POST"])
def reset():
    _reset_state()
    return jsonify({"ok": True})


@app.route("/status")
def status():
    with _lock:
        return jsonify(
            {
                "stage": _state["stage"],
                "running": _state["running"],
                "current_report": _state["current_report"],
                "message": _state["message"],
            }
        )


@app.route("/logs")
def logs():
    with _lock:
        lines = list(_state["log_lines"])
    return "\n".join(lines), 200, {"Content-Type": "text/plain; charset=utf-8"}


@app.route("/checklist")
def checklist():
    with _lock:
        items = list(_state["checklist"])
    return jsonify(items)


def _get_pd_category_info() -> dict:
    """Return filename and modified date of the latest PD Category file."""
    MASTER_PD_CATEGORY_DIR.mkdir(parents=True, exist_ok=True)
    matches = sorted(MASTER_PD_CATEGORY_DIR.glob("Test_PD_Cate - *.xlsx"))
    if not matches:
        return {"filename": None, "modified": None}
    f = matches[-1]
    modified = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d %b %Y, %H:%M")
    return {"filename": f.name, "modified": modified}


MASTER_INPUTS_DIR = BASE_DIR / "Master Inputs"

_SUMMARY_MASTER_DIRS = [
    "1611", "2128", "Assetline Related Party", "CadreSUM",
    "Insurance Premium", "Litigation Report", "Mediation Detail Report",
    "Moratorium Contracts Details", "Yard Final",
]

MASTER_OVERLAY_DIR     = MASTER_INPUTS_DIR / "Overlay"
MASTER_LGD_YARD_DIR    = MASTER_INPUTS_DIR / "Yard FInal LGD"
MASTER_LGD_RECEIPT_DISCOUNT_DIR = MASTER_INPUTS_DIR / "Receipt Discount LGD"
MASTER_LGD_DC_LIST_DIR          = MASTER_INPUTS_DIR / "DC List LGD"


def _get_overlay_master_file() -> str | None:
    """Return the filename of the latest file in Master Inputs/Overlay, or None."""
    MASTER_OVERLAY_DIR.mkdir(parents=True, exist_ok=True)
    files = [f for f in MASTER_OVERLAY_DIR.iterdir() if f.is_file() and not f.name.startswith("~$")]
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime).name


def _get_lgd_master_files() -> dict:
    """Return the latest filename in each LGD master-input folder."""
    result = {}
    for label, d in (
        ("Yard FInal LGD", MASTER_LGD_YARD_DIR),
        ("Receipt Discount LGD", MASTER_LGD_RECEIPT_DISCOUNT_DIR),
        ("DC List LGD", MASTER_LGD_DC_LIST_DIR),
    ):
        d.mkdir(parents=True, exist_ok=True)
        files = [f for f in d.iterdir() if f.is_file() and not f.name.startswith("~$")]
        result[label] = max(files, key=lambda f: f.stat().st_mtime).name if files else None
    return result


def _get_pd_economic_factor_file() -> str | None:
    """Return the filename of the latest file in Master Inputs/Economic Factor, or None."""
    MASTER_PD_ECONOMIC_FACTOR_DIR.mkdir(parents=True, exist_ok=True)
    files = [f for f in MASTER_PD_ECONOMIC_FACTOR_DIR.iterdir() if f.is_file() and not f.name.startswith("~$")]
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime).name


def _find_pd_outputs_subfolder() -> Path | None:
    """Return outputs/PD/<latest-date-folder> using completed_dates.txt."""
    outputs_pd = OUTPUTS_DIR / "PD"
    dates_file = outputs_pd / "completed_dates.txt"
    if not dates_file.exists():
        return None
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    if not lines:
        return None
    date_folder = _find_date_folder_with_variations(outputs_pd, lines[-1])
    return date_folder


def _find_lgd_outputs_lgd_subfolder() -> Path | None:
    """Return outputs/LGD/<latest-date-folder>/LGD using completed_dates.txt."""
    outputs_lgd = OUTPUTS_DIR / "LGD"
    dates_file = outputs_lgd / "completed_dates.txt"
    if not dates_file.exists():
        return None
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    if not lines:
        return None
    date_folder = _find_date_folder_with_variations(outputs_lgd, lines[-1])
    if not date_folder:
        return None
    return date_folder / "LGD"


def _get_summary_master_files() -> dict:
    """Return the latest filename in each Summary master-input folder."""
    result = {}
    for name in _SUMMARY_MASTER_DIRS:
        d = MASTER_INPUTS_DIR / name
        if d.is_dir():
            files = [f for f in d.iterdir() if f.is_file() and not f.name.startswith("~$")]
            if files:
                latest = max(files, key=lambda f: f.stat().st_mtime)
                result[name] = latest.name
            else:
                result[name] = None
        else:
            result[name] = None
    return result


def _find_summary_outputs_input() -> Path | None:
    """Return outputs/Summary/<latest-date-folder>/Input using completed_dates.txt."""
    outputs_subdir = OUTPUTS_DIR / "Summary"
    dates_file = outputs_subdir / "completed_dates.txt"
    if not dates_file.exists():
        return None
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    if not lines:
        return None
    last_date = lines[-1]
    date_folder = _find_date_folder_with_variations(outputs_subdir, last_date)
    if not date_folder:
        return None
    return date_folder / "Input"


def _copy_summary_files_to_working(summary_uploads: dict) -> None:
    """Copy uploaded Summary master files into outputs/Summary/<latest-date>/Input with cleanup."""
    input_dir = _find_summary_outputs_input()
    if not input_dir:
        return
    input_dir.mkdir(parents=True, exist_ok=True)

    cleanup_patterns = {
        "1611":                         "1611-*.xlsx",
        "2128":                         "2128-*.xlsx",
        "CadreSUM":                     "Cadre as at *.xlsx",
        "Insurance Premium":            "Insurance Premium *.xlsx",
        "Litigation Report":            "Litigation Report - *.xlsx",
        "Mediation Detail Report":      "Mediation Detail Report - *.xlsx",
        "Moratorium Contracts Details": "Moratorium Contracts details-*.xlsx",
        "Yard Final":                   "Yard Final - *.xlsx",
    }
    fixed_names = {
        "Assetline Related Party": "Assetline Finance Limited - Related Party Register_All.xlsb",
    }

    for category, file_path_str in summary_uploads.items():
        file_path = Path(file_path_str)
        if not file_path.exists():
            continue
        new_name = file_path.name

        if category in fixed_names:
            old = input_dir / fixed_names[category]
            if old.exists():
                try:
                    old.unlink()
                except Exception:
                    pass
        elif category in cleanup_patterns:
            for old in input_dir.glob(cleanup_patterns[category]):
                if old.name != new_name:
                    try:
                        old.unlink()
                    except Exception:
                        pass

        shutil.copy2(file_path, input_dir / new_name)


@app.route("/master-data", methods=["GET"])
def master_data():
    pd_category = _get_pd_category_info()
    pd_master_data = _read_latest_pd_master_values()
    report_id = request.args.get("report_id", "")
    run_date  = request.args.get("date", "")
    summary_files = _get_summary_master_files()
    overlay_file  = _get_overlay_master_file()
    lgd_files     = _get_lgd_master_files()
    pd_economic_factor_file = _get_pd_economic_factor_file()
    return render_template("ifrs_master_data.html",
                           pd_category=pd_category,
                           pd_master_data=pd_master_data,
                           report_id=report_id,
                           run_date=run_date,
                           summary_files=summary_files,
                           overlay_file=overlay_file,
                           lgd_files=lgd_files,
                           pd_economic_factor_file=pd_economic_factor_file)


@app.route("/master-data", methods=["POST"])
def upload_master_data():
    """Save optional PD Category file, then trigger the report run."""
    report_id = request.form.get("report_id", "").strip()
    run_date  = request.form.get("run_date", "").strip()

    pd_master_values = {
        "nsw": request.form.get("pd_lease_nsw", "").strip(),
        "bcs": request.form.get("pd_lease_bcs", "").strip(),
        "wcs": request.form.get("pd_lease_wcs", "").strip(),
    }

    f = request.files.get("pd_category_file")
    if f and f.filename:
        MASTER_PD_CATEGORY_DIR.mkdir(parents=True, exist_ok=True)
        now = datetime.now()
        ts = f"{now.month}-{now.day}-{now.year}"
        dest_name = f"Test_PD_Cate - {ts}.xlsx"
        dest_path = MASTER_PD_CATEGORY_DIR / dest_name
        f.save(str(dest_path))

    # ── Summary master-data file uploads ──────────────────────────────
    summary_uploads = {}

    def _save_summary(field: str, category: str, filename: str) -> Path | None:
        fobj = request.files.get(field)
        if not fobj or not fobj.filename:
            return None
        d = MASTER_INPUTS_DIR / category
        d.mkdir(parents=True, exist_ok=True)
        p = d / filename
        if p.exists():
            p.unlink()
        fobj.save(str(p))
        return p

    # 1611
    f1611 = request.files.get("sum_1611_file")
    if f1611 and f1611.filename:
        month = request.form.get("sum_1611_month", "").strip()
        year  = request.form.get("sum_1611_year", "").strip()
        if month and year and year.isdigit() and len(year) == 4:
            p = _save_summary("sum_1611_file", "1611", f"1611-{month} {year}.xlsx")
            if p:
                summary_uploads["1611"] = str(p)

    # 2128
    f2128 = request.files.get("sum_2128_file")
    if f2128 and f2128.filename:
        month = request.form.get("sum_2128_month", "").strip()
        year  = request.form.get("sum_2128_year", "").strip()
        if month and year and year.isdigit() and len(year) == 4:
            p = _save_summary("sum_2128_file", "2128", f"2128-{month} {year}.xlsx")
            if p:
                summary_uploads["2128"] = str(p)

    # Assetline Related Party
    farp = request.files.get("sum_assetline_rp_file")
    if farp and farp.filename:
        p = _save_summary("sum_assetline_rp_file", "Assetline Related Party",
                          "Assetline Finance Limited - Related Party Register_All.xlsb")
        if p:
            summary_uploads["Assetline Related Party"] = str(p)

    # Cadre
    fcadre = request.files.get("sum_cadre_file")
    if fcadre and fcadre.filename:
        cadre_date = request.form.get("sum_cadre_date", "").strip()
        if cadre_date:
            date_obj = datetime.strptime(cadre_date, "%Y-%m-%d")
            date_display = date_obj.strftime("%d %m %Y")
            p = _save_summary("sum_cadre_file", "CadreSUM", f"Cadre as at {date_display}.xlsx")
            if p:
                summary_uploads["CadreSUM"] = str(p)

    # Insurance Premium
    fins = request.files.get("sum_insurance_file")
    if fins and fins.filename:
        month = request.form.get("sum_insurance_month", "").strip()
        year  = request.form.get("sum_insurance_year", "").strip()
        if month and year and year.isdigit() and len(year) == 4:
            p = _save_summary("sum_insurance_file", "Insurance Premium",
                              f"Insurance Premium {month} -{year[2:]}.xlsx")
            if p:
                summary_uploads["Insurance Premium"] = str(p)

    # Litigation Report
    flit = request.files.get("sum_litigation_file")
    if flit and flit.filename:
        lit_date = request.form.get("sum_litigation_date", "").strip()
        if lit_date:
            date_obj = datetime.strptime(lit_date, "%Y-%m-%d")
            p = _save_summary("sum_litigation_file", "Litigation Report",
                              f"Litigation Report - {date_obj.strftime('%d.%m.%Y')}.xlsx")
            if p:
                summary_uploads["Litigation Report"] = str(p)

    # Mediation Detail Report
    fmed = request.files.get("sum_mediation_file")
    if fmed and fmed.filename:
        month = request.form.get("sum_mediation_month", "").strip()
        year  = request.form.get("sum_mediation_year", "").strip()
        if month and year and year.isdigit() and len(year) == 4:
            p = _save_summary("sum_mediation_file", "Mediation Detail Report",
                              f"Mediation Detail Report - {month} {year}.xlsx")
            if p:
                summary_uploads["Mediation Detail Report"] = str(p)

    # Moratorium Contracts Details
    fmor = request.files.get("sum_moratorium_file")
    if fmor and fmor.filename:
        month = request.form.get("sum_moratorium_month", "").strip()
        year  = request.form.get("sum_moratorium_year", "").strip()
        if month and year and year.isdigit() and len(year) == 4:
            p = _save_summary("sum_moratorium_file", "Moratorium Contracts Details",
                              f"Moratorium Contracts details-{month}-{year}.xlsx")
            if p:
                summary_uploads["Moratorium Contracts Details"] = str(p)

    # Yard Final
    fyard = request.files.get("sum_yard_final_file")
    if fyard and fyard.filename:
        month = request.form.get("sum_yard_final_month", "").strip()
        year  = request.form.get("sum_yard_final_year", "").strip()
        if month and year and year.isdigit() and len(year) == 4:
            p = _save_summary("sum_yard_final_file", "Yard Final",
                              f"Yard Final - {month} {year[2:]}.xlsx")
            if p:
                summary_uploads["Yard Final"] = str(p)

    # ── PD master: Economic Factor ────────────────────────────────────
    fpd_econ = request.files.get("pd_economic_factor_file")
    if fpd_econ and fpd_econ.filename:
        MASTER_PD_ECONOMIC_FACTOR_DIR.mkdir(parents=True, exist_ok=True)
        # Clear existing files in the master dir before saving
        for _old in MASTER_PD_ECONOMIC_FACTOR_DIR.iterdir():
            if _old.is_file():
                try:
                    _old.unlink()
                except Exception:
                    pass
        econ_dest = MASTER_PD_ECONOMIC_FACTOR_DIR / "Economic Factors.xlsx"
        fpd_econ.save(str(econ_dest))
        # Copy to outputs/PD/<latest-date>/ with cleanup
        pd_sub = _find_pd_outputs_subfolder()
        if pd_sub:
            pd_sub.mkdir(parents=True, exist_ok=True)
            old_econ = pd_sub / "Economic Factors.xlsx"
            if old_econ.exists():
                old_econ.unlink()
            shutil.copy2(str(econ_dest), str(old_econ))

    # ── LGD master: Yard Final LGD ────────────────────────────────────
    flgd_yard = request.files.get("lgd_yard_final_file")
    if flgd_yard and flgd_yard.filename:
        lgd_yard_month = request.form.get("lgd_yard_final_month", "").strip()
        lgd_yard_year  = request.form.get("lgd_yard_final_year", "").strip()
        if lgd_yard_month and lgd_yard_year and lgd_yard_year.isdigit() and len(lgd_yard_year) == 4:
            MASTER_LGD_YARD_DIR.mkdir(parents=True, exist_ok=True)
            lgd_yard_name = f"Yard Final - {lgd_yard_month} {lgd_yard_year[2:]}.xlsx"
            # Remove all existing files in master dir before saving new one
            for _old in MASTER_LGD_YARD_DIR.iterdir():
                if _old.is_file():
                    try:
                        _old.unlink()
                    except Exception:
                        pass
            lgd_yard_dest = MASTER_LGD_YARD_DIR / lgd_yard_name
            flgd_yard.save(str(lgd_yard_dest))
            # Copy to outputs/LGD/<latest>/LGD/ with cleanup
            lgd_sub = _find_lgd_outputs_lgd_subfolder()
            if lgd_sub:
                lgd_sub.mkdir(parents=True, exist_ok=True)
                for _old in lgd_sub.glob("Yard Final*.xlsx"):
                    try:
                        _old.unlink()
                    except Exception:
                        pass
                shutil.copy2(str(lgd_yard_dest), str(lgd_sub / lgd_yard_name))

    # ── LGD master: Receipt Discount ──────────────────────────────────
    flgd_receipt_disc = request.files.get("lgd_receipt_discount_file")
    if flgd_receipt_disc and flgd_receipt_disc.filename:
        lgd_rd_month = request.form.get("lgd_receipt_discount_month", "").strip()
        if lgd_rd_month:
            MASTER_LGD_RECEIPT_DISCOUNT_DIR.mkdir(parents=True, exist_ok=True)
            # "Receipts Discount" (plural) matches the file naming already in
            # use in outputs/LGD — find_receipt_discount_file() in
            # LGD_computation.py globs "Receipt*Discount*" so it accepts
            # either spelling regardless.
            lgd_rd_name = f"Receipts Discount {lgd_rd_month}.xlsb"
            # Remove all existing files in master dir before saving new one
            for _old in MASTER_LGD_RECEIPT_DISCOUNT_DIR.iterdir():
                if _old.is_file():
                    try:
                        _old.unlink()
                    except Exception:
                        pass
            lgd_rd_dest = MASTER_LGD_RECEIPT_DISCOUNT_DIR / lgd_rd_name
            flgd_receipt_disc.save(str(lgd_rd_dest))
            # Copy to outputs/LGD/<latest>/LGD/ with cleanup — .xlsb, matches
            # load_recipts_discounts_dataframe()'s engine="pyxlsb" requirement.
            # Glob tolerates "Receipt"/"Receipts" so any prior variant is cleared.
            lgd_sub = _find_lgd_outputs_lgd_subfolder()
            if lgd_sub:
                lgd_sub.mkdir(parents=True, exist_ok=True)
                for _old in lgd_sub.glob("Receipt*Discount*.xlsb"):
                    try:
                        _old.unlink()
                    except Exception:
                        pass
                shutil.copy2(str(lgd_rd_dest), str(lgd_sub / lgd_rd_name))

    # ── LGD master: DC List ──────────────────────────────────────────
    flgd_dc_list = request.files.get("lgd_dc_list_file")
    if flgd_dc_list and flgd_dc_list.filename:
        lgd_dc_month = request.form.get("lgd_dc_list_month", "").strip()
        if lgd_dc_month:
            MASTER_LGD_DC_LIST_DIR.mkdir(parents=True, exist_ok=True)
            lgd_dc_name = f"DC List {lgd_dc_month}.xlsx"
            # Remove all existing files in master dir before saving new one
            for _old in MASTER_LGD_DC_LIST_DIR.iterdir():
                if _old.is_file():
                    try:
                        _old.unlink()
                    except Exception:
                        pass
            lgd_dc_dest = MASTER_LGD_DC_LIST_DIR / lgd_dc_name
            flgd_dc_list.save(str(lgd_dc_dest))
            # Copy to outputs/LGD/<latest>/LGD/ with cleanup — matches
            # LGD_computation.py's find_dc_list_file() glob "DC List*.xlsx"
            lgd_sub = _find_lgd_outputs_lgd_subfolder()
            if lgd_sub:
                lgd_sub.mkdir(parents=True, exist_ok=True)
                for _old in lgd_sub.glob("DC List*.xlsx"):
                    try:
                        _old.unlink()
                    except Exception:
                        pass
                shutil.copy2(str(lgd_dc_dest), str(lgd_sub / lgd_dc_name))

    # Overlay
    foverlay = request.files.get("sum_overlay_file")
    if foverlay and foverlay.filename:
        overlay_date = request.form.get("sum_overlay_date", "").strip()
        if overlay_date:
            MASTER_OVERLAY_DIR.mkdir(parents=True, exist_ok=True)
            # Remove all existing files in Master Inputs/Overlay before saving new one
            for old in MASTER_OVERLAY_DIR.iterdir():
                if old.is_file():
                    try:
                        old.unlink()
                    except Exception:
                        pass
            overlay_dest = MASTER_OVERLAY_DIR / f"Overlay-{overlay_date}.xlsx"
            foverlay.save(str(overlay_dest))
            # Copy to outputs/Summary/<latest-date>/Input with cleanup
            input_dir = _find_summary_outputs_input()
            if input_dir:
                input_dir.mkdir(parents=True, exist_ok=True)
                for old in input_dir.glob("Overlay-*.xlsx"):
                    try:
                        old.unlink()
                    except Exception:
                        pass
                shutil.copy2(str(overlay_dest), str(input_dir / overlay_dest.name))

    if summary_uploads:
        _copy_summary_files_to_working(summary_uploads)

    try:
        if any(v for v in pd_master_values.values()):
            _write_pd_master_values(pd_master_values)
        _pd_master_cache.clear()
        _pd_master_cache.update(pd_master_values)
    except Exception as exc:
        flash(f"Failed to save PD master data to Master_Data.xlsx: {exc}", "error")
        return redirect(url_for("master_data", report_id=report_id, date=run_date))

    # If called from the run flow, kick off the report
    if report_id and run_date and report_id in _REPORT_BY_ID:
        with _lock:
            if _state["running"]:
                flash("A report is already running. Please wait.", "warning")
                return redirect(url_for("master_data",
                                        report_id=report_id, date=run_date))
        _reset_state()
        report = _REPORT_BY_ID[report_id]
        t = threading.Thread(target=_run_worker, args=(report, run_date), daemon=True)
        t.start()
        return redirect(url_for("ifrs_dashboard"))

    flash("Master data saved.", "success")
    return redirect(url_for("master_data"))


_REPORT_OUTPUTS_SUBDIR = {
    "IFRS-01": "Summary",
    "IFRS-02": "PD",
    "IFRS-03": "LGD",
}


@app.route("/download-report/<report_id>", methods=["GET"])
def download_report(report_id):
    """Download the output files from a report's most recent completed run.

    Independent of the in-run _download_files/_state used by the
    auto-download-after-run flow, so it works as a standalone backup and
    can be clicked any time without disturbing an active/queued run.
    """
    subdir = _REPORT_OUTPUTS_SUBDIR.get(report_id)
    if subdir is None:
        return jsonify({"error": f"Unknown report: {report_id}"}), 404

    outputs_dir = OUTPUTS_DIR / subdir
    if not outputs_dir.exists():
        return jsonify({"error": "This report has not completed a run yet."}), 404

    date_folders = [p for p in outputs_dir.iterdir() if p.is_dir()]
    if not date_folders:
        return jsonify({"error": "This report has not completed a run yet."}), 404

    # Resolve the folder from completed_dates.txt (last line + highest "(n)"
    # suffix), like the rest of the pipeline. A folder's own mtime is NOT
    # reliable: it doesn't change when files inside its subfolders are replaced,
    # and sibling "(1)"/"(2)" folders can carry identical timestamps.
    latest_folder = None
    dates_file = outputs_dir / "completed_dates.txt"
    if dates_file.exists():
        try:
            lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
            if lines:
                latest_folder = _find_date_folder_with_variations(outputs_dir, lines[-1])
        except Exception:
            latest_folder = None
    if latest_folder is None:
        latest_folder = max(date_folders, key=lambda p: p.stat().st_mtime)

    # Prefer the manifest written by the auto-download flow for this run, so
    # this button selects the exact same files — not everything sitting in
    # the folder (which can include large intermediate/input files that were
    # archived alongside the real output, e.g. LGD's firstDefault inputs).
    files: list[Path] = []
    manifest_path = latest_folder / ".download_manifest.json"
    if manifest_path.exists():
        try:
            rel_paths = json.loads(manifest_path.read_text(encoding="utf-8")).get("files", [])
            files = [latest_folder / p for p in rel_paths if (latest_folder / p).is_file()]
        except Exception:
            files = []
    if not files:
        files = sorted(p for p in latest_folder.rglob("*") if p.is_file())
    if not files:
        return jsonify({"error": "No output files found for this report's last run."}), 404

    safe_id = report_id.replace("/", "_")

    if len(files) == 1:
        return send_file(str(files[0]), as_attachment=True, download_name=files[0].name)

    temp_zip_dir = BASE_DIR / "temp_zips"
    temp_zip_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = temp_zip_dir / f"{safe_id}_{timestamp}.zip"

    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for f in files:
                zipf.write(f, arcname=f.relative_to(latest_folder))
    except Exception as e:
        return jsonify({"error": f"Failed to prepare files for download: {e}"}), 500

    download_name = f"{safe_id}_{latest_folder.name}.zip"
    return send_file(str(zip_path), as_attachment=True, download_name=download_name)


@app.route("/download-latest", methods=["GET"])
def download_latest():
    """Serve the queued download file(s) and reset state."""
    global _download_files

    if not _download_files:
        return redirect(url_for("ifrs_dashboard"))

    path = _download_files[0]
    if not path.exists():
        _download_files = []
        return redirect(url_for("ifrs_dashboard"))

    if path.suffix.lower() in (".xlsx", ".xlsb", ".xlsm"):
        download_name = path.name
    else:
        download_name = path.name  # zip already has a descriptive name

    _download_files = []

    with _lock:
        _state["stage"] = "idle"
        _state["message"] = ""

    return send_file(str(path), as_attachment=True, download_name=download_name)


@app.route("/scheduler-log")
def scheduler_log():
    """Return recent scheduler log entries as JSON for the dashboard panel."""
    limit = int(request.args.get("limit", 50))
    return jsonify({
        "entries": _sched.read_log(limit),
        "jobs":    _sched.get_scheduled_jobs(),
    })


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5002, debug=True)
