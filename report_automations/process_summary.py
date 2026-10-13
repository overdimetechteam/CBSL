import pandas as pd
import numpy as np
import os
import sys
import glob
import subprocess
import time
import traceback
from contextlib import contextmanager
from datetime import datetime, timedelta
import argparse
from pathlib import Path
import logging
from openpyxl import load_workbook

from FixedLoans import create_fixed_loans_df
from MT import create_marginal_loans_df
from FDLquarter import create_FDL_quarter_df
from disbursement_processor import get_disbursement_df
from ltv_working_disbursement import create_ltv_working_disbursement_df
from unutilized import update_unutilized_sheet
from mt_report_consolidator import consolidate_mt_reports, finalize_mt_workbook
from Loan_Base_Report_transfer import download_loan_base_report_and_trf

# Force unbuffered output for proper logging when run as subprocess
# errors='replace': when run as a subprocess stdout is a cp1252 pipe, and an
# imported module printing e.g. "✓" (bots/erp_login.py does at import) would
# otherwise raise UnicodeEncodeError and abort that step.
sys.stdout.reconfigure(line_buffering=True, errors='replace')
sys.stderr.reconfigure(line_buffering=True, errors='replace')

# Module-level logger placeholder; configured for real in setup_logging() once
# --log-dir is known. Referenced as a global by helper functions below.
logger = logging.getLogger(__name__)

REQUIRED_FILE_KEYS = [
    'file_path', 'file_path_portfolio', 'file_path_netPortfolio', 'cbsl_provision',
    'file_path_cbsl_sec', 'file_path_district', 'file_path_micro', 'file_path_reschedule',
    'file_path_monthlyreport', 'file_path_fixedloans', 'file_path_MT', 'report_file',
    'file_path_fdl_prev', 'file_path_fdl_prev2', 'file_path_unutilized_credit',
]


def setup_logging(log_dir_path=None):
    """Setup logging to both console and file"""
    if log_dir_path:
        log_dir = Path(log_dir_path)
    else:
        script_root = Path(__file__).parent.parent
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = script_root / "logs" / f"run_{timestamp}"

    handlers = [logging.StreamHandler(sys.stdout)]
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / "process_summary.log"
        handlers.insert(0, logging.FileHandler(log_file, encoding='utf-8'))
    except OSError as e:
        # Don't let a logging-setup problem abort the whole run; fall back to
        # console-only logging and keep going.
        print(f"Warning: could not set up file logging in '{log_dir}': {e}. "
              f"Continuing with console logging only.", flush=True)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%H:%M:%S',
        handlers=handlers,
        force=True,
    )

    return logging.getLogger(__name__)


def log(message):
    """Info-level log, written to console and (if available) the log file."""
    logger.info(message)


@contextmanager
def log_step(name):
    """Wrap a pipeline phase: logs start/duration and captures failures with
    full context (step name + traceback) before propagating them."""
    logger.info(f"--- {name} ---")
    start = time.perf_counter()
    try:
        yield
    except Exception:
        logger.exception(f"FAILED: {name}")
        raise
    else:
        logger.info(f"Done: {name} ({time.perf_counter() - start:.1f}s)")


def find_file_with_pattern(folder_path, pattern):
    """Find file using glob pattern within folder_path (including subfolders) and return the first match"""
    search_pattern = os.path.join(folder_path, "**", pattern)
    matches = glob.glob(search_pattern, recursive=True)
    if matches:
        if len(matches) > 1:
            logger.warning(f"Multiple files match pattern '{pattern}'; using first: {matches[0]} "
                            f"(others: {matches[1:]})")
        return matches[0]
    logger.warning(f"No file found matching pattern: {pattern}")
    return None


def find_newest_file_with_pattern(folder_path, pattern):
    """Like find_file_with_pattern, but returns the most recently modified
    match - for inputs refreshed by a download step, where an older copy may
    still be present."""
    matches = [m for m in glob.glob(os.path.join(folder_path, "**", pattern), recursive=True)
               if not os.path.basename(m).startswith("~$")]
    if not matches:
        logger.warning(f"No file found matching pattern: {pattern}")
        return None
    newest = max(matches, key=os.path.getmtime)
    if len(matches) > 1:
        logger.warning(f"Multiple files match pattern '{pattern}'; using newest: {newest}")
    return newest


def generate_dynamic_file_paths(input_folder, output_folder, current_month_str, current_month_name,
                                 next_month_name, prev_month_name, last_q_monthname,
                                 current_month_name_full, prev_month_name_full, prev2_month_name_full,
                                 prev2_month_name):
    """Generate all file paths dynamically based on GUI date from app.py"""
    return {
        'file_path': find_file_with_pattern(input_folder, f"3. Summary*.xlsb"),
        'file_path_portfolio': find_file_with_pattern(input_folder, f"1. Portfolio_4_ECL*.xlsb"),
        'file_path_netPortfolio': find_file_with_pattern(input_folder, f"Net Portfolio*.xlsx"),
        'cbsl_provision': find_file_with_pattern(input_folder, f"CBSL Provision Comparison*.xlsb"),
        'file_path_cbsl_sec': find_file_with_pattern(input_folder, f"Sector Report*.xlsb"),
        'file_path_district': find_file_with_pattern(input_folder, "JOBHD*.xlsx"),
        'file_path_micro': find_file_with_pattern(input_folder, "Copy of Information Request from Credit*.xlsx"),
        'file_path_reschedule': find_newest_file_with_pattern(input_folder, "Reschedule Contract Details - Internal*.xlsx"),
        # Pinned to the current month's name: the Input folder now also holds the
        # previous two months' "Prod. wise Class. of Loans" files (for FDL Quarter
        # below), so an unqualified pattern would match ambiguously.
        # Whoever saves these files isn't consistent about full vs abbreviated
        # month names (e.g. "September 2025" vs "Sep 2025") - both are tried;
        # Windows' filesystem is already case-insensitive, so "SEP"/"sep"/"Sep"
        # all match the same pattern without any extra handling.
        'file_path_monthlyreport': (
            find_file_with_pattern(input_folder, f"Prod. wise Class. of Loans - {current_month_name_full}*.xlsb")
            or find_file_with_pattern(input_folder, f"Prod. wise Class. of Loans - {current_month_name}*.xlsb")
        ),
        'file_path_fixedloans': find_file_with_pattern(input_folder, "LoanAsAtDate*.csv") or find_file_with_pattern(input_folder, "LoanAsAtDate*.xlsx"),
        'file_path_MT': find_file_with_pattern(input_folder, f"Unutilized Amount*.xlsx"),
        'file_path_unutilized_credit': find_file_with_pattern(input_folder, f"Unutilized Credit Limits*.xlsx"),
        'report_file': find_file_with_pattern(output_folder, f"Quarter Classification Loans & Advance*.xlsx"),
        # Previous two months' "Prod. wise Class. of Loans" files, used only to
        # build the FDL Quarter sheet's 3-month view (see FDLquarter.py).
        'file_path_fdl_prev': (
            find_file_with_pattern(input_folder, f"Prod. wise Class. of Loans - {prev_month_name_full}*.xlsb")
            or find_file_with_pattern(input_folder, f"Prod. wise Class. of Loans - {prev_month_name}*.xlsb")
        ),
        'file_path_fdl_prev2': (
            find_file_with_pattern(input_folder, f"Prod. wise Class. of Loans - {prev2_month_name_full}*.xlsb")
            or find_file_with_pattern(input_folder, f"Prod. wise Class. of Loans - {prev2_month_name}*.xlsb")
        ),
    }


def validate_file_paths(file_paths_dict):
    """Fail fast with one clear message instead of a cryptic pandas/openpyxl
    error the first time a missing path (None) is used."""
    missing = [key for key in REQUIRED_FILE_KEYS if not file_paths_dict.get(key)]
    if missing:
        raise FileNotFoundError(
            "Required input file(s) could not be located: " + ", ".join(missing) +
            ". Check that the Input folder contains files matching the expected naming "
            "pattern for this report date, and that the file names haven't changed."
        )


def resolve_working_directory(args, base_working_dir):
    if args.working_dir:
        provided_path = Path(args.working_dir)
        if not provided_path.exists() or not provided_path.is_dir():
            raise FileNotFoundError(f"Working directory not found: {provided_path}")

        if provided_path.parent.name == "NBD-QF-23-C3-C10-SF-FI":
            date_folder = provided_path
        else:
            subdirs = [d for d in provided_path.iterdir() if d.is_dir() and not d.name.startswith('.')]
            if subdirs:
                if len(subdirs) > 1:
                    raise ValueError(f"Multiple date folders found in {provided_path}. app.py should ensure only one exists.")
                date_folder = subdirs[0]
            else:
                date_folder = provided_path
    else:
        if not base_working_dir.exists():
            raise FileNotFoundError(f"Working directory not found: {base_working_dir}")

        date_folders = [d for d in base_working_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]

        if not date_folders:
            raise FileNotFoundError(f"No date folders found in: {base_working_dir}. app.py should copy the folder first.")

        if len(date_folders) > 1:
            raise ValueError(f"Multiple date folders found in {base_working_dir}. app.py should ensure only one exists. "
                              f"Found: {[d.name for d in date_folders]}")

        date_folder = date_folders[0]
        logger.info(f"Using date folder (already copied by app.py): {date_folder.name}")

    return date_folder


def resolve_report_date(report_date_str):
    if report_date_str:
        try:
            report_date_obj = datetime.strptime(report_date_str, "%Y-%m-%d")
            logger.info(f"Using GUI date for file operations: {report_date_str}")
            return report_date_obj
        except ValueError:
            logger.warning(f"Could not parse report-date '{report_date_str}', using current date instead")
            return datetime.now()

    fallback = datetime.now()
    logger.warning(f"No report-date provided, using current date: {fallback.strftime('%Y-%m-%d')}")
    return fallback


def read_excel_safe(path, description, **kwargs):
    """pd.read_excel wrapper that logs progress/shape and turns any failure
    into a message that names the file and what we were trying to read."""
    logger.info(f"Reading {description}: {path}")
    try:
        df = pd.read_excel(path, **kwargs)
    except Exception as e:
        raise RuntimeError(f"Failed to read {description} from '{path}': {e}") from e
    logger.info(f"Loaded {description}: {df.shape[0]} rows x {df.shape[1]} columns")
    return df


def warn_if_nulls(df, column, label):
    """Surface silent map()/merge misses (unmatched keys become NaN) instead
    of letting them slip into the report unnoticed."""
    if column not in df.columns:
        return
    n_null = int(df[column].isna().sum())
    if n_null:
        logger.warning(f"{label}: {n_null} of {len(df)} row(s) could not be matched (null '{column}')")


def convert_xlsb_to_xlsx(report_file):
    import xlwings as xw
    logger.info("Converting .xlsb to .xlsx via Excel (xlwings)...")
    app = xw.App(visible=False)
    try:
        wb_temp = xw.Book(report_file)
        xlsx_file = report_file.replace('.xlsb', '.xlsx')
        wb_temp.save(xlsx_file)
        wb_temp.close()
        return xlsx_file
    except Exception as e:
        raise RuntimeError(f"Failed to convert report file '{report_file}' from .xlsb to .xlsx: {e}") from e
    finally:
        app.quit()


def save_workbook_safe(wb, path):
    try:
        # openpyxl writes formula text only, never a cached result. Without
        # this, Excel can show a stale/error state (e.g. #NAME?) for freshly
        # pasted formulas until the cell is manually touched (click + Enter).
        # Forcing auto mode + a full recalc on load makes every formula
        # evaluate immediately when the file is opened.
        wb.calculation.calcMode = 'auto'
        wb.calculation.fullCalcOnLoad = True
        wb.save(str(path))
    except PermissionError as e:
        raise PermissionError(
            f"Could not save '{path}' - the file may currently be open in Excel. Close it and re-run."
        ) from e
    except Exception as e:
        raise RuntimeError(f"Failed to save workbook to '{path}': {e}") from e


def write_dataframe_to_sheet(ws, df, start_row, start_col=1):
    """Shared cell-by-cell writer used by every 'paste dataframe into sheet'
    step below (keeps the formula-vs-value handling in one place)."""
    clean_df = df.replace([np.inf, -np.inf], np.nan).fillna('')
    for i, row in enumerate(clean_df.values, start=start_row):
        for j, val in enumerate(row, start=start_col):
            cell = ws.cell(row=i, column=j)
            if isinstance(val, str) and val.startswith('='):
                cell.value = val  # openpyxl auto-detects formulas
            else:
                cell.value = val.item() if isinstance(val, (np.integer, np.floating)) else val
    return clean_df


def update_moi_reschedule_block(ws, df_reschedule):
    """Replace the reschedule block of the 'MOI List -31 Jan' sheet (everything
    below the first blank row in column A) with CLO_NEWCONNO from the Reschedule
    Contract Details file, sorted A-Z, each marked "Yes" in column B. Rows above
    the blank row (the moratorium list) are left untouched."""
    blank_row = None
    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if value is None or str(value).strip() == '':
            blank_row = row
            break
    if blank_row is None:
        # No separator yet - keep the existing list and start a new block after it.
        blank_row = ws.max_row + 1
        logger.warning(f"No blank separator row found in MOI sheet; using row {blank_row}")

    old_contracts = {
        str(ws.cell(row=row, column=1).value).strip()
        for row in range(blank_row + 1, ws.max_row + 1)
        if ws.cell(row=row, column=1).value not in (None, '')
    }

    new_contracts = sorted(
        {str(v).strip() for v in df_reschedule['CLO_NEWCONNO'].dropna() if str(v).strip()},
        key=str.upper,
    )

    for row in range(blank_row + 1, ws.max_row + 1):
        ws.cell(row=row, column=1).value = None
        ws.cell(row=row, column=2).value = None

    for i, contract in enumerate(new_contracts, start=blank_row + 1):
        ws.cell(row=i, column=1).value = contract
        ws.cell(row=i, column=2).value = "Yes"

    added = len(set(new_contracts) - old_contracts)
    removed = len(old_contracts - set(new_contracts))
    log(f"MOI reschedule block (from row {blank_row + 1}): {len(old_contracts)} -> "
        f"{len(new_contracts)} contracts (+{added} added, -{removed} removed)")


def _run():
    global logger

    script_root = os.path.dirname(os.path.abspath(__file__))
    parent_directory = Path(script_root).parent

    parser = argparse.ArgumentParser(description='Process summary - runs after app.py completes copying and master data collection')
    parser.add_argument('--working-dir', type=str, help='Path to working date folder (optional, will auto-detect if not provided)')
    parser.add_argument('--report-date', type=str, help='Report date in YYYY-MM-DD format (from app.py GUI, used for file naming)')
    parser.add_argument('--log-dir', type=str, help='Path to log directory (from app.py timestamp folder)')
    args, unknown = parser.parse_known_args()

    logger = setup_logging(args.log_dir)
    logger.info("=" * 70)
    logger.info("Process Summary - starting run")
    if unknown:
        logger.debug(f"Ignoring unrecognized CLI arguments: {unknown}")

    with log_step("Resolve working directory"):
        base_working_dir = parent_directory / "working" / "NBD-QF-23-C3-C10-SF-FI"
        date_folder = resolve_working_directory(args, base_working_dir)

    output_path = date_folder
    log(f"Working directory: {output_path}")

    with log_step("Resolve report date"):
        report_date_obj = resolve_report_date(args.report_date)

    with log_step("Refresh Loan Base Report (Scienter download)"):
        try:
            capital_balance = download_loan_base_report_and_trf(report_date_obj.strftime("%Y-%m-%d"))
            if capital_balance is not None:
                logger.info(f"Loan Base Report refreshed. Capital Balance: {capital_balance:,.2f}")
            else:
                logger.warning("Loan Base Report download did not complete; continuing with whatever "
                                "LoanAsAtDate file is already present in the Input folder.")
        except Exception as e:
            logger.warning(f"Loan Base Report refresh failed ({e}); continuing with whatever "
                            f"LoanAsAtDate file is already present in the Input folder.")

    with log_step("Refresh Reschedule Contract Details (ERP download)"):
        try:
            if str(parent_directory) not in sys.path:
                sys.path.insert(0, str(parent_directory))
            from bots.reschedule_contract_details_bot import download_reschedule_contract_details
            new_file = download_reschedule_contract_details(date_folder / "Input")
            logger.info(f"Reschedule Contract Details refreshed: {new_file}")
        except Exception as e:
            logger.warning(f"Reschedule Contract Details download failed ({e}); continuing with whatever "
                            f"Reschedule Contract Details file is already present in the Input folder.")

    current_month_date = report_date_obj
    current_month_str = current_month_date.strftime("%Y-%m-%d")
    current_month_name = current_month_date.strftime("%b %Y")
    current_month_short = current_month_date.strftime("%b %y")
    current_month_name_dash = current_month_date.strftime("%b-%Y")

    if current_month_date.month == 1:
        prev_month = current_month_date.replace(year=current_month_date.year - 1, month=12)
    else:
        prev_month = current_month_date.replace(month=current_month_date.month - 1)
    prev_month_name = prev_month.strftime("%b %Y")

    if prev_month.month == 1:
        prev2_month = prev_month.replace(year=prev_month.year - 1, month=12)
    else:
        prev2_month = prev_month.replace(month=prev_month.month - 1)
    prev2_month_name = prev2_month.strftime("%b %Y")

    # Full month names ("September 2025") to match the "Prod. wise Class. of
    # Loans - <Month> <Year>" file-naming convention, for FDL Quarter's 3-month source files.
    # Whoever saves these files month to month isn't consistent about full
    # ("September 2025") vs abbreviated ("Sep 2025") - both are tried below.
    current_month_name_full = current_month_date.strftime("%B %Y")
    prev_month_name_full = prev_month.strftime("%B %Y")
    prev2_month_name_full = prev2_month.strftime("%B %Y")

    if current_month_date.month == 12:
        next_month = current_month_date.replace(year=current_month_date.year + 1, month=1)
    else:
        next_month = current_month_date.replace(month=current_month_date.month + 1)
    next_month_name = next_month.strftime("%b-%Y")

    last_q_monthname = (current_month_date.replace(day=1) - timedelta(days=90)).strftime("%b %Y")

    INPUT_FOLDER = str(date_folder / "Input")
    OUTPUT_FOLDER = str(date_folder)

    with log_step("Resolve and validate input file paths"):
        log(f"Looking for files in: {INPUT_FOLDER}")
        log(f"Date folder: {date_folder.name}")

        file_paths_dict = generate_dynamic_file_paths(
            INPUT_FOLDER, OUTPUT_FOLDER, current_month_str, current_month_name,
            next_month_name, prev_month_name, last_q_monthname,
            current_month_name_full, prev_month_name_full, prev2_month_name_full,
            prev2_month_name,
        )

        log("File paths resolved:")
        for key, path in file_paths_dict.items():
            status = "found" if path else "NOT FOUND"
            log(f"  [{status}] {key}: {path if path else ''}")

        validate_file_paths(file_paths_dict)

    file_path = file_paths_dict['file_path']
    file_path_portfolio = file_paths_dict['file_path_portfolio']
    file_path_netPortfolio = file_paths_dict['file_path_netPortfolio']
    cbsl_provision_path = file_paths_dict['cbsl_provision']
    file_path_cbsl_sec = file_paths_dict['file_path_cbsl_sec']
    file_path_district = file_paths_dict['file_path_district']
    file_path_micro = file_paths_dict['file_path_micro']
    file_path_reschedule = file_paths_dict['file_path_reschedule']
    file_path_monthlyreport = file_paths_dict['file_path_monthlyreport']
    file_path_fixedloans = file_paths_dict['file_path_fixedloans']
    file_path_MT = file_paths_dict['file_path_MT']
    file_path_unutilized_credit = file_paths_dict['file_path_unutilized_credit']
    report_file = file_paths_dict['report_file']
    file_path_fdl_prev = file_paths_dict['file_path_fdl_prev']
    file_path_fdl_prev2 = file_paths_dict['file_path_fdl_prev2']

    with log_step("Load source Excel files"):
        df = read_excel_safe(file_path, "Summary file", sheet_name='SUMMARY', skiprows=2, engine='pyxlsb')
        portfolio_df = read_excel_safe(file_path_portfolio, "Portfolio file", sheet_name='Portfolio', skiprows=2, engine='pyxlsb')
        net_portfolio_df = read_excel_safe(file_path_netPortfolio, "Net Portfolio file")
        cbsl_provision = read_excel_safe(cbsl_provision_path, "CBSL Provision file", sheet_name='Portfolio', skiprows=2, engine='pyxlsb')
        df_property_mortgage = read_excel_safe(
            cbsl_provision_path, "CBSL Provision file (PropertyMortgage sheet)",
            sheet_name='PropertyMortgage', skiprows=5, header=None, usecols='A', engine='pyxlsb'
        )
        df_property_mortgage.columns = ['Contract No']
        df_property_mortgage = df_property_mortgage.dropna(subset=['Contract No']).reset_index(drop=True)

    with log_step("Build main dataframe"):
        try:
            stage_col_index = df.columns.get_loc('STAGE')
        except KeyError as e:
            raise KeyError(f"Expected column 'STAGE' not found in Summary file. "
                            f"The input file format may have changed. Columns present: {list(df.columns)}") from e

        df = df.iloc[:, :stage_col_index + 1]

        try:
            maindf = df.drop(columns=['RELATED-PARTY CONTRACT', 'STAFF CONTRACT', 'ACTUAL LEASE/LOAN'])
        except KeyError as e:
            raise KeyError(f"Expected column(s) missing from Summary file: {e}") from e

        maindf = maindf.dropna(subset=['CONTRACT NO'])
        log(f"Main dataframe after cleanup: {maindf.shape[0]} rows")

        contract_to_freq = dict(zip(net_portfolio_df["CONTRACT_NO"], net_portfolio_df["CON_RNTFREQ"]))
        contract_to_sector = dict(zip(net_portfolio_df["CONTRACT_NO"], net_portfolio_df["SEC_DESC"]))
        contract_to_sub_sec = dict(zip(net_portfolio_df["CONTRACT_NO"], net_portfolio_df["SUB_SECTOR"]))

        contract_to_prov_guid = dict(zip(cbsl_provision["CONTRACT NO"], cbsl_provision["FINAL PROVISION (CBSL GUIDELINE)"]))
        contract_to_prov_pnp = dict(zip(cbsl_provision["CONTRACT NO"], cbsl_provision["CBSL P/NP"]))

        # Map the frequency values to the main dataframe
        maindf['Frequency'] = maindf['CONTRACT NO'].map(contract_to_freq)
        warn_if_nulls(maindf, 'Frequency', "Frequency mapping from Net Portfolio")

        # Add Repayment Cycle column with Excel formulas
        # Excel row 3 corresponds to df index 0 (since we skipped 2 rows)
        maindf['Repayment Cycle'] = [f'=IF(OR(AE{i+3}="M",AE{i+3}="D"),"Monthly Basis","Other")' for i in range(len(maindf))]
        # MAXIFS is a post-2007 "future function" - openpyxl writes formula
        # text verbatim, so it needs the _xlfn. prefix or Excel shows #NAME?
        # until the cell is manually re-entered (same reason Gender uses
        # _xlfn.XLOOKUP below).
        maindf['CBSL DPD'] = [f'=_xlfn.MAXIFS(F:F,B:B,B{i + 3})' for i in range(len(maindf))]

        # DPD-based CBSL provision bucketing off CBSL DPD's own output cell
        # (column AG, where the formula above lands):
        #   <=90 Performing, 91-180 Special Mention, 181-270 Substandard,
        #   271-360 Doubtful, >360 Loss.
        maindf['Provision Category'] = [
            f'=IF(AG{i+3}<=90,"Performing",'
            f'IF(AND(AG{i+3}>90,AG{i+3}<=180),"Special Mention",'
            f'IF(AND(AG{i+3}>180,AG{i+3}<=270),"Substandard",'
            f'IF(AND(AG{i+3}>270,AG{i+3}<=360),"Doubtful",'
            f'IF(AG{i+3}>360,"Loss",0)))))'
            for i in range(len(maindf))
        ]
        maindf['Provision - CBSL Guideline'] = maindf['CONTRACT NO'].map(contract_to_prov_guid)
        maindf['CBSL P/NP'] = maindf['CONTRACT NO'].map(contract_to_prov_pnp)
        maindf['Forbone Loans (Yes/No)'] = [f'=IF(ISNUMBER(MATCH(A{i + 3},\'MOI List -31 Jan\'!A:A, 0)), "Yes", "No")' for i in range(len(maindf))]
        maindf['EXP+DP-IIS'] = [f'=Y{i + 3}-T{i + 3}' for i in range(len(maindf))]
        maindf['Gross-Imp-IIS'] = [f'=E{i+3}-J{i+3}-T{i+3}' for i in range(len(maindf))]
        maindf['Exp+DP-IMP-IIS'] = [f'=Y{i+3}-Z{i+3}-T{i+3}' for i in range(len(maindf))]
        maindf['Collateral/Security Type'] = [f'=IF(ISNUMBER(MATCH(A{i+3},\'Property Mortgage List\'!A:A,0)),"Immovable Properties",IF(OR(AA{i+3}="UV",AA{i+3}="LE",AA{i+3}="AT"),"Vehicles and Machinery",IF(ISNUMBER(MATCH(R{i+3},Cat_List!$A$2:$A$9,0)),"Vehicles and Machinery",IF(A{i+3}="Margin Trading","Shares and Debt Securities-Listed","Personal and Corporate Guarantees"))))' for i in range(len(maindf))]
        maindf['Sector'] = maindf['CONTRACT NO'].map(contract_to_sector)
        maindf['Sub-Sector'] = maindf['CONTRACT NO'].map(contract_to_sub_sec)
        warn_if_nulls(maindf, 'Sector', "Sector mapping from Net Portfolio")

        df_sector = read_excel_safe(file_path_cbsl_sec, "CBSL Sector report", sheet_name='Portfolio', skiprows=1, engine='pyxlsb')
        contract_to_cbsl_sec = dict(zip(df_sector["Contract No"], df_sector["CBSL Sector 1 Final"]))
        maindf['CBSL Sector'] = maindf['CONTRACT NO'].map(contract_to_cbsl_sec)
        warn_if_nulls(maindf, 'CBSL Sector', "CBSL Sector mapping")

        df_district = read_excel_safe(file_path_district, "District (JOBHD) file")
        contract_to_district = dict(zip(df_district["CLM_CODE"], df_district["DISTRICT"]))
        maindf['CustomerDistrict(ClientMain)'] = maindf['CLIENT NO'].map(contract_to_district)

        # Clean up column names
        net_portfolio_df.columns = net_portfolio_df.columns.str.strip()

        df_unique = net_portfolio_df.drop_duplicates(subset=['CLIENT_CODE'], keep='first')
        contract_to_client_name = dict(zip(df_unique["CLIENT_CODE"], df_unique["CLM_NAME"]))
        maindf['Customer Name'] = maindf['CLIENT NO'].map(contract_to_client_name)
        warn_if_nulls(maindf, 'Customer Name', "Customer Name mapping")

        df_micro_dis = read_excel_safe(file_path_micro, "Micro/Small/Medium disbursement sheet", sheet_name=0, skiprows=4)
        df_micro_port = read_excel_safe(file_path_micro, "Micro/Small/Medium portfolio sheet", sheet_name=1, skiprows=3)

        df1_relevant = df_micro_port[['Client Code', 'Contract No', 'Product', 'Customer Name', 'Contract Amount', 'Micro/Small/Medium', 'Remarks']]
        df2_relevant = df_micro_dis[['Client Code', 'Contract No', 'Product', 'Customer Name', 'Net Monthly Disbursement (LKR)', 'Micro/Small/Medium']]
        combined_df = pd.concat([df1_relevant, df2_relevant], ignore_index=True)

        contract_to_ct_micro = dict(zip(combined_df["Contract No"], combined_df["Micro/Small/Medium"]))
        maindf['Micro/Small/Medium'] = maindf['CONTRACT NO'].map(contract_to_ct_micro)

        maindf['Gender'] = [f'=IF(LEFT(B{i + 3},1)="2",_xlfn.XLOOKUP(B{i + 3},BusinessGender!A:A,BusinessGender!C:C,"No Data"),IF(LEFT(AT{i + 3},3)="Mr.","Male",IF(LEFT(AT{i + 3},3)="Mr ","Male",IF(LEFT(AT{i + 3},3)="Rev","Male",IF(LEFT(AT{i + 3},4)="Miss","Female",IF(LEFT(AT{i + 3},3)="Ms.","Female",IF(LEFT(AT{i + 3},4)="Mrs.","Female")))))))' for i in range(len(maindf))]

        df_reschedule = read_excel_safe(file_path_reschedule, "Reschedule Contract Details file")
        contract_to_old_contract = dict(zip(df_reschedule["CLO_NEWCONNO"], df_reschedule["CON_NO"]))
        maindf['Old Contract No (Before Reschedule)'] = maindf['CONTRACT NO'].map(contract_to_old_contract)

        contract_to_grantAm = dict(zip(net_portfolio_df["CONTRACT_NO"], net_portfolio_df["CONTRACT_AMOUNT"]))
        maindf['Grant Amount'] = maindf['CONTRACT NO'].map(contract_to_grantAm)

        df_monthly_report = read_excel_safe(file_path_monthlyreport, "Product-wise Classification of Loans file",
                                             sheet_name='C1 & C2 Working', skiprows=1, engine='pyxlsb')
        contract_to_initialVal = dict(zip(df_monthly_report["Contract No"], df_monthly_report["Initial Valuation"]))
        maindf['Initial Valuation'] = maindf['CONTRACT NO'].map(contract_to_initialVal)

        maindf['LTV %'] = np.where(maindf['Initial Valuation'].isna() | maindf['Grant Amount'].isna(), np.nan,
                                    maindf['Grant Amount'] / maindf['Initial Valuation'])

        maindf['WALTV %'] = [f'=IF(AO{i+3}="Vehicles and Machinery",(E{i+3}/$AZ$1)*AZ{i+3},0)' for i in range(len(maindf))]

        fixedloan = create_fixed_loans_df(file_path_fixedloans, df_monthly_report, len(maindf))
        fixedloan.columns = maindf.columns
        maindf = pd.concat([maindf, fixedloan], ignore_index=True)
        maindf['Top 50 Clients'] = [f'=VLOOKUP(B{i+3},\'Top50\'!A:E,5,0)' for i in range(len(maindf))]

        marginal_loan = create_marginal_loans_df(file_path_MT, df_district, combined_df)
        # FDL Quarter sheet: one file per month, newest first, each read from its
        # own "IA Working" sheet - see create_FDL_quarter_df for the filter logic.
        fdl_month_files = [
            (current_month_date.strftime('%B'), file_path_monthlyreport),
            (prev_month.strftime('%B'), file_path_fdl_prev),
            (prev2_month.strftime('%B'), file_path_fdl_prev2),
        ]
        df_FDL_quarter = create_FDL_quarter_df(fdl_month_files)
        # Disbursement sheet: same 3 monthly IA Working files/month labels as FDL Quarter.
        disbursement_df = get_disbursement_df(fdl_month_files, maindf, df_sector)
        # LTV Working - Disbursement sheet: Vehicles and Machinery-collateral
        # disbursements sourced straight from each of the quarter's 3 monthly
        # IA Working files (same file set as FDL Quarter), not filtered by
        # activation date - see create_ltv_working_disbursement_df.
        df_ltv_working_disbursement = create_ltv_working_disbursement_df(fdl_month_files, maindf)
        # MT workbook (MT - <Mon YYYY>.xlsx in Input/MT). Not used by this report's
        # sheets, so a failure here (e.g. a missing upload) shouldn't stop the run.
        mt_workbook = None
        try:
            mt_workbook = consolidate_mt_reports(report_date_obj, working_dir=date_folder)
        except Exception as e:
            logger.warning(f"MT workbook consolidation failed ({e}); continuing without it")

        log("Data loaded successfully!")

        log(f"Disbursement shape: {disbursement_df.shape}")
        log(f"LTV Working - Disbursement shape: {df_ltv_working_disbursement.shape}")
        log(f"Main dataframe shape: {maindf.shape}")

    with log_step(f"Load report workbook: {report_file}"):
        if report_file.endswith('.xlsb'):
            report_file = convert_xlsb_to_xlsx(report_file)
        try:
            wb = load_workbook(report_file)
        except Exception as e:
            raise RuntimeError(f"Failed to open report workbook '{report_file}': {e}") from e

    with log_step("Update Summary sheet"):
        if 'Summary' in wb.sheetnames:
            ws_summary = wb['Summary']
            ws_summary.delete_rows(3, ws_summary.max_row - 2)
            write_dataframe_to_sheet(ws_summary, maindf, start_row=3, start_col=1)
            log("Updated SUMMARY sheet")
        else:
            logger.warning("'Summary' sheet not found in report workbook - skipping")

    with log_step("Update MT sheet"):
        if 'MT' in wb.sheetnames:
            ws_mt = wb['MT']
            ws_mt.delete_rows(2, ws_mt.max_row - 1)
            write_dataframe_to_sheet(ws_mt, marginal_loan, start_row=2, start_col=3)
            log("Updated MT sheet")
        else:
            logger.warning("'MT' sheet not found in report workbook - skipping")

    with log_step("Update FDL Quarter sheet"):
        if 'FDL Quarter' in wb.sheetnames:
            ws_fdl = wb['FDL Quarter']
            ws_fdl.delete_rows(3, ws_fdl.max_row - 2)
            write_dataframe_to_sheet(ws_fdl, df_FDL_quarter, start_row=3, start_col=1)
            # delete_rows() above drops the deleted rows' number formats along
            # with their values, so percentage columns (Minimum Rate,
            # Contractual Interest Rate, Base Rate, EIR (%), LTV %, WALTV %)
            # would otherwise render as plain decimals instead of percentages.
            last_data_row = max(3, 2 + len(df_FDL_quarter))
            for row in range(3, last_data_row + 1):
                for col in (13, 14, 15, 19, 24, 25):
                    ws_fdl.cell(row=row, column=col).number_format = '0.00%'
            # Row 1's WALTV % total (col Y) sums a fixed row range left over
            # from the template - keep it in sync with the data just written.
            ws_fdl.cell(row=1, column=25).value = f'=SUM(Y3:Y{last_data_row})'
            log("Updated FDL Quarter sheet")
        else:
            logger.warning("'FDL Quarter' sheet not found in report workbook - skipping")

    with log_step("Update Disbursement sheet"):
        if 'Disbursement' in wb.sheetnames:
            ws_disb = wb['Disbursement']
            ws_disb.delete_rows(5, ws_disb.max_row - 2)
            disbursement_df = write_dataframe_to_sheet(ws_disb, disbursement_df, start_row=5, start_col=1)
            log("Updated Disbursement sheet")
        else:
            logger.warning("'Disbursement' sheet not found in report workbook - skipping")

    with log_step("Update LTV Working - Disbursement sheet"):
        if 'LTV Working - Disbursement' in wb.sheetnames:
            ws_ltv = wb['LTV Working - Disbursement']
            ws_ltv.delete_rows(5, ws_ltv.max_row - 2)
            # Row 2 names the quarter period this sheet covers (e.g. "Jul 25 to Sep 25").
            ws_ltv.cell(row=2, column=1).value = (
                f"{prev2_month.strftime('%b %y')} to {current_month_date.strftime('%b %y')}"
            )
            write_dataframe_to_sheet(ws_ltv, df_ltv_working_disbursement, start_row=5, start_col=1)
            # Row 3's totals drive the WALTV % formula ($E$3) - keep their
            # SUM range in sync with however many data rows were just written.
            last_data_row = max(5, 4 + len(df_ltv_working_disbursement))
            ws_ltv['E3'] = f'=SUM(E5:E{last_data_row})'
            ws_ltv['J3'] = f'=SUM(J5:J{last_data_row})'
            # LTV %/WALTV % (columns I/J) are stored as raw decimals - without
            # an explicit percentage format, Excel renders small WALTV %
            # values in scientific notation (e.g. 1.07647E-06) instead of %.
            for row in range(5, last_data_row + 1):
                ws_ltv.cell(row=row, column=9).number_format = '0.00%'
                ws_ltv.cell(row=row, column=10).number_format = '0.00%'
            log("Updated LTV Working - Disbursement sheet")
        else:
            logger.warning("'LTV Working - Disbursement' sheet not found in report workbook - skipping")

    with log_step("Update Unutilized sheet"):
        if 'Unutilized' in wb.sheetnames:
            update_unutilized_sheet(wb['Unutilized'], file_path_unutilized_credit)
            log("Updated Unutilized sheet")
        else:
            logger.warning("'Unutilized' sheet not found in report workbook - skipping")

    with log_step("Update Property Mortgage List sheet"):
        if 'Property Mortgage List' in wb.sheetnames:
            ws_pm = wb['Property Mortgage List']
            ws_pm.cell(row=3, column=1).value = current_month_date
            ws_pm.delete_rows(6, ws_pm.max_row - 5)
            write_dataframe_to_sheet(ws_pm, df_property_mortgage, start_row=6, start_col=1)
            log("Updated Property Mortgage List sheet")
        else:
            logger.warning("'Property Mortgage List' sheet not found in report workbook - skipping")

    with log_step("Update MOI List (reschedule contracts)"):
        if 'MOI List -31 Jan' in wb.sheetnames:
            update_moi_reschedule_block(wb['MOI List -31 Jan'], df_reschedule)
            log("Updated MOI List -31 Jan sheet")
        else:
            logger.warning("'MOI List -31 Jan' sheet not found in report workbook - skipping")

    with log_step("Save report (first pass)"):
        current_q_monthname = (current_month_date.replace(day=1)).strftime("%b %Y")
        output_report = date_folder / f"Quarter Classification Loans & Advance-{current_q_monthname}.xlsx"

        log("Checking for old files matching pattern: Quarter Classification Loans & Advance*.xlsx")
        old_files = list(date_folder.glob("Quarter Classification Loans & Advance*.xlsx"))
        if old_files:
            for old_file in old_files:
                try:
                    old_file.unlink()
                    log(f"Removed old file: {old_file.name}")
                except Exception as e:
                    logger.warning(f"Could not remove old file {old_file.name}: {e}")
        else:
            log("No old files found to remove")

        log(f"Saving report to: {output_report}")
        save_workbook_safe(wb, output_report)
        wb.close()
        log(f"Report saved successfully: {output_report}")

    with log_step("Compute Top 50 clients"):
        updated_report = read_excel_safe(output_report, "just-saved report (Summary sheet)", sheet_name='Summary', skiprows=1)
        updated_report_top50 = updated_report[['Contract No', 'Client No', 'Customer Name']].copy()
        updated_report_top50['EXP+DP-IIS'] = updated_report['EXP+DP'] - updated_report['IIS']

        updated_MT_report_subset = marginal_loan[['ICAM code', '(ICAM) Debtor']].copy()
        updated_MT_report_subset.rename(columns={
            'ICAM code': 'Client No',
            '(ICAM) Debtor': 'EXP+DP-IIS',
        }, inplace=True)

        combined_top50_df = pd.concat([updated_report_top50, updated_MT_report_subset], ignore_index=True)
        pivoted_df = combined_top50_df.groupby('Client No', as_index=False)['EXP+DP-IIS'].sum()
        final_df = pivoted_df.sort_values(by='EXP+DP-IIS', ascending=False).reset_index(drop=True)
        log(f"Top 50 candidate pool: {len(final_df)} unique clients")

    with log_step("Update Top50 sheet"):
        if 'Top50' in wb.sheetnames:
            ws_top50 = wb['Top50']
            max_row = ws_top50.max_row
            for row in range(1, max_row + 1):
                ws_top50.cell(row=row, column=1).value = None
                ws_top50.cell(row=row, column=2).value = None

            final_df_clean = final_df.replace([np.inf, -np.inf], np.nan).fillna('')
            for j, col_name in enumerate(final_df_clean.columns, start=1):
                ws_top50.cell(row=1, column=j, value=col_name)
            for i, row in enumerate(final_df_clean.values, start=2):
                for j, val in enumerate(row, start=1):
                    cell = ws_top50.cell(row=i, column=j)
                    if isinstance(val, str) and val.startswith('='):
                        cell.value = val
                    else:
                        cell.value = val.item() if isinstance(val, (np.integer, np.floating)) else val
            log("Updated Top50 sheet with final_df in columns A and B")
        else:
            logger.warning("'Top50' sheet not found in report workbook - skipping")

        save_workbook_safe(wb, output_report)

    try:
        final_df_top50 = final_df.head(50).copy()
        final_df_top50['Top 50 Clients'] = "TOP 50"
        final_df_top50.to_excel('top50.xlsx', index=False)
    except Exception as e:
        logger.warning(f"Could not write auxiliary debug export 'top50.xlsx': {e}")

    contract_toTop50 = dict(zip(final_df_top50["Client No"], final_df_top50['Top 50 Clients']))

    temp = pd.DataFrame()
    temp['Contract No'] = np.nan
    temp['Client No'] = marginal_loan['ICAM code']
    temp['Customer Name'] = marginal_loan['Name of the Client']
    temp['EXP+DP-IIS'] = marginal_loan['(ICAM) Debtor']

    updated_report_top50 = pd.concat([updated_report_top50, temp], ignore_index=True)
    updated_report_top50['Top 50 Clients'] = updated_report_top50['Client No'].map(contract_toTop50)

    try:
        updated_report_top50.to_excel('finaltop50.xlsx', index=False)
    except Exception as e:
        logger.warning(f"Could not write auxiliary debug export 'finaltop50.xlsx': {e}")

    Filtered_T50 = updated_report_top50[updated_report_top50['Top 50 Clients'] == "TOP 50"].copy()
    log(f"Filtered Top 50 rows: {len(Filtered_T50)}")

    with log_step("Update C6 Working sheet"):
        contract_to_id = dict(zip(df_district["CLM_CODE"], df_district["CLM_IDNO"]))
        Filtered_T50['NIC/Company Registration No.'] = Filtered_T50['Client No'].map(contract_to_id)

        contract_to_prod_type = dict(zip(df_monthly_report["Contract No"], df_monthly_report["Product Type"]))
        Filtered_T50['Type of Facility'] = Filtered_T50['Contract No'].map(contract_to_prod_type)

        datasett50 = Filtered_T50[['Contract No', 'Client No', 'Customer Name', 'NIC/Company Registration No.', 'Type of Facility']]

        if 'C6 Working' in wb.sheetnames:
            ws_c6 = wb['C6 Working']

            for row in range(1, ws_c6.max_row + 1):
                for col in range(1, 6):
                    ws_c6.cell(row=row, column=col).value = None

            filtered_clean = datasett50.replace([np.inf, -np.inf], np.nan).fillna('')
            for j, col_name in enumerate(filtered_clean.columns[:5], start=1):
                ws_c6.cell(row=1, column=j).value = col_name
            for i, row in enumerate(filtered_clean.values, start=2):
                for j, val in enumerate(row[:5], start=1):
                    cell = ws_c6.cell(row=i, column=j)
                    if isinstance(val, str) and val.startswith('='):
                        cell.value = val
                    else:
                        cell.value = val.item() if isinstance(val, (np.integer, np.floating)) else val

            log("Updated C6 Working sheet with Filtered_T50")

            rows_updated = 0
            for row in range(2, ws_c6.max_row + 1):
                col_b_value = ws_c6.cell(row=row, column=2).value
                col_e_value = ws_c6.cell(row=row, column=5).value

                if col_b_value is not None and col_b_value != '':
                    if col_e_value is None or col_e_value == '':
                        ws_c6.cell(row=row, column=5).value = "Margin Trading Loans"
                        rows_updated += 1
                else:
                    break

            log(f"Added 'Margin Trading Loans' to column E for {rows_updated} rows")
        else:
            logger.warning("'C6 Working' sheet not found in report workbook - skipping")

    with log_step("Copy C6 Working to NBD-QF-23-C6 sorted"):
        if 'NBD-QF-23-C6 sorted' in wb.sheetnames:
            # Close and reopen with data_only=True so formulas resolve to their
            # calculated values before copying, then reopen writable again.
            save_workbook_safe(wb, output_report)
            wb.close()

            try:
                wb_data_only = load_workbook(output_report, data_only=True)
            except Exception as e:
                raise RuntimeError(f"Failed to reopen '{output_report}' in data-only mode: {e}") from e

            try:
                ws_c6_data_only = wb_data_only['C6 Working']

                last_data_row = 1
                for row in range(2, ws_c6_data_only.max_row + 1):
                    if ws_c6_data_only.cell(row=row, column=2).value is not None and ws_c6_data_only.cell(row=row, column=2).value != '':
                        last_data_row = row
                    else:
                        break

                wb = load_workbook(output_report)
                ws_c6_sorted = wb['NBD-QF-23-C6 sorted']

                if ws_c6_sorted.max_row >= 6:
                    for row in range(6, ws_c6_sorted.max_row + 1):
                        for col in range(1, 4):
                            ws_c6_sorted.cell(row=row, column=col).value = None

                column_mapping = {1: 1, 3: 2, 4: 3}  # C6 Working A/C/D -> sorted A/B/C

                for i, source_row in enumerate(range(2, last_data_row + 1), start=6):
                    for source_col, dest_col in column_mapping.items():
                        value = ws_c6_data_only.cell(row=source_row, column=source_col).value
                        if isinstance(value, (np.integer, np.floating)):
                            value = value.item()
                        ws_c6_sorted.cell(row=i, column=dest_col).value = value

                log(f"Copied data from C6 Working to NBD-QF-23-C6 sorted ({last_data_row - 1} rows)")
            finally:
                wb_data_only.close()
        else:
            logger.warning("'NBD-QF-23-C6 sorted' sheet not found in workbook - skipping")

    with log_step("Final save"):
        save_workbook_safe(wb, output_report)
        wb.close()

    with log_step("Backfill Disbursement sheet gaps from ERP"):
        # Best-effort: this looks up blank Customer Name/Gender/Sector/
        # Sub-Sector cells in the ERP and writes them back into the already-
        # saved report. Run as a subprocess (needs Chrome/Selenium) so a
        # failure here - ERP down, Chrome missing, network issue - can't take
        # down the rest of report generation; caught and logged instead.
        try:
            gap_fill_script = Path(__file__).resolve().parent.parent / "bots" / "disbursement_gap_fill_bot.py"
            gap_fill_result = subprocess.run(
                [sys.executable, str(gap_fill_script), "--report-file", str(output_report)],
                capture_output=True,
                text=True,
                cwd=str(gap_fill_script.parent)
            )
            if gap_fill_result.stdout:
                log(f"disbursement_gap_fill_bot.py output:\n{gap_fill_result.stdout}")
            if gap_fill_result.stderr:
                log(f"disbursement_gap_fill_bot.py log:\n{gap_fill_result.stderr}")
            if gap_fill_result.returncode == 0:
                log("Disbursement gap-fill completed")
            else:
                logger.warning(f"Disbursement gap-fill failed (rc={gap_fill_result.returncode}); "
                                f"continuing with report as-is")
        except Exception as e:
            logger.warning(f"Disbursement gap-fill step failed: {e}; continuing with report as-is")

    with log_step("Finalize MT workbook and copy it to MT-Disbursement"):
        # Runs after the gap-fill so the genders it found (report Disbursement
        # sheet) can be reused for MT clients before searching the ERP; also
        # fetches the GL 1846 balances. Best-effort, like the steps around it.
        if mt_workbook is None:
            logger.warning("No MT workbook was produced this run - 'MT-Disbursement' left unchanged")
        else:
            try:
                wb = load_workbook(output_report)
                finalize_mt_workbook(mt_workbook, wb, report_date_obj, date_folder)
                save_workbook_safe(wb, output_report)
                wb.close()
                log(f"MT-Disbursement updated from {mt_workbook}")
            except Exception as e:
                logger.warning(f"MT workbook finalization failed ({e}); 'MT-Disbursement' may be out of date")

    with log_step("Fill LTV Working - Disbursement Initial Valuation (#N/A) from ERP"):
        # Best-effort, same subprocess pattern as the gap-fill above. Runs after
        # it so both bots never open the report at once. Contracts with no
        # valuation in the ERP are written to Missing_Records.xlsx.
        try:
            ltv_script = Path(__file__).resolve().parent.parent / "bots" / "ltv_initial_valuation_bot.py"
            ltv_result = subprocess.run(
                [sys.executable, str(ltv_script), "--report-file", str(output_report)],
                capture_output=True,
                text=True,
                cwd=str(ltv_script.parent)
            )
            if ltv_result.stdout:
                log(f"ltv_initial_valuation_bot.py output:\n{ltv_result.stdout}")
            if ltv_result.stderr:
                log(f"ltv_initial_valuation_bot.py log:\n{ltv_result.stderr}")
            if ltv_result.returncode == 0:
                log("LTV Initial Valuation fill completed")
            else:
                logger.warning(f"LTV Initial Valuation fill failed (rc={ltv_result.returncode}); "
                                f"continuing with report as-is")
        except Exception as e:
            logger.warning(f"LTV Initial Valuation fill step failed: {e}; continuing with report as-is")

    log(f"Report successfully saved to: {output_report}")


def main():
    global logger
    try:
        _run()
    except KeyboardInterrupt:
        logger.warning("Interrupted by user")
        return 130
    except Exception as e:
        try:
            logger.exception("Process summary failed")
        except Exception:
            pass
        print(f"FATAL: process_summary.py failed: {e}", file=sys.stderr)
        traceback.print_exc()
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
