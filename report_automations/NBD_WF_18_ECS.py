import sys
import os
import re
import logging
import subprocess
import time
import shutil
import glob as glob_module
from datetime import datetime, timedelta, date
from calendar import monthrange
from functools import wraps
from contextlib import contextmanager
from pathlib import Path
from copy import copy

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import column_index_from_string, get_column_letter
import pandas as pd
import xlwings as xw

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def get_sheet(wb, name: str):
    """Get an openpyxl sheet by name, tolerating trailing/leading spaces."""
    if name in wb.sheetnames:
        return wb[name]
    # Try with/without trailing space
    for sname in wb.sheetnames:
        if sname.strip() == name.strip():
            return wb[sname]
    raise KeyError(f"Sheet '{name}' not found. Available: {wb.sheetnames}")


def com_retry(func, *args, retries=6, delay=2.0, **kwargs):
    """
    Call func(*args, **kwargs) retrying on COM RPC errors (Excel busy), and on
    AttributeError caused by a COM property transiently returning None instead
    of the live object (e.g. sheet.api / range.api occasionally returning None
    right after a workbook opens). Both are symptoms of the same underlying
    Excel-COM-not-ready-yet flakiness — func is re-evaluated fresh each retry,
    so a transient None resolves itself once Excel catches up.
    -2147418111 = RPC_E_CALL_REJECTED — Excel is busy; retry after a short wait.
    """
    import pywintypes
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            if e.args[0] == -2147418111:  # RPC_E_CALL_REJECTED
                logging.warning(f"COM call rejected (attempt {attempt}/{retries}), retrying in {delay}s...")
                time.sleep(delay)
                last_exc = e
            else:
                raise
        except AttributeError as e:
            logging.warning(f"COM property returned None (attempt {attempt}/{retries}): {e}, retrying in {delay}s...")
            time.sleep(delay)
            last_exc = e
    raise last_exc


# ---------------------------------------------------------
# Constants
# ---------------------------------------------------------
LARGE_VALUE_THRESHOLD = 50_000_000

# ---------------------------------------------------------
# TB Bot path configuration – update these placeholders
# ---------------------------------------------------------
TB_DOWNLOADS_FOLDER = Path(r"C:\Users\overdime\Downloads")   # Chrome download destination
_TB_ECS_ROOT        = Path(r"C:\CBSL\Script\outputs\NBD_WF_18_ECS")
def _resolve_tb_output_folder():
    import re as _re
    lines = [l.strip() for l in (_TB_ECS_ROOT / "completed_dates.txt").read_text().splitlines() if l.strip()]
    last_date = lines[-1]
    candidates = []
    for _f in _TB_ECS_ROOT.iterdir():
        if _f.is_dir() and (_f.name == last_date or _f.name.startswith(f"{last_date}(")):
            _m = _re.search(r'\((\d+)\)$', _f.name)
            candidates.append((int(_m.group(1)) if _m else 0, _f))
    if not candidates:
        folder = _TB_ECS_ROOT / last_date
        folder.mkdir(parents=True, exist_ok=True)
        return folder
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]
TB_OUTPUT_FOLDER = _resolve_tb_output_folder()
# ---------------------------------------------------------
# Timing utilities
# ---------------------------------------------------------
@contextmanager
def timed_operation(operation_name):
    """Context manager for timing operations."""
    start_time = time.perf_counter()
    logging.info(f"[START] {operation_name}")
    print(f"[START] {operation_name}")
    try:
        yield
    finally:
        elapsed = time.perf_counter() - start_time
        logging.info(f"[END] {operation_name} - Elapsed: {elapsed:.2f}s")
        print(f"[END] {operation_name} - Elapsed: {elapsed:.2f}s")

def log_time(func):
    """Decorator for timing functions."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        func_name = func.__name__
        logging.info(f"[START] {func_name}")
        print(f"\n[START] {func_name}")
        try:
            result = func(*args, **kwargs)
            return result
        finally:
            elapsed = time.perf_counter() - start_time
            logging.info(f"[END] {func_name} - Elapsed: {elapsed:.2f}s")
            print(f"[END] {func_name} - Elapsed: {elapsed:.2f}s")
    return wrapper

# ---------------------------------------------------------
# Logger setup: creates timestamped log file under /logs
# ---------------------------------------------------------
def setup_logger(script_dir):
    logs_dir = os.path.join(script_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)

    log_file = os.path.join(
        logs_dir,
        f"NBD_WF_18_ECS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logging.info("Logger initialized")
    print(f"Logger initialized. Log file: {log_file}")
    return log_file

# ---------------------------------------------------------
# File discovery helpers
# ---------------------------------------------------------
def find_file_by_pattern(working_path, starts_with=None, contains=None, extension=".xlsx"):
    """Generic file finder that searches for files matching given patterns."""
    for f in os.listdir(working_path):
        if f.startswith("~$"):
            continue
        if not f.lower().endswith(extension):
            continue
        if starts_with and not f.startswith(starts_with):
            continue
        if contains and contains not in f:
            continue
        return os.path.join(working_path, f)
    return None

def find_report_file(working_path):
    return find_file_by_pattern(working_path, starts_with="Cash-Flow-Forecast", extension=".xlsm")

def find_daily_bank_file(working_path):
    return find_file_by_pattern(working_path, starts_with="Daily Bank Balances")

def find_cash_flow_file(working_path):
    return find_file_by_pattern(working_path, starts_with="Cash Flow")

def find_dm_deposit_liability_file(working_path):
    return find_file_by_pattern(working_path, contains="Deposit Liability")

def find_afl_borrowing_schedule_file(working_path):
    # Matches "AFIN Borrowing Schedule - Week <N>.xlsx" regardless of week number.
    return find_file_by_pattern(working_path, starts_with="AFIN Borrowing Schedule")

def find_audit_receipt_file(working_path):
    # Match the weekly file specifically — not the monthly one, which also
    # contains "Audit Receipt" and would otherwise be picked ambiguously.
    return find_file_by_pattern(working_path, starts_with="Audit Receipt - Internal (weekly)")

def find_gl_account_details_file(working_path):
    return find_file_by_pattern(working_path, contains="GL Account Details")

def find_payments_file(working_path):
    return find_file_by_pattern(working_path, starts_with="Payments -")

def find_types_of_receipt_file(working_path):
    return find_file_by_pattern(working_path, contains="Types_of_Receipt")

def find_sasianet_file(working_path):
    return find_file_by_pattern(working_path, starts_with="SasiaNet -")

def find_tax_payments_file(working_path):
    return find_file_by_pattern(working_path, starts_with="Tax Payments -")

def find_cashflow_movement_file(working_path):
    return find_file_by_pattern(working_path, starts_with="Cashflow Movement -")

# ---------------------------------------------------------
# Cached label finder using DataFrame
# ---------------------------------------------------------
class SheetCache:
    """Cache for sheet data to avoid repeated cell-by-cell reads."""

    def __init__(self):
        self._label_cache = {}

    def find_row_by_label(self, ws, label, col=2):
        """Locate row number by label in specified column with caching."""
        cache_key = (id(ws), label, col)
        if cache_key in self._label_cache:
            return self._label_cache[cache_key]

        for row in ws.iter_rows(min_col=col, max_col=col):
            if row[0].value and str(row[0].value).strip() == label:
                self._label_cache[cache_key] = row[0].row
                return row[0].row
        return None

    def clear(self):
        self._label_cache.clear()

# Global cache instance
sheet_cache = SheetCache()

def find_row_by_label(ws, label):
    """Locate row number by label in column B."""
    return sheet_cache.find_row_by_label(ws, label, col=2)

# ---------------------------------------------------------
# Bulk read worksheet to DataFrame
# ---------------------------------------------------------
def worksheet_to_dataframe(ws, min_row=1, max_row=None, min_col=1, max_col=None, header_row=None):
    """Read worksheet data into DataFrame efficiently."""
    max_row = max_row or ws.max_row
    max_col = max_col or ws.max_column

    data = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
        data.append(row)

    if header_row is not None and data:
        header_idx = header_row - min_row
        if 0 <= header_idx < len(data):
            headers = data[header_idx]
            df = pd.DataFrame(data[header_idx + 1:], columns=headers)
            return df

    return pd.DataFrame(data)

# ---------------------------------------------------------
# Bulk write DataFrame to worksheet
# ---------------------------------------------------------
def dataframe_to_worksheet(ws, df, start_row, start_col=1, include_header=False):
    """Write DataFrame to worksheet efficiently."""
    rows = list(dataframe_to_rows(df, index=False, header=include_header))
    for r_idx, row in enumerate(rows):
        for c_idx, value in enumerate(row):
            ws.cell(row=start_row + r_idx, column=start_col + c_idx, value=value)
    return len(rows)

# ---------------------------------------------------------
# Extract account blocks and prepare dataframe - OPTIMIZED
# ---------------------------------------------------------
@log_time
def process_accounts(ws, start_label, skip_copy: bool = False):
    """Extract account data using bulk operations.

    skip_copy=True suppresses the column E -> C 'previous balance' copy below —
    used when this looks like a rerun for a week already processed, so a second
    pass doesn't overwrite last week's snapshot in C with a duplicate of E."""
    start_row = find_row_by_label(ws, start_label)
    if not start_row:
        raise ValueError(f"Label '{start_label}' not found")

    # Find total row — check cols A, B, C; case-insensitive; stop at next section header
    total_row = None
    for r in range(start_row + 1, ws.max_row + 1):
        for col in (1, 2, 3):
            cell_val = ws.cell(r, col).value
            if cell_val and str(cell_val).strip().lower() == "total":
                total_row = r
                break
        if total_row:
            break
        # Stop if we've hit another section header (non-empty col B that isn't a GL code)
        b_val = ws.cell(r, 2).value
        if b_val and str(b_val).strip().endswith("Balances") and r > start_row + 2:
            break

    if not total_row:
        logging.warning(f"Total row not found after '{start_label}' — using last non-empty row as boundary")
        # Fallback: find last row with data in col E between start_row+1 and next section
        for r in range(start_row + 1, ws.max_row + 1):
            b_val = ws.cell(r, 2).value
            if b_val and str(b_val).strip().endswith("Balances") and r > start_row + 2:
                total_row = r
                break
        if not total_row:
            raise ValueError(f"Total row not found after '{start_label}'")

    data_start = start_row + 1
    data_end = total_row - 1

    # Bulk read using iter_rows with values_only for speed
    data = []
    for row in ws.iter_rows(min_row=data_start, max_row=data_end, min_col=1, max_col=5, values_only=True):
        data.append(row)

    df = pd.DataFrame(data, columns=["GL", "Account details", "Col3", "Col4", "Amounts"])

    # Copy column E to column C — skipped on a same-week rerun (see skip_copy)
    if skip_copy:
        logging.info(f"{start_label}: skip_copy=True — not overwriting column C")
    else:
        for i, row in enumerate(data):
            ws.cell(data_start + i, 3).value = row[4]  # Copy col 5 to col 3

    df = df[["GL", "Account details", "Amounts"]]
    logging.info(f"{start_label}: extracted {len(df)} rows")

    return df, data_start, data_end

# ---------------------------------------------------------
# Read Daily Bank Balances - OPTIMIZED with full DataFrame operations
# ---------------------------------------------------------
@log_time
def process_daily_bank_balances(daily_file, input_date):
    """Process daily bank balances using pandas for bulk operations."""
    wb = load_workbook(daily_file, data_only=True, read_only=True)
    ws = wb.active

    # Bulk read all data at once
    data = list(ws.values)
    wb.close()

    df = pd.DataFrame(data[2:], columns=data[1])
    df.columns = [str(c).strip() if c is not None else "" for c in df.columns]

    base_cols = ["Bank Name", "GL", "Account Number"]

    # Vectorized column normalization
    normalized_cols = []
    for c in df.columns:
        if c in base_cols:
            normalized_cols.append(c)
        else:
            try:
                parsed_date = pd.to_datetime(c, errors='coerce')
                if pd.notna(parsed_date):
                    normalized_cols.append(parsed_date.strftime("%Y-%m-%d"))
                else:
                    normalized_cols.append(c)
            except Exception:
                normalized_cols.append(c)

    df.columns = normalized_cols

    date_cols = [c for c in df.columns if c not in base_cols and len(c) == 10 and c.count('-') == 2]

    if not date_cols:
        raise ValueError("No valid date columns found in Daily Bank Balances file")

    logging.info(f"Found {len(date_cols)} date columns")
    print(f"Found {len(date_cols)} date columns in Daily Bank Balances")

    input_str = input_date.strftime("%Y-%m-%d")

    if input_str in date_cols:
        chosen = input_str
        logging.info(f"Exact date match found: {chosen}")
    else:
        valid_dates = [d for d in date_cols if d <= input_str]
        if valid_dates:
            chosen = max(valid_dates)
            logging.info(f"Using latest available date: {chosen}")
        else:
            chosen = min(date_cols)
            logging.warning(f"No dates before input date. Using earliest date: {chosen}")

    df_daily = df[base_cols + [chosen]].copy()
    df_daily.rename(columns={chosen: "latest_date"}, inplace=True)
    df_daily["latest_date"] = pd.to_numeric(df_daily["latest_date"], errors="coerce")

    total = df_daily["latest_date"].sum(skipna=True)
    logging.info(f"Daily bank balances total: {total}")
    print(f"Daily bank balances total: {total}")

    return df_daily, total

# ---------------------------------------------------------
# Map GL balances - OPTIMIZED with vectorized operations
# ---------------------------------------------------------
@log_time
def apply_vlookup(ws, df_target, df_daily, data_start):
    """Apply VLOOKUP logic using pandas merge for speed."""
    # Create lookup dictionary once
    lookup = df_daily.set_index("GL")["latest_date"].to_dict()

    # Vectorized mapping
    df_target["Amounts"] = df_target["GL"].map(lookup).fillna(0)

    # Bulk write - prepare all values first
    amounts = df_target["Amounts"].tolist()
    for i, val in enumerate(amounts):
        ws.cell(data_start + i, 5).value = val

    return df_target["Amounts"].sum()

# ---------------------------------------------------------
# Week calculation helpers
# ---------------------------------------------------------
def week_of_month(dt):
    return ((dt.day - 1) // 7) + 1

def week_date_range(year, month, wom):
    start_day = (wom - 1) * 7 + 1
    end_day = min(start_day + 6, monthrange(year, month)[1])
    return date(year, month, start_day), date(year, month, end_day)

# ---------------------------------------------------------
# Step 03 – Weekly Distributions - OPTIMIZED with single xlwings session
# ---------------------------------------------------------
@log_time
def process_weekly_distributions(report_file, payments_file, input_date):
    """Process weekly distributions with consolidated xlwings usage."""
    wom = week_of_month(input_date)
    woy = input_date.isocalendar()[1]

    logging.info(f"Processing Weekly Distributions (WeekOfMonth={wom}, WeekOfYear={woy})")
    print(f"Processing Weekly Distributions - Week {wom} of month, Week {woy} of year")

    # Source col in Payments and Receipts.xlsx - target col in Weekly Distributions
    col_mappings = {"B": "G", "C": "I", "D": "K", "E": "M"}
    target_map = {1: ("B", "B1"), 2: ("C", "C1"), 3: ("D", "D1"), 4: ("E", "E1")}

    # PART A: Read B1:B7, C1:C7, D1:D7, E1:E7 from Payments and Receipts.xlsx
    #         and write to G49:G55, I49:I55, K49:K55, M49:M55 in Weekly Distributions
    with timed_operation("Write Payments and Receipts -> Weekly Distributions G49:M55"):
        wb_pr = load_workbook(payments_file, data_only=True, read_only=True)
        ws_pr = wb_pr.active
        payments_values = {}
        for src_col in col_mappings:
            col_idx = ord(src_col) - ord("A") + 1
            vals = [ws_pr.cell(row=r, column=col_idx).value for r in range(1, 8)]
            payments_values[src_col] = vals
            logging.info(f"Read Payments and Receipts {src_col}1:{src_col}7: {vals}")
        wb_pr.close()

        wb_report_write = load_workbook(report_file, keep_vba=True)
        ws_weekly_write = get_sheet(wb_report_write, "Weekly Distributions")
        for src_col, tgt_col in col_mappings.items():
            for i, val in enumerate(payments_values[src_col]):
                ws_weekly_write[f"{tgt_col}{49 + i}"].value = val
                logging.info(f"Wrote {tgt_col}{49 + i} = {val}")
        wb_report_write.save(report_file)
        logging.info("Saved Weekly Distributions G49:M55 from Payments and Receipts")

    # Single xlwings session to read calculated values from Weekly Distributions
    app = xw.App(visible=False)
    try:
        with timed_operation("xlwings read Weekly Distributions G18:G22"):
            source_vals = []
            wb_report_xw = app.books.open(report_file)
            time.sleep(2)  # Allow Excel COM to fully initialise before accessing sheets
            try:
                ws_weekly_xw = wb_report_xw.sheets["Weekly Distributions "]
                if ws_weekly_xw is None:
                    raise KeyError("sheet returned None")
            except (KeyError, Exception):
                ws_weekly_xw = wb_report_xw.sheets[2]
                logging.info(f"Accessed Weekly Distributions by index 2 (name: {ws_weekly_xw.name!r})")

            source_vals = ws_weekly_xw.range("G18:G22").value
            if not isinstance(source_vals, list):
                source_vals = [source_vals]
            logging.info(f"Read Weekly Distributions G18:G22: {source_vals}")
            wb_report_xw.close()
    finally:
        app.quit()

    # PART A2: Week 1 only — write week numbers to B25, B26, B27
    # C2:L2 hold dates: C-D=week1, E-F=week2, G-H=week3, I-J=week4, K-L=week1(next)
    # B25 - week that contains the 20th; B26+B27 - week that contains the 15th
    if wom == 1:
        with timed_operation("Write week numbers to B25:B27 (Weekly Distributions)"):
            try:
                def _to_date(val):
                    """Coerce cell value to a date object."""
                    if val is None:
                        return None
                    if isinstance(val, datetime):
                        return val.date()
                    if isinstance(val, date):
                        return val
                    # xlwings may return a float (Excel serial date)
                    if isinstance(val, float):
                        import datetime as _dt
                        try:
                            return (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=int(val))).date()
                        except Exception:
                            return None
                    return None

                # Use xlwings to read C2:L2 so formula-computed dates are resolved
                _app = xw.App(visible=False)
                try:
                    _wb = _app.books.open(report_file)
                    time.sleep(2)
                    try:
                        _ws = _wb.sheets["Weekly Distributions "]
                    except Exception:
                        _ws = _wb.sheets[2]
                    raw_dates = _ws.range("C2:L2").value
                    _wb.close()
                finally:
                    _app.quit()

                # range returns a list of lists for multi-row; flatten if needed
                if raw_dates and isinstance(raw_dates[0], list):
                    raw_dates = raw_dates[0]

                dates = [_to_date(v) for v in (raw_dates or [])]
                # Pad to 10 in case fewer cells returned
                dates += [None] * (10 - len(dates))

                # week2=E-F - idx 2,3 | week3=G-H - idx 4,5 | week4=I-J - idx 6,7
                w3_start, w3_end = dates[4], dates[5]
                w4_start, w4_end = dates[6], dates[7]
                logging.info(f"C2:L2 dates (via xlwings): {dates}")

                def _in_range(d, start, end):
                    return start and end and start <= d <= end

                # Determine the month from the current week dates (C2/D2)
                ref_date = dates[0] or dates[1]
                month = ref_date.month if ref_date else input_date.month
                year  = ref_date.year  if ref_date else input_date.year

                d20 = date(year, month, 20)
                d15 = date(year, month, 15)

                b25 = None
                if _in_range(d20, w3_start, w3_end):
                    b25 = 3
                elif _in_range(d20, w4_start, w4_end):
                    b25 = 4

                b26_b27 = None
                if _in_range(d15, w3_start, w3_end):
                    b26_b27 = 3
                elif _in_range(d15, w4_start, w4_end):
                    b26_b27 = 4

                logging.info(f"20th={d20}, 15th={d15} | B25={b25}, B26/B27={b26_b27}")

                if b25 is not None or b26_b27 is not None:
                    wb_write = load_workbook(report_file, keep_vba=True)
                    ws_write = get_sheet(wb_write, "Weekly Distributions")
                    if b25 is not None:
                        ws_write["B25"] = b25
                        logging.info(f"Set B25 = {b25}")
                    if b26_b27 is not None:
                        ws_write["B26"] = b26_b27
                        ws_write["B27"] = b26_b27
                        logging.info(f"Set B26 = B27 = {b26_b27}")
                    wb_write.save(report_file)
                    logging.info("Saved week numbers to B25:B27")
                else:
                    logging.info("20th and 15th not in week3 or week4 range — B25:B27 not updated")
            except Exception as e:
                logging.warning(f"Could not write week numbers to B25:B27: {e}")

    # PART B: Write to Payments file
    if wom in target_map:
        with timed_operation("Write to Payments file"):
            col, header = target_map[wom]
            wb_pay_write = load_workbook(payments_file)
            ws_pay_write = wb_pay_write.active

            # Bulk write source values
            for i, val in enumerate(source_vals):
                ws_pay_write[f"{col}{2+i}"].value = val
                logging.info(f"Pasted Payments {col}{2+i}: {val}")

            ws_pay_write[header].value = f"Week {woy}"
            wb_pay_write.save(payments_file)
            logging.info("Saved Payments file")

    # PART C: DM Deposit Liability logic
    with timed_operation("Process DM Deposit Liability"):
        folder = os.path.dirname(report_file)
        dm_file = find_file_by_pattern(folder, starts_with="NBD-WF-18-DM Deposit Liability")

        if not dm_file:
            logging.warning("DM Deposit Liability file not found - skipping")
            return

        wb_dm = load_workbook(dm_file, data_only=True, read_only=True)
        ws_dm = wb_dm["Portfolio"]

        # Cache headers - use enumerate for read_only mode compatibility
        headers = {}
        for col_idx, cell in enumerate(ws_dm[4], start=1):
            if cell.value:
                headers[cell.value] = col_idx
        mat_col = headers.get("Maturity Date")
        pay_col = headers.get("Total Payable")

        if not mat_col or not pay_col:
            wb_dm.close()
            raise ValueError("Required columns not found in DM Portfolio sheet")

        prev_month_date = input_date.replace(day=1) - timedelta(days=1)
        prev_year = prev_month_date.year
        prev_month = prev_month_date.month

        start_day = (wom - 1) * 7 + 1
        end_day = min(start_day + 6, monthrange(prev_year, prev_month)[1])
        start_dt = date(prev_year, prev_month, start_day)
        end_dt = date(prev_year, prev_month, end_day)

        # Bulk read and filter using DataFrame
        data = []
        for row in ws_dm.iter_rows(min_row=5, values_only=True):
            if len(row) >= max(mat_col, pay_col):
                data.append((row[mat_col - 1], row[pay_col - 1]))

        wb_dm.close()

        df_dm = pd.DataFrame(data, columns=["Maturity Date", "Total Payable"])
        df_dm["Maturity Date"] = pd.to_datetime(df_dm["Maturity Date"], errors="coerce").dt.date
        df_dm["Total Payable"] = pd.to_numeric(df_dm["Total Payable"], errors="coerce")

        # Vectorized filter
        mask = (df_dm["Maturity Date"] >= start_dt) & (df_dm["Maturity Date"] <= end_dt)
        total_payable = df_dm.loc[mask, "Total Payable"].sum()

        target_cells = {1: "B7", 2: "C7", 3: "D7", 4: "E7"}

        if wom in target_cells:
            wb_pay_dm = load_workbook(payments_file)
            ws_pay_dm = wb_pay_dm.active
            ws_pay_dm[target_cells[wom]].value = total_payable
            wb_pay_dm.save(payments_file)
            logging.info(f"DM Week {wom}: Total Payable {total_payable} -> {target_cells[wom]}")

# ---------------------------------------------------------
# Step 04 – Write Manual Payments to Weekly Distributions - from Master_Data.xlsx
# ---------------------------------------------------------
@log_time
def write_manual_payments_to_weekly_distributions(report_file):
    """Read manual payments from last row of Master_Data.xlsx (cols Q-Y)
    and write to Weekly Distributions sheet cells C51-C59. MT Loan Payment
    (col Z) is written by looking up the row in column B of Weekly
    Distributions whose label is 'MT Loan Payment' and writing to column C
    of that row."""
    logging.info("Processing Manual Payments -> Weekly Distributions")

    _col_to_cell = {
        "Q": "C51",  # Salary/Bonus/APIT current week
        "R": "C52",  # CBSL
        "S": "C53",  # EPF/ETF/OT
        "T": "C54",  # Outsource Staff
        "U": "C55",  # APIIT
        "V": "C56",  # FD
        "W": "C57",  # Taxes
        "X": "C58",  # Dividend Payment
        "Y": "C59",  # SSCL
    }
    _mt_loan_payment_col = "Z"
    _mt_loan_payment_label = "MT Loan Payment"

    try:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        master_data_path = os.path.join(root, "Master_Data.xlsx")
        wb_md = load_workbook(master_data_path, data_only=True)
        ws_md = wb_md["NBD_WF_18_ECS"]
        last_row = ws_md.max_row
        if last_row < 2:
            logging.warning("Master_Data.xlsx NBD_WF_18_ECS has no data rows — skipping manual payments")
            wb_md.close()
            return
        values = {}
        for col_letter, cell_ref in _col_to_cell.items():
            col_idx = ord(col_letter) - ord("A") + 1
            values[cell_ref] = ws_md.cell(row=last_row, column=col_idx).value
        mt_loan_payment_val = ws_md.cell(row=last_row, column=ord(_mt_loan_payment_col) - ord("A") + 1).value
        wb_md.close()
        logging.info(f"Read manual payments from Master_Data row {last_row}: {values}, MT Loan Payment={mt_loan_payment_val}")
    except Exception as e:
        logging.warning(f"Could not read manual payments from Master_Data.xlsx: {e}")
        return

    if all(v in (None, "") for v in values.values()) and mt_loan_payment_val in (None, ""):
        logging.warning("All manual payment values are empty — skipping Weekly Distributions update")
        return

    wb_report = load_workbook(report_file, keep_vba=True)
    ws_weekly = get_sheet(wb_report, "Weekly Distributions")
    if ws_weekly is None:
        logging.warning("'Weekly Distributions' sheet not found — skipping manual payments write")
        wb_report.close()
        return

    for cell_ref, val in values.items():
        ws_weekly[cell_ref] = val
        logging.info(f"Set {cell_ref} = {val}")

    if mt_loan_payment_val not in (None, ""):
        mt_row = None
        for (cell,) in ws_weekly.iter_rows(min_col=2, max_col=2):
            if str(cell.value).strip().lower() == _mt_loan_payment_label.lower():
                mt_row = cell.row
                break
        if mt_row:
            ws_weekly[f"C{mt_row}"] = mt_loan_payment_val
            logging.info(f"Set C{mt_row} = {mt_loan_payment_val} ('{_mt_loan_payment_label}' matched in column B, row {mt_row})")
        else:
            logging.warning(f"'{_mt_loan_payment_label}' not found in Weekly Distributions column B — skipping")

    wb_report.save(report_file)
    logging.info("Manual payments written to Weekly Distributions")

# ---------------------------------------------------------
# Step 04b – Write UCL Adjustments from Cashflows to NBD-WF-18-ECS - from Master_Data.xlsx
# ---------------------------------------------------------
def write_ucl_adjustments_from_cashflows(report_file):
    """Read the 4 'Adjustments from cashflows (+/-)' values (Next 0-7 / 8-14 / 15-21 /
    22-28 days) from the last row of Master_Data.xlsx (cols AB-AE) and write them to
    columns D-G of the row with code '18.9.1.2.0.0' — the 'Adjustments from cashflows
    (+/-)' row under 'Movement in unutilized Liquid Assets' — in the NBD-WF-18-ECS sheet.
    That code is looked up rather than the row label, since the same label
    ('Adjustments from cashflows (+/-)') also appears earlier under 'Movement in
    Liquid Assets' (code 18.7.1.2.0.0)."""
    logging.info("Processing UCL Adjustments from cashflows -> NBD-WF-18-ECS")

    _target_code = "18.9.1.2.0.0"
    _col_to_dest = {
        "AB": 4,  # D - Next 0-7 days
        "AC": 5,  # E - Next 8-14 days
        "AD": 6,  # F - Next 15-21 days
        "AE": 7,  # G - Next 22-28 days
    }

    try:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        master_data_path = os.path.join(root, "Master_Data.xlsx")
        wb_md = load_workbook(master_data_path, data_only=True)
        ws_md = wb_md["NBD_WF_18_ECS"]
        last_row = ws_md.max_row
        if last_row < 2:
            logging.warning("Master_Data.xlsx NBD_WF_18_ECS has no data rows — skipping UCL adjustments")
            wb_md.close()
            return
        values = {}
        for col_letter, dest_col in _col_to_dest.items():
            values[dest_col] = ws_md.cell(row=last_row, column=column_index_from_string(col_letter)).value
        wb_md.close()
        logging.info(f"Read UCL adjustment values from Master_Data row {last_row}: {values}")
    except Exception as e:
        logging.warning(f"Could not read UCL adjustment values from Master_Data.xlsx: {e}")
        return

    if all(v in (None, "") for v in values.values()):
        logging.warning("All UCL adjustment values are empty — skipping NBD-WF-18-ECS update")
        return

    wb_report = load_workbook(report_file, keep_vba=True)
    ws_ecs = get_sheet(wb_report, "NBD-WF-18-ECS")

    target_row = None
    for (cell,) in ws_ecs.iter_rows(min_col=1, max_col=1):
        if cell.value and str(cell.value).strip() == _target_code:
            target_row = cell.row
            break

    if not target_row:
        logging.warning(f"Code '{_target_code}' not found in NBD-WF-18-ECS column A — skipping UCL adjustments write")
        wb_report.close()
        return

    for dest_col, val in values.items():
        if val not in (None, ""):
            ws_ecs.cell(target_row, dest_col).value = val
            logging.info(f"Set NBD-WF-18-ECS row {target_row} col {dest_col} = {val}")

    wb_report.save(report_file)
    logging.info(f"UCL adjustments from cashflows written to NBD-WF-18-ECS row {target_row}, cols D-G")

# ---------------------------------------------------------
# Step 05 – Update Treasury Input Sheet - OPTIMIZED
# ---------------------------------------------------------
@log_time
def update_treasury_input_sheet(report_file, input_date):
    """Update Treasury Input with consolidated xlwings session.

    Also reads master data treasury values from Master_Data.xlsx (NBD_WF_18_ECS sheet,
    columns H-P of last row) and writes them to D4, D5, D8, D9, D13, D14, D18, D19, D20.
    """
    logging.info("Processing Treasury Input sheet update")

    target_fields = [
        "Securitization Loans - Capital",
        "Securitization Loans - Interest Payable",
        "Cost of Investment (Cost + Price Accrued)",
        "Interest Receivable (Coupon Accrued Interest)",
        "M2M Gain"
    ]

    # Read master data treasury values from Master_Data.xlsx
    treasury_master_values = {}
    try:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        master_data_path = os.path.join(root, "Master_Data.xlsx")
        wb_md = load_workbook(master_data_path, data_only=True)
        ws_md = wb_md["NBD_WF_18_ECS"]
        last_row = ws_md.max_row
        if last_row >= 2:
            col_map = {
                "H": "com_paper_cap",
                "I": "com_paper_int_pay",
                "J": "sec_loan_cap",
                "K": "sec_loan_int_pay",
                "L": "repo_cap",
                "M": "repo_int_rec",
                "N": "cost_of_inv",
                "O": "int_rec",
                "P": "m2m_gain",
            }
            for col_letter, key in col_map.items():
                col_idx = ord(col_letter) - ord("A") + 1
                val = ws_md.cell(row=last_row, column=col_idx).value
                treasury_master_values[key] = val
            logging.info(f"Read treasury master values from Master_Data.xlsx row {last_row}: {treasury_master_values}")
        wb_md.close()
    except Exception as e:
        logging.warning(f"Could not read treasury master values from Master_Data.xlsx: {e}")

    # Cell mapping: master data key -> Treasury Input sheet D-column row
    treasury_cell_map = {
        "com_paper_cap":    4,
        "com_paper_int_pay": 5,
        "sec_loan_cap":     8,
        "sec_loan_int_pay": 9,
        "repo_cap":         13,
        "repo_int_rec":     14,
        "cost_of_inv":      18,
        "int_rec":          19,
        "m2m_gain":         20,
    }

    # Single xlwings session
    app = xw.App(visible=False)
    field_values = {}

    try:
        with timed_operation("xlwings read Treasury Input"):
            wb_xw = app.books.open(report_file)
            time.sleep(2)  # Allow Excel COM to fully initialise before accessing sheets

            # Direct name lookup instead of iterating wb_xw.sheets — the iterator
            # (_NewEnum) has been the source of intermittent "does not support
            # enumeration" COM failures elsewhere in this file; Item()-style
            # named access avoids that code path entirely.
            try:
                ws_xw = com_retry(lambda: wb_xw.sheets["Treasury Input"])
            except KeyError:
                logging.warning("Treasury Input sheet not found")
                wb_xw.close()
                return None

            # Batch read column B to find target fields
            b_col_data = com_retry(lambda: ws_xw.range("B1:B100").value)
            for row_num, cell_val in enumerate(b_col_data, start=1):
                if cell_val and str(cell_val).strip() in target_fields:
                    field_name = str(cell_val).strip()
                    value_from_d = com_retry(lambda r=row_num: ws_xw.range(f"D{r}").value)
                    field_values[field_name] = {'row': row_num, 'value': value_from_d}
                    logging.info(f"Read D{row_num} ('{field_name}'): {value_from_d}")

            wb_xw.close()
    finally:
        app.quit()

    # Compute date strings for C1 and D1
    week_start = input_date - timedelta(days=input_date.weekday())  # Monday
    prev_day = week_start - timedelta(days=1)                        # Sunday before
    c1_date_str = f"{prev_day.day}-{prev_day.strftime('%b')}"
    d1_date_str = f"{input_date.day}-{input_date.strftime('%b')}"

    # Write using openpyxl
    with timed_operation("openpyxl write Treasury Input"):
        wb_report = load_workbook(report_file, keep_vba=True)
        ws_treasury = wb_report["Treasury Input"]

        # C1 = week_start - 1 day (DD-Mon), D1 = week ending date (DD-Mon)
        ws_treasury.cell(1, 3).value = c1_date_str
        ws_treasury.cell(1, 4).value = d1_date_str
        logging.info(f"Treasury Input C1={c1_date_str}, D1={d1_date_str}")

        # Copy existing D values to C for all treasury rows, then clear D
        for row_num in treasury_cell_map.values():
            d_val = ws_treasury.cell(row_num, 4).value
            ws_treasury.cell(row_num, 3).value = d_val
            ws_treasury.cell(row_num, 4).value = None
            logging.info(f"Moved D{row_num}->C{row_num}: {d_val}, cleared D{row_num}")

        # Also move D->C for any target_fields rows captured via xlwings
        for field_name, data in field_values.items():
            row_num = data['row']
            if row_num not in treasury_cell_map.values():
                ws_treasury.cell(row_num, 3).value = data['value']
                logging.info(f"'{field_name}': Moved D->C: {data['value']}")

        # Write new master data treasury values to D-column cells
        for key, row_num in treasury_cell_map.items():
            val = treasury_master_values.get(key)
            if val not in (None, ""):
                try:
                    val = float(val)
                except (TypeError, ValueError):
                    pass
                ws_treasury.cell(row_num, 4).value = val
                logging.info(f"Wrote master treasury '{key}' -> D{row_num}: {val}")

        wb_report.save(report_file)

    return input_date

# ---------------------------------------------------------
# Step 06 – Update TB Sheets - OPTIMIZED with DataFrame bulk copy
# ---------------------------------------------------------
@log_time
def update_tb_sheets(report_file):
    """Update TB sheets using DataFrame for fast bulk operations (values only, no styles)."""
    logging.info("Processing Opening TB and Close TB sheet updates")

    wb = load_workbook(report_file, keep_vba=True)

    opening_tb_sheet = None
    close_tb_sheet = None
    close_tb_date = None

    opening_tb_sheet = wb["Opening TB"] if "Opening TB" in wb.sheetnames else None
    close_tb_sheet   = wb["Close TB"]   if "Close TB"   in wb.sheetnames else None

    if not opening_tb_sheet or not close_tb_sheet:
        logging.warning("Opening TB or Close TB sheet not found - skipping")
        wb.close()
        return None

    with timed_operation("Read Close TB into DataFrame"):
        # Open file separately in read_only mode to get only actual data cells
        # (read_only mode streams from XML and doesn't inflate to 16384 columns)
        close_tb_name = close_tb_sheet.title
        wb_ro = load_workbook(report_file, read_only=True, data_only=True)
        ws_ro = wb_ro[close_tb_name]
        data = []
        for row in ws_ro.iter_rows(values_only=True):
            data.append(row)
        wb_ro.close()
        df = pd.DataFrame(data)
        logging.info(f"Close TB actual dimensions: {df.shape[0]} rows x {df.shape[1]} columns")
        logging.info(f"Read {len(df)} rows into DataFrame")

    with timed_operation("Clear Opening TB"):
        # Unmerge cells first
        for merged_range in list(opening_tb_sheet.merged_cells.ranges):
            opening_tb_sheet.unmerge_cells(str(merged_range))

        # Clear all cells - just delete rows and let them be recreated
        if opening_tb_sheet.max_row > 0:
            opening_tb_sheet.delete_rows(1, opening_tb_sheet.max_row)

    with timed_operation("Write DataFrame to Opening TB"):
        # Re-apply merged cells from source
        for merged_range in close_tb_sheet.merged_cells.ranges:
            opening_tb_sheet.merge_cells(str(merged_range))

        # Bulk write using DataFrame - much faster than cell by cell
        for r_idx, row in enumerate(df.values, start=1):
            for c_idx, value in enumerate(row, start=1):
                if value is not None:
                    opening_tb_sheet.cell(row=r_idx, column=c_idx, value=value)

    with timed_operation("Copy dimensions"):
        # Copy column widths
        for col_letter, col_dim in close_tb_sheet.column_dimensions.items():
            opening_tb_sheet.column_dimensions[col_letter].width = col_dim.width

        # Copy row heights
        for row_idx, row_dim in close_tb_sheet.row_dimensions.items():
            opening_tb_sheet.row_dimensions[row_idx].height = row_dim.height

    # new_opening_name = f"Opening TB-{close_tb_date}"
    # opening_tb_sheet.title = new_opening_name
    # logging.info(f"Renamed: Opening TB -> {new_opening_name}")

    wb.save(report_file)

    opening_tb_name = opening_tb_sheet.title

    with timed_operation("Copy entire Close TB to Opening TB via xlwings"):
        app = xw.App(visible=True)  # TEMP DIAGNOSTIC: visible so we can watch what Excel does live — revert to False after
        try:
            wb_xw = com_retry(lambda: app.books.open(str(report_file)))
            # Disable events/alerts immediately after opening — this workbook has VBA
            # (keep_vba=True) and may run a Workbook_Open macro or show an alert/msgbox
            # on open, which can block or corrupt COM automation calls on an invisible
            # instance since Excel is single-threaded for macros and COM alike.
            com_retry(lambda: setattr(app.api, "EnableEvents", False))
            com_retry(lambda: setattr(app, "display_alerts", False))
            # Look up sheets by direct name index instead of iterating wb_xw.sheets —
            # iterating the Sheets collection's enumerator is what was intermittently
            # throwing "This object does not support enumeration"; direct named/indexed
            # access uses a different COM call path and avoids it. Falls back to None
            # (handled below) on KeyError, same as a miss during iteration would.
            try:
                close_tb_xw = com_retry(lambda: wb_xw.sheets[close_tb_name])
            except KeyError:
                close_tb_xw = None
            try:
                opening_tb_xw = com_retry(lambda: wb_xw.sheets[opening_tb_name])
            except KeyError:
                opening_tb_xw = None
            if close_tb_xw is None or opening_tb_xw is None:
                logging.warning(
                    f"Sheet lookup failed in xlwings — "
                    f"wanted close='{close_tb_name}' opening='{opening_tb_name}'"
                )
            else:
                com_retry(lambda: opening_tb_xw.api.UsedRange.UnMerge())
                com_retry(lambda: close_tb_xw.api.UsedRange.Copy(opening_tb_xw.range("A1").api))
            com_retry(wb_xw.save)
            wb_xw.close()
        finally:
            app.quit()

    logging.info("TB sheet updates complete")

    return close_tb_date

# ---------------------------------------------------------
# Step 07 – Manual Process Reminder: Close TB Data Entry
# ---------------------------------------------------------
def log_manual_close_tb_reminder(close_tb_date):
    _log_manual_reminder(
        "Close TB Data Entry",
        f"Close TB sheet must be updated manually from ERP",
        [
            "INSTRUCTIONS FOR CLOSE TB DATA ENTRY:",
            "",
            "1. Open ERP system",
            "2. Navigate to: Finance --> Management Account --> TB-COA",
            "3. Select the required date",
            "4. Generate the report",
            "5. Copy the entire sheet from the report",
            "6. Open the Cash-Flow-Forecast file",
            f"7. Go to the 'Close TB-{close_tb_date}' sheet",
            "8. Paste the copied data into the Close TB sheet",
        ],
        "Please complete the manual Close TB data entry process"
    )

# ---------------------------------------------------------
# Step 08 – Update utilized_credit_lines - OPTIMIZED with DataFrame
# ---------------------------------------------------------
@log_time
def update_utilized_credit_lines(report_file, afl_file, input_date):
    """Update utilized credit lines sheet.

    1. Read C1-C5,C8 computed values via xlwings - write to B1-B5,B8.
    2. Set C1 = "Week {week_number}".
    3. Clear columns E-I of the UCL sheet.
    4. From AFL Funding Arrangement sheet, find header row containing "Institute",
       read columns: Institute, Institute (duplicate), Limits, Utilized,
       Type of facility, Unutilized balance as of * — values only (no formulas),
       skip last data row, paste to UCL E1 onwards.
    5. Find section label rows in column E:
         Short term, Over Draft Facility, Long term, Inter Company
    6. Write SUM/ref formulas to C2,C3,C4,C8 based on discovered I-column ranges.
    """
    logging.info("Processing utilized credit lines sheet update")

    week_number = input_date.isocalendar()[1]

    # ── Step 1: Read C1-C5,C8 computed values via xlwings ──────────────────────
    c_values = {}
    with timed_operation("Read C1:C5,C8 via xlwings"):
        _app = xw.App(visible=False)
        try:
            _wb = _app.books.open(report_file)
            time.sleep(1)
            _ws = _wb.sheets["utilized credit lines"]
            for r in [1, 2, 3, 4, 5, 8]:
                c_values[r] = _ws.range(f"C{r}").value
            _wb.close()
        finally:
            _app.quit()
    logging.info(f"Read UCL C values: {c_values}")

    wb_report = load_workbook(report_file, keep_vba=True)
    if "utilized credit lines" not in wb_report.sheetnames:
        logging.warning("'utilized credit lines' sheet not found — skipping")
        wb_report.close()
        return

    ws_ucl = wb_report["utilized credit lines"]

    # Write C values - B column; set week label
    for r in [1, 2, 3, 4, 5, 8]:
        ws_ucl.cell(r, 2).value = c_values.get(r)
    ws_ucl.cell(1, 3).value = f"Week {week_number}"
    logging.info("Written C-B values and week label")

    # ── Step 2: Clear columns E-I (cols 5-9) ───────────────────────────────────
    with timed_operation("Clear UCL columns E-I"):
        max_clear = ws_ucl.max_row
        for r in range(1, max_clear + 1):
            for c in range(5, 10):   # E=5 … I=9
                ws_ucl.cell(r, c).value = None
    logging.info("Cleared columns E-I")

    if not afl_file:
        logging.warning("AFL Borrowing Schedule file not found — skipping E-I paste")
        wb_report.save(report_file)
        return

    # ── Step 3: Read AFL Funding Arrangement via xlwings (computed values) ──────
    with timed_operation("Read AFL Funding Arrangement via xlwings"):
        _app2 = xw.App(visible=False)
        try:
            _wb2 = _app2.books.open(afl_file)
            time.sleep(1)

            # The AFIN Borrowing Schedule file always has exactly one sheet
            # ("Sheet1") — no name search needed, just grab it directly.
            _ws_afl = com_retry(lambda: _wb2.sheets[0])

            # Find header row containing "Institute"
            used = _ws_afl.used_range
            last_row = used.last_cell.row
            last_col = used.last_cell.column

            afl_header_row_idx = None
            afl_header_vals = None
            for _r in range(1, min(30, last_row) + 1):
                _row_vals = _ws_afl.range((_r, 1), (_r, last_col)).value
                if not isinstance(_row_vals, list):
                    _row_vals = [_row_vals]
                _strs = [str(v).strip() if v else "" for v in _row_vals]
                if "Institute" in _strs:
                    afl_header_row_idx = _r
                    afl_header_vals = _strs
                    break

            if afl_header_row_idx is None:
                logging.warning("Could not find 'Institute' header in AFL file — skipping paste")
                _wb2.close()
                wb_report.save(report_file)
                return

            logging.info(f"AFL header row: {afl_header_row_idx}, headers: {afl_header_vals}")

            # Identify required column indices (0-based within the row)
            # Target columns: Institute, Institute(dup), Limits, Utilized,
            #                  Type of facility, Unutilized balance as of *
            institute_indices = [i for i, h in enumerate(afl_header_vals) if h == "Institute"]
            limits_idx    = next((i for i, h in enumerate(afl_header_vals) if h == "Limits"), None)
            utilized_idx  = next((i for i, h in enumerate(afl_header_vals) if h == "Utilized"), None)
            type_idx      = next((i for i, h in enumerate(afl_header_vals) if h == "Type of facility"), None)
            unutilized_idx = next(
                (i for i, h in enumerate(afl_header_vals) if h.startswith("Unutilized balance")),
                None
            )

            # Build ordered list of (dest_col_offset 0-based, src_col_offset 0-based)
            # dest order: E=0, F=1, G=2, H=3, I=4  -  6 target cols but we have 6 source cols
            dest_src = []
            # Two Institute columns
            for inst_i in institute_indices[:2]:
                dest_src.append(inst_i)
            if limits_idx is not None:
                dest_src.append(limits_idx)
            if utilized_idx is not None:
                dest_src.append(utilized_idx)
            if type_idx is not None:
                dest_src.append(type_idx)
            if unutilized_idx is not None:
                dest_src.append(unutilized_idx)

            logging.info(f"Source col indices (0-based): {dest_src}")

            # Read all data rows (header row to last row) as computed values
            all_rows_raw = _ws_afl.range(
                (afl_header_row_idx, 1), (last_row, last_col)
            ).value
            # xlwings returns list-of-lists for multi-row ranges
            if last_row - afl_header_row_idx == 0:
                all_rows_raw = [all_rows_raw]

            _wb2.close()
        finally:
            _app2.quit()

    # ── Step 4: Determine rows to paste ────────────────────────────────────────
    # all_rows_raw[0] = header row, [1]+ = data rows
    data_rows_raw = all_rows_raw[1:]   # strip header

    # Find "Inter Company" row index within data_rows_raw (0-based)
    # last_row = ic_row_in_data + 2  -  paste up to ic_row_in_data + 1 (inclusive),
    # i.e. slice [:ic_row_in_data + 2]
    inst_src = dest_src[0] if dest_src else 0
    ic_data_idx = None   # 0-based index of "Inter Company" label row in data_rows_raw
    for _i, _row in enumerate(data_rows_raw):
        if not isinstance(_row, list):
            _row = [_row]
        _inst_val = _row[inst_src] if inst_src < len(_row) else None
        if _inst_val and str(_inst_val).strip().lower() == "inter company":
            ic_data_idx = _i
            break

    if ic_data_idx is None:
        # Fallback: skip last data row (original behaviour)
        logging.warning("'Inter Company' label not found in AFL data — falling back to skip-last-row logic")
        last_data_idx = None
        for _i in range(len(data_rows_raw) - 1, -1, -1):
            _row = data_rows_raw[_i]
            if not isinstance(_row, list):
                _row = [_row]
            if _row[inst_src] not in (None, ""):
                last_data_idx = _i
                break
        if last_data_idx is None:
            logging.warning("No data rows found in AFL Funding Arrangement — skipping paste")
            wb_report.save(report_file)
            return
        rows_to_paste = data_rows_raw[:last_data_idx]
    else:
        # Paste rows up to ic_row + 1 (ic_row itself + one data row after it)
        rows_to_paste = data_rows_raw[:ic_data_idx + 2]
        logging.info(f"Inter Company found at data index {ic_data_idx} — pasting {len(rows_to_paste)} rows")

    logging.info(f"Rows to paste to UCL (excl Total row): {len(rows_to_paste)}")

    # ── Step 5: Paste header + data rows to UCL E1 onwards ─────────────────────
    with timed_operation("Paste AFL data to UCL E-I"):
        # Row 1: write header labels
        header_row_data = all_rows_raw[0]
        if not isinstance(header_row_data, list):
            header_row_data = [header_row_data]

        for dest_offset, src_offset in enumerate(dest_src):
            if src_offset < len(header_row_data):
                ws_ucl.cell(1, 5 + dest_offset).value = header_row_data[src_offset]

        # Data rows starting from row 2
        for r_offset, row in enumerate(rows_to_paste):
            if not isinstance(row, list):
                row = [row]
            for dest_offset, src_offset in enumerate(dest_src):
                val = row[src_offset] if src_offset < len(row) else None
                ws_ucl.cell(2 + r_offset, 5 + dest_offset).value = val

        # Total row: immediately after the last pasted data row
        total_row_ucl = 2 + len(rows_to_paste)   # UCL sheet row number
        ws_ucl.cell(total_row_ucl, 5).value = "Total"                              # E
        ws_ucl.cell(total_row_ucl, 6).value = f"=SUM(F3:F{total_row_ucl - 1})"    # F
        ws_ucl.cell(total_row_ucl, 7).value = f"=SUM(G3:G{total_row_ucl - 1})"    # G
        ws_ucl.cell(total_row_ucl, 9).value = f"=SUM(I3:I{total_row_ucl - 1})"    # I
        logging.info(f"Added Total row at UCL row {total_row_ucl}")

    logging.info("Pasted AFL data to UCL columns E-I")

    # ── Step 6: Find section label rows in column E (col 5) ─────────────────────
    # Labels: "Short term", "Over Draft Facility", "Long term", "Inter Company",
    # plus the "Total STL" / "Total OD" / "Total LT" subtotal rows already
    # present in the pasted data (used directly for C2/C3/C8 instead of
    # summing a row range).
    # (search from row 2 since row 1 is headers)
    section_labels = {
        "short term": None,
        "over draft facility": None,
        "long term": None,
        "inter company": None,
        "total stl": None,
        "total od": None,
        "total lt": None,
    }
    ucl_last_row = ws_ucl.max_row
    for _r in range(2, ucl_last_row + 1):
        _val = ws_ucl.cell(_r, 5).value   # column E
        if _val is not None:
            _lower = str(_val).strip().lower()
            if _lower in section_labels and section_labels[_lower] is None:
                section_labels[_lower] = _r

    logging.info(f"Section label rows in col E: {section_labels}")

    ic_row  = section_labels.get("inter company")
    total_stl_row = section_labels.get("total stl")
    total_od_row  = section_labels.get("total od")
    total_lt_row  = section_labels.get("total lt")

    # ── Step 7: Write formulas to C2, C3, C4, C8 ────────────────────────────────
    # C2/C3/C8 reference the existing "Total STL"/"Total LT"/"Total OD" subtotal
    # rows directly, rather than re-summing the individual bank rows.
    # IC_range:  I(ic_row)

    if total_stl_row:
        ws_ucl["C2"] = f"=+I{total_stl_row}"
        logging.info(f"C2 = +I{total_stl_row}")
    else:
        logging.warning("'Total STL' row not found — C2 not updated")

    if total_lt_row:
        ws_ucl["C3"] = f"=+I{total_lt_row}"
        logging.info(f"C3 = +I{total_lt_row}")
    else:
        logging.warning("'Total LT' row not found — C3 not updated")

    if ic_row:
        ws_ucl["C4"] = f"=+I{ic_row}"
        logging.info(f"C4 = +I{ic_row}")
    else:
        logging.warning("Inter Company row not found — C4 not updated")

    if total_od_row:
        ws_ucl["C8"] = f"=+I{total_od_row}"
        logging.info(f"C8 = +I{total_od_row}")
    else:
        logging.warning("'Total OD' row not found — C8 not updated")

    wb_report.save(report_file)
    logging.info("utilized_credit_lines update complete")

# ---------------------------------------------------------
# Step 09 – Update FD Sheet - OPTIMIZED with DataFrame
# ---------------------------------------------------------
@log_time
def update_fd_sheet(report_file, dm_file):
    """Update FD sheet from DM Deposit Liability Portfolio sheet.

    Uses xlwings to read calculated 'Total Payable' values (formula results).
    Copies System Deposit No, Maturity Date, Total Payable - FD sheet cols A, B, C from row 3.
    """
    logging.info("Processing FD sheet update")

    if not dm_file:
        logging.warning("DM Deposit Liability file not found - skipping")
        return

    with timed_operation("Read DM Portfolio data via xlwings"):
        _app = xw.App(visible=False)
        try:
            _wb = _app.books.open(dm_file)
            time.sleep(1)
            _ws = _wb.sheets["Portfolio"]

            # Find header row (row 4) — read cols A onwards to locate required columns
            header_row = _ws.range("A4").expand("right").value
            if not isinstance(header_row, list):
                header_row = [header_row]

            required = ["System Deposit No:", "Maturity Date", "Total Payable"]
            col_indices = {}
            for i, h in enumerate(header_row):
                if h and str(h).strip() in required:
                    col_indices[str(h).strip()] = i  # 0-based offset

            missing = [h for h in required if h not in col_indices]
            if missing:
                logging.warning(f"Missing headers in Portfolio: {missing}")
                _wb.close()
                return

            # Read all data from row 5 downwards for the 3 required columns
            # Use used range to determine last row
            last_row = _ws.range("A4").expand("down").last_cell.row
            if last_row < 5:
                logging.warning("No data rows found in Portfolio sheet")
                _wb.close()
                return

            # Read each required column as a range (gets computed values)
            def _col_letter(zero_idx):
                # Convert 0-based column index to Excel column letter(s)
                import string
                result = ""
                n = zero_idx
                while True:
                    result = string.ascii_uppercase[n % 26] + result
                    n = n // 26 - 1
                    if n < 0:
                        break
                return result

            col_a_letter = _col_letter(col_indices["System Deposit No:"])
            col_b_letter = _col_letter(col_indices["Maturity Date"])
            col_c_letter = _col_letter(col_indices["Total Payable"])

            dep_vals    = _ws.range(f"{col_a_letter}5:{col_a_letter}{last_row}").value
            mat_vals    = _ws.range(f"{col_b_letter}5:{col_b_letter}{last_row}").value
            total_vals  = _ws.range(f"{col_c_letter}5:{col_c_letter}{last_row}").value

            _wb.close()
        finally:
            _app.quit()

    # Normalise to lists (single-cell ranges return scalar)
    def _to_list(v):
        if v is None:
            return []
        if not isinstance(v, list):
            return [v]
        return v

    dep_vals   = _to_list(dep_vals)
    mat_vals   = _to_list(mat_vals)
    total_vals = _to_list(total_vals)

    data = list(zip(dep_vals, mat_vals, total_vals))
    df = pd.DataFrame(data, columns=["System Deposit No:", "Maturity Date", "Total Payable"])
    df = df.dropna(how="all")
    logging.info(f"Read {len(df)} rows from Portfolio sheet (via xlwings)")

    # Sort by maturity date
    df["Maturity Date"] = pd.to_datetime(df["Maturity Date"], errors="coerce")
    df = df.sort_values("Maturity Date", na_position="last")
    df["Total Payable"] = pd.to_numeric(df["Total Payable"], errors="coerce")

    # Split: rows where integer part of Total Payable has more than 7 digits
    large_mask = df["Total Payable"].apply(
        lambda v: len(str(int(abs(v)))) > 7 if pd.notna(v) and v != 0 else False
    )
    large_df = df[large_mask].reset_index(drop=True)
    df = df[~large_mask].reset_index(drop=True)
    logging.info(f"Main df: {len(df)} rows, Large df (> 7 digits): {len(large_df)} rows")

    def _write_date_cell(ws, row_idx, col_idx, val):
        """Write a date value with DD-MM-YYYY number format."""
        if pd.isna(val):
            val = None
        elif isinstance(val, pd.Timestamp):
            val = val.to_pydatetime()
        elif isinstance(val, date) and not isinstance(val, datetime):
            val = datetime(val.year, val.month, val.day)
        cell = ws.cell(row_idx, col_idx)
        cell.value = val
        if isinstance(val, datetime):
            cell.number_format = "DD-MM-YYYY"

    with timed_operation("Write to FD sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        if "FD" not in wb_report.sheetnames:
            logging.warning("FD sheet not found - skipping")
            wb_report.close()
            return

        ws_fd = wb_report["FD"]

        # Clear A:C from row 3 to last data row
        if ws_fd.max_row >= 3:
            ws_fd.delete_rows(3, ws_fd.max_row - 2)

        # Paste main df - A, B, C from row 3
        for idx, row in enumerate(df.itertuples(index=False), start=3):
            ws_fd.cell(idx, 1).value = row[0]        # System Deposit No
            _write_date_cell(ws_fd, idx, 2, row[1])  # Maturity Date
            ws_fd.cell(idx, 3).value = row[2]        # Total Payable

        # Clear I:K from row 3 to last data row (before writing large_df)
        max_row_after = ws_fd.max_row
        if max_row_after >= 3:
            for r in ws_fd.iter_rows(min_row=3, max_row=max_row_after, min_col=9, max_col=11):
                for cell in r:
                    cell.value = None

        # Paste large df - I, J, K from row 3
        for idx, row in enumerate(large_df.itertuples(index=False), start=3):
            ws_fd.cell(idx, 9).value = row[0]         # System Deposit No
            _write_date_cell(ws_fd, idx, 10, row[1])  # Maturity Date
            ws_fd.cell(idx, 11).value = row[2]        # Total Payable

        wb_report.save(report_file)
        logging.info(f"FD sheet updated: {len(df)} rows - A:C, {len(large_df)} rows - I:K from row 3")

# ---------------------------------------------------------
# Step 09b – Update Types of Receipt Sheet from uploaded file
# ---------------------------------------------------------
def update_types_of_receipt_sheet(report_file, types_of_receipt_file):
    """Clear 'Types of Receipt' sheet from row 2 and paste data from Types_of_Receipt.xlsx row 2 onwards."""
    if not types_of_receipt_file:
        logging.info("No Types_of_Receipt.xlsx found — skipping Types of Receipt sheet update")
        return

    logging.info(f"Updating Types of Receipt sheet from: {os.path.basename(types_of_receipt_file)}")

    with timed_operation("Read Types of Receipt source data"):
        wb_src = load_workbook(types_of_receipt_file, data_only=True, read_only=True)
        ws_src = wb_src.active
        src_rows = list(ws_src.iter_rows(min_row=2, values_only=True))
        wb_src.close()
    logging.info(f"Read {len(src_rows)} rows from Types_of_Receipt.xlsx")

    with timed_operation("Write to Types of Receipt sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        if "Types of Receipt" not in wb_report.sheetnames:
            logging.warning("'Types of Receipt' sheet not found in report — skipping")
            wb_report.close()
            return

        ws_tor = wb_report["Types of Receipt"]

        # Clear from row 2 to end
        for row in ws_tor.iter_rows(min_row=2, max_row=ws_tor.max_row):
            for cell in row:
                cell.value = None

        # Paste from row 2
        for r_idx, row_data in enumerate(src_rows, start=2):
            for c_idx, val in enumerate(row_data, start=1):
                ws_tor.cell(r_idx, c_idx).value = val

        wb_report.save(report_file)
        logging.info(f"Types of Receipt sheet updated: {len(src_rows)} rows")

# ---------------------------------------------------------
# Step 09c – Update SasiaNet Sheet from uploaded file
# ---------------------------------------------------------
def update_sasianet_sheet(report_file, sasianet_file):
    """Clear SasiaNet sheet A:L from row 2 and paste data from SasiaNet - *.xlsx row 2 onwards."""
    if not sasianet_file:
        logging.warning("No SasiaNet file found — skipping SasiaNet sheet update")
        return

    logging.info(f"Updating SasiaNet sheet from: {os.path.basename(sasianet_file)}")

    with timed_operation("Read SasiaNet source data"):
        wb_src = load_workbook(sasianet_file, data_only=True, read_only=True)
        ws_src = wb_src.active
        src_rows = list(ws_src.iter_rows(min_row=2, max_col=12, values_only=True))
        wb_src.close()
    logging.info(f"Read {len(src_rows)} rows from SasiaNet file")

    with timed_operation("Write to SasiaNet sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        if "SasiaNet" not in wb_report.sheetnames:
            logging.warning("'SasiaNet' sheet not found in report — skipping")
            wb_report.close()
            return

        ws_sn = wb_report["SasiaNet"]

        # Clear A:L from row 2
        for row in ws_sn.iter_rows(min_row=2, max_row=ws_sn.max_row, min_col=1, max_col=12):
            for cell in row:
                cell.value = None

        # Paste from row 2
        for r_idx, row_data in enumerate(src_rows, start=2):
            for c_idx, val in enumerate(row_data, start=1):
                ws_sn.cell(r_idx, c_idx).value = val

        wb_report.save(report_file)
        logging.info(f"SasiaNet sheet updated: {len(src_rows)} rows")

def update_tax_payments_sheet(report_file, tax_payments_file):
    """Copy A1:D15 from Tax Payments file and replace Tax Payments sheet A1:D15."""
    if not tax_payments_file:
        logging.info("No Tax Payments file found — keeping existing Tax Payments sheet data")
        return

    logging.info(f"Updating Tax Payments sheet from: {os.path.basename(tax_payments_file)}")

    with timed_operation("Read Tax Payments source data"):
        wb_src = load_workbook(tax_payments_file, data_only=True, read_only=True)
        ws_src = wb_src.active
        src_rows = list(ws_src.iter_rows(min_row=1, max_row=15, min_col=1, max_col=4, values_only=True))
        wb_src.close()
    logging.info(f"Read {len(src_rows)} rows from Tax Payments file")

    with timed_operation("Write to Tax Payments sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        # Try both sheet name variants
        sheet_name = None
        for candidate in ("TAX Payments", "Tax Payments"):
            if candidate in wb_report.sheetnames:
                sheet_name = candidate
                break
        if sheet_name is None:
            logging.warning("'TAX Payments' / 'Tax Payments' sheet not found in report — skipping")
            wb_report.close()
            return

        ws_tp = wb_report[sheet_name]

        # Paste A1:D15
        for r_idx, row_data in enumerate(src_rows, start=1):
            for c_idx, val in enumerate(row_data, start=1):
                ws_tp.cell(r_idx, c_idx).value = val

        wb_report.save(report_file)
        logging.info(f"Tax Payments sheet updated: {len(src_rows)} rows")

def update_cashflow_movement_sheet(report_file, cashflow_movement_file):
    """Wipe col A and B from row 2 in 'Cashflow Movement' sheet, then paste A and B from row 2 of source file."""
    if not cashflow_movement_file:
        logging.info("No Cashflow Movement file found — skipping Cashflow Movement sheet update")
        return

    logging.info(f"Updating Cashflow Movement sheet from: {os.path.basename(cashflow_movement_file)}")

    with timed_operation("Read Cashflow Movement source data"):
        wb_src = load_workbook(cashflow_movement_file, data_only=True, read_only=True)
        ws_src = wb_src.active
        src_rows = list(ws_src.iter_rows(min_row=2, max_col=2, values_only=True))
        wb_src.close()
    logging.info(f"Read {len(src_rows)} rows from Cashflow Movement file")

    with timed_operation("Write to Cashflow Movement sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        if "Cashflow Movement" not in wb_report.sheetnames:
            logging.warning("'Cashflow Movement' sheet not found in report — skipping")
            wb_report.close()
            return

        ws_cm = wb_report["Cashflow Movement"]

        # Wipe A and B from row 2 to end
        for row in ws_cm.iter_rows(min_row=2, max_row=ws_cm.max_row, min_col=1, max_col=2):
            for cell in row:
                cell.value = None

        # Paste A and B from row 2
        for r_idx, row_data in enumerate(src_rows, start=2):
            for c_idx, val in enumerate(row_data, start=1):
                ws_cm.cell(r_idx, c_idx).value = val

        wb_report.save(report_file)
        logging.info(f"Cashflow Movement sheet updated: {len(src_rows)} rows")

# ---------------------------------------------------------
# Step 10 – Update Receipt Sheet - OPTIMIZED with DataFrame
# ---------------------------------------------------------
@log_time
def update_receipt_sheet(report_file, audit_receipt_file):
    """Update Receipt sheet using DataFrame bulk operations."""
    logging.info("Processing Receipt sheet update")

    if not audit_receipt_file:
        logging.warning("Audit Receipt file not found - skipping")
        return

    with timed_operation("Read Audit Receipt data"):
        wb_audit = load_workbook(audit_receipt_file, data_only=True, read_only=True)
        ws_audit = wb_audit.active

        # Bulk read all data
        audit_data = list(ws_audit.iter_rows(values_only=True))
        wb_audit.close()
        logging.info(f"Read {len(audit_data)} rows from Audit Receipt")

    with timed_operation("Write to Receipt sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        if "Receipt" not in wb_report.sheetnames:
            logging.warning("Receipt sheet not found - skipping")
            wb_report.close()
            return

        ws_receipt = wb_report["Receipt"]

        # Read W1 as formula template (col 23, row 1)
        w1_formula = ws_receipt.cell(1, 23).value
        logging.info(f"W1 formula template: {w1_formula}")

        # Clear existing data from row 3 (all columns)
        max_col = ws_receipt.max_column
        for row in ws_receipt.iter_rows(min_row=3, max_row=ws_receipt.max_row, max_col=max_col):
            for cell in row:
                cell.value = None

        # Bulk write audit data (all columns)
        for r_idx, row_data in enumerate(audit_data, start=3):
            for c_idx, value in enumerate(row_data, start=1):
                ws_receipt.cell(r_idx, c_idx).value = value

        # Apply W formula from row 4 onwards, only where column R (col 18) has data
        if w1_formula and str(w1_formula).startswith("="):
            w_applied = 0
            for r_idx in range(4, ws_receipt.max_row + 1):
                if ws_receipt.cell(r_idx, 18).value not in (None, ""):
                    adjusted = re.sub(r'(\$?[A-Za-z]+)\$?1(?!\d)', lambda m, r=r_idx: f"{m.group(1)}{r}", str(w1_formula))
                    ws_receipt.cell(r_idx, 23).value = adjusted
                    w_applied += 1
            logging.info(f"Applied W formula to {w_applied} rows from row 4 with data in column R")
        else:
            logging.warning("W1 has no formula — W column left as pasted")

        wb_report.save(report_file)
        logging.info(f"Receipt update complete: {len(audit_data)} rows")

def update_payment_sheet(report_file, payments_file, audit_receipt_file):
    """Update Payment sheet in report_file.

    If payments_file (Payments - *.xlsx) is provided: use it (rows 2+, cols A-G).
    Otherwise: fall back to Audit Receipt - Internal *.xlsx (rows 2+, cols A-G).
    """
    logging.info("Processing Payment sheet update")

    if payments_file:
        source_file = payments_file
        source_label = f"Payments file: {os.path.basename(payments_file)}"
    elif audit_receipt_file:
        source_file = audit_receipt_file
        source_label = f"Audit Receipt file (fallback): {os.path.basename(audit_receipt_file)}"
    else:
        logging.warning("No Payments file and no Audit Receipt file — skipping Payment sheet update")
        return

    def _to_number(val):
        """Convert a numeric-looking text value (e.g. ERP export's '24,900.5')
        to a real float so Excel doesn't flag it as a number stored as text."""
        if isinstance(val, str):
            cleaned = val.replace(",", "").strip()
            try:
                return float(cleaned)
            except ValueError:
                return val
        return val

    logging.info(f"Using {source_label}")
    with timed_operation("Read Payment source data"):
        wb_src = load_workbook(source_file, data_only=True, read_only=True)
        ws_src = wb_src.active

        # Locate the real header row (col A = "Date") instead of assuming row 2.
        # ERP exports (e.g. the GL Account Details report Payments now uses)
        # include a metadata block — Sub Ledger Code, Transaction From/To Date,
        # Opening/Closing Balance — plus a blank row, before the actual column
        # headers and data; a plain manually-uploaded file has headers on row 1
        # with no such block, so this still works for that case too.
        header_row_idx = None
        header_row_vals = None
        for _r, _row in enumerate(ws_src.iter_rows(max_row=30, max_col=8, values_only=True), start=1):
            if _row and str(_row[0]).strip().lower() == "date":
                header_row_idx = _r
                header_row_vals = _row
                break
        data_start_row = (header_row_idx + 1) if header_row_idx else 2

        # A source file only has real per-transaction categories when its own
        # header row labels column H 'Category' — that's true for a manually
        # uploaded Payments file, but the raw ERP GL Account Details export
        # the bot downloads has only 7 columns and no Category at all. This
        # is checked from the actual header text rather than by which source
        # was picked, so it works correctly regardless of which file ends up
        # as payments_file.
        has_source_category_col = bool(
            header_row_vals and len(header_row_vals) > 7 and header_row_vals[7]
            and str(header_row_vals[7]).strip().lower() == "category"
        )

        # max_col=8 also pulls col H (Category): the authoritative source for
        # every row when has_source_category_col is True (looked up by
        # Transaction No — col B — below), and otherwise a fallback for rows
        # the List-sheet keyword search doesn't match.
        src_rows = []
        src_categories = []
        src_txn_to_category = {}
        for row in ws_src.iter_rows(min_row=data_start_row, max_col=8, values_only=True):
            if row and any(v not in (None, "") for v in row):
                row = list(row)
                for _idx in (4, 5, 6):  # Debit Amount, Credit Amount, Running Balance
                    row[_idx] = _to_number(row[_idx])
                cat = row[7] if len(row) > 7 else None
                src_categories.append(cat)
                src_rows.append(tuple(row[:7]))
                if has_source_category_col:
                    txn_no = row[1]   # column B = Transaction No
                    if txn_no not in (None, "") and cat not in (None, ""):
                        src_txn_to_category[str(txn_no).strip()] = str(cat).strip()
        wb_src.close()
    logging.info(
        f"Read {len(src_rows)} data rows from {source_label} "
        f"(header row: {header_row_idx or 'none found, defaulted to row 1'}"
        f"{', with per-transaction Category column' if has_source_category_col else ''})"
    )

    with timed_operation("Write to Payment sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        if "Payment" not in wb_report.sheetnames:
            logging.warning("Payment sheet not found — skipping")
            wb_report.close()
            return

        ws_pay = wb_report["Payment"]

        # Read H1 formula template (col 8) before clearing
        h1_formula = ws_pay.cell(1, 8).value
        logging.info(f"H1 formula template: {h1_formula}")

        # Step 1: Clear existing data from row 3, cols A-G only
        for row in ws_pay.iter_rows(min_row=3, max_row=ws_pay.max_row, min_col=1, max_col=7):
            for cell in row:
                cell.value = None

        # Step 2: Paste data into cols A-G
        last_data_row = 2
        for r_idx, row_data in enumerate(src_rows, start=3):
            for c_idx, val in enumerate(row_data, start=1):
                ws_pay.cell(r_idx, c_idx).value = val
            last_data_row = r_idx

        # Step 3: Clear H column from row 3 to last data row, then apply formula
        for r_idx in range(3, last_data_row + 1):
            ws_pay.cell(r_idx, 8).value = None

        def _fallback_category(src_cat):
            """Fall back to the source file's own Category when the List-sheet
            keyword search finds no match, instead of a bare 'No Category'."""
            if src_cat not in (None, "") and str(src_cat).strip().lower() != "no category":
                return str(src_cat).strip().replace('"', '""')
            return "No Category"

        if has_source_category_col:
            # The source file has its own authoritative Category per
            # transaction — use that directly as a literal value instead of
            # the List-sheet keyword-search formula, looked up by each row's
            # own Transaction No (col B) rather than assumed row order, so a
            # transaction always gets its own correct category regardless of
            # how the destination rows are ordered.
            matched = 0
            for r_idx in range(3, last_data_row + 1):
                txn_no = ws_pay.cell(r_idx, 2).value
                cat = src_txn_to_category.get(str(txn_no).strip()) if txn_no not in (None, "") else None
                ws_pay.cell(r_idx, 8).value = cat or "No Category"
                if cat:
                    matched += 1
            logging.info(
                f"Set Category directly from source file's own column, matched by Transaction No: "
                f"{matched}/{last_data_row - 2} rows"
            )
        elif h1_formula and str(h1_formula).startswith("="):
            for r_idx in range(3, last_data_row + 1):
                adjusted = re.sub(r'(\$?[A-Za-z]+)\$?1(?!\d)', lambda m, r=r_idx: f"{m.group(1)}{r}", str(h1_formula))
                fallback = _fallback_category(src_categories[r_idx - 3])
                adjusted = re.sub(r',\s*"[^"]*"\s*\)\s*$', f',"{fallback}")', adjusted)
                ws_pay.cell(r_idx, 8).value = adjusted
            logging.info(f"Applied H formula (from H1 template) to rows 3-{last_data_row}")
        else:
            # H1 has no template — write hardcoded IFERROR/INDEX/MATCH formula as array formula
            # Using ArrayFormula prevents openpyxl from inserting the @ implicit intersection operator
            from openpyxl.worksheet.formula import ArrayFormula
            for r_idx in range(3, last_data_row + 1):
                fallback = _fallback_category(src_categories[r_idx - 3])
                formula = (
                    f'IFERROR(INDEX(List!$B$2:$B$205,'
                    f'MATCH(TRUE,ISNUMBER(SEARCH(List!$A$2:$A$205,C{r_idx})),0)),"{fallback}")'
                )
                ws_pay.cell(r_idx, 8).value = ArrayFormula(f"H{r_idx}", f"={formula}")
            logging.info(f"Applied hardcoded IFERROR/INDEX/MATCH H formula to rows 3-{last_data_row}")

        wb_report.save(report_file)
        source_name = "Payments file" if payments_file else "Audit Receipt file"
        logging.info(f"Payment sheet update complete: {len(src_rows)} rows from {source_name}")

# ---------------------------------------------------------
# Manual Process Reminder Helper
# ---------------------------------------------------------
def _log_manual_reminder(title, summary, instructions, closing_message):
    """Helper function to log manual process reminders consistently."""
    separator = "=" * 80

    logging.info(separator)
    logging.info(f"MANUAL PROCESS REQUIRED: {title}")
    logging.info(separator)

    print(f"\n{separator}")
    print(f"MANUAL PROCESS REQUIRED: {title}")
    print(separator)

    logging.info(summary)
    print(f"\n{summary}\n")

    for instruction in instructions:
        logging.info(instruction)
        print(instruction)

    print("")
    logging.info(separator)
    logging.info(closing_message)
    logging.info(separator)

    print(separator)
    print(closing_message)
    print(separator + "\n")

# ---------------------------------------------------------
# Step 11 – Manual Process Reminder: SasiaNet Cashflow Data Entry
# ---------------------------------------------------------
def log_manual_sasianet_reminder(input_date):
    date_str = input_date.strftime("%m/%d/%Y")
    _log_manual_reminder(
        "SasiaNet Cashflow Data Entry",
        "SasiaNet sheet must be updated manually",
        [
            "INSTRUCTIONS FOR SASIANET DATA ENTRY:",
            "",
            "1. Login to SasiaNet",
            "2. Go to Cashflow",
            f"3. Select the required date (starting from {date_str})",
            "4. Copy the data",
            "5. Paste to the SasiaNet sheet in Cash-Flow-Forecast file",
            "6. Repeat steps 3-5 for each date until the final date",
        ],
        "Please complete the manual SasiaNet data entry process"
    )

# ---------------------------------------------------------
# Step 12 – Manual Process Reminder: Audit Receipt File Download
# ---------------------------------------------------------
def log_manual_audit_receipt_reminder():
    _log_manual_reminder(
        "Audit Receipt File Download",
        "Audit Receipt file must be downloaded manually from ERP",
        [
            "INSTRUCTIONS FOR AUDIT RECEIPT FILE DOWNLOAD:",
            "",
            "1. Open ERP system",
            "2. Navigate to: ERP > Report > Reports > Audit Receipts - Internal",
            "3. Generate the report",
            "4. Download/Export the file",
            "5. Save the file in the working folder with 'Audit Receipt' in the filename",
        ],
        "Please ensure the Audit Receipt file is available before running the script"
    )

# ---------------------------------------------------------
# Step 13 – Manual Process Reminder: Treasury Input Data Entry
# ---------------------------------------------------------
def log_manual_treasury_input_reminder(treasury_date):
    date_str = treasury_date.strftime("%m/%d/%Y")
    _log_manual_reminder(
        "Treasury Input Data Entry",
        f"Treasury Input Column D data must be entered manually for date: {date_str}",
        [
            "INSTRUCTIONS FOR TREASURY INPUT DATA ENTRY:",
            "",
            f"Date to process: {date_str}",
            "",
            "1. Open SasiaNet",
            "2. Navigate to: Front Office --> Reports --> M2M Report - Debt",
            f"3. Enter the required date: {date_str}",
            "4. Set Instrument type --> TBOND",
            "5. Set Portfolio Code --> ALCL",
            "6. Run the file",
            "",
            "COLLECT THE FOLLOWING VALUES:",
            "7. M2M Gain --> Net M2M Gain / (Loss) - Total",
            "8. Interest Receivable (Coupon Accrued Interest) --> Coupon Accrued Interest (Total)",
            "9. Cost of Investment (Cost + Price Accrued) --> Cost (Total) + Price Accrued (Total)",
        ],
        "Please complete the manual Treasury Input data entry process"
    )

# ---------------------------------------------------------
# Step 16b – Monthly Cum Collection Pivot (week 1 only)
# ---------------------------------------------------------
def update_cum_monthly_collection(report_file: str, date_folder: str, input_date: datetime):
    """
    Week-1 only.  Build a pivot of last month's Audit Receipt (monthly) file
    and write it to the 'Last Cum month collection' sheet of report_file.

    Layout:
      - Rows    = cat (from Types of Receipt Type2 lookup)
      - Columns = day-of-month (1–31), with Grand Total always at column index 31
                  so it lands in sheet column AH regardless of month length.
      - Values  = SUM of RECIPT_AMOUNT

    Columns A (row labels) + B…AH (days 1–31 + Grand Total) = 34 columns total.
    """
    logging.info("Processing monthly cum collection pivot (week 1)")

    # ── 1. Find monthly audit receipt file ──────────────────────────────────
    monthly_file = None
    for _f in Path(date_folder).glob("Audit Receipt - Internal (monthly) *.xlsx"):
        monthly_file = str(_f)
        break
    if not monthly_file:
        logging.warning("No 'Audit Receipt - Internal (monthly) *.xlsx' found — skipping cum collection")
        return

    # ── 2. Load monthly audit receipt into dataframe ─────────────────────────
    with timed_operation("Read monthly Audit Receipt"):
        wb_m = load_workbook(monthly_file, data_only=True, read_only=True)
        ws_m = wb_m.active
        rows_m = list(ws_m.iter_rows(values_only=True))
        wb_m.close()

    if not rows_m:
        logging.warning("Monthly Audit Receipt file is empty — skipping")
        return

    headers_m = [str(h).strip() if h is not None else "" for h in rows_m[0]]
    df_monthly = pd.DataFrame(rows_m[1:], columns=headers_m)
    df_monthly = df_monthly.dropna(how="all")
    logging.info(f"Monthly Audit Receipt: {len(df_monthly)} rows, columns: {list(df_monthly.columns)}")

    # Normalise required column names (case-insensitive match)
    col_map = {}
    for col in df_monthly.columns:
        col_map[col.upper()] = col
    recipt_type_col   = col_map.get("RECIPT_TYPE",   col_map.get("RECEIPT_TYPE",   None))
    recipt_date_col   = col_map.get("RECIPT_DATE",   col_map.get("RECEIPT_DATE",   None))
    recipt_amount_col = col_map.get("RECIPT_AMOUNT", col_map.get("RECEIPT_AMOUNT", None))

    missing = [n for n, c in [("RECIPT_TYPE", recipt_type_col),
                               ("RECIPT_DATE", recipt_date_col),
                               ("RECIPT_AMOUNT", recipt_amount_col)] if c is None]
    if missing:
        logging.warning(f"Monthly Audit Receipt missing columns {missing} — skipping")
        return

    # ── 3. Load Types of Receipt B,C,D - df_typ_of_rec ──────────────────────
    with timed_operation("Read Types of Receipt from report_file"):
        wb_r = load_workbook(report_file, keep_vba=True, data_only=True)
        if "Types of Receipt" not in wb_r.sheetnames:
            logging.warning("'Types of Receipt' sheet not found — skipping cum collection")
            wb_r.close()
            return
        ws_tor = wb_r["Types of Receipt"]
        tor_rows = list(ws_tor.iter_rows(min_col=2, max_col=4, values_only=True))
        wb_r.close()

    if not tor_rows:
        logging.warning("Types of Receipt sheet is empty — skipping")
        return

    # Row 1 = header (Type1 in B=col0, something in C=col1, Type2 in D=col2)
    tor_headers = [str(h).strip() if h else "" for h in tor_rows[0]]
    df_tor = pd.DataFrame(tor_rows[1:], columns=tor_headers)
    df_tor = df_tor.dropna(how="all")
    logging.info(f"Types of Receipt: {len(df_tor)} rows, columns: {tor_headers}")

    # Identify Type1 (B) and Type2 (C) columns by header name, falling back to position
    if len(tor_headers) < 3:
        logging.warning("Types of Receipt has fewer than 3 columns (B,C,D) — skipping")
        return
    # Try to find by name first
    tor_headers_upper = [h.upper() for h in tor_headers]
    type1_idx = next((i for i, h in enumerate(tor_headers_upper) if "TYPE1" in h or h == "TYPE 1"), 0)
    type2_idx = next((i for i, h in enumerate(tor_headers_upper) if "TYPE2" in h or h == "TYPE 2"), 2)
    type1_col = tor_headers[type1_idx]
    type2_col = tor_headers[type2_idx]
    logging.info(f"Using Type1 col='{type1_col}' (idx {type1_idx}), Type2 col='{type2_col}' (idx {type2_idx})")

    # Build lookup dict: Type1 value - Type2 value
    type_lookup = df_tor.dropna(subset=[type1_col, type2_col]).set_index(type1_col)[type2_col].to_dict()
    logging.info(f"Type lookup entries: {len(type_lookup)}")

    # ── 4. Add 'cat' column via vlookup ──────────────────────────────────────
    df_monthly["cat"] = df_monthly[recipt_type_col].map(type_lookup)
    logging.info(f"df_monthly after cat lookup (first 10 rows):\n{df_monthly.head(10).to_string()}")
    logging.info(f"Unique cat values: {df_monthly['cat'].dropna().unique().tolist()}")
    logging.info(f"Unmapped rows (cat is NaN): {df_monthly['cat'].isna().sum()}")

    # ── 5. Prepare for pivot ──────────────────────────────────────────────────
    df_monthly[recipt_date_col] = pd.to_datetime(df_monthly[recipt_date_col], errors="coerce")
    df_monthly[recipt_amount_col] = pd.to_numeric(df_monthly[recipt_amount_col], errors="coerce").fillna(0)

    # Derive last month's year/month from input_date
    _prev_last    = input_date.replace(day=1) - timedelta(days=1)
    prev_year     = _prev_last.year
    prev_month    = _prev_last.month
    days_in_month = monthrange(prev_year, prev_month)[1]

    # Full date sequence for previous month
    all_dates = [date(prev_year, prev_month, d) for d in range(1, days_in_month + 1)]

    # ── 6. Build pivot: rows=cat, columns=actual date ────────────────────────
    df_monthly["_date_only"] = df_monthly[recipt_date_col].dt.date

    pivot = df_monthly.pivot_table(
        index="cat",
        columns="_date_only",
        values=recipt_amount_col,
        aggfunc="sum",
        fill_value=0,
    )

    # Ensure every day of the month has a column, in date order
    for _d in all_dates:
        if _d not in pivot.columns:
            pivot[_d] = 0
    pivot = pivot[all_dates]   # strictly ordered, only prev-month dates

    # Grand Total column (row sum)
    pivot["Grand Total"] = pivot.sum(axis=1)

    # Grand Total row (column sum)
    grand_total_row        = pivot.sum(axis=0)
    grand_total_row.name   = "Grand Total"
    pivot = pd.concat([pivot, grand_total_row.to_frame().T])

    # Reset index — cat becomes a plain column
    pivot = pivot.reset_index().rename(columns={"cat": "cat", "index": "cat"})
    logging.info(f"Pivot table (cum monthly collection):\n{pivot.to_string()}")

    # ── 7. Write to 'Last Cum month collection' sheet ────────────────────────
    # Column layout (1-based sheet columns):
    #   col A  (1)          = cat / row label  (no header cell in A1)
    #   col B  (2)          = date 1
    #   col B+days-1        = date last-day-of-month
    #   cols after last date up to col 32 = empty (padding for short months)
    #   col AH (33)         = Grand Total  (always col 33 regardless of month length)
    GRAND_TOTAL_SHEET_COL = 34   # AH

    with timed_operation("Write cum monthly collection pivot to sheet"):
        wb_w = load_workbook(report_file, keep_vba=True)

        sheet_name = None
        for _sn in wb_w.sheetnames:
            if "cum" in _sn.lower() and "collection" in _sn.lower():
                sheet_name = _sn
                break
        if sheet_name is None:
            logging.warning("'Last Cum month collection' sheet not found in report — skipping")
            wb_w.close()
            return

        from openpyxl.styles import PatternFill, Font

        ws_cum = wb_w[sheet_name]

        _month_header_exists = any(
            isinstance(ws_cum.cell(_r, 3).value, datetime)
            and ws_cum.cell(_r, 3).value.year == all_dates[0].year
            and ws_cum.cell(_r, 3).value.month == all_dates[0].month
            and ws_cum.cell(_r, GRAND_TOTAL_SHEET_COL).value == "Grand Total"
            for _r in range(1, ws_cum.max_row + 1)
        )
        if _month_header_exists:
            logging.info(f"'Last Cum month collection' already has the {all_dates[0]:%b-%Y} block — not appending again")
            wb_w.close()
            return

        _BLUE_FILL = PatternFill(fill_type="solid", fgColor="D8E9F1")
        _BOLD_FONT = Font(bold=True)

        # Find last row with any data, then write starting at last_row + 1
        last_data_row = 0
        for _r in range(ws_cum.max_row, 0, -1):
            if any(ws_cum.cell(_r, _c).value not in (None, "") for _c in range(1, GRAND_TOTAL_SHEET_COL + 1)):
                last_data_row = _r
                break
        start_row = last_data_row + 2
        logging.info(f"Last Cum month collection: last data row={last_data_row}, writing from row={start_row}")

        # Header row: col B = blank, col C onward = dates, col AH = "Grand Total"
        # Style: blue fill + bold
        for idx, _d in enumerate(all_dates, start=3):        # C=3 onward (B reserved for cat label)
            _dt = datetime(_d.year, _d.month, _d.day)
            cell = ws_cum.cell(start_row, idx)
            cell.value = _dt
            cell.number_format = "M/D/YYYY"
            cell.fill = _BLUE_FILL
            cell.font = _BOLD_FONT
        _gt_header_cell = ws_cum.cell(start_row, GRAND_TOTAL_SHEET_COL)
        _gt_header_cell.value = "Grand Total"
        _gt_header_cell.fill = _BLUE_FILL
        _gt_header_cell.font = _BOLD_FONT

        # Data rows (cat in col B, values in C onward, Grand Total in AH)
        total_data_rows = len(pivot)
        for r_offset, row in enumerate(pivot.itertuples(index=False), start=1):
            row_dict = dict(zip(pivot.columns, row))
            r = start_row + r_offset
            is_last_row = (r_offset == total_data_rows)  # Grand Total row

            cat_cell = ws_cum.cell(r, 2)
            cat_cell.value = row_dict.get("cat")
            if is_last_row:
                cat_cell.fill = _BLUE_FILL
                cat_cell.font = _BOLD_FONT

            for idx, _d in enumerate(all_dates, start=3):
                cell = ws_cum.cell(r, idx)
                cell.value = row_dict.get(_d, 0)
                if is_last_row:
                    cell.fill = _BLUE_FILL
                    cell.font = _BOLD_FONT

            gt_cell = ws_cum.cell(r, GRAND_TOTAL_SHEET_COL)
            gt_cell.value = row_dict.get("Grand Total", 0)
            if is_last_row:
                gt_cell.fill = _BLUE_FILL
                gt_cell.font = _BOLD_FONT

        wb_w.save(report_file)
        logging.info(f"Written cum monthly collection pivot: {len(pivot)-1} cat rows + Grand Total row - '{sheet_name}'")

# ---------------------------------------------------------
# Step 16c – Append new rows to Collection sheet (week 1 only)
# ---------------------------------------------------------
def update_collection_sheet(report_file: str, input_date: datetime, week_number: int = 0):
    """
    Week-1 only.  In the 'Collection' sheet, append a new data row to three sections.
    Every appended row copies the full cell formatting (font, fill, borders,
    number formats) of the row above it so the sections stay visually consistent.

    Sections are processed in top-to-bottom sheet order, because each insert
    shifts all rows below it without updating formula text:

    1. "Colelction per day" (top):
       - Col A = previous-month label (same type/format as the row above)
       - Col C–AH = previous row's sheet-ref formula +6 on the row index

    2. "Average collection per day" (middle):
       - Col A = prev-month label
       - Col C–AH = previous row's formula with row refs +1, so the new row
         divides the section-1 row just appended by its AH total
         (e.g. =AC39/$AH39 → =AC40/$AH40)

    3. "Fixed Deposits" (bottom, last so the section-2 insert has already
       shifted it into final position):
       - Same logic as section 1: Col A = prev-month label, +6 ref formulas

    After each insert, every range formula in the totals/average row directly
    below the section (SUM, AVERAGE, … across all columns) has its range end
    re-anchored to the newly inserted row.

    Then AJ21 ("Current month") is re-pointed at the AH cell of the
    prior-year same-month row (same target row as "Update to past same
    month"), e.g. a Dec-25 run points AJ21 at Dec-24's AH cell.
    """
    import re as _re

    logging.info("Updating Collection sheet")

    wb = load_workbook(report_file, keep_vba=True)

    coll_sheet = None
    for _sn in wb.sheetnames:
        if _sn.strip().lower() == "collection":
            coll_sheet = wb[_sn]
            break
    if coll_sheet is None:
        logging.warning("'Collection' sheet not found — skipping")
        wb.close()
        return

    ws = coll_sheet

    # ── Update C1:AG1 for the 'Next 22-28 days' window only ──────────────────
    # Each column is a FIXED day-of-month slot (C=day1, D=day2, … AG=day31),
    # not a rolling 31-day window — so only the 7 columns matching this run's
    # 'Next 22-28 days' dates (Weekly Distributions row 2, cols K:L) get
    # touched; the other 24 keep whatever they were set to on a previous run.
    # K2:L2 there is itself just week_start+28 .. week_start+34 (built off
    # 'NBD-WF-18-ECS'!H2, which this script sets to week_start_date), so it's
    # computed directly here rather than read back from the sheet.
    # Writing straight to the day-of-month column (instead of a fixed offset
    # from week_start) means a window that crosses a month boundary (e.g.
    # 29-Dec..4-Jan) naturally updates the right columns to the new month/
    # year while leaving the rest alone — and a short month (Feb, or any
    # 30-day month) just never produces a day-31 (or day-30) date to write,
    # so that column is left untouched rather than forced to an invalid date.
    week_start = input_date - timedelta(days=input_date.weekday())
    _updated_cols = []
    for _day_offset in range(28, 35):   # the 7 days of 'Next 22-28 days'
        _d = week_start + timedelta(days=_day_offset)
        _col = 3 + (_d.day - 1)   # day-of-month 1 -> C(3) … day-of-month 31 -> AG(33)
        cell = ws.cell(1, _col)
        cell.value = datetime(_d.year, _d.month, _d.day)
        cell.number_format = "M/D/YYYY"
        _updated_cols.append(f"{get_column_letter(_col)}1={_d.strftime('%m/%d/%Y')}")
    logging.info(f"Updated 'Next 22-28 days' date columns: {', '.join(_updated_cols)}")

    prev_month_first = (input_date.replace(day=1) - timedelta(days=1)).replace(day=1)

    def _find_label_row(label_substr: str, start_from: int = 1) -> int | None:
        """Return the first row (>= start_from) whose col A contains label_substr (case-insensitive)."""
        for _r in range(start_from, ws.max_row + 1):
            _v = ws.cell(_r, 1).value
            if _v and label_substr.lower() in str(_v).strip().lower():
                return _r
        return None

    def _find_last_data_row(after_row: int) -> int:
        """Return the last non-empty col-A row in the block after after_row.
        Skips any leading empty rows, then tracks consecutive non-empty rows."""
        last = after_row
        in_data = False
        for _r in range(after_row + 1, ws.max_row + 2):
            _v = ws.cell(_r, 1).value
            if _v not in (None, ""):
                in_data = True
                last = _r
            elif in_data:
                # first empty row after data started — stop
                break
        return last

    def _increment_ref(formula: str, increment: int = 6) -> str:
        """Increment row numbers inside cross-sheet references by increment."""
        return _re.sub(
            r"((?:'[^']+'|[A-Za-z0-9_]+)![A-Z]+)(\d+)",
            lambda m: f"{m.group(1)}{int(m.group(2)) + increment}",
            formula
        )

    def _update_totals_row(total_row: int, new_end: int):
        """Re-anchor the end row of every range reference in the totals/average
        row below an inserted row (e.g. =SUM(AH14:AH40) → =SUM(AH14:AH41),
        =AVERAGE(AF42:AF77) → =AVERAGE(AF42:AF79)) so the range includes the
        row just inserted. Scans cols B..AJ; cells without range formulas are
        left untouched."""
        updated = 0
        for _c in range(2, 37):   # B=2 … AJ=36
            _cell = ws.cell(total_row, _c)
            _f = _cell.value
            if _f and isinstance(_f, str) and _f.startswith("=") and ":" in _f:
                _new_f = _re.sub(
                    r"(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)",
                    lambda m: f"{m.group(1)}{new_end}",
                    _f
                )
                if _new_f != _f:
                    _cell.value = _new_f
                    updated += 1
        if updated:
            logging.info(f"  Totals row {total_row}: re-anchored {updated} range formulas to end row {new_end}")
        else:
            logging.info(f"  Totals row {total_row}: no range formulas found — nothing updated")

    def _copy_row_format(src_row: int, dst_row: int, max_col: int = 36):
        """Copy cell styles (font, fill, border, alignment, number format) from
        src_row onto dst_row for cols A..AJ, so appended rows match the section's
        existing formatting."""
        from copy import copy as _style_copy
        for _c in range(1, max_col + 1):
            _s = ws.cell(src_row, _c)
            _d = ws.cell(dst_row, _c)
            if _s.has_style:
                _d.font = _style_copy(_s.font)
                _d.fill = _style_copy(_s.fill)
                _d.border = _style_copy(_s.border)
                _d.alignment = _style_copy(_s.alignment)
                _d.protection = _style_copy(_s.protection)
                _d.number_format = _s.number_format

    def _write_col_a(new_row: int, prev_row: int):
        """Write the previous-month label in col A, mirroring the type of the row
        above: a real date if the section stores dates (renders as Mmm-yy via the
        copied number format), otherwise a 'Mon-YY' string."""
        cell_a = ws.cell(new_row, 1)
        prev_a = ws.cell(prev_row, 1).value
        if isinstance(prev_a, (datetime, date)):
            cell_a.value = datetime(prev_month_first.year, prev_month_first.month, prev_month_first.day)
        else:
            cell_a.value = prev_month_first.strftime("%b-%y")

    def _append_ref_row(label_substr: str, start_from: int = 1) -> int:
        """
        Find section label, insert new row after last data row, write:
          - Col A: previous-month label, same type/format as the row above
          - Col C–AH: previous row's cross-sheet formula with +6 row index
          - Totals/average row below: re-anchor all range formulas to the new row
        Returns the new_row number (after insertion), or 0 on failure.
        """
        label_row = _find_label_row(label_substr, start_from)
        if label_row is None:
            logging.warning(f"Label '{label_substr}' not found in Collection col A — skipping")
            return 0

        logging.info(f"Found '{label_substr}' at row {label_row}")
        last_data_row = _find_last_data_row(label_row)
        new_row = last_data_row + 1
        logging.info(f"  last data row={last_data_row}, inserting at row={new_row}")

        ws.insert_rows(new_row)
        _copy_row_format(last_data_row, new_row)
        _write_col_a(new_row, last_data_row)

        # Col C–AH: +6 cross-sheet ref from previous row
        prev_c = ws.cell(last_data_row, 3).value
        logging.info(f"  Previous C formula: {prev_c}")
        if prev_c and str(prev_c).startswith("="):
            ws.cell(new_row, 3).value = _increment_ref(str(prev_c))
            for _col in range(4, 35):   # D=4 … AH=34
                prev_f = ws.cell(last_data_row, _col).value
                if prev_f and str(prev_f).startswith("="):
                    ws.cell(new_row, _col).value = _increment_ref(str(prev_f))
        else:
            logging.warning(f"  Col C of row {last_data_row} has no formula — formulas not copied")

        # Totals/average row immediately below the new row: extend its ranges
        # to include the inserted row (all columns, not just AH)
        _update_totals_row(new_row + 1, new_row)

        return new_row

    if week_number == 1:
      # Log all non-empty col A values to identify exact label text
      logging.info("Collection sheet col A labels (all non-empty):")
      for _r in range(1, ws.max_row + 1):
          _v = ws.cell(_r, 1).value
          if _v not in (None, ""):
              logging.info(f"  row {_r}: {repr(str(_v).strip())}")

    # ── Week 1 only: append new rows to the three sections ───────────────────
    if week_number != 1:
        wb.save(report_file)
        logging.info("Collection sheet C1:AG1 updated")
        return

    # Sections are processed in TOP-TO-BOTTOM sheet order (Colelction per day →
    # Average collection per day → Fixed Deposits). Each insert shifts every row
    # below it down by one WITHOUT updating formula text, so a section's totals
    # must only be located and re-anchored after all inserts above it are done.

    # ── Section 1: "Colelction per day" — data rows BELOW the label ──────────
    _sec1_row = _find_label_row("colelction per day")
    _sec2_row = _find_label_row("average collection per day")
    _month_row_exists = _sec1_row is not None and any(
        isinstance(ws.cell(_r, 1).value, datetime)
        and ws.cell(_r, 1).value.year == prev_month_first.year
        and ws.cell(_r, 1).value.month == prev_month_first.month
        for _r in range(_sec1_row + 1, (_sec2_row or ws.max_row + 1))
    )
    if _month_row_exists:
        logging.info(f"Collection sheet already has the {prev_month_first:%b-%Y} rows — not appending again")
        wb.save(report_file)
        return

    nr1 = _append_ref_row("colelction per day")

    # ── Section 2: "Average collection per day" ──────────────────────────────
    # Each row here divides the matching "Colelction per day" row by its AH total
    # (e.g. Oct-25 → =AC39/$AH39). Continue the pattern: copy the previous row's
    # formulas with row refs advanced by 1 so the new row points at the section-1
    # row just appended (=AC39/$AH39 → =AC40/$AH40).
    _label3 = "average collection per day"
    _label3_row = _find_label_row(_label3, start_from=1)
    if _label3_row is None:
        logging.warning(f"Label '{_label3}' not found in Collection col A — skipping")
    else:
        logging.info(f"Found '{_label3}' at row {_label3_row}")
        _last3 = _find_last_data_row(_label3_row)
        _new3 = _last3 + 1
        logging.info(f"  last data row={_last3}, inserting at row={_new3}")
        ws.insert_rows(_new3)
        _copy_row_format(_last3, _new3)
        _write_col_a(_new3, _last3)
        _prev_c3 = ws.cell(_last3, 3).value
        logging.info(f"  Previous C formula: {_prev_c3}")
        if _prev_c3 and str(_prev_c3).startswith("="):
            for _col in range(3, 35):   # C=3 … AH=34
                _prev_f = ws.cell(_last3, _col).value
                if _prev_f and str(_prev_f).startswith("="):
                    ws.cell(_new3, _col).value = _re.sub(
                        r"(\$?[A-Z]+\$?)(\d+)",
                        lambda m: f"{m.group(1)}{int(m.group(2)) + 1}",
                        str(_prev_f)
                    )
            logging.info(f"  Copied row-{_last3} formulas with row refs +1 into row {_new3}")
        else:
            logging.warning(f"  Col C of row {_last3} has no formula — formulas not copied")

        # Averages row immediately below: extend its ranges to the new row
        # (e.g. =AVERAGE(AF42:AF77) → =AVERAGE(AF42:AF79))
        _update_totals_row(_new3 + 1, _new3)

    # ── Section 3: "Fixed Deposits" — the lowest section, processed LAST so the
    # average-section insert above has already shifted it into final position
    # before its rows are located and its totals row is re-anchored ───────────
    nr2 = _append_ref_row("fixed deposits", start_from=nr1 if nr1 else 1)

    # ── AI21: re-anchor AVERAGE range end to the new section-1 row ───────────
    ai21 = ws.cell(21, 35)   # AI = col 35
    ai21_formula = ai21.value
    logging.info(f"AI21 formula: {ai21_formula}")
    if ai21_formula and str(ai21_formula).startswith("=") and nr1:
        updated_ai21 = _re.sub(
            r"(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)",
            lambda m: f"{m.group(1)}{nr1}",
            str(ai21_formula)
        )
        ai21.value = updated_ai21
        logging.info(f"Updated AI21: {updated_ai21}")
    else:
        logging.warning("AI21 has no formula or section-1 row missing — skipping")

    # ── AJ21 "Current month": same calendar month as the RUN DATE, one year
    # earlier — the same target row as "Update to past same month" below
    # (nr1 - 11, see that block's comment for the derivation). Recomputed
    # every run from nr1 so it never drifts, regardless of how many rows
    # have been inserted above it.
    if nr1 and nr1 - 11 >= 4:
        ws.cell(21, 36).value = f"=+AH{nr1 - 11}"   # AJ = col 36
        logging.info(f"AJ21 = =+AH{nr1 - 11} (prior-year same-month row in 'Colelction per day')")
    else:
        logging.warning("Section-1 row missing or no prior-year row available — AJ21 not updated")

    # ── "Update to past same month": same calendar month as the RUN DATE
    # (input_date), one year earlier. nr1 is the row just appended for the
    # PREVIOUS month (this code only runs in week 1, when the prior month's
    # data is finalized — e.g. a Dec-2025 run appends "Nov-25" at nr1). Since
    # the run month is always exactly one month after nr1's month, the same
    # month one year back is nr1 - 12 + 1 = nr1 - 11 (e.g. new row = Nov-25 at
    # row 40 → run month is December → past-year Dec-24 is at row 29, not
    # Nov-24 at row 28). Unlike AJ21, nothing else recomputes this cell, so it
    # was left pinned to whatever row was correct when it was last set by
    # hand, drifting one month stale every week-1 run thereafter.
    upsm_row = _find_label_row("update to past same month", start_from=1)
    if upsm_row is None:
        logging.warning("'Update to past same month' label not found — skipping")
    elif nr1 and nr1 - 11 >= 4:
        ws.cell(upsm_row, 2).value = f"=+AH{nr1 - 11}"
        logging.info(f"'Update to past same month' (row {upsm_row}) B = =+AH{nr1 - 11}")
    else:
        logging.warning("Section-1 row missing or no prior-year row available — 'Update to past same month' not updated")

    # ── "Debtor" row: for each day-of-month column (C:AG), the value is
    # (Debtor total, this row's own B cell) x (that day's average-collection
    # % from the row directly above 'Factor') x (Factor, Factor row's B cell).
    # This whole block ('Average'/Total row, Factor, Update to past same
    # month, Debtor) shifts down by one row each week-1 run, in lockstep,
    # because it sits immediately below the 'Average collection per day'
    # section's newly inserted row. Row positions are looked up by label
    # every run rather than carried forward, so the formula stays correctly
    # anchored regardless of how far the block has drifted.
    factor_row = _find_label_row("factor", start_from=1)
    debtor_row = _find_label_row("debtor", start_from=1)
    if factor_row is None or debtor_row is None:
        logging.warning("'Factor' or 'Debtor' label not found — skipping Debtor row formula refresh")
    else:
        avg_row = factor_row - 1  # the section-2 average/total row, directly above 'Factor'

        # ── "Factor" row's own B-column formula: (Update to past same
        # month's B value) / $AI$21. Drifts the same way as Debtor below —
        # openpyxl's insert_rows() shifts the cell down without updating the
        # row number in its formula text — so it's re-derived by label every
        # run instead of being carried forward.
        if upsm_row is not None:
            ws.cell(factor_row, 2).value = f"=+B{upsm_row}/$AI$21"
            logging.info(f"Factor row {factor_row}: B = =+B{upsm_row}/$AI$21 (Update to past same month row {upsm_row})")
        else:
            logging.warning("'Update to past same month' row not found — skipping Factor B formula refresh")

        updated_cols = 0
        for _col in range(3, 34):  # C=3 … AG=33
            _col_letter = get_column_letter(_col)
            ws.cell(debtor_row, _col).value = f"=$B${factor_row}*{_col_letter}{avg_row}*$B${debtor_row}"
            updated_cols += 1
        logging.info(f"Debtor row {debtor_row}: refreshed {updated_cols} formulas (Factor row {factor_row}, Average row {avg_row})")

    # ── "Other Collections" row: for each day-of-month column (C:AG), the
    # value is (Other Collections' own B total) x (that day's average-
    # collection % from the row directly above 'Factor'). Same drift pattern
    # as Debtor above, but only 2 factors, not 3 — the reference template
    # (Week 48 file) has this row's formula as =$B$83*C78 (its OWN B column,
    # not Debtor's), so re-derived here as $B$<other_collections_row> rather
    # than reusing debtor_row.
    other_collections_row = _find_label_row("other collections", start_from=1)
    if other_collections_row is None or factor_row is None:
        logging.warning("'Other Collections' or 'Factor' label not found — skipping Other Collections row formula refresh")
    else:
        _avg_row = factor_row - 1
        updated_oc_cols = 0
        for _col in range(3, 34):  # C=3 … AG=33
            _col_letter = get_column_letter(_col)
            ws.cell(other_collections_row, _col).value = f"=$B${other_collections_row}*{_col_letter}{_avg_row}"
            updated_oc_cols += 1
        logging.info(f"Other Collections row {other_collections_row}: refreshed {updated_oc_cols} formulas (Average row {_avg_row})")

    # ── AH-column "self total" formulas (=SUM(C{row}:AG{row})) drift the same
    # way as the Debtor row above, and it isn't limited to the Factor/Debtor/
    # Other Collections block — the 'Average collection per day' section's own
    # historical monthly rows (its data sits ABOVE 'Factor', so Section 1's
    # insert shifts the whole section down by one row every week-1 run) and
    # the 'Fixed Deposits' section's monthly rows further down are affected
    # the exact same way: insert_rows() shifts the cell down without updating
    # the row number baked into its own formula text, so each ends up
    # referencing a row a few positions too high (e.g. AH58 pointing at
    # =SUM(C57:AG57) instead of its own =SUM(C58:AG58)). Re-anchor every such
    # formula from the 'Average collection per day' section's first data row
    # all the way down through the end of the 'Fixed Deposits' data block
    # (nr2, the row just appended there) — bounded and re-derived every run,
    # so it keeps working regardless of how many rows have been inserted
    # above it. The regex only matches self-referencing SUM(C:AG) formulas,
    # so it naturally skips the cross-sheet-linked rows within Fixed Deposits
    # (e.g. ='Last  Cum month collection'!AH220) without needing to special-
    # case them.
    #
    # A second, unrelated pattern also shows up further down in the Average-
    # collection-per-day section: some rows carry '=AH{x}/$AH{x}' instead of
    # a real self-total — always evaluating to 100%, not an actual sum of
    # that row's own C:AG values like every row above it. Any row using that
    # ratio pattern is unconditionally replaced with a proper self-total too,
    # so the whole section (through to the Average/Totals row) uses one
    # consistent formula shape.
    if factor_row is not None:
        _self_sum_re = _re.compile(r"^=\s*SUM\(\s*C(\d+)\s*:\s*AG\d+\s*\)\s*$", _re.IGNORECASE)
        _ratio_re = _re.compile(r"^=\s*AH\d+\s*/\s*\$AH\d+\s*$", _re.IGNORECASE)
        block_start = (_label3_row + 1) if _label3_row else factor_row
        block_end = (nr2 + 1) if nr2 else (_find_label_row("fixed deposits", start_from=factor_row) or (factor_row + 10))
        fixed_count = 0
        for _r in range(block_start, block_end):
            _cell = ws.cell(_r, 34)  # AH = col 34
            _f = _cell.value
            if _f and isinstance(_f, str):
                _f_stripped = _f.strip()
                _m = _self_sum_re.match(_f_stripped)
                if _m and int(_m.group(1)) != _r:
                    _cell.value = f"=SUM(C{_r}:AG{_r})"
                    fixed_count += 1
                    logging.info(f"  Re-anchored AH{_r} self-total: {_f} -> =SUM(C{_r}:AG{_r})")
                elif _ratio_re.match(_f_stripped):
                    _cell.value = f"=SUM(C{_r}:AG{_r})"
                    fixed_count += 1
                    logging.info(f"  Replaced AH{_r} ratio formula with self-total: {_f} -> =SUM(C{_r}:AG{_r})")
        if fixed_count:
            logging.info(f"Fixed {fixed_count} AH formula(s) in rows {block_start}-{block_end - 1}")
        else:
            logging.info(f"No AH formula issues found in rows {block_start}-{block_end - 1}")

    # ── Re-point 'Weekly Distributions' cross-sheet links at the Collection
    # rows they actually mean, not whatever row they happened to sit at when
    # the template was last hand-aligned. Three rows there pull from Collection
    # via SUMIFS(Collection!$C${row}:$AG${row}, ...): 'Loan Recoveries' from
    # 'Debtor', 'Any Other cash Inflows' from 'Other Collections', and 'New
    # Deposits' from the Fixed Deposits section's totals row (nr2 + 1, the row
    # right after the newly appended Fixed Deposits data row). All three of
    # those Collection rows shift every week-1 run, but the Weekly
    # Distributions formulas are static text that never gets touched by
    # Excel's own row-insert logic (this is a different sheet), so they were
    # left pointing at stale rows. Re-derived by label every run instead of
    # carried forward, same as the fixes above. other_collections_row was
    # already resolved above for the Other Collections formula refresh.
    fd_totals_row = (nr2 + 1) if nr2 else None

    wd_sheet = None
    for _sn in wb.sheetnames:
        if _sn.strip().lower() == "weekly distributions":
            wd_sheet = wb[_sn]
            break

    if wd_sheet is None:
        logging.warning("'Weekly Distributions' sheet not found — skipping cross-sheet link refresh")
    else:
        # count=1: each formula's FIRST Collection!$C$x:$AG$x is the actual
        # value range being summed; the two that follow are always the fixed
        # Collection!$C$1:$AG$1 header/date-row criteria refs for the SUMIFS
        # ">=" / "<=" conditions, which must stay pointed at row 1.
        _coll_row_re = _re.compile(r"(Collection!\$C\$)\d+(:\$AG\$)\d+")

        def _repoint(row: int, target_row: int | None, target_label: str):
            if target_row is None:
                logging.warning(f"{target_label} row not found — skipping Weekly Distributions row {row} link refresh")
                return
            updated = 0
            for _col in (5, 7, 9, 11):  # E, G, I, K
                _cell = wd_sheet.cell(row, _col)
                _f = _cell.value
                if _f and isinstance(_f, str) and _f.startswith("="):
                    _new_f = _coll_row_re.sub(rf"\g<1>{target_row}\g<2>{target_row}", _f, count=1)
                    if _new_f != _f:
                        _cell.value = _new_f
                        updated += 1
            if updated:
                logging.info(f"Weekly Distributions row {row}: re-pointed {updated} cell(s) to Collection row {target_row} ({target_label})")

        _repoint(4, debtor_row, "Debtor")
        _repoint(5, fd_totals_row, "Fixed Deposits totals")
        _repoint(6, other_collections_row, "Other Collections")

    wb.save(report_file)
    logging.info("Collection sheet updated — C1:AG1 dates filled, three sections appended")


# ---------------------------------------------------------
# Step 17 – M2M Sheet Update from Master_Data.xlsx
# ---------------------------------------------------------
def update_m2m_sheet(report_file: str, input_date: datetime):
    """
    Read the last row of Master_Data.xlsx NBD_WF_18_ECS sheet (columns A-G = day1-day7).
    The 7 dates are Mon-Sun of the week ending on input_date.
    Append 7 rows to the M2M sheet of report_file: col A = date (MM/DD/YYYY), col B = value.
    Then write formulas in col C (=B[n]-B[n-1]) and col D (=MIN(C[n]/1000,0)) for the new rows.
    """
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    master_data_path = os.path.join(root, "Master_Data.xlsx")

    # Read M2M values from Master_Data.xlsx
    try:
        wb_md = load_workbook(master_data_path, data_only=True)
        ws_md = wb_md["NBD_WF_18_ECS"]
        last_row = ws_md.max_row
        if last_row < 2:
            logging.warning("Master_Data.xlsx NBD_WF_18_ECS sheet has no data rows — skipping M2M update")
            wb_md.close()
            return
        m2m_values = [ws_md.cell(row=last_row, column=c).value for c in range(1, 8)]

        # If user did not enter M2M values this run, carry forward from the previous row
        if all(v in (None, "") for v in m2m_values):
            logging.warning("Last row M2M values are empty — carrying forward values from previous row")
            if last_row >= 3:
                m2m_values = [ws_md.cell(row=last_row - 1, column=c).value for c in range(1, 8)]
                logging.info(f"Carried forward M2M values from row {last_row - 1}: {m2m_values}")
            else:
                logging.warning("No previous row to carry forward — using zeros")
                m2m_values = [None] * 7

            # Write the carried-forward (or zero) values to the empty last row so it is persisted
            for c, v in enumerate(m2m_values, start=1):
                ws_md.cell(row=last_row, column=c).value = v if v not in (None, "") else 0
            wb_md.save(master_data_path)
            logging.info(f"Persisted carried-forward M2M values to Master_Data row {last_row}")

        wb_md.close()
    except Exception as e:
        logging.error(f"Failed to read M2M data from Master_Data.xlsx: {e}")
        return

    # Replace any remaining None/"" with 0
    m2m_values = [v if v not in (None, "") else 0 for v in m2m_values]

    # 7 dates: Mon–Sun of the week ending on input_date
    week_start = input_date - timedelta(days=input_date.weekday())
    week_dates = [
        week_start + timedelta(days=i)
        for i in range(7)
    ]

    # Write to M2M sheet in report_file
    try:
        wb = load_workbook(report_file, keep_vba=True)
        ws_m2m = None
        for sname in wb.sheetnames:
            if sname == "M2M":
                ws_m2m = wb[sname]
                break
        if ws_m2m is None:
            logging.warning("M2M sheet not found in report file — skipping M2M update")
            wb.close()
            return

        # Find last row with actual data in column A (avoids ghost rows from formatting)
        last_data_row = 0
        for r_idx in range(1, ws_m2m.max_row + 1):
            if ws_m2m.cell(row=r_idx, column=1).value not in (None, ""):
                last_data_row = r_idx
        first_new_row = last_data_row + 1

        for i, (date_val, value) in enumerate(zip(week_dates, m2m_values)):
            r = first_new_row + i
            # Copy formatting from the row directly above so new rows match
            # existing ones (number format, font, borders, etc.) instead of
            # falling back to Excel's default "General" style.
            for col in (1, 2, 3, 4):
                src_cell = ws_m2m.cell(row=r - 1, column=col)
                dest_cell = ws_m2m.cell(row=r, column=col)
                if src_cell.has_style:
                    dest_cell.font = copy(src_cell.font)
                    dest_cell.border = copy(src_cell.border)
                    dest_cell.fill = copy(src_cell.fill)
                    dest_cell.number_format = src_cell.number_format
                    dest_cell.protection = copy(src_cell.protection)
                    dest_cell.alignment = copy(src_cell.alignment)

            ws_m2m.cell(row=r, column=1, value=date_val)
            ws_m2m.cell(row=r, column=2, value=value if value not in (None, "") else 0)
            # Column C: movement = B[r] - B[r-1]
            ws_m2m.cell(row=r, column=3, value=f"=B{r}-B{r-1}")
            # Column D: in 000' = MIN(C[r]/1000, 0)
            ws_m2m.cell(row=r, column=4, value=f"=MIN(C{r}/1000,0)")

        wb.save(report_file)
        wb.close()
        logging.info(f"M2M sheet updated with {len(week_dates)} rows starting at row {first_new_row}")
    except Exception as e:
        logging.error(f"Failed to update M2M sheet: {e}")

# ---------------------------------------------------------
# TB Bot runner for ECS
# ---------------------------------------------------------
def _read_tb_file_data(tb_file: Path):
    """Read TB file into a list of rows. Tries multiple engines for compatibility."""
    import pandas as _pd
    data = None
    for _engine in ("calamine", "xlrd", None):
        try:
            if _engine is None:
                wb_tb = load_workbook(str(tb_file), read_only=True, data_only=True)
                ws_tb = wb_tb.active
                data = [row for row in ws_tb.iter_rows(values_only=True)]
                wb_tb.close()
            else:
                df_tb = _pd.read_excel(str(tb_file), header=None, engine=_engine)
                data = [tuple(v if v == v else None for v in row) for row in df_tb.values.tolist()]
            logging.info(f"TB file read successfully with engine={_engine!r}")
            break
        except Exception as _e:
            logging.warning(f"TB read with engine={_engine!r} failed: {_e}")
    return data


def _paste_to_close_tb(close_tb_file: Path, report_file: str) -> None:
    """Paste data from the week-end TB file into the Close TB sheet."""
    logging.info(f"Close TB: reading from {close_tb_file.name}")
    data = _read_tb_file_data(close_tb_file)
    if data is None:
        logging.error("Close TB: all read methods failed — skipping")
        return

    wb = load_workbook(report_file, keep_vba=True)
    if "Close TB" not in wb.sheetnames:
        logging.warning("Close TB sheet not found in report file — skipping")
        wb.close()
        return

    ws = wb["Close TB"]
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)

    for r_idx, row in enumerate(data, start=1):
        for c_idx, value in enumerate(row, start=1):
            if value is not None:
                ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(report_file)
    logging.info(f"Close TB: pasted {len(data)} rows from {close_tb_file.name}")


def _download_single_tb(bot_class, enc_username: str, enc_password: str,
                        report_date: datetime, label: str) -> Path | None:
    """
    Download one TB report for report_date, rename it to {YYYY_MM_DD}_Report.xlsx,
    move it to TB_OUTPUT_FOLDER, and return the destination Path.
    Retries indefinitely until the file is downloaded.
    """
    report_date_str = report_date.strftime("%d/%m/%Y")
    date_prefix = report_date.strftime("%Y_%m_%d")
    dest_name = f"{date_prefix}_Report.xlsx"
    dest_path = TB_OUTPUT_FOLDER / dest_name

    download_patterns = ["*Report*.xlsx", "TB-COA*.xlsx", "*TB_COA*.xlsx"]

    # Remove all existing *_Report.xlsx files in the output folder before moving the new one
    for _old in TB_OUTPUT_FOLDER.glob("*_Report.xlsx"):
        try:
            _old.unlink()
            logging.info(f"Removed existing TB file: {_old.name}")
        except Exception as _e:
            logging.warning(f"Could not remove {_old.name}: {_e}")

    retry_count = 0
    while True:
        retry_count += 1
        logging.info(f"TB bot [{label}] attempt #{retry_count} for date {report_date_str}...")

        existing_files = {f.name for f in TB_DOWNLOADS_FOLDER.iterdir() if f.is_file()}

        try:
            bot = bot_class(download_folder=str(TB_DOWNLOADS_FOLDER))
            success = bot.run_full_automation(enc_username, enc_password, report_date_str)
            bot.close_browser()
        except Exception as e:
            logging.warning(f"TB bot [{label}] attempt #{retry_count} raised exception: {e}")
            time.sleep(5)
            continue

        if not success:
            logging.warning(f"TB bot [{label}] reported failure on attempt #{retry_count}. Retrying...")
            continue

        downloaded = [
            f for pattern in download_patterns
            for f in TB_DOWNLOADS_FOLDER.glob(pattern)
            if f.name not in existing_files
        ]
        if not downloaded:
            logging.warning(f"TB bot [{label}] succeeded but no new file found. Retrying...")
            continue

        new_file = max(downloaded, key=lambda f: f.stat().st_mtime)
        logging.info(f"TB [{label}] downloaded: {new_file.name}")

        shutil.move(str(new_file), str(dest_path))
        logging.info(f"TB [{label}] renamed and moved to: {dest_path}")
        return dest_path


def run_tb_bot_for_ecs(report_date: datetime) -> Path | None:
    """
    Download one TB report from ERP for report_date (week-end / report date).
    File is saved as {YYYY_MM_DD}_Report.xlsx in TB_OUTPUT_FOLDER.
    Returns the path, or None on failure.
    """
    logging.info("=" * 60)
    logging.info("Running TB from ERP Bot for ECS (report date)...")
    logging.info("=" * 60)

    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))

    try:
        from dotenv import load_dotenv
        env_path = script_dir.parent / ".env"
        load_dotenv(dotenv_path=env_path, override=True)
    except ImportError:
        pass

    import os as _os
    ENC_USERNAME = _os.getenv("ENC_USERNAME")
    ENC_PASSWORD = _os.getenv("ENC_PASSWORD")
    if not ENC_USERNAME or not ENC_PASSWORD:
        logging.error("ENC_USERNAME / ENC_PASSWORD not set in .env — cannot run TB bot")
        return None

    sys.path.insert(0, str(script_dir))
    try:
        from tb_from_erp_bot import ERPLoginBot
    except ImportError as e:
        logging.error(f"Failed to import tb_from_erp_bot: {e}")
        return None
    finally:
        sys.path.pop(0)

    TB_DOWNLOADS_FOLDER.mkdir(parents=True, exist_ok=True)
    TB_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # Remove all existing *_Report.xlsx files before downloading
    for _old in TB_OUTPUT_FOLDER.glob("*_Report.xlsx"):
        try:
            _old.unlink()
            logging.info(f"Removed existing TB file: {_old.name}")
        except Exception as _e:
            logging.warning(f"Could not remove {_old.name}: {_e}")

    return _download_single_tb(ERPLoginBot, ENC_USERNAME, ENC_PASSWORD, report_date, "report")


# ---------------------------------------------------------
# Pre-run bots
# ---------------------------------------------------------
def run_audit_receipt_bot(input_date: datetime, working_folder, outputs_ecs_dir: Path,
                          start_date_override: datetime = None, dest_filename: str = None) -> bool:
    """
    Run bots/audit_reciept_erp.py as a subprocess.
    The bot downloads 'Audit Receipt - Internal *.xlsx' and moves it to the
    latest ECS outputs date folder.
    If working_folder is not None, also copies the file into it.

    Args:
        start_date_override: If provided, passed as argv[2] to the bot so it uses
                             this as the range start instead of end-6 days.
        dest_filename:       If provided, the downloaded file in the ECS outputs
                             folder is renamed to this name after the bot finishes.
    Returns True on success.
    """
    logging.info("=" * 60)
    logging.info("Running Audit Receipt ERP Bot...")
    logging.info("=" * 60)

    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    bot_script = script_dir.parent / "bots" / "audit_reciept_erp.py"

    if not bot_script.exists():
        logging.error(f"Audit Receipt bot not found: {bot_script}")
        return False

    date_str = input_date.strftime("%d/%m/%Y")
    cmd = [sys.executable, str(bot_script), date_str]
    if start_date_override:
        cmd.append(start_date_override.strftime("%d/%m/%Y"))
        cmd.append("monthly")   # argv[3] — signals bot to replace monthly file

    result = subprocess.run(cmd, capture_output=False)

    if result.returncode != 0:
        logging.error(f"Audit Receipt bot exited with code {result.returncode}")
        return False

    # Rename downloaded file in ECS outputs folder if dest_filename provided
    if dest_filename:
        try:
            dates_file = outputs_ecs_dir / "completed_dates.txt"
            if dates_file.exists():
                _lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
                if _lines:
                    _last_date = _lines[-1]
                    _target = None
                    _base = outputs_ecs_dir / _last_date
                    if _base.exists():
                        _target = _base
                    _max_ver = 0
                    for _d in outputs_ecs_dir.iterdir():
                        if _d.is_dir() and _d.name.startswith(f"{_last_date}("):
                            try:
                                _ver = int(_d.name[len(_last_date)+1:-1])
                                if _ver > _max_ver:
                                    _max_ver = _ver
                                    _target = _d
                            except ValueError:
                                pass
                    if _target:
                        # Remove any existing file with dest_filename
                        _existing = _target / dest_filename
                        if _existing.exists():
                            _existing.unlink()
                            logging.info(f"Removed existing: {dest_filename}")
                        # Also remove old monthly files matching glob pattern
                        _stem = Path(dest_filename).stem
                        _glob_prefix = _stem.rsplit(" ", 1)[0]  # e.g. "Audit Receipt - Internal (monthly)"
                        for _old in _target.glob(f"{_glob_prefix} *.xlsx"):
                            _old.unlink(missing_ok=True)
                            logging.info(f"Removed old monthly file: {_old.name}")
                        # Find the newly downloaded file and rename it
                        for _f in sorted(_target.glob("Audit Receipt - Internal *.xlsx"),
                                         key=lambda x: x.stat().st_mtime, reverse=True):
                            if _f.name != dest_filename:
                                _f.rename(_target / dest_filename)
                                logging.info(f"Renamed {_f.name} - {dest_filename}")
                                break
        except Exception as e:
            logging.warning(f"Could not rename downloaded file to '{dest_filename}': {e}")

    # Copy to working folder if provided
    if working_folder is not None:
        try:
            dates_file = outputs_ecs_dir / "completed_dates.txt"
            if dates_file.exists():
                lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
                if lines:
                    last_date = lines[-1]
                    target_folder = None
                    base = outputs_ecs_dir / last_date
                    if base.exists():
                        target_folder = base
                    max_ver = 0
                    for d in outputs_ecs_dir.iterdir():
                        if d.is_dir() and d.name.startswith(f"{last_date}("):
                            try:
                                ver = int(d.name[len(last_date)+1:-1])
                                if ver > max_ver:
                                    max_ver = ver
                                    target_folder = d
                            except ValueError:
                                pass
                    if target_folder:
                        for f in target_folder.glob("Audit Receipt - Internal *.xlsx"):
                            dest = Path(working_folder) / f.name
                            shutil.copy2(str(f), str(dest))
                            logging.info(f"Copied Audit Receipt to working folder: {dest}")
        except Exception as e:
            logging.warning(f"Could not copy Audit Receipt to working folder: {e}")

    return True


def run_gl_codebook_bot(input_date: datetime) -> bool:
    """
    Run bots/GL_CodeBase_Report.py by importing its run() function.
    Returns the closing balance string (e.g. '3,059,942,299.21'), or None on failure.
    """
    logging.info("=" * 60)
    logging.info("Running GL CodeBase Report Bot...")
    logging.info("=" * 60)

    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    bots_dir = script_dir.parent / "bots"

    sys.path.insert(0, str(bots_dir))
    try:
        import importlib
        gl_mod = importlib.import_module("GL_CodeBase_Report")
        start_str = (input_date - timedelta(days=input_date.weekday())).strftime("%d/%m/%Y")
        end_str   = input_date.strftime("%d/%m/%Y")
        gl_mod.run(start_str, end_str, gl_code="1173", dest_outputs_name="NBD_WF_18_ECS")
        logging.info("GL CodeBase bot completed")
        return True
    except Exception as e:
        logging.error(f"GL CodeBase bot failed: {e}")
        return False
    finally:
        sys.path.pop(0)


def run_payments_bot(input_date: datetime, outputs_ecs_dir: Path) -> bool:
    """
    Run bots/payments.py as a subprocess to download the Payments report from
    the ERP for the reporting week (input_date - 6 days through input_date) —
    the same weekly window used for the Audit Receipt bot.

    Moves the downloaded file into the current ECS outputs date folder
    (resolved via completed_dates.txt, same lookup used elsewhere) under a
    fixed name, "Payments - DD-MM-YYYY.xlsx", replacing any existing one —
    this matches find_payments_file()'s starts_with="Payments -" pattern, so
    update_payment_sheet picks it up exactly like a manually uploaded file.
    Returns True on success.
    """
    logging.info("=" * 60)
    logging.info("Running Payments ERP Bot...")
    logging.info("=" * 60)

    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    bot_script = script_dir.parent / "bots" / "payments.py"

    if not bot_script.exists():
        logging.error(f"Payments bot not found: {bot_script}")
        return False

    week_start = input_date - timedelta(days=6)
    from_date_str = week_start.strftime("%d/%m/%Y")
    to_date_str = input_date.strftime("%d/%m/%Y")

    downloads_dir = Path.home() / "Downloads"
    existing_files = {f.name for f in downloads_dir.iterdir() if f.is_file()} if downloads_dir.exists() else set()

    cmd = [sys.executable, str(bot_script), from_date_str, to_date_str]
    result = subprocess.run(cmd, capture_output=False)

    if result.returncode != 0:
        logging.error(f"Payments bot exited with code {result.returncode}")
        return False

    downloaded = [
        f for f in downloads_dir.iterdir()
        if f.is_file() and f.name not in existing_files and f.suffix.lower() == ".xlsx"
    ]
    if not downloaded:
        logging.error("Payments bot completed but no new file found in Downloads")
        return False

    new_file = max(downloaded, key=lambda f: f.stat().st_mtime)
    logging.info(f"Payments bot downloaded: {new_file.name}")

    # Resolve the current ECS outputs date folder from completed_dates.txt —
    # same version-suffix-aware lookup used by run_audit_receipt_bot.
    dates_file = outputs_ecs_dir / "completed_dates.txt"
    if not dates_file.exists():
        logging.warning("completed_dates.txt not found — cannot place downloaded Payments file")
        return False
    lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
    if not lines:
        logging.warning("completed_dates.txt is empty — cannot place downloaded Payments file")
        return False
    last_date = lines[-1]
    target = None
    base = outputs_ecs_dir / last_date
    if base.exists():
        target = base
    max_ver = 0
    for d in outputs_ecs_dir.iterdir():
        if d.is_dir() and d.name.startswith(f"{last_date}("):
            try:
                ver = int(d.name[len(last_date) + 1:-1])
                if ver > max_ver:
                    max_ver = ver
                    target = d
            except ValueError:
                pass
    if not target:
        logging.warning(f"No outputs folder found for {last_date} — cannot place downloaded Payments file")
        return False

    dest_filename = f"Payments - {input_date.strftime('%d-%m-%Y')}.xlsx"
    dest_path = target / dest_filename
    if dest_path.exists():
        dest_path.unlink()
        logging.info(f"Removed existing: {dest_filename}")
    new_file.rename(dest_path)
    logging.info(f"Moved Payments file to: {dest_path}")

    return True


# ---------------------------------------------------------
# Main execution
# ---------------------------------------------------------
def main():
    total_start_time = time.perf_counter()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = setup_logger(script_dir)

    if len(sys.argv) < 2:
        sys.exit("Usage: python NBD_WF_18_ECS_optimized.py <date>")

    date_formats = [
        "%m/%d/%Y", "%Y-%m-%d",
        "%Y/%m/%d", "%m-%d-%Y",
    ]

    input_date = None
    for fmt in date_formats:
        try:
            input_date = datetime.strptime(sys.argv[1], fmt)
            logging.info(f"Date parsed using format: {fmt}")
            print(f"Date parsed: {input_date.strftime('%d/%m/%Y')}")
            break
        except ValueError:
            continue

    if input_date is None:
        sys.exit(f"Error: Could not parse date '{sys.argv[1]}'")

    # Week start = Monday of the week containing input_date (week-end date)
    week_start_date = input_date - timedelta(days=input_date.weekday())
    first_day_of_month = input_date.replace(day=1)
    week_number = ((input_date.day + first_day_of_month.weekday()) - 1) // 7 + 1

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    outputs_ecs_dir = Path(root) / "outputs" / "NBD_WF_18_ECS"

    # ----------------------------------------------------------------
    # Step 1: Run all bots BEFORE copying outputs - working
    # ----------------------------------------------------------------

    # Run Audit Receipt bot — downloads to ECS outputs folder (not working yet)
    # dest_filename gives this a fixed name immediately after download, so a
    # later monthly download (which can generate the same timestamp-based
    # filename if run within the same minute) can never silently overwrite it.
    audit_ok = run_audit_receipt_bot(
        input_date, None, outputs_ecs_dir,
        dest_filename="Audit Receipt - Internal (weekly) Current.xlsx"
    )
    if not audit_ok:
        logging.warning("Audit Receipt bot failed — receipt sheet may be skipped")

    # Week 1 only: also run Audit Receipt for the full previous month
    if week_number == 1:
        _prev_month_last = input_date.replace(day=1) - timedelta(days=1)  # last day of prev month
        _prev_month_first = _prev_month_last.replace(day=1)               # 1st day of prev month
        _month_label = _prev_month_last.strftime("%b")                    # e.g. "Jan"
        _monthly_dest = f"Audit Receipt - Internal (monthly) {_month_label}.xlsx"
        logging.info(f"Week 1 — running monthly Audit Receipt bot for {_prev_month_first.strftime('%d/%m/%Y')} to {_prev_month_last.strftime('%d/%m/%Y')}")
        _monthly_ok = run_audit_receipt_bot(
            _prev_month_last, None, outputs_ecs_dir,
            start_date_override=_prev_month_first,
            dest_filename=_monthly_dest,
        )
        if not _monthly_ok:
            logging.warning("Monthly Audit Receipt bot failed — monthly file may be missing")

    # Run GL CodeBase bot — retrieves closing balance from ERP
    gl_ok = run_gl_codebook_bot(input_date)
    if not gl_ok:
        logging.warning("GL CodeBase bot failed — file may not have been downloaded")

    # Run TB bot — downloads report date TB file to TB_OUTPUT_FOLDER
    tb_path = run_tb_bot_for_ecs(input_date)
    if not tb_path:
        logging.error("TB bot failed — cannot proceed without TB file")
        sys.exit(1)

    # Run Payments bot — downloads the reporting week's Payments report from ERP
    payments_ok = run_payments_bot(input_date, outputs_ecs_dir)
    if not payments_ok:
        logging.warning("Payments bot failed — Payment sheet will fall back to Audit Receipt data")

    # ----------------------------------------------------------------
    # Step 2: Copy latest ECS outputs date folder - working
    # ----------------------------------------------------------------
    working_path = os.path.join(root, "working", "NBD_WF_18_ECS")
    ecs_working_dir = Path(working_path)

    # Find latest completed date folder from completed_dates.txt
    latest_outputs_folder = None
    _dates_file = outputs_ecs_dir / "completed_dates.txt"
    if _dates_file.exists():
        _lines = [l.strip() for l in _dates_file.read_text().splitlines() if l.strip()]
        if _lines:
            _last_date = _lines[-1]
            import re as _re
            _candidates = []
            for _d in outputs_ecs_dir.iterdir():
                if _d.is_dir() and (_d.name == _last_date or _d.name.startswith(f"{_last_date}(")):
                    _m = _re.search(r'\((\d+)\)$', _d.name)
                    _candidates.append((int(_m.group(1)) if _m else 0, _d))
            if _candidates:
                _candidates.sort(key=lambda x: x[0], reverse=True)
                latest_outputs_folder = _candidates[0][1]
    if not latest_outputs_folder or not latest_outputs_folder.exists():
        sys.exit("No completed ECS outputs date folder found — cannot copy to working")

    # Tax Payments and Cashflow Movement files are preserved in outputs.
    # They are only replaced when the user uploads a new file via the UI.

    # Remove existing working folder contents (not the root, to avoid PermissionError
    # when the process cwd is inside it)
    if ecs_working_dir.exists():
        os.chdir(os.path.dirname(os.path.abspath(__file__)))  # move cwd away first
        for _item in list(ecs_working_dir.iterdir()):
            try:
                if _item.is_dir():
                    shutil.rmtree(_item, ignore_errors=True)
                else:
                    _item.unlink(missing_ok=True)
            except Exception as _e:
                logging.warning(f"Could not remove {_item.name}: {_e}")
        logging.info("Cleared existing ECS working folder")
    ecs_working_dir.mkdir(parents=True, exist_ok=True)

    target_working_folder = ecs_working_dir / latest_outputs_folder.name
    shutil.copytree(str(latest_outputs_folder), str(target_working_folder))
    logging.info(f"Copied ECS date folder from outputs to working: {latest_outputs_folder.name}")

    date_folder = str(target_working_folder)

    # ----------------------------------------------------------------
    # Step 3: Copy bot output files into working date folder
    # ----------------------------------------------------------------

    # Copy audit receipt from ECS outputs - working
    try:
        for f in latest_outputs_folder.glob("Audit Receipt - Internal *.xlsx"):
            dest = target_working_folder / f.name
            shutil.copy2(str(f), str(dest))
            logging.info(f"Copied Audit Receipt to working folder: {dest.name}")
    except Exception as e:
        logging.warning(f"Could not copy Audit Receipt to working folder: {e}")

    # Copy Types_of_Receipt.xlsx from ECS outputs - working
    try:
        tor_src = latest_outputs_folder / "Types_of_Receipt.xlsx"
        if tor_src.exists():
            shutil.copy2(str(tor_src), str(target_working_folder / "Types_of_Receipt.xlsx"))
            logging.info("Copied Types_of_Receipt.xlsx to working folder")
        else:
            logging.info("Types_of_Receipt.xlsx not found in outputs — skipping copy")
    except Exception as e:
        logging.warning(f"Could not copy Types_of_Receipt.xlsx to working folder: {e}")

    # Copy SasiaNet - *.xlsx from ECS outputs - working
    try:
        sn_files = list(latest_outputs_folder.glob("SasiaNet - *.xlsx"))
        if sn_files:
            sn_src = max(sn_files, key=lambda f: f.stat().st_mtime)
            shutil.copy2(str(sn_src), str(target_working_folder / sn_src.name))
            logging.info(f"Copied SasiaNet file to working folder: {sn_src.name}")
        else:
            logging.info("No SasiaNet file found in outputs — skipping copy")
    except Exception as e:
        logging.warning(f"Could not copy SasiaNet file to working folder: {e}")

    # Copy Tax Payments - *.xlsx → working folder
    # Priority: 1) Master Inputs (user-uploaded via UI)  2) ECS outputs (last run)
    try:
        master_inputs_tp = Path(root) / "Master Inputs" / "Tax Payments"
        mi_tp_files = list(master_inputs_tp.glob("Tax Payments - *.xlsx")) if master_inputs_tp.exists() else []
        out_tp_files = list(latest_outputs_folder.glob("Tax Payments - *.xlsx"))

        if mi_tp_files:
            tp_src = max(mi_tp_files, key=lambda f: f.stat().st_mtime)
            shutil.copy2(str(tp_src), str(target_working_folder / tp_src.name))
            logging.info(f"Copied Tax Payments from Master Inputs to working folder: {tp_src.name}")
        elif out_tp_files:
            tp_src = max(out_tp_files, key=lambda f: f.stat().st_mtime)
            shutil.copy2(str(tp_src), str(target_working_folder / tp_src.name))
            logging.info(f"Copied Tax Payments from outputs to working folder: {tp_src.name}")
        else:
            logging.info("No Tax Payments file found in Master Inputs or outputs — sheet will not be updated")
    except Exception as e:
        logging.warning(f"Could not copy Tax Payments file to working folder: {e}")

    # Copy Cashflow Movement - *.xlsx → working folder
    # Priority: 1) Master Inputs (user-uploaded via UI)  2) ECS outputs (last run)
    try:
        master_inputs_cm = Path(root) / "Master Inputs" / "Cashflow Movement"
        mi_cm_files = list(master_inputs_cm.glob("Cashflow Movement - *.xlsx")) if master_inputs_cm.exists() else []
        out_cm_files = list(latest_outputs_folder.glob("Cashflow Movement - *.xlsx"))

        if mi_cm_files:
            # User uploaded via UI — use Master Inputs file
            cm_src = max(mi_cm_files, key=lambda f: f.stat().st_mtime)
            shutil.copy2(str(cm_src), str(target_working_folder / cm_src.name))
            logging.info(f"Copied Cashflow Movement from Master Inputs to working folder: {cm_src.name}")
        elif out_cm_files:
            # No UI upload — fall back to last outputs copy
            cm_src = max(out_cm_files, key=lambda f: f.stat().st_mtime)
            shutil.copy2(str(cm_src), str(target_working_folder / cm_src.name))
            logging.info(f"Copied Cashflow Movement from outputs to working folder: {cm_src.name}")
        else:
            logging.info("No Cashflow Movement file found in Master Inputs or outputs — sheet will not be updated")
    except Exception as e:
        logging.warning(f"Could not copy Cashflow Movement file to working folder: {e}")

    report_file_for_tb = find_report_file(date_folder)

    # Step 3a: Copy existing Close TB - Opening TB (preserves last week's closing as this week's opening)
    if report_file_for_tb:
        with timed_operation("Copy Close TB to Opening TB (update_tb_sheets)"):
            update_tb_sheets(report_file_for_tb)

    # Step 3b: Paste new TB download into Close TB (this week's closing data)
    if tb_path and tb_path.exists():
        shutil.copy2(str(tb_path), str(target_working_folder / tb_path.name))
        logging.info(f"Copied TB to working folder: {tb_path.name}")
        if report_file_for_tb:
            with timed_operation("Paste TB data to Close TB sheet"):
                _paste_to_close_tb(tb_path, report_file_for_tb)
    else:
        logging.warning("TB file not found — skipping Close TB paste")

    # Step 3c: Update Types of Receipt and SasiaNet sheets from uploaded files (if present)
    report_file_for_tor = find_report_file(date_folder)
    if report_file_for_tor:
        types_of_receipt_file = find_types_of_receipt_file(date_folder)
        update_types_of_receipt_sheet(report_file_for_tor, types_of_receipt_file)

        sasianet_file = find_sasianet_file(date_folder)
        update_sasianet_sheet(report_file_for_tor, sasianet_file)

        tax_payments_file = find_tax_payments_file(date_folder)
        update_tax_payments_sheet(report_file_for_tor, tax_payments_file)

        cashflow_movement_file = find_cashflow_movement_file(date_folder)
        update_cashflow_movement_sheet(report_file_for_tor, cashflow_movement_file)

    # ----------------------------------------------------------------
    # Step 4: Run report
    # ----------------------------------------------------------------
    report_file = find_report_file(date_folder)
    daily_file = find_daily_bank_file(date_folder)

    # Main report updates
    with timed_operation("Main report header and account updates"):
        wb = load_workbook(report_file, keep_vba=True)
        ws_ecs = wb["NBD-WF-18-ECS"]
        ws_opening = wb["Opening&End LA"]

        # H2 already holds the week-start date from whichever run touched this
        # report_file last. If that matches this run's week_start_date, this is
        # a rerun for a week already processed — read it before overwriting so
        # process_accounts() can skip the E->C 'previous balance' copy (a
        # second copy would just duplicate this week's E into C, destroying
        # the real previous-week snapshot — see Opening&End LA notes).
        _prev_h2 = ws_ecs["H2"].value
        _same_week_rerun = (
            isinstance(_prev_h2, (datetime, date))
            and _prev_h2.year == week_start_date.year
            and _prev_h2.month == week_start_date.month
            and _prev_h2.day == week_start_date.day
        )
        if _same_week_rerun:
            logging.info(
                f"Opening&End LA: H2 already shows week starting {week_start_date:%d/%m/%Y} — "
                f"rerun for the same week detected, skipping E->C copy for Current/Savings blocks"
            )

        # H2: week start date as short date (DD/MM/YYYY)
        ws_ecs["H2"] = week_start_date
        ws_ecs["H2"].number_format = "DD/MM/YYYY"
        ws_ecs["I2"] = week_number
        ws_ecs["I12"] = None

        df_current, current_start, _ = process_accounts(ws_opening, "Current Accounts Balances", skip_copy=_same_week_rerun)
        df_savings, savings_start, _ = process_accounts(ws_opening, "Savings Accounts Balances", skip_copy=_same_week_rerun)

        df_daily, expected_total = process_daily_bank_balances(daily_file, input_date)
        sum_current = apply_vlookup(ws_opening, df_current, df_daily, current_start)
        sum_savings = apply_vlookup(ws_opening, df_savings, df_daily, savings_start)

        if round(sum_current + sum_savings, 2) != round(expected_total, 2):
            logging.warning("Validation failed: totals do not match")

        wb.save(report_file)
        logging.info("Saved main report updates")

    payments_file = os.path.join(date_folder, "Payments and Receipts.xlsx")

    process_weekly_distributions(report_file, payments_file, input_date)

    write_manual_payments_to_weekly_distributions(report_file)

    write_ucl_adjustments_from_cashflows(report_file)

    update_treasury_input_sheet(report_file, input_date)

    dm_file = find_dm_deposit_liability_file(date_folder)
    update_fd_sheet(report_file, dm_file)

    # Write DM sheet D14 computed value - ECS sheet C41
    if dm_file and os.path.exists(dm_file):
        with timed_operation("Read DM D14 for ECS C41"):
            _dm_d14 = None
            _xw_app = xw.App(visible=False)
            try:
                _wb_dm = _xw_app.books.open(dm_file)
                # Index-based scan instead of iterating _wb_dm.sheets — see note
                # above re: avoiding the COM enumerator (_NewEnum).
                _ws_dm = None
                _dm_sheet_count = com_retry(lambda: len(_wb_dm.sheets))
                for _i in range(_dm_sheet_count):
                    _sh = com_retry(lambda i=_i: _wb_dm.sheets[i])
                    if _sh.name.strip() == "NBD-WF-18-DM":
                        _ws_dm = _sh
                        break
                if _ws_dm is not None:
                    _dm_d14 = _ws_dm.range("D14").value
                    logging.info(f"DM D14 value: {_dm_d14}")
                else:
                    logging.warning("'NBD-WF-18-DM' sheet not found in DM file — C41 not updated")
                _wb_dm.close()
            finally:
                _xw_app.quit()

            if _dm_d14 is not None:
                _wb_ecs = load_workbook(report_file, keep_vba=True)
                _ws_ecs = _wb_ecs["NBD-WF-18-ECS"]
                _ws_ecs["C41"] = _dm_d14
                _wb_ecs.save(report_file)
                _wb_ecs.close()
                logging.info(f"Written DM D14 - ECS C41: {_dm_d14}")
    else:
        logging.warning("DM file not found — ECS C41 not updated")

    afl_file = find_afl_borrowing_schedule_file(date_folder)
    update_utilized_credit_lines(report_file, afl_file, input_date)

    audit_receipt_file = find_audit_receipt_file(date_folder)
    update_receipt_sheet(report_file, audit_receipt_file)

    payments_file = find_payments_file(date_folder)
    update_payment_sheet(report_file, payments_file, audit_receipt_file)

    update_m2m_sheet(report_file, input_date)

    # Read I17 computed value via xlwings; if negative, add its absolute value to I12
    with timed_operation("Check I17 and adjust I12 in NBD-WF-18-ECS"):
        _i17_val = None
        _xw_app3 = xw.App(visible=False)
        try:
            _wb3 = _xw_app3.books.open(report_file)
            time.sleep(2)  # Allow Excel COM to fully initialise before accessing sheets
            _ws3 = _wb3.sheets["NBD-WF-18-ECS"]
            _i17_val = _ws3.range("I17").value
            _wb3.close()
        finally:
            _xw_app3.quit()

        logging.info(f"NBD-WF-18-ECS I17 computed value: {_i17_val}")

        if _i17_val is not None and isinstance(_i17_val, (int, float)) and _i17_val < 0:
            _adjustment = abs(_i17_val)
            _wb_adj = load_workbook(report_file, keep_vba=True)
            _ws_adj = _wb_adj["NBD-WF-18-ECS"]
            _current_i12 = _ws_adj["I12"].value
            _current_i12 = _current_i12 if isinstance(_current_i12, (int, float)) else 0
            _ws_adj["I12"] = _current_i12 + _adjustment
            _wb_adj.save(report_file)
            _wb_adj.close()
            logging.info(f"I17={_i17_val} is negative — added {_adjustment} to I12 (was {_current_i12}, now {_current_i12 + _adjustment})")
        else:
            logging.info("I17 is not negative — I12 not adjusted")

    # Week 1 only: build monthly cum collection pivot from last month's audit receipt
    if week_number == 1:
        update_cum_monthly_collection(report_file, date_folder, input_date)

    # Every week: fill Collection sheet C1:AG1 with week dates (+ week 1 section appends)
    update_collection_sheet(report_file, input_date, week_number)

    # Clear cache
    sheet_cache.clear()

    total_elapsed = time.perf_counter() - total_start_time
    logging.info(f"TOTAL EXECUTION TIME: {total_elapsed:.2f}s")
    print(f"\n{'='*80}")
    print(f"TOTAL EXECUTION TIME: {total_elapsed:.2f}s")
    print(f"{'='*80}")
    print(f"Report updated successfully. Log: {log_file}")

if __name__ == "__main__":
    main()