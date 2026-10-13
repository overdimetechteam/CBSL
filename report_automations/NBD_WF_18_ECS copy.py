import sys
import os
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

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
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
    return find_file_by_pattern(working_path, contains="AFL Borrowing schedule")

def find_audit_receipt_file(working_path):
    return find_file_by_pattern(working_path, contains="Audit Receipt")

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
def process_accounts(ws, start_label):
    """Extract account data using bulk operations."""
    start_row = find_row_by_label(ws, start_label)
    if not start_row:
        raise ValueError(f"Label '{start_label}' not found")

    # Find total row
    total_row = None
    for r in range(start_row + 1, ws.max_row + 1):
        cell_val = ws.cell(r, 2).value
        if cell_val and str(cell_val).strip() == "Total":
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

    # Copy column E to column C
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

    mappings = {"B": "G", "C": "I", "D": "K", "E": "M"}
    target_map = {1: ("B", "B1"), 2: ("C", "C1"), 3: ("D", "D1"), 4: ("E", "E1")}

    # Single xlwings session for all operations requiring formula values
    app = xw.App(visible=False)
    try:
        with timed_operation("xlwings batch read"):
            payments_values = {}
            source_vals = []

            # Read from payments file if week 1
            if wom == 1:
                wb_pay_xw = app.books.open(payments_file)
                ws_pay_xw = wb_pay_xw.sheets[0]

                # Batch read all payment values at once
                for src_col in mappings.keys():
                    # Read range in one call instead of cell by cell
                    range_vals = ws_pay_xw.range(f"{src_col}2:{src_col}6").value
                    if not isinstance(range_vals, list):
                        range_vals = [range_vals]
                    payments_values[src_col] = range_vals
                    logging.info(f"Read Payments {src_col}2:{src_col}6: {range_vals}")

                wb_pay_xw.close()

            # Read calculated values from Weekly Distributions
            wb_report_xw = app.books.open(report_file)
            time.sleep(2)  # Allow Excel COM to fully initialise before accessing sheets
            # Try exact name first (has trailing space), fall back to index 2
            try:
                ws_weekly_xw = wb_report_xw.sheets["Weekly Distributions "]
                if ws_weekly_xw is None:
                    raise KeyError("sheet returned None")
            except (KeyError, Exception):
                ws_weekly_xw = wb_report_xw.sheets[2]
                logging.info(f"Accessed Weekly Distributions by index 2 (name: {ws_weekly_xw.name!r})")

            # Batch read G18:G22 in one call
            source_vals = ws_weekly_xw.range("G18:G22").value
            if not isinstance(source_vals, list):
                source_vals = [source_vals]
            logging.info(f"Read Weekly Distributions G18:G22: {source_vals}")

            wb_report_xw.close()
    finally:
        app.quit()

    # PART A: First week only - write to Weekly Distributions
    if wom == 1:
        with timed_operation("Write Weekly Distributions (Week 1)"):
            wb_report_write = load_workbook(report_file, keep_vba=True)
            ws_weekly_write = get_sheet(wb_report_write, "Weekly Distributions")

            for src_col, tgt_col in mappings.items():
                # Clear and write in bulk
                for i, val in enumerate(payments_values[src_col]):
                    ws_weekly_write[f"{tgt_col}{50+i}"].value = val
                    logging.info(f"Wrote {tgt_col}{50+i} = {val}")

            wb_report_write.save(report_file)
            logging.info("Saved Weekly Distributions updates")

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
# Step 04 – Copy Cash Flow to Weekly Distributions - OPTIMIZED
# ---------------------------------------------------------
@log_time
def copy_cash_flow_to_weekly_distributions(cash_flow_file, report_file):
    """Copy values using DataFrame operations."""
    logging.info("Processing Cash Flow to Weekly Distributions mapping")

    mappings = {
        "APIT": "APIIT",
        "Central Bank": "CBSL",
        "Taxes": "Taxes",
        "Outsource staff": "Outsource Staff"
    }

    # Read cash flow file efficiently
    wb_cash_flow = load_workbook(cash_flow_file, data_only=True, read_only=True)
    ws_cash_flow = wb_cash_flow.active

    # Bulk read into list
    cash_flow_values = {}
    for row in ws_cash_flow.iter_rows(min_col=1, max_col=2, values_only=True):
        if row[0] and str(row[0]).strip() in mappings.keys():
            label = str(row[0]).strip()
            cash_flow_values[label] = row[1]
            logging.info(f"Found '{label}': {row[1]}")

    wb_cash_flow.close()

    # Update report file
    wb_report = load_workbook(report_file, keep_vba=True)
    ws_weekly = get_sheet(wb_report, "Weekly Distributions")

    # Build lookup for target labels
    reverse_mappings = {v: k for k, v in mappings.items()}
    updates_made = 0

    for row in ws_weekly.iter_rows(min_col=2, max_col=3):
        cell_b = row[0]
        cell_c = row[1]

        if cell_b.value and str(cell_b.value).strip() in reverse_mappings:
            target_label = str(cell_b.value).strip()
            source_label = reverse_mappings[target_label]

            if source_label in cash_flow_values:
                cell_c.value = cash_flow_values[source_label]
                logging.info(f"Updated '{target_label}': {cash_flow_values[source_label]}")
                updates_made += 1

    wb_report.save(report_file)
    logging.info(f"Cash Flow mapping complete: {updates_made} cells updated")

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

            if "Treasury Input" not in [sheet.name for sheet in wb_xw.sheets]:
                logging.warning("Treasury Input sheet not found")
                wb_xw.close()
                return None

            ws_xw = wb_xw.sheets["Treasury Input"]

            # Batch read column B to find target fields
            b_col_data = ws_xw.range("B1:B100").value
            for row_num, cell_val in enumerate(b_col_data, start=1):
                if cell_val and str(cell_val).strip() in target_fields:
                    field_name = str(cell_val).strip()
                    value_from_d = ws_xw.range(f"D{row_num}").value
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
        app = xw.App(visible=False)
        try:
            wb_xw = app.books.open(str(report_file))
            close_tb_xw = wb_xw.sheets[close_tb_name]
            opening_tb_xw = wb_xw.sheets[opening_tb_name]
            opening_tb_xw.used_range.api.UnMerge()
            close_tb_xw.used_range.copy(opening_tb_xw.range("A1"))
            wb_xw.save()
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
    """Update credit lines using DataFrame operations."""
    logging.info("Processing utilized credit lines sheet update")

    week_number = input_date.isocalendar()[1]

    wb_report = load_workbook(report_file, keep_vba=True)

    if "utilized credit lines" not in wb_report.sheetnames:
        logging.warning("utilized credit lines sheet not found - skipping")
        wb_report.close()
        return

    ws_ucl = wb_report["utilized credit lines"]

    with timed_operation("Update left table (STL, TL, IC, OD)"):
        target_items = {"STL", "TL", "IC", "OD"}
        items_updated = 0

        for row in ws_ucl.iter_rows(min_row=1, max_row=ws_ucl.max_row, min_col=1, max_col=3):
            cell_a = row[0]
            if cell_a.value and str(cell_a.value).strip() in target_items:
                cell_b = row[1]
                cell_c = row[2]
                cell_b.value = cell_c.value
                items_updated += 1

        # Update week headers
        c1_value = ws_ucl.cell(1, 3).value
        ws_ucl.cell(1, 2).value = c1_value
        ws_ucl.cell(1, 3).value = f"Week {week_number}"
        logging.info(f"Updated {items_updated} items, week headers updated")

    with timed_operation("Find and process sections"):
        # Find Institute header
        institute_col = None
        header_row = None

        for row_num in range(1, min(50, ws_ucl.max_row + 1)):
            for col_num in range(1, min(20, ws_ucl.max_column + 1)):
                cell_val = ws_ucl.cell(row_num, col_num).value
                if cell_val and str(cell_val).strip() == "Institute":
                    institute_col = col_num
                    header_row = row_num
                    break
            if institute_col:
                break

        if not institute_col:
            logging.warning("Could not find 'Institute' header")
            wb_report.save(report_file)
            return

        # Cache UCL headers
        ucl_headers = {}
        for col_num in range(institute_col, ws_ucl.max_column + 1):
            header_val = ws_ucl.cell(header_row, col_num).value
            if header_val:
                ucl_headers[str(header_val).strip()] = col_num

        # Find sections
        sections = {
            "Short Term": {"start": None, "end": None},
            "Over Draft Facility": {"start": None, "end": None},
            "Long Term": {"start": None, "end": None}
        }

        # Log all non-empty Institute column values for diagnosis
        for row_num in range(header_row + 1, ws_ucl.max_row + 1):
            _val = ws_ucl.cell(row_num, institute_col).value
            if _val not in (None, ""):
                logging.info(f"UCL Institute col row {row_num}: {_val!r}")

        current_section = None
        for row_num in range(header_row + 1, ws_ucl.max_row + 1):
            institute_val = ws_ucl.cell(row_num, institute_col).value
            institute_str = str(institute_val).strip() if institute_val else ""

            if institute_str in sections:
                # Close previous section when a new section label is found
                if current_section and sections[current_section]["start"] and sections[current_section]["end"] is None:
                    sections[current_section]["end"] = row_num - 1
                current_section = institute_str
                sections[current_section]["start"] = row_num + 1
            elif not institute_str and current_section and sections[current_section]["start"] and sections[current_section]["end"] is None:
                # Close section on empty row
                sections[current_section]["end"] = row_num - 1
                current_section = None

        # Close last section if it runs to end of sheet with no trailing empty row
        if current_section and sections[current_section]["start"] and sections[current_section]["end"] is None:
            sections[current_section]["end"] = ws_ucl.max_row

        # Clear sections
        for section_name, bounds in sections.items():
            if bounds["start"] and bounds["end"]:
                for row_num in range(bounds["start"], bounds["end"] + 1):
                    for col_num in range(institute_col, ws_ucl.max_column + 1):
                        ws_ucl.cell(row_num, col_num).value = None
                logging.info(f"Cleared section '{section_name}'")

    # Load AFL data if available
    if not afl_file:
        logging.warning("AFL Borrowing Schedule file not found")
        wb_report.save(report_file)
        return

    with timed_operation("Load and copy AFL data"):
        wb_afl = load_workbook(afl_file, data_only=True, read_only=True)
        ws_afl = wb_afl["Funding Arrangement"] if "Funding Arrangement" in wb_afl.sheetnames else wb_afl.active

        # Find AFL headers
        afl_institute_col = None
        afl_header_row = None

        for row_num in range(1, min(20, ws_afl.max_row + 1)):
            for col_num in range(1, min(20, ws_afl.max_column + 1)):
                cell_val = ws_afl.cell(row_num, col_num).value
                if cell_val and str(cell_val).strip() == "Institute":
                    afl_institute_col = col_num
                    afl_header_row = row_num
                    break
            if afl_institute_col:
                break

        if not afl_institute_col:
            wb_afl.close()
            wb_report.save(report_file)
            return

        # Cache AFL headers
        afl_headers = {}
        for col_num in range(afl_institute_col, ws_afl.max_column + 1):
            header_val = ws_afl.cell(afl_header_row, col_num).value
            if header_val:
                afl_headers[str(header_val).strip()] = col_num

        matching_headers = set(ucl_headers.keys()) & set(afl_headers.keys())

        # Read AFL data into sections
        afl_sections = {
            "Short Term": [],
            "Over Draft Facility": [],
            "Long Term": []
        }

        current_section = None
        for row in ws_afl.iter_rows(min_row=afl_header_row + 1, values_only=True):
            if len(row) < afl_institute_col:
                continue
            institute_val = row[afl_institute_col - 1]
            institute_str = str(institute_val).strip() if institute_val else ""

            if institute_str in afl_sections:
                # New section label — switch section (don't reset on empty rows)
                current_section = institute_str
            elif institute_str and current_section:
                # Non-empty, non-label row — data row belonging to current section
                row_data = {}
                for header in matching_headers:
                    col_idx = afl_headers[header] - 1
                    if col_idx < len(row):
                        row_data[header] = row[col_idx]
                if row_data:
                    afl_sections[current_section].append(row_data)

        wb_afl.close()

        # Write AFL data to UCL
        for section_name, section_data in afl_sections.items():
            if section_data and sections[section_name]["start"]:
                start_row = sections[section_name]["start"]
                for idx, row_data in enumerate(section_data):
                    target_row = start_row + idx
                    for header, value in row_data.items():
                        if header in ucl_headers:
                            ws_ucl.cell(target_row, ucl_headers[header]).value = value
                logging.info(f"Copied {len(section_data)} rows to '{section_name}'")

    wb_report.save(report_file)
    logging.info("utilized_credit_lines update complete")

# ---------------------------------------------------------
# Step 09 – Update FD Sheet - OPTIMIZED with DataFrame
# ---------------------------------------------------------
@log_time
def update_fd_sheet(report_file, dm_file):
    """Update FD sheet using DataFrame operations."""
    logging.info("Processing FD sheet update")

    if not dm_file:
        logging.warning("DM Deposit Liability file not found - skipping")
        return

    with timed_operation("Read DM Portfolio data"):
        wb_dm = load_workbook(dm_file, data_only=True, read_only=True)
        ws_portfolio = wb_dm["Portfolio"]

        # Cache headers
        # Use enumerate for read_only mode compatibility
        headers = {}
        for col_idx, cell in enumerate(ws_portfolio[4], start=1):
            if cell.value:
                headers[cell.value] = col_idx

        required_headers = ["System Deposit No:", "Maturity Date", "Total Payable"]
        missing_headers = [h for h in required_headers if h not in headers]

        if missing_headers:
            logging.warning(f"Missing headers: {missing_headers}")
            wb_dm.close()
            return

        # Bulk read into DataFrame
        col_indices = [headers[h] - 1 for h in required_headers]
        data = []
        for row in ws_portfolio.iter_rows(min_row=5, values_only=True):
            if len(row) > max(col_indices):
                data.append([row[i] for i in col_indices])

        wb_dm.close()

    df = pd.DataFrame(data, columns=required_headers)
    df = df.dropna(how='all')
    logging.info(f"Read {len(df)} rows from Portfolio sheet")

    # Sort by maturity date
    df["Maturity Date"] = pd.to_datetime(df["Maturity Date"], errors="coerce")
    df = df.sort_values("Maturity Date", na_position='last')
    df["Total Payable"] = pd.to_numeric(df["Total Payable"], errors="coerce")

    with timed_operation("Write to FD sheet"):
        wb_report = load_workbook(report_file, keep_vba=True)

        if "FD" not in wb_report.sheetnames:
            logging.warning("FD sheet not found - skipping")
            wb_report.close()
            return

        ws_fd = wb_report["FD"]

        # Clear existing data
        if ws_fd.max_row >= 3:
            ws_fd.delete_rows(3, ws_fd.max_row - 2)

        # Bulk write data
        moved_count = 0
        for idx, row in enumerate(df.itertuples(index=False), start=3):
            ws_fd.cell(idx, 1).value = row[0]  # System Deposit No
            ws_fd.cell(idx, 2).value = row[1]  # Maturity Date

            total_payable = row[2]
            if isinstance(total_payable, (int, float)) and total_payable > LARGE_VALUE_THRESHOLD:
                ws_fd.cell(idx, 4).value = total_payable
                moved_count += 1
            else:
                ws_fd.cell(idx, 3).value = total_payable

        wb_report.save(report_file)
        logging.info(f"FD update complete: {len(df)} rows, {moved_count} moved to column D")

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

        # Get formula from column W
        first_w_formula = None
        for row_num in range(3, min(10, ws_receipt.max_row + 1)):
            cell = ws_receipt.cell(row_num, 23)
            if cell.value:
                first_w_formula = cell.value
                break

        # Clear existing data (except column W and rows 1-2)
        max_col = ws_receipt.max_column
        for row in ws_receipt.iter_rows(min_row=3, max_row=ws_receipt.max_row, max_col=max_col):
            for col_idx, cell in enumerate(row, start=1):
                if col_idx != 23:
                    cell.value = None

        # Bulk write audit data
        for r_idx, row_data in enumerate(audit_data, start=3):
            for c_idx, value in enumerate(row_data, start=1):
                if c_idx != 23:
                    ws_receipt.cell(r_idx, c_idx).value = value
            # Apply formula to column W
            if first_w_formula:
                ws_receipt.cell(r_idx, 23).value = first_w_formula

        wb_report.save(report_file)
        logging.info(f"Receipt update complete: {len(audit_data)} rows")

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
        wb_md.close()
    except Exception as e:
        logging.error(f"Failed to read M2M data from Master_Data.xlsx: {e}")
        return

    if all(v in (None, "") for v in m2m_values):
        logging.warning("All M2M values in Master_Data.xlsx are empty — skipping M2M update")
        return

    # 7 dates: Mon–Sun of the week ending on input_date
    week_start = input_date - timedelta(days=input_date.weekday())
    week_dates = [
        f"{(week_start + timedelta(days=i)).day}-{(week_start + timedelta(days=i)).strftime('%b')}"
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

        for i, (date_str, value) in enumerate(zip(week_dates, m2m_values)):
            r = first_new_row + i
            ws_m2m.cell(row=r, column=1, value=date_str)
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

    # Remove any previously downloaded file with the same dest name
    if dest_path.exists():
        dest_path.unlink()
        logging.info(f"Removed existing {dest_name}")

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
def run_audit_receipt_bot(input_date: datetime, working_folder, outputs_ecs_dir: Path) -> bool:
    """
    Run bots/audit_reciept_erp.py as a subprocess.
    The bot downloads 'Audit Receipt - Internal *.xlsx' and moves it to the
    latest ECS outputs date folder.
    If working_folder is not None, also copies the file into it.
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
    result = subprocess.run(
        [sys.executable, str(bot_script), date_str],
        capture_output=False,
    )

    if result.returncode != 0:
        logging.error(f"Audit Receipt bot exited with code {result.returncode}")
        return False

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


def run_gl_codebook_bot(input_date: datetime) -> str | None:
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
        date_str = input_date.strftime("%d/%m/%Y")
        closing_balance = gl_mod.run(date_str)
        if closing_balance:
            logging.info(f"GL CodeBase closing balance: {closing_balance}")
        else:
            logging.warning("GL CodeBase bot returned no closing balance")
        return closing_balance
    except Exception as e:
        logging.error(f"GL CodeBase bot failed: {e}")
        return None
    finally:
        sys.path.pop(0)


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
    # Step 1: Run all bots BEFORE copying outputs → working
    # ----------------------------------------------------------------

    # Run Audit Receipt bot — downloads to ECS outputs folder (not working yet)
    audit_ok = run_audit_receipt_bot(input_date, None, outputs_ecs_dir)
    if not audit_ok:
        logging.warning("Audit Receipt bot failed — receipt sheet may be skipped")

    # Run GL CodeBase bot — retrieves closing balance from ERP
    gl_closing_balance = run_gl_codebook_bot(input_date)
    if gl_closing_balance:
        logging.info(f"GL closing balance retrieved: {gl_closing_balance}")
    else:
        logging.warning("GL CodeBase bot failed — closing balance not available")

    # Run TB bot — downloads report date TB file to TB_OUTPUT_FOLDER
    tb_path = run_tb_bot_for_ecs(input_date)
    if not tb_path:
        logging.error("TB bot failed — cannot proceed without TB file")
        sys.exit(1)

    # ----------------------------------------------------------------
    # Step 2: Copy latest ECS outputs date folder → working
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

    # Copy audit receipt from ECS outputs → working
    try:
        for f in latest_outputs_folder.glob("Audit Receipt - Internal *.xlsx"):
            dest = target_working_folder / f.name
            shutil.copy2(str(f), str(dest))
            logging.info(f"Copied Audit Receipt to working folder: {dest.name}")
    except Exception as e:
        logging.warning(f"Could not copy Audit Receipt to working folder: {e}")

    report_file_for_tb = find_report_file(date_folder)

    # Step 3a: Copy existing Close TB → Opening TB (preserves last week's closing as this week's opening)
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

    # ----------------------------------------------------------------
    # Step 4: Run report
    # ----------------------------------------------------------------
    report_file = find_report_file(date_folder)
    daily_file = find_daily_bank_file(date_folder)
    cash_flow_file = find_cash_flow_file(date_folder)

    # Main report updates
    with timed_operation("Main report header and account updates"):
        wb = load_workbook(report_file, keep_vba=True)
        ws_ecs = wb["NBD-WF-18-ECS"]
        ws_opening = wb["Opening&End LA"]

        # H2: week start date (Monday) in DD/MM/YYYY, no leading zeros
        ws_ecs["H2"] = f"{week_start_date.month}/{week_start_date.day}/{week_start_date.year}"
        ws_ecs["I2"] = week_number
        ws_ecs["I12"] = None

        df_current, current_start, _ = process_accounts(ws_opening, "Current Accounts Balances")
        df_savings, savings_start, _ = process_accounts(ws_opening, "Savings Accounts Balances")

        df_daily, expected_total = process_daily_bank_balances(daily_file, input_date)
        sum_current = apply_vlookup(ws_opening, df_current, df_daily, current_start)
        sum_savings = apply_vlookup(ws_opening, df_savings, df_daily, savings_start)

        if round(sum_current + sum_savings, 2) != round(expected_total, 2):
            logging.warning("Validation failed: totals do not match")

        wb.save(report_file)
        logging.info("Saved main report updates")

    payments_file = os.path.join(date_folder, "Payments and Receipts.xlsx")

    process_weekly_distributions(report_file, payments_file, input_date)

    if cash_flow_file:
        copy_cash_flow_to_weekly_distributions(cash_flow_file, report_file)
    else:
        logging.warning("Cash Flow file not found - skipping")

    new_treasury_date = update_treasury_input_sheet(report_file, input_date)

    close_tb_date = None  # TB sheets handled in Step 3 (before paste)

    dm_file = find_dm_deposit_liability_file(date_folder)
    update_fd_sheet(report_file, dm_file)

    afl_file = find_afl_borrowing_schedule_file(date_folder)
    update_utilized_credit_lines(report_file, afl_file, input_date)

    audit_receipt_file = find_audit_receipt_file(date_folder)
    update_receipt_sheet(report_file, audit_receipt_file)

    # Manual reminders
    log_manual_audit_receipt_reminder()
    log_manual_sasianet_reminder(input_date)

    if new_treasury_date:
        log_manual_treasury_input_reminder(new_treasury_date)

    if close_tb_date:
        log_manual_close_tb_reminder(close_tb_date)

    update_m2m_sheet(report_file, input_date)

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
