import os
import glob
import subprocess
import sys
import pandas as pd
import xlwings as xw
from datetime import datetime, datetime as _datetime

# Force UTF-8 output on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


class _TeeWriter:
    def __init__(self, original, log_file_path):
        self._orig = original
        self._fh = open(log_file_path, "a", encoding="utf-8", errors="replace")
    def write(self, data):
        self._orig.write(data)
        self._fh.write(data)
        self._fh.flush()
    def flush(self):
        self._orig.flush()
        self._fh.flush()
    def fileno(self):
        return self._orig.fileno()


def _setup_log_tee():
    log_dir = os.environ.get("IFRS_LGD_LOG_DIR", "").strip()
    if not log_dir:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "logs", f"run_{ts}"
        )
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "LGD_computation.log")
    sys.stdout = _TeeWriter(sys.stdout, log_file)
    sys.stderr = _TeeWriter(sys.stderr, log_file)
    print(f"Log file: {log_file}")


_setup_log_tee()

from BasicExcelFunctionsClass import ExcelPortfolioAutomation

# ==========================
# PATH SETUP
# ==========================
CURRENT_DIR  = os.path.dirname(os.path.abspath(__file__))
_SCRIPT_ROOT = os.path.dirname(CURRENT_DIR)


def _find_lgd_working_folder() -> str:
    """Scan working/LGD/ and return <latest-date-folder>/LGD/."""
    base = os.path.join(_SCRIPT_ROOT, "working", "LGD")
    if os.path.isdir(base):
        subfolders = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d))]
        if subfolders:
            subfolders.sort(key=lambda d: os.path.getmtime(os.path.join(base, d)), reverse=True)
            return os.path.join(base, subfolders[0], "LGD")
    return base


_lgd_base   = os.environ.get("LGD_BASE_FOLDER", "").strip()
INPUTS_DIR  = os.path.join(_lgd_base, "LGD") if _lgd_base else _find_lgd_working_folder()
ROOT_DIR    = INPUTS_DIR
OUTPUTS_DIR = os.path.join(CURRENT_DIR, "output")
os.makedirs(OUTPUTS_DIR, exist_ok=True)

# ==========================
# FIND INPUT FILE
# ==========================
def find_input_file() -> str:
    for ext in ("xlsx", "xlsb"):
        matches = glob.glob(os.path.join(INPUTS_DIR, f"First Default Date Working *.{ext}"))
        if matches:
            return matches[0]
    raise FileNotFoundError(
        f"No 'First Default Date Working' file (.xlsx or .xlsb) found in {INPUTS_DIR}"
    )

def find_Receipt_file() -> str:
    matches = glob.glob(os.path.join(INPUTS_DIR, "JOBHD*.xlsx"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'JOBHD*.xlsx' found in {INPUTS_DIR}"
        )
    return matches[0]

def find_writebacks_file() -> str:
    matches = glob.glob(os.path.join(INPUTS_DIR, "Ledger Details Report*.xlsx"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'Ledger Details Report*.xlsx' found in {INPUTS_DIR}"
        )
    return matches[0]

def find_writebacks130_file() -> str:
    matches = glob.glob(os.path.join(INPUTS_DIR, "Write Off Working*.xlsx"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'Write Off Working*.xlsx' found in {INPUTS_DIR}"
        )
    return matches[0]

def find_write_backs_file() -> str:
    matches = glob.glob(os.path.join(INPUTS_DIR, "Write Backs 2024-25*.xlsx"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'Write Backs 2024-25*.xlsx' found in {INPUTS_DIR}"
        )
    return matches[0]

def find_Yard_file() -> str:
    matches = glob.glob(os.path.join(INPUTS_DIR, "1. Portfolio_4_ECL*.xlsb"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching '1. Portfolio_4_ECL*.xlsb' found in {INPUTS_DIR}"
        )
    return matches[0]

def find_LGD_file() -> str:
    matches = glob.glob(os.path.join(INPUTS_DIR, "LGD *.xlsb"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'LGD *.xlsb' found in {INPUTS_DIR}"
        )
    return matches[0]

def find_dc_list_file() -> str:
    matches = glob.glob(os.path.join(INPUTS_DIR, "DC List*.xlsx"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'DC List*.xlsx' found in {INPUTS_DIR}"
        )
    return matches[0]

def find_receipt_discount_file() -> str:
    # "Receipt*Discount*" tolerates either word being singular or plural
    # (e.g. "Receipt Discount", "Receipts Discount") since the source file's
    # naming has been inconsistent about it.
    matches = glob.glob(os.path.join(INPUTS_DIR, "Receipt*Discount*.xlsb"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'Receipt*Discount*.xlsb' found in {INPUTS_DIR}"
        )
    return matches[0]
# ==========================
# LOAD CLOSED LIST SHEET
# ==========================
def load_dataframes(file_path: str) -> pd.DataFrame:
    print(f"Reading 'Closed List' sheet from: {file_path}")
    automation = ExcelPortfolioAutomation(workbook_path=file_path, visible=False)
    df, num_rows, num_cols = automation.read_excel_data(
        excel_path=file_path,
        num_columns=15,
        skip_rows=1,
        sheet_name="Closed List",
        engine="openpyxl",
    )
    print(f"Loaded {num_rows} rows x {num_cols} columns from 'Closed List'")

    # Drop rows where First Default Date (column F) is blank or #N/A
    before = len(df)
    df = df[
        df["First Default Date"].notna() &
        (df["First Default Date"].astype(str).str.strip().isin(["", "#N/A"]) == False)
    ].reset_index(drop=True)
    print(f"   Dropped {before - len(df)} rows with missing/N/A 'First Default Date'. Remaining: {len(df)}")

    # print(f"Reading 'DCF Receipts' sheet from: {file_path}")
    # automation = ExcelPortfolioAutomation(workbook_path=file_path, visible=False)
    # df_DCF_receipts, num_rows, num_cols = automation.read_excel_data(
    #     excel_path=file_path,
    #     num_columns=15,
    #     skip_rows=2,
    #     sheet_name="DCF Receipts",
    #     engine="pyxlsb",
    # )
    # print(f"Loaded {num_rows} rows x {num_cols} columns from 'DCF Receipts'")

    return df, # df_DCF_receipts

# ==========================
# LOAD REFINANCE DPD REPORT
# ==========================
def load_refinance_dpd_dataframe() -> pd.DataFrame:
    matches = glob.glob(os.path.join(INPUTS_DIR, "Refinance DPD Report *.xlsx"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'Refinance DPD Report *.xlsx' found in {INPUTS_DIR}"
        )
    file_path = matches[0]
    print(f"Reading Refinance DPD Report: {file_path}")

    automation = ExcelPortfolioAutomation(workbook_path=file_path, visible=False)
    df_refinance, num_rows, num_cols = automation.read_excel_data(
        excel_path=file_path,
        num_columns=20,
        skip_rows=0,
        sheet_name=0,
        engine="openpyxl",
    )
    print(f"Loaded {num_rows} rows x {num_cols} columns from Refinance DPD Report")

    df_refinance = df_refinance[["NEW_CON_NO", "OLD_CON_NO"]]
    print(df_refinance.head())
    print(f"Keeping columns: NEW_CON_NO, OLD_CON_NO")
    return df_refinance


# ==========================
# LOAD LEDGER DETAIL REPORT
# ==========================
def load_ledger_detail_dataframe() -> pd.DataFrame:
    matches = glob.glob(os.path.join(INPUTS_DIR, "Ledger Details Report - Internal *.xlsx"))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'Ledger Details Report - Internal *.xlsx' found in {INPUTS_DIR}"
        )
    file_path = matches[0]
    print(f"Reading Ledger Details Report: {file_path}")

    automation = ExcelPortfolioAutomation(workbook_path=file_path, visible=False)
    df_leg_details, num_rows, num_cols = automation.read_excel_data(
        excel_path=file_path,
        num_columns=20,
        skip_rows=0,
        sheet_name=0,
        engine="openpyxl",
    )
    print(f"Loaded {num_rows} rows x {num_cols} columns from Ledger Details Report")
    return df_leg_details

    
# ==========================
# POPULATE RECEIPTS FORMULAS
# ==========================
def populate_receipts_discount_formulas(file_path: str) -> None:
    """
    Copy formulas from row 1 (above header row) in columns H–M
    down to all data rows (row 3 onwards) in the 'Part1' sheet,
    then save the file.
    """
    print(f"Populating H:M formulas in 'Part1' sheet of: {file_path}")
    with ExcelPortfolioAutomation(workbook_path=file_path, visible=False) as automation:
        sheet = automation.workbook.sheets["Part1"]

        # Row 1 = formula row, Row 2 = headers, Row 3+ = data
        # Find last data row using column A
        last_row = sheet.range("A3").end("down").row
        print(f"   Last data row detected: {last_row}")

        # Copy H1:M1 formulas down to all data rows
        source = sheet.range("H1:M1")
        target = sheet.range(f"H3:M{last_row}")
        source.copy(target)
        print(f"   Formulas copied to H3:M{last_row}")

        automation.workbook.save()
    print("   Part1 formulas populated and file saved.")


# ==========================
# LOAD LGD FINAL - PART I
# ==========================
def load_recipts_discounts_dataframe(file_path: str) -> pd.DataFrame:
    print(f"Reading 'Part1' sheet from: {file_path}")

    automation = ExcelPortfolioAutomation(workbook_path=file_path, visible=False)
    df, num_rows, num_cols = automation.read_excel_data(
        excel_path=file_path,
        num_columns=20,
        skip_rows=1,
        sheet_name="Part1",
        engine="pyxlsb",
    )
    print(f"Loaded {num_rows} rows x {num_cols} columns from 'Part1'")
    return df

# ==========================
# EXCEL SERIAL DATE HELPER
# ==========================
def xlsb_serial_to_date(val):
    """
    Convert an Excel date value to a pd.Timestamp.
    - xlsb / numeric: stored as a float serial (days since 1899-12-30)
    - xlsx / openpyxl: already a Python datetime object — handled directly
    - text dates (e.g. the Portfolio file's DATE column, which exports as a
      "DD/MM/YYYY" string rather than a real Excel date): parsed with
      dayfirst=True. Without this branch, a string like "30/09/2025" falls
      through to int(val), raises, and silently becomes NaT — which is
      exactly what was happening to every Virtual Contracts row's Closed
      Date.
    Returns pd.NaT if conversion fails or value is missing.
    """
    try:
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (_datetime, pd.Timestamp)):
            return pd.Timestamp(val)
        if isinstance(val, str):
            return pd.to_datetime(val, dayfirst=True, errors="coerce")
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))
    except Exception:
        return pd.NaT

# ==========================
# FISCAL YEAR HELPER
# ==========================
def to_fiscal_year(date_val) -> str:
    """
    Returns the fiscal year string for a given date.
    Fiscal year runs April–March.
    e.g. 11/06/2023 → '2023/24'
         15/03/2022 → '2021/22'
    """
    try:
        dt = pd.to_datetime(date_val)
        if pd.isnull(dt):
            return "2025/26"
        if dt.month >= 4:
            return f"{dt.year}/{str(dt.year + 1)[-2:]}"
        else:
            return f"{dt.year - 1}/{str(dt.year)[-2:]}"
    except Exception:
        return "2025/26"
    


# ==========================
# BUILD LGD DATAFRAME
# ==========================
def build_lgd(df_closed: pd.DataFrame) -> pd.DataFrame:
    print("Building df_lgd...")

    df_lgd = pd.DataFrame()

    df_lgd["CONTRACT_NO"] = df_closed["CONTRACT_NO"]
    df_lgd["Fac 12"] = df_closed["Fac 12"]

    df_lgd["Closure Type"] = df_closed["CLOSURE_TYPE"]
    df_lgd["LGD Type"] = df_closed["LGD"]
    closed_dates = pd.to_datetime(df_closed["CLOSE_DATE"].apply(xlsb_serial_to_date), errors="coerce")
    df_lgd["Closed Date"] = closed_dates.dt.strftime("%Y/%m/%d")
    df_lgd["Closed Year"] = closed_dates.apply(to_fiscal_year)
    fdd_dates = pd.to_datetime(df_closed["First Default Date"].apply(xlsb_serial_to_date), errors="coerce")
    df_lgd["First Default Date"] = fdd_dates.dt.strftime("%Y/%m/%d")
    df_lgd["Outstanding at First Default"] = df_closed["Outstanding at First Default"]

    # Report as at the run date: leave out contracts closed after it (the Closure Category
    # report is pulled live, so it also contains later closures).
    _run_dt = _run_date_or_none()
    if _run_dt is not None:
        _after = (closed_dates > pd.Timestamp(_run_dt)).to_numpy()
        if _after.any():
            print(f"Closed after the run date ({_run_dt:%d-%m-%Y}): {int(_after.sum()):,} rows left out "
                  f"(latest {closed_dates[_after].max():%Y-%m-%d})")
            df_lgd = df_lgd[~_after].reset_index(drop=True)

    # One row per Fac 12. A rescheduled facility shows up as several contracts whose
    # CONTRACT_NO differ only in the last digit (...0 original, ...1 / ...2 re-issues).
    # Keep the one ending in the lowest digit (0); for exact duplicates (same
    # CONTRACT_NO) the first row wins. Rows with a blank Fac 12 are left alone.
    before = len(df_lgd)
    fac_blank = df_lgd["Fac 12"].isna() | (df_lgd["Fac 12"].astype(str).str.strip() == "")
    suffix = df_lgd["CONTRACT_NO"].astype(str).str.strip().str[-1:]
    keep_order = suffix.where(~fac_blank, "").sort_values(kind="stable").index
    dedup = df_lgd.loc[keep_order]
    dedup = dedup[fac_blank.loc[dedup.index] | ~dedup["Fac 12"].duplicated(keep="first")]
    df_lgd = dedup.sort_index().reset_index(drop=True)
    print(f"Fac 12 de-duplication: removed {before - len(df_lgd)} rows "
          f"(kept CONTRACT_NO ending in lowest digit). Remaining: {len(df_lgd)}")

    # Dup Check: equivalent of =COUNTIF($B:$B, B1) on the Fac 12 column (1 after the de-dup above)
    df_lgd.insert(2, "Dup Check", df_lgd["Fac 12"].map(df_lgd["Fac 12"].value_counts()))

    print(f"df_lgd built: {len(df_lgd)} rows x {len(df_lgd.columns)} columns")
    return df_lgd

# ==========================
# APPEND TO LGD SHEET
# ==========================
def select_fac12_duplicate_rows(contracts: list, fac12s: list) -> list[int]:
    """
    Given parallel lists of CONTRACT_NO and Fac 12 (sheet order), return the list
    indexes to delete so that each Fac 12 keeps exactly one row: the one whose
    CONTRACT_NO ends in the lowest digit (0 beats 1 / 2); exact ties keep the
    first row. Blank Fac 12 values are never treated as duplicates.
    """
    best: dict = {}
    for i, (c, f) in enumerate(zip(contracts, fac12s)):
        key = str(f).strip() if f is not None else ""
        if not key or key.lower() == "nan":
            continue
        suffix = str(c).strip()[-1:] if c is not None else ""
        if key not in best or suffix < best[key][1]:
            best[key] = (i, suffix)
    keep = {v[0] for v in best.values()}
    return [
        i for i, f in enumerate(fac12s)
        if (str(f).strip() if f is not None else "") not in ("", "nan") and i not in keep
    ]


def dedupe_lgd_sheet_by_fac12(sheet, header_map: dict) -> int:
    """
    Second-pass check on the LGD sheet after new rows were appended: remove any
    Fac 12 duplicates still present (existing or new) and make sure Dup Check
    reads 1 for the survivors. Returns the number of rows deleted.
    """
    contract_ci = header_map.get("CONTRACT_NO")
    fac_ci = header_map.get("Fac 12")
    dup_ci = header_map.get("Dup Check")
    if contract_ci is None or fac_ci is None:
        print("   Dup verification skipped: CONTRACT_NO / Fac 12 column not found")
        return 0

    last_row = sheet.range(f"B{sheet.api.Rows.Count}").end("up").row
    n = last_row - 2
    if n <= 0:
        return 0

    def _col(ci: int) -> list:
        v = sheet.range((3, ci + 1), (last_row, ci + 1)).value
        return v if isinstance(v, list) else [v]

    contracts, fac12s = _col(contract_ci), _col(fac_ci)
    drop = select_fac12_duplicate_rows(contracts, fac12s)
    print(f"   Dup verification: {len(drop)} duplicate Fac 12 row(s) found in {n:,} rows")
    if not drop:
        return 0

    for i in drop:
        print(f"      removing row {i + 3}: {contracts[i]} (Fac 12 {fac12s[i]})")
    for i in sorted(drop, reverse=True):          # bottom-up so row numbers stay valid
        sheet.api.Rows(i + 3).Delete()

    # Dup Check: leave it alone if it is a COUNTIF formula (it recalculates);
    # if it holds plain values, reset the survivors of the de-duped groups to 1.
    if dup_ci is not None:
        new_last = sheet.range(f"B{sheet.api.Rows.Count}").end("up").row
        dropped_facs = {str(fac12s[i]).strip() for i in drop}
        facs_now = sheet.range((3, fac_ci + 1), (new_last, fac_ci + 1)).value
        facs_now = facs_now if isinstance(facs_now, list) else [facs_now]
        for j, f in enumerate(facs_now):
            if f is not None and str(f).strip() in dropped_facs:
                cell = sheet.range((j + 3, dup_ci + 1))
                if not str(cell.formula).startswith("="):
                    cell.value = 1
    print(f"   Dup verification: removed {len(drop)} row(s)")
    return len(drop)


def append_to_lgd_sheet(df: pd.DataFrame) -> None:
    """
    Append df rows to the LGD sheet in LGD *.xlsb, matching columns by name.
    Then copy K:M formulas from the last existing data row into the new rows.
    """
    matches = glob.glob(os.path.join(INPUTS_DIR, "LGD *.xlsb"))
    if not matches:
        raise FileNotFoundError(f"No file matching 'LGD *.xlsb' found in {INPUTS_DIR}")
    file_path = matches[0]
    print(f"Appending to LGD sheet in: {file_path}")

    with ExcelPortfolioAutomation(workbook_path=file_path, visible=False) as automation:
        sheet = automation.workbook.sheets["LGD"]

        # Row 1 = formula row, Row 2 = headers, Row 3+ = data
        # Read headers from row 2
        headers = sheet.range("A2").expand("right").value
        if not isinstance(headers, list):
            headers = [headers]
        print(f"   LGD sheet columns ({len(headers)}): {headers}")

        # Map header name → 0-based index
        header_map = {str(h).strip(): i for i, h in enumerate(headers) if h is not None}

        # DCF Receipts column index (needed for both new-row writes and existing-row updates)
        dcf_col_idx = header_map.get("DCF Receipts")

        # Build Fac 12 → DCF Receipts lookup from df (used to update ALL rows in sheet)
        dcf_lookup: dict = {}
        if dcf_col_idx is not None and "DCF Receipts" in df.columns and "Fac 12" in df.columns:
            dcf_lookup = (
                df.drop_duplicates(subset=["Fac 12"], keep="first")
                .set_index("Fac 12")["DCF Receipts"]
                .fillna(0)
                .to_dict()
            )

        # Find last row with existing data — search from bottom up, immune to gaps
        total_rows = sheet.api.Rows.Count
        last_data_row = sheet.range(f"B{total_rows}").end("up").row
        print(f"   Last existing data row: {last_data_row}")

        # Read existing Fac 12 values to avoid duplicates on re-run
        existing_fac12 = set(
            str(v).strip()
            for v in sheet.range(f"B3:B{last_data_row}").value or []
            if v is not None
        )
        print(f"   Existing Fac 12 count in sheet: {len(existing_fac12)}")

        # Build rows aligned to LGD sheet column order, skipping already-present Fac 12s
        rows_to_write = []
        skipped = 0
        for _, row in df.iterrows():
            fac12_val = str(row.get("Fac 12", "")).strip()
            if fac12_val in existing_fac12:
                skipped += 1
                continue
            new_row = [None] * len(headers)
            for col_name, value in row.items():
                if str(col_name).strip() in header_map:
                    new_row[header_map[str(col_name).strip()]] = value
            rows_to_write.append(new_row)
        print(f"   Skipped {skipped} rows (Fac 12 already in sheet), writing {len(rows_to_write)} new rows")

        if rows_to_write:
            # Normalise None/NaN → None so numeric columns stay numeric.
            # Excel's existing cell format (e.g. custom format showing 0 as "-") handles display.
            # Writing the string '-' would break formula columns (K = I - J → #VALUE!).
            def _clean(v):
                try:
                    if v is None or pd.isna(v):
                        return None
                except TypeError:
                    pass
                return v
            rows_to_write = [[_clean(v) for v in row] for row in rows_to_write]

            # DCF Receipts: blanks should be 0, not None
            if dcf_col_idx is not None:
                for row in rows_to_write:
                    if row[dcf_col_idx] is None:
                        row[dcf_col_idx] = 0

            # Write data starting after last existing row
            start_row = last_data_row + 1
            end_row = start_row + len(rows_to_write) - 1
            sheet.range(f"A{start_row}").value = rows_to_write
            print(f"   Written {len(rows_to_write)} rows to LGD sheet (rows {start_row}:{end_row})")
        else:
            print("   No new rows to append — updating existing DCF Receipts values in place.")

        # Determine actual last row after any new writes
        total_rows = sheet.api.Rows.Count
        actual_last_row = sheet.range(f"B{total_rows}").end("up").row

        # Copy K:M formulas only for newly added rows (existing rows already have them)
        if rows_to_write:
            sheet.range("K1:M1").copy(sheet.range(f"K{start_row}:M{end_row}"))
            print(f"   K:M formulas copied to new rows {start_row}:{end_row}")

        # Update DCF Receipts for ALL rows in sheet via single batch write
        if dcf_col_idx is not None:
            dcf_col_num = dcf_col_idx + 1      # xlwings uses 1-based column numbers

            # Read all Fac 12 values (column B) and existing DCF Receipts in one call each
            b_values = sheet.range(f"B3:B{actual_last_row}").value or []
            if b_values and not isinstance(b_values[0], list):
                b_values = [[v] for v in b_values]

            existing_dcf = sheet.range((3, dcf_col_num), (actual_last_row, dcf_col_num)).value or []
            if existing_dcf and not isinstance(existing_dcf[0], list):
                existing_dcf = [[v] for v in existing_dcf]

            # Build the full column as a list-of-lists for one batch write
            # Rule: if mapping provides a value → update; else keep existing value
            dcf_col_values = []
            updated = 0
            for i, row_vals in enumerate(b_values):
                fac12 = str(row_vals[0]).strip() if row_vals and row_vals[0] is not None else ""
                if fac12 in dcf_lookup:
                    dcf_col_values.append([dcf_lookup[fac12]])
                    updated += 1
                else:
                    existing_val = existing_dcf[i][0] if i < len(existing_dcf) and existing_dcf[i] else None
                    dcf_col_values.append([existing_val if existing_val is not None else 0])

            # Single range write — replaces all 153k individual cell writes
            sheet.range((3, dcf_col_num), (actual_last_row, dcf_col_num)).value = dcf_col_values
            print(f"   DCF Receipts updated for {updated} rows from receipt data (batch write)")

            # Apply accounting format to entire column
            dcf_range = sheet.range((3, dcf_col_num), (actual_last_row, dcf_col_num))
            accounting_fmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
            dcf_range.api.NumberFormat = accounting_fmt
            print(f"   Accounting format applied to DCF Receipts (col {dcf_col_num}) rows 3:{actual_last_row}")

        # Second-pass dup check: clears Fac 12 duplicates already in the sheet and
        # any that slipped in with the new rows (keeps the CONTRACT_NO ending in 0).
        dedupe_lgd_sheet_by_fac12(sheet, header_map)

        automation.workbook.save()
    print("   LGD sheet updated and saved.")


# ==========================
# NORMALISE LGD TYPE IN SHEET
# ==========================
def normalise_lgd_type_in_sheet() -> None:
    """
    Find-and-replace LGD Type values directly in the LGD sheet of LGD *.xlsb,
    operating on the full column so both existing and newly appended rows are covered.
    """
    matches = glob.glob(os.path.join(INPUTS_DIR, "LGD *.xlsb"))
    if not matches:
        raise FileNotFoundError(f"No file matching 'LGD *.xlsb' found in {INPUTS_DIR}")
    file_path = matches[0]
    print(f"Normalising LGD Type values in sheet: {file_path}")

    replacements = {
        "Mini Truck":      "MINI T",
        "Other Equipment": "O.EQP",
        "Revolving Loans": "RL",
        "SME Loans":       "SME",
        "Term Loans":      "TL",
        "Revolving Loan": "RL",
        "SME Loan":       "SME",
        "Term Loan":      "TL",
        "Commercial Vehicles": "CV",
        "Commercial Vehicle": "CV",
    }

    with ExcelPortfolioAutomation(workbook_path=file_path, visible=False) as automation:
        sheet = automation.workbook.sheets["LGD"]

        # Locate the LGD Type column from the header row (row 2)
        headers = sheet.range("A2").expand("right").value
        if not isinstance(headers, list):
            headers = [headers]
        try:
            lgd_type_col = next(
                i + 1 for i, h in enumerate(headers)
                if h is not None and str(h).strip() == "LGD Type"
            )
        except StopIteration:
            print("   'LGD Type' column not found in LGD sheet — skipping normalisation")
            return

        total_rows = sheet.api.Rows.Count
        last_row = sheet.range(f"A{total_rows}").end("up").row

        col_range = sheet.range((3, lgd_type_col), (last_row, lgd_type_col))

        for find_val, replace_val in replacements.items():
            col_range.api.Replace(
                What=find_val,
                Replacement=replace_val,
                LookAt=2,          # xlPart=2 / xlWhole=1  — use xlWhole for exact match
                SearchOrder=1,
                MatchCase=False,
                SearchFormat=False,
                ReplaceFormat=False,
            )
            print(f"   Replaced '{find_val}' → '{replace_val}'")

        automation.workbook.save()
    print("   LGD Type normalisation complete and saved.")


# ==========================
# REFRESH LGD PIVOT TABLE
# ==========================
def refresh_lgd_pivot() -> None:
    """
    Update PivotTable2 in the 'Pivot' sheet of LGD *.xlsb.
    Source range: LGD sheet from A2 (header row) to the last cell
    with data in the last column, detected dynamically.
    """
    matches = glob.glob(os.path.join(INPUTS_DIR, "LGD *.xlsb"))
    if not matches:
        raise FileNotFoundError(f"No file matching 'LGD *.xlsb' found in {INPUTS_DIR}")
    file_path = matches[0]
    print(f"Refreshing PivotTable2 in: {file_path}")

    with ExcelPortfolioAutomation(workbook_path=file_path, visible=False) as automation:
        # update_pivot_source_and_refresh detects last row and last column dynamically
        # start_cell="A2" — headers are in row 2 of the LGD sheet
        automation.update_pivot_source_and_refresh(
            sheet_name="Pivot",
            pivot_table_name="PivotTable2",
            data_sheet_name="LGD",
            start_cell="A2",
            end_column=None,        # auto-detect last column
            include_headers=True,
        )
        automation.workbook.save()
    print("   PivotTable2 refreshed and saved.")


# ==========================
# LOAD YARD FILE
# ==========================
def load_yard_dataframe() -> pd.DataFrame:
    """
    Load Yard Provision file, sheet '1.YardList', skip 3 rows (4th row = headers).
    Keep 'Contract No.' and 'Market Value', then compute:
        Amount     = Market Value * 0.8
        Facility 12 = first 12 chars of Contract No.
    Saves the working df to 'YARD_DF.xlsx' in ROOT_DIR.
    """
    file_path = find_Yard_file()
    print(f"Reading Yard file: {file_path}")

    YARD_DF = pd.read_excel(
        file_path,
        sheet_name="1.YardList",
        skiprows=3,
        engine="pyxlsb",
    )

    YARD_DF = YARD_DF[["Contract No.", "Market Value","Valuation Provision"]].copy()
    YARD_DF = YARD_DF[YARD_DF["Contract No."].notna()].reset_index(drop=True)
    YARD_DF = YARD_DF[
        YARD_DF["Valuation Provision"].notna() & (YARD_DF["Valuation Provision"] != 0)
    ].reset_index(drop=True)



    YARD_DF["Amount"]      = YARD_DF["Valuation Provision"].copy() #YARD_DF["Market Value"] * 0.8
    YARD_DF["Facility 12"] = YARD_DF["Contract No."].astype(str).str[:12]

    out_path = os.path.join(ROOT_DIR, "YARD_DF.xlsx")
    YARD_DF.to_excel(out_path, index=False)
    print(f"YARD_DF saved to: {out_path} ({len(YARD_DF)} rows)")

    return YARD_DF


# ==========================
# LOAD WRITEBACKS FILE
# ==========================
def load_writebacks_dataframe() -> pd.DataFrame:
    """
    Load the Ledger Details Report (writebacks) file, filter DITRECT CREDIT rows,
    keep relevant columns, and derive Facility 12 (first 12 chars of SLT_CONNO).
    """
    file_path = find_writebacks_file()
    print(f"Reading writebacks file: {file_path}")
    df = pd.read_excel(file_path, engine="openpyxl")

    mid_chars = df["SLT_TRNNO"].astype(str).str[4:6]
    writebacks = df[mid_chars.isin(["DC", "RC", "RQ"])].copy()
    
    cols = ["SLT_CONNO", "SLT_TRNNO", "SLT_TRNAMOUNT", "SLST_TRNDRCR", "SLT_DESC", "SLT_TRNDATE"]
    writebacks = writebacks[cols].reset_index(drop=True)

    # Remove DC CANCELLATION rows and their matching counterpart transactions.
    # A cancellation row matches another row on SLT_CONNO + SLT_TRNAMOUNT; both are dropped.
    cancellations = writebacks[writebacks["SLT_DESC"] == "DC CANCELLATION"]
    indices_to_drop = set(cancellations.index)

    for _, cancel_row in cancellations.iterrows():
        match_mask = (
            (writebacks["SLT_CONNO"]    == cancel_row["SLT_CONNO"])    &
            (writebacks["SLT_TRNAMOUNT"] == cancel_row["SLT_TRNAMOUNT"]) &
            (writebacks["SLT_DESC"]     != "DC CANCELLATION")
        )
        indices_to_drop.update(writebacks[match_mask].index.tolist())

    before = len(writebacks)
    writebacks = writebacks.drop(index=list(indices_to_drop)).reset_index(drop=True)
    print(f"   Removed {before - len(writebacks)} rows (DC CANCELLATION + matched transactions)")

    writebacks["Facility 12"] = writebacks["SLT_CONNO"].astype(str).str[:12]

    print(f"Writebacks loaded: {len(writebacks)} rows (SLT_TRNNO[4:6] in DC/RC/RQ)")
    return writebacks


# ==========================
# LOAD WRITEBACKS130 FILE
# ==========================
def load_writebacks130_dataframe() -> pd.DataFrame:
    """
    Load all Write Off Working*.xlsx files, concatenate them, filter Transaction Type == 'DC'/'RC'/'RQ',
    and reshape to match the writebacks schema (SLT_* column names).
    Column mapping:
        Original Contract No → SLT_CONNO
        Transaction No       → SLT_TRNNO
        Trans Date           → SLT_TRNDATE
        Credit               → SLT_TRNAMOUNT
    """
    matches = sorted(glob.glob(os.path.join(INPUTS_DIR, "Write Off Working*.xlsx")))
    if not matches:
        raise FileNotFoundError(
            f"No file matching 'Write Off Working*.xlsx' found in {INPUTS_DIR}"
        )
    print(f"Found {len(matches)} Write Off Working file(s): {[os.path.basename(m) for m in matches]}")

    dfs = []
    for file_path in matches:
        print(f"   Reading: {os.path.basename(file_path)}")
        _df = pd.read_excel(file_path, sheet_name='0130-GL_Working', engine="openpyxl", skiprows=0)
        dfs.append(_df)
    df = pd.concat(dfs, ignore_index=True)
    print(f"   Combined shape: {df.shape}")
    print(f"{df.columns.tolist()}")
    print(f"{df.head()}")
    filtered = df[df["Transaction Type"].isin(["DC", "RC", "RQ"])].copy()
    filtered = filtered[(filtered["Credit"].notna()) & (filtered["Credit"] != 0)]

    
    # Remove offsetting duplicate contracts: if a contract appears more than once
    # and its Debit and Credit values net to zero, drop all those rows.
    dup_mask = filtered.duplicated(subset=["Original Contract No"], keep=False)
    if dup_mask.any():
        dup_contracts = filtered[dup_mask].copy()
        def _nets_to_zero(grp):
            return grp["Debit"].fillna(0).sum() == grp["Credit"].fillna(0).sum()
        offsetting = dup_contracts.groupby("Original Contract No").filter(_nets_to_zero)
        if not offsetting.empty:
            print(f"   Removing {len(offsetting)} offsetting rows for contracts: "
                  f"{offsetting['Original Contract No'].unique().tolist()}")
            filtered = filtered.drop(index=offsetting.index)

    wb130 = pd.DataFrame()
    wb130["SLT_CONNO"]    = filtered["Original Contract No"].astype(str)
    wb130["SLT_TRNNO"]    = filtered["Transaction No"]
    # 'Trans Date' is stored as a raw Excel serial (e.g. 45747) in this sheet. Left as
    # a number, pandas later reads it as nanoseconds since 1970 -> every date 1970-01-01.
    wb130["SLT_TRNDATE"]  = filtered["Trans Date"].apply(xlsb_serial_to_date)
    wb130["SLT_TRNAMOUNT"] = filtered["Credit"]
    wb130["SLST_TRNDRCR"] = None
    wb130["SLT_DESC"]     = None
    wb130["SLT_LDCONNO"] = filtered['Contract No'].astype(str)  # For potential future use, not currently in DC List schema

    wb130["Facility 12"] = wb130["SLT_CONNO"].str[:12]

    print(f"Writebacks130 loaded: {len(wb130)} 'DC' rows")
    return wb130
# ==========================
# PERIOD FILE NAMES (writeoff / writebacks outputs)
# ==========================
def _run_date_or_none():
    """Run date from sys.argv[1] (DD/MM/YYYY, as ifrs.py passes it), or None."""
    if len(sys.argv) < 2:
        return None
    try:
        return datetime.strptime(sys.argv[1].strip(), "%d/%m/%Y")
    except ValueError:
        return None


def drop_closed_after_run_date(df: pd.DataFrame, closed_col: str, label: str) -> pd.DataFrame:
    """
    Remove rows whose Closed Date is later than the run date (sys.argv[1]). The Closure
    Category Wise Report is pulled from the ERP when the run starts, so it also holds
    contracts closed AFTER the month being reported (e.g. Oct 2025 .. Oct 2026 in a
    30-09-2025 report). Rows with a blank Closed Date are kept. No run date -> no change.
    """
    run_dt = _run_date_or_none()
    if run_dt is None or closed_col not in df.columns:
        return df
    closed = pd.to_datetime(df[closed_col].apply(xlsb_serial_to_date), errors="coerce")
    after = (closed > pd.Timestamp(run_dt)).to_numpy()
    n = int(after.sum())
    if n:
        print(f"{label}: removed {n:,} rows closed after the run date {run_dt:%d-%m-%Y} "
              f"(latest {closed[after].max():%Y-%m-%d}) - {len(df) - n:,} rows remain")
        return df.loc[~after].reset_index(drop=True)
    print(f"{label}: no rows closed after the run date {run_dt:%d-%m-%Y}")
    return df


def writeoff_output_name() -> str:
    """'writeoff DD-MM-YYYY' for the run date (falls back to the old fixed name)."""
    run_dt = _run_date_or_none()
    return f"writeoff {run_dt:%d-%m-%Y}" if run_dt else "writeoff 31-08-2025"


def writebacks_output_name() -> str:
    """'writebacks <financial-year start> to <run date>' (falls back to the old fixed name)."""
    run_dt = _run_date_or_none()
    if not run_dt:
        return "writebacks 01-04-2025 to 31-08-2025"
    fy_year = run_dt.year if run_dt.month >= 4 else run_dt.year - 1
    return f"writebacks 01-04-{fy_year} to {run_dt:%d-%m-%Y}"


def _remove_other_versions(pattern: str, keep_path: str) -> None:
    """Delete older copies of a period file (e.g. 'writeoff 31-08-2025.xlsx') so only one remains."""
    for old in glob.glob(os.path.join(ROOT_DIR, pattern)):
        if os.path.abspath(old) != os.path.abspath(keep_path):
            try:
                os.remove(old)
                print(f"Removed previous period file: {os.path.basename(old)}")
            except OSError as exc:
                print(f"Could not remove {os.path.basename(old)}: {exc}")


def drop_duplicate_dc_rows(wb: pd.DataFrame) -> tuple:
    """
    Remove exact duplicate DC transactions (same Contract No, Transaction No, Amount, Date).
    process_writebacks_dcf adds the Ledger rows to the DC List and writes the result back to the
    same file, so each saved run used to add the same Ledger rows again. Values are normalised
    (trimmed text, amount to 2 dp, date without time) so a row read back from Excel still
    matches. First occurrence wins. Returns (frame, number_removed).
    """
    key = pd.DataFrame({
        "c": wb["Contract No"].astype(str).str.strip(),
        "t": wb["Transaction No"].astype(str).str.strip(),
        "a": pd.to_numeric(wb["Amount"], errors="coerce").round(2),
        "d": pd.to_datetime(wb["Date"], errors="coerce").dt.normalize(),
    })
    dup = key.duplicated(keep="first").to_numpy()
    return wb.loc[~dup].reset_index(drop=True), int(dup.sum())


# ==========================
# PROCESS WRITEBACKS DCF
# ==========================
def process_writebacks_dcf(
    writebacks: pd.DataFrame,
    _df_closed: pd.DataFrame,
    output_name: str,
) -> pd.Series:
    """
    Compute full DCF pipeline on a writebacks df, save to '{output_name}.xlsx',
    and return aggregated DCF Series keyed by Facility 12.

    DCF rules:
      - SLST_TRNDRCR blank → DCF = SLT_TRNAMOUNT (no discounting)
      - Otherwise          → DCF = SLT_TRNAMOUNT * DF
        where DF = IF(Date Gap >= 0, 1/(1+Int Rate%)^(Date Gap/365), 0)
    """
    wb = writebacks.copy()

    # ── Select & rename writebacks columns to match DC List headings ────────
    wb = (
        wb[["SLT_CONNO", "SLT_TRNNO", "SLT_TRNAMOUNT", "SLT_TRNDATE", "Facility 12"]]
        .rename(columns={
            "SLT_CONNO":    "Contract No",
            "SLT_TRNNO":    "Transaction No",
            "SLT_TRNAMOUNT": "Amount",
            "SLT_TRNDATE":  "Date",
        })
    )

    # ── Load DC List and concat to wb ─────────────────────────────────────
    dc_list_file = find_dc_list_file()
    print(f"Reading 'DC' sheet from: {dc_list_file}")
    df_dc_list = pd.read_excel(
        dc_list_file,
        sheet_name="DC",
        skiprows=1,
        engine="openpyxl",
        usecols=["Transaction No", "Contract No", "Amount", "Date"],
    )
    df_dc_list["Facility 12"] = df_dc_list["Contract No"].astype(str).str[:12]
    df_dc_list = df_dc_list[["Contract No", "Transaction No", "Amount", "Date", "Facility 12"]]
    print(f"DC List loaded: {len(df_dc_list)} rows, {df_dc_list['Facility 12'].nunique()} unique Facility 12")
    wb = pd.concat([wb, df_dc_list], ignore_index=True)
    print(f"wb after concat with DC List: {len(wb)} rows")
    wb, n_dup = drop_duplicate_dc_rows(wb)
    print(f"Removed {n_dup:,} duplicate DC rows (same contract, transaction, amount and date) "
          f"- {len(wb):,} rows remain")

    # ── Write combined data back to DC List*.xlsx → DC sheet from A2 ──────
    with pd.ExcelWriter(
        dc_list_file,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="overlay",
    ) as writer:
        wb.to_excel(writer, sheet_name="DC", startrow=1, index=False)
    print(f"Combined data written to 'DC' sheet in: {dc_list_file} ({len(wb)} rows, headers at A2)")

    # ── Save to Excel ──────────────────────────────────────────────────────
    out_path = os.path.join(ROOT_DIR, f"{output_name}.xlsx")
    wb.to_excel(out_path, index=False)
    print(f"{output_name} saved to: {out_path} ({len(wb)} rows)")
    _remove_other_versions("writeoff *.xlsx", out_path)

    # ── Aggregate by Facility 12 and return ───────────────────────────────
    wb_agg = wb.groupby("Facility 12").size()
    print(f"{output_name}: {len(wb_agg)} unique Facility 12 values")
    return wb_agg


# ==========================
# PROCESS WRITEBACKS DCF
# ==========================
def process_writebacks130_dcf(
    writebacks: pd.DataFrame,
    _df_closed: pd.DataFrame,
    output_name: str,
) -> pd.Series:
    """
    Compute full DCF pipeline on a writebacks df, save to '{output_name}.xlsx',
    and return aggregated DCF Series keyed by Facility 12.

    DCF rules:
      - SLST_TRNDRCR blank → DCF = SLT_TRNAMOUNT (no discounting)
      - Otherwise          → DCF = SLT_TRNAMOUNT * DF
        where DF = IF(Date Gap >= 0, 1/(1+Int Rate%)^(Date Gap/365), 0)

    Int Rate and First Default Date are sourced from the 'D List and Map' sheet
    of 'Write Backs 2024-25.xlsx', keyed by Fac 12.
    """
    wb = writebacks.copy()

    # ── Build lookups from Write Backs 2024-25 → "D List and Map" sheet ───
    wb_file = find_write_backs_file()
    print(f"Reading 'D List and Map' sheet from: {wb_file}")
    df_dmap = pd.read_excel(wb_file, sheet_name="D List", engine="openpyxl")
    dmap_dedup      = df_dmap.drop_duplicates(subset=["Fac 12"], keep="first").set_index("Fac 12")
    fdd_lookup      = dmap_dedup["First Default Date"]
    int_rate_lookup = dmap_dedup["Int Rate"]

    # ── First Default Date ─────────────────────────────────────────────────
    wb["First Default Date"] = pd.to_datetime(
        wb["Facility 12"].map(fdd_lookup), errors="coerce"
    )

    # ── SLT_TRNDATE → Timestamp ────────────────────────────────────────────
    wb["SLT_TRNDATE"] = pd.to_datetime(wb["SLT_TRNDATE"], dayfirst=True, errors="coerce")

    # ── Date Gap (calendar days) ───────────────────────────────────────────
    wb["Date Gap"] = (wb["SLT_TRNDATE"] - wb["First Default Date"]).dt.days

    # ── Int Rate ───────────────────────────────────────────────────────────
    wb["Int Rate"] = wb["Facility 12"].map(int_rate_lookup)

    # ── DF = IF(Date Gap >= 0, 1/(1+Int Rate%)^(Date Gap/365), 0) ─────────
    def _calc_df(row):
        try:
            date_gap = row["Date Gap"]
            int_rate = row["Int Rate"]
            if pd.isna(date_gap) or pd.isna(int_rate):
                return 0
            return 1 / (1 + int_rate / 100) ** (date_gap / 365) if date_gap >= 0 else 0
        except Exception:
            return 0

    wb["DF"] = wb.apply(_calc_df, axis=1)

    # ── DCF: blank SLST_TRNDRCR → use SLT_TRNAMOUNT directly ──────────────
    
    wb["DCF"] = wb["SLT_TRNAMOUNT"] * wb["DF"]

    # ── Save to Excel ──────────────────────────────────────────────────────
    out_path = os.path.join(ROOT_DIR, f"{output_name}.xlsx")
    wb.to_excel(out_path, index=False)
    print(f"{output_name} saved to: {out_path} ({len(wb)} rows)")
    _remove_other_versions("writebacks * to *.xlsx", out_path)

    # ── Aggregate DCF by Facility 12 and return ────────────────────────────
    wb_agg = wb.groupby("Facility 12")["DCF"].sum()
    print(f"{output_name}: {len(wb_agg)} unique Facility 12 values with DCF")
    return wb_agg


# ==========================
# LOAD JOBHD RECEIPT FILES
# ==========================
def load_receipt_dataframe() -> pd.DataFrame:
    """
    Read all JOBHD*.xlsx files from ROOT_DIR (LGD Calculation folder),
    concatenate them, and create a Facility 12 column (first 12 chars of RSM_CONNO).
    Amount column in these files is 'SUM(RSM_STLDAMT)'.
    """
    matches = sorted(glob.glob(os.path.join(ROOT_DIR, "JOBHD*.xlsx")))
    if not matches:
        raise FileNotFoundError(f"No file matching 'JOBHD*.xlsx' found in {ROOT_DIR}")
    print(f"Found {len(matches)} JOBHD receipt file(s): {[os.path.basename(m) for m in matches]}")

    dfs = []
    for f in matches:
        print(f"   Reading: {os.path.basename(f)}")
        df = pd.read_excel(f, engine="openpyxl")
        dfs.append(df)
    df_receipt = pd.concat(dfs, ignore_index=True)

    # Rename amount column to a consistent name
    if "SUM(RSM_STLDAMT)" in df_receipt.columns:
        df_receipt = df_receipt.rename(columns={"SUM(RSM_STLDAMT)": "RSM_STLDAMT"})

    # Facility 12 = first 12 digits of RSM_CONNO
    df_receipt["Facility 12"] = df_receipt["RSM_CONNO"].astype(str).str[:12]

    print(f"Receipt df loaded: {len(df_receipt)} rows")
    return df_receipt


# ==========================
# COMPUTE DCF FROM RECEIPTS
# ==========================
def _default_lookup_from_d_lists() -> tuple:
    """
    Fallback source of First Default Date / Int Rate by Fac 12, read from the
    'D List' sheet of every 'Receipt*Discount*.xlsb' in the LGD folder.

    Used for facilities the Closed List has no First Default Date for (FDDW's
    lookup left them blank, so compute_dcf_from_receipts() could not discount
    their receipts). If a Fac 12 appears more than once, the row with the
    earliest First Default Date wins.
    Returns (first_default_date, int_rate): two Series indexed by Fac 12.
    """
    frames = []
    for path in sorted(glob.glob(os.path.join(INPUTS_DIR, "Receipt*Discount*.xlsb"))):
        try:
            d = pd.read_excel(path, sheet_name="D List", engine="pyxlsb")
        except Exception as exc:
            print(f"   D List not read from {os.path.basename(path)}: {exc}")
            continue
        need = {"Fac 12", "First Default Date", "Int Rate"}
        if not need <= set(d.columns):
            print(f"   {os.path.basename(path)} D List skipped - missing columns {sorted(need - set(d.columns))}")
            continue
        d = d[["Fac 12", "First Default Date", "Int Rate"]].copy()
        d["First Default Date"] = pd.to_datetime(d["First Default Date"].apply(xlsb_serial_to_date), errors="coerce")
        d = d[d["Fac 12"].notna() & d["First Default Date"].notna()]
        d["Fac 12"] = d["Fac 12"].astype(str).str.strip()
        print(f"   {os.path.basename(path)} D List: {len(d):,} rows with a First Default Date")
        frames.append(d)
    if not frames:
        empty = pd.Series(dtype="object")
        return empty, empty
    allrows = pd.concat(frames, ignore_index=True).sort_values("First Default Date")
    allrows = allrows.drop_duplicates(subset=["Fac 12"], keep="first").set_index("Fac 12")
    return allrows["First Default Date"], allrows["Int Rate"]


def compute_dcf_from_receipts(df_receipt: pd.DataFrame, df_closed: pd.DataFrame) -> pd.DataFrame:
    """
    Adds Default Date, Date Gap, Int Rate, DF, and DCF columns to df_receipt.

    - Default Date : First Default Date from df_closed matched on Facility 12 → Fac 12
    - Date Gap     : RSM_STLDDATE - Default Date (days)
    - Int Rate     : Int Rate from df_closed matched on Facility 12 → Fac 12
    - DF           : IF(Date Gap >= 0, 1 / (1 + Int Rate%) ^ (Date Gap / 365), 0)
    - DCF          : RSM_STLDAMT * DF
    """
    # Build lookup tables from df_closed (drop duplicates on Fac 12, keep first)
    closed_dedup = df_closed.drop_duplicates(subset=["Fac 12"], keep="first").set_index("Fac 12")

    fdd_lookup      = closed_dedup["First Default Date"]
    int_rate_lookup = closed_dedup["Int Rate"]

    # Default Date — map then ensure Timestamp
    df_receipt["Default Date"] = pd.to_datetime(
        df_receipt["Facility 12"]
        .map(fdd_lookup)
        .apply(xlsb_serial_to_date),
        errors="coerce",
    )

    # RSM_STLDDATE — parse date strings like "05-May-24"
    df_receipt["RSM_STLDDATE"] = pd.to_datetime(
        df_receipt["RSM_STLDDATE"], dayfirst=True, errors="coerce"
    )

    # Int Rate (stored as percentage points, e.g. 20.0 means 20%)
    df_receipt["Int Rate"] = df_receipt["Facility 12"].map(int_rate_lookup)

    # Fallback: facilities with no First Default Date / Int Rate in the Closed List
    # take them from the Receipts Discount 'D List' sheets.
    need_fb = df_receipt["Default Date"].isna() | df_receipt["Int Rate"].isna()
    if need_fb.any():
        fb_fdd, fb_rate = _default_lookup_from_d_lists()
        fb_dates = pd.to_datetime(df_receipt["Facility 12"].map(fb_fdd), errors="coerce")
        fb_rates = df_receipt["Facility 12"].map(fb_rate)
        fill_date = df_receipt["Default Date"].isna() & fb_dates.notna()
        fill_rate = df_receipt["Int Rate"].isna() & fb_rates.notna()
        df_receipt.loc[fill_date, "Default Date"] = fb_dates[fill_date]
        df_receipt.loc[fill_rate, "Int Rate"] = fb_rates[fill_rate]
        print(f"D List fallback: Default Date filled on {int(fill_date.sum()):,} rows "
              f"({df_receipt.loc[fill_date, 'Facility 12'].nunique():,} facilities), "
              f"Int Rate on {int(fill_rate.sum()):,} rows; "
              f"{int((df_receipt['Default Date'].isna() | df_receipt['Int Rate'].isna()).sum()):,} rows still blank")

    # Date Gap in calendar days
    df_receipt["Date Gap"] = (
        df_receipt["RSM_STLDDATE"] - df_receipt["Default Date"]
    ).dt.days

    # DF = IF(Date Gap >= 0, 1 / (1 + Int Rate/100) ^ (Date Gap/365), 0)
    def _calc_df(row):
        try:
            date_gap = row["Date Gap"]
            int_rate = row["Int Rate"]
            if pd.isna(date_gap) or pd.isna(int_rate):
                return 0
            if date_gap >= 0:
                return 1 / (1 + int_rate / 100) ** (date_gap / 365)
            return 0
        except Exception:
            return 0

    df_receipt["DF"]  = df_receipt.apply(_calc_df, axis=1)
    df_receipt["DCF"] = df_receipt["RSM_STLDAMT"] * df_receipt["DF"]

    matched = df_receipt["Default Date"].notna().sum()
    print(f"DCF computation done — {matched}/{len(df_receipt)} rows matched a Default Date")
    return df_receipt


# ==========================
# SAVE DCF PIVOT
# ==========================
def save_dcf_pivot(df_receipt: pd.DataFrame) -> pd.DataFrame:
    """
    Pivot df_receipt on Facility 12, summing DCF for each unique value.
    Saves the pivot to ROOT_DIR as 'DCF_Receipts_Pivot.xlsx'.
    Returns the pivot DataFrame.
    """
    pivot = (
        df_receipt
        .groupby("Facility 12", as_index=False)["DCF"]
        .sum()
        .rename(columns={"DCF": "Sum of DCF"})
    )

    out_path = os.path.join(ROOT_DIR, "DCF_Receipts_Pivot.xlsx")
    pivot.to_excel(out_path, index=False)
    print(f"DCF pivot saved to: {out_path} ({len(pivot)} unique Facility 12 values)")
    return pivot


# ==========================
# APPLY WRITEBACKS TO LGD FINAL
# ==========================
def normalise_virtual_contract_type(df: pd.DataFrame) -> pd.DataFrame:
    """
    Fold one-off batch labels such as 'VC 31.03.2025' into 'Virtual Contracts', so
    that is the only Virtual Contract type in the LGD sheet. FDDW only ever writes
    'Virtual Contracts'; the dated labels were relabelled by hand.
    """
    ct = df["Closure Type"].astype(str).str.strip()
    relabel = ct.str.match(r"^VC\b") & ct.ne("Virtual Contracts")
    n = int(relabel.sum())
    if n:
        print(f"Closure Type: relabelled {n:,} rows {sorted(ct[relabel].unique())} -> 'Virtual Contracts'")
        df = df.copy()
        df.loc[relabel, "Closure Type"] = "Virtual Contracts"
    return df


def _virtual_closure_dates() -> tuple:
    """
    Earliest virtual-closure date per Contract No and per Fac 12, read from the
    'D List' sheets kept in the LGD folder:
      * 'Receipt*Discount*.xlsb'  -> D List (CLOSURE_TYPE / CLOSE_DATE)
      * 'Write Backs 2024-25*.xlsx' -> D List (Settlement Type / CLOSE_DATE)
    Only virtual types count ('Virtual Contracts', 'Virtual Closure',
    'Virtually Closed', 'VC 31.03.2025'). A contract is re-stamped in every batch
    it is still >= the DPD threshold, so the EARLIEST date is the one it first
    became virtual - the same date the LGD sheet carries for older Virtual
    Contract rows (their real ERP closures come later and are not used).
    Returns (by_contract, by_fac12) dicts of pd.Timestamp.
    """
    sources = (
        [(p, "CLOSURE_TYPE", "pyxlsb") for p in glob.glob(os.path.join(INPUTS_DIR, "Receipt*Discount*.xlsb"))]
        + [(p, "Settlement Type", "openpyxl") for p in glob.glob(os.path.join(INPUTS_DIR, "Write Backs 2024-25*.xlsx"))]
    )
    by_contract: dict = {}
    by_fac12: dict = {}
    for path, type_col, engine in sources:
        try:
            d = pd.read_excel(path, sheet_name="D List", engine=engine)
        except Exception as exc:
            print(f"   D List not read from {os.path.basename(path)}: {exc}")
            continue
        need = {"CONTRACT_NO", "Fac 12", type_col, "CLOSE_DATE"}
        if not need <= set(d.columns):
            print(f"   {os.path.basename(path)} D List skipped — missing columns {sorted(need - set(d.columns))}")
            continue
        d = d[d[type_col].astype(str).str.match(r"(?i)^\s*(virtual|vc\b)")]
        dates = d["CLOSE_DATE"].apply(xlsb_serial_to_date)
        n0 = len(by_contract)
        for key_col, target in (("CONTRACT_NO", by_contract), ("Fac 12", by_fac12)):
            for k, dt in zip(d[key_col].astype(str).str.strip(), dates):
                if pd.notna(dt) and (k not in target or dt < target[k]):
                    target[k] = dt
        print(f"   {os.path.basename(path)} D List: {len(d):,} virtual rows, {len(by_contract) - n0:,} new contracts")
    return by_contract, by_fac12


def fill_virtual_contract_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Fill the blank Closed Date on 'Virtual Contracts' rows.

    Those rows were written before the xlsb_serial_to_date string fix, when the
    Portfolio DATE text was silently dropped, so Closed Date is blank, Closed Year is
    the "2025/26" fallback and D Gap (=Closed Date - First Default Date) is a large
    negative number. append_to_lgd_sheet skips any Fac 12 already in the sheet, so
    nothing else ever corrects them.

    Date used, in order:
      1. the earliest virtual-closure date in the 'D List' sheets (see
         _virtual_closure_dates), matched on Contract No, then on Fac 12;
      2. otherwise this run's date (sys.argv[1], DD/MM/YYYY) - the date FDDW stamps
         on a Virtual Contract first seen in this run. Listed in the log.
    Closed Year is re-derived from the date used. Only blank cells are touched, so
    re-running is harmless. Closed Date is written as "YYYY/MM/DD" text, as
    build_lgd does, which Excel converts to a date.
    """
    is_vc = df["Closure Type"].astype(str).str.strip().eq("Virtual Contracts")
    blank = df["Closed Date"].isna() | df["Closed Date"].astype(str).str.strip().eq("")
    fix = is_vc & blank
    n = int(fix.sum())
    print(f"Virtual Contracts: {int(is_vc.sum()):,} rows, {n:,} with a blank Closed Date")
    if n == 0:
        return df

    by_contract, by_fac12 = _virtual_closure_dates()
    found = pd.to_datetime(df["CONTRACT_NO"].astype(str).str.strip().map(by_contract))
    found = found.fillna(pd.to_datetime(df["Fac 12"].astype(str).str.strip().map(by_fac12)))
    use_found = fix & found.notna()
    no_source = fix & found.isna()

    run_dt = None
    if len(sys.argv) > 1:
        try:
            run_dt = datetime.strptime(sys.argv[1].strip(), "%d/%m/%Y")
        except ValueError:
            pass

    df = df.copy()
    df["Closed Date"] = df["Closed Date"].astype(object)
    df["Closed Year"] = df["Closed Year"].astype(object)
    df.loc[use_found, "Closed Date"] = found[use_found].dt.strftime("%Y/%m/%d")
    df.loc[use_found, "Closed Year"] = found[use_found].apply(to_fiscal_year)
    print(f"   {int(use_found.sum()):,} rows dated from the D Lists "
          f"({found[use_found].min():%d-%m-%Y} .. {found[use_found].max():%d-%m-%Y})"
          if use_found.any() else "   no dates found in the D Lists")

    if no_source.any():
        if run_dt is None:
            print(f"   WARNING: {int(no_source.sum()):,} rows have no D List date and no run date "
                  f"was supplied — left blank")
        else:
            df.loc[no_source, "Closed Date"] = run_dt.strftime("%Y/%m/%d")
            df.loc[no_source, "Closed Year"] = to_fiscal_year(run_dt)
            print(f"   {int(no_source.sum()):,} rows not in any D List — used run date "
                  f"{run_dt:%d-%m-%Y}: {df.loc[no_source, 'CONTRACT_NO'].astype(str).str.strip().tolist()}")
    return df


def apply_writebacks_to_lgd_final() -> None:
    """
    1. Read Consolidated Receipts by AFIN.xlsx (header row 3) → condf.
    2. Read 'LGD * Final.xlsb' LGD sheet → df_lgd_new (skiprows=1).
    3. Replace df_lgd_new.DCF Receipts with condf.Final Receipts
       where condf.Fac 12 == df_lgd_new.Fac 12.
    4. Add DC column from writeoff 31-08-2025.xlsx:
       Yes if Fac 12 found in writoffdf.Facility 12, else No.
    5. Remove rows where DC == Yes.
    6. Clear A3:O{last_row} in LGD Final LGD sheet, write df_lgd_new
       back at A3 without headers.
    """
    # ── Find LGD Final file ───────────────────────────────────────────────────
    final_matches = glob.glob(os.path.join(INPUTS_DIR, "LGD * Final.xlsb"))
    if not final_matches:
        raise FileNotFoundError(f"No file matching 'LGD * Final.xlsb' found in {INPUTS_DIR}")
    lgd_final_path = final_matches[0]

    # ── Read Consolidated Receipts by AFIN.xlsx ───────────────────────────────
    con_path = os.path.join(ROOT_DIR, "Consolidated Receipts by AFIN.xlsx")
    print(f"Reading Consolidated Receipts by AFIN: {con_path}")
    condf = pd.read_excel(con_path, header=2)
    condf["Fac 12"] = condf["Fac 12"].astype(str).str.strip()
    con_lookup = condf.set_index("Fac 12")["Final Receipts"].to_dict()
    print(f"condf: {len(condf)} rows, {len(con_lookup)} unique Fac 12 values")

    # ── Read LGD Final LGD sheet ──────────────────────────────────────────────
    print(f"Reading LGD Final LGD sheet: {lgd_final_path}")
    df_lgd_new = pd.read_excel(lgd_final_path, sheet_name="LGD", skiprows=1, engine="pyxlsb")
    df_lgd_new["Fac 12"] = df_lgd_new["Fac 12"].astype(str).str.strip()
    print(f"df_lgd_new: {len(df_lgd_new)} rows, columns: {df_lgd_new.columns.tolist()}")

    # ── Read receipt details.xlsx → receiptdfnew ─────────────────────────────
    receipt_path = os.path.join(ROOT_DIR, "receipt details.xlsx")
    print(f"Reading receipt details: {receipt_path}")
    receiptdfnew = pd.read_excel(receipt_path, engine="openpyxl")
    receiptdfnew["Facility 12"] = receiptdfnew["Facility 12"].astype(str).str.strip()
    receiptdfnew["RSM_CONNO"]   = receiptdfnew["RSM_CONNO"].astype(str).str.strip()
    print(f"receiptdfnew: {len(receiptdfnew)} rows")

    # # ── Find unmatched Facility 12s (in receiptdfnew but not in df_lgd_new) ──
    # existing_fac12 = set(df_lgd_new["Fac 12"].unique())
    # unmatched = receiptdfnew[~receiptdfnew["Facility 12"].isin(existing_fac12)].copy()
    # print(f"Unmatched receipt rows (Facility 12 not in df_lgd_new): {len(unmatched)}")

    # if not unmatched.empty:
    #     # Aggregate by Facility 12: keep first RSM_CONNO/Default Date, sum RSM_STLDAMT
    #     unmatched_agg = (
    #         unmatched.groupby("Facility 12", as_index=False)
    #         .agg(
    #             RSM_CONNO=("RSM_CONNO", "first"),
    #             RSM_STLDAMT=("RSM_STLDAMT", "sum"),
    #             Default_Date=("Default Date", "first"),
    #         )
    #     )

    #     new_rows = pd.DataFrame()
    #     new_rows["CONTRACT_NO"]                  = unmatched_agg["RSM_CONNO"]
    #     new_rows["Fac 12"]                       = unmatched_agg["Facility 12"]
    #     new_rows["Closed Date"]                  = ""
    #     new_rows["Closed Year"]                  = "2025/26"
    #     new_rows["First Default Date"]           = pd.to_datetime(
    #         unmatched_agg["Default_Date"], errors="coerce"
    #     ).dt.strftime("%Y/%m/%d")
    #     new_rows["Outstanding at First Default"] = unmatched_agg["RSM_STLDAMT"]

    #     df_lgd_new = pd.concat([df_lgd_new, new_rows], ignore_index=True)
    #     print(f"Added {len(new_rows)} unmatched receipt rows — df_lgd_new now {len(df_lgd_new)} rows")

    # # ── Read Portfolio file → portfdf, map PD CATEGORY by CONTRACT_NO ────────
    # portf_matches = glob.glob(os.path.join(INPUTS_DIR, "1. Portfolio_4_ECL*.xlsb"))
    # if portf_matches:
    #     portf_path = portf_matches[0]
    #     print(f"Reading Portfolio sheet from: {portf_path}")
    #     portfdf = pd.read_excel(portf_path, sheet_name="Portfolio", skiprows=2, engine="pyxlsb")
    #     portfdf["CONTRACT NO"] = portfdf["CONTRACT NO"].astype(str).str.strip()
    #     portf_lookup = (
    #         portfdf.drop_duplicates(subset=["CONTRACT NO"])
    #         .set_index("CONTRACT NO")["PD CATEGORY"]
    #     )
    #     df_lgd_new["CONTRACT_NO"] = df_lgd_new["CONTRACT_NO"].astype(str).str.strip()
    #     df_lgd_new["PD CATEGORY"] = df_lgd_new["CONTRACT_NO"].map(portf_lookup)
    #     matched_pd = df_lgd_new["PD CATEGORY"].notna().sum()
    #     print(f"PD CATEGORY mapped for {matched_pd}/{len(df_lgd_new)} rows from Portfolio")
    # else:
    #     print("Warning: No Portfolio_4_ECL*.xlsb file found — PD CATEGORY not mapped")

    # ── Replace DCF Receipts with Final Receipts from condf ───────────────────
    df_lgd_new["DCF Receipts"] = df_lgd_new["Fac 12"].map(con_lookup)
    matched = df_lgd_new["DCF Receipts"].notna().sum()
    print(f"DCF Receipts replaced from condf for {matched} matched contracts")

    # ── Add DC column from writeoff file ─────────────────────────────────────
    writeoff_path = os.path.join(ROOT_DIR, f"{writeoff_output_name()}.xlsx")
    if not os.path.exists(writeoff_path):          # fall back to the newest writeoff file present
        _found = sorted(glob.glob(os.path.join(ROOT_DIR, "writeoff *.xlsx")), key=os.path.getmtime)
        if _found:
            writeoff_path = _found[-1]
    print(f"Reading writeoff file: {writeoff_path}")
    writoffdf = pd.read_excel(writeoff_path)
    writoffdf["Facility 12"] = writoffdf["Facility 12"].astype(str).str.strip()
    writeoff_fac12 = set(writoffdf["Facility 12"].dropna().unique())

    df_lgd_new["DC"] = df_lgd_new["Fac 12"].apply(
        lambda x: "Yes" if x in writeoff_fac12 else "No"
    )
    dc_yes = (df_lgd_new["DC"] == "Yes").sum()
    print(f"DC column added — {dc_yes} Yes, {len(df_lgd_new) - dc_yes} No")

    # ── Remove DC == Yes rows ─────────────────────────────────────────────────
    df_lgd_new = df_lgd_new[df_lgd_new["DC"] != "Yes"].reset_index(drop=True)
    print(f"Removed {dc_yes} DC rows — {len(df_lgd_new)} rows remaining")

    # ── Single 'Virtual Contracts' type; backfill their blank Closed Dates ────
    df_lgd_new = normalise_virtual_contract_type(df_lgd_new)
    df_lgd_new = fill_virtual_contract_dates(df_lgd_new)

    # ── Keep only contracts closed on or before the run date ─────────────────
    df_lgd_new = drop_closed_after_run_date(df_lgd_new, "Closed Date", "LGD sheet")

    # ── Write df_lgd_new back to LGD Final ───────────────────────────────────
    with ExcelPortfolioAutomation(workbook_path=lgd_final_path, visible=False) as automation:
        sheet = automation.workbook.sheets["LGD"]

        last_row = sheet.used_range.last_cell.row
        num_cols = max(len(df_lgd_new.columns), 15)  # at least A:O (15 cols)
        clear_range = sheet.range((3, 1), (last_row, num_cols))
        clear_range.clear_contents()
        print(f"Cleared LGD Final LGD sheet rows 3:{last_row}, cols 1:{num_cols}")

        def _clean(v):
            try:
                if v is None or pd.isna(v):
                    return None
            except TypeError:
                pass
            return v

        rows_to_write = [
            [_clean(v) for v in row]
            for row in df_lgd_new.values.tolist()
        ]

        sheet.range("A3").value = rows_to_write
        print(f"Written {len(rows_to_write)} rows to LGD Final LGD sheet (A3, no header)")

        automation.workbook.save()

    print("LGD Final updated and saved.")


# ==========================
# DCF RECEIPTS SHEET — PERIOD SNAPSHOT COLUMN
# ==========================
DCF_SNAPSHOT_MONTHS = (3, 9)   # the sheet gets a column only for March / September runs


def update_dcf_receipts_sheet(df_receipt: pd.DataFrame) -> None:
    """
    Add (or refresh) this run's "<Mon YYYY> Additions" column in the
    'DCF Receipts' sheet of 'LGD * Final.xlsb', following the sheet's own pattern:

      * Every period column is a facility's CUMULATIVE DCF of receipts up to that
        period end — not a slice of the period.
      * 'Receipts to Consider' = MAX of the snapshot columns (never a sum, since
        the snapshots are cumulative and would otherwise double count).
      * 'Final Receipts' = SUM('Receipts to Consider' : 'Yard Adjustment').

    The new column is inserted immediately before 'Receipts to Consider'. Excel
    shifts 'Final Receipts' (=SUM(H4:K4)) and the row-1 totals by itself, and that
    is verified below before anything is saved. The 'Receipts to Consider' MAX
    formula is rewritten to include the new column, because Excel does not extend
    a reference when a column is inserted at its edge. Facilities present in the
    receipts data but not yet in the sheet are appended with the same formulas.
    Re-running the same period refreshes the column in place rather than
    inserting a second one.

    Receipts dated after the run date (sys.argv[1], DD/MM/YYYY) are ignored.
    Runs for any month other than those in DCF_SNAPSHOT_MONTHS leave the sheet alone.
    """
    if len(sys.argv) < 2:
        print("update_dcf_receipts_sheet: no run date argument supplied — skipping.")
        return
    try:
        run_dt = datetime.strptime(sys.argv[1].strip(), "%d/%m/%Y")
    except ValueError:
        print(f"update_dcf_receipts_sheet: could not parse run date '{sys.argv[1]}' "
              f"as DD/MM/YYYY — skipping.")
        return

    if run_dt.month not in DCF_SNAPSHOT_MONTHS:
        print(f"update_dcf_receipts_sheet: {run_dt:%b %Y} is not a half-year snapshot month "
              f"(months {DCF_SNAPSHOT_MONTHS}) — no column added.")
        return

    # Existing headers read "Sep 2024 Additions" / "March 2025 Additions"
    month_label = "March" if run_dt.month == 3 else run_dt.strftime("%b")
    new_header = f"{month_label} {run_dt.year} Additions"

    # ── Cumulative DCF per Facility 12 up to and including the run date ──────
    rec = df_receipt[["Facility 12", "RSM_STLDDATE", "DCF"]].copy()
    rec["Facility 12"] = rec["Facility 12"].astype(str).str.strip()
    rec["DCF"]         = pd.to_numeric(rec["DCF"], errors="coerce")
    rec["_d"]          = pd.to_datetime(rec["RSM_STLDDATE"], errors="coerce")
    after_cutoff       = rec["_d"] > run_dt
    print(f"\nDCF Receipts sheet — '{new_header}' (cut-off {run_dt:%d-%m-%Y}): "
          f"ignoring {int(after_cutoff.sum()):,} receipt rows dated after the run date")
    rec = rec[~after_cutoff]
    cum = rec.groupby("Facility 12")["DCF"].sum()
    cum = cum[cum.fillna(0) != 0]
    print(f"   {len(cum):,} facilities with cumulative DCF <> 0, total {cum.sum():,.0f}")

    matches = glob.glob(os.path.join(INPUTS_DIR, "LGD * Final.xlsb"))
    if not matches:
        raise FileNotFoundError(f"No file matching 'LGD * Final.xlsb' found in {INPUTS_DIR}")
    file_path = matches[0]
    print(f"   Workbook: {file_path}")

    HEADER_ROW, FIRST_ROW, TOTALS_LIMIT = 3, 4, 264744   # row-1 totals are =SUM(x4:x264744)

    with ExcelPortfolioAutomation(workbook_path=file_path, visible=False) as automation:
        ws = automation.workbook.sheets["DCF Receipts"]

        def headers() -> list:
            h = ws.range((HEADER_ROW, 1)).expand("right").value
            h = h if isinstance(h, list) else [h]
            return [str(x).strip() if x is not None else "" for x in h]

        col = xw.utils.col_name   # 1-based index -> column letters

        hdr = headers()
        for needed in ("Fac 12", "Dup Check", "Status", "Final Receipts",
                       "Receipts Until March 2024", "Receipts to Consider", "Yard Adjustment"):
            if needed not in hdr:
                raise RuntimeError(f"'DCF Receipts' header '{needed}' not found in row {HEADER_ROW}: {hdr}")

        base_i = hdr.index("Receipts Until March 2024") + 1     # first snapshot column
        inserted = False
        if new_header in hdr:
            new_i = hdr.index(new_header) + 1
            print(f"   Column '{new_header}' already present ({col(new_i)}) — refreshing in place")
        else:
            insert_at = hdr.index("Receipts to Consider") + 1
            ws.api.Columns(insert_at).Insert()                  # old 'Receipts to Consider' moves right
            new_i = insert_at
            ws.range((HEADER_ROW, new_i)).value = new_header
            inserted = True
            print(f"   Inserted '{new_header}' as column {col(new_i)}")

        hdr   = headers()
        rtc_i = hdr.index("Receipts to Consider") + 1
        fin_i = hdr.index("Final Receipts") + 1
        yard_i = hdr.index("Yard Adjustment") + 1
        ncols = len(hdr)

        # ── Verify Excel shifted 'Final Receipts' = SUM(RtC : Yard) as expected ─
        expected_fin = f"=SUM({col(rtc_i)}{FIRST_ROW}:{col(yard_i)}{FIRST_ROW})"
        actual_fin   = str(ws.range(f"{col(fin_i)}{FIRST_ROW}").formula).replace("$", "")
        if actual_fin != expected_fin:
            raise RuntimeError(
                f"'Final Receipts' formula is {actual_fin!r}, expected {expected_fin!r} after the "
                f"column insert — aborting without saving so the sheet is not corrupted."
            )

        last_row = ws.range(f"A{ws.api.Rows.Count}").end("up").row
        fac_vals = ws.range(f"A{FIRST_ROW}:A{last_row}").value
        fac_vals = fac_vals if isinstance(fac_vals, list) else [fac_vals]
        facs     = [str(v).strip() if v is not None else "" for v in fac_vals]
        existing = set(facs)

        # ── New column values for the facilities already in the sheet ────────
        ws.range((FIRST_ROW, new_i), (last_row, new_i)).value = [[float(cum.get(f, 0.0))] for f in facs]
        updated = sum(1 for f in facs if f in cum.index)

        # ── 'Receipts to Consider' = MAX over every snapshot column ──────────
        snap_cols  = [col(i) for i in range(base_i, new_i + 1)]
        max_formula = "=MAX(" + ",".join(f"{c}{FIRST_ROW}" for c in snap_cols) + ")"
        rtc_l = col(rtc_i)
        ws.range(f"{rtc_l}{FIRST_ROW}:{rtc_l}{last_row}").formula = max_formula
        print(f"   'Receipts to Consider' ({rtc_l}) set to {max_formula} on rows {FIRST_ROW}:{last_row}")

        # ── Append facilities that have receipts but no row yet ──────────────
        new_facs = sorted(f for f in cum.index if f not in existing)
        end_row = last_row
        if new_facs:
            n = len(new_facs)
            start, end_row = last_row + 1, last_row + n
            if end_row > TOTALS_LIMIT:
                print(f"   WARNING: rows now extend past {TOTALS_LIMIT}; the row-1 =SUM(x4:x{TOTALS_LIMIT}) "
                      f"totals will not include rows beyond it")
            last_l = col(ncols)
            # Copy row 4 (formulas + formats) down, then overwrite the constants.
            ws.range(f"A{FIRST_ROW}:{last_l}{FIRST_ROW}").copy(ws.range(f"A{start}:{last_l}{end_row}"))

            def put(ci: int, values: list) -> None:
                ws.range((start, ci), (end_row, ci)).value = [[v] for v in values]

            put(1, new_facs)
            put(hdr.index("Dup Check") + 1, [1] * n)
            put(hdr.index("Status") + 1, ["Y"] * n)
            for ci in range(base_i, new_i):                     # earlier snapshot columns
                put(ci, [0] * n)
            put(new_i, [float(cum[f]) for f in new_facs])
            for name in ("Write Backs Until March 2024", "Write Backs 2024/25"):
                if name in hdr:
                    put(hdr.index(name) + 1, [0] * n)
            print(f"   Appended {n:,} new facilities at rows {start}:{end_row}")

        # ── Row-1 total, number format and width for the new column ──────────
        if inserted:
            ws.range(f"{col(new_i - 1)}1").copy(ws.range(f"{col(new_i)}1"))
        ws.range((FIRST_ROW, new_i), (end_row, new_i)).number_format = \
            ws.range((FIRST_ROW, new_i - 1)).number_format
        ws.api.Columns(new_i).ColumnWidth = ws.api.Columns(new_i - 1).ColumnWidth

        print(f"   Existing rows given a value: {updated:,} | total in column: {cum.sum():,.0f}")
        print(f"   Final Receipts {col(fin_i)}{FIRST_ROW}: {ws.range(f'{col(fin_i)}{FIRST_ROW}').formula} | "
              f"Receipts to Consider {rtc_l}{FIRST_ROW}: {ws.range(f'{rtc_l}{FIRST_ROW}').formula}")
        automation.workbook.save()
    print("   DCF Receipts sheet updated and saved.")


# ==========================
# RENAME LGD FINAL TO RUN MONTH
# ==========================
def rename_lgd_final_to_run_date() -> None:
    """
    Rename 'LGD * Final.xlsb' to 'LGD {Mon} {YYYY} Final.xlsb' based on the
    run date passed on the command line (sys.argv[1], DD/MM/YYYY — this is
    how ifrs.py invokes this script). Nothing else renames this file, so
    without this it silently keeps whatever name it was first given,
    regardless of which month's data actually ends up inside it.
    """
    if len(sys.argv) < 2:
        print("No run date argument supplied — skipping LGD Final rename.")
        return
    try:
        run_dt = datetime.strptime(sys.argv[1].strip(), "%d/%m/%Y")
    except ValueError:
        print(f"Could not parse run date '{sys.argv[1]}' as DD/MM/YYYY — skipping LGD Final rename.")
        return

    matches = glob.glob(os.path.join(INPUTS_DIR, "LGD * Final.xlsb"))
    if not matches:
        print("No file matching 'LGD * Final.xlsb' found — skipping LGD Final rename.")
        return
    current_path = matches[0]

    new_name = f"LGD {run_dt.strftime('%b %Y')} Final.xlsb"
    new_path = os.path.join(INPUTS_DIR, new_name)

    if os.path.abspath(current_path) == os.path.abspath(new_path):
        print(f"LGD Final already named '{new_name}' — no rename needed.")
        return

    try:
        if os.path.exists(new_path):
            os.remove(new_path)
        os.rename(current_path, new_path)
        print(f"Renamed LGD Final: {os.path.basename(current_path)} -> {new_name}")
    except Exception as exc:
        print(f"WARNING: LGD Final rename failed: {exc}")


# ==========================
# RUN
# ==========================
def run():
    input_file = find_input_file()
    print(f"Input file: {input_file}")

    (df_closed,) = load_dataframes(input_file)
    df_closed = df_closed.reset_index(drop=True)

    df_refinance      = load_refinance_dpd_dataframe()
    df_ledger_details = load_ledger_detail_dataframe()
    print(df_ledger_details.head())

    # Load JOBHD receipt files and compute DCF (uses original Fac 12 before remapping)
    df_receipt = load_receipt_dataframe()

    # Populate H:M formulas in 'Part1' sheet, then load the resulting data.
    # Not yet merged into any downstream calculation.
    receipt_discount_file = find_receipt_discount_file()
    populate_receipts_discount_formulas(receipt_discount_file)
    df_receipts_discount = load_recipts_discounts_dataframe(receipt_discount_file)

    # Process each writebacks df independently — compute DCF and aggregate by Facility 12
    writebacks = load_writebacks_dataframe()
    wb1562_agg = process_writebacks_dcf(writebacks, df_closed, writeoff_output_name())


    wb130 = load_writebacks130_dataframe()
    df = pd.read_excel(find_LGD_file(),sheet_name='LGD',skiprows=1, engine="pyxlsb")
    wb130_agg = process_writebacks130_dcf(wb130, df, writebacks_output_name())

    # ── Remove DC contracts from df_receipt using writebacks Facility 12s ──────
    # (writebacks = load_writebacks_dataframe() output, already loaded above)
    dc_fac12_from_wb = set(writebacks["Facility 12"].astype(str).unique())
    before_len = len(df_receipt)
    df_receipt = df_receipt[~df_receipt["Facility 12"].isin(dc_fac12_from_wb)].reset_index(drop=True)
    print(f"Removed {before_len - len(df_receipt)} DC-matched rows from receipts — {len(df_receipt)} rows remaining")

    df_receipt = compute_dcf_from_receipts(df_receipt, df_closed)

    receipt_details_path = os.path.join(ROOT_DIR, "receipt details.xlsx")
    df_receipt.to_excel(receipt_details_path, index=False)
    print(f"Receipt details saved to: {receipt_details_path} ({len(df_receipt)} rows)")

    dcf_pivot = save_dcf_pivot(df_receipt)

    # Apply CONTRACT_NO remapping from Refinance DPD Report directly to df_closed
    lookup = dict(zip(df_refinance["OLD_CON_NO"], df_refinance["NEW_CON_NO"]))
    df_closed["CONTRACT_NO"] = df_closed["CONTRACT_NO"].map(lookup).fillna(df_closed["CONTRACT_NO"])
    df_closed["Fac 12"] = df_closed["CONTRACT_NO"].astype(str).str[:12]
    print("CONTRACT_NO and Fac 12 updated in df_closed via Refinance DPD Report")

    # Build LGD dataframe (called last, right before saving)
    df_lgd = build_lgd(df_closed)

    # Map receipt DCF pivot to df_lgd, taking max with any existing DCF Receipts value
    dcf_pivot_lookup = dcf_pivot.set_index("Facility 12")["Sum of DCF"]
    receipt_mapped = df_lgd["Fac 12"].map(dcf_pivot_lookup).fillna(0)
    existing_dcf = df_lgd["DCF Receipts"].fillna(0) if "DCF Receipts" in df_lgd.columns else pd.Series(0, index=df_lgd.index)
    df_lgd["DCF Receipts"] = existing_dcf.combine(receipt_mapped, max)
    print("DCF Receipts mapped from receipt pivot (max with existing) to all contracts in df_lgd")

    # Mark and remove DC contracts using writebacks Facility 12s
    df_lgd["DC"] = df_lgd["Fac 12"].isin(dc_fac12_from_wb).map({True: "Yes", False: "No"})
    dc_yes = (df_lgd["DC"] == "Yes").sum()
    print(f"DC column added — {dc_yes} matched (Yes), {len(df_lgd) - dc_yes} not matched (No)")
    df_lgd = df_lgd[df_lgd["DC"] != "Yes"].reset_index(drop=True)
    print(f"Removed {dc_yes} DC rows from df_lgd — {len(df_lgd)} rows remaining")

    # Add back Yard amounts to DCF Receipts
    YARD_DF = load_yard_dataframe()
    yard_lookup = YARD_DF.groupby("Facility 12")["Amount"].sum()
    yard_matched = df_lgd["Fac 12"].isin(yard_lookup.index).sum()
    df_lgd["DCF Receipts"] += df_lgd["Fac 12"].map(yard_lookup).fillna(0)
    print(f"Yard amounts added back to DCF Receipts for {yard_matched} matched contracts")

    # Add back Writebacks130 DCF to DCF Receipts (wb130_agg is a DCF sum Series keyed by Facility 12)
    wb130_matched = df_lgd["Fac 12"].isin(wb130_agg.index).sum()
    df_lgd["DCF Receipts"] += df_lgd["Fac 12"].map(wb130_agg).fillna(0)
    print(f"Writebacks130 DCF added back to DCF Receipts for {wb130_matched} matched contracts")

    # ── DEBUG: data quality report ──────────────────────────────────────────
    total = len(df_lgd)

    dcf_nonzero = (df_lgd["DCF Receipts"] != 0).sum()
    print(f"\n[DEBUG] DCF Receipts — non-zero: {dcf_nonzero}, zero/unmatched: {total - dcf_nonzero}")
    zero_fac12 = df_lgd.loc[df_lgd["DCF Receipts"] == 0, "Fac 12"].unique()
    if len(zero_fac12) > 0:
        print(f"   Sample Fac 12s with zero DCF Receipts (up to 10): {list(zero_fac12[:10])}")

    fdd_missing = df_lgd["First Default Date"].isna().sum()
    fdd_missing += (df_lgd["First Default Date"] == "NaT").sum()
    print(f"[DEBUG] First Default Date — missing: {fdd_missing} / {total}")
    if fdd_missing > 0:
        sample = df_lgd.loc[
            df_lgd["First Default Date"].isna() | (df_lgd["First Default Date"] == "NaT"),
            ["CONTRACT_NO", "Fac 12", "Closure Type"]
        ].head(10)
        print(f"   Sample rows with missing First Default Date:\n{sample.to_string(index=False)}")

    lgd_missing = df_lgd["LGD Type"].isna().sum()
    print(f"[DEBUG] LGD Type        — missing: {lgd_missing} / {total}")
    if lgd_missing:
        _miss = df_lgd.loc[df_lgd["LGD Type"].isna(), "CONTRACT_NO"].astype(str).tolist()
        print(f"[WARN] LGD Type not found (Portfolio file / PD data file) for {len(_miss):,} contracts: "
              f"{', '.join(_miss[:50])}" + (" ..." if len(_miss) > 50 else ""))
    if lgd_missing > 0:
        sample = df_lgd.loc[df_lgd["LGD Type"].isna(), ["CONTRACT_NO", "Fac 12", "Closure Type"]].head(10)
        print(f"   Sample rows with missing LGD Type:\n{sample.to_string(index=False)}")

    ofd_missing = df_lgd["Outstanding at First Default"].isna().sum()
    print(f"[DEBUG] Outstanding     — missing: {ofd_missing} / {total}")
    print(f"── END DEBUG ───────────────────────────────────────────────────────\n")

    # # Normalise LGD Type values before appending
    # df_lgd["LGD Type"] = df_lgd["LGD Type"].replace({
    #     "Mini Truck":       "MINI T",
    #     "Other Equipment":  "O.EQP",
    #     "Revolving Loans":  "RL",
    #     "SME Loans":        "SME",
    #     "Term Loans":       "TL",
    # })

    # Append df_lgd to LGD sheet in LGD file
    append_to_lgd_sheet(df_lgd)

    # Normalise LGD Type values directly in the LGD sheet before refreshing pivot
    normalise_lgd_type_in_sheet()
    # Run update_afin_receipts.py to refresh Consolidated Receipts by AFIN.xlsx
    afin_script = os.path.join(CURRENT_DIR, "update_afin_receipts.py")
    print(f"\nRunning update_afin_receipts.py: {afin_script}")
    # Pass the run date (DD/MM/YYYY) through so it can ignore receipts dated after it.
    _afin_cmd = [sys.executable, afin_script] + ([sys.argv[1]] if len(sys.argv) > 1 else [])
    subprocess.run(_afin_cmd, check=True)
    print("update_afin_receipts.py completed.\n")

    # Apply writebacks DCF to LGD Final file
    apply_writebacks_to_lgd_final()

    # Add this run's "<Mon YYYY> Additions" snapshot column to the DCF Receipts sheet.
    # That sheet is a reference sheet (the LGD sheet's DCF Receipts values come from the
    # AFIN file), and a failure inside leaves the workbook unsaved and unchanged, so it
    # must not abort the rest of the run (K1:M1 formulas, pivot refresh, rename).
    try:
        update_dcf_receipts_sheet(df_receipt)
    except Exception as exc:
        print("\n" + "!" * 70)
        print(f"WARNING: DCF Receipts sheet was NOT updated: {exc}")
        print("         The rest of the run continues; fix and re-run that step separately.")
        print("!" * 70 + "\n")

    # Copy K1:M1 formulas down to all data rows in LGD * Final.xlsb LGD sheet
    final_matches = glob.glob(os.path.join(INPUTS_DIR, "LGD * Final.xlsb"))
    if not final_matches:
        raise FileNotFoundError(f"No file matching 'LGD * Final.xlsb' found in {INPUTS_DIR}")
    lgd_final_path = final_matches[0]
    print(f"Copying K1:M1 formulas in LGD Final: {lgd_final_path}")
    app = xw.App(visible=False)
    try:
        wb_final = app.books.open(lgd_final_path)
        sheet_final = wb_final.sheets["LGD"]
        total_rows = sheet_final.api.Rows.Count
        last_row = sheet_final.range(f"B{total_rows}").end("up").row
        print(f"   Last data row: {last_row}")
        sheet_final.range("K1:M1").copy(sheet_final.range(f"K3:M{last_row}"))
        print(f"   K1:M1 formulas copied to K3:M{last_row}")
        wb_final.save()
        wb_final.close()
        print("   LGD Final formulas populated, workbook saved and closed.")
    finally:
        app.quit()

    # Refresh PivotTable2 with full updated data range
    refresh_lgd_pivot()

    # Rename LGD Final to reflect the month actually just processed
    rename_lgd_final_to_run_date()

    print(df_lgd)

if __name__ == "__main__":
    run()
