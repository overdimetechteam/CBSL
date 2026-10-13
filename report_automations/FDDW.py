"""
LGD Closed List + First Default Date DPD90 Processing Script
=============================================================
FIXED VERSION — Key changes from original:
  FIX 1 : normalize_columns() — all DataFrame column names forced to
           UPPER_SNAKE_CASE so "CLIENT DPD", "Client DPD", "CLIENT_DPD"
           all resolve to "CLIENT_DPD" consistently.
  FIX 2 : XLOOKUP looks up by Fac 12 (col B / $B:$B) in both Closed List
           and year sheets — lookup_cell uses col B of Closed List (FAC12),
           lookup_arr uses $B:$B in year sheets.
  FIX 3 : build_nested_xlookup() updated: lookup_arr uses $B:$B in year
           sheets; lookup_cell uses col B of Closed List (FAC12).
  FIX 4 : generate_fac12() now also strips trailing whitespace before
           checking for trailing zero.
  FIX 5 : Sort column resolution uses fuzzy name matching (case-insensitive
           substring) so it works even if headers differ slightly.
  FIX 6 : Added diagnose_xlookup() helper to quickly spot key mismatches.
  FIX 7 : step7 is skipped with a warning if step6 raises an exception.
  FIX 8 : step7 no longer converts working_copy.xlsx back to .xlsb.
           Output is saved permanently to BASE_DIR (LGD folder).
  FIX 12: SaveAs uses EnsureDispatch + keyword-only Filename/FileFormat
           args to avoid "Member not found" COM error on certain Excel builds.
  FIX 9 : First Default Date column (col F) in Closed List is formatted
           as DD/MM/YYYY so values display as dates, not serial numbers.
  FIX 10: Output filename is dynamic:
           "First Default Date Working {DD-MM-YYYY} to {DD-MM-YYYY}.xlsx"
           reflecting the actual report date range.
  FIX 11: step7 is called ONCE at the end of main() for the full date
           range — no per-month output files are produced.

PROCESS:
  STEP 1 : Read Closure Category Wise Report
  STEP 2 : Read Portfolio (DPD >= 366, Virtual Contracts)
  STEP 3 : Append combined rows to 'Closed List' sheet (cols A-D)
  STEP 4 : Read Portfolio (DPD > 90) for year-sheet data
  STEP 5 : Build 10-column DPD90 rows
  STEP 6 : Append rows to calendar-year sheet (create if needed), sort
  STEP 7 : Write nested XLOOKUP formulas into Closed List cols E-I,
           format col F as DD/MM/YYYY, yellow-highlight date range rows.
           Single output file saved to BASE_DIR (LGD folder).

Usage:
  python lgd_processor.py 2025-08-31
  python lgd_processor.py 2025-04-30 2025-08-31
  python lgd_processor.py 2025-08-31 --force      (bypass date order check)
  python lgd_processor.py 2025-08-31 --diagnose   (print key diagnostics)
"""

import os
import sys
import glob
import re
import calendar
import pandas as pd
import xlwings as xw
from datetime import datetime

# Force UTF-8 output on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


class _TeeWriter:
    """Write to both the original stdout and a log file."""
    def __init__(self, original, log_file_path: str):
        self._orig = original
        self._fh   = open(log_file_path, "a", encoding="utf-8", errors="replace")

    def write(self, data):
        self._orig.write(data)
        self._fh.write(data)
        self._fh.flush()

    def flush(self):
        self._orig.flush()
        self._fh.flush()

    def fileno(self):
        return self._orig.fileno()


def _setup_log_tee() -> str:
    log_dir = os.environ.get("IFRS_LGD_LOG_DIR", "").strip()
    if not log_dir:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "logs", f"run_{ts}"
        )
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "FDDW.log")
    sys.stdout = _TeeWriter(sys.stdout, log_file)
    sys.stderr = _TeeWriter(sys.stderr, log_file)
    print(f"Log file: {log_file}")
    return log_file


_setup_log_tee()

# ============================================================================
# BASE PATHS
# LGD_BASE_FOLDER env var (set by ifrs.py) points to the working folder.
# Falls back to hardcoded path when running standalone.
# ============================================================================
BASE_ROOT = r"c:\CBSL\Script"
sys.path.insert(0, BASE_ROOT)
from BasicExcelFunctionsClass import ExcelPortfolioAutomation

def _find_lgd_working_folder() -> str:
    """Return the firstDefault subfolder inside the single date folder in working/LGD."""
    base = os.path.join(BASE_ROOT, "working", "LGD")
    if os.path.isdir(base):
        subfolders = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d))]
        if subfolders:
            return os.path.join(base, subfolders[0], "firstDefault")
    return base

_lgd_env = os.environ.get("LGD_BASE_FOLDER", "").strip()
BASE_DIR  = os.path.join(_lgd_env, "firstDefault") if _lgd_env else _find_lgd_working_folder()

LAST_COMPLETE_DATE_FILE = os.path.join(BASE_DIR, "Last_Complete_Date.txt")

def _find_target_file() -> str:
    """Find the First Default Date Working *.xlsb file in BASE_DIR."""
    for ext in ("xlsb", "xlsx"):
        matches = [
            f for f in glob.glob(os.path.join(BASE_DIR, f"First Default Date Working*.{ext}"))
            if not os.path.basename(f).startswith("~$")
        ]
        if matches:
            return sorted(matches, key=os.path.getmtime)[-1]
    raise FileNotFoundError(
        f"No 'First Default Date Working' file found in:\n  {BASE_DIR}"
    )

TARGET_FILE = _find_target_file()

# ============================================================================
# CLOSED LIST SHEET LAYOUT
# ============================================================================
TARGET_SHEET    = "Closed List"
VIRTUAL_CONTRACT_MIN_DPD = 366   # Portfolio CLIENT_DPD at/above which a contract is a Virtual Contract
HEADER_ROW_CL   = 2      # row 2 = column headers
DATA_START_CL   = 3      # row 3 = first data row

# Closed List column letters
CL_COL_CONTRACT   = "A"  # CONTRACT_NO
CL_COL_FAC12      = "B"  # Fac 12            <-- XLOOKUP lookup key
CL_COL_CLOSURE    = "C"  # CLOSURE_TYPE
CL_COL_CLOSE_DATE = "D"  # CLOSE_DATE
CL_COL_INT_RATE   = "E"  # Int Rate          (direct value from All Contracts Sector Report)
CL_COL_FDD        = "F"  # First Default Date (XLOOKUP -> year sheet col C)
CL_COL_OUTSTAND   = "G"  # Outstanding       (XLOOKUP -> year sheet col F)
CL_COL_DPD        = "H"  # DPD at First Default (XLOOKUP -> year sheet col I)
CL_COL_LGD        = "I"  # LGD Category      (XLOOKUP -> year sheet col J)

# ============================================================================
# YEAR SHEET LAYOUT  (1-based column numbers)
# ============================================================================
HEADER_ROW_YS         = 1
DATA_START_YS         = 2
COMBINED_EARLY_SHEET  = "2016-2021"
INDIVIDUAL_YEAR_START = 2022

YS_COL_LOAN_NO    = 1   # A  Loan number  / CONTRACT_NO
YS_COL_FAC12      = 2   # B  Fac 12       <-- XLOOKUP lookup array ($B:$B)
YS_COL_FDD        = 3   # C  First Default Date
YS_COL_CLIENT_NO  = 4   # D  Client No
YS_COL_GRANT_DATE = 5   # E  Grant date
YS_COL_OUTSTAND   = 6   # F  Outstanding at First Default
YS_COL_INT_RATE   = 7   # G  Int Rate
YS_COL_MONTHS     = 8   # H  Number of months in arrears
YS_COL_DPD        = 9   # I  DPD at First Default
YS_COL_LGD        = 10  # J  LGD Category

DEFAULT_YS_HEADERS = [
    "Loan number", "Fac 12", "First Default Date", "Client No",
    "Grant date", "Outstanding at First Default", "Int Rate",
    "no of months in arreas", "DPD at First Default", "LGD Category",
]

# All Contracts Sector Report — source for Interest Rate (col E in Closed List)
ALL_CONTRACTS_SHEET = "All Contracts Sector Report - I"  # Excel truncates sheet name at 31 chars

# Closed List formula columns:  Closed List col -> (year sheet return col, description)
# Lookup key: Closed List col B (Fac 12) matched against year sheet $B:$B (Fac 12)
# NOTE: cols E and I are intentionally excluded — they are filled directly:
#   E (Int Rate)    <- All Contracts Sector Report, by CONTRACT_NO
#   I (LGD Category)<- Portfolio file "Portfolio" sheet, PD CATEGORY by CONTRACT_NO;
#                      if not there: outputs/PD/<run date>/*PD_data* (Portfolio_1/_2, col C -> col E)
FORMULA_COLUMN_MAP = {
    "F": ("C", "First Default Date"),
    "G": ("F", "Outstanding at First Default"),
    "H": ("I", "DPD at First Default"),
}

# ============================================================================
# FIX 9 — Date number format for First Default Date column (col F)
# ============================================================================
FDD_NUMBER_FORMAT = "DD/MM/YYYY"   # applied to Closed List col F after formula write

# ============================================================================
# FIX 13 — PRE-RUN: Clear Closed List data rows (preserve header row 2)
# ============================================================================
def clear_closed_list_data() -> None:
    """
    Clear all data rows (row DATA_START_CL onward) in the 'Closed List'
    sheet of TARGET_FILE, leaving the header row (HEADER_ROW_CL = row 2)
    and any formulas it contains fully intact.

    Uses openpyxl directly — no COM, no Excel process required — so it
    works regardless of file size or Windows COM dispatch quirks.

    Strategy: iterate over every row >= DATA_START_CL in the 'Closed List'
    worksheet and set each cell value to None for columns A-I (cols 1-9).
    openpyxl preserves cell formatting (fill, font, number_format) while
    clearing only the value/formula, which is exactly what 'clear contents'
    does in Excel.  The file is then saved back in-place.
    """
    section("PRE-RUN : Clearing Closed List data rows (preserving row 2 header)")

    import openpyxl as _openpyxl

    LAST_COL_NUM = 9   # cols A-I

    print(f"  Opening {os.path.basename(TARGET_FILE)} with openpyxl ...")
    wb_opy = _openpyxl.load_workbook(TARGET_FILE)
    ws_opy = wb_opy[TARGET_SHEET]

    max_row = ws_opy.max_row
    if max_row < DATA_START_CL:
        print(f"  No data rows found below header (max_row = {max_row}). Nothing to clear.")
        wb_opy.save(TARGET_FILE)
        wb_opy.close()
        print(f"  PRE-RUN complete. Header row {HEADER_ROW_CL} preserved.")
        return

    print(f"  Clearing rows {DATA_START_CL}:{max_row} cols A-I in '{TARGET_SHEET}' ...")

    cleared = 0
    for row in ws_opy.iter_rows(
        min_row=DATA_START_CL, max_row=max_row,
        min_col=1,             max_col=LAST_COL_NUM,
    ):
        for cell in row:
            cell.value = None
        cleared += 1

    print(f"  Cleared {cleared:,} rows x {LAST_COL_NUM} cols (A{DATA_START_CL}:I{max_row}).")

    wb_opy.save(TARGET_FILE)
    wb_opy.close()
    print(f"  PRE-RUN complete. Header row {HEADER_ROW_CL} preserved.")


# ============================================================================
# FIX 10 — Dynamic output filename helper
# ============================================================================
def build_output_path(from_date: datetime, to_date: datetime) -> str:
    """
    Return the full path for the output .xlsx file.

    Filename pattern:
      First Default Date Working {DD-MM-YYYY} to {DD-MM-YYYY}.xlsx

    Example (single month):
      First Default Date Working 30-04-2025 to 30-04-2025.xlsx

    Example (multi-month):
      First Default Date Working 30-04-2025 to 31-08-2025.xlsx
    """
    from_str  = from_date.strftime("%d-%m-%Y")
    to_str    = to_date.strftime("%d-%m-%Y")
    filename  = f"First Default Date Working {from_str} to {to_str}.xlsx"
    return os.path.join(BASE_DIR, filename)

# ============================================================================
# FIX 1 — COLUMN NAME NORMALISER
# ============================================================================
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise every DataFrame column name to UPPER_SNAKE_CASE.
    'Client DPD'  -> 'CLIENT_DPD'
    'CONTRACT NO' -> 'CONTRACT_NO'
    'PD CATEGORY' -> 'PD_CATEGORY'
    This makes all downstream .get() / [] access case- and space-insensitive.
    """
    df.columns = [
        str(c).strip().upper().replace(" ", "_") if c else ""
        for c in df.columns
    ]
    return df

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================
def section(title: str) -> None:
    print("\n" + "=" * 65)
    print(f"  {title}")
    print("=" * 65)

def col_num_to_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def generate_fac12(contract_no) -> str:
    """
    FIX 4: Strip ALL trailing whitespace before checking for trailing '0'.
    'ABC123 '  -> 'ABC123'   (was incorrectly kept as-is)
    'ABC1230'  -> 'ABC123'
    'ABC1231'  -> 'ABC1231'
    """
    s = str(contract_no).strip()
    return s[:-1] if s.endswith("0") else s

def safe_numeric(value) -> float:
    if value is None:
        return 0.0
    try:
        v = float(value)
        return 0.0 if v != v else v
    except (TypeError, ValueError):
        return 0.0

def get_month_end_date(report_date: datetime) -> datetime:
    last_day = calendar.monthrange(report_date.year, report_date.month)[1]
    return datetime(report_date.year, report_date.month, last_day)

def get_year_sheet_name(report_date: datetime) -> str:
    return str(report_date.year)

# ============================================================================
# LAST_COMPLETE_DATE.TXT
# ============================================================================
def read_last_complete_date() -> datetime:
    if not os.path.exists(LAST_COMPLETE_DATE_FILE):
        raise FileNotFoundError(
            f"Last_Complete_Date.txt not found:\n  {LAST_COMPLETE_DATE_FILE}\n"
            f"  Option 1: Create it with content: YYYY-MM-DD  (e.g. 2025-03-31)\n"
            f"  Option 2: Provide two date arguments:  python FDDW.py 2025-03-31 2025-04-30"
        )
    with open(LAST_COMPLETE_DATE_FILE, "rb") as f:
        raw_bytes = f.read()
    print(f"  [DEBUG] Raw bytes in Last_Complete_Date.txt : {raw_bytes!r}")
    try:
        raw = raw_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        raw = raw_bytes.decode("latin-1")
    for ch in ("\r\n", "\r", "\n", "\x00", "\u00a0", "\ufeff"):
        raw = raw.replace(ch, "")
    raw = "".join(c for c in raw if c.isprintable()).strip()
    print(f"  [DEBUG] Cleaned string : '{raw}'")
    if not raw:
        raise ValueError(f"Last_Complete_Date.txt is empty. Raw bytes: {raw_bytes!r}")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    for sep in ("-", "/", "."):
        if sep not in raw:
            continue
        parts = [p.strip() for p in raw.split(sep)]
        if len(parts) != 3:
            continue
        try:
            if len(parts[0]) == 4:
                year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
            else:
                day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
            return datetime(year, month, day)
        except (ValueError, TypeError):
            continue
    raise ValueError(
        f"Cannot parse date from Last_Complete_Date.txt.\n"
        f"  Raw bytes  : {raw_bytes!r}\n"
        f"  Cleaned    : '{raw}'\n"
        f"  Supported  : YYYY-MM-DD, YYYY/MM/DD, DD-MM-YYYY, DD/MM/YYYY"
    )

def update_last_complete_date(date: datetime) -> None:
    with open(LAST_COMPLETE_DATE_FILE, "w", encoding="utf-8", newline="") as f:
        f.write(date.strftime("%Y-%m-%d"))
    print(f"  Last_Complete_Date.txt updated -> {date.strftime('%Y-%m-%d')}")

# ============================================================================
# FILE LOCATORS
# ============================================================================
def get_closure_file() -> str:
    candidates = []
    for ext in ("xlsx", "xlsb", "xls"):
        pattern = os.path.join(BASE_DIR, f"Closure Category Wise Report - Internal *.{ext}")
        candidates.extend(
            f for f in glob.glob(pattern)
            if not os.path.basename(f).startswith("~$")
        )
    if not candidates:
        raise FileNotFoundError(
            f"No Closure Category Wise Report found in:\n  {BASE_DIR}"
        )
    candidates.sort(key=os.path.getmtime, reverse=True)
    if len(candidates) > 1:
        print(f"  Multiple closure files -- selecting most recently modified:")
        for i, fp in enumerate(candidates):
            print(f"    {os.path.basename(fp)}{'  <-- SELECTED' if i == 0 else ''}")
    return candidates[0]

def get_all_contracts_file() -> str:
    candidates = []
    for ext in ("xlsx", "xlsb", "xls"):
        pattern = os.path.join(BASE_DIR, f"All Contracts Sector Report - Internal *.{ext}")
        candidates.extend(
            f for f in glob.glob(pattern)
            if not os.path.basename(f).startswith("~$")
        )
    if not candidates:
        raise FileNotFoundError(
            f"No 'All Contracts Sector Report - Internal' file found in:\n  {BASE_DIR}"
        )
    candidates.sort(key=os.path.getmtime, reverse=True)
    if len(candidates) > 1:
        print(f"  Multiple All Contracts files -- selecting most recently modified:")
        for i, fp in enumerate(candidates):
            print(f"    {os.path.basename(fp)}{'  <-- SELECTED' if i == 0 else ''}")
    return candidates[0]

def get_portfolio_file(report_date: datetime) -> str:
    date_str = report_date.strftime("%d.%m.%Y")
    for ext in ("xlsb", "xlsx", "xls"):
        matches = [
            f for f in glob.glob(os.path.join(BASE_DIR, f"*Portfolio_4_ECL-{date_str}*.{ext}"))
            if not os.path.basename(f).startswith("~$")
        ]
        if matches:
            return matches[0]
    raise FileNotFoundError(
        f"Portfolio file not found for {report_date.strftime('%Y-%m-%d')}.\n"
        f"  Looked in : {BASE_DIR}\n"
        f"  Expected  : 1 .Portfolio_4_ECL-{date_str}*.xlsb"
    )

# ============================================================================
# LGD CATEGORY FALLBACK: PD data file (outputs/PD/<run date folder>/*PD_data*.xlsb)
# ============================================================================
PD_OUTPUTS_DIR = os.environ.get("LGD_PD_OUTPUTS_DIR", "").strip() or os.path.join(BASE_ROOT, "outputs", "PD")
PD_DATA_SHEETS = ("Portfolio_1", "Portfolio_2")
PD_COL_MONTH, PD_COL_FAC12, PD_COL_CATEGORY = 0, 2, 4     # columns A, C, E (0-based)


def find_pd_data_file(report_date: datetime):
    """
    outputs/PD/<DD-MM-YYYY> or <DD-MM-YYYY>(n) for the report date - the highest version
    number wins - then the newest '*PD_data*' workbook inside it (the date in the file
    name changes between runs, so only 'PD_data' is used to recognise it).
    Returns the path, or None (with a warning) if it cannot be found.
    """
    date_str = report_date.strftime("%d-%m-%Y")
    if not os.path.isdir(PD_OUTPUTS_DIR):
        print(f"  [WARN] PD outputs folder not found: {PD_OUTPUTS_DIR}")
        return None
    pattern = re.compile(r"^" + re.escape(date_str) + r"(?:\((\d+)\))?$")
    candidates = []
    for name in os.listdir(PD_OUTPUTS_DIR):
        full = os.path.join(PD_OUTPUTS_DIR, name)
        m = pattern.match(name)
        if m and os.path.isdir(full):
            candidates.append((int(m.group(1) or 0), os.path.getmtime(full), full))
    if not candidates:
        print(f"  [WARN] No PD folder for {date_str} in {PD_OUTPUTS_DIR}")
        return None
    folder = sorted(candidates)[-1][2]
    files = [
        f for ext in ("xlsb", "xlsx")
        for f in glob.glob(os.path.join(folder, f"*PD_data*.{ext}"))
        if not os.path.basename(f).startswith("~$")
    ]
    if not files:
        print(f"  [WARN] No '*PD_data*' workbook in {folder}")
        return None
    return sorted(files, key=os.path.getmtime)[-1]


def _pd_month_to_ts(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, (int, float)):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(value))
    return pd.to_datetime(value, dayfirst=True, errors="coerce")


def load_pd_category_map(report_date: datetime) -> dict:
    """
    {Fac 12 (PD file column C) -> PD_CATEGORY (column E)} from the Portfolio_1 and
    Portfolio_2 sheets of the PD data file. If a facility appears in several months,
    the latest month (column A) wins.
    """
    path = find_pd_data_file(report_date)
    if not path:
        return {}
    print(f"  PD data file for LGD fallback : {path}")
    engine = "pyxlsb" if path.lower().endswith(".xlsb") else "openpyxl"
    frames = []
    for sheet in PD_DATA_SHEETS:
        try:
            raw = pd.read_excel(path, sheet_name=sheet, engine=engine, header=0)
        except Exception as exc:
            print(f"  [WARN] Could not read sheet '{sheet}' of {os.path.basename(path)}: {exc}")
            continue
        part = pd.DataFrame({
            "month": raw.iloc[:, PD_COL_MONTH].apply(_pd_month_to_ts),
            "fac12": raw.iloc[:, PD_COL_FAC12].astype(str).str.strip().str.upper(),
            "cat":   raw.iloc[:, PD_COL_CATEGORY],
        })
        part = part[part["cat"].notna() & (part["fac12"] != "") & (part["fac12"] != "NAN")]
        part["cat"] = part["cat"].astype(str).str.strip()
        part = part[part["cat"] != ""]
        print(f"    sheet {sheet:12s}: {len(part):,} rows with a Fac 12 and PD_CATEGORY")
        frames.append(part)
    if not frames:
        return {}
    allrows = pd.concat(frames, ignore_index=True).sort_values("month", kind="stable", na_position="first")
    allrows = allrows.drop_duplicates(subset="fac12", keep="last")
    print(f"    PD category map: {len(allrows):,} facilities")
    return dict(zip(allrows["fac12"], allrows["cat"]))


# ============================================================================
# FIX 2 & 3 — FORMULA BUILDER: lookup by FAC12 ($B:$B)
# ============================================================================
def get_formula_year_sheets(wb) -> list:
    """Return ordered list of year sheet names present in the workbook."""
    existing = {s.name for s in wb.sheets}
    sheets = []
    if COMBINED_EARLY_SHEET in existing:
        sheets.append(COMBINED_EARLY_SHEET)
    max_year = datetime.today().year + 2
    for yr in range(INDIVIDUAL_YEAR_START, max_year):
        name = str(yr)
        if name in existing:
            sheets.append(name)
    return sheets

def build_nested_xlookup(lookup_cell: str, year_sheets: list, return_col: str) -> str:
    """
    Build a nested VLOOKUP formula (named build_nested_xlookup for call-site
    compatibility — kept as VLOOKUP, not XLOOKUP, because the Excel used to
    recalculate this file doesn't support XLOOKUP: it writes formulas as
    _xludf.XLOOKUP and caches every cell as #NAME?, which pandas then reads
    back as NaN and silently drops every row downstream. VLOOKUP works on
    every Excel version. Requires return_col to sit to the right of col B
    in every year sheet (true for C/F/I here).

    FIX 2 & 3: The lookup array is $B:$B (Fac 12) in every year sheet.
    The lookup_cell is col B of the Closed List (also Fac 12).

    Generated formula example (Int Rate, row 3):
      =IFERROR(VLOOKUP(B3,'2016-2021'!$B:$G,6,FALSE),
         IFERROR(VLOOKUP(B3,'2022'!$B:$G,6,FALSE),
           VLOOKUP(B3,'2025'!$B:$G,6,FALSE)))
    """
    if not year_sheets:
        return ""
    from openpyxl.utils import column_index_from_string
    col_index = column_index_from_string(return_col) - column_index_from_string("B") + 1
    core = None
    for sheet in reversed(year_sheets):
        lookup_range = f"'{sheet}'!$B:${return_col}"
        vlookup_expr = f"VLOOKUP({lookup_cell},{lookup_range},{col_index},FALSE)"
        if core is None:
            core = vlookup_expr
        else:
            core = f"IFERROR({vlookup_expr},{core})"
    return "=" + core

# ============================================================================
# STEP 1: Read Closure Category Wise Report
# ============================================================================
# The ERP export started returning raw ERP field names instead of the
# descriptive aliases this pipeline was built against — same data, same
# column positions, just different header text at the source. Renaming
# is a no-op for files that still use the old aliases (rename() ignores
# keys that aren't present), so this works against either file version.
CLOSURE_REPORT_COLUMN_ALIASES = {
    "CLO_OLDCONNO":    "CONTRACT_NO",
    "CTP_DESC":        "CLOSURE_TYPE",
    "CLO_DATE":        "CLOSE_DATE",
    # EQT_CATEGORY is only ever "A"/"N" (an active/inactive flag), not a loan
    # type, despite the name — EQT_SUBCATEGORY holds the real values (e.g.
    # "3W LEASING", "MINI TRUCK LEASING").
    "EQT_SUBCATEGORY": "CATEGORY",
}


def step1_closure_data(closure_file: str) -> pd.DataFrame:
    section("STEP 1 : Closure Category Wise Report")
    print(f"  File : {os.path.basename(closure_file)}")
    ext    = os.path.splitext(closure_file)[1].lower()
    engine = "pyxlsb" if ext == ".xlsb" else "openpyxl"

    filtered_df = pd.read_excel(closure_file, header=0, engine=engine)
    filtered_df = filtered_df.rename(columns=CLOSURE_REPORT_COLUMN_ALIASES)
    # Drop every Facility Cancellation row. The source spells it several ways
    # ('Facility Cancellation', 'Facility Cancellation & Refund' and the typo
    # 'Facility Cancallation & Refund'), so match on the pattern, not exact text.
    is_cancel = filtered_df["CLOSURE_TYPE"].astype(str).str.contains(
        r"(?i)facility\s+canc[ae]l+ation", regex=True, na=False
    )
    print(f"  Facility Cancellation rows excluded : {int(is_cancel.sum()):,}")
    filtered_df = filtered_df[~is_cancel].copy()

    result = filtered_df[["CONTRACT_NO", "CLOSURE_TYPE", "CLOSE_DATE"]].copy()
    result = result[
        result["CONTRACT_NO"].notna() &
        (result["CONTRACT_NO"].astype(str).str.strip() != "")
    ].reset_index(drop=True)

    print(f"  Records after exclusions : {len(result):,}")
    print(f"  STEP 1 complete.")
    return result

# ============================================================================
# STEP 2: Portfolio — Virtual Contracts (Client DPD >= VIRTUAL_CONTRACT_MIN_DPD)
# ============================================================================
def step2_virtual_contracts(portfolio_file: str) -> pd.DataFrame:
    section(f"STEP 2 : Portfolio - Virtual Contracts (Client DPD >= {VIRTUAL_CONTRACT_MIN_DPD})")
    print(f"  File : {os.path.basename(portfolio_file)}")

    app_xw = xw.App(visible=False)
    try:
        wb_xw    = app_xw.books.open(portfolio_file)
        ws       = wb_xw.sheets["Portfolio"]
        last_row = ws.range("E1048576").end("up").row
        print(f"  Portfolio last row (col E) : {last_row}")
        headers = ws.range("A3:BZ3").value
        data    = ws.range(f"A4:BZ{last_row}").value
        wb_xw.close()
    finally:
        app_xw.quit()

    df = pd.DataFrame(data, columns=headers)
    df = normalize_columns(df)   # FIX 1

    df["CLIENT_DPD"] = pd.to_numeric(df["CLIENT_DPD"], errors="coerce")
    filtered = df[df["CLIENT_DPD"] >= VIRTUAL_CONTRACT_MIN_DPD].copy()
    print(f"  Rows with CLIENT_DPD >= {VIRTUAL_CONTRACT_MIN_DPD} : {len(filtered):,}")

    result = pd.DataFrame({
        "CONTRACT_NO" : filtered["CONTRACT_NO"].astype(str).str.strip(),
        "CLOSURE_TYPE": "Virtual Contracts",
        "CLOSE_DATE"  : filtered["DATE"],
    })
    result = result[
        result["CONTRACT_NO"].notna() & (result["CONTRACT_NO"] != "")
    ].reset_index(drop=True)

    print(f"  STEP 2 complete -> {len(result):,} Virtual Contract records.")
    return result

# ============================================================================
# STEP 3: Append combined results to Closed List (Columns A-D)
# ============================================================================
def step3_append_closed_list(combined_df: pd.DataFrame) -> None:
    section("STEP 3 : Append records to Closed List (cols A-D)")

    rows_to_write = []
    for _, row in combined_df.iterrows():
        contract_no  = str(row["CONTRACT_NO"]).strip()
        fac12        = generate_fac12(contract_no)
        closure_type = str(row["CLOSURE_TYPE"]).strip()
        close_date   = row["CLOSE_DATE"]
        rows_to_write.append([contract_no, fac12, closure_type, close_date])

    with ExcelPortfolioAutomation(TARGET_FILE, visible=False) as auto:
        ws = auto.workbook.sheets[TARGET_SHEET]
        last_used_row = ws.range("A1048576").end("up").row
        if last_used_row < (DATA_START_CL - 1):
            last_used_row = DATA_START_CL - 1
        append_start = last_used_row + 1

        print(f"  Closed List last used row   : {last_used_row}")
        print(f"  Appending new rows from row : {append_start}")
        print(f"  New rows to append          : {len(rows_to_write):,}")

        auto.write_data_to_range(
            sheet_name = TARGET_SHEET,
            start_cell = f"A{append_start}",
            data       = rows_to_write,
        )
        auto.workbook.save()
    print(f"  STEP 3 complete -> {len(rows_to_write):,} rows appended.")

# ============================================================================
# STEP 4: Portfolio — Filter Client DPD > 90
# ============================================================================
def step4_read_dpd90(portfolio_file: str) -> pd.DataFrame:
    section("STEP 4 : Portfolio - Filter Client DPD > 90")
    print(f"  File : {os.path.basename(portfolio_file)}")

    app_xw = xw.App(visible=False)
    try:
        wb_xw    = app_xw.books.open(portfolio_file)
        ws       = wb_xw.sheets["Portfolio"]
        last_row = ws.range("E1048576").end("up").row
        headers  = ws.range("A3:BT3").value
        data     = ws.range(f"A4:BT{last_row}").value
        wb_xw.close()
    finally:
        app_xw.quit()

    df = pd.DataFrame(data, columns=headers)
    df = normalize_columns(df)   # FIX 1

    df["CLIENT_DPD"] = pd.to_numeric(df["CLIENT_DPD"], errors="coerce")
    filtered = df[df["CLIENT_DPD"] > 90].copy().reset_index(drop=True)

    print(f"  Total portfolio rows       : {len(df):,}")
    print(f"  Rows with CLIENT_DPD > 90 : {len(filtered):,}")
    print(f"  STEP 4 complete.")
    return filtered

# ============================================================================
# STEP 5: Build 10-column DPD90 output rows
# ============================================================================
def step5_build_dpd90_rows(filtered_df: pd.DataFrame,
                            first_default_date: datetime) -> list:
    section("STEP 5 : Build DPD90 output rows")
    rows = []

    for _, row in filtered_df.iterrows():
        contract_no   = str(row.get("CONTRACT_NO", "")).strip()
        fac12         = generate_fac12(contract_no)
        client_no     = row.get("CLIENT_CODE")
        grant_date    = row.get("ACTIVATE_DATE")
        outstanding   = (
            safe_numeric(row.get("FUTURE_CAPITAL"))
            + safe_numeric(row.get("ARRCAP"))
            + safe_numeric(row.get("ARRINT"))
        )
        interest_rate  = row.get("INTEREST_RATE")
        client_dpd     = row.get("CLIENT_DPD")
        try:
            months_arrears = round(float(client_dpd) / 30, 2)
        except (TypeError, ValueError):
            months_arrears = None
        lgd_category  = row.get("PD_CATEGORY")

        rows.append([
            contract_no,        # A  Loan number
            fac12,              # B  Fac 12
            first_default_date, # C  First Default Date
            client_no,          # D  Client No
            grant_date,         # E  Grant date
            outstanding,        # F  Outstanding at First Default
            interest_rate,      # G  Int Rate
            months_arrears,     # H  Months in arrears
            client_dpd,         # I  DPD at First Default
            lgd_category,       # J  LGD Category
        ])

    print(f"  DPD90 output rows built : {len(rows):,}")
    print(f"  STEP 5 complete.")
    return rows

# ============================================================================
# STEP 6: Append DPD90 rows to year sheet, create if needed, then sort
# ============================================================================
def _year_row_key(loan_no, fac12, first_default_date):
    """Identity of a DPD90 year-sheet row: (loan number, Fac 12, month-end First Default Date)."""
    d = pd.to_datetime(first_default_date, errors="coerce")
    return (
        str(loan_no).strip(),
        str(fac12).strip(),
        None if pd.isna(d) else d.normalize(),
    )


def drop_rows_already_in_sheet(rows: list, existing_values: list) -> tuple:
    """
    Remove from `rows` (new DPD90 rows, columns A-C = loan no, Fac 12, First Default Date)
    those already present in the sheet (`existing_values`, same A-C layout).
    Steps 4-6 are purely additive, so re-processing a month (e.g. Last_Complete_Date.txt
    rewound) used to append the same month's rows again.
    Returns (new_rows, number_skipped).
    """
    existing = set()
    for r in existing_values or []:
        if r is None or len(r) < 3:
            continue
        existing.add(_year_row_key(r[0], r[1], r[2]))
    kept, seen = [], set()
    for row in rows:
        key = _year_row_key(row[YS_COL_LOAN_NO - 1], row[YS_COL_FAC12 - 1], row[YS_COL_FDD - 1])
        if key in existing or key in seen:
            continue
        seen.add(key)
        kept.append(row)
    return kept, len(rows) - len(kept)


def step6_append_year_sheet(rows: list, sheet_name: str) -> None:
    section(f"STEP 6 : Append {len(rows):,} rows to year sheet '{sheet_name}'")

    with ExcelPortfolioAutomation(TARGET_FILE, visible=False) as auto:
        wb             = auto.workbook
        existing_names = [s.name for s in wb.sheets]

        # -- Create year sheet if missing ----------------------------------
        if sheet_name not in existing_names:
            print(f"  Sheet '{sheet_name}' not found -- creating ...")
            header_values = None

            try:
                prev_year_name = str(int(sheet_name) - 1)
            except ValueError:
                prev_year_name = None

            for candidate in ([prev_year_name] if prev_year_name else []) + [COMBINED_EARLY_SHEET]:
                if candidate and candidate in existing_names:
                    prev_ws       = wb.sheets[candidate]
                    last_col_cell = prev_ws.range("A1").end("right")
                    hdr           = prev_ws.range(f"A1:{last_col_cell.address}").value
                    header_values = hdr if isinstance(hdr, list) else [hdr]
                    print(f"  Headers from sheet '{candidate}' ({len(header_values)} cols)")
                    break

            if header_values is None:
                header_values = DEFAULT_YS_HEADERS
                print(f"  No predecessor -- using default headers ({len(header_values)} cols)")

            wb.sheets.add(name=sheet_name)
            new_ws = wb.sheets[sheet_name]
            new_ws.range("A1").value = [header_values]
            append_start_row = DATA_START_YS
            print(f"  Sheet '{sheet_name}' created.")
        else:
            ws_ys            = wb.sheets[sheet_name]
            last_used        = ws_ys.range("A1048576").end("up").row

            # Sanity check: the last-row lookup once returned 1 on a sheet that really had
            # ~39,000 rows, so the new rows were written from row 2 on top of the data.
            # Refuse to write if the sheet's used range is far larger than the detected end.
            used_last = ws_ys.used_range.last_cell.row
            if last_used <= DATA_START_YS and used_last > DATA_START_YS + 100:
                raise RuntimeError(
                    f"Year sheet '{sheet_name}': last used row detected as {last_used} but the "
                    f"sheet's used range extends to row {used_last}. Refusing to append - the "
                    f"existing data would be overwritten."
                )

            # Skip rows already in the sheet (re-processing a month must not duplicate it).
            if last_used >= DATA_START_YS:
                existing_vals = ws_ys.range((DATA_START_YS, 1), (last_used, 3)).value
                if existing_vals and not isinstance(existing_vals[0], list):
                    existing_vals = [existing_vals]
                rows, skipped = drop_rows_already_in_sheet(rows, existing_vals)
                print(f"  Rows already present in '{sheet_name}' (skipped): {skipped:,}")
                if not rows:
                    print("  Nothing new to append - existing rows are still checked for duplicates below.")

            append_start_row = max(last_used + 1, DATA_START_YS)
            print(f"  Sheet '{sheet_name}' exists. Last used row : {last_used}")
            print(f"  New rows start from row                    : {append_start_row}")

        # -- Append rows ---------------------------------------------------
        if rows:
            auto.write_data_to_range(
                sheet_name = sheet_name,
                start_cell = f"A{append_start_row}",
                data       = rows,
            )
            print(f"  {len(rows):,} rows written to sheet '{sheet_name}'.")

        # -- Sort by Fac 12 ASC then First Default Date ASC ---------------
        ws_ys        = auto.workbook.sheets[sheet_name]
        last_row_now = ws_ys.range("A1048576").end("up").row

        if last_row_now >= DATA_START_YS:
            print(f"  Sorting sheet '{sheet_name}' ...")
            last_col_cell = ws_ys.range("A1").end("right")
            num_cols      = last_col_cell.column
            last_col_ltr  = col_num_to_letter(num_cols)

            header_row = ws_ys.range(
                f"A{HEADER_ROW_YS}:{last_col_ltr}{HEADER_ROW_YS}"
            ).value
            data_block = ws_ys.range(
                f"A{DATA_START_YS}:{last_col_ltr}{last_row_now}"
            ).value

            if data_block and not isinstance(data_block[0], list):
                data_block = [data_block]

            if data_block:
                df_sort = pd.DataFrame(data_block, columns=header_row)

                # FIX 5: fuzzy column name resolution
                fac12_col = next(
                    (c for c in df_sort.columns if "fac" in str(c).lower()),
                    df_sort.columns[YS_COL_FAC12 - 1],
                )
                fdd_col = next(
                    (c for c in df_sort.columns
                     if "default" in str(c).lower() and "first" in str(c).lower()),
                    df_sort.columns[YS_COL_FDD - 1],
                )

                # Remove rows that are identical in EVERY column (e.g. a month that was
                # processed more than once). Rows with the same loan / Fac 12 / date but
                # different values (outstanding, DPD ...) are different data and are kept.
                n_before = len(df_sort)
                df_sort = df_sort.drop_duplicates(keep="first")
                n_removed = n_before - len(df_sort)
                if n_removed:
                    print(f"  Removed {n_removed:,} duplicate rows already in '{sheet_name}'")

                df_sort = df_sort.sort_values(
                    by          = [fac12_col, fdd_col],
                    ascending   = [True, True],
                    na_position = "last",
                )
                ws_ys.range(f"A{DATA_START_YS}").value = df_sort.values.tolist()
                if n_removed:
                    # the sheet is now shorter: clear the leftover rows below the data
                    tail_first = DATA_START_YS + len(df_sort)
                    ws_ys.range(
                        f"A{tail_first}:{last_col_ltr}{last_row_now}"
                    ).clear_contents()
                print(f"  Sort complete ('{fac12_col}' ASC, '{fdd_col}' ASC).")

        auto.workbook.save()
    print(f"  STEP 6 complete -> {len(rows):,} rows appended, sorted, saved.")

# ============================================================================
# STEP 7: Update XLOOKUP formulas in Closed List (cols E-I)
#
# FIX 9  : Col F (First Default Date) cells are formatted as DD/MM/YYYY
#           via openpyxl number_format so Excel renders dates, not serials.
# FIX 10 : Output filename is built by build_output_path() and reflects
#           the actual from_date / to_date range.
# FIX 11 : This function is called ONCE at the end of main() — not once
#           per month — so only a single output file is ever produced.
# ============================================================================
def step7_update_formulas(from_date: datetime, to_date: datetime) -> None:
    """
    Write nested XLOOKUP formulas into Closed List cols F-H (no '@').
    Col E (Int Rate)      : direct value from All Contracts Sector Report by CONTRACT_NO.
    Col I (LGD Category)  : Portfolio sheet PD CATEGORY by CONTRACT_NO; contracts not found
                            there are looked up in the PD data file (see load_pd_category_map).
    Format col F (First Default Date) as DD/MM/YYYY.
    Yellow-highlight rows where CLOSE_DATE falls within [from_date, to_date].

    All changes are written directly to TARGET_FILE and saved in-place
    (equivalent to Ctrl+S). No new file or copy is created.
    """
    section("STEP 7 : Update VLOOKUP formulas in Closed List (cols E-I)")

    import openpyxl as _openpyxl
    from openpyxl.styles import PatternFill

    # ------------------------------------------------------------------ #
    # 1. Collect metadata via xlwings (read-only pass)                   #
    # ------------------------------------------------------------------ #
    print(f"  Opening TARGET_FILE to collect metadata ...")
    print(f"  File : {TARGET_FILE}")
    app_xw = xw.App(visible=False)
    app_xw.display_alerts = False
    try:
        wb_xw       = app_xw.books.open(TARGET_FILE)
        ws_cl_xw    = wb_xw.sheets[TARGET_SHEET]
        year_sheets = get_formula_year_sheets(wb_xw)
        last_row    = ws_cl_xw.range("A1048576").end("up").row
        wb_xw.close()
    finally:
        app_xw.quit()

    print(f"  Year sheets detected  : {year_sheets}")
    print(f"  Closed List last row  : {last_row}")

    if not year_sheets:
        print("  WARNING: No year sheets found -- skipping formula update.")
        return
    if last_row < DATA_START_CL:
        print("  Closed List has no data rows -- skipping formula update.")
        return

    total_rows = last_row - DATA_START_CL + 1
    print(f"  Data rows to update   : {total_rows:,}  (rows {DATA_START_CL}–{last_row})")

    # ------------------------------------------------------------------ #
    # 2. Build all formula strings per column                             #
    # ------------------------------------------------------------------ #
    col_formulas = {}
    for cl_col, (ys_return_col, _) in FORMULA_COLUMN_MAP.items():
        col_formulas[cl_col] = [
            build_nested_xlookup(f"B{r}", year_sheets, ys_return_col)
            for r in range(DATA_START_CL, last_row + 1)
        ]
    print("  Formula strings built.")

    # ------------------------------------------------------------------ #
    # 3. Open TARGET_FILE with openpyxl, write formulas + date format    #
    #    NOTE: TARGET_FILE is opened and saved in-place — no copy made.  #
    # ------------------------------------------------------------------ #
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    no_fill     = PatternFill(fill_type=None)
    range_start = from_date.date()
    range_end   = to_date.date()

    # ------------------------------------------------------------------ #
    # 3a. Build CONTRACT_NO -> INTEREST_RATE lookup from All Contracts   #
    # ------------------------------------------------------------------ #
    all_contracts_file = get_all_contracts_file()
    print(f"  Loading interest rates from : {os.path.basename(all_contracts_file)}")
    wb_ac  = _openpyxl.load_workbook(all_contracts_file, read_only=True, data_only=True)
    ws_ac  = wb_ac[ALL_CONTRACTS_SHEET]
    ac_rows = ws_ac.iter_rows(values_only=True)
    ac_headers = [str(h).strip().upper() if h is not None else "" for h in next(ac_rows)]
    ac_contract_col = ac_headers.index("CONTRACT_NO")
    ac_rate_col     = ac_headers.index("INTEREST_RATE")
    int_rate_map: dict = {}
    for ac_row in ac_rows:
        contract = ac_row[ac_contract_col]
        rate     = ac_row[ac_rate_col]
        if contract is not None:
            int_rate_map[str(contract).strip()] = rate
    wb_ac.close()
    print(f"  Interest rate map loaded    : {len(int_rate_map):,} contracts")

    # ------------------------------------------------------------------ #
    # 3b. Build CONTRACT_NO -> PD_CATEGORY lookup from Portfolio file     #
    # ------------------------------------------------------------------ #
    portfolio_file = get_portfolio_file(to_date)
    print(f"  Loading LGD (PD CATEGORY) from : {os.path.basename(portfolio_file)}")
    port_ext    = os.path.splitext(portfolio_file)[1].lower()
    port_engine = "pyxlsb" if port_ext == ".xlsb" else "openpyxl"
    port_df = pd.read_excel(
        portfolio_file, sheet_name="Portfolio",
        engine=port_engine, header=2, dtype=str,
    )
    port_df = normalize_columns(port_df)
    lgd_map: dict = {
        str(r["CONTRACT_NO"]).strip(): r.get("PD_CATEGORY")
        for _, r in port_df.iterrows()
        if r.get("CONTRACT_NO") is not None and str(r.get("CONTRACT_NO")).strip()
    }
    print(f"  LGD map loaded              : {len(lgd_map):,} contracts")

    # ------------------------------------------------------------------ #
    # 3c. (no closure-report fallback any more: contracts missing from the
    #     Portfolio file are looked up in the PD data file further down)
    # ------------------------------------------------------------------ #

    # ------------------------------------------------------------------ #
    # Open TARGET_FILE directly with openpyxl for in-place editing       #
    # ------------------------------------------------------------------ #
    print(f"  Opening TARGET_FILE with openpyxl for in-place editing ...")
    wb_opy = _openpyxl.load_workbook(TARGET_FILE)
    ws_opy = wb_opy[TARGET_SHEET]

    # ------------------------------------------------------------------ #
    # 3d. Write interest rates (col E) from All Contracts lookup          #
    # ------------------------------------------------------------------ #
    matched = 0
    for row_num in range(DATA_START_CL, last_row + 1):
        contract_no = ws_opy[f"{CL_COL_CONTRACT}{row_num}"].value
        rate = int_rate_map.get(str(contract_no).strip() if contract_no is not None else "")
        ws_opy[f"{CL_COL_INT_RATE}{row_num}"].value = rate
        if rate is not None:
            matched += 1
    print(f"  Col E (Int Rate)                         : {last_row - DATA_START_CL + 1:,} rows written "
          f"({matched:,} matched, {last_row - DATA_START_CL + 1 - matched:,} unmatched)")

    # ------------------------------------------------------------------ #
    # 3e. Write LGD Category (col I) - Portfolio file first, PD data file  #
    #     (outputs/PD/<run date>/ *PD_data*) for contracts not in it        #
    # ------------------------------------------------------------------ #
    lgd_rows = []                     # [excel row, contract key, category or None]
    for row_num in range(DATA_START_CL, last_row + 1):
        contract_no = ws_opy[f"{CL_COL_CONTRACT}{row_num}"].value
        key = str(contract_no).strip() if contract_no is not None else ""
        lgd_rows.append([row_num, key, lgd_map.get(key)])

    matched_lgd = sum(1 for r in lgd_rows if r[2] is not None)
    from_pd_file = 0
    missing_idx = [i for i, r in enumerate(lgd_rows) if r[2] is None and r[1]]
    if missing_idx:
        print(f"  {len(missing_idx):,} contracts not in the Portfolio file - looking them up in the PD data file ...")
        pd_map = load_pd_category_map(to_date)
        if pd_map:
            for i in missing_idx:
                key = lgd_rows[i][1]
                # PD file column C is the contract number without its last digit
                for cand in dict.fromkeys((key[:-1], key[:12], generate_fac12(key))):
                    cat = pd_map.get(cand.upper())
                    if cat is not None:
                        lgd_rows[i][2] = cat
                        from_pd_file += 1
                        break

    not_found = [r[1] for r in lgd_rows if r[2] is None and r[1]]
    for row_num, _key, lgd in lgd_rows:
        ws_opy[f"{CL_COL_LGD}{row_num}"].value = lgd
    print(f"  Col I (LGD Category)                     : {last_row - DATA_START_CL + 1:,} rows written "
          f"({matched_lgd:,} from portfolio, {from_pd_file:,} from PD data file, {len(not_found):,} not found)")
    if not_found:
        print(f"  [WARN] LGD Category not found in the Portfolio file or the PD data file for "
              f"{len(not_found):,} contracts (includes closed contracts with no First Default Date, "
              f"which the LGD report does not use). First 50: {', '.join(not_found[:50])}")
        log_dir = os.environ.get("IFRS_LGD_LOG_DIR", "").strip()
        if log_dir and os.path.isdir(log_dir):
            list_path = os.path.join(log_dir, "LGD_category_not_found.txt")
            with open(list_path, "w", encoding="utf-8") as fh:
                fh.write("\n".join(not_found))
            print(f"  [WARN] Full list written to: {list_path}")

    # ------------------------------------------------------------------ #
    # 3f. Write VLOOKUP formulas for cols F, G, H                         #
    # ------------------------------------------------------------------ #
    print("  Writing VLOOKUP formulas via openpyxl ...")
    for cl_col, (_, description) in FORMULA_COLUMN_MAP.items():
        formulas = col_formulas[cl_col]
        for idx, formula in enumerate(formulas):
            row_num    = DATA_START_CL + idx
            cell       = ws_opy[f"{cl_col}{row_num}"]
            cell.value = formula
            if cl_col == CL_COL_FDD:
                cell.number_format = FDD_NUMBER_FORMAT
        print(f"  Col {cl_col} ({description:<30}) : {len(formulas):,} formulas written"
              + (f"  [format: {FDD_NUMBER_FORMAT}]" if cl_col == CL_COL_FDD else ""))

    # ------------------------------------------------------------------ #
    # 4. Yellow-highlight rows where CLOSE_DATE (col D) is in range      #
    # ------------------------------------------------------------------ #
    print(f"  Applying yellow highlight for CLOSE_DATE in {range_start} – {range_end} ...")
    highlighted = 0
    for row_num in range(DATA_START_CL, last_row + 1):
        close_val = ws_opy.cell(row=row_num, column=4).value   # col D
        in_range  = False
        if close_val is not None:
            if hasattr(close_val, "date"):
                close_dt = close_val.date()
            else:
                try:
                    close_dt = datetime.strptime(str(close_val)[:10], "%Y-%m-%d").date()
                except ValueError:
                    close_dt = None
            if close_dt and range_start <= close_dt <= range_end:
                in_range = True
        fill = yellow_fill if in_range else no_fill
        for col in range(1, 10):   # cols A–I
            ws_opy.cell(row=row_num, column=col).fill = fill
        if in_range:
            highlighted += 1
    print(f"  Yellow highlighted rows : {highlighted:,}")

    # ------------------------------------------------------------------ #
    # 5. Save TARGET_FILE in-place (no new file)                         #
    # ------------------------------------------------------------------ #
    wb_opy.save(TARGET_FILE)
    print(f"  STEP 7 complete — {total_rows:,} rows x 5 columns updated.")
    print(f"  First Default Date column formatted as : {FDD_NUMBER_FORMAT}")
    print(f"  File saved in-place : {TARGET_FILE}")

    # ------------------------------------------------------------------ #
    # 6. Force a full recalculation + re-save via Excel COM              #
    #    openpyxl only writes formula TEXT — it never evaluates          #
    #    formulas, so cols F-H have no cached value until real Excel     #
    #    opens, recalculates, and saves the file. Without this step,     #
    #    anything reading the file with data_only=True (e.g. the LGD     #
    #    computation script) sees every formula cell as blank.           #
    # ------------------------------------------------------------------ #
    print("  Recalculating formulas via Excel COM ...")
    app_recalc = xw.App(visible=False)
    app_recalc.display_alerts = False
    try:
        wb_recalc = app_recalc.books.open(TARGET_FILE)
        app_recalc.api.CalculateFull()
        app_recalc.api.CalculateFull()
        wb_recalc.save()
        wb_recalc.close()
        print("  Formulas recalculated — cached values now saved in file.")
    finally:
        app_recalc.quit()

# ============================================================================
# DIAGNOSTIC HELPER  (run with --diagnose flag)
# ============================================================================
def diagnose_xlookup(first_n_rows: int = 10) -> None:
    """
    Print CONTRACT_NO (col A) and Fac 12 (col B) for the first N rows of the
    Closed List AND every year sheet, then show the resolved formula values.
    Run:  python lgd_processor.py 2025-08-18 --diagnose
    """
    section("DIAGNOSTIC : XLOOKUP key comparison")
    app_xw = xw.App(visible=False)
    try:
        wb    = app_xw.books.open(TARGET_FILE)
        ws_cl = wb.sheets[TARGET_SHEET]

        print(f"\n  Closed List (first {first_n_rows} data rows):")
        print(f"  {'Row':<5} {'CONTRACT_NO (A)':<35} {'Fac12 (B)':<30} "
              f"{'Int Rate (E)':<15} {'FDD (F)'}")
        for r in range(DATA_START_CL, DATA_START_CL + first_n_rows):
            a = ws_cl.range(f"A{r}").value
            b = ws_cl.range(f"B{r}").value
            e = ws_cl.range(f"E{r}").value
            f = ws_cl.range(f"F{r}").value
            if a is None:
                break
            print(f"  {r:<5} {str(a):<35} {str(b):<30} {str(e):<15} {str(f)}")

        for sheet_name in get_formula_year_sheets(wb):
            ws_ys  = wb.sheets[sheet_name]
            last_r = ws_ys.range("A1048576").end("up").row
            if last_r < DATA_START_YS:
                continue
            print(f"\n  Year sheet '{sheet_name}' (first {first_n_rows} data rows):")
            print(f"  {'Row':<5} {'LoanNo/CONTRACT_NO (A)':<35} {'Fac12 (B)':<30} "
                  f"{'Int Rate (G)'}")
            for r in range(DATA_START_YS, min(DATA_START_YS + first_n_rows, last_r + 1)):
                a = ws_ys.range(f"A{r}").value
                b = ws_ys.range(f"B{r}").value
                g = ws_ys.range(f"G{r}").value
                print(f"  {r:<5} {str(a):<35} {str(b):<30} {str(g)}")

        wb.close()
    finally:
        app_xw.quit()

# ============================================================================
# MONTH-END GENERATOR
# ============================================================================
def month_ends_between(from_date: datetime, to_date: datetime) -> list:
    """Return all month-end dates where month_end > from_date and <= to_date."""
    ends  = []
    year  = from_date.year
    month = from_date.month
    while True:
        month += 1
        if month > 12:
            month = 1
            year  += 1
        last_day  = calendar.monthrange(year, month)[1]
        month_end = datetime(year, month, last_day)
        if month_end > to_date:
            break
        ends.append(month_end)
    return ends

# ============================================================================
# SINGLE-MONTH PROCESSOR  (Steps 4-6 only — Step 7 runs once in main)
# ============================================================================
def process_month(to_date: datetime) -> bool:
    """
    Run Steps 4–6 for one month (to_date).

    FIX 11: Step 7 is intentionally NOT called here.  It is called exactly
    once at the end of main() over the full date range, producing a single
    output file rather than one per month.

    Returns True on success, False on failure.
    """
    first_default_date = get_month_end_date(to_date)
    year_sheet_name    = get_year_sheet_name(to_date)

    section(f"PROCESSING : {to_date.strftime('%Y-%m-%d')}")
    print(f"  First Default Date : {first_default_date.strftime('%Y-%m-%d')}")
    print(f"  Year sheet         : '{year_sheet_name}'")

    # Locate portfolio file for this month
    try:
        portfolio_file = get_portfolio_file(to_date)
        print(f"  Portfolio file     : {os.path.basename(portfolio_file)}")
    except FileNotFoundError as exc:
        print(f"\nERROR: {exc}")
        return False

    # Steps 4–6: build and append year-sheet rows
    dpd90_df = step4_read_dpd90(portfolio_file)
    step6_ok = False
    if dpd90_df.empty:
        print("\n  No records with CLIENT_DPD > 90 -- skipping Steps 5 & 6.")
        step6_ok = True
    else:
        dpd90_rows = step5_build_dpd90_rows(dpd90_df, first_default_date)
        try:
            step6_append_year_sheet(dpd90_rows, year_sheet_name)
            step6_ok = True
        except Exception as exc:
            print(f"\nERROR in Step 6: {exc}")

    if not step6_ok:
        return False

    # Update Last_Complete_Date.txt after each successful month
    update_last_complete_date(to_date)

    section(f"MONTH COMPLETE : {to_date.strftime('%Y-%m-%d')}")
    print(f"  First Default Date written : {first_default_date.strftime('%Y-%m-%d')}")
    print(f"  Year sheet used            : '{year_sheet_name}'")
    return True

# ============================================================================
# MAIN
# ============================================================================
def main():
    section("LGD Closed List + First Default Date DPD90 Processor (FIXED)")
    print(f"  Script started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Parse args — filter out flags
    date_args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force_run = "--force"    in sys.argv
    run_diag  = "--diagnose" in sys.argv

    if len(date_args) == 0:
        print("\nUsage (single month):")
        print("  python lgd_processor.py 2025-08-31")
        print("\nUsage (date range — processes every month-end in range):")
        print("  python lgd_processor.py 2025-04-30 2025-08-31")
        print("\nFlags:")
        print("  --force     bypass date order check")
        print("  --diagnose  print XLOOKUP key diagnostics after run")
        sys.exit(1)

    # Parse from_date / to_date
    try:
        if len(date_args) == 1:
            to_date = datetime.strptime(date_args[0], "%Y-%m-%d")
            section("Reading From_Date from Last_Complete_Date.txt")
            try:
                from_date = read_last_complete_date()
            except (FileNotFoundError, ValueError) as exc:
                print(f"\nERROR: {exc}")
                sys.exit(1)
        else:
            from_date = datetime.strptime(date_args[0], "%Y-%m-%d")
            to_date   = datetime.strptime(date_args[1], "%Y-%m-%d")
    except ValueError as exc:
        print(f"\nERROR: Invalid date — {exc}. Expected YYYY-MM-DD.")
        sys.exit(1)

    print(f"\n  From Date : {from_date.strftime('%Y-%m-%d')}")
    print(f"  To Date   : {to_date.strftime('%Y-%m-%d')}")

    if not force_run and to_date <= from_date:
        print(f"\nERROR: To_Date must be after From_Date. Use --force to bypass.")
        sys.exit(1)

    # Build list of month-ends to process
    month_list = month_ends_between(from_date, to_date)
    if not month_list:
        print("\n  No month-end dates to process in the given range.")
        sys.exit(0)

    print(f"\n  Months to process : {len(month_list)}")
    for me in month_list:
        print(f"    {me.strftime('%Y-%m-%d')}")

    # PRE-RUN: clear Closed List data rows before processing
    try:
        clear_closed_list_data()
    except Exception as exc:
        print(f"\nERROR: Pre-run clear of Closed List failed: {exc}")
        sys.exit(1)

    # Locate closure file once — same for all months
    section("Locating Closure file")
    try:
        closure_file = get_closure_file()
        print(f"  Closure file : {os.path.basename(closure_file)}")
    except FileNotFoundError as exc:
        print(f"\nERROR: {exc}")
        sys.exit(1)

    # ── ONE-TIME: Closure Category records ──────────────────────────────────
    section("STEP 1 + 3 (One-time) : Closure Category → Closed List")
    try:
        closure_df = step1_closure_data(closure_file)
        step3_append_closed_list(closure_df)
    except Exception as exc:
        print(f"\nERROR writing Closure Category records: {exc}")
        sys.exit(1)

    # ── ONE-TIME: Virtual Contracts from final month ─────────────────────────
    final_month = month_list[-1]
    section(f"STEP 2 + 3 (One-time) : Virtual Contracts from final month "
            f"({final_month.strftime('%Y-%m-%d')}) → Closed List")
    try:
        final_portfolio_file = get_portfolio_file(final_month)
        print(f"  Portfolio file : {os.path.basename(final_portfolio_file)}")
        virtual_df = step2_virtual_contracts(final_portfolio_file)
        if virtual_df.empty:
            print("  No Virtual Contract records found.")
        else:
            step3_append_closed_list(virtual_df)
    except FileNotFoundError as exc:
        print(f"\nERROR: {exc}")
        sys.exit(1)
    except Exception as exc:
        print(f"\nERROR writing Virtual Contract records: {exc}")
        sys.exit(1)

    # ── Process each month: Steps 4–6 only (Step 7 runs once at the end) ────
    succeeded = []
    failed_at = None

    for me in month_list:
        ok = process_month(me)   # FIX 11: no step7 inside here
        if ok:
            succeeded.append(me)
        else:
            failed_at = me.strftime("%Y-%m-%d")
            break

    # Diagnostics (optional)
    if run_diag:
        diagnose_xlookup(10)

    # ── BATCH SUMMARY ────────────────────────────────────────────────────────
    section("BATCH SUMMARY")
    for me in succeeded:
        print(f"  [OK]   {me.strftime('%Y-%m-%d')}")
    if failed_at:
        print(f"  [FAIL] {failed_at}")
        print(f"\n  Last_Complete_Date.txt holds the last successful date.")
        print(f"  Fix the issue and re-run.")
        sys.exit(1)

    # ── STEP 7 (ONCE) — formulas + date format + highlight + output file ────
    # FIX 11: Called exactly once here, covering the full date range.
    # The output filename reflects from_date → last successful month-end.
    if succeeded:
        last_me = succeeded[-1]
        section("STEP 7 (FINAL — single output file) : VLOOKUP formulas + date format")
        print(f"  Covering range : {from_date.strftime('%Y-%m-%d')} → "
              f"{last_me.strftime('%Y-%m-%d')}")
        try:
            step7_update_formulas(from_date, last_me)
        except Exception as exc:
            print(f"\nERROR in Step 7: {exc}")
            sys.exit(1)

    print(f"\n  All {len(succeeded)} month(s) processed successfully.")
    print(f"  Output : {os.path.basename(build_output_path(from_date, succeeded[-1]))}")
    print()

# ============================================================================
if __name__ == "__main__":
    main()