
"""
Write Off Working - Filtered File Creator
==========================================
Runs AFTER all other processes complete.

PURPOSE:
  1.  Open "Write Off Working - <Month> <Year>.xlsx" from Input Files/LGD/.
  2.  In sheet "0130-GL_Working", keep ONLY rows whose "Final Type" column
      contains "Write Back 0130" or "Write Back 1528".  Delete all other rows.
  3.  SaveAs "Write Off Working_New - <Month> <Year>.xlsx" in the same folder.
  4.  In the _New file:
        - Clear the 9 transaction columns (data only, keep headers):
            Transaction No, Contract No, Initials, Client Name,
            Debit, Credit, Amount, Trans Date, Trans Description
        - Keep (and later AutoFill) the 6 formula columns:
            Txn Net Amount, Transaction Type, Contract Type,
            Original Contract No, 1528_GL, Final Type
  5.  Import transaction data from:
        "Backend Due Types - Transaction wise Report - Internal *.xlsx"
      Sheet: "Backend Due Types - Transaction"
      Column mapping  ->  destination column:
        SLT_TRNNO    -> Transaction No
        SLT_CONNO    -> Contract No
        CLM_INITIALS -> Initials
        CLM_NAME     -> Client Name
        DRAMT        -> Debit
        CRAMT        -> Credit
        DRAMT-CRAMT  -> Amount  (computed plain value)
        SLT_TRNDATE  -> Trans Date
        SLT_DESC     -> Trans Description
  6.  Copy OLD_CON_NO / LD_NO from:
        "Legal Bal New* - Internal *.xlsx"  ->  sheet "Legal_Bal_New"
      Paste starting at row 3 under columns OLD_CON_NO / LD_NO.
  7.  AutoFill the 6 formula columns from their row-2 template down to the
      last data row (via Excel COM so row references update correctly).
  8.  Refresh the pivot table in sheet "NetOff_Pivot" (full data range).
  9.  Move all rows where Txn Net Amount = 0 from "0130-GL_Working" to
      sheet "NetOff_Zero_Txns" (clear that sheet first, then paste).
  10. Move all rows where Trans Description = "OVERPAYMENT SETTLEMENT"
      from "0130-GL_Working" and APPEND them to "NetOff_Zero_Txns".
  11. Refresh the pivot table in sheet "Final Pivot" (full data range).

PIPELINE:
  COM  Session 1 : Filter rows; delete utility row; SaveAs _New.
  openpyxl Pass 2: Clear 9 cols; import Backend Due Types data.
  openpyxl Pass 3: Copy Legal Bal New -> Legal_Bal_New sheet.
  COM  Session 2 : AutoFill formulas; refresh pivots; move zero/OP rows;
                   final save & close.

FIXES IN THIS VERSION:
  - Replaced win32com.client.dynamic.Dispatch with a _launch_excel() helper
    that uses DispatchEx (always starts a fresh Excel instance) and retries
    up to 3 times with a delay, avoiding CO_E_SERVER_EXEC_FAILURE (-2147221021).
  - Clear diagnostic message printed when COM launch fails, with actionable
    remediation steps.

Usage:
  python write_off.py
  python write_off.py "Input Files/LGD/Write Off Working - July 2025.xlsx"
"""

import os
import sys
import glob
import re
import time
import warnings
from datetime import datetime

# Force UTF-8 output on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


# ============================================================================
# LOGGING — mirror all print() output to a log file in IFRS_LGD_LOG_DIR
# ============================================================================

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
    log_file = os.path.join(log_dir, "write_off.log")
    sys.stdout = _TeeWriter(sys.stdout, log_file)
    sys.stderr = _TeeWriter(sys.stderr, log_file)
    print(f"Log file: {log_file}")


_setup_log_tee()

import openpyxl

# Clear the win32com gen_py type-library cache before loading win32com.
# A stale cache (built for a different Excel version) produces
# "Member not found" (-2147352573) on methods like SaveAs and AutoFilter.
# Deleting it here forces regeneration against the currently installed Excel.
import tempfile as _tempfile, shutil as _shutil
_gen_py = os.path.join(_tempfile.gettempdir(), "gen_py")
if os.path.isdir(_gen_py):
    try:
        _shutil.rmtree(_gen_py)
    except Exception:
        pass  # non-fatal: worst case the stale cache remains

import win32com.client

# ============================================================================
# CONFIGURATION
# ============================================================================
_SCRIPT_ROOT = r"c:\CBSL\Script"


def _find_lgd_working_folder() -> str:
    base = os.path.join(_SCRIPT_ROOT, "working", "LGD")
    if os.path.isdir(base):
        subfolders = [
            d for d in os.listdir(base)
            if os.path.isdir(os.path.join(base, d))
        ]
        if subfolders:
            # Use latest modified date folder so changing LGD dates are handled automatically.
            subfolders.sort(
                key=lambda d: os.path.getmtime(os.path.join(base, d)),
                reverse=True
            )
            return os.path.join(base, subfolders[0], "firstDefault")
    return base


_lgd_env = os.environ.get("LGD_BASE_FOLDER", "").strip()
BASE_ROOT = _SCRIPT_ROOT
if _lgd_env:
    # If env already points to firstDefault, use it directly.
    LGD_DIR = _lgd_env if os.path.basename(_lgd_env).lower() == "firstdefault" else os.path.join(_lgd_env, "firstDefault")
else:
    LGD_DIR = _find_lgd_working_folder()

SOURCE_GLOB = "Write Off Working - *.xlsx"
NEW_SUFFIX  = "_New"

WORKING_SHEET  = "0130-GL_Working"
# Header row in the SOURCE file is detected at runtime (see _detect_header_row) —
# some exports include a utility row above the headers, some don't.

# Keep ANY row whose Final Type starts with "Write Back" (case-insensitive).
# Everything else (Disposal Loss, OCH 0130, OCH 1528, blank, 0, …) is deleted.
KEEP_PREFIX = "write back"   # compared in lower-case

def _normalise_ft(ft_value) -> str:
    """
    Normalise a Final Type cell value for comparison:
      • Strip leading/trailing whitespace
      • Collapse any run of internal whitespace (spaces, tabs, etc.) to a
        single space  →  "Write Back  0130" becomes "write back 0130"
      • Lower-case the result
    This handles cells that contain double spaces, tabs, or non-breaking
    spaces between words, which would otherwise prevent a startswith() match.
    """
    if ft_value is None:
        return ""
    return re.sub(r"\s+", " ", str(ft_value)).strip().lower()


def _is_keep_type(ft_value) -> bool:
    """Return True if the Final Type value should be KEPT (any Write Back …)."""
    return _normalise_ft(ft_value).startswith(KEEP_PREFIX)

# ---- Auxiliary file glob patterns (timestamp in filename is ignored) ----
BACKEND_DUE_GLOB  = "Backend Due Types - Contract wise Report - Internal *.xlsx"
BACKEND_DUE_SHEET = "Backend Due Types - Contract wi"

LEGAL_BAL_GLOB         = "Legal Bal New* - Internal *.xlsx"
LEGAL_BAL_SOURCE_SHEET = "Legal Bal New - Internal"
LEGAL_BAL_DEST_SHEET   = "Legal_Bal_New"
LEGAL_BAL_DEST_START   = 3    # paste data starting at this row

# Column mapping: destination column -> Backend Due Types source column
COLUMN_MAPPING = {
    "Transaction No"   : "SLT_TRNNO",
    "Contract No"      : "SLT_CONNO",
    "Initials"         : "CLM_INITIALS",
    "Client Name"      : "CLM_NAME",
    "Debit"            : "DRAMT",
    "Credit"           : "CRAMT",
    "Trans Date"       : "SLT_TRNDATE",
    "Trans Description": "SLT_DESC",
    # Amount = DRAMT - CRAMT  (computed separately, not in this map)
}

# 9 columns whose data is cleared and then refilled from Backend Due Types
COLUMNS_TO_CLEAR = {
    "Transaction No", "Contract No", "Initials", "Client Name",
    "Debit", "Credit", "Amount", "Trans Date", "Trans Description",
}

# 6 columns whose row-2 formula is AutoFilled down to all data rows
COLUMNS_TO_KEEP = {
    "Txn Net Amount", "Transaction Type", "Contract Type",
    "Original Contract No", "1528_GL", "Final Type",
}

# Columns written to / read from NetOff_Zero_Txns (in this order)
MOVE_COLUMNS = [
    "Transaction No", "Contract No", "Initials", "Client Name",
    "Debit", "Credit", "Amount", "Trans Date", "Trans Description",
    "Txn Net Amount",
]

# Sheet names for pivot / zero-txn destinations
NETOFF_ZERO_SHEET  = "NetOff_Zero_Txns"
NETOFF_PIVOT_SHEET = "NetOff_Pivot"
FINAL_PIVOT_SHEET  = "Final Pivot"

# Excel COM constants
XL_UP          = -4162   # xlUp
XL_LEFT        = -4159   # xlToLeft
XL_UPDATE_YES  = 1       # UpdateLinks = yes (resolve cross-sheet refs)
XL_XLSX_FORMAT = 51      # xlOpenXMLWorkbook (.xlsx, no macros)


# ============================================================================
# EXCEL COM LAUNCH HELPER  (replaces direct Dispatch / dynamic.Dispatch calls)
# ============================================================================

def _launch_excel(max_attempts: int = 3, delay: float = 4.0):
    """
    Launch a fresh Excel COM server and return the Application object.

    Uses DispatchEx() rather than Dispatch() so that each call always creates
    a brand-new Excel.exe process, avoiding conflicts with already-running
    (possibly hung) instances.

    Retries up to max_attempts times with a delay between attempts.
    On total failure prints a clear remediation checklist and raises
    RuntimeError so the caller can handle it gracefully.

    Common cause of CO_E_SERVER_EXEC_FAILURE (-2147221021):
      • Script is running as Administrator — Excel COM refuses to start
        under an elevated token when no interactive session is available.
      • A previous Excel.exe is hung in the background (check Task Manager).
      • DCOM identity for Excel is not set to "The interactive user".
    """
    last_exc = None
    for attempt in range(1, max_attempts + 1):
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible          = False
            xl.DisplayAlerts    = False
            xl.AskToUpdateLinks = False
            print(f"  Excel COM launched (attempt {attempt}/{max_attempts}).")
            return xl
        except Exception as exc:
            last_exc = exc
            print(f"  Excel COM launch attempt {attempt}/{max_attempts} failed: {exc}")
            if attempt < max_attempts:
                print(f"  Retrying in {delay}s …")
                time.sleep(delay)

    raise RuntimeError(
        f"Could not launch Excel COM server after {max_attempts} attempts.\n"
        f"Last error: {last_exc}\n\n"
        "Remediation steps (try in order):\n"
        "  1. Do NOT run this script as Administrator — open CMD normally.\n"
        "  2. Kill any hung Excel processes:\n"
        "       taskkill /F /IM EXCEL.EXE\n"
        "  3. Clear the Excel registry entry (run as Admin once, then revert):\n"
        "       reg delete \"HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\" /f\n"
        "     (Replace 16.0 with your Office version: 15.0=2013, 14.0=2010)\n"
        "  4. In dcomcnfg → Component Services → My Computer → DCOM Config\n"
        "     → Microsoft Excel Application → Properties → Identity tab\n"
        "     → select 'The interactive user' → OK.\n"
    )


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def _section(title: str) -> None:
    """Print a visible section heading."""
    print("\n" + "=" * 65)
    print(f"  {title}")
    print("=" * 65)


def _looks_like_run_date(value: str) -> bool:
    """Return True when a CLI arg looks like a date token rather than a filename."""
    if not value:
        return False
    raw = value.strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            datetime.strptime(raw, fmt)
            return True
        except ValueError:
            continue
    return False


def _col_letter(col_idx: int) -> str:
    """Convert 1-based column index to Excel column letter(s). e.g. 1->'A', 27->'AA'."""
    result = ""
    while col_idx:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def _get_sheet(wb_com, sheet_name: str):
    """
    Return the worksheet named sheet_name from a COM workbook.

    Tries in order:
      1. Exact name match.
      2. Case-insensitive match (handles capitalisation differences).
      3. Partial match — sheet whose name CONTAINS the key part of sheet_name.

    If no match is found, raises a clear ValueError that lists every sheet
    present in the workbook so the user can correct WORKING_SHEET if needed.
    """
    # 1. Exact
    try:
        return wb_com.Sheets(sheet_name)
    except Exception:
        pass

    # Build a list of all sheet names via COM
    all_sheets = []
    try:
        for i in range(1, wb_com.Sheets.Count + 1):
            all_sheets.append(wb_com.Sheets(i).Name)
    except Exception:
        all_sheets = ["<could not enumerate sheets>"]

    # 2. Case-insensitive exact
    target_lower = sheet_name.strip().lower()
    for name in all_sheets:
        if name.strip().lower() == target_lower:
            print(f"  NOTE: Sheet found as '{name}' "
                  f"(config says '{sheet_name}') — using it.")
            return wb_com.Sheets(name)

    # 3. Partial match (sheet name contains the target, ignoring case)
    for name in all_sheets:
        if target_lower in name.strip().lower():
            print(f"  NOTE: Partial sheet match '{name}' "
                  f"used for '{sheet_name}'.")
            return wb_com.Sheets(name)

    raise ValueError(
        f"Sheet '{sheet_name}' not found in workbook.\n"
        f"  Sheets present: {all_sheets}\n"
        f"  Tip: update WORKING_SHEET in the configuration section "
        f"to match the exact sheet name shown above."
    )


def _build_col_map_com(ws_com, header_row: int) -> dict:
    """Return {header_name: 1-based_col_index} by scanning a COM worksheet row."""
    col_map  = {}
    last_col = ws_com.Cells(
        header_row, ws_com.Columns.Count).End(XL_LEFT).Column
    for c in range(1, last_col + 1):
        hdr = ws_com.Cells(header_row, c).Value
        if hdr:
            col_map[str(hdr).strip()] = c
    return col_map


def _build_col_map_opxl(ws, header_row: int) -> dict:
    """Return {header_name: 1-based_col_index} by scanning an openpyxl worksheet row."""
    col_map = {}
    for col in range(1, ws.max_column + 1):
        hdr = ws.cell(header_row, col).value
        if hdr is not None:
            col_map[str(hdr).strip()] = col
    return col_map


def _detect_header_row(ws_com, max_scan_rows: int = 5) -> tuple[int, int]:
    """
    Locate the header row in the raw SOURCE sheet by scanning column A of
    the first max_scan_rows rows for "Transaction No" — the same check
    already relied on elsewhere in this file to tell the two known source
    layouts apart:
      - utility row (row 1) + headers (row 2) + data (row 3+)
      - headers (row 1) + data (row 2+), no utility row
    Returns (header_row, data_start_row).
    """
    for r in range(1, max_scan_rows + 1):
        val = str(ws_com.Cells(r, 1).Value or "").strip()
        if val == "Transaction No":
            return r, r + 1
    raise ValueError(
        f"Could not find header row (column A = 'Transaction No') "
        f"in the first {max_scan_rows} rows of sheet '{WORKING_SHEET}'."
    )


def find_source_file(lgd_dir: str) -> str:
    """Find the Write Off Working source file (excludes _New variants and temp files)."""
    pattern    = os.path.join(lgd_dir, SOURCE_GLOB)
    candidates = [
        f for f in glob.glob(pattern)
        if "_New" not in os.path.basename(f)
        and not os.path.basename(f).startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(
            f"No source file matching '{SOURCE_GLOB}' found in:\n  {lgd_dir}"
        )
    candidates.sort(key=os.path.getmtime, reverse=True)
    if len(candidates) > 1:
        print("  Multiple source files found — using most recently modified:")
        for i, fp in enumerate(candidates):
            print(f"    {os.path.basename(fp)}"
                  f"{'  <-- SELECTED' if i == 0 else ''}")
    return candidates[0]


def find_aux_file(lgd_dir: str, pattern: str, label: str) -> str:
    """
    Locate an auxiliary input file inside lgd_dir using a glob pattern.
    - Skips temp files (prefix ~$).
    - If multiple matches, picks the most recently modified one.
    - On failure, lists all .xlsx files present to assist diagnosis.
    """
    full_pattern = os.path.join(lgd_dir, pattern)
    candidates   = [
        f for f in glob.glob(full_pattern)
        if not os.path.basename(f).startswith("~$")
    ]
    if not candidates:
        try:
            all_xlsx = sorted(
                f for f in os.listdir(lgd_dir)
                if f.lower().endswith(".xlsx")
            )
        except OSError:
            all_xlsx = ["<directory not accessible>"]
        raise FileNotFoundError(
            f"{label} file not found.\n"
            f"  Pattern searched : {full_pattern}\n"
            f"  Directory        : {lgd_dir}\n"
            f"  .xlsx files present:\n"
            + "\n".join(f"    {f}" for f in all_xlsx)
        )
    candidates.sort(key=os.path.getmtime, reverse=True)
    if len(candidates) > 1:
        print(f"  [{label}] Multiple matches — using most recently modified:")
        for i, fp in enumerate(candidates):
            print(f"    {os.path.basename(fp)}"
                  f"{'  <-- SELECTED' if i == 0 else ''}")
    else:
        print(f"  [{label}] Found : {os.path.basename(candidates[0])}")
    return candidates[0]


def build_dest_path(source_path: str) -> str:
    """Derive the _New output path from the source path."""
    folder   = os.path.dirname(source_path)
    basename = os.path.basename(source_path)
    new_name = basename.replace(
        "Write Off Working - ", "Write Off Working_New - ", 1)
    if new_name == basename:          # fallback if prefix did not match
        name, ext = os.path.splitext(basename)
        new_name  = name + NEW_SUFFIX + ext
    return os.path.join(folder, new_name)


def _full_used_range(ws_data) -> object:
    """
    Return a Range object that covers EVERY used cell in ws_data.

    Strategy (three attempts, most reliable first):
      1. UsedRange  — Excel's own record of the used extent; always A1-anchored
                      because we anchor to (1,1) explicitly.
      2. SpecialCells(xlCellTypeLastCell) — last cell that ever held data.
      3. End(XL_UP) / End(XL_LEFT) fallback — scan from last row/col.

    All three methods anchor the top-left corner to cell (1,1) so the range
    always starts at A1 (the header row).
    """
    XL_LAST_CELL = 11   # xlCellTypeLastCell

    # --- Method 1: UsedRange (preferred) ---
    try:
        ur       = ws_data.UsedRange
        last_row = ur.Row + ur.Rows.Count - 1
        last_col = ur.Column + ur.Columns.Count - 1
        if last_row >= 1 and last_col >= 1:
            return ws_data.Range(
                ws_data.Cells(1, 1),
                ws_data.Cells(last_row, last_col)
            )
    except Exception:
        pass

    # --- Method 2: SpecialCells(xlCellTypeLastCell) ---
    try:
        lc       = ws_data.Cells.SpecialCells(XL_LAST_CELL)
        last_row = lc.Row
        last_col = lc.Column
        if last_row >= 1 and last_col >= 1:
            return ws_data.Range(
                ws_data.Cells(1, 1),
                ws_data.Cells(last_row, last_col)
            )
    except Exception:
        pass

    # --- Method 3: End(XL_UP) / End(XL_LEFT) fallback ---
    last_row = ws_data.Cells(ws_data.Rows.Count, 1).End(XL_UP).Row
    last_col = ws_data.Cells(1, ws_data.Columns.Count).End(XL_LEFT).Column
    return ws_data.Range(
        ws_data.Cells(1, 1),
        ws_data.Cells(max(last_row, 1), max(last_col, 1))
    )


def _bulk_delete_rows(ws_com, row_list: list, last_row: int) -> None:
    """Delete scattered rows in one AutoFilter pass (~5 COM calls total).

    Writes a temp marker column in bulk, filters to show only marked rows,
    deletes all visible rows in one call, then removes the temp column.
    """
    bad_set  = set(row_list)
    temp_col = ws_com.UsedRange.Columns.Count + 2

    # Build marker array and write in a single COM call
    marker = tuple(
        ("DEL",) if r in bad_set else (None,)
        for r in range(1, last_row + 1)
    )
    ws_com.Range(
        ws_com.Cells(1, temp_col),
        ws_com.Cells(last_row, temp_col)
    ).Value = marker

    # Filter to show only "DEL" rows, then delete visible rows in one shot
    ws_com.AutoFilterMode = False
    ws_com.Range("1:1").AutoFilter(Field=temp_col, Criteria1="DEL")
    try:
        ws_com.Range(
            ws_com.Cells(2, 1), ws_com.Cells(last_row, 1)
        ).SpecialCells(12).EntireRow.Delete()   # 12 = xlCellTypeVisible
    except Exception:
        pass  # nothing matched (shouldn't happen if row_list is non-empty)

    ws_com.AutoFilterMode = False
    ws_com.Columns(temp_col).Delete()


def _bulk_copy_and_delete(ws_src, ws_dest, row_list: list,
                          col_map: dict, last_row: int,
                          dest_start: int) -> None:
    """Bulk-read MOVE_COLUMNS from ws_src for row_list, write to ws_dest,
    then delete those rows from ws_src.  Uses one COM call per column to
    read, one COM call to write, plus _bulk_delete_rows to delete.
    """
    def _flatten(raw, n):
        if raw is None:
            return [None] * n
        if not isinstance(raw, tuple):
            return [raw]
        return [r[0] if isinstance(r, tuple) else r for r in raw]

    n_data = last_row - 1          # number of data rows (row 2 … last_row)

    # Bulk-read each MOVE_COLUMN (one COM call per column)
    col_arrays = {}
    for col_name in MOVE_COLUMNS:
        if col_name not in col_map:
            col_arrays[col_name] = [None] * n_data
            continue
        letter = _col_letter(col_map[col_name])
        raw = ws_src.Range(f"{letter}2:{letter}{last_row}").Value2
        col_arrays[col_name] = _flatten(raw, n_data)

    # Build 2-D array of rows to write (only the matching rows, in order)
    rows_data = []
    for src_row in row_list:
        i = src_row - 2            # 0-based index into col_arrays
        rows_data.append(tuple(
            col_arrays[c][i] if i < len(col_arrays[c]) else None
            for c in MOVE_COLUMNS
        ))

    # Write all rows in one COM call
    n = len(rows_data)
    ws_dest.Range(
        ws_dest.Cells(dest_start, 1),
        ws_dest.Cells(dest_start + n - 1, len(MOVE_COLUMNS))
    ).Value = rows_data

    # Delete the source rows
    _bulk_delete_rows(ws_src, row_list, last_row)


def _refresh_pivot_sheet(wb_com, sheet_name: str,
                         data_source_sheet: str = None) -> None:
    """
    Refresh every pivot table in sheet_name.

    If data_source_sheet is given:
      • Use _full_used_range() to get the COMPLETE extent of that sheet
        (UsedRange-anchored to A1, so the header row is always included).
      • Build a brand-new PivotCache from the Range object
        (avoids string-escaping problems with sheet names containing
        digits, hyphens, or other special characters).
      • Swap the new cache in with ChangePivotCache() then call Refresh().

    Using a Range object as SourceData avoids error -2146827284
    ("Cannot change this part of a PivotTable report") which is triggered
    by badly-formatted SourceData strings.
    """
    XL_DATABASE = 1   # xlDatabase PivotCache SourceType

    try:
        ws = wb_com.Sheets(sheet_name)
    except Exception:
        print(f"    Sheet '{sheet_name}' not found — pivot refresh skipped.")
        return

    count = ws.PivotTables().Count
    if count == 0:
        print(f"    No pivot tables found in '{sheet_name}'.")
        return

    for i in range(1, count + 1):
        pt = ws.PivotTables(i)
        try:
            # ---- Update source range to full extent of data sheet ----
            if data_source_sheet:
                try:
                    ws_data   = wb_com.Sheets(data_source_sheet)
                    src_range = _full_used_range(ws_data)

                    # Log exactly what range was detected
                    r1 = src_range.Row
                    c1 = src_range.Column
                    r2 = r1 + src_range.Rows.Count - 1
                    c2 = c1 + src_range.Columns.Count - 1
                    print(f"    Full range detected : "
                          f"{_col_letter(c1)}{r1}:{_col_letter(c2)}{r2}  "
                          f"({src_range.Rows.Count:,} rows x "
                          f"{src_range.Columns.Count} cols)")

                    # Create a fresh cache from the Range object and swap it in
                    new_cache = wb_com.PivotCaches().Create(
                        SourceType=XL_DATABASE,
                        SourceData=src_range
                    )
                    pt.ChangePivotCache(new_cache)
                    print(f"    PivotCache updated  -> "
                          f"'{data_source_sheet}' full range")

                except Exception as exc:
                    # Range update failed — still attempt a plain Refresh()
                    print(f"    WARNING: Source range update failed: {exc}")
                    print(f"    Falling back to plain Refresh() ...")

            pt.PivotCache().Refresh()
            print(f"    Refreshed pivot '{pt.Name}' in sheet '{sheet_name}'")

        except Exception as exc:
            print(f"    WARNING: Pivot '{pt.Name}' refresh failed: {exc}")


# ============================================================================
# MAIN PIPELINE
# ============================================================================

def create_write_off_filtered_file(source_path: str) -> str:

    dest_path = build_dest_path(source_path)
    lgd_dir   = os.path.dirname(source_path)

    _section("Write Off Filtered File Creator")
    print(f"  Source  : {os.path.basename(source_path)}")
    print(f"  Output  : {os.path.basename(dest_path)}")
    print(f"  LGD dir : {lgd_dir}\n")

    # Locate auxiliary input files by glob pattern (timestamp ignored)
    backend_path = find_aux_file(lgd_dir, BACKEND_DUE_GLOB, "Backend Due Types")
    legal_path   = find_aux_file(lgd_dir, LEGAL_BAL_GLOB,   "Legal Bal New")

    print(f"\n  Backend Due   : {os.path.basename(backend_path)}")
    print(f"  Legal Bal New : {os.path.basename(legal_path)}")

    # ==========================================================================
    # COM SESSION 1
    #   Read Final Type values; delete non-matching rows; delete utility row;
    #   SaveAs _New file; close source.
    # ==========================================================================
    _section("COM Session 1 — Filter rows & SaveAs _New")

    # Copy source → dest before opening, so we can use wb.Save() instead of
    # SaveAs (SaveAs via dynamic COM dispatch is unreliable across Excel versions).
    if os.path.exists(dest_path):
        try:
            os.remove(dest_path)
            print(f"  Removed existing output file.")
        except OSError as exc:
            raise OSError(
                f"Cannot overwrite existing output file (is it open in Excel?):\n"
                f"  {dest_path}\n  Original error: {exc}"
            )
    import shutil as _shutil
    _shutil.copy2(source_path, dest_path)
    print(f"  Copied source to output path.")

    # ---- Launch Excel (with retry logic) ----
    xl = _launch_excel()

    try:
        wb_src = xl.Workbooks.Open(
            os.path.abspath(dest_path),
            UpdateLinks=XL_UPDATE_YES
        )
        # Force full recalculation so formula-driven Final Type values are correct
        xl.CalculateFull()
        xl.CalculateFull()

        ws_src = _get_sheet(wb_src, WORKING_SHEET)

        # ---- Detect header row, then build column map from it ----
        # Source files come in two known layouts (utility row + headers on
        # row 2, or headers directly on row 1, no utility row) — detected
        # here instead of assumed, so either one works without editing code.
        header_row, data_start_row = _detect_header_row(ws_src)
        print(f"  Detected header row : {header_row} "
              f"(data starts at row {data_start_row})")

        src_col_map = _build_col_map_com(ws_src, header_row)
        print(f"  Source columns found ({len(src_col_map)}): "
              f"{list(src_col_map.keys())}")

        ft_col = src_col_map.get("Final Type")
        if not ft_col:
            raise ValueError(
                f"'Final Type' column not found in row {header_row} "
                f"of sheet '{WORKING_SHEET}'."
            )

        # ---- Read every Final Type value — BULK read (single COM call) ----
        # Reading cells one-by-one over COM drops the object reference on large
        # sheets (OLE error 0x800a01a8).  A single Range.Value2 call returns the
        # entire column as a Python tuple in one round-trip, which is both safe
        # and ~100x faster.
        last_src_row = ws_src.Cells(ws_src.Rows.Count, 1).End(XL_UP).Row
        print(f"  Source last row : {last_src_row}  "
              f"(data rows {data_start_row}-{last_src_row})")

        ft_col_letter = _col_letter(ft_col)
        ft_range_addr = (f"{ft_col_letter}{data_start_row}:"
                         f"{ft_col_letter}{last_src_row}")
        ft_range      = ws_src.Range(ft_range_addr)

        # Value2 returns a 2-D tuple ((v,), (v,), …) for a single-column range,
        # or a scalar if only one cell.  Normalise to a flat list.
        raw_values = ft_range.Value2
        if raw_values is None:
            flat_values = [None] * (last_src_row - data_start_row + 1)
        elif not isinstance(raw_values, tuple):
            flat_values = [raw_values]          # single cell
        else:
            # Multi-row single-column: ((v1,), (v2,), …)
            flat_values = [
                row[0] if isinstance(row, tuple) else row
                for row in raw_values
            ]

        final_type_map = {
            data_start_row + i: v
            for i, v in enumerate(flat_values)
        }
        print(f"  Final Type values read (bulk) : {len(final_type_map):,}")

        # ---- Count and report all Final Type values ----
        value_counts = {}
        for ft in final_type_map.values():
            value_counts[ft] = value_counts.get(ft, 0) + 1

        # Build all lines first, then print once — with thousands of unique
        # Final Type values, printing one line at a time was extremely slow
        # (_TeeWriter flushes to disk on every print() call).
        report_lines = ["\n  Final Type value counts in source:",
                         "  (repr column shows exact raw value — reveals hidden spaces/chars)"]
        for val, cnt in sorted(value_counts.items(), key=lambda x: str(x[0])):
            tag  = " KEPT  " if _is_keep_type(val) else "DELETED"
            disp = str(val) if val is not None else "(blank/None)"
            raw  = repr(val)   # shows double-spaces, \t, \xa0, etc.
            report_lines.append(f"    [{tag}]  {disp:<32}  {cnt:>5} rows   raw={raw}")
        print("\n".join(report_lines))

        # ---- Identify rows to delete (Final Type does NOT start with "Write Back") ----
        rows_to_delete = sorted(
            r for r, ft in final_type_map.items()
            if not _is_keep_type(ft)
        )
        rows_to_keep_count = len(final_type_map) - len(rows_to_delete)
        print(f"\n  Rows to delete : {len(rows_to_delete):,}")
        print(f"  Rows to keep   : {rows_to_keep_count:,}")

        # ---- Delete non-matching rows ----
        if rows_to_delete:
            _bulk_delete_rows(ws_src, rows_to_delete, last_src_row)

        # ---- Delete utility/notes row, if present ----
        # header_row was already positively identified above (column A ==
        # "Transaction No"), so there's no need to re-guess here.
        if header_row > 1:
            ws_src.Rows(1).Delete()
            print(f"  Utility row deleted — headers now at row 1 (were at row {header_row}).")
        else:
            print("  No utility row detected — headers already at row 1.")

        # ---- Save in-place (file was already copied to dest_path before open) ----
        wb_src.Save()
        wb_src.Close(SaveChanges=False)
        print(f"\n  Saved as : {os.path.basename(dest_path)}")

    finally:
        try:
            xl.Quit()
        except Exception:
            pass

    # ==========================================================================
    # openpyxl  PASS 2
    #   Read Backend Due Types; open _New file (header=row1, data from row2);
    #   resize to match Backend row count; clear 9 cols; clear formula cols 3+;
    #   fill 9 cols from Backend; Amount = DRAMT-CRAMT; save.
    # ==========================================================================
    _section("Pass 2 — Import Backend Due Types data (openpyxl)")

    # ---- Read Backend Due Types ----
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb_bd = openpyxl.load_workbook(backend_path, data_only=True)

    # Find sheet (exact name first, then case-insensitive)
    if BACKEND_DUE_SHEET in wb_bd.sheetnames:
        ws_bd_name = BACKEND_DUE_SHEET
    else:
        ws_bd_name = next(
            (s for s in wb_bd.sheetnames
             if s.strip().lower() == BACKEND_DUE_SHEET.lower()),
            None
        )
        if not ws_bd_name:
            wb_bd.close()
            raise ValueError(
                f"Sheet '{BACKEND_DUE_SHEET}' not found in:\n  {backend_path}\n"
                f"  Available sheets: {wb_bd.sheetnames}"
            )
        print(f"  Note: using sheet '{ws_bd_name}' "
              f"(expected '{BACKEND_DUE_SHEET}')")

    ws_bd  = wb_bd[ws_bd_name]
    bd_hdr = _build_col_map_opxl(ws_bd, 1)
    print(f"  Backend Due Types sheet   : '{ws_bd_name}'")
    print(f"  Backend Due Types columns : {list(bd_hdr.keys())}")

    # Warn about missing source columns
    missing_src = [v for v in COLUMN_MAPPING.values() if v not in bd_hdr]
    if missing_src:
        print(f"  WARNING: These source columns were not found: {missing_src}")

    bd_rows = []
    for r in range(2, ws_bd.max_row + 1):
        row_data = {col: ws_bd.cell(r, idx).value
                    for col, idx in bd_hdr.items()}
        if any(v is not None for v in row_data.values()):
            bd_rows.append(row_data)
    wb_bd.close()
    n_bd = len(bd_rows)
    print(f"  Backend Due Types data rows : {n_bd:,}")

    # ---- Open _New file ----
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb2 = openpyxl.load_workbook(dest_path)

    # Case-insensitive sheet lookup so minor name differences don't crash
    _ws2_name = WORKING_SHEET
    if _ws2_name not in wb2.sheetnames:
        _ws2_name = next(
            (s for s in wb2.sheetnames
             if s.strip().lower() == WORKING_SHEET.strip().lower()),
            None
        )
        if _ws2_name is None:
            wb2.close()
            raise ValueError(
                f"Sheet '{WORKING_SHEET}' not found in output file.\n"
                f"  Available sheets: {wb2.sheetnames}"
            )
        print(f"  NOTE: Using sheet '{_ws2_name}' "
              f"(config says '{WORKING_SHEET}')")
    ws2      = wb2[_ws2_name]
    dest_hdr = _build_col_map_opxl(ws2, 1)     # header is now row 1
    cur_rows = ws2.max_row - 1                  # current data row count
    print(f"\n  Dest current data rows : {cur_rows:,}")
    print(f"  Backend rows to import : {n_bd:,}")

    # ---- Resize dest to have exactly n_bd data rows ----
    if n_bd > cur_rows:
        extra = n_bd - cur_rows
        # Append blank rows after existing last row
        n_cols = max(dest_hdr.values(), default=1)
        for _ in range(extra):
            ws2.append([None] * n_cols)
        print(f"  Appended {extra} blank rows to accommodate all Backend data rows.")
    elif n_bd < cur_rows:
        excess = cur_rows - n_bd
        ws2.delete_rows(2 + n_bd, amount=excess)
        print(f"  Deleted {excess} excess rows.")

    # ---- Clear 9 transaction columns (all data rows) ----
    for col_name in COLUMNS_TO_CLEAR:
        if col_name in dest_hdr:
            col_idx = dest_hdr[col_name]
            for r in range(2, ws2.max_row + 1):
                ws2.cell(r, col_idx).value = None

    # ---- Clear 6 formula columns rows 3+ (row 2 kept as AutoFill template) ----
    for col_name in COLUMNS_TO_KEEP:
        if col_name in dest_hdr:
            col_idx = dest_hdr[col_name]
            for r in range(3, ws2.max_row + 1):
                ws2.cell(r, col_idx).value = None

    # ---- Column indices for Amount calculation ----
    debit_col  = dest_hdr.get("Debit")
    credit_col = dest_hdr.get("Credit")
    amount_col = dest_hdr.get("Amount")

    # ---- Fill rows from Backend Due Types ----
    filled = 0
    for i, row_data in enumerate(bd_rows):
        out_row = 2 + i

        # Map the 8 direct columns
        for dest_col_name, src_col_name in COLUMN_MAPPING.items():
            if dest_col_name in dest_hdr:
                ws2.cell(out_row, dest_hdr[dest_col_name]).value = (
                    row_data.get(src_col_name)
                )

        # Amount = DRAMT - CRAMT  (plain computed value)
        if amount_col:
            dramt = row_data.get("DRAMT") or 0
            cramt = row_data.get("CRAMT") or 0
            try:
                ws2.cell(out_row, amount_col).value = (
                    float(dramt) - float(cramt)
                )
            except (TypeError, ValueError):
                ws2.cell(out_row, amount_col).value = None

        filled += 1

    print(f"  Rows filled : {filled:,}")

    wb2.save(dest_path)
    wb2.close()
    print(f"\n  Saved : {os.path.basename(dest_path)}")

    # ==========================================================================
    # openpyxl  PASS 3
    #   Read OLD_CON_NO / LD_NO from Legal Bal New source file.
    #   Open _New file; clear Legal_Bal_New sheet from row 3 down;
    #   paste data starting at row 3; save.
    # ==========================================================================
    _section("Pass 3 — Copy Legal Bal New data into Legal_Bal_New sheet")

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb_lb = openpyxl.load_workbook(legal_path, data_only=True)

    # Find source sheet (exact name first, then partial match)
    if LEGAL_BAL_SOURCE_SHEET in wb_lb.sheetnames:
        ws_lb_name = LEGAL_BAL_SOURCE_SHEET
    else:
        ws_lb_name = next(
            (s for s in wb_lb.sheetnames
             if "legal bal" in s.lower()),
            None
        )
        if not ws_lb_name:
            wb_lb.close()
            raise ValueError(
                f"Could not find a 'Legal Bal' sheet in:\n  {legal_path}\n"
                f"  Available sheets: {wb_lb.sheetnames}"
            )
        print(f"  Note: using sheet '{ws_lb_name}' "
              f"(expected '{LEGAL_BAL_SOURCE_SHEET}')")

    ws_lb  = wb_lb[ws_lb_name]
    lb_hdr = _build_col_map_opxl(ws_lb, 1)
    print(f"  Legal Bal source sheet   : '{ws_lb_name}'")
    print(f"  Legal Bal source columns : {list(lb_hdr.keys())}")

    old_con_col = lb_hdr.get("OLD_CON_NO")
    ld_no_col   = lb_hdr.get("LD_NO")

    if not old_con_col:
        print("  WARNING: 'OLD_CON_NO' column not found in source.")
    if not ld_no_col:
        print("  WARNING: 'LD_NO' column not found in source.")

    lb_data = []
    for r in range(2, ws_lb.max_row + 1):
        old_con = ws_lb.cell(r, old_con_col).value if old_con_col else None
        ld_no   = ws_lb.cell(r, ld_no_col).value   if ld_no_col   else None
        if old_con is not None or ld_no is not None:
            lb_data.append((old_con, ld_no))
    wb_lb.close()
    print(f"  Legal Bal data rows read : {len(lb_data):,}")

    # ---- Open _New file and find destination sheet ----
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb3 = openpyxl.load_workbook(dest_path)

    if LEGAL_BAL_DEST_SHEET not in wb3.sheetnames:
        wb3.close()
        raise ValueError(
            f"Sheet '{LEGAL_BAL_DEST_SHEET}' not found in output file.\n"
            f"  Available sheets: {wb3.sheetnames}"
        )

    ws_dest_lb = wb3[LEGAL_BAL_DEST_SHEET]

    # ---- Locate OLD_CON_NO / LD_NO headers in dest sheet (scan rows 1-3) ----
    dest_lb_cols = {}
    for scan_r in range(1, 4):
        for col in range(1, ws_dest_lb.max_column + 1):
            val = ws_dest_lb.cell(scan_r, col).value
            if val and str(val).strip() in ("OLD_CON_NO", "LD_NO"):
                dest_lb_cols[str(val).strip()] = col

    dest_old_con = dest_lb_cols.get("OLD_CON_NO")
    dest_ld_no   = dest_lb_cols.get("LD_NO")
    print(f"  Dest OLD_CON_NO column : {dest_old_con}  |  "
          f"LD_NO column : {dest_ld_no}")

    if not dest_old_con and not dest_ld_no:
        print("  WARNING: Neither OLD_CON_NO nor LD_NO found in dest sheet "
              "— Legal Bal data will NOT be written.")
    else:
        # Clear existing data from LEGAL_BAL_DEST_START downwards
        if ws_dest_lb.max_row >= LEGAL_BAL_DEST_START:
            for r in range(LEGAL_BAL_DEST_START, ws_dest_lb.max_row + 1):
                if dest_old_con:
                    ws_dest_lb.cell(r, dest_old_con).value = None
                if dest_ld_no:
                    ws_dest_lb.cell(r, dest_ld_no).value = None

        # Write data starting at LEGAL_BAL_DEST_START (row 3)
        for i, (old_con, ld_no) in enumerate(lb_data):
            r = LEGAL_BAL_DEST_START + i
            if dest_old_con:
                ws_dest_lb.cell(r, dest_old_con).value = old_con
            if dest_ld_no:
                ws_dest_lb.cell(r, dest_ld_no).value = ld_no

        print(f"  Legal Bal rows written : {len(lb_data):,}")

    wb3.save(dest_path)
    wb3.close()
    print(f"\n  Saved : {os.path.basename(dest_path)}")

    # ==========================================================================
    # COM SESSION 2 — Passes 4 through 9
    #
    #   Pass 4 : Open _New; recalculate; AutoFill 6 formula columns from row-2
    #            template down to last data row.
    #   Pass 4B: Delete rows where Final Type is NOT a Write Back variant
    #            (formula-driven values now resolved after AutoFill + recalc).
    #   Pass 5 : Refresh NetOff_Pivot (full source range).
    #   Pass 6 : Clear NetOff_Zero_Txns; move Txn Net Amount = 0 rows there.
    #   Pass 7 : Append OVERPAYMENT SETTLEMENT rows to NetOff_Zero_Txns.
    #   Pass 8 : Refresh Final Pivot (full source range).
    #   Pass 9 : Final CalculateFull; Save; Close.
    # ==========================================================================
    _section("COM Session 2 — Passes 4-9 (Formulas / NetOff / Pivots)")

    # ---- Launch Excel (with retry logic) ----
    xl = _launch_excel()

    try:
        wb_com = xl.Workbooks.Open(
            os.path.abspath(dest_path),
            UpdateLinks=XL_UPDATE_YES
        )
        xl.CalculateFull()
        xl.CalculateFull()

        # ==================================================================
        # PASS 4 — AutoFill 6 formula columns from row-2 template
        # ==================================================================
        _section("Pass 4 — Extend formula columns via AutoFill")

        ws_com      = _get_sheet(wb_com, WORKING_SHEET)
        com_col_map = _build_col_map_com(ws_com, 1)    # header is row 1
        last_row    = ws_com.Cells(
            ws_com.Rows.Count, 1).End(XL_UP).Row
        print(f"  Working sheet data rows : {last_row - 1:,}  "
              f"(rows 2-{last_row})")

        if last_row >= 3:
            for col_name in COLUMNS_TO_KEEP:
                if col_name not in com_col_map:
                    print(f"  WARNING: '{col_name}' not found — AutoFill skipped.")
                    continue
                col_idx    = com_col_map[col_name]
                col_ltr    = _col_letter(col_idx)
                src_range  = ws_com.Range(f"{col_ltr}2")
                fill_range = ws_com.Range(f"{col_ltr}2:{col_ltr}{last_row}")

                # Skip AutoFill if the template cell has no content
                try:
                    has_formula = src_range.HasFormula
                    has_value   = src_range.Value2 is not None
                except Exception:
                    has_formula = False
                    has_value   = False

                if not has_formula and not has_value:
                    print(f"  SKIP (empty template): {col_name} [{col_ltr}2]")
                    continue
                try:
                    src_range.AutoFill(Destination=fill_range)
                    print(f"  AutoFill OK : {col_name}  "
                          f"{col_ltr}2:{col_ltr}{last_row}")
                except Exception as exc:
                    print(f"  WARNING: AutoFill failed for '{col_name}': {exc}")
        else:
            print(f"  {last_row - 1} data row(s) — AutoFill not needed.")

        xl.CalculateFull()
        xl.CalculateFull()   # second pass ensures formula-driven Final Type is stable
        wb_com.Save()
        print(f"\n  Saved after Pass 4.")

        # ==================================================================
        # PASS 4B — Remove non-Write-Back rows after AutoFill recalc
        #
        # COM row deletion is too slow on 50k+ rows. Strategy:
        #   1. Bulk-read Final Type + all data columns via COM (values only).
        #   2. Close COM session.
        #   3. Reopen with openpyxl, overwrite sheet with only KEEP rows.
        #   4. Re-open COM for passes 5-9.
        # ==================================================================
        _section("Pass 4B — Remove non-Write-Back rows after AutoFill recalc")

        ws_com      = _get_sheet(wb_com, WORKING_SHEET)
        com_col_map = _build_col_map_com(ws_com, 1)
        ft_col_new  = com_col_map.get("Final Type")

        if not ft_col_new:
            print("  WARNING: 'Final Type' column not found — Pass 4B skipped.")
            wb_com.Save()
            print(f"  Saved after Pass 4B.")
        else:
            last_row  = ws_com.Cells(ws_com.Rows.Count, 1).End(XL_UP).Row
            n_cols    = ws_com.UsedRange.Columns.Count

            print(f"  Reading {last_row - 1:,} data rows x {n_cols} cols via COM ...")
            # Bulk-read entire data range in one COM call (values only, no formulas)
            all_data = ws_com.Range(
                ws_com.Cells(1, 1),
                ws_com.Cells(last_row, n_cols)
            ).Value2

            # Normalise to list-of-lists
            if not isinstance(all_data, tuple):
                all_data = ((all_data,),)
            rows_raw = [
                list(r) if isinstance(r, tuple) else [r]
                for r in all_data
            ]

            header_row = rows_raw[0]
            ft_idx     = ft_col_new - 1   # 0-based

            ft_counts: dict = {}
            keep_rows: list = []
            for r in rows_raw[1:]:        # skip header
                ft_val = r[ft_idx] if ft_idx < len(r) else None
                ft_counts[ft_val] = ft_counts.get(ft_val, 0) + 1
                if _is_keep_type(ft_val):
                    keep_rows.append(r)

            # Build all lines first, then print once — see the matching note
            # on the earlier "Final Type value counts in source" report.
            ft_report_lines = ["  Final Type values after AutoFill recalc:"]
            for val, cnt in sorted(ft_counts.items(), key=lambda x: str(x[0])):
                tag  = " KEPT  " if _is_keep_type(val) else "DELETED"
                disp = str(val) if val is not None else "(blank/None)"
                ft_report_lines.append(f"    [{tag}]  {disp:<32}  {cnt:>5} rows   raw={repr(val)}")
            print("\n".join(ft_report_lines))

            deleted = (last_row - 1) - len(keep_rows)
            print(f"\n  Rows to delete : {deleted:,}  |  Rows to keep : {len(keep_rows):,}")

            # Close COM before openpyxl rewrite
            wb_com.Save()
            wb_com.Close(SaveChanges=False)
            xl.Quit()
            xl = None
            print(f"  COM session closed. Rewriting sheet with openpyxl ...")

            # Rewrite sheet: header + keep_rows only
            wb4b = openpyxl.load_workbook(dest_path)
            ws4b_name = WORKING_SHEET
            if ws4b_name not in wb4b.sheetnames:
                ws4b_name = next(
                    (s for s in wb4b.sheetnames
                     if s.strip().lower() == WORKING_SHEET.strip().lower()),
                    wb4b.sheetnames[0]
                )
            ws4b = wb4b[ws4b_name]

            # Delete all existing data rows (keep only row 1 = header)
            if ws4b.max_row > 1:
                ws4b.delete_rows(2, ws4b.max_row - 1)

            # Write keep_rows back
            for r_data in keep_rows:
                ws4b.append(r_data)

            wb4b.save(dest_path)
            wb4b.close()
            print(f"  Rewrite complete. {len(keep_rows):,} rows retained.")

            # Re-launch COM for passes 5-9
            xl = _launch_excel()
            wb_com = xl.Workbooks.Open(
                os.path.abspath(dest_path),
                UpdateLinks=XL_UPDATE_YES
            )
            xl.CalculateFull()
            print(f"  COM re-launched for passes 5-9.")
            print(f"  Saved after Pass 4B.")

        # ==================================================================
        # PASS 5 — Refresh NetOff_Pivot
        # ==================================================================
        _section("Pass 5 — Refresh NetOff_Pivot")
        _refresh_pivot_sheet(wb_com, NETOFF_PIVOT_SHEET, WORKING_SHEET)
        wb_com.Save()
        print(f"  Saved after Pass 5.")

        # ==================================================================
        # PASS 6 — Move Txn Net Amount = 0 rows to NetOff_Zero_Txns
        # ==================================================================
        _section("Pass 6 — Move Txn Net Amount = 0 rows to NetOff_Zero_Txns")

        ws_com  = _get_sheet(wb_com, WORKING_SHEET)
        ws_zero = _get_sheet(wb_com, NETOFF_ZERO_SHEET)

        # Clear the entire NetOff_Zero_Txns sheet (records + old headers)
        ws_zero.Cells.Clear()
        print(f"  Cleared '{NETOFF_ZERO_SHEET}'.")

        # Rebuild column map (row count may have changed)
        com_col_map = _build_col_map_com(ws_com, 1)
        txn_net_col = com_col_map.get("Txn Net Amount")

        if not txn_net_col:
            print("  WARNING: 'Txn Net Amount' column not found — Pass 6 skipped.")
        else:
            # Write MOVE_COLUMNS headers to NetOff_Zero_Txns row 1
            for dest_c, col_name in enumerate(MOVE_COLUMNS, start=1):
                ws_zero.Cells(1, dest_c).Value = col_name

            last_row = ws_com.Cells(ws_com.Rows.Count, 1).End(XL_UP).Row

            # Bulk-read Txn Net Amount column (one COM call)
            net_letter = _col_letter(txn_net_col)
            raw_net    = ws_com.Range(f"{net_letter}2:{net_letter}{last_row}").Value2
            if raw_net is None:
                flat_net = [None] * (last_row - 1)
            elif not isinstance(raw_net, tuple):
                flat_net = [raw_net]
            else:
                flat_net = [r[0] if isinstance(r, tuple) else r for r in raw_net]

            zero_rows = []
            for i, val in enumerate(flat_net):
                try:
                    if float(val) == 0.0:
                        zero_rows.append(2 + i)
                except (TypeError, ValueError):
                    if str(val).strip() in ("0", "0.0"):
                        zero_rows.append(2 + i)

            print(f"  Zero Txn Net Amount rows found : {len(zero_rows):,}")

            if zero_rows:
                _bulk_copy_and_delete(
                    ws_com, ws_zero, zero_rows, com_col_map, last_row, dest_start=2
                )

            print(f"  Moved {len(zero_rows):,} zero-Txn rows "
                  f"-> '{NETOFF_ZERO_SHEET}'.")

        wb_com.Save()
        print(f"  Saved after Pass 6.")

        # ==================================================================
        # PASS 7 — Append OVERPAYMENT SETTLEMENT rows to NetOff_Zero_Txns
        # ==================================================================
        _section("Pass 7 — Append OVERPAYMENT SETTLEMENT rows to NetOff_Zero_Txns")

        ws_com  = _get_sheet(wb_com, WORKING_SHEET)
        ws_zero = _get_sheet(wb_com, NETOFF_ZERO_SHEET)

        # Rebuild column map after Pass 6 deletions
        com_col_map    = _build_col_map_com(ws_com, 1)
        trans_desc_col = com_col_map.get("Trans Description")

        if not trans_desc_col:
            print("  WARNING: 'Trans Description' not found — Pass 7 skipped.")
        else:
            last_row = ws_com.Cells(ws_com.Rows.Count, 1).End(XL_UP).Row

            # Bulk-read Trans Description column (one COM call)
            td_letter = _col_letter(trans_desc_col)
            raw_td    = ws_com.Range(f"{td_letter}2:{td_letter}{last_row}").Value2
            if raw_td is None:
                flat_td = [None] * (last_row - 1)
            elif not isinstance(raw_td, tuple):
                flat_td = [raw_td]
            else:
                flat_td = [r[0] if isinstance(r, tuple) else r for r in raw_td]

            op_rows = [
                2 + i for i, v in enumerate(flat_td)
                if isinstance(v, str) and v.strip().upper() == "OVERPAYMENT SETTLEMENT"
            ]

            print(f"  OVERPAYMENT SETTLEMENT rows found : {len(op_rows):,}")

            if op_rows:
                dest_start = ws_zero.Cells(ws_zero.Rows.Count, 1).End(XL_UP).Row + 1
                _bulk_copy_and_delete(
                    ws_com, ws_zero, op_rows, com_col_map, last_row, dest_start
                )

            print(f"  Moved {len(op_rows):,} OVERPAYMENT rows "
                  f"-> '{NETOFF_ZERO_SHEET}'.")

        wb_com.Save()
        print(f"  Saved after Pass 7.")

        # ==================================================================
        # PASS 8 — Refresh Final Pivot
        # ==================================================================
        _section("Pass 8 — Refresh Final Pivot")
        _refresh_pivot_sheet(wb_com, FINAL_PIVOT_SHEET, WORKING_SHEET)
        wb_com.Save()
        print(f"  Saved after Pass 8.")

        # ==================================================================
        # PASS 9 — Final recalculate, save and close
        # ==================================================================
        _section("Pass 9 — Final recalculate & close")
        xl.CalculateFull()
        wb_com.Save()
        # Mark as saved so Close() does not show a save-changes prompt
        wb_com.Saved = True
        wb_com.Close(SaveChanges=False)
        print(f"\n  Closed : {os.path.basename(dest_path)}")

    finally:
        try:
            xl.Quit()
        except Exception:
            pass

    return dest_path


# ============================================================================
# ENTRY POINT
# ============================================================================

def main():
    _section("WRITE OFF — FILTERED FILE CREATION")

    if len(sys.argv) >= 2:
        arg1 = sys.argv[1].strip()
        if _looks_like_run_date(arg1):
            # Orchestration may pass run date as first arg; write_off does not need it.
            try:
                source_path = find_source_file(LGD_DIR)
            except FileNotFoundError as exc:
                print(f"\nERROR: {exc}")
                sys.exit(1)
        else:
            source_path = arg1
            if not os.path.isabs(source_path):
                source_path = os.path.join(LGD_DIR, source_path)
            if not os.path.exists(source_path):
                print(f"\nERROR: Source file not found:\n  {source_path}")
                sys.exit(1)
    else:
        try:
            source_path = find_source_file(LGD_DIR)
        except FileNotFoundError as exc:
            print(f"\nERROR: {exc}")
            sys.exit(1)

    try:
        output_file = create_write_off_filtered_file(source_path)
    except Exception as exc:
        print(f"\nERROR: {exc}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # Remove original source file and rename _New → original name (one file remains)
    try:
        final_path = output_file.replace("_New - ", " - ", 1)
        if os.path.exists(source_path) and os.path.abspath(source_path) != os.path.abspath(output_file):
            os.remove(source_path)
            print(f"  Removed source file : {os.path.basename(source_path)}")
        if os.path.exists(final_path) and os.path.abspath(final_path) != os.path.abspath(output_file):
            os.remove(final_path)
        os.rename(output_file, final_path)
        output_file = final_path
        print(f"  Renamed to : {os.path.basename(final_path)}")
    except Exception as exc:
        print(f"  WARNING: Rename to final name failed: {exc}")

    # Copy Write Off Working*.xlsx and First Default Date Working*.xlsx → sibling LGD/ folder
    import shutil as _shutil2
    lgd_sub = os.path.join(os.path.dirname(LGD_DIR), "LGD")
    os.makedirs(lgd_sub, exist_ok=True)

    # Write Off Working: remove all Write Off Working*.xlsx in LGD/, then copy final file
    try:
        for old in glob.glob(os.path.join(lgd_sub, "Write Off Working*.xlsx")):
            try:
                os.remove(old)
            except Exception:
                pass
        _shutil2.copy2(output_file, os.path.join(lgd_sub, os.path.basename(output_file)))
        print(f"  Write Off Working file copied to LGD folder: {os.path.basename(output_file)}")
    except Exception as exc:
        print(f"  WARNING: Write Off Working copy to LGD failed: {exc}")

    # First Default Date Working: remove all First Default Date*.xlsx in LGD/, then copy
    try:
        fdd_matches = [
            f for f in glob.glob(os.path.join(LGD_DIR, "First Default Date Working*.xlsx"))
            if not os.path.basename(f).startswith("~$")
        ]
        if fdd_matches:
            fdd_src = max(fdd_matches, key=os.path.getmtime)
            for old in glob.glob(os.path.join(lgd_sub, "First Default Date*.xlsx")):
                try:
                    os.remove(old)
                except Exception:
                    pass
            _shutil2.copy2(fdd_src, os.path.join(lgd_sub, os.path.basename(fdd_src)))
            print(f"  First Default Date file copied to LGD folder: {os.path.basename(fdd_src)}")
        else:
            print("  WARNING: No 'First Default Date Working*.xlsx' found in firstDefault — skipping copy.")
    except Exception as exc:
        print(f"  WARNING: First Default Date copy to LGD failed: {exc}")

    _section("COMPLETED")
    print(f"  Output file : {os.path.basename(output_file)}")
    print(f"  Location    : {os.path.dirname(output_file)}")
    print()


if __name__ == "__main__":
    main()
