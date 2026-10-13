"""
Update AFIN Receipts Script
- Reads Consolidated Receipts by AFIN.xlsx as maindf
- Maps current period from receipt details (receipts dated on/before the run date)
- Maps Writeback Current period from writebacks
- Maps yard Current period from YARD_DF
- Recalculates Final Receipts
- Clears A3:end in the original workbook and writes maindf back at A3

Receipt figures are CUMULATIVE snapshots, not period slices: 'Receipts Until March
2025' is the cumulative DCF as at that date and 'current period' is the cumulative
DCF of every receipt in receipt details up to the run date. They overlap, so
Final Receipts takes the LARGER of the two (the same MAX rule the 'DCF Receipts'
sheet in LGD Final uses for 'Receipts to Consider'); adding them counts every
receipt already in the base twice. Write-backs and yard are additive.

The run date is taken from argv[1] (DD/MM/YYYY, passed by LGD_computation.py) or,
failing that, the LGD_RUN_DATE environment variable. Without one, no cut-off is
applied and a warning is printed.
"""

import sys
from datetime import datetime

import pandas as pd
import numpy as np
import os
import xlwings as xw
import glob

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _parse_cutoff():
    """Run date as a datetime, or None if it can't be determined."""
    raw = sys.argv[1].strip() if len(sys.argv) > 1 else ""
    if raw:
        try:
            return datetime.strptime(raw, "%d/%m/%Y")
        except ValueError:
            print(f"WARNING: could not parse run date argument '{raw}' as DD/MM/YYYY")
    env = os.environ.get("LGD_RUN_DATE", "").strip()
    if env:
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):   # ifrs.py sets M/D/YYYY
            try:
                return datetime.strptime(env, fmt)
            except ValueError:
                continue
        print(f"WARNING: could not parse LGD_RUN_DATE '{env}'")
    return None


CUTOFF = _parse_cutoff()


def _find_lgd_working_folder() -> str:
    """Return working/LGD/<latest-date-folder>/LGD/, same logic as LGD_computation.py."""
    base = os.path.join(os.path.dirname(_SCRIPT_DIR), "working", "LGD")
    if os.path.isdir(base):
        subfolders = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d))]
        if subfolders:
            subfolders.sort(key=lambda d: os.path.getmtime(os.path.join(base, d)), reverse=True)
            return os.path.join(base, subfolders[0], "LGD")
    return _SCRIPT_DIR


_lgd_base = os.environ.get("LGD_BASE_FOLDER", "").strip()
ROOT = os.path.join(_lgd_base, "LGD") if _lgd_base else _find_lgd_working_folder()
print(f"update_afin_receipts: using folder: {ROOT}")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: Read Consolidated Receipts by AFIN.xlsx
#         Header is at row index 2 (rows 0-1 are metadata)
# ─────────────────────────────────────────────────────────────────────────────
print("Reading Consolidated Receipts by AFIN.xlsx ...")
maindf = pd.read_excel(
    os.path.join(ROOT, "Consolidated Receipts by AFIN.xlsx"),
    header=2,          # row index 2 contains the actual column names
)
print(f"  maindf shape: {maindf.shape}")
print(f"  Columns: {maindf.columns.tolist()}")

# Ensure Fac 12 is clean string for matching
maindf["Fac 12"] = maindf["Fac 12"].astype(str).str.strip()

# Initialise new column
maindf["current period"] = np.nan

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: Read receipt details.xlsx, aggregate DCF by Facility 12
#         Map to maindf.current period where Fac 12 == Facility 12
# ─────────────────────────────────────────────────────────────────────────────
print("\nReading receipt details.xlsx ...")
pattern = os.path.join(ROOT, "receipt details*.xlsx")

files = glob.glob(pattern)
receiptdf = pd.read_excel(files[0])

receiptdf["Facility 12"] = receiptdf["Facility 12"].astype(str).str.strip()

# Ignore receipts dated after the run date. The receipts export runs well past the
# reporting date, and current period is a cumulative total, so anything later would
# leak into this period's figures.
if CUTOFF is not None:
    _rec_dates = pd.to_datetime(receiptdf["RSM_STLDDATE"], errors="coerce")
    _after = _rec_dates > CUTOFF
    print(f"  Run date {CUTOFF:%d-%m-%Y}: excluding {int(_after.sum()):,} of {len(receiptdf):,} "
          f"receipt rows dated after it")
    receiptdf = receiptdf[~_after].copy()
else:
    print("  WARNING: no run date supplied - receipts are NOT cut off at the run date")

# Aggregate DCF per Facility 12
receipt_agg = receiptdf.groupby("Facility 12")["DCF"].sum().reset_index()
receipt_map = receipt_agg.set_index("Facility 12")["DCF"].to_dict()

maindf["current period"] = maindf["Fac 12"].map(receipt_map)
matched_count = maindf["current period"].notna().sum()
print(f"  Matched {matched_count} rows in maindf")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: Find unmatched records in receiptdf
#         (Facility 12 not present in maindf.Fac 12)
#         Append to maindf with Fac 12=Facility 12, current period=DCF
# ─────────────────────────────────────────────────────────────────────────────
matched_fac12 = set(maindf["Fac 12"].dropna().unique())
unmatched_receipts = receiptdf[~receiptdf["Facility 12"].isin(matched_fac12)].copy()
print(f"  Unmatched receipt rows to append: {len(unmatched_receipts)}")

if not unmatched_receipts.empty:
    new_rows = pd.DataFrame({
        "Fac 12": unmatched_receipts["Facility 12"].values,
        "current period": unmatched_receipts["DCF"].values,
    })
    maindf = pd.concat([maindf, new_rows], ignore_index=True)

print(f"  maindf after appending unmatched: {maindf.shape[0]} rows")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: Read writebacks 1-04-2025 to 31-08-2025.xlsx
#         Map DCF (sum per Facility 12) to new column Writeback Current period
# ─────────────────────────────────────────────────────────────────────────────

pattern = os.path.join(ROOT, "writebacks*.xlsx")

files = glob.glob(pattern)
writebacksdf = pd.read_excel(files[0])
writebacksdf["Facility 12"] = writebacksdf["Facility 12"].astype(str).str.strip()

wb_agg = writebacksdf.groupby("Facility 12")["DCF"].sum().reset_index()
wb_map = wb_agg.set_index("Facility 12")["DCF"].to_dict()

maindf["Writeback Current period"] = maindf["Fac 12"].map(wb_map)
print(f"  Matched {maindf['Writeback Current period'].notna().sum()} rows")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5: Read YARD_DF 31-08-2025.xlsx
#         Map Amount (by Facility 12) to new column yard Current period
# ─────────────────────────────────────────────────────────────────────────────
print("\nReading YARD_DF 31-08-2025.xlsx ...")
pattern = os.path.join(ROOT, "YARD_DF*.xlsx")

files = glob.glob(pattern)
yarddf = pd.read_excel(files[0])

yarddf["Facility 12"] = yarddf["Facility 12"].astype(str).str.strip()

yard_map = yarddf.set_index("Facility 12")["Amount"].to_dict()
maindf["yard Current period"] = maindf["Fac 12"].map(yard_map)
print(f"  Matched {maindf['yard Current period'].notna().sum()} rows")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: Recalculate Final Receipts
#         = MAX(Receipts Until March 2025, current period)
#         + Write Backs Until March 2025
#         + Writeback Current period
#         + yard Current period
#
#         Both receipt columns are cumulative, so they are combined with MAX
#         (as 'Receipts to Consider' is in the DCF Receipts sheet) — summing
#         them would count every receipt already in the March 2025 base twice.
# ─────────────────────────────────────────────────────────────────────────────
print("\nRecalculating Final Receipts ...")

_num = lambda c: pd.to_numeric(maindf[c], errors="coerce").fillna(0)

_base = _num("Receipts Until March 2025")
_cur  = _num("current period")
receipts_to_consider = np.maximum(_base, _cur)
print(f"  current period > base on {int(((_cur > _base)).sum()):,} facilities, "
      f"base >= current period on {int(((_base >= _cur) & ((_base != 0) | (_cur != 0))).sum()):,}")
print(f"  Reduction in Final Receipts vs the old base + current period sum: "
      f"{float(np.minimum(_base, _cur).sum()):,.0f}")

additive_cols = [
    "Write Backs Until March 2025",
    "Writeback Current period",
    "yard Current period",
]

# Ensure numeric, treat NaN as 0 for summation
maindf["Final Receipts"] = receipts_to_consider + sum(_num(col) for col in additive_cols)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7: Reorder columns before writing
# ─────────────────────────────────────────────────────────────────────────────
col_order = [
    "Fac 12",
    "Dup Check",
    "Status",
    "Final Receipts",
    "Receipts Until March 2025",
    "current period",
    "Write Backs Until March 2025",
    "Writeback Current period",
    "yard Current period",
    "A",
]
# Only include columns that actually exist (guards against missing cols)
col_order = [c for c in col_order if c in maindf.columns]
maindf = maindf[col_order]
print(f"\nColumn order: {col_order}")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 8: Read writeoff 31-08-2025.xlsx → mark DC, remove Yes rows
# ─────────────────────────────────────────────────────────────────────────────
_writeoff_matches = glob.glob(os.path.join(ROOT, "writeoff *.xlsx"))
if not _writeoff_matches:
    raise FileNotFoundError(f"No writeoff *.xlsx found in {ROOT}")
writeoff_path = max(_writeoff_matches, key=os.path.getmtime)
print(f"\nReading writeoff file: {writeoff_path}")
writeoffdf = pd.read_excel(writeoff_path)
writeoffdf["Facility 12"] = writeoffdf["Facility 12"].astype(str).str.strip()
writeoff_fac12 = set(writeoffdf["Facility 12"].dropna().unique())

maindf["DC"] = maindf["Fac 12"].apply(lambda x: "Yes" if x in writeoff_fac12 else "No")
dc_yes = (maindf["DC"] == "Yes").sum()
print(f"  DC column added — {dc_yes} Yes, {len(maindf) - dc_yes} No")

maindf = maindf[maindf["DC"] != "Yes"].reset_index(drop=True)
print(f"  Removed {dc_yes} DC rows — {len(maindf)} rows remaining")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 9: Write maindf back into Consolidated Receipts by AFIN.xlsx
#         Clear A3 to end of data, then write headers + data starting at A3
# ─────────────────────────────────────────────────────────────────────────────
file_path = os.path.join(ROOT, "Consolidated Receipts by AFIN.xlsx")
print(f"\nOpening {file_path} with xlwings ...")

app = xw.App(visible=False)
try:
    wb = app.books.open(file_path)
    ws = wb.sheets[0]

    # Find last used row/column to determine clear range
    last_row = ws.used_range.last_cell.row
    last_col = ws.used_range.last_cell.column
    last_col_letter = xw.utils.col_name(last_col)

    clear_range = f"A3:{last_col_letter}{last_row}"
    print(f"  Clearing range {clear_range} ...")
    ws.range(clear_range).clear_contents()

    # Write column headers at row 3 (A3)
    headers = maindf.columns.tolist()
    ws.range("A3").value = headers

    # Write data starting at row 4 (A4)
    print(f"  Writing {len(maindf)} rows × {len(headers)} columns starting at A4 ...")
    data = maindf.where(pd.notna(maindf), None).values.tolist()
    ws.range("A4").value = data

    wb.save()
    print(f"  Saved: {file_path}")
    wb.close()
finally:
    app.quit()

print(f"\nDone. Final shape: {maindf.shape}")
print("\nSample (first 5 rows, key columns):")
sample_cols = [c for c in ["Fac 12", "Receipts Until March 2025", "current period",
                            "Write Backs Until March 2025", "Writeback Current period",
                            "yard Current period", "Final Receipts"] if c in maindf.columns]
print(maindf[sample_cols].head(5).to_string())
