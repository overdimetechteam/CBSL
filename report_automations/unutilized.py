import os

import pandas as pd
import numpy as np

# Unutilized sheet layout: header row 2, blank row 3, data from row 4.
UNUTILIZED_SHEET_START_ROW = 4


UNUTILIZED_SOURCE_COLUMNS = [
    'Facility Number', 'Client Name', 'Facility Amount', 'Amount Utilized',
    'Unutilized Loan Balances', 'Loan Type',
]


def _read_working_sheet_via_excel(unutilized_file_path, sheet_name):
    """Open the source in its own hidden Excel instance, force a full
    recalculation, and read the working sheet's calculated values (same layout
    as the pandas read: row 2 is the header, data from row 3)."""
    import shutil
    import tempfile
    import xlwings as xw

    # Work on a temp copy: the original may already be open in someone's Excel
    # (locked), which makes Excel refuse to open it a second time.
    temp_dir = tempfile.mkdtemp(prefix="unutilized_")
    temp_copy = os.path.join(temp_dir, "source.xlsx")
    shutil.copyfile(unutilized_file_path, temp_copy)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        # corrupt_load=1 (xlRepairFile): this source file fails a normal open
        # ("Open method of Workbooks class failed") but opens fine when Excel
        # repairs it on load, as it does when opened by hand.
        book = app.books.open(temp_copy, update_links=False, ignore_read_only_recommended=True, corrupt_load=1)
        app.calculate()
        rows = book.sheets[sheet_name].used_range.options(ndim=2).value
        book.close()
    finally:
        app.quit()
        shutil.rmtree(temp_dir, ignore_errors=True)

    df = pd.DataFrame([r[:6] for r in rows[2:]], columns=UNUTILIZED_SOURCE_COLUMNS)
    df = df.dropna(subset=['Facility Number']).copy()
    for col in ('Facility Amount', 'Amount Utilized', 'Unutilized Loan Balances'):
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df


def _read_unutilized_working_df(unutilized_file_path):
    """Read Facility Number/Client Name/Facility Amount/Amount Utilized/
    Unutilized Loan Balances/Loan Type from the Unutilized Credit Limits
    file's own working sheet (its name varies - 'Working' or 'Working-<Mon>' -
    so it's matched by prefix, same as NBD_MF_20_C4.py does for this file)."""
    sheet_names = pd.ExcelFile(unutilized_file_path).sheet_names
    working_sheet_name = next((s for s in sheet_names if s.startswith('Working')), None)
    if working_sheet_name is None:
        raise ValueError(
            f"No sheet starting with 'Working' found in '{unutilized_file_path}'. "
            f"Available sheets: {sheet_names}"
        )

    df = pd.read_excel(unutilized_file_path, sheet_name=working_sheet_name, skiprows=1)
    df = df.iloc[:, :6].copy()
    df.columns = UNUTILIZED_SOURCE_COLUMNS
    df = df.dropna(subset=['Facility Number']).copy()

    # Amount Utilized/Unutilized Loan Balances are formulas in the source; a
    # file that was never calculated+saved in Excel carries no stored results,
    # so they read back empty. Ask Excel to calculate them instead.
    formula_cols = ['Amount Utilized', 'Unutilized Loan Balances']
    if df[formula_cols].isna().any().any():
        print("Unutilized source has no stored values for some formula cells - recalculating in Excel...")
        df = _read_working_sheet_via_excel(unutilized_file_path, working_sheet_name)
        if df[formula_cols].isna().all().all():
            raise ValueError(
                f"'{working_sheet_name}' in '{unutilized_file_path}' has no values for Amount Utilized/"
                f"Unutilized Loan Balances even after recalculating in Excel - refusing to leave the "
                f"Unutilized sheet with its previous values."
            )
        if df[formula_cols].isna().any().any():
            print(f"WARNING: {int(df[formula_cols].isna().any(axis=1).sum())} facility row(s) still have "
                  f"blank Amount Utilized/Unutilized Loan Balances after recalculating")

    df['Facility Number'] = df['Facility Number'].astype('int64')
    return df.reset_index(drop=True)


def _cell_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (np.integer, np.floating)):
        return value.item()
    return value


def _is_totals_row(ws, row):
    """The sheet's bottom row is a =SUM(...) totals row (e.g. '=SUM(C4:C15)' in
    Facility Amount/Unutilized Loan Balances) rather than a facility row."""
    if ws.cell(row=row, column=1).value not in (None, ''):
        return False
    for col in (3, 5):
        value = ws.cell(row=row, column=col).value
        if isinstance(value, str) and value.startswith('='):
            return True
    return False


def update_unutilized_sheet(ws, unutilized_file_path):
    """
    Refreshes the 'Unutilized' sheet in place, matched by Facility Number.

    For a Facility Number already on the sheet, overwrites just Facility
    Amount/Amount Utilized/Unutilized Loan Balances/Loan Type (columns C-F) -
    Client Name and the Sector/Sub-sector/Stage/CBSL Sector columns already on
    that row are left untouched. A Facility Number with no existing row is
    added as a new row (Facility Number/Client Name plus those same 4
    columns; Sector/Sub-sector/Stage/CBSL Sector left blank).

    New rows are inserted above the sheet's trailing SUM(...) totals row
    (shifting it down) rather than appended below it, and that row's SUM
    ranges are widened to cover the newly added rows.
    """
    source_df = _read_unutilized_working_df(unutilized_file_path)

    existing_row_by_facility = {}
    last_data_row = UNUTILIZED_SHEET_START_ROW - 1
    totals_row = None
    for row in range(UNUTILIZED_SHEET_START_ROW, ws.max_row + 1):
        facility_no = ws.cell(row=row, column=1).value
        if facility_no is None or facility_no == '':
            if _is_totals_row(ws, row):
                totals_row = row
            continue
        existing_row_by_facility[int(facility_no)] = row
        last_data_row = row

    new_facilities = [
        source_row for _, source_row in source_df.iterrows()
        if int(source_row['Facility Number']) not in existing_row_by_facility
    ]

    if new_facilities and totals_row is not None:
        ws.insert_rows(totals_row, amount=len(new_facilities))
        insert_at = totals_row
        totals_row += len(new_facilities)
    else:
        insert_at = last_data_row + 1

    updated = 0
    added = 0
    next_new_row = insert_at

    for _, source_row in source_df.iterrows():
        facility_no = int(source_row['Facility Number'])

        if facility_no in existing_row_by_facility:
            target_row = existing_row_by_facility[facility_no]
            updated += 1
        else:
            target_row = next_new_row
            ws.cell(row=target_row, column=1, value=facility_no)
            ws.cell(row=target_row, column=2, value=_cell_value(source_row['Client Name']))
            next_new_row += 1
            last_data_row = max(last_data_row, target_row)
            added += 1

        # Assign .value (not ws.cell(..., value=...)): openpyxl ignores a
        # None passed to cell(), which silently kept stale numbers from an
        # earlier month whenever the source cell was blank.
        ws.cell(row=target_row, column=3).value = _cell_value(source_row['Facility Amount'])
        ws.cell(row=target_row, column=4).value = _cell_value(source_row['Amount Utilized'])
        ws.cell(row=target_row, column=5).value = _cell_value(source_row['Unutilized Loan Balances'])
        # Loan Type is only filled on the first row of each group in the source
        # and is a merged cell in the output, so only write it when present.
        loan_type = _cell_value(source_row['Loan Type'])
        if loan_type is not None:
            ws.cell(row=target_row, column=6).value = loan_type

    if new_facilities and totals_row is not None:
        ws.cell(row=totals_row, column=3).value = f'=SUM(C{UNUTILIZED_SHEET_START_ROW}:C{last_data_row})'
        ws.cell(row=totals_row, column=5).value = f'=SUM(E{UNUTILIZED_SHEET_START_ROW}:E{last_data_row})'

    print(f"Unutilized sheet: updated {updated} existing facilities, added {added} new facilities")
