import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re
from datetime import datetime
import os
import glob
from datetime import datetime, timedelta    
import argparse
from pathlib import Path

SCRIPT_ROOT = os.path.dirname(os.path.abspath(__file__))

parent_directory = Path(SCRIPT_ROOT).parent

# Allow passing the <date-folder> explicitly; fallback to auto-detect
parser = argparse.ArgumentParser(description='Process summary using new folder structure')
parser.add_argument('--working-dir', type=str, help='Path to working date folder (working/NBD-QF-23-C3-C10-SF-FI/<date-folder>)')
args, unknown = parser.parse_known_args()
if args.working_dir:
    provided_path = Path(args.working_dir)
    if not provided_path.exists() or not provided_path.is_dir():
        raise FileNotFoundError(f"Working directory not found: {provided_path}")
    # If the provided path is the base folder (working/NBD-QF-23-C3-C10-SF-FI), pick its single <date-folder>
    subdirs = [d for d in provided_path.iterdir() if d.is_dir()]
    if subdirs:
        if len(subdirs) > 1:
            print(f"Warning: multiple <date-folder>s found under base directory, using the first: {subdirs[0].name}")
        date_folder = subdirs[0]
    else:
        # Treat the provided path as the <date-folder>
        date_folder = provided_path
    output_path = date_folder
else:
    # New structure: working/NBD-QF-23-C3-C10-SF-FI/<date-folder>/ (exactly one folder)
    base_working_dir = parent_directory / "outputs" / "NBD-QF-23-C3-C10-SF-FI"
    if not base_working_dir.exists():
        raise FileNotFoundError(f"Working directory not found: {base_working_dir}")
    
    # Look for completed_dates.txt file
    completed_dates_file = base_working_dir / "completed_dates.txt"
    if not completed_dates_file.exists():
        raise FileNotFoundError(f"completed_dates.txt not found in: {base_working_dir}")
    
    # Read the date from the text file (get the last line)
    with open(completed_dates_file, 'r') as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]
    
    if not lines:
        raise ValueError(f"completed_dates.txt is empty in: {base_working_dir}")
    
    # Get the last date from the file
    target_date = lines[-1]
    print(f"Target date from completed_dates.txt: {target_date}")
    
    # Find all folders that match the target date (including versioned ones)
    date_folders = []
    pattern = re.compile(rf"^{re.escape(target_date)}(\(\d+\))?$")
    
    for folder in base_working_dir.iterdir():
        if folder.is_dir() and pattern.match(folder.name):
            date_folders.append(folder)
    
    if not date_folders:
        raise FileNotFoundError(f"No folder found for date '{target_date}' in: {base_working_dir}")
    
    # Sort folders to get the highest version
    # Folders are sorted by: base date, then version number (0 if no version)
    def get_version_number(folder_name):
        """Extract version number from folder name like '31-07-2025(1)' -> 1, or '31-07-2025' -> 0"""
        match = re.search(r'\((\d+)\)$', folder_name)
        return int(match.group(1)) if match else 0
    
    # Sort by version number (highest first)
    date_folders.sort(key=lambda f: get_version_number(f.name), reverse=True)
    
    # Select the folder with the highest version
    date_folder = date_folders[0]
    print(f"Selected date folder (highest version): {date_folder.name}")
    
    # Store the output path
    output_path = date_folder

print(f"Using date folder: {date_folder}")
print(f"Output path: {output_path}")

# Get current date for file name generation
now = datetime.now()
current_month_date = now.replace(day=1) - timedelta(days=1)  # Last day of previous month

# Folder paths (operate strictly within <date-folder>)
INPUT_FOLDER = date_folder / "Input" / "MT"
OUTPUT_FOLDER = date_folder

print(f"Input folder: {INPUT_FOLDER}")
print(f"Output folder: {OUTPUT_FOLDER}")


def unmerge_cells(file_path):
    """
    Unmerge all cells in the Excel file and fill them with the merged cell's value.

    Args:
    - file_path: Path to the Excel file
    """
    wb = load_workbook(file_path)

    for sheet in wb.worksheets:
        # Get all merged cell ranges
        merged_cells = list(sheet.merged_cells.ranges)

        # Unmerge and fill each range
        for merged_range in merged_cells:
            # Get the value from the top-left cell
            min_col, min_row, max_col, max_row = merged_range.bounds
            top_left_value = sheet.cell(min_row, min_col).value

            # Unmerge the cells
            sheet.unmerge_cells(str(merged_range))

            # Fill all cells in the range with the value
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    sheet.cell(row, col).value = top_left_value

    # Save the workbook
    wb.save(file_path)
    wb.close()


def consolidate_mt_reports():
    """
    Consolidate three MT reports into a single dataframe with opening balance from Jun 2025 MT report.

    Returns:
    - consolidated_df: DataFrame with MT No., Client Name, Opening Balance (Mar), and Actual Lending from each report
    """
    from datetime import datetime
    from calendar import monthrange

    # Get current date
    today = datetime.today()

    # Generate list of last 3 months excluding current month
    months = []
    current_date = today.replace(day=1)

    for i in range(1, 4):  # Get last 3 months (excluding current)
        # Go back i months
        if current_date.month - i <= 0:
            month = 12 + (current_date.month - i)
            year = current_date.year - 1
        else:
            month = current_date.month - i
            year = current_date.year

        # Get last day of that month
        last_day = monthrange(year, month)[1]
        date_obj = datetime(year, month, last_day)
        months.append(date_obj)

    # Reverse to get chronological order (oldest to newest)
    months.reverse()

    # Generate file paths dynamically using INPUT_FOLDER
    report_files = [
        INPUT_FOLDER / f"{date.strftime('%d-%m-%Y')}  Client Available Limits Report.xlsx"
        for date in months
    ]

    # Get the month before the oldest report month for opening balance
    # This will be 4 months back from current month
    oldest_month = months[0]
    if oldest_month.month - 1 <= 0:
        opening_month = 12
        opening_year = oldest_month.year - 1
    else:
        opening_month = oldest_month.month - 1
        opening_year = oldest_month.year

    opening_date = datetime(opening_year, opening_month, 1)
    mt_jun_2025_file = INPUT_FOLDER / f"MT - {opening_date.strftime('%b %Y')}.xlsx"

    # Column names for the output
    report_dates = [date.strftime('%d-%m-%Y') for date in months]

    # Read opening balance from MT - Jun 2025.xlsx (Dis worksheet, Mar column)
    print(f"Reading opening balance from {mt_jun_2025_file}...")
    opening_balance_df = pd.read_excel(mt_jun_2025_file, sheet_name='Dis')

    # Extract MT No., Client Name, and Opening Balance (Mar column)
    opening_balance_df = opening_balance_df[['Unnamed: 0','Unnamed: 1', 'Unnamed: 2', opening_balance_df.columns[-1]]].copy()
    opening_balance_df.columns = ['MT No.','Gender', 'Client Name', 'Opening Balance']

    # Clean up any NaN values in MT No. or Client Name
    opening_balance_df = opening_balance_df.dropna(subset=['MT No.', 'Client Name'])

    print(f"Loaded opening balance for {len(opening_balance_df)} MT entries")

    # Read all three reports
    dfs = []
    for i, file_path in enumerate(report_files):
        if not file_path.exists():
            print(f"Warning: File not found - {file_path}")
            continue

        # Unmerge cells before reading
        print(f"Unmerging cells in {file_path}...")
        unmerge_cells(str(file_path))

        # Read Excel file and skip first 7 columns
        df = pd.read_excel(file_path, skiprows=7)

        # Display available columns for debugging
        if i == 0:
            print(f"Available columns in reports: {df.columns.tolist()}")

        # Select relevant columns (MT No., Client Name, Actual Lending)
        # Adjust column names if they differ in your files
        df_filtered = df[['MT No.', 'Client Name', 'Actual Lending(Rs.)']].copy()
        df_filtered.columns = ['MT No.', 'Client Name', f'Actual Lending {report_dates[i]}']

        dfs.append(df_filtered)

    if len(dfs) == 0:
        print("Error: No files were successfully loaded")
        return pd.DataFrame()

    # Merge all dataframes on MT No. and Client Name
    consolidated_df = dfs[0].reset_index(drop=True)

    for i in range(1, len(dfs)):
        consolidated_df = consolidated_df.merge(
            dfs[i],
            on=['MT No.', 'Client Name'],
            how='outer'
        ).reset_index(drop=True)

    # Merge opening balance as the 3rd column (after MT No. and Client Name)
    consolidated_df = consolidated_df.merge(
        opening_balance_df,
        on=['MT No.', 'Client Name'],
        how='left'
    ).reset_index(drop=True)

    # Reorder columns to place Column 2 and Opening Balance after Client Name
    cols = ['MT No.', 'Client Name', 'Gender', 'Opening Balance'] + [col for col in consolidated_df.columns if col not in ['MT No.', 'Client Name', 'Column 2', 'Opening Balance', 'Gender']]
    consolidated_df = consolidated_df[cols]

    # Fill NaN values with 0 for missing data
    lending_columns = [col for col in consolidated_df.columns if 'Actual Lending' in col or 'Opening Balance' in col]
    consolidated_df[lending_columns] = consolidated_df[lending_columns].fillna(0)

    # Sort by MT No. and reset index
    consolidated_df = consolidated_df.sort_values('MT No.').reset_index(drop=True)

    # Identify missing records (records that have NaN in Gender or Opening Balance)
    # Create boolean masks separately and use them on the same DataFrame to avoid index alignment issues
    gender_mask = consolidated_df['Gender'].isna()
    balance_mask = consolidated_df['Opening Balance'].isna()
    missing_records = consolidated_df[gender_mask | balance_mask].copy()

    # Export missing records to separate Excel file in OUTPUT_FOLDER
    if len(missing_records) > 0:
        missing_records_file = OUTPUT_FOLDER / "Missing_Records.xlsx"
        missing_records.to_excel(missing_records_file, index=False)
        print(f"\nFound {len(missing_records)} missing records")
        print(f"Missing records exported to: {missing_records_file}")
    else:
        print("\nNo missing records found")

    print(f"\nConsolidated Report Summary:")
    print(f"Total unique MT entries: {len(consolidated_df)}")
    print(f"\nDataFrame Shape: {consolidated_df.shape}")
    print(f"\nFirst few rows:")
    print(consolidated_df.head())

    # Generate filename with last month's name
    from datetime import datetime, timedelta
    import shutil

    # Get last month's name and year
    today = datetime.today()
    last_month_date = today.replace(day=1) - timedelta(days=1)
    last_month_name = last_month_date.strftime('%b %Y')  # e.g., "May 2025"

    # Extract base workbook name
    base_name = mt_jun_2025_file.stem  # "MT - Jun 2025"
    workbook_prefix = base_name.split(' - ')[0] if ' - ' in base_name else base_name  # "MT"

    # Create new filename with last month
    new_filename = f"{workbook_prefix} - {last_month_name}.xlsx"
    mt_output_file = INPUT_FOLDER / new_filename

    # Copy original file to new filename if it doesn't exist or if we're updating
    if mt_output_file != mt_jun_2025_file:
        shutil.copy2(mt_jun_2025_file, mt_output_file)
        print(f"Created new workbook: {mt_output_file}")

    # Write consolidated_df to the workbook file, Dis sheet, starting from A1
    print(f"\nWriting consolidated data to {mt_output_file}, sheet 'Dis'...")
    with pd.ExcelWriter(mt_output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        consolidated_df.to_excel(writer, sheet_name='Dis', startrow=0, startcol=0, index=False)
    print(f"Data written successfully to 'Dis' sheet starting from cell A1")

    # Create gender summary table
    print(f"\nCreating gender summary table...")
    gender_summary = consolidated_df.groupby('Gender').agg({
        'Client Name': 'count',
        'Opening Balance': 'sum'
    }).reset_index()
    gender_summary.columns = ['Gender', 'Count of Client Name', 'Sum of Opening Balance']
    print(f"Gender summary table:")
    print(gender_summary)

    # Copy values only from Dis sheet A1:C1000 to Disbursement sheet starting at A3
    print(f"\nCopying values from 'Dis' sheet (A1:C1000) to 'Disbursement' sheet (A3)...")
    wb = load_workbook(mt_output_file)

    # Check if sheets exist
    if 'Dis' not in wb.sheetnames:
        print("Warning: 'Dis' sheet not found")
    elif 'Disbursement' not in wb.sheetnames:
        print("Warning: 'Disbursement' sheet not found")
    else:
        dis_sheet = wb['Dis']
        disbursement_sheet = wb['Disbursement']

        # Copy values only from A1:C1000 in Dis sheet to A3:C1002 in Disbursement sheet
        for row in range(1, 1001):  # Rows 1 to 1000 in Dis sheet
            for col in range(1, 4):  # Columns A, B, C (1, 2, 3)
                cell_value = dis_sheet.cell(row=row, column=col).value
                disbursement_sheet.cell(row=row+2, column=col).value = cell_value

        print(f"Values copied successfully from 'Dis' A1:C1000 to 'Disbursement' A3:C1002")

        
        # Find the last row with data in column A of Disbursement sheet
        last_row = 3  # Start from row 3 (where data starts)
        for row in range(3, disbursement_sheet.max_row + 1):
            if disbursement_sheet.cell(row=row, column=1).value is not None:
                last_row = row
            else:
                break

        print(f"\nCopying formulas from D10, E10, F10, G10 to end of data (row {last_row})...")

        # Get formulas from row 10, columns D, E, F, G (4, 5, 6, 7)
        formula_columns = [4, 5, 6, 7]  # D, E, F, G
        source_row = 10

        print(f"Source formulas from row {source_row}:")
        for col in formula_columns:
            cell = disbursement_sheet.cell(row=source_row, column=col)
            if cell.value:
                print(f"  Column {get_column_letter(col)}{source_row}: {cell.value}")

        # Copy formulas to all rows from 11 to last_row
        from openpyxl.formula.translate import Translator

        for row in range(11, last_row):
            for col in formula_columns:
                source_cell = disbursement_sheet.cell(row=source_row, column=col)
                if source_cell.value and isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                    # Calculate row offset
                    row_offset = row - source_row
                    # Translate formula with proper row offset
                    translated_formula = Translator(source_cell.value, origin=f"{get_column_letter(col)}{source_row}").translate_formula(
                        f"{get_column_letter(col)}{row}", row_delta=row_offset, col_delta=0
                    )
                    disbursement_sheet.cell(row=row, column=col).value = translated_formula

        print(f"Formulas copied to rows 11 through {last_row}")

        # Add total row after the last data row
        total_row = last_row + 1
        disbursement_sheet.cell(row=total_row, column=1).value = "Total"

        print(f"\nAdding totals in row {total_row} for columns D, E, F, G...")

        # Add SUM formulas for columns D, E, F, G
        for col in formula_columns:
            col_letter = get_column_letter(col)
            sum_formula = f"=SUM({col_letter}3:{col_letter}{last_row})"
            disbursement_sheet.cell(row=total_row, column=col).value = sum_formula
            print(f"Column {col_letter}: {sum_formula}")

        # Read Disbursement sheet data (skip first 2 rows) and create pivot table
        print(f"\nReading Disbursement sheet data (skipping 2 rows) for pivot table...")
        disbursement_df = pd.read_excel(mt_output_file, sheet_name='Disbursement', skiprows=2)

        # Create pivot table with Gender as rows, Count of Client Name and Sum of Total
        pivot_table = disbursement_df.pivot_table(
            values=['Client Name', 'Total'],
            index='Male/Female/Other',
            aggfunc={'Client Name': 'count', 'Total': 'sum'}
        ).reset_index()

        # Rename columns for clarity
        pivot_table.columns = ['Male/Female/Other', 'Count of Client Name', 'Sum of Total']

        print(f"Pivot table created:")
        print(pivot_table)

        # Write pivot table to Disbursement sheet starting at J9
        print(f"\nWriting pivot table to Disbursement sheet at J9...")
        start_row = 9
        start_col = 10  # Column J

        # Write headers
        for col_idx, header in enumerate(pivot_table.columns):
            disbursement_sheet.cell(row=start_row, column=start_col + col_idx).value = header

        # Write data rows
        for row_idx, row_data in pivot_table.iterrows():
            for col_idx, value in enumerate(row_data):
                disbursement_sheet.cell(row=start_row + 1 + row_idx, column=start_col + col_idx).value = value

        print(f"Pivot table written to J9:L{start_row + len(pivot_table)}")


        wb.save(mt_output_file)
        print(f"\nWorkbook saved with formulas and totals")

    wb.close()

    return consolidated_df


if __name__ == "__main__":
    # Execute the consolidation
    result_df = consolidate_mt_reports()

    # Optionally save to Excel in OUTPUT_FOLDER
    output_file = OUTPUT_FOLDER / "MT_Consolidated_Report.xlsx"
    result_df.to_excel(output_file, index=False)
    print(f"\nConsolidated report saved to: {output_file}")