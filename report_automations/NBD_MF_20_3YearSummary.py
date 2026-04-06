import os
from pathlib import Path
import openpyxl
from openpyxl.styles import numbers
from copy import copy
import sys
from datetime import datetime
import calendar
import xlwings as xw
import logging

def setup_logging():
    """Setup organized logging for 3 Year Summary report"""
    # Configure logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)

    # Create organized log directory: logs/YYYY-MM-DD/frequency/report_name.log
    try:
        frequency = "monthly"
        date_folder = datetime.now().strftime('%Y-%m-%d')
        run_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        logs_dir = Path(__file__).parent.parent / "logs" / date_folder / frequency
        logs_dir.mkdir(parents=True, exist_ok=True)
        run_log_file = logs_dir / f"NBD_MF_20_3YearSummary_{run_timestamp}.log"
    except Exception:
        # Fallback to simple structure
        logs_dir = Path(__file__).parent / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        run_log_file = logs_dir / f"NBD_MF_20_3YearSummary_{run_timestamp}.log"

    file_handler = logging.FileHandler(run_log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)

    logger.info(f"Detailed log file: {run_log_file}")
    return logger

def get_month_year_from_filename(filename):
    # Example: "NBD-MF-20-C1 to C6 July 2025.xlsx"
    parts = filename.split()
    for i, part in enumerate(parts):
        if part in {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","January","February","March","April","May","June","July","August","September","October","November","December"}:
            month = part
            year = parts[i+1].replace(".xlsx","")
            return month, year
    return None, None

def find_latest_working_folder(base_dir, report_name):
    """Return the latest date folder inside working/<report_name>."""
    working_report_dir = base_dir / "working" / report_name
    if not working_report_dir.exists():
        print(f"ERROR: Working folder not found: {working_report_dir}")
        return None

    subfolders = [f for f in working_report_dir.iterdir() if f.is_dir()]
    if not subfolders:
        print(f"ERROR: No subfolders found under {working_report_dir}")
        return None

    # Since there is only one date folder in working, just return it
    latest_folder = subfolders[0]
    print(f"Working folder for {report_name}: {latest_folder}")
    return latest_folder


def find_file(base_dir, pattern, description):
    """Safely find a single file using glob and provide detailed errors."""
    print(f"\nLooking for {description} in: {base_dir}")
    files = list(base_dir.glob(pattern))
    if not files:
        print(f"No files found for {description} with pattern: {pattern}")
        all_files = [f.name for f in base_dir.iterdir() if f.is_file()]
        if all_files:
            print("Available files in this folder:")
            for f in sorted(all_files):
                print(f"  - {f}")
        else:
            print("No files found in this folder.")
        return None
    print(f"Found {description}: {files[0].name}")
    return files[0]


def find_row_by_code(worksheet, code_value, code_column='A'):
    """Find row number by searching for a specific code in the specified column"""
    for row in range(1, worksheet.max_row + 1):
        cell_value = worksheet[f"{code_column}{row}"].value
        if cell_value and str(cell_value).strip() == str(code_value).strip():
            return row
    return None


def get_numeric_value(worksheet, cell_ref, description=""):
    """Safely extract numeric value from a cell, handling text/None values"""
    try:
        value = worksheet[cell_ref].value
        print(f"  {description} ({cell_ref}): {value}")
        
        if value is None:
            return 0
        elif isinstance(value, (int, float)):
            return value
        elif isinstance(value, str):
            try:
                clean_value = value.replace(',', '').strip()
                return float(clean_value)
            except ValueError:
                print(f"    Warning: Cell {cell_ref} contains text '{value}' - using 0")
                return 0
        else:
            print(f"    Warning: Cell {cell_ref} contains unexpected type {type(value)} - using 0")
            return 0
    except Exception as e:
        print(f"    Error reading {cell_ref}: {e} - using 0")
        return 0


def get_value_by_code(worksheet, code, value_column='H', code_column='A'):
    """Get value from a specific column by matching code in column A"""
    row = find_row_by_code(worksheet, code, code_column)
    if row is None:
        print(f"  WARNING: Code '{code}' not found")
        return 0

    cell_ref = f"{value_column}{row}"
    return get_numeric_value(worksheet, cell_ref, f"Value for code {code}")


def get_value_by_code_xlwings(worksheet_xw, code, value_column='H', code_column='A'):
    """Get value from xlwings worksheet by matching code in column A"""
    # Find the row with the matching code
    max_row = worksheet_xw.range(f'{code_column}1').end('down').row

    for row in range(1, max_row + 1):
        cell_value = worksheet_xw.range(f'{code_column}{row}').value
        if cell_value and str(cell_value).strip() == str(code).strip():
            # Found the code, get the value from the specified column
            value = worksheet_xw.range(f'{value_column}{row}').value
            print(f"  Value for code {code} ({value_column}{row}): {value}")

            if value is None:
                return 0
            elif isinstance(value, (int, float)):
                return value
            else:
                print(f"    Warning: Unexpected value type {type(value)} - using 0")
                return 0

    print(f"  WARNING: Code '{code}' not found")
    return 0


def parse_month_year(month_str):
    """Parse month string from various formats to a standardized format"""
    if not month_str:
        return None
    
    month_str = str(month_str).strip()
    
    # Try parsing various date formats
    formats = [
        "%m/%d/%Y",  # 7/31/2025 or 6/30/2025
        "%m-%d-%Y",  # 7-31-2025
        "%b-%y",     # Jul-25
        "%B-%Y",     # July-2025
        "%b %Y",     # Jul 2025
        "%B %Y",     # July 2025
    ]
    
    for fmt in formats:
        try:
            date_obj = datetime.strptime(month_str, fmt)
            return date_obj
        except ValueError:
            continue
    
    # Try parsing just "M/D/Y" formats with varying day values
    try:
        parts = month_str.split('/')
        if len(parts) == 3:
            month_num = int(parts[0])
            year_num = int(parts[2])
            # Create datetime with month and year (day doesn't matter for comparison)
            return datetime(year_num, month_num, 1)
    except:
        pass
    
    return None


def find_or_create_month_row(worksheet, target_month, target_year):
    """Find existing row for the month/year or create a new one as part of the table"""
    print(f"\nSearching for existing record: {target_month}-{target_year}")
    
    # Month mapping for comparison
    month_map = {
        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
        "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
        "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
        "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12
    }
    
    target_month_num = month_map.get(target_month, 0)
    target_year_num = int(target_year)
    
    # Find the last row with data in column A (the table)
    last_data_row = 1
    for row in range(2, worksheet.max_row + 1):  # Start from row 2 (assuming row 1 is header)
        cell_value = worksheet[f"A{row}"].value
        
        if cell_value:
            last_data_row = row
            # Try to parse the existing date
            parsed_date = parse_month_year(cell_value)
            
            if parsed_date:
                # Compare month and year
                if parsed_date.month == target_month_num and parsed_date.year == target_year_num:
                    print(f"  Found existing record at row {row}")
                    return row, True  # Row number, is_existing
    
    # If not found, add new row right after the last data row (append to table)
    new_row = last_data_row + 1
    
    # Copy formatting from the previous row (last_data_row)
    if last_data_row >= 2:
        print(f"  Copying formatting from row {last_data_row}")
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            source_cell = worksheet[f"{col}{last_data_row}"]
            target_cell = worksheet[f"{col}{new_row}"]
            
            # Copy cell formatting (font, fill, border, alignment, number format)
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.number_format = copy(source_cell.number_format)
    
    # Format the date as "Aug-25" format (Month-YY)
    # Map month number to short month name
    month_abbr_map = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }

    month_abbr = month_abbr_map.get(target_month_num, "")
    year_short = str(target_year_num)[-2:]  # Get last 2 digits of year

    date_string = f"{month_abbr}-{year_short}"

    worksheet[f"A{new_row}"] = date_string
    print(f"  Appended new record at row {new_row} (table row) with date: {date_string}")
    
    return new_row, False  # Row number, is_existing


def main():
    # Setup logging
    logger = setup_logging()

    base_dir = Path(__file__).resolve().parent.parent
    report_name = "NBD_MF_20_C1_C6"

    # Get the latest working folder
    working_folder = find_latest_working_folder(base_dir, report_name)
    if not working_folder:
        return

    # Find files inside the working folder
    file_c1_c6 = find_file(working_folder, "**/NBD-MF-20-C1 to C6*.xlsx", "C1-C6 file")
    file_afl = find_file(working_folder, "**/NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx", "AFL file")

    if not file_c1_c6 or not file_afl:
        print("Required files not found. Cannot continue.")
        return

    # Extract month and year from filename
    month, year = get_month_year_from_filename(file_afl.name)
    if not month or not year:
        print("Could not extract month/year from filename")
        return

    print("\n--- Summary ---")
    print(f"C1-C6 File: {file_c1_c6}")
    print(f"AFL File: {file_afl}")
    print(f"Report Month: {month}")
    print(f"Report Year: {year}")
    print(f"Working folder: {working_folder}")
    
    try:
        # Load C1-C6 workbook with openpyxl (for writing)
        print(f"\nLoading C1-C6 workbook with openpyxl...")
        wb_c1_c6 = openpyxl.load_workbook(file_c1_c6, data_only=False)

        # Check if required sheet exists in C1-C6 file
        print(f"Available sheets in C1-C6 file: {wb_c1_c6.sheetnames}")

        if "3 Year Summary" not in wb_c1_c6.sheetnames:
            print(f"Required sheet '3 Year Summary' not found in C1-C6 file")
            return

        summary_ws = wb_c1_c6["3 Year Summary"]

        # Load AFL file with xlwings (for reading calculated values)
        print(f"\nLoading AFL file with xlwings (to calculate formulas)...")
        app = xw.App(visible=False)

        try:
            wb_afl_xw = app.books.open(str(file_afl))

            # Check if SOCI sheet exists
            if "NBD-MF-02-SOCI" not in [sheet.name for sheet in wb_afl_xw.sheets]:
                print(f"Required sheet 'NBD-MF-02-SOCI' not found in AFL file")
                print(f"Available sheets: {[s.name for s in wb_afl_xw.sheets]}")
                wb_afl_xw.close()
                return

            soci_ws_xw = wb_afl_xw.sheets["NBD-MF-02-SOCI"]

            print(f"All required sheets found")

            print(f"\nProcessing 3 Year Summary data...")

            # Find or create the row for this month/year
            target_row, is_existing = find_or_create_month_row(summary_ws, month, year)

            if is_existing:
                print(f"Updating existing record for {month}-{year}")
            else:
                print(f"Creating new record for {month}-{year}")

            # Extract values from SOCI sheet using direct cell references (column H)
            # Reading calculated formula values from column H and dividing by 1000
            print(f"\nExtracting values from SOCI sheet (column H - calculated values divided by 1000)...")

            # Helper function to safely get value from xlwings cell and divide by 1000
            def get_cell_value(worksheet_xw, cell_ref):
                """Get numeric value from a specific cell using xlwings and divide by 1000"""
                try:
                    value = worksheet_xw.range(cell_ref).value
                    if value is None:
                        return 0
                    elif isinstance(value, (int, float)):
                        return value / 1000
                    else:
                        print(f"    Warning: Cell {cell_ref} contains non-numeric value - using 0")
                        return 0
                except Exception as e:
                    print(f"    Error reading {cell_ref}: {e} - using 0")
                    return 0

            # Column B: Net Interest Income (Code 1.1.1.0.0.0)
            value_b = get_cell_value(soci_ws_xw, 'H4')
            summary_ws[f"B{target_row}"] = value_b
            print(f"Column B (Net Interest Income from H4/1000): {value_b}")

            # Column C: Interest Expenses (Code 1.1.2.0.0.0)
            value_c = get_cell_value(soci_ws_xw, 'H13')
            summary_ws[f"C{target_row}"] = value_c
            print(f"Column C (Interest Expenses from H13/1000): {value_c}")

            # Column D: Sum of Net Fee and Commission Income + Net Other Operating Income
            # Code 1.2.0.0.0.0 (H31) + Code 1.4.0.0.0.0 (H51)
            value_d1 = get_cell_value(soci_ws_xw, 'H31')
            value_d2 = get_cell_value(soci_ws_xw, 'H51')
            value_d = value_d1 + value_d2
            summary_ws[f"D{target_row}"] = value_d
            print(f"Column D (Fee + Other Income from H31/1000 + H51/1000): {value_d1} + {value_d2} = {value_d}")

            # Column E: Net Gains/(Losses) from Financial Instruments (Code 1.3.0.0.0.0)
            value_e = get_cell_value(soci_ws_xw, 'H42')
            summary_ws[f"E{target_row}"] = value_e
            print(f"Column E (Net Gains/Losses from H42/1000): {value_e}")

            # Column F: Sum of various Gains/(Losses)
            # Code 1.4.1.0.0.0 (H52) + Code 1.4.2.0.0.0 (H53) + Code 1.4.3.0.0.0 (H54) + Code 1.4.4.0.0.0 (H55)
            value_f1 = get_cell_value(soci_ws_xw, 'H52')
            value_f2 = get_cell_value(soci_ws_xw, 'H53')
            value_f3 = get_cell_value(soci_ws_xw, 'H54')
            value_f4 = get_cell_value(soci_ws_xw, 'H55')
            value_f = value_f1 + value_f2 + value_f3 + value_f4
            summary_ws[f"F{target_row}"] = value_f
            print(f"Column F (Total Gains/Losses from H52/1000 + H53/1000 + H54/1000 + H55/1000): {value_f1} + {value_f2} + {value_f3} + {value_f4} = {value_f}")

            print(f"\n--- Update Summary ---")
            print(f"Row {target_row} updated for {month}-{year}")
            print(f"  B (Net Interest Income): {value_b}")
            print(f"  C (Interest Expenses): {value_c}")
            print(f"  D (Fee + Other Income): {value_d}")
            print(f"  E (Net Gains/Losses): {value_e}")
            print(f"  F (Total Gains/Losses): {value_f}")

            # Update formulas in J104:N106 - increment row references by 1
            print(f"\nUpdating formulas in range J104:N106...")
            import re

            for row in range(104, 107):  # Rows 104, 105, 106
                for col in ['J', 'K', 'L', 'M', 'N']:
                    cell_ref = f"{col}{row}"
                    cell = summary_ws[cell_ref]

                    # Check if cell has a formula
                    if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                        original_formula = cell.value

                        # Use regex to find and increment row numbers in cell references
                        # Pattern matches: letter(s) followed by number(s)
                        def increment_row(match):
                            col_letters = match.group(1)
                            row_num = int(match.group(2))
                            return f"{col_letters}{row_num + 1}"

                        # Replace all cell references like B87, D111, etc.
                        updated_formula = re.sub(r'([A-Z]+)(\d+)', increment_row, original_formula)

                        # Update the cell with the new formula
                        summary_ws[cell_ref] = updated_formula
                        print(f"  {cell_ref}: {original_formula} -> {updated_formula}")

            # Copy calculated values from J104:N106, clear J94:N96, and paste values
            print(f"\nCopying calculated values from J104:N106 to J94:N96...")

            # Step 1: Read calculated values from J104:N106
            calculated_values = {}
            for row in range(104, 107):  # Rows 104, 105, 106
                for col in ['J', 'K', 'L', 'M', 'N']:
                    cell_ref = f"{col}{row}"
                    cell = summary_ws[cell_ref]
                    # Get the calculated value (formulas already exist and calculated)
                    # Since formulas are just updated, we need to read the calculated result
                    calculated_values[f"{col}{row}"] = cell.value
                    print(f"  Read {cell_ref}: {cell.value}")

            # Step 2: Clear range J94:N96
            print(f"\nClearing range J94:N96...")
            for row in range(94, 97):  # Rows 94, 95, 96
                for col in ['J', 'K', 'L', 'M', 'N']:
                    cell_ref = f"{col}{row}"
                    summary_ws[cell_ref].value = None
                    print(f"  Cleared {cell_ref}")

            # Step 3: Paste values to J94:N96 (mapping J104->J94, J105->J95, J106->J96)
            print(f"\nPasting calculated values to J94:N96...")
            for source_row, target_row in zip(range(104, 107), range(94, 97)):
                for col in ['J', 'K', 'L', 'M', 'N']:
                    source_ref = f"{col}{source_row}"
                    target_ref = f"{col}{target_row}"
                    # Paste only the value, not the formula
                    value = calculated_values[source_ref]
                    summary_ws[target_ref] = value
                    print(f"  Pasted {source_ref} -> {target_ref}: {value}")

            print(f"[OK] Values copied from J104:N106 to J94:N96")

            # Close AFL workbook
            wb_afl_xw.close()

        finally:
            app.quit()

        # Save updated C1-C6 workbook back to the same original file
        wb_c1_c6.save(file_c1_c6)
        print(f"\nReport saved successfully to original location: {file_c1_c6}")
        
    except FileNotFoundError as e:
        print(f"File not found: {e}")
    except PermissionError as e:
        print(f"Permission error (file might be open in Excel): {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='NBD-MF-20 3 Year Summary Report Automation')
    parser.add_argument('--working-dir', type=str, help='Working directory (not used, for compatibility)')
    parser.add_argument('--month', type=str, help='Report month (e.g., Jul)')
    parser.add_argument('--year', type=str, help='Report year (e.g., 2025)')
    args = parser.parse_args()

    main()
