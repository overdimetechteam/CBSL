import os
from pathlib import Path
import openpyxl
import sys
import subprocess

def get_month_year_from_filename(filename):
    # Example: "NBD-AF-20-C1 to C6 July 2025.xlsx"
    parts = filename.split()
    for i, part in enumerate(parts):
        if part in {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","January","February","March","April","May","June","July","August","September","October","November","December"}:
            month = part
            year = parts[i+1].replace(".xlsx","")
            return month, year

def find_files_safely(base_dir, pattern, description):
    """Safely find files and provide detailed error messages"""
    print(f"Looking for {description} in: {base_dir}")
    print(f"Search pattern: {pattern}")
    
    files = list(base_dir.glob(pattern))
    print(f"Found {len(files)} files matching pattern:")
    
    if not files:
        print(f"No files found matching pattern: {pattern}")
        print("Available files in directory:")
        try:
            all_files = [f.name for f in base_dir.iterdir() if f.is_file() and f.suffix == '.xlsx']
            if all_files:
                for file in sorted(all_files):
                    print(f"  - {file}")
            else:
                print("  No .xlsx files found in directory")
        except Exception as e:
            print(f"  Error listing files: {e}")
        return None
    
    for file in files:
        print(f"  {file.name}")
    
    return files[0]  # Return the first match

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


def prepare_output_folder(base_dir, report_name, working_folder):
    """Create output folder matching working folder name under outputs/<report_name>/"""
    output_report_dir = base_dir / "outputs" / report_name
    output_folder = output_report_dir / working_folder.name
    output_folder.mkdir(parents=True, exist_ok=True)
    print(f"Output folder created: {output_folder}")
    return output_folder

def read_tier2_values_from_master_data(base_dir):
    """Read Tier 2 capital values from Master_Data.xlsx file"""
    try:
        # Master_Data.xlsx is located at the base directory (not in working folder)
        master_data_file = base_dir / "Master_Data.xlsx"

        if not master_data_file.exists():
            print(f"Warning: Master_Data.xlsx not found at {master_data_file}. Using default values of 0 for Tier 2 capital.")
            return {
                "tier2_instruments": 0,
                "tier2_revaluation": 0,
                "tier2_provisions": 0
            }

        # Load the Master_Data.xlsx file
        print(f"\nLoading Master_Data.xlsx from: {master_data_file}")
        wb_master = openpyxl.load_workbook(master_data_file, data_only=True)

        # Check if NBD-MF-20-C1-C6 sheet exists
        if "NBD-MF-20-C1-C6" not in wb_master.sheetnames:
            print(f"Warning: 'NBD-MF-20-C1-C6' sheet not found in Master_Data.xlsx. Using default values of 0.")
            print(f"Available sheets: {wb_master.sheetnames}")
            return {
                "tier2_instruments": 0,
                "tier2_revaluation": 0,
                "tier2_provisions": 0
            }

        ws_master = wb_master["NBD-MF-20-C1-C6"]

        # Find the last row with data
        last_row = ws_master.max_row
        print(f"Master Data sheet has {last_row} rows")

        # Column mapping: A=Timestamp, B=Guarantees, C=Tier2 Instruments, D=Tier2 Revaluation, E=Tier2 Provisions
        # Read directly from known columns
        tier2_instruments_col = 3  # Column C
        tier2_revaluation_col = 4  # Column D
        tier2_provisions_col = 5   # Column E

        print(f"Reading from fixed columns: C (Tier 2 instruments), D (Revaluation gains), E (General provisions)")

        # Read values from the last row
        value = ws_master.cell(row=last_row, column=tier2_instruments_col).value
        tier2_instruments = float(value) if value is not None else 0
        print(f"Read Tier 2 instruments value: {tier2_instruments} from row {last_row}, column C")

        value = ws_master.cell(row=last_row, column=tier2_revaluation_col).value
        tier2_revaluation = float(value) if value is not None else 0
        print(f"Read Revaluation gains value: {tier2_revaluation} from row {last_row}, column D")

        value = ws_master.cell(row=last_row, column=tier2_provisions_col).value
        tier2_provisions = float(value) if value is not None else 0
        print(f"Read General provisions value: {tier2_provisions} from row {last_row}, column E")

        wb_master.close()

        return {
            "tier2_instruments": tier2_instruments,
            "tier2_revaluation": tier2_revaluation,
            "tier2_provisions": tier2_provisions
        }

    except Exception as e:
        print(f"Error reading Tier 2 values from Master_Data.xlsx: {e}")
        import traceback
        traceback.print_exc()
        print("Using default values of 0 for Tier 2 capital.")
        return {
            "tier2_instruments": 0,
            "tier2_revaluation": 0,
            "tier2_provisions": 0
        }


def read_capital_values_from_master_data(base_dir):
    """Read capital values from Master_Data.xlsx file based on audit completion status"""
    try:
        # Master_Data.xlsx is located at the base directory
        master_data_file = base_dir / "Master_Data.xlsx"

        if not master_data_file.exists():
            print(f"Warning: Master_Data.xlsx not found at {master_data_file}. Using default values of 0 for capital values.")
            return {
                "audit_completed": False,
                "stated_capital": 0,
                "statutory_reserve_fund": 0,
                "fair_value_through_other_comprehensive_income_reserve": 0,
                "revaluation_reserve": 0,
                "retained_earnings": 0
            }

        # Load the Master_Data.xlsx file
        print(f"\nLoading Master_Data.xlsx for capital values from: {master_data_file}")
        wb_master = openpyxl.load_workbook(master_data_file, data_only=True)

        # Check if NBD-MF-20-C1-C6 sheet exists
        if "NBD-MF-20-C1-C6" not in wb_master.sheetnames:
            print(f"Warning: 'NBD-MF-20-C1-C6' sheet not found in Master_Data.xlsx. Using default values of 0.")
            print(f"Available sheets: {wb_master.sheetnames}")
            return {
                "audit_completed": False,
                "stated_capital": 0,
                "statutory_reserve_fund": 0,
                "fair_value_through_other_comprehensive_income_reserve": 0,
                "revaluation_reserve": 0,
                "retained_earnings": 0
            }

        ws_master = wb_master["NBD-MF-20-C1-C6"]

        # Find the last row with data
        last_row = ws_master.max_row
        print(f"Reading capital values from row {last_row}")

        # Column mapping: F=audited?, G=stated_capital, H=statutory_reserve_fund,
        # I=fair_value_through_other_comprehensive_income_reserve, J=revaluation_reserve, K=retained_earnings

        # Read audit completion status from column F
        audit_status = ws_master.cell(row=last_row, column=6).value  # Column F
        audit_completed = str(audit_status).strip().upper() == "YES"
        print(f"Audit completed status: {audit_status} -> {audit_completed}")

        # If audit is not completed, use previous row's capital values
        if not audit_completed:
            print("Audit not completed. Using previous row's capital values.")
            # Find the previous row with "Yes" in column F
            previous_row = None
            for row in range(last_row - 1, 0, -1):
                prev_audit_status = ws_master.cell(row=row, column=6).value
                if str(prev_audit_status).strip().upper() == "YES":
                    previous_row = row
                    print(f"Found previous audited data at row {previous_row}")
                    break

            if previous_row:
                last_row = previous_row
            else:
                print("No previous audited data found. Using default values of 0.")
                wb_master.close()
                return {
                    "audit_completed": False,
                    "stated_capital": 0,
                    "statutory_reserve_fund": 0,
                    "fair_value_through_other_comprehensive_income_reserve": 0,
                    "revaluation_reserve": 0,
                    "retained_earnings": 0
                }

        # Read capital values from the identified row
        value = ws_master.cell(row=last_row, column=7).value  # Column G
        stated_capital = float(value) if value is not None else 0
        print(f"Read Stated Capital: {stated_capital} from row {last_row}, column G")

        value = ws_master.cell(row=last_row, column=8).value  # Column H
        statutory_reserve_fund = float(value) if value is not None else 0
        print(f"Read Statutory Reserve Fund: {statutory_reserve_fund} from row {last_row}, column H")

        value = ws_master.cell(row=last_row, column=9).value  # Column I
        fair_value_through_other_comprehensive_income_reserve = float(value) if value is not None else 0
        print(f"Read Fair Value through OCI Reserve: {fair_value_through_other_comprehensive_income_reserve} from row {last_row}, column I")

        value = ws_master.cell(row=last_row, column=10).value  # Column J
        revaluation_reserve = float(value) if value is not None else 0
        print(f"Read Revaluation Reserve: {revaluation_reserve} from row {last_row}, column J")

        value = ws_master.cell(row=last_row, column=11).value  # Column K
        retained_earnings = float(value) if value is not None else 0
        print(f"Read Retained Earnings: {retained_earnings} from row {last_row}, column K")

        wb_master.close()

        return {
            "audit_completed": audit_completed,
            "stated_capital": stated_capital,
            "statutory_reserve_fund": statutory_reserve_fund,
            "fair_value_through_other_comprehensive_income_reserve": fair_value_through_other_comprehensive_income_reserve,
            "revaluation_reserve": revaluation_reserve,
            "retained_earnings": retained_earnings
        }

    except Exception as e:
        print(f"Error reading capital values from Master_Data.xlsx: {e}")
        import traceback
        traceback.print_exc()
        print("Using default values of 0 for capital values.")
        return {
            "audit_completed": False,
            "stated_capital": 0,
            "statutory_reserve_fund": 0,
            "fair_value_through_other_comprehensive_income_reserve": 0,
            "revaluation_reserve": 0,
            "retained_earnings": 0
        }

def create_output_folder(working_folder):
    """Create output folder matching working folder structure"""
    # Get the base directory (parent of report_automations)
    base_dir = Path(__file__).resolve().parent.parent

    # Create outputs directory with same structure as working
    # working_folder structure: working/NBD_AF_20_C1_C6/YYYY-MM-DD/
    # outputs structure: outputs/NBD_AF_20_C1_C6/YYYY-MM-DD/

    report_name = working_folder.parent.name  # NBD_AF_20_C1_C6
    date_folder = working_folder.name  # YYYY-MM-DD

    output_folder = base_dir / "outputs" / report_name / date_folder
    output_folder.mkdir(parents=True, exist_ok=True)
    print(f"Output folder created: {output_folder}")
    return output_folder

def main():

    base_dir = Path(__file__).resolve().parent.parent
    report_name = "NBD_AF_20_C1_C6"  # Example report

    # Get the latest working folder (single date folder)
    working_folder = find_latest_working_folder(base_dir, report_name)
    if not working_folder:
        return

    # Find files inside the working folder
    # Try AF files first, then fallback to MF files (they share the same structure)
    file_c1_c6 = find_file(working_folder, "**/NBD-AF-20-C1 to C6*.xlsx", "C1-C6 file (AF)")
    if not file_c1_c6:
        print("AF C1-C6 file not found, trying MF C1-C6 file as fallback...")
        file_c1_c6 = find_file(working_folder, "**/NBD-MF-20-C1 to C6*.xlsx", "C1-C6 file (MF fallback)")

    file_afl = find_file(working_folder, "**/NBD-AF-01-SOFP & SOCI AFL Annual FS*.xlsx", "AFL file (AF)")
    is_af_file = bool(file_afl)
    if not file_afl:
        print("AF SOFP file not found, trying MF SOFP file as fallback...")
        file_afl = find_file(working_folder, "**/NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx", "AFL file (MF fallback)")
        is_af_file = False

    if not file_c1_c6 or not file_afl:
        print("Required files not found. Cannot continue.")
        return

    # Read Tier 2 capital values from Master_Data.xlsx
    tier2_values = read_tier2_values_from_master_data(base_dir)

    # Read capital values from Master_Data.xlsx
    capital_values = read_capital_values_from_master_data(base_dir)

    # Prepare output folder
    output_folder = prepare_output_folder(base_dir, report_name, working_folder)

    print("\n--- Summary ---")
    print(f"C1-C6 File: {file_c1_c6}")
    print(f"AFL File: {file_afl}")
    print(f"Output folder: {output_folder}")
    
    try:
        # Load workbooks
        print(f"\nLoading workbooks...")
        wb_c1_c6 = openpyxl.load_workbook(file_c1_c6, data_only=False)
        wb_afl = openpyxl.load_workbook(file_afl, data_only=True)
        print(f"Workbooks loaded successfully")
        
        # Check if required sheets exist
        print(f"\nAvailable sheets in C1-C6 file: {wb_c1_c6.sheetnames}")
        print(f"Available sheets in AFL file: {wb_afl.sheetnames}")
        
        # Check for C2 sheet in C1-C6 file
        if "NBD_NBL-MF-20-C2" not in wb_c1_c6.sheetnames:
            print(f"Required sheet 'NBD_NBL-MF-20-C2' not found in C1-C6 file")
            print(f"Available sheets: {wb_c1_c6.sheetnames}")
            return
        
        # Check for SOCI sheet in AFL file (use AF or MF sheet name based on file type)
        soci_sheet_name = "NBD-AF-02-SOCI" if is_af_file else "NBD-MF-02-SOCI"
        if soci_sheet_name not in wb_afl.sheetnames:
            print(f"Required sheet '{soci_sheet_name}' not found in AFL file")
            print(f"Available sheets: {wb_afl.sheetnames}")
            return
        
        print(f"All required sheets found")
        
        # Get worksheets
        src_ws = wb_afl[soci_sheet_name]
        tgt_ws = wb_c1_c6["NBD_NBL-MF-20-C2"]
        
        print(f"\nProcessing data...")
        
        # Helper function to safely get numeric values
        def get_numeric_value(worksheet, cell_ref, description=""):
            """Safely extract numeric value from a cell, handling text/None values"""
            try:
                value = worksheet[cell_ref].value
                print(f"{description} ({cell_ref}): {value}")
                
                if value is None:
                    return 0
                elif isinstance(value, (int, float)):
                    return value
                elif isinstance(value, str):
                    # Try to extract number from string if possible
                    try:
                        # Remove common text and try to parse number
                        clean_value = value.replace(',', '').strip()
                        return float(clean_value)
                    except ValueError:
                        print(f"  Warning: Cell {cell_ref} contains text '{value}' - using 0")
                        return 0
                else:
                    print(f"  Warning: Cell {cell_ref} contains unexpected type {type(value)} - using 0")
                    return 0
            except Exception as e:
                print(f"  Error reading {cell_ref}: {e} - using 0")
                return 0
        
        # Find "Total Comprehensive Income for the Year" row in AFL file
        print(f"\nSearching for 'Total Comprehensive Income for the Year' in SOCI sheet...")
        
        comprehensive_income_value = None
        comprehensive_income_row = None
        
        # First, find the "Total Comprehensive Income for the Year" row
        for row in range(1, src_ws.max_row + 1):
            for col in range(1, src_ws.max_column + 1):
                cell_value = src_ws.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str) and "Total Comprehensive Income for the Year" in cell_value:
                    print(f"Found 'Total Comprehensive Income for the Year' at row {row}, column {col}")
                    comprehensive_income_row = row
                    break
            if comprehensive_income_row:
                break
        
        if comprehensive_income_row is None:
            print("Could not find 'Total Comprehensive Income for the Year' row in AFL file")
            print("Available cell values in the SOCI sheet:")
            for row in range(1, min(20, src_ws.max_row + 1)):  # Show first 20 rows
                for col in range(1, min(10, src_ws.max_column + 1)):  # Show first 10 columns
                    cell_value = src_ws.cell(row=row, column=col).value
                    if cell_value:
                        print(f"  Row {row}, Col {col}: {cell_value}")
            return
        
        # Find the "Rs.'000" column and get the value
        print(f"Looking for 'Rs.'000' column value...")
        
        # Look for Rs.'000 column header
        rs_column = None
        for col in range(1, src_ws.max_column + 1):
            cell_value = src_ws.cell(row=1, column=col).value
            if cell_value and isinstance(cell_value, str) and "Rs.'000" in cell_value:
                print(f"Found 'Rs.'000' column at column {col}")
                rs_column = col
                break
        
        if rs_column is None:
            print("Could not find 'Rs.'000' column in AFL file")
            return
        
        # Get the comprehensive income value
        comprehensive_income_value = get_numeric_value(src_ws, f"{openpyxl.utils.get_column_letter(rs_column)}{comprehensive_income_row}", "Total Comprehensive Income for the Year")
        
        # Only update if the value is negative (loss)
        current_year_profit_loss = comprehensive_income_value
        current_year_updated = False
        
        if comprehensive_income_value < 0:
            print(f"Comprehensive income is negative ({comprehensive_income_value}), will update Current year's profit(losses)")
            
            # Find "Current year's profit(losses)" row in C2 sheet
            print(f"\nSearching for 'Current year's profit(losses)' in C2 sheet...")
            
            current_year_row = None
            for row in range(1, tgt_ws.max_row + 1):
                for col in range(1, tgt_ws.max_column + 1):
                    cell_value = tgt_ws.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and "Current year's profit(losses)" in cell_value:
                        print(f"Found 'Current year's profit(losses)' at row {row}, column {col}")
                        current_year_row = row
                        break
                if current_year_row:
                    break
            
            if current_year_row is None:
                print("Could not find 'Current year's profit(losses)' row in C2 sheet")
                print("Available cell values in C2 sheet:")
                for row in range(1, min(20, tgt_ws.max_row + 1)):  # Show first 20 rows
                    for col in range(1, min(10, tgt_ws.max_column + 1)):  # Show first 10 columns
                        cell_value = tgt_ws.cell(row=row, column=col).value
                        if cell_value:
                            print(f"  Row {row}, Col {col}: {cell_value}")
                return
            
            # Update the current year's profit(losses) value
            # Assuming the value goes in column C (3rd column)
            # CRITICAL: Skip C13 as it contains a formula that must be preserved
            target_cell = f"C{current_year_row}"
            if current_year_row == 13:
                print(f"WARNING: 'Current year's profit(losses)' found at row 13, but C13 contains a formula that must be preserved. Skipping update.")
                current_year_updated = False
            else:
                print(f"Updating cell {target_cell} with value: {current_year_profit_loss}")
                tgt_ws[target_cell] = current_year_profit_loss
                current_year_updated = True
        else:
            print(f"Comprehensive income is positive ({comprehensive_income_value}), NOT updating Current year's profit(losses)")
        
        # Now calculate Adjustments to Tier I capital values
        print(f"\nCalculating Adjustments to Tier I capital values...")
        
        # Find the required values in C2 sheet
        stated_capital = 0
        statutory_reserve_fund = 0
        retained_earnings = 0
        
        # Search for these values in the C2 sheet
        for row in range(1, tgt_ws.max_row + 1):
            for col in range(1, tgt_ws.max_column + 1):
                cell_value = tgt_ws.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str):
                    if "Stated Capital" in cell_value:
                        # Get value from column C of the same row
                        stated_capital = get_numeric_value(tgt_ws, f"C{row}", "Stated Capital")
                    elif "Statutory Reserve Fund" in cell_value:
                        # Get value from column C of the same row
                        statutory_reserve_fund = get_numeric_value(tgt_ws, f"C{row}", "Statutory Reserve Fund")
                    elif "Retained Earnings" in cell_value:
                        # Get value from column C of the same row
                        retained_earnings = get_numeric_value(tgt_ws, f"C{row}", "Retained Earnings")
        
        # Calculate Adjustments to Tier I capital
        adjustments_to_tier1 = stated_capital + statutory_reserve_fund + retained_earnings
        print(f"Adjustments to Tier I capital calculation:")
        print(f"  Stated Capital: {stated_capital}")
        print(f"  Statutory Reserve Fund: {statutory_reserve_fund}")
        print(f"  Retained Earnings: {retained_earnings}")
        print(f"  Total Adjustments to Tier I capital: {adjustments_to_tier1}")
        
        # Find "Adjustments to Tier I capital" row and update it
        print(f"\nSearching for 'Adjustments to Tier I capital' in C2 sheet...")

        adjustments_row = None
        for row in range(1, tgt_ws.max_row + 1):
            for col in range(1, tgt_ws.max_column + 1):
                cell_value = tgt_ws.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str) and "Adjustments to Tier I capital" in cell_value:
                    print(f"Found 'Adjustments to Tier I capital' at row {row}, column {col}")
                    adjustments_row = row
                    break
            if adjustments_row:
                break

        # Update the adjustments value
        # CRITICAL: Skip C13 as it contains a formula that must be preserved
        if adjustments_row is not None:
            if adjustments_row == 13:
                print(f"WARNING: 'Adjustments to Tier I capital' found at row 13, but C13 contains a formula that must be preserved. Skipping update.")
            else:
                adjustments_cell = f"C{adjustments_row}"
                print(f"Updating cell {adjustments_cell} with value: {adjustments_to_tier1}")
                tgt_ws[adjustments_cell] = adjustments_to_tier1
        else:
            print("Warning: Could not find 'Adjustments to Tier I capital' row in C2 sheet - skipping adjustment update")

        # MANDATORY: Update Tier 2 capital values from Master Data
        # These cells MUST be updated every time, regardless of other conditions
        print(f"\nUpdating Tier 2 capital values from Master Data...")
        print(f"  - Instruments qualified as Tier 2 capital: {tier2_values['tier2_instruments']}")
        print(f"  - Revaluation gains: {tier2_values['tier2_revaluation']}")
        print(f"  - General provisions/collective impairment allowances: {tier2_values['tier2_provisions']}")

        # CRITICAL: Cell C13 contains a formula that must be preserved
        # Store C13 formula before any updates
        c13_formula = tgt_ws['C13'].value
        print(f"\nPreserving C13 formula: {c13_formula}")

        print(f"\nUpdating Tier 2 capital values at fixed cell positions (MANDATORY)...")
        print(f"Note: Cell C13 contains a formula and will be preserved")

        # Update "Instruments qualified as Tier 2 capital" at C25
        tier2_instruments_cell = "C25"
        print(f"Updating cell {tier2_instruments_cell} (Instruments qualified as Tier 2 capital) with value: {tier2_values['tier2_instruments']}")
        tgt_ws[tier2_instruments_cell] = tier2_values['tier2_instruments']

        # Update "Revaluation gains" at C26
        tier2_revaluation_cell = "C26"
        print(f"Updating cell {tier2_revaluation_cell} (Revaluation gains) with value: {tier2_values['tier2_revaluation']}")
        tgt_ws[tier2_revaluation_cell] = tier2_values['tier2_revaluation']

        # Update "General provisions/collective impairment allowances" at C27
        tier2_provisions_cell = "C27"
        print(f"Updating cell {tier2_provisions_cell} (General provisions/collective impairment allowances - 20.2.4.3.0.0) with value: {tier2_values['tier2_provisions']}")
        tgt_ws[tier2_provisions_cell] = tier2_values['tier2_provisions']

        # Restore C13 formula if it was modified
        if c13_formula and tgt_ws['C13'].value != c13_formula:
            print(f"Restoring C13 formula: {c13_formula}")
            tgt_ws['C13'] = c13_formula
        else:
            print(f"C13 formula preserved: {tgt_ws['C13'].value}")

        # Update capital values from Master Data to cells G6-G10
        print(f"\nUpdating capital values from Master Data...")
        print(f"  - Stated Capital: {capital_values['stated_capital']}")
        print(f"  - Statutory Reserve Fund: {capital_values['statutory_reserve_fund']}")
        print(f"  - Fair Value through OCI Reserve: {capital_values['fair_value_through_other_comprehensive_income_reserve']}")
        print(f"  - Revaluation Reserve: {capital_values['revaluation_reserve']}")
        print(f"  - Retained Earnings: {capital_values['retained_earnings']}")

        # Direct cell updates (no search needed as positions are fixed: G6-G10)
        print(f"\nWriting capital values to C2 sheet cells G6-G10...")
        tgt_ws["G6"] = capital_values['stated_capital']
        print(f"✓ Updated G6 (Stated Capital): {capital_values['stated_capital']}")

        tgt_ws["G7"] = capital_values['statutory_reserve_fund']
        print(f"✓ Updated G7 (Statutory Reserve Fund): {capital_values['statutory_reserve_fund']}")

        tgt_ws["G8"] = capital_values['fair_value_through_other_comprehensive_income_reserve']
        print(f"✓ Updated G8 (Fair Value through OCI Reserve): {capital_values['fair_value_through_other_comprehensive_income_reserve']}")

        tgt_ws["G9"] = capital_values['revaluation_reserve']
        print(f"✓ Updated G9 (Revaluation Reserve): {capital_values['revaluation_reserve']}")

        tgt_ws["G10"] = capital_values['retained_earnings']
        print(f"✓ Updated G10 (Retained Earnings): {capital_values['retained_earnings']}")

        print(f"\nData processing completed")
        print(f"Updated C2 sheet:")
        if current_year_updated:
            print(f"  - Current year's profit(losses): {current_year_profit_loss} (UPDATED)")
        else:
            print(f"  - Current year's profit(losses): NOT UPDATED (comprehensive income was positive)")
        if adjustments_row:
            print(f"  - Adjustments to Tier I capital: {adjustments_to_tier1} (UPDATED)")
        else:
            print(f"  - Adjustments to Tier I capital: NOT UPDATED (row not found)")
        print(f"  - Tier 2 Capital Values (C25-C27) - MANDATORY UPDATES:")
        print(f"    • C25 Instruments qualified as Tier 2 capital: {tier2_values['tier2_instruments']} (from Master Data)")
        print(f"    • C26 Revaluation gains: {tier2_values['tier2_revaluation']} (from Master Data)")
        print(f"    • C27 General provisions/collective impairment allowances: {tier2_values['tier2_provisions']} (from Master Data)")
        print(f"  - Capital Values (G6-G10) - MANDATORY UPDATES:")
        print(f"    • G6 Stated Capital: {capital_values['stated_capital']}")
        print(f"    • G7 Statutory Reserve Fund: {capital_values['statutory_reserve_fund']}")
        print(f"    • G8 Fair Value through OCI Reserve: {capital_values['fair_value_through_other_comprehensive_income_reserve']}")
        print(f"    • G9 Revaluation Reserve: {capital_values['revaluation_reserve']}")
        print(f"    • G10 Retained Earnings: {capital_values['retained_earnings']}")

        # Save updated workbook back to the same original file (in-place like Ctrl+S)
        wb_c1_c6.save(file_c1_c6)
        print(f"\nReport saved successfully to original location: {file_c1_c6}")

        # IMPORTANT: Copy the file to outputs directory BEFORE running 3YearSummary
        # This ensures the downloaded file doesn't include 3YearSummary modifications
        print(f"\n{'='*80}")
        print(f"COPYING FILE TO OUTPUTS (BEFORE 3-YEAR SUMMARY)")
        print(f"{'='*80}")

        import shutil
        try:
            # Create outputs directory structure
            output_folder = create_output_folder(working_folder)
            output_file = output_folder / file_c1_c6.name

            # Copy the file to outputs
            shutil.copy2(file_c1_c6, output_file)
            print(f"✓ Copied C1-C6 file to outputs: {output_file}")
            print(f"  This ensures the download file does not include 3-Year Summary changes")
        except Exception as e:
            print(f"Warning: Failed to copy file to outputs before 3YearSummary: {e}")
            # Continue anyway - 3YearSummary will still run

        # Run NBD_MF_20_3YearSummary.py to generate 3-Year Summary
        print(f"\n{'='*80}")
        print(f"RUNNING 3-YEAR SUMMARY SCRIPT")
        print(f"{'='*80}")

        try:
            # Get the path to NBD_MF_20_3YearSummary.py
            script_dir = Path(__file__).resolve().parent
            summary_script = script_dir / "NBD_MF_20_3YearSummary.py"

            # Run the 3-Year Summary script
            print(f"Running 3-Year Summary script: {summary_script}")
            result = subprocess.run(
                [sys.executable, str(summary_script)],
                capture_output=True,
                text=True,
                cwd=str(script_dir)
            )

            # Log output
            if result.stdout:
                print(f"3-Year Summary script output:\n{result.stdout}")

            if result.stderr:
                print(f"3-Year Summary script stderr:\n{result.stderr}")

            if result.returncode == 0:
                print(f"✓ 3-Year Summary script completed successfully")
            else:
                print(f"✗ 3-Year Summary script failed with return code {result.returncode}")

        except Exception as e:
            print(f"Error running 3-Year Summary script: {e}")
            import traceback
            traceback.print_exc()
        
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
    parser = argparse.ArgumentParser(description='NBD-AF-20-C2 Report Automation')
    parser.add_argument('--working-dir', type=str, help='Working directory (not used, for compatibility)')
    parser.add_argument('--month', type=str, help='Report month (e.g., Jul)')
    parser.add_argument('--year', type=str, help='Report year (e.g., 2025)')
    args = parser.parse_args()

    main()
