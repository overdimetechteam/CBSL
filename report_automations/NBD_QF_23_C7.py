import sys
import argparse
from pathlib import Path
from datetime import datetime

# Add the parent directory of the project to Python path
project_root = Path(__file__).resolve().parent.parent
sys.path.append(str(project_root))

# Commented out until full release - uncomment when needed
# from unwindOfDiscount import UnwindOfDiscount
# from report_automations.closedSheet import update_closed_sheet
# from report_automations.newActivations import update_new_activations_sheet
# from report_automations.grantAmounts import update_grant_amounts_sheet
from report_automations.wo import update_wo_sheet

def calculate_quarter_from_date(date_str):
    """
    Calculate quarter and financial year from a date string.
    Financial Year: April - March
    Q1 (Apr-Jun), Q2 (Jul-Sep), Q3 (Oct-Dec), Q4 (Jan-Mar)

    Args:
        date_str: Date in format YYYY-MM-DD or DD-MM-YYYY

    Returns:
        tuple: (quarter, financial_year)
    """
    # Parse date
    if '-' in date_str:
        parts = date_str.split('-')
        if len(parts[0]) == 4:  # YYYY-MM-DD
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        else:  # DD-MM-YYYY
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
    else:
        raise ValueError(f"Invalid date format: {date_str}")

    month = date_obj.month
    year = date_obj.year

    # Determine quarter and financial year
    # Financial year starts in April
    if month >= 4:  # Apr-Dec of same year
        fy_year = year
        if month in [4, 5, 6]:
            quarter = 1  # Q1: Apr-Jun
        elif month in [7, 8, 9]:
            quarter = 2  # Q2: Jul-Sep
        else:  # month in [10, 11, 12]
            quarter = 3  # Q3: Oct-Dec
    else:  # Jan-Mar of next year
        fy_year = year - 1
        quarter = 4  # Q4: Jan-Mar

    return quarter, fy_year

def get_quarter_input():
    """Get quarter selection from user."""
    print("\n" + "="*60)
    print("CBSL Quarterly Reporting Automation")
    print("="*60)
    print("\nFinancial Year: April - March")
    print("\nAvailable Quarters:")
    print("  1. Q1 (Apr-Jun)")
    print("  2. Q2 (Jul-Sep)")
    print("  3. Q3 (Oct-Dec)")
    print("  4. Q4 (Jan-Mar)")
    
    while True:
        try:
            choice = input("\nSelect quarter (1-4): ").strip()
            if choice in ['1', '2', '3', '4']:
                return int(choice)
            else:
                print("Invalid input. Please enter 1, 2, 3, or 4.")
        except KeyboardInterrupt:
            print("\n\nProcess cancelled by user.")
            sys.exit(0)
        except Exception:
            print("Invalid input. Please enter a number between 1 and 4.")

def get_financial_year():
    """Get financial year from user."""
    print("\nFinancial Year (format: YYYY for year starting in April)")
    print("Example: Enter 2025 for FY 2025-26 (Apr 2025 - Mar 2026)")
    
    while True:
        try:
            year = input("Enter financial year: ").strip()
            if len(year) == 4 and year.isdigit():
                fy_year = int(year)
                if 2020 <= fy_year <= 2050:
                    return fy_year
                else:
                    print("Please enter a year between 2020 and 2050.")
            else:
                print("Invalid format. Please enter a 4-digit year (e.g., 2025).")
        except KeyboardInterrupt:
            print("\n\nProcess cancelled by user.")
            sys.exit(0)
        except Exception:
            print("Invalid input. Please enter a valid year.")

def get_quarter_months(quarter: int) -> tuple:
    """Get quarter start and end months."""
    quarter_months_map = {
        1: ("Apr", "Jun"),
        2: ("Jul", "Sep"),
        3: ("Oct", "Dec"),
        4: ("Jan", "Mar")
    }
    return quarter_months_map[quarter]

def find_dynamic_folder(base_path: Path) -> Path:
    """
    Locate the latest date folder inside the base path.
    Similar to NBD_MF_23_C1C2.py's find_dynamic_folder function.

    Example:
        base_path/
            30-09-2025
            31-10-2025
            31-1-2025(1)  <- latest (by modification time)

    Returns:
        Path to the latest date folder, or None if not found
    """
    if not base_path.exists():
        print(f"ERROR: Working folder not found at {base_path}")
        return None

    # List subfolders (date folders)
    subfolders = [f for f in base_path.iterdir() if f.is_dir()]

    if not subfolders:
        print(f"ERROR: No date folders found under {base_path}")
        return None

    # Sort by last modified date, newest first
    subfolders.sort(key=lambda f: f.stat().st_mtime, reverse=True)

    latest_folder = subfolders[0]
    print(f"Found latest working folder: {latest_folder.name}")
    return latest_folder

def process_c7_stage_movement(quarter: int, fy_year: int, working_dir: Path = None) -> Path:
    """Process C7 Stage Movement report. Returns the output path."""
    print("\n" + "="*60)
    print("Processing C7 Stage Movement Report")
    print("="*60)

    # Setup paths - use provided working_dir or default
    if working_dir is None:
        # Interactive mode: find the latest date folder
        script_dir = Path(__file__).resolve().parent
        project_root = script_dir.parent
        base_working_dir = project_root / "working" / "NBD_QF_23_C7"
        working_dir = find_dynamic_folder(base_working_dir)
        if not working_dir:
            return None
    else:
        # GUI mode: working_dir is already the date folder from app.py
        working_dir = Path(working_dir)

    # Get quarter months
    start_month, end_month = get_quarter_months(quarter)

    # Search for C7 Stage Movement file in working directory
    print(f"\nSearching for C7 Stage Movement file in: {working_dir}")

    # Pattern to match: any file starting with "C7" and ending with ".xlsb"
    # This will match files like "1. C7- Report for Stage - Movement.xlsb" or "C7- Report for Stage - Movement.xlsb"
    pattern = "*C7*.xlsb"
    c7_files = list(working_dir.glob(pattern))

    if not c7_files:
        print(f"✗ C7 Stage Movement file not found in {working_dir}")
        print(f"   Expected pattern: {pattern}")
        return None

    # Use the first match
    source_file = c7_files[0]
    print(f"✓ Found: {source_file.name}")

    # Generate new filename: C7- Report for Stage - Movement Q{quarter} FY{fy_year}.xlsb
    new_filename = f"C7- Report for Stage - Movement Q{quarter} FY{fy_year}.xlsb"
    output_path = working_dir / new_filename

    # Rename the file in place
    print(f"\nRenaming file...")
    print(f"  From: {source_file.name}")
    print(f"  To: {new_filename}")

    try:
        if source_file != output_path:
            source_file.rename(output_path)
        print(f"✓ File renamed successfully to: {output_path.name}")
    except Exception as e:
        print(f"✗ Failed to rename file: {e}")
        return None

    print("\n" + "="*60)
    print("C7 STAGE MOVEMENT REPORT SUMMARY")
    print("="*60)
    print(f"Quarter: Q{quarter} ({start_month}-{end_month})")
    print(f"Financial Year: FY{fy_year}-{str(fy_year+1)[-2:]}")
    print(f"Output: {output_path}")
    print("="*60 + "\n")

    return output_path

def find_closure_report(working_dir: Path) -> Path:
    """Find the most recent Closure Category Wise Report in working directory."""
    pattern = "Closure Category Wise Report - Internal*.xlsx"
    closure_files = list(working_dir.glob(pattern))
    
    if not closure_files:
        print(f"✗ Closure report not found in {working_dir}")
        print(f"   Expected pattern: {pattern}")
        return None
    
    # Sort by modification time and get the most recent
    closure_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return closure_files[0]

def process_closed_sheet_update(quarter: int, fy_year: int, c7_output_path: Path, working_dir: Path = None):
    """Process Closed Sheet update in C7 report."""
    print("\n" + "="*60)
    print("Processing Closed Sheet Update")
    print("="*60)

    # Setup paths - use provided working_dir or default
    if working_dir is None:
        # Interactive mode: find the latest date folder
        script_dir = Path(__file__).resolve().parent
        project_root = script_dir.parent
        base_working_dir = project_root / "working" / "NBD_QF_23_C7"
        working_dir = find_dynamic_folder(base_working_dir)
        if not working_dir:
            return False
    else:
        # GUI mode: working_dir is already the date folder from app.py
        working_dir = Path(working_dir)

    # Get quarter months for the Month column
    start_month, end_month = get_quarter_months(quarter)
    quarter_months = f"{start_month}-{end_month}"

    # Find closure report
    print(f"\nSearching for Closure Category Wise Report in: {working_dir}")
    closure_report_path = find_closure_report(working_dir)

    if not closure_report_path:
        return False

    print(f"✓ Found: {closure_report_path.name}")

    # Update closed sheet
    # TODO: Uncomment when full release is done
    # success = update_closed_sheet(
    #     c7_report_path=c7_output_path,
    #     closure_report_path=closure_report_path,
    #     quarter_months=quarter_months
    # )
    # return success

    print("⚠ Closed sheet update not yet implemented")
    return True  # Return True for now to allow workflow to continue

def find_info_request_report(working_dir: Path) -> Path:
    """Find the most recent Information Request from Credit report in working directory."""
    pattern = "Information Request from Credit*.xlsx"
    info_files = list(working_dir.glob(pattern))
    
    if not info_files:
        print(f"✗ Information Request report not found in {working_dir}")
        print(f"   Expected pattern: {pattern}")
        return None
    
    # Sort by modification time and get the most recent
    info_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return info_files[0]

def process_new_activations_update(quarter: int, fy_year: int, c7_output_path: Path, working_dir: Path = None):
    """Process New Activations sheet update in C7 report."""
    print("\n" + "="*60)
    print("Processing New Activations Sheet Update")
    print("="*60)

    # Setup paths - use provided working_dir or default
    if working_dir is None:
        # Interactive mode: find the latest date folder
        script_dir = Path(__file__).resolve().parent
        project_root = script_dir.parent
        base_working_dir = project_root / "working" / "NBD_QF_23_C7"
        working_dir = find_dynamic_folder(base_working_dir)
        if not working_dir:
            return False
    else:
        # GUI mode: working_dir is already the date folder from app.py
        working_dir = Path(working_dir)

    # Get quarter months for the Month column
    start_month, end_month = get_quarter_months(quarter)
    quarter_months = f"{start_month}-{end_month}"

    # Find Information Request report
    print(f"\nSearching for Information Request from Credit report in: {working_dir}")
    info_request_path = find_info_request_report(working_dir)

    if not info_request_path:
        return False

    print(f"✓ Found: {info_request_path.name}")

    # Update New Activations sheet
    # TODO: Uncomment when full release is done
    # success = update_new_activations_sheet(
    #     c7_report_path=c7_output_path,
    #     info_request_path=info_request_path,
    #     quarter_months=quarter_months
    # )
    # return success

    print("⚠ New Activations sheet update not yet implemented")
    return True  # Return True for now to allow workflow to continue

def process_grant_amounts_update(quarter: int, fy_year: int, c7_output_path: Path, working_dir: Path = None):
    """Process Grant Amounts sheet update in C7 report."""
    print("\n" + "="*60)
    print("Processing Grant Amounts Sheet Update")
    print("="*60)

    # Setup paths - use provided working_dir or default
    if working_dir is None:
        # Interactive mode: find the latest date folder
        script_dir = Path(__file__).resolve().parent
        project_root = script_dir.parent
        base_working_dir = project_root / "working" / "NBD_QF_23_C7"
        working_dir = find_dynamic_folder(base_working_dir)
        if not working_dir:
            return False
    else:
        # GUI mode: working_dir is already the date folder from app.py
        working_dir = Path(working_dir)

    # Find Information Request report
    print(f"\nSearching for Information Request from Credit report in: {working_dir}")
    info_request_path = find_info_request_report(working_dir)

    if not info_request_path:
        return False

    print(f"✓ Found: {info_request_path.name}")

    # Update Grant Amounts sheet with quarter and fy_year for Stage lookup
    # TODO: Uncomment when full release is done
    # success = update_grant_amounts_sheet(
    #     c7_report_path=c7_output_path,
    #     info_request_path=info_request_path,
    #     quarter=quarter,
    #     fy_year=fy_year
    # )
    # return success

    print("⚠ Grant Amounts sheet update not yet implemented")
    return True  # Return True for now to allow workflow to continue

def process_wo_update(quarter: int, fy_year: int, c7_output_path: Path):
    """Process WO (Write Off) sheet update in C7 report."""
    print("\n" + "="*60)
    print("Processing WO Sheet Update")
    print("="*60)
    
    # Update WO sheet
    success = update_wo_sheet(
        c7_report_path=c7_output_path,
        quarter=quarter,
        fy_year=fy_year
    )
    
    return success

def main():
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='NBD-QF-23-C7: C7 Stage Movement Report')
    parser.add_argument('--quarter', type=int, choices=[1, 2, 3, 4],
                       help='Quarter (1-4): Q1(Apr-Jun), Q2(Jul-Sep), Q3(Oct-Dec), Q4(Jan-Mar)')
    parser.add_argument('--financial-year', type=int,
                       help='Financial year (e.g., 2025 for FY 2025-26)')
    parser.add_argument('--gui-date', type=str,
                       help='Date from GUI (YYYY-MM-DD or DD-MM-YYYY) - will calculate quarter and FY automatically')
    parser.add_argument('--working-dir', type=str,
                       help='Working directory path')

    args = parser.parse_args()

    # Determine quarter and financial year
    if args.gui_date:
        quarter, fy_year = calculate_quarter_from_date(args.gui_date)
        print(f"Calculated from date {args.gui_date}: Quarter {quarter}, Financial Year {fy_year}")
    elif args.quarter and args.financial_year:
        quarter = args.quarter
        fy_year = args.financial_year
    else:
        # Interactive mode (backward compatibility)
        quarter = get_quarter_input()
        fy_year = get_financial_year()

    # Determine working directory
    if args.working_dir:
        working_dir = Path(args.working_dir)
        print(f"Using working directory: {working_dir}")
    else:
        working_dir = None  # Will use default

    print("\n" + "="*60)
    print("Processing C7 Stage Movement Report...")
    print("="*60)

    # Step 1: Process C7 Stage Movement (rename in place) - returns the output path
    c7_output_path = process_c7_stage_movement(quarter=quarter, fy_year=fy_year, working_dir=working_dir)

    if not c7_output_path or not c7_output_path.exists():
        print("✗ Failed to process C7 output file. Exiting...")
        sys.exit(1)

    print(f"\n📁 Working with C7 file: {c7_output_path.name}")

    # Step 2: Process Unwind of Discount (full automation)
    #!!!UnComment this when full release is done
    # processor = UnwindOfDiscount(quarter=quarter, financial_year=fy_year)
    # processor.process()

    #!!!UnComment this when full release is done
    # Step 3: Update Closed Sheet in C7 report (works directly with .xlsb)
    # print(f"\n🔄 Updating same file: {c7_output_path.name}")
    # process_closed_sheet_update(quarter=quarter, fy_year=fy_year, c7_output_path=c7_output_path, working_dir=working_dir)

    #!!!UnComment this when full release is done
    # Step 4: Update New Activations Sheet in C7 report (same file)
    # print(f"\n🔄 Updating same file: {c7_output_path.name}")
    # process_new_activations_update(quarter=quarter, fy_year=fy_year, c7_output_path=c7_output_path, working_dir=working_dir)

    #!!!UnComment this when full release is done
    # Step 5: Update Grant Amounts Sheet in C7 report (same file)
    # print(f"\n🔄 Updating same file: {c7_output_path.name}")
    # process_grant_amounts_update(quarter=quarter, fy_year=fy_year, c7_output_path=c7_output_path, working_dir=working_dir)

    # Step 6: Update WO (Write Off) Sheet in C7 report (same file)
    print(f"\n🔄 Updating same file: {c7_output_path.name}")
    process_wo_update(quarter=quarter, fy_year=fy_year, c7_output_path=c7_output_path)

    print("\n" + "="*60)
    print("✓ C7 processing completed successfully!")
    print(f"📄 Final output: {c7_output_path}")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
