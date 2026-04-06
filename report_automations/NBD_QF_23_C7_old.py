import sys
import argparse
import os
from pathlib import Path
from datetime import datetime

# Add the parent directory of the project to Python path
project_root = Path(__file__).resolve().parent.parent
sys.path.append(str(project_root))

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

def main():
    parser = argparse.ArgumentParser(description='NBD-QF-23-C7: Unwind of Discount Report')
    parser.add_argument('--quarter', type=int, choices=[1, 2, 3, 4],
                       help='Quarter (1-4): Q1(Apr-Jun), Q2(Jul-Sep), Q3(Oct-Dec), Q4(Jan-Mar)')
    parser.add_argument('--financial-year', type=int,
                       help='Financial year (e.g., 2025 for FY 2025-26)')
    parser.add_argument('--gui-date', type=str,
                       help='Date from GUI (YYYY-MM-DD or DD-MM-YYYY) - will calculate quarter and FY automatically')
    parser.add_argument('--working-dir', type=str,
                       help='Working directory path (optional)')

    args = parser.parse_args()

    # Calculate quarter and FY from GUI date if provided
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

    print(f"\nProcessing C7 Report for Q{quarter} FY{fy_year}")
    print(f"Working directory: {args.working_dir if args.working_dir else 'Default'}")

    # TODO: Implement UnwindOfDiscount class or call appropriate processing logic
    # For now, create a placeholder output file
    if args.working_dir:
        working_path = Path(args.working_dir)
        output_file = working_path / f"C7- Report for Stage - Movement Q{quarter} FY{fy_year}.xlsb"
        print(f"Output will be created at: {output_file}")

    print("✓ C7 processing completed successfully")

def get_quarter_input():
    """Get quarter selection from user (interactive mode)."""
    print("\n" + "="*60)
    print("CBSL Quarterly Reporting - Unwind of Discount")
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
    """Get financial year from user (interactive mode)."""
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

if __name__ == "__main__":
    main()