#!/usr/bin/env python3
"""
Master runner script for all CBSL reporting automations.

This script loads the main workbook once and passes it through all report processing
scripts in sequence, then saves the final result. This approach is much more efficient
than loading and saving the workbook separately for each script.

Usage:
    python run_all_reports.py

Output:
    outputs/monthly/final_C1_to_C6.xlsx - Contains all modifications from scripts C2 through C6
"""

import openpyxl
from pathlib import Path
import sys
import traceback

# Import all report processing functions
from NBD_MF_20_C2 import main as run_c2
from NBD_MF_20_C3 import main as run_c3
from NBD_MF_20_C4 import main as run_c4
from NBD_MF_20_C5 import main as run_c5
from NBD_MF_20_C6 import main as run_c6

def find_files_safely(base_dir, pattern, description):
    """Safely find files and provide detailed error messages"""
    print(f"Looking for {description} in: {base_dir}")
    print(f"Search pattern: {pattern}")
    
    files = list(base_dir.glob(pattern))
    print(f"Found {len(files)} files matching pattern:")
    
    if not files:
        print(f"[ERROR] No files found matching pattern: {pattern}")
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
        print(f"  [OK] {file.name}")
    
    return files[0]  # Return the first match

def find_first_matching(search_dirs, pattern):
    """Find first matching file in search directories"""
    for search_dir in search_dirs:
        if search_dir.exists():
            files = list(search_dir.glob(pattern))
            if files:
                return files[0]
    return None

def get_month_year_from_filename(filename):
    """Extract month and year from filename"""
    parts = filename.split()
    for i, part in enumerate(parts):
        if part in {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","January","February","March","April","May","June","July","August","September","October","November","December"}:
            month = part
            year = parts[i+1].replace(".xlsx","").replace(".xlsb","")
            return month, year
    return None, None

def main():
    """Main function to run all report processing scripts in sequence"""
    print("Starting CBSL Reporting Automation - Master Runner")
    print("=" * 60)
    
    # Get the base directory (parent of report_automations)
    base_dir = Path(__file__).resolve().parent.parent
    working_dir = base_dir / "working"
    monthly_working_dir = working_dir / "monthly"
    outputs_monthly_dir = base_dir / "outputs" / "monthly"
    
    print(f"Base directory: {base_dir}")
    print(f"Working directory: {working_dir}")
    print(f"Monthly working directory: {monthly_working_dir}")
    
    # Create outputs monthly directory if it doesn't exist
    outputs_monthly_dir.mkdir(parents=True, exist_ok=True)
    
    # Search directories
    search_dirs = [monthly_working_dir, working_dir, base_dir]
    
    # Locate required files
    print(f"\n[INFO] Locating required files...")
    
    # Main C1-C6 file (required)
    file_c1_c6 = find_first_matching(search_dirs, "**/NBD-MF-20-C1 to C6*.xlsx")
    if not file_c1_c6:
        print("[ERROR] Could not find C1 to C6 file")
        return
    
    # AFL file (required for C2 and C6)
    file_afl = find_first_matching(search_dirs, "**/NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx")
    if not file_afl:
        print("[ERROR] Could not find AFL file")
        return
    
    # Additional files for C3
    file_car = find_first_matching(search_dirs, "**/CAR Working*.xlsb")
    file_prod = find_first_matching(search_dirs, "**/Prod. wise Class. of Loans*.xlsb")
    file_cbsl = find_first_matching(search_dirs, "**/CBSL Provision Comparison*.xlsb")
    file_sofp = find_first_matching(search_dirs, "**/NBD-MF-01-SOFP*.xlsx")
    
    # Unutilized file for C4
    file_unutilized = find_first_matching(search_dirs, "**/Unutilized Credit Limits*.xlsx")
    
    print(f"[OK] C1-C6: {file_c1_c6}")
    print(f"[OK] AFL: {file_afl}")
    if file_car:
        print(f"[OK] CAR Working: {file_car}")
    if file_prod:
        print(f"[OK] Prod wise: {file_prod}")
    if file_cbsl:
        print(f"[OK] CBSL Provision: {file_cbsl}")
    if file_sofp:
        print(f"[OK] SOFP: {file_sofp}")
    if file_unutilized:
        print(f"[OK] Unutilized: {file_unutilized}")
    
    # Extract month and year from filename
    month, year = get_month_year_from_filename(file_c1_c6.name)
    if not month or not year:
        print(f"[ERROR] Could not parse month/year from filename: {file_c1_c6.name}")
        return
    
    print(f"\n[OK] Parsed month: {month}, year: {year}")
    
    # Determine output folder
    working_folder_name = None
    for search_dir in search_dirs:
        if search_dir.exists():
            for subfolder in search_dir.iterdir():
                if subfolder.is_dir():
                    c1 = list(subfolder.glob("**/NBD-MF-20-C1 to C6*.xlsx"))
                    if c1:
                        working_folder_name = subfolder.name
                        break
            if working_folder_name:
                break
    
    if working_folder_name:
        out_folder = outputs_monthly_dir / working_folder_name
    else:
        out_folder = outputs_monthly_dir / f"{year}_{month}"
    
    out_folder.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {out_folder}")
    
    try:
        # Load main workbook once
        print(f"\n[LOAD] Loading main workbook...")
        wb_c1_c6 = openpyxl.load_workbook(file_c1_c6, data_only=False)
        print(f"[OK] Main workbook loaded successfully")
        
        # Load AFL workbook
        print(f"[LOAD] Loading AFL workbook...")
        wb_afl = openpyxl.load_workbook(file_afl, data_only=True)
        print(f"[OK] AFL workbook loaded successfully")
        
        # Process all reports in sequence
        print(f"\n[PROCESS] Processing all reports in sequence...")
        print("=" * 60)
        
        # C2 Report
        print(f"\n[PROCESS] Processing C2 Report...")
        try:
            wb_c1_c6 = run_c2(wb_c1_c6, wb_afl)
            if wb_c1_c6:
                print(f"[OK] C2 Report completed successfully")
            else:
                print(f"[ERROR] C2 Report failed")
                return
        except Exception as e:
            print(f"[ERROR] C2 Report failed with error: {e}")
            traceback.print_exc()
            return
        
        # C3 Report
        print(f"\n[PROCESS] Processing C3 Report...")
        try:
            if all([file_car, file_prod, file_cbsl]):
                wb_c1_c6 = run_c3(wb_c1_c6, file_car, file_prod, file_cbsl, file_sofp, out_folder)
                if wb_c1_c6:
                    print(f"[OK] C3 Report completed successfully")
                else:
                    print(f"[ERROR] C3 Report failed")
                    return
            else:
                print(f"[WARN] Skipping C3 Report - missing required files")
        except Exception as e:
            print(f"[ERROR] C3 Report failed with error: {e}")
            traceback.print_exc()
            return
        
        # C4 Report
        print(f"\n[PROCESS] Processing C4 Report...")
        try:
            if file_unutilized:
                wb_c1_c6 = run_c4(wb_c1_c6, file_unutilized)
                if wb_c1_c6:
                    print(f"[OK] C4 Report completed successfully")
                else:
                    print(f"[ERROR] C4 Report failed")
                    return
            else:
                print(f"[WARN] Skipping C4 Report - missing unutilized file")
        except Exception as e:
            print(f"[ERROR] C4 Report failed with error: {e}")
            traceback.print_exc()
            return
        
        # C5 Report
        print(f"\n[PROCESS] Processing C5 Report...")
        try:
            wb_c1_c6 = run_c5(wb_c1_c6)
            if wb_c1_c6:
                print(f"[OK] C5 Report completed successfully")
            else:
                print(f"[ERROR] C5 Report failed")
                return
        except Exception as e:
            print(f"[ERROR] C5 Report failed with error: {e}")
            traceback.print_exc()
            return
        
        # C6 Report
        print(f"\n[PROCESS] Processing C6 Report...")
        try:
            wb_c1_c6 = run_c6(wb_c1_c6, wb_afl)
            if wb_c1_c6:
                print(f"[OK] C6 Report completed successfully")
            else:
                print(f"[ERROR] C6 Report failed")
                return
        except Exception as e:
            print(f"[ERROR] C6 Report failed with error: {e}")
            traceback.print_exc()
            return
        
        # Save final output
        print(f"\n[SAVE] Saving final combined file...")
        # Use input file name as base for output file name
        input_filename = file_c1_c6.stem  # Get filename without extension
        output_file = out_folder / f"{input_filename}.xlsx"
        wb_c1_c6.save(output_file)
        
        print(f"\n[SUCCESS] All reports completed successfully!")
        print(f"[SAVE] Final combined file saved to: {output_file}")
        print(f"[PROCESS] Contains all modifications from scripts C2 through C6")
        
    except FileNotFoundError as e:
        print(f"[ERROR] File not found: {e}")
    except PermissionError as e:
        print(f"[ERROR] Permission error (file might be open in Excel): {e}")
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
