import os
from pathlib import Path
import openpyxl
import sys
import xlwings as xw

def get_month_year_from_filename(filename):
    # Example: "NBD-MF-20-C1 to C6 July 2025.xlsx"
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


def main():

    base_dir = Path(__file__).resolve().parent.parent
    report_name = "NBD_MF_20_C1_C6"  # Example report

    # Get the latest working folder (single date folder)
    working_folder = find_latest_working_folder(base_dir, report_name)
    if not working_folder:
        return

    # Find files inside the working folder
    file_c1_c6 = find_file(working_folder, "**/NBD-MF-20-C1 to C6*.xlsx", "C1-C6 file")
    file_afl = find_file(working_folder, "**/NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx", "AFL file")

    if not file_c1_c6 or not file_afl:
        print("Required files not found. Cannot continue.")
        return

    print("\n--- Summary ---")
    print(f"C1-C6 File: {file_c1_c6}")
    print(f"AFL File: {file_afl}")
    print(f"Working folder: {working_folder}")
    
    try:
        # Use xlwings to read values (calculates formulas through Excel)
        # and openpyxl to write (faster for writing)
        print(f"\nLoading workbooks with xlwings (to calculate formulas)...")

        app = xw.App(visible=False)

        try:
            # Open SOFP source file with xlwings to get calculated values
            wb_afl_xw = app.books.open(str(file_afl))

            # Check if SOFP sheet exists
            if "NBD-MF-01-SOFP" not in [sheet.name for sheet in wb_afl_xw.sheets]:
                print(f"Required sheet 'NBD-MF-01-SOFP' not found in AFL file")
                print(f"Available sheets: {[s.name for s in wb_afl_xw.sheets]}")
                wb_afl_xw.close()
                return

            src_ws_xw = wb_afl_xw.sheets["NBD-MF-01-SOFP"]

            print(f"\n=== STEP 1: Reading VALUES from SOFP file ===")
            print(f"Note: Using xlwings to calculate formulas through Excel")
            print(f"      - Formulas → Evaluated/calculated values")
            print(f"      - Direct values → Direct values")

            # Read the date from D2
            date_value = src_ws_xw.range("D2").value
            print(f"✓ Read date from D2: {date_value}")

            # Read all values from D3:E137 as a 2D array
            values = src_ws_xw.range("D3:E137").value
            print(f"✓ Read {len(values)} rows × 2 columns = {len(values) * 2} cells")

            # Close SOFP file
            wb_afl_xw.close()

        finally:
            app.quit()

        # Now use openpyxl to write the values to C1-C6 file
        print(f"\n=== STEP 2: Writing values to C1-C6 file ===")
        print(f"Loading C1-C6 file with openpyxl...")
        wb_c1_c6 = openpyxl.load_workbook(file_c1_c6, data_only=False)

        # Check if SOFP sheet exists
        if "NBD-MF-01-SOFP" not in wb_c1_c6.sheetnames:
            print(f"Required sheet 'NBD-MF-01-SOFP' not found in C1-C6 file")
            print(f"Available sheets: {wb_c1_c6.sheetnames}")
            return

        tgt_ws = wb_c1_c6["NBD-MF-01-SOFP"]

        # Write the date to D2
        print(f"Writing date to D2...")
        tgt_ws["D2"].value = date_value
        print(f"✓ Written date to D2: {date_value}")

        print(f"\nClearing range D3:E137 in C1-C6 SOFP sheet...")
        # Clear the target range first
        for row in range(3, 138):
            tgt_ws[f"D{row}"].value = None
            tgt_ws[f"E{row}"].value = None
        print(f"✓ Cleared 270 cells")

        print(f"\nPasting values to D3:E137...")
        # Write the values
        cells_copied = 0
        for i, row_values in enumerate(values):
            excel_row = i + 3  # Start at row 3

            # row_values is [D_value, E_value]
            if row_values is not None:
                if isinstance(row_values, list) and len(row_values) >= 2:
                    tgt_ws[f"D{excel_row}"].value = row_values[0]
                    tgt_ws[f"E{excel_row}"].value = row_values[1]
                    cells_copied += 2
                else:
                    # Single value (shouldn't happen for a 2-column range)
                    tgt_ws[f"D{excel_row}"].value = row_values
                    cells_copied += 1

            # Log progress every 20 rows
            if excel_row % 20 == 0:
                print(f"  Written rows 3-{excel_row} ({cells_copied} cells so far)...")

        print(f"\n=== Copy Summary ===")
        print(f"✓ Successfully copied date from D2")
        print(f"✓ Successfully copied {cells_copied} cells (values only)")
        print(f"✓ Range: D2 and D3:E137 from SOFP file to C1-C6 file")

        # Save updated workbook
        wb_c1_c6.save(file_c1_c6)
        print(f"\n✓ Report saved successfully to: {file_c1_c6}")
        
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
    parser = argparse.ArgumentParser(description='NBD-MF-01-SOFP Report Automation')
    parser.add_argument('--working-dir', type=str, help='Working directory (not used, for compatibility)')
    parser.add_argument('--month', type=str, help='Report month (e.g., Jul)')
    parser.add_argument('--year', type=str, help='Report year (e.g., 2025)')
    args = parser.parse_args()

    main()
