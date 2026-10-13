import os
import sys
from pathlib import Path
import re
import shutil
import csv
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import logging

# Add parent directory to path for imports
script_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(script_dir)
sys.path.insert(0, parent_dir)

# Import the download bots
from report_automations.tb_from_erp_bot import run_tb_download
from bots.DM_Scienter_Daily_Deposit_Bal import run_dm_daily_deposit_download

# ------------------------------------------------------------
# STEP 0 – Download Required Files
# ------------------------------------------------------------
def download_required_files(input_date, working_folder, max_retries=3):
    """
    Download all required files for NBD-WF-18-DM report in parallel threads
    with retry logic for failed downloads.

    Args:
        input_date (datetime): Date selected by user
        working_folder (Path): Path to working folder date directory
        max_retries (int): Maximum number of retry attempts for each failed download

    Returns:
        bool: True if all downloads successful, False otherwise
    """
    import threading
    import time

    print("=" * 60)
    print("STEP 0: Downloading Required Files")
    print("=" * 60)
    print(f"Input Date: {input_date.strftime('%Y-%m-%d')}")
    print(f"Working Folder: {working_folder}")
    print(f"Max Retries: {max_retries}")
    print("\nRunning downloads in 2 parallel threads:")
    print("  Thread 1: TB_DetailReport.xlsx from ERP")
    print("  Thread 2: Scienter CSV Reports (sequentially)")
    print("=" * 60)

    # Ensure working folder exists
    os.makedirs(working_folder, exist_ok=True)

    # Track success of each download
    tb_success = False
    scienter_success = False
    tb_error = None
    scienter_error = None

    # ========================================
    # Thread 1: Download TB_DetailReport.xlsx
    # ========================================
    def download_tb():
        nonlocal tb_success, tb_error
        print("\n[Thread 1] Starting TB_DetailReport.xlsx download from ERP...")
        try:
            result = run_tb_download(input_date, str(working_folder))
            if result:
                print("[Thread 1] ✓ TB_DetailReport.xlsx downloaded successfully")
                tb_success = True
            else:
                print("[Thread 1] ✗ Failed to download TB_DetailReport.xlsx")
                tb_error = "Download returned False"
        except Exception as e:
            print(f"[Thread 1] ✗ Error downloading TB_DetailReport.xlsx: {e}")
            tb_error = str(e)

    # ========================================
    # Thread 2: Download Scienter CSV Reports (Sequential)
    # ========================================
    def download_scienter():
        nonlocal scienter_success, scienter_error
        print("\n[Thread 2] Starting Scienter CSV Reports download...")
        print("[Thread 2] Report 1/4: CentralBankMovementOfDailyDepositBalances.csv")
        print("[Thread 2] Report 2/4: MovementOfDepositDetails.csv")
        print("[Thread 2] Report 3/4: FDBaseNew_UserBranch.csv")
        print("[Thread 2] Report 4/4: InterestPayableAsAtDateReport.csv")
        print("[Thread 2] Note: All reports downloaded sequentially from same webpage")
        try:
            result = run_dm_daily_deposit_download(input_date, str(working_folder))
            if result:
                print("[Thread 2] ✓ All four Scienter CSV reports downloaded successfully")
                scienter_success = True
            else:
                print("[Thread 2] ✗ Failed to download Scienter CSV reports")
                scienter_error = "Download returned False"
        except Exception as e:
            print(f"[Thread 2] ✗ Error downloading Scienter CSV reports: {e}")
            scienter_error = str(e)

    # Start both downloads in parallel
    tb_thread = threading.Thread(target=download_tb, name="TB-Download")
    scienter_thread = threading.Thread(target=download_scienter, name="Scienter-Download")

    tb_thread.start()
    scienter_thread.start()

    # Wait for both threads to complete
    print("\nWaiting for downloads to complete...")
    tb_thread.join()
    scienter_thread.join()

    # ========================================
    # Retry Logic for Failed Downloads
    # ========================================
    retry_count = 0
    while (not tb_success or not scienter_success) and retry_count < max_retries:
        retry_count += 1
        print("\n" + "=" * 60)
        print(f"RETRY ATTEMPT {retry_count}/{max_retries}")
        print("=" * 60)

        retry_threads = []

        # Retry TB download if failed
        if not tb_success:
            print(f"\n[Retry {retry_count}] Retrying TB_DetailReport.xlsx download...")
            tb_error = None  # Reset error
            tb_retry_thread = threading.Thread(target=download_tb, name=f"TB-Retry-{retry_count}")
            retry_threads.append(tb_retry_thread)
            tb_retry_thread.start()

        # Retry Scienter download if failed
        if not scienter_success:
            print(f"\n[Retry {retry_count}] Retrying Scienter CSV Reports download...")
            scienter_error = None  # Reset error
            scienter_retry_thread = threading.Thread(target=download_scienter, name=f"Scienter-Retry-{retry_count}")
            retry_threads.append(scienter_retry_thread)
            scienter_retry_thread.start()

        # Wait for retry threads to complete
        for thread in retry_threads:
            thread.join()

        # Brief pause before next retry if still failing
        if (not tb_success or not scienter_success) and retry_count < max_retries:
            print(f"\nWaiting 5 seconds before next retry...")
            time.sleep(5)

    # ========================================
    # Final Summary
    # ========================================
    print("\n" + "=" * 60)
    print("Download Summary (Final):")
    print(f"  TB Report: {'✓ Success' if tb_success else '✗ Failed'}")
    print(f"  Scienter CSVs: {'✓ Success' if scienter_success else '✗ Failed'}")
    if retry_count > 0:
        print(f"  Total Retry Attempts: {retry_count}")

    downloads_successful = tb_success and scienter_success

    if downloads_successful:
        print("\n✓ All required files downloaded successfully")
    else:
        print("\n⚠ Some downloads failed after all retries:")
        if not tb_success:
            print(f"  - TB Report: {tb_error or 'Unknown error'}")
        if not scienter_success:
            print(f"  - Scienter CSVs: {scienter_error or 'Unknown error'}")
    print("=" * 60)

    return downloads_successful

# ------------------------------------------------------------
# STEP 1 – Find and rename NBD-WF-18-DM Deposit Liability file
# ------------------------------------------------------------
def rename_wf18dm_report(working_dir, report_name, old_week, new_week):
    """Find and rename the NBD-WF-18-DM file in the working date folder"""
    working_path = Path(working_dir)

    # Search for the template file directly in the working folder
    patterns = [
        f"{report_name}*.xlsx",
        f"{report_name}*.xlsb",
        f"{report_name}*.xls",
        "NBD*WF*18*DM*.xlsx",
        "NBD*WF*18*DM*.xlsb",
        "NBD*WF*18*DM*.xls"
    ]

    template_file = None
    for pattern in patterns:
        files = list(working_path.glob(pattern))
        if files:
            template_file = files[0]
            break

    if not template_file:
        print(f"⚠️ No template file found for {report_name}. Creating a basic template...")
        # Create a basic template in the working folder
        template_file = working_path / f"{report_name} Week {old_week}.xlsx"

        # Create a basic workbook with required sheets
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet

        # Create required sheets
        wb.create_sheet("Summary")
        wb.create_sheet("Rec")
        wb.create_sheet("Detailed")
        wb.create_sheet("Portfolio")

        wb.save(template_file)
        print(f"✅ Created basic template: {template_file.name}")

    # Rename the file in the same directory
    new_filename = re.sub(
        fr"Week\s*{old_week}\b",
        f"Week {new_week}",
        template_file.name
    )
    if new_filename == template_file.name:
        # If no week pattern found, add week number
        stem, suffix = os.path.splitext(template_file.name)
        new_filename = f"{stem} Week {new_week}{suffix}"

    new_path = working_path / new_filename

    # Rename the file
    if template_file != new_path:
        shutil.move(template_file, new_path)
        print(f"✅ Renamed: {template_file.name} → {new_filename}")
    else:
        print(f"✅ File already has correct name: {new_filename}")

    return new_path

# ------------------------------------------------------------
# STEP 2 & 3 – Update Summary & Rec last week values
# ------------------------------------------------------------
def update_summary_and_rec(filepath):
    wb = openpyxl.load_workbook(filepath)
    
    # Summary sheet
    if "Summary" in wb.sheetnames:
        ws_summary = wb["Summary"]
        for col in range(1, ws_summary.max_column + 1):
            ws_summary.cell(row=16, column=col).value = ws_summary.cell(row=11, column=col).value
        print("✅ Copied Summary row 11 → row 16")
    
    # Rec sheet
    if "Rec" in wb.sheetnames:
        ws_rec = wb["Rec"]
        for col in range(1, 5):  # A-D
            ws_rec.cell(row=13, column=col).value = ws_rec.cell(row=10, column=col).value
        print("✅ Copied Rec row 10 (A–D) → row 13")
    
    wb.save(filepath)
    wb.close()
    print(f"💾 Summary & Rec updated in {os.path.basename(filepath)}")

# ------------------------------------------------------------
# Helper function to convert accounting format to negative numbers
# ------------------------------------------------------------
def convert_accounting_to_number(value):
    """Convert accounting format (15280275.12) to negative number -15280275.12"""
    if isinstance(value, str):
        value = value.strip()
        # Check if value is in accounting format with parentheses
        if value.startswith('(') and value.endswith(')'):
            # Remove parentheses and convert to negative
            try:
                number_str = value[1:-1].replace(',', '').strip()
                return -float(number_str)
            except (ValueError, AttributeError):
                pass
        # Try normal conversion
        try:
            return float(value.replace(',', '').strip())
        except (ValueError, AttributeError):
            pass
    return value

# ------------------------------------------------------------
# STEP 4 – Paste data from CentralBankMovementOfDailyDepositBalances
# ------------------------------------------------------------
def paste_from_centralbank_report(target_file, working_dir):
    """Look for CSV/Excel files directly in the working date folder"""
    working_path = Path(working_dir)

    # Locate CentralBankMovement report (Excel or CSV) directly in the working folder
    cb_files = list(working_path.glob("CentralBankMovementOfDailyDepositBalances*.xls*")) + \
               list(working_path.glob("CentralBankMovementOfDailyDepositBalances*.csv"))
    mod_files = list(working_path.glob("MovementOfDepositDetails*.xls*")) + \
                list(working_path.glob("MovementOfDepositDetails*.csv"))
    
    if not cb_files:
        raise FileNotFoundError("No file starting with 'CentralBankMovementOfDailyDepositBalances' found.")
    if not mod_files:
        raise FileNotFoundError("No file starting with 'MovementOfDepositDetails' found.")
    
    cb_file = cb_files[0]
    mod_file = mod_files[0]

    # Load target workbook
    wb_target = openpyxl.load_workbook(target_file)

    # Handle CSV files
    if cb_file.suffix.lower() == '.csv' and mod_file.suffix.lower() == '.csv':
        # Read CSV data
        def read_csv_rows(path):
            with open(path, newline="", encoding="utf-8-sig") as f:
                return list(csv.reader(f))

        cb_rows = read_csv_rows(cb_file)
        mod_rows = read_csv_rows(mod_file)

        # Summary CSV: copy rows 5-12, cols A-I (0..8) → A4:I11
        if "Summary" in wb_target.sheetnames:
            ws_tgt_summary = wb_target["Summary"]
            for r in range(8):
                for c in range(9):
                    source_val = cb_rows[4 + r][0 + c] if len(cb_rows) > 4 + r and len(cb_rows[4 + r]) > c else None
                    ws_tgt_summary.cell(row=4 + r, column=1 + c).value = source_val

            # Format A4:A11 as dates (column A, rows 4-11)
            from datetime import datetime
            for r in range(4, 12):  # rows 4-11
                cell = ws_tgt_summary.cell(row=r, column=1)
                if isinstance(cell.value, str):
                    try:
                        # Try to parse date in DD/MM/YYYY or other formats
                        date_obj = datetime.strptime(cell.value.strip(), "%d/%m/%Y")
                        cell.value = date_obj
                        cell.number_format = 'DD/MM/YYYY'
                    except (ValueError, AttributeError):
                        try:
                            # Try DD-MM-YYYY format
                            date_obj = datetime.strptime(cell.value.strip(), "%d-%m-%Y")
                            cell.value = date_obj
                            cell.number_format = 'DD/MM/YYYY'
                        except:
                            pass

            # Format B4:I11 as numbers (columns 2 to 9, rows 4 to 11)
            for r in range(4, 12):  # rows 4-11
                for c in range(2, 10):  # columns B(2) to I(9)
                    cell = ws_tgt_summary.cell(row=r, column=c)
                    # Convert accounting format and string numbers to actual numbers
                    cell.value = convert_accounting_to_number(cell.value)
                    # Apply number format
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.00'

            # Format A16 as date
            cell_a16 = ws_tgt_summary.cell(row=16, column=1)
            if isinstance(cell_a16.value, str):
                try:
                    date_obj = datetime.strptime(cell_a16.value.strip(), "%d/%m/%Y")
                    cell_a16.value = date_obj
                    cell_a16.number_format = 'DD/MM/YYYY'
                except (ValueError, AttributeError):
                    try:
                        date_obj = datetime.strptime(cell_a16.value.strip(), "%d-%m-%Y")
                        cell_a16.value = date_obj
                        cell_a16.number_format = 'DD/MM/YYYY'
                    except:
                        pass

            print("✅ Pasted Summary A5:I12 (CSV) → A4:I11")
            print("   Formatted A4:A11, A16 as dates")
            print("   Formatted B4:I11 as numbers")

        # Detailed CSV: copy rows 5-12, cols A-W (0..22) → A8:W15
        if "Detailed" in wb_target.sheetnames:
            ws_tgt_detailed = wb_target["Detailed"]
            for r in range(8):
                for c in range(23):
                    source_val = mod_rows[4 + r][0 + c] if len(mod_rows) > 4 + r and len(mod_rows[4 + r]) > c else None
                    ws_tgt_detailed.cell(row=8 + r, column=1 + c).value = source_val

            # Format A8:A15 as dates (column A, rows 8-15)
            for r in range(8, 16):  # rows 8-15
                cell = ws_tgt_detailed.cell(row=r, column=1)
                if isinstance(cell.value, str):
                    try:
                        date_obj = datetime.strptime(cell.value.strip(), "%d/%m/%Y")
                        cell.value = date_obj
                        cell.number_format = 'DD/MM/YYYY'
                    except (ValueError, AttributeError):
                        try:
                            date_obj = datetime.strptime(cell.value.strip(), "%d-%m-%Y")
                            cell.value = date_obj
                            cell.number_format = 'DD/MM/YYYY'
                        except:
                            pass

            # Format C8:W15 as numbers (columns 3 to 23, rows 8 to 15)
            for r in range(8, 16):  # rows 8-15
                for c in range(3, 24):  # columns C(3) to W(23)
                    cell = ws_tgt_detailed.cell(row=r, column=c)
                    # Convert accounting format and string numbers to actual numbers
                    cell.value = convert_accounting_to_number(cell.value)
                    # Apply number format
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.00'

            print("✅ Pasted Detailed A5:W12 (CSV) → A8:W15")
            print("   Formatted A8:A15 as dates")
            print("   Formatted C8:W15 as numbers")

    else:
        # Handle Excel files
        wb_source = openpyxl.load_workbook(cb_file, data_only=True)

        # Summary sheet paste A5:I12 → Summary!A4:I11
        for sheet_name in wb_source.sheetnames:
            if "CentralBankMovementOfDailyDepositBalances" in sheet_name:
                ws_src_summary = wb_source[sheet_name]
                break
        else:
            raise ValueError("Source sheet 'CentralBankMovementOfDailyDepositBalances' not found.")

        if "Summary" in wb_target.sheetnames:
            ws_tgt_summary = wb_target["Summary"]
            for r in range(8):
                for c in range(9):
                    ws_tgt_summary.cell(row=4+r, column=1+c).value = ws_src_summary.cell(row=5+r, column=1+c).value

            # Format A4:A11 as dates (column A, rows 4-11)
            for r in range(4, 12):  # rows 4-11
                cell = ws_tgt_summary.cell(row=r, column=1)
                if isinstance(cell.value, str):
                    try:
                        date_obj = datetime.strptime(cell.value.strip(), "%d/%m/%Y")
                        cell.value = date_obj
                        cell.number_format = 'DD/MM/YYYY'
                    except (ValueError, AttributeError):
                        try:
                            date_obj = datetime.strptime(cell.value.strip(), "%d-%m-%Y")
                            cell.value = date_obj
                            cell.number_format = 'DD/MM/YYYY'
                        except:
                            pass

            # Format B4:I11 as numbers (columns 2 to 9, rows 4 to 11)
            for r in range(4, 12):  # rows 4-11
                for c in range(2, 10):  # columns B(2) to I(9)
                    cell = ws_tgt_summary.cell(row=r, column=c)
                    # Convert accounting format and string numbers to actual numbers
                    cell.value = convert_accounting_to_number(cell.value)
                    # Apply number format
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.00'

            # Format A16 as date
            cell_a16 = ws_tgt_summary.cell(row=16, column=1)
            if isinstance(cell_a16.value, str):
                try:
                    date_obj = datetime.strptime(cell_a16.value.strip(), "%d/%m/%Y")
                    cell_a16.value = date_obj
                    cell_a16.number_format = 'DD/MM/YYYY'
                except (ValueError, AttributeError):
                    try:
                        date_obj = datetime.strptime(cell_a16.value.strip(), "%d-%m-%Y")
                        cell_a16.value = date_obj
                        cell_a16.number_format = 'DD/MM/YYYY'
                    except:
                        pass

            print("✅ Pasted Summary A5:I12 → A4:I11")
            print("   Formatted A4:A11, A16 as dates")
            print("   Formatted B4:I11 as numbers")

        # Detailed sheet paste A5:W12 → Detailed!A8:W15
        for sheet_name in wb_source.sheetnames:
            if "MovementOfDepositDetails" in sheet_name:
                ws_src_detailed = wb_source[sheet_name]
                break
        else:
            raise ValueError("Source sheet 'MovementOfDepositDetails' not found.")

        if "Detailed" in wb_target.sheetnames:
            ws_tgt_detailed = wb_target["Detailed"]
            for r in range(8):
                for c in range(23):
                    ws_tgt_detailed.cell(row=8+r, column=1+c).value = ws_src_detailed.cell(row=5+r, column=1+c).value

            # Format A8:A15 as dates (column A, rows 8-15)
            for r in range(8, 16):  # rows 8-15
                cell = ws_tgt_detailed.cell(row=r, column=1)
                if isinstance(cell.value, str):
                    try:
                        date_obj = datetime.strptime(cell.value.strip(), "%d/%m/%Y")
                        cell.value = date_obj
                        cell.number_format = 'DD/MM/YYYY'
                    except (ValueError, AttributeError):
                        try:
                            date_obj = datetime.strptime(cell.value.strip(), "%d-%m-%Y")
                            cell.value = date_obj
                            cell.number_format = 'DD/MM/YYYY'
                        except:
                            pass

            # Format C8:W15 as numbers (columns 3 to 23, rows 8 to 15)
            for r in range(8, 16):  # rows 8-15
                for c in range(3, 24):  # columns C(3) to W(23)
                    cell = ws_tgt_detailed.cell(row=r, column=c)
                    # Convert accounting format and string numbers to actual numbers
                    cell.value = convert_accounting_to_number(cell.value)
                    # Apply number format
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.00'

            print("✅ Pasted Detailed A5:W12 → A8:W15")
            print("   Formatted A8:A15 as dates")
            print("   Formatted C8:W15 as numbers")

        wb_source.close()

    wb_target.save(target_file)
    wb_target.close()
    print(f"💾 CentralBankMovement data pasted to {os.path.basename(target_file)}")

# ------------------------------------------------------------
# STEP 5 – Validate Detailed sheet
# ------------------------------------------------------------
def validate_detailed_sheet(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    if "Detailed" not in wb.sheetnames:
        raise ValueError("'Detailed' sheet not found.")
    ws = wb["Detailed"]
    exceptions = []

    # Check A8:W15 columns not empty
    for col in range(1, 24):
        col_values = [ws.cell(row=r, column=col).value for r in range(8, 16)]
        if all(v is None or str(v).strip() == "" for v in col_values):
            col_letter = get_column_letter(col)
            exceptions.append(f"Column {col_letter} (A8:W15) is completely empty.")

    # Check O17:O24, W17:W24, E17:E24, J17:J24 = 0
    for cell_range in ["O17:O24", "W17:W24", "E17:E24", "J17:J24"]:
        for row in ws[cell_range]:
            for cell in row:
                val = cell.value
                if val not in (0, 0.0):
                    exceptions.append(f"Cell {cell.coordinate} = {val} (should be 0)")

    if exceptions:
        report_path = Path(filepath).parent / "NBD_WF_18_DM_Detailed_Exceptions.txt"
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("⚠️ Detailed Sheet Exception Report ⚠️\n\n")
            for e in exceptions:
                f.write(f"- {e}\n")
        print(f"❗ {len(exceptions)} exceptions found. Report saved to {report_path}")
    else:
        print("✅ No exceptions in Detailed sheet.")

    wb.close()

# ------------------------------------------------------------
# STEP 6 – Update Rec sheet with dates and Scienter data
# ------------------------------------------------------------
def update_rec_sheet_with_scienter_data(filepath, input_date):
    """Update Rec sheet with dates from Summary and collect data from Scienter

    Args:
        filepath: Path to the target Excel file
        input_date: datetime object for calculating date range
    """
    print("Collecting Deposit Base and Interest Payable data from Scienter...")

    # Import the Scienter rec bot
    from bots.DM_Scienter_rec import collect_rec_data

    # Collect data from Scienter
    rec_data = collect_rec_data(input_date)

    if not rec_data:
        raise Exception("Failed to collect Rec data from Scienter")

    # Load the workbook
    wb = openpyxl.load_workbook(filepath)
    ws_summary = wb["Summary"]
    ws_rec = wb["Rec"]

    # Step 1: Copy A3:D3 to A13:D13 (backup current week data to last week row)
    print("  Backing up A3:D3 to A13:D13...")
    for col in range(1, 5):  # Columns A to D
        value = ws_rec.cell(row=3, column=col).value
        ws_rec.cell(row=13, column=col).value = value

    # Step 2: Clear A3:D10 before writing new values
    print("  Clearing A3:D10...")
    for row in range(3, 11):  # Rows 3 to 10
        for col in range(1, 5):  # Columns A to D
            ws_rec.cell(row=row, column=col).value = None

    # Step 3: Copy dates from Summary A4:A11 to Rec A3:A10
    print("  Copying dates from Summary to Rec...")
    for i in range(8):  # Summary A4:A11 → Rec A3:A10
        date_val = ws_summary.cell(row=4+i, column=1).value
        ws_rec.cell(row=3+i, column=1).value = date_val

    # Step 4: Populate C3:C10 with deposit base values
    print("  Writing Deposit Base values to C3:C10...")
    for i in range(8):
        ws_rec.cell(row=3+i, column=3).value = rec_data['deposit_base'][i]

    # Step 5: Populate D3:D10 with interest payable values
    print("  Writing Interest Payable values to D3:D10...")
    for i in range(8):
        ws_rec.cell(row=3+i, column=4).value = rec_data['interest_payable'][i]

    # Step 6: Format A3:A13 as dates
    print("  Formatting A3:A13 as dates...")
    from datetime import datetime
    for r in range(3, 14):  # rows 3-13
        cell = ws_rec.cell(row=r, column=1)
        if isinstance(cell.value, str):
            try:
                date_obj = datetime.strptime(cell.value.strip(), "%d/%m/%Y")
                cell.value = date_obj
                cell.number_format = 'DD/MM/YYYY'
            except (ValueError, AttributeError):
                try:
                    date_obj = datetime.strptime(cell.value.strip(), "%d-%m-%Y")
                    cell.value = date_obj
                    cell.number_format = 'DD/MM/YYYY'
                except:
                    pass
        elif cell.value is not None and not isinstance(cell.value, datetime):
            # If it's already a datetime object from Summary, just apply format
            cell.number_format = 'DD/MM/YYYY'

    wb.save(filepath)
    wb.close()
    print("✅ Rec sheet updated with Scienter data")
    print(f"   - Deposit Base: {len(rec_data['deposit_base'])} values")
    print(f"   - Interest Payable: {len(rec_data['interest_payable'])} values")
    print(f"   - Formatted A3:A13 as dates")

# ------------------------------------------------------------
# STEP 7 – Copy FDBaseNew_UserBranch data with date formatting
# ------------------------------------------------------------
def parse_and_format_date(date_val, output_format="%d/%m/%Y"):
    """Convert various date formats to specified output format
    Args:
        date_val: Date value to convert
        output_format: Output format string (default: "%d/%m/%Y" for DD/MM/YYYY)
    """
    if date_val is None or date_val == "":
        return None
    
    # If already a datetime object
    if isinstance(date_val, datetime):
        return date_val.strftime(output_format)
    
    date_str = str(date_val).strip()
    
    # Skip if it's clearly not a date (too long or contains letters except month names)
    if len(date_str) > 50:
        return date_val
    
    # Try multiple date formats
    date_formats = [
        "%d/%m/%Y",      # 18/09/2025
        "%m/%d/%Y",      # 09/18/2025
        "%Y-%m-%d",      # 2025-09-18
        "%d-%m-%Y",      # 18-09-2025
        "%Y/%m/%d",      # 2025/09/18
        "%d.%m.%Y",      # 18.09.2025
        "%d %m %Y",      # 18 09 2025
        "%d/%m/%y",      # 18/09/25
        "%m/%d/%y",      # 09/18/25
        "%y/%m/%d",      # 25/09/18
    ]
    
    for fmt in date_formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime(output_format)
        except ValueError:
            continue
    
    # Try to handle Excel serial dates (numeric)
    try:
        # Excel stores dates as numbers (days since 1900-01-01)
        excel_date = float(date_str)
        # Valid Excel date range: 1 (1900-01-01) to 2958465 (9999-12-31)
        if 1 <= excel_date <= 2958465:
            # Excel incorrectly considers 1900 a leap year, so adjust
            if excel_date < 60:
                dt = datetime(1899, 12, 31) + timedelta(days=excel_date)
            else:
                dt = datetime(1899, 12, 31) + timedelta(days=excel_date - 1)
            return dt.strftime(output_format)
    except (ValueError, TypeError, OverflowError):
        pass
    
    # Return original if all parsing fails
    return date_val

def copy_fdbase_columns(file_path):
    path = Path(file_path)
    copied_data = []
    
    if path.suffix.lower() == '.csv':
        # CSV: Columns B,C,D,H,J,K,L,M,N,Q -> indices 1,2,3,7,9,10,11,12,13,16 (0-based)
        # Mapping: B=0, C=1, D=2, H=3, J=4, K=5, L=6, M=7, N=8, Q=9 in our extracted data
        csv_indices = [1, 2, 3, 7, 9, 10, 11, 12, 13, 16]
        
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = list(csv.reader(f))
            # Start at row 5 (1-based) -> index 4
            for r in range(4, len(reader)):
                row = reader[r]
                row_data = []
                for col_idx in csv_indices:
                    val = row[col_idx] if col_idx < len(row) else None
                    # Keep raw values - will be processed during paste
                    row_data.append(val)
                
                if any(v not in (None, "") for v in row_data):
                    copied_data.append(row_data)
    else:
        # Excel file - Load with data_only to get values
        wb_values = openpyxl.load_workbook(path, data_only=True)
        ws_values = wb_values.active
        
        # Columns: B,C,D,H,J,K,L,M,N,Q
        columns_to_copy = ["B", "C", "D", "H", "J", "K", "L", "M", "N", "Q"]
        
        max_row = ws_values.max_row
        for row in range(5, max_row + 1):
            row_data = []
            for col in columns_to_copy:
                # Get the value - keep datetime objects and numbers as-is
                val = ws_values[f"{col}{row}"].value
                row_data.append(val)
            
            if any(v not in (None, "") for v in row_data):
                copied_data.append(row_data)
        
        wb_values.close()

    print(f"✅ Copied {len(copied_data)} rows from FDBaseNew_UserBranch.")
    return copied_data

# ------------------------------------------------------------
# STEP 8 & 9 – Paste FDBaseNew_UserBranch → Portfolio
# ------------------------------------------------------------
def convert_date_string_to_excel_serial(date_str):
    """Convert date string to Excel serial number"""
    if date_str is None or date_str == "":
        return None
    
    # If it's already a datetime object
    if isinstance(date_str, datetime):
        excel_epoch = datetime(1899, 12, 30)
        delta = date_str - excel_epoch
        return delta.days
    
    # If it's already a number, return it
    if isinstance(date_str, (int, float)):
        return date_str
    
    # Parse the date string
    date_formats = [
        "%d/%m/%Y",      # 18/09/2025
        "%m/%d/%Y",      # 09/18/2025
        "%Y-%m-%d",      # 2025-09-18
        "%d-%m-%Y",      # 18-09-2025
        "%Y/%m/%d",      # 2025/09/18
        "%d/%m/%y",      # 18/09/25
        "%m/%d/%y",      # 09/18/25
    ]
    
    for fmt in date_formats:
        try:
            dt = datetime.strptime(str(date_str).strip(), fmt)
            # Convert to Excel serial number
            # Excel's date system: 1 = January 1, 1900
            excel_epoch = datetime(1899, 12, 30)
            delta = dt - excel_epoch
            return delta.days
        except ValueError:
            continue
    
    # If parsing fails, return original value
    return date_str

def paste_fdbase_to_portfolio(target_file, copied_data):
    wb = openpyxl.load_workbook(target_file)
    ws = wb["Portfolio"]

    # Clear A5:J[last row]
    max_row = ws.max_row
    for row in range(5, max_row + 1):
        for col in range(1, 11):
            ws.cell(row=row, column=col).value = None

    # Mapping of our extracted columns to Portfolio columns:
    # Extracted: B=0, C=1, D=2, H=3, J=4, K=5, L=6, M=7, N=8, Q=9
    # Portfolio: A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10
    # So: extracted index 2 (D) -> Portfolio column 4 (D)
    #     extracted index 3 (H) -> Portfolio column 5 (E)
    #     extracted index 5 (K) -> Portfolio column 6 (F) - Maturity Date
    
    # Date columns in extracted data: index 2 (source D) and index 3 (source H)
    # These go to Portfolio columns D (4) and E (5)
    date_column_indices = [2, 3]  # Indices in extracted data
    
    # Maturity Date column (F): index 5 in extracted data
    maturity_date_column_index = 5
    
    # Paste data starting row 5
    for i, row_data in enumerate(copied_data):
        for j, val in enumerate(row_data):
            cell = ws.cell(row=5+i, column=j+1)
            
            # Convert date strings to Excel serial numbers for date columns
            if j in date_column_indices:
                val = convert_date_string_to_excel_serial(val)
                cell.value = val
                # Apply date format to the cell
                if val is not None and isinstance(val, (int, float)):
                    # Portfolio E (column 5) corresponds to extracted index 3
                    if j == 3:
                        cell.number_format = 'MM/DD/YYYY'
                    else:
                        cell.number_format = 'M/D/YYYY'
            # Handle Maturity Date column (F)
            elif j == maturity_date_column_index:
                val = convert_date_string_to_excel_serial(val)
                cell.value = val
                # Apply M/D/YYYY format (e.g., 7/30/2025)
                if val is not None and isinstance(val, (int, float)):
                    cell.number_format = 'M/D/YYYY'
            # Force H, I, J columns (indices 7, 8, 9) as NUMBERS
            elif j in [7, 8, 9]:
                try:
                    # Convert string numbers to float, remove commas
                    if val is not None and val != "":
                        val = float(str(val).replace(",", "").strip())
                    else:
                        val = None
                except (ValueError, AttributeError):
                    val = None  # fallback if conversion fails
                
                cell.value = val
                # Apply different formats: H=General, I=Accounting, J=Number with commas
                if j == 7:  # H column
                    cell.number_format = 'General'
                elif j == 8:  # I column
                    cell.number_format = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'
                elif j == 9:  # J column
                    cell.number_format = '#,##0'
            else:
                cell.value = val

    # Find the last row with data in column B
    last_data_row = 4  # Start from row 4 (before data rows start at 5)
    for row in range(5, ws.max_row + 1):
        if ws.cell(row=row, column=2).value is not None and ws.cell(row=row, column=2).value != "":
            last_data_row = row
        else:
            break

    # Apply formulas for K,L,M,N,O,P,Q (columns 11,12,13,14,15,16,17)
    # Mimic Excel's Ctrl+D behavior: copy row 5 formulas down to all rows
    formula_cols = [11, 12, 13, 14, 15, 16, 17]  # K, L, M, N, O, P, Q
    template_row = 5  # Row 5 contains the actual formulas (row 4 is headers)

    print(f"   Extending formulas from row {template_row} to row {last_data_row}...")

    # Use openpyxl's Translator to properly handle formula translation (like Excel's Ctrl+D)
    from openpyxl.formula.translate import Translator

    for col in formula_cols:
        # Get the template cell from row 5
        template_cell = ws.cell(row=template_row, column=col)

        # Copy formula to all rows from 6 to last_data_row
        for row in range(6, last_data_row + 1):
            target_cell = ws.cell(row=row, column=col)

            # Copy cell value (formula or value)
            if template_cell.value:
                if str(template_cell.value).startswith("="):
                    # It's a formula - use openpyxl's translate_formula
                    # Calculate row offset from template
                    row_offset = row - template_row
                    # Translate formula with row offset (like Excel's Ctrl+D)
                    translated_formula = Translator(str(template_cell.value), origin=f"{get_column_letter(col)}{template_row}").translate_formula(row_delta=row_offset, col_delta=0)
                    target_cell.value = translated_formula
                else:
                    # It's a value, just copy it
                    target_cell.value = template_cell.value

    # Apply number formats to formula columns
    print(f"   Applying number formats to columns K-Q...")
    for row in range(5, last_data_row + 1):
        # K: Accounting format, rounded, no decimals
        ws.cell(row=row, column=11).number_format = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'

        # L: Number format
        ws.cell(row=row, column=12).number_format = '0'

        # M: Date format (MM/DD/YYYY)
        ws.cell(row=row, column=13).number_format = 'MM/DD/YYYY'

        # N: Accounting format
        ws.cell(row=row, column=14).number_format = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'

        # O: Accounting format
        ws.cell(row=row, column=15).number_format = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'

        # P: Accounting format
        ws.cell(row=row, column=16).number_format = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'

        # Q: Accounting format
        ws.cell(row=row, column=17).number_format = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'

    # Update SUM formulas in row 3 for columns N, O, Q with dynamic last row
    # N3 = SUM(N5:N[last_row])
    ws.cell(row=3, column=14).value = f"=SUM(N5:N{last_data_row})"
    # O3 = SUM(O5:O[last_row])
    ws.cell(row=3, column=15).value = f"=SUM(O5:O{last_data_row})"
    # Q3 = SUM(Q5:Q[last_row])
    ws.cell(row=3, column=17).value = f"=SUM(Q5:Q{last_data_row})"

    print(f"   Updated SUM formulas in row 3 (N3, O3, Q3) with last row: {last_data_row}")

    # Normalize existing Portfolio column E (5) dates to true Date type with MM/DD/YYYY
    last_row = ws.max_row
    for r in range(5, last_row + 1):
        cell_e = ws.cell(row=r, column=5)
        val = cell_e.value
        if isinstance(val, str):
            s = val.strip()
            m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", s)
            if m:
                d_str, m_str, y_str = m.groups()
                if len(y_str) == 2:
                    y_str = "20" + y_str
                try:
                    year = int(y_str)
                    month = int(m_str)
                    day = int(d_str)
                    dt = datetime(year, month, day)
                    cell_e.value = dt
                    cell_e.number_format = 'MM/DD/YYYY'
                except ValueError:
                    # Leave value as-is if invalid date
                    pass
        elif isinstance(val, datetime):
            cell_e.number_format = 'MM/DD/YYYY'
        elif isinstance(val, (int, float)):
            cell_e.number_format = 'MM/DD/YYYY'

    # Extra safety: Force H, I, J columns (8, 9, 10) to be numeric
    for r in range(5, ws.max_row + 1):
        for col in [8, 9, 10]:  # H, I, J
            cell = ws.cell(row=r, column=col)
            if isinstance(cell.value, str):
                try:
                    cell.value = float(cell.value.replace(",", "").strip())
                    # Apply different formats: H=General, I=Accounting, J=Number with commas
                    if col == 8:  # H column
                        cell.number_format = 'General'
                    elif col == 9:  # I column
                        cell.number_format = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'
                    elif col == 10:  # J column
                        cell.number_format = '#,##0'
                except (ValueError, AttributeError):
                    pass

    wb.save(target_file)
    wb.close()
    print("✅ FDBaseNew_UserBranch data pasted to Portfolio with formulas.")
    print("   - Date columns D and E converted to Excel date format (MM/DD/YYYY)")

# ------------------------------------------------------------
# STEP 10 – Update Portfolio column P from InterestPayableAsAtDateReport
# ------------------------------------------------------------
def update_portfolio_column_p(portfolio_file, interest_file):
    wb_portfolio = openpyxl.load_workbook(portfolio_file)
    ws_portfolio = wb_portfolio["Portfolio"]
    max_row_portfolio = ws_portfolio.max_row

    # Find "System Deposit No:" column in Portfolio sheet (header in row 4)
    portfolio_deposit_col = None
    for col in range(1, ws_portfolio.max_column + 1):
        header = ws_portfolio.cell(row=4, column=col).value
        if header and "System Deposit No:" in str(header):
            portfolio_deposit_col = col
            break
    
    if not portfolio_deposit_col:
        print("⚠️ Warning: 'System Deposit No:' header not found in Portfolio sheet. Defaulting to column C.")
        portfolio_deposit_col = 3

    # Clear column P
    for row in range(5, max_row_portfolio + 1):
        ws_portfolio.cell(row=row, column=16).value = None

    interest_path = Path(interest_file)
    interest_data = []

    if interest_path.suffix.lower() == '.csv':
        # CSV file - headers in row 4 (index 3)
        with open(interest_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            # Read first 3 rows to skip to header row
            for _ in range(3):
                next(reader, None)
            
            # Read header row (4th row)
            header_row = next(reader, None)
            if not header_row:
                raise ValueError("InterestPayableAsAtDateReport CSV is empty or missing header row")
            
            # Find fdno and INTPAYABLE columns
            fdno_col = None
            intpayable_col = None
            for idx, header in enumerate(header_row):
                if header and "fdno" in str(header).lower():
                    fdno_col = idx
                if header and "intpayable" in str(header).lower():
                    intpayable_col = idx
            
            if fdno_col is None:
                print("⚠️ Warning: 'fdno' column not found in CSV. Defaulting to column 3.")
                fdno_col = 2  # index 2 = column 3
            if intpayable_col is None:
                print("⚠️ Warning: 'INTPAYABLE' column not found in CSV. Defaulting to column 18.")
                intpayable_col = 17  # index 17 = column 18
            
            # Read data rows
            for row in reader:
                fdno = row[fdno_col] if len(row) > fdno_col else None
                intpayable = row[intpayable_col] if len(row) > intpayable_col else None
                try:
                    intpayable_val = float(intpayable) if intpayable not in (None, "") else None
                except ValueError:
                    intpayable_val = None
                if fdno not in (None, "") and intpayable_val is not None:
                    interest_data.append((str(fdno), intpayable_val))
    else:
        # Excel file - headers in row 4
        wb_interest = openpyxl.load_workbook(interest_path, data_only=True)
        ws_interest = wb_interest.active
        
        # Find fdno and INTPAYABLE columns from header row (row 4)
        fdno_col = None
        intpayable_col = None
        for col in range(1, ws_interest.max_column + 1):
            header = ws_interest.cell(row=4, column=col).value
            if header and "fdno" in str(header).lower():
                fdno_col = col
            if header and "intpayable" in str(header).lower():
                intpayable_col = col
        
        if fdno_col is None:
            print("⚠️ Warning: 'fdno' column not found in Excel. Defaulting to column 3.")
            fdno_col = 3
        if intpayable_col is None:
            print("⚠️ Warning: 'INTPAYABLE' column not found in Excel. Defaulting to column 18.")
            intpayable_col = 18
        
        # Read data rows (starting from row 5)
        for row in range(5, ws_interest.max_row + 1):
            fdno = ws_interest.cell(row=row, column=fdno_col).value
            intpayable = ws_interest.cell(row=row, column=intpayable_col).value
            if fdno is not None and intpayable is not None:
                interest_data.append((str(fdno), float(intpayable)))
        wb_interest.close()

    # Sort by INTPAYABLE ascending
    interest_data.sort(key=lambda x: x[1])
    lookup_dict = {fdno: intpayable for fdno, intpayable in interest_data}

    # VLOOKUP: fill Portfolio!P
    for row in range(5, max_row_portfolio + 1):
        fdno = ws_portfolio.cell(row=row, column=portfolio_deposit_col).value
        if fdno is not None:
            ws_portfolio.cell(row=row, column=16).value = lookup_dict.get(str(fdno), 0)

    wb_portfolio.save(portfolio_file)
    wb_portfolio.close()
    print("✅ Portfolio column P updated from InterestPayableAsAtDateReport (using header-based lookup).")

# ------------------------------------------------------------
# STEP 11 – Q3 check for Scienter
# ------------------------------------------------------------
def check_portfolio_q3_for_scienter(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb["Portfolio"]
    q3_val = ws["Q3"].value
    if q3_val is not None and isinstance(q3_val, (int, float)) and q3_val > 100:
        print(f"⚠️ Portfolio!Q3 = {q3_val} > 100 → Scienter system check required (ToDo).")
    else:
        print("✅ Portfolio!Q3 <= 100 → no Scienter check needed.")
    wb.close()

# ------------------------------------------------------------
# STEP 12 – Add VLOOKUP formulas in TB Tally sheet
# ------------------------------------------------------------
def add_tb_tally_vlookup(target_file, working_dir):
    """
    Perform VLOOKUP: Match TB Tally column A with Report.xlsx column A (LEDGER_CODE),
    import AMOUNT from Report.xlsx column C to TB Tally column C.

    Report.xlsx structure:
    - Row 5: Headers (LEDGER_CODE | DESCRIPTION | AMOUNT)
    - Row 6+: Data
    - Column A: LEDGER_CODE (lookup key)
    - Column C: AMOUNT (value to import)
    """
    import pandas as pd

    working_path = Path(working_dir)

    # Find Report.xlsx file (with date prefix like 2026_01_27_15_27_Report.xlsx)
    tb_files = list(working_path.glob("*Report.xlsx"))
    if not tb_files:
        tb_files = list(working_path.glob("*TB_DetailReport.xlsx"))
    if not tb_files:
        print("⚠️ Report.xlsx file not found. Skipping TB Tally VLOOKUP.")
        return

    tb_file = tb_files[0]
    print(f"   Loading {tb_file.name}...")

    # Build lookup dictionary: LEDGER_CODE (A) -> AMOUNT (C)
    # Read directly from ZIP/XML to avoid openpyxl merged cell corruption issues
    lookup_dict = {}

    try:
        import zipfile
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(tb_file, 'r') as z:
            # Read the first sheet
            sheet_xml = z.read('xl/worksheets/sheet1.xml')
            shared_strings_xml = None
            try:
                shared_strings_xml = z.read('xl/sharedStrings.xml')
            except KeyError:
                pass

        # Parse shared strings if available
        shared_strings = []
        if shared_strings_xml:
            ss_root = ET.fromstring(shared_strings_xml)
            for si in ss_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                shared_strings.append(si.text if si.text else '')

        # Parse sheet data
        root = ET.fromstring(sheet_xml)

        # Extract rows starting from row 6 (data rows, row 5 is header)
        for row_elem in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
            row_num = int(row_elem.get('r', 0))
            if row_num >= 6:  # Data starts from row 6
                cells = {}
                for cell in row_elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    cell_ref = cell.get('r', '')
                    cell_type = cell.get('t', '')
                    value_elem = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')

                    if value_elem is not None and value_elem.text:
                        # Get column letter from cell reference (e.g., "A6" -> "A")
                        col_letter = ''.join(c for c in cell_ref if c.isalpha())

                        if cell_type == 's':  # Shared string
                            idx = int(value_elem.text)
                            if idx < len(shared_strings):
                                cells[col_letter] = shared_strings[idx]
                        else:
                            cells[col_letter] = value_elem.text

                # Extract column A (LEDGER_CODE) and column C (AMOUNT)
                if 'A' in cells:
                    ledger_code = str(cells['A']).strip()
                    amount = None
                    if 'C' in cells:
                        try:
                            amount = float(cells['C'])
                        except (ValueError, TypeError):
                            amount = cells['C']
                    if ledger_code:
                        lookup_dict[ledger_code] = amount

        print(f"   Loaded {len(lookup_dict)} lookup entries from Report.xlsx (LEDGER_CODE -> AMOUNT)")

    except Exception as e:
        print(f"⚠️ Failed to load Report.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return

    if not lookup_dict:
        print("⚠️ No lookup data extracted from Report.xlsx")
        return

    # Update TB Tally sheet in target DM file
    wb = openpyxl.load_workbook(target_file)

    if "TB Tally" not in wb.sheetnames:
        print("⚠️ TB Tally sheet not found in target file. Skipping.")
        wb.close()
        return

    ws = wb["TB Tally"]
    lookup_count = 0
    matched_count = 0

    # Iterate from row 3 to last row with data in column A
    for row in range(3, ws.max_row + 1):
        cell_a = ws.cell(row=row, column=1).value  # Column A: lookup key
        if cell_a is not None and cell_a != "":
            lookup_key = str(cell_a).strip()
            amount = lookup_dict.get(lookup_key)

            if amount is not None:
                ws.cell(row=row, column=3).value = amount  # Column C: AMOUNT
                matched_count += 1

            lookup_count += 1

    wb.save(target_file)
    wb.close()
    print(f"✅ VLOOKUP completed: {matched_count}/{lookup_count} AMOUNT values populated in TB Tally column C")

# ------------------------------------------------------------
# STEP 13 – Copy DM file to NBD_WF_15_LA output folder
# ------------------------------------------------------------
def copy_dm_file_to_la_output(source_file, input_date):
    """
    Copy the created NBD-WF-18-DM file to NBD_WF_15_LA output folder.

    Finds the target folder by:
    1. Reading the last date from outputs/NBD_WF_15_LA/completed_dates.txt
    2. Finding the folder with that date (or highest version if multiple exist)
    3. Replacing any existing NBD-WF-18-DM Deposit Liability*.xlsx file
    4. Renaming the file to NBD-WF-18-DM Deposit Liability (input_date).xlsx

    Args:
        source_file: Path to the created NBD-WF-18-DM file
        input_date: datetime object for the input date (used for renaming)
    """
    script_dir = Path(__file__).parent
    la_output_dir = script_dir.parent / "outputs" / "NBD_WF_15_LA"
    completed_dates_file = la_output_dir / "completed_dates.txt"

    # Check if completed_dates.txt exists
    if not completed_dates_file.exists():
        print(f"⚠️ completed_dates.txt not found at {completed_dates_file}")
        print("   Skipping copy to NBD_WF_15_LA output folder.")
        return False

    # Read the last non-empty date from completed_dates.txt
    with open(completed_dates_file, 'r') as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]

    if not lines:
        print("⚠️ completed_dates.txt is empty. Skipping copy.")
        return False

    target_date = lines[-1]  # Last written date
    print(f"   Target date from completed_dates.txt: {target_date}")

    # Find folders matching the target date (could be "30-06-2025" or "30-06-2025(1)", etc.)
    matching_folders = []
    for folder in la_output_dir.iterdir():
        if folder.is_dir():
            folder_name = folder.name
            # Check if folder name starts with target_date
            if folder_name == target_date or folder_name.startswith(f"{target_date}("):
                matching_folders.append(folder)

    if not matching_folders:
        print(f"⚠️ No folder found for date {target_date} in {la_output_dir}")
        return False

    # Find the maximum version folder
    # Sort by: base name first, then by version number
    def get_version(folder):
        name = folder.name
        if name == target_date:
            return 0  # Base folder (no version)
        # Extract version number from "date(n)" format
        if "(" in name and name.endswith(")"):
            try:
                version_str = name.split("(")[-1].rstrip(")")
                return int(version_str)
            except ValueError:
                return 0
        return 0

    target_folder = max(matching_folders, key=get_version)
    print(f"   Target folder: {target_folder}")

    # Find and delete existing NBD-WF-18-DM Deposit Liability*.xlsx files
    existing_files = list(target_folder.glob("NBD-WF-18-DM*.xlsx"))
    for existing_file in existing_files:
        print(f"   Removing existing file: {existing_file.name}")
        existing_file.unlink()

    # Format input_date as dd-mm-yyyy for the filename
    date_str = input_date.strftime("%d-%m-%Y")
    new_filename = f"NBD-WF-18-DM Deposit Liability ({date_str}).xlsx"

    # Copy the source file to target folder with new name
    source_path = Path(source_file)
    target_path = target_folder / new_filename
    shutil.copy2(source_path, target_path)
    print(f"✅ Copied and renamed to {new_filename} in {target_folder}")

    return True

# ------------------------------------------------------------
# STEP 14 – Copy DM file to NBD_WF_18_ECS output folder
# ------------------------------------------------------------
def copy_dm_file_to_ecs_output(source_file, input_date):
    """
    Copy the created NBD-WF-18-DM file to NBD_WF_18_ECS output folder.

    Finds the target folder by:
    1. Reading the last date from outputs/NBD_WF_18_ECS/completed_dates.txt
    2. Finding the folder with that date (or highest version if multiple exist)
    3. Replacing any existing NBD-WF-18-DM Deposit Liability*.xlsx file
    4. Renaming the file to NBD-WF-18-DM Deposit Liability (input_date).xlsx

    Args:
        source_file: Path to the created NBD-WF-18-DM file
        input_date: datetime object for the input date (used for renaming)
    """
    script_dir = Path(__file__).parent
    ecs_output_dir = script_dir.parent / "outputs" / "NBD_WF_18_ECS"
    completed_dates_file = ecs_output_dir / "completed_dates.txt"

    # Check if completed_dates.txt exists
    if not completed_dates_file.exists():
        print(f"⚠️ completed_dates.txt not found at {completed_dates_file}")
        return False

    # Read the last non-empty date from completed_dates.txt
    with open(completed_dates_file, 'r') as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]

    if not lines:
        print("⚠️ completed_dates.txt is empty. Skipping copy.")
        return False

    target_date = lines[-1]
    print(f"   Target date from completed_dates.txt: {target_date}")

    # Find folders matching the target date (could be "26-02-2026" or "26-02-2026(1)", etc.)
    matching_folders = []
    for folder in ecs_output_dir.iterdir():
        if folder.is_dir():
            if folder.name == target_date or folder.name.startswith(f"{target_date}("):
                matching_folders.append(folder)

    if not matching_folders:
        print(f"⚠️ No folder found for date {target_date} in {ecs_output_dir}")
        return False

    # Find the maximum version folder
    def get_version(folder):
        name = folder.name
        if name == target_date:
            return 0
        if "(" in name and name.endswith(")"):
            try:
                return int(name.split("(")[-1].rstrip(")"))
            except ValueError:
                return 0
        return 0

    target_folder = max(matching_folders, key=get_version)
    print(f"   Target folder: {target_folder}")

    # Find and delete existing NBD-WF-18-DM Deposit Liability*.xlsx files
    existing_files = list(target_folder.glob("NBD-WF-18-DM*.xlsx"))
    for existing_file in existing_files:
        print(f"   Removing existing file: {existing_file.name}")
        existing_file.unlink()

    # Format input_date as dd-mm-yyyy for the filename
    date_str = input_date.strftime("%d-%m-%Y")
    new_filename = f"NBD-WF-18-DM Deposit Liability ({date_str}).xlsx"

    # Copy the source file to target folder with new name
    source_path = Path(source_file)
    target_path = target_folder / new_filename
    shutil.copy2(source_path, target_path)
    print(f"✅ Copied and renamed to {new_filename} in {target_folder}")

    return True

# ------------------------------------------------------------
# MAIN EXECUTION
# ------------------------------------------------------------
def main():
    # Parse command line arguments for input_date
    if len(sys.argv) < 2:
        print("Usage: python NBD_WF_18_DM.py MM/DD/YYYY")
        sys.exit(1)

    try:
        input_date = datetime.strptime(sys.argv[1], "%m/%d/%Y")
    except ValueError:
        print("Error: Invalid date format. Use MM/DD/YYYY")
        sys.exit(1)

    # Get relative path to working/NBD-WF-18-DM directory
    script_dir = Path(__file__).parent
    base_dir = script_dir.parent / "working" / "NBD-WF-18-DM"

    # Check if working directory exists
    if not base_dir.exists():
        raise FileNotFoundError(f"Working directory not found: {base_dir}")

    # Find the single date folder inside NBD-WF-18-DM
    date_folders = [f for f in base_dir.iterdir() if f.is_dir()]
    if not date_folders:
        raise FileNotFoundError(f"No date folder found inside '{base_dir}' directory.")

    # Use the first (and should be only) date folder
    if len(date_folders) > 1:
        print(f"Warning: Multiple date folders found: {[f.name for f in date_folders]}. Using the first one: {date_folders[0].name}")

    # Working directory is the date folder
    working_date_folder = date_folders[0]
    print(f"Using working directory: {working_date_folder}")

    # Remove any stale TB report files carried over from a previous run
    for _old_tb in working_date_folder.glob("*Report.xlsx"):
        try:
            _old_tb.unlink()
            print(f"Removed stale TB file: {_old_tb.name}")
        except Exception as _e:
            print(f"Warning: could not remove {_old_tb.name}: {_e}")

    # Output will be in the same working folder
    output_dir = working_date_folder
    report_name = "NBD-WF-18-DM"
    old_week = 40
    new_week = 41

    print("=" * 60)
    print("Starting NBD-WF-18-DM Deposit Liability Report Processing")
    print("=" * 60)

    try:
        # Step 0 – Download required files BEFORE processing
        print("\n[STEP 0] Downloading required files...")
        download_success = download_required_files(input_date, working_date_folder)
        if not download_success:
            raise Exception("Failed to download one or more required files. Check errors above.")
        print("All required files downloaded successfully. Proceeding with report processing...\n")

        # Note: From STEP 1 onwards, the processing logic is from NBD_WF_18_DM (1).py
        # Step 1 – Rename template file
        print("\n[STEP 1] Renaming template file...")
        renamed_file = rename_wf18dm_report(working_date_folder, report_name, old_week, new_week)

        # Step 2 & 3 – Summary & Rec update
        print("\n[STEP 2-3] Updating Summary & Rec sheets...")
        update_summary_and_rec(renamed_file)

        # Step 4 – Paste CentralBankMovement data (look in the same date folder)
        print("\n[STEP 4] Pasting CentralBankMovement data...")
        paste_from_centralbank_report(renamed_file, working_date_folder)

        # Step 5 – Validate Detailed sheet
        print("\n[STEP 5] Validating Detailed sheet...")
        validate_detailed_sheet(renamed_file)

        # Step 6 – Update Rec sheet with Scienter data
        print("\n[STEP 6] Updating Rec sheet with Scienter data...")
        update_rec_sheet_with_scienter_data(renamed_file, input_date)

        # Step 7 – Copy FDBaseNew_UserBranch (look in the same date folder)
        print("\n[STEP 7] Copying FDBaseNew_UserBranch data...")
        fdbase_candidates = list(working_date_folder.glob("FDBaseNew_UserBranch*.xlsx")) + \
                            list(working_date_folder.glob("FDBaseNew_UserBranch*.csv"))
        if not fdbase_candidates:
            raise FileNotFoundError("FDBaseNew_UserBranch file not found in the working folder.")
        copied_data = copy_fdbase_columns(fdbase_candidates[0])

        # Step 8 & 9 – Paste FDBaseNew_UserBranch → Portfolio
        print("\n[STEP 8-9] Pasting data to Portfolio sheet...")
        paste_fdbase_to_portfolio(renamed_file, copied_data)

        # Step 10 – Update Portfolio column P
        print("\n[STEP 10] Updating Portfolio column P...")
        interest_candidates = list(working_date_folder.glob("InterestPayableAsAtDateReport*.xlsx")) + \
                              list(working_date_folder.glob("InterestPayableAsAtDateReport*.csv"))
        if not interest_candidates:
            raise FileNotFoundError("InterestPayableAsAtDateReport file not found in the working folder.")
        update_portfolio_column_p(renamed_file, interest_candidates[0])

        # Step 11 – Q3 check for Scienter
        print("\n[STEP 11] Checking Portfolio Q3 for Scienter...")
        check_portfolio_q3_for_scienter(renamed_file)

        # Step 12 – Add VLOOKUP formulas in TB Tally sheet
        print("\n[STEP 12] Adding VLOOKUP formulas to TB Tally sheet...")
        add_tb_tally_vlookup(renamed_file, working_date_folder)

        # Step 13 – Copy DM file to NBD_WF_15_LA output folder
        print("\n[STEP 13] Copying DM file to NBD_WF_15_LA output folder...")
        copy_dm_file_to_la_output(renamed_file, input_date)

        # Step 14 – Copy DM file to NBD_WF_18_ECS output folder
        print("\n[STEP 14] Copying DM file to NBD_WF_18_ECS output folder...")
        copy_dm_file_to_ecs_output(renamed_file, input_date)

        print("\n" + "=" * 60)
        print("✅ PROCESSING COMPLETED SUCCESSFULLY!")
        print(f"📁 Output file: {renamed_file}")
        print("=" * 60)

    except Exception as e:
        print("\n" + "=" * 60)
        print(f"❌ ERROR: {str(e)}")
        print("=" * 60)
        raise

if __name__ == "__main__":
    main()