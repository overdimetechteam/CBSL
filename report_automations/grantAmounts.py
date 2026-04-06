import pandas as pd
from pathlib import Path
import win32com.client
import pythoncom
from datetime import datetime

try:
    from pyxlsb import open_workbook as open_xlsb
    XLSB_SUPPORT = True
except ImportError:
    XLSB_SUPPORT = False
    print("Warning: pyxlsb not installed. .xlsb files will not be supported.")
    print("Install with: pip install pyxlsb")

def get_portfolio_summary_file(portfolios_dir: Path, quarter: int, fy_year: int) -> Path:
    """
    Find portfolio summary file for the last month of the quarter.
    Matches any summary file from the quarter months and returns the most recent.
    Mirrors the logic from UnwindOfDiscount.py
    
    Args:
        portfolios_dir: Path to Portfolios directory
        quarter: Quarter number (1-4)
        fy_year: Financial year
    
    Returns:
        Path to the Summary file or None
    """
    # Determine quarter months
    quarter_months_map = {
        1: [(4, fy_year), (5, fy_year), (6, fy_year)],        # Apr, May, Jun
        2: [(7, fy_year), (8, fy_year), (9, fy_year)],        # Jul, Aug, Sep
        3: [(10, fy_year), (11, fy_year), (12, fy_year)],     # Oct, Nov, Dec
        4: [(1, fy_year + 1), (2, fy_year + 1), (3, fy_year + 1)]  # Jan, Feb, Mar (next year)
    }
    
    quarter_months = quarter_months_map[quarter]
    
    month_map = {
        1: "01", 2: "02", 3: "03", 4: "04", 5: "05", 6: "06",
        7: "07", 8: "08", 9: "09", 10: "10", 11: "11", 12: "12"
    }
    
    if not portfolios_dir.exists():
        print(f"✗ Portfolios directory not found: {portfolios_dir}")
        return None
    
    # Search for matching files from all quarter months
    matching_files = []
    
    for month_num, year in quarter_months:
        month_str = month_map[month_num]
        pattern_prefix = f"3. Summary_{year}-{month_str}-"
        
        # Search for files matching this month
        for file in portfolios_dir.iterdir():
            if file.is_file() and file.name.startswith(pattern_prefix):
                # Check for valid Excel extensions
                if file.suffix.lower() in ['.xlsx', '.xlsb', '.xlsm']:
                    matching_files.append(file)
    
    if not matching_files:
        # Build helpful error message
        month_patterns = []
        for month_num, year in quarter_months:
            month_str = month_map[month_num]
            month_patterns.append(f"3. Summary_{year}-{month_str}-DD")
        
        print(f"✗ Summary file not found in {portfolios_dir}")
        print(f"   Expected patterns (any of):")
        for pattern in month_patterns:
            print(f"     - {pattern}*.xlsb")
        print(f"   Example: {month_patterns[-1]} Final V7.1.xlsb")
        return None
    
    # Sort by modification time and get the most recent (last month)
    # This ensures we get the latest month's file in the quarter
    matching_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    
    selected_file = matching_files[0]
    print(f"   Found {len(matching_files)} matching file(s) across quarter months")
    print(f"   Selected most recent: {selected_file.name}")
    
    return selected_file

def load_stage_from_summary_xlsb(summary_file: Path, contract_no: str) -> int:
    """
    Load stage value for a specific contract from .xlsb Summary file.
    
    Args:
        summary_file: Path to Summary .xlsb file
        contract_no: Contract number to lookup
    
    Returns:
        Stage value as integer, or None if not found
    """
    if not XLSB_SUPPORT:
        print(f"✗ Cannot read .xlsb file: pyxlsb not installed")
        return None
    
    try:
        with open_xlsb(summary_file) as wb:
            if "SUMMARY" not in wb.sheets:
                return None
            
            with wb.get_sheet("SUMMARY") as sheet:
                rows = list(sheet.rows())
                
                if not rows or len(rows) < 3:
                    return None
                
                # Header is in row 3 (index 2)
                header_row = rows[2]
                contract_col = None
                stage_col = None
                
                # Find CONTRACT NO and STAGE columns
                for col_idx, cell in enumerate(header_row):
                    if cell.v:
                        cell_str = str(cell.v).strip().upper()
                        if "CONTRACT" in cell_str and ("NO" in cell_str or "NUMBER" in cell_str):
                            contract_col = col_idx
                        elif "STAGE" in cell_str:
                            stage_col = col_idx
                
                if contract_col is None or stage_col is None:
                    return None
                
                # Search for contract number
                contract_key = str(contract_no).strip()
                for row in rows[3:]:
                    if len(row) > max(contract_col, stage_col):
                        contract_cell = row[contract_col]
                        
                        if contract_cell and contract_cell.v:
                            row_contract = str(contract_cell.v).strip()
                            
                            if row_contract == contract_key:
                                stage_cell = row[stage_col]
                                if stage_cell and stage_cell.v is not None:
                                    try:
                                        return int(float(stage_cell.v))
                                    except (ValueError, TypeError):
                                        pass
    except Exception as e:
        print(f"✗ Error reading {summary_file.name}: {e}")
    
    return None

def load_stage_from_summary_xlsx(summary_file: Path, contract_no: str) -> int:
    """
    Load stage value for a specific contract from .xlsx/.xlsm Summary file.
    
    Args:
        summary_file: Path to Summary .xlsx/.xlsm file
        contract_no: Contract number to lookup
    
    Returns:
        Stage value as integer, or None if not found
    """
    try:
        from openpyxl import load_workbook
        
        wb = load_workbook(summary_file, data_only=True, read_only=True)
        
        if "SUMMARY" not in wb.sheetnames:
            wb.close()
            return None
        
        sheet = wb["SUMMARY"]
        
        # Check row 3 for headers
        contract_col = None
        stage_col = None
        header_row = 3
        
        # Find CONTRACT NO and STAGE columns
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=header_row, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip().upper()
                if "CONTRACT" in cell_str and ("NO" in cell_str or "NUMBER" in cell_str):
                    contract_col = col
                elif "STAGE" in cell_str:
                    stage_col = col
        
        if not contract_col or not stage_col:
            wb.close()
            return None
        
        # Search for contract number
        contract_key = str(contract_no).strip()
        for row in range(header_row + 1, sheet.max_row + 1):
            row_contract = sheet.cell(row=row, column=contract_col).value
            
            if row_contract and str(row_contract).strip() == contract_key:
                stage = sheet.cell(row=row, column=stage_col).value
                wb.close()
                
                if stage is not None:
                    try:
                        return int(float(stage))
                    except (ValueError, TypeError):
                        pass
                break
        
        wb.close()
    except Exception as e:
        print(f"✗ Error reading {summary_file.name}: {e}")
    
    return None

def update_grant_amounts_sheet(c7_report_path: Path, info_request_path: Path, quarter: int = None, fy_year: int = None):
    """
    Update the Grant Amounts sheet in C7 report with Contract No and Net Monthly Disbursement 
    from Information Request report, and Stage from Summary file using direct lookup.
    Works directly with .xlsb files using Excel COM automation (background process).
    
    Args:
        c7_report_path: Path to C7 Stage Movement xlsb file
        info_request_path: Path to Information Request xlsx file
        quarter: Quarter number (1-4) - optional, needed for Stage lookup
        fy_year: Financial year - optional, needed for Stage lookup
    """
    print("\n" + "="*60)
    print("Updating Grant Amounts Sheet")
    print("="*60)
    
    # Validate files exist
    if not c7_report_path.exists():
        print(f"✗ C7 report not found: {c7_report_path}")
        return False
    
    if not info_request_path.exists():
        print(f"✗ Information Request report not found: {info_request_path}")
        return False
    
    print(f"✓ C7 Report: {c7_report_path.name}")
    print(f"✓ Information Request Report: {info_request_path.name}")
    
    excel = None
    workbook = None
    
    try:
        # Read Contract No and Net Monthly Disbursement from Information Request report - Disbursements sheet
        # Headers are in row 5, so we skip the first 4 rows
        print("\nReading Contract No and Net Monthly Disbursement from Information Request report (Disbursements sheet)...")
        info_df = pd.read_excel(info_request_path, sheet_name='Disbursements', header=4)
        
        # Check if required columns exist
        if 'Contract No' not in info_df.columns:
            print("✗ Contract No column not found in Disbursements sheet")
            print(f"Available columns: {list(info_df.columns)}")
            return False
        
        if 'Net Monthly Disbursement (LKR)' not in info_df.columns:
            print("✗ Net Monthly Disbursement (LKR) column not found in Disbursements sheet")
            print(f"Available columns: {list(info_df.columns)}")
            return False
        
        # Filter out "Margin trading" entries and remove NaN from Contract No
        filtered_df = info_df[
            (info_df['Contract No'].notna()) & 
            (info_df['Contract No'].astype(str).str.lower() != 'margin trading')
        ].copy()
        
        # Get Contract No and corresponding Net Monthly Disbursement values
        grant_data = filtered_df[['Contract No', 'Net Monthly Disbursement (LKR)']].values.tolist()
        print(f"✓ Found {len(grant_data)} Contract No entries with disbursement amounts (excluding 'Margin trading')")
        
        if len(grant_data) == 0:
            print("⚠ No data to copy")
            return True
        
        # Find Summary file for Stage lookup if quarter and fy_year provided
        summary_path = None
        if quarter and fy_year:
            script_dir = Path(__file__).resolve().parent
            project_root = script_dir.parent
            portfolios_dir = project_root / "working" / "quarterly" / "Portfolios"  # Fixed path
            
            print(f"\nSearching for Summary file in: {portfolios_dir}")
            summary_path = get_portfolio_summary_file(portfolios_dir, quarter, fy_year)
            
            if summary_path:
                print(f"✓ Found Summary file: {summary_path.name}")
            else:
                print("⚠ Summary file not found. Stage column will be empty.")
        
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Create Excel application object (completely hidden - no UI)
        print(f"\nOpening C7 workbook in background: {c7_report_path.name}...")
        excel = win32com.client.DispatchEx("Excel.Application")  # DispatchEx for better server isolation
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False  # Disable screen updates for better performance
        excel.Interactive = False  # Disable user interaction
        excel.EnableEvents = False  # Disable events
        
        # Open the xlsb file
        workbook = excel.Workbooks.Open(str(c7_report_path.absolute()), ReadOnly=False, UpdateLinks=False)
        
        # Check if Grant Amounts sheet exists, create if not
        sheet_exists = False
        for sheet in workbook.Sheets:
            if sheet.Name == 'Grant Amounts':
                ws = sheet
                sheet_exists = True
                print("✓ Found existing 'Grant Amounts' sheet")
                break
        
        if not sheet_exists:
            print("Creating new 'Grant Amounts' sheet...")
            ws = workbook.Sheets.Add()
            ws.Name = 'Grant Amounts'
            # Add headers
            ws.Cells(1, 1).Value = 'New Contract'
            ws.Cells(1, 2).Value = 'Grant Amount'
            ws.Cells(1, 3).Value = 'Stage'
            start_row = 2
        else:
            # Find the last row with data in column A
            start_row = 2
            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 is xlUp constant
            
            # If there's data, start after the last row
            if last_row >= 2 and ws.Cells(last_row, 1).Value is not None:
                start_row = last_row + 1
            elif last_row < 2:
                start_row = 2
            else:
                # Find first empty row starting from row 2
                for row in range(2, last_row + 2):
                    if ws.Cells(row, 1).Value is None or ws.Cells(row, 1).Value == "":
                        start_row = row
                        break
            
            # Ensure headers exist
            if ws.Cells(1, 1).Value != 'New Contract':
                ws.Cells(1, 1).Value = 'New Contract'
            if ws.Cells(1, 2).Value != 'Grant Amount':
                ws.Cells(1, 2).Value = 'Grant Amount'
            if ws.Cells(1, 3).Value != 'Stage':
                ws.Cells(1, 3).Value = 'Stage'
        
        print(f"\nAppending {len(grant_data)} entries starting from row {start_row}...")
        
        # Determine which function to use for stage lookup
        stage_lookup_func = None
        if summary_path:
            if summary_path.suffix.lower() == '.xlsb':
                stage_lookup_func = lambda contract: load_stage_from_summary_xlsb(summary_path, contract)
            else:
                stage_lookup_func = lambda contract: load_stage_from_summary_xlsx(summary_path, contract)
            print("Using direct stage lookup from Summary file...")
        
        # Append new data
        for idx, (contract_no, grant_amount) in enumerate(grant_data):
            row_num = start_row + idx
            ws.Cells(row_num, 1).Value = str(contract_no)
            
            # Handle grant_amount - could be None/NaN
            if pd.notna(grant_amount):
                ws.Cells(row_num, 2).Value = grant_amount
            else:
                ws.Cells(row_num, 2).Value = 0
            
            # Direct stage lookup if Summary file is available
            if stage_lookup_func:
                stage_value = stage_lookup_func(contract_no)
                if stage_value is not None:
                    ws.Cells(row_num, 3).Value = stage_value
                else:
                    ws.Cells(row_num, 3).Value = ""
            else:
                ws.Cells(row_num, 3).Value = ""
        
        # Save workbook
        print(f"Saving changes to {c7_report_path.name}...")
        workbook.Save()
        
        print(f"✓ Successfully updated Grant Amounts sheet")
        print(f"  - Added rows: {start_row} to {start_row + len(grant_data) - 1}")
        print(f"  - Stage column: {'Populated via direct lookup' if summary_path else 'Empty (Summary file not found)'}")
        
        return True
        
    except Exception as e:
        print(f"✗ Error updating Grant Amounts sheet: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        # Clean up Excel objects properly for server environment
        try:
            if workbook:
                workbook.Close(SaveChanges=True)
            if excel:
                excel.Quit()
            # Release COM objects
            if workbook:
                del workbook
            if excel:
                del excel
            pythoncom.CoUninitialize()
        except Exception as cleanup_error:
            print(f"⚠ Warning during cleanup: {cleanup_error}")
