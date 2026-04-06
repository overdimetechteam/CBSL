import pandas as pd
from pathlib import Path
import win32com.client
import pythoncom
from datetime import datetime

def find_write_off_files(wo_dir: Path, quarter: int, fy_year: int) -> list:
    """
    Find Write Off Working files for the quarter months.
    
    Args:
        wo_dir: Path to WO directory
        quarter: Quarter number (1-4)
        fy_year: Financial year
    
    Returns:
        List of tuples (file_path, month_label)
    """
    # Determine quarter months
    quarter_months_map = {
        1: [(4, fy_year, "Apr"), (5, fy_year, "May"), (6, fy_year, "Jun")],
        2: [(7, fy_year, "Jul"), (8, fy_year, "Aug"), (9, fy_year, "Sep")],
        3: [(10, fy_year, "Oct"), (11, fy_year, "Nov"), (12, fy_year, "Dec")],
        4: [(1, fy_year + 1, "Jan"), (2, fy_year + 1, "Feb"), (3, fy_year + 1, "Mar")]
    }
    
    quarter_months = quarter_months_map[quarter]
    
    if not wo_dir.exists():
        print(f"✗ WO directory not found: {wo_dir}")
        return []
    
    matching_files = []
    
    for month_num, year, month_name in quarter_months:
        # Pattern: 01. Write Off Working - MMM YYYY - DR Analysis.xlsx
        pattern = f"01. Write Off Working - {month_name} {year} - DR Analysis.xlsx"
        file_path = wo_dir / pattern
        
        if file_path.exists():
            # Create date in format MMM-YY (e.g., Apr-25)
            year_short = str(year)[-2:]
            month_label = f"{month_name}-{year_short}"
            matching_files.append((file_path, month_label))
            print(f"  ✓ Found: {file_path.name} -> {month_label}")
        else:
            print(f"  ⚠ Not found: {pattern}")
    
    return matching_files

def read_write_off_data(file_path: Path) -> list:
    """
    Read Original Contract No, Trans Description, and Amount from Write Off file.
    
    Args:
        file_path: Path to Write Off Working file
    
    Returns:
        List of tuples (contract_no, trans_description, amount)
    """
    data = []
    
    try:
        # Read from 0130-GL_Working sheet, headers in row 3 (index 2)
        print(f"\n  Reading: {file_path.name}")
        df = pd.read_excel(file_path, sheet_name='0130-GL_Working', header=2)
        
        # Check if required columns exist
        required_columns = ['Original Contract No', 'Trans Description', 'Amount']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"  ✗ Missing columns: {missing_columns}")
            print(f"  Available columns: {list(df.columns)}")
            return data
        
        # Extract data, filtering out NaN values
        for _, row in df.iterrows():
            contract_no = row['Original Contract No']
            trans_desc = row['Trans Description']
            amount = row['Amount']
            
            # Only include rows with valid contract number
            if pd.notna(contract_no):
                contract_str = str(contract_no).strip()
                trans_str = str(trans_desc).strip() if pd.notna(trans_desc) else ""
                amount_val = float(amount) if pd.notna(amount) else 0.0
                
                data.append((contract_str, trans_str, amount_val))
        
        print(f"  ✓ Extracted {len(data)} records")
        
    except Exception as e:
        print(f"  ✗ Error reading {file_path.name}: {e}")
    
    return data

def load_movement_data(workbook) -> dict:
    """
    Load contract to AG WO mapping from Movement sheet.
    
    Args:
        workbook: Excel workbook object
    
    Returns:
        Dictionary mapping contract_no to AG WO value
    """
    movement_map = {}
    
    try:
        # Check if Movement sheet exists
        movement_sheet = None
        for sheet in workbook.Sheets:
            if sheet.Name == 'Movement':
                movement_sheet = sheet
                break
        
        if not movement_sheet:
            print("⚠ Movement sheet not found in workbook")
            return movement_map
        
        print("  Reading Movement sheet data for AG WO lookup...")
        
        # Read from row 3 onwards (data starts at row 3, column A is contract, column K is AG WO)
        # Column A = 1, Column K = 11
        for row in range(3, 57746 + 1):  # Up to row 57746 as per VLOOKUP
            contract_no = movement_sheet.Cells(row, 1).Value  # Column A
            ag_wo_value = movement_sheet.Cells(row, 11).Value  # Column K
            
            if contract_no:
                contract_key = str(contract_no).strip()
                try:
                    ag_wo = float(ag_wo_value) if ag_wo_value is not None else 0.0
                except (ValueError, TypeError):
                    ag_wo = 0.0
                
                movement_map[contract_key] = ag_wo
        
        print(f"  ✓ Loaded {len(movement_map)} contracts from Movement sheet")
        
    except Exception as e:
        print(f"  ✗ Error reading Movement sheet: {e}")
    
    return movement_map

def update_wo_sheet(c7_report_path: Path, quarter: int, fy_year: int):
    """
    Update the WO (Write Off) sheet in C7 report with data from Write Off Working files.
    Also update Disposal Loss sheet with filtered Disposal Loss and Event Based entries.
    Works directly with .xlsb files using Excel COM automation (background process).
    
    Args:
        c7_report_path: Path to C7 Stage Movement xlsb file
        quarter: Quarter number (1-4)
        fy_year: Financial year
    """
    print("\n" + "="*60)
    print("Updating WO (Write Off) Sheet")
    print("="*60)
    
    # Validate C7 file exists
    if not c7_report_path.exists():
        print(f"✗ C7 report not found: {c7_report_path}")
        return False
    
    print(f"✓ C7 Report: {c7_report_path.name}")
    
    # Setup paths - look inside the WO subfolder next to the C7 report
    # The C7 report is in working/NBD_QF_23_C7/{date}/
    # WO files are stored under working/NBD_QF_23_C7/{date}/WO/
    wo_dir = c7_report_path.parent / "WO"

    print(f"\nSearching for Write Off files in: {wo_dir}")
    
    # Find Write Off files for the quarter
    wo_files = find_write_off_files(wo_dir, quarter, fy_year)
    
    if not wo_files:
        print("\n⚠ No Write Off files found for the quarter")
        return False
    
    print(f"\n✓ Found {len(wo_files)} Write Off file(s)")
    
    # Collect all data from Write Off files
    all_wo_data = []
    disposal_loss_data = []  # For Disposal Loss sheet
    
    for file_path, month_label in wo_files:
        wo_data = read_write_off_data(file_path)
        
        # Add month_label to each record and separate Disposal Loss entries
        for contract_no, trans_desc, amount in wo_data:
            all_wo_data.append((contract_no, amount, trans_desc, month_label))
            
            # Filter for Disposal Loss and Event Based
            trans_desc_upper = trans_desc.upper()
            if 'DISPOSAL LOSS' in trans_desc_upper or 'EVENT BASED' in trans_desc_upper:
                disposal_loss_data.append((contract_no, amount, trans_desc))
    
    if not all_wo_data:
        print("\n⚠ No data extracted from Write Off files")
        return False
    
    print(f"\n✓ Total records to add to WO sheet: {len(all_wo_data)}")
    print(f"✓ Total Disposal Loss/Event Based records: {len(disposal_loss_data)}")
    
    excel = None
    workbook = None
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Create Excel application object (completely hidden - no UI)
        print(f"\nOpening C7 workbook in background: {c7_report_path.name}...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.Interactive = False
        excel.EnableEvents = False
        
        # Open the xlsb file
        workbook = excel.Workbooks.Open(str(c7_report_path.absolute()), ReadOnly=False, UpdateLinks=False)
        
        # Load Movement sheet data for AG WO lookup
        print("\nLoading Movement sheet data...")
        movement_map = load_movement_data(workbook)
        
        # ===== Update WO Sheet =====
        print("\n--- Updating WO Sheet ---")
        sheet_exists = False
        for sheet in workbook.Sheets:
            if sheet.Name == 'WO':
                ws = sheet
                sheet_exists = True
                print("✓ Found existing 'WO' sheet")
                break
        
        if not sheet_exists:
            print("Creating new 'WO' sheet...")
            ws = workbook.Sheets.Add()
            ws.Name = 'WO'
            # Add headers in row 1
            ws.Cells(1, 1).Value = 'Original Contract No'
            ws.Cells(1, 2).Value = 'Amount'
            ws.Cells(1, 3).Value = 'Final Type'
            ws.Cells(1, 4).Value = 'Month'
            ws.Cells(1, 5).Value = 'Pro'
            start_row = 2
        else:
            # Find the last row with data in column A
            start_row = 2
            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            
            if last_row >= 2 and ws.Cells(last_row, 1).Value is not None:
                start_row = last_row + 1
            elif last_row < 2:
                start_row = 2
            else:
                for row in range(2, last_row + 2):
                    if ws.Cells(row, 1).Value is None or ws.Cells(row, 1).Value == "":
                        start_row = row
                        break
            
            # Ensure headers exist
            if ws.Cells(1, 1).Value != 'Original Contract No':
                ws.Cells(1, 1).Value = 'Original Contract No'
            if ws.Cells(1, 2).Value != 'Amount':
                ws.Cells(1, 2).Value = 'Amount'
            if ws.Cells(1, 3).Value != 'Final Type':
                ws.Cells(1, 3).Value = 'Final Type'
            if ws.Cells(1, 4).Value != 'Month':
                ws.Cells(1, 4).Value = 'Month'
            if ws.Cells(1, 5).Value != 'Pro':
                ws.Cells(1, 5).Value = 'Pro'
        
        print(f"Appending {len(all_wo_data)} entries to WO sheet starting from row {start_row}...")
        
        # Append new data to WO sheet
        for idx, (contract_no, amount, final_type, month_label) in enumerate(all_wo_data):
            row_num = start_row + idx
            ws.Cells(row_num, 1).Value = contract_no
            ws.Cells(row_num, 2).Value = amount
            ws.Cells(row_num, 3).Value = final_type
            ws.Cells(row_num, 4).Value = month_label
            pro_formula = f"=MID(A{row_num},3,2)"
            ws.Cells(row_num, 5).Formula = pro_formula
        
        print(f"✓ WO sheet updated ({len(all_wo_data)} records)")
        
        # ===== Update Disposal Loss Sheet =====
        if disposal_loss_data:
            print("\n--- Updating Disposal Loss Sheet ---")
            
            # Check if Disposal Loss sheet exists
            dl_sheet_exists = False
            for sheet in workbook.Sheets:
                if sheet.Name == 'Disposal Loss':
                    dl_ws = sheet
                    dl_sheet_exists = True
                    print("✓ Found existing 'Disposal Loss' sheet")
                    break
            
            if not dl_sheet_exists:
                print("Creating new 'Disposal Loss' sheet...")
                dl_ws = workbook.Sheets.Add()
                dl_ws.Name = 'Disposal Loss'
                
                # Headers in row 2 (column J starts at 10)
                # J=10: Contact, K=11: Amount, L=12: Blank, M=13: Pro AG WO, N=14: Pro
                dl_ws.Cells(2, 10).Value = 'Contact'
                dl_ws.Cells(2, 11).Value = 'Amount'
                dl_ws.Cells(2, 12).Value = 'Blank'
                dl_ws.Cells(2, 13).Value = 'Pro AG WO'
                dl_ws.Cells(2, 14).Value = 'Pro'
                dl_start_row = 3
            else:
                # Find the last row with data in column J (column 10)
                dl_start_row = 3
                last_row = dl_ws.Cells(dl_ws.Rows.Count, 10).End(-4162).Row
                
                if last_row >= 3 and dl_ws.Cells(last_row, 10).Value is not None:
                    dl_start_row = last_row + 1
                elif last_row < 3:
                    dl_start_row = 3
                else:
                    for row in range(3, last_row + 2):
                        if dl_ws.Cells(row, 10).Value is None or dl_ws.Cells(row, 10).Value == "":
                            dl_start_row = row
                            break
                
                # Ensure headers exist in row 2
                if dl_ws.Cells(2, 10).Value != 'Contact':
                    dl_ws.Cells(2, 10).Value = 'Contact'
                if dl_ws.Cells(2, 11).Value != 'Amount':
                    dl_ws.Cells(2, 11).Value = 'Amount'
                if dl_ws.Cells(2, 12).Value != 'Blank':
                    dl_ws.Cells(2, 12).Value = 'Blank'
                if dl_ws.Cells(2, 13).Value != 'Pro AG WO':
                    dl_ws.Cells(2, 13).Value = 'Pro AG WO'
                if dl_ws.Cells(2, 14).Value != 'Pro':
                    dl_ws.Cells(2, 14).Value = 'Pro'
            
            print(f"Appending {len(disposal_loss_data)} Disposal Loss entries starting from row {dl_start_row}...")
            
            # Append Disposal Loss data with direct lookup
            for idx, (contract_no, amount, trans_desc) in enumerate(disposal_loss_data):
                row_num = dl_start_row + idx
                
                # Column J (10): Contact (Contract No)
                dl_ws.Cells(row_num, 10).Value = contract_no
                
                # Column K (11): Amount
                dl_ws.Cells(row_num, 11).Value = amount
                
                # Column L (12): Blank (leave blank)
                dl_ws.Cells(row_num, 12).Value = ""
                
                # Column M (13): Pro AG WO (Direct lookup from movement_map)
                contract_key = str(contract_no).strip()
                ag_wo_value = movement_map.get(contract_key, 0.0)
                
                # Try with leading zeros stripped if not found
                if ag_wo_value == 0.0 and contract_key:
                    try_key = contract_key.lstrip("0")
                    if try_key:
                        ag_wo_value = movement_map.get(try_key, 0.0)
                
                dl_ws.Cells(row_num, 13).Value = ag_wo_value
                
                # Column N (14): Pro (extract from contract number using MID)
                pro_formula = f"=MID(J{row_num},3,2)"
                dl_ws.Cells(row_num, 14).Formula = pro_formula
            
            print(f"✓ Disposal Loss sheet updated ({len(disposal_loss_data)} records)")
            if movement_map:
                print(f"  ✓ Pro AG WO values populated via direct lookup from Movement sheet")
            else:
                print(f"  ⚠ Pro AG WO values set to 0 (Movement sheet data not available)")
        
        # Save workbook
        print(f"\nSaving changes to {c7_report_path.name}...")
        workbook.Save()
        
        print(f"\n✓ Successfully updated WO and Disposal Loss sheets")
        print(f"  - WO sheet: {len(all_wo_data)} records")
        print(f"  - Disposal Loss sheet: {len(disposal_loss_data)} records")
        
        return True
        
    except Exception as e:
        print(f"✗ Error updating sheets: {e}")
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
