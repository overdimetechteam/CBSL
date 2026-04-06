import pandas as pd
from pathlib import Path
import win32com.client
import pythoncom

def update_new_activations_sheet(c7_report_path: Path, info_request_path: Path, quarter_months: str):
    """
    Update the New Activations sheet in C7 report with Contract No from Information Request report.
    Works directly with .xlsb files using Excel COM automation.
    
    Args:
        c7_report_path: Path to C7 Stage Movement xlsb file
        info_request_path: Path to Information Request xlsx file
        quarter_months: Quarter months in format "Apr-Jun"
    """
    print("\n" + "="*60)
    print("Updating New Activations Sheet")
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
        # Read Contract No from Information Request report - Disbursements sheet
        # Headers are in row 5, so we skip the first 4 rows
        print("\nReading Contract No from Information Request report (Disbursements sheet)...")
        info_df = pd.read_excel(info_request_path, sheet_name='Disbursements', header=4)
        
        if 'Contract No' not in info_df.columns:
            print("✗ Contract No column not found in Disbursements sheet")
            print(f"Available columns: {list(info_df.columns)}")
            return False
        
        # Filter out "Margin trading" entries and remove NaN
        filtered_df = info_df[
            (info_df['Contract No'].notna()) & 
            (info_df['Contract No'].astype(str).str.lower() != 'margin trading')
        ]
        
        # Get unique Contract No values
        contract_numbers = filtered_df['Contract No'].unique().tolist()
        print(f"✓ Found {len(contract_numbers)} unique Contract No entries (excluding 'Margin trading')")
        
        if len(contract_numbers) == 0:
            print("⚠ No Contract No data to copy")
            return True
        
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
        
        # Check if New Activations sheet exists, create if not
        sheet_exists = False
        for sheet in workbook.Sheets:
            if sheet.Name == 'New Activations':
                ws = sheet
                sheet_exists = True
                print("✓ Found existing 'New Activations' sheet")
                break
        
        if not sheet_exists:
            print("Creating new 'New Activations' sheet...")
            ws = workbook.Sheets.Add()
            ws.Name = 'New Activations'
            # Add headers
            ws.Cells(1, 1).Value = 'Contract No'
            ws.Cells(1, 2).Value = 'Month'
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
            if ws.Cells(1, 1).Value != 'Contract No':
                ws.Cells(1, 1).Value = 'Contract No'
            if ws.Cells(1, 2).Value != 'Month':
                ws.Cells(1, 2).Value = 'Month'
        
        print(f"\nAppending {len(contract_numbers)} Contract No entries starting from row {start_row}...")
        
        # Append new data
        for idx, contract_no in enumerate(contract_numbers):
            row_num = start_row + idx
            ws.Cells(row_num, 1).Value = str(contract_no)
            ws.Cells(row_num, 2).Value = quarter_months
        
        # Save workbook
        print(f"Saving changes to {c7_report_path.name}...")
        workbook.Save()
        
        print(f"✓ Successfully updated New Activations sheet")
        print(f"  - Added rows: {start_row} to {start_row + len(contract_numbers) - 1}")
        print(f"  - Quarter: {quarter_months}")
        
        return True
        
    except Exception as e:
        print(f"✗ Error updating New Activations sheet: {e}")
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
            del workbook
            del excel
            pythoncom.CoUninitialize()
        except Exception as cleanup_error:
            print(f"⚠ Warning during cleanup: {cleanup_error}")
