import pandas as pd
from pathlib import Path
import win32com.client
import pythoncom

def update_closed_sheet(c7_report_path: Path, closure_report_path: Path, quarter_months: str):
    """
    Update the Closed sheet in C7 report with CONTRACT_NO from closure report.
    Works directly with .xlsb files using Excel COM automation.
    
    Args:
        c7_report_path: Path to C7 Stage Movement xlsb file
        closure_report_path: Path to Closure Category Wise Report xlsx file
        quarter_months: Quarter months in format "Apr-Jun"
    """
    print("\n" + "="*60)
    print("Updating Closed Sheet")
    print("="*60)
    
    # Validate files exist
    if not c7_report_path.exists():
        print(f"✗ C7 report not found: {c7_report_path}")
        return False
    
    if not closure_report_path.exists():
        print(f"✗ Closure report not found: {closure_report_path}")
        return False
    
    print(f"✓ C7 Report: {c7_report_path.name}")
    print(f"✓ Closure Report: {closure_report_path.name}")
    
    excel = None
    workbook = None
    
    try:
        # Read CONTRACT_NO from closure report
        print("\nReading CONTRACT_NO from closure report...")
        closure_df = pd.read_excel(closure_report_path, sheet_name=0)
        
        if 'CONTRACT_NO' not in closure_df.columns:
            print("✗ CONTRACT_NO column not found in closure report")
            return False
        
        # Get unique CONTRACT_NO values and remove NaN
        contract_numbers = closure_df['CONTRACT_NO'].dropna().unique().tolist()
        print(f"✓ Found {len(contract_numbers)} unique CONTRACT_NO entries")
        
        if len(contract_numbers) == 0:
            print("⚠ No CONTRACT_NO data to copy")
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
        
        # Check if Closed sheet exists, create if not
        sheet_exists = False
        for sheet in workbook.Sheets:
            if sheet.Name == 'Closed':
                ws = sheet
                sheet_exists = True
                print("✓ Found existing 'Closed' sheet")
                break
        
        if not sheet_exists:
            print("Creating new 'Closed' sheet...")
            ws = workbook.Sheets.Add()
            ws.Name = 'Closed'
            # Add headers
            ws.Cells(1, 1).Value = 'CONTRACT_NO'
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
            if ws.Cells(1, 1).Value != 'CONTRACT_NO':
                ws.Cells(1, 1).Value = 'CONTRACT_NO'
            if ws.Cells(1, 2).Value != 'Month':
                ws.Cells(1, 2).Value = 'Month'
        
        print(f"\nAppending {len(contract_numbers)} CONTRACT_NO entries starting from row {start_row}...")
        
        # Append new data
        for idx, contract_no in enumerate(contract_numbers):
            row_num = start_row + idx
            ws.Cells(row_num, 1).Value = str(contract_no)
            ws.Cells(row_num, 2).Value = quarter_months
        
        # Save workbook
        print(f"Saving changes to {c7_report_path.name}...")
        workbook.Save()
        
        print(f"✓ Successfully updated Closed sheet")
        print(f"  - Added rows: {start_row} to {start_row + len(contract_numbers) - 1}")
        print(f"  - Quarter: {quarter_months}")
        
        return True
        
    except Exception as e:
        print(f"✗ Error updating closed sheet: {e}")
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
