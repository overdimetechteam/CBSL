import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import pyxlsb
from pyxlsb import open_workbook as open_xlsb
import re
from pathlib import Path
import win32com.client as win32
import pythoncom

def extract_quarter_from_filename(filename):
    """Extract quarter period from C7 report filename"""
    # Example: "C7- Report for Stage - Movement Apr-Jun2025.xlsb"
    match = re.search(r'(\w+)-(\w+)(\d{4})', filename)
    if match:
        start_month = match.group(1)
        end_month = match.group(2)
        year = match.group(3)
        return f"{start_month}-{end_month} {year}"
    return None

def read_unwind_values(unwind_file_path):
    """Read Stage values from Unwind of Discount file"""
    try:
        wb = load_workbook(unwind_file_path, data_only=True)
        sheet = wb['Sheet1']
        
        values = {
            'Stage 1': sheet['B4'].value,
            'Stage 2': sheet['B5'].value,
            'Stage 3': sheet['B6'].value,
            'Grand Total': sheet['B7'].value
        }
        
        wb.close()
        return values
    except Exception as e:
        print(f"✗ Error reading Unwind file: {e}")
        return None

def write_to_c7_report(c7_file_path, values):
    """Write values to C7 report xlsb file"""
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Open Excel application
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open the xlsb file
        workbook = excel.Workbooks.Open(str(Path(c7_file_path).resolve()))
        
        # Access the target sheet
        target_sheet_name = 'Unwind of Dicount-2425'
        sheet_found = False
        
        for sheet in workbook.Sheets:
            if sheet.Name == target_sheet_name:
                worksheet = sheet
                sheet_found = True
                break
        
        if not sheet_found:
            print(f"✗ Sheet '{target_sheet_name}' not found!")
            workbook.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()
            return False
        
        # Clear existing values first (optional, but ensures clean replacement)
        worksheet.Range("C5:F5").ClearContents()
        
        # Write new values to row 5, columns C, D, E, F
        worksheet.Range("C5").Value = values.get('Stage 1', 0)
        worksheet.Range("D5").Value = values.get('Stage 2', 0)
        worksheet.Range("E5").Value = values.get('Stage 3', 0)
        worksheet.Range("F5").Value = values.get('Grand Total', 0)
        
        print(f"  Stage 1: {values.get('Stage 1', 0)} → Cell C5")
        print(f"  Stage 2: {values.get('Stage 2', 0)} → Cell D5")
        print(f"  Stage 3: {values.get('Stage 3', 0)} → Cell E5")
        print(f"  Grand Total: {values.get('Grand Total', 0)} → Cell F5")
        
        # Save and close
        workbook.Save()
        workbook.Close(SaveChanges=False)
        excel.Quit()
        
        # Uninitialize COM
        pythoncom.CoUninitialize()
        
        print(f"✓ Successfully updated C7 report: {Path(c7_file_path).name}")
        
        return True
        
    except Exception as e:
        print(f"✗ Error writing to C7 report: {e}")
        try:
            excel.Quit()
            pythoncom.CoUninitialize()
        except:
            pass
        return False

def process_unwind_to_c7(c7_file_path, unwind_file_path):
    """
    Copy Unwind values to C7 report
    
    Args:
        c7_file_path: Path to C7 report file (.xlsb)
        unwind_file_path: Path to Unwind of Discount file (.xlsx)
    """
    print(f"  C7 report: {Path(c7_file_path).name}")
    print(f"  Unwind file: {Path(unwind_file_path).name}")
    
    print("\nReading values from Unwind of Discount file...")
    values = read_unwind_values(unwind_file_path)
    
    if values:
        print(f"✓ Values read: {values}")
        
        print("\nWriting values to C7 report...")
        success = write_to_c7_report(c7_file_path, values)
        
        if success:
            print("\n✓ Process completed successfully!")
            return True
        else:
            print("\n✗ Process completed with errors.")
            return False
    else:
        print("✗ Failed to read values from Unwind file.")
        return False

def main():
    """Main function to orchestrate the data transfer"""
    # Define output directory using relative path
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent
    output_dir = project_root / "outputs" / "quarterly"
    
    if not output_dir.exists():
        print(f"Output directory not found: {output_dir}")
        return
    
    # Find files
    print("Searching for files in output directory...")
    c7_file, unwind_file = find_files_in_output_dir(str(output_dir), None)
    
    if not c7_file or not unwind_file:
        print(f"Required files not found in {output_dir}")
        print(f"C7 file: {c7_file}")
        print(f"Unwind file: {unwind_file}")
        return
    
    print(f"Found C7 report: {os.path.basename(c7_file)}")
    print(f"Found Unwind file: {os.path.basename(unwind_file)}")
    
    # Extract quarter info
    quarter = extract_quarter_from_filename(os.path.basename(c7_file))
    print(f"Processing quarter: {quarter}")
    
    # Read values from Unwind file
    print("\nReading values from Unwind of Discount file...")
    values = read_unwind_values(unwind_file)
    
    if values:
        print(f"Values read: {values}")
        
        # Write to C7 report
        print("\nWriting values to C7 report...")
        success = write_to_c7_report(c7_file, values)
        
        if success:
            print("\nProcess completed successfully!")
        else:
            print("\nProcess completed with errors.")
    else:
        print("Failed to read values from Unwind file.")

if __name__ == "__main__":
    # For standalone testing
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent
    output_dir = project_root / "outputs" / "quarterly"
    
    # Find files for testing
    c7_files = list(output_dir.glob("*C7- Report for Stage*.xlsb"))
    unwind_files = list(output_dir.glob("Unwind of Discount*.xlsx"))
    
    if c7_files and unwind_files:
        process_unwind_to_c7(str(c7_files[0]), str(unwind_files[0]))
    else:
        print("Files not found for testing")

