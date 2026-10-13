"""
Portfolio Automation Script using xlwings
Automates the process of:
1. Clearing data range A4:AB in Portfolio sheet of xlsb file
2. Reading CSV file columns A-AA and pasting to Portfolio sheet starting at B4
3. Filling column A with today's date
4. Clearing range A2:E in 1.DPD sheet
5. Reading xlsx file and pasting first 5 columns to 1.DPD sheet starting at A2
6. Saving as "Updated ECL Portfolio" in xlsb format
"""
from calendar import calendar
from scipy.optimize import brentq
import xlwings as xw

import pandas as pd
import xlwings as xw
from datetime import datetime
import os
import glob
import shutil
import traceback
from typing import Optional, Tuple
import re
import sys
from typing import Union, List
from typing import Optional
from pathlib import Path
import threading
import time
import win32gui
import win32con
import calendar
from ol_adjustment import OLAdjustment

DOWNLOADS_FOLDER = os.environ.get("IFRS_DOWNLOADS_FOLDER", os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")), "Downloads"))
ROOT_FOLDER = os.environ.get("IFRS_ROOT_FOLDER", r"C:\CBSL\Script")
OUTPUTS_FOLDER = os.path.join(ROOT_FOLDER, "outputs")


class _DialogDismisser:
    """Background thread that auto-clicks OK on any matching Excel VBA MsgBox dialog."""

    def __init__(self, titles, interval=0.5):
        self._titles = [t.lower() for t in titles]
        self._interval = interval
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def __enter__(self):
        self._stop.clear()
        self._thread.start()
        return self

    def __exit__(self, *_):
        self._stop.set()
        self._thread.join(timeout=3)

    def _run(self):
        def _callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd).lower()
            if any(t in title for t in self._titles):
                # Find the OK button (IDOK=1) and click it
                btn = win32gui.FindWindowEx(hwnd, 0, "Button", None)
                while btn:
                    btn_text = win32gui.GetWindowText(btn).lower().strip()
                    if btn_text in ("ok", "&ok"):
                        win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
                        print(f"[AutoDismiss] Clicked OK on dialog: '{win32gui.GetWindowText(hwnd)}'")
                        return
                    btn = win32gui.FindWindowEx(hwnd, btn, "Button", None)
                # Fallback: send Enter to the dialog itself
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                print(f"[AutoDismiss] Sent Enter to dialog: '{win32gui.GetWindowText(hwnd)}'")

        while not self._stop.is_set():
            try:
                win32gui.EnumWindows(_callback, None)
            except Exception:
                pass
            time.sleep(self._interval)


class ExcelPortfolioAutomation:
    """
    Excel Portfolio Automation class using xlwings for .xlsb file manipulation
    """

    def __init__(self, workbook_path: str, visible: bool = True):
        """
        Initialize the automation class

        Args:
            workbook_path: Path to the .xlsb Excel file
            visible: Whether to show Excel application (default: False)
        """
        self.workbook_path = workbook_path
        self.visible = visible
        self.app = None
        self.workbook = None
        self.processing_date = None  # Store the date from process_portfolio_sheet
        self.month_end_date = None  # Store the calculated month-end date

    def __enter__(self):
        """Context manager entry - opens the workbook"""
        self.open_workbook()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - closes the workbook"""
        self.close_workbook()

    def open_workbook(self):
        """Open the Excel workbook"""
        print(f"Opening workbook: {self.workbook_path}")
        self.app = xw.App(visible=self.visible)
        self.app.api.DisplayAlerts = False
        self.app.api.AutomationSecurity = 1  # msoAutomationSecurityLow: allow all macros silently
        self.app.api.AskToUpdateLinks = False
        self.workbook = self.app.books.open(self.workbook_path, update_links=False)
        self.app.api.DisplayAlerts = False  # re-apply after open in case workbook reset it
        print(f"    Workbook opened successfully")

    def close_workbook(self, save: bool = False):
        """
        Close the Excel workbook

        Args:
            save: Whether to save before closing (default: False)
        """
        if self.workbook:
            if save:
                self.workbook.save()
                print(f"    Workbook saved")
            self.workbook.close()
            print(f"    Workbook closed")
        if self.app:
            self.app.quit()

    def clear_range(self, sheet_name: str, range_address: str) -> None:
        """
        Clear data in a specific range of a worksheet

        Args:
            sheet_name: Name of the worksheet
            range_address: Range address (e.g., 'A4:AB10000')
        """
        print(f"   Clearing range {range_address} in sheet '{sheet_name}'...")
        sheet = self.workbook.sheets[sheet_name]
        sheet.range(range_address).clear_contents()
        print(f"    Range cleared successfully")

    def clear_range_dynamic(self, sheet_name: str, start_cell: str, end_column: str) -> None:
        """
        Clear data from start_cell to the last row with data in specified columns

        Args:
            sheet_name: Name of the worksheet
            start_cell: Starting cell (e.g., 'A4', 'AX3', 'AB10')
            end_column: Last column letter (e.g., 'AB', 'AX')
        """
        sheet = self.workbook.sheets[sheet_name]

        # Extract column letters and row number
        start_column = ''.join(filter(str.isalpha, start_cell))
        start_row = int(''.join(filter(str.isdigit, start_cell)))

        print(f"   Analyzing range starting from {start_cell} in sheet '{sheet_name}'...")

        try:
            # Find the last row with data in the start column
            # Check from the row after start_cell
            check_cell = sheet.range(f'{start_column}{start_row + 1}')
            
            # Use end('down') to find last row
            last_row = check_cell.end('down').row
            
            # Handle case where end('down') goes to very last row (means no data)
            # Check if the last_row actually has data
            if last_row > 1048576 or last_row < start_row:
                # No data found, just clear the start row
                last_row = start_row

            # If there's data, clear it
            if last_row >= start_row:
                range_address = f"{start_cell}:{end_column}{last_row}"
                print(f"   Clearing range {range_address} in sheet '{sheet_name}'...")
                sheet.range(range_address).clear_contents()
                print(f"    Range cleared successfully ({last_row - start_row + 1} rows)")
            else:
                print(f"   No data to clear in sheet '{sheet_name}'")
                
        except Exception as e:
            print(f"   Warning: {e}")
            # Fallback: try to clear used range
            try:
                used_range = sheet.used_range
                last_row = used_range.last_cell.row
                
                if last_row >= start_row:
                    range_address = f"{start_cell}:{end_column}{last_row}"
                    print(f"   Using fallback method, clearing {range_address}...")
                    sheet.range(range_address).clear_contents()
                    print(f"    Range cleared successfully")
                else:
                    print(f"   No data to clear")
            except Exception as e2:
                print(f"    Could not clear range: {e2}")


    def read_csv_data(self, csv_path: str, num_columns: int = 27) -> Tuple[pd.DataFrame, int, int]:
        """
        Read data from CSV file

        Args:
            csv_path: Path to the CSV file
            num_columns: Number of columns to read (default: 27 for A-AA)

        Returns:
            Tuple of (DataFrame, num_rows, num_cols)
        """
        print(f"   Reading CSV file: {csv_path}")
        df = pd.read_csv(csv_path, header=0)

        # Get specified number of columns
        if len(df.columns) >= num_columns:
            df = df.iloc[:, 0:num_columns]
        else:
            print(f"   Note: CSV has only {len(df.columns)} columns (expected {num_columns})")

        num_rows = len(df)
        num_cols = len(df.columns)
        print(f"    Successfully read {num_rows} rows and {num_cols} columns")

        return df, num_rows, num_cols

    def normalize_date(self, date_str):
        formats = [
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d.%m.%Y",
            "%d/%m/%y",
            "%Y/%m/%d",
            "%d %b %Y",
            "%d %B %Y"
        ]
        
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt).strftime("%d.%m.%Y")
            except ValueError:
                continue
        
        raise ValueError(f"Unsupported date format: {date_str}")

    def read_excel_data(self, excel_path: str, num_columns: int = 5, skip_rows: int = 0, sheet_name: str = None, engine='openpyxl') -> Tuple[pd.DataFrame, int, int]:
        """
        Read data from Excel file

        Args:
            excel_path: Path to the Excel file
            num_columns: Number of columns to read (default: 5 for A-E)
            skip_rows: Number of rows to skip (default: 0)
            sheet_name: Sheet name to read (default: None = first sheet or 0)
            engine: Engine to use for reading (default: openpyxl)

        Returns:
            Tuple of (DataFrame, num_rows, num_cols)
        """
        print(f"   Reading Excel file: {excel_path}")

        # If sheet_name is None, default to first sheet (index 0) instead of all sheets
        if sheet_name is None:
            sheet_name = 0

        df = pd.read_excel(excel_path, header=0, skiprows=skip_rows, sheet_name=sheet_name, engine=engine)

        # Get specified number of columns
        if len(df.columns) >= num_columns:
            df = df.iloc[:, 0:num_columns]
        else:
            print(f"   Note: Excel has only {len(df.columns)} columns (expected {num_columns})")

        num_rows = len(df)
        num_cols = len(df.columns)
        print(f"    Successfully read {num_rows} rows and {num_cols} columns")

        return df, num_rows, num_cols

    def write_data_to_range(self, sheet_name: str, start_cell: str, data: list) -> None:
        """
        Write data to a worksheet starting at specified cell

        Args:
            sheet_name: Name of the worksheet
            start_cell: Starting cell (e.g., 'A4', 'B4')
            data: List of lists containing data to write
        """
        if not data:
            print(f"   No data to write to sheet '{sheet_name}'")
            return

        sheet = self.workbook.sheets[sheet_name]
        num_rows = len(data)
        num_cols = len(data[0]) if data else 0

        print(f"   Writing {num_rows} rows x {num_cols} columns to sheet '{sheet_name}' starting at {start_cell}...")

        # Write data to sheet
        sheet.range(start_cell).value = data

        print(f"    Data written successfully")

    def fill_column_with_value(self, sheet_name: str, column: str, start_row: int,
                                end_row: int, value: str) -> None:
        """
        Fill a column with a specific value

        Args:
            sheet_name: Name of the worksheet
            column: Column letter (e.g., 'A')
            start_row: Starting row number
            end_row: Ending row number
            value: Value to fill
        """
        sheet = self.workbook.sheets[sheet_name]
        range_address = f"{column}{start_row}:{column}{end_row}"

        print(f"   Filling {range_address} with value: {value}")
        sheet.range(range_address).value = value
        print(f"    Column filled successfully")

    def copy_formulas_to_range(self, sheet_name: str, source_range: str,
                                target_start_row: int, target_end_row: int) -> None:
        """
        Copy formulas from the last row with data in the column range to target rows

        Args:
            sheet_name: Name of the worksheet
            source_range: Column range to search (e.g., 'F2:O2' or 'F:O')
            target_start_row: Starting row for pasting formulas
            target_end_row: Ending row for pasting formulas
        """
        sheet = self.workbook.sheets[sheet_name]

        # Parse source range to get column range
        parts = source_range.split(':')
        if len(parts) != 2:
            raise ValueError(f"Invalid range format: {source_range}. Expected format like 'F2:O2'")

        start_cell = parts[0]
        end_cell = parts[1]

        # Extract column letters from cells
        start_col = ''.join([c for c in start_cell if c.isalpha()])
        end_col = ''.join([c for c in end_cell if c.isalpha()])

        # Find the last row with data in the start column of the range
        last_row_with_data = sheet.range(f'{start_col}5').end('down').row

        # Build the actual source range from the last row with data
        actual_source_range = f"{start_col}{last_row_with_data}:{end_col}{last_row_with_data}"

        print(f"   Found last row with data at row {last_row_with_data}")
        print(f"   Copying formulas from {actual_source_range} to rows {target_start_row}:{target_end_row}")

        # Copy the source range from the last row with data
        source_range_obj = sheet.range(actual_source_range)
        target_start_row = max(target_start_row, last_row_with_data + 1)  # Ensure we start below the source row
        # Paste to entire target range in a single operation
        if target_start_row >= target_end_row:
            print("   No rows to copy formulas to (target start row is after target end row)")
        else:
            # Define the entire target range at once
            target_range = f"{start_col}{target_start_row}:{end_col}{target_end_row}"
            # Copy the source to the entire target range in one operation
            source_range_obj.copy(sheet.range(target_range))

            print(f"    Formulas copied successfully to {target_end_row - target_start_row + 1} rows")

    def save_as(self, output_path: str) -> None:
        """
        Save workbook as a new file

        Args:
            output_path: Path for the output file
        """
        print(f"   Saving workbook as: {output_path}")

        # Get full path
        full_path = os.path.abspath(output_path)

        # Suppress any format-change dialogs during save
        self.app.api.DisplayAlerts = False
        self.workbook.save(full_path)
        self.app.api.DisplayAlerts = False  # re-apply in case save reset it
        print(f"    Workbook saved successfully")

    def copy_formula_and_paste_values(self, sheet_name: str, formula_cell: str, 
                                   target_start_cell: str, find_last_row_column: str = 'A') -> None:
        """
        Copy formula from a single cell and paste to a range, then convert to values
        Uses Excel's native PasteSpecial for better performance
        
        Args:
            sheet_name: Name of the worksheet
            formula_cell: Source cell with formula (e.g., 'A1')
            target_start_cell: Starting cell for pasting (e.g., 'A4')
            find_last_row_column: Column to check for last row with data (default: 'A')
        """
        sheet = self.workbook.sheets[sheet_name]
        
        # Find the last row with data in the specified column
        last_row = sheet.range(f'{find_last_row_column}4').end('down').row
        
        # Extract column letter from target_start_cell
        target_col = ''.join([c for c in target_start_cell if c.isalpha()])
        target_row = int(''.join([c for c in target_start_cell if c.isdigit()]))
        
        # Build target range
        target_range = f"{target_col}{target_row}:{target_col}{last_row}"
        
        print(f"   Copying formula from {formula_cell} to {target_range} in sheet '{sheet_name}'...")
        
        # Copy formula to range
        formula_source = sheet.range(formula_cell)
        target_range_obj = sheet.range(target_range)
        formula_source.copy(target_range_obj)
        
        print(f"   Converting formulas to values in {target_range}...")
        
        # Convert to values using Excel's native method
        # Copy the range
        target_range_obj.api.Copy()
        
        # Paste as values using Excel constant (xlPasteValues = -4163)
        target_range_obj.api.PasteSpecial(Paste=-4163)
        
        # Clear clipboard
        self.app.api.CutCopyMode = False
        
        print(f"    Formula copied and converted to values successfully")

    def copy_formula_and_paste_value_single(self, sheet_name: str,
                                        formula_cell: str,
                                        target_cell: str) -> None:
        """
        Copy formula from one cell and paste as value into another single cell.
        """

        sheet = self.workbook.sheets[sheet_name]

        print(f"Copying formula from {formula_cell} to {target_cell} in sheet '{sheet_name}'...")

        # Source and target
        source = sheet.range(formula_cell)
        target = sheet.range(target_cell)

        # Convert to value
        source.api.Copy()
        target.api.PasteSpecial(Paste=-4163)  # xlPasteValues

        # Clear clipboard
        self.app.api.CutCopyMode = False

        print("Formula copied and converted to value successfully")

    def update_pivot_source_and_refresh(self,
                                       sheet_name: str,
                                       pivot_table_name: str,
                                       data_sheet_name: str,
                                       start_cell: str = 'A1',
                                       end_column: str = None,
                                       include_headers: bool = True) -> None:
        """
        Dynamically detect data range, update pivot table source, and refresh it

        Args:
            sheet_name: Name of the worksheet containing the pivot table
            pivot_table_name: Name of the pivot table to update
            data_sheet_name: Name of the worksheet containing the source data
            start_cell: Starting cell of the data range (default: 'A1')
            end_column: Last column letter (e.g., 'Z'). If None, will auto-detect
            include_headers: Whether to include header row in the range (default: True)
        """
        try:
            print(f"   Updating data source for pivot table '{pivot_table_name}'...")

            # Get the data sheet
            data_sheet = self.workbook.sheets[data_sheet_name]

            # Extract start column and row from start_cell
            start_column = ''.join(filter(str.isalpha, start_cell))
            start_row = int(''.join(filter(str.isdigit, start_cell)))

            # Find the last row with data
            # Start checking from one row below the start_cell
            check_row = start_row + 1 if include_headers else start_row
            last_row = data_sheet.range(f'{start_column}{check_row}').end('down').row

            # Handle edge case where no data exists
            if last_row > 1048576 or last_row < start_row:
                last_row = start_row
                print(f"   Warning: No data found, using only header row")

            # If end_column not specified, find it dynamically
            if end_column is None:
                # Find last column with data in the header row
                end_column_obj = data_sheet.range(f'{start_column}{start_row}').end('right')
                end_column = end_column_obj.get_address(False, False).split(':')[0]
                end_column = ''.join(filter(str.isalpha, end_column))

            # Build the dynamic range address
            range_address = f"{start_cell}:{end_column}{last_row}"
            full_range_address = f"{data_sheet_name}!{range_address}"

            print(f"   Detected data range: {full_range_address}")
            print(f"   Range contains {last_row - start_row + 1} rows (including headers)")

            # Get the pivot table
            sheet = self.workbook.sheets[sheet_name]
            pivot_table = sheet.api.PivotTables(pivot_table_name)

            # Update the source data range
            # Build the full workbook reference for the range
            workbook_name = self.workbook.name
            full_source = f"'{workbook_name}'!{full_range_address}"

            print(f"   Updating pivot table source to: {full_source}")
            pivot_table.ChangePivotCache(
                self.workbook.api.PivotCaches().Create(
                    SourceType=1,  # xlDatabase
                    SourceData=full_source
                )
            )

            # Refresh the pivot table
            pivot_table.RefreshTable()

            print(f"    Pivot table '{pivot_table_name}' source updated and refreshed successfully")

        except Exception as e:
            print(f"    Error updating pivot table source: {e}")
            raise
    
    def convert_text_to_date_dmy(self, worksheet_name, date_range):
        """
        Convert a text column to date format with DMY (Day-Month-Year) order using xlwings.
        
        Parameters:
        -----------
        workbook_path : str
            Full path to the Excel workbook
        worksheet_name : str
            Name of the worksheet containing the date column
        date_range : str
            Range of the date column (e.g., 'A1:A100' or 'A:A' for entire column)
        
        Example:
        --------
        convert_text_to_date_dmy('C:/data/myfile.xlsx', 'Sheet1', 'A2:A50')
        """
        
        wb = self.workbook
        
        try:
            # Select the worksheet
            ws = wb.sheets[worksheet_name]
            
            # Select the date range
            date_column = ws.range(date_range)
            
            # Get the Excel API object to access Text to Columns functionality
            excel_range = date_column.api
            
            # Perform Text to Columns with DMY date format
            # Parameters for TextToColumns method:
            # Destination, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, 
            # Comma, Space, Other, OtherChar, FieldInfo
            
            excel_range.TextToColumns(
                Destination=excel_range,
                DataType=1,  # xlDelimited = 1
                TextQualifier=1,  # xlTextQualifierDoubleQuote = 1
                ConsecutiveDelimiter=False,
                Tab=False,
                Semicolon=False,
                Comma=False,
                Space=False,
                Other=False,
                FieldInfo=[[1, 4]],  # [[column_number, xlColumnDataType]] where 4 = xlDMYFormat
                TrailingMinusNumbers=True
            )
            
            # Save the workbook
            wb.save()
            print(f"Successfully converted date column and saved workbook: {worksheet_name}")
            
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            raise
        
        finally:
            pass

    def copy_paste_formulas(self,worksheet, copy_range, paste_range):
        """
        Copy formulas from one range to another in Excel using xlwings.
        
        Parameters:
        -----------
        worksheet : xlwings.Sheet
            The worksheet object to work with
        copy_range : str
            Source range in A1 notation (e.g., 'A1:C10')
        paste_range : str
            Destination range in A1 notation (e.g., 'E1:G10')
        
        Example:
        --------
        wb = xw.Book('myfile.xlsx')
        ws = wb.sheets['Sheet1']
        copy_paste_formulas(ws, 'A1:C10', 'E1:G10')
        """
        wb = self.workbook
        
        
        # Select the worksheet
        ws = wb.sheets[worksheet]
        # Get the source range
        source = ws.range(copy_range)
        
        # Get the destination range
        destination = ws.range(paste_range)
        
        # Copy formulas (not values)
        destination.formula = source.formula

    def insert_rows(self, sheet_name: str, start_row: int, num_rows: int) -> None:
        """
        Insert n number of rows starting from a specific row in a worksheet
        
        Args:
            sheet_name: Name of the worksheet
            start_row: The row number where insertion begins (rows will be inserted above this row)
            num_rows: Number of rows to insert
        """
        sheet = self.workbook.sheets[sheet_name]
        
        print(f"   Inserting {num_rows} rows at row {start_row} in sheet '{sheet_name}'...")
        
        # Select the range of rows to insert
        # Format: "start_row:end_row" e.g., "5:9" to insert 5 rows starting at row 5
        end_row = start_row + num_rows - 1
        row_range = f"{start_row}:{end_row}"
        
        # Use Excel's API to insert entire rows
        sheet.range(row_range).api.EntireRow.Insert()
        
        print(f"    Successfully inserted {num_rows} rows at row {start_row}")


    def color_column_text_black(self, worksheet_name, column_letter):
        """
        Colors all text in a specified column to black.
        
        Args:
            worksheet_name (str): Name of the worksheet
            column_letter (str): Column letter (e.g., 'A', 'B', 'C')
        """
        try:
            # Access the worksheet
            ws = self.workbook.sheets[worksheet_name]
            
            # Select the entire column
            column_range = ws.range(f"{column_letter}:{column_letter}")
            
            # Set font color to black
            # Option 1: Using RGB color code (0 = black)
            column_range.api.Font.Color = 0
            
            # Option 2: Alternative using RGB values
            # column_range.api.Font.Color = self.workbook.app.rgb_to_int((0, 0, 0))
            
        except Exception as e:
            print(f"Error coloring column {column_letter} in {worksheet_name}: {e}")


    def get_row_count(self, sheet_name: str, column: str, start_row: int = 1) -> int:
        """
        Get the number of rows with data in a specific column
        
        Args:
            sheet_name: Name of the worksheet
            column: Column letter (e.g., 'A', 'B', 'AB')
            start_row: Row to start counting from (default: 1)
        
        Returns:
            Number of rows with data in the specified column
        """
        sheet = self.workbook.sheets[sheet_name]
        
        # Get the last row with data in the column
        last_row = sheet.range(f'{column}{start_row}').end('down').row
        
        # Handle case where column is empty or has only one cell
        if last_row >= 1048576:  # Excel's max row limit
            # Check if the start cell itself has data
            if sheet.range(f'{column}{start_row}').value is None:
                return 0
            return 1
        
        row_count = last_row - start_row + 1
        return row_count

    def validate_yardlist_checkbox(self) -> None:
        """
        Validate that the checkbox control in cell M1 of the YardList sheet is ticked.
        This check must be performed when executing month-end calculations.
        
        Raises:
            ValueError: If the checkbox in M1 is not ticked
        """
        print("\n" + "=" * 70)
        print("VALIDATING YARDLIST CHECKBOX")
        print("=" * 70)
        
        try:
            # Get the YardList sheet
            yardlist_sheet = self.workbook.sheets['1.YardList']
            
            print(f"   Checking checkbox in cell M1 in 1.YardList sheet...")
            
            # Find the checkbox control in M1
            checkbox_found = False
            checkbox_checked = False
            
            for ole in yardlist_sheet.api.OLEObjects():
                # Check if this OLE object is positioned at M1
                if ole.TopLeftCell.Address == yardlist_sheet.range('M1').api.Address:
                    checkbox_found = True
                    # Check if the checkbox is ticked (Value = True or 1)
                    checkbox_checked = ole.Object.Value == True or ole.Object.Value == 1
                    break
            
            if not checkbox_found:
                print(f"    ERROR: No checkbox found in cell M1!")
                print("=" * 70)
                raise ValueError(
                    "Month-end calculation cannot proceed: "
                    "No checkbox control found in cell M1 of the 1.YardList sheet."
                )
            
            if not checkbox_checked:
                print(f"    ERROR: Checkbox in cell M1 is NOT ticked!")
                print("=" * 70)
                raise ValueError(
                    "Month-end calculation cannot proceed: "
                    "The checkbox in cell M1 of the 1.YardList sheet must be ticked before performing month-end calculation."
                )
            
            print(f"    [OK] Checkbox in cell M1 is ticked")
            print("=" * 70)
            
        except Exception as e:
            if "1.YardList" in str(e) and "not found" in str(e).lower():
                raise ValueError("YardList sheet not found in workbook")
            raise
    
    def is_month_end(self, date_str: str) -> bool:
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                date = datetime.strptime(date_str, fmt)
                last_day = calendar.monthrange(date.year, date.month)[1]
                return date.day == last_day
            except ValueError:
                continue
        raise ValueError(f"Unable to parse date: {date_str}")
    
    def add_yardlist_completion_tick(self, strDate) -> None:
        is_month_end = self.is_month_end(strDate)
        yardlist_sheet = self.workbook.sheets['1.YardList']
        p1_cell = yardlist_sheet.range('P1')
        p1_value = p1_cell.value

        # Normalize: handles bool, int (1/0), and text ("TRUE"/"FALSE")
        def normalize(val) -> bool:
            if isinstance(val, bool):
                return val
            if isinstance(val, int):
                return val != 0
            if isinstance(val, str):
                return val.strip().upper() == "TRUE"
            return False  # None or unexpected type defaults to False

        is_checked = normalize(p1_value)

        if not is_month_end and not is_checked:
            print(f"   Date {strDate} is not month-end, skipping checkbox tick")
            return

        elif not is_month_end and is_checked:
            p1_cell.value = "FALSE"
            print(f"   Date {strDate} is not month-end, checkbox unticked")

        elif is_month_end and not is_checked:
            p1_cell.value = "TRUE"
            print(f"   Date {strDate} is month-end, checkbox ticked")

        elif is_month_end and is_checked:
            print(f"   Date {strDate} is month-end, checkbox already ticked, no change")

    def filter_copy_delete(
                                self,
                                source_sheet_name: str,
                                target_sheet_name: str,
                                header_row: int,
                                filter_column: str,       # e.g., "A", "C", "M"
                                filter_criteria,          # value to filter by e.g., "DONE", 1, True
                                paste_start_cell: str = "A1"  # where to start pasting in target sheet
                            ) -> None:

        wb = self.workbook
        source_sheet = wb.sheets[source_sheet_name]
        sheet_names = [s.name for s in wb.sheets]
        if target_sheet_name not in sheet_names:
            target_sheet = wb.sheets.add(target_sheet_name, after=wb.sheets[-1])
            print(f"   Sheet '{target_sheet_name}' not found — created automatically")
        else:
            target_sheet = wb.sheets[target_sheet_name]

        # ── 1. Find last used row and column in source ──────────────────────────
        last_row = source_sheet.range('A1').current_region.last_cell.row
        last_col = source_sheet.range('A1').current_region.last_cell.column
        last_col_letter = source_sheet.range(1, last_col).address.split('$')[1]

        print(f"   Source: '{source_sheet_name}' | Data range: A{header_row}:{last_col_letter}{last_row}")

        # ── 2. Define full data range starting from header row ──────────────────
        data_range = source_sheet.range(f'A{header_row}:{last_col_letter}{last_row}')

        # ── 3. Get column index (e.g., "C" -> 3) for AutoFilter field number ─────
        filter_col_index = source_sheet.range(f'{filter_column}1').column

        # ── 4. Clear any existing filters ───────────────────────────────────────
        if source_sheet.api.AutoFilterMode:
            source_sheet.api.AutoFilterMode = False

        # ── 5. Apply AutoFilter on header row ───────────────────────────────────
        data_range.api.AutoFilter(
            Field=filter_col_index,
            Criteria1=str(filter_criteria)
        )
        print(f"   Filter applied on column '{filter_column}' for criteria '{filter_criteria}'")

        # ── 6. Get visible (filtered) rows excluding header ──────────────────────
        try:
            visible_range = source_sheet.range(
                f'A{header_row + 1}:{last_col_letter}{last_row}'
            ).api.SpecialCells(12)  # 12 = xlCellTypeVisible
        except Exception:
            print(f"   No visible rows found matching criteria '{filter_criteria}', aborting")
            source_sheet.api.AutoFilterMode = False
            return

        # ── 7. Find next empty row in target sheet ───────────────────────────────
        target_last_row = target_sheet.range(paste_start_cell).current_region.last_cell.row
        target_paste_row = target_last_row + 1 if target_sheet.range(paste_start_cell).value else int(
            paste_start_cell[1:] if len(paste_start_cell) > 1 else 1
        )
        target_col_letter = paste_start_cell[0]  # e.g., "A" from "A1"
        target_paste_cell = f'{target_col_letter}{target_paste_row}'

        # ── 8. Copy visible rows to target sheet ─────────────────────────────────
        visible_range.Copy(target_sheet.range(target_paste_cell).api)
        print(f"   Copied filtered rows to '{target_sheet_name}' starting at '{target_paste_cell}'")

        # ── 9. Delete filtered (visible) rows from source ────────────────────────
        visible_range.Delete()
        print(f"   Deleted filtered rows from '{source_sheet_name}'")

        # ── 10. Remove the AutoFilter ─────────────────────────────────────────────
        source_sheet.api.AutoFilterMode = False
        print(f"   AutoFilter cleared | Operation complete")

    def calculate_sum_range(self, ws , start_cell, end_column, data_lookup_cell, direction="down"):
        # Find EndRow using Ctrl+Down or Ctrl+Up
        xl_dir = xw.constants.Direction.xlDown if direction == "down" else xw.constants.Direction.xlUp
        ws = self.workbook.sheets[ws]
        end_row = ws.range(data_lookup_cell).end(xl_dir).row

        # Build range and sum
        start_col = re.match(r"([A-Za-z]+)", start_cell).group(1).upper()
        start_row = re.match(r"[A-Za-z]+(\d+)", start_cell).group(1)
        sum_range = f"{start_col}{start_row}:{end_column.upper()}{end_row}"


        return f"=SUM({sum_range})"


    def filter_excel_data(worksheet, filter_column: Union[int, str], filter_criteria: Union[str, List[str], dict]):
        """
        Apply filter to Excel worksheet
        
        Parameters:
        -----------
        worksheet : xlwings.Sheet
            The worksheet object to filter
        filter_column : int or str
            Column number (1-based) or column letter (e.g., 'A', 'B')
        filter_criteria : str, list, or dict
            Filter criteria:
            - String: "Apple", ">100", "<=50", "<>0", "=#N/A", "=*text*"
            - List: ["Apple", "Banana", "Orange"]
            - Dict: {"operator": ">=", "value": 100}
                    {"operator": "between", "value1": 10, "value2": 100}
                    {"operator": "top", "value": 10}
                    {"operator": "contains", "value": "text"}
        
        Returns:
        --------
        bool : True if successful, False otherwise
        
        Examples:
        ---------
        # Single value
        filter_excel_data(ws, 'A', "Apple")
        
        # Comparison
        filter_excel_data(ws, 2, ">100")
        
        # Multiple values
        filter_excel_data(ws, 'C', ["Red", "Blue", "Green"])
        
        # Between values
        filter_excel_data(ws, 3, {"operator": "between", "value1": 10, "value2": 100})
        
        # Contains text
        filter_excel_data(ws, 'E', {"operator": "contains", "value": "apple"})
        
        # Filter errors
        filter_excel_data(ws, 'G', "=#N/A")
        """
        try:
            # Convert column letter to number if needed
            if isinstance(filter_column, str):
                filter_column = xw.utils.col_name_to_num(filter_column)
            
            # Get the used range
            used_range = worksheet.used_range
            
            if used_range.rows.count == 0:
                print("Worksheet is empty")
                return False
            
            # Clear any existing filters
            if worksheet.api.AutoFilterMode:
                worksheet.api.AutoFilterMode = False
            
            # Apply AutoFilter
            used_range.api.AutoFilter()
            
            # Apply criteria based on type
            if isinstance(filter_criteria, str):
                # Single string criteria
                used_range.api.AutoFilter(Field=filter_column, Criteria1=filter_criteria)
                
            elif isinstance(filter_criteria, list):
                # Multiple values
                if len(filter_criteria) == 1:
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=filter_criteria[0])
                else:
                    used_range.api.AutoFilter(Field=filter_column, 
                                            Criteria1=filter_criteria, 
                                            Operator=7)  # xlFilterValues
                    
            elif isinstance(filter_criteria, dict):
                # Advanced criteria
                operator = filter_criteria.get('operator', '').lower()
                
                if operator == 'between':
                    value1 = filter_criteria.get('value1')
                    value2 = filter_criteria.get('value2')
                    used_range.api.AutoFilter(Field=filter_column,
                                            Criteria1=f">={value1}",
                                            Operator=1,  # xlAnd
                                            Criteria2=f"<={value2}")
                
                elif operator == 'top':
                    value = filter_criteria.get('value', 10)
                    is_percent = filter_criteria.get('percent', False)
                    op = 6 if is_percent else 5  # xlTop10Percent or xlTop10Items
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=value, Operator=op)
                
                elif operator == 'bottom':
                    value = filter_criteria.get('value', 10)
                    is_percent = filter_criteria.get('percent', False)
                    op = 8 if is_percent else 9  # xlBottom10Percent or xlBottom10Items
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=value, Operator=op)
                
                elif operator in ['>', '<', '>=', '<=', '=', '<>']:
                    value = filter_criteria.get('value')
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=f"{operator}{value}")
                
                elif operator == 'contains':
                    value = filter_criteria.get('value', '')
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=f"=*{value}*")
                
                elif operator == 'not_contains':
                    value = filter_criteria.get('value', '')
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=f"<>*{value}*")
                
                elif operator == 'begins_with':
                    value = filter_criteria.get('value', '')
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=f"{value}*")
                
                elif operator == 'ends_with':
                    value = filter_criteria.get('value', '')
                    used_range.api.AutoFilter(Field=filter_column, Criteria1=f"*{value}")
                
                elif operator == 'blanks':
                    used_range.api.AutoFilter(Field=filter_column, Criteria1="=")
                
                elif operator == 'non_blanks':
                    used_range.api.AutoFilter(Field=filter_column, Criteria1="<>")
                
                else:
                    print(f"Unknown operator: {operator}")
                    return False
            
            print(f"Filter applied successfully on column {filter_column}")
            return True
            
        except Exception as e:
            print(f"Error applying filter: {e}")
            return False
        
    def delete_filtered_data_fast(worksheet, exclude_header: bool = True, remove_filter: bool = True):
        """
        Faster version using SpecialCells - deletes visible filtered rows
        Works better for large datasets
        
        Parameters:
        -----------
        worksheet : xlwings.Sheet
            The worksheet object with filtered data
        exclude_header : bool, optional
            If True, keeps the header row (default: True)
        remove_filter : bool, optional
            If True, removes filter after deletion (default: True)
        
        Returns:
        --------
        int : Number of rows deleted
        """
        try:
            # Check if filter is applied
            if not worksheet.api.AutoFilterMode:
                print("No filter applied. Nothing to delete.")
                return 0
            
            used_range = worksheet.used_range
            first_row = used_range.row
            
            # Get only visible cells
            try:
                visible_range = used_range.api.SpecialCells(12)  # xlCellTypeVisible
            except:
                print("No visible cells found")
                return 0
            
            rows_deleted = 0
            
            # Process each visible area
            for area in visible_range.Areas:
                area_first_row = area.Row
                
                # Skip header if needed
                if exclude_header and area_first_row == first_row:
                    # If this area includes the header, we need to handle it carefully
                    if area.Rows.Count > 1:
                        # Delete all rows except the first one in this area
                        delete_range = area.Offset(1, 0).Resize(area.Rows.Count - 1, area.Columns.Count)
                        delete_range.EntireRow.Delete()
                        rows_deleted += area.Rows.Count - 1
                else:
                    # Delete entire area
                    area.EntireRow.Delete()
                    rows_deleted += area.Rows.Count
            
            # Remove filter if requested
            if remove_filter and worksheet.api.AutoFilterMode:
                worksheet.api.AutoFilterMode = False
            
            print(f"Successfully deleted {rows_deleted} filtered rows")
            return rows_deleted
            
        except Exception as e:
            print(f"Error deleting filtered data: {e}")
            try:
                if worksheet.api.AutoFilterMode:
                    worksheet.api.AutoFilterMode = False
            except:
                pass
            return 0

    def refresh_pivot_table(self, sheet_name: str, pivot_table_name: str):
        """
        Refresh a specific pivot table in a given sheet.
        
        Args:
            sheet_name (str): Name of the sheet containing the pivot table
            pivot_table_name (str): Name of the pivot table to refresh
        """
        try:
            # Get the sheet
            sheet = self.workbook.sheets[sheet_name]
            
            # Access the pivot table via the sheet's API
            pivot_table = sheet.api.PivotTables(pivot_table_name)
            
            # Refresh the pivot table
            pivot_table.RefreshTable()
            
            print(f"Pivot table '{pivot_table_name}' in sheet '{sheet_name}' refreshed successfully.")

        except Exception as e:
            print(f"Error refreshing pivot table '{pivot_table_name}' in sheet '{sheet_name}': {e}")
            raise

    def _get_last_row(self, ws, col: int) -> int:
        return ws.range((1048576, col)).end('up').row

    def _copy_col_values(self, src_ws, src_col, src_start, src_end, dst_ws, dst_col, dst_start):
        values = src_ws.range((src_start, src_col), (src_end, src_col)).options(transpose=True).value
        dst_ws.range((dst_start, dst_col), (dst_start + (src_end - src_start), dst_col)).options(transpose=True).value = values


    def ol_adjustment(self, wb):
        """
        Python equivalent of the VBA macro OLAdjstmnt.

        Expects the workbook to contain sheets: "ECL", "toSummary", "Pivot".
        Returns True on success, False if the BR23 guard fails.

        Notes:
        - Checkbox toggling (Dynamic Mode) is not possible via openpyxl;
            add xlwings/win32com calls where the comments indicate if needed.
        - GoalSeek (BS25 = 0, changing BQ2) requires you to supply the
            BS25 formula as a Python callable — see the marked section below.
        - time.sleep() calls mirror the VBA Application.Wait pauses;
            remove them if your data does not depend on recalculation settling.
        """
        import time

        ws_ecl     = wb.sheets["ECL"]
        ws_summary = wb.sheets["toSummary"]
        ws_pivot   = wb.sheets["Pivot"]

        # Guard: BR23 must contain a numeric value
        br23 = ws_ecl.range("BR23").value
        if br23 is None or not isinstance(br23, (int, float)):
            return False

        # Clear Adjusted LGD percentage
        ws_ecl.range("BQ2").value = None

        # --- Untick "Dynamic Mode" (Check Box 1) ---
        # Add: ws_ecl.api.CheckBoxes("Check Box 1").Value = -4146  (xlwings)
        time.sleep(5)

        # Copy col 49 (AW) -> col 57 (BE) as values, rows 3..lastRow
        last_row = self._get_last_row(ws_ecl, col=2)
        self._copy_col_values(ws_ecl, 49, 3, last_row, ws_ecl, 57, 3)

        time.sleep(5)

        # Copy col 64 (BL) -> col 51 (AY) as values
        last_row = self._get_last_row(ws_ecl, col=2)
        self._copy_col_values(ws_ecl, 64, 3, last_row, ws_ecl, 51, 3)

        # Clear toSummary from row 2 down
        summary_region = ws_summary.range('A1').current_region
        summary_last_row = summary_region.last_cell.row
        summary_last_col = summary_region.last_cell.column
        if summary_last_row >= 2:
            ws_summary.range((2, 1), (summary_last_row, summary_last_col)).clear_contents()

        # Copy data blocks from ECL -> toSummary
        last_row = self._get_last_row(ws_ecl, col=2)
        self._copy_col_values(ws_ecl,  2, 2, last_row, ws_summary, 1, 2)   # B  -> A
        self._copy_col_values(ws_ecl, 57, 2, last_row, ws_summary, 2, 2)   # BE -> B
        self._copy_col_values(ws_ecl, 56, 2, last_row, ws_summary, 3, 2)   # BD -> C
        self._copy_col_values(ws_ecl, 50, 2, last_row, ws_summary, 4, 2)   # AX -> D
        self._copy_col_values(ws_ecl, 51, 2, last_row, ws_summary, 5, 2)   # AY -> E

        # --- Re-tick "Dynamic Mode" (Check Box 1) ---
        # Add: ws_ecl.api.CheckBoxes("Check Box 1").Value = 1  (xlwings)
        time.sleep(5)

        # Copy col 63 (BK) -> col 62 (BJ) as values
        last_row = self._get_last_row(ws_ecl, col=2)
        self._copy_col_values(ws_ecl, 63, 3, last_row, ws_ecl, 62, 3)

        self._copy_col_values(ws_ecl, 61, 2, last_row, ws_summary, 6, 2)   # BI -> F
        self._copy_col_values(ws_ecl, 62, 2, last_row, ws_summary, 7, 2)   # BJ -> G

        # Paste Pivot A2:D36 -> H2 as values (single bulk write)
        ws_pivot.range('H2').value = ws_pivot.range('A2:D36').value

        # Copy BQ21 -> BR21 as value (Lease original IMP)
        ws_ecl.range("BR21").value = ws_ecl.range("BQ21").value

        # GoalSeek: make BS25 = 0 by changing BQ2
        # ----------------------------------------------------------------
        
        app = xw.apps.active
        wb = app.books.active
        ws_ecl = wb.sheets["ECL"]

        target = 0  # solve for BS25 == this value

        def formula_fn(x):
            ws_ecl.range("BQ2").value = x
            ws_ecl.book.app.calculate()          # recalculate
            return ws_ecl.range("BS25").value - target

        result = brentq(formula_fn, -1e6, 1e6)
        ws_ecl.range("BQ2").value = result       # write final answer back
        print(f"Solution: BQ2 = {result}")
        # ----------------------------------------------------------------

        # Copy col 55 (BC) -> toSummary col H
        last_row = self._get_last_row(ws_ecl, col=2)
        self._copy_col_values(ws_ecl, 55, 2, last_row, ws_summary, 8, 2)   # BC -> H

        return True


    def write_to_cell(self, sheet_name: str, cell_address: str, value):
        """
        Write a value to a specific cell in an Excel workbook.
        
        Args:
            sheet_name (str): Name of the worksheet
            cell_address (str): Cell address (e.g., 'A1', 'B2')
            value: Value to write to the cell
        """
        sheet = self.workbook.sheets[sheet_name]
        sheet.range(cell_address).value = value

    def process_portfolio_sheet(self, csv_path: str, date_str: str = None) -> int:
        """
        Process Portfolio sheet:
        1. Clear range A4:AB (to last row)
        2. Read CSV data (columns A-AA)
        3. Write CSV data starting at B4
        4. Fill column A with date

        Args:
            csv_path: Path to CSV file
            date_str: Date string to fill in column A (format: M/D/YYYY)
                     If None, uses today's date

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING PORTFOLIO SHEET")
        print("=" * 70)

        # Parse the date
        # Get date from command line argument
        date_str = sys.argv[1] if len(sys.argv) > 1 else None

        if date_str:
            for _fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
                try:
                    date_obj = datetime.strptime(date_str, _fmt)
                    date_str = date_obj.strftime("%d/%m/%Y")
                    print(f"   Using date: {date_str}")
                    break
                except ValueError:
                    continue
            else:
                print(f"   Invalid date format. Using today's date instead.")
                date_obj = datetime.now()
                date_str = date_obj.strftime('%d/%m/%Y')
                print(f"   Using today's date: {date_str}")
        else:
            date_obj = datetime.now()
            date_str = date_obj.strftime('%m/%d/%Y')
            print(f"   Using today's date: {date_str}")
        
        # Store the processing date
        self.processing_date = date_obj
        
        # Calculate month-end date
        # Get the last day of the month for the processing date
        if date_obj.month == 12:
            # December: month-end is Dec 31
            month_end = date_obj.replace(day=31)
        else:
            # Get first day of next month, then subtract 1 day
            next_month = date_obj.replace(day=1, month=date_obj.month + 1)
            month_end = next_month - pd.Timedelta(days=1)
        
        self.month_end_date = month_end
        print(f"   Month-end date calculated: {month_end.strftime('%d/%m/%Y')}")

        # Step 1: Clear existing data in Portfolio sheet
        print("\n1. Clearing existing data in Portfolio sheet...")
        self.clear_range_dynamic('Portfolio', 'A4', 'AB')

        # Step 2: Read CSV data
        print("\n2. Reading CSV data...")
        try:
            df_csv, num_rows, num_cols = self.read_csv_data(csv_path, num_columns=27)
        except:
            df_csv, num_rows, num_cols = self.read_excel_data(csv_path, num_columns=27)

        # Step 3: Write CSV data to columns B onwards (starting at B4)
        print("\n3. Writing CSV data to Portfolio sheet (starting at B4)...")
        # Convert DataFrame to list of lists
        csv_data = df_csv.values.tolist()
        self.write_data_to_range('Portfolio', 'B4', csv_data)

        # Step 4: Fill column A with date
        print("\n4. Filling column A with date...")
        end_row = 3 + num_rows  # Start at row 4, so end at 3 + num_rows
        self.fill_column_with_value('Portfolio', 'A', 4, end_row, date_str)
        data_range1 = f'G4:G{end_row}'   

        data_range2 = f'U4:U{end_row}'     
        # step 5: Convert Columns to Date 
        self.convert_text_to_date_dmy('Portfolio',data_range1)
        print("Succesfully Converted to Date format")
        self.convert_text_to_date_dmy('Portfolio',data_range2)
        print("Succesfully Converted to Date format")

        copy_range1 = f'AK4:AM4'
        Paste_range1 = f'AK4:AM{end_row}'  

        self.copy_paste_formulas('Portfolio', copy_range1, Paste_range1)

        copy_range1 = f'AO4:BS4'
        Paste_range1 = f'AO4:BS{end_row}'  

        self.copy_paste_formulas('Portfolio', copy_range1, Paste_range1)

        print("\n" + "=" * 70)
        print("PORTFOLIO SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Rows processed: {num_rows}")
        print(f"Date in column A: {date_str}")
        print(f"CSV data in columns B-{chr(65 + num_cols)}")  # B to corresponding column
        print("=" * 70)

        return num_rows

    def process_dpd_sheet(self, excel_path: str, date_str: str = None) -> int:
        """
        Process 1.DPD sheet:
        1. Clear range A2:E (to last row)
        2. Read Excel data (first 5 columns)
        3. Write data starting at A2
        4. Copy formulas from F2:O2 to end of data range

        Args:
            excel_path: Path to Excel file
            date_str: Date string for R1 (any common format); writes first day of that month as MM/DD/YYYY

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING 1.DPD SHEET")
        print("=" * 70)
        # Derive first and last date of the processing month in MM/DD/YYYY format
        _parsed_date = None
        if date_str:
            for _fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d\\%m\\%Y", "%m\\%d\\%Y"):
                try:
                    _parsed_date = datetime.strptime(date_str, _fmt)
                    break
                except ValueError:
                    continue
        if _parsed_date is None:
            _parsed_date = datetime.now()
        _month_first_date = _parsed_date.replace(day=1).strftime("%m/%d/%Y")
        import calendar
        _last_day = calendar.monthrange(_parsed_date.year, _parsed_date.month)[1]
        _month_last_date = _parsed_date.replace(day=_last_day).strftime("%m/%d/%Y")
        self.write_to_cell('1.DPD', 'R1', _month_first_date)
        self.write_to_cell('1.DPD', 'S1', _month_last_date)
        # Step 1: Clear existing data in 1.DPD sheet
        print("\n1. Clearing existing data in 1.DPD sheet...")
        self.clear_range_dynamic('1.DPD', 'A2', 'E')

        # Step 2: Read Excel data
        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(excel_path, num_columns=5, sheet_name='Month DPD Report New')

        # Step 3: Write Excel data starting at A2
        print("\n3. Writing data to 1.DPD sheet (starting at A2)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('1.DPD', 'A2', excel_data)

        # Step 4: Copy formulas from F2:O2 to the end of data range
        print("\n4. Copying formulas from F2:O2 to end of data range...")
        end_row = 1 + num_rows  # Data starts at row 2, so end at 1 + num_rows
        if end_row > 2:  # Only copy if there are rows beyond row 2
            self.copy_formulas_to_range('1.DPD', 'F2:O2', 3, end_row)
        else:
            print("   No additional rows to copy formulas to (only header row)")

        print("\n" + "=" * 70)
        print("1.DPD SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Rows processed: {num_rows}")
        print(f"Data range: A2:E{1 + num_rows}")
        print(f"Formulas range: F2:O{end_row}")
        print("=" * 70)

        return num_rows
    
    def process_NPA_sheet(self, excel_path: str, last_pf: str) -> int:
        """
        Process 2.NPA List sheet:
        1. Clear range A3:C (to last row)
        2. Read Excel data ('STR_DPD_CON_NO','NOOF_DATES','PERFORMING_MONTH_BASE')
        3. Write data starting at A3
        4. Filter Column J of 1.DPD sheet for "Add to NPA List" and append to 2.NPA List
        5. Copy formulas from D2:J2 to end of data range

        Args:
            excel_path: Path to Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING 2.NPA LIST SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in 2.NPA List sheet
        print("\n1. Clearing existing data in 2.NPA List sheet...")
        self.clear_range_dynamic('2.NPA List', 'A3', 'C')

        # Step 2: Read Excel data
        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(excel_path, num_columns=10, sheet_name='Month DPD Report New')
        if last_pf.endswith('.xlsb'):
            engine = 'pyxlsb'
        else:
            engine = 'openpyxl'
        df_excel_dpd, num_rows_dpd, num_cols = self.read_excel_data(last_pf, num_columns=27, sheet_name='1.DPD',engine=engine)
        df_excel_npa, num_rows_npa, num_cols = self.read_excel_data(last_pf, num_columns=27, sheet_name='2.NPA List',skip_rows=1, engine=engine)
        # Select only the required columns
        required_columns = ['STR_DPD_CON_NO']
        df_excel = df_excel[required_columns]
        num_cols = len(df_excel.columns)
        # map final dpd to contract number in dfexcel from df_excel_dpd
        contract_to_dpd = dict(zip(df_excel_dpd["CON_NO"], df_excel_dpd["FINAL DPD (NORMAL)"]))
        df_excel['dpd'] = df_excel['STR_DPD_CON_NO'].map(contract_to_dpd).fillna(0)
        contract_to_npa = dict(zip(df_excel_npa["CON_NO"], df_excel_npa["Contract-wise P/NP"]))
        df_excel['npa'] = df_excel['STR_DPD_CON_NO'].map(contract_to_npa).fillna('P')
        # Step 3: Write Excel data starting at A3
        print("\n3. Writing data to 2.NPA List sheet (starting at A3)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('2.NPA List', 'A3', excel_data)
        
        # Calculate the last row after writing excel data
        last_row_after_excel = 2 + num_rows  # Row 3 (A3) + num_rows - 1
        
        # Step 4: Filter and copy data from 1.DPD sheet
        print("\n4. Filtering Column J of 1.DPD sheet for 'Add to NPA List'...")
        
        # Get the workbook that's currently open
        dpd_sheet = self.workbook.sheets['1.DPD']
        
        # Get the used range of the sheet
        used_range = dpd_sheet.used_range
        dpd_data = used_range.value
        
        if dpd_data and len(dpd_data) > 1:  # Check if there's data beyond header
            # Filter rows where column J contains "Add to NPA List"
            filtered_rows = []
            for i, row in enumerate(dpd_data[1:], start=2):  # Skip header row, start counting from row 2
                if row and len(row) > 9 and row[9] == "Add to NPA List":  # Column J is index 9
                    # Add column A value, 0 for column B, and 'P' for column C
                    filtered_rows.append([row[0], 0, 'P'])
            
            if filtered_rows:
                print(f"   Found {len(filtered_rows)} rows with 'Add to NPA List' in column J")
                
                # Append filtered data to 2.NPA List sheet
                next_empty_row = last_row_after_excel + 1
                print(f"   Appending filtered data starting at row {next_empty_row}...")
                self.write_data_to_range('2.NPA List', f'A{next_empty_row}', filtered_rows)
                
                # Update the total number of rows
                num_rows += len(filtered_rows)
                last_row_after_excel = next_empty_row + len(filtered_rows) - 1
            else:
                print("   No rows found with 'Add to NPA List' in column J")
        else:
            print("   1.DPD sheet has no data rows to filter")
        
        # Step 5: Copy formulas from D2:J2 to the end of data range
        print("\n5. Copying formulas from D2:J2 to end of data range...")
        end_row = last_row_after_excel  # Use updated last row
        if end_row > 2:  # Only copy if there are rows beyond row 2
            self.copy_formulas_to_range('2.NPA List', 'D2:J2', 4, end_row)
        else:
            print("   No additional rows to copy formulas to (only header row)")


        print("\n" + "=" * 70)
        print("2.NPA LIST SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {num_rows}")
        print(f"Data range: A3:C{last_row_after_excel}")
        print(f"Formulas range: D3:J{end_row}")
        print("=" * 70)

        return num_rows

    def process_100_percent_sheet(self) -> int:
            """
            Process 3.100%_Provision sheet:
            1. Clear range A5:A (to last row)
            2. Read data from 1.DPD sheet Column A2:A (to last row)
            3. Write data starting at A5 in 3.100%_Provision sheet
            4. Copy formulas from B5:F5 to end of data range

            Returns:
                Number of rows processed
            """
            print("\n" + "=" * 70)
            print("PROCESSING 3.100%_Provision SHEET")
            print("=" * 70)

            # Step 1: Clear existing data in 3.100%_Provision sheet
            print("\n1. Clearing existing data in 3.100%_Provision sheet...")
            self.clear_range_dynamic('3.100%_Provision', 'A5', 'A')
            
            # Step 2: Read data from 1.DPD sheet Column A
            print("\n2. Reading data from 1.DPD sheet Column A...")
            
            # Get the 1.DPD sheet
            dpd_sheet = self.workbook.sheets['1.DPD']
            
            # Get the used range of the sheet
            used_range = dpd_sheet.used_range
            dpd_data = used_range.value
            
            if dpd_data and len(dpd_data) > 1:  # Check if there's data beyond header
                # Extract Column A data (starting from row 2, skipping header)
                column_a_data = []
                for i, row in enumerate(dpd_data[1:], start=2):  # Skip header row
                    if row and row[0]:  # Column A is index 0
                        column_a_data.append([row[0]])  # Wrap in list for write_data_to_range
                
                num_rows = len(column_a_data)
                
                if column_a_data:
                    print(f"   Found {num_rows} rows in Column A of 1.DPD sheet")
                    
                    # Step 3: Write data to 3.100%_Provision sheet starting at A5
                    print(f"\n3. Writing data to 3.100%_Provision sheet (starting at A5)...")
                    self.write_data_to_range('3.100%_Provision', 'A5', column_a_data)
                    
                    # Calculate the last row after writing data
                    last_row = 4 + num_rows  # Data starts at row 5, so last row is 4 + num_rows
                    
                    # Step 4: Copy formulas from B5:F5 to end of data range
                    print("\n4. Copying formulas from B5:F5 to end of data range...")
                    if last_row > 5:  # Only copy if there are rows beyond row 5
                        self.copy_formulas_to_range('3.100%_Provision', 'B5:F5', 6, last_row)
                    else:
                        print("   No additional rows to copy formulas to (only first data row)")
                    
                    print("\n" + "=" * 70)
                    print("3.100%_Provision SHEET PROCESSING COMPLETED!")
                    print("=" * 70)
                    print(f"Total rows processed: {num_rows}")
                    print(f"Data range: A5:A{last_row}")
                    print(f"Formulas range: B5:F{last_row}")
                    print("=" * 70)
                    
                    return num_rows
                else:
                    print("   No data found in Column A of 1.DPD sheet")
                    return 0
            else:
                print("   1.DPD sheet has no data rows")
                return 0
            

    def process_AllConSec_sheet(self, excel_path: str) -> int:
        """
        Process 5.AllConSec sheet:
        1. Clear range A30:C (to last row)
        2. Read Excel data (three Columns)
        3. Write data starting at A30

        Args:
            excel_path: Path to Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING 5.ALLCONSEC SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in 5.AllConSec sheet
        print("\n1. Clearing existing data in 5.AllConSec sheet...")
        self.clear_range_dynamic('5.AllConSec', 'A30', 'C')

        # Step 2: Read Excel data
        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(excel_path, num_columns=3, sheet_name='All Contracts Sector Report - I')
        # Select only the required columns


        # Step 3: Write Excel data starting at A30
        print("\n3. Writing data to 5.AllConSec sheet (starting at A30)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('5.AllConSec', 'A30', excel_data)

        print("\n" + "=" * 70)
        print("2.NPA LIST SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {num_rows}")
        print(f"Data range: A30:C{29 + num_rows}")
        print("=" * 70)

        return num_rows 
    
    def process_int_receivable_sheet(self, int_rec: str) -> int:
        """
        Process Int_Receivable sheet:
        1. Clear range A2:B (to last row)
        2. Read Excel data (two Columns)
        3. Write data starting at A2

        Args:
            excel_path: Path to Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING Int_Receivable SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in Int_Receivable sheet
        print("\n1. Clearing existing data in Int_Receivable sheet...")
        self.clear_range_dynamic('Int_Receivable', 'A2', 'B')
        xl = pd.ExcelFile(int_rec)
        print(xl.sheet_names)
        # Step 2: Read Excel data
        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(int_rec, num_columns=3, skip_rows=0, sheet_name=xl.sheet_names[0])
        #remove first column

        # Step 3: Write Excel data starting at A2
        print("\n3. Writing data to Int_Receivable sheet (starting at A2)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('Int_Receivable', 'A2', excel_data)

        print("\n" + "=" * 70)
        print("Int_Receivable SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {num_rows}")
        print(f"Data range: A2:B{1 + num_rows}")
        print("=" * 70)

        return num_rows 
    
    def process_moi_sheet(self, moi: str) -> int:
        """
        Process MOI_Adjustment sheet:
        1. Clear range A2:B (to last row)
        2. Read Excel data (two Columns)
        3. Write data starting at A2

        Args:
            excel_path: Path to Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING MOI_Adjustment SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in MOI_Adjustment sheet
        print("\n1. Clearing existing data in MOI_Adjustment sheet...")
        self.clear_range_dynamic('MOI_Adjustment', 'A2', 'B')

        # Step 2: Read Excel data
        print("\n2. Reading Excel data...")
        xl = pd.ExcelFile(moi)
        sheet_name = 'Final' if 'Final' in xl.sheet_names else xl.sheet_names[0]
        print(f"   Using sheet: '{sheet_name}'")
        df_excel, num_rows, num_cols = self.read_excel_data(moi, num_columns=2, skip_rows=2, sheet_name=sheet_name)

        # Step 3: Write Excel data starting at A2
        print("\n3. Writing data to MOI_Adjustment sheet (starting at A2)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('MOI_Adjustment', 'A2', excel_data)

        print("\n" + "=" * 70)
        print("MOI_Adjustment SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {num_rows}")
        print(f"Data range: A2:B{1 + num_rows}")
        print("=" * 70)

        return num_rows
    
    def process_YARD_sheet(self, excel_path: str) -> int:
        """
        Process 1.YardList sheet:
        1. Clear range A5:M (to last row)
        2. Read Excel data (13 Columns)
        3. Write data starting at A5
        4. Copy formulas from N5:T5 to end of data range

        Args:
            excel_path: Path to Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING 1.YARDLIST SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in 1.YardList sheet
        print("\n1. Clearing existing data in 1.YardList sheet...")
        self.clear_range_dynamic('1.YardList', 'A5', 'M')


        # Step 2: Read Excel data
        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(excel_path, num_columns=13, skip_rows=4)
        # Select only the required columns


        # Step 3: Write Excel data starting at A5
        print("\n3. Writing data to 1.YardList sheet (starting at A5)...")
        # Rows where Valuer Name is 'ASSUMED' carry a bogus date (1/0/1900) that
        # pandas reads as datetime.time, which COM cannot handle — clear those cells.
        if 'Valuer Name' in df_excel.columns and 'Date of Valuation' in df_excel.columns:
            mask = df_excel['Valuer Name'].astype(str).str.strip().str.upper() == 'ASSUMED'
            df_excel.loc[mask, 'Date of Valuation'] = None
            print(f"   Cleared 'Date of Valuation' for {mask.sum()} ASSUMED row(s)")
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('1.YardList', 'A5', excel_data)
        self.clear_range_dynamic('1.YardList', 'N10', 'T')
        self.copy_formulas_to_range('1.YardList', 'N1:T1', 5, num_rows + 4)

        print("\n" + "=" * 70)
        print("1.YARDLIST SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {num_rows}")
        print(f"Data range: A5:M{4 + num_rows}")
        print("=" * 70)

        return num_rows

    def process_YARDCalculation_sheet(self) -> int:
        """
        Process 4.YardCalculation sheet:
        1. Clear range A4:A (to last row)
        2. Read Excel Data from Same workbook 1.YardList (Column A5 : A to last row)
        3. Write data starting at A4 in 4.YardCalculation sheet

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING 4.YARDCALCULATION SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in 4.YardCalculation sheet
        print("\n1. Clearing existing data in 4.YardCalculation sheet...")
        self.clear_range_dynamic('4.YardCalculation', 'A4', 'A')

        # Step 2: Read data from 1.YardList sheet Column A
        print("\n2. Reading data from 1.YardList sheet Column A...")

        # Get the 1.YardList sheet
        yardlist_sheet = self.workbook.sheets['1.YardList']

        # Get the used range of the sheet
        used_range = yardlist_sheet.used_range
        yardlist_data = used_range.value

        if yardlist_data and len(yardlist_data) > 4:  # Check if there's data beyond row 4
            # Extract Column A data (starting from row 5, index 4)
            column_a_data = []
            for row in yardlist_data[4:]:  # Skip rows 0-3 (rows 1-4 in Excel)
                if row and row[0]:  # Column A is index 0
                    column_a_data.append([row[0]])  # Wrap in list for write_data_to_range

            num_rows = len(column_a_data)

            if column_a_data:
                print(f"   Found {num_rows} rows in Column A of 1.YardList sheet")

                # Step 3: Write data to 4.YardCalculation sheet starting at A4
                print(f"\n3. Writing data to 4.YardCalculation sheet (starting at A4)...")
                self.write_data_to_range('4.YardCalculation', 'A4', column_a_data)

                # Calculate the last row after writing data
                last_row = 3 + num_rows  # Data starts at row 4, so last row is 3 + num_rows
                self.copy_formulas_to_range('4.YardCalculation', 'B4:F4', 4, last_row)
                # df_excel, num_rows_d, num_cols = self.read_excel_data( num_columns=6, skip_rows=2, sheet_name='4.YardCalculation')
                # # get df #N/A values drop from the df
                # df_excel = df_excel.replace('#N/A', pd.NA).dropna()
                # nd = len(df_excel)
                # num_rows = len(df_excel)
                # self.clear_range_dynamic('4.YardCalculation', f'A6', 'F')
                # self.write_data_to_range('4.YardCalculation', f'A4', df_excel['CONTRACT NO'].values.tolist())
                # self.copy_formulas_to_range('4.YardCalculation', 'B4:F4', 4, 4 + nd)
                ws = self.workbook.sheets['4.YardCalculation']
                start_range = ws.range("D4")
                start_row = start_range.row
                start_col = start_range.column
                
                rows_to_delete = []
                row = start_row
                
                # Loop through rows
                while True:
                    cell = ws.range(row, start_col)
                    # Stop at blank
                    if cell.api.Text is None or cell.api.Text == '':
                        break
                    # Check for #N/A
                    if cell.api.Text == "#N/A":
                        rows_to_delete.append(row)

                    row += 1
                    
                    # Safety limit
                    if row > start_row + 10000:
                        break
                
                # Delete rows in reverse
                for row_num in reversed(rows_to_delete):
                    ws.range(f'{row_num}:{row_num}').api.Delete()
                
                print("=" * 70)
                print(f"Total rows processed: {num_rows}")
                print(f"Data range: A4:A{last_row + 4}")
                print("=" * 70)

                return num_rows
            else:
                print("   No data found in Column A of 1.YardList sheet")
                return 0
        else:
            print("   1.YardList sheet has no data rows")
            return 0  

    def process_Legal_sheet(self, excel_path_mediation: str,legal_cases: str) -> int:
        """
        Process 2.Legal sheet:
        1. Clear range A7:C (to last row)
        2. Read Excel data (2 Columns)
        3. Write data starting at A7

        Args:
            excel_path: Path to Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING 2.LEGAL SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in 2.Legal sheet
        print("\n1. Clearing existing data in 2.Legal sheet...")
        self.clear_range_dynamic('2.Legal', 'A7', 'C')

        # Get first sheet name using pandas (without opening Excel)
        print("\n2. Getting sheet names...")
        excel_file = pd.ExcelFile(excel_path_mediation, engine='openpyxl')
        first_sheet_name = excel_file.sheet_names[0]
        excel_file.close()
        print(f"   First sheet name: {first_sheet_name}")

        # Step 3: Read Excel data
        print("\n3. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(excel_path_mediation, num_columns=2, sheet_name=first_sheet_name)
        # Select only the required columns
        #df excel second column name using index
        status_column = df_excel.columns[1]
        #Filter df_excel to only have following values in the July status column: "SETTLEMENT","NON SETTLEMENT", "PROCESSING", "TO BE HANDED OVER"
        df_excel = df_excel[df_excel[status_column].isin(["SETTLEMENT","NON SETTLEMENT", "PROCESSING", "TO BE HANDED OVER"])]
        df_excel['Type'] = "Legal Cases"
        df_excel = df_excel[['STR_CONTRACT_NO', 'Type', status_column]]

        # Step 4: Write Excel data starting at A7
        print("\n4. Writing mediation data to 2.Legal sheet (starting at A7)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('2.Legal', 'A7', excel_data)

        # Step 5: Read legal cases data
        print("\n5. Reading legal cases data...")
        df_excel, num_rows, num_cols = self.read_excel_data(legal_cases, num_columns=1, sheet_name='Export')
        # Select only the required columns
        df_excel['Type'] = "Legal Cases"

        # Step 6: Write legal cases data
        print("\n6. Writing legal cases data to 2.Legal sheet (appending after mediation data)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        sheet = self.workbook.sheets['2.Legal']
        last_row_with_data = sheet.range('A7').end('down').row
        self.write_data_to_range('2.Legal', f'A{last_row_with_data + 1}', excel_data)

        print("\n" + "=" * 70)
        print("2.LEGAL SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {num_rows}")
        print(f"Data range: A7:C{6 + num_rows}")
        print("=" * 70)

        return num_rows

    def process_worst_data_sheet(self, MATCAP: str,MATINT: str) -> int:
        """
        Process Worst_Data sheet:
        1. Clear range A3:I (to last row)
        2. Read Excel data (2 Columns)
        3. Write data starting at A3
        4. Read Excel data (2 Columns)
        5. Write data starting at D3
        Args:
            excel_path: Path to Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING Worst_Data SHEET")
        print("=" * 70)

        # Step 1: Clear existing data in Worst_Data sheet
        print("\n1. Clearing existing data in Worst_Data sheet...")
        self.clear_range_dynamic('Worst_Data', 'A3', 'I')


        # Step 2: Read Excel data
        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_csv_data(MATCAP, num_columns=2)
        # Select only the required columns
        # Step 3: Write Excel data starting at A7
        print("\n3. Writing data to Worst_Data sheet (starting at A3)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('Worst_Data', 'A3', excel_data)

        df_excel, num_rows, num_cols = self.read_csv_data(MATINT, num_columns=2)
        # Step 3: Write Excel data starting at A3
        print("\n3. Writing data to Worst_Data sheet (starting at A3 last row)...")
        # Convert DataFrame to list of lists
        excel_data = df_excel.values.tolist()

        self.write_data_to_range('Worst_Data','D3', excel_data)

        print("\n" + "=" * 70)
        print("WORST_DATA SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {num_rows}")
        print(f"Data range: A3:C{2 + num_rows}")
        print("=" * 70)

        return num_rows
    
    def process_Ins_Rec_sheet(self, excel_path: str, date_str) -> int:
        """
        Process 3.Ins_Rec sheet:
        - If day >= 25: Copy Portfolio sheet E4:E{last_row} -> 3.Ins_Rec B6
        - If day < 25:  Read external Excel and write to 3.Ins_Rec A6

        Args:
            excel_path: Path to Excel file
            date_str:   Date string in any supported format

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING 3.INS_REC SHEET")
        print("=" * 70)

        # ── Parse date_str to get the day ───────────────────────────────────────
        parsed_date = None
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                parsed_date = datetime.strptime(date_str, fmt)
                break
            except ValueError:
                continue

        if parsed_date is None:
            raise ValueError(f"Unable to parse date: {date_str}")

        day = parsed_date.day
        print(f"\n   Parsed date: {parsed_date.strftime('%d/%m/%Y')} | Day: {day}")

        # ── Step 1: Clear existing data in 3.Ins_Rec sheet ─────────────────────
        print("\n1. Clearing existing data in 3.Ins_Rec sheet (A6:Q)...")
        self.clear_range_dynamic('3.Ins_Rec', 'A6', 'Q')

        # ── Branch: Day >= 25 ───────────────────────────────────────────────────
        if day >= 25:
            print(f"\n   Day {day} >= 25: Copying Portfolio E column -> 3.Ins_Rec B6")

            portfolio_sheet = self.workbook.sheets['Portfolio']
            ins_rec_sheet   = self.workbook.sheets['3.Ins_Rec']

            # Find last data row in Portfolio column E (starting from row 4)
            last_row = portfolio_sheet.range('E4').end('down').row
            print(f"   Portfolio data range: E4:E{last_row}")

            # Copy E4:E{last_row} from Portfolio -> paste to B6 in 3.Ins_Rec
            source_range = portfolio_sheet.range(f'E4:E{last_row}')
            target_cell  = ins_rec_sheet.range('B6')

            source_range.copy(destination=target_cell)
            print(f"   Copied Portfolio E4:E{last_row} -> 3.Ins_Rec B6:B{5 + (last_row - 3)}")

            num_rows = last_row - 3  # E4 is row 4, so offset by 3

        # ── Branch: Day < 25 ────────────────────────────────────────────────────
        else:
            print(f"\n   Day {day} < 25: Reading external Excel and writing to 3.Ins_Rec A6")

            # Step 2: Read Excel data
            print("\n2. Reading Excel data...")
            df_excel, num_rows, num_cols = self.read_excel_data(
                excel_path,
                num_columns=16,
                sheet_name='Backend Due Types - Contract wi'
            )

            # Step 3: Write data starting at A6
            print("\n3. Writing data to 3.Ins_Rec sheet (starting at A6)...")
            excel_data = df_excel.values.tolist()
            self.write_data_to_range('3.Ins_Rec', 'A6', excel_data)

        # ── Summary ─────────────────────────────────────────────────────────────
        print("\n" + "=" * 70)
        print("3.INS_REC SHEET PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"   Mode          : {'Portfolio Copy (day >= 25)' if day >= 25 else 'Excel Import (day < 25)'}")
        print(f"   Total rows    : {num_rows}")
        print(f"   Date processed: {parsed_date.strftime('%d/%m/%Y')}")
        print("=" * 70)

        return num_rows

    def insurance_premium(self, insurance_premium_path: str) -> int:
        """
        Process insurance premium data:
        1. Read Insurance Premium Excel file
        2. Remove all blank rows
        3. Map columns: CONTRACT NUMBER -> CONTRACT, POLICY START DATE -> RENEWAL DATE, TOTAL PREMIUM -> TOTAL PREMIUM
        4. Paste data to Worst_Data sheet (columns: CONTRACT, RENEWAL DATE, TOTAL PREMIUM)

        Args:
            insurance_premium_path: Path to Insurance Premium Excel file

        Returns:
            Number of rows processed
        """
        print("\n" + "=" * 70)
        print("PROCESSING INSURANCE PREMIUM DATA")
        print("=" * 70)

        # Step 1: Read Insurance Premium Excel data
        print("\n1. Reading Insurance Premium Excel data...")
        df_insurance, num_rows, num_cols = self.read_excel_data(insurance_premium_path)

        # Step 2: Remove blank rows
        print("\n2. Removing blank rows...")
        df_insurance = df_insurance.dropna(how='all')
        rows_after_cleanup = len(df_insurance)
        print(f"   Removed {num_rows - rows_after_cleanup} blank rows")

        # Step 3: Map and select required columns
        print("\n3. Mapping columns...")

        # Strip whitespace from all column names for easier matching
        df_insurance.columns = df_insurance.columns.str.strip()

        # Map: CONTRACT NUMBER -> CONTRACT, POLICY START DATE -> RENEWAL DATE, TOTAL PREMIUM -> TOTAL PREMIUM
        required_cols = ['CONTRACT NUMBER', 'POLICY START DATE', 'TOTAL PREMIUM']

        # Check if required columns exist
        missing_cols = [col for col in required_cols if col not in df_insurance.columns]

        if missing_cols:
            print(f"   Warning: Missing columns: {missing_cols}")
            print(f"   Available columns: {list(df_insurance.columns)}")

        # Select and rename columns
        df_mapped = df_insurance[required_cols].copy()
        df_mapped.columns = ['CONTRACT', 'RENEWAL DATE', 'TOTAL PREMIUM']

        # Step 4: Write data to Worst_Data sheet
        print("\n4. Writing data to Worst_Data sheet (CONTRACT, RENEWAL DATE, TOTAL PREMIUM columns)...")
        # Convert DataFrame to list of lists
        insurance_data = df_mapped.values.tolist()

        # Find the target columns in Worst_Data sheet (you may need to adjust the starting cell)
        # Assuming we need to find the columns or write to specific column positions
        # For now, writing to columns G, H, I (adjust as needed based on actual sheet structure)
        self.write_data_to_range('Worst_Data', 'G3', insurance_data)

        print("\n" + "=" * 70)
        print("INSURANCE PREMIUM PROCESSING COMPLETED!")
        print("=" * 70)
        print(f"Total rows processed: {rows_after_cleanup}")
        print(f"Data written to Worst_Data sheet starting at G3")
        print("=" * 70)

        return rows_after_cleanup

    def process_complete_automation(self, csv_path: str, excel_path: str, dpd_month_path: str,
                                     allconsec_path: str , yardListPath: str, excel_path_mediation: str, legal_cases: str,
                                    excel_path_back_end: str, MATCAP: str, MATINT: str, moi_adjustment: str,
                                     int_receivable: str, insurane_Premium: str, output_path: str = None,
                                      date_str: str = None,last_pf: str = None) -> None:
        """
        Complete automation process:
        1. Process Portfolio sheet with CSV data
        2. Process 1.DPD sheet with Excel data
        3. Save as "Updated ECL Portfolio" file

        Args:
            csv_path: Path to CSV file for Portfolio sheet
            excel_path: Path to Excel file for 1.DPD sheet
            output_path: Output file path (if None, auto-generates)
            date_str: Date string for Portfolio sheet (format: M/D/YYYY)
        """
        print("\n" + "=" * 70)
        print("COMPLETE PORTFOLIO AUTOMATION")
        print("=" * 70)

        # Process Portfolio sheet
        portfolio_rows = self.process_portfolio_sheet(csv_path, date_str)

        # Process DPD sheet
        dpd_rows = self.process_dpd_sheet(excel_path, date_str)

        # Process DPD Month sheet
        dpd_month_rows = self.process_NPA_sheet(dpd_month_path, last_pf)

        # Process 100% Provision sheet
        provision_rows = self.process_100_percent_sheet()
        # Process AllConSec sheet
        allconsec_rows = self.process_AllConSec_sheet(allconsec_path)

        # Process YardList sheet
        yardlist_rows = self.process_YARD_sheet(yardListPath)

        # Process Legal sheet
        
        legal_rows = self.process_Legal_sheet(excel_path_mediation, legal_cases)
        # Process Ins_Rec sheet
        insrec_rows = self.process_Ins_Rec_sheet(excel_path_back_end, date_str)

        # Process Worst_Data sheet
        worst_data_rows = self.process_worst_data_sheet(MATCAP, MATINT)

        # Process insurance premium file to Worst_Data sheet columns G, H, I
        insurance_premium_rows = self.insurance_premium(insurane_Premium)
        # Process Int_Receivable sheet
        intrec_rows = self.process_int_receivable_sheet(int_receivable)
        # Process MOI_Adjustment sheet
        moiadj_rows = self.process_moi_sheet(moi_adjustment)

        self.copy_formula_and_paste_values('Portfolio', 'AN2', 'AN4')
        self.copy_formula_and_paste_values('Portfolio', 'AT1', 'AT4')
        self.copy_formula_and_paste_values('Portfolio', 'AU1', 'AU4')

        self.process_YARDCalculation_sheet()

        self.color_column_text_black("Portfolio", 'AT')
        self.color_column_text_black("Portfolio", 'AU')
        # Generate output filename if not specified
        if output_path is None:
            base_name = os.path.basename(self.workbook_path)
            if base_name.startswith("1. Portfolio_4_ECL_-"):
                # Calculate last month's ending date
                today = self.processing_date
                date_part = today.strftime("%d.%m.%Y")
                output_path = f"1. Portfolio_4_ECL-{date_part}.xlsb"
            else:
                output_path = base_name.replace('.xlsb', '_1. Portfolio_4_ECL.xlsb')
        self.copy_formula_and_paste_values('Portfolio','BH1','BH4','A')
        self.color_column_text_black('Portfolio','BH')
        # Add completion tick to YardList sheet M1
        self.add_yardlist_completion_tick(date_str)

        # Save the file
        print("\n" + "=" * 70)
        print("SAVING WORKBOOK")
        print("=" * 70)
        # Remove any existing Portfolio output files before saving
        _save_dir = os.path.dirname(os.path.abspath(output_path)) if os.path.dirname(output_path) else os.getcwd()
        for _pat in ("1. Portfolio_4_ECL_-*.xlsb", "1. Portfolio_4_ECL-*.xlsb"):
            for _old in glob.glob(os.path.join(_save_dir, _pat)):
                if os.path.abspath(_old) != os.path.abspath(output_path):
                    try:
                        os.remove(_old)
                        print(f"   Removed old file: {os.path.basename(_old)}")
                    except Exception:
                        pass
        self.save_as(output_path)
        
        # Summary of operations
        print("\n" + "=" * 70)
        print("COMPLETE AUTOMATION FINISHED SUCCESSFULLY!")
        print("=" * 70)
        print(f"\nOutput file: {output_path}")
        print(f"\nPortfolio sheet:")
        print(f"  - Rows processed: {portfolio_rows}")
        print(f"  - Date in column A: {date_str}")
        print(f"  - CSV data in columns B-AB")
        print(f"\n1.DPD sheet:")
        print(f"  - Rows processed: {dpd_rows}")
        print(f"  - NPA Rows processed: {dpd_month_rows}")
        print(f"  - 100% Provision Rows processed: {provision_rows}")
        print(f"  - all contract sector Rows processed: {allconsec_rows}")
        print(f"  - Yard List Rows processed: {yardlist_rows}")
        print(f"  - Legal Cases Rows processed: {legal_rows}")
        print(f"  - Insurance Recovery Rows processed: {insrec_rows}")
        print(f"  - Worst Data Rows processed: {worst_data_rows}")
        print(f"  - Insurance Premium Rows processed: {insurance_premium_rows}")
        print(f"  - Int Receivable Rows processed: {intrec_rows}")
        print(f"  - MOI Adjustment Rows processed: {moiadj_rows}")
        print(f"\nAll other sheets and formatting preserved!")
        print("=" * 70)
        return output_path

    def ECL_calculation(self, excel_path_portfolio: str, output_path, overlayfile: str, date_str) -> None:
        """
        Perform ECL calculation by copying formulas and pasting values in specific ranges
        """
        print("\n" + "=" * 70)
        print("PERFORMING ECL CALCULATION")
        print("=" * 70)

        print("\n1. Clearing existing data in ECL sheet...")
        self.clear_range_dynamic('ECL', 'A3', 'N')
        self.clear_range_dynamic('ECL', 'AX3', 'AX')
        self.clear_range_dynamic('ECL', 'BF3', 'BH')
        #get portfolio sheet to DF
        self.copy_formula_and_paste_value_single("ECL", "BR26","BR27")
        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(excel_path_portfolio,skip_rows=2, num_columns=70, sheet_name='Portfolio',engine='pyxlsb')
        # filter and get NON YARDs
        df_excel = df_excel[df_excel['YARD STATUS'] == 'NON-YARD']
        num_rows = len(df_excel)
        # Process the DataFrame as needed
        # Extract the specified columns
        columns_to_extract = [
            'DATE',
            'CONTRACT NO',
            'CLIENT CODE',
            'CLIENT P/NP',
            'PD CATEGORY',
            'LEASE/LOAN',
            'EQT_DESC',
            'FUTURE CAPITAL',
            'TOTAL ARREARS',
            'GROSS EXPOSURE (GROSS-IIS) [CBSL]',
            'CLIENT DPD',
            'INTEREST RATE',
            'ACTIVATE DATE',
            'EXPIRY DATE'
        ]

        df_1 = df_excel[columns_to_extract]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('ECL', 'A3', excel_data)
        df_2 = df_excel[['100% PROVISION']]
        excel_data = df_2.values.tolist()
        self.write_data_to_range('ECL', 'AX3', excel_data)
        df_3 = df_excel[['SEC_DESC','SUB_SECTOR','PURPOSE']]
        excel_data = df_3.values.tolist()
        self.write_data_to_range('ECL', 'BF3', excel_data)

        df_excel_y, num_foo, num_cols = self.read_excel_data(excel_path_portfolio,skip_rows=2, num_columns=6, sheet_name='4.YardCalculation',engine='pyxlsb')
        Yard_provision = df_excel_y['YARD PROVISION'].sum()
        print(f'YardProvision:{Yard_provision}')
        self.clear_range('ECL', 'BR23')
        self.write_to_cell('ECL', 'BR23', Yard_provision)
        # Copy paste formulas
        end_row = 2 + num_rows  # Data starts at row 2, so end at 1 + num_rows
        if end_row > 3:  # Only copy if there are rows beyond row 2
            self.copy_formulas_to_range('ECL', 'O3:AW3', 3, end_row)
            self.copy_formulas_to_range('ECL', 'AY3:BE3', 3, end_row)
            self.copy_formulas_to_range('ECL', 'BI3:BM3', 3, end_row)
        else:
            print("no formulas to copy paste")

        self.clear_range_dynamic('ECL', f'A{end_row+1}', 'CA')

        if overlayfile.endswith('.xlsb'):
            engine = 'pyxlsb'
        else:
            engine = 'openpyxl'

        self.update_pivot_source_and_refresh('Pivot','PivotTable1','ECL', "A2", 'BM')
        self.update_pivot_source_and_refresh('Pivot','PivotTable2','ECL', "A2", 'BM')
        self.update_pivot_source_and_refresh('Pivot','PivotTable7','ECL', "A2", 'BM')

        OL_df, num_rows, num_cols = self.read_excel_data(overlayfile,skip_rows=1, num_columns=4, sheet_name='OL',engine=engine)
        olval = OL_df.values.tolist()
        self.app.api.EnableEvents = False  # suppress VBA worksheet Change event / MsgBox on OL write
        self.write_data_to_range('OL', 'A4', olval)
        self.app.api.EnableEvents = True
        #run the excel macro named "ECL_Calculation"
        self.app.display_alerts = False
        self.app.screen_updating = False
        
        adj = OLAdjustment(self.workbook)
        adj.run()
        # self.workbook.macro('OLAdjstmnt')()
        
        if output_path is None:
            base_name = os.path.basename(self.workbook_path)
            if base_name.startswith("2. ECL Model_"):
                date = self.normalize_date(date_str)
                # date_part = today.strftime("%d.%m.%Y")
                output_path = f"2. ECL Model_{date}.xlsb"
            else:
                output_path = base_name.replace('.xlsb', '2. ECL Model_2025-08-18.xlsb')
        # Remove any existing ECL Model output files before saving
        _save_dir = os.path.dirname(os.path.abspath(output_path)) if os.path.dirname(output_path) else os.getcwd()
        for _old in glob.glob(os.path.join(_save_dir, "2. ECL Model_*.xlsb")):
            if os.path.abspath(_old) != os.path.abspath(output_path):
                try:
                    os.remove(_old)
                    print(f"   Removed old file: {os.path.basename(_old)}")
                except Exception:
                    pass
        self.save_as(output_path)

        
        return output_path

    def Summary_calculation(self, excel_path_portfolio: str, output_path, ecl_file_path, net_portfolio, related_party, cader, dp_file, date_str: str = None, gl_closing_balance: float = None, loan_capital_balance: float = None) -> None:
        """
        Perform ECL calculation by copying formulas and pasting values in specific ranges
        """
        print("\n" + "=" * 70)
        print("PERFORMING Summary CALCULATION")
        print("=" * 70)
        rowcountSum = self.get_row_count('SUMMARY','A',3)
        print("\n1. Clearing existing data in SUMMARY sheet...")
        self.clear_range_dynamic('SUMMARY', 'A4', 'C')
        self.clear_range_dynamic('SUMMARY', 'E4', 'G')
        self.clear_range_dynamic('SUMMARY', 'K4', 'L')
        self.clear_range_dynamic('SUMMARY', 'N4', 'O')
        self.clear_range_dynamic('SUMMARY', 'R4', 'R')
        self.clear_range_dynamic('SUMMARY', 'T4', 'U')
        self.clear_range_dynamic('SUMMARY', 'AB4', 'AC')
        self.clear_range_dynamic('ECL Provision', 'A3', 'H')
        self.clear_range_dynamic('RPT', 'A5', 'D')
        #get portfolio sheet to DF

        print("\n2. Reading Excel data...")
        df_excel, num_rows, num_cols = self.read_excel_data(excel_path_portfolio,skip_rows=2, num_columns=70, sheet_name='Portfolio',engine='pyxlsb')
        p_num_rows = num_rows
        # get the row count
        
        #rows to insert
        rowstoinsert = num_rows - rowcountSum

        ## insert rows 
        self.insert_rows("SUMMARY", rowcountSum + 1, rowstoinsert)
        # Process the DataFrame as needed
        # Extract the specified columns
        columns_to_extract_1 = [
            'CONTRACT NO',
            'CLIENT CODE',
            'BRANCH'
        ]

        df_1 = df_excel[columns_to_extract_1]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('SUMMARY', 'A4', excel_data)

        columns_to_extract_1 = [
            'GROSS OUTSTANDING',
            'CLIENT DPD',
            'CLIENT P/NP'
        ]

        df_1 = df_excel[columns_to_extract_1]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('SUMMARY', 'E4', excel_data)
        
        columns_to_extract_1 = [
            'YARD STATUS',
            'LEASE/LOAN'
        ]

        df_1 = df_excel[columns_to_extract_1]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('SUMMARY', 'K4', excel_data)

        columns_to_extract_1 = [
            'ARREARS',
            'EQT_DESC'
        ]

        df_1 = df_excel[columns_to_extract_1]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('SUMMARY', 'N4', excel_data)

        columns_to_extract_1 = [
            'PD CATEGORY'
        ]

        df_1 = df_excel[columns_to_extract_1]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('SUMMARY', 'R4', excel_data)

        columns_to_extract_1 = [
            'IIS',
            'GROSS EXPOSURE (GROSS-IIS) [CBSL]'
        ]

        df_1 = df_excel[columns_to_extract_1]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('SUMMARY', 'T4', excel_data)

        columns_to_extract_1 = [
            'FUTURE CAPITAL',
            'TOTAL ARREARS'
        ]

        df_1 = df_excel[columns_to_extract_1]
        excel_data = df_1.values.tolist()
        self.write_data_to_range('SUMMARY', 'AB4', excel_data)

        df_excel, num_rows, num_cols = self.read_excel_data(ecl_file_path,skip_rows=1, num_columns=8, sheet_name='toSummary',engine='pyxlsb')
        self.clear_range_dynamic('ECL Provision', 'A3', 'H')
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('ECL Provision', 'A3', excel_data)
        self.copy_formulas_to_range('ECL Provision', 'J3:K3', 3, num_rows+3)

        df_excel, num_rows, num_cols = self.read_excel_data(excel_path_portfolio,skip_rows=2, num_columns=6, sheet_name='4.YardCalculation',engine='pyxlsb')
        with _DialogDismisser(["missing data", "yard provision", "microsoft excel"], interval=0.2):
            self.app.api.DisplayAlerts = False
            self.app.api.EnableEvents = False
            self.clear_range_dynamic('Yard Provision', 'A3', 'F')
            excel_data = df_excel.values.tolist()
            self.write_data_to_range('Yard Provision', 'A4', excel_data)
            self.app.api.EnableEvents = True
            self.app.api.DisplayAlerts = False
        print("Yard Provision Copy Pasted Successfully")
        df_portfolio, num_rows_pf, num_cols_pf = self.read_excel_data(
            excel_path_portfolio, skip_rows=2, num_columns=11, sheet_name='Portfolio', engine='pyxlsb'
        )
        # Row 3 is the header (skip_rows=2 skips rows 1-2), so df_portfolio row 0 = Excel row 3
        # Column K is index 10 (0-based) — verify header name matches
        future_interest_col = df_portfolio.columns[10]
        print(f"   Column K header: '{future_interest_col}'")
        future_interest_sum = df_portfolio[future_interest_col].sum() / 1000
        print(f"future interest Sum : {future_interest_sum}")
        sheet = self.workbook.sheets['NPA']
        sheet.range('D39').value = future_interest_sum

        df_excel, num_rows, num_cols = self.read_excel_data(related_party,skip_rows=3, num_columns=6, sheet_name='Related Parties - All',engine='pyxlsb')
        excel_data = df_excel[['Client code','Name','NIC/PP','Relationship Type_1']]
        excel_data = excel_data.values.tolist()
        self.write_data_to_range('RPT', 'A5', excel_data)

        df_excel, num_rows, num_cols = self.read_excel_data(dp_file,skip_rows=0, num_columns=2, sheet_name='DP',engine='openpyxl')
        excel_data = df_excel
        excel_data = excel_data.values.tolist()
        self.clear_range_dynamic('DP', 'A2', 'B')
        self.write_data_to_range('DP', 'A2', excel_data)

        df_excel, num_rows, num_cols = self.read_excel_data(cader, num_columns=7, sheet_name='Sheet1')
        self.clear_range_dynamic('Staff', 'A5', 'G')
        excel_data = df_excel.values.tolist()
        self.write_data_to_range('Staff', 'A5', excel_data)

        # Write pre-fetched selenium values to NPA sheet in the Summary workbook
        npa_sheet = self.workbook.sheets['NPA']
        if gl_closing_balance is not None:
            npa_sheet.range('I12').value = gl_closing_balance
            print(f"   [OK] Closing balance written to NPA I12: {gl_closing_balance}")
        else:
            print("   [!] No GL closing balance provided — I12 not updated")
        if loan_capital_balance is not None:
            npa_sheet.range('J12').value = loan_capital_balance
            print(f"   [OK] Capital balance written to NPA J12: {loan_capital_balance}")
        else:
            print("   [!] No loan capital balance provided — J12 not updated")

        # Copy paste formulas
        end_row = 2 + p_num_rows  # Data starts at row 2, so end at 1 + num_rows
        if end_row > 3:  # Only copy if there are rows beyond row 2
            self.copy_formulas_to_range('SUMMARY', 'D4:D4', 4, end_row+1)
            self.copy_formulas_to_range('SUMMARY', 'H4:J4', 4, end_row+1)
            self.copy_formulas_to_range('SUMMARY', 'M4:M4', 4, end_row+1)
            self.copy_formulas_to_range('SUMMARY', 'P4:Q4', 4, end_row+1)
            self.copy_formulas_to_range('SUMMARY', 'S4:S4', 4, end_row+1)
            self.copy_formulas_to_range('SUMMARY', 'V4:AA4', 4, end_row+1)
            self.copy_formulas_to_range('SUMMARY', 'AD4:AG4', 4, end_row+1)
        else:
            print("no formulas to copy paste")

        foo = self.filter_copy_delete(
                                        source_sheet_name="SUMMARY",
                                        target_sheet_name="ZeroValues",
                                        header_row=3,
                                        filter_column="E",
                                        filter_criteria='-',
                                        paste_start_cell="B2"
                                    )

        self.update_pivot_source_and_refresh('Pivot','PivotTable1','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Pivot','PivotTable2','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Pivot','PivotTable3','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Pivot','PivotTable4','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Pivot','PivotTable5','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('NPA','PivotTable1','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Check','PivotTable1','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('PBI_Data','PivotTable3','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('JEs','PivotTable1','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('JEs','PivotTable2','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable8','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable7','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable4','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable6','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable1','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable3','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable2','SUMMARY', "A3", 'AU')
        self.update_pivot_source_and_refresh('Other Pivots','PivotTable5','SUMMARY', "A3", 'AU')

        self.write_to_cell('SUMMARY','D2', self.calculate_sum_range('SUMMARY','D4','D','A3','down'))
        self.write_to_cell('SUMMARY','E2', self.calculate_sum_range('SUMMARY','E4','E','A3','down'))
        self.write_to_cell('SUMMARY','F2', self.calculate_sum_range('SUMMARY','F4','F','A3','down'))
        self.write_to_cell('SUMMARY','H2', self.calculate_sum_range('SUMMARY','H4','H','A3','down'))
        self.write_to_cell('SUMMARY','J2', self.calculate_sum_range('SUMMARY','J4','J','A3','down'))
        self.write_to_cell('SUMMARY','N2', self.calculate_sum_range('SUMMARY','N4','N','A3','down'))
        self.write_to_cell('SUMMARY','T2', self.calculate_sum_range('SUMMARY','T4','T','A3','down'))
        self.write_to_cell('SUMMARY','U2', self.calculate_sum_range('SUMMARY','U4','U','A3','down'))
        self.write_to_cell('SUMMARY','W2', self.calculate_sum_range('SUMMARY','W4','W','A3','down'))
        self.write_to_cell('SUMMARY','X2', self.calculate_sum_range('SUMMARY','X4','X','A3','down'))
        self.write_to_cell('SUMMARY','Y2', self.calculate_sum_range('SUMMARY','Y4','Y','A3','down'))
        self.write_to_cell('SUMMARY','Z2', self.calculate_sum_range('SUMMARY','Z4','Z','A3','down'))
        self.write_to_cell('SUMMARY','AB2', self.calculate_sum_range('SUMMARY','AB4','AB','A3','down'))
        self.write_to_cell('SUMMARY','AC2', self.calculate_sum_range('SUMMARY','AC4','AC','A3','down'))

        if output_path is None:
            workbook_dir = os.path.dirname(os.path.abspath(self.workbook_path))
            if date_str:
                d = None
                for _fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
                    try:
                        d = datetime.strptime(date_str, _fmt)
                        break
                    except ValueError:
                        continue
                date_suffix = d.strftime("%d.%m.%Y") if d else date_str.replace("/", ".")
                output_path = os.path.join(workbook_dir, f"3. Summary_{date_suffix}.xlsb")
            else:
                base_name = os.path.basename(self.workbook_path)
                output_path = os.path.join(workbook_dir, re.sub(r'_\d{4}-\d{2}-\d{2}', '', base_name))
        # Remove any existing Summary output files before saving
        _save_dir = os.path.dirname(os.path.abspath(output_path))
        for _old in glob.glob(os.path.join(_save_dir, "3. Summary_*.xlsb")):
            if os.path.abspath(_old) != os.path.abspath(output_path):
                try:
                    os.remove(_old)
                    print(f"   Removed old file: {os.path.basename(_old)}")
                except Exception:
                    pass
        self.save_as(output_path)


def get_highest_version_date(completed_dates_file: str, file_category: str) -> Optional[str]:
    """
    Read completed_dates.txt and return the date with the highest version number.
    Example: If file contains "31-07-2025,31-07-2025(1),31-07-2025(2)", returns "31-07-2025(2)"

    Args:
        completed_dates_file: Path to the completed_dates.txt file

    Returns:
        The date string with highest version, or None if file doesn't exist
    """
    completed_dates_file = os.path.join(completed_dates_file, file_category, "completed_dates.txt")
    if not os.path.exists(completed_dates_file):
        print(f"Warning: {completed_dates_file} not found. Using default date.")
        return None

    try:
        with open(completed_dates_file, 'r') as f:
            content = f.read().strip()

        if not content:
            print("Warning: completed_dates.txt is empty.")
            return None

        # Split by comma and strip whitespace
        dates = [d.strip() for d in content.split(',')]

        if not dates:
            return None

        # Parse dates to find the one with the highest version
        # Format: "DD-MM-YYYY" or "DD-MM-YYYY(N)"
        date_versions = []
        for date_str in dates:
            # Extract base date and version number
            match = re.match(r'(\d{2}-\d{2}-\d{4})(?:\((\d+)\))?', date_str)
            if match:
                base_date = match.group(1)
                version = int(match.group(2)) if match.group(2) else 0
                date_versions.append((base_date, version, date_str))

        if not date_versions:
            return dates[0]  # Return first date if no valid format found

        # Sort by base date and version, get the highest
        date_versions.sort(key=lambda x: (x[0], x[1]), reverse=True)
        highest_version_date = date_versions[0][2]

        print(f"Selected date from completed_dates.txt: {highest_version_date}")
        return highest_version_date

    except Exception as e:
        print(f"Error reading completed_dates.txt: {e}")
        return None


def remove_xlsb_files(folder_path: str) -> None:
    """
    Remove all .xlsb files from the specified folder.

    Args:
        folder_path: Path to the folder where .xlsb files should be removed
    """
    print("\n" + "=" * 70)
    print("REMOVING EXISTING .XLSB FILES FROM ROOT FOLDER")
    print("=" * 70)

    try:
        xlsb_files = glob.glob(os.path.join(folder_path, "*.xlsb"))

        if not xlsb_files:
            print("No .xlsb files found to remove.")
            return

        removed_count = 0
        for file_path in xlsb_files:
            try:
                os.remove(file_path)
                print(f"Removed: {os.path.basename(file_path)}")
                removed_count += 1
            except Exception as e:
                print(f"Error removing {os.path.basename(file_path)}: {e}")

        print(f"\nTotal .xlsb files removed: {removed_count}")
        print("=" * 70)

    except Exception as e:
        print(f"Error scanning for .xlsb files: {e}")
  


def transfer_xlsb_files(source_folder: str, target_date: str, file_category: str = "NBD_MF_23_C4") -> None:
    """
    Transfer all .xlsb files from source folder to the appropriate output folder.

    Args:
        source_folder: Path to the folder containing .xlsb files
        target_date: Date string in format DD-MM-YYYY or DD-MM-YYYY(N)
        file_category: Category folder name (default: "NBD_MF_23_C4")
    """
    print("\n" + "=" * 70)
    print("TRANSFERRING .XLSB FILES TO OUTPUT FOLDER")
    print("=" * 70)

    try:
        # Find all .xlsb files in source folder
        xlsb_files = glob.glob(os.path.join(source_folder, "*.xlsb"))

        if not xlsb_files:
            print("No .xlsb files found to transfer.")
            return

        # Create target directory path
        target_dir = os.path.join(OUTPUTS_FOLDER, file_category, target_date)

        # Create target directory if it doesn't exist
        os.makedirs(target_dir, exist_ok=True)
        print(f"Target directory: {target_dir}")

        transferred_count = 0
        for file_path in xlsb_files:
            try:
                file_name = os.path.basename(file_path)
                target_path = os.path.join(target_dir, file_name)

                # Copy file to target directory
                shutil.copy2(file_path, target_path)
                print(f"Transferred: {file_name}")
                transferred_count += 1

            except Exception as e:
                print(f"Error transferring {os.path.basename(file_path)}: {e}")

        print(f"\nTotal .xlsb files transferred: {transferred_count}")
        print("=" * 70)

    except Exception as e:
        print(f"Error during file transfer: {e}")


def main():
    """
    Main execution function
    """
    
    Daily = os.environ.get("IFRS_DAILY_FOLDER", r"C:\Users\overdime\Documents\Scheduler\Weekly and Daily\Supporting\Daily")
    csv_file = None

    # Helper function to find file by glob pattern (returns most recent if multiple matches)
    def find_file(pattern):
        matches = glob.glob(pattern)
        if not matches:
            return None
        # Return the most recently modified file if multiple matches
        return max(matches, key=os.path.getmtime)

    # File paths - all combined with Daily folder path (using wildcards for dates)
    # Match both "1. Portfolio_4_ECL_-*.xlsb" and "1. Portfolio_4_ECL-*.xlsb"
    xlsb_file = (find_file(os.path.join(Daily, r"1. Portfolio_4_ECL_-*.xlsb"))
                 or find_file(os.path.join(Daily, r"1. Portfolio_4_ECL-*.xlsb")))
    xlsb_file_last = (find_file(os.path.join(Daily, r"Input\1. Portfolio_4_ECL_-*.xlsb"))
                      or find_file(os.path.join(Daily, r"Input\1. Portfolio_4_ECL-*.xlsb")))
    
    csv_file = find_file(os.path.join(Daily, r"Input\PORTFOLIO REPORT NEW-*.csv"))

    if csv_file is None:
        csv_file = find_file(os.path.join(Daily, r"Input\PORTFOLIO REPORT NEW-*.xlsx"))

    xlsx_dpd_file = find_file(os.path.join(Daily, r"Input\Month DPD Report New *.xlsx"))
    dpd_month_file = find_file(os.path.join(Daily, r"Input\Month DPD Report New *.xlsx"))
    output_file = None  # Will auto-generate as "Updated ECL Portfolio-<date>.xlsb"
    allconsec_path = find_file(os.path.join(Daily, r"Input\All Contracts Sector Report - Internal *.xlsx"))
    yardListPath = find_file(os.path.join(Daily, r"Input\Yard Final - *.xlsx"))
    legal_cases = find_file(os.path.join(Daily, r"Input\Litigation Report - *.xlsx"))
    excel_path_mediation = find_file(os.path.join(Daily, r"Input\Mediation Detail Report - *.xlsx"))
    excel_path_back_end = find_file(os.path.join(Daily, r"Input\Backend Due Types - Contract wise Report - Internal *.xlsx"))
    MATINT = find_file(os.path.join(Daily, r"Input\MATINT-*.csv"))
    MATCAP = find_file(os.path.join(Daily, r"Input\MATCAP-*.csv"))
    moi_adjustment = find_file(os.path.join(Daily, r"Input\Moratorium Contracts details-*.xlsx"))
    int_receivable = find_file(os.path.join(Daily, r"Input\2128-*.xlsx"))
    ecl_workbook_path = find_file(os.path.join(Daily, r"2. ECL Model_*.xlsb"))
    summary_workbook_path = find_file(os.path.join(Daily, r"3. Summary_*.xlsb"))
    related_party = find_file(os.path.join(Daily, r"Input\Assetline Finance Limited - Related Party Register*.xlsb"))
    carder = find_file(os.path.join(Daily, r"Input\Cadre *.xlsx"))
    insurane_Premium = find_file(os.path.join(Daily, r"Input\Insurance Premium *.xlsx"))
    overlay_file = find_file(os.path.join(Daily, r"Input\Overlay-*.xlsx"))
    dp_file = find_file(os.path.join(Daily, r"Input\1611-*.xlsx"))
    # Normalise date to DD/MM/YYYY regardless of what ifrs.py passes in
    _raw_date = sys.argv[1] if len(sys.argv) > 1 else datetime.today().strftime("%d/%m/%Y")
    if isinstance(_raw_date, str):
        for _fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                date_str = datetime.strptime(_raw_date.strip(), _fmt).strftime("%d/%m/%Y")
                break
            except ValueError:
                continue
        else:
            print(f"   Warning: unrecognised date format '{_raw_date}', using today.")
            date_str = datetime.today().strftime("%d/%m/%Y")
    else:
        date_str = _raw_date.strftime("%d/%m/%Y")

    # Check if files exist
    if xlsb_file is None or not os.path.exists(xlsb_file):
        print(f"Error: File not found - Pattern: 1. Portfolio_4_ECL_-*.xlsb in {Daily}")
        return

    if csv_file is None or not os.path.exists(csv_file):
        print(f"Error: File not found - Pattern: PORTFOLIO REPORT NEW-*.csv in {Daily}/Input")
        return

    if xlsx_dpd_file is None or not os.path.exists(xlsx_dpd_file):
        print(f"Error: File not found - Pattern: Month DPD Report New *.xlsx in {Daily}/Input")
        return

    if dpd_month_file is None or not os.path.exists(dpd_month_file):
        print(f"Error: File not found - Pattern: Month DPD Report New *.xlsx in {Daily}/Input")
        return

    if allconsec_path is None or not os.path.exists(allconsec_path):
        print(f"Error: File not found - Pattern: All Contracts Sector Report - Internal *.xlsx in {Daily}/Input")
        return

    if yardListPath is None or not os.path.exists(yardListPath):
        print(f"Error: File not found - Pattern: Yard Final - *.xlsx in {Daily}/Input")
        return

    if legal_cases is None or not os.path.exists(legal_cases):
        print(f"Error: File not found - Pattern: Litigation Report - *.xlsx in {Daily}/Input")
        return

    if excel_path_mediation is None or not os.path.exists(excel_path_mediation):
        print(f"Error: File not found - Pattern: Mediation Detail Report - *.xlsx in {Daily}/Input")
        return

    if excel_path_back_end is None or not os.path.exists(excel_path_back_end):
        print(f"Error: File not found - Pattern: Backend Due Types - Contract wise Report - Internal *.xlsx in {Daily}/Input")
        return

    if MATCAP is None or not os.path.exists(MATCAP):
        print(f"Error: File not found - Pattern: MATCAP-*.csv in {Daily}/Input")
        return

    if MATINT is None or not os.path.exists(MATINT):
        print(f"Error: File not found - Pattern: MATINT-*.csv in {Daily}/Input")
        return

    if moi_adjustment is None or not os.path.exists(moi_adjustment):
        print(f"Error: File not found - Pattern: Moratorium Contracts details-*.xlsx in {Daily}/Input")
        return

    if int_receivable is None or not os.path.exists(int_receivable):
        print(f"Error: File not found - Pattern: 2128-*.xlsx in {Daily}/Input")
        return

    if insurane_Premium is None or not os.path.exists(insurane_Premium):
        print(f"Error: File not found - Pattern: Insurance Premium *.xlsx in {Daily}/Input")
        return

    if ecl_workbook_path is None or not os.path.exists(ecl_workbook_path):
        print(f"Error: File not found - Pattern: 2. ECL Model_*.xlsb in {Daily}")
        return

    if related_party is None or not os.path.exists(related_party):
        print(f"Error: File not found - Pattern: 2. Assett line Finance - Related Party Register*.xlsb in {Daily}/Input")
        return
    # Run selenium bots before any Excel workbook opens (avoids COM conflicts)
    print("\nRunning GL CodeBase Report to fetch closing balance for NPA I12...")
    gl_closing_balance = None
    try:
        import GL_CodeBase_Report
        closing_balance_raw = GL_CodeBase_Report.run(date_str)

        if closing_balance_raw is not None and closing_balance_raw != '':
            if isinstance(closing_balance_raw, str):
                gl_closing_balance = float(closing_balance_raw.replace(',', ''))
            else:
                gl_closing_balance = float(closing_balance_raw)
            print(f"   [OK] GL closing balance fetched: {gl_closing_balance}")
        else:
            print("   [!] GL CodeBase Report returned no value — I12 will not be updated")
    except Exception as e:
        print(f"    GL CodeBase Report failed: {e}")

    print("\nRunning Loan Base Report to fetch capital balance for NPA J12...")
    loan_capital_balance = None
    try:
        import Loan_Base_Report
        loan_capital_balance = Loan_Base_Report.run(date_str)
        if loan_capital_balance is not None:
            print(f"   [OK] Loan capital balance fetched: {loan_capital_balance}")
        else:
            print("   [!] Loan Base Report returned no value — J12 will not be updated")
    except Exception as e:
        print(f"    Loan Base Report failed: {e}")

    # Step 1: Remove all .xlsb files from root folder at the start (unless caller disables this)
    if os.environ.get("IFRS_SKIP_ROOT_XLSB_CLEANUP") != "1":
        remove_xlsb_files(ROOT_FOLDER)
    else:
        print("Skipping root .xlsb cleanup (IFRS_SKIP_ROOT_XLSB_CLEANUP=1).")

    # Step 2: Get the highest version date from completed_dates.txt
    target_date = get_highest_version_date(OUTPUTS_FOLDER, "NBD-QF-23-C3-C10-SF-FI")
    target_date1 = get_highest_version_date(OUTPUTS_FOLDER, "NBD_MF_23_C4")

    # If no date found in file, use default date format
    if target_date is None:
        today = datetime.today()
        last_month_end = (today.replace(day=1) - pd.Timedelta(days=1))
        target_date = last_month_end.strftime("%d-%m-%Y")
        print(f"Using default date: {target_date}")



    with ExcelPortfolioAutomation(xlsb_file, visible=False) as automation:
        newpffile = automation.process_complete_automation(csv_file, xlsx_dpd_file,dpd_month_file,allconsec_path,
                                               yardListPath,excel_path_mediation, legal_cases,excel_path_back_end,
                                               MATCAP,MATINT,moi_adjustment,int_receivable,insurane_Premium, output_file, date_str,xlsb_file_last)
    print("\nPortfolio Automation completed. Excel closed.")

    # Perform ECL calculation

    with ExcelPortfolioAutomation(ecl_workbook_path, visible=False) as automation:
        newECLfile = automation.ECL_calculation(newpffile, output_file, overlay_file, date_str)
    print("\nECL Calculation completed. Excel closed.")

    with ExcelPortfolioAutomation(summary_workbook_path, visible=False) as automation:
        automation.Summary_calculation(newpffile, output_file,newECLfile, csv_file, related_party, carder, dp_file, date_str, gl_closing_balance, loan_capital_balance)
        print("\nSummary Calculation completed. Excel closed.")

    # Step 3: Transfer all .xlsb files to the output folder at the end (unless caller disables this)
    if os.environ.get("IFRS_SKIP_TRANSFER") != "1":
        transfer_xlsb_files(ROOT_FOLDER, target_date, file_category="NBD-QF-23-C3-C10-SF-FI")
        transfer_xlsb_files(ROOT_FOLDER, target_date1, file_category="NBD_MF_23_C4")
    else:
        print("Skipping transfer step (IFRS_SKIP_TRANSFER=1).")
        

if __name__ == "__main__":
    _base_dir = Path(__file__).resolve().parent.parent
    _ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Honour log dir injected by the UI (IFRS_LOG_DIR); fall back to auto-timestamped folder
    _log_dir = Path(os.environ.get("IFRS_LOG_DIR", "") or (_base_dir / "logs" / f"run_{_ts}"))
    _log_dir.mkdir(parents=True, exist_ok=True)
    _log_file = _log_dir / "portfolio_automation_xlwings.log"

    _real_stdout = sys.stdout
    _real_stderr = sys.stderr

    class _Tee:
        """Write to both the log file and the original stdout so the UI pipe stays live."""
        def __init__(self, fh, passthrough):
            self._fh = fh
            self._pt = passthrough
        def write(self, s):
            self._fh.write(s)
            self._fh.flush()
            self._pt.write(s)
            self._pt.flush()
        def flush(self):
            self._fh.flush()
            self._pt.flush()

    with open(_log_file, "w", encoding="utf-8") as _fh:
        _fh.write(f"Portfolio Automation — started {datetime.now().isoformat()}\n")
        _fh.write("-" * 60 + "\n")
        _fh.flush()
        sys.stdout = _Tee(_fh, _real_stdout)
        sys.stderr = _Tee(_fh, _real_stderr)
        try:
            main()
        except Exception:
            traceback.print_exc()
            raise
        finally:
            sys.stdout = _real_stdout
            sys.stderr = _real_stderr

    print(f"Log written: {_log_file}")
