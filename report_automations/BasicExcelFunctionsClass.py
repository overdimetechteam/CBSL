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
import pandas as pd
import xlwings as xw
from datetime import datetime
import os
import glob
import shutil
from typing import Optional, Tuple, List
import re
import calendar


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
        self.workbook = self.app.books.open(self.workbook_path)
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

    def read_excel_data(self, excel_path: str, num_columns: int = 5, skip_rows: int = 0, sheet_name: str = None, engine='openpyxl') -> Tuple[pd.DataFrame, int, int]:
        """
        Read data from Excel file

        Args:
            excel_path: Path to the Excel file
            num_columns: Number of columns to read (default: 5 for A-E)

        Returns:
            Tuple of (DataFrame, num_rows, num_cols)
        """
        print(f"   Reading Excel file: {excel_path}")
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
        # Paste to each target row (starting from the row after last data row)
        if target_start_row >= target_end_row:
            print("   No rows to copy formulas to (target start row is after target end row)")
        else:
            for row in range(target_start_row, target_end_row + 1):
                target_range = f"{start_col}{row}:{end_col}{row}"
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

        # Save as new file
        self.workbook.save(full_path)
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

    def refresh_pivot_table(self, sheet_name: str, pivot_table_name: str) -> None:
        """
        Refresh a specific pivot table in a worksheet

        Args:
            sheet_name: Name of the worksheet containing the pivot table
            pivot_table_name: Name of the pivot table to refresh
        """
        try:
            print(f"   Refreshing pivot table '{pivot_table_name}' in sheet '{sheet_name}'...")

            sheet = self.workbook.sheets[sheet_name]

            # Access the pivot table through Excel API
            pivot_table = sheet.api.PivotTables(pivot_table_name)

            # Refresh the pivot table
            pivot_table.RefreshTable()

            print(f"    Pivot table '{pivot_table_name}' refreshed successfully")

        except Exception as e:
            print(f"    Error refreshing pivot table: {e}")
            raise

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

    def read_sheet_range_to_dataframe(self, sheet_name: str, range_address: str = None) -> pd.DataFrame:
        """
        Read data from a specific range in a sheet to a pandas DataFrame
        If range_address is None, reads the entire used range

        Args:
            sheet_name: Name of the worksheet to read from
            range_address: Range to read (e.g., 'A1:Z100'). If None, reads used range

        Returns:
            pd.DataFrame: DataFrame containing the data from the range
        """
        print(f"   Reading data from sheet '{sheet_name}'...")

        sheet = self.workbook.sheets[sheet_name]

        if range_address:
            # Read specific range
            data = sheet.range(range_address).value
            print(f"    Read range {range_address}")
        else:
            # Read entire used range
            used_range = sheet.used_range
            data = used_range.value
            range_address = used_range.address
            print(f"    Read used range {range_address}")

        # Convert to DataFrame
        if data:
            # If data is a list of lists, convert to DataFrame
            if isinstance(data, list) and len(data) > 0:
                # Use first row as headers
                df = pd.DataFrame(data[1:], columns=data[0])
            else:
                # Single cell value
                df = pd.DataFrame([data])

            print(f"    Successfully read {len(df)} rows and {len(df.columns)} columns")
            return df
        else:
            print(f"    No data found in range")
            return pd.DataFrame()

    def write_dataframe_to_sheet(self, sheet_name: str, start_cell: str, df: pd.DataFrame,
                                  include_headers: bool = True, clear_existing: bool = False) -> None:
        """
        Write a pandas DataFrame to a specific location in a sheet

        Args:
            sheet_name: Name of the worksheet to write to
            start_cell: Starting cell for writing (e.g., 'A1')
            df: DataFrame to write
            include_headers: Whether to include column headers (default: True)
            clear_existing: Whether to clear existing data in the target range (default: False)
        """
        print(f"   Writing DataFrame to sheet '{sheet_name}' starting at {start_cell}...")

        sheet = self.workbook.sheets[sheet_name]

        # Prepare data
        if include_headers:
            # Include headers
            data = [df.columns.tolist()] + df.values.tolist()
        else:
            # Data only
            data = df.values.tolist()

        # Calculate target range if we need to clear
        if clear_existing and data:
            num_rows = len(data)
            num_cols = len(data[0]) if data else 0

            # Calculate end cell by offsetting from start cell
            end_cell = sheet.range(start_cell).offset(num_rows - 1, num_cols - 1)
            end_cell_address = end_cell.get_address(False, False)

            # Clear the range
            clear_range = f"{start_cell}:{end_cell_address}"
            print(f"   Clearing existing data in range {clear_range}...")
            sheet.range(clear_range).clear_contents()

        # Write data
        if data:
            sheet.range(start_cell).value = data
            print(f"    Successfully wrote {len(df)} rows and {len(df.columns)} columns")
        else:
            print(f"    No data to write")

    def delete_rows_after_last_data(self, sheet_name: str, check_column: str = 'A',
                                     start_row: int = 2, max_delete_rows: int = 1000000) -> None:
        """
        Delete all empty rows after the last row with data in a sheet
        This helps clean up sheets that have extra empty rows extending the used range

        Args:
            sheet_name: Name of the worksheet to clean up
            check_column: Column to check for last data (default: 'A')
            start_row: Row to start checking from (default: 2, to preserve headers)
            max_delete_rows: Maximum number of rows to delete in one go (default: 1000000)
        """
        print(f"   Deleting extra rows after last data in sheet '{sheet_name}'...")

        sheet = self.workbook.sheets[sheet_name]

        try:
            # Find the last row with data in the check column
            last_row_with_data = sheet.range(f'{check_column}{start_row}').end('down').row

            # Handle case where end('down') goes to the very last row
            if last_row_with_data > 1048576 or last_row_with_data < start_row:
                last_row_with_data = start_row
                print(f"   No data found below row {start_row}")
                return

            print(f"   Last row with data: {last_row_with_data}")

            # Find the actual used range last row
            used_range_last_row = sheet.used_range.last_cell.row

            if used_range_last_row > last_row_with_data:
                # Calculate how many rows to delete
                rows_to_delete = min(used_range_last_row - last_row_with_data, max_delete_rows)
                start_delete_row = last_row_with_data + 1

                print(f"   Deleting {rows_to_delete} empty rows starting from row {start_delete_row}...")

                # Delete the entire rows
                end_delete_row = start_delete_row + rows_to_delete - 1
                sheet.range(f'{start_delete_row}:{end_delete_row}').api.EntireRow.Delete()

                print(f"    Successfully deleted {rows_to_delete} empty rows")
            else:
                print(f"   No extra rows to delete (last data row matches used range)")

        except Exception as e:
            print(f"    Error deleting extra rows: {e}")
            raise

    def find_all_pivot_tables(self) -> dict:
        """
        Find all pivot tables in the entire workbook

        Returns:
            Dictionary with sheet names as keys and list of pivot table names as values
        """
        print(f"   Searching for all pivot tables in workbook...")

        pivot_tables_dict = {}

        for sheet in self.workbook.sheets:
            sheet_name = sheet.name
            try:
                pivot_tables = sheet.api.PivotTables()
                if pivot_tables.Count > 0:
                    pivot_list = []
                    for i in range(1, pivot_tables.Count + 1):
                        pivot_list.append(pivot_tables(i).Name)
                    pivot_tables_dict[sheet_name] = pivot_list
                    print(f"    Found {pivot_tables.Count} pivot table(s) in sheet '{sheet_name}': {pivot_list}")
            except Exception as e:
                # Sheet might not support pivot tables
                pass

        if not pivot_tables_dict:
            print(f"    No pivot tables found in workbook")

        return pivot_tables_dict

    # =============================================================================
    # NEW METHODS - Moved from Main.py and made dynamic
    # =============================================================================

    def read_portfolio_data(self, sheet_name: str) -> pd.DataFrame:
        """
        Read portfolio data from a sheet.
        Reads only data columns (A:B, D:F) in chunks to avoid COM errors
        on large sheets or formula columns (column C is skipped).

        Args:
            sheet_name: Name of the portfolio sheet to read

        Returns:
            pd.DataFrame: Portfolio data with columns MONTH, CONTRACT_NO,
                          EQT_DESC, PD_CATEGORY, DPD
        """
        sheet = self.workbook.sheets[sheet_name]
        last_row = sheet.range('A2').end('down').row

        # Handle empty sheet
        if last_row > 1000000:
            print(f"   {sheet_name}: Empty sheet detected")
            return pd.DataFrame()

        print(f"   Reading data from sheet '{sheet_name}' ({last_row - 1} rows)...")

        CHUNK_SIZE = 50000
        ab_data = []   # columns A, B
        def_data = []  # columns D, E, F

        for start in range(2, last_row + 1, CHUNK_SIZE):
            end = min(start + CHUNK_SIZE - 1, last_row)

            # A:B  (MONTH, CONTRACT_NO)
            chunk_ab = sheet.range(f'A{start}:B{end}').value
            if chunk_ab is not None:
                if not isinstance(chunk_ab[0], list):
                    chunk_ab = [chunk_ab]
                ab_data.extend(chunk_ab)

            # D:F  (EQT_DESC, PD_CATEGORY, DPD) — skip formula column C
            chunk_def = sheet.range(f'D{start}:F{end}').value
            if chunk_def is not None:
                if not isinstance(chunk_def[0], list):
                    chunk_def = [chunk_def]
                def_data.extend(chunk_def)

        if not ab_data:
            return pd.DataFrame()

        df_left  = pd.DataFrame(ab_data,  columns=['MONTH', 'CONTRACT_NO'])
        df_right = pd.DataFrame(def_data, columns=['EQT_DESC', 'PD_CATEGORY', 'DPD'])
        df = pd.concat([df_left, df_right], axis=1)

        # Drop rows where both key fields are blank
        df = df[df['MONTH'].notna() & df['CONTRACT_NO'].notna()]

        print(f"    Successfully read {len(df)} rows")
        return df

    def write_portfolio_data(self, sheet_name: str, df: pd.DataFrame,
                            columns_to_write: List[str],
                            column_positions: dict = None,
                            full_clear: bool = False) -> None:
        """
        Write specific columns to Portfolio sheet at specified positions.

        Args:
            sheet_name: Name of the portfolio sheet
            df: DataFrame containing data to write
            columns_to_write: List of column names to write
            column_positions: Dict mapping column names to Excel columns.
                              Defaults to A=MONTH, B=CONTRACT_NO, D=EQT_DESC, E=PD_CATEGORY, F=DPD
            full_clear: When True AND df is empty, wipe the entire used range including
                        formula columns (DC_BUCKET, STAGE, etc.).
                        When False (default), only clear data columns A:F — preserves
                        formula columns beyond F (use this for Portfolio_1).
        """
        sheet = self.workbook.sheets[sheet_name]

        # Default column positions if not provided
        if column_positions is None:
            column_positions = {
                'MONTH': 'A',
                'CONTRACT_NO': 'B',
                'EQT_DESC': 'D',
                'PD_CATEGORY': 'E',
                'DPD': 'F'
            }

        if df.empty:
            if full_clear:
                # Nuclear clear: wipe entire used range from row 2 down
                # Removes DC_BUCKET, STAGE, and every other formula column
                used = sheet.used_range
                if used is not None:
                    last_row = used.last_cell.row
                    last_col = used.last_cell.column
                    if last_row >= 2 and last_col >= 1:
                        last_col_letter = self._col_number_to_letter(last_col)
                        sheet.range(f'A2:{last_col_letter}{last_row}').clear_contents()
                        print(f"  {sheet_name}: Fully cleared (A2:{last_col_letter}{last_row})")
                        return
                print(f"  {sheet_name}: Already empty — nothing to clear")
            else:
                # Partial clear: A:F and H (STAGE values) — keeps G (DC_BUCKET formula)
                last_row_a = sheet.range('A2').end('down').row
                last_row_c = sheet.range('C2').end('down').row
                candidates = [r for r in [last_row_a, last_row_c] if r < 1000000]
                last_row = max(candidates) if candidates else 0
                if last_row > 0:
                    sheet.range(f'A2:F{last_row}').clear_contents()
                    sheet.range(f'H2:H{last_row}').clear_contents()
                    print(f"  {sheet_name}: Data columns cleared (A2:F{last_row}, H2:H{last_row})")
                else:
                    print(f"  {sheet_name}: Already empty — nothing to clear")
            return

        # Before writing data: clear A:F and H (STAGE values) — preserve G (DC_BUCKET formula)
        last_row_a = sheet.range('A2').end('down').row
        last_row_c = sheet.range('C2').end('down').row
        candidates = [r for r in [last_row_a, last_row_c] if r < 1000000]
        last_row = max(candidates) if candidates else 0
        if last_row > 0:
            try:
                sheet.range(f'A2:F{last_row}').clear_contents()
                sheet.range(f'H2:H{last_row}').clear_contents()
            except:
                pass

        # Write each column to its specified position
        num_rows = len(df)
        for col_name in columns_to_write:
            if col_name in df.columns and col_name in column_positions:
                excel_col = column_positions[col_name]
                sheet.range(f'{excel_col}2').value = df[col_name].values.reshape(-1, 1).tolist()

        print(f"  {sheet_name}: {num_rows} rows written ({len(columns_to_write)} columns: {', '.join(columns_to_write)})")

    # =============================================================================
    # STATIC UTILITY METHODS - Date and File Operations
    # =============================================================================

    @staticmethod
    def months_between(start_date: datetime, end_date: datetime) -> int:
        """
        Calculate number of months between two dates.

        Args:
            start_date: Starting date
            end_date: Ending date

        Returns:
            int: Number of months between dates
        """
        return (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)

    @staticmethod
    def parse_month_string(month_str: str, date_format: str = '%m/%d/%Y') -> datetime:
        """
        Parse month string to datetime object.

        Args:
            month_str: Date string to parse
            date_format: Format of the date string (default: '%m/%d/%Y')

        Returns:
            datetime: Parsed datetime object
        """
        return datetime.strptime(month_str.strip(), date_format)

    @staticmethod
    def format_month_string(date: datetime, date_format: str = '%m/%d/%Y') -> str:
        """
        Format datetime to string.

        Args:
            date: Datetime object to format
            date_format: Desired output format (default: '%m/%d/%Y')

        Returns:
            str: Formatted date string
        """
        return date.strftime(date_format)

    @staticmethod
    def read_config_file(config_file_path: str, date_format: str = '%m/%d/%Y') -> datetime:
        """
        Read the latest recorded month from a config file.

        Args:
            config_file_path: Path to the config file
            date_format: Expected date format in config file (default: '%m/%d/%Y')

        Returns:
            datetime: Latest month from config file

        Raises:
            FileNotFoundError: If config file doesn't exist
            ValueError: If date format is invalid
        """
        if not os.path.exists(config_file_path):
            raise FileNotFoundError(f"Config file not found: {config_file_path}")

        with open(config_file_path, 'r') as f:
            date_str = f.read().strip()

        try:
            return ExcelPortfolioAutomation.parse_month_string(date_str, date_format)
        except ValueError:
            raise ValueError(f"Invalid date format in config file: {date_str}. Expected format: {date_format}")

    @staticmethod
    def save_config_file(config_file_path: str, latest_month: datetime, date_format: str = '%m/%d/%Y') -> None:
        """
        Save the latest month to a config file.

        Args:
            config_file_path: Path to the config file
            latest_month: Date to save
            date_format: Date format to use (default: '%m/%d/%Y')
        """
        with open(config_file_path, 'w') as f:
            f.write(ExcelPortfolioAutomation.format_month_string(latest_month, date_format))
        print(f"Config updated: {ExcelPortfolioAutomation.format_month_string(latest_month, date_format)}")

    @staticmethod
    def get_latest_file_in_folder(folder_path: str, file_pattern: str, 
                                  fallback_file: str = None) -> str:
        """
        Get the latest file matching a pattern in a folder.
        If no files found, returns fallback file.

        Args:
            folder_path: Path to search for files
            file_pattern: Glob pattern to match files (e.g., "*.xlsb", "PD_data_*_Updated_*.xlsb")
            fallback_file: File to return if no matches found (default: None)

        Returns:
            str: Path to the latest file or fallback file
        """
        # Search for files matching pattern
        search_pattern = os.path.join(folder_path, file_pattern)
        matching_files = glob.glob(search_pattern)

        if matching_files:
            # Sort by modification time (most recent first)
            matching_files.sort(key=os.path.getmtime, reverse=True)
            latest_file = matching_files[0]
            print(f"Using latest file: {os.path.basename(latest_file)}")
            return latest_file
        else:
            if fallback_file:
                print(f"Using fallback file: {os.path.basename(fallback_file)}")
                return fallback_file
            else:
                raise FileNotFoundError(f"No files found matching pattern: {file_pattern}")

    @staticmethod
    def find_summary_files_by_date_range(input_folder: str, start_month: datetime, 
                                        num_months: int, file_prefix: str = "3. Summary_",
                                        file_extension: str = ".xlsb",
                                        date_format_in_filename: str = '%Y-%m-%d') -> List[Tuple[str, datetime]]:
        """
        Find summary files for a specified date range.

        Args:
            input_folder: Folder containing summary files
            start_month: Starting month
            num_months: Number of months to find
            file_prefix: Prefix of summary files (default: "3. Summary_")
            file_extension: File extension (default: ".xlsb")
            date_format_in_filename: Date format used in filenames (default: '%Y-%m-%d')

        Returns:
            List of (file_path, month_date) tuples in chronological order
        """
        files = []

        for i in range(1, num_months + 1):
            target_date = ExcelPortfolioAutomation.add_months(start_month, i)
            # Summary files use specified date format in filename
            date_pattern = target_date.strftime(date_format_in_filename)

            pattern = os.path.join(input_folder, f"{file_prefix}{date_pattern}*{file_extension}")
            matches = glob.glob(pattern)

            if matches:
                files.append((matches[0], target_date))
                print(f"  Found: {os.path.basename(matches[0])}")
            else:
                print(f"  WARNING: No file for {date_pattern}")

        return files

    @staticmethod
    def extract_data_from_summary_files(file_paths: List[Tuple[str, datetime]], 
                                       column_mapping: dict,
                                       output_columns: List[str],
                                       sheet_name: str = 'SUMMARY',
                                       date_format: str = '%m/%d/%Y',
                                       max_header_search_rows: int = 10) -> pd.DataFrame:
        """
        Extract data from summary files into a DataFrame.

        Args:
            file_paths: List of (file_path, month_date) tuples
            column_mapping: Dict mapping source column names to target column names
                          e.g., {'CONTRACT NO': 'CONTRACT_NO', 'EQUIPMENT DESCRIPTION': 'EQT_DESC'}
            output_columns: List of columns to include in output (in order)
            sheet_name: Name of sheet to read (default: 'SUMMARY')
            date_format: Format for MONTH column (default: '%m/%d/%Y')
            max_header_search_rows: Maximum rows to search for headers (default: 10)

        Returns:
            pd.DataFrame: Consolidated data from all summary files
        """
        from concurrent.futures import ThreadPoolExecutor, as_completed

        def _extract_one(file_path, month_date):
            fname = os.path.basename(file_path)
            try:
                # --- Step 1: probe only the first row to find the header line ---
                found_header_row = 0
                for header_row in range(max_header_search_rows):
                    try:
                        probe = pd.read_excel(file_path, sheet_name=sheet_name,
                                              header=header_row, nrows=1,
                                              engine='pyxlsb')
                        if any(col in probe.columns for col in column_mapping.keys()):
                            found_header_row = header_row
                            break
                    except Exception:
                        continue

                # --- Step 2: one full read with the correct header row ---
                df = pd.read_excel(file_path, sheet_name=sheet_name,
                                   header=found_header_row, engine='pyxlsb')

                # Map columns (first non-null match wins — supports multiple src names -> same target)
                extracted = pd.DataFrame()
                for src_col, tgt_col in column_mapping.items():
                    found_col = next(
                        (c for c in df.columns if str(c).strip().upper() == src_col.upper()),
                        None
                    )
                    if found_col is not None:
                        # Only assign if target not yet set or currently all-null
                        if tgt_col not in extracted.columns or extracted[tgt_col].isna().all():
                            extracted[tgt_col] = df[found_col]
                    elif tgt_col not in extracted.columns:
                        extracted[tgt_col] = None

                # Diagnostic: report STAGE fill rate so column-name mismatches are visible
                if 'STAGE' in extracted.columns:
                    stage_found = int(extracted['STAGE'].notna().sum())
                    print(f"    STAGE values extracted: {stage_found}/{len(extracted)}"
                          + (" (column not found in source)" if stage_found == 0 else ""))

                # Add MONTH column
                extracted['MONTH'] = ExcelPortfolioAutomation.format_month_string(
                    month_date, date_format)

                # Filter empty rows
                first_col = list(column_mapping.values())[0]
                extracted = extracted[extracted[first_col].notna()]
                extracted = extracted[
                    ~extracted[first_col].astype(str).isin(['', '-', 'nan', 'None'])]

                extracted = extracted[output_columns]
                return fname, len(extracted), extracted, None

            except Exception as e:
                return fname, 0, None, e

        print(f"Extracting data from {len(file_paths)} file(s)...")
        max_workers = min(len(file_paths), 4)  # cap at 4 to avoid memory pressure
        results_map = {}

        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            futures = {pool.submit(_extract_one, fp, md): (fp, md)
                       for fp, md in file_paths}
            for fut in as_completed(futures):
                fname, nrows, df_result, err = fut.result()
                if err:
                    print(f"  ERROR {fname}: {err}")
                else:
                    print(f"  Extracting: {fname}\n    Rows: {nrows}")
                    results_map[fname] = df_result

        # Preserve original file order
        all_data = [results_map[os.path.basename(fp)]
                    for fp, _ in file_paths
                    if os.path.basename(fp) in results_map]

        if all_data:
            return pd.concat(all_data, ignore_index=True)
        return pd.DataFrame(columns=output_columns)

    @staticmethod
    def get_unique_months_from_dataframe(df: pd.DataFrame, month_column: str = 'MONTH',
                                        date_format: str = '%m/%d/%Y') -> List[datetime]:
        """
        Get sorted list of unique months from a DataFrame.

        Args:
            df: DataFrame containing month data
            month_column: Name of the column containing month data (default: 'MONTH')
            date_format: Format of date strings in the column (default: '%m/%d/%Y')

        Returns:
            List of unique month datetime objects in ascending order
        """
        if df.empty or month_column not in df.columns:
            return []

        months = pd.to_datetime(df[month_column], format=date_format, errors='coerce')
        unique_months = sorted(months.dropna().unique())
        return [pd.Timestamp(m).to_pydatetime() for m in unique_months]

    @staticmethod
    def filter_dataframe_by_months(df: pd.DataFrame, months_to_keep: List[datetime],
                                   month_column: str = 'MONTH',
                                   date_format: str = '%m/%d/%Y') -> pd.DataFrame:
        """
        Filter DataFrame to keep only rows matching specified months.

        Args:
            df: DataFrame to filter
            months_to_keep: List of datetime objects representing months to keep
            month_column: Name of the month column (default: 'MONTH')
            date_format: Format of date strings in the column (default: '%m/%d/%Y')

        Returns:
            pd.DataFrame: Filtered dataframe
        """
        if df.empty:
            return df

        df = df.copy()
        df['_MONTH_DT'] = pd.to_datetime(df[month_column], format=date_format, errors='coerce')

        # Convert months_to_keep to timestamps for comparison
        keep_timestamps = [pd.Timestamp(m) for m in months_to_keep]
        filtered = df[df['_MONTH_DT'].isin(keep_timestamps)]

        return filtered.drop(columns=['_MONTH_DT'])

    @staticmethod
    def consolidate_summary_files(input_folder: str, file_pattern: str = "3. Summary_*.xlsb",
                                   sheet_name: str = "SUMMARY", header_row: int = 0) -> pd.DataFrame:
        """
        Consolidate data from multiple summary Excel files into a single DataFrame

        Args:
            input_folder: Path to folder containing summary files
            file_pattern: Pattern to match summary files (default: "3. Summary_*.xlsb")
            sheet_name: Name of the sheet to read from (default: "SUMMARY")
            header_row: Row number containing column headers (default: 0)

        Returns:
            pd.DataFrame: Consolidated dataframe with columns CONTRACT_NO, EQT_DESC, PD_CATEGORY, DPD, MONTH
        """
        print(f"   Consolidating summary files from: {input_folder}")
        print(f"   File pattern: {file_pattern}")
        print()

        # Find all matching files
        search_pattern = os.path.join(input_folder, file_pattern)
        all_summary_files = sorted(glob.glob(search_pattern))

        # Take only the latest 6 files
        summary_files = all_summary_files[-6:] if len(all_summary_files) >= 6 else all_summary_files

        print(f"   Found {len(all_summary_files)} summary files, using latest {len(summary_files)}:")
        for file in summary_files:
            print(f"     - {os.path.basename(file)}")
        print()

        # Column mapping: source column name -> target column name
        column_mapping = {
            'CONTRACT NO': 'CONTRACT_NO',
            'EQUIPMENT DESCRIPTION': 'EQT_DESC',
            'PD/LGD CATEGORY': 'PD_CATEGORY',
            'CLIENT DPD': 'DPD'
        }

        consolidated_data = []

        for file_path in summary_files:
            try:
                print(f"   Processing: {os.path.basename(file_path)}")

                # Read the SUMMARY sheet using pandas (works with closed files)
                # Try to find the correct header row
                df = None
                for try_header in range(header_row, min(header_row + 10, 20)):
                    try:
                        temp_df = pd.read_excel(file_path, sheet_name=sheet_name, header=try_header, engine='pyxlsb')
                        # Check if this row contains the columns we're looking for
                        if 'CONTRACT NO' in temp_df.columns or 'CONTRACT_NO' in temp_df.columns:
                            df = temp_df
                            if try_header != header_row:
                                print(f"     Found headers at row {try_header} (tried starting from row {header_row})")
                            break
                    except:
                        continue

                if df is None:
                    # Fall back to specified header row
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, engine='pyxlsb')

                print(f"     Read {len(df)} rows")
                print(f"     Available columns: {list(df.columns)[:10]}...")  # Show first 10 columns only

                # Select and rename the required columns
                selected_data = pd.DataFrame()

                for source_col, target_col in column_mapping.items():
                    # Try to find the column (case-insensitive, with or without spaces)
                    found_col = None
                    for col in df.columns:
                        if str(col).strip().upper() == source_col.upper():
                            found_col = col
                            break

                    if found_col:
                        selected_data[target_col] = df[found_col]
                    else:
                        print(f"     Warning: Column '{source_col}' not found in {os.path.basename(file_path)}")
                        selected_data[target_col] = None

                # Extract date from filename (e.g., "3. Summary_2025-04-30_Final_V2.xlsb" -> "04/30/2025")
                filename = os.path.basename(file_path)
                date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
                if date_match:
                    year, month, day = date_match.groups()
                    month_date = f"{month}/{day}/{year}"
                else:
                    month_date = "Unknown"

                selected_data['MONTH'] = month_date
                print(f"     Extracted month: {month_date}")

                consolidated_data.append(selected_data)
                print(f"     Successfully extracted {len(selected_data)} rows with {len(selected_data.columns)} columns")

            except Exception as e:
                print(f"     Error processing {os.path.basename(file_path)}: {e}")
                import traceback
                traceback.print_exc()

        # Combine all dataframes
        if consolidated_data:
            final_df = pd.concat(consolidated_data, ignore_index=True)
            print()
            print(f"   Consolidation complete!")
            print(f"     Total rows: {len(final_df)}")
            print(f"     Total columns: {len(final_df.columns)}")
            print(f"     Columns: {list(final_df.columns)}")
            return final_df
        else:
            print()
            print(f"   No data consolidated")
            return pd.DataFrame()

    @staticmethod
    def add_months(date: datetime, months: int) -> datetime:
        """Add months to a date, handling month-end correctly.
        If original date is month-end, result will also be month-end."""
        # Check if original date is the last day of its month
        original_month_last_day = calendar.monthrange(date.year, date.month)[1]
        is_month_end = (date.day == original_month_last_day)

        # Calculate target month/year
        month = date.month - 1 + months
        year = date.year + month // 12
        month = month % 12 + 1

        # Get last day of target month
        target_month_last_day = calendar.monthrange(year, month)[1]

        # If original was month-end, use target month-end; otherwise use original day (capped)
        if is_month_end:
            day = target_month_last_day
        else:
            day = min(date.day, target_month_last_day)

        return datetime(year, month, day)

    # =============================================================================
    # PIVOT TABLE & HISTORIC PD METHODS
    # =============================================================================

    def extract_pivot_table_to_dataframe(self, sheet_name: str, 
                                        start_cell: str = 'A4',
                                        contract_col: str = 'A',
                                        pd_category_col: str = 'B',
                                        first_data_col: str = 'C') -> pd.DataFrame:
        """
        Extract pivot table data from a sheet to DataFrame.
        Assumes pivot structure:
        - Column A: CONTRACT_NO_NOLASTDIG
        - Column B: PD_CATEGORY  
        - Column C onwards: Date columns (e.g., 2024-09, 2024-10, etc.)

        Args:
            sheet_name: Name of the sheet containing pivot table
            start_cell: Starting cell of pivot data (default: 'A4')
            contract_col: Column letter for contract numbers (default: 'A')
            pd_category_col: Column letter for PD category (default: 'B')
            first_data_col: First data column letter (default: 'C')

        Returns:
            pd.DataFrame: Pivot data with multi-index (CONTRACT_NO_NOLASTDIG, PD_CATEGORY) 
                         and date columns
        """
        print(f"\n   Extracting pivot table from '{sheet_name}'...")
        
        sheet = self.workbook.sheets[sheet_name]
        
        # Read the entire used range
        used_range = sheet.used_range
        data = used_range.value
        
        if not data or len(data) < 4:
            print(f"   ERROR: No data found in sheet '{sheet_name}'")
            return pd.DataFrame()
        
        def _to_yyyymm(cell) -> str | None:
            """
            Normalise a pivot column header cell to 'YYYY-MM' string.
            Handles: str 'YYYY-MM', datetime/date objects, Excel date serials (float).
            Returns None if the cell cannot be recognised as a date.
            """
            if cell is None:
                return None
            # Already a YYYY-MM string
            if isinstance(cell, str):
                s = cell.strip()
                if re.match(r'^\d{4}-\d{2}', s):
                    return s[:7]           # keep only YYYY-MM
                # MM/DD/YYYY or M/D/YYYY (e.g. "04/30/2025") — MONTH field format
                m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
                if m:
                    return f"{m.group(3)}-{int(m.group(1)):02d}"
                # Mon-YY or Mon YYYY (e.g. "Apr-25", "Apr 2025") — Excel pivot display
                m = re.match(r'^([A-Za-z]{3})[-\s](\d{2,4})$', s)
                if m:
                    try:
                        from datetime import datetime as _dt
                        year_s = m.group(2)
                        year = int(year_s) + (2000 if len(year_s) == 2 else 0)
                        month = _dt.strptime(m.group(1), '%b').month
                        return f"{year}-{month:02d}"
                    except Exception:
                        pass
                return None                # (blank), Grand Total, etc.
            # datetime / date objects (xlwings returns these for date-typed cells)
            if hasattr(cell, 'year') and hasattr(cell, 'month'):
                return f"{cell.year}-{cell.month:02d}"
            # Excel date serial (float): days since 1899-12-30
            # Valid financial range: ~29000 (1979) to ~60000 (2064)
            # Anything outside this range is a row count or other number, not a date
            if isinstance(cell, (int, float)):
                if 29000 <= cell <= 60000:
                    try:
                        from datetime import datetime as _dt, timedelta as _td
                        excel_epoch = _dt(1899, 12, 30)
                        d = excel_epoch + _td(days=int(cell))
                        return f"{d.year}-{d.month:02d}"
                    except Exception:
                        return None
                return None  # out of valid date serial range
            return None

        # Find header row — any cell from column 3 onwards that normalises to YYYY-MM
        header_row_idx = None
        for i, row in enumerate(data):
            if row and len(row) > 2:
                for cell in row[2:]:
                    if _to_yyyymm(cell) is not None:
                        header_row_idx = i
                        break
                if header_row_idx is not None:
                    break

        if header_row_idx is None:
            print(f"   ERROR: Could not find header row with date columns")
            # Debug: show first 5 rows to diagnose the format
            for i, row in enumerate(data[:5]):
                if row:
                    preview = [repr(c) for c in row[:8]]
                    print(f"   Row {i}: {preview}")
            return pd.DataFrame()

        print(f"   Found header row at index: {header_row_idx}")

        # Extract headers — only keep YYYY-MM date columns; skip (blank), Grand Total, etc.
        headers = ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY']
        date_headers = []
        date_col_indices = []  # raw column indices in the data row for each date header
        for idx, cell in enumerate(data[header_row_idx][2:]):  # Skip first 2 columns
            if cell is None or (isinstance(cell, str) and cell.strip() == ''):
                break  # Stop at first empty column
            ym = _to_yyyymm(cell)
            if ym is None:
                if isinstance(cell, str) and cell.strip() == 'Grand Total':
                    break
                # skip non-date columns like (blank)
                continue
            headers.append(ym)
            date_headers.append(ym)
            date_col_indices.append(idx + 2)  # +2 because we enumerate from col index 2

        print(f"   Found {len(date_headers)} date columns: {date_headers[:5]}...")

        # Extract data rows — use explicit column index mapping to honour skipped columns
        data_rows = []
        for row in data[header_row_idx + 1:]:
            if row and row[0]:  # If first column has value (contract number)
                row_data = [row[0], row[1]] + [
                    row[i] if i < len(row) else None for i in date_col_indices
                ]
                data_rows.append(row_data)
        
        print(f"   Extracted {len(data_rows)} data rows")
        
        # Create DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Filter out empty rows and pivot artifacts like (blank) and Grand Total
        df = df[df['CONTRACT_NO_NOLASTDIG'].notna()]
        df = df[~df['CONTRACT_NO_NOLASTDIG'].astype(str).str.strip().isin(['(blank)', 'blank', '', 'Grand Total'])]
        df = df[df['PD_CATEGORY'].notna()]
        df = df[~df['PD_CATEGORY'].astype(str).str.strip().isin(['(blank)', 'blank', '', 'Grand Total'])]
        
        print(f"   Final DataFrame: {len(df)} rows x {len(df.columns)} columns")
        print(f"{df.head}")
        
        return df

    @staticmethod
    def convert_pivot_date_to_year_month(date_str: str) -> tuple:
        """
        Convert pivot date format to year and month name.
        
        Args:
            date_str: Date string in format "YYYY-MM" (e.g., "2024-09")
        
        Returns:
            tuple: (year, month_name) e.g., (2024, "Sep")
        """
        try:
            # Parse the date string
            parts = date_str.split('-')
            year = int(parts[0])
            month_num = int(parts[1])
            
            # Convert month number to name
            month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            month_name = month_names[month_num - 1]
            
            return (year, month_name)
        except:
            return (None, None)

    def write_historic_pd_format(self, sheet_name: str, pivot_df: pd.DataFrame,
                                year_row: int = 1, month_row: int = 2, data_start_row: int = 3,
                                contract_col: str = 'A', pd_category_col: str = 'B') -> None:
        """
        Write pivot data to Historic PD sheet format using direct positional mapping.

        Dynamically writes:
        - Row 1: Year headers (e.g., 2024 for C1:G1, 2025 for H1:O1)
        - Row 2: Month abbreviation headers (e.g., Aug, Sep, ..., Jul, Aug)
        - Row 3+: Data (CONTRACT_NO_NOLASTDIG in col A, PD_CATEGORY in col B, values in C:O)

        Mapping is positional: first pivot date column -> col C, second -> col D, etc.
        Total months is always 13.

        Args:
            sheet_name: Name of the sheet to update
            pivot_df: DataFrame with pivot data (from extract_pivot_table_to_dataframe)
            year_row: Row number for years (default: 1)
            month_row: Row number for month names (default: 2)
            data_start_row: First data row (default: 3)
            contract_col: Column letter for contract numbers (default: 'A')
            pd_category_col: Column letter for PD category (default: 'B')
        """
        print(f"\n   Writing data to '{sheet_name}'...")

        sheet = self.workbook.sheets[sheet_name]

        # Get date columns from pivot DataFrame (only YYYY-MM format columns)
        date_columns = [col for col in pivot_df.columns
                       if col not in ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY']
                       and re.match(r'\d{4}-\d{2}', str(col))]

        if not date_columns:
            print(f"   ERROR: No date columns found in pivot data")
            return

        print(f"   Date columns ({len(date_columns)}): {date_columns}")

        # Parse each date column to extract year and month abbreviation
        month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        year_headers = []
        month_headers = []

        for date_col in date_columns:
            parts = str(date_col).split('-')
            year = int(parts[0])
            month_num = int(parts[1])
            year_headers.append(year)
            month_headers.append(month_names[month_num - 1])

        print(f"   Year headers: {year_headers}")
        print(f"   Month headers: {month_headers}")

        # Column C = 3 (first data column)
        first_data_col_num = 3
        last_data_col_letter = self._col_number_to_letter(first_data_col_num + len(date_columns) - 1)

        # Formula columns sit immediately after the last date column
        # e.g. for 6 months (C-H): I=Sum, J=Max, K=First Value, L=DC Bucket
        formula_start_col_num = first_data_col_num + len(date_columns)   # 3+6=9 → I
        sum_col   = self._col_number_to_letter(formula_start_col_num)         # I
        max_col   = self._col_number_to_letter(formula_start_col_num + 1)     # J
        first_col = self._col_number_to_letter(formula_start_col_num + 2)     # K
        dc_col    = self._col_number_to_letter(formula_start_col_num + 3)     # L
        last_formula_col = self._col_number_to_letter(formula_start_col_num + 3)  # L

        # Disable screen updating, auto-calculation, alerts, and events for performance
        print(f"   Disabling screen updating and auto-calculation...")
        self.app.screen_updating = False
        original_calculation = self.app.api.Calculation
        self.app.api.Calculation = -4135  # xlCalculationManual
        self.app.api.DisplayAlerts = False
        self.app.api.EnableEvents = False

        try:
            # Step 1: Write year headers to Row 1 (C1 onwards)
            # Clear the full old header range C1:S1 to remove leftover headers from prior runs
            print(f"   Writing year headers to row {year_row} (C{year_row}:{last_data_col_letter}{year_row})...")
            full_year_range = sheet.range(f'C{year_row}:S{year_row}')
            try:
                full_year_range.api.UnMerge()
            except:
                pass
            full_year_range.clear_contents()

            # Write all year values at once (batch)
            sheet.range(f'C{year_row}').value = [year_headers]
            print(f"   Year values written")

            # Merge consecutive columns that share the same year
            i = 0
            while i < len(year_headers):
                j = i + 1
                while j < len(year_headers) and year_headers[j] == year_headers[i]:
                    j += 1
                if j - i > 1:
                    start_col = self._col_number_to_letter(first_data_col_num + i)
                    end_col = self._col_number_to_letter(first_data_col_num + j - 1)
                    merge_range = f'{start_col}{year_row}:{end_col}{year_row}'
                    sheet.range(merge_range).api.Merge()
                    print(f"     Merged {merge_range} = {year_headers[i]}")
                i = j

            # Step 2: Write month abbreviation headers to Row 2 (C2:last_data_col_letter)
            # Also write A2/B2 headers — pivot source starts at row 2, so ALL columns
            # must have valid labels or Excel rejects the source range as invalid
            print(f"   Writing month headers to row {month_row} (C{month_row}:{last_data_col_letter}{month_row})...")
            sheet.range(f'{contract_col}{month_row}').value    = 'CONTRACT_NO_NOLASTDIG'
            sheet.range(f'{pd_category_col}{month_row}').value = 'PD_CATEGORY'
            sheet.range(f'C{month_row}:S{month_row}').clear_contents()
            sheet.range(f'C{month_row}').value = [month_headers]

            # Write formula column headers immediately after month data
            # Total | WORST | FIRST | LAST
            sheet.range(f'{sum_col}{month_row}').value   = 'Total'
            sheet.range(f'{max_col}{month_row}').value   = 'WORST'
            sheet.range(f'{first_col}{month_row}').value = 'FIRST'
            sheet.range(f'{dc_col}{month_row}').value    = 'LAST'
            # Header cells (WORST/FIRST/LAST) keep their existing green color — do not override
            print(f"   Formula column headers written: {sum_col}=Total, {max_col}=WORST, {first_col}=FIRST, {dc_col}=LAST")

            # Step 3: Clear existing data from data_start_row down (A through S — full old range)
            last_row = sheet.range(f'{contract_col}{data_start_row}').end('down').row
            if last_row < 1000000 and last_row >= data_start_row:
                clear_range = f'{contract_col}{data_start_row}:S{last_row}'
                print(f"   Clearing data range {clear_range}...")
                sheet.range(clear_range).clear_contents()

            if pivot_df.empty:
                print(f"   No data to write")
                return

            num_rows = len(pivot_df)
            print(f"   Writing {num_rows} rows of data starting at row {data_start_row}...")

            # Step 4: Write CONTRACT_NO_NOLASTDIG to column A
            print(f"   Writing column A (CONTRACT_NO)...")
            contract_values = pivot_df['CONTRACT_NO_NOLASTDIG'].values.reshape(-1, 1).tolist()
            sheet.range(f'{contract_col}{data_start_row}').value = contract_values

            # Step 5: Write PD_CATEGORY to column B
            print(f"   Writing column B (PD_CATEGORY)...")
            category_values = pivot_df['PD_CATEGORY'].values.reshape(-1, 1).tolist()
            sheet.range(f'{pd_category_col}{data_start_row}').value = category_values

            # Step 6: Write date values (columns C through last_data_col_letter)
            print(f"   Writing columns C-{last_data_col_letter} (date values)...")
            data_values = pivot_df[date_columns].values.tolist()
            sheet.range(f'C{data_start_row}').value = data_values

            print(f"   Successfully wrote {num_rows} rows x {len(date_columns) + 2} columns")
            print(f"   Year headers: Row {year_row} (C-{last_data_col_letter})")
            print(f"   Month headers: Row {month_row} (C-{last_data_col_letter})")
            print(f"   Data: Rows {data_start_row}-{data_start_row + num_rows - 1} (A-{last_data_col_letter})")

            # Step 7: Write formulas to dynamic formula columns (I,J,K,L for 6 months)
            last_data_row = data_start_row + num_rows - 1
            r   = data_start_row       # first formula row
            lcl = last_data_col_letter  # e.g. 'H' for 6 months
            print(f"   Writing formulas to columns {sum_col}-{last_formula_col} (rows {r}-{last_data_row})...")

            # Set formulas in the first data row
            sheet.range(f'{sum_col}{r}').formula   = f'=SUM(C{r}:{lcl}{r})'
            sheet.range(f'{max_col}{r}').formula   = f'=MAX(C{r}:{lcl}{r})'
            sheet.range(f'{first_col}{r}').formula = f'=INDEX(C{r}:{lcl}{r},MATCH(TRUE,INDEX((C{r}:{lcl}{r}<>""),0),0))'
            sheet.range(f'{dc_col}{r}').formula    = f'=LOOKUP(2,1/(C{r}:{lcl}{r}<>""),C{r}:{lcl}{r})'

            # AutoFill down to last data row
            if last_data_row > data_start_row:
                source = sheet.range(f'{sum_col}{r}:{last_formula_col}{r}')
                dest   = sheet.range(f'{sum_col}{r}:{last_formula_col}{last_data_row}')
                source.api.AutoFill(Destination=dest.api, Type=0)  # xlFillDefault

            print(f"   Formulas written to {sum_col}{r}:{last_formula_col}{last_data_row}")

            # Step 8: Apply light yellow fill to WORST, FIRST, LAST data rows
            # Header row keeps its existing green color
            light_yellow = (255, 255, 204)
            sheet.range(f'{max_col}{data_start_row}:{max_col}{last_data_row}').color   = light_yellow
            sheet.range(f'{first_col}{data_start_row}:{first_col}{last_data_row}').color = light_yellow
            sheet.range(f'{dc_col}{data_start_row}:{dc_col}{last_data_row}').color     = light_yellow
            print(f"   Light yellow fill applied to {max_col}:{dc_col} data rows")

            # Step 10: Delete any extra columns to the right of the last formula column
            # These are leftover from previous runs that had more month columns
            last_formula_col_num = formula_start_col_num + 3  # e.g. L = 12
            last_used_col = sheet.used_range.last_cell.column
            if last_used_col > last_formula_col_num:
                delete_start_letter = self._col_number_to_letter(last_formula_col_num + 1)
                delete_end_letter   = self._col_number_to_letter(last_used_col)
                print(f"   Deleting extra columns {delete_start_letter}:{delete_end_letter}...")
                sheet.range(
                    f'{delete_start_letter}1:{delete_end_letter}1'
                ).api.EntireColumn.Delete()
                print(f"   Extra columns deleted")
            else:
                print(f"   No extra columns to delete beyond {last_formula_col}")

        finally:
            # Re-enable screen updating, auto-calculation, alerts, and events
            print(f"   Re-enabling screen updating and auto-calculation...")
            self.app.api.EnableEvents = True
            self.app.api.DisplayAlerts = True
            self.app.api.Calculation = original_calculation
            self.app.screen_updating = True

        print(f"   Historic PD update complete!")
    
    def _col_number_to_letter(self, col_num: int) -> str:
        """Convert column number to Excel column letter (e.g., 1->A, 27->AA)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + 65) + result
            col_num //= 26
        return result

    def setup_historic_pivot_tables(self, pivot_sheet_name: str, data_sheet_name: str,
                                 big_pivot_name: str, small_pivot_name: str,
                                 last_month_field: str, last_data_row: int,
                                 last_col: str = 'L') -> None:
        """
        Set up pivot tables in Historic PD file after data write.

        Steps:
        1. Delete existing 'Select PD Category' slicer
        2. PivotTable2 (E4) -> Change Data Source -> select all rows -> OK
        3. PivotTable1 (O4) -> Change Data Source -> select all rows -> OK
        4. Refresh both pivots
        5. PivotTable1 -> drag last_month_field to Rows -> filter out (blank)
        6. PivotTable2 -> Add PD_CATEGORY slicer
        7. Slicer Report Connections -> tick PivotTable2 + PivotTable1
        """
        import time
        import win32com.client

        print(f"\n   Setting up pivot tables in '{pivot_sheet_name}'...")

        pivot_sheet = self.workbook.sheets[pivot_sheet_name]
        source_range = f"'{data_sheet_name}'!$A$2:${last_col}${last_data_row}"

        # Get raw win32com workbook/worksheet objects via xlwings .api
        # This is the most reliable approach — no GetActiveObject/name-matching needed
        wb_com    = self.workbook.api
        excel_com = wb_com.Application
        ws_com    = wb_com.Worksheets(pivot_sheet_name)

        def _get_field(pt, name):
            try:
                return pt.PivotFields(name)
            except (TypeError, AttributeError):
                return pt.PivotFields.Item(name)

        def _get_item(pivot_items, idx):
            try:
                return pivot_items(idx)
            except (TypeError, AttributeError):
                return pivot_items.Item(idx)

        try:
            # ====================================================================
            # Step 1: Delete existing slicers via raw COM (wb_com, not xlwings
            # wrapper) so Excel fully commits the disconnection before Step 2.
            # ====================================================================
            print(f"   Step 1: Deleting existing slicers...")
            try:
                while wb_com.SlicerCaches.Count > 0:
                    wb_com.SlicerCaches(1).Delete()
                print(f"     Slicers deleted")
            except Exception as e:
                print(f"     No slicers to delete or error: {e}")
            time.sleep(2)   # Let Excel fully process the disconnection

            # ====================================================================
            # Steps 2+3: Shared cache strategy.
            #
            # The problem on Run 2: after Run 1, PivotTable1 and PivotTable2
            # already share the same cache.  Calling ChangePivotCache on
            # PivotTable2 while PivotTable1 still holds a reference to the old
            # shared cache causes "PivotTable field name is not valid" because
            # Excel tries to reconcile the new column set against PivotTable1's
            # old field list and fails.
            #
            # Solution: delete PivotTable1 FIRST (breaks the shared-cache link),
            # then safely change PivotTable2's cache, then recreate PivotTable1
            # from the new shared cache so both are re-linked on the new source.
            # ====================================================================

            # --- Delete PivotTable1 first (break old shared-cache link) ---
            print(f"   Step 2a: Removing {small_pivot_name} to break old cache link...")
            pivot_sheet.range("O4").select()
            try:
                small_pt_com = ws_com.PivotTables(small_pivot_name)
                old_row = small_pt_com.TableRange2.Row
                old_col = small_pt_com.TableRange2.Column
                small_pt_com.TableRange2.Clear()
                print(f"     {small_pivot_name} cleared (was at row={old_row} col={old_col})")
            except Exception as _del_e:
                # First run — pivot may not exist yet; default position to O4
                print(f"     {small_pivot_name} not found ({_del_e}), defaulting to row=4 col=15")
                old_row, old_col = 4, 15
            time.sleep(1)

            # --- Now safely change PivotTable2's cache ---
            print(f"   Step 2b: Assigning shared cache to {big_pivot_name}...")
            big_pt_com = ws_com.PivotTables(big_pivot_name)

            shared_cache = wb_com.PivotCaches().Create(SourceType=1, SourceData=source_range)
            print(f"     Shared cache created ({source_range})")

            pivot_sheet.range("E4").select()
            big_pt_com.ChangePivotCache(shared_cache)
            print(f"     {big_pivot_name} -> shared cache assigned")
            time.sleep(1)

            # --- Recreate PivotTable1 from the same shared_cache ---
            print(f"   Step 3: Recreating {small_pivot_name} from shared cache...")
            target_cell = ws_com.Cells(old_row, old_col)
            small_pt_com = shared_cache.CreatePivotTable(
                TableDestination=target_cell,
                TableName=small_pivot_name
            )
            print(f"     {small_pivot_name} recreated from shared cache")
            time.sleep(1)

            # ====================================================================
            # Step 4: Refresh both pivots
            # ====================================================================
            print(f"   Step 4: Refreshing pivots...")
            self.refresh_pivot_table(pivot_sheet_name, big_pivot_name)
            self.refresh_pivot_table(pivot_sheet_name, small_pivot_name)
            print(f"     Both pivots refreshed")
            time.sleep(2)

            # ====================================================================
            # Step 5: PivotTable1 -> drag last_month_field to Rows
            #         then filter out (blank) via Filters trick
            # When 13 months span two years and share a month name
            # (e.g. Mar 2025 + Mar 2026), Excel creates "Mar" AND "Mar2".
            # We always want "Mar2" (the most recent duplicate) in Rows.
            # ====================================================================
            print(f"   Step 5: Setting up {small_pivot_name} Rows (last month field)...")
            pivot_sheet.range("O4").select()

            # 5a. Remove ALL existing row fields so none linger from previous run
            print(f"     Clearing existing Row fields...")
            try:
                while small_pt_com.RowFields.Count > 0:
                    small_pt_com.RowFields(1).Orientation = 0  # xlHidden
                print(f"     Row fields cleared")
            except Exception as _e:
                print(f"     Note clearing rows: {_e}")

            # 5b. Prefer "fieldname2" (duplicate month) over bare "fieldname"
            preferred_field_name = last_month_field + "2"
            try:
                field = _get_field(small_pt_com, preferred_field_name)
                print(f"     Detected duplicate month — using '{preferred_field_name}'")
            except Exception:
                preferred_field_name = last_month_field
                field = _get_field(small_pt_com, preferred_field_name)
                print(f"     Using field '{preferred_field_name}'")

            # 5c. Drag field to Rows
            try:
                field.Orientation = 1  # xlRowField
                print(f"     '{preferred_field_name}' added to Rows")
            except Exception as e:
                print(f"     Error adding field to Rows: {e}")
                import traceback
                traceback.print_exc()

            # 5c2. Add Count of CONTRACT_NO_NOLASTDIG as the Values field.
            # Using a DIFFERENT field from the row field avoids Excel moving
            # the row field to Values when AddDataField is called.
            try:
                count_field = _get_field(small_pt_com, "CONTRACT_NO_NOLASTDIG")
                small_pt_com.AddDataField(count_field, "Count", -4112)  # -4112 = xlCount
                print(f"     Count (CONTRACT_NO_NOLASTDIG) added to Values")
            except Exception as _cf:
                print(f"     WARNING: Could not add Count field: {_cf}")

            print(f"     Waiting 3 seconds...")
            time.sleep(3)

            # 5d-5h. Filter (blank) via Filters trick:
            #   Move field to Filters -> enable multi-select -> untick blank -> move back to Rows
            print(f"   Filtering (blank) from {small_pivot_name}...")

            pivot_sheet.range("O4").select()
            time.sleep(1)

            print(f"     Moving '{preferred_field_name}' to Filters area...")
            try:
                field.Orientation = 3  # xlPageField
                print(f"     Moved to Filters")
            except Exception as e:
                print(f"     Error moving to Filters: {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)

            pivot_sheet.range("P2").select()
            try:
                field.EnableMultiplePageItems = True
                print(f"     Select Multiple Items enabled")
            except Exception as e:
                print(f"     Error enabling multi-select: {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)

            print(f"     Unticking (blank)...")
            try:
                try:
                    pi_items = field.PivotItems()
                except TypeError:
                    pi_items = field.PivotItems
                item_count = pi_items.Count
                print(f"     '{preferred_field_name}' has {item_count} items")

                blank_hidden = False
                for i in range(1, item_count + 1):
                    try:
                        pi = _get_item(pi_items, i)
                        name = str(pi.Name)
                        print(f"       Item {i}: '{name}'")
                        if name.lower() in ['(blank)', 'blank', '']:
                            pi.Visible = False
                            blank_hidden = True
                            print(f"     >>> Unticked: '{name}'")
                    except Exception as e:
                        print(f"       Error with item {i}: {e}")

                if not blank_hidden and item_count > 1:
                    try:
                        pi = _get_item(pi_items, item_count)
                        print(f"     >>> FALLBACK: Unticking item {item_count} '{pi.Name}'")
                        pi.Visible = False
                    except Exception as e:
                        print(f"     Fallback failed: {e}")

                print(f"     (blank) unticked")
            except Exception as e:
                print(f"     ERROR unticking (blank): {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)

            print(f"     Moving '{preferred_field_name}' back to Rows...")
            try:
                field.Orientation = 1  # xlRowField
                print(f"     '{preferred_field_name}' back in Rows")
            except Exception as e:
                print(f"     Error moving back to Rows: {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)


            # ====================================================================
            # Step 6: PivotTable2 -> Insert Slicer -> PD_CATEGORY
            # Use raw win32com (wb_com / ws_com) throughout so AddPivotTable
            # receives the same COM type as small_pt_com — no xlwings mismatch.
            # ====================================================================
            print(f"   Step 6: Adding PD_CATEGORY slicer to {big_pivot_name}...")
            pivot_sheet.range("E4").select()
            try:
                big_pt_com2 = ws_com.PivotTables(big_pivot_name)
                slicer_cache_com = wb_com.SlicerCaches.Add2(big_pt_com2, "PD_CATEGORY")
                slicer_com = slicer_cache_com.Slicers.Add(ws_com)
                slicer_com.Caption = "Select PD Category"

                # Position slicer to the right of the big pivot
                big_pt_range = big_pt_com2.TableRange2
                slicer_com.Width = 144
                slicer_com.Height = 200
                slicer_com.Left = big_pt_range.Left + big_pt_range.Width + 10
                slicer_com.Top = big_pt_range.Top
                print(f"     Slicer added to {big_pivot_name}")

                # ================================================================
                # Step 7: Report Connections -> tick PivotTable2 + PivotTable1
                # Both objects are raw win32com — no type mismatch.
                # ================================================================
                print(f"   Step 7: Connecting slicer to {small_pivot_name} via Report Connections...")
                try:
                    small_pt_com2 = ws_com.PivotTables(small_pivot_name)
                    slicer_cache_com.PivotTables.AddPivotTable(small_pt_com2)
                    print(f"     Slicer connected to {small_pivot_name}")
                except Exception as _se:
                    print(f"     WARNING: Could not connect slicer to {small_pivot_name}: {_se}")

                print(f"     PD_CATEGORY slicer connected to both pivot tables")

            except Exception as e:
                print(f"     ERROR adding slicer: {e}")
                import traceback
                traceback.print_exc()

        except Exception as e:
            print(f"\n   FATAL ERROR: {e}")
            import traceback
            traceback.print_exc()
            raise

        print(f"\n   Pivot setup complete!")

    @staticmethod
    def copy_pivot_to_historic(latest_pd_file: str,
                              input_folder: str,
                              output_folder: str,
                              historic_filename: str = "02. Historic PD Calculation 2024-25.xlsb",
                              pivot_sheet: str = "01.Pivoted_Portfolio",
                              historic_sheet: str = "02.Working") -> str:
        """
        Complete workflow to copy pivot data to historic PD format.

        Run 1: pivot has 7 months (Portfolio_1 only) — written directly.
        Run 2: pivot has 13 months (both portfolios combined) — written directly.
        No merging needed; the big pivot already combines both portfolios.

        Args:
            latest_pd_file: Path to the latest PD file with pivot data
            input_folder: Folder containing the original historic file
            output_folder: Folder to save the updated historic file
            historic_filename: Name of the historic file
            pivot_sheet: Sheet name with pivot data
            historic_sheet: Sheet name to write historic format

        Returns:
            str: Path to the saved historic file
        """
        print("\n" + "="*80)
        print("COPYING PIVOT TO HISTORIC PD FORMAT")
        print("="*80)

        historic_input_file = os.path.join(input_folder, historic_filename)

        # Step 1: Extract pivot data from latest PD file
        print(f"\n1. Opening PD file to extract pivot...")
        print(f"   File: {os.path.basename(latest_pd_file)}")

        with ExcelPortfolioAutomation(latest_pd_file, visible=True) as excel:
            pivot_df = excel.extract_pivot_table_to_dataframe(pivot_sheet)

            if pivot_df.empty:
                print("\n   ERROR: No pivot data extracted!")
                return None

        date_col_count = len([c for c in pivot_df.columns
                               if c not in ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY']])
        print(f"   Extracted pivot data: {len(pivot_df)} rows x {date_col_count} date cols")

        # Step 2: Open historic file and write data
        print(f"\n2. Opening Historic PD file...")
        print(f"   File: {os.path.basename(historic_input_file)}")
        
        with ExcelPortfolioAutomation(historic_input_file, visible=True) as excel:
            # Write in historic format
            excel.write_historic_pd_format(historic_sheet, pivot_df)

            # Step 3: Set up pivot tables in 03.PD_Pivot
            print(f"\n3. Setting up pivot tables...")

            # Compute last month field name from pivot date columns
            date_columns = [col for col in pivot_df.columns
                           if col not in ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY']
                           and re.match(r'\d{4}-\d{2}', str(col))]
            last_date = date_columns[-1]  # e.g., '2025-09'
            month_num = int(last_date.split('-')[1])
            month_names_list = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            last_month_field = month_names_list[month_num - 1]  # e.g., 'Sep'
            last_data_row = 3 + len(pivot_df) - 1  # data starts at row 3

            # Last column with data = first data col (C=3) + date cols + 4 formula cols - 1
            # For 6 months: 3 + 6 + 4 - 1 = 12 = 'L'
            last_col_num    = 3 + len(date_columns) + 4 - 1
            last_col_letter = excel._col_number_to_letter(last_col_num)

            excel.setup_historic_pivot_tables(
                pivot_sheet_name="03.PD_Pivot",
                data_sheet_name=historic_sheet,
                big_pivot_name="PivotTable2",
                small_pivot_name="PivotTable1",
                last_month_field=last_month_field,
                last_data_row=last_data_row,
                last_col=last_col_letter
            )

            # Save as new file with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = historic_filename.replace(".xlsb", f"_Updated_{timestamp}.xlsb")
            output_file = os.path.join(output_folder, output_filename)

            excel.save_as(output_file)
            print(f"\n4. Saved Historic file:")
            print(f"   {output_file}")
        
        print("\n" + "="*80)
        print("HISTORIC PD UPDATE COMPLETED!")
        print("="*80)
        
        return output_file