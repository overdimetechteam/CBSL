"""
Excel Automation Script for CBSL Provision Comparison
Handles both .xlsx and .xlsb files with comprehensive logging and validation
"""

import os
import sys
import time
import pandas as pd
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings('ignore')

# Auto-install dependencies
def install_dependencies():
    """Install required packages if not available"""
    required_packages = ['pandas', 'openpyxl', 'xlwings']
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Installing {package}...")
            os.system(f'pip install {package}')

install_dependencies()

import xlwings as xw

# Configuration
FILE_PATH = r"C:\Users\ASUS\Documents\WORK\Provision\Files"

class ExcelAutomation:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.log_messages = []
        self.start_time = time.time()
        
    def log(self, message, level="INFO"):
        """Log messages with timestamp"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        elapsed = time.time() - self.start_time
        log_entry = f"[{timestamp}] [{level}] [+{elapsed:.2f}s] {message}"
        print(log_entry)
        self.log_messages.append(log_entry)
    
    def find_file(self, keyword):
        """Find file containing keyword in filename"""
        self.log(f"Searching for file with keyword: '{keyword}'")
        for file in self.base_path.glob("*"):
            if keyword.lower() in file.name.lower():
                self.log(f"Found file: {file.name}", "SUCCESS")
                return file
        self.log(f"File with keyword '{keyword}' not found!", "ERROR")
        return None
    
    def find_header_row(self, df, header_keywords):
        """Find header row by searching for keywords"""
        for idx in range(len(df)):
            row = df.iloc[idx]
            row_str = ' '.join([str(cell).upper() for cell in row if pd.notna(cell)])
            if all(keyword.upper() in row_str for keyword in header_keywords):
                return idx
        return None
    
    def read_excel_file(self, file_path, sheet_name=None):
        """Read Excel file (both .xlsx and .xlsb)"""
        self.log(f"Reading file: {file_path.name}, Sheet: {sheet_name}")
        
        try:
            if file_path.suffix == '.xlsb':
                # Use xlwings for .xlsb files
                app = xw.App(visible=False)
                try:
                    wb = app.books.open(str(file_path))
                    if sheet_name:
                        sheet = wb.sheets[sheet_name]
                    else:
                        sheet = wb.sheets[0]
                    # Debug used_range info
                    used = sheet.used_range
                    self.log(f"  used_range.address: {used.address}")
                    self.log(f"  used_range starts at col {used.column}, last_cell col {used.last_cell.column}")
                    self.log(f"  Cell A3 value: {sheet.range('A3').value}")

                    # Force read from column A to the absolute last column
                    last_row = used.last_cell.row
                    # Ensure we capture all columns including any skipped by used_range
                    last_col = max(used.last_cell.column, used.column + used.columns.count - 1)
                    full_range = sheet.range((1, 1), (last_row, last_col))
                    self.log(f"  Reading range: A1 to ({last_row}, {last_col})")
                    df = full_range.options(pd.DataFrame, header=False, index=False).value
                    wb.close()
                    app.quit()
                    self.log(f"Successfully read {file_path.name} using xlwings", "SUCCESS")
                    return df
                except Exception as e:
                    app.quit()
                    raise e
            else:
                # Use pandas for .xlsx files
                if sheet_name:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                else:
                    # Read first sheet
                    df = pd.read_excel(file_path, sheet_name=0, header=None)
                self.log(f"Successfully read {file_path.name} using pandas", "SUCCESS")
                return df
        except Exception as e:
            self.log(f"Error reading file: {str(e)}", "ERROR")
            raise
    
    def step_01_copy_summary_to_portfolio(self):
        """Step 01: Copy data from Summary sheet to Portfolio sheet"""
        self.log("=" * 80)
        self.log("STEP 01: Copying Summary to Portfolio")
        self.log("=" * 80)
        step_start = time.time()
        
        # Find files
        summary_file = self.find_file("summary")
        provision_file = self.find_file("provision")
        
        if not summary_file or not provision_file:
            self.log("Required files not found!", "ERROR")
            return False
        
        # Read Summary sheet
        summary_df = self.read_excel_file(summary_file, "SUMMARY")
        
        # Debug: Print first few rows to understand structure
        self.log("First 20 rows of Summary sheet (checking for header):")
        for i in range(min(20, len(summary_df))):
            row_preview = ' | '.join([str(summary_df.iloc[i, j])[:20] for j in range(min(5, len(summary_df.columns)))])
            self.log(f"  Row {i}: {row_preview}")
        
        # Find header row in summary
        header_row_idx = self.find_header_row(summary_df, ["CONTRACT", "CLIENT"])
        if header_row_idx is None:
            self.log("Header row not found in Summary sheet!", "ERROR")
            self.log("Trying alternative search with more lenient criteria...", "WARNING")
            # Try more lenient search
            for idx in range(min(20, len(summary_df))):
                row = summary_df.iloc[idx]
                row_str = ' '.join([str(cell).upper() for cell in row if pd.notna(cell)])
                if 'CONTRACT' in row_str or 'CLIENT' in row_str:
                    header_row_idx = idx
                    self.log(f"Found potential header at row {idx}", "WARNING")
                    break
            
            if header_row_idx is None:
                self.log("Still could not find header row!", "ERROR")
                return False
        
        self.log(f"Header row found at index: {header_row_idx}")
        
        # Set header
        summary_df.columns = summary_df.iloc[header_row_idx]
        summary_df = summary_df.iloc[header_row_idx + 1:].reset_index(drop=True)
        
        # Find last row with data in first column
        first_col = summary_df.columns[0]
        last_row = summary_df[first_col].last_valid_index()
        if last_row is not None:
            summary_df = summary_df.iloc[:last_row + 1]
        
        self.log(f"Data rows in Summary: {len(summary_df)}")
        
        # Extract columns A to AG (first 33 columns)
        data_frame_01 = summary_df.iloc[:, :33].copy()
        
        self.log("Column names from Summary (A to AG):")
        for i, col in enumerate(data_frame_01.columns[:10]):  # Show first 10
            self.log(f"  Column {chr(65 + i)}: {col}")
        self.log(f"  ... and {len(data_frame_01.columns) - 10} more columns")
        
        # Open provision file once for both reading header info and writing
        app = xw.App(visible=False)
        try:
            wb = app.books.open(str(provision_file))
            sheet = wb.sheets["Portfolio"]

            # Read only the first 10 rows to find the header (instead of reading entire sheet)
            self.log("Reading Portfolio header rows...")
            header_scan_range = sheet.range((1, 1), (10, 33))
            header_scan = header_scan_range.value  # list of lists

            # Debug: Print first rows of Portfolio sheet
            self.log("First 10 rows of Portfolio sheet (checking for header):")
            for i, row in enumerate(header_scan):
                row_preview = ' | '.join([str(cell)[:25] for cell in row[:5]])
                self.log(f"  Row {i}: {row_preview}")

            # Find header row
            portfolio_header_idx = None
            for idx, row in enumerate(header_scan):
                row_str = ' '.join([str(cell).upper() for cell in row if cell is not None])
                if 'CONTRACT' in row_str and 'CLIENT' in row_str:
                    portfolio_header_idx = idx
                    self.log(f"Found Portfolio header at row {idx}", "SUCCESS")
                    break

            if portfolio_header_idx is None:
                self.log("Header row not found in Portfolio sheet!", "ERROR")
                wb.close()
                app.quit()
                return False

            self.log(f"Portfolio header row found at index: {portfolio_header_idx}")

            # Get portfolio columns from the header row
            portfolio_cols = header_scan[portfolio_header_idx]

            self.log("Column names from Portfolio (A to AG):")
            for i, col in enumerate(portfolio_cols[:10]):  # Show first 10
                self.log(f"  Column {chr(65 + i)}: {col}")
            self.log(f"  ... and {len(portfolio_cols) - 10} more columns")

            # Compare column names
            self.log("\nColumn mapping validation:")
            mismatches = []
            for i, (sum_col, port_col) in enumerate(zip(data_frame_01.columns, portfolio_cols)):
                if str(sum_col).strip().upper() != str(port_col).strip().upper():
                    mismatches.append(f"  Column {chr(65 + i)}: '{sum_col}' != '{port_col}'")

            if mismatches:
                self.log("WARNING: Column name mismatches detected:", "WARNING")
                for mismatch in mismatches[:5]:  # Show first 5
                    self.log(mismatch, "WARNING")
                if len(mismatches) > 5:
                    self.log(f"  ... and {len(mismatches) - 5} more mismatches", "WARNING")
                self.log("Proceeding with positional mapping...")
            else:
                self.log("All column names match!", "SUCCESS")

            # Rename columns to match portfolio
            data_frame_01.columns = portfolio_cols

            # Get old data row count from used range
            last_used_row = sheet.used_range.last_cell.row
            old_data_rows = last_used_row - (portfolio_header_idx + 1)  # rows below header
            self.log(f"Clearing {old_data_rows} existing rows in Portfolio")
            
            if old_data_rows > 0:
                clear_range = sheet.range((portfolio_header_idx + 2, 1), (portfolio_header_idx + 1 + old_data_rows, 33))
                clear_range.clear_contents()
            
            # Write new data
            new_data_rows = len(data_frame_01)
            self.log(f"Writing {new_data_rows} new rows to Portfolio")
            
            sheet.range((portfolio_header_idx + 2, 1)).value = data_frame_01.values
            
            # Handle formulas in columns after AG
            if new_data_rows != old_data_rows:
                self.log(f"Row count changed: {old_data_rows} -> {new_data_rows}")
                
                if new_data_rows > old_data_rows and old_data_rows > 0:
                    # Copy formulas down using bulk range copy (not cell-by-cell)
                    self.log("Extending formulas for additional rows...")
                    last_col = sheet.used_range.last_cell.column

                    if last_col > 33:
                        source_row_num = portfolio_header_idx + 2
                        source_range = sheet.range((source_row_num, 34), (source_row_num, last_col))

                        first_new_row = portfolio_header_idx + 2 + old_data_rows
                        last_new_row = portfolio_header_idx + 1 + new_data_rows
                        target_range = sheet.range((first_new_row, 34), (last_new_row, last_col))

                        source_range.copy(target_range)
                        extra_rows = new_data_rows - old_data_rows
                        self.log(f"Extended formulas to {extra_rows} new rows (columns AH to col {last_col})", "SUCCESS")
                    else:
                        self.log("No formula columns found after AG", "INFO")
                
                elif new_data_rows < old_data_rows:
                    # Clear extra rows
                    self.log("Removing excess rows...")
                    delete_range = sheet.range((portfolio_header_idx + 2 + new_data_rows, 1),
                                              (portfolio_header_idx + 1 + old_data_rows, sheet.used_range.last_cell.column))
                    delete_range.clear_contents()
            
            wb.save()
            wb.close()
            app.quit()
            
            self.log(f"Step 01 completed in {time.time() - step_start:.2f}s", "SUCCESS")
            return True
            
        except Exception as e:
            try:
                app.quit()
            except:
                pass
            self.log(f"Error in Step 01: {str(e)}", "ERROR")
            return False
    
    def step_02_sync_property_mortgages(self):
        """Step 02: Sync Property Mortgages data"""
        self.log("=" * 80)
        self.log("STEP 02: Syncing Property Mortgages")
        self.log("=" * 80)
        step_start = time.time()
        
        # Find files
        property_file = self.find_file("Property mortgages")
        provision_file = self.find_file("provision")
        
        if not property_file or not provision_file:
            self.log("Required files not found!", "ERROR")
            return False
        
        # Read both sheets
        property_df = self.read_excel_file(property_file)
        property_mortgage_df = self.read_excel_file(provision_file, "PropertyMortgage")
        
        # Find headers
        prop_header_idx = 0  # Assuming first row is header
        prop_mortgage_header_idx = 4  # Given in requirements
        
        property_df.columns = property_df.iloc[prop_header_idx]
        property_df = property_df.iloc[prop_header_idx + 1:].reset_index(drop=True)

        property_mortgage_df.columns = property_mortgage_df.iloc[prop_mortgage_header_idx]
        existing_data = property_mortgage_df.iloc[prop_mortgage_header_idx + 1:].reset_index(drop=True)

        self.log(f"Property Mortgages file rows: {len(property_df)}")
        self.log(f"PropertyMortgage sheet rows: {len(existing_data)}")

        # Debug: show column names from both sources
        self.log("Property Mortgages file columns:")
        for i, col in enumerate(property_df.columns):
            self.log(f"  Col {i}: '{col}'")
        self.log("PropertyMortgage sheet columns:")
        for i, col in enumerate(existing_data.columns):
            self.log(f"  Col {i}: '{col}'")

        # Compare and find new records
        contract_col_july = None
        contract_col_provision = None

        # Find CONTRACT_NO column in Property Mortgages file
        for col in property_df.columns:
            col_str = str(col).upper().replace(' ', '_')
            if 'CONTRACT' in col_str and ('NO' in col_str or '_NO' in col_str):
                contract_col_july = col
                break

        # Find CONTRACT_NO column in PropertyMortgage sheet (more flexible search)
        for col in existing_data.columns:
            col_str = str(col).upper().replace(' ', '_')
            if 'CONTRACT' in col_str and ('NO' in col_str or '_NO' in col_str):
                contract_col_provision = col
                break

        # If still not found, try matching with known typos and variations
        if contract_col_provision is None:
            for col in existing_data.columns:
                col_str = str(col).upper().strip()
                if col_str in ('CONTRACT_NO', 'CONTRACT NO', 'CONTRACTNO', 'CONTRACT',
                               'CONTRTACT NO', 'CONTRTACT_NO'):
                    contract_col_provision = col
                    break

        self.log(f"Contract column in Property Mortgages file: {contract_col_july}")
        self.log(f"Contract column in PropertyMortgage sheet: {contract_col_provision}")
        
        if contract_col_july and contract_col_provision:
            property_contracts = set(property_df[contract_col_july].dropna().astype(str))
            existing_contracts = set(existing_data[contract_col_provision].dropna().astype(str))
            
            new_contracts = property_contracts - existing_contracts
            
            if new_contracts:
                self.log(f"Found {len(new_contracts)} new contracts to add:", "WARNING")
                for contract in list(new_contracts)[:5]:  # Show first 5
                    self.log(f"  - {contract}")
                if len(new_contracts) > 5:
                    self.log(f"  ... and {len(new_contracts) - 5} more")
                
                self.log("Note: Adding new records requires manual review", "WARNING")
            else:
                self.log("No new contracts found - sheets are in sync", "SUCCESS")
        else:
            self.log("Could not identify CONTRACT columns for comparison", "WARNING")
        
        self.log(f"Step 02 completed in {time.time() - step_start:.2f}s", "SUCCESS")
        return True
    
    def step_03_update_property_mortgage_columns(self, reporting_date=None):
        """Step 03: Update specific columns in PropertyMortgage sheet (matched by contract number)"""
        self.log("=" * 80)
        self.log("STEP 03: Updating PropertyMortgage Columns")
        self.log("=" * 80)
        step_start = time.time()

        property_file = self.find_file("Property mortgages")
        provision_file = self.find_file("provision")

        if not property_file or not provision_file:
            self.log("Required files not found!", "ERROR")
            return False

        # Read Property Mortgages source file
        property_df = self.read_excel_file(property_file)
        property_df.columns = property_df.iloc[0]
        property_df = property_df.iloc[1:].reset_index(drop=True)

        # Find source columns
        src_contract_col = src_client_code_col = src_property_val_col = None
        src_valuation_date_col = src_activation_date_col = None

        for col in property_df.columns:
            col_str = str(col).upper()
            if 'CONTRACT' in col_str and 'NO' in col_str:
                src_contract_col = col
            elif 'CLIENT' in col_str and 'CODE' in col_str:
                src_client_code_col = col
            elif 'PROPERTY' in col_str and 'VALUE' in col_str:
                src_property_val_col = col
            elif 'VALUATION' in col_str and 'DATE' in col_str:
                src_valuation_date_col = col
            elif 'ACTIVATION' in col_str and 'DATE' in col_str:
                src_activation_date_col = col

        self.log("Source file columns found:")
        self.log(f"  CONTRACT_NO: {src_contract_col}")
        self.log(f"  CLIENT_CODE: {src_client_code_col}")
        self.log(f"  Property value: {src_property_val_col}")
        self.log(f"  Valuation Date: {src_valuation_date_col}")
        self.log(f"  ACTIVATION_DATE: {src_activation_date_col}")

        if not src_contract_col:
            self.log("Could not find CONTRACT column in source file!", "ERROR")
            return False

        # Build lookup dict
        source_lookup = {}
        for _, row in property_df.iterrows():
            contract = str(row[src_contract_col]).strip() if pd.notna(row[src_contract_col]) else None
            if contract and contract != 'nan':
                source_lookup[contract] = {
                    'client_code': row[src_client_code_col] if src_client_code_col else None,
                    'property_value': row[src_property_val_col] if src_property_val_col else None,
                    'valuation_date': row[src_valuation_date_col] if src_valuation_date_col else None,
                    'activation_date': row[src_activation_date_col] if src_activation_date_col else None,
                }
        self.log(f"Source file has {len(source_lookup)} contracts")

        def is_valid(val):
            """Return True only if value is non-null and non-nan"""
            if val is None:
                return False
            if isinstance(val, float) and pd.isna(val):
                return False
            return True

        def ensure_date(val):
            """Convert value to proper datetime if it's a date string"""
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return None
            if isinstance(val, datetime):
                return val
            if isinstance(val, str):
                # Try to parse common date formats
                for fmt in ['%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d.%m.%Y', '%d-%m-%Y']:
                    try:
                        return datetime.strptime(val.strip(), fmt)
                    except:
                        continue
            return val  # Return as-is if can't convert

        def read_single_col(rng):
            """Read a single-column xlwings range into a flat list"""
            vals = rng.value
            if vals is None:
                return []
            if not isinstance(vals, list):
                return [vals]
            # xlwings returns list of scalars for single-column ranges
            return vals

        app = xw.App(visible=False)
        try:
            wb = app.books.open(str(provision_file))
            sheet = wb.sheets["PropertyMortgage"]

            # Update A3 with reporting date
            if reporting_date:
                sheet.range('A3').value = reporting_date
                self.log(f"Updated A3 date to {reporting_date.strftime('%d/%m/%Y')}", "SUCCESS")

            header_row = 5
            data_start_row = header_row + 1
            last_row = sheet.used_range.last_cell.row
            last_col = sheet.used_range.last_cell.column

            if last_row < data_start_row:
                self.log("No data rows in PropertyMortgage sheet", "WARNING")
                wb.close()
                app.quit()
                return True

            n_rows = last_row - data_start_row + 1
            self.log(f"Data rows: {n_rows} (rows {data_start_row}–{last_row})")

            # --- Step 1: Unmerge column C (VALUATION) in data range ---
            # Writing to merged cells breaks the merge, so unmerge first then re-merge after
            self.log("Unmerging VALUATION column (col C)...")
            seen_merge_starts = set()
            for i in range(n_rows):
                cell_api = sheet.range((data_start_row + i, 3)).api
                if cell_api.MergeCells:
                    merge_start = cell_api.MergeArea.Row
                    if merge_start not in seen_merge_starts:
                        cell_api.MergeArea.UnMerge()
                        seen_merge_starts.add(merge_start)
            self.log(f"  Unmerged {len(seen_merge_starts)} groups")

            # --- Step 2: Read existing contracts from col A ---
            existing_contracts = read_single_col(
                sheet.range((data_start_row, 1), (last_row, 1))
            )
            existing_contracts = [
                str(c).strip() if c is not None else None for c in existing_contracts
            ]
            self.log(f"Read {len(existing_contracts)} existing contracts")

            # --- Step 2b: Delete blank rows (col A empty) from data range ---
            blank_indices = [
                i for i, c in enumerate(existing_contracts)
                if not c or c == 'None' or c == 'nan'
            ]
            if blank_indices:
                self.log(f"Removing {len(blank_indices)} blank row(s) from data range...")
                # Delete bottom-to-top so row numbers don't shift mid-loop
                for i in reversed(blank_indices):
                    excel_row = data_start_row + i
                    sheet.range((excel_row, 1)).api.EntireRow.Delete()
                    self.log(f"  Deleted blank row {excel_row}")
                # Rebuild state after deletions
                existing_contracts = [c for c in existing_contracts if c and c != 'None' and c != 'nan']
                last_row = data_start_row + len(existing_contracts) - 1
                n_rows = len(existing_contracts)
                self.log(f"  {len(existing_contracts)} rows remain after cleanup")

            # --- Step 3: Update matched rows (only write non-null values) ---
            matched = 0
            matched_set = set()
            formula_template_row = None
            unmatched_rows = []

            for i, contract in enumerate(existing_contracts):
                excel_row = data_start_row + i
                if contract and contract != 'nan' and contract in source_lookup:
                    src = source_lookup[contract]
                    # Col B: CLIENT CODE
                    if is_valid(src['client_code']):
                        sheet.range((excel_row, 2)).value = src['client_code']
                    # Col C: VALUATION - only write if source has a real value
                    if is_valid(src['property_value']):
                        sheet.range((excel_row, 3)).value = src['property_value']
                    # Col D: VALUATION DATE (ensure proper date format)
                    if is_valid(src['valuation_date']):
                        date_val = ensure_date(src['valuation_date'])
                        if date_val:
                            sheet.range((excel_row, 4)).value = date_val
                            sheet.range((excel_row, 4)).api.NumberFormat = "dd/mm/yyyy"
                    # Col H: DATE OF ACTIVATION (ensure proper date format)
                    if is_valid(src['activation_date']):
                        date_val = ensure_date(src['activation_date'])
                        if date_val:
                            sheet.range((excel_row, 8)).value = date_val
                            sheet.range((excel_row, 8)).api.NumberFormat = "dd/mm/yyyy"
                    matched += 1
                    matched_set.add(contract)
                    formula_template_row = excel_row
                else:
                    unmatched_rows.append((excel_row, contract))

            self.log(f"Updated {matched} matched rows", "SUCCESS")
            if unmatched_rows:
                self.log(f"{len(unmatched_rows)} rows had no source match (existing values preserved):", "WARNING")
                for row_num, c in unmatched_rows[:5]:
                    self.log(f"  Row {row_num}: '{c}'")
                if len(unmatched_rows) > 5:
                    self.log(f"  ... and {len(unmatched_rows) - 5} more")

            # --- Step 3b: Fill blank VALUATION DATEs within CLIENT CODE groups ---
            # Rows sharing the same CLIENT CODE (merged VALUATION group) should all
            # carry the same VALUATION DATE. Propagate the first available date in
            # each consecutive group to any blank cells below it.
            self.log("Filling blank VALUATION DATEs within CLIENT CODE groups...")
            dates_filled = 0
            client_codes_now = read_single_col(sheet.range((data_start_row, 2), (last_row, 2)))
            date_col_now     = read_single_col(sheet.range((data_start_row, 4), (last_row, 4)))
            i = 0
            while i < len(client_codes_now):
                code = client_codes_now[i]
                if code is None:
                    i += 1
                    continue
                # Find extent of this consecutive client code group
                j = i + 1
                while j < len(client_codes_now) and client_codes_now[j] == code:
                    j += 1
                if j > i + 1:  # Multi-row group — propagate date if any row has one
                    group_date = next(
                        (date_col_now[k] for k in range(i, j) if date_col_now[k] is not None),
                        None
                    )
                    if group_date is not None:
                        for k in range(i, j):
                            if date_col_now[k] is None:
                                excel_row_k = data_start_row + k
                                sheet.range((excel_row_k, 4)).value = group_date
                                sheet.range((excel_row_k, 4)).api.NumberFormat = "dd/mm/yyyy"
                                date_col_now[k] = group_date
                                dates_filled += 1
                i = j
            self.log(f"  Filled {dates_filled} blank VALUATION DATE cells within client groups",
                     "SUCCESS")

            # --- Step 4: Re-merge column C by consecutive CLIENT CODE groups ---
            # Rows sharing the same CLIENT CODE get one merged VALUATION cell
            self.log("Re-merging VALUATION column by CLIENT CODE groups...")
            client_codes = read_single_col(
                sheet.range((data_start_row, 2), (last_row, 2))
            )

            merge_count = 0
            app.api.DisplayAlerts = False  # Suppress "merging discards values" dialog
            try:
                i = 0
                while i < len(client_codes):
                    code = client_codes[i]
                    if code is not None:
                        j = i + 1
                        while j < len(client_codes) and client_codes[j] == code:
                            j += 1
                        if j > i + 1:  # Two or more consecutive rows with same CLIENT CODE
                            merge_rng = sheet.range(
                                (data_start_row + i, 3),
                                (data_start_row + j - 1, 3)
                            )
                            merge_rng.api.Merge()
                            merge_rng.api.VerticalAlignment = -4108   # xlVAlignCenter
                            merge_rng.api.HorizontalAlignment = -4108 # xlHAlignCenter
                            merge_count += 1
                        i = j
                    else:
                        i += 1
            finally:
                app.api.DisplayAlerts = True  # Always restore alerts
            self.log(f"  Created {merge_count} merge groups", "SUCCESS")

            # --- Step 5: Append new contracts (copy formulas from template row) ---
            new_contracts = [c for c in source_lookup if c not in matched_set]
            if new_contracts:
                self.log(f"Appending {len(new_contracts)} new contracts...")
                append_row = last_row + 1

                # Use last matched row as formula template, fallback to first data row
                if formula_template_row is None:
                    formula_template_row = data_start_row

                for contract in new_contracts:
                    src = source_lookup[contract]
                    # Copy entire template row to get formulas in E,F,G,M,N,O etc.
                    src_rng = sheet.range((formula_template_row, 1), (formula_template_row, last_col))
                    tgt_rng = sheet.range((append_row, 1), (append_row, last_col))
                    src_rng.copy(tgt_rng)
                    # Overwrite with actual values
                    sheet.range((append_row, 1)).value = contract
                    if is_valid(src['client_code']):
                        sheet.range((append_row, 2)).value = src['client_code']
                    if is_valid(src['property_value']):
                        sheet.range((append_row, 3)).value = src['property_value']
                    # Col D: VALUATION DATE (ensure proper date format)
                    if is_valid(src['valuation_date']):
                        date_val = ensure_date(src['valuation_date'])
                        if date_val:
                            sheet.range((append_row, 4)).value = date_val
                            sheet.range((append_row, 4)).api.NumberFormat = "dd/mm/yyyy"
                    # Col H: DATE OF ACTIVATION (ensure proper date format)
                    if is_valid(src['activation_date']):
                        date_val = ensure_date(src['activation_date'])
                        if date_val:
                            sheet.range((append_row, 8)).value = date_val
                            sheet.range((append_row, 8)).api.NumberFormat = "dd/mm/yyyy"
                    self.log(f"  Appended contract '{contract}' at row {append_row}")
                    append_row += 1
                self.log(f"Appended {len(new_contracts)} new rows with formulas", "SUCCESS")
            else:
                self.log("No new contracts to append", "SUCCESS")

            # --- Step 6: Correct column P (Legal status) ---
            # "Legal" should appear ONLY when J, K, L, M all equal "CLOSED"
            self.log("Correcting Legal status in column P...")
            final_last_row = sheet.used_range.last_cell.row
            legal_set = 0
            legal_cleared = 0
            for excel_row in range(data_start_row, final_last_row + 1):
                jklm = sheet.range((excel_row, 10), (excel_row, 13)).value  # J=10, K=11, L=12, M=13
                if jklm is None:
                    all_closed = False
                else:
                    all_closed = all(str(v).strip().upper() == 'CLOSED' for v in jklm)
                p_cell = sheet.range((excel_row, 16))  # P = column 16
                current_p = p_cell.value
                if all_closed:
                    if current_p != 'Legal':
                        p_cell.value = 'Legal'
                        legal_set += 1
                else:
                    if current_p == 'Legal':
                        p_cell.value = None
                        legal_cleared += 1
            self.log(f"  Legal: set in {legal_set} rows, cleared from {legal_cleared} rows", "SUCCESS")

            # --- Step 7: Fix formulas in column F based on merge status ---
            # Column F should have:
            #   - =E[row]/$I$[row]*$C$[merge_start_row]  for rows where column C (VALUATION) is merged
            #   - =C[row]                                 for rows where column C is NOT merged
            # When cells are merged, all rows in the merge group reference the first row's cell address
            self.log("Fixing formulas in column F based on VALUATION merge status...")
            formulas_fixed = 0
            for excel_row in range(data_start_row, final_last_row + 1):
                c_cell = sheet.range((excel_row, 3))  # Column C
                f_cell = sheet.range((excel_row, 6))  # Column F

                # Check if cell C is merged
                is_merged = c_cell.api.MergeCells

                if is_merged:
                    # Get the first row of the merged cell area (e.g., if C10:C12 merged, this returns 10)
                    merge_start_row = c_cell.api.MergeArea.Row
                    # For merged cells: E and I use current row, C uses merge start row
                    # Example: Row 11 with C10:C12 merged - =E11/$I$11*$C$10
                    expected_formula = f"=IFERROR(E{excel_row}/$I${excel_row}*$C${merge_start_row},0)"
                else:
                    # For unmerged cells: =C[row]
                    expected_formula = f"=C{excel_row}"

                current_formula = f_cell.formula
                if current_formula != expected_formula:
                    f_cell.formula = expected_formula
                    formulas_fixed += 1

            self.log(f"  Fixed {formulas_fixed} formulas in column F", "SUCCESS")

            # --- Step 8: Fix all existing dates in columns D and H ---
            # Some rows may have dates in wrong format (e.g., "14.09.2022" with dots)
            # from previous data. Convert all dates to proper Excel date format.
            self.log("Fixing all existing dates in VALUATION DATE (D) and DATE OF ACTIVATION (H)...")
            dates_fixed = 0
            for excel_row in range(data_start_row, final_last_row + 1):
                # Fix VALUATION DATE (column D)
                d_cell = sheet.range((excel_row, 4))
                d_val = d_cell.value
                if d_val is not None:
                    # If it's a string, try to convert it to proper date
                    if isinstance(d_val, str):
                        date_val = ensure_date(d_val)
                        if date_val and isinstance(date_val, datetime):
                            d_cell.value = date_val
                            d_cell.api.NumberFormat = "dd/mm/yyyy"
                            dates_fixed += 1
                    elif isinstance(d_val, datetime):
                        # Already a datetime, just ensure proper format
                        d_cell.api.NumberFormat = "dd/mm/yyyy"

                # Fix DATE OF ACTIVATION (column H)
                h_cell = sheet.range((excel_row, 8))
                h_val = h_cell.value
                if h_val is not None:
                    # If it's a string, try to convert it to proper date
                    if isinstance(h_val, str):
                        date_val = ensure_date(h_val)
                        if date_val and isinstance(date_val, datetime):
                            h_cell.value = date_val
                            h_cell.api.NumberFormat = "dd/mm/yyyy"
                            dates_fixed += 1
                    elif isinstance(h_val, datetime):
                        # Already a datetime, just ensure proper format
                        h_cell.api.NumberFormat = "dd/mm/yyyy"

            self.log(f"  Fixed {dates_fixed} date values to proper format", "SUCCESS")

            # --- Step 9: Update pivot table source range to cover full data ---
            # The pivot table in columns R and S must reference all data rows
            # (header row through final last row) so newly appended rows are included.
            self.log("Updating PropertyMortgage pivot table source range...")
            try:
                pivot_count = sheet.api.PivotTables().Count
                self.log(f"  Found {pivot_count} pivot table(s) in PropertyMortgage sheet")
                if pivot_count > 0:
                    # Source data covers header row to final data row, cols A–O (1–15).
                    # Columns P and Q have no headers; the pivot table itself sits at R and S.
                    # The old source was R5C1:R36C15 (= A5:O36), so we match that column extent.
                    sheet_name = sheet.name
                    new_source_address = f"'{sheet_name}'!$A${header_row}:$O${final_last_row}"
                    self.log(f"  New pivot source range: {new_source_address}")
                    for i in range(1, pivot_count + 1):
                        pt = sheet.api.PivotTables(i)
                        try:
                            old_source = pt.SourceData
                            self.log(f"  Pivot {i} old SourceData: {old_source}")
                            new_cache = wb.api.PivotCaches().Create(
                                SourceType=1,   # xlDatabase
                                SourceData=new_source_address
                            )
                            pt.ChangePivotCache(new_cache)
                            pt.RefreshTable()
                            self.log(f"  Pivot table {i} source updated and refreshed", "SUCCESS")
                        except Exception as pe:
                            self.log(f"  Could not update pivot {i} source: {str(pe)}", "WARNING")
                            try:
                                pt.RefreshTable()
                                self.log(f"  Pivot table {i} refreshed (source unchanged)", "WARNING")
                            except Exception as pe2:
                                self.log(f"  Pivot table {i} refresh also failed: {str(pe2)}", "WARNING")
            except Exception as pte:
                self.log(f"  Pivot table update error: {str(pte)}", "WARNING")

            wb.save()
            wb.close()
            app.quit()

            self.log(f"Step 03 completed in {time.time() - step_start:.2f}s", "SUCCESS")
            return True

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            self.log(f"Error in Step 03: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            return False
    
    def step_04_update_comparison_tables(self):
        """Step 04: Update Comparison sheet tables"""
        self.log("=" * 80)
        self.log("STEP 04: Updating Comparison Sheet")
        self.log("=" * 80)
        step_start = time.time()

        provision_file = self.find_file("provision")
        summary_file = self.find_file("summary")
        if not provision_file:
            self.log("Provision file not found!", "ERROR")
            return False

        # Derive current month from summary filename (e.g. "Summary_2025-08-31" -> "Aug-25")
        new_month_label = None
        if summary_file:
            import re
            match = re.search(r'(\d{4})-(\d{2})-\d{2}', summary_file.name)
            if match:
                year, month = int(match.group(1)), int(match.group(2))
                month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                new_month_label = f"{month_names[month - 1]}-{str(year)[-2:]}"
        self.log(f"New month label: {new_month_label}")

        app = xw.App(visible=False)
        try:
            wb = app.books.open(str(provision_file))
            sheet = wb.sheets["Comparison"]

            month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                           'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

            def read_label(cell_ref):
                """Read a month label cell, handling case where Excel re-interpreted it as a date."""
                val = sheet.range(cell_ref).value
                if val is None:
                    return ''
                if isinstance(val, datetime):
                    # Excel turned "Aug-25" into a date object — convert back to Mon-YY
                    return f"{month_names[val.month - 1]}-{str(val.year)[-2:]}"
                return str(val).strip()

            def write_label(cell_ref, label):
                """Write a month label as plain text so Excel won't reformat it as a date."""
                cell = sheet.range(cell_ref)
                cell.api.NumberFormat = "@"  # Force text format
                cell.value = label

            # --- Step 1: Read current labels ---
            current_label = read_label('A14')   # e.g. "Jul-25"  (current month slot)
            previous_label = read_label('E14')  # e.g. "Jun-25"  (previous month slot)
            self.log(f"Current labels: A14='{current_label}', E14='{previous_label}'")

            # --- Same-month guard: don't shift if we're re-running for the same month ---
            same_month = (new_month_label is not None and current_label == new_month_label)
            if same_month:
                self.log(f"Same-month re-run detected ({new_month_label}) — skipping table shift", "WARNING")
            else:
                # --- Step 2: Shift current month table (A) -> previous month position (E) ---
                self.log(f"Shifting '{current_label}' - previous month position...")
                jul_table_data = sheet.range('A15:C23').value
                sheet.range('E15:G23').value = jul_table_data
                write_label('E14', current_label)
                self.log(f"  Wrote '{current_label}' data to E14:G23", "SUCCESS")

            # --- Step 3: Update pivot source range then refresh ---
            self.log("Updating Comparison pivot table source range...")
            try:
                port_sheet = wb.sheets["Portfolio"]
                port_last_row = port_sheet.used_range.last_cell.row
                port_last_col = port_sheet.used_range.last_cell.column

                # Find Portfolio header row (same logic as Step 01)
                port_header_scan = port_sheet.range((1, 1), (10, 5)).value
                port_header_row = 1  # fallback
                for idx, row in enumerate(port_header_scan):
                    row_str = ' '.join([str(c).upper() for c in row if c is not None])
                    if 'CONTRACT' in row_str and 'CLIENT' in row_str:
                        port_header_row = idx + 1  # convert to 1-based Excel row
                        self.log(f"  Portfolio header at Excel row {port_header_row}")
                        break

                # Convert column number to letter (e.g. 33 -> 'AG')
                def col_letter(n):
                    s = ''
                    while n > 0:
                        n, r = divmod(n - 1, 26)
                        s = chr(65 + r) + s
                    return s

                new_source = f"'Portfolio'!$A${port_header_row}:${col_letter(port_last_col)}${port_last_row}"
                self.log(f"  New pivot source range: {new_source}")

                pivot_count = sheet.api.PivotTables().Count
                self.log(f"  Found {pivot_count} pivot table(s)")
                for i in range(1, pivot_count + 1):
                    pt = sheet.api.PivotTables(i)
                    try:
                        self.log(f"  Pivot {i} old SourceData: {pt.SourceData}")
                        new_cache = wb.api.PivotCaches().Create(
                            SourceType=1,   # xlDatabase
                            SourceData=new_source
                        )
                        pt.ChangePivotCache(new_cache)
                        pt.RefreshTable()
                        self.log(f"  Pivot {i} source updated and refreshed", "SUCCESS")
                    except Exception as pe:
                        self.log(f"  Could not update pivot {i} source: {str(pe)}", "WARNING")
                        try:
                            pt.RefreshTable()
                            self.log(f"  Pivot {i} refreshed (source unchanged)", "WARNING")
                        except Exception as pe2:
                            self.log(f"  Pivot {i} refresh also failed: {str(pe2)}", "WARNING")
            except Exception as e:
                self.log(f"  Pivot source update error: {str(e)}", "WARNING")
                try:
                    for pc in wb.api.PivotCaches():
                        pc.Refresh()
                    self.log("  Refreshed via PivotCaches (source unchanged)", "WARNING")
                except Exception as e2:
                    self.log(f"  PivotCaches refresh also failed: {str(e2)}", "ERROR")

            # --- Step 4: Copy refreshed PIVOT values -> current month position (A) ---
            self.log(f"Copying refreshed PIVOT values to current month position...")
            pivot_data = sheet.range('A5:C12').value
            sheet.range('A16:C23').value = pivot_data
            pivot_headers = sheet.range('A4:C4').value
            sheet.range('A15:C15').value = pivot_headers
            if new_month_label:
                write_label('A14', new_month_label)
                self.log(f"  Updated current month label to '{new_month_label}'", "SUCCESS")
            self.log(f"  Wrote PIVOT data to A15:C23", "SUCCESS")

            # --- Step 5: Update Month to Month summary labels ---
            if new_month_label:
                prev_for_summary = current_label if not same_month else previous_label
                sheet.range('I5').value = f"{new_month_label} CBSL Provision"
                sheet.range('I6').value = f"{prev_for_summary} CBSL Provision"
                self.log(f"  Month-to-Month labels: {new_month_label} / {prev_for_summary}", "SUCCESS")

            wb.save()
            wb.close()
            app.quit()

            self.log(f"Step 04 completed in {time.time() - step_start:.2f}s", "SUCCESS")
            return True

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            self.log(f"Error in Step 04: {str(e)}", "ERROR")
            return False
    
    def step_05_update_repossessed_reporting_date(self, reporting_date):
        """Step 05: Update Reporting Date in RepossessedList sheet"""
        self.log("=" * 80)
        self.log("STEP 05: Updating RepossessedList Reporting Date")
        self.log("=" * 80)
        step_start = time.time()

        provision_file = self.find_file("provision")
        if not provision_file:
            self.log("Provision file not found!", "ERROR")
            return False

        app = xw.App(visible=False)
        try:
            wb = app.books.open(str(provision_file))
            sheet = wb.sheets["RepossessedList"]

            old_date = sheet.range('K2').value
            self.log(f"Current Reporting Date: {old_date}")

            sheet.range('K2').value = reporting_date
            self.log(f"Updated Reporting Date to: {reporting_date.strftime('%m/%d/%Y')}", "SUCCESS")

            wb.save()
            wb.close()
            app.quit()

            self.log(f"Step 05 completed in {time.time() - step_start:.2f}s", "SUCCESS")
            return True

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            self.log(f"Error in Step 05: {str(e)}", "ERROR")
            return False

    def step_06_update_portfolio_ecl_columns(self):
        """Step 06: Update Portfolio sheet columns AP-AW from ECL Portfolio file"""
        self.log("=" * 80)
        self.log("STEP 06: Updating Portfolio ECL Columns (AP-AW)")
        self.log("=" * 80)
        step_start = time.time()

        ecl_file = self.find_file("Portfolio")
        provision_file = self.find_file("provision")

        if not ecl_file or not provision_file:
            self.log("Required files not found!", "ERROR")
            return False

        def col_letter_to_idx(col_str):
            """Convert column letter(s) like 'A', 'M', 'AP' to 0-indexed integer."""
            if not col_str or not isinstance(col_str, str):
                return None
            col_str = col_str.strip().upper()
            if not col_str.isalpha():
                return None
            num = 0
            for char in col_str:
                num = num * 26 + (ord(char) - ord('A') + 1)
            return num - 1  # 0-indexed

        def idx_to_col_letter(idx):
            """Convert 0-indexed integer to column letter(s) like 'A', 'M', 'AP'."""
            result = ''
            n = idx + 1
            while n:
                n, r = divmod(n - 1, 26)
                result = chr(65 + r) + result
            return result

        # ---------------------------------------------------------------
        # Phase 1: Read the ECL Portfolio file
        # ---------------------------------------------------------------
        self.log(f"Reading ECL file: {ecl_file.name}")
        app = xw.App(visible=False)
        try:
            ecl_wb = app.books.open(str(ecl_file))

            # Detect the Portfolio sheet
            sheet_names = [s.name for s in ecl_wb.sheets]
            self.log(f"ECL file sheets: {sheet_names}")
            ecl_sheet = None
            for s in ecl_wb.sheets:
                if 'portfolio' in s.name.lower():
                    ecl_sheet = s
                    self.log(f"Using ECL sheet: '{s.name}'")
                    break
            if ecl_sheet is None:
                ecl_sheet = ecl_wb.sheets[0]
                self.log(f"No 'Portfolio' sheet found — using first sheet: '{ecl_sheet.name}'", "WARNING")

            # Read all data from ECL sheet (force A1 origin)
            used = ecl_sheet.used_range
            ecl_last_row = used.last_cell.row
            ecl_last_col = max(used.last_cell.column, used.column + used.columns.count - 1)
            self.log(f"ECL used range: {ecl_last_row} rows × {ecl_last_col} cols")

            full_range = ecl_sheet.range((1, 1), (ecl_last_row, ecl_last_col))
            ecl_df_raw = full_range.options(pd.DataFrame, header=False, index=False).value
            self.log(f"ECL data loaded: {ecl_df_raw.shape[0]} rows × {ecl_df_raw.shape[1]} cols", "SUCCESS")

            ecl_wb.close()
            app.quit()

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            self.log(f"Error reading ECL file: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            return False

        # Find header row in ECL data
        # Use "CONTRACT NO" (not just "CONTRACT") to avoid matching "CONTRACT COUNT VERIFY" in row 1
        ecl_header_idx = self.find_header_row(ecl_df_raw, ["CONTRACT NO"])
        if ecl_header_idx is None:
            # Fallback: try with CLIENT CODE as an additional discriminator
            ecl_header_idx = self.find_header_row(ecl_df_raw, ["CONTRACT", "CLIENT CODE"])
        if ecl_header_idx is None:
            self.log("Header row not found in ECL Portfolio data!", "ERROR")
            return False
        self.log(f"ECL header row: index {ecl_header_idx} (Excel row {ecl_header_idx + 1})")

        # Debug: show ECL header columns
        ecl_header_row = ecl_df_raw.iloc[ecl_header_idx]
        self.log("ECL column headers (non-empty):")
        for i, col in enumerate(ecl_header_row):
            if col is not None and str(col).strip():
                self.log(f"  {idx_to_col_letter(i)} (idx {i}): '{col}'")

        # Find CONTRACT NO column position in ECL
        ecl_contract_col_pos = None
        for j, val in enumerate(ecl_header_row):
            if val and 'CONTRACT' in str(val).upper() and 'NO' in str(val).upper():
                ecl_contract_col_pos = j
                self.log(f"ECL CONTRACT NO column: {idx_to_col_letter(j)} (idx {j})")
                break
        if ecl_contract_col_pos is None:
            self.log("CONTRACT NO column not found in ECL Portfolio!", "ERROR")
            return False

        # Build lookup: contract_no -> DataFrame row
        ecl_data_rows = ecl_df_raw.iloc[ecl_header_idx + 1:].reset_index(drop=True)
        self.log(f"ECL data rows to index: {len(ecl_data_rows)}")
        ecl_lookup = {}
        contract_series = ecl_data_rows.iloc[:, ecl_contract_col_pos].astype(str).str.strip()
        for i, contract in enumerate(contract_series):
            if contract and contract != 'nan':
                ecl_lookup[contract] = ecl_data_rows.iloc[i]
        self.log(f"ECL lookup built: {len(ecl_lookup)} unique contracts", "SUCCESS")

        # ---------------------------------------------------------------
        # Phase 2: Update Provision Comparison Portfolio sheet
        # ---------------------------------------------------------------
        # Target columns: AP=42, AQ=43, AR=44, AS=45, AT=46, AU=47, AV=48, AW=49 (1-indexed)
        TARGET_START = 42   # AP
        TARGET_END   = 49   # AW
        N_TARGET     = TARGET_END - TARGET_START + 1  # 8

        app = xw.App(visible=False)
        try:
            prov_wb = app.books.open(str(provision_file))
            prov_sheet = prov_wb.sheets["Portfolio"]

            # Read row 1 cells AP-AW to get source column letter mapping
            mapping_raw = prov_sheet.range((1, TARGET_START), (1, TARGET_END)).value
            if not isinstance(mapping_raw, list):
                mapping_raw = [mapping_raw]
            self.log(f"Row 1 mapping (AP-AW): {mapping_raw}")

            source_col_indices = [col_letter_to_idx(str(v)) if v is not None else None
                                   for v in mapping_raw]
            self.log("Resolved source column indices (0-indexed in ECL):")
            for j, (raw, idx) in enumerate(zip(mapping_raw, source_col_indices)):
                target_col_letter = idx_to_col_letter(TARGET_START - 1 + j)
                ecl_col_letter = idx_to_col_letter(idx) if idx is not None else 'N/A'
                self.log(f"  Portfolio {target_col_letter} ← ECL col {ecl_col_letter} (idx {idx}) [mapping: '{raw}']")

            # Find header row in Portfolio sheet
            header_scan = prov_sheet.range((1, 1), (10, 5)).value
            prov_header_idx = None
            for idx, row in enumerate(header_scan):
                row_str = ' '.join([str(c).upper() for c in row if c is not None])
                if 'CONTRACT' in row_str:
                    prov_header_idx = idx
                    self.log(f"Portfolio header at Excel row {idx + 1}")
                    break
            if prov_header_idx is None:
                self.log("Header row not found in Portfolio sheet!", "ERROR")
                prov_wb.close()
                app.quit()
                return False

            data_start_row = prov_header_idx + 2  # 1-indexed Excel row
            prov_last_row = prov_sheet.used_range.last_cell.row
            n_data_rows = prov_last_row - data_start_row + 1
            self.log(f"Portfolio data: Excel rows {data_start_row}–{prov_last_row} ({n_data_rows} rows)")

            # Read CONTRACT NO column (col A) from Portfolio sheet
            contracts_raw = prov_sheet.range((data_start_row, 1), (prov_last_row, 1)).value
            if not isinstance(contracts_raw, list):
                contracts_raw = [contracts_raw]
            prov_contracts = [str(c).strip() if c is not None else None for c in contracts_raw]
            self.log(f"Read {len(prov_contracts)} contracts from Portfolio sheet")

            # Build output array (bulk write — much faster than cell-by-cell)
            output = [[None] * N_TARGET for _ in range(n_data_rows)]
            matched = 0
            unmatched_sample = []

            for i, contract in enumerate(prov_contracts):
                if contract and contract != 'nan' and contract in ecl_lookup:
                    ecl_row = ecl_lookup[contract]
                    for j, src_idx in enumerate(source_col_indices):
                        if src_idx is not None and src_idx < len(ecl_row):
                            val = ecl_row.iloc[src_idx]
                            output[i][j] = val if (not isinstance(val, float) or not pd.isna(val)) else None
                    matched += 1
                else:
                    if contract and contract != 'nan' and len(unmatched_sample) < 5:
                        unmatched_sample.append(contract)

            unmatched_total = n_data_rows - matched
            self.log(f"Match result: {matched} matched, {unmatched_total} unmatched")
            if unmatched_sample:
                self.log(f"Sample unmatched contracts: {unmatched_sample}", "WARNING")

            # Bulk write all 8 columns at once
            self.log(f"Writing {n_data_rows} rows to Portfolio columns AP-AW...")
            output_range = prov_sheet.range(
                (data_start_row, TARGET_START),
                (data_start_row + n_data_rows - 1, TARGET_END)
            )
            output_range.value = output
            self.log(f"Written to Portfolio AP-AW", "SUCCESS")

            prov_wb.save()
            prov_wb.close()
            app.quit()

            self.log(f"Step 06 completed in {time.time() - step_start:.2f}s", "SUCCESS")
            return True

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            self.log(f"Error in Step 06: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            return False

    def resolve_sheet_name(self, file_path, keyword):
        """Find the exact sheet name in an Excel file by keyword (case-insensitive substring)."""
        if file_path.suffix == '.xlsb':
            app = xw.App(visible=False)
            try:
                wb = app.books.open(str(file_path))
                for s in wb.sheets:
                    if keyword.lower() in s.name.lower():
                        name = s.name
                        wb.close()
                        app.quit()
                        self.log(f"Resolved sheet '{keyword}' - '{name}'")
                        return name
                wb.close()
                app.quit()
            except Exception as e:
                try:
                    app.quit()
                except:
                    pass
                self.log(f"Error resolving sheet name: {e}", "ERROR")
        else:
            try:
                xf = pd.ExcelFile(str(file_path))
                for name in xf.sheet_names:
                    if keyword.lower() in name.lower():
                        self.log(f"Resolved sheet '{keyword}' - '{name}'")
                        return name
            except Exception as e:
                self.log(f"Error resolving sheet name: {e}", "ERROR")
        self.log(f"Sheet matching '{keyword}' not found in {file_path.name}", "ERROR")
        return None

    def step_07_update_repossessed_list(self):
        """Step 07: Replace RepossessedList cols A-N with matching data from YardList in portfolio_4_ECL"""
        self.log("=" * 80)
        self.log("STEP 07: Updating RepossessedList from YardList (portfolio_4_ECL)")
        self.log("=" * 80)
        step_start = time.time()

        ecl_file = self.find_file("portfolio_4")
        provision_file = self.find_file("provision")

        if not ecl_file or not provision_file:
            self.log("Required files not found!", "ERROR")
            return False

        # ── Phase 1: Read YardList sheet from ECL file into a DataFrame ──────
        self.log(f"Reading YardList from: {ecl_file.name}")

        # Resolve the exact sheet name (handles "1.YardList", "YardList", etc.)
        yard_sheet_name = self.resolve_sheet_name(ecl_file, "yardlist")
        if yard_sheet_name is None:
            return False

        yard_df_raw = self.read_excel_file(ecl_file, yard_sheet_name)

        # Use "CONTRACT NO" + "YARD NAME" to avoid matching title rows like "YARD VEHICLE CONTRACT LIST"
        yard_header_idx = self.find_header_row(yard_df_raw, ["CONTRACT NO", "YARD NAME"])
        if yard_header_idx is None:
            yard_header_idx = self.find_header_row(yard_df_raw, ["CONTRACT NO", "VEHICLE"])
        if yard_header_idx is None:
            yard_header_idx = self.find_header_row(yard_df_raw, ["CONTRACT", "VEHICLE NO"])
        if yard_header_idx is None:
            self.log("Header row not found in YardList!", "ERROR")
            return False

        self.log(f"YardList header at index {yard_header_idx}")
        yard_df_raw.columns = yard_df_raw.iloc[yard_header_idx]
        yard_df = yard_df_raw.iloc[yard_header_idx + 1:].reset_index(drop=True)

        # Drop rows where the first column is empty
        first_yard_col = yard_df.columns[0]
        yard_df = yard_df[yard_df[first_yard_col].notna()].reset_index(drop=True)
        new_rows = len(yard_df)
        self.log(f"YardList data rows: {new_rows}")

        self.log("YardList columns:")
        for c in yard_df.columns:
            self.log(f"  '{c}'")

        # ── Phase 2: Open provision file ─────────────────────────────────────
        app = xw.App(visible=False)
        try:
            wb = app.books.open(str(provision_file))
            sheet = wb.sheets["RepossessedList"]

            N_COLS = 14  # columns A (1) to N (14)

            # Find header row in RepossessedList (scan first 15 rows, A-N only)
            header_scan = sheet.range((1, 1), (15, N_COLS)).value  # list of lists
            reposs_header_idx = None
            for idx, row in enumerate(header_scan):
                row_str = ' '.join([str(c).upper() for c in row if c is not None])
                # Use VEHICLE + YARD — both appear in the header row without typo issues
                # (avoids "CONTRTACT" typo in col A which breaks a "CONTRACT" substring check)
                if 'VEHICLE' in row_str and 'YARD' in row_str:
                    reposs_header_idx = idx
                    break
            if reposs_header_idx is None:
                # Fallback: any row with VEHICLE NO
                for idx, row in enumerate(header_scan):
                    row_str = ' '.join([str(c).upper() for c in row if c is not None])
                    if 'VEHICLE' in row_str and 'PRODUCT' in row_str:
                        reposs_header_idx = idx
                        self.log(f"Header found via fallback match at scan row {idx}", "WARNING")
                        break

            if reposs_header_idx is None:
                self.log("Header row not found in RepossessedList!", "ERROR")
                wb.close()
                app.quit()
                return False

            reposs_header_excel_row = reposs_header_idx + 1   # 1-indexed Excel row
            data_start_excel_row    = reposs_header_excel_row + 1
            reposs_cols = header_scan[reposs_header_idx][:N_COLS]  # up to 14 headers

            self.log(f"RepossessedList header at Excel row {reposs_header_excel_row}")
            self.log(f"RepossessedList A-N headers: {reposs_cols}")

            # Snapshot used range BEFORE any modifications
            last_used_row = sheet.used_range.last_cell.row
            last_used_col = sheet.used_range.last_cell.column
            old_data_rows = max(0, last_used_row - reposs_header_excel_row)
            self.log(f"Old data rows: {old_data_rows}, last used col: {last_used_col}")

            # ── Column mapping: fuzzy match RepossessedList A-N - YardList cols ──
            import re
            from difflib import SequenceMatcher

            def _clean_hdr(h):
                """Strip all non-alphanumeric chars and uppercase for comparison."""
                return re.sub(r'[^A-Z0-9]', '', str(h).upper())

            def _headers_match(a, b):
                ca, cb = _clean_hdr(a), _clean_hdr(b)
                if not ca or not cb:
                    return False
                if ca == cb:
                    return True
                # Substring: handles "VALUER" ↔ "VALUERNAME", "STATUS" ↔ "CURRENTSTATUS"
                if ca in cb or cb in ca:
                    return True
                # Sequence similarity: handles typos like "CONTRTACTNO" ↔ "CONTRACTNO"
                return SequenceMatcher(None, ca, cb).ratio() >= 0.85

            # Build list of (cleaned_name, original_col) for YardList
            yard_col_clean = [(_clean_hdr(c), c) for c in yard_df.columns]

            col_mapping = {}  # {0-based A-N index - original YardList column name}
            for i, reposs_col in enumerate(reposs_cols):
                if reposs_col is None:
                    continue
                # Find best matching YardList column
                best_col = None
                best_ratio = 0.0
                cleaned_reposs = _clean_hdr(reposs_col)
                for yard_cleaned, yard_orig in yard_col_clean:
                    if _headers_match(reposs_col, yard_orig):
                        # Prefer higher similarity when multiple match
                        ratio = SequenceMatcher(None, cleaned_reposs, yard_cleaned).ratio()
                        if ratio > best_ratio:
                            best_ratio = ratio
                            best_col = yard_orig
                if best_col is not None:
                    col_mapping[i] = best_col

            self.log(f"Column mapping — {len(col_mapping)} columns matched:")
            for col_idx, yard_col in col_mapping.items():
                self.log(f"  Col {chr(65 + col_idx)} ← YardList '{yard_col}'")

            # ── Build output using DataFrames (vectorised, no row-by-row loops) ─
            output_df = pd.DataFrame(None, index=range(new_rows), columns=range(N_COLS))
            for col_idx, yard_col in col_mapping.items():
                output_df[col_idx] = yard_df[yard_col].reset_index(drop=True)

            # Convert NaN - None so xlwings writes blank cells
            def _clean(v):
                return None if (isinstance(v, float) and pd.isna(v)) else v

            output_list = [[_clean(v) for v in row] for row in output_df.values.tolist()]

            # ── Clear old A-N data (below header) ────────────────────────────
            if old_data_rows > 0:
                sheet.range(
                    (data_start_excel_row, 1),
                    (data_start_excel_row + old_data_rows - 1, N_COLS)
                ).clear_contents()
                self.log(f"Cleared {old_data_rows} old rows in columns A-N")

            # ── Bulk-write new data into A-N ──────────────────────────────────
            sheet.range((data_start_excel_row, 1)).value = output_list
            self.log(f"Wrote {new_rows} rows to RepossessedList columns A-N", "SUCCESS")

            # ── Handle formula columns O onwards (col 15+) ───────────────────
            if last_used_col > N_COLS:
                if new_rows > old_data_rows and old_data_rows > 0:
                    # Extra rows in new data — extend formulas downward
                    extra_rows = new_rows - old_data_rows
                    self.log(f"Extending formulas for {extra_rows} additional rows "
                             f"(cols O–{last_used_col})...")
                    # Use first data row as the formula template
                    src_range = sheet.range(
                        (data_start_excel_row, N_COLS + 1),
                        (data_start_excel_row, last_used_col)
                    )
                    first_new_row = data_start_excel_row + old_data_rows
                    last_new_row  = data_start_excel_row + new_rows - 1
                    tgt_range = sheet.range(
                        (first_new_row, N_COLS + 1),
                        (last_new_row, last_used_col)
                    )
                    src_range.copy(tgt_range)
                    self.log(f"Formulas extended to Excel rows {first_new_row}–{last_new_row}",
                             "SUCCESS")

                elif new_rows < old_data_rows:
                    # Fewer rows — clear excess formula rows (O+)
                    excess_start = data_start_excel_row + new_rows
                    excess_end   = data_start_excel_row + old_data_rows - 1
                    self.log(f"Clearing {old_data_rows - new_rows} excess formula rows "
                             f"(rows {excess_start}–{excess_end})...")
                    sheet.range(
                        (excess_start, N_COLS + 1),
                        (excess_end, last_used_col)
                    ).clear_contents()
                    self.log(f"Excess formula rows cleared", "SUCCESS")

            wb.save()
            wb.close()
            app.quit()

            self.log(f"Step 07 completed in {time.time() - step_start:.2f}s", "SUCCESS")
            return True

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            self.log(f"Error in Step 07: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            return False

    def run_all_steps(self, reporting_date=None):
        """Execute all automation steps"""
        self.log("=" * 80)
        self.log("EXCEL AUTOMATION STARTED")
        self.log("=" * 80)
        self.log(f"Base path: {self.base_path}")
        self.log(f"Reporting date: {reporting_date.strftime('%d/%m/%Y') if reporting_date else 'N/A'}")

        try:
            # Execute all steps
            success_01 = self.step_01_copy_summary_to_portfolio()
            success_02 = self.step_02_sync_property_mortgages()
            success_03 = self.step_03_update_property_mortgage_columns(reporting_date=reporting_date)
            success_04 = self.step_04_update_comparison_tables()
            success_05 = self.step_05_update_repossessed_reporting_date(reporting_date)
            success_06 = self.step_06_update_portfolio_ecl_columns()
            success_07 = self.step_07_update_repossessed_list()

            all_steps = [success_01, success_02, success_03, success_04, success_05, success_06, success_07]
            total_time = time.time() - self.start_time
            self.log("=" * 80)

            if all(all_steps):
                self.log(f"ALL STEPS COMPLETED SUCCESSFULLY in {total_time:.2f}s", "SUCCESS")
            else:
                self.log(f"COMPLETED WITH WARNINGS in {total_time:.2f}s", "WARNING")
                self.log(f"  Step 01: {'SUCCESS' if success_01 else 'FAILED'}")
                self.log(f"  Step 02: {'SUCCESS' if success_02 else 'FAILED'}")
                self.log(f"  Step 03: {'SUCCESS' if success_03 else 'FAILED'}")
                self.log(f"  Step 04: {'SUCCESS' if success_04 else 'FAILED'}")
                self.log(f"  Step 05: {'SUCCESS' if success_05 else 'FAILED'}")
                self.log(f"  Step 06: {'SUCCESS' if success_06 else 'FAILED'}")
                self.log(f"  Step 07: {'SUCCESS' if success_07 else 'FAILED'}")

            self.log("=" * 80)

            # ── Rename provision file to reflect the reporting date ───────────
            import re
            provision_file_ref = self.find_file("provision")
            if provision_file_ref:
                date_for_name = reporting_date if reporting_date else datetime.now()
                new_date_str  = date_for_name.strftime("%b %Y")   # e.g. "Aug 2025"
                old_name      = provision_file_ref.name
                # Replace the existing "Mon YYYY" pattern anywhere in the filename
                new_name = re.sub(r'[A-Za-z]{3} \d{4}', new_date_str, old_name)
                if new_name != old_name:
                    new_path = provision_file_ref.parent / new_name
                    provision_file_ref.rename(new_path)
                    self.log(f"Renamed: '{old_name}' - '{new_name}'", "SUCCESS")
                else:
                    self.log(f"Provision filename already up to date: '{old_name}'")

        except Exception as e:
            self.log(f"FATAL ERROR: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")

# Main execution
if __name__ == "__main__":
    # Parse optional date argument: python working.py 25/08/2025
    if len(sys.argv) > 1:
        try:
            reporting_date = datetime.strptime(sys.argv[1], "%d/%m/%Y")
            print(f"Using provided reporting date: {reporting_date.strftime('%d/%m/%Y')}")
        except ValueError:
            print(f"Invalid date format: '{sys.argv[1]}'. Expected DD/MM/YYYY")
            sys.exit(1)
    else:
        reporting_date = datetime.now()
        print(f"No date provided, using current date: {reporting_date.strftime('%d/%m/%Y')}")

    automation = ExcelAutomation(FILE_PATH)
    automation.run_all_steps(reporting_date=reporting_date)