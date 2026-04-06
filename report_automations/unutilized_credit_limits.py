import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import subprocess
import sys
import re
import traceback
from pathlib import Path
from datetime import datetime
import time
import gc

# Configuration
BASE_PATH = r"C:\Users\ASUS\Documents\WORK\Unutilized\Files"

# Keywords used to identify files (matched case-insensitively against filenames)
FILE_KEYWORDS = {
    "summary": "Summary",
    "unutilized": "Unutilized Credit Limits",
    "net_portfolio": "Net Portfolio",
}

class ExcelProcessor:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.summary_path = self._find_file(FILE_KEYWORDS["summary"])
        self.unutilized_path = self._find_file(FILE_KEYWORDS["unutilized"])
        self.net_portfolio_path = self._find_file(FILE_KEYWORDS["net_portfolio"])

    def _build_unutilized_save_path(self, date_str=None):
        """Build save path for Unutilized file with the given or current date.

        Args:
            date_str: Optional date string in DD/MM/YYYY format (e.g., '08/08/2025').
                      If provided, formats as '08-Aug-2025'.
                      If not provided, uses current date as 'Aug 2025'.
        """
        name = self.unutilized_path.name
        # Match both possible formats in existing filename: "Aug 2025" or "08-Aug-2025"
        pattern = r'\d{2}-[A-Z][a-z]{2}-\d{4}|[A-Z][a-z]{2}\s\d{4}'
        if date_str:
            date_obj = datetime.strptime(date_str, "%d/%m/%Y")
            new_date = date_obj.strftime("%d-%b-%Y")  # e.g. 08-Aug-2025
        else:
            new_date = datetime.now().strftime("%b %Y")  # e.g. Aug 2025
        new_name = re.sub(pattern, new_date, name)
        return self.unutilized_path.parent / new_name

    def _find_file(self, keyword):
        """Find a file in base_path whose name contains the keyword (case-insensitive)."""
        keyword_lower = keyword.lower()
        matches = [
            f for f in self.base_path.iterdir()
            if f.is_file() and keyword_lower in f.name.lower()
        ]
        if len(matches) == 1:
            return matches[0]
        if len(matches) == 0:
            raise FileNotFoundError(f"No file found containing '{keyword}' in {self.base_path}")
        # Multiple matches – pick the most recently modified one
        matches.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        print(f"  ⚠ Multiple files match '{keyword}', using most recent: {matches[0].name}")
        return matches[0]
        
    def log(self, step, message):
        """Log processing steps"""
        print(f"\n[STEP {step}] {message}")

    def debug(self, data_info):
        """Print debug information"""
        print(f"  → {data_info}")

    def ensure_file_closed(self, wait_time=0.5):
        """Ensure Excel file is fully closed and released by forcing garbage collection"""
        gc.collect()
        time.sleep(wait_time)
        
    def read_xlsb(self, file_path, sheet_name, skiprows=2):
        """Read XLSB file with proper header row handling"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', skiprows=skiprows)
            df.columns = [str(col).strip().upper() for col in df.columns]
            self.debug(f"Read {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
            return df
        except Exception as e:
            self.log("ERROR", f"Failed to read {sheet_name}: {str(e)}")
            raise
    
    def read_xlsx(self, file_path, sheet_name, skiprows=2):
        """Read XLSX file with proper header row handling"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=skiprows)
            df.columns = [str(col).strip().upper() for col in df.columns]
            self.debug(f"Read {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
            return df
        except Exception as e:
            self.log("ERROR", f"Failed to read {sheet_name}: {str(e)}")
            raise
    
    def map_columns_by_header(self, source_df, dest_columns):
        """Map source dataframe columns to destination columns by matching header names"""
        mapped_data = pd.DataFrame()
        source_cols = set(source_df.columns)
        matched_cols = []
        unmatched_cols = []
        
        for dest_col in dest_columns:
            if dest_col in source_cols:
                mapped_data[dest_col] = source_df[dest_col]
                matched_cols.append(dest_col)
            else:
                mapped_data[dest_col] = None
                unmatched_cols.append(dest_col)
        
        self.debug(f"Matched {len(matched_cols)}/{len(dest_columns)} columns")
        if unmatched_cols and len(unmatched_cols) <= 10:
            self.debug(f"Unmatched: {unmatched_cols}")
        
        return mapped_data
    
    def write_data_preserve_structure(self, wb, sheet_name, data_df, header_row=3, start_col=1, preserve_formulas=False):
        """Write data to Excel while preserving existing structure and optionally formulas"""
        ws = wb[sheet_name]
        
        if preserve_formulas:
            # Preserve formulas: only update cells with data, skip formula cells
            data_start_row = header_row + 1
            
            for row_idx, row_data in enumerate(data_df.itertuples(index=False), start=data_start_row):
                for col_idx, value in enumerate(row_data, start=start_col):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    
                    # Only update if cell doesn't have a formula
                    if not cell.value or (isinstance(cell.value, str) and not cell.value.startswith('=')):
                        if pd.notna(value):
                            if isinstance(value, (int, float)):
                                cell.value = float(value) if not isinstance(value, int) else int(value)
                            else:
                                cell.value = value
                        else:
                            cell.value = None
        else:
            # Clear existing data rows (keep headers) and write fresh data
            max_row = ws.max_row
            if max_row > header_row:
                ws.delete_rows(header_row + 1, max_row - header_row)
            
            # Write data starting from the row after header
            data_start_row = header_row + 1
            
            for row_idx, row_data in enumerate(data_df.itertuples(index=False), start=data_start_row):
                for col_idx, value in enumerate(row_data, start=start_col):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if pd.notna(value):
                        if isinstance(value, (int, float)):
                            cell.value = float(value) if not isinstance(value, int) else int(value)
                        else:
                            cell.value = value
                    else:
                        cell.value = None
        
        self.debug(f"Written {len(data_df)} rows to {sheet_name}")
    
    def step_01_filter_dl(self):
        """STEP 01: Filter DL from SUMMARY and replace in DF sheet"""
        self.log("01", "Processing DL data from SUMMARY sheet")
        
        try:
            summary_df = self.read_xlsb(self.summary_path, 'SUMMARY', skiprows=2)
            
            if 'PRODUCT' not in summary_df.columns:
                raise ValueError(f"PRODUCT column not found in SUMMARY")
            
            dl_data = summary_df[summary_df['PRODUCT'] == 'DL'].copy()
            self.debug(f"Filtered {len(dl_data)} DL records")
            
            if len(dl_data) == 0:
                self.log("WARNING", "No DL records found")
                return
            
            df_sheet = self.read_xlsx(self.unutilized_path, 'DF', skiprows=2)
            mapped_data = self.map_columns_by_header(dl_data, df_sheet.columns.tolist())
            
            wb = openpyxl.load_workbook(self.unutilized_path)
            self.write_data_preserve_structure(wb, 'DF', mapped_data, header_row=3, start_col=1)
            wb.save(self.unutilized_path)
            wb.close()
            del wb

            self.debug(f"✓ DF sheet updated with {len(mapped_data)} records")
            
        except Exception as e:
            self.log("ERROR", f"Step 01 failed: {str(e)}")
            traceback.print_exc()
    
    def step_02_filter_pg(self):
        """STEP 02: Filter PG from SUMMARY and replace in PG sheet"""
        self.log("02", "Processing PG data from SUMMARY sheet")
        
        try:
            summary_df = self.read_xlsb(self.summary_path, 'SUMMARY', skiprows=2)
            pg_data = summary_df[summary_df['PRODUCT'] == 'PG'].copy()
            self.debug(f"Filtered {len(pg_data)} PG records")
            
            if len(pg_data) == 0:
                self.log("WARNING", "No PG records found")
                return
            
            pg_sheet = self.read_xlsx(self.unutilized_path, 'PG', skiprows=2)
            mapped_data = self.map_columns_by_header(pg_data, pg_sheet.columns.tolist())
            
            wb = openpyxl.load_workbook(self.unutilized_path)
            self.write_data_preserve_structure(wb, 'PG', mapped_data, header_row=3, start_col=1)
            wb.save(self.unutilized_path)
            wb.close()
            del wb

            self.debug(f"✓ PG sheet updated with {len(mapped_data)} records")
            
        except Exception as e:
            self.log("ERROR", f"Step 02 failed: {str(e)}")
            traceback.print_exc()
    
    def step_03_create_pivot_tables(self):
        """STEP 03: Placeholder - Pivot tables will be refreshed in next step using xlwings"""
        self.log("03", "Pivot table refresh will be handled by xlwings")
        self.debug("✓ Skipping manual pivot computation - using Excel native pivot tables")
    
    def step_04_sunmatch_update(self):
        """STEP 04: Update SunMatch sheet with filtered PG data"""
        self.log("04", "Updating SunMatch sheet")
        
        try:
            sunmatch_df = self.read_xlsx(self.unutilized_path, 'SunMatch', skiprows=2)
            
            if len(sunmatch_df) == 0 or 'CLIENT NO' not in sunmatch_df.columns:
                self.log("WARNING", "No data or CLIENT NO column not found in SunMatch")
                return
            
            sunmatch_client = sunmatch_df.iloc[0]['CLIENT NO']
            self.debug(f"SunMatch CLIENT NO: {sunmatch_client}")
            
            dest_columns = sunmatch_df.columns.tolist()
            summary_df = self.read_xlsb(self.summary_path, 'SUMMARY', skiprows=2)
            
            filtered_data = summary_df[
                (summary_df['PRODUCT'] == 'PG') & 
                (summary_df['CLIENT NO'] == sunmatch_client)
            ].copy()
            
            self.debug(f"Filtered {len(filtered_data)} PG records for SunMatch client")
            
            if len(filtered_data) == 0:
                self.log("WARNING", "No matching PG data found for SunMatch client")
                return
            
            mapped_data = self.map_columns_by_header(filtered_data, dest_columns)
            
            wb = openpyxl.load_workbook(self.unutilized_path)
            self.write_data_preserve_structure(wb, 'SunMatch', mapped_data, header_row=3, start_col=1, preserve_formulas=False)
            wb.save(self.unutilized_path)
            wb.close()
            del wb

            self.debug(f"✓ SunMatch sheet updated with {len(mapped_data)} records")
            
        except Exception as e:
            self.log("ERROR", f"Step 04 failed: {str(e)}")
            traceback.print_exc()
    
    def step_05_copy_overpayment(self):
        """STEP 05: Lookup CONTRACT NO from SunMatch in Net Portfolio and copy OVERPAYMENT values"""
        self.log("05", "Looking up OVERPAYMENT values based on CONTRACT NO")
        
        try:
            # Read Net Portfolio with auto-detection of header row
            net_portfolio_df = None
            for skip in range(5):
                try:
                    temp_df = pd.read_excel(self.net_portfolio_path, skiprows=skip)
                    temp_df.columns = [str(col).strip().upper() for col in temp_df.columns]
                    
                    # Check if this looks like a header row
                    has_contract = any('CONTRACT' in str(col).upper() for col in temp_df.columns)
                    if has_contract:
                        net_portfolio_df = temp_df
                        self.debug(f"Found Net Portfolio headers (skiprows={skip})")
                        break
                except:
                    continue
            
            if net_portfolio_df is None:
                self.log("ERROR", "Could not read Net Portfolio file")
                return
            
            # Find CONTRACT NO column in Net Portfolio
            contract_col_net = None
            for col in net_portfolio_df.columns:
                col_str = str(col).upper()
                if 'CONTRACT' in col_str and 'NO' in col_str:
                    contract_col_net = col
                    break
            
            # Fallback to column E (5th column, index 4)
            if not contract_col_net and len(net_portfolio_df.columns) > 4:
                contract_col_net = net_portfolio_df.columns[4]
                self.debug(f"Using positional column 5 (E) for contract")
            
            # Find OVERPAYMENT column in Net Portfolio
            overpayment_col_net = None
            for col in net_portfolio_df.columns:
                if 'OVERPAYMENT' in str(col).upper() or 'OVER PAYMENT' in str(col).upper() or 'OVERPAY' in str(col).upper():
                    overpayment_col_net = col
                    break
            
            # Fallback to column AC (29th column, index 28)
            if not overpayment_col_net and len(net_portfolio_df.columns) > 28:
                overpayment_col_net = net_portfolio_df.columns[28]
                self.debug(f"Using positional column 29 (AC) for overpayment")
            
            if not contract_col_net or not overpayment_col_net:
                self.log("ERROR", "Could not find required columns in Net Portfolio")
                return
            
            self.debug(f"Net Portfolio CONTRACT column: '{contract_col_net}'")
            self.debug(f"Net Portfolio OVERPAYMENT column: '{overpayment_col_net}'")
            
            # Debug: Show raw sample values from Net Portfolio before cleaning
            raw_samples = net_portfolio_df[contract_col_net].dropna().head(5).tolist()
            self.debug(f"Net Portfolio raw contract samples: {raw_samples}")
            raw_overpay_samples = net_portfolio_df[overpayment_col_net].dropna().head(5).tolist()
            self.debug(f"Net Portfolio raw overpayment samples: {raw_overpay_samples}")

            # Clean contract numbers and create lookup dictionary
            net_portfolio_df[contract_col_net] = net_portfolio_df[contract_col_net].astype(str).str.strip()
            net_portfolio_df[overpayment_col_net] = pd.to_numeric(net_portfolio_df[overpayment_col_net], errors='coerce').fillna(0)

            # Debug: Show non-zero overpayment entries in lookup
            non_zero_overpay = net_portfolio_df[net_portfolio_df[overpayment_col_net] != 0]
            self.debug(f"Net Portfolio entries with non-zero overpayment: {len(non_zero_overpay)}")
            if len(non_zero_overpay) > 0:
                self.debug(f"  Sample non-zero entries: {list(zip(non_zero_overpay[contract_col_net].head(3).tolist(), non_zero_overpay[overpayment_col_net].head(3).tolist()))}")

            # Create lookup dictionary: {contract_no: overpayment_value}
            lookup_dict = dict(zip(net_portfolio_df[contract_col_net], net_portfolio_df[overpayment_col_net]))
            self.debug(f"Created lookup dictionary with {len(lookup_dict)} contract entries")

            # Debug: Show sample keys from lookup dict (cleaned)
            sample_keys = list(lookup_dict.keys())[:5]
            self.debug(f"Lookup dict sample keys (cleaned): {sample_keys}")
            
            # Load workbook with openpyxl to preserve formulas
            wb = openpyxl.load_workbook(self.unutilized_path)
            ws = wb['SunMatch']
            
            # Find header row and columns in SunMatch
            header_row = 3
            contract_col_idx = None
            overpayment_col_idx = None
            sun_match_amount_col_idx = None
            
            for col_idx in range(1, ws.max_column + 1):
                cell_value = ws.cell(header_row, col_idx).value
                if cell_value:
                    cell_str = str(cell_value).upper()
                    if 'CONTRACT' in cell_str and 'NO' in cell_str:
                        contract_col_idx = col_idx
                    if 'OVERPAYMENT' in cell_str and 'SUN MATCH' not in cell_str:
                        overpayment_col_idx = col_idx
                    if 'SUN MATCH AMOUNT' in cell_str or 'SUNMATCH AMOUNT' in cell_str:
                        sun_match_amount_col_idx = col_idx
            
            if not contract_col_idx:
                self.log("ERROR", "Could not find CONTRACT NO column in SunMatch sheet")
                wb.close()
                return
            
            if not overpayment_col_idx:
                self.log("ERROR", "Could not find OVERPAYMENT column in SunMatch sheet")
                wb.close()
                return
            
            contract_col_letter = get_column_letter(contract_col_idx)
            overpayment_col_letter = get_column_letter(overpayment_col_idx)
            
            self.debug(f"SunMatch CONTRACT NO column: {contract_col_letter} (column {contract_col_idx})")
            self.debug(f"SunMatch OVERPAYMENT column: {overpayment_col_letter} (column {overpayment_col_idx})")
            
            # Lookup and copy OVERPAYMENT values based on CONTRACT NO
            data_start_row = header_row + 1
            rows_processed = 0
            matches_found = 0
            unmatched_samples = []

            # Debug: Show first few SunMatch contract values before lookup
            sunmatch_sample_vals = []
            for row_idx in range(data_start_row, min(data_start_row + 5, ws.max_row + 1)):
                raw_val = ws.cell(row=row_idx, column=contract_col_idx).value
                sunmatch_sample_vals.append(repr(raw_val))
            self.debug(f"SunMatch raw contract samples (first 5): {sunmatch_sample_vals}")

            for row_idx in range(data_start_row, ws.max_row + 1):
                contract_cell = ws.cell(row=row_idx, column=contract_col_idx)
                contract_value = str(contract_cell.value).strip() if contract_cell.value else None

                if contract_value and contract_value not in ['None', '']:
                    rows_processed += 1

                    # Lookup overpayment value for this contract number
                    if contract_value in lookup_dict:
                        overpayment_value = lookup_dict[contract_value]
                        overpayment_cell = ws.cell(row=row_idx, column=overpayment_col_idx)
                        overpayment_cell.value = float(overpayment_value)
                        matches_found += 1
                    else:
                        # Contract not found in Net Portfolio, set to 0
                        overpayment_cell = ws.cell(row=row_idx, column=overpayment_col_idx)
                        overpayment_cell.value = 0
                        if len(unmatched_samples) < 5:
                            unmatched_samples.append(contract_value)

            self.debug(f"✓ Processed {rows_processed} contracts")
            self.debug(f"✓ Found {matches_found} matches in Net Portfolio")
            if unmatched_samples:
                self.debug(f"⚠ Unmatched contract samples (not found in Net Portfolio): {unmatched_samples}")
            
            # Add SUN MATCH AMOUNT formula (=AB#-AP#) to column AQ
            if sun_match_amount_col_idx:
                sun_match_col_letter = get_column_letter(sun_match_amount_col_idx)
                self.debug(f"Adding SUN MATCH AMOUNT formula to column {sun_match_col_letter}")
                
                formula_count = 0
                for row_idx in range(data_start_row, ws.max_row + 1):
                    # Check if there's data in this row
                    contract_cell = ws.cell(row=row_idx, column=contract_col_idx)
                    if contract_cell.value and str(contract_cell.value).strip() not in ['None', '']:
                        # Add formula: =AB#-AP#
                        formula = f"=AB{row_idx}-AP{row_idx}"
                        ws.cell(row=row_idx, column=sun_match_amount_col_idx, value=formula)
                        formula_count += 1
                
                self.debug(f"✓ Added formula to {formula_count} rows in SUN MATCH AMOUNT column")
            else:
                self.debug("⚠ WARNING: Could not find SUN MATCH AMOUNT column")
            
            # Save workbook
            wb.save(self.unutilized_path)
            wb.close()
            del wb

            self.debug("✓ Data written successfully with formulas preserved")
            
        except Exception as e:
            self.log("ERROR", f"Step 05 failed: {str(e)}")
            traceback.print_exc()


    def update_pivot_source_and_refresh(self,
                                       sheet_name: str,
                                       pivot_table_name: str,
                                       data_sheet_name: str,
                                       start_cell: str = 'A1',
                                       end_column: str = None,
                                       include_headers: bool = True) -> None:
        """
        Create/update pivot table by aggregating data from source sheet
        Uses pandas and openpyxl only (no COM automation to avoid errors)

        Args:
            sheet_name: Name of the worksheet to write pivot table to (e.g., 'DF Summary')
            pivot_table_name: Name of the pivot table (for logging only)
            data_sheet_name: Name of the worksheet containing source data (e.g., 'DF')
            start_cell: Starting cell of the data range (default: 'A1')
            end_column: Last column letter (not used with openpyxl method)
            include_headers: Whether to include header row (default: True)
        """
        try:
            self.log("PIVOT", f"Updating {sheet_name} from {data_sheet_name}")

            # Extract start row from start_cell
            start_row = int(''.join(filter(str.isdigit, start_cell))) if start_cell else 1

            # Calculate skiprows for pandas
            skiprows = start_row - 1 if start_row > 1 else 0

            # Read entire full range from data_sheet_name (all rows until end)
            source_data = self.read_xlsx(self.unutilized_path, data_sheet_name, skiprows=skiprows)

            # Get the full data range dimensions
            total_rows = len(source_data)
            total_cols = len(source_data.columns)
            self.debug(f"Captured full data range from {data_sheet_name}: {total_rows} rows × {total_cols} columns")

            # Find CLIENT NO column
            if 'CLIENT NO' not in source_data.columns:
                self.log("WARNING", f"CLIENT NO column not found in {data_sheet_name}")
                return

            # Find GROSS EXP + DP column
            gross_exp_col = None
            for col in source_data.columns:
                if 'GROSS' in col and ('EXP' in col or 'EXPOSURE' in col) and 'DP' in col:
                    gross_exp_col = col
                    break

            if not gross_exp_col:
                self.log("WARNING", f"GROSS EXP + DP column not found in {data_sheet_name}")
                return

            self.debug(f"Aggregating all {total_rows} rows: GROUP BY '{gross_exp_col.strip()}' SUM BY 'CLIENT NO'")

            # Create pivot aggregation from entire dataset
            pivot_data = source_data.groupby('CLIENT NO')[gross_exp_col].sum().reset_index()
            pivot_data.columns = ['Row Labels', 'Sum of GROSS EXP + DP']

            # Calculate grand total
            grand_total = pivot_data['Sum of GROSS EXP + DP'].sum()

            # Add grand total row
            pivot_data = pd.concat([
                pivot_data,
                pd.DataFrame({
                    'Row Labels': ['Grand Total'],
                    'Sum of GROSS EXP + DP': [grand_total]
                })
            ], ignore_index=True)

            self.debug(f"Pivot computed: {len(pivot_data)-1} clients + grand total")

            # Write to destination sheet using openpyxl
            wb = openpyxl.load_workbook(self.unutilized_path)
            ws = wb[sheet_name]

            # Remove any existing pivot table objects from the sheet
            if hasattr(ws, '_pivots') and ws._pivots:
                self.debug(f"Removing {len(ws._pivots)} existing pivot table(s) from {sheet_name}")
                ws._pivots = []

            # Clear existing data from row after headers onwards
            header_row = start_row
            max_row = ws.max_row
            if max_row > header_row:
                ws.delete_rows(header_row + 1, max_row - header_row)

            # Write headers with bold formatting
            ws.cell(row=header_row, column=1, value='Row Labels').font = Font(bold=True)
            ws.cell(row=header_row, column=2, value='Sum of GROSS EXP + DP').font = Font(bold=True)

            # Write static aggregated data to columns A-B
            data_start_row = header_row + 1
            grand_total_row = None
            for row_idx, (row_label, sum_value) in enumerate(pivot_data.itertuples(index=False), start=data_start_row):
                ws.cell(row=row_idx, column=1, value=row_label)

                # Format sum value as currency (with thousand separators)
                sum_cell = ws.cell(row=row_idx, column=2, value=float(sum_value) if pd.notna(sum_value) else 0)
                sum_cell.number_format = '#,##0.00'

                # Bold the Grand Total row
                if row_label == 'Grand Total':
                    ws.cell(row=row_idx, column=1).font = Font(bold=True)
                    sum_cell.font = Font(bold=True)
                    grand_total_row = row_idx

            self.debug(f"✓ Static aggregated data written to columns A-B: {len(pivot_data)} rows")

            # Now copy this data to the table starting in column D
            self.debug("Copying data to table starting in column D...")

            # Find headers in the table starting at column D (same header row)
            d_table_headers = {}
            col_idx = 4  # Column D is index 4
            while col_idx <= ws.max_column and col_idx < 20:  # Limit search to reasonable range
                header_value = ws.cell(row=header_row, column=col_idx).value
                if header_value and str(header_value).strip():
                    d_table_headers[str(header_value).strip().upper()] = col_idx
                else:
                    if d_table_headers:  # Found headers, then empty - table ended
                        break
                col_idx += 1

            if not d_table_headers:
                self.debug("⚠ No table found starting at column D")
            else:
                self.debug(f"Found table starting at column D with {len(d_table_headers)} headers")

                # Map columns from A-B table to D-table by matching headers
                mapping = {}

                # Try to find Row Labels / CLIENT NO column
                for header, col_idx in d_table_headers.items():
                    if 'ROW' in header and 'LABEL' in header:
                        mapping['Row Labels'] = col_idx
                    elif 'CLIENT' in header and ('NO' in header or 'NUMBER' in header):
                        mapping['Row Labels'] = col_idx

                # Try to find Sum column
                for header, col_idx in d_table_headers.items():
                    if 'SUM' in header and ('GROSS' in header or 'EXP' in header):
                        mapping['Sum of GROSS EXP + DP'] = col_idx
                    elif 'GROSS' in header and 'EXP' in header and 'DP' in header:
                        mapping['Sum of GROSS EXP + DP'] = col_idx

                if mapping:
                    self.debug(f"Mapped columns: {len(mapping)} matches found")

                    # Bold headers in D-table
                    for col_idx in d_table_headers.values():
                        ws.cell(row=header_row, column=col_idx).font = Font(bold=True)

                    # Clear existing data in D-table (including column F)
                    max_data_row = data_start_row + len(pivot_data) + 10  # Clear extra rows
                    for row_idx in range(data_start_row, max_data_row):
                        for col_idx in d_table_headers.values():
                            ws.cell(row=row_idx, column=col_idx).value = None
                        # Also clear column F
                        ws.cell(row=row_idx, column=6).value = None

                    # Write pivot data to mapped D-table columns
                    for row_idx, (row_label, sum_value) in enumerate(pivot_data.itertuples(index=False), start=data_start_row):
                        if 'Row Labels' in mapping:
                            label_cell = ws.cell(row=row_idx, column=mapping['Row Labels'], value=row_label)
                            # Bold Grand Total
                            if row_label == 'Grand Total':
                                label_cell.font = Font(bold=True)

                        if 'Sum of GROSS EXP + DP' in mapping:
                            sum_cell = ws.cell(row=row_idx, column=mapping['Sum of GROSS EXP + DP'],
                                              value=float(sum_value) if pd.notna(sum_value) else 0)
                            # Format as currency
                            sum_cell.number_format = '#,##0.00'
                            # Bold Grand Total
                            if row_label == 'Grand Total':
                                sum_cell.font = Font(bold=True)

                        # Apply VLOOKUP formula to column F for all data rows (NOT Grand Total)
                        if row_label != 'Grand Total':
                            vlookup_formula = f"=VLOOKUP(D{row_idx},'Working'!$A$4:$B$14,2,FALSE)"
                            ws.cell(row=row_idx, column=6, value=vlookup_formula)

                    self.debug(f"✓ Data copied to table in column D: {len(pivot_data)} rows")
                    self.debug(f"✓ VLOOKUP formulas applied to column F for all data rows")
                else:
                    self.debug("⚠ Could not find matching column headers in table at column D")

            wb.save(self.unutilized_path)
            wb.close()
            del wb

            self.debug(f"✓ {sheet_name} fully updated: {len(pivot_data)} rows (from {total_rows} source rows)")

        except Exception as e:
            self.log("ERROR", f"Failed to update {sheet_name}: {str(e)}")
            traceback.print_exc()

    def run_all_steps(self, date_str=None):
        """Execute all processing steps"""
        print("="*60)
        print("EXCEL AUTOMATION - CREDIT LIMITS PROCESSING")
        print("="*60)
        
        try:
            # Verify file paths
            if not self.summary_path.exists():
                raise FileNotFoundError(f"Summary file not found: {self.summary_path}")
            if not self.unutilized_path.exists():
                raise FileNotFoundError(f"Unutilized file not found: {self.unutilized_path}")
            if not self.net_portfolio_path.exists():
                raise FileNotFoundError(f"Net Portfolio file not found: {self.net_portfolio_path}")
            
            self.log("INIT", "All file paths verified ✓")

            # Rename Unutilized file with given or current date
            save_path = self._build_unutilized_save_path(date_str)
            if save_path != self.unutilized_path:
                self.unutilized_path.rename(save_path)
                self.debug(f"Renamed to: {save_path.name}")
                self.unutilized_path = save_path

            # Execute steps
            self.step_01_filter_dl()
            self.step_02_filter_pg()
            self.step_03_create_pivot_tables()
            self.update_pivot_source_and_refresh(
                sheet_name='DF Summary',
                pivot_table_name='PivotTable1',
                data_sheet_name='DF',
                start_cell='A3'
            )
            self.step_04_sunmatch_update()
            self.step_05_copy_overpayment()

            print("\n" + "="*60)
            print("✓ PROCESSING COMPLETED SUCCESSFULLY")
            print("="*60)
            
        except Exception as e:
            print("\n" + "="*60)
            print(f"✗ FATAL ERROR: {str(e)}")
            print("="*60)
            traceback.print_exc()
            sys.exit(1)

if __name__ == "__main__":
    date_str = None
    if len(sys.argv) > 1:
        date_str = sys.argv[1]
        try:
            datetime.strptime(date_str, "%d/%m/%Y")
        except ValueError:
            print(f"Invalid date format: '{date_str}'. Expected DD/MM/YYYY (e.g., 08/08/2025)")
            sys.exit(1)
    processor = ExcelProcessor(BASE_PATH)
    processor.run_all_steps(date_str)