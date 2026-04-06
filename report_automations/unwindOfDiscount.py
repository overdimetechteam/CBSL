import os
import shutil
from pathlib import Path
from datetime import datetime
import re
import csv
from typing import List, Tuple, Dict
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

try:
    from pyxlsb import open_workbook as open_xlsb
    XLSB_SUPPORT = True
except ImportError:
    XLSB_SUPPORT = False
    print("Warning: pyxlsb not installed. .xlsb files will not be supported.")
    print("Install with: pip install pyxlsb")


class UnwindOfDiscount:
    def __init__(self, quarter: int = 1, financial_year: int = 2025):
        # --- Path setup ---
        self.SCRIPT_DIR = Path(__file__).resolve().parent
        self.PROJECT_ROOT = self.SCRIPT_DIR.parent
        self.WORKING_DIR = self.PROJECT_ROOT / "working" / "quarterly"
        self.OUTPUTS_DIR = self.PROJECT_ROOT / "outputs" / "quarterly"
        self.MATURITY_REPORTS_DIR = self.WORKING_DIR / "Maturity Reports"
        self.PORTFOLIOS_DIR = self.WORKING_DIR / "Portfolios"

        # --- Quarter and Financial Year Configuration ---
        self.QUARTER = quarter
        self.FY_YEAR = financial_year
        
        # Define quarter months mapping
        quarter_months_map = {
            1: ["Apr", "May", "Jun"],
            2: ["Jul", "Aug", "Sep"],
            3: ["Oct", "Nov", "Dec"],
            4: ["Jan", "Feb", "Mar"]
        }
        
        self.QUARTER_MONTHS = quarter_months_map[quarter]
        
        # Generate canonical months for the full financial year
        self.CANONICAL_MONTHS = self._generate_canonical_months()
        
        # File naming based on financial year start month and quarter last month
        fy_start_month = "Apr"  # Financial year always starts in April
        quarter_last_month = self.QUARTER_MONTHS[-1]  # Last month of the quarter
        self.TARGET_FILENAME = f"Unwind of Discount - {fy_start_month}-{quarter_last_month} {self.FY_YEAR}.xlsx"
        self.FY_START_LABEL = f"{self.CANONICAL_MONTHS[0]}"
        self.OUTPUT_SUFFIX = f"_FYStart-{self.FY_START_LABEL}.xlsx"
        self.SHEET_NAME = "Maturity Int"
        self.STAGE_SHEET_NAME = "Stage Sorted to Month"

    def _generate_canonical_months(self) -> List[str]:
        """Generate 12 canonical months for the financial year starting April."""
        months_order = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
        canonical = []
        
        for i, month in enumerate(months_order):
            # Jan, Feb, Mar belong to next calendar year
            if month in ["Jan", "Feb", "Mar"]:
                year = self.FY_YEAR + 1
            else:
                year = self.FY_YEAR
            canonical.append(f"{month}-{year}")
        
        return canonical

    # ---------- Utility ----------
    def timestamp(self) -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def find_unwind_file(self) -> Tuple[Path, bool]:
        """Find the Unwind Excel file in working folder."""
        if not self.WORKING_DIR.exists():
            raise FileNotFoundError(f"Working directory not found: {self.WORKING_DIR}")

        target = self.WORKING_DIR / self.TARGET_FILENAME
        if target.exists():
            return target, True

        # Flexible pattern matching for quarter files
        pattern = f"Unwind of Discount*{self.QUARTER_MONTHS[0]}*{self.QUARTER_MONTHS[-1]}*{self.FY_YEAR}*.xlsx"
        for p in self.WORKING_DIR.glob(pattern):
            return p, False
        
        # Fallback: any Unwind file
        for p in self.WORKING_DIR.glob("Unwind of Discount*.xlsx"):
            print(f"⚠ Using fallback file: {p.name}")
            return p, False

        return None, False

    def format_month_name(self, month: str, current_year: int) -> str:
        month_name, year = month.split("-")
        year = int(year)
        full_names = {
            "Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April",
            "May": "May", "Jun": "June", "Jul": "July", "Aug": "August",
            "Sep": "September", "Oct": "October", "Nov": "November", "Dec": "December"
        }
        if year == current_year:
            return full_names.get(month_name, month_name)
        else:
            return f"{month_name}-{str(year)[-2:]}"

    def get_unique_output_path(self, base_path: Path) -> Path:
        """Generate unique file name if same exists."""
        if not base_path.exists():
            return base_path
        stem, suffix = base_path.stem, base_path.suffix
        copy_num = 1
        while True:
            new_path = base_path.parent / f"{stem} (Copy {copy_num}){suffix}"
            if not new_path.exists():
                return new_path
            copy_num += 1

    # ---------- Excel Handling ----------
    def ensure_maturity_months(self, wb_path: Path, save: bool = True) -> Dict:
        """Ensure correct month headers (B2:M2)."""
        log = {"inserted": [], "errors": []}
        print("Loading workbook for header update...")

        try:
            wb = load_workbook(wb_path, data_only=False)
        except Exception as e:
            log["errors"].append(f"Failed to load workbook: {e}")
            return log

        if self.SHEET_NAME not in wb.sheetnames:
            log["errors"].append(f"Sheet '{self.SHEET_NAME}' not found")
            wb.close()
            return log

        sheet = wb[self.SHEET_NAME]
        current_year = int(self.FY_START_LABEL.split("-")[1])

        print("Updating month headers in B2:M2...")
        for idx, month in enumerate(self.CANONICAL_MONTHS[:12]):
            col = idx + 2
            formatted = self.format_month_name(month, current_year)
            old_val = sheet.cell(row=2, column=col).value
            sheet.cell(row=2, column=col, value=formatted)
            if old_val != formatted:
                log["inserted"].append(f"Col {col}: '{old_val}' → '{formatted}'")

        if save:
            try:
                wb.save(wb_path)
                wb.close()
                print(f"✓ Header updated and saved: {wb_path}")
            except PermissionError:
                print(f"✗ File is open in Excel. Close it and rerun: {wb_path}")
                wb.close()
            except Exception as e:
                log["errors"].append(f"Error saving workbook: {e}")
                wb.close()
        else:
            wb.close()

        return log

    # ---------- CSV Handling ----------
    def get_maturity_sheet_name(self, month: str) -> str:
        month_map = {
            "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
            "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
            "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"
        }
        month_name, year = month.split("-")
        return f"MATINT-01-{month_map.get(month_name, '01')}-{year}.csv"

    def load_maturity_data(self, month: str) -> Dict[str, float]:
        """Read maturity report CSV for a given month."""
        data = {}
        csv_file = self.MATURITY_REPORTS_DIR / self.get_maturity_sheet_name(month)

        if not csv_file.exists():
            print(f"⚠ Maturity file not found: {csv_file.name}")
            return data

        print(f"Reading: {csv_file.name}")
        try:
            with open(csv_file, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                
                # Find CON_NO and ARS_INTAMT1 columns (case-insensitive)
                contract_col = None
                amount_col = None
                
                for fieldname in reader.fieldnames:
                    fieldname_upper = fieldname.strip().upper()
                    if fieldname_upper in ["CON_NO", "CONTRACT_NO"]:
                        contract_col = fieldname
                    elif fieldname_upper == "ARS_INTAMT1":
                        amount_col = fieldname
                
                if not contract_col or not amount_col:
                    print(f"⚠ Required columns not found. Available: {reader.fieldnames}")
                    return data

                for row in reader:
                    contract = row.get(contract_col, "").strip()
                    amount = row.get(amount_col, "")
                    
                    if contract:
                        try:
                            data[contract] = float(amount) if amount else 0.0
                        except ValueError:
                            data[contract] = 0.0
                            
        except Exception as e:
            print(f"✗ Error reading {csv_file.name}: {e}")

        print(f"✓ {len(data)} contracts loaded")
        return data

    def collect_all_contracts(self) -> Dict[str, Dict[str, float]]:
        """Collect all contract values from quarter months only."""
        print("\nCollecting maturity report data for current quarter...")
        all_data = {}
        missing_files = []
        
        # Get only the months for the current quarter
        quarter_canonical_months = []
        for month_abbr in self.QUARTER_MONTHS:
            # Find the corresponding canonical month (with year)
            for canonical_month in self.CANONICAL_MONTHS:
                if canonical_month.startswith(month_abbr):
                    quarter_canonical_months.append(canonical_month)
                    break
        
        print(f"Quarter months to process: {quarter_canonical_months}")
        
        # Load data for each month in the quarter
        for month in quarter_canonical_months:
            csv_filename = self.get_maturity_sheet_name(month)
            csv_file = self.MATURITY_REPORTS_DIR / csv_filename
            
            # Check if file exists
            if not csv_file.exists():
                missing_files.append(csv_filename)
                print(f"✗ MISSING: {csv_filename}")
                continue
            
            month_data = self.load_maturity_data(month)
            if not month_data:
                print(f"⚠ Warning: No data loaded from {csv_filename}")
            else:
                for contract, value in month_data.items():
                    all_data.setdefault(contract, {})[month] = value
        
        # Raise error if any required files are missing
        if missing_files:
            error_msg = f"Required maturity report files not found:\n" + "\n".join(f"  - {f}" for f in missing_files)
            raise FileNotFoundError(error_msg)
        
        print(f"✓ Total unique contracts: {len(all_data)}")
        return all_data

    # ---------- Populate Workbook ----------
    def populate_intamt_values(self, wb_path: Path, save: bool = True) -> Dict:
        """Fill ARS_INTAMT1 data into Unwind file."""
        log = {"populated": [], "errors": []}
        print("\nPopulating ARS_INTAMT1 values...")

        try:
            wb = load_workbook(wb_path, data_only=False)
        except Exception as e:
            log["errors"].append(f"Cannot open workbook: {e}")
            return log

        if self.SHEET_NAME not in wb.sheetnames:
            log["errors"].append(f"Sheet '{self.SHEET_NAME}' not found")
            wb.close()
            return log

        sheet = wb[self.SHEET_NAME]
        
        # Collect quarter-specific contracts (will raise error if files missing)
        try:
            all_contracts = self.collect_all_contracts()
        except FileNotFoundError as e:
            log["errors"].append(str(e))
            print(f"\n✗ {e}")
            wb.close()
            return log
        
        if not all_contracts:
            log["errors"].append("No maturity data found for the selected quarter")
            wb.close()
            return log

        # Get only the months for the current quarter
        quarter_canonical_months = []
        for month_abbr in self.QUARTER_MONTHS:
            for canonical_month in self.CANONICAL_MONTHS:
                if canonical_month.startswith(month_abbr):
                    quarter_canonical_months.append(canonical_month)
                    break

        # Existing contracts
        existing = {}
        for row in range(3, sheet.max_row + 1):
            val = sheet.cell(row=row, column=1).value
            if val:
                existing[str(val).strip()] = row

        added, updated = 0, 0
        next_row = max(existing.values(), default=2) + 1

        # Track monthly sums for row 1 (only for quarter months)
        monthly_sums = {month: 0.0 for month in quarter_canonical_months}
        grand_total = 0.0

        for contract, monthly_data in sorted(all_contracts.items()):
            if contract in existing:
                row = existing[contract]
                updated += 1
            else:
                row = next_row
                sheet.cell(row=row, column=1, value=contract)
                existing[contract] = row
                next_row += 1
                added += 1

            # Calculate row total (only for quarter months)
            row_total = 0.0
            for idx, month in enumerate(self.CANONICAL_MONTHS[:12]):
                col = idx + 2
                
                # Only populate quarter months with actual data
                if month in quarter_canonical_months:
                    val = monthly_data.get(month, 0.0)
                    sheet.cell(row=row, column=col, value=val)
                    monthly_sums[month] += val
                    row_total += val
                else:
                    # Clear non-quarter months or leave as 0
                    sheet.cell(row=row, column=col, value=0.0)

            # Write row total to column N (column 14)
            sheet.cell(row=row, column=14, value=row_total)
            grand_total += row_total
            
            # Add VLOOKUP formula for STAGE in column O (column 15)
            stage_formula = f"=VLOOKUP(A{row},'Stage Sorted to Month'!A:C,2,FALSE)"
            sheet.cell(row=row, column=15, value=stage_formula)

        # Write monthly sums to row 1 (only for quarter columns)
        print("\nAdding monthly sum totals to row 1 (quarter months only)...")
        for idx, month in enumerate(self.CANONICAL_MONTHS[:12]):
            col = idx + 2
            if month in quarter_canonical_months:
                total = monthly_sums[month]
                sheet.cell(row=1, column=col, value=total)
                print(f"  {month}: {total:,.2f}")
            else:
                # Clear non-quarter month totals
                sheet.cell(row=1, column=col, value=0.0)

        # Write grand total to row 1, column N (Total column)
        sheet.cell(row=1, column=14, value=grand_total)
        print(f"  Grand Total: {grand_total:,.2f}")

        # Update headers in row 2 if needed
        if not sheet.cell(row=2, column=14).value:
            sheet.cell(row=2, column=14, value="Total")
        if not sheet.cell(row=2, column=15).value:
            sheet.cell(row=2, column=15, value="Stage")

        print(f"\n✓ Added {added} new contracts, updated {updated}")
        print(f"✓ Added STAGE column with VLOOKUP formulas")
        log["populated"].append(f"Added {added}, Updated {updated}, Total {len(all_contracts)}")
        log["populated"].append(f"Quarter months processed: {', '.join(quarter_canonical_months)}")
        log["populated"].append(f"Monthly sums and totals added to row 1")
        log["populated"].append(f"STAGE column added with VLOOKUP formulas")

        if save:
            try:
                wb.save(wb_path)
                print(f"✓ Workbook saved: {wb_path}")
            except PermissionError:
                print(f"✗ Cannot save workbook — file might be open in Excel: {wb_path}")
            except Exception as e:
                log["errors"].append(f"Failed to save workbook: {e}")
            finally:
                wb.close()
        else:
            wb.close()

        return log

    # ---------- Portfolio Summary Handling ----------
    def get_portfolio_summary_file(self, month: str) -> Path:
        """Find portfolio summary file for a given month."""
        month_name, year = month.split("-")
        month_map = {
            "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
            "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
            "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"
        }
        month_num = month_map.get(month_name, "01")
        
        # Pattern to match: 3. Summary_YYYY-MM-DD*
        pattern_prefix = f"3. Summary_{year}-{month_num}-"
        
        if not self.PORTFOLIOS_DIR.exists():
            return None
        
        # Search for matching files with various extensions (.xlsx, .xlsb, .xlsm)
        for file in self.PORTFOLIOS_DIR.iterdir():
            if file.is_file() and file.name.startswith(pattern_prefix):
                # Check for valid Excel extensions
                if file.suffix.lower() in ['.xlsx', '.xlsb', '.xlsm']:
                    return file
        
        return None

    def load_portfolio_stage_data(self, month: str) -> List[Tuple[str, str]]:
        """Load contract numbers and stages from portfolio summary file."""
        data = []
        portfolio_file = self.get_portfolio_summary_file(month)
        
        if not portfolio_file or not portfolio_file.exists():
            print(f"⚠ Portfolio summary file not found for {month}")
            return data
        
        print(f"Reading: {portfolio_file.name}")
        
        # Check file extension and use appropriate library
        if portfolio_file.suffix.lower() == '.xlsb':
            if not XLSB_SUPPORT:
                print(f"✗ Cannot read .xlsb file: pyxlsb not installed")
                print(f"   Install with: pip install pyxlsb")
                return data
            return self._load_xlsb_portfolio_data(portfolio_file)
        else:
            return self._load_xlsx_portfolio_data(portfolio_file)

    def _load_xlsx_portfolio_data(self, portfolio_file: Path) -> List[Tuple[str, int]]:
        """Load portfolio data from .xlsx/.xlsm files using openpyxl."""
        data = []
        try:
            wb = load_workbook(portfolio_file, data_only=True, read_only=True)
            
            if "SUMMARY" not in wb.sheetnames:
                print(f"⚠ 'SUMMARY' sheet not found in {portfolio_file.name}")
                wb.close()
                return data
            
            sheet = wb["SUMMARY"]
            
            # Check row 3 for headers
            contract_col = None
            stage_col = None
            header_row = 3
            
            # First, print all headers in row 3 for debugging
            print(f"  Headers in row {header_row}:")
            for col in range(1, min(sheet.max_column + 1, 50)):
                cell_value = sheet.cell(row=header_row, column=col).value
                if cell_value:
                    print(f"    Col {col}: '{cell_value}'")
            
            # Find CONTRACT NO and STAGE columns
            for col in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row=header_row, column=col).value
                if cell_value:
                    cell_str = str(cell_value).strip().upper()
                    if "CONTRACT" in cell_str and ("NO" in cell_str or "NUMBER" in cell_str):
                        contract_col = col
                        print(f"  Found CONTRACT column at: {col} ('{cell_value}')")
                    elif "STAGE" in cell_str:
                        stage_col = col
                        print(f"  Found STAGE column at: {col} ('{cell_value}')")
            
            if not contract_col or not stage_col:
                print(f"⚠ Required columns not found in {portfolio_file.name}")
                print(f"   CONTRACT NO column: {contract_col}, STAGE column: {stage_col}")
                wb.close()
                return data
            
            # Read data rows (starting from row after header)
            for row in range(header_row + 1, sheet.max_row + 1):
                contract_no = sheet.cell(row=row, column=contract_col).value
                stage = sheet.cell(row=row, column=stage_col).value
                
                if contract_no:
                    contract_str = str(contract_no).strip()
                    # Convert stage to integer
                    stage_int = None
                    if stage is not None:
                        try:
                            stage_float = float(stage)
                            stage_int = int(stage_float)
                        except (ValueError, TypeError):
                            pass
                    
                    if stage_int is not None:
                        data.append((contract_str, stage_int))
            
            wb.close()
            print(f"✓ Loaded {len(data)} contracts with stages")
            
        except Exception as e:
            print(f"✗ Error reading {portfolio_file.name}: {e}")
        
        return data

    def _load_xlsb_portfolio_data(self, portfolio_file: Path) -> List[Tuple[str, int]]:
        """Load portfolio data from .xlsb files using pyxlsb."""
        data = []
        try:
            with open_xlsb(portfolio_file) as wb:
                # Check if SUMMARY sheet exists
                if "SUMMARY" not in wb.sheets:
                    print(f"⚠ 'SUMMARY' sheet not found in {portfolio_file.name}")
                    return data
                
                with wb.get_sheet("SUMMARY") as sheet:
                    # Read all rows
                    rows = list(sheet.rows())
                    
                    if not rows or len(rows) < 3:
                        print(f"⚠ Not enough rows in SUMMARY sheet")
                        return data
                    
                    # Header is in row 3 (index 2)
                    header_row = rows[2]
                    contract_col = None
                    stage_col = None
                    
                    # First, print all headers in row 3 for debugging
                    print(f"  Headers in row 3:")
                    for col_idx, cell in enumerate(header_row[:50]):
                        if cell and cell.v:
                            print(f"    Col {col_idx}: '{cell.v}'")
                    
                    # Find CONTRACT NO and STAGE columns
                    for col_idx, cell in enumerate(header_row):
                        if cell.v:
                            cell_str = str(cell.v).strip().upper()
                            if "CONTRACT" in cell_str and ("NO" in cell_str or "NUMBER" in cell_str):
                                contract_col = col_idx
                                print(f"  Found CONTRACT column at: {col_idx} ('{cell.v}')")
                            elif "STAGE" in cell_str:
                                stage_col = col_idx
                                print(f"  Found STAGE column at: {col_idx} ('{cell.v}')")
                    
                    if contract_col is None or stage_col is None:
                        print(f"⚠ Required columns not found in {portfolio_file.name}")
                        print(f"   CONTRACT NO column: {contract_col}, STAGE column: {stage_col}")
                        return data
                    
                    # Read data rows (starting from row 4, index 3)
                    for row in rows[3:]:
                        if len(row) > max(contract_col, stage_col):
                            contract_cell = row[contract_col]
                            stage_cell = row[stage_col]
                            
                            if contract_cell and contract_cell.v:
                                contract_str = str(contract_cell.v).strip()
                                # Convert stage to integer
                                stage_int = None
                                if stage_cell and stage_cell.v is not None:
                                    try:
                                        stage_float = float(stage_cell.v)
                                        stage_int = int(stage_float)
                                    except (ValueError, TypeError):
                                        pass
                            
                            if stage_int is not None:
                                data.append((contract_str, stage_int))
            
            print(f"✓ Loaded {len(data)} contracts with stages")
            
        except Exception as e:
            print(f"✗ Error reading {portfolio_file.name}: {e}")
            import traceback
            traceback.print_exc()
        
        return data

    def format_month_label(self, month: str) -> str:
        """Format month as MMM-YY (e.g., JAN-25)."""
        month_name, year = month.split("-")
        year_short = year[-2:]
        return f"{month_name.upper()}-{year_short}"

    def populate_stage_sorted_data(self, wb_path: Path, save: bool = True) -> Dict:
        """Populate Stage Sorted to Month sheet with portfolio data."""
        log = {"populated": [], "errors": []}
        print("\nPopulating Stage Sorted to Month sheet...")
        
        try:
            wb = load_workbook(wb_path, data_only=False)
        except Exception as e:
            log["errors"].append(f"Cannot open workbook: {e}")
            return log
        
        # Check if Stage Sorted to Month sheet exists, create if not
        if self.STAGE_SHEET_NAME not in wb.sheetnames:
            print(f"Creating '{self.STAGE_SHEET_NAME}' sheet...")
            ws = wb.create_sheet(self.STAGE_SHEET_NAME)
            # Add headers
            ws.cell(row=1, column=1, value="Contract No")
            ws.cell(row=1, column=2, value="Stage")
            ws.cell(row=1, column=3, value="Month")
        else:
            ws = wb[self.STAGE_SHEET_NAME]
        
        # Get quarter months
        quarter_canonical_months = []
        for month_abbr in self.QUARTER_MONTHS:
            for canonical_month in self.CANONICAL_MONTHS:
                if canonical_month.startswith(month_abbr):
                    quarter_canonical_months.append(canonical_month)
                    break
        
        print(f"Processing portfolio summaries for: {quarter_canonical_months}")
        
        # Clear existing data (keep headers)
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        
        current_row = 2
        total_contracts = 0
        missing_files = []
        
        # Process each quarter month
        for month in quarter_canonical_months:
            portfolio_data = self.load_portfolio_stage_data(month)
            
            if not portfolio_data:
                missing_files.append(month)
                continue
            
            month_label = self.format_month_label(month)
            
            # Write data to sheet
            for contract_no, stage in portfolio_data:
                ws.cell(row=current_row, column=1, value=contract_no)
                ws.cell(row=current_row, column=2, value=stage)  # stage is now int
                ws.cell(row=current_row, column=3, value=month_label)
                current_row += 1
                total_contracts += 1
            
            log["populated"].append(f"{month} ({month_label}): {len(portfolio_data)} contracts")
            print(f"  ✓ {month} ({month_label}): {len(portfolio_data)} contracts added")
        
        if missing_files:
            warning_msg = f"Portfolio summary files not found for: {', '.join(missing_files)}"
            log["errors"].append(warning_msg)
            print(f"\n⚠ {warning_msg}")
        
        print(f"\n✓ Total contracts added to Stage Sorted sheet: {total_contracts}")
        log["populated"].append(f"Total: {total_contracts} contracts")
        
        if save:
            try:
                wb.save(wb_path)
                print(f"✓ Workbook saved with Stage Sorted data")
            except PermissionError:
                print(f"✗ Cannot save workbook — file might be open in Excel")
            except Exception as e:
                log["errors"].append(f"Failed to save workbook: {e}")
            finally:
                wb.close()
        else:
            wb.close()
        
        return log

    # ---------- Pivot Table Handling ----------
    def update_pivot_table_sheet(self, wb_path: Path, save: bool = True) -> Dict:
        """Update pivot table in Sheet1 with Stage 1–3 totals and grand total.
           This version reads Stage Sorted to Month directly (no dependency on VLOOKUP formulas)."""
        log = {"updated": [], "errors": []}
        print("\nUpdating pivot table in Sheet1 (direct lookup)...")

        try:
            wb = load_workbook(wb_path, data_only=False)
        except Exception as e:
            log["errors"].append(f"Cannot open workbook: {e}")
            return log

        if self.SHEET_NAME not in wb.sheetnames:
            log["errors"].append(f"Sheet '{self.SHEET_NAME}' not found")
            wb.close()
            return log

        if self.STAGE_SHEET_NAME not in wb.sheetnames:
            log["errors"].append(f"'{self.STAGE_SHEET_NAME}' not found")
            wb.close()
            return log

        maturity_sheet = wb[self.SHEET_NAME]
        stage_sheet = wb[self.STAGE_SHEET_NAME]

        # Build contract -> stage map from Stage Sorted to Month
        contract_stage_map = {}
        print("  Building contract->stage map from Stage Sorted sheet...")
        for r in range(2, stage_sheet.max_row + 1):
            contract = stage_sheet.cell(row=r, column=1).value
            stage_val = stage_sheet.cell(row=r, column=2).value
            if not contract:
                continue
            key = str(contract).strip()
            try:
                stage_int = int(float(stage_val))
            except Exception:
                continue
            contract_stage_map[key] = stage_int

        print(f"  Contracts in stage map: {len(contract_stage_map)}")

        # Initialize totals for stages 1..3
        stage_totals = {1: 0.0, 2: 0.0, 3: 0.0}
        contracts_by_stage = {1: 0, 2: 0, 3: 0}

        # Iterate maturity sheet rows and sum totals using lookup
        print("  Summing totals from Maturity Int sheet by matching contract numbers...")
        for row in range(3, maturity_sheet.max_row + 1):
            contract = maturity_sheet.cell(row=row, column=1).value
            total_value = maturity_sheet.cell(row=row, column=14).value
            if not contract:
                continue
            
            key = str(contract).strip()
            stage_int = contract_stage_map.get(key)
            
            # Fallback: try stripping leading zeros
            if stage_int is None:
                try_key = key.lstrip("0").strip()
                if try_key:
                    stage_int = contract_stage_map.get(try_key)

            if stage_int in (1, 2, 3):
                try:
                    total_float = float(total_value) if total_value is not None else 0.0
                except Exception:
                    total_float = 0.0

                stage_totals[stage_int] += total_float
                contracts_by_stage[stage_int] += 1

        print("\n  Calculated stage totals:")
        for s in (1, 2, 3):
            print(f"    Stage {s}: {contracts_by_stage[s]} contracts, Total = {stage_totals[s]:,.2f}")

        # Get or create Sheet1
        if "Sheet1" not in wb.sheetnames:
            wb.create_sheet("Sheet1")
        
        sheet1 = wb["Sheet1"]
        
        # Clear ALL data from Sheet1
        print("  Clearing all data from Sheet1...")
        for row in sheet1.iter_rows():
            for cell in row:
                cell.value = None
                cell.style = 'Normal'
        
        # Create complete table structure from scratch
        print("  Creating complete table structure...")
        
        # Define styles
        header_font = Font(name='Calibri', size=11, bold=True)
        data_font = Font(name='Calibri', size=11)
        total_font = Font(name='Calibri', size=11, bold=True)
        
        center_alignment = Alignment(horizontal='center', vertical='center')
        right_alignment = Alignment(horizontal='right', vertical='center')
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        total_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        
        # Row 3: Headers
        sheet1.cell(row=3, column=1, value="Row Labels")
        sheet1.cell(row=3, column=2, value="Sum of Total")
        
        # Format headers
        for col in [1, 2]:
            cell = sheet1.cell(row=3, column=col)
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = thin_border
            cell.fill = header_fill
        
        # Set column widths
        sheet1.column_dimensions['A'].width = 15
        sheet1.column_dimensions['B'].width = 20
        
        # Rows 4-6: Stage data
        stage_rows = {1: 4, 2: 5, 3: 6}
        grand_total = 0.0
        
        for stage, row_num in stage_rows.items():
            total = stage_totals.get(stage, 0.0)
            
            # Column A: Stage number
            cell_a = sheet1.cell(row=row_num, column=1, value=stage)
            cell_a.font = data_font
            cell_a.alignment = center_alignment
            cell_a.border = thin_border
            
            # Column B: Total value with number format
            cell_b = sheet1.cell(row=row_num, column=2, value=total)
            cell_b.font = data_font
            cell_b.alignment = right_alignment
            cell_b.border = thin_border
            cell_b.number_format = '#,##0'  # Number format with thousand separator
            
            grand_total += total
            log["updated"].append(f"Stage {stage}: {total:,.2f}")
            print(f"    Row {row_num}: Stage {stage} = {total:,.2f}")

        # Row 7: Grand Total
        cell_a = sheet1.cell(row=7, column=1, value="Grand Total")
        cell_a.font = total_font
        cell_a.alignment = center_alignment
        cell_a.border = thin_border
        cell_a.fill = total_fill
        
        cell_b = sheet1.cell(row=7, column=2, value=grand_total)
        cell_b.font = total_font
        cell_b.alignment = right_alignment
        cell_b.border = thin_border
        cell_b.fill = total_fill
        cell_b.number_format = '#,##0'  # Number format with thousand separator
        
        log["updated"].append(f"Grand Total: {grand_total:,.2f}")
        print(f"    Row 7: Grand Total = {grand_total:,.2f}")

        print(f"\n✓ Complete table structure created in Sheet1 with formatting")
        print(f"    Grand Total: {grand_total:,.2f}")

        # Close workbook with optional save
        if save:
            try:
                wb.save(wb_path)
                print(f"✓ Pivot table saved to workbook")
            except Exception as e:
                log["errors"].append(f"Failed to save workbook: {e}")
                print(f"✗ Failed to save: {e}")
            finally:
                wb.close()
        else:
            wb.close()

        return log

    # ---------- Main Process ----------
    def process_all_in_one(self, wb_path: Path) -> Dict:
        """Perform all operations on the workbook in a single session."""
        log = {
            "headers": {"inserted": [], "errors": []},
            "intamt": {"populated": [], "errors": []},
            "stage": {"populated": [], "errors": []},
            "pivot": {"updated": [], "errors": []}
        }
        
        print("\nOpening workbook for all operations...")
        try:
            wb = load_workbook(wb_path, data_only=False)
            print("✓ Workbook loaded successfully")
        except Exception as e:
            error_msg = f"Cannot open workbook: {e}"
            log["headers"]["errors"].append(error_msg)
            print(f"✗ {error_msg}")
            return log

        # ===== STEP 1: Update Headers =====
        print("\n--- Step 1: Updating Headers ---")
        if self.SHEET_NAME in wb.sheetnames:
            sheet = wb[self.SHEET_NAME]
            current_year = int(self.FY_START_LABEL.split("-")[1])
            
            for idx, month in enumerate(self.CANONICAL_MONTHS[:12]):
                col = idx + 2
                formatted = self.format_month_name(month, current_year)
                old_val = sheet.cell(row=2, column=col).value
                sheet.cell(row=2, column=col, value=formatted)
                if old_val != formatted:
                    log["headers"]["inserted"].append(f"Col {col}: '{old_val}' → '{formatted}'")
            print(f"✓ Headers updated: {len(log['headers']['inserted'])} changes")
        else:
            log["headers"]["errors"].append(f"Sheet '{self.SHEET_NAME}' not found")

        # ===== STEP 2: Populate Maturity Data =====
        print("\n--- Step 2: Populating Maturity Data ---")
        try:
            all_contracts = self.collect_all_contracts()
            
            if self.SHEET_NAME in wb.sheetnames and all_contracts:
                sheet = wb[self.SHEET_NAME]
                
                quarter_canonical_months = []
                for month_abbr in self.QUARTER_MONTHS:
                    for canonical_month in self.CANONICAL_MONTHS:
                        if canonical_month.startswith(month_abbr):
                            quarter_canonical_months.append(canonical_month)
                            break
                
                existing = {}
                for row in range(3, sheet.max_row + 1):
                    val = sheet.cell(row=row, column=1).value
                    if val:
                        existing[str(val).strip()] = row
                
                added, updated = 0, 0
                next_row = max(existing.values(), default=2) + 1
                monthly_sums = {month: 0.0 for month in quarter_canonical_months}
                grand_total = 0.0
                
                for contract, monthly_data in sorted(all_contracts.items()):
                    if contract in existing:
                        row = existing[contract]
                        updated += 1
                    else:
                        row = next_row
                        sheet.cell(row=row, column=1, value=contract)
                        existing[contract] = row
                        next_row += 1
                        added += 1
                    
                    row_total = 0.0
                    for idx, month in enumerate(self.CANONICAL_MONTHS[:12]):
                        col = idx + 2
                        if month in quarter_canonical_months:
                            val = monthly_data.get(month, 0.0)
                            sheet.cell(row=row, column=col, value=val)
                            monthly_sums[month] += val
                            row_total += val
                        else:
                            sheet.cell(row=row, column=col, value=0.0)
                    
                    sheet.cell(row=row, column=14, value=row_total)
                    grand_total += row_total
                    
                    stage_formula = f"=VLOOKUP(A{row},'Stage Sorted to Month'!A:C,2,FALSE)"
                    sheet.cell(row=row, column=15, value=stage_formula)
                
                for idx, month in enumerate(self.CANONICAL_MONTHS[:12]):
                    col = idx + 2
                    if month in quarter_canonical_months:
                        sheet.cell(row=1, column=col, value=monthly_sums[month])
                    else:
                        sheet.cell(row=1, column=col, value=0.0)
                
                sheet.cell(row=1, column=14, value=grand_total)
                
                if not sheet.cell(row=2, column=14).value:
                    sheet.cell(row=2, column=14, value="Total")
                if not sheet.cell(row=2, column=15).value:
                    sheet.cell(row=2, column=15, value="Stage")

                log["intamt"]["populated"].append(f"Added {added}, Updated {updated}, Total {len(all_contracts)}")
                print(f"✓ Added {added} new contracts, updated {updated}")
        except FileNotFoundError as e:
            log["intamt"]["errors"].append(str(e))
            print(f"✗ {e}")

        # ===== STEP 3: Populate Stage Sorted to Month =====
        print("\n--- Step 3: Populating Stage Sorted to Month ---")
        if self.STAGE_SHEET_NAME not in wb.sheetnames:
            ws = wb.create_sheet(self.STAGE_SHEET_NAME)
            ws.cell(row=1, column=1, value="Contract No")
            ws.cell(row=1, column=2, value="Stage")
            ws.cell(row=1, column=3, value="Month")
        else:
            ws = wb[self.STAGE_SHEET_NAME]
        
        quarter_canonical_months = []
        for month_abbr in self.QUARTER_MONTHS:
            for canonical_month in self.CANONICAL_MONTHS:
                if canonical_month.startswith(month_abbr):
                    quarter_canonical_months.append(canonical_month)
                    break
        
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        
        current_row = 2
        total_contracts = 0
        
        for month in quarter_canonical_months:
            portfolio_data = self.load_portfolio_stage_data(month)
            
            if not portfolio_data:
                continue
            
            month_label = self.format_month_label(month)
            
            for contract_no, stage in portfolio_data:
                ws.cell(row=current_row, column=1, value=contract_no)
                ws.cell(row=current_row, column=2, value=stage)
                ws.cell(row=current_row, column=3, value=month_label)
                current_row += 1
                total_contracts += 1
            
            log["stage"]["populated"].append(f"{month} ({month_label}): {len(portfolio_data)} contracts")
        
        log["stage"]["populated"].append(f"Total: {total_contracts} contracts")
        print(f"✓ Total {total_contracts} contracts added to Stage Sorted sheet")

        # Save before updating pivot table (to calculate VLOOKUP formulas)
        print("\n--- Saving Changes Before Pivot Update ---")
        try:
            wb.save(wb_path)
            wb.close()
            print(f"✓ Changes saved to calculate formulas")
        except Exception as e:
            log["pivot"]["errors"].append(f"Failed to save before pivot update: {e}")
            wb.close()
            print(f"✗ Failed to save: {e}")
            return log

        # ===== STEP 4: Update Pivot Table with save=True =====
        print("\n--- Step 4: Updating Pivot Table ---")
        pivot_log = self.update_pivot_table_sheet(wb_path, save=True)
        log["pivot"] = pivot_log
        
        # No need for additional save - update_pivot_table_sheet handles it
        print(f"✓ All operations completed successfully")
        
        return log

    # ---------- Main Process ----------
    def process(self):
        print(f"\n{'='*65}")
        print(f"Unwind of Discount Processor - Q{self.QUARTER} FY{self.FY_YEAR}-{str(self.FY_YEAR+1)[-2:]}")
        print(f"Started at {self.timestamp()}")
        print(f"{'='*65}\n")

        try:
            src_path, found = self.find_unwind_file()
            if not src_path:
                print(f"✗ Unwind file not found in {self.WORKING_DIR}")
                print(f"   Expected: {self.TARGET_FILENAME}")
                return

            print(f"✓ Found: {src_path.name}")
            self.OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

            # Generate output filename
            fy_start_month = "Apr"
            quarter_last_month = self.QUARTER_MONTHS[-1]
            out_file = f"Unwind of Discount - {fy_start_month}-{quarter_last_month} {self.FY_YEAR}.xlsx"
            out_path = self.get_unique_output_path(self.OUTPUTS_DIR / out_file)

            print("Copying file...")
            shutil.copy(src_path, out_path)
            print(f"✓ Copied to: {out_path}")

            # Process all operations in one session
            combined_log = self.process_all_in_one(out_path)

            # --- Summary ---
            print(f"\n{'='*65}")
            print("PROCESS SUMMARY")
            print(f"{'='*65}")
            
            print(f"\nSheet1 Pivot Table:")
            for item in combined_log['pivot'].get('updated', []):
                print(f"  {item}")
            
            all_errors = (combined_log['headers'].get('errors', []) + 
                         combined_log['intamt'].get('errors', []) + 
                         combined_log['stage'].get('errors', []) +
                         combined_log['pivot'].get('errors', []))
            if all_errors:
                print(f"\nErrors/Warnings:")
                for err in all_errors:
                    print(f"  - {err}")
            else:
                print("\n✓ No errors found.")
            print(f"{'='*65}\n")

        except Exception as e:
            print(f"\n✗ UNHANDLED EXCEPTION: {e}")
            import traceback
            traceback.print_exc()
