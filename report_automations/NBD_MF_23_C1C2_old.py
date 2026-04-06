"""
NBD_MF_23_C1C2 - Excel Automation Script for CBSL Reporting

BULK 2 DATA FLOW & VALIDATION:
==============================
1. Read contracts matching pattern ^[A-Z]{4}\d{9}$ from Summary and Net Portfolio
2. Extract data from Summary sheet columns: AG, E, R, T, Y, B, J, X
3. Map to C1C2 Working sheet columns:
   - Summary AG -> C1C2 R (R_VALUE)
   - Summary E  -> C1C2 AB (AB_VALUE)
   - Summary R  -> C1C2 AD (AD_VALUE)
   - Summary T  -> C1C2 O (IIS_VALUE)
   - Summary Y  -> C1C2 K (GROSS_PORTFOLIO)
   - Summary B  -> C1C2 C (CLIENT_CODE/CLIENT_NO)
   - Summary J  -> C1C2 L (PROVISION)
   - Summary X  -> C1C2 M (DP_PROVISION)
4. Preserve numeric values throughout processing (no premature formatting)
5. Validate totals match between Summary source and C1C2 destination
6. Track all data with numeric columns for sum validation
7. Report detailed debugging information at each stage

Key validations:
- Pre-paste: Print expected sums from source data
- Post-paste: Read back pasted data and compare sums
- Final validation: Compare Summary totals vs C1C2 totals for all numeric columns
"""

import os
import pandas as pd
import xlwings as xw
import re
import time
import subprocess
from datetime import datetime
import sys

# Fix Unicode encoding for Windows console
if sys.platform.startswith('win'):
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'replace')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'replace')

# ==================== UTILITY FUNCTIONS ====================

def kill_excel_processes():
    """Kill any hanging Excel processes"""
    try:
        subprocess.call(['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
        time.sleep(2)
    except:
        pass

def log_step(step_name, start_time):
    """Log step completion with time taken"""
    elapsed = time.time() - start_time
    print(f"  ✓ Completed in {elapsed:.2f} seconds")
    return time.time()

def get_column_letter(col_num):
    """Convert column number to Excel column letter"""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result

def column_letter_to_index(col_letter):
    """Convert Excel column letter to 0-based index"""
    col_letter = col_letter.upper()
    result = 0
    for i, char in enumerate(reversed(col_letter)):
        result += (ord(char) - ord('A') + 1) * (26 ** i)
    return result - 1

def get_first_sheet_name(file_path):
    """Get the name of the first sheet in an Excel file"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(file_path)
            sheet_name = wb.sheets[0].name
            wb.close()
            app.quit()
            return sheet_name
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def is_valid_contract_number(value):
    """Check if value matches valid contract number patterns"""
    if not value or not isinstance(value, str):
        return False

    value_clean = str(value).strip().upper()

    if len(value_clean) < 5:
        return False

    if re.match(r'^LR\d{8}$', value_clean):
        return True

    if re.match(r'^[A-Z]{4}\d{9}$', value_clean):
        return True

    if re.match(r'^[A-Z]{2,3}\d{8,10}$', value_clean):
        return True

    if len(value_clean) >= 8 and any(c.isalpha() for c in value_clean) and any(c.isdigit() for c in value_clean):
        header_keywords = ['CONTRACT', 'NUMBER', 'NO', 'LOAN', 'YARD', 'PROPERTY', 'MORTGAGE', 'LIST', 'TOTAL', 'DATE']
        if not any(keyword in value_clean for keyword in header_keywords):
            return True

    return False

def is_valid_value(value):
    """Check if value is valid (not N/A, 0, -, or empty)"""
    if pd.isna(value) or value is None or value == '':
        return False

    value_str = str(value).strip().upper()

    if value_str in ['N/A', 'NA', '#N/A', 'NULL', 'NONE', 'NAN']:
        return False

    try:
        num_value = float(str(value).replace(',', ''))
        return num_value != 0
    except:
        return False

# ==================== FILE OPERATIONS ====================
def find_dynamic_folder():
    """
    Locate the date folder inside working/NBD_MF_23_IA.
    Uses the same pattern as NBD_MF_23_IA.py - finds project root relative to script location.
    Only one date folder exists in the directory.
    Example:
        project_root/
            working/
                NBD_MF_23_IA/
                    30-09-2025  <- only one date folder
    """
    try:
        # Step 1: Locate project root (one level above report_automations)
        root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

        # Step 2: Path to working/NBD_MF_23_IA
        ia_root = os.path.join(root_dir, "working", "NBD_MF_23_IA")
        if not os.path.exists(ia_root):
            error_msg = f"'NBD_MF_23_IA' folder not found at '{ia_root}'."
            print(f"ERROR: {error_msg}")
            return None

        # Step 3: Get all subfolders (date-based)
        date_folders = [f for f in os.listdir(ia_root) if os.path.isdir(os.path.join(ia_root, f))]
        if not date_folders:
            error_msg = f"No date folders found under '{ia_root}'."
            print(f"ERROR: {error_msg}")
            return None

        # Step 4: Get the single date folder (only one should exist)
        if len(date_folders) > 1:
            print(f"WARNING: Multiple date folders found, using first one: {date_folders[0]}")
        
        target_folder = os.path.join(ia_root, date_folders[0])
        print(f"Found working folder: {target_folder}")
        return target_folder

    except Exception as e:
        print(f"ERROR: Error finding IA folder: {e}")
        return None


def find_files(folder_path):
    """Locate all required files under the given folder by keyword matching."""
    files = {
        'summary': None, 'net': None, 'prod': None, 'loan_base': None,
        'reschedule': None, 'yard_stock': None, 'cbsl_provision': None,
        'cadre': None, 'unutilized': None, 'po_listing': None, 'portfolio_recovery': None,
        'sofp': None, 'portfoliorec': None
    }

    keywords = {
        'summary': 'Summary',
        'net': 'Net Portfolio',
        'prod': 'Prod. wise Class. of Loans',
        'loan_base': 'Loan Base',
        'reschedule': 'Reschedule Contract Details',
        'yard_stock': 'YARD STOCK',
        'cbsl_provision': 'CBSL Provision Comparison',
        'cadre': 'Cadre',
        'unutilized': 'Unutilized Amount',
        'po_listing': 'Po Listing - Internal',
        'portfolio_recovery': 'Portfolio Report Recovery - Internal',
        'sofp': 'AFL Monthly FS',
        'portfoliorec': 'PORTFOLIOREC'
    }

    for root, _, file_list in os.walk(folder_path):
        for f in file_list:
            if f.startswith("~$"):
                continue
            # Check for .xlsb, .xlsx, and .csv files (PORTFOLIOREC is CSV)
            if not f.endswith((".xlsb", ".xlsx", ".csv")):
                continue
            for key, keyword in keywords.items():
                if keyword in f and files[key] is None:
                    files[key] = os.path.join(root, f)

    return (
        files['summary'], files['net'], files['prod'], files['loan_base'],
        files['reschedule'], files['yard_stock'], files['cbsl_provision'],
        files['cadre'], files['unutilized'], files['po_listing'], files['portfolio_recovery'],
        files['sofp'], files['portfoliorec']
    )


def locate_latest_report_files(base_path):
    """Wrapper function: find latest working folder and extract report files."""
    target_folder = find_dynamic_folder(base_path)
    if not target_folder:
        print("ERROR: Could not locate target working folder.")
        return None

    print(f"Searching for files inside: {target_folder}")
    return find_files(target_folder)

# ==================== EXCEPTION FILE ====================

class ExceptionTracker:
    """Track exceptions and write to a single file"""

    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.exception_file = os.path.join(folder_path, 'EXCEPTIONS.xlsx')
        self.exceptions = {}

    def add_exception(self, sheet_name, data_dict):
        """Add exception data for a specific sheet"""
        if sheet_name not in self.exceptions:
            self.exceptions[sheet_name] = []
        self.exceptions[sheet_name].append(data_dict)

    def write_exceptions(self):
        """Write all exceptions to a single Excel file"""
        if not self.exceptions:
            return

        with pd.ExcelWriter(self.exception_file, engine='openpyxl') as writer:
            for sheet_name, exception_list in self.exceptions.items():
                df = pd.DataFrame(exception_list)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"  Exception file: {os.path.basename(self.exception_file)}")

# ==================== DATA READING ====================

def read_contracts_fast(file_path, sheet_name, column_letter):
    """Fast contract reading for large files"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if file_path.endswith('.xlsb'):
                app = xw.App(visible=False)
                app.display_alerts = False
                wb = app.books.open(file_path)
                sheet = wb.sheets[sheet_name]
                last_row = sheet.used_range.last_cell.row

                chunk_size = 50000
                all_data = []

                for start in range(1, last_row + 1, chunk_size):
                    end = min(start + chunk_size - 1, last_row)
                    chunk = sheet.range(f'{column_letter}{start}:{column_letter}{end}').value
                    all_data.extend(chunk if isinstance(chunk, list) else [chunk])

                wb.close()
                app.quit()
                contracts = all_data
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                col_idx = ord(column_letter.upper()) - ord('A')
                contracts = df.iloc[:, col_idx].tolist()

            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

    header_values = {'CONTRACT_NO', 'CONTRACT NO', 'CONTRACT NUMBER'}
    contracts_cleaned = [
        str(c).strip() for c in contracts
        if c and str(c).strip() and str(c).strip().upper() not in header_values
    ]

    return contracts_cleaned

def classify_contracts(contracts, pattern):
    """Classify contracts by regex pattern"""
    compiled_pattern = re.compile(pattern, re.IGNORECASE)
    return [c for c in contracts if compiled_pattern.match(c)]

def read_net_portfolio_data(net_file, matching_contracts):
    """Read and format Net Portfolio data for matched contracts"""
    print("  [BULK2 DEBUG] Reading Net Portfolio data...")
    print(f"  [BULK2 DEBUG] Source: {os.path.basename(net_file)} (Sheet 0)")
    print(f"  [BULK2 DEBUG] Reading columns: E(4), F(5), H(7), AB(27), AH(33), AI(34), AL(38)")

    usecols = [4, 5, 7, 27, 33, 34, 38]
    df = pd.read_excel(net_file, sheet_name=0, header=None, usecols=usecols)

    df.columns = ['CONTRACT_NO', 'CONTRACT_PERIOD', 'CONTRACT_AMOUNT', 'EQT_DESC',
                  'CON_RNTFREQ', 'CON_INTRATE', 'PURPOSE']

    print("  [BULK2 DEBUG] Data mapping from Net Portfolio:")
    print("  [BULK2 DEBUG]   Net Col E (Contract No)        -> CONTRACT_NO")
    print("  [BULK2 DEBUG]   Net Col F (Contract Period)    -> CONTRACT_PERIOD")
    print("  [BULK2 DEBUG]   Net Col H (Contract Amount)    -> CONTRACT_AMOUNT")
    print("  [BULK2 DEBUG]   Net Col AB (Equipment Desc)    -> EQT_DESC")
    print("  [BULK2 DEBUG]   Net Col AH (Rental Frequency)  -> CON_RNTFREQ")
    print("  [BULK2 DEBUG]   Net Col AI (Interest Rate)     -> CON_INTRATE")
    print("  [BULK2 DEBUG]   Net Col AL (Purpose)           -> PURPOSE")

    df['CONTRACT_NO'] = df['CONTRACT_NO'].astype(str).str.strip().str.upper()

    # Filter to matching contracts while preserving order
    contract_set = set(matching_contracts)
    df = df[df['CONTRACT_NO'].isin(contract_set)].copy()

    print(f"  [BULK2 DEBUG] Filtered to {len(df)} matching contracts from {len(matching_contracts)} input contracts")

    # Keep numeric values for validation before formatting
    df['CONTRACT_AMOUNT_NUMERIC'] = pd.to_numeric(df['CONTRACT_AMOUNT'], errors='coerce')
    df['CONTRACT_AMOUNT'] = df['CONTRACT_AMOUNT'].apply(lambda x: f"{float(x):,.0f}" if pd.notna(x) and x != '' else '')

    def safe_format_rate(x):
        """Safely format interest rate, handling invalid data"""
        if pd.notna(x) and x != '':
            try:
                return f"{float(x):.2f}%"
            except (ValueError, TypeError):
                return ''
        return ''

    df['CON_INTRATE_NUMERIC'] = pd.to_numeric(df['CON_INTRATE'], errors='coerce')
    df['CON_INTRATE'] = df['CON_INTRATE'].apply(safe_format_rate)

    for col in ['EQT_DESC', 'PURPOSE', 'CON_RNTFREQ', 'CONTRACT_PERIOD']:
        df[col] = df[col].fillna('').astype(str).str.strip()

    return df

def read_multiple_columns_from_file(file_path, sheet_name, lookup_col, data_cols, contracts):
    """Read multiple columns at once from a file for given contracts"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(file_path)

            if isinstance(sheet_name, int):
                sheet = wb.sheets[sheet_name]
            else:
                sheet = None
                for s in wb.sheets:
                    if sheet_name.upper() in s.name.upper():
                        sheet = s
                        break
                if not sheet:
                    sheet = wb.sheets[0]

            last_row = sheet.used_range.last_cell.row

            lookup_data = sheet.range(f'{lookup_col}1:{lookup_col}{last_row}').value
            lookup_list = lookup_data if isinstance(lookup_data, list) else [lookup_data]

            result_data = {}
            for col in data_cols:
                col_data = sheet.range(f'{col}1:{col}{last_row}').value
                col_list = col_data if isinstance(col_data, list) else [col_data]
                result_data[col] = col_list

            wb.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

    # Build lookup dictionary
    contract_data_map = {}
    for i, contract in enumerate(lookup_list):
        if contract and str(contract).strip():
            contract_clean = str(contract).strip().upper()
            contract_data_map[contract_clean] = {col: result_data[col][i] for col in data_cols}

    # Match contracts and preserve order
    result = {col: [] for col in data_cols}
    missing_contracts = []

    for contract in contracts:
        contract_clean = str(contract).strip().upper()
        if contract_clean in contract_data_map:
            for col in data_cols:
                result[col].append(contract_data_map[contract_clean][col])
        else:
            missing_contracts.append(contract)
            for col in data_cols:
                result[col].append(None)

    # Debug missing contracts
    if missing_contracts:
        print(f"  [DEBUG] Warning: {len(missing_contracts)} contracts not found in {os.path.basename(file_path)} sheet '{sheet_name}'")
        if len(missing_contracts) <= 5:
            print(f"  [DEBUG] Missing contracts: {missing_contracts}")

    return result, missing_contracts

def read_columns_from_large_file(file_path, lookup_col_letter, data_col_letters, contracts):
    """Read columns from large Excel files using pandas"""
    if not os.path.exists(file_path):
        raise Exception(f"File not found: {file_path}")

    try:
        lookup_col_idx = column_letter_to_index(lookup_col_letter)
        data_col_indices = [column_letter_to_index(col) for col in data_col_letters]

        usecols = [lookup_col_idx] + data_col_indices
        df = pd.read_excel(file_path, sheet_name=0, header=None, usecols=usecols)

        df.columns = ['LOOKUP'] + [f'DATA_{i}' for i in range(len(data_col_letters))]
        df['LOOKUP'] = df['LOOKUP'].astype(str).str.strip().str.upper()

        lookup_dict = {}
        for idx, row in df.iterrows():
            contract = row['LOOKUP']
            if contract and contract != 'NAN':
                lookup_dict[contract] = {data_col_letters[i]: row[f'DATA_{i}']
                                        for i in range(len(data_col_letters))}

        result = {col: [] for col in data_col_letters}
        for contract in contracts:
            contract_clean = str(contract).strip().upper()
            if contract_clean in lookup_dict:
                for col in data_col_letters:
                    result[col].append(lookup_dict[contract_clean][col])
            else:
                for col in data_col_letters:
                    result[col].append(None)

        return result

    except Exception as e:
        raise Exception(f"Failed to read file with pandas: {str(e)}")

def calculate_contract_periods(loan_base_file, loan_base_sheet, bulk1_contracts):
    """Calculate contract periods using DATEDIF between Loan Date and Mat Date"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(loan_base_file)
            sheet = wb.sheets[loan_base_sheet]

            last_row = sheet.used_range.last_cell.row

            contracts_col = sheet.range(f'B1:B{last_row}').value
            loan_dates_col = sheet.range(f'E1:E{last_row}').value
            mat_dates_col = sheet.range(f'F1:F{last_row}').value

            wb.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

    contracts_list = contracts_col if isinstance(contracts_col, list) else [contracts_col]
    loan_dates_list = loan_dates_col if isinstance(loan_dates_col, list) else [loan_dates_col]
    mat_dates_list = mat_dates_col if isinstance(mat_dates_col, list) else [mat_dates_col]

    contract_period_map = {}
    for contract, loan_date, mat_date in zip(contracts_list, loan_dates_list, mat_dates_list):
        if contract and loan_date and mat_date:
            contract_clean = str(contract).strip().upper()
            try:
                if isinstance(loan_date, str):
                    loan_date = pd.to_datetime(loan_date)
                if isinstance(mat_date, str):
                    mat_date = pd.to_datetime(mat_date)

                months_diff = (mat_date.year - loan_date.year) * 12 + (mat_date.month - loan_date.month)
                contract_period_map[contract_clean] = months_diff
            except:
                contract_period_map[contract_clean] = ''

    periods = []
    for contract in bulk1_contracts:
        contract_clean = str(contract).strip().upper()
        period = contract_period_map.get(contract_clean, '')
        periods.append(period if period != '' else '')

    return periods

def load_old_initial_valuations(prod_file):
    """Load Contract No and Initial Valuation from C1 & C2 Working and IA Working sheets at script start"""
    print("  Loading old Initial Valuation data from existing sheets...")

    df_old_c1c2 = pd.DataFrame()
    df_old_ia = pd.DataFrame()

    try:
        # Load C1 & C2 Working sheet (Contract No in col A, Initial Valuation in col AF)
        print(f"    Reading 'C1 & C2 Working' sheet...")
        df_c1c2_full = pd.read_excel(prod_file, sheet_name='C1 & C2 Working', engine='pyxlsb')
        if len(df_c1c2_full.columns) >= 32:  # AF is column 32 (0-indexed: 31)
            df_old_c1c2 = df_c1c2_full.iloc[:, [0, 31]].copy()  # Columns A and AF
            df_old_c1c2.columns = ['Contract No', 'Initial Valuation']
            df_old_c1c2 = df_old_c1c2[df_old_c1c2['Contract No'].notna()].copy()
            df_old_c1c2['Initial Valuation'] = pd.to_numeric(df_old_c1c2['Initial Valuation'], errors='coerce')
            print(f"    Loaded {len(df_old_c1c2)} rows from 'C1 & C2 Working'")
        else:
            print(f"    WARNING: 'C1 & C2 Working' sheet doesn't have enough columns")
    except Exception as e:
        print(f"    WARNING: Could not load 'C1 & C2 Working': {e}")

    try:
        # Load IA Working sheet (Contract No in col A, Initial Valuation in col V)
        print(f"    Reading 'IA Working' sheet...")
        df_ia_full = pd.read_excel(prod_file, sheet_name='IA Working', engine='pyxlsb')
        if len(df_ia_full.columns) >= 22:  # V is column 22 (0-indexed: 21)
            df_old_ia = df_ia_full.iloc[:, [0, 21]].copy()  # Columns A and V
            df_old_ia.columns = ['Contract No', 'Initial Valuation']
            df_old_ia = df_old_ia[df_old_ia['Contract No'].notna()].copy()
            df_old_ia['Initial Valuation'] = pd.to_numeric(df_old_ia['Initial Valuation'], errors='coerce')
            print(f"    Loaded {len(df_old_ia)} rows from 'IA Working'")
        else:
            print(f"    WARNING: 'IA Working' sheet doesn't have enough columns")
    except Exception as e:
        print(f"    WARNING: Could not load 'IA Working': {e}")

    return df_old_c1c2, df_old_ia

def calculate_initial_valuation_bulk2(contracts, po_listing_file, portfolio_recovery_file, prod_file, df_old_c1c2, df_old_ia, portfoliorec_file, exception_tracker):
    """Calculate Initial Valuation for Bulk 2 with 5-tier fallback logic and max value logic

    Tier 1: C1 & C2 Working (old data)
    Tier 2: IA Working (old data) - max value logic
    Tier 3: PORTFOLIOREC - max value logic
    Tier 4: Po Listing - max value logic
    Tier 5: Scienter bot for LR contracts - max value logic
    """
    print("  [BULK2 DEBUG] Calculating Initial Valuation (AF_VALUE) using 5-tier fallback with max value logic...")

    initial_valuations = [None] * len(contracts)

    # Tier 1: C1 & C2 Working (old data)
    print(f"  [BULK2 DEBUG] TIER 1: Attempting C1 & C2 Working sheet")
    if not df_old_c1c2.empty:
        try:
            print(f"  [BULK2 DEBUG]   Source: C1 & C2 Working Col A (Contract No) -> Col AF (Initial Valuation)")
            c1c2_mapping = dict(zip(df_old_c1c2['Contract No'].astype(str), df_old_c1c2['Initial Valuation']))
            c1c2_mapping = {k: v for k, v in c1c2_mapping.items() if pd.notna(v) and v > 0}

            tier1_count = 0
            for i, contract in enumerate(contracts):
                val = c1c2_mapping.get(str(contract))
                if val is not None and val > 0:
                    initial_valuations[i] = val
                    tier1_count += 1
            print(f"  [BULK2 DEBUG]   Found {tier1_count} valid values from C1 & C2 Working")
        except Exception as e:
            print(f"  [BULK2 DEBUG]   C1 & C2 Working failed: {str(e)}")
    else:
        print("  [BULK2 DEBUG]   C1 & C2 Working data not available")

    # Tier 2: IA Working (old data) - use max value logic
    print(f"  [BULK2 DEBUG] TIER 2: Attempting IA Working sheet (with max value logic)")
    if not df_old_ia.empty:
        try:
            print(f"  [BULK2 DEBUG]   Source: IA Working Col A (Contract No) -> Col V (Initial Valuation)")
            ia_mapping = dict(zip(df_old_ia['Contract No'].astype(str), df_old_ia['Initial Valuation']))
            ia_mapping = {k: v for k, v in ia_mapping.items() if pd.notna(v) and v > 0}

            tier2_count = 0
            for i, contract in enumerate(contracts):
                ia_val = ia_mapping.get(str(contract))
                if ia_val is not None and ia_val > 0:
                    current_val = initial_valuations[i]
                    if current_val is None:
                        initial_valuations[i] = ia_val
                        tier2_count += 1
                    elif ia_val > current_val:
                        initial_valuations[i] = ia_val
                        tier2_count += 1
            print(f"  [BULK2 DEBUG]   Updated/Added {tier2_count} values from IA Working")
        except Exception as e:
            print(f"  [BULK2 DEBUG]   IA Working failed: {str(e)}")
    else:
        print("  [BULK2 DEBUG]   IA Working data not available")

    # Tier 3: PortfolioRec (PORTFOLIOREC CSV) - use max value logic
    print(f"  [BULK2 DEBUG] TIER 3: Attempting PORTFOLIOREC file - {os.path.basename(portfoliorec_file) if portfoliorec_file else 'N/A'}")
    if portfoliorec_file and os.path.exists(portfoliorec_file):
        try:
            print(f"  [BULK2 DEBUG]   Source: PORTFOLIOREC Col C (Contract No) -> Col AA (SELLING_PRICE)")
            # Read CSV columns C (index 2) and AA (index 26)
            df_portfoliorec = pd.read_csv(portfoliorec_file, usecols=[2, 26])
            df_portfoliorec.columns = ['Contract No', 'Initial Valuation']
            df_portfoliorec = df_portfoliorec[df_portfoliorec['Contract No'].notna()].copy()
            df_portfoliorec['Initial Valuation'] = pd.to_numeric(df_portfoliorec['Initial Valuation'], errors='coerce')

            rec_mapping = dict(zip(df_portfoliorec['Contract No'].astype(str), df_portfoliorec['Initial Valuation']))
            rec_mapping = {k: v for k, v in rec_mapping.items() if pd.notna(v) and v > 0}

            tier3_count = 0
            for i, contract in enumerate(contracts):
                rec_val = rec_mapping.get(str(contract))
                if rec_val is not None and rec_val > 0:
                    current_val = initial_valuations[i]
                    if current_val is None:
                        initial_valuations[i] = rec_val
                        tier3_count += 1
                    elif rec_val > current_val:
                        initial_valuations[i] = rec_val
                        tier3_count += 1
            print(f"  [BULK2 DEBUG]   Updated/Added {tier3_count} values from PORTFOLIOREC")
        except Exception as e:
            print(f"  [BULK2 DEBUG]   PORTFOLIOREC failed: {str(e)}")
    else:
        print("  [BULK2 DEBUG]   PORTFOLIOREC file not available")

    # Tier 4: Po Listing - use max value logic
    print(f"  [BULK2 DEBUG] TIER 4: Attempting Po Listing file - {os.path.basename(po_listing_file) if po_listing_file else 'N/A'}")
    if po_listing_file and os.path.exists(po_listing_file):
        try:
            print(f"  [BULK2 DEBUG]   Source: Po Listing Col H (Contract No) -> Col AH (Initial Valuation)")
            po_data = read_columns_from_large_file(po_listing_file, 'H', ['AH'], contracts)
            tier4_count = 0
            for i, val in enumerate(po_data['AH']):
                if is_valid_value(val):
                    current_val = initial_valuations[i]
                    if current_val is None:
                        initial_valuations[i] = val
                        tier4_count += 1
                    elif val > current_val:
                        initial_valuations[i] = val
                        tier4_count += 1
            print(f"  [BULK2 DEBUG]   Updated/Added {tier4_count} values from Po Listing")
        except Exception as e:
            print(f"  [BULK2 DEBUG]   Po Listing failed: {str(e)}")
    else:
        print("  [BULK2 DEBUG]   Po Listing file not available")

    # Tier 5: Scienter bot for LR contracts - use max value logic
    print(f"  [BULK2 DEBUG] TIER 5: Attempting Scienter bot for LR contracts with missing Initial Valuation")
    lr_contracts_to_fetch = []
    lr_indices = []
    for i, contract in enumerate(contracts):
        # Check if this is an LR contract and has no Initial Valuation yet
        if str(contract).startswith('LR') and (initial_valuations[i] is None or initial_valuations[i] == ''):
            lr_contracts_to_fetch.append(str(contract))
            lr_indices.append(i)

    if lr_contracts_to_fetch:
        print(f"  [BULK2 DEBUG]   Found {len(lr_contracts_to_fetch)} LR contracts with missing Initial Valuation")
        print(f"  [BULK2 DEBUG]   LR contracts to process: {lr_contracts_to_fetch[:10]}..." if len(lr_contracts_to_fetch) > 10 else f"  [BULK2 DEBUG]   LR contracts to process: {lr_contracts_to_fetch}")
        scienter_bot = load_scienter_bot()
        if scienter_bot:
            try:
                print(f"  [BULK2 DEBUG]   Calling Scienter bot to fetch Initial Valuations...")
                lr_valuations = fetch_lr_valuations(lr_contracts_to_fetch, scienter_bot)

                if lr_valuations:
                    print(f"  [BULK2 DEBUG]   Received {len(lr_valuations)} valuations from Scienter bot")
                    tier5_count = 0
                    for contract, value in lr_valuations.items():
                        # Find the index of this contract
                        for i, c in enumerate(contracts):
                            if str(c) == contract:
                                current_val = initial_valuations[i]
                                if current_val is None or current_val == '':
                                    initial_valuations[i] = value
                                    tier5_count += 1
                                    print(f"  [BULK2 DEBUG]     ✓ {contract}: Set to {value:,.0f}")
                                elif value > current_val:
                                    initial_valuations[i] = value
                                    tier5_count += 1
                                    print(f"  [BULK2 DEBUG]     ✓ {contract}: Updated from {current_val:,.0f} to {value:,.0f} (max value)")
                                break
                    print(f"  [BULK2 DEBUG]   ✓ Successfully updated/added {tier5_count} values from Scienter bot")
                else:
                    print(f"  [BULK2 DEBUG]   Scienter bot returned no valuations")
            except Exception as e:
                print(f"  [BULK2 DEBUG]   Scienter bot failed: {str(e)}")
                import traceback
                print(f"  [BULK2 DEBUG]   Traceback: {traceback.format_exc()}")
        else:
            print("  [BULK2 DEBUG]   Scienter bot module not available - cannot fetch LR valuations")
    else:
        print("  [BULK2 DEBUG]   No LR contracts with missing Initial Valuation")

    # Track missing contracts
    missing_contracts = []
    missing_count = 0
    for i, val in enumerate(initial_valuations):
        if val is None:
            missing_contracts.append({
                'Contract_No': contracts[i],
                'Issue': 'Initial Valuation Missing',
                'Description': 'Not found in C1C2 Working, IA Working, PORTFOLIOREC, Po Listing, or Scienter bot',
                'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            initial_valuations[i] = ''
            missing_count += 1

    print(f"  [BULK2 DEBUG] Initial Valuation Summary:")
    print(f"  [BULK2 DEBUG]   Total contracts: {len(contracts)}")
    print(f"  [BULK2 DEBUG]   Found values: {len(contracts) - missing_count}")
    print(f"  [BULK2 DEBUG]   Missing values: {missing_count}")

    if missing_contracts:
        for exception in missing_contracts:
            exception_tracker.add_exception('Bulk2_Initial_Valuation', exception)

    formatted_values = []
    for val in initial_valuations:
        if val and val != '':
            try:
                formatted_values.append(f"{float(val):,.0f}")
            except:
                formatted_values.append(str(val))
        else:
            formatted_values.append('')

    return formatted_values

def build_bulk1_dataframe(net_file, loan_base_file, loan_base_sheet, bulk1_contracts):
    """Build complete Bulk 1 DataFrame with all required data"""
    df_bulk1 = pd.DataFrame()
    df_bulk1['CONTRACT_NO'] = bulk1_contracts

    net_data, _ = read_multiple_columns_from_file(net_file, 0, 'E', ['Z'], bulk1_contracts)
    df_bulk1['GROSS_PORTFOLIO'] = net_data['Z']

    df_bulk1['GROSS_PORTFOLIO'] = df_bulk1['GROSS_PORTFOLIO'].apply(
        lambda x: f"{float(x):,.0f}" if pd.notna(x) and x != '' and x is not None else ''
    )

    periods = calculate_contract_periods(loan_base_file, loan_base_sheet, bulk1_contracts)
    df_bulk1['CONTRACT_PERIOD'] = periods

    loan_base_data, loan_base_missing = read_multiple_columns_from_file(loan_base_file, loan_base_sheet, 'B', ['H', 'M', 'J'], bulk1_contracts)

    df_bulk1['CON_INTRATE'] = [
        f"{float(val):.2f}%" if pd.notna(val) and val != '' and val is not None else ''
        for val in loan_base_data['H']
    ]

    df_bulk1['CONTRACT_AMOUNT'] = [
        f"{float(val):,.0f}" if pd.notna(val) and val != '' and val is not None else ''
        for val in loan_base_data['M']
    ]

    df_bulk1['AF_VALUE'] = loan_base_data['J']

    # Run Scienter bot for missing LR contracts
    if loan_base_missing:
        lr_missing = [c for c in loan_base_missing if str(c).startswith('LR')]
        if lr_missing:
            print(f"  [BULK1 DEBUG] Running Scienter bot for {len(lr_missing)} missing LR contracts from Loan Base")
            print(f"  [BULK1 DEBUG] LR contracts: {lr_missing}")
            scienter_bot = load_scienter_bot()
            if scienter_bot:
                try:
                    lr_valuations = fetch_lr_valuations(lr_missing, scienter_bot)
                    if lr_valuations:
                        print(f"  [BULK1 DEBUG] Received {len(lr_valuations)} valuations from Scienter bot")
                        # Update AF_VALUE for contracts that got valuations from bot
                        for i, contract in enumerate(bulk1_contracts):
                            if str(contract) in lr_valuations:
                                df_bulk1.at[i, 'AF_VALUE'] = lr_valuations[str(contract)]
                                print(f"  [BULK1 DEBUG]   ✓ {contract}: Set Initial Valuation to {lr_valuations[str(contract)]:,.0f}")
                except Exception as e:
                    print(f"  [BULK1 DEBUG] Scienter bot failed: {str(e)}")

    return df_bulk1

def build_bulk2_dataframe(summary_file, net_file, po_listing_file, portfolio_recovery_file,
                          prod_file, portfoliorec_file, df_old_c1c2, df_old_ia, matching_contracts, exception_tracker):
    """Build complete Bulk 2 DataFrame with all required data including Initial Valuation"""
    print("\n  [BULK2 DEBUG] ========== BULK 2 DATA BUILDING ==========")
    print(f"  [BULK2 DEBUG] Processing {len(matching_contracts)} matching contracts")

    # Read base data from Net Portfolio
    df_bulk2 = read_net_portfolio_data(net_file, matching_contracts)

    # Read additional data from Summary file
    print(f"\n  [BULK2 DEBUG] Reading additional data from Summary file...")
    print(f"  [BULK2 DEBUG] Source: {os.path.basename(summary_file)} (SUMMARY sheet)")
    print(f"  [BULK2 DEBUG] Lookup: Summary Col A (Contract No)")
    print(f"  [BULK2 DEBUG] Reading: Summary Col AG -> R_VALUE")
    print(f"  [BULK2 DEBUG] Reading: Summary Col E  -> AB_VALUE")
    print(f"  [BULK2 DEBUG] Reading: Summary Col R  -> AD_VALUE")
    print(f"  [BULK2 DEBUG] Reading: Summary Col T  -> IIS_VALUE")
    print(f"  [BULK2 DEBUG] Reading: Summary Col Y  -> GROSS_PORTFOLIO (replaces Net Portfolio Col U)")
    print(f"  [BULK2 DEBUG] Reading: Summary Col B  -> CLIENT_CODE (for C1C2 Col C)")
    print(f"  [BULK2 DEBUG] Reading: Summary Col J  -> PROVISION (for C1C2 Col L)")
    print(f"  [BULK2 DEBUG] Reading: Summary Col X  -> DP_PROVISION (for C1C2 Col M)")

    summary_data, summary_missing = read_multiple_columns_from_file(summary_file, 'SUMMARY', 'A', ['AG', 'E', 'R', 'T', 'Y', 'B', 'J', 'X'], df_bulk2['CONTRACT_NO'].tolist())

    # Store numeric values for validation
    df_bulk2['R_VALUE_NUMERIC'] = pd.to_numeric(pd.Series(summary_data['AG']), errors='coerce')
    df_bulk2['AB_VALUE_NUMERIC'] = pd.to_numeric(pd.Series(summary_data['E']), errors='coerce')
    df_bulk2['AD_VALUE_NUMERIC'] = pd.to_numeric(pd.Series(summary_data['R']), errors='coerce')
    df_bulk2['IIS_VALUE_NUMERIC'] = pd.to_numeric(pd.Series(summary_data['T']), errors='coerce')
    df_bulk2['GROSS_PORTFOLIO_NUMERIC'] = pd.to_numeric(pd.Series(summary_data['Y']), errors='coerce')
    df_bulk2['PROVISION_NUMERIC'] = pd.to_numeric(pd.Series(summary_data['J']), errors='coerce')
    df_bulk2['DP_PROVISION_NUMERIC'] = pd.to_numeric(pd.Series(summary_data['X']), errors='coerce')

    # Assign raw values (will be formatted during paste)
    # R_VALUE contains Stage values from Summary AG column (already read and mapped via read_multiple_columns_from_file)
    df_bulk2['R_VALUE'] = summary_data['AG']
    df_bulk2['AB_VALUE'] = summary_data['E']
    df_bulk2['AD_VALUE'] = summary_data['R']
    df_bulk2['IIS_VALUE'] = summary_data['T']
    df_bulk2['GROSS_PORTFOLIO'] = summary_data['Y']
    df_bulk2['CLIENT_CODE'] = [
        str(val).strip() if pd.notna(val) and val != '' else ''
        for val in summary_data['B']
    ]
    df_bulk2['PROVISION'] = summary_data['J']
    df_bulk2['DP_PROVISION'] = summary_data['X']

    print(f"  [BULK2 DEBUG] Successfully mapped {len(df_bulk2)} contracts from Summary file")

    # Print summary statistics for validation
    print(f"\n  [BULK2 DEBUG] ========== DATA VALIDATION SUMS ==========")
    print(f"  [BULK2 DEBUG] R_VALUE (Summary AG)         sum: {df_bulk2['R_VALUE_NUMERIC'].sum():,.2f}")
    print(f"  [BULK2 DEBUG] AB_VALUE (Summary E)         sum: {df_bulk2['AB_VALUE_NUMERIC'].sum():,.2f}")
    print(f"  [BULK2 DEBUG] AD_VALUE (Summary R)         sum: {df_bulk2['AD_VALUE_NUMERIC'].sum():,.2f}")
    print(f"  [BULK2 DEBUG] IIS_VALUE (Summary T)        sum: {df_bulk2['IIS_VALUE_NUMERIC'].sum():,.2f}")
    print(f"  [BULK2 DEBUG] GROSS_PORTFOLIO (Summary Y)  sum: {df_bulk2['GROSS_PORTFOLIO_NUMERIC'].sum():,.2f}")
    print(f"  [BULK2 DEBUG] PROVISION (Summary J)        sum: {df_bulk2['PROVISION_NUMERIC'].sum():,.2f}")
    print(f"  [BULK2 DEBUG] DP_PROVISION (Summary X)     sum: {df_bulk2['DP_PROVISION_NUMERIC'].sum():,.2f}")

    # Calculate Initial Valuation
    print(f"\n  [BULK2 DEBUG] Calculating Initial Valuation (AF_VALUE)...")
    initial_valuations = calculate_initial_valuation_bulk2(
        df_bulk2['CONTRACT_NO'].tolist(),
        po_listing_file,
        portfolio_recovery_file,
        prod_file,
        df_old_c1c2,
        df_old_ia,
        portfoliorec_file,
        exception_tracker
    )

    df_bulk2['AF_VALUE'] = initial_valuations

    print("\n  [BULK2 DEBUG] ========== BULK 2 DATAFRAME COMPLETE ==========")
    print(f"  [BULK2 DEBUG] Final DataFrame columns: {list(df_bulk2.columns)}")
    print(f"  [BULK2 DEBUG] Total records: {len(df_bulk2)}")

    return df_bulk2

def validate_bulk2_summary_totals(summary_file, prod_file, bulk2_rows, df_bulk2, matching_contracts):
    """Validate that C1C2 totals match Summary totals for Bulk 2 contracts"""
    print("\n  [BULK2 VALIDATE] ========== VALIDATING BULK 2 TOTALS ==========")
    print("  [BULK2 VALIDATE] Comparing Summary file totals vs C1C2 pasted totals...")

    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False

            # Read Summary totals for matching contracts
            wb_summary = app.books.open(summary_file)
            summary_sheet = wb_summary.sheets['SUMMARY']
            last_row_summary = summary_sheet.used_range.last_cell.row

            summary_contracts = summary_sheet.range(f'A1:A{last_row_summary}').value
            summary_contracts_list = summary_contracts if isinstance(summary_contracts, list) else [summary_contracts]

            # Read all columns from Summary
            summary_col_ag = summary_sheet.range(f'AG1:AG{last_row_summary}').value
            summary_col_e = summary_sheet.range(f'E1:E{last_row_summary}').value
            summary_col_r = summary_sheet.range(f'R1:R{last_row_summary}').value
            summary_col_t = summary_sheet.range(f'T1:T{last_row_summary}').value
            summary_col_y = summary_sheet.range(f'Y1:Y{last_row_summary}').value
            summary_col_j = summary_sheet.range(f'J1:J{last_row_summary}').value
            summary_col_x = summary_sheet.range(f'X1:X{last_row_summary}').value

            wb_summary.close()

            # Convert to lists
            summary_col_ag_list = summary_col_ag if isinstance(summary_col_ag, list) else [summary_col_ag]
            summary_col_e_list = summary_col_e if isinstance(summary_col_e, list) else [summary_col_e]
            summary_col_r_list = summary_col_r if isinstance(summary_col_r, list) else [summary_col_r]
            summary_col_t_list = summary_col_t if isinstance(summary_col_t, list) else [summary_col_t]
            summary_col_y_list = summary_col_y if isinstance(summary_col_y, list) else [summary_col_y]
            summary_col_j_list = summary_col_j if isinstance(summary_col_j, list) else [summary_col_j]
            summary_col_x_list = summary_col_x if isinstance(summary_col_x, list) else [summary_col_x]

            # Filter to matching contracts and sum
            matching_set = set(str(c).strip().upper() for c in matching_contracts)
            summary_totals = {
                'AG': 0, 'E': 0, 'R': 0, 'T': 0, 'Y': 0, 'J': 0, 'X': 0
            }

            for i, contract in enumerate(summary_contracts_list):
                if contract and str(contract).strip().upper() in matching_set:
                    summary_totals['AG'] += pd.to_numeric(summary_col_ag_list[i], errors='coerce') or 0
                    summary_totals['E'] += pd.to_numeric(summary_col_e_list[i], errors='coerce') or 0
                    summary_totals['R'] += pd.to_numeric(summary_col_r_list[i], errors='coerce') or 0
                    summary_totals['T'] += pd.to_numeric(summary_col_t_list[i], errors='coerce') or 0
                    summary_totals['Y'] += pd.to_numeric(summary_col_y_list[i], errors='coerce') or 0
                    summary_totals['J'] += pd.to_numeric(summary_col_j_list[i], errors='coerce') or 0
                    summary_totals['X'] += pd.to_numeric(summary_col_x_list[i], errors='coerce') or 0

            # Read C1C2 totals
            wb_prod = app.books.open(prod_file)
            c1c2_sheet = wb_prod.sheets['C1 & C2 Working']

            start_row = bulk2_rows[0]
            end_row = bulk2_rows[-1]

            c1c2_col_r = c1c2_sheet.range(f'R{start_row}:R{end_row}').value
            c1c2_col_ab = c1c2_sheet.range(f'AB{start_row}:AB{end_row}').value
            c1c2_col_ad = c1c2_sheet.range(f'AD{start_row}:AD{end_row}').value
            c1c2_col_o = c1c2_sheet.range(f'O{start_row}:O{end_row}').value
            c1c2_col_k = c1c2_sheet.range(f'K{start_row}:K{end_row}').value
            c1c2_col_l = c1c2_sheet.range(f'L{start_row}:L{end_row}').value
            c1c2_col_m = c1c2_sheet.range(f'M{start_row}:M{end_row}').value

            wb_prod.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

    # Convert to lists and sum
    c1c2_col_r_list = c1c2_col_r if isinstance(c1c2_col_r, list) else [c1c2_col_r]
    c1c2_col_ab_list = c1c2_col_ab if isinstance(c1c2_col_ab, list) else [c1c2_col_ab]
    c1c2_col_ad_list = c1c2_col_ad if isinstance(c1c2_col_ad, list) else [c1c2_col_ad]
    c1c2_col_o_list = c1c2_col_o if isinstance(c1c2_col_o, list) else [c1c2_col_o]
    c1c2_col_k_list = c1c2_col_k if isinstance(c1c2_col_k, list) else [c1c2_col_k]
    c1c2_col_l_list = c1c2_col_l if isinstance(c1c2_col_l, list) else [c1c2_col_l]
    c1c2_col_m_list = c1c2_col_m if isinstance(c1c2_col_m, list) else [c1c2_col_m]

    c1c2_totals = {
        'R': sum(pd.to_numeric(v, errors='coerce') or 0 for v in c1c2_col_r_list),
        'AB': sum(pd.to_numeric(v, errors='coerce') or 0 for v in c1c2_col_ab_list),
        'AD': sum(pd.to_numeric(v, errors='coerce') or 0 for v in c1c2_col_ad_list),
        'O': sum(pd.to_numeric(v, errors='coerce') or 0 for v in c1c2_col_o_list),
        'K': sum(pd.to_numeric(v, errors='coerce') or 0 for v in c1c2_col_k_list),
        'L': sum(pd.to_numeric(v, errors='coerce') or 0 for v in c1c2_col_l_list),
        'M': sum(pd.to_numeric(v, errors='coerce') or 0 for v in c1c2_col_m_list)
    }

    # Compare and report
    print(f"\n  [BULK2 VALIDATE] Comparison for {len(matching_contracts)} Bulk 2 contracts:")
    print("  [BULK2 VALIDATE] " + "="*80)
    print(f"  [BULK2 VALIDATE] {'Column':<25} {'Summary Total':>18} {'C1C2 Total':>18} {'Difference':>15} {'Status':>10}")
    print("  [BULK2 VALIDATE] " + "="*80)

    comparisons = [
        ('R_VALUE (AG->R)', summary_totals['AG'], c1c2_totals['R']),
        ('AB_VALUE (E->AB)', summary_totals['E'], c1c2_totals['AB']),
        ('AD_VALUE (R->AD)', summary_totals['R'], c1c2_totals['AD']),
        ('IIS_VALUE (T->O)', summary_totals['T'], c1c2_totals['O']),
        ('GROSS_PORT (Y->K)', summary_totals['Y'], c1c2_totals['K']),
        ('PROVISION (J->L)', summary_totals['J'], c1c2_totals['L']),
        ('DP_PROVISION (X->M)', summary_totals['X'], c1c2_totals['M'])
    ]

    all_match = True
    for col_name, summary_val, c1c2_val in comparisons:
        diff = abs(summary_val - c1c2_val)
        status = "✓ MATCH" if diff < 0.01 else "✗ MISMATCH"
        if diff >= 0.01:
            all_match = False

        print(f"  [BULK2 VALIDATE] {col_name:<25} {summary_val:>18,.2f} {c1c2_val:>18,.2f} {diff:>15,.2f} {status:>10}")

    print("  [BULK2 VALIDATE] " + "="*80)

    if all_match:
        print("  [BULK2 VALIDATE] ✓✓✓ ALL TOTALS MATCH! Data integrity confirmed.")
    else:
        print("  [BULK2 VALIDATE] ✗✗✗ MISMATCH DETECTED! Review data extraction and pasting logic.")

    print("  [BULK2 VALIDATE] ========== VALIDATION COMPLETE ==========\n")

    return all_match

def verify_bulk2_data_differences(summary_file, prod_file, bulk2_rows, df_bulk2, exception_tracker):
    """Verify and report differences between Summary and C1 & C2 Working sheet for Bulk 2 data"""
    print("\n  [BULK2 VERIFY] ========== VERIFYING BULK 2 DATA CONSISTENCY ==========")
    print("  [BULK2 VERIFY] Comparing Summary sheet vs C1 & C2 Working sheet...")

    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False

            # Read from Summary
            wb_summary = app.books.open(summary_file)
            summary_sheet = wb_summary.sheets['SUMMARY']
            last_row_summary = summary_sheet.used_range.last_cell.row

            # Read all relevant columns from Summary
            summary_col_a = summary_sheet.range(f'A1:A{last_row_summary}').value
            summary_col_e = summary_sheet.range(f'E1:E{last_row_summary}').value
            summary_col_r = summary_sheet.range(f'R1:R{last_row_summary}').value
            summary_col_t = summary_sheet.range(f'T1:T{last_row_summary}').value
            summary_col_ag = summary_sheet.range(f'AG1:AG{last_row_summary}').value

            wb_summary.close()

            # Read from C1 & C2 Working
            wb_prod = app.books.open(prod_file)
            c1c2_sheet = wb_prod.sheets['C1 & C2 Working']

            start_row = bulk2_rows[0]
            end_row = bulk2_rows[-1]

            # Read all relevant columns from C1 & C2 Working
            c1c2_col_a = c1c2_sheet.range(f'A{start_row}:A{end_row}').value
            c1c2_col_o = c1c2_sheet.range(f'O{start_row}:O{end_row}').value
            c1c2_col_r = c1c2_sheet.range(f'R{start_row}:R{end_row}').value
            c1c2_col_ab = c1c2_sheet.range(f'AB{start_row}:AB{end_row}').value
            c1c2_col_ad = c1c2_sheet.range(f'AD{start_row}:AD{end_row}').value

            wb_prod.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

    # Convert to lists
    summary_col_a = summary_col_a if isinstance(summary_col_a, list) else [summary_col_a]
    summary_col_e = summary_col_e if isinstance(summary_col_e, list) else [summary_col_e]
    summary_col_r = summary_col_r if isinstance(summary_col_r, list) else [summary_col_r]
    summary_col_t = summary_col_t if isinstance(summary_col_t, list) else [summary_col_t]
    summary_col_ag = summary_col_ag if isinstance(summary_col_ag, list) else [summary_col_ag]

    c1c2_col_a = c1c2_col_a if isinstance(c1c2_col_a, list) else [c1c2_col_a]
    c1c2_col_o = c1c2_col_o if isinstance(c1c2_col_o, list) else [c1c2_col_o]
    c1c2_col_r = c1c2_col_r if isinstance(c1c2_col_r, list) else [c1c2_col_r]
    c1c2_col_ab = c1c2_col_ab if isinstance(c1c2_col_ab, list) else [c1c2_col_ab]
    c1c2_col_ad = c1c2_col_ad if isinstance(c1c2_col_ad, list) else [c1c2_col_ad]

    # Build Summary lookup dictionary
    summary_dict = {}
    for i, contract in enumerate(summary_col_a):
        if contract and str(contract).strip():
            contract_clean = str(contract).strip().upper()
            summary_dict[contract_clean] = {
                'AB_VALUE': summary_col_e[i],
                'AD_VALUE': summary_col_r[i],
                'IIS_VALUE': summary_col_t[i],
                'R_VALUE': summary_col_ag[i]
            }

    # Compare data
    differences = []
    total_comparisons = 0
    mismatch_count = 0

    print(f"  [BULK2 VERIFY] Checking {len(c1c2_col_a)} contracts in Bulk 2 section...")
    print(f"  [BULK2 VERIFY] Row range in C1 & C2 Working: {start_row} to {end_row}")

    column_mappings = {
        'AB_VALUE': ('Summary Col E', 'C1C2 Col AB'),
        'AD_VALUE': ('Summary Col R', 'C1C2 Col AD'),
        'IIS_VALUE': ('Summary Col T', 'C1C2 Col O'),
        'R_VALUE': ('Summary Col AG', 'C1C2 Col R')
    }

    for i, contract in enumerate(c1c2_col_a):
        if not contract or not str(contract).strip():
            continue

        contract_clean = str(contract).strip().upper()

        if contract_clean not in summary_dict:
            differences.append({
                'Contract_No': contract,
                'Row_in_C1C2': start_row + i,
                'Issue': 'Contract not found in Summary',
                'Column': 'All',
                'Summary_Value': 'N/A',
                'C1C2_Value': 'N/A',
                'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            continue

        summary_vals = summary_dict[contract_clean]
        c1c2_vals = {
            'AB_VALUE': c1c2_col_ab[i],
            'AD_VALUE': c1c2_col_ad[i],
            'IIS_VALUE': c1c2_col_o[i],
            'R_VALUE': c1c2_col_r[i]
        }

        # Compare each column
        for col_name, (summary_col, c1c2_col) in column_mappings.items():
            total_comparisons += 1
            summary_val = summary_vals[col_name]
            c1c2_val = c1c2_vals[col_name]

            # Normalize values for comparison
            def normalize_value(val):
                if pd.isna(val) or val is None or val == '':
                    return ''
                return str(val).strip()

            summary_val_norm = normalize_value(summary_val)
            c1c2_val_norm = normalize_value(c1c2_val)

            if summary_val_norm != c1c2_val_norm:
                mismatch_count += 1
                differences.append({
                    'Contract_No': contract,
                    'Row_in_C1C2': start_row + i,
                    'Issue': 'Value Mismatch',
                    'Column': f'{summary_col} -> {c1c2_col}',
                    'Summary_Value': summary_val_norm,
                    'C1C2_Value': c1c2_val_norm,
                    'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })

    # Report findings
    print(f"\n  [BULK2 VERIFY] ========== VERIFICATION RESULTS ==========")
    print(f"  [BULK2 VERIFY] Total value comparisons: {total_comparisons}")
    print(f"  [BULK2 VERIFY] Matching values: {total_comparisons - mismatch_count}")
    print(f"  [BULK2 VERIFY] Mismatched values: {mismatch_count}")
    print(f"  [BULK2 VERIFY] Total issues found: {len(differences)}")

    if differences:
        print(f"  [BULK2 VERIFY] ⚠ WARNING: Found {len(differences)} differences!")
        print(f"  [BULK2 VERIFY] Sample differences (first 5):")
        for idx, diff in enumerate(differences[:5], 1):
            print(f"  [BULK2 VERIFY]   {idx}. Contract: {diff['Contract_No']}, Row: {diff['Row_in_C1C2']}")
            print(f"  [BULK2 VERIFY]      {diff['Column']}: Summary='{diff['Summary_Value']}' vs C1C2='{diff['C1C2_Value']}'")

        # Add to exception tracker
        for diff in differences:
            exception_tracker.add_exception('Bulk2_Data_Verification', diff)
    else:
        print(f"  [BULK2 VERIFY] ✓ All data matches between Summary and C1 & C2 Working!")

    print(f"  [BULK2 VERIFY] ========== VERIFICATION COMPLETE ==========\n")

    return differences

def build_bulk3_dataframe(net_file, unutilized_file, exception_tracker):
    """Build complete Bulk 3 DataFrame with Margin Trading data"""
    print("\n  [BULK3 DEBUG] ========== BULK 3 DATA BUILDING (MARGIN TRADING) ==========")

    # Read Net Portfolio for validation
    df_net = pd.read_excel(net_file, sheet_name=0, header=None)
    df_net.columns = ['Type'] + [f'Col_{i}' for i in range(1, len(df_net.columns))]

    margin_mask = df_net['Type'].astype(str).str.strip().str.upper() == 'MARGIN TRADING'

    if margin_mask.any():
        net_portfolio_value = df_net.loc[margin_mask, 'Col_25'].values[0]
        net_portfolio_value = float(net_portfolio_value) if pd.notna(net_portfolio_value) else 0
        print(f"  [BULK3 DEBUG] Net Portfolio MARGIN TRADING value: {net_portfolio_value:,.2f}")
    else:
        net_portfolio_value = 0
        print(f"  [BULK3 DEBUG] WARNING: No MARGIN TRADING row found in Net Portfolio")

    # Read Unutilized Amount file
    print(f"\n  [BULK3 DEBUG] Reading Unutilized Amount file...")
    print(f"  [BULK3 DEBUG] Source: {os.path.basename(unutilized_file)}")

    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(unutilized_file)

            # Use the first available sheet
            target_sheet = wb.sheets[0]

            print(f"  [BULK3 DEBUG] Using sheet: {target_sheet.name}")
            last_row = target_sheet.used_range.last_cell.row

            print(f"  [BULK3 DEBUG] Reading columns:")
            print(f"  [BULK3 DEBUG]   Unutilized Col C (ICAM Code)      -> CLIENT_CODE")
            print(f"  [BULK3 DEBUG]   Unutilized Col G (Closing Balance)-> GROSS_PORTFOLIO")
            print(f"  [BULK3 DEBUG]   Unutilized Col E (Interest Rate)  -> CON_INTRATE")

            icam_codes = target_sheet.range(f'C1:C{last_row}').value
            closing_balances = target_sheet.range(f'G1:G{last_row}').value
            interest_rates = target_sheet.range(f'E1:E{last_row}').value

            wb.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

    icam_codes = icam_codes if isinstance(icam_codes, list) else [icam_codes]
    closing_balances = closing_balances if isinstance(closing_balances, list) else [closing_balances]
    interest_rates = interest_rates if isinstance(interest_rates, list) else [interest_rates]

    df_unutilized = pd.DataFrame({
        'ICAM_CODE': icam_codes,
        'CLOSING_BALANCE': closing_balances,
        'INTEREST_RATE': interest_rates
    })

    df_unutilized = df_unutilized.iloc[:-1]

    print(f"  [BULK3 DEBUG] Total rows read: {len(df_unutilized)}")

    df_unutilized['CLOSING_BALANCE_NUM'] = pd.to_numeric(df_unutilized['CLOSING_BALANCE'], errors='coerce')
    
    print(f"unutilized data =+++++++++++++++++++++++++++++++++++++\n{df_unutilized.head()}")
    df_positive = df_unutilized[df_unutilized['CLOSING_BALANCE_NUM'] > 1].copy()

    print(f"  [BULK3 DEBUG] Rows with positive balances: {len(df_positive)}")

    df_positive = df_positive[~df_positive['ICAM_CODE'].astype(str).str.upper().str.contains('ICAM|CODE|TOTAL', na=False)]

    print(f"  [BULK3 DEBUG] Rows after filtering headers: {len(df_positive)}")

    unutilized_sum = df_positive['CLOSING_BALANCE_NUM'].sum()
    print(f"  [BULK3 DEBUG] Sum of Unutilized balances: {unutilized_sum:,.2f}")

    difference = abs(net_portfolio_value - unutilized_sum)
    tolerance = 0.01
    values_tally = difference < tolerance

    print(f"\n  [BULK3 DEBUG] ========== BALANCE VALIDATION ==========")
    print(f"  [BULK3 DEBUG] Net Portfolio (MARGIN TRADING): {net_portfolio_value:,.2f}")
    print(f"  [BULK3 DEBUG] Unutilized Amount (Sum):        {unutilized_sum:,.2f}")
    print(f"  [BULK3 DEBUG] Difference:                     {difference:,.2f}")
    print(f"  [BULK3 DEBUG] Tolerance:                      {tolerance:,.2f}")
    print(f"  [BULK3 DEBUG] Values Match: {'✓ YES' if values_tally else '✗ NO (MISMATCH)'}")

    if not values_tally:
        exception_tracker.add_exception('Bulk3_Margin_Trading', {
            'Check': 'Margin Trading Balance Comparison',
            'Net_Portfolio_Value': f"{net_portfolio_value:,.2f}",
            'Unutilized_Sum': f"{unutilized_sum:,.2f}",
            'Difference': f"{difference:,.2f}",
            'Status': 'MISMATCH',
            'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })

    df_bulk3 = df_positive[['ICAM_CODE', 'CLOSING_BALANCE_NUM', 'INTEREST_RATE']].copy()
    df_bulk3.columns = ['CLIENT_CODE', 'GROSS_PORTFOLIO', 'CON_INTRATE']

    df_bulk3['GROSS_PORTFOLIO'] = df_bulk3['GROSS_PORTFOLIO'].apply(
        lambda x: f"{float(x):,.0f}" if pd.notna(x) else ''
    )

    def safe_format_rate(x):
        """Safely format interest rate, handling invalid data"""
        if pd.notna(x) and x != '':
            try:
                return f"{float(x):.2f}%"
            except (ValueError, TypeError):
                print(f"  [BULK3 DEBUG] Warning: Invalid interest rate value: {x}")
                return ''
        return ''

    df_bulk3['CON_INTRATE'] = df_bulk3['CON_INTRATE'].apply(safe_format_rate)

    print("\n  [BULK3 DEBUG] ========== BULK 3 DATAFRAME COMPLETE ==========")
    print(f"  [BULK3 DEBUG] Final DataFrame columns: {list(df_bulk3.columns)}")
    print(f"  [BULK3 DEBUG] Total records for C1 & C2 Working: {len(df_bulk3)}")
    print(f"  [BULK3 DEBUG] Data will be written to C1 & C2 Working sheet:")
    print(f"  [BULK3 DEBUG]   DataFrame 'CLIENT_CODE'    -> C1C2 Col C  (ICAM Code)")
    print(f"  [BULK3 DEBUG]   DataFrame 'GROSS_PORTFOLIO'-> C1C2 Col K  (Closing Balance)")
    print(f"  [BULK3 DEBUG]   DataFrame 'CON_INTRATE'    -> C1C2 Col V  (Interest Rate)")

    return df_bulk3, values_tally

# ==================== SHEET OPERATIONS ====================

def process_moi_list(prod_file, reschedule_file):
    """Process MOI List sheet - add missing reschedule contracts"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            moi_sheet = wb.sheets['MOI List -31 Jan']

            last_row = moi_sheet.used_range.last_cell.row

            col_a = moi_sheet.range(f'A1:A{last_row}').value
            col_b = moi_sheet.range(f'B1:B{last_row}').value
            col_c = moi_sheet.range(f'C1:C{last_row}').value

            wb.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

    col_a_list = col_a if isinstance(col_a, list) else [col_a]
    col_b_list = col_b if isinstance(col_b, list) else [col_b]
    col_c_list = col_c if isinstance(col_c, list) else [col_c]

    df_moi = pd.DataFrame({'A': col_a_list, 'B': col_b_list, 'C': col_c_list})

    reschedule_mask = df_moi['C'].astype(str).str.strip() == 'Reschedule Total List'
    reschedule_indices = df_moi[reschedule_mask].index.tolist()

    if not reschedule_indices:
        return

    moi_reschedule_contracts = set(
        str(val).strip().upper() for val in df_moi.loc[reschedule_mask, 'A']
        if val and str(val).strip() and is_valid_contract_number(str(val))
    )

    df_reschedule = pd.read_excel(reschedule_file, sheet_name=0)

    if 'CLO_NEWCONNO' in df_reschedule.columns:
        reschedule_contracts = set(
            str(val).strip().upper() for val in df_reschedule['CLO_NEWCONNO']
            if pd.notna(val) and is_valid_contract_number(str(val))
        )
    else:
        reschedule_contracts = set(
            str(val).strip().upper() for val in df_reschedule.iloc[:, 1]
            if pd.notna(val) and is_valid_contract_number(str(val))
        )

    missing_contracts = reschedule_contracts - moi_reschedule_contracts

    if not missing_contracts:
        return

    last_reschedule_idx = reschedule_indices[-1]
    template_b = df_moi.loc[last_reschedule_idx, 'B']
    template_c = df_moi.loc[last_reschedule_idx, 'C']

    insert_position = last_reschedule_idx + 1

    new_rows = [{'A': contract, 'B': template_b, 'C': template_c}
                for contract in sorted(missing_contracts)]

    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            moi_sheet = wb.sheets['MOI List -31 Jan']

            if new_rows:
                moi_sheet.range(f'{insert_position + 1}:{insert_position + len(new_rows)}').insert(shift='down')

                for i, row_data in enumerate(new_rows):
                    row_num = insert_position + i + 1
                    moi_sheet.range(f'A{row_num}').value = row_data['A']
                    moi_sheet.range(f'B{row_num}').value = row_data['B']
                    moi_sheet.range(f'C{row_num}').value = row_data['C']

            wb.save()
            wb.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def process_yard_and_property(prod_file, yard_stock_file, cbsl_provision_file):
    """Process YardandProperty List sheet"""
    df_yard = pd.read_excel(yard_stock_file, sheet_name=0)
    yard_all_values = df_yard.iloc[:, 0].tolist()

    yard_contracts = [
        str(val).strip() for val in yard_all_values
        if is_valid_contract_number(str(val))
    ]

    df_property = pd.read_excel(cbsl_provision_file, sheet_name='PropertyMortgage')
    property_all_values = df_property.iloc[:, 0].tolist()

    property_contracts = [
        str(val).strip() for val in property_all_values
        if is_valid_contract_number(str(val))
    ]

    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            yard_sheet = wb.sheets['YardandProperty List']

            last_row = yard_sheet.used_range.last_cell.row
            if last_row > 1:
                yard_sheet.range(f'A2:A{last_row}').clear_contents()
                yard_sheet.range(f'D2:D{last_row}').clear_contents()

            current_header = yard_sheet.range('D1').value
            if current_header and 'June' in str(current_header):
                new_header = str(current_header).replace('June', 'July')
                yard_sheet.range('D1').value = new_header

            if yard_contracts:
                yard_array = [[c] for c in yard_contracts]
                yard_sheet.range(f'A2:A{1 + len(yard_contracts)}').value = yard_array

            if property_contracts:
                property_array = [[c] for c in property_contracts]
                yard_sheet.range(f'D2:D{1 + len(property_contracts)}').value = property_array

            wb.save()
            wb.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def process_staff_loan(prod_file, cadre_file):
    """Process Staff_Loan sheet"""
    max_retries = 3

    df_cadre = pd.read_excel(cadre_file, sheet_name=0, header=0)
    df_cadre.columns = df_cadre.columns.str.strip()

    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb_prod = app.books.open(prod_file)

            staff_sheet = wb_prod.sheets['Staff_Loan']

            header_row = None
            for row_num in range(1, 11):
                row_values = []
                for col_num in range(1, 20):
                    col_letter = get_column_letter(col_num)
                    val = staff_sheet.range(f'{col_letter}{row_num}').value
                    if val:
                        row_values.append(val)

                if len(row_values) >= len(df_cadre.columns) * 0.5:
                    header_row = row_num
                    break

            if not header_row:
                header_row = 1

            last_col = staff_sheet.used_range.last_cell.column
            existing_headers = []
            for col_num in range(1, last_col + 1):
                col_letter = get_column_letter(col_num)
                header_val = staff_sheet.range(f'{col_letter}{header_row}').value
                if header_val:
                    existing_headers.append((col_num, col_letter, str(header_val).strip()))

            last_row = staff_sheet.used_range.last_cell.row
            data_start_row = header_row + 1
            if last_row >= data_start_row:
                staff_sheet.range(f'A{data_start_row}:{get_column_letter(last_col)}{last_row}').clear_contents()

            for col_num, col_letter, header in existing_headers:
                matching_cadre_col = None

                if header in df_cadre.columns:
                    matching_cadre_col = header
                else:
                    for cadre_col in df_cadre.columns:
                        if cadre_col.upper() == header.upper():
                            matching_cadre_col = cadre_col
                            break

                if matching_cadre_col:
                    data = df_cadre[matching_cadre_col].tolist()
                    data_array = [[d] for d in data]
                    staff_sheet.range(f'{col_letter}{data_start_row}:{col_letter}{data_start_row + len(data) - 1}').value = data_array

            wb_prod.save()
            wb_prod.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def process_mgt_acc(summary_file, prod_file):
    """Process MGT ACC values from Summary to FORM1 Working and NBD-MF-23-C1"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False

            # Read from Summary
            wb_summary = app.books.open(summary_file)
            summary_sheet = wb_summary.sheets['SUMMARY']
            last_row = summary_sheet.used_range.last_cell.row

            col_w = summary_sheet.range(f'W1:W{last_row}').value
            col_x = summary_sheet.range(f'X1:X{last_row}').value
            col_y = summary_sheet.range(f'Y1:Y{last_row}').value

            col_w_list = col_w if isinstance(col_w, list) else [col_w]
            col_x_list = col_x if isinstance(col_x, list) else [col_x]
            col_y_list = col_y if isinstance(col_y, list) else [col_y]

            net_portfolio_x = None
            net_portfolio_y = None
            total_adjustments_x = None

            for i, val in enumerate(col_w_list):
                if val and 'NET PORTFOLIO MGT ACC' in str(val).upper():
                    net_portfolio_x = col_x_list[i]
                    net_portfolio_y = col_y_list[i]
                elif val and 'TOTAL ADJUSTMENTS' in str(val).upper():
                    total_adjustments_x = col_x_list[i]

            wb_summary.close()

            # Write to FORM1 Working and NBD-MF-23-C1
            wb_prod = app.books.open(prod_file)

            # FORM1 Working sheet
            form1_sheet = wb_prod.sheets['FORM1 Working']
            form1_last_row = form1_sheet.used_range.last_cell.row

            col_f = form1_sheet.range(f'F1:F{form1_last_row}').value
            col_f_list = col_f if isinstance(col_f, list) else [col_f]

            as_per_ma_row = None
            acc_diff_row = None

            for i, val in enumerate(col_f_list, start=1):
                if val and 'AS PER MA' in str(val).upper():
                    as_per_ma_row = i
                elif val and 'ACC. DIFFERENCE' in str(val).upper():
                    acc_diff_row = i

            if as_per_ma_row and net_portfolio_x is not None:
                form1_sheet.range(f'G{as_per_ma_row}').value = net_portfolio_x
                if net_portfolio_y is not None:
                    form1_sheet.range(f'J{as_per_ma_row}').value = net_portfolio_y

            if acc_diff_row and total_adjustments_x is not None:
                form1_sheet.range(f'G{acc_diff_row}').value = total_adjustments_x

            # NBD-MF-23-C1 sheet
            c1_sheet = wb_prod.sheets['NBD-MF-23-C1']
            c1_last_row = c1_sheet.used_range.last_cell.row

            col_b = c1_sheet.range(f'B1:B{c1_last_row}').value
            col_b_list = col_b if isinstance(col_b, list) else [col_b]

            adjustments_row = None
            for i, val in enumerate(col_b_list, start=1):
                if val and 'ADJUSTMENTS' in str(val).upper():
                    adjustments_row = i
                    break

            if adjustments_row:
                c1_sheet.range(f'C{adjustments_row}').formula = "='FORM1 Working'!G11/1000"

            wb_prod.save()
            wb_prod.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def process_sofp(sofp_file, prod_file):
    """Process SOFP value from NBD-MF-01-SOFP sheet to FORM1 Working"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False

            # Read from SOFP file
            wb_sofp = app.books.open(sofp_file)
            sofp_sheet = wb_sofp.sheets['NBD-MF-01-SOFP']
            last_row = sofp_sheet.used_range.last_cell.row

            col_a = sofp_sheet.range(f'A1:A{last_row}').value
            col_e = sofp_sheet.range(f'E1:E{last_row}').value

            col_a_list = col_a if isinstance(col_a, list) else [col_a]
            col_e_list = col_e if isinstance(col_e, list) else [col_e]

            sofp_value = None
            for i, val in enumerate(col_a_list):
                if val and str(val).strip() == '1.1.6.1.0.0':
                    sofp_value = col_e_list[i]
                    break

            wb_sofp.close()

            # Write to FORM1 Working
            wb_prod = app.books.open(prod_file)
            form1_sheet = wb_prod.sheets['FORM1 Working']

            if sofp_value is not None:
                form1_sheet.range('B34').value = sofp_value

            # Refresh pivot tables in FORM1 Working sheet before saving
            try:
                for pivot in form1_sheet.api.PivotTables():
                    if pivot.Name in ('PivotTable1', 'PivotTable2'):
                        pivot.RefreshTable()
                        print(f"Refreshed {pivot.Name} in FORM1 Working sheet")
            except Exception as pivot_error:
                print(f"Warning: Could not refresh pivot tables: {pivot_error}")

            wb_prod.save()
            wb_prod.close()
            app.quit()
            break

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

# ==================== BULK OPERATIONS ====================

def find_bulk_rows_in_prod(prod_file, pattern):
    """Find bulk section rows by pattern"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']

            last_row = sheet.used_range.last_cell.row
            compiled_pattern = re.compile(pattern, re.IGNORECASE)
            bulk_rows = []

            chunk_size = 10000
            for start in range(1, last_row + 1, chunk_size):
                end = min(start + chunk_size - 1, last_row)
                data = sheet.range(f'A{start}:A{end}').value
                data_list = data if isinstance(data, list) else [data]

                for i, cell in enumerate(data_list):
                    if cell and compiled_pattern.match(str(cell).strip()):
                        bulk_rows.append(start + i)

            wb.close()
            app.quit()
            return bulk_rows

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def adjust_rows_in_prod(sheet, bulk_rows, required_rows, data_columns):
    """Adjust rows and copy formulas/formats from row above"""
    if not bulk_rows:
        return []

    current_rows = len(bulk_rows)

    if current_rows == required_rows:
        return bulk_rows

    start_row, end_row = bulk_rows[0], bulk_rows[-1]

    if required_rows > current_rows:
        rows_to_add = required_rows - current_rows

        sheet.range(f'{end_row + 1}:{end_row + rows_to_add}').insert(shift='down')

        template_row = sheet.range(f'{end_row}:{end_row}')
        template_row.copy()

        new_rows_range = sheet.range(f'{end_row + 1}:{end_row + rows_to_add}')
        new_rows_range.paste()

        try:
            sheet.api.Application.CutCopyMode = False
        except:
            pass

        for col in data_columns:
            sheet.range(f'{col}{end_row + 1}:{col}{end_row + rows_to_add}').clear_contents()

        return list(range(start_row, end_row + rows_to_add + 1))

    else:
        rows_to_remove = current_rows - required_rows
        delete_start = end_row - rows_to_remove + 1
        sheet.range(f'{delete_start}:{end_row}').delete(shift='up')

        return list(range(start_row, delete_start))

def paste_bulk1_dataframe_to_prod(prod_file, bulk1_rows, df_bulk1):
    """Paste Bulk 1 DataFrame to prod"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            app.calculation = 'manual'

            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']

            data_columns = ['A', 'C', 'K', 'T', 'V', 'W', 'AF']

            required_rows = len(df_bulk1)
            adjusted_rows = adjust_rows_in_prod(sheet, bulk1_rows, required_rows, data_columns)

            if not adjusted_rows:
                wb.close()
                app.quit()
                return adjusted_rows

            start_row = adjusted_rows[0]
            end_row = start_row + required_rows - 1

            data_mapping = {
                'A': df_bulk1['CONTRACT_NO'].tolist(),
                'C': df_bulk1['CONTRACT_NO'].tolist(),
                'K': df_bulk1['GROSS_PORTFOLIO'].tolist(),
                'T': df_bulk1['CONTRACT_PERIOD'].tolist(),
                'V': df_bulk1['CON_INTRATE'].tolist(),
                'W': df_bulk1['CONTRACT_AMOUNT'].tolist(),
                'AF': df_bulk1['AF_VALUE'].tolist()
            }

            for col_letter, values in data_mapping.items():
                values_array = [[v] for v in values]
                sheet.range(f'{col_letter}{start_row}:{col_letter}{end_row}').value = values_array

            app.calculation = 'automatic'
            wb.save()
            wb.close()
            app.quit()

            return adjusted_rows

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def paste_bulk2_dataframe_to_prod(prod_file, bulk2_rows, df_bulk2):
    """Paste Bulk 2 DataFrame to prod"""
    print("\n  [BULK2 DEBUG] ========== PASTING BULK 2 TO PROD FILE ==========")
    print(f"  [BULK2 DEBUG] Target: {os.path.basename(prod_file)} ('C1 & C2 Working' sheet)")
    print(f"  [BULK2 DEBUG] Number of rows to paste: {len(df_bulk2)}")

    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            app.calculation = 'manual'

            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']

            data_columns = ['A', 'C', 'D', 'E', 'K', 'L', 'M', 'O', 'S', 'T', 'V', 'W', 'R', 'AB', 'AD', 'AF']

            required_rows = len(df_bulk2)
            adjusted_rows = adjust_rows_in_prod(sheet, bulk2_rows, required_rows, data_columns)

            if not adjusted_rows:
                wb.close()
                app.quit()
                return adjusted_rows

            start_row = adjusted_rows[0]
            end_row = start_row + required_rows - 1

            print(f"  [BULK2 DEBUG] Pasting to rows {start_row} to {end_row}")
            print("\n  [BULK2 DEBUG] DESTINATION COLUMN MAPPING:")
            print("  [BULK2 DEBUG]   DataFrame 'CONTRACT_NO'    -> Prod Col A  (Contract Number)")
            print("  [BULK2 DEBUG]   DataFrame 'CLIENT_CODE'    -> Prod Col C  (Client Code) [FROM SUMMARY COL B]")
            print("  [BULK2 DEBUG]   DataFrame 'EQT_DESC'       -> Prod Col D  (Equipment Description)")
            print("  [BULK2 DEBUG]   DataFrame 'PURPOSE'        -> Prod Col E  (Purpose)")
            print("  [BULK2 DEBUG]   DataFrame 'GROSS_PORTFOLIO'-> Prod Col K  (Gross Portfolio) [FROM SUMMARY COL Y]")
            print("  [BULK2 DEBUG]   DataFrame 'PROVISION'      -> Prod Col L  (Provision) [FROM SUMMARY COL J]")
            print("  [BULK2 DEBUG]   DataFrame 'DP_PROVISION'   -> Prod Col M  (DP Provision) [FROM SUMMARY COL X]")
            print("  [BULK2 DEBUG]   DataFrame 'IIS_VALUE'      -> Prod Col O  (IIS from Summary)")
            print("  [BULK2 DEBUG]   DataFrame 'CON_RNTFREQ'    -> Prod Col S  (Rental Frequency)")
            print("  [BULK2 DEBUG]   DataFrame 'CONTRACT_PERIOD'-> Prod Col T  (Contract Period)")
            print("  [BULK2 DEBUG]   DataFrame 'CON_INTRATE'    -> Prod Col V  (Interest Rate)")
            print("  [BULK2 DEBUG]   DataFrame 'CONTRACT_AMOUNT'-> Prod Col W  (Contract Amount)")
            print("  [BULK2 DEBUG]   DataFrame 'R_VALUE'        -> Prod Col R  (R Value from Summary)")
            print("  [BULK2 DEBUG]   DataFrame 'AB_VALUE'       -> Prod Col AB (AB Value from Summary)")
            print("  [BULK2 DEBUG]   DataFrame 'AD_VALUE'       -> Prod Col AD (AD Value from Summary)")
            print("  [BULK2 DEBUG]   DataFrame 'AF_VALUE'       -> Prod Col AF (Initial Valuation)")

            # Prepare data - keep numeric values as-is, format text
            def prepare_value(val):
                """Convert value for Excel - preserve numeric types"""
                if pd.isna(val) or val is None or val == '':
                    return ''
                if isinstance(val, str):
                    # Remove formatting from strings that represent numbers
                    cleaned = val.replace(',', '').replace('%', '').strip()
                    try:
                        return float(cleaned)
                    except:
                        return val
                return val

            # Prepare R_VALUE (Stage) - use values directly (VLOOKUP already done in dataframe via read_multiple_columns_from_file)
            r_values = [prepare_value(v) for v in df_bulk2['R_VALUE'].tolist()]

            # Prepare CON_INTRATE (Interest Rate from Net Portfolio) - divide by 100 since final output multiplies by 100
            v_values = []
            for v in df_bulk2['CON_INTRATE'].tolist():
                prep_val = prepare_value(v)
                if prep_val != '' and prep_val is not None:
                    try:
                        v_values.append(float(prep_val) / 100)
                    except:
                        v_values.append(prep_val)
                else:
                    v_values.append('')

            data_mapping = {
                'A': df_bulk2['CONTRACT_NO'].tolist(),
                'C': df_bulk2['CLIENT_CODE'].tolist(),
                'D': df_bulk2['EQT_DESC'].tolist(),
                'E': df_bulk2['PURPOSE'].tolist(),
                'K': [prepare_value(v) for v in df_bulk2['GROSS_PORTFOLIO'].tolist()],
                'L': [prepare_value(v) for v in df_bulk2['PROVISION'].tolist()],
                'M': [prepare_value(v) for v in df_bulk2['DP_PROVISION'].tolist()],
                'O': [prepare_value(v) for v in df_bulk2['IIS_VALUE'].tolist()],
                'S': df_bulk2['CON_RNTFREQ'].tolist(),
                'T': df_bulk2['CONTRACT_PERIOD'].tolist(),
                'V': v_values,
                'W': [prepare_value(v) for v in df_bulk2['CONTRACT_AMOUNT'].tolist()],
                'R': r_values,
                'AB': [prepare_value(v) for v in df_bulk2['AB_VALUE'].tolist()],
                'AD': [prepare_value(v) for v in df_bulk2['AD_VALUE'].tolist()],
                'AF': [prepare_value(v) for v in df_bulk2['AF_VALUE'].tolist()]
            }

            # Paste data column by column
            for col_letter, values in data_mapping.items():
                values_array = [[v] for v in values]
                sheet.range(f'{col_letter}{start_row}:{col_letter}{end_row}').value = values_array

            print(f"  [BULK2 DEBUG] Successfully pasted {len(data_mapping)} columns to Prod file")

            # Validate sums after pasting
            print("\n  [BULK2 DEBUG] ========== POST-PASTE VALIDATION ==========")

            # Read back and validate numeric columns
            validation_cols = {
                'R': ('R_VALUE', df_bulk2['R_VALUE_NUMERIC']),
                'AB': ('AB_VALUE', df_bulk2['AB_VALUE_NUMERIC']),
                'AD': ('AD_VALUE', df_bulk2['AD_VALUE_NUMERIC']),
                'O': ('IIS_VALUE', df_bulk2['IIS_VALUE_NUMERIC']),
                'K': ('GROSS_PORTFOLIO', df_bulk2['GROSS_PORTFOLIO_NUMERIC']),
                'L': ('PROVISION', df_bulk2['PROVISION_NUMERIC']),
                'M': ('DP_PROVISION', df_bulk2['DP_PROVISION_NUMERIC'])
            }

            for col_letter, (col_name, expected_series) in validation_cols.items():
                pasted_data = sheet.range(f'{col_letter}{start_row}:{col_letter}{end_row}').value
                pasted_list = pasted_data if isinstance(pasted_data, list) else [pasted_data]
                pasted_numeric = pd.to_numeric(pd.Series(pasted_list), errors='coerce')

                expected_sum = expected_series.sum()
                actual_sum = pasted_numeric.sum()
                difference = abs(expected_sum - actual_sum)

                status = "✓ MATCH" if difference < 0.01 else "✗ MISMATCH"
                print(f"  [BULK2 DEBUG] Col {col_letter:2} ({col_name:18}): Expected: {expected_sum:15,.2f} | Actual: {actual_sum:15,.2f} | Diff: {difference:10,.2f} {status}")

            app.calculation = 'automatic'
            wb.save()
            wb.close()
            app.quit()

            print("  [BULK2 DEBUG] ========== BULK 2 PASTE COMPLETE ==========\n")
            return adjusted_rows

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def paste_bulk3_dataframe_to_prod(prod_file, bulk3_rows, df_bulk3):
    """Paste Bulk 3 DataFrame to prod"""
    print("\n  [BULK3 DEBUG] ========== PASTING BULK 3 TO PROD FILE ==========")
    print(f"  [BULK3 DEBUG] Target: {os.path.basename(prod_file)} ('C1 & C2 Working' sheet)")
    print(f"  [BULK3 DEBUG] Number of rows to paste: {len(df_bulk3)}")

    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            app.calculation = 'manual'

            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']

            data_columns = ['C', 'K', 'V']

            required_rows = len(df_bulk3)
            adjusted_rows = adjust_rows_in_prod(sheet, bulk3_rows, required_rows, data_columns)

            if not adjusted_rows:
                wb.close()
                app.quit()
                return adjusted_rows

            start_row = adjusted_rows[0]
            end_row = start_row + required_rows - 1

            print(f"  [BULK3 DEBUG] Pasting to rows {start_row} to {end_row}")
            print("\n  [BULK3 DEBUG] DESTINATION COLUMN MAPPING:")
            print("  [BULK3 DEBUG]   DataFrame 'CLIENT_CODE'    -> Prod Col C  (ICAM Code from Unutilized)")
            print("  [BULK3 DEBUG]   DataFrame 'GROSS_PORTFOLIO'-> Prod Col K  (Closing Balance from Unutilized)")
            print("  [BULK3 DEBUG]   DataFrame 'CON_INTRATE'    -> Prod Col V  (Interest Rate from Unutilized)")

            # Prepare value helper function
            def prepare_value(val):
                """Convert value for Excel - preserve numeric types"""
                if pd.isna(val) or val is None or val == '':
                    return ''
                if isinstance(val, str):
                    # Remove formatting from strings that represent numbers
                    cleaned = val.replace(',', '').replace('%', '').strip()
                    try:
                        return float(cleaned)
                    except:
                        return val
                return val

            # Prepare CON_INTRATE (Interest Rate) - divide by 100 since final output multiplies by 100
            v_values = []
            for v in df_bulk3['CON_INTRATE'].tolist():
                prep_val = prepare_value(v)
                if prep_val != '' and prep_val is not None:
                    try:
                        v_values.append(float(prep_val) / 100)
                    except:
                        v_values.append(prep_val)
                else:
                    v_values.append('')

            data_mapping = {
                'C': df_bulk3['CLIENT_CODE'].tolist(),
                'K': df_bulk3['GROSS_PORTFOLIO'].tolist(),
                'V': v_values
            }

            for col_letter, values in data_mapping.items():
                values_array = [[v] for v in values]
                sheet.range(f'{col_letter}{start_row}:{col_letter}{end_row}').value = values_array

            print(f"  [BULK3 DEBUG] Successfully pasted {len(data_mapping)} columns to Prod file")

            app.calculation = 'automatic'
            wb.save()
            wb.close()
            app.quit()

            print("  [BULK3 DEBUG] ========== BULK 3 PASTE COMPLETE ==========\n")
            return adjusted_rows

        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

# ==================== SCIENTER BOT FUNCTIONS ====================

def load_scienter_bot():
    """Load Scienter valuation bot with error handling."""
    try:
        print("  Loading IA_Working_Initial_valuation_Scienter_bot module...")
        import sys
        import os

        # Add the parent "bots" folder relative to this file
        current_dir = os.path.dirname(__file__)
        bot_dir = os.path.join(current_dir, "..", "bots")
        sys.path.insert(0, os.path.abspath(bot_dir))

        import IA_Working_Initial_valuation_Scienter_bot as scienter_bot
        print("  Successfully imported IA_Working_Initial_valuation_Scienter_bot module")
        return scienter_bot
    except ImportError as e:
        print(f"  Could not import IA_Working_Initial_valuation_Scienter_bot: {e}")
        return None

def fetch_lr_valuations(contracts_list, scienter_bot):
    """
    Fetch initial valuations for LR-type contracts using Scienter bot.

    Args:
        contracts_list: List of LR contract numbers
        scienter_bot: The imported Scienter bot module

    Returns:
        dict: Dictionary mapping contract numbers to their valuations
    """
    valuations = {}

    if not scienter_bot:
        print("  Scienter bot not available, cannot fetch LR valuations")
        return valuations

    if not contracts_list:
        print("  No LR contracts to process")
        return valuations

    try:
        print(f"  Fetching valuations for {len(contracts_list)} LR contracts via Scienter bot")

        # Call the Scienter bot's valuation function
        response = scienter_bot.run_scienter_valuation_bot(contracts_list)

        if isinstance(response, dict):
            for contract, value in response.items():
                # Validate the value from bot
                if value and value not in ["", "Not Valued", "#N/A", "N/A"]:
                    try:
                        numeric_value = float(value)
                        if numeric_value > 0:
                            valuations[contract] = numeric_value
                            print(f"  ✓ {contract}: {numeric_value:,.0f}")
                    except (ValueError, TypeError):
                        print(f"  ✗ {contract}: Invalid value '{value}'")

        print(f"  Successfully fetched {len(valuations)} valuations from Scienter bot")
        return valuations

    except Exception as e:
        print(f"  Error fetching valuations from Scienter bot: {e}")
        return valuations

# ==================== MAIN EXECUTION ====================

def main():
    script_start = time.time()

    print("Checking for existing Excel processes...")
    kill_excel_processes()

    print("="*60)
    print("EXCEL AUTOMATION - COMPLETE")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)

    folder = find_dynamic_folder()
    if not folder:
        return

    exception_tracker = ExceptionTracker(folder)

    (summary_file, net_file, prod_file, loan_base_file, reschedule_file,
     yard_stock_file, cbsl_provision_file, cadre_file, unutilized_file,
     po_listing_file, portfolio_recovery_file, sofp_file, portfoliorec_file) = find_files(folder)

    missing = []
    if not summary_file: missing.append("Summary")
    if not net_file: missing.append("Net Portfolio")
    if not prod_file: missing.append("Prod. wise Class. of Loans")
    if not loan_base_file: missing.append("Loan Base")
    if not reschedule_file: missing.append("Reschedule Contract Details")
    if not yard_stock_file: missing.append("YARD STOCK")
    if not cbsl_provision_file: missing.append("CBSL Provision Comparison")
    if not cadre_file: missing.append("Cadre")
    if not unutilized_file: missing.append("Unutilized Amount")
    if not po_listing_file: missing.append("Po Listing - Internal")
    if not portfolio_recovery_file: missing.append("Portfolio Report Recovery - Internal")
    if not sofp_file: missing.append("NBD-MF-01-SOFP & SOCI AFL Monthly FS")
    # PORTFOLIOREC is optional

    if missing:
        print(f"\nERROR: Missing files: {', '.join(missing)}")
        return

    print(f"\nAll files found")
    if portfoliorec_file:
        print(f"  PORTFOLIOREC file: {os.path.basename(portfoliorec_file)}")
    else:
        print(f"  PORTFOLIOREC file: Not found (optional)")

    # Load old Initial Valuation data from C1C2 Working and IA Working sheets
    print("\n" + "="*60)
    print("LOADING OLD INITIAL VALUATION DATA")
    print("="*60)
    step = time.time()
    df_old_c1c2, df_old_ia = load_old_initial_valuations(prod_file)
    log_step("Loading old valuations", step)

    # Step 1: MOI List
    print("\n" + "="*60)
    print("STEP 1: MOI LIST")
    print("="*60)
    step = time.time()
    process_moi_list(prod_file, reschedule_file)
    log_step("Step 1", step)

    # Step 2: Yard & Property
    print("\n" + "="*60)
    print("STEP 2: YARD AND PROPERTY")
    print("="*60)
    step = time.time()
    process_yard_and_property(prod_file, yard_stock_file, cbsl_provision_file)
    log_step("Step 2", step)

    # Step 3: Staff Loan
    print("\n" + "="*60)
    print("STEP 3: STAFF LOAN")
    print("="*60)
    step = time.time()
    process_staff_loan(prod_file, cadre_file)
    log_step("Step 3", step)

    # Step 4-7: Bulk 1
    print("\n" + "="*60)
    print("STEP 4-7: BULK 1 (LR Pattern)")
    print("="*60)
    step = time.time()
    net_contracts = read_contracts_fast(net_file, 0, 'E')
    net_bulk1 = classify_contracts(net_contracts, r'^LR\d{8}$')
    print(f"  Bulk 1: {len(net_bulk1)} contracts")
    log_step("Step 4-5", step)

    step = time.time()
    loan_base_sheet = get_first_sheet_name(loan_base_file)
    df_bulk1 = build_bulk1_dataframe(net_file, loan_base_file, loan_base_sheet, net_bulk1)
    bulk1_rows = find_bulk_rows_in_prod(prod_file, r'^LR\d{8}$')
    adjusted_bulk1_rows = paste_bulk1_dataframe_to_prod(prod_file, bulk1_rows, df_bulk1)
    log_step("Step 6-7", step)

    # Step 8-12: Bulk 2
    print("\n" + "="*60)
    print("STEP 8-12: BULK 2 (4-Letter + 9-Digit Pattern)")
    print("="*60)
    step = time.time()
    summary_contracts = read_contracts_fast(summary_file, 'SUMMARY', 'A')
    summary_bulk2 = classify_contracts(summary_contracts, r'^[A-Z]{4}\d{9}$')
    net_bulk2 = classify_contracts(net_contracts, r'^[A-Z]{4}\d{9}$')
    matches = set(summary_bulk2) & set(net_bulk2)
    print(f"  Bulk 2: {len(matches)} contracts")
    log_step("Step 8-9", step)

    step = time.time()
    df_bulk2 = build_bulk2_dataframe(summary_file, net_file, po_listing_file,
                                      portfolio_recovery_file, prod_file, portfoliorec_file,
                                      df_old_c1c2, df_old_ia,
                                      list(matches), exception_tracker)
    bulk2_rows = find_bulk_rows_in_prod(prod_file, r'^[A-Z]{4}\d{9}$')
    adjusted_bulk2_rows = paste_bulk2_dataframe_to_prod(prod_file, bulk2_rows, df_bulk2)
    log_step("Step 10-12", step)

    # Validate Bulk 2 totals
    print("\n" + "="*60)
    print("STEP 12.1: VALIDATE BULK 2 TOTALS")
    print("="*60)
    step = time.time()
    validate_bulk2_summary_totals(summary_file, prod_file, adjusted_bulk2_rows, df_bulk2, list(matches))
    log_step("Step 12.1", step)

    # Verify Bulk 2 row-by-row data
    print("\n" + "="*60)
    print("STEP 12.2: VERIFY BULK 2 DATA ROW-BY-ROW")
    print("="*60)
    step = time.time()
    verify_bulk2_data_differences(summary_file, prod_file, adjusted_bulk2_rows, df_bulk2, exception_tracker)
    log_step("Step 12.2", step)

    # Step 13-15: Bulk 3
    print("\n" + "="*60)
    print("STEP 13-15: BULK 3 (MARGIN TRADING)")
    print("="*60)
    step = time.time()
    df_bulk3, values_tally = build_bulk3_dataframe(net_file, unutilized_file, exception_tracker)
    print(f"  Bulk 3: {len(df_bulk3)} entries")
    bulk3_rows = find_bulk_rows_in_prod(prod_file, r'^MARGIN\s+TRADING$')
    adjusted_bulk3_rows = paste_bulk3_dataframe_to_prod(prod_file, bulk3_rows, df_bulk3)
    log_step("Step 13-15", step)

    # Step 16: MGT ACC
    print("\n" + "="*60)
    print("STEP 16: MGT ACC PROCESSING")
    print("="*60)
    step = time.time()
    process_mgt_acc(summary_file, prod_file)
    log_step("Step 16", step)

    # Step 17: SOFP
    print("\n" + "="*60)
    print("STEP 17: SOFP PROCESSING")
    print("="*60)
    step = time.time()
    process_sofp(sofp_file, prod_file)
    log_step("Step 17", step)

    # Write exceptions
    print("\n" + "="*60)
    print("WRITING EXCEPTIONS")
    print("="*60)
    exception_tracker.write_exceptions()

    # Summary
    total_time = time.time() - script_start
    print("\n" + "="*60)
    print("COMPLETED SUCCESSFULLY")
    print("="*60)
    print(f"\nBulk 1: {len(net_bulk1)} | Bulk 2: {len(matches)} | Bulk 3: {len(df_bulk3)}")
    print(f"Total time: {total_time:.2f}s ({total_time/60:.2f} min)")
    print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)

if __name__ == "__main__":
    main()
