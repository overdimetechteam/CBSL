import pandas as pd
import numpy as np
import xlwings as xw
import logging
import glob
import os
import re
from datetime import datetime
from pathlib import Path

import sys, os
# Add parent folder to Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from bots.na_contract_numbers_search_bot_api import (
    get_authenticated_session,
    process_contract_with_retry,
)

def setup_logging():
    """Setup organized logging for NBD-MF-23-C4 report with both file and console output"""
    # Get logger
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    logger.propagate = False  # Prevent propagation to root logger to avoid duplicates

    # Clear any existing handlers to avoid duplicates
    logger.handlers.clear()

    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Console handler (for app.py subprocess to capture)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Create organized log directory: report_automations/logs/
    try:
        run_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        # Get script directory (c:\CBSL\Script\report_automations)
        script_dir = Path(__file__).resolve().parent

        # Create log directory structure: report_automations/logs/
        logs_dir = script_dir / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        run_log_file = logs_dir / f"NBD_MF_23_C4_{run_timestamp}.log"

        # File handler
        file_handler = logging.FileHandler(run_log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        logger.info("="*80)
        logger.info(f"NBD-MF-23-C4 Report Started at {datetime.now()}")
        logger.info(f"Log file: {run_log_file}")
        logger.info("="*80)

        return logger
    except Exception as e:
        logger.error(f"Failed to setup file logging: {e}")
        return logger

logger = setup_logging()

def get_working_directory():
    """Find the working directory dynamically: working/NBD_MF_23_C4/<date>"""
    # Get script's parent directory (c:\CBSL\Script)
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    report_path = os.path.join(script_dir, "working", "NBD_MF_23_C4")

    if not os.path.exists(report_path):
        raise FileNotFoundError(f"Report path not found: {report_path}")

    # Get the single subdirectory inside NBD_MF_23_C4 folder (the date folder)
    subdirs = [d for d in os.listdir(report_path) if os.path.isdir(os.path.join(report_path, d))]
    if not subdirs:
        raise FileNotFoundError(f"No date subdirectory found in {report_path}")

    # Use the first (and should be only) subdirectory
    date_folder = subdirs[0]
    working_dir = os.path.join(report_path, date_folder)

    if not os.path.exists(working_dir):
        raise FileNotFoundError(f"Date folder not found at: {working_dir}")

    logger.info(f"Working directory found: {working_dir}")
    return working_dir

def find_file(pattern, working_dir=None):
    """Find file by pattern in working directory or Input subdirectory"""
    search_paths = []

    # If working directory provided, search there first
    if working_dir:
        search_paths.append(os.path.join(working_dir, pattern))
        search_paths.append(os.path.join(working_dir, "Input", pattern))
        search_paths.append(os.path.join(working_dir, f"*{pattern}*"))
        search_paths.append(os.path.join(working_dir, "Input", f"*{pattern}*"))

    # Also try current directory
    search_paths.extend([
        pattern,
        os.path.join("Input", pattern),
        f"*{pattern}*",
        os.path.join("Input", f"*{pattern}*")
    ])

    for path in search_paths:
        matches = glob.glob(path)
        if matches:
            return matches[0]

    raise FileNotFoundError(f"Could not find file matching: {pattern}")

def find_last_two_files(file_pattern, file_type_name, working_dir=None):
    """
    Find the last two files matching the pattern by date in filename.
    Returns a tuple: (last_month_file, current_month_file)
    """
    # Search in working directory and Input folder
    search_paths = []
    if working_dir:
        search_paths.append(os.path.join(working_dir, file_pattern))
        search_paths.append(os.path.join(working_dir, "Input", file_pattern))

    # Also search current folder and Input folder
    search_paths.append(file_pattern)
    search_paths.append(os.path.join("Input", file_pattern))

    files = []
    for path in search_paths:
        files.extend(glob.glob(path))

    if not files:
        raise FileNotFoundError(f"No {file_type_name} files found in the directory or Input folder.")

    file_dates = []
    for f in files:
        basename = os.path.basename(f)
        # Match date like 2025-06-30 or 2025_06_30
        match = re.search(r'(\d{4}[-_]\d{2}[-_]\d{2})', basename)
        if match:
            date_str = match.group(1).replace("_", "-")
            try:
                file_dates.append((f, datetime.strptime(date_str, "%Y-%m-%d")))
            except ValueError:
                continue  # skip if not valid date

    if len(file_dates) < 2:
        raise FileNotFoundError(f"Not enough {file_type_name} files with valid dates to determine last and current month.")

    # Sort by date
    file_dates.sort(key=lambda x: x[1])

    last_month_file = file_dates[-1][0]
    current_month_file = file_dates[0][0]

    return current_month_file, last_month_file

# Get the working directory dynamically
working_dir = get_working_directory()

CE = pd.read_excel(find_file("CE*.xlsx", working_dir), sheet_name='Sheet1')
logger.info("CE Report Loaded")

# Find Portfolio files automatically by date
portfolio_last_month_file, portfolio_current_month_file = find_last_two_files("*Portfolio*.xlsb", "Portfolio", working_dir)
logger.info(f"Portfolio Last Month File: {portfolio_last_month_file}")
logger.info(f"Portfolio Current Month File: {portfolio_current_month_file}")

PortfolioLastMonth = pd.read_excel(portfolio_last_month_file,
                                    sheet_name='Portfolio',
                                    engine='pyxlsb',
                                    skiprows=2)
logger.info("Portfolio Last Month Loaded")
PortfoliocurrentMonth = pd.read_excel(portfolio_current_month_file,
                                    sheet_name='Portfolio',
                                    engine='pyxlsb',
                                    skiprows=2)
logger.info("Portfolio Current Month Loaded")

# Find Summary files automatically by date
summary_last_month_file, summary_current_month_file = find_last_two_files("*Summary*.xlsb", "Summary", working_dir)
logger.info(f"Summary Last Month File: {summary_last_month_file}")
logger.info(f"Summary Current Month File: {summary_current_month_file}")

SummaryLastMonth = pd.read_excel(summary_last_month_file,
                                    sheet_name='SUMMARY',
                                    engine='pyxlsb',
                                    skiprows=2)
logger.info("Summary Last Month Loaded")
SummaryCurrentMonth = pd.read_excel(summary_current_month_file,
                                    sheet_name='SUMMARY',
                                    engine='pyxlsb',
                                    skiprows=2)
logger.info("Summary Current Month Loaded")
try:
    productCategory = find_file("Product_Cat.xlsx", working_dir)
    logger.info(f"Loading AT RECOVERY file: {productCategory}")
    productCategory= pd.read_excel(productCategory, engine='openpyxl')
    logger.info(f"Atrecovery Loaded - Shape: {productCategory.shape}")
except Exception as e:
    logger.error(f"Failed to load AT RECOVERY file: {e}")
    raise
# combine all five columns as lookup column

def concat_non_blank(row):
    fields = ['Product', 'Equipment', 'Purpose', 'Corporate Clients']
    return ''.join(str(row[field]) for field in fields if pd.notna(row[field]) and row[field] != '')

productCategory['LOOKUP'] = productCategory.apply(concat_non_blank, axis=1)
try:
    # Load the Excel file using pyxlsb engine for .xlsb format
    file_path = find_file("Loan CF Analysis*.xlsb", working_dir)
    #productCategory = pd.read_excel(file_path, sheet_name='Product_Cat', engine='pyxlsb')

    # Create a new Excel instance (or connect to an existing one)
    app = xw.App(visible=False)  # visible=True for debugging, set to False in production
    wb = app.books.open(file_path)

    # Access the 'Product_Cat' sheet
    sheet = wb.sheets['Product_Cat']

    # Write DataFrame to the 'Product_Cat' sheet starting from cell A1
    # Write headers and data from the DataFrame
    sheet.range('A1').value = productCategory.columns.tolist()  # Write headers
    sheet.range('A2').value = productCategory.values.tolist()  # Write data

    # Save the workbook
    wb.save()
    wb.close()

    # Close the Excel application (optional)
    app.quit()

    logger.info("Data written successfully to Product_Cat sheet.")
except Exception as e:
    logger.error(f"Error occurred while writing to the Excel file: {e}")


productCategory = pd.read_excel(find_file("Loan CF Analysis*.xlsb", working_dir), sheet_name='Product_Cat', engine='pyxlsb')
logger.info("Product Category Loaded")
LoanCategory = pd.read_excel(find_file("Loan CF Analysis*.xlsb", working_dir), sheet_name='Mor. Loan Category', engine='pyxlsb')
logger.info("Loan Category Loaded")

try:
    atrecovery_file = find_file("AT RECOVERY*.xlsx", working_dir)
    logger.info(f"Loading AT RECOVERY file: {atrecovery_file}")
    Atrecovery = pd.read_excel(atrecovery_file, sheet_name='RECOVERY COLLECTION EFFICIENCY ', engine='openpyxl')
    logger.info(f"Atrecovery Loaded - Shape: {Atrecovery.shape}")
except Exception as e:
    logger.error(f"Failed to load AT RECOVERY file: {e}")
    raise

try:
    setoff_file = find_file("Rec Target*.xlsx", working_dir)
    logger.info(f"Loading Expected Setoff file: {setoff_file}")
    expectedSetoff = pd.read_excel(setoff_file, engine='openpyxl')
    logger.info(f"Expected Setoff Loaded - Shape: {expectedSetoff.shape}")
except Exception as e:
    logger.error(f"Failed to load Expected Setoff file: {e}")
    raise
logger.info(f"Expected Setoff columns: {list(expectedSetoff.columns)}")


# Add source identifier
PortfolioLastMonth_copy = PortfolioLastMonth.copy()
PortfoliocurrentMonth_copy = PortfoliocurrentMonth.copy()

# Concatenate both DataFrames
merged_df_pf = pd.concat([PortfolioLastMonth_copy, PortfoliocurrentMonth_copy], ignore_index=True)
logger.info("Merged portfolio sample:\n%s", merged_df_pf.head())
# Remove duplicates, keeping the last occurrence (current month data)
merged_df_pf = merged_df_pf.drop_duplicates(subset=['CONTRACT NO'], keep='last')
logger.info("Merged portfolio sample:\n%s", merged_df_pf.head())
Report = pd.DataFrame()

CE_without_NA = CE[~CE['CON_NO'].isna()]

logger.info("CE (non-NA) sample:\n%s", CE_without_NA.head())

Report['CONTRACT NO'] = CE_without_NA['CON_NO']

Report['MONTH DUE (RENTAL)'] = CE_without_NA['DUE_RNT_AMOUNT'].fillna(0)
Report['MONTH DUE (OTHCHG)'] = CE_without_NA['DUE_OTH_AMOUNT'].fillna(0)
Report['TOTAL MONTH DUE'] = CE_without_NA['Due new'].fillna(0)
Report['TOTAL SETOFF'] = CE_without_NA['Setoff new'].fillna(0)

Report["PRODUCT"] = Report["CONTRACT NO"].apply(lambda x: "FDL" if x[:2] == "LR" else x[2:4])

# Step 1: Filter DF1 for "AT" product type rows
# Create a boolean mask for AT products
at_mask = Report['PRODUCT'] == 'AT'

# Perform merge for AT rows only
at_merged = Report[at_mask].merge(
    Atrecovery[['STR_TRCO_CON_NO', 'NUM_TRCO_DUE_RNT_AMOUNT', 'NUM_TRCO_DUE_OTH_AMOUNT', 'Due', 'Settle']], 
    left_on='CONTRACT NO', 
    right_on='STR_TRCO_CON_NO', 
    how='left'
)
    
# Reset index to ensure proper alignment
at_merged = at_merged.reset_index(drop=True)

# Assign values ignoring the original indices
Report.loc[at_mask, 'MONTH DUE (RENTAL)'] = at_merged['NUM_TRCO_DUE_RNT_AMOUNT'].fillna(0).values
Report.loc[at_mask, 'MONTH DUE (OTHCHG)'] = at_merged['NUM_TRCO_DUE_OTH_AMOUNT'].fillna(0).values
Report.loc[at_mask, 'AT SETOFF'] = at_merged['Settle'].fillna(0).values

Report['AT SETOFF'] = Report['AT SETOFF'].fillna(0)

contract_to_client = dict(zip(merged_df_pf["CONTRACT NO"], merged_df_pf["CLIENT CODE"]))
contract_to_quip = dict(zip(merged_df_pf["CONTRACT NO"], merged_df_pf["EQT_DESC"]))
contract_to_purpose = dict(zip(merged_df_pf["CONTRACT NO"], merged_df_pf["PURPOSE"]))

map_primary   = dict(zip(productCategory['SPECIAL CATEGORIES (CONTRACT WISE) *Add here to pick contract wise cat'], productCategory['Unnamed: 12']))   # A -> M (via L)
map_moratorium= dict(zip(productCategory['MORATORIUM CATEGORY'], productCategory['Unnamed: 9']))   # E -> J (via I)
map_composite = dict(zip(productCategory['LOOKUP'], productCategory['Classification']))

contract_mot = dict(zip(LoanCategory['New Contract (Loan)'], LoanCategory['Old Contract Category']))  # (C+D+F+G) -> E (via F)

contract_to_arrcap = dict(zip(PortfolioLastMonth["CONTRACT NO"], PortfolioLastMonth["ARRCAP"]))
contract_to_arrint = dict(zip(PortfolioLastMonth["CONTRACT NO"], PortfolioLastMonth["ARRINT"]))
contract_to_arrtax = dict(zip(PortfolioLastMonth["CONTRACT NO"], PortfolioLastMonth["ARRTAX"]))
contract_to_othout = dict(zip(PortfolioLastMonth["CONTRACT NO"], PortfolioLastMonth["OTHOUT"]))

# Find the SET-OFF TARGET column dynamically (it changes by month)
# Try different possible column name patterns
setoff_col = None
for col in expectedSetoff.columns:
    col_upper = str(col).upper()
    if "SET-OFF TARGET" in col_upper or "SETOFF TARGET" in col_upper or "SET OFF TARGET" in col_upper:
        setoff_col = col
        break

if not setoff_col:
    # If still not found, look for any column with "TARGET" and "SET" or just "TARGET"
    for col in expectedSetoff.columns:
        col_upper = str(col).upper()
        if "TARGET" in col_upper and ("SET" in col_upper or "SETOFF" in col_upper):
            setoff_col = col
            break

    if not setoff_col:
        # Last resort: use second column if it exists (assuming CON_NO is first)
        if len(expectedSetoff.columns) >= 2:
            setoff_col = expectedSetoff.columns[1]
            logger.warning(f"Could not find SET-OFF TARGET column, using second column: {setoff_col}")
        else:
            raise ValueError(f"No column starting with 'SET-OFF TARGET' found. Available columns: {list(expectedSetoff.columns)}")

logger.info(f"Using setoff column: {setoff_col}")
expectedSetoffDict = dict(zip(expectedSetoff["CON_NO"], expectedSetoff[setoff_col]))

Report['CLM_CODE'] = Report['CONTRACT NO'].map(contract_to_client)

Report['EQP_TYPE'] = Report['CONTRACT NO'].map(contract_to_quip)
Report['MOR CAT'] = Report['CONTRACT NO'].map(contract_mot)
Report['PURPOSE'] = Report['CONTRACT NO'].map(contract_to_purpose)
Report["CLIENT TYPE"] = Report["CLM_CODE"].apply(
    lambda x: np.nan if pd.isna(x) else ("Corporate Client" if str(x)[0] == "2" else "Non-Corporate")
)
result = Report['CONTRACT NO'].map(map_primary)

# 3) Fallback A: if D == "Moratorium Loan" → XLOOKUP(E, I, J)
mask_moratorium = result.isna() & (Report['EQP_TYPE'] == 'Moratorium Loan')
result[mask_moratorium] = Report.loc[mask_moratorium, 'MOR CAT'].map(map_moratorium)

# 4) Fallback B: if C == "FDL" → literal "Loans against Cash/Deposits"
mask_fdl = result.isna() & (Report['PRODUCT'] == 'FDL')
result[mask_fdl] = "Loans against Cash/Deposits"

# 5) Fallback C: else → XLOOKUP(C&D&F&G, F, E)
mask_composite = result.isna()
# build the composite key exactly like Excel's C8&D8&F8&G8 (coerce to string to avoid NaN issues)
composite_key = (
    Report['PRODUCT'].fillna('').astype(str)
    + Report['EQP_TYPE'].fillna('').astype(str)
    + Report['PURPOSE'].fillna('').astype(str)
    + Report['CLIENT TYPE'].fillna('').astype(str)
)
result[mask_composite] = composite_key[mask_composite].map(map_composite)

# 6) Attach to df (e.g., as a new column matching the Excel formula output)
Report['CBSL CATEGORY'] = result
Report['ARREARS CAPITAL'] = Report['CONTRACT NO'].map(contract_to_arrcap).fillna(0)
Report['ARREARS INTEREST'] = Report['CONTRACT NO'].map(contract_to_arrint).fillna(0)
Report['ARREARS TAX'] = Report['CONTRACT NO'].map(contract_to_arrtax).fillna(0)
Report['OTHER CHARGES'] = Report['CONTRACT NO'].map(contract_to_othout).fillna(0)
Report['EXPECTED SETOFF'] = Report['CONTRACT NO'].map(expectedSetoffDict).fillna(0)

Report['TOTAL OPENING ARREARS'] = Report['ARREARS CAPITAL'] + Report['ARREARS INTEREST'] + Report['ARREARS TAX'] + Report['OTHER CHARGES']
# Step 1: Calculate 'ADD TO EXPECTED' based on the given formula
Report['ADD TO EXPECTED'] = Report['TOTAL MONTH DUE'] - Report['MONTH DUE (OTHCHG)'] - Report['MONTH DUE (RENTAL)'] + Report['AT SETOFF']
Report['FINAL EXPECTED SETOFF'] = Report['ADD TO EXPECTED'] + Report['EXPECTED SETOFF']
# Step 1: Calculate 'OPENING (RENTAL)' based on the given formula
Report['OPENING (RENTAL)'] = Report['ARREARS CAPITAL'] + Report['ARREARS INTEREST']
# Step 1: Calculate 'OPENING (OTHCHG)' based on the given formula
Report['OPENING (OTHCHG)'] = Report['ARREARS TAX'] + Report['OTHER CHARGES']
Report['EXPECTED SETOFF MONTH DUE'] = np.where(
    (Report['FINAL EXPECTED SETOFF'] - Report['OPENING (RENTAL)'] - Report['OPENING (OTHCHG)']) < 0, 
    0, 
    Report['FINAL EXPECTED SETOFF'] - Report['OPENING (RENTAL)'] - Report['OPENING (OTHCHG)']
)
# Step 1: Apply the conditional formula to the ACTUAL COLLECTION column
Report['ACTUAL COLLECTION'] = np.where(
    (Report['TOTAL SETOFF'] - Report['OPENING (OTHCHG)'] - Report['OPENING (RENTAL)']) < 0, 
    0, 
    Report['TOTAL SETOFF'] - Report['OPENING (OTHCHG)'] - Report['OPENING (RENTAL)']
)
# Step 1: Ensure the columns are numeric and fill NaN values with 0
Report['TOTAL OPENING ARREARS'] = pd.to_numeric(Report['TOTAL OPENING ARREARS'], errors='coerce').fillna(0)
Report['TOTAL MONTH DUE'] = pd.to_numeric(Report['TOTAL MONTH DUE'], errors='coerce').fillna(0)
Report['TOTAL SETOFF'] = pd.to_numeric(Report['TOTAL SETOFF'], errors='coerce').fillna(0)
Report['FINAL EXPECTED SETOFF'] = pd.to_numeric(Report['FINAL EXPECTED SETOFF'], errors='coerce').fillna(0)

# Step 2: Apply the conditional formula to the 'ALL ZERO REMOVE' column
Report['ALL ZERO REMOVE'] = np.where(
    (Report['TOTAL OPENING ARREARS'] + Report['TOTAL MONTH DUE'] + Report['TOTAL SETOFF'] + Report['FINAL EXPECTED SETOFF']) == 0, 
    'ALL ZERO', 
    'OK'
)
logger.info("Report dtypes before filtering:\n%s", Report.dtypes)
# Step 1: Delete rows where 'ALL ZERO REMOVE' is "ALL ZERO"
Report = Report[Report['ALL ZERO REMOVE'] != 'ALL ZERO']

# Step 2: Optionally, reset the index if needed (to avoid gaps in the index)
Report.reset_index(drop=True, inplace=True)

lastMonthStage = dict(zip(SummaryLastMonth["CONTRACT NO"], SummaryLastMonth["STAGE"]))
currentMonthStage = dict(zip(SummaryCurrentMonth["CONTRACT NO"], SummaryCurrentMonth["STAGE"]))

# Step 1: Update 'STAGE' based on the conditions
def update_stage(contract_no):
    # Check if the contract is in currentMonthStage
    if contract_no in currentMonthStage:
        return currentMonthStage[contract_no]
    # If not in currentMonthStage, check in lastMonthStage
    elif contract_no in lastMonthStage:
        return lastMonthStage[contract_no]
    # If not found in either, set to 1
    else:
        return 1

# Step 2: Apply the function to update the 'STAGE' column in the Report dataframe
Report['STAGE'] = Report['CONTRACT NO'].apply(update_stage)
Report['STAGE TARGET %'] = 0
Report['WEIGHTAGE'] = 0

# Fill missing CLM_CODE/EQP_TYPE via API for the affected contracts
filtered_df = Report[Report['CLM_CODE'].isna()][['CONTRACT NO', 'CLM_CODE']]

logger.info(f"filtered_df Shape: {filtered_df.shape}")
logger.info(f"Initial contracts with missing CLM_CODE: {filtered_df['CONTRACT NO'].tolist()}")

if not filtered_df.empty:
    try:
        session, verification_token = get_authenticated_session()
        updated_count = 0
        failed_contracts = []
        
        for _, row in filtered_df.iterrows():
            contract_no = str(row['CONTRACT NO']).strip()
            logger.info("Missing CLM_CODE for CONTRACT NO: %s - querying API", contract_no)
            display_data, success = process_contract_with_retry(session, contract_no, verification_token)
            logger.info("API Response - Data: %s, Success: %s", display_data, success)
            
            if success and isinstance(display_data, dict):
                client_code = display_data.get('client_code')
                equipment = display_data.get('equipment')
                
                if client_code or equipment:
                    mask = Report['CONTRACT NO'] == contract_no
                    if client_code:
                        Report.loc[mask, 'CLM_CODE'] = client_code
                        logger.info("✓ Updated CLM_CODE for %s: %s", contract_no, client_code)
                    if equipment:
                        Report.loc[mask, 'EQP_TYPE'] = equipment
                        logger.info("✓ Updated EQP_TYPE for %s: %s", contract_no, equipment)
                    updated_count += 1
                else:
                    logger.warning("✗ API returned success but no client_code/equipment for %s", contract_no)
                    failed_contracts.append(contract_no)
            else:
                logger.warning("✗ API could not provide data for %s", contract_no)
                failed_contracts.append(contract_no)

        logger.info("="*60)
        logger.info("API Enrichment Summary:")
        logger.info(f"  - Total contracts queried: {len(filtered_df)}")
        logger.info(f"  - Successfully updated: {updated_count}")
        logger.info(f"  - Failed to update: {len(failed_contracts)}")
        if failed_contracts:
            logger.warning(f"  - Failed contracts: {failed_contracts}")
        logger.info("="*60)
        
    except Exception as e:
        logger.exception("Failed during API enrichment step: %s", e)

    # Recompute dependent fields after enrichment
    try:
        logger.info("Recomputing dependent fields after API enrichment...")
        
        # Recompute CLIENT TYPE - be explicit about NaN handling
        Report["CLIENT TYPE"] = Report["CLM_CODE"].apply(
            lambda x: (
                np.nan if pd.isna(x) or str(x).strip() == '' 
                else ("Corporate Client" if str(x)[0] == "2" else "Non-Corporate")
            )
        )

        # Re-run category resolution
        result = Report['CONTRACT NO'].map(map_primary)
        mask_moratorium = result.isna() & (Report['EQP_TYPE'] == 'Moratorium Loan')
        result[mask_moratorium] = Report.loc[mask_moratorium, 'MOR CAT'].map(map_moratorium)

        mask_fdl = result.isna() & (Report['PRODUCT'] == 'FDL')
        result[mask_fdl] = "Loans against Cash/Deposits"

        mask_composite = result.isna()
        composite_key = (
            Report['PRODUCT'].fillna('').astype(str)
            + Report['EQP_TYPE'].fillna('').astype(str)
            + Report['PURPOSE'].fillna('').astype(str)
            + Report['CLIENT TYPE'].fillna('').astype(str)
        )
        result[mask_composite] = composite_key[mask_composite].map(map_composite)
        Report['CBSL CATEGORY'] = result
        
        logger.info("✓ Recomputed CLIENT TYPE and CBSL CATEGORY after API enrichment")
        
        # Check for remaining blanks after recomputation
        remaining_blanks = Report[Report['CLM_CODE'].isna()]
        logger.info(f"Remaining contracts with blank CLM_CODE after enrichment: {len(remaining_blanks)}")
        if not remaining_blanks.empty:
            logger.warning(f"Contracts still missing CLM_CODE: {remaining_blanks['CONTRACT NO'].tolist()}")
            
    except Exception as e:
        logger.exception("Failed to recompute dependent fields after enrichment: %s", e)

# Final validation before writing to Excel
final_blanks = Report[Report['CLM_CODE'].isna()]
logger.info("="*60)
logger.info(f"FINAL CHECK - Contracts with blank CLM_CODE: {len(final_blanks)}")
if not final_blanks.empty:
    logger.warning("Contracts with missing CLM_CODE that will be written to Excel:")
    for idx, row in final_blanks.iterrows():
        logger.warning(f"  - {row['CONTRACT NO']}")
logger.info("="*60)

# Step 1: Define the desired column order
column_order = [
    'CONTRACT NO', 'CLM_CODE', 'PRODUCT', 'EQP_TYPE', 'MOR CAT', 'PURPOSE', 
    'CLIENT TYPE', 'CBSL CATEGORY', 'ARREARS CAPITAL', 'ARREARS INTEREST', 
    'ARREARS TAX', 'OTHER CHARGES', 'TOTAL OPENING ARREARS', 'STAGE TARGET %', 
    'WEIGHTAGE', 'FINAL EXPECTED SETOFF', 'OPENING (RENTAL)', 'MONTH DUE (RENTAL)', 
    'OPENING (OTHCHG)', 'MONTH DUE (OTHCHG)', 'TOTAL MONTH DUE', 
    'EXPECTED SETOFF MONTH DUE', 'TOTAL SETOFF', 'ACTUAL COLLECTION', 
    'ALL ZERO REMOVE', 'STAGE', 'ADD TO EXPECTED', 'EXPECTED SETOFF', 'AT SETOFF'
]

# Step 2: Reorder the columns in the Report dataframe
Report = Report[column_order]
#exception report CBSL CATEGORY NULL RECORDS
exceptionReport = Report[Report['CBSL CATEGORY'].isnull()]
logger.info("Exception Report CBSL CATEGORY NULL RECORDS:\n%s", exceptionReport.head())
# Step 1: Open the Excel .xlsb file - use the same file we found earlier
loan_cf_file = find_file("Loan CF Analysis*.xlsb", working_dir)

exceptionReport = exceptionReport[['PRODUCT','EQP_TYPE','PURPOSE','CLIENT TYPE']].drop_duplicates()


try: 
    # Load the Excel file using pyxlsb engine for .xlsb format 
    file_path = find_file("Product_Cat.xlsx", working_dir)
    # Assuming exceptionReport is a DataFrame, and you need to drop duplicates 
    exceptionReport = exceptionReport.drop_duplicates() 
 
    # Create a new Excel instance (or connect to an existing one) 
    app = xw.App(visible=False)  # visible=True for debugging, set to False in production 
    wb = app.books.open(file_path) 
 
    # Access the 'Product_Cat' sheet 
    sheet = wb.sheets['Sheet1'] 
 
    # Find the last used row in column A (or use another column if needed)
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    
    # Write data starting from the next empty row
    sheet.range(f'A{last_row + 1}').value = exceptionReport.values.tolist()
 
    # Save the workbook 
    wb.save() 
    wb.close() 
 
    # Close the Excel application 
    app.quit() 
 
    logger.info("exceptionReport written successfully to the end of Product_Cat sheet.") 
except Exception as e: 
    logger.error(f"Error occurred while writing to the Excel file: {e}") 
 
wb = None
try:
    logger.info("Opening workbook for Report data: %s", loan_cf_file)
    # Step 1: Open the workbook
    wb = xw.Book(loan_cf_file)
    logger.info("Workbook opened successfully")

    # Step 2: Select the 'WORKING' sheet
    sheet = wb.sheets['Working']
    logger.info("Selected 'Working' sheet")

    # Step 3: Delete rows after row 3
    # Get the last used row in the sheet using a more reliable method
    try:
        last_row = sheet.range('A4').end('down').row
        logger.info("Found last row using end('down'): %d", last_row)
    except Exception as e:
        logger.warning("end('down') failed: %s, using UsedRange instead", e)
        # If end() fails, use a safer method
        last_row = sheet.api.UsedRange.Rows.Count
        logger.info("Found last row using UsedRange: %d", last_row)

    if last_row > 3:
        logger.info("Deleting rows from 4 to %d", last_row)
        sheet.range(f'A4:A{last_row}').api.EntireRow.Delete()  # Deletes rows from 4 to the end
        logger.info("Rows deleted successfully")

    # Step 4: Add your data to A4
    logger.info("Writing Report data to A4 (%d rows)", len(Report))
    sheet.range('A4').value = Report.values
    logger.info("Data written successfully")

    # Step 5: Add SUM functions to cells I1 through AC1
    logger.info("Adding SUM formulas from I1 to AC1")
    # Define the range of columns from I to AC
    start_col = 'I'  # Column I (9th column)
    end_col = 'AC'   # Column AC (29th column)

    # Get column numbers for iteration
    start_col_num = ord(start_col) - ord('A') + 1  # I = 9
    end_col_num = 26 + (ord(end_col[1]) - ord('A') + 1)  # AC = 26 + 3 = 29

    # Add SUM formulas to each cell from I1 to AC1
    for col_num in range(start_col_num, end_col_num + 1):
        if col_num <= 26:
            # Single letter columns (A-Z)
            col_letter = chr(ord('A') + col_num - 1)
        else:
            # Double letter columns (AA, AB, AC, etc.)
            first_letter = chr(ord('A') + (col_num - 27) // 26)
            second_letter = chr(ord('A') + (col_num - 1) % 26)
            col_letter = first_letter + second_letter

        # Create the SUM formula for the current column
        formula = f"=SUM({col_letter}3:{col_letter}1000000)"

        # Apply the formula to the cell in row 1
        sheet.range(f'{col_letter}1').formula = formula

    logger.info("SUM formulas added successfully")

    # Step 6: Save the workbook (Ctrl+S - save with same filename)
    logger.info("Saving workbook: %s", loan_cf_file)
    wb.save()  # Save to the same file (Ctrl+S behavior)
    logger.info("Workbook processed successfully - data written and formulas updated; saved as %s", loan_cf_file)
except Exception as e:
    logger.error("Error processing workbook: %s", e)
    logger.exception("Full traceback:")
finally:
    # Step 7: Close the workbook
    if wb is not None:
        try:
            logger.info("Closing workbook")
            wb.close()
            logger.info("Workbook closed successfully")
        except Exception as e:
            logger.error("Error closing workbook: %s", e)

logger.info("Product category sample:\n%s", productCategory.head())
logger.info("Final Report sample:\n%s", Report.head())
logger.info("="*80)
logger.info(f"NBD-MF-23-C4 Report Completed Successfully at {datetime.now()}")
logger.info("="*80)