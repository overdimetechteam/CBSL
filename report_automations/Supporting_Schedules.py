import os
import logging
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pyxlsb import open_workbook

# Configure logging
def setup_logging():
    """Setup logging — writes to /logs/run_YYYYMMDD_HHMMSS/Supporting_Schedules.log"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_dir = Path(__file__).parent.parent / "logs" / f"run_{timestamp}"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "Supporting_Schedules.log"

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(fmt)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(fmt)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info(f"Log file: {log_file}")
    return logger

logger = setup_logging()


def get_monthly_folder(base_path: Path) -> Optional[Path]:
    
    #Get the single subdirectory inside the monthly folder.
    
    logger.info(f"Searching for monthly folder in: {base_path}")
    
    if not base_path.exists() or not base_path.is_dir():
        logger.error(f"Base path does not exist or is not a directory: {base_path}")
        return None
    
    subdirs = [p for p in base_path.iterdir() if p.is_dir()]
    logger.info(f"Found {len(subdirs)} subdirectories: {[p.name for p in subdirs]}")
    
    if len(subdirs) != 1:
        logger.error(f"Expected exactly 1 subdirectory in monthly folder, found {len(subdirs)}")
        return None
    
    monthly_folder = subdirs[0]
    logger.info(f"Selected monthly folder: {monthly_folder}")
    return monthly_folder


def find_gl_account_details_file(monthly_folder: Path) -> Optional[Path]:
    
    #Find the GL Account Details Excel file directly in the date folder.
    logger.info(f"Searching for GL Account Details file in: {monthly_folder}")

    if not monthly_folder.exists():
        logger.error(f"Folder not found: {monthly_folder}")
        return None

    identifier = "GL Account Details"
    found_files = []

    for file in monthly_folder.glob("*.xlsx"):
        if file.name.startswith("~$"):
            continue
        if identifier in file.name:
            found_files.append(file)
            logger.info(f"Found matching file: {file.name}")
    
    if not found_files:
        logger.error(f"No file containing '{identifier}' found in {monthly_folder}")
        return None

    if len(found_files) > 1:
        logger.warning(f"Multiple files found, using first: {found_files[0].name}")

    gl_file = found_files[0]
    logger.info(f"Selected file: {gl_file}")
    return gl_file


def get_month_from_gl_file(working_dir: Path, sheet_name: str = "GL Account Details") -> Optional[tuple]:
    """
    Extract the month name and date from the GL Account Details file by reading
    'Transaction From Date' and 'Transaction To Date' from columns A and B.
    
    Returns: Tuple of (month_name, date_object) or None if not found.
             e.g., ('August', datetime(2025, 8, 1))
    """
    logger.info("Extracting month from GL Account Details file...")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return None
        
        gl_file = find_gl_account_details_file(monthly_folder)
        if not gl_file:
            return None
        
        logger.info(f"Reading header data from: {gl_file.name}")
        
        # Read the first few rows without skipping to get the header information
        df_header = pd.read_excel(gl_file, sheet_name=sheet_name, engine='openpyxl', header=None, nrows=15)
        
        logger.info(f"Header data shape: {df_header.shape}")
        
        # Look for "Transaction From Date" and "Transaction To Date" in column A (index 0)
        transaction_from_date = None
        transaction_to_date = None
        
        for idx, row in df_header.iterrows():
            cell_value = str(row[0]).strip() if pd.notna(row[0]) else ""
            
            if "Transaction From Date" in cell_value:
                # Get the date from column B (index 1)
                if pd.notna(row[1]):
                    transaction_from_date = row[1]
                    logger.info(f"Found 'Transaction From Date' at row {idx}: {transaction_from_date}")
            
            elif "Transaction To Date" in cell_value:
                # Get the date from column B (index 1)
                if pd.notna(row[1]):
                    transaction_to_date = row[1]
                    logger.info(f"Found 'Transaction To Date' at row {idx}: {transaction_to_date}")
        
        # Extract month from either date (prefer From Date)
        date_to_use = transaction_from_date if transaction_from_date else transaction_to_date
        
        if date_to_use is None:
            logger.error("Could not find Transaction From Date or Transaction To Date")
            return None
        
        # Convert to datetime - handle dd/mm/yyyy format
        if isinstance(date_to_use, str):
            # Try dd/mm/yyyy format first (common in Sri Lanka/international formats)
            try:
                date_obj = pd.to_datetime(date_to_use, format='%d/%m/%Y')
                logger.info(f"Parsed date as dd/mm/yyyy: {date_obj}")
            except:
                # Fallback to automatic parsing
                date_obj = pd.to_datetime(date_to_use, dayfirst=True)
                logger.info(f"Parsed date with dayfirst=True: {date_obj}")
        else:
            date_obj = pd.to_datetime(date_to_use)
        
        month_name = date_obj.strftime("%B")  # Full month name
        logger.info(f"Extracted month: {month_name}")
        logger.info(f"Date object: {date_obj}")
        
        return (month_name, date_obj)
        
    except Exception as e:
        logger.error(f"Error extracting month from GL file: {e}", exc_info=True)
        return None


def get_gl_account_data(working_dir: Path, sheet_name: str = "GL Account Details") -> Optional[pd.DataFrame]:
    # Get GL Account Details data in a DataFrame.
    logger.info("Getting GL Account Details data...")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return None
        
        gl_file = find_gl_account_details_file(monthly_folder)
        if not gl_file:
            return None
        
        logger.info(f"Reading Excel file: {gl_file.name}")
        df_glacc = pd.read_excel(gl_file, sheet_name=sheet_name, engine='openpyxl', skiprows=11)
        
        logger.info(f"Excel file read successfully. Shape: {df_glacc.shape}")
        logger.info(f"Columns found: {df_glacc.columns.tolist()}")

        debit_col = None
        credit_col = None

        for col in df_glacc.columns:
            col_str = str(col).strip()
            if "Debit Amount" in col_str or col_str == "Debit Amount":
                debit_col = col
                logger.info(f"Found 'Debit Amount' column: {col}")
            elif "Credit Amount" in col_str or col_str == "Credit Amount":
                credit_col = col
                logger.info(f"Found 'Credit Amount' column: {col}")
        
        if debit_col is None or credit_col is None:
            logger.error("Required columns not found in Excel file")
            logger.error(f"Available columns: {df_glacc.columns.tolist()}")
            return None

        debit_values = pd.to_numeric(df_glacc[debit_col], errors='coerce')
        credit_values = pd.to_numeric(df_glacc[credit_col], errors='coerce')

        mask_filled = debit_values.notna() | credit_values.notna()
        if mask_filled.any():
            last_filled_idx = mask_filled[::-1].idxmax()
            last_filled_pos = df_glacc.index.get_loc(last_filled_idx)
            logger.info(f"Last filled row position: {last_filled_pos} (Excel row: {last_filled_pos + 12})")
            debit_values = debit_values.iloc[:last_filled_pos + 1]
            credit_values = credit_values.iloc[:last_filled_pos + 1]
        else:
            logger.warning("No data found in Debit Amount or Credit Amount columns")
            return None
        
        df_result = pd.DataFrame({
            'Debit Amount': debit_values.fillna(0),
            'Credit Amount': credit_values.fillna(0)
        })
        
        df_result = df_result.reset_index(drop=True)
        df_result['Net Amount'] = df_result['Debit Amount'] - df_result['Credit Amount']
        
        logger.info(f"Data stored in DataFrame. Shape: {df_result.shape}")
        logger.info("Full DataFrame contents:\n" + df_result.to_string(index=True))

        return df_result
        
    except Exception as e:
        logger.error(f"Error getting GL Account data: {e}", exc_info=True)
        return None


def find_supporting_schedule_file(monthly_folder: Path) -> Optional[Path]:
    """
    Find the Supporting Schedule Excel file in the Supporting_Schedules subdirectory.
    """
    logger.info(f"Searching for Supporting Schedule file in: {monthly_folder}")

    if not monthly_folder.exists():
        logger.error(f"Folder not found: {monthly_folder}")
        return None

    identifier = "Supporting Schedules"
    found_files = []

    for file in monthly_folder.glob("*.xlsx"):
        if file.name.startswith("~$"):
            continue
        if identifier in file.name:
            found_files.append(file)
            logger.info(f"Found matching file: {file.name}")

    if not found_files:
        logger.error(f"No file containing '{identifier}' found in {monthly_folder}")
        return None
    
    if len(found_files) > 1:
        logger.warning(f"Multiple files found, using first: {found_files[0].name}")
    
    ss_file = found_files[0]
    logger.info(f"Selected file: {ss_file}")
    return ss_file


def get_new_supporting_schedule_filename(ss_file: Path, date_obj: datetime) -> Path:
    """
    Generate new filename for Supporting Schedule file based on the month/year from data.
    
    Args:
        ss_file: Current Supporting Schedule file path
        date_obj: Date object containing the month/year to use
    
    Returns:
        Path object with new filename
    """
    # Format: "Supporting Schedules - Aug 2025.xlsx"
    month_abbr = date_obj.strftime("%b")  # 3-letter month abbreviation (Jan, Feb, etc.)
    year = date_obj.strftime("%Y")  # 4-digit year
    
    new_filename = f"Supporting Schedules - {month_abbr} {year}.xlsx"
    new_file_path = ss_file.parent / new_filename
    
    logger.info(f"New filename will be: {new_filename}")
    return new_file_path


def paste_data_to_supporting_schedule(working_dir: Path, df_data: pd.DataFrame, month_name: str, 
                                      sheet_name: str = "New Shcedule") -> bool:
    """
    Paste Net Amount data (rows 1-6) from dataframe to the Supporting Schedule file
    under the matching month column.
    
    Args:
        working_dir: Path to the working directory
        df_data: DataFrame with Net Amount column
        month_name: Month name to find in the Supporting Schedule (e.g., 'August')
        sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info(f"Pasting data to Supporting Schedule for month: {month_name}")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        logger.info(f"Opening workbook: {ss_file.name}")
        wb = load_workbook(ss_file)
        
        if sheet_name not in wb.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}")
            wb.close()
            return False
        
        ws = wb[sheet_name]
        logger.info(f"Working with sheet: {sheet_name}")
        
        # Find the row with "Description" in column A
        description_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_col=1, max_col=1, values_only=False), start=1):
            cell_value = str(row[0].value).strip() if row[0].value else ""
            if "Description" in cell_value:
                description_row = row_idx
                logger.info(f"Found 'Description' in column A at row {description_row}")
                break
        
        if description_row is None:
            logger.error("Could not find 'Description' in column A")
            wb.close()
            return False
        
        # Find the month column by searching the header row (description_row)
        month_col = None
        logger.info(f"Searching for month '{month_name}' in row {description_row}")
        
        # Log all column headers for debugging
        for col_idx, cell in enumerate(ws[description_row], start=1):
            cell_value = cell.value
            cell_str = str(cell_value).strip() if cell_value else ""
            logger.info(f"Column {get_column_letter(col_idx)}: '{cell_str}'")
            
            # Check if cell contains a datetime object and extract month name
            if isinstance(cell_value, datetime):
                cell_month = cell_value.strftime("%B")  # Full month name
                if cell_month == month_name:
                    month_col = col_idx
                    month_col_letter = get_column_letter(col_idx)
                    logger.info(f"Found month '{month_name}' at column {month_col_letter} (index {month_col})")
                    break
            # Also check for text-based month names
            elif month_name in cell_str or month_name[:3] in cell_str:
                month_col = col_idx
                month_col_letter = get_column_letter(col_idx)
                logger.info(f"Found month '{month_name}' at column {month_col_letter} (index {month_col})")
                break
        
        if month_col is None:
            logger.error(f"Could not find month '{month_name}' in row {description_row}")
            wb.close()
            return False
        
        # Paste rows 1-6 from Net Amount column to the month column
        # Starting from the row after Description row
        start_row = description_row + 1
        
        logger.info(f"Pasting data starting at row {start_row}, column {get_column_letter(month_col)}")
        
        # Clear existing values before pasting
        for i in range(1, 7):
            target_row = start_row + (i - 1)
            ws.cell(row=target_row, column=month_col, value=None)
        logger.info(f"Cleared cells {get_column_letter(month_col)}{start_row} to {get_column_letter(month_col)}{start_row+5}")
        
        for i in range(1, 7):  # Rows 1 to 6 from dataframe
            if i < len(df_data):
                net_amount = df_data.loc[i, 'Net Amount']
                target_row = start_row + (i - 1)
                ws.cell(row=target_row, column=month_col, value=net_amount)
                logger.info(f"Pasted row {i}: {net_amount} to cell {get_column_letter(month_col)}{target_row}")
            else:
                logger.warning(f"DataFrame row {i} does not exist, skipping")
        
        # Save the workbook
        wb.save(ss_file)
        wb.close()
        logger.info(f"Successfully saved data to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error pasting data to Supporting Schedule: {e}", exc_info=True)
        return False


def find_cbsl_provision_file(monthly_folder: Path) -> Optional[Path]:
    """
    Find the CBSL Provision Comparison file in the Supporting_Schedules subdirectory.
    """
    logger.info(f"Searching for CBSL Provision Comparison file in: {monthly_folder}")

    if not monthly_folder.exists():
        logger.error(f"Folder not found: {monthly_folder}")
        return None

    identifier = "CBSL Provision Comparison"
    found_files = []

    for file in monthly_folder.glob("*.xlsb"):
        if file.name.startswith("~$"):
            continue
        if identifier in file.name:
            found_files.append(file)
            logger.info(f"Found matching file: {file.name}")

    if not found_files:
        logger.error(f"No .xlsb file containing '{identifier}' found in {monthly_folder}")
        return None
    
    if len(found_files) > 1:
        logger.warning(f"Multiple files found, using first: {found_files[0].name}")
    
    cbsl_file = found_files[0]
    logger.info(f"Selected file: {cbsl_file}")
    return cbsl_file


def get_cbsl_provision_values(working_dir: Path, sheet_name: str = "Comparison") -> Optional[dict]:
    """
    Get values from CBSL Provision Comparison file at specific cells.
    
    Returns:
        dict with keys 'J11', 'J12', 'L14' containing the values, or None if error
    """
    logger.info("Getting CBSL Provision values from Comparison sheet...")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return None
        
        cbsl_file = find_cbsl_provision_file(monthly_folder)
        if not cbsl_file:
            return None
        
        logger.info(f"Reading .xlsb file: {cbsl_file.name}")
        
        # Read .xlsb file using pyxlsb
        with open_workbook(str(cbsl_file)) as wb:
            # Get sheet names - wb.sheets is a list of sheet names as strings
            sheet_names = wb.sheets
            logger.info(f"Available sheets: {sheet_names}")
            
            if sheet_name not in sheet_names:
                logger.error(f"Sheet '{sheet_name}' not found. Available sheets: {sheet_names}")
                return None
            
            with wb.get_sheet(sheet_name) as sheet:
                logger.info(f"Reading from sheet: {sheet_name}")
                
                # Convert to DataFrame for easier cell access
                data = []
                for row in sheet.rows():
                    data.append([cell.v for cell in row])
                
                df = pd.DataFrame(data)
                
                logger.info(f"DataFrame shape: {df.shape}")
                logger.info(f"First 20 rows of column J (index 9):\n{df.iloc[:20, 9] if df.shape[1] > 9 else 'Column J not found'}")
                logger.info(f"First 20 rows of column L (index 11):\n{df.iloc[:20, 11] if df.shape[1] > 11 else 'Column L not found'}")
                
                # Get values from J11, J12, L14
                # J is column 9 (0-indexed), L is column 11 (0-indexed)
                # Row 11 is index 10, row 12 is index 11, row 14 is index 13 (0-indexed)
                j11_value = df.iloc[10, 9] if df.shape[0] > 10 and df.shape[1] > 9 else None
                j12_value = df.iloc[11, 9] if df.shape[0] > 11 and df.shape[1] > 9 else None
                l14_value = df.iloc[13, 11] if df.shape[0] > 13 and df.shape[1] > 11 else None
                
                logger.info(f"J11 (row 11, col J, index [10,9]) value: {j11_value}")
                logger.info(f"J12 (row 12, col J, index [11,9]) value: {j12_value}")
                logger.info(f"L14 (row 14, col L, index [13,11]) value: {l14_value}")
                
                return {
                    'J11': j11_value,
                    'J12': j12_value,
                    'L14': l14_value
                }
        
    except Exception as e:
        logger.error(f"Error reading CBSL Provision file: {e}", exc_info=True)
        return None


def paste_cbsl_provision_to_supporting_schedule(working_dir: Path, cbsl_values: dict, month_name: str,
                                                sheet_name: str = "New Shcedule") -> bool:
    """
    Paste CBSL Provision values to Supporting Schedule file at specific rows.
    
    Mapping:
    - J11 → row 25
    - J12 → row 24
    - L14 → row 26
    
    Args:
        working_dir: Path to the working directory
        cbsl_values: Dictionary with keys 'J11', 'J12', 'L14'
        month_name: Month name to find the column
        sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info(f"Pasting CBSL Provision data to Supporting Schedule for month: {month_name}")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        logger.info(f"Opening workbook: {ss_file.name}")
        wb = load_workbook(ss_file)
        
        if sheet_name not in wb.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}")
            wb.close()
            return False
        
        ws = wb[sheet_name]
        logger.info(f"Working with sheet: {sheet_name}")
        
        # Find the row with "Description" in column A
        description_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_col=1, max_col=1, values_only=False), start=1):
            cell_value = str(row[0].value).strip() if row[0].value else ""
            if "Description" in cell_value:
                description_row = row_idx
                logger.info(f"Found 'Description' in column A at row {description_row}")
                break
        
        if description_row is None:
            logger.error("Could not find 'Description' in column A")
            wb.close()
            return False
        
        # Find the month column
        month_col = None
        logger.info(f"Searching for month '{month_name}' in row {description_row}")
        
        for col_idx, cell in enumerate(ws[description_row], start=1):
            cell_value = cell.value
            cell_str = str(cell_value).strip() if cell_value else ""
            
            # Check if cell contains a datetime object and extract month name
            if isinstance(cell_value, datetime):
                cell_month = cell_value.strftime("%B")  # Full month name
                if cell_month == month_name:
                    month_col = col_idx
                    month_col_letter = get_column_letter(col_idx)
                    logger.info(f"Found month '{month_name}' at column {month_col_letter} (index {month_col})")
                    break
            # Also check for text-based month names
            elif month_name in cell_str or month_name[:3] in cell_str:
                month_col = col_idx
                month_col_letter = get_column_letter(col_idx)
                logger.info(f"Found month '{month_name}' at column {month_col_letter} (index {month_col})")
                break
        
        if month_col is None:
            logger.error(f"Could not find month '{month_name}' in row {description_row}")
            wb.close()
            return False
        
        # Clear existing values in target cells before pasting
        ws.cell(row=25, column=month_col, value=None)
        ws.cell(row=24, column=month_col, value=None)
        ws.cell(row=26, column=month_col, value=None)
        logger.info(f"Cleared cells {get_column_letter(month_col)}24, {get_column_letter(month_col)}25, {get_column_letter(month_col)}26")
        
        # Paste values to specific rows
        # J11 → row 25
        if cbsl_values.get('J11') is not None:
            ws.cell(row=25, column=month_col, value=cbsl_values['J11'])
            logger.info(f"Pasted J11 value ({cbsl_values['J11']}) to cell {get_column_letter(month_col)}25")
        else:
            logger.warning(f"J11 value is None, skipping paste to row 25")
        
        # J12 → row 24
        if cbsl_values.get('J12') is not None:
            ws.cell(row=24, column=month_col, value=cbsl_values['J12'])
            logger.info(f"Pasted J12 value ({cbsl_values['J12']}) to cell {get_column_letter(month_col)}24")
        else:
            logger.warning(f"J12 value is None, skipping paste to row 24")
        
        # L14 → row 26
        if cbsl_values.get('L14') is not None:
            ws.cell(row=26, column=month_col, value=cbsl_values['L14'])
            logger.info(f"Pasted L14 value ({cbsl_values['L14']}) to cell {get_column_letter(month_col)}26")
        else:
            logger.warning(f"L14 value is None, skipping paste to row 26")
        
        # Save the workbook
        wb.save(ss_file)
        wb.close()
        logger.info(f"Successfully saved CBSL Provision data to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error pasting CBSL Provision data to Supporting Schedule: {e}", exc_info=True)
        return False


def copy_property_mortgage_to_new_schedule(working_dir: Path,
                                           ss_sheet_name: str = "New Shcedule") -> bool:
    """
    Copy values from CBSL Provision Comparison file PropertyMortgage sheet 
    to Supporting Schedule New Shcedule sheet.
    
    Mapping:
    - PropertyMortgage S2 → New Shcedule B80
    - PropertyMortgage S3 → New Shcedule B81
    
    Args:
        working_dir: Path to the working directory
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying PropertyMortgage data to New Shcedule sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find CBSL Provision file
        cbsl_file = find_cbsl_provision_file(monthly_folder)
        if not cbsl_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        logger.info(f"Reading .xlsb file: {cbsl_file.name}")
        
        # Read from CBSL Provision file PropertyMortgage sheet
        s2_value = None
        s3_value = None
        
        with open_workbook(str(cbsl_file)) as wb:
            sheet_names = wb.sheets
            logger.info(f"Available sheets: {sheet_names}")
            
            if "PropertyMortgage" not in sheet_names:
                logger.error(f"PropertyMortgage sheet not found. Available sheets: {sheet_names}")
                return False
            
            with wb.get_sheet("PropertyMortgage") as sheet:
                logger.info("Reading from PropertyMortgage sheet")
                
                # Convert to DataFrame for easier cell access
                data = []
                row_count = 0
                for row in sheet.rows():
                    data.append([cell.v for cell in row])
                    row_count += 1
                    if row_count >= 4:  # Only need rows 1-3 (indices 0-2)
                        break
                
                df = pd.DataFrame(data)
                
                logger.info(f"DataFrame shape: {df.shape}")
                
                # Get values from S2 and S3
                # S is column 18 (0-indexed), row 2 is index 1, row 3 is index 2
                s2_value = df.iloc[1, 18] if df.shape[0] > 1 and df.shape[1] > 18 else None
                s3_value = df.iloc[2, 18] if df.shape[0] > 2 and df.shape[1] > 18 else None
                
                logger.info(f"S2 (row 2, col S, index [1,18]) value: {s2_value}")
                logger.info(f"S3 (row 3, col S, index [2,18]) value: {s3_value}")
        
        # Write to Supporting Schedule file
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        ss_ws = ss_wb[ss_sheet_name]
        logger.info(f"Working with sheet: {ss_sheet_name}")
        
        # Clear existing values before pasting
        ss_ws['B80'].value = None
        ss_ws['B81'].value = None
        logger.info("Cleared cells B80 and B81")
        
        # Paste values
        if s2_value is not None:
            ss_ws['B80'].value = s2_value
            logger.info(f"Pasted S2 value ({s2_value}) to B80")
        else:
            logger.warning("S2 value is None, skipping paste to B80")
        
        if s3_value is not None:
            ss_ws['B81'].value = s3_value
            logger.info(f"Pasted S3 value ({s3_value}) to B81")
        else:
            logger.warning("S3 value is None, skipping paste to B81")
        
        # Save the workbook
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved PropertyMortgage data to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying PropertyMortgage data to New Shcedule: {e}", exc_info=True)
        return False


def update_b31_and_reorganize_cells(working_dir: Path, date_obj: datetime, 
                                   sheet_name: str = "New Shcedule") -> bool:
    """
    Update B31 with the relevant month date, then reorganize B40, B41, B44, B45.
    
    Steps:
    1. Update B31 to the date in m/d/yyyy format
    2. Clear B40 and B44
    3. Copy B41 → B40
    4. Copy B45 → B44
    5. Clear B41 and B45
    
    Args:
        working_dir: Path to the working directory
        date_obj: Date object to format for B31
        sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    # Format date as m/d/yyyy (e.g., "8/1/2025")
    formatted_date = f"{date_obj.month}/{date_obj.day}/{date_obj.year}"
    logger.info(f"Updating B31 and reorganizing cells for date: {formatted_date}")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        logger.info(f"Opening workbook: {ss_file.name}")
        wb = load_workbook(ss_file)
        
        if sheet_name not in wb.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}")
            wb.close()
            return False
        
        ws = wb[sheet_name]
        logger.info(f"Working with sheet: {sheet_name}")
        
        # Step 1: Update B31 with formatted date
        ws['B31'] = formatted_date
        logger.info(f"Updated B31 with date: {formatted_date}")
        
        # Step 2: Get values before clearing
        b41_value = ws['B41'].value
        b45_value = ws['B45'].value
        logger.info(f"Read B41 value: {b41_value}")
        logger.info(f"Read B45 value: {b45_value}")
        
        # Step 3: Clear B40 and B44
        ws['B40'].value = None
        ws['B44'].value = None
        logger.info("Cleared B40 and B44")
        
        # Step 4: Copy B41 → B40 and B45 → B44
        ws['B40'].value = b41_value
        ws['B44'].value = b45_value
        logger.info(f"Copied B41 ({b41_value}) → B40")
        logger.info(f"Copied B45 ({b45_value}) → B44")
        
        # Step 5: Clear B41 and B45
        ws['B41'].value = None
        ws['B45'].value = None
        logger.info("Cleared B41 and B45")
        
        # Save the workbook
        wb.save(ss_file)
        wb.close()
        logger.info(f"Successfully saved changes to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error updating B31 and reorganizing cells: {e}", exc_info=True)
        return False


def find_summary_file(monthly_folder: Path) -> Optional[Path]:
    """
    Find the Summary Excel file in the Supporting_Schedules subdirectory.
    """
    logger.info(f"Searching for Summary file in: {monthly_folder}")

    if not monthly_folder.exists():
        logger.error(f"Folder not found: {monthly_folder}")
        return None

    identifier = "Summary"
    found_files = []

    # Search for both .xlsx and .xlsb files
    for extension in ["*.xlsx", "*.xlsb"]:
        for file in monthly_folder.glob(extension):
            if file.name.startswith("~$"):
                continue
            if identifier in file.name:
                found_files.append(file)
                logger.info(f"Found matching file: {file.name}")

    if not found_files:
        logger.error(f"No file containing '{identifier}' found in {monthly_folder}")
        return None
    
    if len(found_files) > 1:
        logger.warning(f"Multiple files found, using first: {found_files[0].name}")
    
    summary_file = found_files[0]
    logger.info(f"Selected file: {summary_file}")
    return summary_file


def copy_from_summary_to_supporting_schedule(working_dir: Path, 
                                            ss_sheet_name: str = "New Shcedule") -> bool:
    """
    Copy values from Summary file to Supporting Schedule file.
    
    Steps:
    1. Copy 'Yard Provision' sheet F1 → Supporting Schedule B41
    2. Copy 'SUMMARY' sheet Z2 → Supporting Schedule B45
    
    Args:
        working_dir: Path to the working directory
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying values from Summary file to Supporting Schedule")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Read from Summary file
        logger.info(f"Opening Summary file: {summary_file.name}")
        
        yard_provision_value = None
        summary_value = None
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets in Summary file: {sheet_names}")
                
                # Get value from Yard Provision F1
                if "Yard Provision" in sheet_names:
                    with wb.get_sheet("Yard Provision") as sheet:
                        # F1 is row 0, column 5 (F is the 6th column, 0-indexed as 5)
                        data = []
                        for row in sheet.rows():
                            data.append([cell.v for cell in row])
                            if len(data) >= 1:  # Only need first row
                                break
                        
                        if data and len(data[0]) > 5:
                            yard_provision_value = data[0][5]
                            logger.info(f"Read 'Yard Provision' F1 value: {yard_provision_value}")
                        else:
                            logger.error("Could not read F1 from Yard Provision sheet")
                else:
                    logger.error(f"'Yard Provision' sheet not found. Available: {sheet_names}")
                    return False
                
                # Get value from SUMMARY Z2
                if "SUMMARY" in sheet_names:
                    with wb.get_sheet("SUMMARY") as sheet:
                        # Z2 is row 1, column 25 (Z is the 26th column, 0-indexed as 25)
                        data = []
                        row_count = 0
                        for row in sheet.rows():
                            data.append([cell.v for cell in row])
                            row_count += 1
                            if row_count >= 2:  # Need first 2 rows
                                break
                        
                        if len(data) > 1 and len(data[1]) > 25:
                            summary_value = data[1][25]
                            logger.info(f"Read 'SUMMARY' Z2 value: {summary_value}")
                        else:
                            logger.error("Could not read Z2 from SUMMARY sheet")
                else:
                    logger.error(f"'SUMMARY' sheet not found. Available: {sheet_names}")
                    return False
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            summary_wb = load_workbook(summary_file, data_only=True)
            
            # Get value from Yard Provision F1
            if "Yard Provision" in summary_wb.sheetnames:
                yard_sheet = summary_wb["Yard Provision"]
                yard_provision_value = yard_sheet['F1'].value
                logger.info(f"Read 'Yard Provision' F1 value: {yard_provision_value}")
            else:
                logger.error(f"'Yard Provision' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            # Get value from SUMMARY Z2
            if "SUMMARY" in summary_wb.sheetnames:
                summary_sheet = summary_wb["SUMMARY"]
                summary_value = summary_sheet['Z2'].value
                logger.info(f"Read 'SUMMARY' Z2 value: {summary_value}")
            else:
                logger.error(f"'SUMMARY' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            summary_wb.close()
        
        # Write to Supporting Schedule file
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        ss_ws = ss_wb[ss_sheet_name]
        
        # Clear existing values before pasting
        ss_ws['B41'].value = None
        ss_ws['B45'].value = None
        logger.info("Cleared cells B41 and B45")
        
        # Paste values
        ss_ws['B41'].value = yard_provision_value
        logger.info(f"Pasted 'Yard Provision' F1 value ({yard_provision_value}) to B41")
        
        ss_ws['B45'].value = summary_value
        logger.info(f"Pasted 'SUMMARY' Z2 value ({summary_value}) to B45")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved values to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying from Summary to Supporting Schedule: {e}", exc_info=True)
        return False


def copy_other_pivots_to_staging(working_dir: Path, 
                                 ss_sheet_name: str = "Staging") -> bool:
    """
    Copy values from Summary file 'Other Pivots' sheet to Supporting Schedule 'Staging' sheet.
    
    Mapping:
    - Other Pivots B4:B6 → Staging B4:B6
    - Other Pivots C4:C6 → Staging C4:C6
    - Other Pivots L4:L6 → Staging D4:D6
    - Other Pivots E4:E6 → Staging E4:E6
    - Other Pivots F4:F6 → Staging F4:F6
    
    Args:
        working_dir: Path to the working directory
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying Other Pivots data to Staging sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Read from Summary file
        logger.info(f"Opening Summary file: {summary_file.name}")
        
        # Dictionary to store the data ranges
        data_ranges = {
            'B': [],  # B4:B6
            'C': [],  # C4:C6
            'L': [],  # L4:L6
            'E': [],  # E4:E6
            'F': []   # F4:F6
        }
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets in Summary file: {sheet_names}")
                
                if "Other Pivots" not in sheet_names:
                    logger.error(f"'Other Pivots' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("Other Pivots") as sheet:
                    # Read rows 1-6 (indices 0-5) to get row 4-6 data (indices 3-5)
                    data = []
                    row_count = 0
                    for row in sheet.rows():
                        data.append([cell.v for cell in row])
                        row_count += 1
                        if row_count >= 6:  # Read up to row 6
                            break
                    
                    df = pd.DataFrame(data)
                    logger.info(f"Read {len(data)} rows from Other Pivots sheet")
                    
                    # Extract columns (0-indexed): B=1, C=2, E=4, F=5, L=11
                    # Rows 4-6 in Excel = indices 3-5
                    if len(data) >= 6:
                        data_ranges['B'] = [df.iloc[3, 1], df.iloc[4, 1], df.iloc[5, 1]] if df.shape[1] > 1 else [None, None, None]
                        data_ranges['C'] = [df.iloc[3, 2], df.iloc[4, 2], df.iloc[5, 2]] if df.shape[1] > 2 else [None, None, None]
                        data_ranges['E'] = [df.iloc[3, 4], df.iloc[4, 4], df.iloc[5, 4]] if df.shape[1] > 4 else [None, None, None]
                        data_ranges['F'] = [df.iloc[3, 5], df.iloc[4, 5], df.iloc[5, 5]] if df.shape[1] > 5 else [None, None, None]
                        data_ranges['L'] = [df.iloc[3, 11], df.iloc[4, 11], df.iloc[5, 11]] if df.shape[1] > 11 else [None, None, None]
                        
                        logger.info(f"B4:B6 values: {data_ranges['B']}")
                        logger.info(f"C4:C6 values: {data_ranges['C']}")
                        logger.info(f"L4:L6 values: {data_ranges['L']}")
                        logger.info(f"E4:E6 values: {data_ranges['E']}")
                        logger.info(f"F4:F6 values: {data_ranges['F']}")
                    else:
                        logger.error("Not enough rows in Other Pivots sheet")
                        return False
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            summary_wb = load_workbook(summary_file, data_only=True)
            
            if "Other Pivots" not in summary_wb.sheetnames:
                logger.error(f"'Other Pivots' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            op_sheet = summary_wb["Other Pivots"]
            
            # Read ranges
            data_ranges['B'] = [op_sheet['B4'].value, op_sheet['B5'].value, op_sheet['B6'].value]
            data_ranges['C'] = [op_sheet['C4'].value, op_sheet['C5'].value, op_sheet['C6'].value]
            data_ranges['L'] = [op_sheet['L4'].value, op_sheet['L5'].value, op_sheet['L6'].value]
            data_ranges['E'] = [op_sheet['E4'].value, op_sheet['E5'].value, op_sheet['E6'].value]
            data_ranges['F'] = [op_sheet['F4'].value, op_sheet['F5'].value, op_sheet['F6'].value]
            
            logger.info(f"B4:B6 values: {data_ranges['B']}")
            logger.info(f"C4:C6 values: {data_ranges['C']}")
            logger.info(f"L4:L6 values: {data_ranges['L']}")
            logger.info(f"E4:E6 values: {data_ranges['E']}")
            logger.info(f"F4:F6 values: {data_ranges['F']}")
            
            summary_wb.close()
        
        # Write to Supporting Schedule Staging sheet
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        staging_ws = ss_wb[ss_sheet_name]
        logger.info(f"Writing to sheet: {ss_sheet_name}")
        
        # Clear existing values before pasting
        for row_num in [4, 5, 6]:
            for col in ['B', 'C', 'D', 'E', 'F']:
                staging_ws[f'{col}{row_num}'].value = None
        logger.info("Cleared cells B4:F6")
        
        # Paste values according to mapping
        # B4:B6 → B4:B6
        staging_ws['B4'].value = data_ranges['B'][0]
        staging_ws['B5'].value = data_ranges['B'][1]
        staging_ws['B6'].value = data_ranges['B'][2]
        logger.info(f"Pasted B4:B6")
        
        # C4:C6 → C4:C6
        staging_ws['C4'].value = data_ranges['C'][0]
        staging_ws['C5'].value = data_ranges['C'][1]
        staging_ws['C6'].value = data_ranges['C'][2]
        logger.info(f"Pasted C4:C6")
        
        # L4:L6 → D4:D6
        staging_ws['D4'].value = data_ranges['L'][0]
        staging_ws['D5'].value = data_ranges['L'][1]
        staging_ws['D6'].value = data_ranges['L'][2]
        logger.info(f"Pasted L4:L6 to D4:D6")
        
        # E4:E6 → E4:E6
        staging_ws['E4'].value = data_ranges['E'][0]
        staging_ws['E5'].value = data_ranges['E'][1]
        staging_ws['E6'].value = data_ranges['E'][2]
        logger.info(f"Pasted E4:E6")
        
        # F4:F6 → F4:F6
        staging_ws['F4'].value = data_ranges['F'][0]
        staging_ws['F5'].value = data_ranges['F'][1]
        staging_ws['F6'].value = data_ranges['F'][2]
        logger.info(f"Pasted F4:F6")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved Other Pivots data to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying Other Pivots to Staging: {e}", exc_info=True)
        return False


def copy_summary_sheet_to_staging(working_dir: Path, date_obj: datetime,
                                  ss_sheet_name: str = "Staging") -> bool:
    """
    Copy values from Summary file 'SUMMARY' sheet to Supporting Schedule 'Staging' sheet.
    
    Mapping:
    - SUMMARY W88822 → Staging B10
    - SUMMARY X88822 → Staging C10
    - SUMMARY Y88822 → Staging D10
    - Update A10 with date in m/d/yyyy format
    
    Args:
        working_dir: Path to the working directory
        date_obj: Date object to format for A10
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    # Format date as m/d/yyyy (e.g., "8/1/2025")
    formatted_date = f"{date_obj.month}/{date_obj.day}/{date_obj.year}"
    logger.info(f"Copying SUMMARY sheet data to Staging sheet (Date: {formatted_date})")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Read from Summary file
        logger.info(f"Opening Summary file: {summary_file.name}")
        
        w88822_value = None
        x88822_value = None
        y88822_value = None
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets in Summary file: {sheet_names}")
                
                if "SUMMARY" not in sheet_names:
                    logger.error(f"'SUMMARY' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("SUMMARY") as sheet:
                    # W88822, X88822, Y88822
                    # W = column 22 (0-indexed), X = column 23, Y = column 24
                    # Row 88822 = index 88821 (0-indexed)
                    
                    logger.info("Reading row 88822 from SUMMARY sheet...")
                    data = []
                    row_count = 0
                    target_row_index = 88821  # 88822 in Excel = 88821 in 0-indexed
                    
                    for row in sheet.rows():
                        if row_count == target_row_index:
                            data.append([cell.v for cell in row])
                            break
                        row_count += 1
                    
                    if data and len(data[0]) > 24:
                        w88822_value = data[0][22]  # Column W (23rd column, 0-indexed as 22)
                        x88822_value = data[0][23]  # Column X (24th column, 0-indexed as 23)
                        y88822_value = data[0][24]  # Column Y (25th column, 0-indexed as 24)
                        
                        logger.info(f"W88822 value: {w88822_value}")
                        logger.info(f"X88822 value: {x88822_value}")
                        logger.info(f"Y88822 value: {y88822_value}")
                    else:
                        logger.error("Could not read row 88822 or columns W, X, Y")
                        return False
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            summary_wb = load_workbook(summary_file, data_only=True)
            
            if "SUMMARY" not in summary_wb.sheetnames:
                logger.error(f"'SUMMARY' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            summary_sheet = summary_wb["SUMMARY"]
            
            # Read values
            w88822_value = summary_sheet['W88822'].value
            x88822_value = summary_sheet['X88822'].value
            y88822_value = summary_sheet['Y88822'].value
            
            logger.info(f"W88822 value: {w88822_value}")
            logger.info(f"X88822 value: {x88822_value}")
            logger.info(f"Y88822 value: {y88822_value}")
            
            summary_wb.close()
        
        # Write to Supporting Schedule Staging sheet
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        staging_ws = ss_wb[ss_sheet_name]
        logger.info(f"Writing to sheet: {ss_sheet_name}")
        
        # Clear existing values before pasting
        staging_ws['A10'].value = None
        staging_ws['B10'].value = None
        staging_ws['C10'].value = None
        staging_ws['D10'].value = None
        logger.info("Cleared cells A10, B10, C10, D10")
        
        # Update A10 with date
        staging_ws['A10'].value = formatted_date
        logger.info(f"Updated A10 with date: {formatted_date}")
        
        # Paste values according to mapping
        staging_ws['B10'].value = w88822_value
        logger.info(f"Pasted W88822 ({w88822_value}) to B10")
        
        staging_ws['C10'].value = x88822_value
        logger.info(f"Pasted X88822 ({x88822_value}) to C10")
        
        staging_ws['D10'].value = y88822_value
        logger.info(f"Pasted Y88822 ({y88822_value}) to D10")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved SUMMARY data to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying SUMMARY sheet to Staging: {e}", exc_info=True)
        return False


def copy_other_pivots_row18_to_staging(working_dir: Path, 
                                       ss_sheet_name: str = "Staging") -> bool:
    """
    Copy values from Summary file 'Other Pivots' sheet row 18 to Supporting Schedule 'Staging' sheet.
    
    Mapping:
    - Other Pivots J18 → Staging B18
    - Other Pivots K18 → Staging C18
    - Other Pivots L18 → Staging D18
    - Other Pivots M18 → Staging E18
    - Other Pivots N18 → Staging F18
    
    Args:
        working_dir: Path to the working directory
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying Other Pivots row 18 data to Staging sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Read from Summary file
        logger.info(f"Opening Summary file: {summary_file.name}")
        
        # Variables to store the values
        j18_value = None
        k18_value = None
        l18_value = None
        m18_value = None
        n18_value = None
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets in Summary file: {sheet_names}")
                
                if "Other Pivots" not in sheet_names:
                    logger.error(f"'Other Pivots' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("Other Pivots") as sheet:
                    # Read rows 1-18 to get row 18 data (index 17)
                    data = []
                    row_count = 0
                    for row in sheet.rows():
                        data.append([cell.v for cell in row])
                        row_count += 1
                        if row_count >= 18:  # Read up to row 18
                            break
                    
                    # Extract columns (0-indexed): J=9, K=10, L=11, M=12, N=13
                    # Row 18 in Excel = index 17
                    if len(data) >= 18:
                        row18_data = data[17]
                        j18_value = row18_data[9] if len(row18_data) > 9 else None
                        k18_value = row18_data[10] if len(row18_data) > 10 else None
                        l18_value = row18_data[11] if len(row18_data) > 11 else None
                        m18_value = row18_data[12] if len(row18_data) > 12 else None
                        n18_value = row18_data[13] if len(row18_data) > 13 else None
                        
                        logger.info(f"J18 value: {j18_value}")
                        logger.info(f"K18 value: {k18_value}")
                        logger.info(f"L18 value: {l18_value}")
                        logger.info(f"M18 value: {m18_value}")
                        logger.info(f"N18 value: {n18_value}")
                    else:
                        logger.error("Not enough rows in Other Pivots sheet")
                        return False
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            summary_wb = load_workbook(summary_file, data_only=True)
            
            if "Other Pivots" not in summary_wb.sheetnames:
                logger.error(f"'Other Pivots' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            op_sheet = summary_wb["Other Pivots"]
            
            # Read values from row 18
            j18_value = op_sheet['J18'].value
            k18_value = op_sheet['K18'].value
            l18_value = op_sheet['L18'].value
            m18_value = op_sheet['M18'].value
            n18_value = op_sheet['N18'].value
            
            logger.info(f"J18 value: {j18_value}")
            logger.info(f"K18 value: {k18_value}")
            logger.info(f"L18 value: {l18_value}")
            logger.info(f"M18 value: {m18_value}")
            logger.info(f"N18 value: {n18_value}")
            
            summary_wb.close()
        
        # Write to Supporting Schedule Staging sheet
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        staging_ws = ss_wb[ss_sheet_name]
        logger.info(f"Writing to sheet: {ss_sheet_name}")
        
        # Clear existing values before pasting
        staging_ws['B18'].value = None
        staging_ws['C18'].value = None
        staging_ws['D18'].value = None
        staging_ws['E18'].value = None
        staging_ws['F18'].value = None
        logger.info("Cleared cells B18:F18")
        
        # Paste values according to mapping
        staging_ws['B18'].value = j18_value
        logger.info(f"Pasted J18 ({j18_value}) to B18")
        
        staging_ws['C18'].value = k18_value
        logger.info(f"Pasted K18 ({k18_value}) to C18")
        
        staging_ws['D18'].value = l18_value
        logger.info(f"Pasted L18 ({l18_value}) to D18")
        
        staging_ws['E18'].value = m18_value
        logger.info(f"Pasted M18 ({m18_value}) to E18")
        
        staging_ws['F18'].value = n18_value
        logger.info(f"Pasted N18 ({n18_value}) to F18")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved Other Pivots row 18 data to {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying Other Pivots row 18 to Staging: {e}", exc_info=True)
        return False


def copy_other_pivots_to_new_schedule(working_dir: Path, 
                                      ss_sheet_name: str = "New Shcedule") -> bool:
    """
    Copy values from Summary file 'Other Pivots' sheet to Supporting Schedule 'New Shcedule' sheet.
    
    Mapping:
    - Other Pivots L11 → New Shcedule B50
    - Other Pivots L12 → New Shcedule B51
    - Other Pivots B31 → New Shcedule C50
    - Other Pivots B32 → New Shcedule C51
    - Other Pivots E31 → New Shcedule D50
    - Other Pivots E32 → New Shcedule D51
    - Other Pivots O11 → New Shcedule F50
    - Other Pivots O12 → New Shcedule F51
    - Other Pivots M11 → New Shcedule G50
    - Other Pivots M12 → New Shcedule G51
    
    Args:
        working_dir: Path to the working directory
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying Other Pivots data to New Shcedule sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Read from Summary file
        logger.info(f"Opening Summary file: {summary_file.name}")
        
        # Variables to store the values
        l11_value = None
        l12_value = None
        b31_value = None
        b32_value = None
        e31_value = None
        e32_value = None
        o11_value = None
        o12_value = None
        m11_value = None
        m12_value = None
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets in Summary file: {sheet_names}")
                
                if "Other Pivots" not in sheet_names:
                    logger.error(f"'Other Pivots' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("Other Pivots") as sheet:
                    # Read rows 1-32 to get all needed data
                    data = []
                    row_count = 0
                    for row in sheet.rows():
                        data.append([cell.v for cell in row])
                        row_count += 1
                        if row_count >= 32:  # Read up to row 32
                            break
                    
                    # Extract values (0-indexed)
                    # Column indices: B=1, E=4, L=11, M=12, O=14
                    # Row indices: 11=10, 12=11, 31=30, 32=31
                    if len(data) >= 32:
                        # Row 11 (index 10)
                        if len(data[10]) > 11:
                            l11_value = data[10][11]  # L11
                        if len(data[10]) > 14:
                            o11_value = data[10][14]  # O11
                        if len(data[10]) > 12:
                            m11_value = data[10][12]  # M11
                        
                        # Row 12 (index 11)
                        if len(data) > 11 and len(data[11]) > 11:
                            l12_value = data[11][11]  # L12
                        if len(data) > 11 and len(data[11]) > 12:
                            m12_value = data[11][12]  # M12
                        if len(data) > 11 and len(data[11]) > 14:
                            o12_value = data[11][14]  # O12
                        
                        # Row 31 (index 30)
                        if len(data[30]) > 1:
                            b31_value = data[30][1]  # B31
                        if len(data[30]) > 4:
                            e31_value = data[30][4]  # E31
                        
                        # Row 32 (index 31)
                        if len(data[31]) > 1:
                            b32_value = data[31][1]  # B32
                        if len(data[31]) > 4:
                            e32_value = data[31][4]  # E32
                        
                        logger.info(f"L11 value: {l11_value}")
                        logger.info(f"L12 value: {l12_value}")
                        logger.info(f"B31 value: {b31_value}")
                        logger.info(f"B32 value: {b32_value}")
                        logger.info(f"E31 value: {e31_value}")
                        logger.info(f"E32 value: {e32_value}")
                        logger.info(f"O11 value: {o11_value}")
                        logger.info(f"O12 value: {o12_value}")
                        logger.info(f"M11 value: {m11_value}")
                        logger.info(f"M12 value: {m12_value}")
                    else:
                        logger.error("Not enough rows in Other Pivots sheet")
                        return False
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            summary_wb = load_workbook(summary_file, data_only=True)
            
            if "Other Pivots" not in summary_wb.sheetnames:
                logger.error(f"'Other Pivots' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            op_sheet = summary_wb["Other Pivots"]
            
            # Read values
            l11_value = op_sheet['L11'].value
            l12_value = op_sheet['L12'].value
            b31_value = op_sheet['B31'].value
            b32_value = op_sheet['B32'].value
            e31_value = op_sheet['E31'].value
            e32_value = op_sheet['E32'].value
            o11_value = op_sheet['O11'].value
            o12_value = op_sheet['O12'].value
            m11_value = op_sheet['M11'].value
            m12_value = op_sheet['M12'].value
            
            logger.info(f"L11 value: {l11_value}")
            logger.info(f"L12 value: {l12_value}")
            logger.info(f"B31 value: {b31_value}")
            logger.info(f"B32 value: {b32_value}")
            logger.info(f"E31 value: {e31_value}")
            logger.info(f"E32 value: {e32_value}")
            logger.info(f"O11 value: {o11_value}")
            logger.info(f"O12 value: {o12_value}")
            logger.info(f"M11 value: {m11_value}")
            logger.info(f"M12 value: {m12_value}")
            
            summary_wb.close()
        
        # Write to Supporting Schedule New Shcedule sheet
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        new_schedule_ws = ss_wb[ss_sheet_name]
        logger.info(f"Writing to sheet: {ss_sheet_name}")
        
        # Clear existing values before pasting
        new_schedule_ws['B50'].value = None
        new_schedule_ws['B51'].value = None
        new_schedule_ws['C50'].value = None
        new_schedule_ws['C51'].value = None
        new_schedule_ws['D50'].value = None
        new_schedule_ws['D51'].value = None
        new_schedule_ws['F50'].value = None
        new_schedule_ws['F51'].value = None
        new_schedule_ws['G50'].value = None
        new_schedule_ws['G51'].value = None
        logger.info("Cleared cells B50, B51, C50, C51, D50, D51, F50, F51, G50, G51")
        
        # Paste values according to mapping
        new_schedule_ws['B50'].value = l11_value
        logger.info(f"Pasted L11 ({l11_value}) to B50")
        
        new_schedule_ws['B51'].value = l12_value
        logger.info(f"Pasted L12 ({l12_value}) to B51")
        
        new_schedule_ws['C50'].value = b31_value
        logger.info(f"Pasted B31 ({b31_value}) to C50")
        
        new_schedule_ws['C51'].value = b32_value
        logger.info(f"Pasted B32 ({b32_value}) to C51")
        
        new_schedule_ws['D50'].value = e31_value
        logger.info(f"Pasted E31 ({e31_value}) to D50")
        
        new_schedule_ws['D51'].value = e32_value
        logger.info(f"Pasted E32 ({e32_value}) to D51")
        
        new_schedule_ws['F50'].value = o11_value
        logger.info(f"Pasted O11 ({o11_value}) to F50")
        
        new_schedule_ws['F51'].value = o12_value
        logger.info(f"Pasted O12 ({o12_value}) to F51")
        
        new_schedule_ws['G50'].value = m11_value
        logger.info(f"Pasted M11 ({m11_value}) to G50")
        
        new_schedule_ws['G51'].value = m12_value
        logger.info(f"Pasted M12 ({m12_value}) to G51")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved Other Pivots data to New Shcedule sheet in {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying Other Pivots to New Shcedule: {e}", exc_info=True)
        return False


def copy_staging_to_new_schedule(working_dir: Path,
                                 staging_sheet_name: str = "Staging",
                                 new_schedule_sheet_name: str = "New Shcedule") -> bool:
    """
    Copy values from Supporting Schedule 'Staging' sheet to 'New Shcedule' sheet.
    
    Mapping:
    - Staging C10 → New Shcedule B52
    - Staging B10 → New Shcedule B53
    
    Args:
        working_dir: Path to the working directory
        staging_sheet_name: Name of the Staging sheet
        new_schedule_sheet_name: Name of the New Shcedule sheet
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying Staging sheet data to New Shcedule sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Open Supporting Schedule file
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        # Check if both sheets exist
        if staging_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{staging_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        if new_schedule_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{new_schedule_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        # Get both sheets
        staging_ws = ss_wb[staging_sheet_name]
        new_schedule_ws = ss_wb[new_schedule_sheet_name]
        
        # Read values from Staging sheet
        c10_value = staging_ws['C10'].value
        b10_value = staging_ws['B10'].value
        
        logger.info(f"Read from Staging - C10: {c10_value}")
        logger.info(f"Read from Staging - B10: {b10_value}")
        
        # Clear existing values before pasting
        new_schedule_ws['B52'].value = None
        new_schedule_ws['B53'].value = None
        logger.info("Cleared cells B52 and B53")
        
        # Paste values to New Shcedule sheet
        new_schedule_ws['B52'].value = c10_value
        logger.info(f"Pasted C10 ({c10_value}) to B52")
        
        new_schedule_ws['B53'].value = b10_value
        logger.info(f"Pasted B10 ({b10_value}) to B53")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved Staging data to New Shcedule sheet in {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying Staging to New Shcedule: {e}", exc_info=True)
        return False


def copy_summary_sum_to_new_schedule(working_dir: Path,
                                     ss_sheet_name: str = "New Shcedule") -> bool:
    """
    Copy and sum values from Summary file 'SUMMARY' sheet to Supporting Schedule 'New Shcedule' sheet.
    
    Calculation:
    - Sum = SUMMARY X88805 + SUMMARY Y88800
    - Paste sum to New Shcedule B56
    
    Args:
        working_dir: Path to the working directory
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying and summing SUMMARY sheet data to New Shcedule sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Read from Summary file
        logger.info(f"Opening Summary file: {summary_file.name}")
        
        x88805_value = None
        y88800_value = None
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets in Summary file: {sheet_names}")
                
                if "SUMMARY" not in sheet_names:
                    logger.error(f"'SUMMARY' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("SUMMARY") as sheet:
                    # Read rows up to 88805 to get the data
                    # For efficiency, we'll read row by row until we reach the needed rows
                    target_rows = {88805, 88800}
                    data = {}
                    row_count = 0
                    
                    for row in sheet.rows():
                        row_count += 1
                        if row_count in target_rows:
                            row_data = [cell.v for cell in row]
                            data[row_count] = row_data
                        
                        if row_count >= 88805:  # Stop after reading row 88805
                            break
                    
                    # Extract values (0-indexed): X=23 (24th column), Y=24 (25th column)
                    if 88805 in data and len(data[88805]) > 23:
                        x88805_value = data[88805][23]  # Column X (24th column, 0-indexed as 23)
                        logger.info(f"X88805 value: {x88805_value}")
                    else:
                        logger.error("Could not read row 88805 or column X")
                        return False
                    
                    if 88800 in data and len(data[88800]) > 24:
                        y88800_value = data[88800][24]  # Column Y (25th column, 0-indexed as 24)
                        logger.info(f"Y88800 value: {y88800_value}")
                    else:
                        logger.error("Could not read row 88800 or column Y")
                        return False
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            summary_wb = load_workbook(summary_file, data_only=True)
            
            if "SUMMARY" not in summary_wb.sheetnames:
                logger.error(f"'SUMMARY' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            summary_sheet = summary_wb["SUMMARY"]
            
            # Read values
            x88805_value = summary_sheet['X88805'].value
            y88800_value = summary_sheet['Y88800'].value
            
            logger.info(f"X88805 value: {x88805_value}")
            logger.info(f"Y88800 value: {y88800_value}")
            
            summary_wb.close()
        
        # Calculate sum
        if x88805_value is None or y88800_value is None:
            logger.error(f"Cannot calculate sum - X88805: {x88805_value}, Y88800: {y88800_value}")
            return False
        
        sum_value = x88805_value + y88800_value
        logger.info(f"Sum calculated: {x88805_value} + {y88800_value} = {sum_value}")
        
        # Write to Supporting Schedule New Shcedule sheet
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        new_schedule_ws = ss_wb[ss_sheet_name]
        logger.info(f"Writing to sheet: {ss_sheet_name}")
        
        # Clear existing value before pasting
        new_schedule_ws['B56'].value = None
        logger.info("Cleared cell B56")
        
        # Paste sum to B56
        new_schedule_ws['B56'].value = sum_value
        logger.info(f"Pasted sum ({sum_value}) to B56")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved SUMMARY sum to New Shcedule sheet in {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying SUMMARY sum to New Shcedule: {e}", exc_info=True)
        return False


def copy_summary_range_to_new_schedule(working_dir: Path,
                                       ss_sheet_name: str = "New Shcedule") -> bool:
    """
    Copy range from Summary file 'SUMMARY' sheet to Supporting Schedule 'New Shcedule' sheet.
    
    Mapping:
    - SUMMARY X88808:X88813 → New Shcedule B60:B65
    
    Args:
        working_dir: Path to the working directory
        ss_sheet_name: Name of the sheet in Supporting Schedule file
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying SUMMARY sheet range X88808:X88813 to New Shcedule sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Read from Summary file
        logger.info(f"Opening Summary file: {summary_file.name}")
        
        # Store values for rows 88808-88813, column X only
        x_values = []
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets in Summary file: {sheet_names}")
                
                if "SUMMARY" not in sheet_names:
                    logger.error(f"'SUMMARY' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("SUMMARY") as sheet:
                    # Read rows 88808-88813
                    target_rows = list(range(88808, 88814))  # 88808 to 88813 inclusive
                    data = {}
                    row_count = 0
                    
                    for row in sheet.rows():
                        row_count += 1
                        if row_count in target_rows:
                            row_data = [cell.v for cell in row]
                            data[row_count] = row_data
                        
                        if row_count >= 88813:  # Stop after reading row 88813
                            break
                    
                    # Extract values (0-indexed): X=23
                    for row_num in target_rows:
                        if row_num in data:
                            if len(data[row_num]) > 23:
                                x_values.append(data[row_num][23])  # Column X
                            else:
                                x_values.append(None)
                        else:
                            x_values.append(None)
                    
                    logger.info(f"X column values (88808-88813): {x_values}")
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            summary_wb = load_workbook(summary_file, data_only=True)
            
            if "SUMMARY" not in summary_wb.sheetnames:
                logger.error(f"'SUMMARY' sheet not found. Available: {summary_wb.sheetnames}")
                summary_wb.close()
                return False
            
            summary_sheet = summary_wb["SUMMARY"]
            
            # Read values from X88808:X88813
            for row_num in range(88808, 88814):
                x_values.append(summary_sheet[f'X{row_num}'].value)
            
            logger.info(f"X column values (88808-88813): {x_values}")
            
            summary_wb.close()
        
        # Write to Supporting Schedule New Shcedule sheet
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        if ss_sheet_name not in ss_wb.sheetnames:
            logger.error(f"Sheet '{ss_sheet_name}' not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        new_schedule_ws = ss_wb[ss_sheet_name]
        logger.info(f"Writing to sheet: {ss_sheet_name}")
        
        # Clear existing values before pasting
        target_rows = [60, 61, 62, 63, 64, 65]
        for row_num in target_rows:
            new_schedule_ws[f'B{row_num}'].value = None
        logger.info("Cleared cells B60:B65")
        
        # Paste values to B60:B65
        for i, row_num in enumerate(target_rows):
            # Column B gets X values
            new_schedule_ws[f'B{row_num}'].value = x_values[i]
            logger.info(f"Pasted X{88808+i} ({x_values[i]}) to B{row_num}")
        
        # Save
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved SUMMARY range to New Shcedule sheet in {ss_file.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error copying SUMMARY range to New Shcedule: {e}", exc_info=True)
        return False


def export_summary_data_to_excel(working_dir: Path) -> bool:
    """
    Read SUMMARY sheet data from Summary file and paste to Supporting Schedule RPT List sheet.
    
    Reads data from row 3 onwards, columns A to AG, until column A "CONTRACT NO" is empty.
    Filters to only include rows where RELATED-PARTY CONTRACT column contains "RPT CONTRACT".
    Clears existing data in RPT List sheet and pastes filtered values starting from row 3, column B.
    
    Args:
        working_dir: Path to the working directory
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Copying filtered SUMMARY data to RPT List sheet")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Summary file
        summary_file = find_summary_file(monthly_folder)
        if not summary_file:
            return False
        
        logger.info(f"Reading SUMMARY sheet from: {summary_file.name}")
        
        # Read from Summary file
        df_summary = None
        
        # Check if it's a .xlsb file
        if summary_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(summary_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets: {sheet_names}")
                
                if "SUMMARY" not in sheet_names:
                    logger.error(f"'SUMMARY' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("SUMMARY") as sheet:
                    # Read all data from the sheet
                    data = []
                    for row in sheet.rows():
                        data.append([cell.v for cell in row])
                    
                    logger.info(f"Read {len(data)} rows from SUMMARY sheet")
                    
                    # Convert to DataFrame
                    df_full = pd.DataFrame(data)
                    
                    # Extract data from row 3 onwards (row 3 = index 2), columns A to AG (0 to 32)
                    # Row 3 should contain headers
                    if len(df_full) < 3:
                        logger.error("Not enough rows in SUMMARY sheet")
                        return False
                    
                    # Get headers from row 3 (index 2)
                    headers = df_full.iloc[2, :33].tolist()  # Columns A to AG (0-32)
                    logger.info(f"Headers from row 3: {headers}")
                    
                    # Get data starting from row 4 (index 3)
                    data_rows = []
                    for idx in range(3, len(df_full)):
                        row_data = df_full.iloc[idx, :33].tolist()  # Columns A to AG
                        
                        # Check if column A (CONTRACT NO) is empty
                        if pd.isna(row_data[0]) or row_data[0] == '' or row_data[0] is None:
                            logger.info(f"Found empty CONTRACT NO at row {idx + 1} (Excel row). Stopping data collection.")
                            break
                        
                        data_rows.append(row_data)
                    
                    logger.info(f"Collected {len(data_rows)} data rows")
                    
                    # Create DataFrame with headers
                    df_summary = pd.DataFrame(data_rows, columns=headers)
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            
            # Read the entire sheet starting from row 3
            df_full = pd.read_excel(summary_file, sheet_name="SUMMARY", engine='openpyxl', 
                                   header=2, usecols="A:AG")  # Row 3 as header, columns A to AG
            
            logger.info(f"Read {len(df_full)} rows from SUMMARY sheet")
            logger.info(f"Columns: {df_full.columns.tolist()}")
            
            # Find the first row where CONTRACT NO is empty/NaN
            contract_no_col = df_full.columns[0]  # First column should be CONTRACT NO
            
            # Find first empty row in CONTRACT NO column
            empty_row_idx = None
            for idx, value in enumerate(df_full[contract_no_col]):
                if pd.isna(value) or value == '' or value is None:
                    empty_row_idx = idx
                    logger.info(f"Found empty CONTRACT NO at row {idx + 4} (Excel row). Stopping data collection.")
                    break
            
            # Slice dataframe up to the empty row
            if empty_row_idx is not None:
                df_summary = df_full.iloc[:empty_row_idx].copy()
            else:
                df_summary = df_full.copy()
                logger.info("No empty CONTRACT NO found, using all data")
            
            logger.info(f"Final dataframe has {len(df_summary)} rows")
        
        if df_summary is None or len(df_summary) == 0:
            logger.error("No data extracted from SUMMARY sheet")
            return False
        
        logger.info(f"Total rows before filtering: {len(df_summary)}")
        
        # Filter to keep only rows where RELATED-PARTY CONTRACT column has "RPT CONTRACT"
        # Find the column that contains "RELATED-PARTY CONTRACT" or similar
        rpt_col = None
        for col in df_summary.columns:
            col_str = str(col).strip().upper()
            if "RELATED" in col_str and "PARTY" in col_str and "CONTRACT" in col_str:
                rpt_col = col
                logger.info(f"Found RELATED-PARTY CONTRACT column: {col}")
                break
        
        if rpt_col is None:
            logger.error(f"RELATED-PARTY CONTRACT column not found. Available columns: {df_summary.columns.tolist()}")
            return False
        
        # Filter rows where the RELATED-PARTY CONTRACT column contains "RPT CONTRACT"
        df_filtered = df_summary[df_summary[rpt_col].astype(str).str.contains("RPT CONTRACT", case=False, na=False)].copy()
        
        logger.info(f"Rows after filtering for 'RPT CONTRACT': {len(df_filtered)}")
        
        if len(df_filtered) == 0:
            logger.warning("No rows found with 'RPT CONTRACT' in RELATED-PARTY CONTRACT column")
            logger.info(f"Sample values from {rpt_col}: {df_summary[rpt_col].head(10).tolist()}")
        
        # Find Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Open Supporting Schedule file
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        # Find the RPT List sheet (using keyword matching)
        rpt_list_sheet_name = None
        for sheet_name in ss_wb.sheetnames:
            if "RPT" in sheet_name.upper() and "LIST" in sheet_name.upper():
                rpt_list_sheet_name = sheet_name
                logger.info(f"Found RPT List sheet: {sheet_name}")
                break
        
        if rpt_list_sheet_name is None:
            logger.error(f"RPT List sheet not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        rpt_ws = ss_wb[rpt_list_sheet_name]
        logger.info(f"Writing to sheet: {rpt_list_sheet_name}")
        
        # Clear existing data from row 3 onwards (keep headers in rows 1-2)
        logger.info("Clearing existing data from row 3 onwards")
        max_row = rpt_ws.max_row
        max_col = rpt_ws.max_column
        
        # Delete all rows from row 3 onwards
        if max_row >= 3:
            rpt_ws.delete_rows(3, max_row - 2)
            logger.info(f"Cleared {max_row - 2} rows")
        
        # Paste filtered data values (without headers) starting from row 3, column B
        logger.info(f"Pasting {len(df_filtered)} rows of data starting from row 3, column B")
        
        for row_idx, (_, row_data) in enumerate(df_filtered.iterrows(), start=3):
            for col_idx, value in enumerate(row_data, start=2):  # Start from column B (column 2)
                rpt_ws.cell(row=row_idx, column=col_idx, value=value)
        
        logger.info(f"Successfully pasted {len(df_filtered)} RPT CONTRACT rows to {rpt_list_sheet_name}")
        
        # Save the workbook
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully saved data to {ss_file.name}")
        
        print(f"Data pasted to RPT List sheet: {len(df_filtered)} RPT CONTRACT rows")
        
        return True
        
    except Exception as e:
        logger.error(f"Error exporting SUMMARY data to Excel: {e}", exc_info=True)
        return False


def fill_rpt_list_client_names(working_dir: Path, date_obj: datetime) -> bool:
    """
    Fill CLIENT NAME column (A) in RPT List sheet by matching CONTRACT NO from Portfolio file.
    Then rename the sheet to include month and year (e.g., "RPT List-May 25").
    
    Steps:
    1. Find Portfolio_4_ECL file
    2. Read Portfolio sheet columns D (CLIENT NAME) and E (CONTRACT NO) from row 3 onwards
    3. Match CONTRACT NO between Portfolio and RPT List
    4. Fill column A in RPT List with matching CLIENT NAME
    5. Rename sheet to "RPT List-{Month} {YY}"
    
    Args:
        working_dir: Path to the working directory
        date_obj: Date object for formatting the sheet name
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Filling CLIENT NAME in RPT List sheet from Portfolio file")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find Portfolio_4_ECL file directly in the date folder
        logger.info(f"Searching for Portfolio_4_ECL file in: {monthly_folder}")

        portfolio_file = None
        for extension in ["*.xlsx", "*.xlsb"]:
            for file in monthly_folder.glob(extension):
                if "Portfolio_4_ECL" in file.name:
                    portfolio_file = file
                    logger.info(f"Found Portfolio file: {file.name}")
                    break
            if portfolio_file:
                break
        
        if not portfolio_file:
            logger.error("Portfolio_4_ECL file not found")
            return False
        
        # Read Portfolio sheet columns D and E starting from row 3
        logger.info(f"Reading Portfolio sheet from: {portfolio_file.name}")
        
        # Dictionary to store CONTRACT NO -> CLIENT NAME mapping
        contract_to_client = {}
        
        if portfolio_file.suffix.lower() == '.xlsb':
            logger.info("Reading .xlsb file using pyxlsb")
            with open_workbook(str(portfolio_file)) as wb:
                sheet_names = wb.sheets
                logger.info(f"Available sheets: {sheet_names}")
                
                if "Portfolio" not in sheet_names:
                    logger.error(f"'Portfolio' sheet not found. Available: {sheet_names}")
                    return False
                
                with wb.get_sheet("Portfolio") as sheet:
                    # Read all data from the sheet
                    data = []
                    for row in sheet.rows():
                        data.append([cell.v for cell in row])
                    
                    logger.info(f"Read {len(data)} rows from Portfolio sheet")
                    
                    # Start from row 3 (index 2), get columns D (index 3) and E (index 4)
                    for idx in range(2, len(data)):
                        row = data[idx]
                        if len(row) > 4:
                            client_name = row[3]  # Column D
                            contract_no = row[4]  # Column E
                            
                            # Skip if contract number is empty
                            if contract_no is not None and contract_no != '':
                                contract_to_client[str(contract_no).strip()] = client_name
                    
                    logger.info(f"Loaded {len(contract_to_client)} contract mappings")
        else:
            # Use openpyxl for .xlsx files
            logger.info("Reading .xlsx file using openpyxl")
            portfolio_wb = load_workbook(portfolio_file, data_only=True)
            
            if "Portfolio" not in portfolio_wb.sheetnames:
                logger.error(f"'Portfolio' sheet not found. Available: {portfolio_wb.sheetnames}")
                portfolio_wb.close()
                return False
            
            portfolio_sheet = portfolio_wb["Portfolio"]
            
            # Read from row 3 onwards, columns D and E
            for row_idx in range(3, portfolio_sheet.max_row + 1):
                client_name = portfolio_sheet.cell(row=row_idx, column=4).value  # Column D
                contract_no = portfolio_sheet.cell(row=row_idx, column=5).value   # Column E
                
                # Skip if contract number is empty
                if contract_no is not None and contract_no != '':
                    contract_to_client[str(contract_no).strip()] = client_name
            
            logger.info(f"Loaded {len(contract_to_client)} contract mappings")
            portfolio_wb.close()
        
        # Now open the Supporting Schedule file and fill RPT List column A
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        logger.info(f"Opening Supporting Schedule file: {ss_file.name}")
        ss_wb = load_workbook(ss_file)
        
        # Find the RPT List sheet
        rpt_list_sheet_name = None
        for sheet_name in ss_wb.sheetnames:
            if "RPT" in sheet_name.upper() and "LIST" in sheet_name.upper():
                rpt_list_sheet_name = sheet_name
                logger.info(f"Found RPT List sheet: {sheet_name}")
                break
        
        if rpt_list_sheet_name is None:
            logger.error(f"RPT List sheet not found. Available sheets: {ss_wb.sheetnames}")
            ss_wb.close()
            return False
        
        rpt_ws = ss_wb[rpt_list_sheet_name]
        
        # Fill column A with client names by matching contract numbers
        # Data starts at row 3, column B has the first data column (contract no should be there)
        # We need to find which column has CONTRACT NO
        
        matched_count = 0
        unmatched_count = 0
        
        # Start from row 3 and iterate through all rows
        for row_idx in range(3, rpt_ws.max_row + 1):
            # Get contract number from column B (first data column after skipping A)
            contract_no_cell = rpt_ws.cell(row=row_idx, column=2).value  # Column B
            
            if contract_no_cell is None or contract_no_cell == '':
                break  # Stop at first empty row
            
            contract_no_str = str(contract_no_cell).strip()
            
            # Look up the client name
            if contract_no_str in contract_to_client:
                client_name = contract_to_client[contract_no_str]
                rpt_ws.cell(row=row_idx, column=1, value=client_name)  # Column A
                matched_count += 1
            else:
                unmatched_count += 1
                logger.warning(f"No match found for contract: {contract_no_str}")
        
        logger.info(f"Matched {matched_count} contracts, {unmatched_count} unmatched")
        
        # Rename the RPT List sheet with month and year
        # Format: "RPT List-May 25" for May 2025
        month_name = date_obj.strftime("%B")  # Full month name
        year_short = date_obj.strftime("%y")  # 2-digit year
        new_sheet_name = f"RPT List-{month_name} {year_short}"
        
        logger.info(f"Renaming sheet '{rpt_list_sheet_name}' to '{new_sheet_name}'")
        rpt_ws.title = new_sheet_name
        
        # Save the workbook
        ss_wb.save(ss_file)
        ss_wb.close()
        logger.info(f"Successfully filled CLIENT NAME and renamed sheet to: {new_sheet_name}")
        
        print(f"CLIENT NAME filled: {matched_count} matched, {unmatched_count} unmatched")
        print(f"RPT List sheet renamed to: {new_sheet_name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error filling RPT List client names: {e}", exc_info=True)
        return False


def rename_supporting_schedule_file(working_dir: Path, date_obj: datetime) -> bool:
    """
    Rename the Supporting Schedule file with the month/year from the data.
    Format: "Supporting Schedules - Aug 2025.xlsx"
    
    Args:
        working_dir: Path to the working directory
        date_obj: Date object containing the month/year to use
    
    Returns:
        True if successful, False otherwise
    """
    logger.info("Renaming Supporting Schedule file with month/year")
    
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False
        
        # Find current Supporting Schedule file
        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False
        
        # Generate new filename
        new_file_path = get_new_supporting_schedule_filename(ss_file, date_obj)
        
        # Check if already has the correct name
        if ss_file == new_file_path:
            logger.info(f"File already has the correct name: {ss_file.name}")
            return True
        
        # Rename the file
        if new_file_path.exists():
            logger.info(f"Target file already exists, deleting: {new_file_path.name}")
            new_file_path.unlink()
        
        ss_file.rename(new_file_path)
        logger.info(f"Successfully renamed {ss_file.name} to {new_file_path.name}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error renaming Supporting Schedule file: {e}", exc_info=True)
        return False


def kill_excel() -> None:
    """Terminate all running Excel processes to release file locks."""
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "EXCEL.EXE"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        logger.info("Excel processes terminated.")
    except Exception as e:
        logger.warning(f"taskkill failed (Excel may not have been running): {e}")


def clear_new_schedule_for_april(working_dir: Path, sheet_name: str = "New Shcedule") -> bool:
    """
    Clear the ranges B5:M13, B15:M20, B22:M26 in the 'New Shcedule' sheet of the
    Supporting Schedule file. Called only when the reporting month is April.
    """
    logger.info("April detected — clearing New Shcedule ranges before processing.")
    try:
        monthly_folder = get_monthly_folder(working_dir)
        if not monthly_folder:
            return False

        ss_file = find_supporting_schedule_file(monthly_folder)
        if not ss_file:
            return False

        kill_excel()
        logger.info(f"Opening {ss_file.name} to clear April ranges.")
        wb = load_workbook(ss_file)

        if sheet_name not in wb.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found in {ss_file.name}")
            wb.close()
            return False

        ws = wb[sheet_name]

        ranges_to_clear = ["B5:M13", "B15:M20", "B22:M26"]
        for rng in ranges_to_clear:
            for row in ws[rng]:
                for cell in row:
                    cell.value = None
            logger.info(f"Cleared range {rng}")

        wb.save(ss_file)
        wb.close()
        logger.info("April cleanup completed successfully.")
        return True

    except Exception as e:
        logger.error(f"Error during April cleanup: {e}", exc_info=True)
        return False


# Example usage
if __name__ == "__main__":
    import sys

    script_dir = Path(__file__).parent.parent
    working_dir = script_dir / "working" / "Supporting Schedules"

    logger.info(f"Working directory: {working_dir}")

    # Parse date from command-line argument (DD/MM/YYYY)
    print("\n" + "="*70)
    print("STEP 1/15: DETERMINING MONTH AND DATE FROM INPUT")
    print("="*70)
    if len(sys.argv) < 2:
        print("Usage: python Supporting_Schedules.py DD/MM/YYYY")
        exit(1)

    try:
        input_date = datetime.strptime(sys.argv[1], "%d/%m/%Y")
    except ValueError:
        print(f"Invalid date format '{sys.argv[1]}'. Expected DD/MM/YYYY (e.g. 31/08/2025).")
        exit(1)

    # Use the first day of the given month as the reference date
    date_obj = input_date.replace(day=1)
    month_name = date_obj.strftime("%B")  # Full month name, e.g. 'August'

    logger.info(f"Input date: {input_date.strftime('%d/%m/%Y')}")
    logger.info(f"Month start date: {date_obj.strftime('%d/%m/%Y')}")
    logger.info(f"Month name: {month_name}")
    print(f"INPUT DATE:  {input_date.strftime('%d/%m/%Y')}")
    print(f"MONTH START: {date_obj.strftime('%d/%m/%Y')}")
    print(f"MONTH NAME:  {month_name}")

    # If April, clear previous year's data from New Shcedule before processing
    if month_name == "April":
        print("\n" + "="*70)
        print("STEP 1a: APRIL CLEANUP — CLEARING NEW SHCEDULE RANGES")
        print("="*70)
        success = clear_new_schedule_for_april(working_dir)
        if success:
            print("APRIL CLEANUP COMPLETED SUCCESSFULLY")
        else:
            print("Failed to clear April ranges. Check logs for details.")
            exit(1)

    # Get the data
    print("\n" + "="*70)
    print("STEP 2/15: LOADING GL ACCOUNT DATA")
    print("="*70)
    df = get_gl_account_data(working_dir)
    
    if df is not None:
        print("GL ACCOUNT DETAILS - DATA LOADED SUCCESSFULLY")
        print(df.to_string(index=True))
    else:
        print("Failed to load data. Check logs for details.")
        exit(1)
    
    # Paste data to Supporting Schedule
    print("\n" + "="*70)
    print("STEP 3/15: PASTING GL DATA TO SUPPORTING SCHEDULE")
    print("="*70)
    success = paste_data_to_supporting_schedule(working_dir, df, month_name)
    
    if success:
        print("DATA PASTED TO SUPPORTING SCHEDULE SUCCESSFULLY")
    else:
        print("Failed to paste data. Check logs for details.")
        exit(1)
    
    # Get CBSL Provision values
    print("\n" + "="*70)
    print("STEP 4/15: EXTRACTING CBSL PROVISION VALUES")
    print("="*70)
    cbsl_values = get_cbsl_provision_values(working_dir)
    
    if cbsl_values:
        print("CBSL PROVISION VALUES EXTRACTED")
        print(f"  J11: {cbsl_values['J11']}")
        print(f"  J12: {cbsl_values['J12']}")
        print(f"  L14: {cbsl_values['L14']}")
    else:
        print("Failed to extract CBSL Provision values. Check logs for details.")
        exit(1)
    
    # Paste CBSL Provision values to Supporting Schedule
    print("\n" + "="*70)
    print("STEP 5/18: PASTING CBSL PROVISION DATA TO SUPPORTING SCHEDULE")
    print("="*70)
    success = paste_cbsl_provision_to_supporting_schedule(working_dir, cbsl_values, month_name)
    
    if success:
        print("CBSL PROVISION DATA PASTED SUCCESSFULLY")
    else:
        print("Failed to paste CBSL Provision data. Check logs for details.")
        exit(1)
    
    # Copy PropertyMortgage values to New Shcedule
    print("\n" + "="*70)
    print("STEP 6/18: COPYING PROPERTY MORTGAGE DATA TO NEW SCHEDULE")
    print("="*70)
    success = copy_property_mortgage_to_new_schedule(working_dir)
    
    if success:
        print("PROPERTY MORTGAGE DATA COPIED SUCCESSFULLY")
    else:
        print("Failed to copy PropertyMortgage data. Check logs for details.")
        exit(1)
    
    # Update B31 and reorganize cells
    print("\n" + "="*70)
    print("STEP 7/18: UPDATING B31 AND REORGANIZING CELLS")
    print("="*70)
    success = update_b31_and_reorganize_cells(working_dir, date_obj)
    
    if success:
        print("B31 UPDATED AND CELLS REORGANIZED SUCCESSFULLY")
    else:
        print("Failed to update B31 and reorganize cells. Check logs for details.")
        exit(1)
    
    # Copy values from Summary file to Supporting Schedule
    print("\n" + "="*70)
    print("STEP 8/18: COPYING SUMMARY VALUES TO SUPPORTING SCHEDULE")
    print("="*70)
    success = copy_from_summary_to_supporting_schedule(working_dir)
    
    if success:
        print("SUMMARY VALUES COPIED SUCCESSFULLY")
    else:
        print("Failed to copy Summary values. Check logs for details.")
        exit(1)
    
    # Copy Other Pivots data to Staging sheet
    print("\n" + "="*70)
    print("STEP 9/18: COPYING OTHER PIVOTS DATA TO STAGING SHEET")
    print("="*70)
    success = copy_other_pivots_to_staging(working_dir)
    
    if success:
        print("OTHER PIVOTS DATA COPIED TO STAGING SUCCESSFULLY")
    else:
        print("Failed to copy Other Pivots data. Check logs for details.")
        exit(1)
    
    # Copy SUMMARY sheet data to Staging sheet
    print("\n" + "="*70)
    print("STEP 10/18: COPYING SUMMARY SHEET DATA TO STAGING")
    print("="*70)
    success = copy_summary_sheet_to_staging(working_dir, date_obj)
    
    if success:
        print("SUMMARY SHEET DATA COPIED TO STAGING SUCCESSFULLY")
    else:
        print("Failed to copy SUMMARY sheet data. Check logs for details.")
        exit(1)
    
    # Copy Other Pivots row 18 data to Staging sheet
    print("\n" + "="*70)
    print("STEP 11/18: COPYING OTHER PIVOTS ROW 18 TO STAGING")
    print("="*70)
    success = copy_other_pivots_row18_to_staging(working_dir)
    
    if success:
        print("OTHER PIVOTS ROW 18 DATA COPIED TO STAGING SUCCESSFULLY")
    else:
        print("Failed to copy Other Pivots row 18 data. Check logs for details.")
        exit(1)
    
    # Copy Other Pivots data to New Shcedule sheet
    print("\n" + "="*70)
    print("STEP 12/18: COPYING OTHER PIVOTS DATA TO NEW SCHEDULE")
    print("="*70)
    success = copy_other_pivots_to_new_schedule(working_dir)
    
    if success:
        print("OTHER PIVOTS DATA COPIED TO NEW SHCEDULE SUCCESSFULLY")
    else:
        print("Failed to copy Other Pivots to New Shcedule. Check logs for details.")
        exit(1)
    
    # Copy Staging sheet data to New Shcedule sheet
    print("\n" + "="*70)
    print("STEP 13/18: COPYING STAGING DATA TO NEW SCHEDULE")
    print("="*70)
    success = copy_staging_to_new_schedule(working_dir)
    
    if success:
        print("STAGING DATA COPIED TO NEW SHCEDULE SUCCESSFULLY")
    else:
        print("Failed to copy Staging to New Shcedule. Check logs for details.")
        exit(1)
    
    # Copy and sum SUMMARY sheet values to New Shcedule sheet
    print("\n" + "="*70)
    print("STEP 14/18: COPYING AND SUMMING SUMMARY VALUES TO NEW SCHEDULE")
    print("="*70)
    success = copy_summary_sum_to_new_schedule(working_dir)
    
    if success:
        print("SUMMARY SUM COPIED TO NEW SHCEDULE SUCCESSFULLY")
    else:
        print("Failed to copy SUMMARY sum to New Shcedule. Check logs for details.")
        exit(1)
    
    # Copy SUMMARY sheet range to New Shcedule sheet
    print("\n" + "="*70)
    print("STEP 15/18: COPYING SUMMARY RANGE TO NEW SCHEDULE")
    print("="*70)
    success = copy_summary_range_to_new_schedule(working_dir)
    
    if success:
        print("SUMMARY RANGE COPIED TO NEW SHCEDULE SUCCESSFULLY")
    else:
        print("Failed to copy SUMMARY range to New Shcedule. Check logs for details.")
        exit(1)
    
    # Export SUMMARY sheet data to Excel file
    print("\n" + "="*70)
    print("STEP 16/18: COPYING FILTERED SUMMARY DATA TO RPT LIST SHEET")
    print("="*70)
    success = export_summary_data_to_excel(working_dir)
    
    if success:
        print("SUMMARY DATA COPIED TO RPT LIST SHEET SUCCESSFULLY")
    else:
        print("Failed to copy SUMMARY data to RPT List. Check logs for details.")
        exit(1)
    
    # Fill CLIENT NAME in RPT List from Portfolio file
    print("\n" + "="*70)
    print("STEP 17/18: FILLING CLIENT NAME IN RPT LIST FROM PORTFOLIO")
    print("="*70)
    success = fill_rpt_list_client_names(working_dir, date_obj)
    
    if success:
        print("CLIENT NAME FILLED IN RPT LIST SUCCESSFULLY")
    else:
        print("Failed to fill CLIENT NAME in RPT List. Check logs for details.")
        exit(1)
    
    # Rename Supporting Schedule file with month/year
    print("\n" + "="*70)
    print("STEP 18/18: RENAMING SUPPORTING SCHEDULE FILE")
    print("="*70)
    success = rename_supporting_schedule_file(working_dir, date_obj)
    
    if success:
        print("SUPPORTING SCHEDULE FILE RENAMED SUCCESSFULLY")
    else:
        print("Failed to rename Supporting Schedule file. Check logs for details.")
        exit(1)
    
    print("\n" + "="*70)
    print("ALL 18 OPERATIONS COMPLETED SUCCESSFULLY!")
    print("="*70)
