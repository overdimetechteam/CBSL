#!/usr/bin/env python3
r"""
AFL Liquidity Report Automation Script - Date-Driven Workflow
Central Bank Sri Lanka - NBD_MF_15_LA

This script automates the complete AFL Liquidity report generation workflow:

USAGE:
python NBD_MF_15_LA.py 10/07/2025
python NBD_MF_15_LA.py 07/10/2025 --no-bots

INTEGRATED BOT EXECUTION:
0a. Run TB from ERP bot (Selenium) - downloads TB-COA file for specified date (retries indefinitely until downloaded)
0b. Run M2M from SASIA bot (Selenium) - downloads M2M file to working folder

FILE MANAGEMENT LOGIC:
- FIXED Working Folder: C:\CBSL1\Script\working\NBD_MF_15_LA\10-10-2025
  * ALL input files are placed here (TB-COA, Loan Schedule, Bank Balances, M2M, Deposit Liability, etc.)
  * Bot downloads go here
  * AFL Liquidity working file is here
  * Script makes changes on existing AFL file and creates a copy with new name

EXAMPLE WORKFLOW:
Run 1: Use existing AFL file -> Update it -> Save new copy with updated week number
Run 2: Use latest AFL file -> Update it -> Save new copy with updated week number
All operations happen within: C:\CBSL1\Script\working\NBD_MF_15_LA\10-10-2025

DATE UPDATES IN FILE:
- Creates: AFL Liquidity - Week_XX.xlsx (week number calculated from date)
- Updates all date cells inside the file:
  * Treasury Input C1: Oct-07
  * Liquid Assets D1: Oct-07
  * Borrowings D1: Oct-07
  * NBD-WF-15-LA C3: Oct 07, 2025
  * ALCO C1, D1, E1: 07-Oct-25

DATA PROCESSING (ALL FILES REQUIRED):
1. Run TB bot (retries until TB-COA downloaded to working folder)
2. Find latest AFL file in working folder
3. Copy TB-COA data to working file's TB sheet (REQUIRED)
4. Extract M2M values and update Treasury Input C23, C24, C25 (OPTIONAL)
5. Extract Loan Summary values from Loan Schedule file (REQUIRED)
6. Update working file's Treasury Input sheet with Loan Summary values (REQUIRED)
7. Extract bank balances from Daily Bank Balances file (for specified date) (REQUIRED)
8. Update working file's Treasury Input sheet with bank balances (REQUIRED)
9. Extract Deposit Liability data and update working file (REQUIRED)

VALIDATION & OUTPUT:
10. Validate Liquid Assets sheet (C93 and D93 must be 0)
11. Create new AFL Liquidity - Week_XX.xlsx file with updated dates (only if validation passes)
12. Generate exceptions report if validation fails

COMMAND-LINE OPTIONS:
date              : Report date in MM/DD/YYYY or DD/MM/YYYY format (default: yesterday)
--no-bots         : Skip Selenium bots and use existing files
"""

import os
import sys
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils import get_column_letter
import glob
import logging
from pathlib import Path
import shutil
import subprocess
import time
from dotenv import load_dotenv
import pandas as pd
import xlwings as xw

# Setup logging with UTF-8 encoding to avoid Windows console errors
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('afl_liquidity_automation.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class AFLLiquidityAutomation:
    def __init__(self, base_dir=".", input_date=None, run_bots=True, month_end=False, time_deposits=""):
        """
        Initialize the AFL Liquidity automation script

        Args:
            base_dir (str): Base directory path
            input_date (datetime): Date for report generation (defaults to yesterday)
            run_bots (bool): Whether to run Selenium bots at startup (default: True)
            month_end (bool): Whether this is a month-end report run
            time_deposits (str): Time deposits value for month-end (15.7.1.1.0.0)
        """
        self.base_dir = Path(base_dir)

        # Set date - if no date provided, use yesterday's date
        if input_date is None:
            self.report_date = datetime.now() - timedelta(days=1)
        else:
            self.report_date = input_date

        self.run_bots = run_bots
        self.month_end = month_end
        self.time_deposits = time_deposits

        # Determine week number for the current month
        self.week_of_month = self.get_week_of_month(self.report_date)

        # Calculate ISO week number for the year (for Deposit Liability)
        self.week_of_year = self.report_date.isocalendar()[1]

        # Create date-based folder name for outputs
        self.date_folder_name = self.report_date.strftime("%m-%d-%Y")

        # Extract month and year names for file patterns
        self.month_name = self.report_date.strftime("%B")  # e.g., "August"
        self.year = self.report_date.strftime("%Y")  # e.g., "2025"
        self.mmdd = self.report_date.strftime("%m%d")  # e.g., "0809"

        # Dynamically find working folder - look for single dated subdirectory under NBD_WF_15_LA
        base_working = Path(r"C:\CBSL\Script\working\NBD_WF_15_LA")
        self.working_folder = self._find_single_subdirectory(base_working)

        if self.working_folder is None:
            logger.error(f"Could not find single dated folder under {base_working}")
            raise FileNotFoundError(f"No dated folder found under {base_working}")

        # Bots folder
        self.bots_folder = self.base_dir / "bots"

        # Create folder structure
        self.setup_folder_structure()

        logger.info(f"Initialized AFL Liquidity Automation for date: {self.report_date.strftime('%Y-%m-%d')}")
        logger.info(f"Week of month: {self.week_of_month}")
        logger.info(f"Working folder: {self.working_folder}")
        logger.info(f"Bots folder: {self.bots_folder}")

        # Verify folder exists
        if not self.working_folder.exists():
            logger.error(f"Working folder does not exist: {self.working_folder}")
        else:
            logger.info("✓ Working folder found successfully")
            files = list(self.working_folder.glob("*"))
            logger.info(f"Files in working folder:")
            for f in files:
                logger.info(f"  - {f.name}")

        # File paths will be determined dynamically
        self.afl_liquidity_file = None
        self.working_file = None  # This will be the latest test file or AFL Liquidity file
        self.tb_coa_file = None
        self.daily_bank_balance_file = None
        self.deposit_liability_file = None
        self.m2m_file = None  # M2M file for Treasury Input C23, C24, C25

    def _find_single_subdirectory(self, parent_dir: Path):
        """
        Find the single subdirectory within a parent directory.

        Args:
            parent_dir: Parent directory to search in

        Returns:
            Path to the single subdirectory, or None if not found or multiple exist
        """
        if not parent_dir.exists():
            logger.error(f"Parent directory does not exist: {parent_dir}")
            return None

        subdirs = [d for d in parent_dir.iterdir() if d.is_dir()]

        if len(subdirs) == 0:
            logger.error(f"No subdirectories found in {parent_dir}")
            return None
        elif len(subdirs) > 1:
            logger.warning(f"Multiple subdirectories found in {parent_dir}, using first one: {subdirs[0].name}")
            return subdirs[0]
        else:
            logger.info(f"Found single subdirectory: {subdirs[0].name}")
            return subdirs[0]

    def get_week_of_month(self, date):
        """
        Calculate which week of the month this date falls into (1-5)
        """
        first_day = date.replace(day=1)
        # Calculate the week number within the month
        day_of_month = date.day
        week_number = ((day_of_month - 1) // 7) + 1
        return min(week_number, 5)  # Cap at 5 weeks

    def get_versioned_folder_path(self, base_folder_path):
        """
        Get a versioned folder path if the folder already exists
        Returns: Path object for the folder (either original or versioned)

        Examples:
        - 07-06-2025 -> 07-06-2025 (if doesn't exist)
        - 07-06-2025 -> 07-06-2025(2) (if exists)
        - 07-06-2025 -> 07-06-2025(3) (if (2) also exists)
        """
        try:
            original_path = Path(base_folder_path)

            # If folder doesn't exist, use it as-is
            if not original_path.exists():
                return original_path

            # Folder exists, find next available version number
            version = 2
            while True:
                versioned_path = Path(f"{base_folder_path}({version})")
                if not versioned_path.exists():
                    logger.info(f"Folder {original_path.name} exists, creating versioned folder: {versioned_path.name}")
                    return versioned_path
                version += 1

        except Exception as e:
            logger.error(f"Error getting versioned folder path: {e}")
            return Path(base_folder_path)

    def setup_folder_structure(self):
        r"""
        Folder structure is managed manually - this method does nothing.
        Keeping method for backward compatibility.
        """
        logger.info(f"Folder structure already managed manually, skipping setup")

    def find_tb_coa_file(self):
        """
        Find the TB-COA Excel file downloaded by the bot in working folder
        Looks for pattern like: "*Report*.xlsx" or "TB-COA*.xlsx"
        """
        try:
            # Try multiple patterns to find the TB-COA report
            patterns = [
                "*Report*.xlsx",
                "TB-COA*.xlsx",
                "*TB_COA*.xlsx"
            ]

            files = []
            for pattern in patterns:
                files = list(self.working_folder.glob(pattern))
                if files:
                    logger.info(f"Found TB-COA file(s) with pattern: {pattern}")
                    break

            if not files:
                logger.error(f"No TB-COA file found in {self.working_folder}")
                logger.info("Searched patterns: *Report*.xlsx, TB-COA*.xlsx, *TB_COA*.xlsx")
                return None

            # Use the first file found (or most recent if multiple)
            self.tb_coa_file = files[0]
            logger.info(f"Found TB-COA file: {files[0].name}")
            return files[0]

        except Exception as e:
            logger.error(f"Error finding TB-COA file: {e}")
            return None
    
    def find_daily_bank_balance_file(self):
        """
        Find the Daily Bank Balances Excel file in working folder
        Pattern: Daily Bank Balances - {Month} {Year}.xlsx (e.g., Daily Bank Balances - August 2025.xlsx)
        """
        try:
            # Try specific pattern first: Daily Bank Balances - {Month} {Year}.xlsx
            specific_pattern = f"Daily Bank Balances - {self.month_name} {self.year}.xlsx"
            files = list(self.working_folder.glob(specific_pattern))

            # Fallback to generic pattern if specific not found
            if not files:
                logger.info(f"Specific pattern not found: {specific_pattern}, trying generic pattern...")
                generic_pattern = "Daily Bank Balance*.xlsx"
                files = list(self.working_folder.glob(generic_pattern))

            if not files:
                logger.error(f"No Daily Bank Balances file found in {self.working_folder}")
                logger.info(f"Expected: Daily Bank Balances - {self.month_name} {self.year}.xlsx or Daily Bank Balance*.xlsx")
                return None

            self.daily_bank_balance_file = files[0]
            logger.info(f"Found Daily Bank Balances file: {files[0].name}")
            return files[0]

        except Exception as e:
            logger.error(f"Error finding Daily Bank Balances file: {e}")
            return None
    
    def find_deposit_liability_file(self):
        """
        Find the Deposit Liability Excel file.

        Primary location: C:\\CBSL\\Script\\outputs\\NBD-WF-18-DM\\<dd-mm-yyyy>[(N)]\\
        This is where the NBD-WF-18-DM automation actually publishes its dated output
        (e.g. "07-12-2025(7)"), which may contain a "(N)" suffix when the report for
        that date was run more than once - in that case the most recently modified
        matching folder is used.

        Fallback location: the NBD_WF_15_LA working folder (legacy behaviour), only
        used if nothing is found in the primary location.

        Only a file whose filename date exactly matches the report date is accepted.
        We deliberately do NOT fall back to "first file found" on a date mismatch -
        that previously caused a wrong-period file's balance to be used silently.
        """
        try:
            import re

            # Get expected date string from report_date (format: dd-mm-yyyy)
            expected_date_str = self.report_date.strftime("%d-%m-%Y")
            expected_filename = f"NBD-WF-18-DM Deposit Liability ({expected_date_str}).xlsx"

            def find_exact_match(folder):
                """Return the exact-dated file inside folder, if present."""
                candidate = folder / expected_filename
                if candidate.exists():
                    return candidate
                return None

            # --- Primary: C:\CBSL\Script\outputs\NBD-WF-18-DM\<dd-mm-yyyy>[(N)]\ ---
            dm_outputs_base = self.base_dir / "outputs" / "NBD-WF-18-DM"
            matching_file = None

            if dm_outputs_base.exists():
                folder_pattern = re.compile(
                    r'^' + re.escape(expected_date_str) + r'(\(\d+\))?$'
                )
                date_folders = [
                    d for d in dm_outputs_base.iterdir()
                    if d.is_dir() and folder_pattern.match(d.name)
                ]
                # Most recently modified matching folder first (latest re-run wins)
                date_folders.sort(key=lambda d: d.stat().st_mtime, reverse=True)

                for folder in date_folders:
                    found = find_exact_match(folder)
                    if found:
                        matching_file = found
                        logger.info(f"Found Deposit Liability file in outputs: {folder.name}\\{found.name}")
                        break
                    else:
                        logger.warning(f"Folder {folder.name} does not contain {expected_filename}")

                if not date_folders:
                    logger.warning(f"No dated subfolder matching '{expected_date_str}[(N)]' found under {dm_outputs_base}")
            else:
                logger.warning(f"NBD-WF-18-DM outputs folder not found: {dm_outputs_base}")

            # --- Fallback: legacy working folder, exact date match only ---
            if not matching_file:
                logger.info(f"Falling back to working folder search: {self.working_folder}")
                found = find_exact_match(self.working_folder)
                if found:
                    matching_file = found
                    logger.info(f"Found Deposit Liability file in working folder: {found.name}")

            if not matching_file:
                logger.error(
                    f"No Deposit Liability file found for {expected_date_str}. "
                    f"Checked {dm_outputs_base}\\{expected_date_str}[(N)] and {self.working_folder}. "
                    f"Expected filename: {expected_filename}"
                )
                return None

            self.deposit_liability_file = matching_file
            logger.info(f"Using Deposit Liability file: {matching_file}")
            return matching_file

        except Exception as e:
            logger.error(f"Error finding Deposit Liability file: {e}")
            return None

    def read_treasury_amounts_from_master_data(self):
        """
        Read treasury amounts from Master_Data.xlsx NBD-WF-15-LA sheet

        Returns:
            dict: Dictionary with all treasury amounts, or None on error
        """
        wb = None
        try:
            # Master_Data.xlsx is in the script root directory (same as base_dir)
            master_data_file = self.base_dir / "Master_Data.xlsx"

            if not master_data_file.exists():
                logger.error(f"Master_Data.xlsx not found at {master_data_file}")
                return None

            logger.info(f"Reading treasury amounts from Master_Data.xlsx")
            wb = openpyxl.load_workbook(master_data_file, data_only=True)

            if "NBD-WF-15-LA" not in wb.sheetnames:
                logger.error("NBD-WF-15-LA sheet not found in Master_Data.xlsx")
                return None

            ws = wb["NBD-WF-15-LA"]

            # Read amounts from row 2 (row 1 has headers)
            amounts = {
                'com_papers_cap': ws.cell(row=2, column=1).value,  # Column A
                'com_papers_init_payable': ws.cell(row=2, column=2).value,  # Column B
                'sec_laon_cap': ws.cell(row=2, column=3).value,  # Column C
                'sec_laon_int_payable': ws.cell(row=2, column=4).value,  # Column D
                'rel_party_loans_cap': ws.cell(row=2, column=5).value,  # Column E
                'rel_party_loans_int_pay': ws.cell(row=2, column=6).value,  # Column F
                'repo_cap': ws.cell(row=2, column=7).value,  # Column G
                'repo_int_recei': ws.cell(row=2, column=8).value,  # Column H
                'cost_invest': ws.cell(row=2, column=9).value,  # Column I
                'int_recei': ws.cell(row=2, column=10).value,  # Column J
                'M2M_gain': ws.cell(row=2, column=11).value,  # Column K
            }

            logger.info("Treasury amounts read from Master_Data.xlsx:")
            for key, value in amounts.items():
                logger.info(f"  {key}: {value}")

            return amounts

        except Exception as e:
            logger.error(f"Error reading treasury amounts from Master_Data.xlsx: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None
        finally:
            if wb:
                wb.close()

    def write_time_deposits_to_master_data(self):
        """
        Write time_deposits value to Master_Data.xlsx NBD-WF-15-LA sheet column L (time_deposits).
        Writes to the same row where the last row has data in other columns.
        """
        wb = None
        try:
            master_data_file = self.base_dir / "Master_Data.xlsx"
            if not master_data_file.exists():
                logger.error(f"Master_Data.xlsx not found at {master_data_file}")
                return False

            wb = openpyxl.load_workbook(master_data_file)
            if "NBD-WF-15-LA" not in wb.sheetnames:
                logger.error("NBD-WF-15-LA sheet not found in Master_Data.xlsx")
                return False

            ws = wb["NBD-WF-15-LA"]

            logger.info(f"Master_Data.xlsx sheet dimensions: {ws.dimensions}, max_row: {ws.max_row}")

            # Find the last row that has data in other columns (A-K)
            last_data_row = 1  # header row
            for row in range(2, ws.max_row + 1):
                has_data = False
                row_values = []
                for col in range(1, 12):  # columns A-K
                    cell_val = ws.cell(row=row, column=col).value
                    row_values.append(cell_val)
                    if cell_val is not None:
                        has_data = True
                if has_data:
                    last_data_row = row
                    logger.info(f"Row {row} has data: {row_values}")

            logger.info(f"Last data row determined: {last_data_row}")

            if last_data_row <= 1:
                # No data rows found, write to row 2
                logger.warning("No data rows found in columns A-K, defaulting to row 2")
                last_data_row = 2

            # Parse the time deposits value
            try:
                td_value = float(self.time_deposits)
                logger.info(f"Parsed time deposits value: {td_value}")
            except (ValueError, TypeError):
                logger.error(f"Invalid time deposits value: {self.time_deposits}")
                return False

            # Check current value in time_deposits column before writing
            current_val = ws.cell(row=last_data_row, column=12).value
            logger.info(f"Current value in row {last_data_row}, column L (time_deposits): {current_val}")

            # Write to column L (12) = time_deposits
            ws.cell(row=last_data_row, column=12).value = td_value
            logger.info(f"Writing time_deposits = {td_value} to row {last_data_row}, column L")

            # Save the file
            try:
                wb.save(master_data_file)
                logger.info(f"Successfully saved Master_Data.xlsx with time_deposits = {td_value} at row {last_data_row}, column L")
            except PermissionError:
                logger.error("=" * 80)
                logger.error("CRITICAL ERROR: Master_Data.xlsx is currently OPEN in Excel!")
                logger.error("Please CLOSE Master_Data.xlsx and run the report again.")
                logger.error("=" * 80)
                return False

            return True

        except Exception as e:
            logger.error(f"Error writing time deposits to Master_Data.xlsx: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            if wb:
                wb.close()

    def extract_m2m_values(self):
        """
        Extract values from M2M file (Sheet1) and return them

        Extracts:
        - AB46: For Treasury Input C25
        - L46: For Treasury Input C24
        - I46 + K46: For Treasury Input C23 (sum)

        Returns:
            tuple: (c23_value, c24_value, c25_value) or (None, None, None) on error
        """
        wb = None
        try:
            if not self.m2m_file:
                logger.error("M2M file not found")
                return None, None, None

            # Load the workbook
            logger.info(f"Opening M2M file: {self.m2m_file}")
            wb = openpyxl.load_workbook(self.m2m_file, data_only=True)

            # Check if Sheet1 exists
            if "Sheet1" not in wb.sheetnames:
                logger.error(f"Sheet 'Sheet1' not found in M2M file. Available sheets: {wb.sheetnames}")
                return None, None, None

            sheet = wb["Sheet1"]

            # Extract values from specified cells
            ab46_value = sheet["AB46"].value  # For C25
            l46_value = sheet["L46"].value    # For C24
            i46_value = sheet["I46"].value    # Part 1 for C23
            k46_value = sheet["K46"].value    # Part 2 for C23

            logger.info(f"Extracted M2M values:")
            logger.info(f"  AB46: {ab46_value} (for Treasury Input C25)")
            logger.info(f"  L46: {l46_value} (for Treasury Input C24)")
            logger.info(f"  I46: {i46_value} (for Treasury Input C23 - part 1)")
            logger.info(f"  K46: {k46_value} (for Treasury Input C23 - part 2)")

            # Calculate C23 value (I46 + K46)
            c23_value = None
            if i46_value is not None and k46_value is not None:
                try:
                    c23_value = float(i46_value) + float(k46_value)
                    logger.info(f"  Calculated C23 value: {i46_value} + {k46_value} = {c23_value}")
                except (TypeError, ValueError) as e:
                    logger.error(f"Could not calculate sum of I46 and K46: {e}")
            else:
                logger.warning("I46 or K46 is None, cannot calculate sum for C23")

            return c23_value, l46_value, ab46_value

        except Exception as e:
            logger.error(f"Error extracting M2M values: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None, None, None
        finally:
            if wb is not None:
                try:
                    wb.close()
                except:
                    pass

    def update_treasury_input_with_amounts(self, amounts):
        """
        Update AFL Liquidity Treasury Input sheet with treasury amounts from Master_Data.xlsx

        Args:
            amounts (dict): Dictionary with all treasury amounts
        """
        wb = None
        try:
            if not self.afl_liquidity_file:
                logger.error("AFL Liquidity file not found")
                return False

            if not amounts:
                logger.error("No amounts provided")
                return False

            # Load the workbook
            logger.info(f"Opening AFL Liquidity file: {self.afl_liquidity_file}")
            wb = openpyxl.load_workbook(self.afl_liquidity_file)

            # Check if the required sheet exists
            if "Treasury Input" not in wb.sheetnames:
                logger.error(f"Sheet 'Treasury Input' not found. Available sheets: {wb.sheetnames}")
                return False

            sheet = wb["Treasury Input"]

            # Mapping: amounts dict key -> Treasury Input cell
            cell_mapping = {
                'com_papers_cap': 'C4',
                'com_papers_init_payable': 'C5',
                'sec_laon_cap': 'C8',
                'sec_laon_int_payable': 'C9',
                'rel_party_loans_cap': 'C12',
                'rel_party_loans_int_pay': 'C13',
                'repo_cap': 'C18',
                'repo_int_recei': 'C19',
                'cost_invest': 'C23',
                'int_recei': 'C24',
                'M2M_gain': 'C25',
            }

            # Update cells with amounts
            for key, cell_ref in cell_mapping.items():
                value = amounts.get(key)
                if value is not None:
                    sheet[cell_ref].value = value
                    logger.info(f"✓ Updated Treasury Input {cell_ref} with {key}: {value}")
                else:
                    logger.warning(f"{key} value is None, skipping {cell_ref} update")

            # Save the workbook
            wb.save(self.afl_liquidity_file)
            wb.close()
            wb = None

            logger.info("✓ Treasury Input sheet updated successfully with treasury amounts")
            return True

        except Exception as e:
            logger.error(f"Error updating Treasury Input with treasury amounts: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            if wb is not None:
                try:
                    wb.close()
                except:
                    pass
    
    def copy_tb_coa_data_with_excel_com(self):
        """
        Copy TB-COA data using Excel COM automation (Windows only)
        This bypasses openpyxl's issues with corrupted files
        """
        excel = None
        tb_wb = None
        afl_wb = None

        try:
            import win32com.client

            logger.info("Using Excel COM automation to handle corrupted file...")

            # Create Excel application
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            # Open TB-COA file
            logger.info(f"Opening TB-COA file with Excel: {self.tb_coa_file}")
            tb_wb = excel.Workbooks.Open(str(self.tb_coa_file.absolute()))

            # Find TB-COA sheet
            sheet_name = None
            for sheet in tb_wb.Sheets:
                if 'TB' in sheet.Name.upper() and 'COA' in sheet.Name.upper():
                    sheet_name = sheet.Name
                    break

            if not sheet_name:
                sheet_name = tb_wb.Sheets(1).Name
                logger.info(f"TB-COA sheet not found by name, using first sheet: {sheet_name}")

            tb_sheet = tb_wb.Sheets(sheet_name)
            logger.info(f"Reading sheet: {sheet_name}")

            # Get used range
            used_range = tb_sheet.UsedRange
            data = used_range.Value

            if data is None:
                logger.error("No data found in TB-COA sheet")
                return False

            # Convert to list if single row
            if not isinstance(data, tuple):
                data = [[data]]
            elif not isinstance(data[0], tuple):
                data = [list(data)]
            else:
                data = [list(row) if row else [] for row in data]

            num_rows = len(data)
            num_cols = len(data[0]) if data else 0
            logger.info(f"Read {num_rows} rows x {num_cols} columns from TB-COA")

            # Close TB workbook
            tb_wb.Close(False)
            tb_wb = None
            excel.Quit()
            excel = None

            # Now write to AFL Liquidity using openpyxl
            logger.info(f"Opening AFL Liquidity file: {self.afl_liquidity_file}")
            afl_wb = openpyxl.load_workbook(self.afl_liquidity_file)

            if "TB" not in afl_wb.sheetnames:
                logger.error(f"Sheet 'TB' not found in AFL Liquidity file. Available sheets: {afl_wb.sheetnames}")
                return False

            tb_dest_sheet = afl_wb["TB"]
            logger.info("TB sheet found in AFL Liquidity file")

            # Unmerge all cells first
            logger.info("Unmerging all cells in TB sheet...")
            merged_ranges = list(tb_dest_sheet.merged_cells.ranges)
            for merged_range in merged_ranges:
                tb_dest_sheet.unmerge_cells(str(merged_range))

            # Clear existing data
            logger.info("Clearing existing data in TB sheet...")
            for row in tb_dest_sheet.iter_rows():
                for cell in row:
                    cell.value = None

            # Write data
            logger.info("Writing data to TB sheet...")
            for row_idx, row_data in enumerate(data, 1):
                for col_idx, value in enumerate(row_data, 1):
                    tb_dest_sheet.cell(row=row_idx, column=col_idx, value=value)

            logger.info(f"Wrote {num_rows} rows x {num_cols} columns to TB sheet")

            # Save
            afl_wb.save(self.afl_liquidity_file)
            afl_wb.close()
            afl_wb = None

            logger.info("TB-COA data copied successfully using Excel COM")
            return True

        except ImportError:
            logger.error("pywin32 not installed. Install with: pip install pywin32")
            return False
        except Exception as e:
            logger.error(f"Error copying TB-COA data with Excel COM: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            # Ensure proper cleanup
            try:
                if afl_wb is not None:
                    afl_wb.close()
            except:
                pass

            try:
                if tb_wb is not None:
                    tb_wb.Close(False)
            except:
                pass

            try:
                if excel is not None:
                    excel.Quit()
            except:
                pass
    
    def copy_tb_coa_data(self):
        """
        Copy all data from TB-COA sheet to AFL Liquidity TB sheet
        """
        try:
            if not self.tb_coa_file:
                logger.error("TB-COA file not found")
                return False
            
            if not self.afl_liquidity_file:
                logger.error("AFL Liquidity file not found")
                return False
            
            # Try Excel COM automation (works on Windows)
            logger.info("Attempting to use Excel COM automation...")
            return self.copy_tb_coa_data_with_excel_com()
            
        except Exception as e:
            logger.error(f"Error copying TB-COA data: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def extract_loan_summary_values(self):
        """
        Extract values from Loan Schedule file (in Input folder)
        Sheet: Loan Summary, Cells: I29 and J29
        """
        wb = None
        try:
            if not self.loan_schedule_file:
                logger.error("Loan Schedule file not found")
                return None, None

            # Load the workbook
            logger.info(f"Opening Loan Schedule file: {self.loan_schedule_file}")
            wb = openpyxl.load_workbook(self.loan_schedule_file, data_only=True)

            # Check if the required sheet exists
            if "Loan Summary" not in wb.sheetnames:
                logger.error(f"Sheet 'Loan Summary' not found. Available sheets: {wb.sheetnames}")
                return None, None

            sheet = wb["Loan Summary"]

            # Get values from I29 and J29
            i29_value = sheet["I29"].value
            j29_value = sheet["J29"].value

            logger.info(f"Extracted values - I29: {i29_value}, J29: {j29_value}")

            return i29_value, j29_value

        except Exception as e:
            logger.error(f"Error extracting Loan Summary values: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None, None
        finally:
            if wb is not None:
                try:
                    wb.close()
                except:
                    pass
    
    def update_deposit_liability_value(self):
        """
        Extract value from Deposit Liability file (NBD-WF-18-DM sheet, D14 cell)
        and update AFL Liquidity file (NBD-WF-15-LA sheet, C7 cell)

        Uses VBScript with Excel COM automation as primary method to handle
        corrupted Excel files that trigger "We found a problem with some content" warning.
        Falls back to ZIP/XML extraction and openpyxl if VBScript fails.
        """

        
        afl_wb = None
        try:
            if not self.deposit_liability_file:
                logger.error("Deposit Liability file not found")
                return False

            if not self.afl_liquidity_file:
                logger.error("AFL Liquidity file not found")
                return False

            logger.info(f"Opening Deposit Liability file: {self.deposit_liability_file}")

            d14_value = None

            # Method 0: Try VBScript with Excel COM (handles corrupted files best)
            try:
                vbs_script = Path(__file__).parent / "copy_deposit_liability_value.vbs"
                if vbs_script.exists():
                    logger.info("Trying VBScript method with Excel COM automation...")

                    # Close any open Excel instances that have the deposit liability or AFL files open
                    try:
                        import win32com.client
                        excel_app = win32com.client.GetObject(Class="Excel.Application")
                        if excel_app:
                            files_to_close = [
                                str(self.deposit_liability_file).lower(),
                                str(self.afl_liquidity_file).lower()
                            ]
                            for wb in list(excel_app.Workbooks):
                                if wb.FullName.lower() in files_to_close:
                                    logger.info(f"Closing open Excel file before VBS: {wb.FullName}")
                                    wb.Close(SaveChanges=True)
                            # Release COM reference
                            del excel_app
                    except Exception as close_err:
                        # No Excel instance running or file not open - that's fine
                        logger.debug(f"No open Excel instance found: {close_err}")

                    result = subprocess.run(
                        ["cscript", "//NoLogo", str(vbs_script),
                         str(self.deposit_liability_file), str(self.afl_liquidity_file)],
                        capture_output=True,
                        text=True,
                        timeout=120
                    )

                    # Log output
                    for line in result.stdout.strip().split('\n'):
                        if line.strip():
                            logger.info(f"VBS: {line.strip()}")

                    if result.returncode == 0 and "SUCCESS" in result.stdout:
                        logger.info("VBScript method succeeded - D14 value copied to C7")
                        return True
                    else:
                        if result.stderr:
                            logger.warning(f"VBScript stderr: {result.stderr}")
                        logger.warning("VBScript method failed, trying fallback methods...")
                else:
                    logger.warning(f"VBScript not found at {vbs_script}, using fallback methods...")
            except subprocess.TimeoutExpired:
                logger.warning("VBScript timed out, trying fallback methods...")
            except Exception as vbs_error:
                logger.warning(f"VBScript method failed: {vbs_error}")

            # # Method 1: Try direct ZIP/XML extraction (bypasses openpyxl corruption issues)
            # try:
            #     import zipfile
            #     import xml.etree.ElementTree as ET

            #     logger.info("Reading D14 value using ZIP/XML extraction...")

            #     with zipfile.ZipFile(self.deposit_liability_file, 'r') as z:
            #         # First, find which sheet is "NBD-WF-18-DM"
            #         workbook_xml = z.read('xl/workbook.xml')
            #         wb_root = ET.fromstring(workbook_xml)

            #         # Find sheet name to sheet ID mapping
            #         sheet_id = None
            #         ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            #               'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

            #         for sheet in wb_root.findall('.//main:sheet', ns):
            #             if sheet.get('name') == 'NBD-WF-18-DM':
            #                 sheet_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            #                 break

            #         if not sheet_id:
            #             logger.error("Sheet 'NBD-WF-18-DM' not found in workbook")
            #             raise Exception("Sheet 'NBD-WF-18-DM' not found")

            #         # Get the sheet file path from relationships
            #         rels_xml = z.read('xl/_rels/workbook.xml.rels')
            #         rels_root = ET.fromstring(rels_xml)

            #         sheet_file = None
            #         for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            #             if rel.get('Id') == sheet_id:
            #                 target = rel.get('Target')
            #                 # Handle both relative and absolute paths
            #                 if target.startswith('xl/') or target.startswith('/xl/'):
            #                     sheet_file = target.lstrip('/')
            #                 else:
            #                     sheet_file = 'xl/' + target
            #                 break

            #         if not sheet_file:
            #             logger.error(f"Could not find sheet file for {sheet_id}")
            #             raise Exception("Sheet file not found")

            #         logger.info(f"Reading sheet from: {sheet_file}")

            #         # Read the sheet XML
            #         sheet_xml = z.read(sheet_file)

            #         # Read shared strings if available
            #         shared_strings = []
            #         try:
            #             ss_xml = z.read('xl/sharedStrings.xml')
            #             ss_root = ET.fromstring(ss_xml)
            #             for si in ss_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
            #                 shared_strings.append(si.text if si.text else '')
            #         except KeyError:
            #             pass

            #     # Parse sheet data
            #     sheet_root = ET.fromstring(sheet_xml)

            #     # Find cell D14
            #     for row_elem in sheet_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
            #         row_num = row_elem.get('r')
            #         if row_num == '14':
            #             for cell in row_elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            #                 cell_ref = cell.get('r', '')
            #                 if cell_ref == 'D14':
            #                     cell_type = cell.get('t', '')
            #                     value_elem = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')

            #                     if value_elem is not None and value_elem.text:
            #                         if cell_type == 's':  # Shared string
            #                             idx = int(value_elem.text)
            #                             if idx < len(shared_strings):
            #                                 d14_value = shared_strings[idx]
            #                         else:
            #                             # Numeric value
            #                             try:
            #                                 d14_value = float(value_elem.text)
            #                             except ValueError:
            #                                 d14_value = value_elem.text

            #                     logger.info(f"D14 value from XML: {d14_value}")
            #                     break
            #             break

            # except Exception as xml_error:
            #     logger.warning(f"ZIP/XML extraction failed: {xml_error}")

            # # Method 2: Fallback to openpyxl if XML extraction failed
            # if d14_value is None:
            #     logger.info("Trying openpyxl with data_only=True...")
            #     try:
            #         dep_wb = openpyxl.load_workbook(self.deposit_liability_file, data_only=True)
            #         if "NBD-WF-18-DM" in dep_wb.sheetnames:
            #             dep_sheet = dep_wb["NBD-WF-18-DM"]
            #             d14_value = dep_sheet["D14"].value
            #             logger.info(f"D14 value from openpyxl (data_only): {d14_value}")
            #         dep_wb.close()
            #     except Exception as e:
            #         logger.warning(f"openpyxl data_only failed: {e}")

            # # Method 3: If still None, try formula calculation
            # if d14_value is None:
            #     logger.info("Trying manual formula calculation...")
            #     try:
            #         dep_wb = openpyxl.load_workbook(self.deposit_liability_file, data_only=False)
            #         dep_wb_values = openpyxl.load_workbook(self.deposit_liability_file, data_only=True)

            #         if "NBD-WF-18-DM" in dep_wb.sheetnames:
            #             dep_sheet = dep_wb["NBD-WF-18-DM"]
            #             dep_sheet_values = dep_wb_values["NBD-WF-18-DM"]

            #             d14_raw = dep_sheet["D14"].value
            #             logger.info(f"D14 formula: {d14_raw}")

            #             if isinstance(d14_raw, str) and d14_raw.startswith("="):
            #                 formula = d14_raw.upper()
            #                 import re

            #                 # Handle SUM formula
            #                 if "SUM" in formula:
            #                     # First try range format: SUM(D3:D10)
            #                     sum_match = re.search(r'SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', formula)
            #                     if sum_match:
            #                         col = sum_match.group(1)
            #                         row_start = int(sum_match.group(2))
            #                         row_end = int(sum_match.group(4))

            #                         total = 0
            #                         for row in range(row_start, row_end + 1):
            #                             cell_val = dep_sheet_values[f"{col}{row}"].value
            #                             if cell_val is None:
            #                                 cell_val = dep_sheet[f"{col}{row}"].value
            #                             if isinstance(cell_val, (int, float)):
            #                                 total += cell_val
            #                         d14_value = total
            #                         logger.info(f"Calculated SUM (range): {d14_value}")
            #                     else:
            #                         # Handle SUM with individual cells: SUM(D5+D8+D9+D12+D13)
            #                         # Extract all cell references from inside SUM()
            #                         sum_content_match = re.search(r'SUM\(([^)]+)\)', formula)
            #                         if sum_content_match:
            #                             sum_content = sum_content_match.group(1)
            #                             cell_refs = re.findall(r'([A-Z]+)(\d+)', sum_content)
            #                             total = 0
            #                             for col, row in cell_refs:
            #                                 cell_val = dep_sheet_values[f"{col}{row}"].value
            #                                 if cell_val is None:
            #                                     cell_val = dep_sheet[f"{col}{row}"].value
            #                                 if isinstance(cell_val, (int, float)):
            #                                     total += cell_val
            #                             d14_value = total
            #                             logger.info(f"Calculated SUM (individual cells): {d14_value}")

            #                 # Handle addition formula (without SUM)
            #                 elif "+" in formula:
            #                     cell_refs = re.findall(r'([A-Z]+)(\d+)', formula)
            #                     total = 0
            #                     for col, row in cell_refs:
            #                         cell_val = dep_sheet_values[f"{col}{row}"].value
            #                         if cell_val is None:
            #                             cell_val = dep_sheet[f"{col}{row}"].value
            #                         if isinstance(cell_val, (int, float)):
            #                             total += cell_val
            #                     d14_value = total
            #                     logger.info(f"Calculated addition: {d14_value}")

            #         dep_wb.close()
            #         dep_wb_values.close()
            #     except Exception as e:
            #         logger.warning(f"Formula calculation failed: {e}")

            if d14_value is None:
                logger.error("Could not extract D14 value from Deposit Liability file")
                return False

            logger.info(f"Final D14 value to copy: {d14_value}")

            # Load AFL Liquidity file and update C7
            logger.info(f"Opening AFL Liquidity file: {self.afl_liquidity_file}")
            afl_wb = openpyxl.load_workbook(self.afl_liquidity_file)

            if "NBD-WF-15-LA" not in afl_wb.sheetnames:
                logger.error(f"Sheet 'NBD-WF-15-LA' not found. Available: {afl_wb.sheetnames}")
                return False

            afl_sheet = afl_wb["NBD-WF-15-LA"]
            afl_sheet["C7"].value = d14_value
            logger.info(f"Updated AFL Liquidity NBD-WF-15-LA C7 with value: {d14_value}")

            afl_wb.save(self.afl_liquidity_file)
            afl_wb.close()
            afl_wb = None

            logger.info("Deposit Liability value updated successfully")
            return True

        except Exception as e:
            logger.error(f"Error updating deposit liability value: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            if afl_wb is not None:
                try:
                    afl_wb.close()
                except:
                    pass
    
    def update_treasury_input_with_bank_balances(self, target_date="07/25/2025"):
        """
        Update AFL Liquidity Treasury Input sheet with bank balances from Daily Bank Balances file

        This function:
        1. Validates and opens AFL Liquidity file
        2. Clears Treasury Input range A29:C150
        3. Copies BankBalanceGLCodes data to columns A & B starting at row 29 (using openpyxl)
        4. Reads BankBalanceGLCodes.xlsx into DataFrame as COA
        5. Reads Daily Bank Balances (Sheet1) into DataFrame:
        - Skips Excel row 1
        - Uses Excel row 2 as column headers/headings (dates become datetime objects)
        - Excel row 3+ becomes the data rows
        6. Finds target date column in Daily Bank Balances:
        - Checks row 1 from left to right using openpyxl
        - Finds column where row 1 is NOT blank AND next column's row 1 IS blank
        - This identifies the last filled column (running date column)
        7. Creates "COA Amount" column in COA by matching GL codes
        - Matches COA column 0 (GL codes) with Daily Bank Balance column 1 (GL codes)
        - Fills COA Amount with values from the target date column found in step 6
        8. Copies COA Amount to Treasury Input column C starting at row 29
        9. Writes formulas to Treasury Input sheet:
        - C74: "=SUM(C30:C73)"
        - C94: "=SUM(C77:C93)"
        10. Adds SUM formulas for subtotal rows in column C
        11. Inserts non-matching values into Liquid Assets sheet

        Args:
            target_date (str): Date to look up in MM/DD/YYYY format (currently not used - uses last filled column)
        """
        bank_wb = None
        afl_wb = None
        try:
            if not self.daily_bank_balance_file:
                logger.error("Daily Bank Balances file not found")
                return False

            if not self.afl_liquidity_file:
                logger.error("AFL Liquidity file not found")
                return False

            # Parse the target date to be flexible with different formats
            try:
                target_dt = datetime.strptime(target_date, '%m/%d/%Y')
            except:
                try:
                    target_dt = datetime.strptime(target_date, '%d/%m/%Y')
                except:
                    logger.error(f"Invalid date format: {target_date}. Use MM/DD/YYYY")
                    return False

            logger.info(f"Looking for date: {target_dt.strftime('%m/%d/%Y')}")

            # Step 1: Validate file exists before opening
            logger.info(f"Validating AFL Liquidity file: {self.afl_liquidity_file}")

            # Check if file exists
            if not self.afl_liquidity_file.exists():
                logger.error(f"File does not exist: {self.afl_liquidity_file}")
                return False

            # Check file properties
            logger.info(f"File exists: True")
            logger.info(f"File size: {self.afl_liquidity_file.stat().st_size} bytes")
            logger.info(f"Full path: {self.afl_liquidity_file.absolute()}")

            # Validate file is a proper Excel file using openpyxl
            logger.info("Validating file format with openpyxl...")
            try:
                test_wb = openpyxl.load_workbook(self.afl_liquidity_file, read_only=True, data_only=False)
                sheet_names = test_wb.sheetnames
                logger.info(f"File is a valid Excel workbook with {len(sheet_names)} sheets")

                # Check for Treasury Input sheet
                if "Treasury Input" not in sheet_names:
                    logger.error("'Treasury Input' sheet not found in workbook")
                    test_wb.close()
                    return False

                test_wb.close()
                logger.info("File validation passed")
            except Exception as validation_error:
                logger.error(f"File validation failed: {validation_error}")
                logger.error("The file may be corrupted or in an unsupported format")
                return False

            # Step 2: Open AFL Liquidity file with openpyxl
            logger.info("Opening AFL Liquidity file with openpyxl...")
            afl_wb = openpyxl.load_workbook(self.afl_liquidity_file, data_only=False)
            treasury_sheet = afl_wb["Treasury Input"]
            logger.info("File opened successfully")

            # Step 3: Clear Treasury Input column C (amounts) range C29:C150
            logger.info("Clearing Treasury Input column C range C29:C150")
            for row_idx in range(29, 151):
                treasury_sheet.cell(row=row_idx, column=3).value = None

            # Step 4: Read Daily Bank Balances file
            logger.info(f"Reading Daily Bank Balances file: {self.daily_bank_balance_file}")

            # Open with openpyxl to find the last filled date column (row 3)
            temp_bank_wb = openpyxl.load_workbook(self.daily_bank_balance_file, data_only=True)
            temp_bank_sheet = temp_bank_wb.active if temp_bank_wb.active else temp_bank_wb['Sheet1']

            target_col_index = None
            # Search row 3 from left to right — find last column where current cell has data and next is blank
            for col_idx in range(1, temp_bank_sheet.max_column + 1):
                current_cell_value = temp_bank_sheet.cell(row=3, column=col_idx).value
                next_cell_value = temp_bank_sheet.cell(row=3, column=col_idx + 1).value if col_idx < temp_bank_sheet.max_column else None

                current_not_blank = current_cell_value is not None and str(current_cell_value).strip() != ""
                next_is_blank = next_cell_value is None or str(next_cell_value).strip() == ""

                if current_not_blank and next_is_blank:
                    target_col_index = col_idx
                    logger.info(f"Found target date column: {get_column_letter(col_idx)} (index {col_idx})")
                    break

            temp_bank_wb.close()

            if not target_col_index:
                logger.error("Could not find target date column in Daily Bank Balances")
                return False

            # Read Daily Bank Balances into DataFrame (skip row 1, use row 2 as headers)
            daily_bank_df = pd.read_excel(
                self.daily_bank_balance_file,
                sheet_name='Sheet1',
                skiprows=1,
                header=0
            )
            logger.info(f"Daily Bank Balance DataFrame: {len(daily_bank_df)} rows, columns: {list(daily_bank_df.columns[:4])}...")

            if len(daily_bank_df.columns) < 2:
                logger.error(f"Daily Bank Balance file has too few columns: {len(daily_bank_df.columns)}")
                return False

            if target_col_index > len(daily_bank_df.columns):
                logger.error(f"Target column index {target_col_index} beyond DataFrame columns ({len(daily_bank_df.columns)})")
                return False

            target_date_column = daily_bank_df.columns[target_col_index - 1]
            logger.info(f"Using amount column: '{target_date_column}'")

            # Step 5: Read existing GL codes from Treasury Input column A (A30 onwards)
            # These are already in the template — match them with Daily Bank Balances
            logger.info("Reading GL codes from Treasury Input column A and matching with Daily Bank Balances...")

            # Build GL -> amount lookup from Daily Bank Balances
            daily_gl_column = daily_bank_df.columns[1]  # GL column (index 1)
            gl_to_amount = dict(zip(
                pd.to_numeric(daily_bank_df[daily_gl_column], errors='coerce').fillna(0).astype(int),
                daily_bank_df[target_date_column].fillna(0)
            ))

            # Fill column C by matching GL codes in column A
            matches_found = 0
            for row_idx in range(30, 151):
                gl_value = treasury_sheet.cell(row=row_idx, column=1).value
                if gl_value is None:
                    continue
                try:
                    gl_int = int(gl_value)
                except (ValueError, TypeError):
                    continue
                amount = gl_to_amount.get(gl_int, 0)
                treasury_sheet.cell(row=row_idx, column=3).value = amount
                if amount != 0:
                    matches_found += 1
                    logger.info(f"  Row {row_idx}: GL {gl_value} = {amount}")

            logger.info(f"Matched {matches_found} GL codes with amounts")

            # Step 6: Write SUM formulas to Treasury Input sheet
            logger.info("Writing SUM formula to cell C74: =SUM(C30:C73)")
            treasury_sheet.cell(row=74, column=3).value = "=SUM(C30:C73)"

            logger.info("Writing SUM formula to cell C94: =SUM(C77:C93)")
            treasury_sheet.cell(row=94, column=3).value = "=SUM(C77:C93)"

            # Step 13: Extract A29:C150 to dataframe
            logger.info("Extracting A29:C150 to dataframe")
            treasury_data = []
            for row_idx in range(29, 151):
                row_data = [
                    treasury_sheet.cell(row=row_idx, column=1).value,
                    treasury_sheet.cell(row=row_idx, column=2).value,
                    treasury_sheet.cell(row=row_idx, column=3).value
                ]
                treasury_data.append(row_data)
            treasury_df = pd.DataFrame(treasury_data, columns=['A', 'B', 'C'])

            # Step 14: Get Liquid Assets sheet range A11:B until "FD & Repo"
            if "Liquid Assets" not in afl_wb.sheetnames:
                logger.error(f"Sheet 'Liquid Assets' not found.")
                return False

            liquid_assets_sheet = afl_wb["Liquid Assets"]
            logger.info("Reading Liquid Assets sheet A11:B until 'FD & Repo'")

            # Find the row with "FD & Repo" in column B
            liquid_assets_data = []
            end_row = None
            max_search_row = 200  # Set a reasonable max to avoid infinite loops

            for row_idx in range(11, max_search_row + 1):
                b_value = liquid_assets_sheet.cell(row=row_idx, column=2).value
                if b_value and str(b_value).strip() == "FD & Repo":
                    end_row = row_idx
                    logger.info(f"Found 'FD & Repo' at row {row_idx}")
                    break
                a_value = liquid_assets_sheet.cell(row=row_idx, column=1).value
                liquid_assets_data.append([a_value, b_value])

            liquid_assets_df = pd.DataFrame(liquid_assets_data, columns=['A', 'B'])

            # Step 15: Find non-matching values
            logger.info("Finding non-matching values")
            liquid_assets_values = set(liquid_assets_df['B'].dropna().astype(str))
            treasury_values = set(treasury_df['B'].dropna().astype(str))
            non_matching = treasury_values - liquid_assets_values

            # Step 16: Insert non-matching values at appropriate rows
            logger.info(f"Found {len(non_matching)} non-matching values to insert")

            # Find position markers in treasury data
            current_accounts_idx = None
            savings_accounts_idx = None

            for idx, row in treasury_df.iterrows():
                b_val = str(row['B']) if row['B'] is not None else ""
                if "Current Accounts Balances" in b_val:
                    current_accounts_idx = idx
                elif "Savings Accounts Balances" in b_val:
                    savings_accounts_idx = idx

            for non_match_value in non_matching:
                # Find which section this value belongs to in treasury_df
                value_idx = treasury_df[treasury_df['B'] == non_match_value].index

                if len(value_idx) > 0:
                    value_position = value_idx[0]
                    insert_row = None

                    # Determine insert position
                    if current_accounts_idx is not None and value_position > current_accounts_idx:
                        if savings_accounts_idx is not None and value_position > savings_accounts_idx:
                            insert_row = 75
                            logger.info(f"Inserting '{non_match_value}' at row 75 (after Savings Accounts)")
                        else:
                            insert_row = 55
                            logger.info(f"Inserting '{non_match_value}' at row 55 (under Current Accounts)")

                    if insert_row:
                        # Insert row in Liquid Assets sheet using openpyxl
                        liquid_assets_sheet.insert_rows(insert_row)

                        # Get the value from treasury_df
                        matching_row = treasury_df[treasury_df['B'] == non_match_value].iloc[0]
                        liquid_assets_sheet.cell(row=insert_row, column=1).value = matching_row['A']
                        liquid_assets_sheet.cell(row=insert_row, column=2).value = matching_row['B']

                        # Copy formulas from above row to columns C and D
                        if insert_row > 11:
                            source_row = insert_row - 1
                            for col_idx in [3, 4]:  # Columns C and D
                                source_cell = liquid_assets_sheet.cell(row=source_row, column=col_idx)
                                dest_cell = liquid_assets_sheet.cell(row=insert_row, column=col_idx)

                                # Check if source has a formula
                                if source_cell.value and isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                                    # Copy formula (openpyxl will adjust references automatically)
                                    dest_cell.value = source_cell.value
                                else:
                                    dest_cell.value = source_cell.value

                                # Copy formatting
                                if source_cell.has_style:
                                    dest_cell.font = source_cell.font.copy()
                                    dest_cell.border = source_cell.border.copy()
                                    dest_cell.fill = source_cell.fill.copy()
                                    dest_cell.number_format = source_cell.number_format
                                    dest_cell.alignment = source_cell.alignment.copy()

            # Save and close the workbook
            logger.info("Saving AFL Liquidity file...")
            afl_wb.save(self.afl_liquidity_file)
            afl_wb.close()
            afl_wb = None

            logger.info("Treasury Input sheet updated successfully with bank balances")
            return True

        except Exception as e:
            logger.error(f"Error updating Treasury Input with bank balances: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            # Close openpyxl workbooks
            if afl_wb is not None:
                try:
                    afl_wb.close()
                except:
                    pass
            if bank_wb is not None:
                try:
                    bank_wb.close()
                except:
                    pass                
    
    def update_treasury_input_sheet(self, i29_value, j29_value):
        """
        Update the AFL Liquidity file (in Inputs folder) with Loan Summary values
        Sheet: Treasury Input, Cells: E8, E9

        NOTE: C8 and C9 are now updated from Master_Data.xlsx treasury amounts (sec_laon_cap, sec_laon_int_payable)
        so we only update E8 and E9 here to avoid overwriting those values.
        """
        wb = None
        try:
            if not self.afl_liquidity_file:
                logger.error("AFL Liquidity file not found")
                return False

            # Load the workbook
            logger.info(f"Opening AFL Liquidity file: {self.afl_liquidity_file}")
            wb = openpyxl.load_workbook(self.afl_liquidity_file)

            # Check if the required sheet exists
            if "Treasury Input" not in wb.sheetnames:
                logger.error(f"Sheet 'Treasury Input' not found. Available sheets: {wb.sheetnames}")
                return False

            sheet = wb["Treasury Input"]

            # Update cells with I29 and J29 values (only E8 and E9, not C8 and C9)
            if i29_value is not None:
                sheet["E8"].value = i29_value
                logger.info(f"Updated E8 with I29 value: {i29_value}")
            else:
                logger.warning("I29 value is None, skipping E8 update")

            if j29_value is not None:
                sheet["E9"].value = j29_value
                logger.info(f"Updated E9 with J29 value: {j29_value}")
            else:
                logger.warning("J29 value is None, skipping E9 update")

            # Save the workbook
            wb.save(self.afl_liquidity_file)
            wb.close()
            wb = None

            logger.info("Treasury Input sheet updated successfully with Loan Summary values")
            return True

        except Exception as e:
            logger.error(f"Error updating Treasury Input sheet: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            if wb is not None:
                try:
                    wb.close()
                except:
                    pass
    
    def get_next_test_version(self):
        """
        Find the highest test version number and return the next one
        Returns: next version number (int)
        """
        try:
            # Find all AFL Test files in working folder
            test_files = list(self.working_folder.glob("AFL Test *.xlsx"))

            if not test_files:
                # No test files found, start with version 1
                return 1

            # Extract version numbers
            import re
            versions = []
            for file in test_files:
                match = re.search(r'AFL Test (\d+)\.xlsx', file.name)
                if match:
                    versions.append(int(match.group(1)))

            if versions:
                return max(versions) + 1
            else:
                return 1

        except Exception as e:
            logger.error(f"Error finding next test version: {e}")
            return 1

    def find_latest_afl_file_across_all_folders(self):
        """
        Find the latest AFL Liquidity file in working folder
        Finds the most recently modified AFL file (excluding working_ files)

        NEW MONTH LOGIC:
        When starting a new month, this will automatically find the 5th week AFL file
        from the previous month (which is kept after monthly report generation).
        This file becomes the base for creating the first week's AFL file of the new month.

        Returns: Path to latest AFL file, or None if not found
        """
        try:
            logger.info("Searching for latest AFL Liquidity file in working folder...")

            # Find AFL Liquidity files in working folder (excluding working_ prefix files)
            # This will find the 5th week file from previous month when starting new month
            all_afl_files = [f for f in self.working_folder.glob("AFL Liquidity - Week_*.xlsx")
                           if not f.name.startswith("working_")]

            if not all_afl_files:
                # Month-end completion moves that month's week files into an
                # "AFL <Mon>-<Year> Expand" folder instead of deleting them (see cbsl.py's
                # LA post-processing step), so a new month can start with no top-level
                # Week_*.xlsx file. Fall back to the last recorded week file in there,
                # copied up (never edited in place, so the archived history stays intact).
                import re
                expand_dirs = list(self.working_folder.glob("AFL * Expand"))
                if expand_dirs:
                    expand_files = [f for f in expand_dirs[0].glob("AFL Liquidity - Week_*.xlsx")
                                     if not f.name.startswith("working_")]
                    if expand_files:
                        source_file = max(
                            expand_files,
                            key=lambda f: int(re.search(r'Week_(\d+)', f.name).group(1))
                        )
                        restored_file = self.working_folder / source_file.name
                        shutil.copy2(source_file, restored_file)
                        logger.info(f"No top-level Week_*.xlsx found; restored carry-forward "
                                    f"seed from '{expand_dirs[0].name}': {restored_file.name}")
                        return restored_file

                logger.info("No AFL Liquidity files found in working folder")
                return None

            # Get the most recent file by modification time
            latest_file = max(all_afl_files, key=lambda f: f.stat().st_mtime)
            logger.info(f"✓ Latest AFL file: {latest_file.name}")
            logger.info(f"  Modified: {datetime.fromtimestamp(latest_file.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')}")
            return latest_file

        except Exception as e:
            logger.error(f"Error finding latest AFL file: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None

    def update_date_cells(self):
        """
        Update date cells in all sheets of the working file.
        - Treasury Input C1: "Aug - 10"
        - Liquid Assets D1: "Aug - 10"
        - Borrowings D1: "Aug - 10"
        - NBD-WF-15-LA C3: "Aug 10, 2025"
        - ALCO C1, D1, E1: "10-Aug-25"
        """
        try:
            wb = openpyxl.load_workbook(self.afl_liquidity_file)

            # Format: "Aug - 10" (month abbreviation - space dash space - day without leading zero)
            short_date = f"{self.report_date.strftime('%b')} - {self.report_date.day}"

            # Treasury Input C1
            if "Treasury Input" in wb.sheetnames:
                wb["Treasury Input"]["C1"].value = short_date
                logger.info(f"Updated Treasury Input C1: {short_date}")

            # Liquid Assets D1
            if "Liquid Assets" in wb.sheetnames:
                wb["Liquid Assets"]["D1"].value = short_date
                logger.info(f"Updated Liquid Assets D1: {short_date}")

            # Borrowings D1
            if "Borrowings" in wb.sheetnames:
                wb["Borrowings"]["D1"].value = short_date
                logger.info(f"Updated Borrowings D1: {short_date}")

            # NBD-WF-15-LA C3: "Aug 10, 2025"
            if "NBD-WF-15-LA" in wb.sheetnames:
                nbd_date = f"{self.report_date.strftime('%b')} {self.report_date.day}, {self.report_date.year}"
                wb["NBD-WF-15-LA"]["C3"].value = nbd_date
                logger.info(f"Updated NBD-WF-15-LA C3: {nbd_date}")

            # ALCO C1, D1, E1: "10-Aug-25"
            if "ALCO" in wb.sheetnames:
                alco_sheet = wb["ALCO"]
                alco_date = self.report_date.strftime("%d-%b-%y")

                # Unmerge cells first if they are merged
                for cell_ref in ["C1", "D1", "E1"]:
                    cell = alco_sheet[cell_ref]
                    if isinstance(cell, openpyxl.cell.cell.MergedCell):
                        for merged_range in list(alco_sheet.merged_cells.ranges):
                            if cell_ref in merged_range:
                                alco_sheet.unmerge_cells(str(merged_range))
                                break

                alco_sheet["C1"].value = alco_date
                alco_sheet["D1"].value = alco_date
                alco_sheet["E1"].value = alco_date
                logger.info(f"Updated ALCO C1, D1, E1: {alco_date}")

            wb.save(self.afl_liquidity_file)
            wb.close()
            logger.info("All date cells updated successfully")
            return True

        except Exception as e:
            logger.error(f"Error updating date cells: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def create_afl_test_copy(self):
        """
        Create a new versioned AFL file from the currently updated working file.
        The working file (which was just updated) becomes the source for the new AFL file.
        Saves to the versioned output Inputs folder (e.g., 07-13-2025/NBD_MF_04_LA/Inputs or 07-13-2025(2)/NBD_MF_04_LA/Inputs)
        and updates date cells inside the file.
        """
        wb = None
        try:
            # Calculate week number from report date
            week_number = self.report_date.isocalendar()[1]

            # Create filename with week number
            copy_filename = f"AFL Liquidity - Week_{week_number}.xlsx"
            copy_path = self.working_folder / copy_filename

            # Use the current working file (which has been updated) as source
            source_file = self.working_file

            if not source_file:
                logger.error("Working file not set")
                return False

            logger.info(f"Creating {copy_filename} from {source_file.name}...")

            # Create the copy in working folder
            shutil.copy2(source_file, copy_path)
            logger.info(f"✓ Created: {copy_path.name} in working folder")

            # Now update date cells inside the new file
            logger.info("Updating date cells inside the new file...")
            wb = openpyxl.load_workbook(copy_path)

            # Update Treasury Input sheet - C1 (e.g., "Oct-07")
            if "Treasury Input" in wb.sheetnames:
                treasury_sheet = wb["Treasury Input"]
                treasury_sheet["C1"].value = self.report_date.strftime("%b-%d")
                logger.info(f"Updated Treasury Input C1: {self.report_date.strftime('%b-%d')}")

            # Update Liquid Assets sheet - D1 (e.g., "Oct-07")
            if "Liquid Assets" in wb.sheetnames:
                liquid_sheet = wb["Liquid Assets"]
                liquid_sheet["D1"].value = self.report_date.strftime("%b-%d")
                logger.info(f"Updated Liquid Assets D1: {self.report_date.strftime('%b-%d')}")

            # Update Borrowings sheet - D1 (e.g., "Oct-07")
            if "Borrowings" in wb.sheetnames:
                borrowings_sheet = wb["Borrowings"]
                borrowings_sheet["D1"].value = self.report_date.strftime("%b-%d")
                logger.info(f"Updated Borrowings D1: {self.report_date.strftime('%b-%d')}")

            # Update NBD-WF-15-LA sheet - C3 (e.g., "Oct 07, 2025")
            if "NBD-WF-15-LA" in wb.sheetnames:
                nbd_sheet = wb["NBD-WF-15-LA"]
                nbd_sheet["C3"].value = self.report_date.strftime("%b %d, %Y")
                logger.info(f"Updated NBD-WF-15-LA C3: {self.report_date.strftime('%b %d, %Y')}")

            # Update ALCO sheet - C1, D1, E1 (e.g., "07-Oct-25")
            if "ALCO" in wb.sheetnames:
                alco_sheet = wb["ALCO"]
                alco_date = self.report_date.strftime("%d-%b-%y")

                # Unmerge cells first if they are merged
                cells_to_update = ["C1", "D1", "E1"]
                for cell_ref in cells_to_update:
                    # Check if cell is part of a merged range
                    cell = alco_sheet[cell_ref]
                    if isinstance(cell, openpyxl.cell.cell.MergedCell):
                        # Find and unmerge the range
                        for merged_range in list(alco_sheet.merged_cells.ranges):
                            if cell_ref in merged_range:
                                alco_sheet.unmerge_cells(str(merged_range))
                                logger.info(f"Unmerged cell range: {merged_range}")
                                break

                # Now update the values
                alco_sheet["C1"].value = alco_date
                alco_sheet["D1"].value = alco_date
                alco_sheet["E1"].value = alco_date
                logger.info(f"Updated ALCO C1, D1, E1: {alco_date}")

            # Save the updated workbook
            wb.save(copy_path)
            wb.close()
            wb = None

            # Update working_file to point to the new test file
            self.working_file = copy_path
            logger.info(f"Working file updated to: {copy_path.name}")
            logger.info("✓ All date cells updated successfully")

            return True

        except Exception as e:
            logger.error(f"Error creating AFL Test copy: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            if wb is not None:
                try:
                    wb.close()
                except:
                    pass

    def run_tb_bot(self):
        """
        Run the TB from ERP bot to download TB-COA file
        Retries indefinitely until the file is successfully downloaded
        Returns: True if successful, False otherwise
        """
        try:
            logger.info("\n" + "=" * 60)
            logger.info("Running TB from ERP Bot (LA version from report_automations)...")
            logger.info("=" * 60)

            # Use TB bot from report_automations folder for LA
            report_automations_folder = self.base_dir / "report_automations"
            bot_script = report_automations_folder / "tb_from_erp_bot.py"

            if not bot_script.exists():
                logger.error(f"TB bot script not found: {bot_script}")
                return False

            logger.info(f"Using TB bot from: {bot_script}")

            # Delete old TB-COA files before downloading new one
            logger.info("Checking for old TB-COA files to delete...")
            old_tb_files = []
            patterns = ["*Report*.xlsx", "TB-COA*.xlsx", "*TB_COA*.xlsx"]
            for pattern in patterns:
                old_tb_files.extend(self.working_folder.glob(pattern))

            if old_tb_files:
                logger.info(f"Found {len(old_tb_files)} old TB-COA file(s) to delete:")
                for old_file in old_tb_files:
                    try:
                        logger.info(f"  Deleting: {old_file.name}")
                        old_file.unlink()
                        logger.info(f"  ✓ Deleted: {old_file.name}")
                    except Exception as e:
                        logger.warning(f"  Could not delete {old_file.name}: {e}")
            else:
                logger.info("No old TB-COA files found")

            # Load environment variables
            env_path = self.base_dir / '.env'
            load_dotenv(dotenv_path=env_path)

            # Import and run the bot from report_automations folder
            sys.path.insert(0, str(report_automations_folder))

            try:
                from tb_from_erp_bot import ERPLoginBot

                # Get encrypted credentials
                ENC_USERNAME = os.getenv('ENC_USERNAME')
                ENC_PASSWORD = os.getenv('ENC_PASSWORD')

                if not ENC_USERNAME or not ENC_PASSWORD:
                    logger.error("ENC_USERNAME and ENC_PASSWORD not found in environment variables")
                    return False

                # Format report date for bot (DD/MM/YYYY)
                report_date = self.report_date.strftime('%d/%m/%Y')

                # Retry loop - keep trying until TB-COA file is downloaded
                retry_count = 0
                while True:
                    retry_count += 1
                    logger.info(f"\nAttempt #{retry_count} to download TB-COA file...")

                    # Create bot instance with working folder as download location
                    bot = ERPLoginBot(download_folder=str(self.working_folder))

                    # Run bot automation
                    success = bot.run_full_automation(ENC_USERNAME, ENC_PASSWORD, report_date)

                    # Clean up
                    bot.close_browser()

                    if success:
                        logger.info("✓ TB bot completed successfully")

                        # Verify TB-COA file actually exists
                        tb_file = self.find_tb_coa_file()
                        if tb_file:
                            logger.info(f"✓ TB-COA file confirmed downloaded: {tb_file.name}")
                            return True
                        else:
                            logger.warning("TB bot reported success but file not found. Retrying immediately...")
                            continue
                    else:
                        logger.warning(f"TB bot failed on attempt #{retry_count}. Retrying immediately...")
                        continue

            except ImportError as e:
                logger.error(f"Failed to import TB bot: {e}")
                return False

        except Exception as e:
            logger.error(f"Error running TB bot: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            # Remove report_automations folder from path
            report_automations_folder = self.base_dir / "report_automations"
            if str(report_automations_folder) in sys.path:
                sys.path.remove(str(report_automations_folder))

    def run_m2m_bot(self):
        """
        Run the M2M from SASIA bot
        Returns: True if successful, False otherwise
        """
        try:
            logger.info("\n" + "=" * 60)
            logger.info("Running M2M from SASIA Bot...")
            logger.info("=" * 60)

            bot_script = self.bots_folder / "m2m_from_sasia_bot.py"

            if not bot_script.exists():
                logger.error(f"M2M bot script not found: {bot_script}")
                return False

            # Check if the script has actual content
            with open(bot_script, 'r') as f:
                content = f.read().strip()
                if len(content) < 50:  # If it's just a placeholder
                    logger.warning("M2M bot script appears to be a placeholder, skipping...")
                    return True

            # Run the bot script
            result = subprocess.run(
                [sys.executable, str(bot_script)],
                cwd=str(self.bots_folder),
                capture_output=True,
                text=True
            )

            if result.returncode == 0:
                logger.info("✓ M2M bot completed successfully")
                logger.info(result.stdout)
                time.sleep(2)
                return True
            else:
                logger.error(f"M2M bot failed with return code {result.returncode}")
                logger.error(result.stderr)
                return False

        except Exception as e:
            logger.error(f"Error running M2M bot: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def determine_working_file(self):
        """
        Determine which file to work on:
        - Search working folder for latest AFL Liquidity file (by modification time)
        - Edit the original file directly (no copy)
        """
        try:
            # Search working folder for the latest AFL file
            latest_afl = self.find_latest_afl_file_across_all_folders()

            if latest_afl:
                self.working_file = latest_afl
                logger.info(f"Working file: {latest_afl.name}")
            else:
                logger.error("No AFL Liquidity files found in working folder")
                logger.info(f"Please place an AFL Liquidity file in {self.working_folder}")
                return None

            return self.working_file

        except Exception as e:
            logger.error(f"Error determining working file: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None

    def validate_liquid_assets(self):
        """
        Validate that Liquid Assets sheet cells C93 and D93 are 0
        Returns: (is_valid, exceptions_list)
        """
        wb = None
        try:
            if not self.afl_liquidity_file:
                logger.error("AFL Liquidity file not found for validation")
                return False, []

            logger.info("Validating Liquid Assets sheet...")
            wb = openpyxl.load_workbook(self.afl_liquidity_file, data_only=True)

            if "Liquid Assets" not in wb.sheetnames:
                logger.error(f"Sheet 'Liquid Assets' not found. Available sheets: {wb.sheetnames}")
                return False, []

            sheet = wb["Liquid Assets"]

            # Get values from C93 and D93
            c93_value = sheet["C93"].value
            d93_value = sheet["D93"].value

            logger.info(f"Liquid Assets validation - C93: {c93_value}, D93: {d93_value}")

            exceptions = []

            # Check if C93 is not 0
            if c93_value is not None and c93_value != 0:
                exceptions.append({
                    'file': self.afl_liquidity_file.name,
                    'sheet': 'Liquid Assets',
                    'cell': 'C93',
                    'value': c93_value,
                    'expected': 0
                })
                logger.warning(f"VALIDATION FAILED: C93 = {c93_value} (expected 0)")

            # Check if D93 is not 0
            if d93_value is not None and d93_value != 0:
                exceptions.append({
                    'file': self.afl_liquidity_file.name,
                    'sheet': 'Liquid Assets',
                    'cell': 'D93',
                    'value': d93_value,
                    'expected': 0
                })
                logger.warning(f"VALIDATION FAILED: D93 = {d93_value} (expected 0)")

            if len(exceptions) == 0:
                logger.info("✓ Validation PASSED: Both C93 and D93 are 0")
                return True, []
            else:
                logger.error(f"✗ Validation FAILED: {len(exceptions)} exception(s) found")
                return False, exceptions

        except Exception as e:
            logger.error(f"Error during validation: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False, []
        finally:
            if wb is not None:
                try:
                    wb.close()
                except:
                    pass

    def create_exceptions_report(self, exceptions):
        """
        Create an exceptions report Excel file with the validation failures

        Args:
            exceptions (list): List of exception dictionaries
        """
        try:
            if not exceptions:
                logger.info("No exceptions to report")
                return True

            # Extract week number from AFL Liquidity filename
            afl_filename = self.afl_liquidity_file.name
            week_number = "XX"  # Default if not found

            # Try to extract week number from filename
            import re
            week_match = re.search(r'Week\s*(\d+)', afl_filename, re.IGNORECASE)
            if week_match:
                week_number = week_match.group(1)

            # Create exceptions report filename
            exceptions_filename = f"AFL Liquidity - week {week_number} - Exceptions.xlsx"
            exceptions_path = self.working_folder / exceptions_filename

            logger.info(f"Creating exceptions report: {exceptions_filename}")

            # Create a new workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Exceptions"

            # Add headers
            headers = ['File', 'Sheet', 'Cell', 'Actual Value', 'Expected Value']
            ws.append(headers)

            # Style headers
            from openpyxl.styles import Font, PatternFill, Alignment
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_num)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Add exception rows
            for exception in exceptions:
                ws.append([
                    exception['file'],
                    exception['sheet'],
                    exception['cell'],
                    exception['value'],
                    exception['expected']
                ])

            # Adjust column widths
            ws.column_dimensions['A'].width = 40
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 10
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 15

            # Save the workbook
            wb.save(exceptions_path)
            wb.close()

            logger.info(f"✓ Exceptions report created: {exceptions_path}")
            return True

        except Exception as e:
            logger.error(f"Error creating exceptions report: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def cleanup_working_files(self):
        """
        Delete all working_AFL_Liquidity files after processing is complete
        """
        try:
            logger.info("Cleaning up temporary working files...")

            # Find all working_ files
            working_files = list(self.working_folder.glob("working_AFL Liquidity*.xlsx"))

            if not working_files:
                logger.info("No working files to clean up")
                return True

            logger.info(f"Found {len(working_files)} working file(s) to delete:")
            for working_file in working_files:
                try:
                    logger.info(f"  Deleting: {working_file.name}")
                    working_file.unlink()
                    logger.info(f"  ✓ Deleted: {working_file.name}")
                except Exception as e:
                    logger.warning(f"  Could not delete {working_file.name}: {e}")

            logger.info("✓ Working files cleanup completed")
            return True

        except Exception as e:
            logger.error(f"Error cleaning up working files: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def run_automation(self):
        """
        Main method to run the AFL Liquidity report automation with integrated bot execution
        """
        logger.info("=" * 60)
        logger.info("Starting AFL Liquidity Report Automation")
        logger.info(f"Report Date: {self.report_date.strftime('%Y-%m-%d')}")
        if self.month_end:
            logger.info(f"Mode: MONTH-END (Time Deposits: {self.time_deposits})")
        else:
            logger.info("Mode: Weekly")
        logger.info("=" * 60)

        try:
            # Step 0: Run Selenium bots if enabled
            if self.run_bots:
                # Run TB from ERP bot first - CRITICAL: Will retry until successful
                logger.info("\nStep 0a: Running TB from ERP Bot...")
                logger.info("NOTE: TB download will retry indefinitely until successful")
                tb_bot_success = self.run_tb_bot()

                if not tb_bot_success:
                    logger.error("TB bot failed to download TB-COA file after retries")
                    logger.error("Cannot proceed without TB-COA file. Stopping automation.")
                    return False

                # M2M bot is no longer used - amounts are now entered manually via web interface
                # and stored in Master_Data.xlsx NBD-WF-15-LA sheet
                logger.info("\nStep 0b: Skipping M2M bot (amounts read from Master_Data.xlsx)")
            else:
                logger.info("\nBot execution disabled - skipping Selenium automation")

            # Step 1: Find required files
            logger.info("\nStep 1: Finding required Excel files...")
            logger.info("  - Looking for TB-COA file in Input folder...")
            tb_coa_file = self.find_tb_coa_file()

            logger.info("  - Looking for Daily Bank Balances in Input folder...")
            bank_balance_file = self.find_daily_bank_balance_file()

            logger.info("  - Looking for Deposit Liability in Input folder...")
            deposit_liability_file = self.find_deposit_liability_file()

            # Check all required files - ALL must be present
            missing_files = []
            if not tb_coa_file:
                missing_files.append("TB-COA file")
            if not bank_balance_file:
                missing_files.append("Daily Bank Balances file")
            if not deposit_liability_file:
                missing_files.append("Deposit Liability file")

            if missing_files:
                logger.error("\n" + "=" * 60)
                logger.error("MISSING REQUIRED FILES - CANNOT PROCEED")
                logger.error("=" * 60)
                logger.error(f"The following required files are missing from Input folder:")
                for file in missing_files:
                    logger.error(f"  ✗ {file}")
                logger.error(f"\nWorking folder location: {self.working_folder}")
                logger.error("\nPlease ensure all required files are present before running the automation.")
                logger.error("=" * 60)
                return False

            # Step 1.5: Determine which file to work on (latest AFL file in working folder)
            logger.info("\nStep 1.5: Determining working file...")
            working_file = self.determine_working_file()

            if not working_file:
                logger.error("Could not determine working file. Stopping automation.")
                return False

            # Important: Update self.afl_liquidity_file to point to working file for all operations
            # This ensures all updates happen on the correct file
            self.afl_liquidity_file = working_file

            # Step 1.6: Write time deposits to Master_Data.xlsx (month-end only)
            if self.month_end and self.time_deposits:
                logger.info("\nStep 1.6: Writing time deposits to Master_Data.xlsx (month-end mode)...")
                td_success = self.write_time_deposits_to_master_data()
                if not td_success:
                    logger.warning("Failed to write time deposits to Master_Data.xlsx, but continuing...")
            elif self.month_end:
                logger.info("\nMonth-end mode enabled but no time deposits value provided, skipping...")

            # Step 2: Copy TB-COA data to working file's TB sheet
            logger.info("\nStep 2: Copying TB-COA data to working file's TB sheet...")
            tb_copy_success = self.copy_tb_coa_data()

            if not tb_copy_success:
                logger.error("Failed to copy TB-COA data")
                return False

            # Step 2.5: Read treasury amounts from Master_Data.xlsx
            logger.info("\nStep 2.5: Reading treasury amounts from Master_Data.xlsx...")
            logger.info("  - Reading amounts from NBD-WF-15-LA sheet...")
            amounts = self.read_treasury_amounts_from_master_data()

            if amounts:
                # Update Treasury Input with treasury amounts
                logger.info("  - Updating Treasury Input with treasury amounts...")
                amounts_update_success = self.update_treasury_input_with_amounts(amounts)

                if not amounts_update_success:
                    logger.warning("Failed to update Treasury Input with treasury amounts, but continuing...")
            else:
                logger.warning("No treasury amounts found in Master_Data.xlsx, skipping treasury amounts update (C4, C5, C8, C9, C12, C13, C18, C19, C23, C24, C25 will not be updated)")

            # Step 3: Update working file's Treasury Input sheet with bank balances
            logger.info("\nStep 3: Updating working file's Treasury Input sheet with bank balances...")
            # Use the report date for bank balances lookup
            target_date_str = self.report_date.strftime('%m/%d/%Y')
            bank_balance_success = self.update_treasury_input_with_bank_balances(target_date=target_date_str)

            if not bank_balance_success:
                logger.error("Failed to update Treasury Input with bank balances")
                return False

            # Step 4: Update Deposit Liability value in working file
            if self.month_end and self.time_deposits:
                # Month-end: use user-provided time deposits value for C7 instead of DM value
                logger.info("\nStep 4: Updating NBD-WF-15-LA C7 with month-end time deposits value...")
                try:
                    td_value = float(self.time_deposits)
                    afl_wb = openpyxl.load_workbook(self.afl_liquidity_file)
                    afl_sheet = afl_wb["NBD-WF-15-LA"]
                    afl_sheet["C7"].value = td_value
                    afl_wb.save(self.afl_liquidity_file)
                    afl_wb.close()
                    logger.info(f"Updated NBD-WF-15-LA C7 with month-end time deposits: {td_value}")
                except Exception as e:
                    logger.error(f"Failed to update C7 with time deposits: {e}")
                    return False
            else:
                logger.info("\nStep 4: Updating Deposit Liability value in working file...")
                deposit_success = self.update_deposit_liability_value()

                if not deposit_success:
                    logger.error("Failed to update Deposit Liability value")
                    return False

            # Step 5: Update date cells in all sheets
            logger.info("\nStep 5: Updating date cells in all sheets...")
            self.update_date_cells()

            # Step 6: Validate working file
            logger.info("\nStep 6: Validating Liquid Assets sheet (C93 and D93 must be 0)...")
            is_valid, exceptions = self.validate_liquid_assets()

            if not is_valid:
                logger.warning("Validation failed - creating exceptions report...")
                self.create_exceptions_report(exceptions)
                logger.info("\n" + "=" * 60)
                logger.warning("COMPLETED WITH EXCEPTIONS: Check the exceptions report")
                logger.warning("NOT creating test file due to validation errors")
                logger.info("=" * 60)
                return True  # Automation completed, but with validation errors

            logger.info("\n" + "=" * 60)
            logger.info("SUCCESS: AFL Liquidity Report Automation completed successfully")
            logger.info("All validations passed!")
            logger.info(f"Output file: {self.working_file.name}")
            logger.info("=" * 60)

            return True
            
        except Exception as e:
            logger.error(f"Error in automation process: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

def main():
    """
    Main function to run the AFL Liquidity automation with integrated workflow

    Usage:
        python NBD_MF_15_LA.py 10/07/2025
        python NBD_MF_15_LA.py 07/10/2025 --no-bots

    Command-line arguments:
        date: Report date in MM/DD/YYYY or DD/MM/YYYY format (optional, defaults to yesterday)
        --no-bots: Skip running Selenium bots (use existing files)
    """
    try:
        # Parse command line arguments
        import argparse
        parser = argparse.ArgumentParser(
            description='AFL Liquidity Report Automation - Date-Driven Workflow',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Examples:
    python NBD_MF_15_LA.py 10/07/2025
    python NBD_MF_15_LA.py 07/10/2025 --no-bots

Workflow:
1. Run TB from ERP bot (Selenium) - downloads TB-COA file for the specified date
2. Run M2M from SASIA bot (Selenium) - downloads M2M file
3. Determine working file (latest AFL Test X or AFL Liquidity)
4. Copy TB-COA data to working file
5. Extract and update Loan Summary values
6. Update bank balances for the specified date
7. Update Deposit Liability values
8. Validate Liquid Assets sheet (C93 and D93 must be 0)
9. Create new AFL Liquidity - Week_XX.xlsx file with updated dates (if validation passes)

Date Updates in File:
- Treasury Input C1: Oct-07
- Liquid Assets D1: Oct-07
- Borrowings D1: Oct-07
- NBD-WF-15-LA C3: Oct 07, 2025
- ALCO C1, D1, E1: 07-Oct-25
            """
        )
        parser.add_argument('date', nargs='?', type=str, help='Report date in MM/DD/YYYY or DD/MM/YYYY format (default: yesterday)')
        parser.add_argument('--no-bots', action='store_true', help='Skip running Selenium bots (use existing files)')
        parser.add_argument('--month-end', action='store_true', help='Mark this run as a month-end report')
        parser.add_argument('--time-deposits', type=str, default='', help='15.7.1.1.0.0 Time Deposits value for month-end')
        args = parser.parse_args()

        report_date = None
        if args.date:
            # Try both MM/DD/YYYY and DD/MM/YYYY formats
            try:
                report_date = datetime.strptime(args.date, '%m/%d/%Y')
            except ValueError:
                try:
                    report_date = datetime.strptime(args.date, '%d/%m/%Y')
                except ValueError:
                    logger.error("Invalid date format. Use MM/DD/YYYY or DD/MM/YYYY (e.g., 10/07/2025 or 07/10/2025)")
                    sys.exit(1)

        run_bots = not args.no_bots

        logger.info("=" * 60)
        logger.info("AFL LIQUIDITY AUTOMATION - DATE-DRIVEN WORKFLOW")
        logger.info("=" * 60)
        if report_date:
            logger.info(f"Report Date: {report_date.strftime('%m/%d/%Y (%B %d, %Y)')}")
        else:
            logger.info("Report Date: Yesterday (default)")
        if run_bots:
            logger.info("Mode: Full automation (with Selenium bots)")
        else:
            logger.info("Mode: Skip bots (use existing files)")
        if args.month_end:
            logger.info(f"Report Type: MONTH-END (Time Deposits: {args.time_deposits})")
        else:
            logger.info("Report Type: Weekly")
        logger.info("=" * 60)

        # Initialize and run automation
        automation = AFLLiquidityAutomation(
            base_dir="C:\\CBSL\\Script",
            input_date=report_date,
            run_bots=run_bots,
            month_end=args.month_end,
            time_deposits=args.time_deposits
        )
        success = automation.run_automation()

        if success:
            logger.info("\n✓ Automation completed successfully")
            sys.exit(0)
        else:
            logger.error("\n✗ Automation failed")
            sys.exit(1)

    except KeyboardInterrupt:
        logger.info("\nAutomation interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"\nUnexpected error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()