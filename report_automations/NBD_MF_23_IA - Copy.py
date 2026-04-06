import os
import glob
import pandas as pd
from datetime import datetime, date
import numpy as np
import subprocess
import json
import logging
import sys
import traceback
from typing import Dict, Optional, Tuple, Any
from pathlib import Path
import xlwings as xw
from pyxlsb import open_workbook
from openpyxl import load_workbook


# Create logs folder if it doesn't exist
logs_dir = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(logs_dir, exist_ok=True)

# Create timestamp-based log file name
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = os.path.join(logs_dir, f"NBD_MF23_IA_Report_{timestamp}.log")

# Enhanced logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s",
    handlers=[
        logging.FileHandler(log_file, mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)
logger.info("="*80)
logger.info("STARTING NEW SESSION - NBD_MF23_IA_Report")
logger.info("Logging initialized. Logs are being written to %s", log_file)
logger.info("="*80)

# Add the parent "bots" folder relative to this file
current_dir = os.path.dirname(__file__)
bot_dir = os.path.join(current_dir, "..", "bots")
sys.path.insert(0, os.path.abspath(bot_dir))

# Import bot modules with error handling
try:
    import na_contract_numbers_search_bot_api as na_contract_bot
    logger.info("Successfully imported na_contract_numbers_search_bot_api module")
except ImportError as e:
    logger.warning(f"Could not import na_contract_numbers_search_bot_api: {e}")
    na_contract_bot = None

try:
    import IA_Working_Initial_valuation_bot as initial_valuation_bot
    logger.info("Successfully imported IA_Working_Initial_valuation_bot module")
except ImportError as e:
    logger.warning(f"Could not import IA_Working_Initial_valuation_bot: {e}")
    initial_valuation_bot = None


class NBD_MF23_IA_Report:
    def __init__(self, base_dir=r"working", working_dir=None, rate_updates=None, ltv_updates=None):
        """Initialize the report generator with error handling.

        Args:
            base_dir: Base directory for default folder searching (used in standalone mode)
            working_dir: Direct path to working directory (used when called from app.py)
            rate_updates: List of dict with contract_number, contractual_rate, base_rate for manual updates
            ltv_updates: List of dict with contract_number, ltv_percentage for manual LTV updates
        """
        try:
            self.base_dir = base_dir
            self.working_dir = working_dir
            self.rate_updates = rate_updates or []
            self.ltv_updates = ltv_updates or []
            self.ia_folder = self._find_ia_folder()
            logger.info(f"Initialized NBD_MF23_IA_Report with base_dir: {base_dir}")
            if working_dir:
                logger.info(f"Using working_dir from app.py: {working_dir}")
            if self.rate_updates:
                logger.info(f"Loaded {len(self.rate_updates)} contract rate updates for manual override")
            if self.ltv_updates:
                logger.info(f"Loaded {len(self.ltv_updates)} contract LTV updates for manual override")
            logger.info(f"IA folder found at: {self.ia_folder}")
        except Exception as e:
            logger.error(f"Failed to initialize NBD_MF23_IA_Report: {e}")
            raise

    @staticmethod
    def parse_date(date_input):
        """Parse date from various formats - for use by app.py or command line."""
        if isinstance(date_input, date):
            return date_input
        elif isinstance(date_input, str):
            try:
                # Try MM/DD/YYYY format first (command line)
                return datetime.strptime(date_input, "%m/%d/%Y").date()
            except ValueError:
                try:
                    # Try YYYY-MM-DD format (app.py)
                    return datetime.strptime(date_input, "%Y-%m-%d").date()
                except ValueError:
                    raise ValueError(f"Invalid date format: {date_input}. Expected MM/DD/YYYY or YYYY-MM-DD")
        else:
            raise ValueError(f"Invalid date type: {type(date_input)}. Expected string or date object")

    def _find_ia_folder(self):
        """Find the latest NBD_MF_23_IA date folder inside the working directory."""
        try:
            # Step 1: Locate project root (one level above report_automations)
            root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

            # Step 2: Path to working/NBD_MF_23_IA
            ia_root = os.path.join(root_dir, "working", "NBD_MF_23_IA")
            if not os.path.exists(ia_root):
                error_msg = f"'NBD_MF_23_IA' folder not found at '{ia_root}'."
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)

            # Step 3: Get all subfolders (date-based)
            date_folders = [f for f in glob.glob(os.path.join(ia_root, "*")) if os.path.isdir(f)]
            if not date_folders:
                error_msg = f"No subfolders found under '{ia_root}'."
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)

            # Step 4: Pick latest modified folder
            date_folders.sort(key=os.path.getmtime, reverse=True)
            latest_folder = date_folders[0]

            logger.info(f"Found latest IA working folder: {latest_folder}")
            return latest_folder

        except Exception as e:
            logger.error(f"Error finding IA folder: {e}")
            raise


    def _find_file(self, keyword: str) -> str:
        """Find a file containing the keyword in the IA folder with detailed logging."""
        try:
            logger.debug(f"Searching for file with keyword: '{keyword}' in {self.ia_folder}")

            files = os.listdir(self.ia_folder)
            # Filter out temporary Excel files that start with ~$
            matching_files = [f for f in files if keyword in f and not f.startswith('~$')]

            if not matching_files:
                error_msg = f"No file found with keyword: '{keyword}' in {self.ia_folder}"
                logger.error(error_msg)
                logger.error(f"Available files (excluding ~$ temp files): {[f for f in files if not f.startswith('~$')]}")
                raise FileNotFoundError(error_msg)

            if len(matching_files) > 1:
                logger.warning(f"Multiple files found with keyword '{keyword}': {matching_files}. Using first match.")

            selected_file = os.path.join(self.ia_folder, matching_files[0])
            logger.debug(f"Found file: {selected_file}")
            return selected_file

        except Exception as e:
            logger.error(f"Error finding file with keyword '{keyword}': {e}")
            raise

    # === File Loaders with Enhanced Error Handling ===
    def load_disbursement(self):
        """Load disbursement data with error handling."""
        try:
            logger.info("Loading Disbursement with Budget file...")
            file_disbursement = self._find_file("Disbursement with Budget")
            
            df = pd.read_excel(
                file_disbursement,
                sheet_name="month",
                usecols="A,H,AC"
            )
            
            # Validate required columns
            required_cols = ["CONTRACT NO", "Net Amount", "Base Rate"]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Missing required columns in Disbursement file: {missing_cols}")
            
            df_filtered = df[df["CONTRACT NO"].notna()]
            logger.info(f"Successfully loaded Disbursement data: {len(df_filtered)} rows")
            return df_filtered
            
        except Exception as e:
            logger.error(f"Failed to load Disbursement data: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_portfolio_rec(self):
        """Load PORTFOLIOREC CSV file (columns C and AA) with error handling."""
        try:
            logger.info("Loading PORTFOLIOREC file...")

            # Locate the file dynamically
            file_portfolio = self._find_file("PORTFOLIOREC")

            # Read CSV - C and AA correspond to columns 2 and 26 (0-indexed: 2, 26)
            df = pd.read_csv(file_portfolio, usecols=[2, 26])

            # Assign proper column names: column C = CONTRACT_NO, column AA = SELLING_PRICE
            df.columns = ["CONTRACT_NO", "SELLING_PRICE"]

            # Filter out rows with null CONTRACT_NO
            df_filtered = df[df["CONTRACT_NO"].notna()]

            logger.info(f"Successfully loaded PORTFOLIOREC data: {len(df_filtered)} rows from {os.path.basename(file_portfolio)}")
            return df_filtered

        except Exception as e:
            logger.error(f"Failed to load PORTFOLIOREC data: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_information_request_from_credit(self):
        """Load information request from credit with error handling."""
        try:
            logger.info("Loading Information Request from Credit file...")
            file_path = self._find_file("Information Request from Credit")
            
            df = pd.read_excel(
                file_path,
                sheet_name="Disbursements",
                usecols="C,G",
                skiprows=4
            )
            
            # Rename columns if needed
            if len(df.columns) == 2:
                df.columns = ["Contract No", "Micro/Small/Medium"]
            
            df_filtered = df[df["Contract No"].notna()]
            logger.info(f"Successfully loaded Information Request data: {len(df_filtered)} rows")
            return df_filtered
            
        except Exception as e:
            logger.error(f"Failed to load Information Request from Credit: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_summary(self):
        """Load summary with error handling."""
        try:
            logger.info("Loading Information Request from Credit file...")
            file_path = self._find_file("Summary")
            
            df = pd.read_excel(
                file_path,
                sheet_name="SUMMARY",
                usecols="A,R",
                skiprows=2
            )

            return df
            
        except Exception as e:
            logger.error(f"Failed to load Summary sheet: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_net_portfolio(self):
        """Load net portfolio with error handling."""
        try:
            logger.info("Loading Net Portfolio file...")
            file_path = self._find_file("Net Portfolio")
            
            df = pd.read_excel(
                file_path,
                usecols="E,C,F,H,AB,AH,AI,AM"
            )
            
            # Validate required columns
            required_cols = ["CLIENT_CODE", "CONTRACT_NO"]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                logger.warning(f"Missing expected columns in Net Portfolio: {missing_cols}")
            
            df_filtered = df[df["CLIENT_CODE"].notna()]
            logger.info(f"Successfully loaded Net Portfolio data: {len(df_filtered)} rows")
            return df_filtered
            
        except Exception as e:
            logger.error(f"Failed to load Net Portfolio: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_yard_and_Property_List(self):
        """Load yard and property list with comprehensive error handling."""
        try:
            logger.info("Loading CBSL Provision Comparison file...")
            file_path = self._find_file("CBSL Provision Comparison")
            
            # Load Property Mortgage
            try:
                df_propertyMortgage = pd.read_excel(
                    file_path,
                    sheet_name="PropertyMortgage",
                    usecols="A",
                    skiprows=5,
                    header=None
                )
                df_propertyMortgage.columns = ["Property Mortgage List"]
                logger.info(f"Loaded PropertyMortgage: {len(df_propertyMortgage)} rows")
            except Exception as e:
                logger.warning(f"Could not load PropertyMortgage sheet: {e}. Creating empty dataframe.")
                df_propertyMortgage = pd.DataFrame(columns=["Property Mortgage List"])

            # Load Repossessed List
            try:
                df_repossessedList = pd.read_excel(
                    file_path,
                    sheet_name="RepossessedList",
                    usecols="A",
                    skiprows=5,
                    header=None
                )
                df_repossessedList.columns = ["Yard Contract"]
                logger.info(f"Loaded RepossessedList: {len(df_repossessedList)} rows")
            except Exception as e:
                logger.warning(f"Could not load RepossessedList sheet: {e}. Creating empty dataframe.")
                df_repossessedList = pd.DataFrame(columns=["Yard Contract"])

            # Build combined dataframe
            df_yard_and_property_list = pd.DataFrame()
            df_yard_and_property_list["Yard Contract"] = df_repossessedList["Yard Contract"]
            df_yard_and_property_list[""] = [""] + ["Yard"] * (max(0, len(df_repossessedList) - 1))
            df_yard_and_property_list["Empty"] = ""
            
            max_len = max(len(df_repossessedList), len(df_propertyMortgage))
            df_yard_and_property_list = df_yard_and_property_list.reindex(range(max_len))
            df_yard_and_property_list["Property Mortgage List"] = df_propertyMortgage.reindex(range(max_len))["Property Mortgage List"].fillna("")
            
            logger.info(f"Successfully created yard and property list: {len(df_yard_and_property_list)} rows")
            return df_yard_and_property_list
            
        except Exception as e:
            logger.error(f"Failed to load Yard and Property List: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_product_cat(self):
        """Load product category from .xlsb and synchronize Product_Cat.xlsx data with error handling."""
        try:
            logger.info("Loading Prod. wise Class. of Loans file...")

            # Step 1: Locate main .xlsb file
            file_path = Path(self._find_file("Prod. wise Class. of Loans"))
            base_dir = file_path.parent

            if not file_path.exists():
                raise FileNotFoundError(f"Prod. wise Class. of Loans file not found at {file_path}")

            # Step 2: Locate Product_Cat.xlsx
            prod_cat_file = base_dir / "Product_Cat.xlsx"
            if not prod_cat_file.exists():
                raise FileNotFoundError(f"Product_Cat.xlsx not found in {base_dir}")

            logger.info(f"Found Product_Cat.xlsx at {prod_cat_file}")

            # Step 3: Read Product_Cat.xlsx
            df_source = pd.read_excel(prod_cat_file, sheet_name="Sheet1")

            required_cols = [
                "Product",
                "Equipment",
                "Purpose",
                "Corporate Clients",
                "Classification",
                "LOOKUP"
            ]
            missing_cols = [c for c in required_cols if c not in df_source.columns]
            if missing_cols:
                raise ValueError(f"Missing columns in Product_Cat.xlsx: {missing_cols}")

            df_source = df_source[required_cols]
            logger.info(f"Loaded {len(df_source)} rows from Product_Cat.xlsx")

            # Step 4: Open .xlsb and overwrite the Product_Cat sheet using xlwings
            logger.info("Opening .xlsb workbook with xlwings...")
            app = xw.App(visible=False)
            wb = app.books.open(str(file_path))
            ws = wb.sheets["Product_Cat"]

            # Clear existing data (columns A–F from row 2 downward)
            used_range = ws.used_range
            if used_range.last_cell.row > 1:
                last_row = used_range.last_cell.row
                clear_range = f"A2:F{last_row}"
                ws.range(clear_range).clear_contents()

            # Write new data (without header)
            start_cell = ws.range("A2")
            ws.range(start_cell).value = df_source.values.tolist()

            wb.save()
            wb.close()
            app.quit()
            logger.info(f"Successfully updated Product_Cat sheet in {file_path.name}")

            # Step 5: Read ALL columns from .xlsb Product_Cat sheet using pandas
            logger.info("Reading all columns from Product_Cat sheet in .xlsb...")
            df = pd.read_excel(
                str(file_path),
                sheet_name="Product_Cat",
                engine='pyxlsb'
            )

            # Step 6: Rename target columns
            df = df.rename(columns={
                "SPECIAL CATEGORIES (CONTRACT WISE) *Add here to pick contract wise cat": "SPECIAL CATEGORIES",
                df.columns[-1]: "M"
            })

            logger.info(f"Successfully loaded Product_Cat data: {len(df)} rows, columns: {df.columns.tolist()}")
            return df

        except Exception as e:
            logger.error(f"Failed to load Product_Cat: {e}")
            logger.error(traceback.format_exc())
            raise




    # def load_C1_C2_Working(self):
    #     """Load C1 & C2 Working with error handling."""
    #     try:
    #         logger.info("Loading C1 & C2 Working sheet...")
    #         file_path = self._find_file("Prod. wise Class. of Loans")
            
    #         df = pd.read_excel(
    #             file_path,
    #             sheet_name="C1 & C2 Working",
    #             header=1,
    #             usecols="A, AB, AD"
    #         )
            
    #         # Rename columns
    #         df = df.rename(columns={
    #             df.columns[0]: "Contract No",
    #             df.columns[1]: "Gross Exposure",
    #             df.columns[2]: "PD Category" if len(df.columns) > 2 else None
    #         })
            
    #         # Keep all contract rows (removed filtering of LR/Margin Trading to ensure proper mapping)
    #         logger.info(f"Keeping all contracts including LR and Margin Trading for proper PD Category and Gross Exposure mapping")
            
    #         logger.info(f"Successfully loaded C1 & C2 Working data: {len(df)} rows")
    #         return df
            
    #     except Exception as e:
    #         logger.error(f"Failed to load C1 & C2 Working: {e}")
    #         logger.error(traceback.format_exc())
    #         raise

    def load_po_listing(self):
        """Load PO listing with error handling."""
        try:
            logger.info("Loading Po Listing - Internal file...")
            file_path = self._find_file("Po Listing - Internal")
            df = pd.read_excel(file_path, usecols="H, AH")
            # Assign proper column names: H = Contract No, AH = SELL_PRICE
            df.columns = ["CONTRACT_NO", "SELL_PRICE"]
            df_filtered = df[df["CONTRACT_NO"].notna() & df["SELL_PRICE"].notna()]
            logger.info(f"Successfully loaded PO Listing data: {len(df_filtered)} rows with valid Contract No and SELL_PRICE")
            return df_filtered
        except Exception as e:
            logger.error(f"Failed to load PO Listing: {e}")
            logger.warning("Returning empty DataFrame for PO Listing")
            return pd.DataFrame(columns=["CONTRACT_NO", "SELL_PRICE"])

    def load_portfolio_report_recovery(self):
        """Load portfolio report recovery with error handling."""
        try:
            logger.info("Loading Portfolio Report Recovery - Internal file...")
            file_path = self._find_file("Portfolio Report Recovery - Internal")
            df = pd.read_excel(file_path, usecols="A,B")
            df.columns = ["CONTRACT_NO", "SELL_PRICE"]
            df_filtered = df[df["CONTRACT_NO"].notna()]
            logger.info(f"Successfully loaded Portfolio Report Recovery: {len(df_filtered)} rows")
            return df_filtered
        except FileNotFoundError:
            logger.warning("Portfolio Report Recovery file not found. Returning empty DataFrame.")
            return pd.DataFrame(columns=["CONTRACT_NO", "SELL_PRICE"])
        except Exception as e:
            logger.error(f"Failed to load Portfolio Report Recovery: {e}")
            return pd.DataFrame(columns=["CONTRACT_NO", "SELL_PRICE"])

    def safe_apply(self, df: pd.DataFrame, func, axis=1, default_value=None) -> pd.Series:
        """Safely apply a function to a dataframe with error handling."""
        try:
            return df.apply(func, axis=axis)
        except Exception as e:
            logger.error(f"Error in apply function: {e}")
            logger.error(f"Function: {func.__name__ if hasattr(func, '__name__') else 'lambda'}")
            logger.error(traceback.format_exc())
            if default_value is not None:
                return pd.Series([default_value] * len(df), index=df.index)
            raise

    def classify_collateral_excel_formula(self, row, property_mortgage_list, vehicles_machinery_list):
        """
        Classify collateral following Excel formula EXACTLY:
        =IF(ISNUMBER(MATCH(A3,'YardandProperty List'!D:D,0)),"Immovable Properties",
            IF(ISNUMBER(MATCH(T3,Product_Cat!$P$2:$P$9,0)),"Vehicles and Machinery",
                IF(B3="Margin Trading","Shares and Debt Securities-Listed",
                    IF(B3="FDL","Deposits (Cash-Backed)","Personal and Corporate Guarantees"))))

        Where: A3=Contract No, T3=PD Category, B3=Product
        """
        try:
            # Extract values exactly as in Excel formula
            contract_no = str(row["Contract No"]).strip() if pd.notna(row["Contract No"]) else ""  # A3
            pd_category = str(row["PD Category"]).strip() if pd.notna(row["PD Category"]) else ""  # T3
            product = str(row["Product"]).strip() if pd.notna(row["Product"]) else ""  # B3

            # Log the values being checked for this row
            logger.debug(f"Classifying: Contract={contract_no}, PD_Category={pd_category}, Product={product}")

            # Step 1: IF(ISNUMBER(MATCH(A3,'YardandProperty List'!D:D,0)),"Immovable Properties"
            if contract_no and contract_no in property_mortgage_list:
                logger.debug(f"  → Immovable Properties (Contract {contract_no} in property list)")
                return "Immovable Properties"

            # Step 2: IF(ISNUMBER(MATCH(T3,Product_Cat!$P$2:$P$9,0)),"Vehicles and Machinery"
            if pd_category and pd_category in vehicles_machinery_list:
                logger.debug(f"  → Vehicles and Machinery (PD Category '{pd_category}' in vehicles list)")
                return "Vehicles and Machinery"

            # Step 3: IF(B3="Margin Trading","Shares and Debt Securities-Listed"
            if product == "Margin Trading":
                logger.debug(f"  → Shares and Debt Securities-Listed (Product = 'Margin Trading')")
                return "Shares and Debt Securities-Listed"

            # Step 4: IF(B3="FDL","Deposits (Cash-Backed)"
            if product == "FDL":
                logger.debug(f"  → Deposits (Cash-Backed) (Product = 'FDL')")
                return "Deposits (Cash-Backed)"

            # Step 5: Default case
            logger.debug(f"  → Personal and Corporate Guarantees (default case)")
            return "Personal and Corporate Guarantees"

        except Exception as e:
            logger.error(f"Error classifying collateral for row: {e}")
            logger.error(f"Row data: Contract={row.get('Contract No')}, PD={row.get('PD Category')}, Product={row.get('Product')}")
            return "Personal and Corporate Guarantees"

    def copy_sofp_to_prod_loans(self, picked_date):
        """Copy values D3:E137 from SOFP file to Prod. wise Class. of Loans file and update D2 date."""
        try:
            logger.info("="*80)
            logger.info("COPYING SOFP DATA TO PROD. WISE CLASS. OF LOANS")
            logger.info("="*80)

            # Find SOFP file (starts with "NBD-MF-01-SOFP & SOCI AFL Monthly FS" and ends with ".xlsx")
            sofp_files = [f for f in os.listdir(self.ia_folder)
                         if f.startswith("NBD-MF-01-SOFP & SOCI AFL Monthly FS")
                         and f.endswith(".xlsx")
                         and not f.startswith('~$')]

            if not sofp_files:
                raise FileNotFoundError("SOFP file not found (NBD-MF-01-SOFP & SOCI AFL Monthly FS*.xlsx)")

            sofp_file_path = os.path.join(self.ia_folder, sofp_files[0])
            logger.info(f"Found SOFP file: {sofp_file_path}")

            # Find Prod. wise Class. of Loans file (starts with "Prod. wise Class. of Loans" and ends with ".xlsb")
            prod_files = [f for f in os.listdir(self.ia_folder)
                         if f.startswith("Prod. wise Class. of Loans")
                         and f.endswith(".xlsb")
                         and not f.startswith('~$')]

            if not prod_files:
                raise FileNotFoundError("Prod. wise Class. of Loans file not found (*.xlsb)")

            prod_file_path = os.path.join(self.ia_folder, prod_files[0])
            logger.info(f"Found Prod. wise Class. of Loans file: {prod_file_path}")

            # Use xlwings to read from SOFP file and write to Prod file
            logger.info("Opening files with xlwings...")
            app = xw.App(visible=False)

            try:
                # Open SOFP file
                wb_sofp = app.books.open(sofp_file_path)

                if "NBD-MF-01-SOFP" not in [sheet.name for sheet in wb_sofp.sheets]:
                    wb_sofp.close()
                    raise ValueError(f"Sheet 'NBD-MF-01-SOFP' not found in {sofp_files[0]}")

                ws_sofp = wb_sofp.sheets["NBD-MF-01-SOFP"]

                # Read values from D3:E137 (values only)
                logger.info("Reading range D3:E137 from NBD-MF-01-SOFP sheet in SOFP file...")
                values_to_copy = ws_sofp.range("D3:E137").value
                logger.info(f"Successfully read {len(values_to_copy) if values_to_copy else 0} rows from SOFP file")

                # Close SOFP file
                wb_sofp.close()

                # Open Prod. wise Class. of Loans file
                wb_prod = app.books.open(prod_file_path)

                if "NBD-MF-01-SOFP" not in [sheet.name for sheet in wb_prod.sheets]:
                    wb_prod.close()
                    raise ValueError(f"Sheet 'NBD-MF-01-SOFP' not found in {prod_files[0]}")

                ws_prod = wb_prod.sheets["NBD-MF-01-SOFP"]

                # Update D2 with GUI input date
                logger.info(f"Updating D2 with date: {picked_date}")
                ws_prod.range("D2").value = picked_date

                # Clear range D3:E137
                logger.info("Clearing range D3:E137 in Prod. wise Class. of Loans file...")
                ws_prod.range("D3:E137").clear_contents()

                # Paste the copied values
                logger.info("Pasting values to D3:E137 in Prod. wise Class. of Loans file...")
                if values_to_copy:
                    ws_prod.range("D3").value = values_to_copy
                    logger.info(f"✓ Successfully pasted {len(values_to_copy)} rows to Prod. wise Class. of Loans file")
                else:
                    logger.warning("No values to paste from SOFP file")

                # Save and close
                wb_prod.save()
                wb_prod.close()
                logger.info("✓ Successfully copied SOFP data and updated date in Prod. wise Class. of Loans file")

            finally:
                app.quit()

        except Exception as e:
            logger.error(f"Failed to copy SOFP data to Prod. wise Class. of Loans: {e}")
            logger.error(traceback.format_exc())
            raise

    def build_prod_loans(self, df_disbursement, df_net_portfolio, df_information_request_from_credit,df_summary,
                        df_product_cat, df_C1_C2_Working, df_yard_and_property_list,
                        df_po_listing, df_portfolio_rec):
        """Build production loans dataframe with comprehensive error handling."""

        logger.info("="*80)
        logger.info("STARTING build_prod_loans METHOD")
        logger.info("="*80)
        
        try:
            # Initialize dataframe
            df_prod_loans = pd.DataFrame()
            df_prod_loans["Contract No"] = df_disbursement["CONTRACT NO"]

            logger.info(f"INITIAL CONTRACT COUNT: {len(df_prod_loans)} contracts from disbursement")
            
            # Process each column with error handling
            try:
                df_prod_loans["Disbursed Amount (Net of DC)"] = (
                    pd.to_numeric(df_disbursement["Net Amount"], errors="coerce")
                    .round(0)
                    .astype("Int64")
                )
                logger.info("✓ Disbursed Amount column created")
            except Exception as e:
                logger.error(f"Error creating Disbursed Amount column: {e}")
                df_prod_loans["Disbursed Amount (Net of DC)"] = None
            
            df_prod_loans["Base Rate"] = df_disbursement["Base Rate"]
            
            # Product column with replacement
            try:
                # Extract product codes and show samples for debugging
                df_prod_loans["Product"] = df_prod_loans["Contract No"].astype(str).str[2:4]

                # Show sample contract numbers and their extracted codes
                sample_contracts = df_prod_loans[["Contract No", "Product"]].head(10)
                logger.info(f"Sample contract numbers and extracted product codes:\n{sample_contracts.to_string()}")

                # Show unique product codes before replacement
                unique_products_before = df_prod_loans["Product"].unique()
                logger.info(f"Unique product codes before replacement: {unique_products_before}")

                df_prod_loans["Product"] = df_prod_loans["Product"].replace({
                    "00": "FDL",
                    "rg": "Margin Trading"
                })

                # Debug product classifications
                product_counts = df_prod_loans["Product"].value_counts()
                logger.info("✓ Product column created with replacements")
                logger.info(f"Product distribution: {product_counts.to_dict()}")

                fdl_count = len(df_prod_loans[df_prod_loans["Product"] == "FDL"])
                margin_trading_count = len(df_prod_loans[df_prod_loans["Product"] == "Margin Trading"])
                logger.info(f"FDL contracts: {fdl_count}, Margin Trading contracts: {margin_trading_count}")
            except Exception as e:
                logger.error(f"Error creating Product column: {e}")
                df_prod_loans["Product"] = ""
            
            # Build mapping dictionaries
            try:
                mapping_dicts = {
                    "Client Code": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CLIENT_CODE"])),
                    "Contract Period": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CONTRACT_PERIOD"])),
                    "Contract Amount": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CONTRACT_AMOUNT"])),
                    "Equipment": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["EQT_DESC"])),
                    "Frequency": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CON_RNTFREQ"])),
                    "Contractual Interest Rate": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CON_INTRATE"])),
                    "Purpose": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["PURPOSE"]))
                }
                
                for col, mapping in mapping_dicts.items():
                    df_prod_loans[col] = df_prod_loans["Contract No"].map(mapping).fillna("")
                    logger.info(f"✓ Mapped {col} column ({len([v for v in mapping.values() if v])} mappings)")

                logger.info("✓ All net portfolio mappings completed")
            except Exception as e:
                logger.error(f"Error in net portfolio mappings: {e}")
            
            # Clean Client Code
            df_prod_loans["Client Code"] = df_prod_loans["Client Code"].astype(str).str.replace(r"\.0$", "", regex=True)
            
            # Fill incomplete contracts
            df_prod_loans = self.fill_incomplete_contracts(df_prod_loans, na_contract_bot)
            
            # Corporate Clients classification
            try:
                df_prod_loans["Corporate Clients"] = df_prod_loans["Client Code"].str[0].apply(
                    lambda x: "Corporate Client" if x == "2" else "Non-Corporate"
                )
                logger.info("✓ Corporate Clients classification completed")
            except Exception as e:
                logger.error(f"Error in Corporate Clients classification: {e}")
                df_prod_loans["Corporate Clients"] = "Non-Corporate"
            
            # Calculate rates and tenure
            try:
                df_prod_loans["Contractual Interest Rate"] = pd.to_numeric(df_prod_loans["Contractual Interest Rate"], errors="coerce")
                df_prod_loans["Base Rate"] = pd.to_numeric(df_prod_loans["Base Rate"], errors="coerce")

                # Apply manual rate updates if provided by user
                if self.rate_updates:
                    logger.info(f"Applying {len(self.rate_updates)} manual rate updates...")
                    updated_count = 0
                    for update in self.rate_updates:
                        contract_no = str(update.get("contract_number", "")).strip()
                        contractual_rate = update.get("contractual_rate")
                        base_rate = update.get("base_rate")

                        if not contract_no:
                            continue

                        # Find matching contract(s)
                        mask = df_prod_loans["Contract No"].astype(str).str.strip() == contract_no

                        if mask.any():
                            if contractual_rate is not None:
                                df_prod_loans.loc[mask, "Contractual Interest Rate"] = float(contractual_rate)
                                logger.info(f"Updated Contractual Rate for {contract_no}: {contractual_rate}%")
                            if base_rate is not None:
                                df_prod_loans.loc[mask, "Base Rate"] = float(base_rate)
                                logger.info(f"Updated Base Rate for {contract_no}: {base_rate}%")
                            updated_count += mask.sum()
                        else:
                            logger.warning(f"Contract {contract_no} not found in data - skipping rate update")

                    logger.info(f"✓ Applied manual rate updates to {updated_count} contract(s)")

                df_prod_loans["Minimum Rate (Final)"] = df_prod_loans[["Contractual Interest Rate", "Base Rate"]].min(axis=1)

                df_prod_loans["Contract Period"] = pd.to_numeric(df_prod_loans["Contract Period"], errors="coerce")
                conditions = [
                    df_prod_loans["Frequency"] == "D",
                    df_prod_loans["Frequency"] == "Q",
                    df_prod_loans["Frequency"] == "M",
                    df_prod_loans["Frequency"] == "W"
                ]
                choices = [
                    df_prod_loans["Contract Period"] / 30,
                    df_prod_loans["Contract Period"] * 3,
                    df_prod_loans["Contract Period"],
                    df_prod_loans["Contract Period"] / 4
                ]
                df_prod_loans["Tenure (Months)"] = np.select(conditions, choices, default=0)
                logger.info("✓ Rate and tenure calculations completed")
            except Exception as e:
                logger.error(f"Error in rate/tenure calculations: {e}")
            
            # MSME Classification
            try:
                contract_to_msme = dict(
                    zip(
                        df_information_request_from_credit["Contract No"],
                        df_information_request_from_credit["Micro/Small/Medium"]
                    )
                )
                df_prod_loans["MSME Classification"] = (
                    df_prod_loans["Contract No"]
                    .map(contract_to_msme)
                    .fillna("0")
                    .replace(
                        to_replace=r"\s*/\s*",
                        value="/",
                        regex=True
                    )
                )

                # Optionally, ensure consistent case or formatting
                df_prod_loans["MSME Classification"] = df_prod_loans["MSME Classification"].str.strip()

                logger.info("✓ MSME Classification completed and normalized ('Corporate/Other Enterprises')")
            except Exception as e:
                logger.error(f"Error in MSME Classification: {e}")
                df_prod_loans["MSME Classification"] = "0"
            
            # EIR Calculation with error handling
            try:
                # Ensure numeric input
                df_prod_loans["Minimum Rate (Final)"] = pd.to_numeric(df_prod_loans["Minimum Rate (Final)"], errors="coerce")
                df_prod_loans["Tenure (Months)"] = pd.to_numeric(df_prod_loans["Tenure (Months)"], errors="coerce")

                # Avoid division by zero or invalid tenure
                mask = df_prod_loans["Tenure (Months)"] > 0

                # EIR (%) formula
                df_prod_loans.loc[mask, "EIR (%)"] = (
                    (1 + df_prod_loans.loc[mask, "Minimum Rate (Final)"] / 100 /
                    (12 / df_prod_loans.loc[mask, "Tenure (Months)"])) **
                    (12 / df_prod_loans.loc[mask, "Tenure (Months)"]) - 1
                ) * 100

                # Clean and round
                df_prod_loans["EIR (%)"] = pd.to_numeric(df_prod_loans["EIR (%)"], errors="coerce").round(2)

                # Convert to decimal form for Excel output (divide by 100 only)
                df_prod_loans["EIR (Decimal)"] = (df_prod_loans["EIR (%)"] / 10000).round(6)

                logger.info("✓ EIR calculation completed")

            except Exception as e:
                logger.error(f"Error in EIR calculation: {e}")
                df_prod_loans["EIR (%)"] = ""
                df_prod_loans["EIR (Decimal)"] = ""

            
            # Type of Loans classification
            # Excel formula: =XLOOKUP(A3,Product_Cat!L:L,Product_Cat!M:M,
            #                 IF(OR(B3="MT",B3="Margin Trading"),"Margin Trading Loans",
            #                 IF(OR(B3="FDL",B3="FD Loan"),"Loans against Cash/Deposits",
            #                 XLOOKUP((B3&D3&E3&F3),Product_Cat!F:F,Product_Cat!E:E))))
            try:
                # Create lookup dictionaries from Product_Cat sheet
                # L:L = SPECIAL CATEGORIES (Contract No), M:M = result
                # Clean NaN values before creating dict
                special_cat_df = df_product_cat[["SPECIAL CATEGORIES", "M"]].dropna(subset=["SPECIAL CATEGORIES"])
                special_cat_map = dict(zip(special_cat_df["SPECIAL CATEGORIES"], special_cat_df["M"]))

                # F:F = LOOKUP (Product&Equipment&Purpose&Corporate), E:E = Classification
                lookup_df = df_product_cat[["LOOKUP", "Classification"]].dropna(subset=["LOOKUP"])
                lookup_map = dict(zip(lookup_df["LOOKUP"], lookup_df["Classification"]))

                logger.info(f"Special categories map has {len(special_cat_map)} entries")
                logger.info(f"Lookup map has {len(lookup_map)} entries")

                def classify_loan_type(row):
                    """Classify loan type matching Excel XLOOKUP formula."""
                    contract_no = str(row["Contract No"]) if pd.notna(row["Contract No"]) else ""
                    product = str(row["Product"]) if pd.notna(row["Product"]) else ""
                    equipment = str(row["Equipment"]) if pd.notna(row["Equipment"]) else ""
                    purpose = str(row["Purpose"]) if pd.notna(row["Purpose"]) else ""
                    corporate_clients = str(row["Corporate Clients"]) if pd.notna(row["Corporate Clients"]) else ""

                    # First XLOOKUP: Check special categories by Contract No (A3 in L:L column)
                    if contract_no and contract_no in special_cat_map:
                        return special_cat_map[contract_no]

                    # IF condition 1: Check for Margin Trading (B3="MT" or "Margin Trading")
                    if product in ["MT", "Margin Trading"]:
                        return "Margin Trading Loans"

                    # IF condition 2: Check for FD Loans (B3="FDL" or "FD Loan")
                    if product in ["FDL", "FD Loan"]:
                        return "Loans against Cash/Deposits"

                    # Second XLOOKUP: Use concatenated key (B3&D3&E3&F3) in F:F column
                    lookup_key = f"{product}{equipment}{purpose}{corporate_clients}"
                    result = lookup_map.get(lookup_key)

                    if result is not None and pd.notna(result):
                        return result

                    return ""

                df_prod_loans["Type of Loans (1.3.1.0.0.0)"] = df_prod_loans.apply(classify_loan_type, axis=1)

                # Log statistics and sample results
                null_count = df_prod_loans["Type of Loans (1.3.1.0.0.0)"].isna().sum()
                logger.info(f"Type of Loans classification: {len(df_prod_loans) - null_count} classified, {null_count} null")

                # Show value counts for debugging
                value_counts = df_prod_loans["Type of Loans (1.3.1.0.0.0)"].value_counts()
                logger.info(f"Type of Loans value distribution:\n{value_counts}")

                logger.info("✓ Type of Loans classification completed")
            except Exception as e:
                logger.error(f"Error in Type of Loans classification: {e}")
                logger.error(traceback.format_exc())
                # Create empty column if classification failed
                df_prod_loans["Type of Loans (1.3.1.0.0.0)"] = ""

            # PD Category mapping (MUST happen before Collateral classification)
            # Excel formula: =XLOOKUP(A504,'C1 & C2 Working'!A:A,'C1 & C2 Working'!AD:AD,XLOOKUP(D504,Product_Cat!R:R,Product_Cat!S:S))
            # Using df_summary instead of 'C1 & C2 Working'
            try:
                logger.info("Mapping PD Category (required for Collateral classification)...")
                # First XLOOKUP: Contract No in df_summary -> PD/LGD CATEGORY
                contract_to_pd_category = dict(
                    zip(df_summary["CONTRACT NO"], df_summary["PD/LGD CATEGORY"])
                )
                # Second XLOOKUP: Equipment in Product_Cat R:R (EQT_DESC) -> S:S (CATEGORY)
                equipment_to_category = dict(
                    zip(df_product_cat["EQT_DESC"], df_product_cat["CATEGORY"])
                )
                df_prod_loans["PD Category"] = df_prod_loans["Contract No"].map(contract_to_pd_category)
                df_prod_loans["PD Category"] = df_prod_loans["PD Category"].fillna(
                    df_prod_loans["Equipment"].map(equipment_to_category)
                )
                logger.info(f"✓ PD Category mapping completed. Non-null values: {df_prod_loans['PD Category'].notna().sum()}")
            except Exception as e:
                logger.error(f"Error in PD Category mapping: {e}")
                logger.error(traceback.format_exc())
                df_prod_loans["PD Category"] = ""

            # Collateral/Security Type classification
            try:
                logger.info("="*60)
                logger.info("STARTING COLLATERAL/SECURITY TYPE CLASSIFICATION")
                logger.info("="*60)
                logger.info(f"Current df_prod_loans columns: {df_prod_loans.columns.tolist()}")
                logger.info(f"df_prod_loans shape: {df_prod_loans.shape}")

                # Get lists for classification
                property_mortgage_list = df_yard_and_property_list["Property Mortgage List"].dropna().astype(str).str.strip().tolist()
                logger.info(f"Property Mortgage List loaded: {len(property_mortgage_list)} items")
                if property_mortgage_list:
                    logger.info(f"Sample Property Mortgage contracts: {property_mortgage_list[:5]}")

                # Get Column P values (df_product_cat columns debugging)
                logger.info(f"df_product_cat columns: {df_product_cat.columns.tolist()}")
                logger.info(f"df_product_cat shape: {df_product_cat.shape}")

                # Get Column "PD Category for LTV working" rows 2-9
                # Excel formula: MATCH(T504,Product_Cat!$P$2:$P$9,0)
                # Column P in Excel = "PD Category for LTV working"
                if "PD Category for LTV working" in df_product_cat.columns:
                    p_column = df_product_cat["PD Category for LTV working"]
                    p_column_index = df_product_cat.columns.get_loc("PD Category for LTV working")
                    logger.info(f"Using 'PD Category for LTV working' column at index {p_column_index}")

                    # Show ALL values in column P for debugging
                    logger.info(f"All values in 'PD Category for LTV working' column (first 20 rows):")
                    for i, val in enumerate(p_column.head(20)):
                        logger.info(f"  P[{i+2}] (Excel row {i+2}): {val}")

                    # Get first 9 values (Excel rows 2-9, df rows 0-8)
                    vehicles_machinery_list = p_column.iloc[0:9].dropna().astype(str).str.strip().tolist()
                    vehicles_machinery_list = [x for x in vehicles_machinery_list if x and x != 'nan' and x]
                    logger.info(f"Vehicles & Machinery List (PD Category for LTV working rows 2-9): {vehicles_machinery_list}")
                else:
                    vehicles_machinery_list = []
                    logger.warning("'PD Category for LTV working' column not found in df_product_cat")

                # Log exact matches that will be checked
                logger.info("="*60)
                logger.info("EXACT MATCHES TO BE CHECKED:")
                logger.info(f"  Property Mortgage matches: Contract No must be IN {property_mortgage_list[:10]}...")
                logger.info(f"  Vehicles & Machinery matches: PD Category must be IN {vehicles_machinery_list}")
                logger.info(f"  Margin Trading matches: Product must EQUAL 'Margin Trading'")
                logger.info(f"  FDL matches: Product must EQUAL 'FDL'")
                logger.info("="*60)

                # Check if PD Category column exists (required for classification)
                if "PD Category" not in df_prod_loans.columns:
                    logger.warning("PD Category column not found! Creating empty column...")
                    df_prod_loans["PD Category"] = ""

                # Show sample PD Category values for debugging
                logger.info("Sample PD Category values in df_prod_loans:")
                pd_sample = df_prod_loans["PD Category"].dropna().head(20).tolist()
                logger.info(f"First 20 PD Categories: {pd_sample}")
                logger.info(f"Total non-null PD Categories: {df_prod_loans['PD Category'].notna().sum()}")

                # Show unique PD Category values
                unique_pd_categories = df_prod_loans["PD Category"].dropna().unique()
                logger.info(f"Unique PD Categories ({len(unique_pd_categories)} total): {list(unique_pd_categories)}")

                # Apply classification using EXACT Excel formula logic
                logger.info("Applying Collateral/Security Type classification (Excel formula)...")
                df_prod_loans["Collateral/Security Type"] = self.safe_apply(
                    df_prod_loans,
                    lambda row: self.classify_collateral_excel_formula(
                        row, property_mortgage_list, vehicles_machinery_list
                    ),
                    default_value="Personal and Corporate Guarantees"
                )

                # Log sample classifications for debugging
                logger.info("SAMPLE CLASSIFICATIONS (first 10 rows):")
                for idx in range(min(10, len(df_prod_loans))):
                    row = df_prod_loans.iloc[idx]
                    classification = row["Collateral/Security Type"]
                    logger.info(f"  Row {idx+1}: Contract={row['Contract No']}, PD={row.get('PD Category', 'N/A')}, Product={row.get('Product', 'N/A')} → {classification}")

                # Log classification results
                classification_counts = df_prod_loans["Collateral/Security Type"].value_counts()
                logger.info("="*40)
                logger.info("COLLATERAL CLASSIFICATION RESULTS:")
                logger.info("="*40)
                for ctype, count in classification_counts.items():
                    logger.info(f"  {ctype}: {count}")
                logger.info("="*40)

                # Special check for Vehicles and Machinery
                vehicles_count = classification_counts.get("Vehicles and Machinery", 0)
                logger.info(f"Found {vehicles_count} Vehicles and Machinery contracts (Expected: ~6426)")

                # Show which PD Categories were classified as Vehicles and Machinery
                vehicles_contracts = df_prod_loans[df_prod_loans["Collateral/Security Type"] == "Vehicles and Machinery"]
                if len(vehicles_contracts) > 0:
                    vehicles_pd_categories = vehicles_contracts["PD Category"].value_counts()
                    logger.info("PD Categories classified as Vehicles and Machinery:")
                    for pd_cat, count in vehicles_pd_categories.items():
                        logger.info(f"  '{pd_cat}': {count} contracts")

                if vehicles_count < 6000:  # Threshold for warning
                    # Find PD Categories that are NOT in vehicles_machinery_list but should be
                    non_vm_contracts = df_prod_loans[df_prod_loans["Collateral/Security Type"] != "Vehicles and Machinery"]
                    unique_pd_categories_non_vm = non_vm_contracts["PD Category"].dropna().unique()

                    logger.warning("="*60)
                    logger.warning(f"MISSING {6426 - vehicles_count} Vehicles and Machinery contracts!")
                    logger.warning(f"Expected vehicles_machinery_list: {vehicles_machinery_list}")
                    logger.warning(f"PD Categories in NON-Vehicles contracts: {list(unique_pd_categories_non_vm)}")

                    # Check which PD Categories from contracts are NOT in vehicles_machinery_list
                    missing_pd_cats = [pd for pd in unique_pd_categories_non_vm if pd not in vehicles_machinery_list]
                    logger.warning(f"PD Categories NOT in vehicles_machinery_list: {missing_pd_cats}")
                    logger.warning("="*60)

                    # Show specific contracts that might be missing
                    potential_missing = non_vm_contracts[non_vm_contracts["PD Category"].notna()]
                    if len(potential_missing) > 0:
                        logger.warning("Sample contracts with PD Categories that aren't classified as Vehicles and Machinery:")
                        for idx in range(min(10, len(potential_missing))):
                            row = potential_missing.iloc[idx]
                            pd_cat = row.get('PD Category', 'N/A')
                            in_list = "IN LIST" if pd_cat in vehicles_machinery_list else "NOT IN LIST"
                            logger.warning(f"  Contract={row['Contract No']}, PD='{pd_cat}' ({in_list}), Classification={row['Collateral/Security Type']}")
                else:
                    logger.info(f"✓ Found expected {vehicles_count} Vehicles and Machinery contracts")

            except Exception as e:
                logger.error(f"Error in Collateral/Security Type classification: {e}")
                df_prod_loans["Collateral/Security Type"] = "Personal and Corporate Guarantees"
            
            # Initial Valuation =====================================================
            try:
                logger.info("Starting Initial Valuation processing...")
                
                df_prod_loans["Initial Valuation"] = None
                mask = df_prod_loans["Collateral/Security Type"] == "Vehicles and Machinery"

                # Step 1: Apply PO Listing SELL_PRICE
                if not df_po_listing.empty and "CONTRACT_NO" in df_po_listing.columns and "SELL_PRICE" in df_po_listing.columns:
                    po_mapping = dict(
                        zip(df_po_listing["CONTRACT_NO"].astype(str),
                            pd.to_numeric(df_po_listing["SELL_PRICE"], errors='coerce'))
                    )
                    po_mapping = {k: v for k, v in po_mapping.items() if pd.notna(v) and v > 0}
                    df_prod_loans.loc[mask, "Initial Valuation"] = (
                        df_prod_loans.loc[mask, "Contract No"].astype(str).map(po_mapping)
                    )
                    logger.info(f"✓ Applied Initial Valuation from PO Listing: {df_prod_loans.loc[mask, 'Initial Valuation'].notna().sum()} contracts")
                else:
                    logger.warning("PO Listing data not available or missing required columns")

                # Step 2: Apply Portfolio Recovery (df_portfolio_rec) SELLING_PRICE
                if not df_portfolio_rec.empty and "CONTRACT_NO" in df_portfolio_rec.columns and "SELLING_PRICE" in df_portfolio_rec.columns:
                    rec_mapping = dict(
                        zip(df_portfolio_rec["CONTRACT_NO"].astype(str),
                            pd.to_numeric(df_portfolio_rec["SELLING_PRICE"], errors='coerce'))
                    )
                    rec_mapping = {k: v for k, v in rec_mapping.items() if pd.notna(v) and v > 0}

                    logger.info(f"Portfolio Recovery has {len(rec_mapping)} contracts with SELLING_PRICE")

                    # Compare and keep max value between existing Initial Valuation and portfolio recovery
                    updated_count = 0
                    for idx, row in df_prod_loans.loc[mask].iterrows():
                        contract = str(row["Contract No"])
                        current_val = row["Initial Valuation"]
                        rec_val = rec_mapping.get(contract)

                        # Only consider non-null values for comparison
                        values_to_compare = []
                        if pd.notna(current_val) and current_val > 0:
                            values_to_compare.append(current_val)
                        if pd.notna(rec_val) and rec_val > 0:
                            values_to_compare.append(rec_val)

                        # Get maximum value, or None if no valid values
                        if values_to_compare:
                            max_val = max(values_to_compare)
                            # Only update if the value changed
                            if pd.isna(current_val) or max_val > current_val:
                                df_prod_loans.at[idx, "Initial Valuation"] = max_val
                                updated_count += 1
                                if updated_count <= 5:  # Log first 5 updates
                                    logger.debug(f"Contract {contract}: Updated from {current_val} to {max_val} (PO: {current_val}, Recovery: {rec_val})")

                    logger.info(f"✓ Updated Initial Valuation with max value from PO Listing & Portfolio Recovery ({updated_count} contracts updated)")
                else:
                    logger.warning("Portfolio Recovery data not available or missing required columns")

                # Step 2.5: Keep the maximum Initial Valuation across df_prod_loans (for info)
                max_initial_valuation = df_prod_loans["Initial Valuation"].max()
                logger.info(f"Maximum Initial Valuation across all vehicles: {max_initial_valuation}")

                # Step 3: Extract contracts still missing Initial Valuation for bot processing
                missing_mask = mask & (
                    df_prod_loans["Initial Valuation"].isna() |
                    (df_prod_loans["Initial Valuation"] <= 0)
                )

                # LR-type contracts
                lr_missing_mask = missing_mask & df_prod_loans["Contract No"].astype(str).str.startswith("LR")
                if lr_missing_mask.any():
                    lr_contracts = df_prod_loans.loc[lr_missing_mask, "Contract No"].astype(str).tolist()
                    logger.info(f"Found {len(lr_contracts)} LR-type contracts needing valuation")
                    scienter_bot = self.load_scienter_valuation()
                    if scienter_bot:
                        try:
                            lr_valuations = self.fetch_lr_valuations(lr_contracts, scienter_bot)
                            for contract, value in lr_valuations.items():
                                df_prod_loans.loc[df_prod_loans["Contract No"] == contract, "Initial Valuation"] = value
                            logger.info(f"✓ Applied {len(lr_valuations)} valuations from Scienter bot for LR contracts")
                        except Exception as e:
                            logger.error(f"Error running Scienter bot: {e}")
                    else:
                        logger.warning("Scienter bot not loaded, skipping LR contract valuation")

                # Non-LR contracts
                non_lr_missing_mask = missing_mask & ~df_prod_loans["Contract No"].astype(str).str.startswith("LR")
                if non_lr_missing_mask.any() and initial_valuation_bot:
                    contracts_to_fetch = df_prod_loans.loc[non_lr_missing_mask, "Contract No"].astype(str).tolist()
                    try:
                        response = initial_valuation_bot.run_valuation_bot(contracts_to_fetch)
                        if isinstance(response, dict):
                            for contract, value in response.items():
                                if value and value not in ["", "Not Valued", "#N/A", "N/A"]:
                                    numeric_value = pd.to_numeric(value, errors='coerce')
                                    if pd.notna(numeric_value) and numeric_value > 0:
                                        df_prod_loans.loc[df_prod_loans["Contract No"] == contract, "Initial Valuation"] = numeric_value
                    except Exception as e:
                        logger.error(f"Error running initial_valuation_bot: {e}")

                # Step 4: Log final stats
                total_vehicles = mask.sum()
                valued_count = df_prod_loans.loc[mask, "Initial Valuation"].notna().sum()
                missing_count = total_vehicles - valued_count

                logger.info("="*60)
                logger.info(f"Total Vehicles & Machinery: {total_vehicles}")
                logger.info(f"Total Valued: {valued_count} ({valued_count/total_vehicles*100:.1f}%)")
                logger.info(f"Total Missing: {missing_count} ({missing_count/total_vehicles*100:.1f}%)")
                logger.info("="*60)
                logger.info("✓ Initial Valuation processing completed")

            except Exception as e:
                logger.error(f"Error in Initial Valuation processing: {e}")
                logger.error(traceback.format_exc())
                df_prod_loans["Initial Valuation"] = None
                        

            # Other calculations with error handling
            try:
                # Annual Interest Cost ========================================================
                df_prod_loans["Annual Interest Cost"] = (
                    pd.to_numeric(df_prod_loans["Contract Amount"], errors="coerce") *
                    pd.to_numeric(df_prod_loans["EIR (%)"], errors="coerce") / 100
                )
                df_prod_loans["Annual Interest Cost"] = df_prod_loans["Annual Interest Cost"].round(0).astype("Int64")
                logger.info("✓ Annual Interest Cost calculated")

                # Gross Exposure ===========================================================
                if not df_C1_C2_Working.empty and "Contract No" in df_C1_C2_Working.columns and "Gross Exposure" in df_C1_C2_Working.columns:
                    contract_to_gross_exposure = dict(
                        zip(df_C1_C2_Working["Contract No"], df_C1_C2_Working["Gross Exposure"])
                    )
                    df_prod_loans["Gross Exposure"] = df_prod_loans["Contract No"].map(contract_to_gross_exposure)
                    df_prod_loans["Gross Exposure"] = df_prod_loans["Gross Exposure"].fillna(df_prod_loans["Contract Amount"])
                else:
                    logger.warning("df_C1_C2_Working is empty or missing required columns, using Contract Amount as Gross Exposure")
                    df_prod_loans["Gross Exposure"] = df_prod_loans["Contract Amount"]

                df_prod_loans["Gross Exposure"] = (
                    pd.to_numeric(df_prod_loans["Gross Exposure"], errors="coerce")
                    .round(0)
                    .astype("Int64")
                )
                logger.info("✓ Gross Exposure calculated")
                
                # Normal/Concessionary classification ===============================================================
                prod_file_path = self._find_file("Prod. wise Class. of Loans")
                df_nbd_sheet = pd.read_excel(prod_file_path, sheet_name="NBD-MF-23-IA", header=None)
                l20_value = df_nbd_sheet.iloc[19, 11]
                
                df_staff_loan = pd.read_excel(prod_file_path, sheet_name="Staff_Loan", usecols="F", skiprows=3)
                icam_list = df_staff_loan["ICAM"].dropna().tolist()
                
                def calc_normal_concessionary(row):
                    try:
                        if row["Minimum Rate (Final)"] > l20_value:
                            return "Normal"
                        elif row["Client Code"] in icam_list:
                            return "Concessionary"
                        else:
                            return "Normal"
                    except:
                        return "Normal"
                
                df_prod_loans["Normal/Concessionary"] = self.safe_apply(
                    df_prod_loans, calc_normal_concessionary, default_value="Normal"
                )
                logger.info("✓ Normal/Concessionary classification completed")

                # PD Category already mapped earlier (before Collateral classification)
                logger.info("✓ PD Category mapping already completed earlier")
                
            except Exception as e:
                logger.error(f"Error in final calculations: {e}")
                logger.error(traceback.format_exc())
            
            logger.info("="*80)
            logger.info(f"build_prod_loans COMPLETED - Generated {len(df_prod_loans)} rows")
            logger.info(f"FINAL COLUMNS: {df_prod_loans.columns.tolist()}")
            logger.info("="*80)

            # Add LTV % column ================================================
            def calculate_ltv(row):
                try:
                    if row["Collateral/Security Type"] == "Vehicles and Machinery":
                        contract_amount = pd.to_numeric(row["Contract Amount"], errors='coerce')
                        initial_valuation = pd.to_numeric(row["Initial Valuation"], errors='coerce')

                        if (pd.notna(contract_amount) and pd.notna(initial_valuation) and
                            initial_valuation != 0 and initial_valuation is not None):
                            ltv_ratio = (contract_amount / initial_valuation) * 100  # Convert to percentage
                            return round(ltv_ratio, 2)  # Round to 2 decimal places
                    return 0.00
                except Exception as e:
                    logger.debug(f"Error calculating LTV for contract {row.get('Contract No', 'Unknown')}: {e}")
                    return 0.00

            df_prod_loans["LTV %"] = df_prod_loans.apply(calculate_ltv, axis=1)

            # Apply manual LTV updates if provided by user
            if self.ltv_updates:
                logger.info(f"Applying {len(self.ltv_updates)} manual LTV updates...")
                updated_count = 0
                for update in self.ltv_updates:
                    contract_no = str(update.get("contract_number", "")).strip()
                    ltv_percentage = update.get("ltv_percentage")

                    if not contract_no:
                        continue

                    # Find matching contract(s)
                    mask = df_prod_loans["Contract No"].astype(str).str.strip() == contract_no

                    if mask.any():
                        if ltv_percentage is not None:
                            df_prod_loans.loc[mask, "LTV %"] = float(ltv_percentage)
                            logger.info(f"Updated LTV % for {contract_no}: {ltv_percentage}%")
                        updated_count += mask.sum()
                    else:
                        logger.warning(f"Contract {contract_no} not found in data - skipping LTV update")

                logger.info(f"✓ Applied manual LTV updates to {updated_count} contract(s)")

            # WALTV % ==========================================================
            try:
                # Check if Gross Exposure column exists
                if "Gross Exposure" not in df_prod_loans.columns:
                    logger.warning("Gross Exposure column not found, creating with Contract Amount values")
                    df_prod_loans["Gross Exposure"] = df_prod_loans["Contract Amount"]

                # Ensure Gross Exposure is numeric
                df_prod_loans["Gross Exposure"] = pd.to_numeric(df_prod_loans["Gross Exposure"], errors='coerce').fillna(0)

                # Compute total gross exposure grouped by Collateral/Security Type
                total_gross_by_type = df_prod_loans.groupby("Collateral/Security Type")["Gross Exposure"].transform("sum")

                # Formula directly
                df_prod_loans["WALTV %"] = 0.00
                mask = df_prod_loans["Collateral/Security Type"] == "Vehicles and Machinery"

                # Only calculate if we have valid data
                if mask.any() and total_gross_by_type[mask].sum() > 0:
                    waltv_calculation = (
                        df_prod_loans.loc[mask, "Gross Exposure"] / total_gross_by_type[mask]
                    ) * (df_prod_loans.loc[mask, "LTV %"] / 100)  # Convert LTV % back to ratio for calculation
                    df_prod_loans.loc[mask, "WALTV %"] = (waltv_calculation * 100).round(2)  # Convert to percentage with 2 decimals

                logger.info("✓ WALTV % calculation completed")
            except Exception as e:
                logger.error(f"Error in WALTV % calculation: {e}")
                df_prod_loans["WALTV %"] = 0


            # Check specifically for the missing columns
            required_columns = ["Collateral/Security Type", "Initial Valuation"]
            missing_columns = [col for col in required_columns if col not in df_prod_loans.columns]
            if missing_columns:
                logger.error(f"MISSING REQUIRED COLUMNS: {missing_columns}")
            else:
                logger.info("✓ All required columns present")

            
            # List of columns to divide by 100 for Excel percentage formatting
            # EIR (%) is already calculated as percentage (e.g., 10.23 for 10.23%), so it needs to be divided by 100
            # so that Excel percentage formatting shows it correctly (0.1023 → 10.23%)
            cols_to_scale = ["Minimum Rate (Final)", "Contractual Interest Rate", "Base Rate", "EIR (%)", "LTV %", "WALTV %"]

            # Divide numeric values by 100
            for col in cols_to_scale:
                if col in df_prod_loans.columns:
                    df_prod_loans[col] = pd.to_numeric(df_prod_loans[col], errors="coerce") / 100

            # Get minimum rate threshold from master data and validate rates
            try:
                # Try to get from master data first, then use command line override if provided
                minimum_rate_final = self.get_minimum_rate_threshold(df_product_cat)

                # Use command line override if master data extraction failed
                # if hasattr(self, 'minimum_rate_override') and self.minimum_rate_override is not None:
                #     logger.info(f"Using command line minimum rate override: {self.minimum_rate_override}%")
                #     minimum_rate_final = self.minimum_rate_override / 100

                exception_report_path = os.path.join(self.ia_folder, "Minimum_Rate_Exceptions.xlsx")

                # The threshold is in decimal format (e.g., 0.0822 for 8.22%)
                # Minimum Rate (Final) values are in percentage format (e.g., 2.3 for 2.3%)
                # We need to convert threshold to percentage format OR divide Minimum Rate by 100
                # For consistency, we'll multiply threshold by 100 to compare in percentage format
                threshold = minimum_rate_final /100

                logger.info(f"Validating Minimum Rate (Final) against threshold: {threshold}%")
                logger.info(f"Threshold value for comparison: {threshold}")

                # Ensure required columns exist
                required_cols = ["Contract No", "Minimum Rate (Final)"]
                missing_cols = [col for col in required_cols if col not in df_prod_loans.columns]
                if missing_cols:
                    logger.warning(f"Missing columns for rate validation: {missing_cols}")
                else:
                    # Convert Minimum Rate (Final) to numeric for comparison
                    df_prod_loans["Minimum Rate (Final)"] = pd.to_numeric(df_prod_loans["Minimum Rate (Final)"], errors='coerce')

                    # Show sample values for debugging
                    sample_rates = df_prod_loans["Minimum Rate (Final)"].dropna().head(10)
                    logger.info(f"Sample Minimum Rate (Final) values: {sample_rates.tolist()}")
                    logger.info(f"Min rate in data: {df_prod_loans['Minimum Rate (Final)'].min()}, Max rate in data: {df_prod_loans['Minimum Rate (Final)'].max()}")
                    logger.info(f"Comparing against threshold: {threshold}")

                    # Find exceptions where rate > threshold (both in percentage format)
                    exceptions_mask = df_prod_loans["Minimum Rate (Final)"] > threshold
                    exceptions = df_prod_loans[exceptions_mask].copy()

                    logger.info(f"Found {len(exceptions)} contracts with rates > {threshold}% out of {len(df_prod_loans)} total contracts")

                    if not exceptions.empty:
                        # Add reference column for clarity (in percentage format)
                        exceptions["Expected Max Rate"] = threshold
                        exceptions["Rate Difference"] = exceptions["Minimum Rate (Final)"] - threshold

                        # Reorder columns for better readability
                        exception_cols = ["Contract No", "Product", "Client Code", "Minimum Rate (Final)",
                                        "Expected Max Rate", "Rate Difference", "Equipment", "Purpose"]
                        available_cols = [col for col in exception_cols if col in exceptions.columns]
                        remaining_cols = [col for col in exceptions.columns if col not in available_cols]
                        exceptions = exceptions[available_cols + remaining_cols]

                        # Export to Excel in same folder
                        exceptions.to_excel(exception_report_path, index=False)
                        logger.warning(f"Exception report generated at {exception_report_path} with {len(exceptions)} records exceeding {threshold}%")
                        logger.warning(f"Contracts with high rates: {exceptions['Contract No'].tolist()[:10]}...")  # Show first 10
                    else:
                        logger.info(f"✓ All {len(df_prod_loans)} contracts have rates within threshold ({threshold}%)")

            except Exception as e:
                logger.error(f"Error in minimum rate validation: {e}")
                logger.error(traceback.format_exc())

            # LTV % Validation - Check for LTV > 100% =====================
            try:
                ltv_exception_report_path = os.path.join(self.ia_folder, "LTV_validation_exceptions.xlsx")

                logger.info("Validating LTV % for contracts exceeding 100%")

                # Ensure required columns exist
                if "LTV %" not in df_prod_loans.columns:
                    logger.warning("LTV % column not found, skipping LTV validation")
                else:
                    # Convert LTV % to numeric for comparison
                    df_prod_loans["LTV %"] = pd.to_numeric(df_prod_loans["LTV %"], errors='coerce')

                    # Show sample values for debugging
                    sample_ltv = df_prod_loans["LTV %"].dropna().head(10)
                    logger.info(f"Sample LTV % values: {sample_ltv.tolist()}")
                    logger.info(f"Min LTV in data: {df_prod_loans['LTV %'].min()}, Max LTV in data: {df_prod_loans['LTV %'].max()}")

                    # Find exceptions where LTV > 100%
                    ltv_exceptions_mask = df_prod_loans["LTV %"] > 1
                    ltv_exceptions = df_prod_loans[ltv_exceptions_mask].copy()

                    logger.info(f"Found {len(ltv_exceptions)} contracts with LTV > 100% out of {len(df_prod_loans)} total contracts")

                    if not ltv_exceptions.empty:
                        # Add reference column for clarity
                        ltv_exceptions["Expected Max LTV %"] = 1
                        ltv_exceptions["LTV Difference"] = ltv_exceptions["LTV %"] - 1

                        # Reorder columns for better readability
                        ltv_exception_cols = ["Contract No", "Product", "Client Code", "LTV %",
                                            "Expected Max LTV %", "LTV Difference", "Equipment",
                                            "Initial Valuation", "Outstanding Amount"]
                        available_cols = [col for col in ltv_exception_cols if col in ltv_exceptions.columns]
                        remaining_cols = [col for col in ltv_exceptions.columns if col not in available_cols]
                        ltv_exceptions = ltv_exceptions[available_cols + remaining_cols]

                        # Export to Excel in same folder
                        ltv_exceptions.to_excel(ltv_exception_report_path, index=False)
                        logger.warning(f"LTV exception report generated at {ltv_exception_report_path} with {len(ltv_exceptions)} records exceeding 100%")
                        logger.warning(f"Contracts with high LTV: {ltv_exceptions['Contract No'].tolist()[:10]}...")  # Show first 10
                    else:
                        logger.info(f"✓ All {len(df_prod_loans)} contracts have LTV within threshold (100%)")

            except Exception as e:
                logger.error(f"Error in LTV validation: {e}")
                logger.error(traceback.format_exc())

            # Mask for FDL and Margin Trading rows (case-insensitive)
            mask_special = df_prod_loans["Product"].str.strip().str.lower().isin(["fdl", "margin trading"])

            # ---- Step 1: Apply transformations to special rows ----

            # Columns to clear
            cols_to_clear = ["Client Code", "Equipment", "Purpose", "Frequency", "Contract Period", "Tenure (Months)"]
            df_prod_loans.loc[mask_special, cols_to_clear] = ""

            # Assign Base Rate to interest columns for the special rows
            interest_cols = ["Minimum Rate (Final)", "Contractual Interest Rate", "EIR (%)"]
            df_prod_loans.loc[mask_special, interest_cols] = df_prod_loans.loc[mask_special, "Base Rate"]

            # Assign Contract Amount and Annual Interest Cost = Disbursed Amount (Net of DC)
            df_prod_loans.loc[mask_special, ["Contract Amount", "Annual Interest Cost"]] = df_prod_loans.loc[mask_special, "Disbursed Amount (Net of DC)"]

            # Additional adjustments for Margin Trading only
            mask_margin = df_prod_loans["Product"].str.strip().str.lower() == "margin trading"
            df_prod_loans.loc[mask_margin, "Corporate Clients"] = ""
            df_prod_loans.loc[mask_margin, "MSME Classification"] = "Small"

            # ---- Step 2: Move special rows to the bottom ----

            # Split into normal and special rows
            df_normal = df_prod_loans[~mask_special]
            df_special = df_prod_loans[mask_special]

            # Concatenate: normal rows + special rows (no empty row separator)
            df_prod_loans = pd.concat([df_normal, df_special], ignore_index=True)


            # Columns to copy
            cols_to_copy = ["Product", "Equipment", "Purpose", "Corporate Clients"]

            # Filter rows where 'Type of Loans (1.3.1.0.0.0)' is '#N/A'
            mask_na_type = df_prod_loans["Type of Loans (1.3.1.0.0.0)"] == "#N/A"

            # Create new dataframe with the selected columns
            df_add_product_type = df_prod_loans.loc[mask_na_type, cols_to_copy].copy()

            # Remove duplicate rows (all columns duplicate)
            df_add_product_type.drop_duplicates(inplace=True)

            # Optional: reset index for clean sequential index
            df_add_product_type.reset_index(drop=True, inplace=True)


            # Rearrange columns in the specified order
            desired_column_order = [
                "Contract No",
                "Product",
                "Client Code",
                "Equipment",
                "Purpose",
                "Corporate Clients",
                "Type of Loans (1.3.1.0.0.0)",
                "Normal/Concessionary",
                "Frequency",
                "Contract Period",
                "Tenure (Months)",
                "Minimum Rate (Final)",
                "Contractual Interest Rate",
                "Base Rate",
                "Disbursed Amount (Net of DC)",
                "Contract Amount",
                "Annual Interest Cost",
                "EIR (%)",
                "Gross Exposure",
                "PD Category",
                "Collateral/Security Type",
                "Initial Valuation",
                "LTV %",
                "WALTV %",
                "MSME Classification"
            ]

            # Only include columns that exist in the dataframe
            existing_columns = [col for col in desired_column_order if col in df_prod_loans.columns]

            # Add any remaining columns that weren't specified
            remaining_columns = [col for col in df_prod_loans.columns if col not in existing_columns]
            final_column_order = existing_columns + remaining_columns

            df_prod_loans = df_prod_loans[final_column_order]
            logger.info(f"✓ Columns rearranged. Final order: {df_prod_loans.columns.tolist()}")

            # Final debugging - check what contracts we ended up with
            final_product_counts = df_prod_loans["Product"].value_counts()
            logger.info(f"FINAL CONTRACT COUNT: {len(df_prod_loans)} contracts")
            logger.info(f"FINAL Product distribution: {final_product_counts.to_dict()}")

            final_fdl_count = len(df_prod_loans[df_prod_loans["Product"] == "FDL"])
            final_margin_trading_count = len(df_prod_loans[df_prod_loans["Product"] == "Margin Trading"])
            logger.info(f"FINAL FDL contracts: {final_fdl_count}, FINAL Margin Trading contracts: {final_margin_trading_count}")

            print(df_prod_loans["Minimum Rate (Final)"])
            return df_prod_loans
            
        except Exception as e:
            logger.error(f"CRITICAL ERROR in build_prod_loans: {e}")
            logger.error(traceback.format_exc())
            raise
    
    
    
    def save_report(self, df_prod_loans, picked_date):
        """Paste data to IA Working sheet using fast pandas Excel writing,
        check Product_Cat classification for missing values,
        and verify C39 of NBD-MF-23-IA sheet after pasting data."""
        try:
            # Find the Prod. wise Class. of Loans file
            prod_file_path = self._find_file("Prod. wise Class. of Loans")

            # Validate dataframe before saving
            if df_prod_loans.empty:
                logger.warning("DataFrame is empty. Nothing to paste.")
                return prod_file_path

            logger.info(f"Fast pasting {len(df_prod_loans)} rows x {len(df_prod_loans.columns)} columns to IA Working sheet")

            # Clean dataframe values
            df_clean = df_prod_loans.copy()
            for col in df_clean.columns:
                if pd.api.types.is_numeric_dtype(df_clean[col]):
                    df_clean[col] = df_clean[col].fillna(0)
                else:
                    df_clean[col] = df_clean[col].fillna("")

            # Extract unique combinations and paste to Product_Cat sheet
            logger.info("Extracting unique Product, Equipment, Purpose, Corporate Clients combinations...")
            cols_for_product_cat = ["Product", "Equipment", "Purpose", "Corporate Clients"]
            df_unique_combinations = df_clean[cols_for_product_cat].drop_duplicates().reset_index(drop=True)
            logger.info(f"Found {len(df_unique_combinations)} unique combinations")

            # Paste to Product_Cat sheet
            try:
                import xlwings as xw

                # Use xlwings to open the workbook
                app = xw.App(visible=False, add_book=False)
                app.display_alerts = False
                app.screen_updating = False

                try:
                    workbook = app.books.open(os.path.abspath(prod_file_path))
                    product_cat_sheet = workbook.sheets["Product_Cat"]

                    # Find the last row with data in Product_Cat sheet (to append below it)
                    used_range = product_cat_sheet.used_range
                    last_row = used_range.last_cell.row
                    logger.info(f"Product_Cat sheet currently has {last_row} rows (from Product_Cat.xlsx sync)")

                    # Append new unique combinations below existing data (no clearing)
                    if not df_unique_combinations.empty:
                        start_row = last_row + 1
                        product_cat_sheet.range(f'A{start_row}').value = df_unique_combinations.values
                        num_rows = len(df_unique_combinations)
                        logger.info(f"✓ Appended {num_rows} new unique combinations to Product_Cat sheet (rows {start_row} to {start_row + num_rows - 1})")
                    else:
                        logger.info("No new unique combinations to append")

                    workbook.save()

                    # Save Product_Cat sheet columns A:F to Product_Cat.xlsx (overwrite existing)
                    try:
                        # Get the updated last row after appending
                        final_last_row = product_cat_sheet.used_range.last_cell.row
                        logger.info(f"Product_Cat sheet has {final_last_row} rows before processing")

                        # Clean and rewrite LOOKUP column with formula
                        logger.info("Cleaning and rewriting LOOKUP column formulas in Product_Cat sheet...")
                        lookup_col = "F"  # LOOKUP column

                        # Clear LOOKUP column from row 2 onwards (keep header)
                        if final_last_row > 1:
                            clear_range = f"F2:F{final_last_row}"
                            product_cat_sheet.range(clear_range).clear_contents()
                            logger.info(f"Cleared LOOKUP column range {clear_range}")

                        # Write formula to all rows from row 2 onwards
                        # Formula: =A2&B2&C2&D2 (concatenate Product, Equipment, Purpose, Corporate Clients)
                        formula_count = 0
                        for row in range(2, final_last_row + 1):
                            cell = product_cat_sheet.range(f"F{row}")
                            # Set formula: concatenate columns A, B, C, D for current row
                            cell.formula = f"=A{row}&B{row}&C{row}&D{row}"
                            formula_count += 1

                        logger.info(f"Applied LOOKUP formula (=A&B&C&D) to {formula_count} rows from row 2 to {final_last_row}")

                        # Save to calculate formulas
                        workbook.save()
                        logger.info("Saved workbook to calculate LOOKUP formulas")

                        # Read columns A:F from row 1 (header) to last row with calculated LOOKUP values
                        product_cat_range = product_cat_sheet.range(f'A1:F{final_last_row}')
                        product_cat_data = product_cat_range.value

                        # Convert to DataFrame
                        df_product_cat_export = pd.DataFrame(product_cat_data[1:], columns=product_cat_data[0])
                        logger.info(f"Loaded {len(df_product_cat_export)} rows from Product_Cat sheet")

                        # Remove duplicate rows based on LOOKUP column
                        original_count = len(df_product_cat_export)
                        df_product_cat_export = df_product_cat_export.drop_duplicates(subset=['LOOKUP'], keep='first')
                        removed_count = original_count - len(df_product_cat_export)
                        logger.info(f"Removed {removed_count} duplicate rows based on LOOKUP column")
                        logger.info(f"Remaining rows after deduplication: {len(df_product_cat_export)}")

                        # Clear A-F columns from row 2 onward in Product_Cat sheet
                        if final_last_row > 1:
                            clear_range = f"A2:F{final_last_row}"
                            product_cat_sheet.range(clear_range).clear_contents()
                            logger.info(f"Cleared range {clear_range} in Product_Cat sheet")

                        # Paste deduplicated data back (without header)
                        if len(df_product_cat_export) > 0:
                            product_cat_sheet.range("A2").value = df_product_cat_export.values.tolist()
                            logger.info(f"Pasted {len(df_product_cat_export)} deduplicated rows back to Product_Cat sheet")

                        # Save the deduplicated data in the workbook
                        workbook.save()
                        logger.info("Saved Product_Cat sheet with deduplicated data")

                        # Save to Product_Cat.xlsx (overwrite)
                        base_dir = Path(prod_file_path).parent
                        product_cat_xlsx_path = base_dir / "Product_Cat.xlsx"

                        # Remove existing file if it exists
                        if product_cat_xlsx_path.exists():
                            os.remove(product_cat_xlsx_path)
                            logger.info(f"Removed existing Product_Cat.xlsx")

                        # Save new file with deduplicated data
                        df_product_cat_export.to_excel(product_cat_xlsx_path, sheet_name="Sheet1", index=False)
                        logger.info(f"✓ Saved {len(df_product_cat_export)} deduplicated rows to Product_Cat.xlsx")

                    except Exception as export_error:
                        logger.error(f"Failed to export Product_Cat to xlsx: {export_error}")
                        logger.error(traceback.format_exc())

                    workbook.close()
                finally:
                    app.quit()

            except Exception as product_cat_error:
                logger.error(f"Failed to paste data to Product_Cat sheet: {product_cat_error}")
                logger.error(traceback.format_exc())

            # Convert to absolute path
            abs_file_path = os.path.abspath(prod_file_path)
            logger.info(f"Working with file: {abs_file_path}")

            if not os.path.exists(abs_file_path):
                raise FileNotFoundError(f"File not found: {abs_file_path}")

            try:
                import xlwings as xw

                # Use xlwings to open the workbook
                app = xw.App(visible=False, add_book=False)
                app.display_alerts = False
                app.screen_updating = False

                try:
                    workbook = app.books.open(abs_file_path)

                    # === Step 1: Update IA Working sheet ===
                    worksheet = workbook.sheets["IA Working"]

                    # Clear existing data from row 3 onwards (keep headers in rows 1-2)
                    # Clear only columns A:Z to preserve any formulas/data in other columns
                    used_range = worksheet.used_range
                    last_row = used_range.last_cell.row

                    logger.info(f"IA Working sheet: Clearing columns A:Z from row 3 to {last_row}")

                    if last_row >= 3:
                        # Clear only columns A through Z from row 3 to the last used row
                        worksheet.range(f'A3:Z{last_row}').clear_contents()
                        logger.info(f"Cleared range A3:Z{last_row}")

                    # Write new data starting from row 3
                    if not df_clean.empty:
                        logger.info(f"Pasting {len(df_clean)} rows to IA Working sheet (final row will be {3 + len(df_clean) - 1})")
                        worksheet.range('A3').value = df_clean.values
                        logger.info(f"✓ Successfully pasted data to IA Working sheet")

                    # === Step 2: Verify C39 in NBD-MF-23-IA sheet ===
                    try:
                        ia_sheet = workbook.sheets["NBD-MF-23-IA"]
                        c39_value = ia_sheet.range("C39").value
                        logger.info(f"C39 value in NBD-MF-23-IA: {c39_value}")

                        if c39_value not in (None, 0, "0"):
                            exception_data = pd.DataFrame(
                                [{"Sheet": "NBD-MF-23-IA", "Cell": "C39", "Value": c39_value}]
                            )
                            exception_filepath = os.path.join(
                                os.path.dirname(abs_file_path), "NBD-MF-23-IA_exceptions.xlsx"
                            )
                            exception_data.to_excel(exception_filepath, index=False)
                            logger.warning(f"⚠ Exception found (C39 != 0). Saved file: {exception_filepath}")
                        else:
                            logger.info("✓ C39 is 0. No exceptions.")

                    except Exception as check_error:
                        logger.error(f"Failed to check C39 value: {check_error}")

                    # Save and close workbook
                    workbook.save()
                    workbook.close()
                finally:
                    app.quit()

            except Exception as xlwings_error:
                logger.error(f"xlwings method failed: {xlwings_error}")
                logger.error(traceback.format_exc())
                raise

            # === Step 3: Rename the file with date format ===
            try:
                month_year = picked_date.strftime("%b %Y")
                file_dir = os.path.dirname(abs_file_path)
                new_filename = f"Prod. wise Class. of Loans - {month_year}.xlsb"
                new_file_path = os.path.join(file_dir, new_filename)

                os.rename(abs_file_path, new_file_path)
                logger.info(f"✓ File renamed to: {new_filename}")
                return new_file_path

            except Exception as e:
                logger.error(f"Failed to rename file: {e}")
                return abs_file_path

        except Exception as e:
            logger.error(f"Failed to paste data to IA Working sheet: {e}")
            logger.error(traceback.format_exc())
            raise


    
    def load_all(self, picked_date):
        """Orchestrator method with comprehensive error handling and recovery."""
        # Parse the date if it's not already a date object
        try:
            picked_date = self.parse_date(picked_date)
        except ValueError as e:
            logger.error(f"Date parsing error: {e}")
            raise

        logger.info("="*80)
        logger.info(f"Starting load_all process for date: {picked_date}")
        logger.info("="*80)
        
        loaded_data = {}
        errors = []
        
        # Define loading steps with fallback options
        loading_steps = [
            ("df_disbursement", self.load_disbursement, True),  # Required
            ("df_net_portfolio", self.load_net_portfolio, True),  # Required
            ("df_summary", self.load_summary, True),
            ("df_po_listing", self.load_po_listing, False),  # Optional
            ("df_information_request_from_credit", self.load_information_request_from_credit, True),  # Required
            ("df_product_cat", self.load_product_cat, True),  # Required
            # ("df_C1_C2_Working", self.load_C1_C2_Working, True),  # Required
            ("df_yard_and_property_list", self.load_yard_and_Property_List, True),  # Required
            ("df_portfolio_rec", self.load_portfolio_rec, False),  # Optional
        ]
        
        # Load each dataset with error handling
        for name, loader_func, is_required in loading_steps:
            try:
                logger.info(f"Loading {name}...")
                loaded_data[name] = loader_func()
                logger.info(f"✓ Successfully loaded {name}")
            except Exception as e:
                error_msg = f"Failed to load {name}: {e}"
                logger.error(error_msg)
                errors.append((name, str(e)))
                
                if is_required:
                    logger.error(f"CRITICAL: Required dataset {name} could not be loaded")
                    raise RuntimeError(f"Cannot proceed without {name}. Error: {e}")
                else:
                    logger.warning(f"Optional dataset {name} failed to load, using empty DataFrame")
                    loaded_data[name] = pd.DataFrame()
        
        # Check if we have minimum required data
        required_datasets = ["df_disbursement", "df_net_portfolio", "df_product_cat"]
        missing_required = [ds for ds in required_datasets if ds not in loaded_data or loaded_data[ds].empty]
        
        if missing_required:
            error_msg = f"Missing required datasets: {missing_required}"
            logger.error(error_msg)
            raise RuntimeError(error_msg)
        
        # Copy SOFP data to Prod. wise Class. of Loans before building production loans
        try:
            logger.info("Copying SOFP data to Prod. wise Class. of Loans...")
            self.copy_sofp_to_prod_loans(picked_date)
            logger.info("✓ SOFP data copy completed successfully")
        except Exception as e:
            logger.error(f"Failed to copy SOFP data: {e}")
            logger.error(traceback.format_exc())
            raise

        # Build production loans
        try:
            logger.info("Starting build_prod_loans process...")
            df_prod_loans = self.build_prod_loans(
                loaded_data.get("df_disbursement"),
                loaded_data.get("df_net_portfolio"),
                loaded_data.get("df_information_request_from_credit", pd.DataFrame()),
                loaded_data.get("df_summary", pd.DataFrame()),
                loaded_data.get("df_product_cat"),
                loaded_data.get("df_C1_C2_Working", pd.DataFrame()),
                loaded_data.get("df_yard_and_property_list", pd.DataFrame()),
                loaded_data.get("df_po_listing", pd.DataFrame()),
                loaded_data.get("df_portfolio_rec", pd.DataFrame())
            )
            logger.info("✓ build_prod_loans completed successfully")
        except Exception as e:
            logger.error(f"Failed to build production loans: {e}")
            logger.error(traceback.format_exc())
            raise
        
        # Save report
        try:
            report_path = self.save_report(df_prod_loans, picked_date)
        except Exception as e:
            logger.error(f"Failed to save report: {e}")
            raise
        
        # Prepare return data
        month_year = picked_date.strftime("%b %Y")

        # Log summary
        logger.info("="*80)
        logger.info("PROCESS SUMMARY")
        logger.info(f"Date processed: {picked_date}")
        logger.info(f"Total rows generated: {len(df_prod_loans)}")
        logger.info(f"Data pasted to IA Working sheet in: {report_path}")
        if errors:
            logger.warning(f"Errors encountered during processing:")
            for name, error in errors:
                logger.warning(f"  - {name}: {error}")
        logger.info("="*80)

        return {
            "df_Prod_wise_Class_of_Loans": df_prod_loans,
            "df_Disbursement": loaded_data.get("df_disbursement"),
            "df_Net_Portfolio": loaded_data.get("df_net_portfolio"),
            "df_PO_Listing": loaded_data.get("df_po_listing"),
            "df_information_request_from_credit": loaded_data.get("df_information_request_from_credit"),
            "final_output_path": report_path,
            "report_path": report_path,
            "processing_errors": errors
        }
    
    def get_minimum_rate_threshold(self, df_product_cat):
        """
        Extract minimum rate threshold from Master_Data.xlsx file.

        Args:
            df_product_cat: Product category dataframe from master file (not used, kept for compatibility)

        Returns:
            float: Minimum rate threshold value from annual_interest_rate column
        """
        try:
            # Get the file path for Master_Data.xlsx (located in the root Script folder)
            script_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            master_data_path = os.path.join(script_folder, "Master_Data.xlsx")

            if not os.path.exists(master_data_path):
                logger.warning(f"Master_Data.xlsx not found at {master_data_path}")
                raise FileNotFoundError(f"Master_Data.xlsx not found at {master_data_path}")

            logger.info(f"Reading threshold from: {master_data_path}")

            # Read the NBD-MF-23-IA-C1&C2-C5-IP sheet
            df_master_sheet = pd.read_excel(master_data_path, sheet_name="NBD-MF-23-IA-C1&C2-C5-IP")

            # Get the annual_interest_rate column
            if "annual_interest_rate" not in df_master_sheet.columns:
                logger.warning("Column 'annual_interest_rate' not found in Master_Data.xlsx sheet 'NBD-MF-23-IA-C1&C2-C5-IP'")
                raise ValueError("annual_interest_rate column not found")

            # Get the last non-null value from annual_interest_rate column
            annual_interest_rate_values = df_master_sheet["annual_interest_rate"].dropna()

            if annual_interest_rate_values.empty:
                logger.warning("No values found in annual_interest_rate column")
                raise ValueError("No values in annual_interest_rate column")

            # Get the LAST value (most recent/bottom row)
            minimum_rate_threshold = annual_interest_rate_values.iloc[-1]

            # Convert to numeric and validate
            minimum_rate_threshold = pd.to_numeric(minimum_rate_threshold, errors='coerce')

            if pd.isna(minimum_rate_threshold):
                logger.warning("Could not extract minimum rate threshold from annual_interest_rate column")
                # Check if command line override is available
                if hasattr(self, 'minimum_rate_override') and self.minimum_rate_override is not None:
                    logger.info(f"Using command line minimum rate override: {self.minimum_rate_override}%")
                    return self.minimum_rate_override/100
                else:
                    logger.warning("No command line override provided, using default value of 0.06 (6.00%)")
                    return 0.06

            logger.info(f"✓ Extracted minimum rate threshold from Master_Data.xlsx annual_interest_rate (last row): {minimum_rate_threshold} (will compare as {minimum_rate_threshold}%)")
            return minimum_rate_threshold

        except Exception as e:
            logger.error(f"Error extracting minimum rate threshold from Master_Data.xlsx: {e}")
            logger.error(traceback.format_exc())
            # Check if command line override is available
            if hasattr(self, 'minimum_rate_override') and self.minimum_rate_override is not None:
                logger.info(f"Using command line minimum rate override: {self.minimum_rate_override}%")
                return self.minimum_rate_override
            else:
                logger.warning("No command line override provided, using default minimum rate threshold of 0.06 (6.00%)")
                return 0.06

    def load_scienter_valuation(self):
        """Load Scienter valuation bot with error handling."""
        try:
            logger.info("Loading IA_Working_Initial_valuation_Scienter_bot module...")
            import sys
            import os
            
            # Add the parent "bots" folder relative to this file
            current_dir = os.path.dirname(__file__)
            bot_dir = os.path.join(current_dir, "..", "bots")
            sys.path.insert(0, os.path.abspath(bot_dir))
            
            import IA_Working_Initial_valuation_Scienter_bot as scienter_bot
            logger.info("Successfully imported IA_Working_Initial_valuation_Scienter_bot module")
            return scienter_bot
        except ImportError as e:
            logger.warning(f"Could not import IA_Working_Initial_valuation_Scienter_bot: {e}")
            return None

    def fetch_lr_valuations(self, contracts_list, scienter_bot):
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
            logger.warning("Scienter bot not available, cannot fetch LR valuations")
            return valuations
        
        if not contracts_list:
            logger.info("No LR contracts to process")
            return valuations
        
        try:
            logger.info(f"Fetching valuations for {len(contracts_list)} LR contracts via Scienter bot")
            
            # Call the Scienter bot's valuation function
            # Adjust the function name based on the actual bot's API
            response = scienter_bot.run_valuation_bot(contracts_list)
            
            if isinstance(response, dict):
                for contract, value in response.items():
                    # Validate the value from bot
                    if value and value not in ["", "Not Valued", "#N/A", "N/A"]:
                        try:
                            numeric_value = pd.to_numeric(value, errors='coerce')
                            if pd.notna(numeric_value) and numeric_value > 0:
                                valuations[contract] = numeric_value
                                logger.debug(f"✓ Scienter bot: {contract} = {numeric_value}")
                            else:
                                logger.debug(f"Scienter bot returned invalid value for {contract}: {value}")
                        except:
                            logger.debug(f"Scienter bot returned non-numeric value for {contract}: {value}")
                    else:
                        logger.debug(f"Scienter bot returned empty/N/A for {contract}")
            else:
                logger.warning(f"Scienter bot returned unexpected response type: {type(response)}")
                
            logger.info(f"✓ Scienter bot returned {len(valuations)} valid valuations")
            
        except Exception as e:
            logger.error(f"Error running Scienter bot for LR contracts: {e}")
            logger.error(traceback.format_exc())
        
        return valuations

    def fill_incomplete_contracts(self, df_prod_loans, na_contract_bot):
        """Fill incomplete contracts with comprehensive error handling."""
        if na_contract_bot is None:
            logger.warning("Contract bot not available, skipping incomplete contract processing")
            return df_prod_loans
        
        try:
            critical_cols = [
                "Client Code", "Equipment", "Frequency",
                "Contract Period", "Contractual Interest Rate", "Contract Amount"
            ]
            
            # Find incomplete contracts
            mask = df_prod_loans[critical_cols].isnull().any(axis=1) | (df_prod_loans[critical_cols] == "").any(axis=1)
            incomplete_contracts = df_prod_loans.loc[mask, "Contract No"].dropna().unique()
            
            # Keep all incomplete contracts (removed LR/Margin Trading exclusion)
            incomplete_contracts = list(incomplete_contracts)
            
            if len(incomplete_contracts) == 0:
                logger.info("No incomplete contracts found")
                return df_prod_loans
            
            logger.info(f"Found {len(incomplete_contracts)} incomplete contracts to process")
            
            # Process contracts with bot
            try:
                session, verification_token = na_contract_bot.get_authenticated_session()
                updated_count = 0
                failed_contracts = []
                
                for i, cn in enumerate(incomplete_contracts, 1):
                    cn_str = str(cn).strip()
                    if not cn_str:
                        continue
                    
                    try:
                        logger.debug(f"Processing contract {i}/{len(incomplete_contracts)}: {cn_str}")
                        data, success = na_contract_bot.process_contract_with_retry(session, cn_str, verification_token)
                        
                        if success and isinstance(data, dict):
                            mapping = {
                                "Client Code": data.get("client_code", ""),
                                "Equipment": data.get("equipment", ""),
                                "Frequency": data.get("frequency", ""),
                                "Contract Period": data.get("contract_period", ""),
                                "Contractual Interest Rate": data.get("interest_rate", ""),
                                "Contract Amount": data.get("contract_amount", "")
                            }
                            
                            for k, v in mapping.items():
                                if v:  # Only update if we have a value
                                    df_prod_loans.loc[df_prod_loans["Contract No"] == cn_str, k] = v
                            
                            updated_count += 1
                            logger.debug(f"✓ Updated contract {cn_str}")
                        else:
                            failed_contracts.append(cn_str)
                            logger.warning(f"✗ Failed to get data for contract {cn_str}")
                            
                    except Exception as e:
                        logger.error(f"Error processing contract {cn_str}: {e}")
                        failed_contracts.append(cn_str)
                
                logger.info(f"Contract update summary: {updated_count} successful, {len(failed_contracts)} failed")
                if failed_contracts:
                    logger.warning(f"Failed contracts: {failed_contracts[:10]}...")  # Show first 10
                    
            except Exception as e:
                logger.error(f"Error in bot processing: {e}")
                logger.error(traceback.format_exc())
            
            return df_prod_loans
            
        except Exception as e:
            logger.error(f"Error in fill_incomplete_contracts: {e}")
            logger.error(traceback.format_exc())
            return df_prod_loans


def run_report(picked_date):
    """
    Function to be called by app.py or other modules.

    Args:
        picked_date: Can be a datetime.date object, string in "YYYY-MM-DD" format, or "MM/DD/YYYY" format

    Returns:
        Dictionary containing the processed data and file paths
    """
    try:
        logger.info("Starting NBD_MF23_IA_Report via run_report function")

        # Initialize report generator
        loader = NBD_MF23_IA_Report()

        # Load and process all data
        dfs = loader.load_all(picked_date)

        # Log results
        logger.info("✓ Processing completed successfully")
        logger.info(f"Final output location: {dfs['final_output_path']}")

        # Check for processing errors
        if dfs.get("processing_errors"):
            logger.warning("Some non-critical errors occurred during processing:")
            for name, error in dfs["processing_errors"]:
                logger.warning(f"  - {name}: {error}")

        return dfs

    except Exception as e:
        logger.error(f"FATAL ERROR in run_report: {e}")
        logger.error(traceback.format_exc())
        raise


if __name__ == "__main__":
    try:
        import argparse

        # Set up command line argument parsing
        parser = argparse.ArgumentParser(description='NBD MF23 IA Report Generator')

        # Support both standalone execution and app.py integration
        parser.add_argument('date', nargs='?', help='Processing date in MM/DD/YYYY or YYYY-MM-DD format')
        parser.add_argument('--base-dir', default=r"working\monthly",
                          help='Base directory for data files (default: working\\monthly)')
        parser.add_argument('--working-dir', help='Working directory (used by app.py)')
        parser.add_argument('--month', help='Report month (e.g., Jan) - used by app.py')
        parser.add_argument('--year', help='Report year (e.g., 2025) - used by app.py')
        parser.add_argument('--minimum-rate', type=float, metavar='RATE',
                          help='Minimum rate threshold (percentage) to use if master data is missing (e.g., 6.5)')
        parser.add_argument('--rate-updates', type=str,
                          help='JSON string of rate updates for specific contracts (used by app.py)')
        parser.add_argument('--ltv-updates', type=str,
                          help='JSON string of LTV updates for specific contracts (used by app.py)')
        parser.add_argument('--verbose', '-v', action='store_true',
                          help='Enable verbose logging')

        # Parse arguments
        args = parser.parse_args()

        if args.verbose:
            logging.getLogger().setLevel(logging.DEBUG)
            logger.info("Verbose logging enabled")

        logger.info("Starting NBD_MF23_IA_Report script")

        # Determine date: either from 'date' argument or from --month/--year
        picked_date = None
        if args.date:
            picked_date = args.date
            logger.info(f"Processing date: {args.date}")
        elif args.month and args.year:
            # Convert month name and year to date (first day of next month for app.py compatibility)
            month_num = datetime.strptime(args.month, "%b").month
            next_month = month_num + 1 if month_num < 12 else 1
            year = int(args.year) if month_num < 12 else int(args.year) + 1
            picked_date = f"{next_month:02d}/01/{year}"
            logger.info(f"Processing date derived from month/year: {picked_date} (for report month {args.month} {args.year})")
        else:
            raise ValueError("Either 'date' or '--month' and '--year' must be provided")

        # Parse rate_updates if provided
        rate_updates = None
        if args.rate_updates:
            try:
                rate_updates = json.loads(args.rate_updates)
                logger.info(f"Parsed {len(rate_updates)} rate updates from command line")
            except json.JSONDecodeError as e:
                logger.warning(f"Failed to parse rate_updates JSON: {e}")
                rate_updates = None

        # Parse ltv_updates if provided
        ltv_updates = None
        if args.ltv_updates:
            try:
                ltv_updates = json.loads(args.ltv_updates)
                logger.info(f"Parsed {len(ltv_updates)} LTV updates from command line")
            except json.JSONDecodeError as e:
                logger.warning(f"Failed to parse ltv_updates JSON: {e}")
                ltv_updates = None

        # Determine directories
        if args.working_dir:
            # app.py provides full working directory path
            logger.info(f"Using working directory from app.py: {args.working_dir}")
            loader = NBD_MF23_IA_Report(working_dir=args.working_dir, rate_updates=rate_updates, ltv_updates=ltv_updates)
        else:
            # Standalone execution uses base_dir
            logger.info(f"Using base directory: {args.base_dir}")
            loader = NBD_MF23_IA_Report(base_dir=args.base_dir, rate_updates=rate_updates, ltv_updates=ltv_updates)

        if args.minimum_rate:
            logger.info(f"Minimum rate override: {args.minimum_rate}%")
            loader.minimum_rate_override = args.minimum_rate

        # Load and process all data
        dfs = loader.load_all(picked_date)

        # Log results
        logger.info("✓ Processing completed successfully")
        logger.info(f"Final output location: {dfs['final_output_path']}")

        # Check for processing errors
        if dfs.get("processing_errors"):
            logger.warning("Some non-critical errors occurred during processing:")
            for name, error in dfs["processing_errors"]:
                logger.warning(f"  - {name}: {error}")

        # Check if exception reports were generated
        exception_file = os.path.join(loader.ia_folder, "Minimum_Rate_Exceptions.xlsx")
        if os.path.exists(exception_file):
            logger.info(f"Minimum Rate exception report generated: {exception_file}")

        ltv_exception_file = os.path.join(loader.ia_folder, "LTV_validation_exceptions.xlsx")
        if os.path.exists(ltv_exception_file):
            logger.info(f"LTV validation exception report generated: {ltv_exception_file}")

        logger.info("NBD MF23 IA Report generation completed successfully!")

    except SystemExit:
        # Allow argparse to handle help and error exits
        pass
    except Exception as e:
        logger.error(f"FATAL ERROR: {e}")
        logger.error(traceback.format_exc())
        sys.exit(1)