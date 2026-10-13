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

            df_filtered = df[df["Contract No"].notna()].copy()

            # Normalize contract numbers and MSME values
            df_filtered["Contract No"] = df_filtered["Contract No"].astype(str).str.strip()
            df_filtered["Micro/Small/Medium"] = df_filtered["Micro/Small/Medium"].astype(str).str.strip()

            logger.info(f"Successfully loaded Information Request data: {len(df_filtered)} rows")
            logger.info(f"Sample data - Contracts: {df_filtered['Contract No'].head(3).tolist()}")
            logger.info(f"Sample data - MSME: {df_filtered['Micro/Small/Medium'].head(3).tolist()}")

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
            
            # Product column - will be created as formula later
            # Placeholder for now
            df_prod_loans["Product"] = ""
            logger.info("✓ Product column placeholder created (will be formula in Excel)")
            
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
            
            # Corporate Clients - will be created as formula later
            df_prod_loans["Corporate Clients"] = ""
            logger.info("✓ Corporate Clients column placeholder created (will be formula in Excel)")
            
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

                # Minimum Rate (Final) - will be created as formula later
                df_prod_loans["Minimum Rate (Final)"] = ""

                # Tenure (Months) - will be created as formula later
                df_prod_loans["Contract Period"] = pd.to_numeric(df_prod_loans["Contract Period"], errors="coerce")
                df_prod_loans["Tenure (Months)"] = ""
                logger.info("✓ Rate and tenure column placeholders created (will be formulas in Excel)")
            except Exception as e:
                logger.error(f"Error in rate/tenure calculations: {e}")
            
            # MSME Classification
            try:
                logger.info(f"Building MSME Classification mapping from Information Request from Credit data...")
                logger.info(f"Information Request from Credit has {len(df_information_request_from_credit)} contracts")
                logger.info(f"Sample contracts from Information Request: {df_information_request_from_credit['Contract No'].head(5).tolist()}")
                logger.info(f"Sample Micro/Small/Medium values: {df_information_request_from_credit['Micro/Small/Medium'].head(5).tolist()}")

                # Create mapping dictionary (already normalized in load_information_request_from_credit)
                contract_to_msme = dict(
                    zip(
                        df_information_request_from_credit["Contract No"],
                        df_information_request_from_credit["Micro/Small/Medium"]
                    )
                )
                logger.info(f"Created MSME mapping for {len(contract_to_msme)} contracts")

                # Normalize contract numbers in prod_loans before mapping
                df_prod_loans_normalized_contracts = df_prod_loans["Contract No"].astype(str).str.strip()
                logger.info(f"Sample contracts from prod_loans: {df_prod_loans_normalized_contracts.head(5).tolist()}")

                # Map using normalized contract numbers
                df_prod_loans["MSME Classification"] = (
                    df_prod_loans_normalized_contracts
                    .map(contract_to_msme)
                    .fillna("0")
                    .replace(
                        to_replace=r"\s*/\s*",
                        value="/",
                        regex=True
                    )
                )

                # Ensure consistent formatting
                df_prod_loans["MSME Classification"] = df_prod_loans["MSME Classification"].str.strip()

                # Standardize MSME Classification values to proper case
                def normalize_msme(value):
                    """Normalize MSME Classification values to standard format"""
                    if pd.isna(value) or value == "0" or value == "":
                        return value

                    value_upper = str(value).strip().upper()

                    # Map variations to standard values
                    if value_upper == "MICRO":
                        return "Micro"
                    elif value_upper == "SMALL":
                        return "Small"
                    elif value_upper == "MEDIUM":
                        return "Medium"
                    elif "CORPORATE" in value_upper or "OTHER ENTERPRISES" in value_upper:
                        return "Corporate/Other Enterprises"
                    else:
                        return value  # Keep original if not recognized

                df_prod_loans["MSME Classification"] = df_prod_loans["MSME Classification"].apply(normalize_msme)
                logger.info("✓ MSME Classification values standardized to proper case")

                # Log MSME Classification results
                msme_counts = df_prod_loans["MSME Classification"].value_counts(dropna=False)
                logger.info(f"MSME Classification value counts:\n{msme_counts}")

                # Check how many contracts were matched vs not matched
                matched_count = (df_prod_loans["MSME Classification"] != "0").sum()
                unmatched_count = (df_prod_loans["MSME Classification"] == "0").sum()
                logger.info(f"Matched contracts: {matched_count}, Unmatched (set to '0'): {unmatched_count}")

                # Log sample of unmatched contracts for debugging
                if unmatched_count > 0:
                    unmatched_contracts = df_prod_loans.loc[df_prod_loans["MSME Classification"] == "0", "Contract No"].head(10).tolist()
                    logger.info(f"Sample unmatched contracts (first 10): {unmatched_contracts}")

                logger.info("✓ MSME Classification completed and normalized")
            except Exception as e:
                logger.error(f"Error in MSME Classification: {e}")
                logger.error(traceback.format_exc())
                df_prod_loans["MSME Classification"] = "0"
            
            # EIR (%) - will be created as formula later
            df_prod_loans["EIR (%)"] = ""
            logger.info("✓ EIR column placeholder created (will be formula in Excel)")


            # Type of Loans - will be created as formula later
            df_prod_loans["Type of Loans (1.3.1.0.0.0)"] = ""
            logger.info("✓ Type of Loans column placeholder created (will be formula in Excel)")

            # PD Category - will be created as formula later
            df_prod_loans["PD Category"] = ""
            logger.info("✓ PD Category column placeholder created (will be formula in Excel)")

            # Collateral/Security Type - will be created as formula later
            df_prod_loans["Collateral/Security Type"] = ""
            logger.info("✓ Collateral/Security Type column placeholder created (will be formula in Excel)")

            # Initial Valuation - create placeholder column
            # IMPORTANT: Initial Valuation processing will be done AFTER reading calculated values
            # because it depends on Collateral/Security Type which is a formula
            df_prod_loans["Initial Valuation"] = None
            logger.info("✓ Initial Valuation placeholder created (will be populated after formulas calculate)")

            # Store bot instances and dataframes for later use in save_report
            self.df_po_listing = df_po_listing
            self.df_portfolio_rec = df_portfolio_rec
            self.df_product_cat = df_product_cat
            self.initial_valuation_bot = initial_valuation_bot

            # Other calculations - will be created as formulas later
            try:
                # Annual Interest Cost - will be formula
                df_prod_loans["Annual Interest Cost"] = ""
                logger.info("✓ Annual Interest Cost column placeholder created (will be formula in Excel)")

                # Gross Exposure - will be formula
                df_prod_loans["Gross Exposure"] = ""
                logger.info("✓ Gross Exposure column placeholder created (will be formula in Excel)")

                # Normal/Concessionary - will be formula
                df_prod_loans["Normal/Concessionary"] = ""
                logger.info("✓ Normal/Concessionary column placeholder created (will be formula in Excel)")

            except Exception as e:
                logger.error(f"Error in final calculations: {e}")
                logger.error(traceback.format_exc())
            
            logger.info("="*80)
            logger.info(f"build_prod_loans COMPLETED - Generated {len(df_prod_loans)} rows")
            logger.info(f"FINAL COLUMNS: {df_prod_loans.columns.tolist()}")
            logger.info("="*80)

            # LTV % - will be created as formula later
            df_prod_loans["LTV %"] = ""
            logger.info("✓ LTV % column placeholder created (will be formula in Excel)")

            # WALTV % - will be created as formula later
            df_prod_loans["WALTV %"] = ""
            logger.info("✓ WALTV % column placeholder created (will be formula in Excel)")


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

            # ===================================================================
            # All validations, transformations, and filtering moved to save_report
            # to run on calculated values (df_clean) instead of formula strings
            # ===================================================================
            logger.info("✓ Skipping validations and transformations in build_prod_loans")
            logger.info("  (Will run on calculated formula values in save_report)")

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

            # ===================================================================
            # Write Excel formulas to calculated columns with correct row references
            # ===================================================================
            logger.info("Creating Excel formulas for calculated columns...")

            # Helper to get Excel column letter from dataframe column index
            def get_col_letter(df, col_name):
                """Get Excel column letter for a given dataframe column name"""
                try:
                    idx = list(df.columns).index(col_name) + 1  # 1-based for Excel
                    result = ""
                    while idx > 0:
                        idx -= 1
                        result = chr(65 + (idx % 26)) + result
                        idx //= 26
                    return result
                except ValueError:
                    logger.warning(f"Column '{col_name}' not found in dataframe")
                    return None

            # Get column letters for all required columns
            col_letters = {}
            for col in df_prod_loans.columns:
                letter = get_col_letter(df_prod_loans, col)
                if letter:
                    col_letters[col] = letter

            logger.info(f"Column letter mapping: {col_letters}")

            # Convert formula columns to object dtype to avoid FutureWarning
            formula_columns = [
                "Tenure (Months)",
                "Minimum Rate (Final)",
                "Annual Interest Cost",
                "EIR (%)",
                "Gross Exposure",
                "PD Category",
                "Collateral/Security Type",
                "LTV %",
                "WALTV %"
            ]
            for col in formula_columns:
                if col in df_prod_loans.columns:
                    df_prod_loans[col] = df_prod_loans[col].astype(object)
            logger.info("✓ Converted formula columns to object dtype")

            # Create formulas for each row (Excel rows start at 3 since row 1-2 are headers)
            for idx in range(len(df_prod_loans)):
                excel_row = idx + 3  # Excel row number (data starts at row 3)

                # Get column letters
                A = col_letters.get("Contract No", "A")
                B = col_letters.get("Product", "B")
                C = col_letters.get("Client Code", "C")
                D = col_letters.get("Equipment", "D")
                E = col_letters.get("Purpose", "E")
                F = col_letters.get("Corporate Clients", "F")
                G = col_letters.get("Type of Loans (1.3.1.0.0.0)", "G")
                H = col_letters.get("Normal/Concessionary", "H")
                I = col_letters.get("Frequency", "I")
                J = col_letters.get("Contract Period", "J")
                K = col_letters.get("Tenure (Months)", "K")
                L = col_letters.get("Minimum Rate (Final)", "L")
                M = col_letters.get("Contractual Interest Rate", "M")
                N = col_letters.get("Base Rate", "N")
                P = col_letters.get("Contract Amount", "P")
                Q = col_letters.get("Annual Interest Cost", "Q")
                R = col_letters.get("EIR (%)", "R")
                S = col_letters.get("Gross Exposure", "S")
                T = col_letters.get("PD Category", "T")
                U = col_letters.get("Collateral/Security Type", "U")
                V = col_letters.get("Initial Valuation", "V")
                W = col_letters.get("LTV %", "W")
                X = col_letters.get("WALTV %", "X")

                # Product: =MID(A{row},3,2)
                df_prod_loans.at[idx, "Product"] = f"=MID({A}{excel_row},3,2)"

                # Corporate Clients: =IF(LEFT(C{row},1)="2","Corporate Client","Non-Corporate")
                df_prod_loans.at[idx, "Corporate Clients"] = f'=IF(LEFT({C}{excel_row},1)="2","Corporate Client","Non-Corporate")'

                # Type of Loans: Complex XLOOKUP formula
                df_prod_loans.at[idx, "Type of Loans (1.3.1.0.0.0)"] = (
                    f"=XLOOKUP({A}{excel_row},Product_Cat!L:L,Product_Cat!M:M,"
                    f'IF(OR({B}{excel_row}="MT",{B}{excel_row}="Margin Trading"),"Margin Trading Loans",'
                    f'IF(OR({B}{excel_row}="FDL",{B}{excel_row}="FD Loan"),"Loans against Cash/Deposits",'
                    f"XLOOKUP(({B}{excel_row}&{D}{excel_row}&{E}{excel_row}&{F}{excel_row}),Product_Cat!F:F,Product_Cat!E:E))))"
                )

                # Normal/Concessionary
                df_prod_loans.at[idx, "Normal/Concessionary"] = f'=IF({L}{excel_row}>\'NBD-MF-23-IA\'!$L$20,"Normal",IF(ISNUMBER(MATCH({C}{excel_row},Staff_Loan!F:F,0)),"Concessionary","Normal"))'

                # Tenure (Months)
                df_prod_loans.at[idx, "Tenure (Months)"] = f'=IF({I}{excel_row}="D",{J}{excel_row}/30,IF({I}{excel_row}="Q",{J}{excel_row}*3,IF({I}{excel_row}="M",{J}{excel_row},IF({I}{excel_row}="W",{J}{excel_row}/4,"Error"))))'

                # Minimum Rate (Final): =MIN(M{row},N{row})
                df_prod_loans.at[idx, "Minimum Rate (Final)"] = f"=MIN({M}{excel_row},{N}{excel_row})"

                # Annual Interest Cost: =P{row}*R{row}
                df_prod_loans.at[idx, "Annual Interest Cost"] = f"={P}{excel_row}*{R}{excel_row}"

                # EIR (%): =(POWER((1+L{row}/(12/K{row})),(12/K{row}))-1)
                df_prod_loans.at[idx, "EIR (%)"] = f"=(POWER((1+{L}{excel_row}/(12/{K}{excel_row})),(12/{K}{excel_row}))-1)"

                # Gross Exposure
                df_prod_loans.at[idx, "Gross Exposure"] = f"=XLOOKUP({A}{excel_row},'C1 & C2 Working'!A:A,'C1 & C2 Working'!AB:AB,{P}{excel_row})"

                # PD Category
                df_prod_loans.at[idx, "PD Category"] = f"=XLOOKUP({A}{excel_row},'C1 & C2 Working'!A:A,'C1 & C2 Working'!AD:AD,XLOOKUP({D}{excel_row},Product_Cat!R:R,Product_Cat!S:S))"

                # Collateral/Security Type
                df_prod_loans.at[idx, "Collateral/Security Type"] = (
                    f"=IF(ISNUMBER(MATCH({A}{excel_row},'YardandProperty List'!D:D,0)),\"Immovable Properties\","
                    f"IF(ISNUMBER(MATCH({T}{excel_row},Product_Cat!$P$2:$P$9,0)),\"Vehicles and Machinery\","
                    f'IF({B}{excel_row}="Margin Trading","Shares and Debt Securities-Listed",'
                    f'IF({B}{excel_row}="FDL","Deposits (Cash-Backed)","Personal and Corporate Guarantees"))))'
                )

                # LTV %: =IF(U{row}="Vehicles and Machinery",P{row}/V{row},0)
                df_prod_loans.at[idx, "LTV %"] = f'=IF({U}{excel_row}="Vehicles and Machinery",{P}{excel_row}/{V}{excel_row},0)'

                # WALTV %: =IF(U{row}="Vehicles and Machinery",(S{row}/$V$1)*W{row},0)
                # Note: Using total Initial Valuation sum in a fixed cell (V1) for WALTV calculation
                df_prod_loans.at[idx, "WALTV %"] = f'=IF({U}{excel_row}="Vehicles and Machinery",({S}{excel_row}/$V$1)*{W}{excel_row},0)'

            logger.info(f"✓ Created formulas for {len(df_prod_loans)} rows")
            logger.info("Formulas will be calculated when written to Excel")

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
                        logger.info(f"Pasting {len(df_clean)} rows with formulas to IA Working sheet (final row will be {3 + len(df_clean) - 1})")
                        worksheet.range('A3').value = df_clean.values
                        logger.info(f"✓ Successfully pasted data with formulas to IA Working sheet")

                        # Save workbook to calculate formulas
                        workbook.save()
                        logger.info("✓ Saved workbook to calculate formulas")

                        # Read back calculated values (VALUES ONLY, not formulas)
                        logger.info("Reading back calculated values from IA Working sheet...")
                        final_row = 3 + len(df_clean) - 1
                        final_col_letter = chr(64 + len(df_clean.columns))  # A=65, so 64+n gives correct letter
                        if len(df_clean.columns) > 26:
                            # Handle columns beyond Z (AA, AB, etc.)
                            final_col_letter = ""
                            n = len(df_clean.columns)
                            while n > 0:
                                n -= 1
                                final_col_letter = chr(65 + (n % 26)) + final_col_letter
                                n //= 26

                        data_range = worksheet.range(f'A3:{final_col_letter}{final_row}')
                        calculated_values = data_range.value

                        # Convert to list of lists if single row
                        if not isinstance(calculated_values[0], (list, tuple)):
                            calculated_values = [calculated_values]

                        # Create new dataframe with calculated values
                        df_calculated = pd.DataFrame(calculated_values, columns=df_clean.columns)
                        logger.info(f"✓ Read back {len(df_calculated)} rows with calculated values")

                        # Log the entire calculated dataframe (values only, no formulas)
                        logger.info("="*80)
                        logger.info("CALCULATED DATAFRAME (VALUES ONLY - NO FORMULAS)")
                        logger.info("="*80)
                        logger.info(f"Shape: {df_calculated.shape} (rows={df_calculated.shape[0]}, columns={df_calculated.shape[1]})")
                        logger.info(f"\nColumns: {list(df_calculated.columns)}")
                        logger.info(f"\nDataframe Info:")
                        logger.info(f"{df_calculated.info()}")
                        logger.info(f"\nFirst 100 rows:")
                        logger.info(f"\n{df_calculated.head(100).to_string()}")
                        logger.info(f"\nLast 100 rows:")
                        logger.info(f"\n{df_calculated.tail(100).to_string()}")
                        logger.info(f"\nData types:")
                        logger.info(f"\n{df_calculated.dtypes.to_string()}")
                        logger.info("="*80)

                        # Log sample of calculated values for verification
                        logger.info("Sample calculated values (first 3 rows):")
                        for idx in range(min(3, len(df_calculated))):
                            row_data = df_calculated.iloc[idx]
                            logger.info(f"  Row {idx+1}: Product={row_data.get('Product', 'N/A')}, "
                                      f"Corporate Clients={row_data.get('Corporate Clients', 'N/A')}, "
                                      f"Type of Loans={row_data.get('Type of Loans (1.3.1.0.0.0)', 'N/A')}")

                        # Now continue with validation using df_calculated instead of df_clean
                        # Note: Validations will now work on calculated values from formulas
                        df_clean = df_calculated.copy()

                        # ========================================================================
                        # Process Initial Valuation (must be done AFTER Collateral/Security Type is calculated)
                        # ========================================================================
                        logger.info("Processing Initial Valuation for calculated dataframe...")
                        try:
                            # Ensure Initial Valuation column exists
                            if "Initial Valuation" not in df_clean.columns:
                                df_clean["Initial Valuation"] = None

                            # Identify Vehicles and Machinery contracts
                            mask = df_clean["Collateral/Security Type"] == "Vehicles and Machinery"
                            total_vehicles = mask.sum()
                            logger.info(f"Found {total_vehicles} Vehicles and Machinery contracts")

                            if total_vehicles > 0:
                                # Step 1: Apply PO Listing SELL_PRICE
                                if hasattr(self, 'df_po_listing') and not self.df_po_listing.empty:
                                    if "CONTRACT_NO" in self.df_po_listing.columns and "SELL_PRICE" in self.df_po_listing.columns:
                                        po_mapping = dict(
                                            zip(self.df_po_listing["CONTRACT_NO"].astype(str),
                                                pd.to_numeric(self.df_po_listing["SELL_PRICE"], errors='coerce'))
                                        )
                                        po_mapping = {k: v for k, v in po_mapping.items() if pd.notna(v) and v > 0}
                                        df_clean.loc[mask, "Initial Valuation"] = (
                                            df_clean.loc[mask, "Contract No"].astype(str).map(po_mapping)
                                        )
                                        logger.info(f"✓ Applied Initial Valuation from PO Listing: {df_clean.loc[mask, 'Initial Valuation'].notna().sum()} contracts")

                                # Step 2: Apply Portfolio Recovery SELLING_PRICE
                                if hasattr(self, 'df_portfolio_rec') and not self.df_portfolio_rec.empty:
                                    if "CONTRACT_NO" in self.df_portfolio_rec.columns and "SELLING_PRICE" in self.df_portfolio_rec.columns:
                                        rec_mapping = dict(
                                            zip(self.df_portfolio_rec["CONTRACT_NO"].astype(str),
                                                pd.to_numeric(self.df_portfolio_rec["SELLING_PRICE"], errors='coerce'))
                                        )
                                        rec_mapping = {k: v for k, v in rec_mapping.items() if pd.notna(v) and v > 0}

                                        # Compare and keep max value
                                        updated_count = 0
                                        for idx, row in df_clean.loc[mask].iterrows():
                                            contract = str(row["Contract No"])
                                            current_val = row["Initial Valuation"]
                                            rec_val = rec_mapping.get(contract)

                                            values_to_compare = []
                                            if pd.notna(current_val) and current_val > 0:
                                                values_to_compare.append(current_val)
                                            if pd.notna(rec_val) and rec_val > 0:
                                                values_to_compare.append(rec_val)

                                            if values_to_compare:
                                                max_val = max(values_to_compare)
                                                if pd.isna(current_val) or max_val > current_val:
                                                    df_clean.at[idx, "Initial Valuation"] = max_val
                                                    updated_count += 1

                                        logger.info(f"✓ Updated Initial Valuation with Portfolio Recovery ({updated_count} contracts)")

                                # Step 3: Run bots for missing valuations (Vehicles & Machinery only)
                                missing_mask = mask & (df_clean["Initial Valuation"].isna() | (df_clean["Initial Valuation"] <= 0))

                            if total_vehicles > 0:
                                # Non-LR contracts
                                non_lr_missing_mask = missing_mask & ~df_clean["Contract No"].astype(str).str.startswith("LR")
                                if non_lr_missing_mask.any() and hasattr(self, 'initial_valuation_bot') and self.initial_valuation_bot:
                                    contracts_to_fetch = df_clean.loc[non_lr_missing_mask, "Contract No"].astype(str).tolist()
                                    logger.info(f"Running valuation bot for {len(contracts_to_fetch)} non-LR contracts...")
                                    try:
                                        response = self.initial_valuation_bot.run_valuation_bot(contracts_to_fetch)
                                        if isinstance(response, dict):
                                            for contract, value in response.items():
                                                if value and value not in ["", "Not Valued", "#N/A", "N/A"]:
                                                    numeric_value = pd.to_numeric(value, errors='coerce')
                                                    if pd.notna(numeric_value) and numeric_value > 0:
                                                        df_clean.loc[df_clean["Contract No"] == contract, "Initial Valuation"] = numeric_value
                                            logger.info(f"✓ Applied valuations from valuation bot")
                                    except Exception as e:
                                        logger.error(f"Error running valuation bot: {e}")

                                # Log final stats for Vehicles & Machinery
                                valued_count = df_clean.loc[mask, "Initial Valuation"].notna().sum()
                                missing_count = total_vehicles - valued_count
                                logger.info("="*60)
                                logger.info(f"Initial Valuation Summary (Vehicles & Machinery):")
                                logger.info(f"  Total Vehicles & Machinery: {total_vehicles}")
                                logger.info(f"  Total Valued: {valued_count} ({valued_count/total_vehicles*100:.1f}%)")
                                logger.info(f"  Total Missing: {missing_count} ({missing_count/total_vehicles*100:.1f}%)")
                                logger.info("="*60)

                            # Step 4: Run Scienter bot for ALL LR contracts with missing Initial Valuation (regardless of collateral type)
                            logger.info("\n" + "="*80)
                            logger.info("STEP 4: RUNNING SCIENTER BOT FOR ALL LR CONTRACTS WITH MISSING INITIAL VALUATION")
                            logger.info("="*80)

                            # Find all LR contracts with missing Initial Valuation
                            all_lr_mask = df_clean["Contract No"].astype(str).str.startswith("LR")
                            lr_missing_mask = all_lr_mask & (df_clean["Initial Valuation"].isna() | (df_clean["Initial Valuation"] <= 0))

                            if lr_missing_mask.any():
                                lr_contracts = df_clean.loc[lr_missing_mask, "Contract No"].astype(str).tolist()
                                logger.info(f"Found {len(lr_contracts)} LR contracts with missing Initial Valuation")
                                logger.info(f"LR contracts to process: {lr_contracts[:10]}..." if len(lr_contracts) > 10 else f"LR contracts to process: {lr_contracts}")

                                scienter_bot = self.load_scienter_valuation()
                                if scienter_bot:
                                    try:
                                        logger.info("Calling Scienter bot to fetch Initial Valuations...")
                                        lr_valuations = self.fetch_lr_valuations(lr_contracts, scienter_bot)

                                        if lr_valuations:
                                            logger.info(f"Received {len(lr_valuations)} valuations from Scienter bot")
                                            for contract, value in lr_valuations.items():
                                                df_clean.loc[df_clean["Contract No"] == contract, "Initial Valuation"] = value
                                                logger.info(f"  ✓ {contract}: {value:,.0f}")
                                            logger.info(f"✓ Successfully applied {len(lr_valuations)} valuations from Scienter bot")
                                        else:
                                            logger.warning("Scienter bot returned no valuations")
                                    except Exception as e:
                                        logger.error(f"Error running Scienter bot for LR contracts: {e}")
                                        logger.error(traceback.format_exc())
                                else:
                                    logger.warning("Scienter bot module not available - cannot fetch LR valuations")
                            else:
                                logger.info("No LR contracts with missing Initial Valuation found")

                            logger.info("="*80 + "\n")

                        except Exception as e:
                            logger.error(f"Error in Initial Valuation processing: {e}")
                            logger.error(traceback.format_exc())

                        # Apply Initial Valuation updates from initial_valuation.xlsx file (after formula calculation)
                        logger.info("="*80)
                        logger.info("APPLYING INITIAL VALUATION UPDATES FROM initial_valuation.xlsx FILE")
                        logger.info("="*80)

                        try:
                            # Look for initial_valuation.xlsx in the working folder
                            from pathlib import Path
                            iv_file_path = Path(self.ia_folder) / "initial_valuation.xlsx"

                            if iv_file_path.exists():
                                logger.info(f"Found initial_valuation.xlsx at: {iv_file_path}")

                                # Read initial_valuation.xlsx file
                                df_iv = pd.read_excel(iv_file_path)
                                logger.info(f"Loaded {len(df_iv)} rows from initial_valuation.xlsx")
                                logger.info(f"Columns in initial_valuation.xlsx: {df_iv.columns.tolist()}")

                                # Ensure required columns exist
                                if "contract_no" not in df_iv.columns or "initial_valuation" not in df_iv.columns:
                                    logger.error("initial_valuation.xlsx must have 'contract_no' and 'initial_valuation' columns")
                                else:
                                    # Filter out rows with empty contract_no or initial_valuation
                                    df_iv_valid = df_iv[df_iv["contract_no"].notna() & df_iv["initial_valuation"].notna()].copy()
                                    logger.info(f"Found {len(df_iv_valid)} valid Initial Valuation updates")

                                    updated_count = 0
                                    skipped_count = 0

                                    for idx, row in df_iv_valid.iterrows():
                                        contract_no = str(row["contract_no"]).strip()
                                        iv_value = float(row["initial_valuation"])

                                        # Find matching contract(s) in calculated data
                                        mask = df_clean["Contract No"].astype(str).str.strip() == contract_no

                                        if mask.any():
                                            # Check if Initial Valuation already has a value
                                            current_iv = df_clean.loc[mask, "Initial Valuation"].values[0]

                                            if pd.notna(current_iv) and current_iv != '' and float(current_iv) > 0:
                                                # Skip if already has value
                                                logger.info(f"Skipped {contract_no}: already has Initial Valuation = {current_iv}")
                                                skipped_count += 1
                                            else:
                                                # Update in dataframe
                                                df_clean.loc[mask, "Initial Valuation"] = iv_value
                                                logger.info(f"Updated Initial Valuation for {contract_no}: {iv_value:,.0f}")
                                                updated_count += mask.sum()
                                        else:
                                            logger.warning(f"Contract {contract_no} not found in data - skipping Initial Valuation update")

                                    logger.info(f"✓ Applied Initial Valuation updates to {updated_count} contract(s)")
                                    logger.info(f"✓ Skipped {skipped_count} contract(s) that already had Initial Valuation")

                            else:
                                logger.info("initial_valuation.xlsx file not found - skipping Initial Valuation updates")

                        except Exception as e:
                            logger.error(f"Error applying Initial Valuation updates from initial_valuation.xlsx: {e}")
                            logger.error(traceback.format_exc())

                        # ===================================================================
                        # Run validations on calculated values (df_clean)
                        # IMPORTANT: Run validations BEFORE writing back to Excel
                        # ===================================================================

                        # === Minimum Rate Validation ===
                        logger.info("="*80)
                        logger.info("MINIMUM RATE VALIDATION (on calculated values)")
                        logger.info("="*80)

                        try:
                            # Get minimum rate threshold from Master_Data.xlsx (as decimal)
                            min_rate_threshold = self.get_minimum_rate_threshold(self.df_product_cat)
                            logger.info(f"Minimum rate threshold: {min_rate_threshold} ({min_rate_threshold*100}%)")

                            # Filter out FDL and Margin Trading rows (they don't need minimum rate validation)
                            # Product codes: FD = FDL, MT = Margin Trading
                            mask_non_fdl_margin = ~df_clean["Product"].astype(str).str.strip().str.upper().isin(["FD", "MT"])
                            df_for_min_rate = df_clean[mask_non_fdl_margin].copy()
                            logger.info(f"Checking minimum rate for {len(df_for_min_rate)} contracts (excluding FD/MT)")

                            if len(df_for_min_rate) > 0 and "Minimum Rate (Final)" in df_for_min_rate.columns:
                                # Convert Minimum Rate to numeric
                                df_for_min_rate["Minimum Rate (Final)"] = pd.to_numeric(df_for_min_rate["Minimum Rate (Final)"], errors='coerce')

                                # Log statistics
                                min_rate = df_for_min_rate["Minimum Rate (Final)"].min()
                                max_rate = df_for_min_rate["Minimum Rate (Final)"].max()
                                logger.info(f"Min rate in data: {min_rate} ({min_rate*100:.2f}%), Max rate: {max_rate} ({max_rate*100:.2f}%)")

                                # Find contracts below threshold
                                below_threshold = df_for_min_rate[df_for_min_rate["Minimum Rate (Final)"] < min_rate_threshold]

                                if not below_threshold.empty:
                                    logger.warning(f"Found {len(below_threshold)} contracts with Minimum Rate < {min_rate_threshold} ({min_rate_threshold*100}%)")

                                    # Prepare exception data
                                    exception_data = below_threshold[["Contract No", "Client Code", "Minimum Rate (Final)"]].copy()
                                    exception_data["Threshold"] = min_rate_threshold
                                    exception_data["Issue"] = "Minimum Rate below threshold"

                                    # Save to Excel
                                    exception_dir = os.path.dirname(abs_file_path)
                                    exception_filepath = os.path.join(exception_dir, "NBD-MF-23-IA_Minimum_Rate_exceptions.xlsx")
                                    exception_data.to_excel(exception_filepath, index=False)
                                    logger.warning(f"⚠ Minimum Rate exceptions saved to: {exception_filepath}")
                                else:
                                    logger.info(f"✓ All contracts have Minimum Rate >= {min_rate_threshold} ({min_rate_threshold*100}%)")
                            else:
                                logger.info("No contracts to validate for minimum rate")

                        except Exception as e:
                            logger.error(f"Minimum Rate validation failed: {e}")
                            logger.error(traceback.format_exc())

                        # ===================================================================
                        # Write non-formula columns back to Excel
                        # (After all transformations and validations are complete)
                        # ===================================================================

                        # Write ONLY non-formula columns back to Excel
                        # Columns with formulas should retain their formulas, only update data columns
                        logger.info("Writing non-formula columns back to IA Working sheet...")

                        # Log Initial Valuation values before writing to Excel (to verify Scienter bot updates are included)
                        if "Initial Valuation" in df_clean.columns:
                            iv_values = df_clean["Initial Valuation"]
                            iv_non_null = iv_values[iv_values.notna()]
                            logger.info(f"Initial Valuation column stats before write: {len(iv_non_null)} non-null values out of {len(iv_values)} total")
                            # Log LR contracts with Initial Valuation
                            lr_mask = df_clean["Contract No"].astype(str).str.startswith("LR")
                            lr_with_iv = df_clean[lr_mask & iv_values.notna()]
                            if not lr_with_iv.empty:
                                logger.info(f"LR contracts with Initial Valuation ({len(lr_with_iv)} contracts):")
                                for idx, row in lr_with_iv.head(20).iterrows():
                                    logger.info(f"  {row['Contract No']}: {row['Initial Valuation']}")

                        # Define columns that contain actual data (not formulas)
                        non_formula_columns = [
                            "Contract No",
                            "Client Code",
                            "Equipment",
                            "Purpose",
                            "Frequency",
                            "Contract Period",
                            "Contractual Interest Rate",
                            "Base Rate",
                            "Disbursed Amount (Net of DC)",
                            "Contract Amount",
                            "Initial Valuation",
                            "MSME Classification"
                        ]

                        # Filter to only include columns that exist in df_clean
                        available_non_formula_cols = [col for col in non_formula_columns if col in df_clean.columns]
                        logger.info(f"Writing {len(available_non_formula_cols)} non-formula columns: {available_non_formula_cols}")

                        # Get column indices for each non-formula column
                        col_map = {col: idx for idx, col in enumerate(df_clean.columns)}

                        # Write each non-formula column to its corresponding Excel column
                        for col_name in available_non_formula_cols:
                            if col_name in col_map:
                                col_idx = col_map[col_name]
                                # Convert column index to Excel column letter
                                excel_col_idx = col_idx + 1  # 1-based
                                excel_col_letter = ""
                                n = excel_col_idx
                                while n > 0:
                                    n -= 1
                                    excel_col_letter = chr(65 + (n % 26)) + excel_col_letter
                                    n //= 26

                                # Get column data as list
                                col_data = df_clean[col_name].values.tolist()
                                col_data_2d = [[val] for val in col_data]  # Convert to 2D list for xlwings

                                # Write to Excel starting at row 3
                                worksheet.range(f'{excel_col_letter}3').value = col_data_2d
                                logger.info(f"  Wrote {len(col_data)} values to column {excel_col_letter} ({col_name})")

                        workbook.save()
                        logger.info("✓ Saved non-formula columns to Excel (formulas preserved)")

                        # ===================================================================
                        # Pre-Process Product Column for FDL and Margin Trading Detection
                        # (Must be done BEFORE post-paste modifications)
                        # ===================================================================
                        logger.info("="*80)
                        logger.info("PRE-PROCESSING: Updating Product column for FDL and Margin Trading detection")
                        logger.info("="*80)

                        # Read Contract No (column A) and Product (column B)
                        contract_nos_pre = worksheet.range(f'A3:A{3 + len(df_clean) - 1}').value
                        product_vals_pre = worksheet.range(f'B3:B{3 + len(df_clean) - 1}').value

                        if not isinstance(contract_nos_pre, list):
                            contract_nos_pre = [contract_nos_pre]
                        if not isinstance(product_vals_pre, list):
                            product_vals_pre = [product_vals_pre]

                        fdl_updated = 0
                        mt_updated = 0

                        # Update Product column based on Contract No patterns
                        for i, (contract, product) in enumerate(zip(contract_nos_pre, product_vals_pre)):
                            row_num = i + 3
                            contract_str = str(contract).strip() if contract else ""
                            product_str = str(product).strip() if product else ""

                            # For LR contracts (FDL) - check both contract number pattern and Product = "00"
                            # LR contracts have format LR00000049 where MID(A,3,2) = "00"
                            is_lr_contract = (contract_str.upper().startswith("LR") and
                                            len(contract_str) >= 4 and
                                            contract_str[2:4] == "00")
                            is_product_00 = product_str == "00"

                            if is_lr_contract or is_product_00:
                                worksheet.range(f'B{row_num}').value = "FDL"
                                fdl_updated += 1
                                logger.info(f"  Row {row_num}: Contract {contract_str}, Product '{product_str}' → Product set to 'FDL'")

                            # For Margin Trading - check if Contract No column itself says "Margin trading"
                            elif contract_str.lower().strip() == "margin trading":
                                worksheet.range(f'B{row_num}').value = "Margin Trading"
                                mt_updated += 1
                                logger.info(f"  Row {row_num}: Contract 'Margin trading', Product '{product_str}' → Product set to 'Margin Trading'")

                        logger.info(f"✓ Updated {fdl_updated} FDL rows and {mt_updated} Margin Trading rows")
                        workbook.save()
                        logger.info("✓ Saved Product column updates")

                        # ===================================================================
                        # Notes on transformations that will happen after final paste
                        # ===================================================================

                        # FDL and Margin Trading transformations
                        logger.info("="*80)
                        logger.info("FDL/Margin Trading transformations will be handled in Excel after final paste")
                        logger.info("="*80)

                        # LTV Validation note
                        logger.info("="*80)
                        logger.info("Note: LTV % validation will be performed in Excel after formula calculation")
                        logger.info("="*80)

                        # ===================================================================
                        # Post-Paste Excel Modifications
                        # ===================================================================

                        logger.info("="*80)
                        logger.info("APPLYING POST-PASTE EXCEL MODIFICATIONS")
                        logger.info("="*80)

                        # Get last data row
                        last_row = 3 + len(df_clean) - 1
                        logger.info(f"Data range: Row 3 to {last_row}")

                        # --- Modification 1: FDL and Margin Trading Transformations ---
                        logger.info("Processing FDL and Margin Trading rows...")

                        # Read column B (Product) to find FDL and Margin Trading rows
                        product_col = worksheet.range(f'B3:B{last_row}').value
                        if not isinstance(product_col, list):
                            product_col = [product_col]

                        fdl_rows = []
                        mt_rows = []
                        for i, val in enumerate(product_col):
                            row_num = i + 3
                            val_str = str(val).strip().upper() if val else ""
                            if val_str == "FDL":
                                fdl_rows.append(row_num)
                            elif val_str == "MARGIN TRADING":
                                mt_rows.append(row_num)

                        logger.info(f"Found {len(fdl_rows)} FDL rows and {len(mt_rows)} Margin Trading rows")

                        # Helper function to check if cell is empty, zero, or #N/A
                        def is_empty_zero_na(value):
                            if value is None or value == "" or value == 0:
                                return True
                            if isinstance(value, str):
                                val_upper = value.strip().upper()
                                if val_upper in ["", "0", "#N/A", "N/A", "NONE"]:
                                    return True
                            return False

                        # Process FDL rows
                        for row in fdl_rows:
                            logger.info(f"Processing FDL row {row}...")

                            # Clear columns C,D,E (Client Code, Equipment, Purpose)
                            worksheet.range(f'C{row}:E{row}').clear_contents()

                            # Clear columns I,J,K (Frequency, Contract Period, Tenure)
                            worksheet.range(f'I{row}:K{row}').clear_contents()

                            # Get current values for L and M
                            l_value = worksheet.range(f'L{row}').value
                            m_value = worksheet.range(f'M{row}').value

                            # Only set L and R if both L and M are empty/zero/#N/A
                            if is_empty_zero_na(l_value) and is_empty_zero_na(m_value):
                                base_rate = worksheet.range(f'N{row}').value
                                if base_rate:
                                    worksheet.range(f'L{row}').value = base_rate  # Minimum Rate
                                    worksheet.range(f'R{row}').value = base_rate  # EIR
                                    logger.info(f"  Set L{row} and R{row} = {base_rate}")
                                else:
                                    logger.info(f"  N{row} (Base Rate) is empty, skipped L and R")
                            else:
                                logger.info(f"  L{row} or M{row} already has value, skipped L and R update")

                            # Set P = O (Contract Amount = Disbursed Amount)
                            disbursed = worksheet.range(f'O{row}').value
                            if disbursed:
                                worksheet.range(f'P{row}').value = disbursed  # Contract Amount
                                logger.info(f"  Set P{row} = {disbursed}")

                            # Set Q = formula (Annual Interest Cost = P * R)
                            # Note: Q already has formula from initial paste, keep it as is
                            logger.info(f"  Annual Interest Cost (Q{row}) kept as formula =P{row}*R{row}")

                            # Calculate W = P / V (only if both have values)
                            try:
                                p_value = worksheet.range(f'P{row}').value
                                v_value = worksheet.range(f'V{row}').value
                                if p_value and v_value and float(v_value) != 0:
                                    w_calculated = float(p_value) / float(v_value)
                                    worksheet.range(f'W{row}').value = w_calculated
                                    logger.info(f"  Calculated W{row} = {p_value}/{v_value} = {w_calculated}")
                                else:
                                    logger.info(f"  Cannot calculate W{row}: P={p_value}, V={v_value}")
                            except Exception as e:
                                logger.warning(f"  Error calculating W for row {row}: {e}")

                            # Calculate X = S / (S + S_next) * W (only if all values available)
                            try:
                                s_value = worksheet.range(f'S{row}').value
                                s_next_value = worksheet.range(f'S{row+1}').value if row < last_row else 0
                                w_value = worksheet.range(f'W{row}').value

                                if s_value and w_value:
                                    sum_s = float(s_value) + (float(s_next_value) if s_next_value else 0)
                                    if sum_s != 0:
                                        x_calculated = (float(s_value) / sum_s) * float(w_value)
                                        worksheet.range(f'X{row}').value = x_calculated
                                        logger.info(f"  Calculated X{row} = ({s_value}/{sum_s}) * {w_value} = {x_calculated}")
                                    else:
                                        logger.info(f"  Cannot calculate X{row}: sum_s is zero")
                                else:
                                    logger.info(f"  Cannot calculate X{row}: S={s_value}, W={w_value}")
                            except Exception as e:
                                logger.warning(f"  Error calculating X for row {row}: {e}")

                        logger.info(f"✓ Processed {len(fdl_rows)} FDL rows")

                        # Process Margin Trading rows
                        for row in mt_rows:
                            logger.info(f"Processing Margin Trading row {row}...")

                            # Clear columns C,D,E (Client Code, Equipment, Purpose)
                            worksheet.range(f'C{row}:E{row}').clear_contents()

                            # Do NOT clear column F (Corporate Clients) - set to "Non-Corporate" if empty
                            corp_client_val = worksheet.range(f'F{row}').value
                            if not corp_client_val or str(corp_client_val).strip() == "":
                                worksheet.range(f'F{row}').value = "Non-Corporate"
                                logger.info(f"  Set F{row} = Non-Corporate (was empty)")
                            else:
                                logger.info(f"  F{row} (Corporate Clients) kept existing value: {corp_client_val}")

                            # Clear columns I,J,K (Frequency, Contract Period, Tenure)
                            worksheet.range(f'I{row}:K{row}').clear_contents()

                            # Clear column V (Initial Valuation)
                            worksheet.range(f'V{row}').clear_contents()

                            # Get current values for L and M
                            l_value = worksheet.range(f'L{row}').value
                            m_value = worksheet.range(f'M{row}').value

                            # Only set L and R if both L and M are empty/zero/#N/A
                            if is_empty_zero_na(l_value) and is_empty_zero_na(m_value):
                                base_rate = worksheet.range(f'N{row}').value
                                if base_rate:
                                    worksheet.range(f'L{row}').value = base_rate  # Minimum Rate
                                    worksheet.range(f'R{row}').value = base_rate  # EIR
                                    logger.info(f"  Set L{row} and R{row} = {base_rate}")
                                else:
                                    logger.info(f"  N{row} (Base Rate) is empty, skipped L and R")
                            else:
                                logger.info(f"  L{row} or M{row} already has value, skipped L and R update")

                            # Set P = O (Contract Amount = Disbursed Amount)
                            disbursed = worksheet.range(f'O{row}').value
                            if disbursed:
                                worksheet.range(f'P{row}').value = disbursed
                                logger.info(f"  Set P{row} (Contract Amount) = {disbursed}")

                            # Set MSME Classification (column Y) = "Medium"
                            worksheet.range(f'Y{row}').value = "Medium"
                            logger.info(f"  Set Y{row} (MSME Classification) = Medium")

                        logger.info(f"✓ Processed {len(mt_rows)} Margin Trading rows")

                        # --- Modification 2: Move FDL/MT to bottom with blank separator ---
                        if fdl_rows or mt_rows:
                            logger.info("Moving FDL/Margin Trading rows to bottom...")

                            # Collect all special rows
                            special_rows = sorted(fdl_rows + mt_rows)

                            # Insert blank row at the end
                            worksheet.range(f'{last_row + 1}:{last_row + 1}').insert(shift='down')

                            # Copy special rows to bottom (in reverse to maintain order)
                            for row in reversed(special_rows):
                                # Copy entire row
                                source_range = worksheet.range(f'{row}:{row}')
                                target_row = last_row + 2
                                source_range.copy(worksheet.range(f'{target_row}:{target_row}'))
                                last_row += 1

                            # Delete original rows (in reverse to avoid index shift)
                            for row in reversed(special_rows):
                                worksheet.range(f'{row}:{row}').delete()

                            logger.info(f"✓ Moved {len(special_rows)} special rows to bottom")

                        # --- Modification 3: Insert blank row before first non-LR contract ---
                        logger.info("Inserting separator between LR and main contracts...")
                        contract_nos = worksheet.range(f'A3:A{last_row}').value
                        if not isinstance(contract_nos, list):
                            contract_nos = [contract_nos]

                        first_non_lr_row = None
                        for i, contract in enumerate(contract_nos):
                            contract_str = str(contract).strip() if contract else ""
                            if contract_str and not contract_str.upper().startswith("LR"):
                                first_non_lr_row = i + 3
                                break

                        if first_non_lr_row and first_non_lr_row > 3:
                            worksheet.range(f'{first_non_lr_row}:{first_non_lr_row}').insert(shift='down')
                            logger.info(f"✓ Inserted blank row at {first_non_lr_row}")
                        else:
                            logger.info("No LR separator needed (no LR contracts or all are LR)")

                        # --- Modification 4: Force formula recalculation ---
                        logger.info("Forcing formula recalculation...")
                        workbook.app.calculate()
                        logger.info("✓ Formulas recalculated")

                        # ===================================================================
                        # Calculate WALTV % for FDL contracts (after formulas are calculated)
                        # ===================================================================

                        # Need to identify FDL rows again after moving them to bottom
                        logger.info("="*80)
                        logger.info("CALCULATING WALTV % FOR FDL CONTRACTS")
                        logger.info("="*80)

                        # Re-read Product column to find FDL rows at their new positions
                        last_row_current = worksheet.used_range.last_cell.row
                        product_col_updated = worksheet.range(f'B3:B{last_row_current}').value
                        if not isinstance(product_col_updated, list):
                            product_col_updated = [product_col_updated]

                        fdl_rows_updated = []
                        for i, val in enumerate(product_col_updated):
                            row_num = i + 3
                            val_str = str(val).strip().upper() if val else ""
                            if val_str == "FDL":
                                fdl_rows_updated.append(row_num)

                        logger.info(f"Found {len(fdl_rows_updated)} FDL rows after moving to bottom")

                        if fdl_rows_updated:

                            # Force Excel to calculate formulas before reading
                            worksheet.api.Calculate()

                            fdl_gross_exposures = []
                            fdl_ltv_percentages = []

                            for row in fdl_rows_updated:
                                # Read computed values from Excel
                                s_value = worksheet.api.Range(f"S{row}").Value   # Gross Exposure (formula result)
                                w_value = worksheet.api.Range(f"W{row}").Value   # LTV %

                                gross_exposure = float(s_value) if s_value not in [None, ""] else 0
                                ltv_pct = float(w_value) if w_value not in [None, ""] else 0

                                fdl_gross_exposures.append(gross_exposure)
                                fdl_ltv_percentages.append(ltv_pct)

                                logger.info(f"  Row {row}: Gross Exposure = {gross_exposure}, LTV % = {ltv_pct}")

                            sum_gross_exposure = sum(fdl_gross_exposures)
                            logger.info(f"Total Gross Exposure for FDL contracts: {sum_gross_exposure}")

                            if sum_gross_exposure > 0:
                                for i, row in enumerate(fdl_rows_updated):
                                    gross_exposure = fdl_gross_exposures[i]
                                    ltv_pct = fdl_ltv_percentages[i]

                                    waltv_pct = (gross_exposure / sum_gross_exposure) * ltv_pct
                                    worksheet.range(f"X{row}").value = waltv_pct

                                    logger.info(
                                        f"  Row {row}: WALTV % = ({gross_exposure}/{sum_gross_exposure}) * {ltv_pct} = {waltv_pct}"
                                    )

                                workbook.save()
                                logger.info("✓ Saved WALTV % calculations")
                            else:
                                logger.warning("Sum of Gross Exposure for FDL contracts is 0, cannot calculate WALTV %")

                        else:
                            logger.info("No FDL rows found, skipping WALTV % calculation")
                        # ===================================================================
                        # Validations After Recalculation (reading calculated values from Excel)
                        # ===================================================================

                        # --- Validation 1: LTV % Validation ---
                        logger.info("="*80)
                        logger.info("LTV % VALIDATION (after formula calculation)")
                        logger.info("="*80)

                        try:
                            # Read LTV % column (W) after formulas are calculated
                            last_row_val = 3 + len(df_clean) - 1
                            ltv_values = worksheet.range(f'W3:W{last_row_val}').value
                            collateral_types = worksheet.range(f'U3:U{last_row_val}').value
                            contract_nos = worksheet.range(f'A3:A{last_row_val}').value

                            if not isinstance(ltv_values, list):
                                ltv_values = [ltv_values]
                            if not isinstance(collateral_types, list):
                                collateral_types = [collateral_types]
                            if not isinstance(contract_nos, list):
                                contract_nos = [contract_nos]

                            # Check for LTV % values that are too high (e.g., > 100% = 1.0)
                            ltv_threshold = 1.0  # 100%
                            high_ltv_rows = []

                            for i, (ltv, coll_type, contract) in enumerate(zip(ltv_values, collateral_types, contract_nos)):
                                row_num = i + 3
                                if coll_type == "Vehicles and Machinery" and ltv is not None:
                                    try:
                                        ltv_float = float(ltv)
                                        if ltv_float > ltv_threshold:
                                            high_ltv_rows.append({
                                                'row': row_num,
                                                'contract': contract,
                                                'ltv': ltv_float,
                                                'collateral_type': coll_type
                                            })
                                    except (ValueError, TypeError):
                                        pass

                            if high_ltv_rows:
                                logger.warning(f"Found {len(high_ltv_rows)} contracts with LTV % > 100%")

                                # Save to Excel
                                ltv_exception_data = pd.DataFrame(high_ltv_rows)
                                ltv_exception_data.columns = ['Row', 'Contract No', 'LTV %', 'Collateral Type']
                                ltv_exception_filepath = os.path.join(os.path.dirname(abs_file_path), "NBD-MF-23-IA_LTV_exceptions.xlsx")
                                ltv_exception_data.to_excel(ltv_exception_filepath, index=False)
                                logger.warning(f"⚠ LTV % exceptions saved to: {ltv_exception_filepath}")
                            else:
                                logger.info("✓ All Vehicles & Machinery contracts have LTV % <= 100%")

                        except Exception as ltv_error:
                            logger.error(f"LTV % validation failed: {ltv_error}")
                            logger.error(traceback.format_exc())

                        # --- Validation 2: FDL/Margin Trading Final Check ---
                        logger.info("="*80)
                        logger.info("FDL/MARGIN TRADING FINAL VALIDATION (after transformations)")
                        logger.info("="*80)

                        try:
                            # Verify FDL and Margin Trading rows have correct values
                            product_col_final = worksheet.range(f'B3:B{last_row_val}').value
                            if not isinstance(product_col_final, list):
                                product_col_final = [product_col_final]

                            fdl_validation_issues = []
                            mt_validation_issues = []

                            for i, product in enumerate(product_col_final):
                                row_num = i + 3
                                product_str = str(product).strip() if product else ""

                                if product_str == "FDL":
                                    # Check if C,D,E,I,J,K are cleared
                                    c_val = worksheet.range(f'C{row_num}').value
                                    d_val = worksheet.range(f'D{row_num}').value
                                    e_val = worksheet.range(f'E{row_num}').value
                                    i_val = worksheet.range(f'I{row_num}').value
                                    j_val = worksheet.range(f'J{row_num}').value
                                    k_val = worksheet.range(f'K{row_num}').value

                                    issues = []
                                    if c_val: issues.append("C not cleared")
                                    if d_val: issues.append("D not cleared")
                                    if e_val: issues.append("E not cleared")
                                    if i_val: issues.append("I not cleared")
                                    if j_val: issues.append("J not cleared")
                                    if k_val: issues.append("K not cleared")

                                    if issues:
                                        fdl_validation_issues.append({'row': row_num, 'issues': ', '.join(issues)})

                                elif product_str == "Margin Trading":
                                    # Check if C,D,E,F,I,J,K,V are cleared
                                    c_val = worksheet.range(f'C{row_num}').value
                                    d_val = worksheet.range(f'D{row_num}').value
                                    e_val = worksheet.range(f'E{row_num}').value
                                    f_val = worksheet.range(f'F{row_num}').value
                                    i_val = worksheet.range(f'I{row_num}').value
                                    j_val = worksheet.range(f'J{row_num}').value
                                    k_val = worksheet.range(f'K{row_num}').value
                                    v_val = worksheet.range(f'V{row_num}').value

                                    issues = []
                                    if c_val: issues.append("C not cleared")
                                    if d_val: issues.append("D not cleared")
                                    if e_val: issues.append("E not cleared")
                                    if f_val: issues.append("F not cleared")
                                    if i_val: issues.append("I not cleared")
                                    if j_val: issues.append("J not cleared")
                                    if k_val: issues.append("K not cleared")
                                    if v_val: issues.append("V not cleared")

                                    if issues:
                                        mt_validation_issues.append({'row': row_num, 'issues': ', '.join(issues)})

                            if fdl_validation_issues or mt_validation_issues:
                                logger.warning(f"Found validation issues: {len(fdl_validation_issues)} FDL, {len(mt_validation_issues)} Margin Trading")

                                if fdl_validation_issues:
                                    for issue in fdl_validation_issues:
                                        logger.warning(f"  FDL row {issue['row']}: {issue['issues']}")

                                if mt_validation_issues:
                                    for issue in mt_validation_issues:
                                        logger.warning(f"  Margin Trading row {issue['row']}: {issue['issues']}")
                            else:
                                logger.info("✓ All FDL and Margin Trading rows validated successfully")

                        except Exception as fdl_mt_error:
                            logger.error(f"FDL/Margin Trading validation failed: {fdl_mt_error}")
                            logger.error(traceback.format_exc())

                        # Save after all modifications
                        workbook.save()
                        logger.info("✓ All post-paste modifications completed and saved")

                        logger.info("="*80)
                        logger.info("✓ Post-recalculation validations completed")
                        logger.info("="*80)

                        # ===================================================================
                        # Fix Gross Exposure Formula (AFTER validations)
                        # ===================================================================
                        logger.info("="*80)
                        logger.info("FIXING GROSS EXPOSURE FORMULA")
                        logger.info("="*80)

                        try:
                            # The Gross Exposure should lookup from C1 & C2 Working column K (Gross Portfolio)
                            # not column AB, since Summary Y (Gross Portfolio) goes to C1C2 K

                            # Get the last row
                            last_row_fix = worksheet.used_range.last_cell.row

                            # Update all Gross Exposure formulas (column S)
                            for row_num in range(3, last_row_fix + 1):
                                # Read current formula to check if it needs updating
                                current_formula = worksheet.range(f'S{row_num}').formula

                                if current_formula and 'XLOOKUP' in current_formula and 'AB:AB' in current_formula:
                                    # Replace AB:AB with K:K in the formula
                                    new_formula = current_formula.replace('AB:AB', 'K:K')
                                    worksheet.range(f'S{row_num}').formula = new_formula

                            logger.info(f"✓ Updated Gross Exposure formulas from column AB to column K (Gross Portfolio)")

                            # Recalculate and save
                            worksheet.api.Calculate()
                            workbook.save()
                            logger.info("✓ Recalculated and saved after Gross Exposure formula fix")

                        except Exception as e:
                            logger.error(f"Failed to fix Gross Exposure formula: {e}")
                            logger.error(traceback.format_exc())

                        # ===================================================================
                        # Product_Cat Sheet Processing (AFTER recalculation)
                        # ===================================================================
                        logger.info("="*80)
                        logger.info("PRODUCT_CAT SHEET PROCESSING (after recalculation)")
                        logger.info("="*80)

                        # Extract unique combinations from rows with #N/A Type of Loans and paste to Product_Cat sheet
                        # Use df_calculated (values-only dataframe after formulas are calculated)
                        logger.info("Extracting unique combinations from rows with #N/A or empty Type of Loans...")
                        logger.info("Using df_calculated (values-only dataframe after formula calculation)")
                        cols_for_product_cat = ["Product", "Equipment", "Purpose", "Corporate Clients"]

                        # Filter rows where 'Type of Loans (1.3.1.0.0.0)' is '#N/A', None, empty, or null (these need classification)
                        # Convert to string first to catch all variations of #N/A
                        type_loans_str = df_calculated["Type of Loans (1.3.1.0.0.0)"].astype(str).str.strip().str.upper()
                        mask_na_type = (
                            (type_loans_str == "#N/A") |
                            (type_loans_str == "N/A") |
                            (type_loans_str == "#N/A!") |
                            (type_loans_str == "") |
                            (type_loans_str == "NONE") |
                            (type_loans_str == "NAN") |
                            (df_calculated["Type of Loans (1.3.1.0.0.0)"].isna()) |
                            (df_calculated["Type of Loans (1.3.1.0.0.0)"].isnull())
                        )
                        na_count = mask_na_type.sum()
                        logger.info(f"Found {na_count} rows with #N/A, None, or empty Type of Loans in df_calculated")

                        # Log ALL unique values in Type of Loans column for complete debugging
                        logger.info(f"Total rows in df_calculated: {len(df_calculated)}")
                        unique_type_loans = df_calculated["Type of Loans (1.3.1.0.0.0)"].unique()
                        logger.info(f"All unique Type of Loans values ({len(unique_type_loans)}): {unique_type_loans}")

                        # Count how many of each value
                        value_counts = df_calculated["Type of Loans (1.3.1.0.0.0)"].value_counts(dropna=False)
                        logger.info(f"Type of Loans value counts:\n{value_counts}")

                        # Log sample of Type of Loans values for debugging
                        if na_count > 0:
                            sample_values = df_calculated.loc[mask_na_type, "Type of Loans (1.3.1.0.0.0)"].head(10).tolist()
                            logger.info(f"Sample Type of Loans values that need classification (first 10): {sample_values}")

                        # Extract unique combinations from those rows only
                        if na_count > 0:
                            # Step 1: Get rows with #N/A Type of Loans (4 columns only)
                            df_na_rows = df_calculated.loc[mask_na_type, cols_for_product_cat].copy()
                            logger.info(f"Step 1: Extracted {len(df_na_rows)} rows with #N/A Type of Loans")

                            # Step 2: Create LOOKUP column (5th column) by concatenating the 4 columns
                            df_na_rows["LOOKUP"] = (
                                df_na_rows["Product"].astype(str) +
                                df_na_rows["Equipment"].astype(str) +
                                df_na_rows["Purpose"].astype(str) +
                                df_na_rows["Corporate Clients"].astype(str)
                            )
                            logger.info(f"Step 2: Created LOOKUP column (5th column)")

                            # Step 3: Remove duplicate rows based on LOOKUP column
                            original_count = len(df_na_rows)
                            df_na_rows = df_na_rows.drop_duplicates(subset=['LOOKUP'], keep='first').reset_index(drop=True)
                            removed_count = original_count - len(df_na_rows)
                            logger.info(f"Step 3: Removed {removed_count} duplicate rows based on LOOKUP column")

                            # Step 4: Remove rows where Product is "rg", "Margin Trading", "00", or "FDL"
                            before_exclusion = len(df_na_rows)
                            df_na_rows = df_na_rows[~df_na_rows["Product"].isin(["rg", "Margin Trading", "00", "FDL"])].reset_index(drop=True)
                            excluded_count = before_exclusion - len(df_na_rows)
                            logger.info(f"Step 4: Excluded {excluded_count} rows with Product = rg/Margin Trading/00/FDL, remaining: {len(df_na_rows)} rows")

                            # Keep all 5 columns (Product, Equipment, Purpose, Corporate Clients, LOOKUP)
                            df_unique_combinations = df_na_rows
                            logger.info(f"Final unique combinations to append (with LOOKUP column): {len(df_unique_combinations)}")
                        else:
                            df_unique_combinations = pd.DataFrame(columns=cols_for_product_cat + ["LOOKUP"])
                            logger.info("No #N/A Type of Loans rows found - nothing to add to Product_Cat")

                        # Step 5: Process Product_Cat sheet - Append to the bottom
                        product_cat_sheet = workbook.sheets["Product_Cat"]

                        # Find the last row with data in Product_Cat sheet (to append below it)
                        used_range = product_cat_sheet.used_range
                        last_row_pc = used_range.last_cell.row
                        logger.info(f"Step 5: Product_Cat sheet currently has {last_row_pc} rows")

                        # Append new unique combinations below existing data (with LOOKUP column - 5 columns total)
                        if not df_unique_combinations.empty:
                            start_row = last_row_pc + 1
                            product_cat_sheet.range(f'A{start_row}').value = df_unique_combinations.values
                            num_rows = len(df_unique_combinations)
                            logger.info(f"Step 5: Appended {num_rows} rows to Product_Cat sheet (rows {start_row} to {start_row + num_rows - 1})")
                        else:
                            logger.info("Step 5: No new rows to append")

                        workbook.save()
                        logger.info("Step 5: Saved workbook after appending")

                        # Step 6: Read the entire Product_Cat sheet to a new dataframe
                        try:
                            final_last_row = product_cat_sheet.used_range.last_cell.row
                            logger.info(f"Step 6: Reading entire Product_Cat sheet ({final_last_row} rows)")

                            # Read columns A:E from Product_Cat sheet (5 columns: Product, Equipment, Purpose, Corporate Clients, LOOKUP)
                            product_cat_range = product_cat_sheet.range(f'A1:E{final_last_row}')
                            product_cat_data = product_cat_range.value

                            # Convert to list of lists if single row
                            if final_last_row == 2 and not isinstance(product_cat_data[0], (list, tuple)):
                                product_cat_data = [product_cat_data]

                            # Convert to DataFrame (with headers from row 1)
                            df_product_cat_full = pd.DataFrame(product_cat_data[1:], columns=product_cat_data[0])
                            logger.info(f"Step 6: Read {len(df_product_cat_full)} rows from Product_Cat sheet")

                            # Step 7: Prepare to write to Product_Cat.xlsx
                            logger.info("Step 7: Writing to Product_Cat.xlsx...")
                            base_dir = Path(prod_file_path).parent
                            product_cat_xlsx_path = base_dir / "Product_Cat.xlsx"

                            # Open or create Product_Cat.xlsx with xlwings
                            if product_cat_xlsx_path.exists():
                                logger.info(f"Opening existing Product_Cat.xlsx at {product_cat_xlsx_path}")
                                pc_wb = workbook.app.books.open(str(product_cat_xlsx_path))
                            else:
                                logger.info(f"Creating new Product_Cat.xlsx at {product_cat_xlsx_path}")
                                pc_wb = workbook.app.books.add()
                                pc_wb.save(str(product_cat_xlsx_path))

                            pc_sheet = pc_wb.sheets[0]  # First sheet

                            # Clear existing data from row 2 onwards
                            if pc_sheet.used_range.last_cell.row > 1:
                                last_row_xlsx = pc_sheet.used_range.last_cell.row
                                clear_range = f"A2:E{last_row_xlsx}"
                                pc_sheet.range(clear_range).clear_contents()
                                logger.info(f"Step 7: Cleared {clear_range} in Product_Cat.xlsx")

                            # Ensure headers are in row 1
                            pc_sheet.range("A1").value = ["Product", "Equipment", "Purpose", "Corporate Clients", "LOOKUP"]
                            logger.info("Step 7: Ensured headers in row 1")

                            # Paste data without headers (from row 2)
                            if len(df_product_cat_full) > 0:
                                pc_sheet.range("A2").value = df_product_cat_full.values
                                logger.info(f"Step 7: Pasted {len(df_product_cat_full)} rows (without headers) to Product_Cat.xlsx")

                            # Save and close Product_Cat.xlsx
                            pc_wb.save()
                            pc_wb.close()
                            logger.info(f"Step 7: Saved and closed Product_Cat.xlsx")
                            logger.info(f"✓ Product_Cat.xlsx will be downloaded with report")

                        except Exception as export_error:
                            logger.error(f"Failed to export Product_Cat to xlsx: {export_error}")
                            logger.error(traceback.format_exc())

                        logger.info("="*80)
                        logger.info("✓ Product_Cat sheet processing completed")
                        logger.info("="*80)

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

            # Convert percentage to decimal (42.0% → 0.42)
            minimum_rate_decimal = minimum_rate_threshold / 100
            logger.info(f"✓ Extracted minimum rate threshold from Master_Data.xlsx: {minimum_rate_threshold}% → {minimum_rate_decimal} (as decimal)")
            return minimum_rate_decimal

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
            response = scienter_bot.run_scienter_valuation_bot(contracts_list)
            
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