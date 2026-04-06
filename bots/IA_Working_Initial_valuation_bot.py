#!/usr/bin/env python3
"""
IA Working Initial Valuation Bot
Extracts valuation amounts from Equipment details API response for contract numbers
"""

import os
import time
import logging
import json
import base64
import requests
from datetime import datetime
from dotenv import load_dotenv
from cryptography.fernet import Fernet
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from requests.exceptions import RequestException, ConnectionError
import random

# Load environment variables from a deterministic path
try:
    # Absolute path to repo root (parent of bots directory)
    _root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _env_path = os.path.join(_root_dir, '.env')
    if os.path.exists(_env_path):
        load_dotenv(_env_path)
    else:
        # Fallback to CWD .env
        load_dotenv()
except Exception:
    # Best-effort load
    load_dotenv()

# ---------------- Config ----------------
BASE_URL = "https://erp.assetline.lk/"
RECOVERY_URL = "https://erp.assetline.lk/Application/Home/RECOVERY"
LOGS_DIR = "logs"

# Create logs directory
os.makedirs(LOGS_DIR, exist_ok=True)

def load_contract_numbers_from_excel():
    """Load contract numbers from the Excel file used by NBD_MF_23_IA.py"""
    try:
        import pandas as pd
        from pathlib import Path
        
        # Get the parent directory (Script directory)
        script_dir = Path(__file__).parent.parent
        
        # Path to the Excel file (same as used in NBD_MF_23_IA.py)
        excel_file = script_dir / "Prod. wise Class. of Loans - Jul 2025.xlsb"
        
        if not excel_file.exists():
            raise FileNotFoundError(f"Excel file not found: {excel_file}")
        
        # Read the Excel file
        logger.info(f"Reading contract numbers from: {excel_file}")
        
        # Try to read as .xlsb first, if that fails, try .xlsx
        try:
            df = pd.read_excel(excel_file, sheet_name="IA Working")
        except Exception as e:
            logger.warning(f"Failed to read .xlsb file: {e}")
            # Try alternative file names
            alternative_files = [
                script_dir / "Prod. wise Class. of Loans - Jul 2025.xlsx",
                script_dir / "Prod. wise Class. of Loans - Jul 2025.xls"
            ]
            
            for alt_file in alternative_files:
                if alt_file.exists():
                    try:
                        df = pd.read_excel(alt_file, sheet_name="IA Working")
                        logger.info(f"Successfully read from alternative file: {alt_file}")
                        break
                    except Exception as alt_e:
                        logger.warning(f"Failed to read {alt_file}: {alt_e}")
                        continue
            else:
                raise Exception("Could not read any Excel file")
        
        # Extract contract numbers from column A starting from row 3 (index 2)
        contract_numbers = []
        for idx, row in df.iterrows():
            if idx >= 2:  # Skip first two rows (0 and 1), start from row 3 (index 2)
                contract = row.iloc[0]  # Column A
                if contract and pd.notna(contract) and str(contract).strip():
                    contract_str = str(contract).strip()
                    if contract_str != "nan" and contract_str != "":
                        contract_numbers.append(contract_str)
        
        logger.info(f"Successfully loaded {len(contract_numbers)} contract numbers from Excel file")
        return contract_numbers
        
    except Exception as e:
        logger.error(f"Failed to load contract numbers from Excel: {e}")
        raise

# ---------------- Logging Setup ----------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, 'ia_working_valuation_bot.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# ---------------- HTTP Session with Retries ----------------
def build_session():
    """Build a resilient requests session with retries and backoff."""
    session = requests.Session()
    retry = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
        raise_on_status=False,
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=100, pool_maxsize=100)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"User-Agent": "InitialValuationBot/1.0"})
    return session

def fetch_valuation_with_retry(session, url, data_payload, headers, max_attempts=4, base_delay=1.0):
    """POST with retries and jitter; returns JSON dict or None."""
    for attempt in range(1, max_attempts + 1):
        try:
            resp = session.post(url, data=data_payload, headers=headers, timeout=(10, 60))
            if resp.ok:
                try:
                    return resp.json()
                except json.JSONDecodeError:
                    return None
            if resp.status_code in (429, 500, 502, 503, 504):
                raise RequestException(f"HTTP {resp.status_code}")
            return None
        except (ConnectionError, RequestException):
            if attempt < max_attempts:
                delay = base_delay * attempt + random.uniform(0, 0.5)
                try:
                    session.close()
                finally:
                    session = build_session()
                time.sleep(delay)
                continue
            return None

# ---------------- Credential Decryption ----------------
def load_credentials():
    """Load and decrypt credentials from environment variables.

    Prefers ENCRYPTION_KEY from environment to decrypt ENC_USERNAME/ENC_PASSWORD.
    Falls back to reading .key file in repo root if ENCRYPTION_KEY is not set.
    """
    try:
        # Get encrypted credentials
        enc_username = os.getenv('ENC_USERNAME')
        enc_password = os.getenv('ENC_PASSWORD')
        
        if not enc_username or not enc_password:
            raise ValueError("ENC_USERNAME and ENC_PASSWORD must be set in environment variables")
        
        logger.info("Found encrypted credentials, attempting to decrypt...")

        key_str = os.environ.get("FERNET_KEY", "").strip()
        if not key_str:
            raise RuntimeError("FERNET_KEY environment variable not set")

        fernet = Fernet(key_str.encode())
        
        # Decrypt credentials
        username = fernet.decrypt(enc_username.encode()).decode()
        password = fernet.decrypt(enc_password.encode()).decode()
        
        logger.info("Successfully decrypted credentials")
        logger.info("Credentials loaded successfully")
        return username, password
        
    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        raise

# ---------------- Selenium Functions ----------------
def start_driver():
    """Start Chrome driver with optimized settings"""
    try:
        options = Options()
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-plugins")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-images")  # Disable image loading
        options.add_argument("--blink-settings=imagesEnabled=false")  # Additional image blocking
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Check if Chrome is installed
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(os.getenv('USERNAME', '')),
            r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(os.getenv('USERPROFILE', '').split('\\')[-1])
        ]
        
        chrome_found = False
        for path in chrome_paths:
            if os.path.exists(path):
                options.binary_location = path
                chrome_found = True
                logger.info(f"Found Chrome at: {path}")
                break
        
        if not chrome_found:
            raise Exception("Chrome browser not found. Please install Google Chrome.")
        
        # Start driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logger.info("Chrome driver started successfully")
        return driver
        
    except Exception as e:
        logger.error(f"Failed to start Chrome driver: {e}")
        raise

def login_to_system(driver):
    """Login to the ERP system"""
    try:
        logger.info("Logging into ERP system...")
        driver.get(BASE_URL)
        time.sleep(3)
        
        # Load credentials
        username, password = load_credentials()
        
        # Find and fill username
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "strUserName"))
        )
        username_input.clear()
        username_input.send_keys(username)
        logger.info("Username entered")
        
        # Find and fill password
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "strPassword"))
        )
        password_input.clear()
        password_input.send_keys(password)
        logger.info("Password entered")
        
        # Click sign in button
        signin_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "idSignIn"))
        )
        signin_button.click()
        logger.info("Sign in button clicked")
        
        # Wait for page to load and click SELECT button
        time.sleep(3)
        try:
            select_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='idSignIn' and @type='submit']"))
            )
            select_button.click()
            logger.info("SELECT button clicked")
        except:
            logger.info("No additional SELECT button found, proceeding with login")
        
        time.sleep(3)
        logger.info("Login successful")
        return True
        
    except Exception as e:
        logger.error(f"Login failed: {e}")
        raise

def navigate_to_contract_inquiry(driver):
    """Navigate to Contract Inquiry page"""
    try:
        logger.info("Navigating to Contract Inquiry...")
        driver.get(RECOVERY_URL)
        time.sleep(3)

        try:
            # Primary path: via tiles
            info_center_tile = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//div[contains(@class, 'metro-tile-body') and contains(@class, 'tile-type-menu')]//i[contains(@class, 'fa-file')]/.."))
            )
            info_center_tile.click()
            logger.info("Information Center tile clicked")
            time.sleep(3)

            contract_inquiry_tile = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//div[contains(@class, 'metro-tile-body')]//i[contains(@class, 'fa-user')]/.."))
            )
            contract_inquiry_tile.click()
            logger.info("Contract Inquiry tile clicked")
            time.sleep(3)

            add_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "btn_02461_Add"))
            )
            add_button.click()
            logger.info("Add button clicked")
            time.sleep(3)
            logger.info("Successfully navigated to Contract Inquiry")
            return True
        except Exception as primary_err:
            logger.warning(f"Primary navigation path failed: {primary_err}. Trying direct URL fallback...")

            # Fallback: direct page URL, then wait for Add button
            try:
                driver.get("https://erp.assetline.lk/Recovery/RECContractInquiry")
                add_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, "btn_02461_Add"))
                )
                add_button.click()
                logger.info("Add button clicked via direct URL fallback")
                time.sleep(2)
                logger.info("Successfully navigated to Contract Inquiry via fallback")
                return True
            except Exception as fallback_err:
                logger.error(f"Fallback navigation also failed: {fallback_err}")
                raise
        
    except Exception as e:
        logger.error(f"Failed to navigate to Contract Inquiry: {e}")
        raise

def process_contract(driver, contract_no):
    """Process a single contract to extract valuation amount using direct API call"""
    try:
        logger.info(f"Processing contract: {contract_no}")
        
        # Step 1: Find the contract number input field and enter contract number
        contract_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "strContractNo_RECContractInquiry"))
        )
        
        # Clear the field and enter contract number
        contract_input.clear()
        contract_input.send_keys(contract_no)
        logger.info(f"Entered contract number: {contract_no}")
        
        # Step 2: Hit Enter to trigger the search
        contract_input.send_keys(Keys.RETURN)
        logger.info("Pressed Enter to trigger search")
        
        # Step 3: Wait for page to load
        time.sleep(3)
        logger.info("Waiting for page to load...")

        # Click Equipment Details button to ensure server triggers the inquiry
        try:
            equipment_details_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "btn_02461_EquipmentDetails_Button_RECContractInquiry"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", equipment_details_button)
            time.sleep(0.2)
            try:
                equipment_details_button.click()
            except Exception:
                driver.execute_script("arguments[0].click();", equipment_details_button)
            logger.info("Clicked Equipment Details button")
            time.sleep(0.5)
        except Exception as click_err:
            logger.warning(f"Equipment Details button click failed or not found: {click_err}")
        
        # Step 4: Make direct API call to get Equipment Details
        try:
            # Get verification token from the page
            verification_token = ""
            try:
                token_element = driver.find_element(By.NAME, "__RequestVerificationToken")
                verification_token = token_element.get_attribute("value")
                logger.info("Found verification token")
            except:
                logger.warning("Could not find verification token")
            
            # Get cookies from the current session
            cookies = driver.get_cookies()
            logger.info(f"Retrieved {len(cookies)} cookies")
            
            # Create a resilient requests session with the cookies
            session = build_session()
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])
            
            # Set headers matching the actual browser request
            headers = {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "Accept": "*/*",
                "Accept-Language": "en-US,en;q=0.9",
                "Origin": "https://erp.assetline.lk",
                "Referer": "https://erp.assetline.lk/Application/Home/RECOVERY",
                "X-Requested-With": "XMLHttpRequest",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0",
                "Priority": "u=1, i",
                "Sec-Ch-Ua": '"Not;A=Brand";v="99", "Microsoft Edge";v="139", "Chromium";v="139"',
                "Sec-Ch-Ua-Mobile": "?0",
                "Sec-Ch-Ua-Platform": '"Windows"',
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin"
            }
            
            # Prepare payload for Equipment Details API call
            payload = {
                "strContractNo_RECContractInquiry": contract_no,
                "strInquireType_RECContractInquiry": "MAININQ_EQUIPMENTDETAILS",
                "STR_FORM_ID": "02461",
                "STR_FUNCTION_ID": "CR",
                "STR_PREMIS": "HOD",
                "STR_INSTANT": "ALCL",
                "STR_APP_ID": "00023"
            }
            
            if verification_token:
                payload["__RequestVerificationToken"] = verification_token
            
            # Make the API request with retries
            api_url = "https://erp.assetline.lk/Recovery/RECContractInquiry/Inquire"
            equipment_api_response = fetch_valuation_with_retry(session, api_url, payload, headers)
            
            if equipment_api_response is not None:
                try:
                    logger.info("Successfully received Equipment Details API response")
                    
                    # Extract EQD_VALAMOUNT from the API response
                    valuation_amount = None
                    
                    if isinstance(equipment_api_response, dict):
                        logger.info(f"API Response structure: {list(equipment_api_response.keys())}")
                        
                        # Check if we have DATA field with dsResult
                        if "DATA" in equipment_api_response and "dsResult" in equipment_api_response["DATA"]:
                            ds_result = equipment_api_response["DATA"]["dsResult"]
                            logger.info(f"dsResult structure: {list(ds_result.keys())}")
                            
                            # Look for ContractInquire_Four which contains equipment details
                            if "ContractInquire_Four" in ds_result:
                                contract_four = ds_result["ContractInquire_Four"]
                                if isinstance(contract_four, list) and len(contract_four) > 0:
                                    # Get the first equipment item
                                    equipment_item = contract_four[0]
                                    if isinstance(equipment_item, dict) and "EQD_VALAMOUNT" in equipment_item:
                                        valuation_amount = str(equipment_item["EQD_VALAMOUNT"]).strip()
                                        logger.info(f"Found EQD_VALAMOUNT in ContractInquire_Four: {valuation_amount}")
                                    else:
                                        logger.warning("ContractInquire_Four item doesn't contain EQD_VALAMOUNT")
                                        logger.info(f"Available fields: {list(equipment_item.keys()) if isinstance(equipment_item, dict) else 'Not a dict'}")
                                else:
                                    logger.warning("ContractInquire_Four is empty or not a list")
                            else:
                                logger.warning("ContractInquire_Four not found in dsResult")
                                # Fallback: search through all dsResult items
                                for key, value in ds_result.items():
                                    if isinstance(value, list) and len(value) > 0:
                                        for item in value:
                                            if isinstance(item, dict) and "EQD_VALAMOUNT" in item:
                                                valuation_amount = str(item["EQD_VALAMOUNT"]).strip()
                                                logger.info(f"Found EQD_VALAMOUNT in {key}: {valuation_amount}")
                                                break
                                            if valuation_amount:
                                                break
                                    if valuation_amount:
                                        break
                        else:
                            # Fallback: search through the entire response
                            logger.info("No DATA/dsResult structure found, searching entire response...")
                            for key, value in equipment_api_response.items():
                                if isinstance(value, list) and len(value) > 0:
                                    for item in value:
                                        if isinstance(item, dict) and "EQD_VALAMOUNT" in item:
                                            valuation_amount = str(item["EQD_VALAMOUNT"]).strip()
                                            logger.info(f"Found EQD_VALAMOUNT in {key}: {valuation_amount}")
                                            break
                                        if valuation_amount:
                                            break
                                elif isinstance(value, dict) and "EQD_VALAMOUNT" in value:
                                    valuation_amount = str(value["EQD_VALAMOUNT"]).strip()
                                    logger.info(f"Found EQD_VALAMOUNT directly in {key}: {valuation_amount}")
                                    break
                                if valuation_amount:
                                    break
                    
                    if valuation_amount:
                        logger.info(f"✓ Successfully extracted valuation amount: {valuation_amount}")
                        return {
                            "contract_number": contract_no,
                            "valuation_amount": valuation_amount,
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                    else:
                        logger.warning("No EQD_VALAMOUNT field found in API response")
                        logger.info(f"API Response keys: {list(equipment_api_response.keys()) if isinstance(equipment_api_response, dict) else 'Not a dict'}")
                        logger.info(f"API Response structure: {json.dumps(equipment_api_response, indent=2)[:500]}...")
                        return {
                            "contract_number": contract_no,
                            "valuation_amount": "NO_EQD_VALAMOUNT_FOUND",
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        
                except json.JSONDecodeError as e:
                    logger.error(f"Failed to parse JSON response: {e}")
                    return {
                        "contract_number": contract_no,
                        "valuation_amount": None,
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
            else:
                logger.error("API call failed after retries")
                return {
                    "contract_number": contract_no,
                    "valuation_amount": None,
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
        except Exception as api_error:
            logger.error(f"Error making API call: {api_error}")
            return {
                "contract_number": contract_no,
                "valuation_amount": None,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
    except Exception as e:
        logger.error(f"Error processing contract {contract_no}: {e}")
        return {
            "contract_number": contract_no,
            "valuation_amount": f"ERROR: {str(e)}",
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

def display_results(contracts_data):
    """Display results in terminal"""
    try:
        if contracts_data:
            print("\n" + "="*80)
            print("VALUATION AMOUNT EXTRACTION SUMMARY")
            print("="*80)
            print(f"Total Contracts Processed: {len(contracts_data)}")
            print(f"Generated At: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("="*80)
            
            # Display each contract's valuation data
            print("\nEXTRACTED VALUATION DATA:")
            print("-" * 80)
            for contract_data in contracts_data:
                if contract_data:
                    contract_no = contract_data.get('contract_number', 'N/A')
                    valuation_amount = contract_data.get('valuation_amount', 'N/A')
                    # Normalize missing/error cases to #N/A for display
                    if (not valuation_amount) or (
                        isinstance(valuation_amount, str) and (
                            valuation_amount.startswith(('NO_', 'API_', 'ERROR'))
                        )
                    ):
                        valuation_amount = 0
                    print(f"\n{contract_no}:")
                    print(f"  Valuation Amount: {valuation_amount}")
                else:
                    print(f"\nNo data extracted")
            
            print("\n" + "="*80)
            return True
        else:
            print("\nNo contract data to display")
            return False
            
    except Exception as e:
        logger.error(f"Failed to display results: {e}")
        return False

def run_valuation_bot(contract_numbers):
    """
    Function to be called from NBD_MF_23_IA.py
    Returns a dictionary with contract numbers as keys and valuation amounts as values
    """
    try:
        logger.info(f"Running valuation bot for {len(contract_numbers)} contracts...")
        
        # Run the main function with provided contract numbers
        results = main(contract_numbers)
        
        if results:
            # Convert results to the expected format: {contract_number: valuation_amount}
            valuation_dict = {}
            for contract_data in results:
                if contract_data and contract_data.get('contract_number'):
                    contract_no = contract_data.get('contract_number')
                    valuation_amount = contract_data.get('valuation_amount', 'N/A')
                    # Normalize to #N/A for Excel mapping when missing/error
                    if (not valuation_amount) or (
                        isinstance(valuation_amount, str) and (
                            valuation_amount.startswith(('NO_', 'API_', 'ERROR'))
                        )
                    ):
                        valuation_amount = 0
                    valuation_dict[contract_no] = valuation_amount
            
            logger.info(f"Valuation bot completed. Processed {len(valuation_dict)} contracts.")
            return valuation_dict
        else:
            logger.warning("Valuation bot returned no results.")
            return {}
            
    except Exception as e:
        logger.error(f"Valuation bot execution failed: {e}")
        raise

# ---------------- Main flow ----------------
def main(contract_numbers=None):
    """Main execution function"""
    driver = None
    try:
        logger.info("Starting IA Working Initial Valuation Bot...")
        
        # Use provided contract numbers or load from Excel file
        if contract_numbers is None:
            contract_numbers = load_contract_numbers_from_excel()
        
        if not contract_numbers:
            logger.error("No contract numbers provided or loaded. Exiting.")
            return
        
        logger.info(f"Processing {len(contract_numbers)} contract numbers")
        
        # Start Chrome driver
        driver = start_driver()
        
        # Login to system
        if not login_to_system(driver):
            logger.error("Login failed. Exiting.")
            return
        
        # Navigate to Contract Inquiry
        if not navigate_to_contract_inquiry(driver):
            logger.error("Navigation to Contract Inquiry failed. Exiting.")
            return
        
        # Process each contract number
        contracts_data = []
        total_contracts = len(contract_numbers)
        
        for i, contract_no in enumerate(contract_numbers, 1):
            logger.info(f"Processing contract {i}/{total_contracts}: {contract_no}")
            
            # Process the contract
            result = process_contract(driver, contract_no)
            contracts_data.append(result)
            
            # Attempt to close Equipment Details popup if open
            try:
                close_btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.ID, "btn_02461_Equipment_DetailsBtnClose"))
                )
                try:
                    close_btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", close_btn)
                logger.info("Closed Equipment Details popup")
                time.sleep(0.5)
            except Exception:
                pass
            
            # Small delay between contracts
            time.sleep(2)
            
            # Clear the input field for the next contract (except for the last one)
            if i < total_contracts:
                try:
                    contract_input = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.ID, "strContractNo_RECContractInquiry"))
                    )
                    contract_input.clear()
                    logger.info("Cleared contract input field for next contract")
                except Exception as e:
                    logger.warning(f"Failed to clear contract input field: {e}")
        
        # Display results in terminal
        display_results(contracts_data)
        
        # Final summary
        successful_contracts = sum(1 for data in contracts_data if data and data.get("valuation_amount") and not data.get("valuation_amount").startswith(("NO_", "API_", "ERROR")))
        logger.info("="*80)
        logger.info("FINAL PROCESSING SUMMARY")
        logger.info("="*80)
        logger.info(f"Total Contracts: {total_contracts}")
        logger.info(f"Successfully Processed: {successful_contracts}")
        logger.info(f"Failed: {total_contracts - successful_contracts}")
        logger.info(f"Success Rate: {(successful_contracts/total_contracts)*100:.1f}%")
        logger.info("="*80)
        
        logger.info(f"Bot completed! Successfully processed {successful_contracts}/{total_contracts} contracts")
        
        return contracts_data
        
    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        raise
    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

if __name__ == "__main__":
    # Support passing contract numbers via CLI for standalone testing
    import sys
    try:
        cli_contracts = [arg.strip() for arg in sys.argv[1:] if arg.strip()]
        if cli_contracts:
            main(cli_contracts)
        else:
            main()
    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        exit(1)
