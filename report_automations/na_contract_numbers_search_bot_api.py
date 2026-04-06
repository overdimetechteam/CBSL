#!/usr/bin/env python3
"""
NA Contract Numbers Search Bot - API Version
Uses Selenium for authentication, then direct HTTP requests for API calls
"""

import requests
import json
import os
import time
from dotenv import load_dotenv
from cryptography.fernet import Fernet
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Load environment variables from multiple likely locations
try:
    _root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _this_dir = os.path.dirname(os.path.abspath(__file__))
    _env_paths = [
        os.path.join(_root_dir, '.env'),                 # repo root .env
        os.path.join(_this_dir, '.env'),                 # same-folder .env
        os.path.join(_this_dir, 'ATT58495.env'),         # provided env file
    ]
    loaded_any = False
    for _path in _env_paths:
        if os.path.exists(_path):
            load_dotenv(_path)
            loaded_any = True
    if not loaded_any:
        load_dotenv()
except Exception:
    load_dotenv()

# ---------------- Config ----------------
# Single API endpoint with multiple inquiry types
API_URL = "https://erp.assetline.lk/Recovery/RECContractInquiry/Inquire"

# Inquiry types to try for each contract (each yields different data)
INQUIRY_TYPES = [
    "MAININQ_ONE",      # Basic contract info
    "MAININQ_TWO",      # Financial data (when available)
    "MAININQ_LODAMT",   # Additional contract details
    "MAININQ_THREE",    # Extended contract info
    "MAININQ_FOUR"      # More contract details
]
BASE_URL = "https://erp.assetline.lk/"
RECOVERY_URL = "https://erp.assetline.lk/Application/Home/RECOVERY"
OUT_DIR = "outputs/contract_search_results"
LOGS_DIR = "logs"

# Create logs directory only (no output directory needed)
os.makedirs(LOGS_DIR, exist_ok=True)

# Retry configuration
MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds between retries

# Contract numbers will be read from Excel file
na_contract_numbers_data = []

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
            logger.warning(f"Excel file not found: {excel_file}")
            logger.info("Using fallback contract numbers for testing")
            return [
                "KRUV251895440",
                "TRUV251889870",
                "TRUV251895290",
                "WMUV251888440",
                "NKUV251891520",
                "AWAT250038070",
                "BRAT250038310"
            ]
        
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

# ---------------- Security Config ----------------
# Security settings
SECURE_LOGGING = True  # Do not log sensitive values in production
ENCRYPT_OUTPUT_FILES = False  # Encrypt output files (not used here)
CLEANUP_SENSITIVE_DATA = True  # Clean up sensitive data after processing

# Inline encryption key to decrypt ENC_USERNAME/ENC_PASSWORD.
# Paste your Fernet key string between the quotes, e.g. "mYFernetKey...".
# This script no longer reads from a .key file.
ENCRYPTION_KEY_INLINE = "xOXvPok_6fGgcEsth0rddzRzw2IPqaiAf38lyBMFy8A="

# ---------------- Logging Setup ----------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, 'contract_search_bot_api.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

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

        # Fallback: attempt to load from ATT58495.env explicitly if not present yet
        if not enc_username or not enc_password:
            try:
                _fallback_env = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ATT58495.env')
                if os.path.exists(_fallback_env):
                    load_dotenv(_fallback_env, override=True)
                    enc_username = os.getenv('ENC_USERNAME')
                    enc_password = os.getenv('ENC_PASSWORD')
                    logger.info("Loaded encrypted credentials from ATT58495.env")
            except Exception as _e:
                logger.warning("Failed loading ATT58495.env: %s", _e)

        # Secondary fallback: allow plaintext USERNAME/PASSWORD for local/debug
        if not enc_username or not enc_password:
            plain_user = os.getenv('USERNAME')
            plain_pass = os.getenv('PASSWORD')
            if plain_user and plain_pass:
                logger.warning("Using plaintext USERNAME/PASSWORD from environment (not encrypted)")
                return plain_user, plain_pass
            raise ValueError("ENC_USERNAME and ENC_PASSWORD must be set in environment variables or provide USERNAME/PASSWORD")
        
        logger.info("Found encrypted credentials, attempting to decrypt...")

        # Use inline key first; fall back to ENCRYPTION_KEY env var. No .key file is used.
        key_str = ENCRYPTION_KEY_INLINE.strip() if ENCRYPTION_KEY_INLINE else ""
        if not key_str:
            key_str = os.getenv('ENCRYPTION_KEY', '').strip()
        if not key_str:
            raise ValueError("Missing ENCRYPTION_KEY_INLINE and ENCRYPTION_KEY environment variable")

        key = key_str.encode()
        fernet = Fernet(key)
        
        # Decrypt credentials
        username = fernet.decrypt(enc_username.encode()).decode()
        password = fernet.decrypt(enc_password.encode()).decode()
        
        logger.info("Successfully decrypted credentials")
        logger.info("Credentials loaded successfully")
        return username, password
        
    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        raise

# ---------------- Selenium Authentication ----------------
def get_authenticated_session():
    """Use Selenium to login and get session cookies for API requests"""
    driver = None
    try:
        logger.info("Starting Selenium authentication...")
        
        # Setup Chrome driver
        options = Options()
        options.add_argument("--headless=new")  # Commented out to see the browser
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
        
        # Login process
        logger.info("Starting login process...")
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
            EC.presence_of_element_located((By.ID, "idSignIn"))
        )
        signin_button.click()
        logger.info("Sign in button clicked")
        
        # Wait for page to load and check if there's another button to click
        time.sleep(3)
        try:
            # Look for any button with id="idSignIn" that might be a submit button
            select_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//button[@id='idSignIn' and @type='submit']"))
            )
            select_button.click()
            logger.info("SELECT button clicked")
        except:
            logger.info("No additional SELECT button found, proceeding with login")
        
        time.sleep(3)
        logger.info("Login successful")
        
        # Navigate to Contract Inquiry page to ensure proper session
        logger.info("Navigating to Contract Inquiry...")
        try:
            driver.get(RECOVERY_URL)
            time.sleep(5)  # Increased wait time
            
            # Click on Information Center tile
            info_center_tile = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, 
                    "//div[contains(@class, 'metro-tile-body') and contains(@class, 'tile-type-menu')]//i[contains(@class, 'fa-file')]/.."))
            )
            info_center_tile.click()
            logger.info("Information Center tile clicked")
            time.sleep(3)
        except Exception as e:
            logger.error(f"Failed to navigate to Contract Inquiry: {e}")
            # Try to get cookies anyway
            logger.info("Attempting to get cookies from current page...")
        
        # Try to navigate to Contract Inquiry if possible
        try:
            # Click on Contract Inquiry tile
            contract_inquiry_tile = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH,
                    "//div[contains(@class, 'metro-tile-body')]//i[contains(@class, 'fa-user')]/.."))
            )
            contract_inquiry_tile.click()
            logger.info("Contract Inquiry tile clicked")
            time.sleep(3)
            
            # Click on Add button
            add_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "btn_02461_Add"))
            )
            add_button.click()
            logger.info("Add button clicked")
            time.sleep(3)
            
            logger.info("Successfully navigated to Contract Inquiry")
            
            # IMPORTANT: Simulate a contract search to activate the API
            logger.info("Activating API by simulating a contract search...")
            
            # Find the contract number input field
            contract_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "strContractNo_RECContractInquiry"))
            )
            logger.info("Found contract number input field")
            
            # Clear the field and paste a test contract number
            contract_input.clear()
            test_contract = na_contract_numbers_data[0]  # Use first contract as test
            contract_input.send_keys(test_contract)
            logger.info(f"Entered test contract number: {test_contract}")
            
            # Hit Enter to trigger the search
            contract_input.send_keys(Keys.RETURN)
            logger.info("Pressed Enter to trigger search")
            
            # Wait for the search to complete
            time.sleep(5)
            logger.info("API activation completed")
            

        except Exception as e:
            logger.warning(f"Could not complete full navigation: {e}")
            logger.info("Proceeding with cookie extraction from current page...")
        
        # Get verification token from the page
        try:
            verification_token = driver.find_element(By.NAME, "__RequestVerificationToken").get_attribute("value")
            if SECURE_LOGGING:
                logger.info("Found verification token successfully")
            else:
                logger.info(f"Found verification token: {verification_token[:10]}...")
        except:
            verification_token = ""
            logger.warning("Could not find verification token")
        
        # Get all cookies
        cookies = driver.get_cookies()
        logger.info(f"Retrieved {len(cookies)} cookies")
        
        # Log cookie details for debugging (secure mode)
        if SECURE_LOGGING:
            logger.info(f"Retrieved {len(cookies)} cookies successfully")
        else:
            for cookie in cookies:
                logger.info(f"Cookie: {cookie['name']} = {cookie['value'][:10]}... (domain: {cookie['domain']})")
        
        # Create a requests session with the cookies
        session = requests.Session()
        for cookie in cookies:
            session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])
        
        # Set common headers matching browser request exactly
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0",
            "Accept": "*/*",
            "Accept-Language": "en-US,en;q=0.9",
            "Origin": "https://erp.assetline.lk",
            "Referer": "https://erp.assetline.lk/Application/Home/RECOVERY",
            "X-Requested-With": "XMLHttpRequest",
            "Priority": "u=1, i",
            "Sec-Ch-Ua": '"Not;A=Brand";v="99", "Microsoft Edge";v="139", "Chromium";v="139"',
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin"
        })
        
        logger.info("Authenticated session created successfully")
        return session, verification_token
        
    except Exception as e:
        logger.error(f"Authentication failed: {e}")
        raise
    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

# ---------------- API Functions ----------------
def search_contract_api(session, contract_no, verification_token=""):
    """Search for a contract using multiple inquiry types and merge responses"""
    try:
        logger.info(f"Searching contract: {contract_no} using multiple inquiry types...")
        
        # Set headers matching the actual browser request exactly
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
        
        # Collect all responses from different inquiry types with context
        all_responses = []
        
        # First, get the client code from MAININQ_ONE
        client_code = ""
        try:
            logger.info(f"Getting client code from MAININQ_ONE for {contract_no}...")
            first_payload = {
                "strContractNo_RECContractInquiry": contract_no,
                "strInquireType_RECContractInquiry": "MAININQ_ONE",
                "STR_FORM_ID": "02461",
                "STR_FUNCTION_ID": "CR",
                "STR_PREMIS": "HOD",
                "STR_INSTANT": "ALCL",
                "STR_APP_ID": "00023"
            }
            
            if verification_token:
                first_payload["__RequestVerificationToken"] = verification_token
            
            first_response = session.post(API_URL, data=first_payload, headers=headers, timeout=30)
            
            if first_response.status_code == 200:
                first_data = first_response.json()
                if first_data != "LOGOUT" and isinstance(first_data, dict):
                    # Extract client code from the first response
                    if "DATA" in first_data and "dsResult" in first_data["DATA"]:
                        ds_result = first_data["DATA"]["dsResult"]
                        if "ContractInquire_One" in ds_result and ds_result["ContractInquire_One"]:
                            contract_info = ds_result["ContractInquire_One"][0]
                            client_code = contract_info.get("CON_CLMCODE", "")
                            logger.info(f"✓ Extracted client code: {client_code} for {contract_no}")
                    
                    # Store the first response
                    all_responses.append(("MAININQ_ONE", first_data))
                    
        except Exception as e:
            logger.warning(f"Failed to get client code for {contract_no}: {e}")
        
        # Now loop through remaining inquiry types with enhanced payloads
        remaining_inquiry_types = [t for t in INQUIRY_TYPES if t != "MAININQ_ONE"]
        
        for i, inquiry_type in enumerate(remaining_inquiry_types, 2):
            try:
                logger.info(f"Making API call {i}/{len(INQUIRY_TYPES)} with inquiry type: {inquiry_type}")
                
                # Prepare payload for this inquiry type - match browser exactly
                payload = {
                    "strContractNo_RECContractInquiry": contract_no,
                    "strInquireType_RECContractInquiry": inquiry_type,
                    "STR_FORM_ID": "02461",
                    "STR_FUNCTION_ID": "CR",
                    "STR_PREMIS": "HOD",
                    "STR_INSTANT": "ALCL",
                    "STR_APP_ID": "00023"
                }
                
                # Add client code for MAININQ_TWO (this is the key fix!)
                if inquiry_type == "MAININQ_TWO" and client_code:
                    payload["strClientCode_RECContractInquiry"] = client_code
                    payload["strActualDate_RECContractInquiry"] = ""
                    logger.info(f"Added client code {client_code} to MAININQ_TWO payload")
                
                # Add verification token if available
                if verification_token:
                    payload["__RequestVerificationToken"] = verification_token
                
                # Make the API request
                response = session.post(API_URL, data=payload, headers=headers, timeout=30)
                
                logger.info(f"API call {i} response status: {response.status_code}")
                
                if response.status_code == 200:
                    try:
                        data = response.json()
                        if data == "LOGOUT":
                            logger.error(f"Session expired for {contract_no} on inquiry type {inquiry_type}")
                            return None
                        
                        logger.info(f"✓ Got data for {contract_no} with {inquiry_type}")
                        
                        # Log response structure for debugging
                        if isinstance(data, dict):
                            logger.info(f"Inquiry type {inquiry_type} response structure: {list(data.keys())}")
                            if "DATA" in data and "dsResult" in data["DATA"]:
                                ds_result = data["DATA"]["dsResult"]
                                logger.info(f"Inquiry type {inquiry_type} dsResult structure: {list(ds_result.keys())}")
                        else:
                            logger.info(f"Inquiry type {inquiry_type} response type: {type(data)}")
                        
                        # Store response with inquiry type context
                        all_responses.append((inquiry_type, data))
                        
                        # Small delay between API calls
                        time.sleep(0.5)
                        
                    except json.JSONDecodeError as e:
                        logger.warning(f"JSON parse failed for {contract_no} / {inquiry_type}: {e}")
                        continue
                else:
                    logger.warning(f"API call failed for {contract_no} with {inquiry_type}: {response.status_code}")
                    continue
                    
            except Exception as e:
                logger.warning(f"API call failed for contract {contract_no} with {inquiry_type}: {e}")
                continue
        
        if all_responses:
            logger.info(f"Successfully retrieved {len(all_responses)} API responses for {contract_no}")
            return all_responses
        else:
            logger.warning(f"No successful API responses for {contract_no}")
            return None
            
    except Exception as e:
        logger.error(f"Search failed for contract {contract_no}: {e}")
        return None

def extract_contract_data(contract_no, api_responses):
    """Extract the required fields from multiple API responses by merging them"""
    try:
        if not api_responses:
            logger.warning(f"No API responses for {contract_no}")
            return None
        
        # Initialize extracted data structure with source tracking
        extracted_data = {
            "contract_number": contract_no,
            "client_code": "",
            "equipment": "",
            "contract_period": "",
            "frequency": "",
            "interest_rate": "",
            "contract_amount": "",
            # AT_limit removed from final response spec; keep only internally if needed
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "sources": {}  # Track which inquiry type provided which data
        }
        
        # Process responses with inquiry type context
        for i, (inquiry_type, api_response) in enumerate(api_responses):
            try:
                if not api_response:
                    logger.warning(f"Empty response {i+1} for {contract_no} with {inquiry_type}")
                    continue
                
                # Log the response state for debugging
                response_state = api_response.get("STATE", "UNKNOWN")
                logger.info(f"Processing {inquiry_type} response {i+1} for {contract_no} - STATE: {response_state}")
                
                # Even if STATE != TRUE, the response might contain useful data
                if response_state != "TRUE":
                    logger.warning(f"{inquiry_type} response {i+1} has STATE != TRUE: {response_state}")
                    # Don't skip - examine the response anyway
                
                data = api_response.get("DATA", {})
                
                # Handle different DATA field types
                if isinstance(data, str):
                    logger.info(f"{inquiry_type} response {i+1}: DATA field is a string: {data}")
                    # Even string responses might contain useful information
                    if "NO DATA FOUND" in data:
                        logger.info(f"{inquiry_type} response {i+1}: Contains 'NO DATA FOUND' message")
                    # Continue processing to see if there are other fields
                
                # Get dsResult - handle both dict and other types
                ds_result = {}
                if isinstance(data, dict):
                    ds_result = data.get("dsResult", {})
                    logger.info(f"{inquiry_type} response {i+1}: Found dsResult with keys: {list(ds_result.keys())}")
                else:
                    logger.info(f"{inquiry_type} response {i+1}: DATA field is not a dict, type: {type(data)}")
                
                logger.info(f"Processing {inquiry_type} response {i+1} for {contract_no}")
                
                # Comprehensive data extraction from ALL responses regardless of state
                logger.info(f"Examining {inquiry_type} response {i+1} for {contract_no} - searching all data structures...")
                
                # Log the complete structure of this response
                if isinstance(ds_result, dict):
                    logger.info(f"{inquiry_type} response {i+1} structure: {list(ds_result.keys())}")
                else:
                    logger.info(f"{inquiry_type} response {i+1}: dsResult is not a dict, type: {type(ds_result)}")
                
                # Extract basic contract info from any response that has it
                if "ContractInquire_One" in ds_result:
                    contract_one = ds_result.get("ContractInquire_One", [])
                    if contract_one and not extracted_data["client_code"]:
                        contract_info = contract_one[0]
                        extracted_data["client_code"] = contract_info.get("CON_CLMCODE", "")
                        extracted_data["sources"]["client_code"] = inquiry_type
                        logger.info(f"{inquiry_type} response {i+1}: Found client_code: {extracted_data['client_code']}")
                
                # Extract equipment info from any response that has it
                if "ContractInquire_Four" in ds_result:
                    contract_four = ds_result.get("ContractInquire_Four", [])
                    if contract_four and not extracted_data["equipment"]:
                        equipment_info = contract_four[0]
                        extracted_data["equipment"] = equipment_info.get("EQT_DESC", "")
                        extracted_data["sources"]["equipment"] = inquiry_type
                        logger.info(f"{inquiry_type} response {i+1}: Found equipment: {extracted_data['equipment']}")
                
                                # Target MAININQ_TWO MainInq_1 specifically for financial data
                if inquiry_type == "MAININQ_TWO" and "MainInq_1" in ds_result:
                    logger.info(f"{inquiry_type} response {i+1}: Found MainInq_1 - extracting financial data")
                    main_inq_1 = ds_result["MainInq_1"]
                    
                    if isinstance(main_inq_1, list) and len(main_inq_1) > 0:
                        financial_data = main_inq_1[0]
                        if isinstance(financial_data, dict):
                            logger.info(f"{inquiry_type} response {i+1}: MainInq_1 contains fields: {list(financial_data.keys())}")
                            
                            # Extract financial fields with correct uppercase names
                            if not extracted_data["contract_period"]:
                                value = financial_data.get("CON_PERIOD", "")
                                if value and value != "":
                                    extracted_data["contract_period"] = value
                                    extracted_data["sources"]["contract_period"] = inquiry_type
                                    logger.info(f"{inquiry_type} response {i+1}: Found CON_PERIOD: {value}")
                            
                            if not extracted_data["frequency"]:
                                value = financial_data.get("CON_RNTFREQ", "")
                                if value and value != "":
                                    extracted_data["frequency"] = value
                                    extracted_data["sources"]["frequency"] = inquiry_type
                                    logger.info(f"{inquiry_type} response {i+1}: Found CON_RNTFREQ: {value}")
                            
                            if not extracted_data["interest_rate"]:
                                value = financial_data.get("CON_INTRATE", "")
                                if value and value != "":
                                    extracted_data["interest_rate"] = value
                                    extracted_data["sources"]["interest_rate"] = inquiry_type
                                    logger.info(f"{inquiry_type} response {i+1}: Found CON_INTRATE: {value}")
                            
                            if not extracted_data["contract_amount"]:
                                value = financial_data.get("CON_CONTAMT", "")
                                if value and value != "":
                                    extracted_data["contract_amount"] = value
                                    extracted_data["sources"]["contract_amount"] = inquiry_type
                                    logger.info(f"{inquiry_type} response {i+1}: Found CON_CONTAMT: {value}")
                            
                            # Capture AT_limit internally to support fallback for contract_amount
                            if "_internal_at_limit" not in extracted_data:
                                value = financial_data.get("TCL_CONAMOUNT", "")
                                if value and value != "0" and value != 0:
                                    extracted_data["_internal_at_limit"] = value
                                    logger.info(f"{inquiry_type} response {i+1}: Found TCL_CONAMOUNT: {value}")
                    else:
                        logger.warning(f"{inquiry_type} response {i+1}: MainInq_1 is not a list or is empty")
                else:
                    # For other inquiry types, just log what we found (no financial data expected)
                    logger.info(f"{inquiry_type} response {i+1}: No financial data expected from this inquiry type")
                
                # Special handling for responses with "NO DATA FOUND" - might contain hidden data
                if response_state == "FALSE" and isinstance(data, str) and "NO DATA FOUND" in data:
                    logger.info(f"{inquiry_type} response {i+1}: Found 'NO DATA FOUND' response - examining for hidden data")
                    # The data might be in a different format or the response might have additional structure
                
                # Examine the ENTIRE API response for any hidden financial data
                logger.info(f"{inquiry_type} response {i+1}: Examining complete API response structure...")
                for response_key, response_value in api_response.items():
                    if response_key not in ["STATE", "DATA"]:
                        logger.info(f"{inquiry_type} response {i+1}: Found additional response field: {response_key} = {response_value}")
                        # Check if this field contains financial data
                        if isinstance(response_value, dict):
                            for sub_key, sub_value in response_value.items():
                                if any(field in sub_key.upper() for field in ["PERIOD", "FREQ", "RATE", "AMOUNT", "LIMIT"]):
                                    logger.info(f"{inquiry_type} response {i+1}: Found potential financial field: {sub_key} = {sub_value}")
                        elif isinstance(response_value, str) and any(field in response_value.upper() for field in ["PERIOD", "FREQ", "RATE", "AMOUNT", "LIMIT"]):
                            logger.info(f"{inquiry_type} response {i+1}: Found potential financial data in {response_key}: {response_value}")
                
            except Exception as e:
                logger.error(f"Error processing {inquiry_type} response {i+1} for {contract_no}: {e}")
                continue
        
        if SECURE_LOGGING:
            logger.info(f"Successfully extracted data for {contract_no}")
        else:
            logger.info(f"Extracted data for {contract_no}: {extracted_data}")
        
        return extracted_data
        
    except Exception as e:
        logger.error(f"Failed to extract data for {contract_no}: {e}")
        return None

def save_contract_data(contract_no, api_responses, extracted_data):
    """Return extracted data for terminal display"""
    try:
        if extracted_data:
            # Normalize and derive contract_amount: if 0/blank, take AT_limit; omit AT_limit in display
            def _to_float(value):
                try:
                    if value is None:
                        return None
                    if isinstance(value, (int, float)):
                        return float(value)
                    s = str(value).replace(",", "").strip()
                    if s == "":
                        return None
                    return float(s)
                except Exception:
                    return None

            amount = extracted_data.get("contract_amount", "")
            amount_num = _to_float(amount)
            if amount_num is None or amount_num == 0.0:
                at_limit = extracted_data.get("_internal_at_limit", "")
                at_limit_num = _to_float(at_limit)
                if at_limit_num is not None and at_limit_num != 0.0:
                    amount_num = at_limit_num

            # Return only the essential contract data for display (without AT_limit)
            display_data = {
                "contract_number": extracted_data.get("contract_number", ""),
                "client_code": extracted_data.get("client_code", ""),
                "equipment": extracted_data.get("equipment", ""),
                "contract_period": extracted_data.get("contract_period", ""),
                "frequency": extracted_data.get("frequency", ""),
                "interest_rate": extracted_data.get("interest_rate", ""),
                "contract_amount": amount_num if amount_num is not None else extracted_data.get("contract_amount", "")
            }
            return display_data
        return None
        
    except Exception as e:
        logger.error(f"Failed to process data for {contract_no}: {e}")
        return None

def _is_blank_display_data(display_data: dict) -> bool:
    """Return True if all principal fields are blank/empty/zero-like."""
    if not isinstance(display_data, dict):
        return True
    keys = [
        "client_code",
        "equipment",
        "contract_period",
        "frequency",
        "interest_rate",
        "contract_amount",
    ]
    for k in keys:
        v = display_data.get(k, "")
        if isinstance(v, str):
            if v.strip() != "":
                return False
        else:
            if v not in (None, "", 0, "0"):
                return False
    return True

def create_summary_json(contracts_data):
    """Create and display summary data in terminal"""
    try:
        if contracts_data:
            # Create summary data structure
            summary_data = {
                "summary_info": {
                    "total_contracts": len(contracts_data),
                    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "api_url": API_URL,
                    "inquiry_types_used": INQUIRY_TYPES
                },
                "contracts": contracts_data
            }
            
            # Display summary info
            print("\n" + "="*80)
            print("CONTRACT SEARCH SUMMARY")
            print("="*80)
            print(f"Total Contracts Processed: {len(contracts_data)}")
            print(f"Generated At: {summary_data['summary_info']['generated_at']}")
            print(f"API URL: {summary_data['summary_info']['api_url']}")
            print(f"Inquiry Types Used: {', '.join(summary_data['summary_info']['inquiry_types_used'])}")
            print("="*80)
            
            # Display each contract's data
            print("\nEXTRACTED CONTRACT DATA:")
            print("-" * 80)
            first_cn = None
            if contracts_data and isinstance(contracts_data[0], dict):
                first_cn = contracts_data[0].get("contract_number")
            total_items = len(contracts_data)
            for idx, contract_data in enumerate(contracts_data, 1):
                if contract_data:
                    label = f"Contract {idx}:"
                    # If the last item is a re-run of the first contract, label accordingly
                    if idx == total_items and first_cn and isinstance(contract_data, dict):
                        if contract_data.get("contract_number") == first_cn:
                            label = "Contract 1 (retry at the end):"
                    print(f"\n{label}")
                    print(json.dumps(contract_data, indent=6, ensure_ascii=False))
                else:
                    print(f"\nContract {idx}: No data extracted")
            
            print("\n" + "="*80)
            return True
        else:
            print("\nNo contract data to display")
            return False
            
    except Exception as e:
        logger.error(f"Failed to create summary: {e}")
        return False

# File saving function removed - data is now displayed in terminal

# Cleanup function removed - no files to clean up

def process_contract_with_retry(session, contract_no, verification_token, max_retries=MAX_RETRIES):
    """Process a contract with automatic retry on failure"""
    for attempt in range(1, max_retries + 1):
        try:
            logger.info(f"Processing {contract_no} - Attempt {attempt}/{max_retries}")
            
            # Search contract via multiple API calls
            api_responses = search_contract_api(session, contract_no, verification_token)
            
            if api_responses:
                # Extract required data from all responses
                extracted_data = extract_contract_data(contract_no, api_responses)
                
                # Get display data for summary
                display_data = save_contract_data(contract_no, api_responses, extracted_data)
                # Treat a completely blank display_data as a failure to force retry
                if display_data and not _is_blank_display_data(display_data):
                    logger.info(f"✓ Successfully processed {contract_no} on attempt {attempt}")
                    return display_data, True
                else:
                    logger.warning(f"Blank or insufficient data for {contract_no} on attempt {attempt}")
            else:
                logger.warning(f"No response data for {contract_no} on attempt {attempt}")
            
            # If we get here, the attempt failed
            if attempt < max_retries:
                logger.info(f"Retrying {contract_no} in {RETRY_DELAY} seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(RETRY_DELAY)
            else:
                logger.error(f"Failed to process {contract_no} after {max_retries} attempts")
                
        except Exception as e:
            logger.error(f"Error processing {contract_no} on attempt {attempt}: {e}")
            if attempt < max_retries:
                logger.info(f"Retrying {contract_no} in {RETRY_DELAY} seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(RETRY_DELAY)
            else:
                logger.error(f"Failed to process {contract_no} after {max_retries} attempts due to errors")
    
    return None, False

def validate_settings():
    """Validate basic settings"""
    logger.info("Running in DEBUG mode - security features disabled for debugging")
    logger.warning("SECURE_LOGGING is disabled - sensitive data may be logged")
    logger.warning("ENCRYPT_OUTPUT_FILES is disabled - output files are not encrypted")
    return True

# ---------------- Main flow ----------------
def main(contract_numbers=None):
    """Main execution function. Optionally accepts a list of contract numbers to process."""
    try:
        logger.info("Starting NA Contract Numbers Search Bot (API Version)...")
        
        # Validate settings
        if not validate_settings():
            logger.error("Settings validation failed.")
            return
        
        # Decide contract numbers source
        global na_contract_numbers_data
        if contract_numbers and isinstance(contract_numbers, list) and len(contract_numbers) > 0:
            na_contract_numbers_data = [str(c).strip() for c in contract_numbers if str(c).strip()]
            logger.info(f"Using contract numbers from CLI: {len(na_contract_numbers_data)} provided")
        else:
            # Load contract numbers from Excel file
            na_contract_numbers_data = load_contract_numbers_from_excel()
        
        # Get authenticated session
        session, verification_token = get_authenticated_session()
        
        # Process each contract number with retry mechanism
        successful_searches = 0
        total_contracts = len(na_contract_numbers_data)
        contracts_data = []  # Collect all contract data for summary
        first_display_data = None  # Hold first contract's data for later appending
        
        for i, contract_no in enumerate(na_contract_numbers_data, 1):
            logger.info(f"Processing contract {i}/{total_contracts}: {contract_no}")
            
            # Use retry mechanism for processing contracts
            display_data, success = process_contract_with_retry(session, contract_no, verification_token)
            
            if success and display_data:
                contracts_data.append(display_data)
                successful_searches += 1
                logger.info(f"✓ Successfully processed {contract_no}")
                # Capture first contract's data
                if i == 1:
                    first_display_data = display_data
            else:
                logger.error(f"✗ Failed to process {contract_no} after all retry attempts")
                contracts_data.append(None)  # Add None for failed contracts
                if i == 1:
                    first_display_data = None
            
            # Small delay between contracts
            time.sleep(1)

        # Re-run Contract 1 via API at the end to get proper data extraction
        if na_contract_numbers_data:
            first_contract = na_contract_numbers_data[0]
            logger.info(f"Re-running Contract 1 via API at the end: {first_contract}")
            try:
                # Re-run the first contract with fresh API calls
                display_data_rerun, success_rerun = process_contract_with_retry(session, first_contract, verification_token)
                # Append if re-run produced any non-blank data; otherwise keep original entry only
                if display_data_rerun and not _is_blank_display_data(display_data_rerun):
                    # Keep original first entry (even if blank) and append the re-run result at the end
                    contracts_data.append(display_data_rerun)
                    logger.info(f"Successfully re-ran Contract 1 via API and appended as last item (keeping original entry)")
                else:
                    logger.warning(f"Re-run of Contract 1 failed; keeping only the original entry")
            except Exception as e:
                logger.error(f"Error re-running Contract 1: {e}")
                # Fallback: do nothing; keep original entry only

        # Recompute summary counts after potential de-duplication
        successful_searches = sum(1 for d in contracts_data if isinstance(d, dict))
        total_contracts = len(contracts_data)

        # Display summary in terminal
        create_summary_json(contracts_data)
        
        # Final summary with retry information
        failed_contracts = total_contracts - successful_searches
        logger.info("="*80)
        logger.info("FINAL PROCESSING SUMMARY")
        logger.info("="*80)
        logger.info(f"Total Contracts: {total_contracts}")
        logger.info(f"Successfully Processed: {successful_searches}")
        logger.info(f"Failed After Retries: {failed_contracts}")
        logger.info(f"Success Rate: {(successful_searches/total_contracts)*100:.1f}%")
        logger.info(f"Retry Configuration: Max {MAX_RETRIES} attempts, {RETRY_DELAY}s delay")
        logger.info("="*80)
        
        if failed_contracts > 0:
            logger.warning(f"⚠️  {failed_contracts} contracts failed after all retry attempts")
            logger.warning("Consider checking network connectivity or API availability")
        else:
            logger.info("All contracts processed successfully!")
        
        logger.info(f"Bot completed! Successfully processed {successful_searches}/{total_contracts} contracts")
        
    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        raise

if __name__ == "__main__":
    import sys
    # Support passing contract numbers via CLI: python na_contract_numbers_search_bot_api.py CN1 CN2 ...
    cli_contracts = [arg.strip() for arg in sys.argv[1:] if arg.strip()]
    if cli_contracts:
        main(cli_contracts)
    else:
        main()