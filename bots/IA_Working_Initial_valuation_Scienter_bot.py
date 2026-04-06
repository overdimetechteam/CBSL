#!/usr/bin/env python3
"""
IA Working Initial Valuation Scienter Bot
Extracts FD amounts from Scienter system for contract numbers from NBD_MF_23_IA.py
"""

import os
import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from cryptography.fernet import Fernet
import pandas as pd
from pathlib import Path

# Load environment variables
try:
    _root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _env_path = os.path.join(_root_dir, '.env')
    if os.path.exists(_env_path):
        load_dotenv(_env_path)
    else:
        load_dotenv()
except Exception:
    load_dotenv()

# Configuration
SCIENTER_LOGIN_URL = "https://efn.assetline.lk:1000/#/"
SCIENTER_REPORT_URL = "https://efn.assetline.lk:1009/Report/FDquery.aspx"
LOGS_DIR = "logs"

# Create logs directory
os.makedirs(LOGS_DIR, exist_ok=True)


# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, 'scienter_valuation_bot.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def load_credentials():
    """Load and decrypt credentials from environment variables"""
    try:
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
        return username, password
        
    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        raise

def load_contract_numbers_from_excel():
    """Load contract numbers from the Excel file used by NBD_MF_23_IA.py"""
    try:
        # Get the parent directory (Script directory)
        script_dir = Path(__file__).parent.parent
        
        # Path to the Excel file (same as used in NBD_MF_23_IA.py)
        excel_file = script_dir / "Prod. wise Class. of Loans - Jul 2025.xlsb"
        
        if not excel_file.exists():
            logger.warning(f"Excel file not found: {excel_file}")
            logger.info("Using fallback contract numbers for testing")
            return [
                "BRAT250040380",
                "BRAT250040440",
                "HQDL250007420"
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

def setup_driver():
    """Setup Chrome driver with appropriate options"""
    try:
        options = Options()
        # options.add_argument("--headless=new")  # Commented out to see browser
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-plugins")
        options.add_argument("--disable-blink-features=AutomationControlled")
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
        logger.error(f"Failed to setup Chrome driver: {e}")
        raise

def login_to_scienter(driver, username, password):
    """Login to Scienter system"""
    try:
        logger.info("Starting Scienter login process...")
        
        # Navigate to login page
        driver.get(SCIENTER_LOGIN_URL)
        time.sleep(5)  # Increased wait time
        
        # Try multiple selectors for username field
        username_input = None
        username_selectors = [
            "input[formcontrolname='username']",
            "input[placeholder='Username']",
            "input[type='text']",
            "input[name='username']"
        ]
        
        for selector in username_selectors:
            try:
                username_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                logger.info(f"Found username field with selector: {selector}")
                break
            except TimeoutException:
                continue
        
        if not username_input:
            raise Exception("Could not find username input field")
        
        username_input.clear()
        username_input.send_keys(username)
        logger.info("Username entered")
        
        # Try multiple selectors for password field
        password_input = None
        password_selectors = [
            "input[formcontrolname='password']",
            "input[placeholder='Password']",
            "input[type='password']",
            "input[name='password']"
        ]
        
        for selector in password_selectors:
            try:
                password_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                logger.info(f"Found password field with selector: {selector}")
                break
            except TimeoutException:
                continue
        
        if not password_input:
            raise Exception("Could not find password input field")
        
        password_input.clear()
        password_input.send_keys(password)
        logger.info("Password entered")
        
        # Try multiple selectors for login button
        login_button = None
        login_selectors = [
            "button[type='submit']",
            "input[type='submit']",
            "button:contains('Login')",
            "button:contains('Sign In')",
            "input[value='Login']",
            "input[value='Sign In']"
        ]
        
        for selector in login_selectors:
            try:
                login_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                logger.info(f"Found login button with selector: {selector}")
                break
            except TimeoutException:
                continue
        
        if not login_button:
            # Try pressing Enter on password field as fallback
            password_input.send_keys(Keys.RETURN)
            logger.info("Pressed Enter on password field")
        else:
            login_button.click()
            logger.info("Login button clicked")
        
        # Wait for login to complete and check for redirects
        time.sleep(10)  # Increased wait time
        
        # Check if we're still on the login page or if we've been redirected
        current_url = driver.current_url
        logger.info(f"Current URL after login attempt: {current_url}")
        
        # If we're still on the login page, the login might have failed
        if "login" in current_url.lower() or "signin" in current_url.lower():
            logger.warning("Still on login page - login might have failed")
            # Try to find error messages
            try:
                error_elements = driver.find_elements(By.CSS_SELECTOR, ".error, .alert, .message")
                for error in error_elements:
                    if error.is_displayed():
                        logger.error(f"Login error message: {error.text}")
            except:
                pass
            return False
        
        logger.info("Login successful")
        return True
        
    except Exception as e:
        logger.error(f"Login failed: {e}")
        return False

def navigate_to_fd_query(driver):
    """Navigate to FD Query page through menu structure"""
    try:
        logger.info("Navigating to FD Query page through menu...")
        
        # Wait for the page to load after login
        time.sleep(2)
        
        # Step 1: Click on "Fixed Deposit" menu item
        logger.info("Looking for Fixed Deposit menu item...")
        fixed_deposit_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "h2.hovicon.effect-3.sub-a"))
        )
        fixed_deposit_menu.click()
        logger.info("Clicked on Fixed Deposit menu")
        time.sleep(3)
        
        # Step 2: Click on first "QUERY" submenu
        logger.info("Looking for first QUERY submenu...")
        query_menu_1 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span.rmText.rmExpandDown"))
        )
        query_menu_1.click()
        logger.info("Clicked on first QUERY submenu")
        time.sleep(2)
        
        # Step 3: Click on second "QUERY" submenu
        logger.info("Looking for second QUERY submenu...")
        query_menus = driver.find_elements(By.CSS_SELECTOR, "span.rmText.rmExpandDown")
        if len(query_menus) >= 2:
            query_menu_2 = query_menus[1]  # Get the second QUERY menu
            query_menu_2.click()
            logger.info("Clicked on second QUERY submenu")
            time.sleep(3)
        else:
            logger.warning("Second QUERY submenu not found, trying alternative approach...")
            # Try clicking on any visible QUERY menu
            for menu in query_menus:
                if menu.is_displayed() and menu.is_enabled():
                    menu.click()
                    logger.info("Clicked on available QUERY submenu")
                    time.sleep(3)
                    break
        
        # Wait for the FD Query page to load
        time.sleep(5)
        
        # Verify we're on the correct page
        current_url = driver.current_url
        logger.info(f"Current URL after navigation: {current_url}")
        
        if "FDquery" in current_url or "fdquery" in current_url.lower():
            logger.info("Successfully navigated to FD Query page")
            return True
        else:
            logger.warning(f"May not be on FD Query page. Current URL: {current_url}")
            # Try direct navigation as fallback
            logger.info("Attempting direct navigation as fallback...")
            driver.get(SCIENTER_REPORT_URL)
            time.sleep(5)
            return True
        
    except Exception as e:
        logger.error(f"Failed to navigate to FD Query through menu: {e}")
        # Try direct navigation as fallback
        try:
            logger.info("Attempting direct navigation as fallback...")
            driver.get(SCIENTER_REPORT_URL)
            time.sleep(5)
            logger.info("Fallback navigation successful")
            return True
        except Exception as fallback_e:
            logger.error(f"Fallback navigation also failed: {fallback_e}")
            return False

def extract_fd_amount_for_contract(driver, contract_number):
    """Extract FD amount for a specific contract number"""
    try:
        logger.info(f"Processing contract: {contract_number}")
        
        # Step 4.1: Select "Loan Number" from dropdown (by exact ID)
        search_by_dropdown = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtSearchBy"))
        )
        select = Select(search_by_dropdown)
        select.select_by_value("LN")  # value="LN" is Loan Number
        logger.info("Selected 'Loan Number' from dropdown")
        
        # Step 4.2: Enter contract number
        fd_no_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtFDNo"))
        )
        fd_no_input.clear()
        fd_no_input.send_keys(contract_number)
        logger.info(f"Entered contract number: {contract_number}")
        
        # Step 4.3: Click search
        search_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnDisplay"))
        )
        search_button.click()
        logger.info("Clicked search button")
        
        # Wait for results table to load (support both specific id and generic class)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ContentPlaceHolder1_FDList_ctl00, table.Vardena_b_8_n"))
            )
            logger.info("Results table loaded for Loan Number search")
        except TimeoutException:
            logger.warning("Results table did not appear after Loan Number search")
            return None
        
        # Step 4.4: Get FD number from the first result (first row, first column link)
        try:
            # Prefer table by id; if not present, fallback to class-based table
            fd_number_link = None
            try:
                fd_number_link = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ContentPlaceHolder1_FDList_ctl00 tbody tr:first-child td:first-child a"))
                )
            except TimeoutException:
                fd_number_link = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table.Vardena_b_8_n tbody tr:first-child td:first-child a"))
                )
            fd_number = fd_number_link.text.strip()
            logger.info(f"Found FD number: {fd_number}")
        except TimeoutException:
            logger.warning(f"No FD number found for contract {contract_number}")
            return None
        
        # Step 4.5: Change to FD Number search and search again (with resilient fallbacks)
        logger.info("Switching dropdown to 'FD Number'")
        try:
            dropdown_el = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtSearchBy"))
            )
            Select(dropdown_el).select_by_value("FD")
            try:
                logger.info(f"Dropdown now selected: {Select(dropdown_el).first_selected_option.text}")
            except Exception:
                logger.info("Dropdown switched to FD (selected text not readable)")
        except Exception as e:
            logger.warning(f"Standard dropdown select failed ({e}); trying JS to set value to FD")
            driver.execute_script(
                "var el=document.getElementById('ctl00_ContentPlaceHolder1_txtSearchBy');"
                "if(el){el.value='FD'; if(document.createEvent){var ev=document.createEvent('HTMLEvents'); ev.initEvent('change',true,false); el.dispatchEvent(ev);} else if(el.fireEvent){el.fireEvent('onchange');}}"
            )
        time.sleep(0.3)

        logger.info(f"Entering FD number into input: {fd_number}")
        try:
            # Re-locate input to avoid stale refs after dropdown change
            fd_no_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtFDNo"))
            )
            fd_no_input.clear()
            fd_no_input.send_keys(fd_number)
        except Exception as e:
            logger.warning(f"Typing FD number failed ({e}); setting via JS")
            driver.execute_script(
                "var el=document.getElementById('ctl00_ContentPlaceHolder1_txtFDNo'); if(el){el.value=arguments[0];}" , fd_number
            )
            # Trigger change event for good measure
            driver.execute_script(
                "var el=document.getElementById('ctl00_ContentPlaceHolder1_txtFDNo');"
                "if(el && document.createEvent){var ev=document.createEvent('HTMLEvents'); ev.initEvent('change',true,false); el.dispatchEvent(ev);}"
            )
        time.sleep(0.3)

        logger.info("Clicking Search button for FD Number search")
        clicked = False
        try:
            # Re-locate search button to avoid stale refs
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnDisplay"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
            time.sleep(0.2)
            search_button.click()
            clicked = True
        except Exception as e:
            logger.warning(f"Standard click failed ({e}); trying JS click")
            try:
                driver.execute_script(
                    "var btn=document.getElementById('ctl00_ContentPlaceHolder1_btnDisplay'); if(btn){btn.click();}"
                )
                clicked = True
            except Exception as e2:
                logger.warning(f"JS click failed ({e2}); invoking __doPostBack")
                driver.execute_script("__doPostBack('ctl00$ContentPlaceHolder1$btnDisplay','')")
                clicked = True
        if clicked:
            logger.info(f"Search clicked. Waiting for results table for FD number {fd_number}...")
        logger.info(f"Search clicked. Waiting for results table for FD number {fd_number}...")
        
        # Wait for FD search results table
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ContentPlaceHolder1_FDList_ctl00, table.Vardena_b_8_n"))
            )
            logger.info("Results table loaded for FD Number search (table present)")
        except TimeoutException:
            logger.warning("Results table did not appear after FD Number search")
            return None
        
        # Step 4.6: Click the last row first column link (prefer class-based table per spec)
        try:
            # Gather links from either table structure
            fd_links = driver.find_elements(By.CSS_SELECTOR, "#ctl00_ContentPlaceHolder1_FDList_ctl00 tbody tr td:first-child a")
            if not fd_links:
                fd_links = driver.find_elements(By.CSS_SELECTOR, "table.Vardena_b_8_n tbody tr td:first-child a")
            if not fd_links:
                logger.warning(f"No FD links found for contract {contract_number}")
                return None

            # Get the index of the last link
            last_link_index = len(fd_links) - 1
            logger.info(f"Found {len(fd_links)} FD links, clicking on last one (index {last_link_index})")

            # Re-find the element right before clicking to avoid stale reference
            # Use a retry mechanism with element re-location
            clicked = False
            for attempt in range(3):
                try:
                    # Re-fetch the links to avoid stale reference
                    fd_links_fresh = driver.find_elements(By.CSS_SELECTOR, "#ctl00_ContentPlaceHolder1_FDList_ctl00 tbody tr td:first-child a")
                    if not fd_links_fresh:
                        fd_links_fresh = driver.find_elements(By.CSS_SELECTOR, "table.Vardena_b_8_n tbody tr td:first-child a")

                    if not fd_links_fresh or last_link_index >= len(fd_links_fresh):
                        logger.warning(f"Link count changed or link not found on attempt {attempt + 1}")
                        time.sleep(1)
                        continue

                    last_fd_link = fd_links_fresh[last_link_index]

                    # Try normal click first
                    try:
                        last_fd_link.click()
                        clicked = True
                        logger.info(f"Clicked on last FD number link (attempt {attempt + 1})")
                        break
                    except Exception as click_e:
                        # Fallback to JavaScript click
                        logger.info(f"Normal click failed on attempt {attempt + 1}, trying JS click: {click_e}")
                        driver.execute_script("arguments[0].click();", last_fd_link)
                        clicked = True
                        logger.info(f"JS clicked on last FD number link (attempt {attempt + 1})")
                        break

                except Exception as e:
                    logger.warning(f"Attempt {attempt + 1} failed to click FD link: {e}")
                    if attempt < 2:  # If not the last attempt
                        time.sleep(1)
                        continue
                    else:
                        raise

            if not clicked:
                logger.warning(f"Failed to click FD link after 3 attempts for contract {contract_number}")
                return None

            # Wait for page to load
            time.sleep(4)

        except Exception as e:
            logger.warning(f"Failed to click FD link for contract {contract_number}: {e}")
            import traceback
            logger.warning(traceback.format_exc())
            return None
        
        # Step 4.7: Extract FD amount
        try:
            amount_input = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtAmount"))
            )
            fd_amount = amount_input.get_attribute("value").strip()
            logger.info(f"Extracted FD amount: {fd_amount} for contract {contract_number}")
            
            return {
                "contract_number": contract_number,
                "fd_amount": fd_amount,
                "fd_number": fd_number
            }
            
        except TimeoutException:
            logger.warning(f"FD amount input not found for contract {contract_number}")
            return None
            
    except Exception as e:
        logger.error(f"Failed to extract FD amount for contract {contract_number}: {e}")
        return None

def run_scienter_valuation_bot(contract_numbers=None):
    """Main function to run the Scienter valuation bot"""
    driver = None
    try:
        logger.info("Starting Scienter Valuation Bot...")
        
        # Load contract numbers
        if contract_numbers is None:
            contract_numbers = load_contract_numbers_from_excel()
        
        logger.info(f"Processing {len(contract_numbers)} contract numbers")
        
        # Setup driver
        driver = setup_driver()
        
        # Load credentials
        username, password = load_credentials()
        
        # Login to Scienter
        if not login_to_scienter(driver, username, password):
            logger.error("Failed to login to Scienter")
            return {}
        
        # Navigate to FD Query page through menu
        if not navigate_to_fd_query(driver):
            logger.error("Failed to navigate to FD Query page")
            return {}
        
        # Process each contract
        results = {}
        for i, contract_number in enumerate(contract_numbers, 1):
            logger.info(f"Processing contract {i}/{len(contract_numbers)}: {contract_number}")
            
            try:
                result = extract_fd_amount_for_contract(driver, contract_number)
                if result:
                    results[contract_number] = result["fd_amount"]
                    logger.info(f"✓ Successfully extracted FD amount for {contract_number}: {result['fd_amount']}")
                else:
                    results[contract_number] = "-"
                    logger.warning(f"✗ Failed to extract FD amount for {contract_number}")
                
                # Small delay between contracts
                time.sleep(2)
                
            except Exception as e:
                logger.error(f"Error processing contract {contract_number}: {e}")
                results[contract_number] = "-"
        
        # Display results
        print("\n" + "="*80)
        print("SCIENTER VALUATION RESULTS")
        print("="*80)
        for contract, amount in results.items():
            print(f"{contract}: {amount}")
        print("="*80)
        
        logger.info(f"Bot completed! Processed {len(contract_numbers)} contracts")
        return results
        
    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        return {}
        
    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

def main():
    """Main entry point"""
    import sys
    
    # Support passing contract numbers via CLI
    cli_contracts = [arg.strip() for arg in sys.argv[1:] if arg.strip()]
    if cli_contracts:
        run_scienter_valuation_bot(cli_contracts)
    else:
        run_scienter_valuation_bot()

if __name__ == "__main__":
    main()