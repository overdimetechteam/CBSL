import os
import time
import logging
import glob
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
import openpyxl

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

def setup_driver(download_dir=None):
    """Setup Chrome driver with appropriate options

    Args:
        download_dir: Optional directory path for downloads
    """
    try:
        # Kill any existing Chrome processes that might interfere
        import subprocess
        import time

        try:
            subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'],
                         capture_output=True, timeout=10)
            subprocess.run(['taskkill', '/F', '/IM', 'chromedriver.exe'],
                         capture_output=True, timeout=10)
            time.sleep(2)  # Wait for processes to fully terminate
            logger.info("Closed existing Chrome/ChromeDriver processes")
        except Exception as e:
            logger.warning(f"Could not kill existing processes: {e}")

        options = Options()
        # options.add_argument("--headless=new")  # Commented out to see browser
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-blink-features=AutomationControlled")

        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        # Configure download directory if provided
        if download_dir:
            prefs = {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            options.add_experimental_option("prefs", prefs)
            logger.info(f"Download directory configured: {download_dir}")

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

def click_fixed_deposit(driver):
    """Click on Fixed Deposit button only"""
    try:
        logger.info("Clicking Fixed Deposit button...")

        # Wait for the page to load after login
        time.sleep(2)

        # Click on "Fixed Deposit" menu item
        logger.info("Looking for Fixed Deposit button...")
        fixed_deposit_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "h2.hovicon.effect-3.sub-a"))
        )
        fixed_deposit_menu.click()
        logger.info("Clicked on Fixed Deposit button")
        time.sleep(3)

        return True

    except Exception as e:
        logger.error(f"Failed to click Fixed Deposit button: {e}")
        return False

def calculate_sunday_dates(input_date):
    """Calculate 8-day date range ending with input_date

    Args:
        input_date: datetime object

    Returns:
        list of datetime objects for 8 consecutive dates ending on input_date
        Example: if input_date is 1/9/2026, returns [1/2/2026, 1/3/2026, ..., 1/9/2026]
    """
    from datetime import timedelta

    # Calculate the start date (7 days before input_date)
    start_date = input_date - timedelta(days=7)

    # Generate list of 8 dates from start_date to input_date (inclusive)
    date_list = [start_date + timedelta(days=i) for i in range(8)]

    logger.info(f"Date range: {start_date.strftime('%d/%m/%Y')} to {input_date.strftime('%d/%m/%Y')}")

    return date_list

def navigate_to_deposit_base(driver):
    """Navigate to Deposit Base page through REPORTS > FD QUERY INFO > DEPOSIT BASE"""
    try:
        logger.info("Navigating to Deposit Base page...")

        # Wait for page to load
        time.sleep(3)

        # Step 1: Click on "REPORTS" menu to expand it
        logger.info("Looking for REPORTS menu...")
        reports_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'REPORTS')]"))
        )
        reports_menu.click()
        logger.info("Clicked on REPORTS menu")
        time.sleep(2)

        # Step 2: Click on "FD QUERY INFO" submenu
        logger.info("Looking for FD QUERY INFO submenu...")
        fd_query_info = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'FD QUERY INFO')]"))
        )
        fd_query_info.click()
        logger.info("Clicked on FD QUERY INFO submenu")
        time.sleep(2)

        # Step 3: Click on "DEPOSIT BASE" link
        logger.info("Looking for DEPOSIT BASE link...")
        deposit_base = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'DEPOSIT BASE')]"))
        )
        deposit_base.click()
        logger.info("Clicked on DEPOSIT BASE link")
        time.sleep(3)

        # Verify navigation
        current_url = driver.current_url
        logger.info(f"Current URL after navigation: {current_url}")

        if "NewDepositBase" in current_url or "depositbase" in current_url.lower():
            logger.info("Successfully navigated to Deposit Base page")
            return True
        else:
            logger.warning(f"May not be on Deposit Base page. Current URL: {current_url}")
            return True

    except Exception as e:
        logger.error(f"Failed to navigate to Deposit Base: {e}")
        return False
    
def collect_deposit_base_value(driver, date_obj, download_dir):
    """Collect deposit base value for a specific date by downloading Excel report

    Args:
        driver: Selenium WebDriver
        date_obj: datetime object
        download_dir: Directory where Excel file will be downloaded

    Returns:
        float: The deposit base value
    """
    try:
        date_str = date_obj.strftime("%d/%m/%Y")
        date_str_filename = date_obj.strftime("%d-%m-%Y")
        logger.info(f"Collecting deposit base for date: {date_str}")

        # Enter date
        date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtDate"))
        )
        date_input.clear()
        date_input.send_keys(date_str)
        logger.info(f"Entered date: {date_str}")
        time.sleep(1)

        # Click Generate button
        generate_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnGen"))
        )
        generate_btn.click()
        logger.info("Clicked Generate button")
        time.sleep(8)  # Wait for report generation

        # Select EXCELOPENXML from dropdown
        logger.info("Selecting Excel format from dropdown...")
        format_dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"))
        )
        select = Select(format_dropdown)
        select.select_by_value("EXCELOPENXML")
        logger.info("Selected EXCELOPENXML format")
        time.sleep(1)

        # Get list of existing files before download
        existing_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))

        # Click Export button
        export_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"))
        )
        export_btn.click()
        logger.info("Clicked Export button")

        # Wait for download to complete
        logger.info("Waiting for Excel file to download...")
        max_wait = 30
        downloaded_file = None
        for i in range(max_wait):
            time.sleep(1)
            current_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
            new_files = current_files - existing_files

            # Filter out temporary files
            new_files = [f for f in new_files if not f.endswith('.crdownload') and not f.endswith('.tmp')]

            if new_files:
                downloaded_file = new_files[0]
                logger.info(f"Download complete: {downloaded_file}")
                break

        if not downloaded_file:
            logger.error("Download timeout - no new Excel file found")
            return 0.0

        # Rename file
        new_filename = f"FDBaseNew_UserBranch {date_str_filename}.xlsx"
        new_filepath = os.path.join(download_dir, new_filename)

        # Remove existing file if it exists
        if os.path.exists(new_filepath):
            os.remove(new_filepath)

        os.rename(downloaded_file, new_filepath)
        logger.info(f"Renamed file to: {new_filename}")

        # Read Excel file and extract value
        logger.info("Reading Excel file to extract value...")
        df = pd.read_excel(new_filepath, header=5)  # Headers are in row 6 (index 5)

        if "Amount" not in df.columns:
            logger.error(f"'Amount' column not found. Available columns: {df.columns.tolist()}")
            return 0.0

        # Get last row value from Amount column
        last_row_value = df["Amount"].iloc[-1]

        # Clean and convert to float
        if pd.isna(last_row_value):
            logger.warning("Last row Amount value is NaN")
            return 0.0

        value_float = float(last_row_value)
        logger.info(f"Extracted deposit base value: {value_float}")

        return value_float

    except Exception as e:
        logger.error(f"Failed to collect deposit base value for {date_str}: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 0.0

def navigate_to_interest_payable(driver):
    """Navigate to Interest Payable As At Date page"""
    try:
        logger.info("Navigating to Interest Payable As At Date page...")
        time.sleep(2)

        # Click REPORTS menu
        reports_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'REPORTS')]"))
        )
        reports_menu.click()
        logger.info("Clicked on REPORTS menu")
        time.sleep(2)

        # Click INTEREST DETAILS submenu
        interest_details = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'INTEREST DETAILS')]"))
        )
        interest_details.click()
        logger.info("Clicked on INTEREST DETAILS submenu")
        time.sleep(2)

        # Click INTEREST PAYABLE AS AT DATE link
        interest_payable = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'INTEREST PAYABLE AS AT DATE')]"))
        )
        interest_payable.click()
        logger.info("Clicked on INTEREST PAYABLE AS AT DATE link")
        time.sleep(3)

        logger.info("Successfully navigated to Interest Payable page")
        return True

    except Exception as e:
        logger.error(f"Failed to navigate to Interest Payable: {e}")
        return False

def collect_interest_payable_value(driver, date_obj, download_dir):
    """Collect interest payable value for a specific date by downloading Excel report

    Args:
        driver: Selenium WebDriver
        date_obj: datetime object
        download_dir: Directory where Excel file will be downloaded

    Returns:
        float: The interest payable value
    """
    try:
        date_str = date_obj.strftime("%d/%m/%Y")
        date_str_filename = date_obj.strftime("%d-%m-%Y")
        logger.info(f"Collecting interest payable for date: {date_str}")

        # Enter date
        date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtAsAtDate"))
        )
        date_input.clear()
        date_input.send_keys(date_str)
        logger.info(f"Entered date: {date_str}")
        time.sleep(1)

        # Click Display button
        display_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnDisplay"))
        )
        display_btn.click()
        logger.info("Clicked Display button")
        time.sleep(8)  # Wait for report generation

        # Select EXCELOPENXML from dropdown
        logger.info("Selecting Excel format from dropdown...")
        format_dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"))
        )
        select = Select(format_dropdown)
        select.select_by_value("EXCELOPENXML")
        logger.info("Selected EXCELOPENXML format")
        time.sleep(1)

        # Get list of existing files before download
        existing_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))

        # Click Export button
        export_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"))
        )
        export_btn.click()
        logger.info("Clicked Export button")

        # Wait for download to complete
        logger.info("Waiting for Excel file to download...")
        max_wait = 30
        downloaded_file = None
        for _ in range(max_wait):
            time.sleep(1)
            current_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
            new_files = current_files - existing_files

            # Filter out temporary files
            new_files = [f for f in new_files if not f.endswith('.crdownload') and not f.endswith('.tmp')]

            if new_files:
                downloaded_file = new_files[0]
                logger.info(f"Download complete: {downloaded_file}")
                break

        if not downloaded_file:
            logger.error("Download timeout - no new Excel file found")
            return 0.0

        # Rename file
        new_filename = f"InterestPayableAsAtDateReport {date_str_filename}.xlsx"
        new_filepath = os.path.join(download_dir, new_filename)

        # Remove existing file if it exists
        if os.path.exists(new_filepath):
            os.remove(new_filepath)

        os.rename(downloaded_file, new_filepath)
        logger.info(f"Renamed file to: {new_filename}")

        # Read Excel file and extract value
        logger.info("Reading Excel file to extract value...")
        df = pd.read_excel(new_filepath, header=6)  # Headers are in row 7 (index 6)

        if "Int\nPayable" not in df.columns:
            logger.error(f"'Int\nPayable' column not found. Available columns: {df.columns.tolist()}")
            return 0.0

        # Get last row value from Int\nPayable column
        last_row_value = df["Int\nPayable"].iloc[-1]

        # Clean and convert to float
        if pd.isna(last_row_value):
            logger.warning("Last row Int\nPayable value is NaN")
            return 0.0

        value_float = float(last_row_value)
        logger.info(f"Extracted interest payable value: {value_float}")

        return value_float

    except Exception as e:
        logger.error(f"Failed to collect interest payable value for {date_str}: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 0.0

def collect_rec_data(input_date):
    """Main function to collect Rec sheet data (deposit base and interest payable)

    Args:
        input_date: datetime object

    Returns:
        dict: Dictionary with 'dates', 'deposit_base', and 'interest_payable' lists
              Returns None if collection fails
    """
    driver = None
    try:
        logger.info("Starting Rec data collection bot...")
        logger.info(f"Input date: {input_date.strftime('%d/%m/%Y')}")

        # Calculate the 8-day date range (7 days back to input_date)
        date_list = calculate_sunday_dates(input_date)
        logger.info(f"Collecting data for {len(date_list)} dates")

        # Find the date folder for downloads
        logger.info("Looking for NBD-WF-18-DM working directory...")
        current_dir = os.getcwd()

        # Check if we're already in a NBD-WF-18-DM date folder
        # Pattern: C:\CBSL\Script\working\NBD-WF-18-DM\30-06-2025
        if "NBD-WF-18-DM" in current_dir and os.path.exists(os.path.join(current_dir, "inputs")):
            # We're already in the date folder (called from NBD_WF_18_DM.py)
            date_folder = current_dir
            logger.info(f"Using current directory as date folder: {date_folder}")
        else:
            # We're running standalone from script root, need to find the working directory
            working_dir = os.path.join(current_dir, "working")

            if not os.path.exists(working_dir):
                logger.error(f"Working directory not found: {working_dir}")
                return None

            nbd_wf_dirs = [d for d in os.listdir(working_dir) if d.startswith("NBD-WF-18-DM") and os.path.isdir(os.path.join(working_dir, d))]

            if not nbd_wf_dirs:
                logger.error("No NBD-WF-18-DM directory found in working directory")
                return None

            nbd_wf_dir = os.path.join(working_dir, nbd_wf_dirs[0])
            logger.info(f"Found NBD-WF-18-DM directory: {nbd_wf_dir}")

            # Find the date folder (only folder inside NBD-WF-18-DM)
            date_folders = [d for d in os.listdir(nbd_wf_dir) if os.path.isdir(os.path.join(nbd_wf_dir, d))]

            if not date_folders:
                logger.error(f"No date folder found in {nbd_wf_dir}")
                return None

            date_folder = os.path.join(nbd_wf_dir, date_folders[0])
            logger.info(f"Found date folder: {date_folder}")

        # Use existing inputs subdirectory
        download_dir = os.path.join(date_folder, "inputs")

        # Clear existing Excel files in inputs folder
        if os.path.exists(download_dir):
            for file in glob.glob(os.path.join(download_dir, "*.xlsx")):
                try:
                    os.remove(file)
                    logger.info(f"Cleared existing file: {os.path.basename(file)}")
                except Exception as e:
                    logger.warning(f"Failed to remove {os.path.basename(file)}: {e}")

        logger.info(f"Using download directory: {download_dir}")

        # Setup driver with download directory
        driver = setup_driver(download_dir)

        # Load credentials
        username, password = load_credentials()

        # Login to Scienter
        if not login_to_scienter(driver, username, password):
            logger.error("Failed to login to Scienter")
            return None

        # Click Fixed Deposit button
        if not click_fixed_deposit(driver):
            logger.error("Failed to click Fixed Deposit button")
            return None

        logger.info("Successfully logged in and clicked Fixed Deposit button")

        # ======================================
        # PART 1: Collect Deposit Base Data
        # ======================================
        logger.info("=" * 60)
        logger.info("PART 1: Collecting Deposit Base Data")
        logger.info("=" * 60)

        # Navigate to Deposit Base page
        if not navigate_to_deposit_base(driver):
            logger.error("Failed to navigate to Deposit Base page")
            return None

        # Collect deposit base values for each date
        deposit_base_values = []
        for i, date_obj in enumerate(date_list, 1):
            logger.info(f"Collecting deposit base {i}/{len(date_list)}: {date_obj.strftime('%d/%m/%Y')}")
            value = collect_deposit_base_value(driver, date_obj, download_dir)
            deposit_base_values.append(value)

        logger.info(f"Collected {len(deposit_base_values)} deposit base values")

        # ======================================
        # PART 2: Collect Interest Payable Data
        # ======================================
        logger.info("=" * 60)
        logger.info("PART 2: Collecting Interest Payable Data")
        logger.info("=" * 60)

        # Navigate to Interest Payable page
        if not navigate_to_interest_payable(driver):
            logger.error("Failed to navigate to Interest Payable page")
            return None

        # Collect interest payable values for each date
        interest_payable_values = []
        for i, date_obj in enumerate(date_list, 1):
            logger.info(f"Collecting interest payable {i}/{len(date_list)}: {date_obj.strftime('%d/%m/%Y')}")
            value = collect_interest_payable_value(driver, date_obj, download_dir)
            interest_payable_values.append(value)

        logger.info(f"Collected {len(interest_payable_values)} interest payable values")

        # ======================================
        # Return collected data
        # ======================================
        result = {
            'dates': date_list,
            'deposit_base': deposit_base_values,
            'interest_payable': interest_payable_values
        }

        logger.info("=" * 60)
        logger.info("Data Collection Summary:")
        logger.info(f"  Dates: {len(date_list)}")
        logger.info(f"  Deposit Base values: {len(deposit_base_values)}")
        logger.info(f"  Interest Payable values: {len(interest_payable_values)}")
        logger.info("=" * 60)

        return result

    except Exception as e:
        logger.error(f"Rec data collection failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return None

    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

def main():
    """Main entry point for testing"""
    import sys
    from datetime import datetime

    # For testing: use current date or date from command line
    if len(sys.argv) > 1:
        try:
            input_date = datetime.strptime(sys.argv[1], "%d/%m/%Y")
        except ValueError:
            print("Usage: python DM_Scienter_rec.py DD/MM/YYYY")
            sys.exit(1)
    else:
        input_date = datetime.now()

    result = collect_rec_data(input_date)

    if result:
        print("\n" + "=" * 60)
        print("Collection successful!")
        print("=" * 60)
        for i, date_obj in enumerate(result['dates']):
            print(f"{date_obj.strftime('%d/%m/%Y')}: Deposit={result['deposit_base'][i]:,.2f}, Interest={result['interest_payable'][i]:,.2f}")
    else:
        print("Collection failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()