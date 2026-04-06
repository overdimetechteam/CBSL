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

def setup_driver(download_folder=None):
    """Setup Chrome driver with appropriate options"""
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

        # Set download preferences if download_folder specified
        if download_folder:
            os.makedirs(download_folder, exist_ok=True)
            download_folder_abs = os.path.abspath(download_folder)
            prefs = {
                "download.default_directory": download_folder_abs,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            options.add_experimental_option("prefs", prefs)
            logger.info(f"Download folder set to: {download_folder_abs}")

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
    """
    Calculate 8-day date range ending with input_date
    Args:
        input_date (datetime): Input date
    Returns:
        tuple: (from_date, to_date) as datetime objects
        Example: if input_date is 1/9/2026, returns (1/2/2026, 1/9/2026)
    """
    from datetime import timedelta

    # to_date is the input_date
    to_date = input_date

    # from_date is 7 days before input_date (giving us 8 days total)
    from_date = input_date - timedelta(days=7)

    return from_date, to_date

def navigate_to_daily_deposit_bal(driver, input_date, download_folder):
    """
    Navigate to Central Bank Movement of Daily Deposit Balances and download CSV reports
    Args:
        driver: Selenium webdriver
        input_date (datetime): Input date for calculation
        download_folder (str): Folder to download files to
    """
    try:
        logger.info("Navigating to Central Bank Movement of Daily Deposit Balances...")

        # Wait for page to load
        time.sleep(3)

        # Step 1: Click on "REPORTS" menu to expand it
        logger.info("Looking for REPORTS menu...")
        reports_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.rmLink.rmRootLink"))
        )

        # Check if it contains "REPORTS" text
        reports_text = reports_menu.find_element(By.CSS_SELECTOR, "span.rmText.rmExpandDown")
        if "REPORTS" in reports_text.text.upper():
            reports_menu.click()
            logger.info("Clicked on REPORTS menu")
            time.sleep(2)
        else:
            # Try finding by text content
            reports_menu = driver.find_element(By.XPATH, "//span[contains(text(), 'REPORTS')]/parent::a")
            reports_menu.click()
            logger.info("Clicked on REPORTS menu (via XPath)")
            time.sleep(2)

        # Step 2: Click on "CENTRAL BANK REPORT_NEW" submenu
        logger.info("Looking for CENTRAL BANK REPORT_NEW submenu...")
        central_bank_menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'CENTRAL BANK REPORT_NEW')]/parent::a"))
        )
        central_bank_menu.click()
        logger.info("Clicked on CENTRAL BANK REPORT_NEW submenu")
        time.sleep(2)

        # Step 3: Click on "CENTRAL BANK MOVEMENT OF DAILY DEPOSIT BALANCES" link
        logger.info("Looking for CENTRAL BANK MOVEMENT OF DAILY DEPOSIT BALANCES link...")
        daily_deposit_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'CENTRAL BANK MOVEMENT OF DAILY DEPOSIT BALANCES')]/parent::a"))
        )
        daily_deposit_link.click()
        logger.info("Clicked on CENTRAL BANK MOVEMENT OF DAILY DEPOSIT BALANCES link")
        time.sleep(3)

        # Calculate from_date and to_date
        from_date, to_date = calculate_sunday_dates(input_date)
        # Scienter expects YYYY/MM/DD format
        from_date_str = from_date.strftime("%Y/%m/%d")
        to_date_str = to_date.strftime("%Y/%m/%d")

        logger.info(f"Date range: {from_date_str} to {to_date_str}")
        print(f"From Date: {from_date_str}")
        print(f"To Date: {to_date_str}")

        # Enter from_date
        logger.info("Entering from date...")
        from_date_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtFromDate"))
        )
        from_date_input.clear()
        from_date_input.send_keys(from_date_str)
        logger.info(f"Entered from date: {from_date_str}")

        # Enter to_date
        logger.info("Entering to date...")
        to_date_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtToDate"))
        )
        to_date_input.clear()
        to_date_input.send_keys(to_date_str)
        logger.info(f"Entered to date: {to_date_str}")

        # Click CBSL Report button
        logger.info("Clicking CBSL Report button...")
        cbsl_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnDisplay"))
        )
        cbsl_button.click()
        logger.info("Clicked CBSL Report button")
        time.sleep(5)  # Wait for report to load

        # Clear existing CentralBankMovementOfDailyDepositBalances*.csv files
        logger.info("Clearing existing CentralBankMovementOfDailyDepositBalances files...")
        for filename in os.listdir(download_folder):
            if "CentralBankMovementOfDailyDepositBalances" in filename and filename.endswith('.csv'):
                file_path = os.path.join(download_folder, filename)
                os.remove(file_path)
                logger.info(f"Removed: {filename}")

        # Wait for report to fully load and ensure we're in main context
        time.sleep(3)
        driver.switch_to.default_content()
        logger.info("Ensured in main page context (not in iframe)")

        # Try multiple selectors for the format dropdown
        logger.info("Selecting CSV format...")
        format_dropdown = None
        dropdown_selectors = [
            (By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"),
            (By.CSS_SELECTOR, "select[title='Export Formats']"),  # Note: plural "Formats"
            (By.CSS_SELECTOR, "select[title='Export Format']"),
            (By.XPATH, "//select[contains(@title, 'Export')]"),
            (By.XPATH, "//select[contains(@id, 'ReportViewer')]"),
            (By.XPATH, "//select[contains(@id, 'ctl05')]"),
        ]

        for selector_type, selector_value in dropdown_selectors:
            try:
                format_dropdown = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((selector_type, selector_value))
                )
                logger.info(f"Found format dropdown with selector: {selector_type}={selector_value}")
                break
            except:
                continue

        if not format_dropdown:
            logger.warning("Could not find export format dropdown, trying to click export directly")
            # Switch back from iframe if we switched
            try:
                driver.switch_to.default_content()
            except:
                pass

            # Try to click export button directly without dropdown
            try:
                logger.info("Attempting to click Export button directly...")
                export_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
                )
                export_button.click()
                logger.info("Clicked Export button directly for CentralBankMovementOfDailyDepositBalances")
                time.sleep(10)  # Wait for download
            except Exception as e:
                logger.error(f"Failed to click export button directly: {e}")
                raise Exception("Could not find export format dropdown or export button")
        else:
            select = Select(format_dropdown)
            select.select_by_value("CSV")
            logger.info("Selected CSV format")
            time.sleep(1)

            # Switch back from iframe if we switched
            try:
                driver.switch_to.default_content()
            except:
                pass

            # Click Export button
            logger.info("Clicking Export button...")
            export_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
            )
            export_button.click()
            logger.info("Clicked Export button for CentralBankMovementOfDailyDepositBalances")
            time.sleep(10)  # Wait for download

        # Now download MovementOfDepositDetails.csv
        logger.info("Proceeding to download MovementOfDepositDetails.csv...")

        # Clear existing MovementOfDepositDetails*.csv files
        logger.info("Clearing existing MovementOfDepositDetails files...")
        for filename in os.listdir(download_folder):
            if "MovementOfDepositDetails" in filename and filename.endswith('.csv'):
                file_path = os.path.join(download_folder, filename)
                os.remove(file_path)
                logger.info(f"Removed: {filename}")

        # Click Details Report button
        logger.info("Clicking Details Report button...")
        details_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnDisplayDetail"))
        )
        details_button.click()
        logger.info("Clicked Details Report button")
        time.sleep(5)  # Wait for report to load

        # Wait for report to fully load and ensure we're in main context
        time.sleep(3)
        driver.switch_to.default_content()
        logger.info("Ensured in main page context for Details Report (not in iframe)")

        # Try multiple selectors for the format dropdown
        logger.info("Selecting CSV format for Details Report...")
        format_dropdown2 = None
        for selector_type, selector_value in dropdown_selectors:
            try:
                format_dropdown2 = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((selector_type, selector_value))
                )
                logger.info(f"Found format dropdown with selector: {selector_type}={selector_value}")
                break
            except:
                continue

        if not format_dropdown2:
            logger.warning("Could not find export format dropdown for Details Report, trying to click export directly")
            # Switch back from iframe if we switched
            try:
                driver.switch_to.default_content()
            except:
                pass

            # Try to click export button directly without dropdown
            try:
                logger.info("Attempting to click Export button directly for Details Report...")
                export_button2 = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
                )
                export_button2.click()
                logger.info("Clicked Export button directly for MovementOfDepositDetails")
                time.sleep(10)  # Wait for download
            except Exception as e:
                logger.error(f"Failed to click export button directly for Details Report: {e}")
                raise Exception("Could not find export format dropdown or export button for Details Report")
        else:
            select2 = Select(format_dropdown2)
            select2.select_by_value("CSV")
            logger.info("Selected CSV format for Details Report")
            time.sleep(1)

            # Switch back from iframe if we switched
            try:
                driver.switch_to.default_content()
            except:
                pass

            # Click Export button for Details Report
            logger.info("Clicking Export button for Details Report...")
            export_button2 = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
            )
            export_button2.click()
            logger.info("Clicked Export button for MovementOfDepositDetails")
            time.sleep(10)  # Wait for download

        logger.info("Successfully downloaded both CSV reports")

        # Additional downloads: FDBaseNew_UserBranch.csv and InterestPayableAsAtDateReport.csv
        logger.info("Proceeding to download FDBaseNew_UserBranch.csv...")

        # Clear existing FDBaseNew_UserBranch*.csv files
        logger.info("Clearing existing FDBaseNew_UserBranch files...")
        for filename in os.listdir(download_folder):
            if "FDBaseNew_UserBranch" in filename and filename.endswith('.csv'):
                file_path = os.path.join(download_folder, filename)
                os.remove(file_path)
                logger.info(f"Removed: {filename}")

        # Navigate to REPORTS -> FD QUERY INFO -> DEPOSIT BASE
        logger.info("Navigating to FD QUERY INFO -> DEPOSIT BASE...")
        time.sleep(2)

        # Click REPORTS menu
        logger.info("Looking for REPORTS menu...")
        reports_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'REPORTS')]"))
        )
        reports_menu.click()
        logger.info("Clicked on REPORTS menu")
        time.sleep(2)

        # Click FD QUERY INFO submenu
        logger.info("Looking for FD QUERY INFO submenu...")
        fd_query_submenu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'FD QUERY INFO')]"))
        )
        fd_query_submenu.click()
        logger.info("Clicked on FD QUERY INFO submenu")
        time.sleep(2)

        # Click DEPOSIT BASE link
        logger.info("Looking for DEPOSIT BASE link...")
        deposit_base_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'DEPOSIT BASE')]"))
        )
        deposit_base_link.click()
        logger.info("Clicked on DEPOSIT BASE link")
        time.sleep(3)

        # Enter input_date in DD/MM/YYYY format
        input_date_str_ddmmyyyy = input_date.strftime("%d/%m/%Y")
        logger.info(f"Entering date: {input_date_str_ddmmyyyy}")
        date_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtDate"))
        )
        date_field.clear()
        date_field.send_keys(input_date_str_ddmmyyyy)
        logger.info(f"Entered date: {input_date_str_ddmmyyyy}")
        time.sleep(1)

        # Click Generate button
        logger.info("Clicking Generate button...")
        generate_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnGen"))
        )
        generate_button.click()
        logger.info("Clicked Generate button")
        time.sleep(5)  # Wait for report to load

        # Ensure in main page context
        driver.switch_to.default_content()
        logger.info("Ensured in main page context for FDBaseNew_UserBranch")

        # Select CSV format and export
        logger.info("Selecting CSV format for FDBaseNew_UserBranch...")
        fdbase_dropdown = None
        for selector_type, selector_value in dropdown_selectors:
            try:
                fdbase_dropdown = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((selector_type, selector_value))
                )
                logger.info(f"Found format dropdown with selector: {selector_type}={selector_value}")
                break
            except:
                continue

        if not fdbase_dropdown:
            logger.warning("Could not find export format dropdown for FDBaseNew_UserBranch, trying to click export directly")
            try:
                driver.switch_to.default_content()
            except:
                pass

            try:
                logger.info("Attempting to click Export button directly for FDBaseNew_UserBranch...")
                export_button_fdbase = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
                )
                export_button_fdbase.click()
                logger.info("Clicked Export button directly for FDBaseNew_UserBranch")
                time.sleep(10)
            except Exception as e:
                logger.error(f"Failed to click export button directly for FDBaseNew_UserBranch: {e}")
                raise Exception("Could not find export format dropdown or export button for FDBaseNew_UserBranch")
        else:
            select_fdbase = Select(fdbase_dropdown)
            select_fdbase.select_by_value("CSV")
            logger.info("Selected CSV format for FDBaseNew_UserBranch")
            time.sleep(1)

            try:
                driver.switch_to.default_content()
            except:
                pass

            # Click Export button
            logger.info("Clicking Export button for FDBaseNew_UserBranch...")
            export_button_fdbase = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
            )
            export_button_fdbase.click()
            logger.info("Clicked Export button for FDBaseNew_UserBranch")
            time.sleep(10)  # Wait for download

        logger.info("Successfully downloaded FDBaseNew_UserBranch.csv")

        # Download InterestPayableAsAtDateReport.csv
        logger.info("Proceeding to download InterestPayableAsAtDateReport.csv...")

        # Clear existing InterestPayableAsAtDateReport*.csv files
        logger.info("Clearing existing InterestPayableAsAtDateReport files...")
        for filename in os.listdir(download_folder):
            if "InterestPayableAsAtDateReport" in filename and filename.endswith('.csv'):
                file_path = os.path.join(download_folder, filename)
                os.remove(file_path)
                logger.info(f"Removed: {filename}")

        # Navigate to REPORTS -> INTEREST DETAILS -> INTEREST PAYABLE AS AT DATE
        logger.info("Navigating to INTEREST DETAILS -> INTEREST PAYABLE AS AT DATE...")
        time.sleep(2)

        # Click REPORTS menu
        logger.info("Looking for REPORTS menu...")
        reports_menu2 = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'REPORTS')]"))
        )
        reports_menu2.click()
        logger.info("Clicked on REPORTS menu")
        time.sleep(3)  # Increased wait for menu to expand

        # Click INTEREST DETAILS submenu
        logger.info("Looking for INTEREST DETAILS submenu...")
        try:
            interest_details_submenu = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'INTEREST DETAILS')]"))
            )
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView(true);", interest_details_submenu)
            time.sleep(1)
            interest_details_submenu.click()
            logger.info("Clicked on INTEREST DETAILS submenu")
        except TimeoutException:
            logger.error("INTEREST DETAILS submenu not found, trying to click REPORTS again...")
            # Sometimes menu needs to be clicked again
            reports_menu2 = driver.find_element(By.XPATH, "//span[contains(text(), 'REPORTS')]")
            reports_menu2.click()
            time.sleep(3)
            interest_details_submenu = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'INTEREST DETAILS')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", interest_details_submenu)
            time.sleep(1)
            interest_details_submenu.click()
            logger.info("Clicked on INTEREST DETAILS submenu (second attempt)")
        time.sleep(2)

        # Click INTEREST PAYABLE AS AT DATE link
        logger.info("Looking for INTEREST PAYABLE AS AT DATE link...")
        interest_payable_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'INTEREST PAYABLE AS AT DATE')]"))
        )
        interest_payable_link.click()
        logger.info("Clicked on INTEREST PAYABLE AS AT DATE link")
        time.sleep(3)

        # Enter input_date in DD/MM/YYYY format
        logger.info(f"Entering date: {input_date_str_ddmmyyyy}")
        asatdate_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtAsAtDate"))
        )
        asatdate_field.clear()
        asatdate_field.send_keys(input_date_str_ddmmyyyy)
        logger.info(f"Entered date: {input_date_str_ddmmyyyy}")
        time.sleep(1)

        # Click Display link
        logger.info("Clicking Display link...")
        display_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnDisplay"))
        )
        display_link.click()
        logger.info("Clicked Display link")
        time.sleep(5)  # Wait for report to load

        # Ensure in main page context
        driver.switch_to.default_content()
        logger.info("Ensured in main page context for InterestPayableAsAtDateReport")

        # Select CSV format and export
        logger.info("Selecting CSV format for InterestPayableAsAtDateReport...")
        interest_dropdown = None
        for selector_type, selector_value in dropdown_selectors:
            try:
                interest_dropdown = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((selector_type, selector_value))
                )
                logger.info(f"Found format dropdown with selector: {selector_type}={selector_value}")
                break
            except:
                continue

        if not interest_dropdown:
            logger.warning("Could not find export format dropdown for InterestPayableAsAtDateReport, trying to click export directly")
            try:
                driver.switch_to.default_content()
            except:
                pass

            try:
                logger.info("Attempting to click Export button directly for InterestPayableAsAtDateReport...")
                export_button_interest = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
                )
                export_button_interest.click()
                logger.info("Clicked Export button directly for InterestPayableAsAtDateReport")
                time.sleep(10)
            except Exception as e:
                logger.error(f"Failed to click export button directly for InterestPayableAsAtDateReport: {e}")
                raise Exception("Could not find export format dropdown or export button for InterestPayableAsAtDateReport")
        else:
            select_interest = Select(interest_dropdown)
            select_interest.select_by_value("CSV")
            logger.info("Selected CSV format for InterestPayableAsAtDateReport")
            time.sleep(1)

            try:
                driver.switch_to.default_content()
            except:
                pass

            # Click Export button
            logger.info("Clicking Export button for InterestPayableAsAtDateReport...")
            export_button_interest = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export']"))
            )
            export_button_interest.click()
            logger.info("Clicked Export button for InterestPayableAsAtDateReport")
            time.sleep(10)  # Wait for download

        logger.info("Successfully downloaded InterestPayableAsAtDateReport.csv")
        logger.info("All four CSV reports downloaded successfully")
        return True

    except Exception as e:
        logger.error(f"Failed to download reports: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return False

def run_DM_scienter_portfolio(contract_numbers=None):
    """Main function to run the DM_scienter_portfolio_Q3 bot"""
    driver = None
    try:
        logger.info("Starting DM_scienter_portfolio_Q3 Bot...")

        # Setup driver
        driver = setup_driver()

        # Load credentials
        username, password = load_credentials()

        # Login to Scienter
        if not login_to_scienter(driver, username, password):
            logger.error("Failed to login to Scienter")
            return False

        # Click Fixed Deposit button
        if not click_fixed_deposit(driver):
            logger.error("Failed to click Fixed Deposit button")
            return False

        logger.info("Successfully logged in and clicked Fixed Deposit button")

        # Navigate and download reports (requires input_date and download_folder parameters)
        # This function needs to be called with those parameters
        logger.error("This function is deprecated. Use run_dm_daily_deposit_download() instead")
        return False

        logger.info("Successfully navigated to Daily Deposit Balance page")
        logger.info("Bot execution completed successfully")

        # Keep browser open for manual interaction (optional - remove if not needed)
        logger.info("Browser will remain open for manual interaction. Press Ctrl+C to close.")
        input("Press Enter to close the browser...")

        return True

    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        return False

    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

def run_dm_daily_deposit_download(input_date, download_folder):
    """
    Main function to download DM Daily Deposit CSV reports
    Args:
        input_date (datetime): Input date for calculation
        download_folder (str): Folder to download files to
    Returns:
        bool: True if successful, False otherwise
    """
    driver = None
    try:
        logger.info("Starting DM Daily Deposit Balance Download...")
        logger.info(f"Input date: {input_date.strftime('%Y-%m-%d')}")
        logger.info(f"Download folder: {download_folder}")

        # Setup driver with download folder
        driver = setup_driver(download_folder=download_folder)

        # Load credentials
        username, password = load_credentials()

        # Login to Scienter
        if not login_to_scienter(driver, username, password):
            logger.error("Failed to login to Scienter")
            return False

        # Click Fixed Deposit button
        if not click_fixed_deposit(driver):
            logger.error("Failed to click Fixed Deposit button")
            return False

        logger.info("Successfully logged in and clicked Fixed Deposit button")

        # Navigate and download CSV reports
        if not navigate_to_daily_deposit_bal(driver, input_date, download_folder):
            logger.error("Failed to download Daily Deposit Balance reports")
            return False

        logger.info("Successfully downloaded Daily Deposit Balance reports")
        logger.info("Bot execution completed successfully")

        return True

    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return False

    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

def main():
    """Main entry point for command line usage"""
    import sys
    from datetime import datetime

    # Parse command line arguments
    if len(sys.argv) > 1:
        try:
            input_date = datetime.strptime(sys.argv[1], "%m/%d/%Y")
        except:
            print("Usage: python DM_Scienter_Daily_Deposit_Bal.py [MM/DD/YYYY] [download_folder]")
            return
    else:
        input_date = datetime.now()

    if len(sys.argv) > 2:
        download_folder = sys.argv[2]
    else:
        download_folder = os.path.join(os.getcwd(), "downloads")

    run_dm_daily_deposit_download(input_date, download_folder)

if __name__ == "__main__":
    main()