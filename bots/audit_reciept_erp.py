#!/usr/bin/env python3
"""
Report Download Bot for IA and C1C2 Reports
Uses Selenium to automate the download of Po Listing - Internal reports from ERP system
"""

import os
import re
import time
import logging
import threading
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from cryptography.fernet import Fernet
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    TimeoutException,
)
from datetime import datetime, timedelta

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
BASE_URL = "https://erp.assetline.lk/"
LOGS_DIR = "logs"
# Use the user's default Downloads folder
DOWNLOAD_DIR = str(Path.home() / "Downloads")

# Create necessary directories
os.makedirs(LOGS_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Security settings
SECURE_LOGGING = True

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, 'report_download_bot.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def click_with_retries(driver, by, value, description="element", max_attempts=3):
    """Robust click: wait, scroll into view, try normal click, then ActionChains, then JS click."""
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            elem = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((by, value))
            )
            WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((by, value))
            )
            # Scroll into center view
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
            # Try standard click
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((by, value))
            )
            elem.click()
            return True
        except (ElementClickInterceptedException, ElementNotInteractableException, StaleElementReferenceException, TimeoutException) as e1:
            last_error = e1
            try:
                # Re-find and attempt ActionChains click
                elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((by, value))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                ActionChains(driver).move_to_element(elem).click().perform()
                return True
            except Exception as e2:
                last_error = e2
                try:
                    # Final fallback: JS click
                    elem = driver.find_element(by, value)
                    driver.execute_script("arguments[0].click();", elem)
                    return True
                except Exception as e3:
                    last_error = e3
                    # Ensure page is ready before next retry
                    try:
                        WebDriverWait(driver, 10).until(lambda d: d.execute_script("return document.readyState") == "complete")
                    except Exception:
                        pass
    logger.error(f"Failed to click {description}: {last_error}")
    return False

def load_credentials():
    """Load and decrypt credentials from environment variables"""
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
        return username, password
        
    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        raise

def setup_chrome_driver(download_dir):
    """Setup Chrome driver with proper options for headless operation and file downloads"""
    try:
        logger.info("Setting up Chrome driver...")
        
        # Chrome options
        options = Options()
        # options.add_argument("--headless=new")
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
        
        # Configure download preferences
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.notifications": 2,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        options.add_experimental_option("prefs", prefs)
        
        # Additional DevTools settings for downloads
        options.add_argument("--disable-web-security")
        options.add_argument("--allow-running-insecure-content")
        options.add_argument("--disable-features=VizDisplayCompositor")
        
        # SSL and security settings to handle handshake issues
        options.add_argument("--ignore-ssl-errors")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--ignore-certificate-errors-spki-list")
        options.add_argument("--ignore-ssl-errors-list")
        options.add_argument("--allow-running-insecure-content")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-plugins")
        options.add_argument("--disable-web-security")
        options.add_argument("--allow-running-insecure-content")
        options.add_argument("--disable-features=VizDisplayCompositor")
        
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

def login_to_erp(driver) -> bool:
    """Login to ERP system"""
    try:
        logger.info("Opening ERP website...")
        driver.get(BASE_URL)

        wait = WebDriverWait(driver, 15)

        username, password = load_credentials()

        logger.info("Entering username...")
        username_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="strUserName"]')))
        username_field.clear()
        username_field.send_keys(username)

        logger.info("Entering password...")
        password_field = driver.find_element(By.XPATH, '//*[@id="strPassword"]')
        password_field.clear()
        password_field.send_keys(password)

        logger.info("Clicking sign in...")
        sign_in_button = driver.find_element(By.XPATH, '//*[@id="idSignIn"]')
        sign_in_button.click()

        time.sleep(3)
        logger.info("Login successful!")
        return True

    except Exception as e:
        logger.error(f"Error during login: {e}")
        return False


# def select_premises(driver) -> bool:
#     """Click the SELECT button to proceed past the company/branch selector"""
#     try:
#         wait = WebDriverWait(driver, 20)

#         time.sleep(3)

#         logger.info("Clicking Select button...")
#         select_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/form/div/button')))
#         select_button.click()

#         time.sleep(3)

#         logger.info("Refreshing the page...")
#         driver.refresh()

#         time.sleep(3)
#         logger.info("Premises selection completed!")
#         return True

#     except Exception as e:
#         logger.error(f"Error clicking Select button: {e}")
#         return False

def navigate_to_reports(driver):
    """Step 1: Navigate to REPORT page, search for Audit Receipt and click it."""
    try:
        wait = WebDriverWait(driver, 15)

        logger.info("Navigating to Reports page...")
        driver.get("https://erp.assetline.lk/Application/Home/REPORT")
        time.sleep(5)

        # Click the search input, type text, press Enter
        logger.info("Searching for 'audit rec'...")
        search_input = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.input-group input.form-control[placeholder='Search']")
        ))
        driver.execute_script("arguments[0].click(); arguments[0].focus();", search_input)
        search_input.clear()
        search_input.send_keys("audit rec")
        search_input.send_keys(webdriver.common.keys.Keys.ENTER)
        time.sleep(2)

        # Click "Audit Receipt - Internal" from results
        logger.info("Clicking Audit Receipt - Internal from search results...")
        result = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@id='divSearchScrollBdy']//a[normalize-space()='Audit Receipt - Internal']")
        ))
        driver.execute_script("arguments[0].click();", result)
        time.sleep(3)

        logger.info("Audit Receipt - Internal clicked — module loading...")
        return True

    except Exception as e:
        logger.error(f"Failed to navigate to Reports: {e}")
        return False


def click_inquiry(driver, reportDate, startDate=None):
    """Step 2: Wait for Inquiry button to appear and click it.

    Args:
        reportDate: End date string in DD/MM/YYYY format.
        startDate:  Optional start date string in DD/MM/YYYY format.
                    If omitted, start = reportDate - 6 days (weekly default).
    """
    try:
        wait = WebDriverWait(driver, 30)

        logger.info("Waiting for Inquiry button...")
        inquiry_btn = wait.until(EC.presence_of_element_located((By.ID, "btn_02013_Inquiry")))
        driver.execute_script("arguments[0].click();", inquiry_btn)
        logger.info("Inquiry button clicked")
        time.sleep(3)

        # Select BETWEEN from the RPTOperator dropdown in the RECIPT_DATE row
        wait.until(EC.presence_of_element_located((By.XPATH, "//tr[td[normalize-space()='RECIPT_DATE']]//select[@name='RPTOperator']")))
        from selenium.webdriver.support.ui import Select
        dropdown = Select(driver.find_element(
            By.XPATH,
            "//tr[td[normalize-space()='RECIPT_DATE']]//select[@name='RPTOperator']"
        ))
        dropdown.select_by_value("BETWEEN")
        logger.info("Selected BETWEEN from RPTOperator dropdown")
        time.sleep(1)

        date_format = "%d/%m/%Y"
        end_date = datetime.strptime(reportDate, date_format)
        if startDate:
            start_date = datetime.strptime(startDate, date_format)
        else:
            start_date = end_date - timedelta(days=6)
        date = f"'{start_date.strftime(date_format)}','{end_date.strftime(date_format)}'"

        logger.info(f"Entering date range: {date}")
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[td[normalize-space()='RECIPT_DATE']]//input[contains(@class,'RPTCriteria')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", date_input)
        driver.execute_script(
            "arguments[0].value = '';"
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            date_input, date
        )
        logger.info(f"Date range entered: {date}")

        # Select '=' operator for RECIPT_STATUS row
        wait.until(EC.presence_of_element_located((By.XPATH, "//tr[td[normalize-space()='RECIPT_STATUS']]//select[@name='RPTOperator']")))
        status_dropdown = Select(driver.find_element(
            By.XPATH,
            "//tr[td[normalize-space()='RECIPT_STATUS']]//select[@name='RPTOperator']"
        ))
        status_dropdown.select_by_value("=")
        logger.info("Selected '=' from RECIPT_STATUS RPTOperator dropdown")
        time.sleep(0.5)

        # Enter "Active" in the RECIPT_STATUS criteria input
        status_input = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[td[normalize-space()='RECIPT_STATUS']]//input[contains(@class,'RPTCriteria')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", status_input)
        driver.execute_script(
            "arguments[0].value = '';"
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            status_input, "Active"
        )
        logger.info("Entered 'Active' in RECIPT_STATUS criteria")

        return True

    except Exception as e:
        logger.error(f"Failed to click Inquiry button: {e}")
        return False


def wait_for_download(download_path: Path, pre_files: set) -> Path:
    """
    Wait indefinitely for a new .xlsx file to finish downloading.
    Mirrors tb_from_erp_bot.py wait_for_download_completion logic.
    Returns the Path of the completed file, or None on error.
    """
    logger.info(f"Monitoring download folder: {download_path}")
    logger.info("Waiting indefinitely until download is detected...")

    start = time.time()
    while True:
        current_files = set(f for f in download_path.glob("*") if f.is_file())
        new_files = current_files - pre_files

        for f in new_files:
            name = f.name
            if name.endswith(".xlsx") and not name.endswith(".crdownload") and not name.endswith(".tmp"):
                try:
                    s1 = f.stat().st_size
                    time.sleep(1)
                    s2 = f.stat().st_size
                    if s1 == s2 and s1 > 0:
                        logger.info(f"Downloaded: {name}  ({s1:,} bytes)")
                        return f
                except Exception:
                    pass

        elapsed = int(time.time() - start)
        if elapsed % 10 == 0 and elapsed > 0:
            logger.info(f"  Still waiting... ({elapsed}s elapsed)")
        time.sleep(1)


def _find_ecs_outputs_folder() -> Path | None:
    """
    Read outputs/NBD_WF_18_ECS/completed_dates.txt, take the last line,
    and return the max-versioned matching date folder.
    """
    root = Path(__file__).resolve().parent.parent
    ecs_outputs = root / "outputs" / "NBD_WF_18_ECS"
    dates_file = ecs_outputs / "completed_dates.txt"

    if not dates_file.exists():
        logger.error(f"completed_dates.txt not found: {dates_file}")
        return None

    lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
    if not lines:
        logger.error("completed_dates.txt is empty")
        return None

    last_date = lines[-1]
    logger.info(f"ECS latest completed date: {last_date}")

    date_variants = {last_date}
    for fmt_in, fmt_out in [("%d-%m-%Y", "%d/%m/%Y"), ("%d/%m/%Y", "%d-%m-%Y")]:
        try:
            date_variants.add(datetime.strptime(last_date, fmt_in).strftime(fmt_out))
        except ValueError:
            pass

    candidates = []
    for child in ecs_outputs.iterdir():
        if not child.is_dir():
            continue
        if any(child.name.startswith(v) for v in date_variants):
            m = re.search(r'\((\d+)\)$', child.name)
            candidates.append((int(m.group(1)) if m else 0, child))

    if not candidates:
        logger.error(f"No folder found for date '{last_date}' in {ecs_outputs}")
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    folder = candidates[0][1]
    logger.info(f"ECS outputs folder: {folder}")
    return folder


def generate_report(driver) -> bool:
    """Click the Generate Report button to trigger the download."""
    try:
        logger.info("Clicking Generate Report button...")
        if not click_with_retries(driver, By.XPATH, "//button[@id='btn_02013_GenerateReport']", description="Generate Report"):
            logger.error("Could not click Generate Report button")
            return False
        logger.info("Generate Report button clicked — waiting for download...")
        return True

    except Exception as e:
        logger.error(f"generate_report failed: {e}")
        return False


def download_and_move(download_path: Path, pre_files: set, is_monthly: bool = False) -> bool:
    """
    Wait for the Audit Receipt file to finish downloading, then move it to the
    latest ECS outputs folder.

    - Weekly run  (is_monthly=False): removes existing 'Audit Receipt - Internal *.xlsx'
      that do NOT contain '(monthly)' in the name.
    - Monthly run (is_monthly=True):  removes existing 'Audit Receipt - Internal (monthly) *.xlsx'.
    """
    try:
        logger.info("Waiting for Audit Receipt file to download...")
        downloaded_file = wait_for_download(download_path, pre_files)

        dest_folder = _find_ecs_outputs_folder()
        if not dest_folder:
            logger.error("Cannot determine ECS outputs destination folder")
            return False

        if is_monthly:
            # Remove only existing monthly files
            for old in dest_folder.glob("Audit Receipt - Internal (monthly) *.xlsx"):
                try:
                    old.unlink()
                    logger.info(f"Removed existing monthly file: {old.name}")
                except Exception as e:
                    logger.warning(f"Could not remove {old.name}: {e}")
        else:
            # Remove only existing weekly files (exclude monthly)
            for old in dest_folder.glob("Audit Receipt - Internal *.xlsx"):
                if "(monthly)" not in old.name:
                    try:
                        old.unlink()
                        logger.info(f"Removed existing file: {old.name}")
                    except Exception as e:
                        logger.warning(f"Could not remove {old.name}: {e}")

        dest_path = dest_folder / downloaded_file.name
        downloaded_file.replace(dest_path)
        logger.info(f"Moved to: {dest_path}")
        return True

    except Exception as e:
        logger.error(f"download_and_move failed: {e}")
        return False


def main(report_date: str = None, start_date: str = None, is_monthly: bool = False):
    """Main execution function.

    Args:
        report_date: End date in DD/MM/YYYY format.
        start_date:  Optional start date in DD/MM/YYYY format.
                     If omitted, start = report_date - 6 days.
        is_monthly:  If True, removes Audit Receipt - Internal (monthly) *.xlsx on move.
                     If False (default), removes Audit Receipt - Internal *.xlsx (weekly).
    """
    import sys as _sys

    if report_date is None:
        if len(_sys.argv) > 1:
            report_date = _sys.argv[1]
        else:
            report_date = datetime.now().strftime("%d/%m/%Y")

    if start_date is None and len(_sys.argv) > 2:
        start_date = _sys.argv[2]

    # argv[3] = "monthly" flag
    if not is_monthly and len(_sys.argv) > 3 and _sys.argv[3].lower() == "monthly":
        is_monthly = True

    logger.info("=" * 60)
    logger.info("Audit Receipt ERP Bot starting...")
    logger.info(f"Report date: {report_date}" + (f", Start date: {start_date}" if start_date else "") + (" [MONTHLY]" if is_monthly else ""))
    logger.info("=" * 60)

    # Snapshot downloads folder once before any attempt
    download_path = Path(DOWNLOAD_DIR)
    pre_files = set(f for f in download_path.glob("*") if f.is_file())

    max_retries = 3
    for attempt in range(1, max_retries + 1):
        driver = None
        try:
            logger.info(f"Attempt {attempt}/{max_retries}...")
            driver = setup_chrome_driver(DOWNLOAD_DIR)

            if not login_to_erp(driver):
                logger.error("Login failed.")
                continue

            # if not select_premises(driver):
            #     logger.error("Failed to select premises.")
            #     continue

            if not navigate_to_reports(driver):
                logger.error("Failed to navigate to Reports.")
                continue

            if not click_inquiry(driver, report_date, start_date):
                logger.error("Failed to click Inquiry button.")
                continue

            if not generate_report(driver):
                logger.error("Failed to generate report.")
                continue

            time.sleep(3)

            if not download_and_move(download_path, pre_files, is_monthly=is_monthly):
                logger.error("Failed to download/move Audit Receipt report.")
                continue

            logger.info("=" * 60)
            logger.info("Audit Receipt ERP Bot completed successfully.")
            logger.info("=" * 60)
            return True

        except Exception as e:
            logger.error(f"Attempt {attempt} failed: {e}")

        finally:
            if driver:
                try:
                    driver.quit()
                    logger.info("Chrome driver closed")
                except Exception:
                    pass

        time.sleep(5)

    logger.error(f"All {max_retries} attempts failed.")
    return False


if __name__ == "__main__":
    success = main()
    if success:
        print("\nAudit Receipt bot completed successfully!")
    else:
        print("\nAudit Receipt bot failed!")
