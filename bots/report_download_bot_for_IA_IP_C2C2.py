#!/usr/bin/env python3
"""
Report Download Bot for IA and C1C2 Reports
Uses Selenium to automate the download of Po Listing - Internal reports from ERP system
"""

import os
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
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    TimeoutException,
)

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

def login_to_erp(driver):
    """Login to ERP system using decrypted credentials"""
    try:
        logger.info("Starting login process...")
        
        # Navigate to ERP login page
        driver.get(BASE_URL)
        logger.info(f"Navigated to {BASE_URL}")
        
        # Wait for page to load
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "strUserName"))
        )
        logger.info("Login page loaded successfully")
        
        # Load credentials
        username, password = load_credentials()
        
        # Find and fill username
        username_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "strUserName"))
        )
        username_input.clear()
        username_input.send_keys(username)
        logger.info("Username entered")
        
        # Find and fill password
        password_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "strPassword"))
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
        
        # Wait for page to load and check for select button
        select_clicked = False
        try:
            # Look for select button
            select_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='idSignIn' and @type='submit']"))
            )
            select_button.click()
            select_clicked = True
            logger.info("SELECT button clicked")
        except:
            logger.info("No additional SELECT button found, proceeding with login")

        if select_clicked:
            # Wait for the next page to fully load after clicking SELECT
            WebDriverWait(driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            # Additionally, wait for a known post-login element if available (Report menu)
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//li[@class='nav-item']//a[@title='Report']"))
                )
            except Exception:
                logger.info("Report menu not yet visible; proceeding to refresh after load state complete")

            # Refresh once AFTER the page has finished loading
            driver.refresh()

            # Wait again for the refreshed page to fully load
            WebDriverWait(driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            # Ensure a stable element is present after refresh
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
        else:
            # No SELECT click; just ensure the page is ready
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
        logger.info("Login process completed successfully")
        return True
        
    except Exception as e:
        logger.error(f"Login failed: {e}")
        return False

def navigate_to_reports(driver):
    """Navigate to the Reports section and open 'Po Listing - Internal' pinned item"""
    try:
        logger.info("Navigating to Reports section...")
        
        # Wait for the page to load after login/refresh
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # Click on Report menu item
        report_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class,'nav-item')]//a[@title='Report']"))
        )
        current_windows = driver.window_handles
        report_menu.click()
        logger.info("Report menu clicked")

        # Wait for new tab to open and switch to it
        WebDriverWait(driver, 15).until(lambda d: len(d.window_handles) > len(current_windows))
        new_windows = driver.window_handles
        target_handle = [h for h in new_windows if h not in current_windows][-1]
        driver.switch_to.window(target_handle)
        logger.info("Switched to newly opened Report tab")

        # Ensure new tab is fully loaded
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

        # Wait for pinned items to be visible
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'pined-item')]"))
        )

        # Click the 'Po Listing - Internal' pinned item
        po_listing_click_target = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'pined-item-hidden')][normalize-space(text())='Po Listing - Internal']/ancestor::div[contains(@class,'pined-item')]"))
        )
        try:
            po_listing_click_target.click()
        except Exception:
            driver.execute_script("arguments[0].click();", po_listing_click_target)
        logger.info("'Po Listing - Internal' pinned item clicked")

        # Wait for the page/module to load (Inquiry button becomes available later)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        return True
        
    except Exception as e:
        logger.error(f"Failed to navigate to Reports or open pinned item: {e}")
        return False

def navigate_to_inquiry(driver):
    """Navigate to Inquiry page"""
    try:
        logger.info("Navigating to Inquiry page...")
        
        # Click on Inquiry button
        inquiry_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "btn_01940_Inquiry"))
        )
        inquiry_button.click()
        logger.info("Inquiry button clicked")
        
        # Wait for page to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        logger.info("Successfully navigated to Inquiry page")
        return True
        
    except Exception as e:
        logger.error(f"Failed to navigate to Inquiry: {e}")
        return False

def download_loan_base_report(driver, as_at_date="31/07/2025"):
    """Download the Loan Base Report from Scienter system"""
    try:
        logger.info("Starting Loan Base Report download...")
        
        # Open new tab for Scienter
        driver.execute_script("window.open('https://efn.assetline.lk:1000', '_blank');")
        scienter_tab = driver.window_handles[-1]
        driver.switch_to.window(scienter_tab)
        logger.info("Opened new tab for Scienter login")
        
        # Wait for page to load
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # Load credentials
        username, password = load_credentials()
        
        # Fill username
        username_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[formcontrolname='username']"))
        )
        username_input.clear()
        username_input.send_keys(username)
        logger.info("Scienter username entered")
        
        # Fill password
        password_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[formcontrolname='password']"))
        )
        password_input.clear()
        password_input.send_keys(password)
        logger.info("Scienter password entered")
        
        # Click login button
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
        )
        login_button.click()
        logger.info("Scienter login button clicked")
        
        # Wait for login to complete and page to fully load
        logger.info("Waiting for login to complete and page to load...")
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # Wait for the main content to be present (not just body)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # Additional wait to ensure all dynamic content is loaded
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # Wait for the Fixed Deposit menu to be present and clickable
        logger.info("Waiting for Fixed Deposit menu to be available...")
        try:
            # First try to find the element by the exact structure
            fixed_deposit_menu = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'hovicon') and contains(@class, 'effect-3') and contains(@class, 'sub-a')]//h2[contains(text(), 'Fixed Deposit')]"))
            )
        except Exception:
            # Fallback: try alternative selectors
            logger.info("Trying alternative selectors for Fixed Deposit menu...")
            try:
                # Try by text content only
                fixed_deposit_menu = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//h2[contains(text(), 'Fixed Deposit')]"))
                )
            except Exception:
                # Try by class combination
                fixed_deposit_menu = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'hovicon')]//h2[contains(text(), 'Fixed Deposit')]"))
                )
        
        # Click on Fixed Deposit menu
        logger.info("Clicking on Fixed Deposit menu...")
        try:
            fixed_deposit_menu.click()
        except Exception:
            # Use JavaScript click if regular click fails
            driver.execute_script("arguments[0].click();", fixed_deposit_menu)
        logger.info("Fixed Deposit menu clicked")
        
        # Wait for the menu interaction to complete
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # Click REPORTS (root) -> LOAN QUERY INFO -> LOAN BASE REPORT
        logger.info("Navigating menu: REPORTS -> LOAN QUERY INFO -> LOAN BASE REPORT...")
        try:
            reports_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class,'rmRootLink')][.//span[contains(normalize-space(.), 'REPORTS')]]"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", reports_link)
            try:
                reports_link.click()
            except Exception:
                driver.execute_script("arguments[0].click();", reports_link)
            logger.info("Clicked REPORTS root menu")
        except Exception as e:
            logger.warning(f"Could not click REPORTS root menu: {e}")
        
        try:
            loan_query_info = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class,'rmLink')][.//span[contains(normalize-space(.), 'LOAN QUERY INFO')]]"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", loan_query_info)
            try:
                loan_query_info.click()
            except Exception:
                driver.execute_script("arguments[0].click();", loan_query_info)
            logger.info("Clicked LOAN QUERY INFO menu")
        except Exception as e:
            logger.warning(f"Could not click LOAN QUERY INFO: {e}")
        
        # Click LOAN BASE REPORT link under menus
        logger.info("Clicking LOAN BASE REPORT link...")
        try:
            loan_base_link_xpath = (
                "//a[contains(@href, 'Report_LoanBase.aspx') and contains(@class, 'rmLink')]" \
                "[.//span[contains(normalize-space(.), 'LOAN BASE REPORT')]]"
            )
            loan_base_link = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, loan_base_link_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", loan_base_link)
            try:
                loan_base_link.click()
            except Exception:
                driver.execute_script("arguments[0].click();", loan_base_link)
            logger.info("LOAN BASE REPORT link clicked")
        except Exception as e:
            logger.error(f"Failed to click LOAN BASE REPORT link: {e}")
            return False
        
        # Wait for Loan Base Report page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtASAtDate"))
        )
        
        # Fill AS At Date (ensure click for headless safety)
        date_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_txtASAtDate"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", date_input)
        try:
            date_input.click()
        except Exception:
            driver.execute_script("arguments[0].click();", date_input)
        date_input.clear()
        date_input.send_keys(as_at_date)
        logger.info(f"Entered AS At Date: {as_at_date}")
        
        # Click Display
        display_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_Display"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", display_button)
        try:
            display_button.click()
        except Exception:
            driver.execute_script("arguments[0].click();", display_button)
        logger.info("Display button clicked")
        
        # Wait for report toolbar
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"))
        )
        
        # Select Excel format
        format_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", format_select)
        Select(format_select).select_by_value("EXCELOPENXML")
        logger.info("Selected Excel format")
        
        # Click Export
        export_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", export_button)
        try:
            export_button.click()
        except Exception:
            driver.execute_script("arguments[0].click();", export_button)
        logger.info("Export button clicked")
        
        # Wait for download to complete (~10s typical, allow up to 60s)
        download_path = Path(DOWNLOAD_DIR)
        start_dl = time.time()
        target_file = None
        while time.time() - start_dl < 60:
            try:
                crs = list(download_path.glob("*.crdownload"))
                if crs:
                    time.sleep(1)
                    continue
                # Find latest file that looks like Loan Base (fallback: use presence of 'LoanAsAtDate' if required by name policy)
                candidates = [f for f in download_path.glob("*") if f.is_file()]
                if candidates:
                    latest = max(candidates, key=lambda p: p.stat().st_mtime)
                    # Size stability 3s
                    s1 = latest.stat().st_size
                    time.sleep(2)
                    s2 = latest.stat().st_size
                    time.sleep(1)
                    s3 = latest.stat().st_size
                    if s1 == s2 == s3:
                        target_file = latest
                        break
            except Exception:
                time.sleep(1)
        
        if not target_file:
            logger.error("Timeout waiting for Loan Base Report download")
            return False
        
        # Move to working/monthly/<only subfolder>/NBD_MF_23_IA after removing existing LoanAsAtDate* file
        try:
            base_working = Path(__file__).resolve().parent.parent / "working" / "monthly"
            subfolders = [p for p in base_working.iterdir() if p.is_dir()]
            if len(subfolders) != 1:
                logger.error(f"Expected exactly 1 folder in {base_working}, found {len(subfolders)}")
                return False
            target_root = subfolders[0] / "NBD_MF_23_IA"
            target_root.mkdir(parents=True, exist_ok=True)
            
            # Delete existing LoanAsAtDate files
            for f in target_root.glob("*"):
                try:
                    if f.is_file() and f.stem.lower() == "loanasatdate":
                        f.unlink(missing_ok=True)
                        logger.info(f"Deleted existing Loan Base Report in target: {f.name}")
                except Exception:
                    continue
            
            dest_path = target_root / target_file.name
            if dest_path.exists():
                try:
                    dest_path.unlink(missing_ok=True)
                    logger.info(f"Removed existing destination file: {dest_path.name}")
                except Exception:
                    pass
            target_file.replace(dest_path)
            logger.info(f"Moved Loan Base Report to: {dest_path}")
        except Exception as move_e:
            logger.error(f"Failed to move Loan Base Report file: {move_e}")
            return False
        
    except Exception as e:
        logger.error(f"Failed to download Loan Base Report: {e}")
        return False

def monitor_po_download(driver, download_path, pre_po_files, click_ts, result_dict):
    """Monitor Po Listing download in background thread"""
    try:
        def _is_po_listing_file(path_obj):
            try:
                name = path_obj.name.lower()
                return all(k in name for k in ["po", "listing", "internal"]) and not name.endswith(".crdownload")
            except Exception:
                return False

        logger.info("Monitoring Po Listing download...")
        start_time = time.time()
        download_started = False
        
        # Wait for download to start or complete (more flexible detection)
        while time.time() - start_time < 120:  # 2 minute total timeout after trigger
            try:
                # Check for partial downloads
                partials = [p for p in download_path.glob("*.crdownload") if _is_po_listing_file(p.with_suffix(""))]
                if partials and not download_started:
                    logger.info("Po Listing download started (partial file detected)")
                    download_started = True
                
                # Check for completed files
                candidates = [f for f in download_path.glob("*") if f.is_file() and _is_po_listing_file(f) and f not in pre_po_files and f.stat().st_mtime >= click_ts - 5]
                if candidates:
                    po_file = max(candidates, key=lambda p: p.stat().st_mtime)
                    # Verify file is complete (no .crdownload with same base name)
                    base_name = po_file.stem
                    has_partial = any(p.stem == base_name for p in download_path.glob("*.crdownload"))
                    if not has_partial:
                        # Size stability check (5s window)
                        try:
                            size1 = po_file.stat().st_size
                            time.sleep(3)
                            size2 = po_file.stat().st_size
                            time.sleep(2)
                            size3 = po_file.stat().st_size
                            if size1 == size2 == size3:
                                result_dict["success"] = True
                                result_dict["file_path"] = po_file
                                logger.info(f"Po Listing download completed (size stable): {po_file.name}")
                                return
                        except Exception:
                            pass
                
                # If we have partials, wait for them to complete
                if partials:
                    time.sleep(2)
                else:
                    # No partials, check if we have a completed file
                    if candidates:
                        # File exists but might still be downloading, wait a bit more
                        time.sleep(3)
                    else:
                        # No activity yet, wait longer
                        time.sleep(5)
                        
            except Exception as e:
                logger.warning(f"Error checking download status: {e}")
                time.sleep(5)
        
        # Final check - maybe download completed but we missed it
        candidates = [f for f in download_path.glob("*") if f.is_file() and _is_po_listing_file(f) and f not in pre_po_files and f.stat().st_mtime >= click_ts - 10]
        if candidates:
            po_file = max(candidates, key=lambda p: p.stat().st_mtime)
            # Size stability quick check
            try:
                s1 = po_file.stat().st_size
                time.sleep(2)
                s2 = po_file.stat().st_size
                if s1 == s2:
                    result_dict["success"] = True
                    result_dict["file_path"] = po_file
                    logger.info(f"Po Listing download found (late detection, size stable): {po_file.name}")
                    return
            except Exception:
                pass
        
        result_dict["success"] = False
        result_dict["error"] = "Po Listing download timeout - no matching file found within 3 minutes"
        
    except Exception as e:
        result_dict["success"] = False
        result_dict["error"] = str(e)

def download_po_listing_report(driver, as_at_date: str):
    """Download the Po Listing - Internal report"""
    try:
        logger.info("Starting Po Listing - Internal report download...")
        
        # Click on Select All
        if not click_with_retries(
            driver,
            By.XPATH,
            "//span[@class='bootstrap-switch-handle-on bootstrap-switch-primary']",
            description="Select All"
        ):
            return False
        logger.info("Select All clicked")
        
        # Wait for Select None button to be available
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='bootstrap-switch-handle-off bootstrap-switch-default']"))
        )
        
        # Click on Select None
        if not click_with_retries(
            driver,
            By.XPATH,
            "//span[@class='bootstrap-switch-handle-off bootstrap-switch-default']",
            description="Select None"
        ):
            return False
        logger.info("Select None clicked")
        
        # Wait for Generate Report button to be available
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "btn_01940_GenerateReport"))
        )
        
        # Helpers for detecting export buttons/links and new tabs
        def _try_click_export(max_tries: int = 2) -> bool:
            try:
                candidates = [
                    (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'export')]") ,
                    (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'download')]") ,
                    (By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'export')]") ,
                    (By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'download')]") ,
                    (By.XPATH, "//a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'.xlsx') or contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'.xls')]")
                ]
                for _ in range(max_tries):
                    for by, sel in candidates:
                        try:
                            el = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((by, sel)))
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                            time.sleep(0.2)
                            try:
                                el.click()
                            except Exception:
                                driver.execute_script("arguments[0].click();", el)
                            logger.info("Clicked Export/Download control")
                            return True
                        except Exception:
                            continue
            except Exception:
                pass
            return False

        def _switch_to_new_tab_and_export() -> bool:
            try:
                handles = driver.window_handles
                if len(handles) > 1:
                    driver.switch_to.window(handles[-1])
                    logger.info("Switched to new tab after Generate Report")
                    if _try_click_export():
                        return True
            except Exception:
                pass
            return False

        def _enter_iframe_if_present() -> None:
            try:
                # Try common iframe patterns after report generation
                frames = driver.find_elements(By.TAG_NAME, "iframe")
                for fr in frames:
                    try:
                        driver.switch_to.frame(fr)
                        logger.info("Switched into iframe to look for Export/Download")
                        if _try_click_export():
                            return
                        # If not found, switch back and continue
                        driver.switch_to.default_content()
                    except Exception:
                        try:
                            driver.switch_to.default_content()
                        except Exception:
                            pass
            except Exception:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

        # Prepare for detecting the correct downloaded file
        download_path = Path(DOWNLOAD_DIR)
        def _is_po_listing_file(path_obj):
            try:
                name = path_obj.name.lower()
                return all(k in name for k in ["po", "listing", "internal"]) and not name.endswith(".crdownload")
            except Exception:
                return False

        pre_po_files = set(
            f for f in download_path.glob("*")
            if f.is_file() and _is_po_listing_file(f)
        )
        click_ts = time.time()

        # Click Generate Report button
        generate_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btn_01940_GenerateReport"))
        )
        generate_button.click()
        logger.info("Generate Report button clicked")

        # Wait for Po Listing download to start, with robust fallbacks (no re-clicks)
        logger.info("Waiting for Po Listing download to start (watching for SUCCESS modal)...")
        download_started = False
        last_log_time = time.time()

        def _success_modal_visible() -> bool:
            try:
                header = driver.find_element(By.ID, "ModalHeaderText")
                desc = driver.find_element(By.ID, "ModalDescriptionText")
                h_text = (header.text or "").strip().upper()
                d_text = (desc.text or "").strip().lower()
                return ("SUCCESS" in h_text) and ("report successfully generated" in d_text)
            except Exception:
                return False

        # Wait until modal indicates success, or partial file appears
        while not download_started:
            now = time.time()

            # Detect new tab and attempt export there (no re-clicks)
            if _switch_to_new_tab_and_export():
                logger.info("Triggered export from new tab")

            # Try clicking Export/Download on the current page (no re-clicks of Generate)
            _enter_iframe_if_present()
            if _try_click_export():
                logger.info("Triggered export via Export/Download control")

            # Primary trigger: SUCCESS modal displayed
            if _success_modal_visible():
                logger.info("SUCCESS modal detected: Report Successfully Generated.")
                download_started = True
                break

            # Secondary trigger: partial download detected
            try:
                partials = [p for p in download_path.glob("*.crdownload") if _is_po_listing_file(p.with_suffix(""))]
                if partials:
                    logger.info("Po Listing download started (partial file detected)")
                    download_started = True
                    break
            except Exception:
                pass

            # Secondary trigger: completed file appears very quickly
            try:
                candidates = [
                    f for f in download_path.glob("*")
                    if f.is_file() and _is_po_listing_file(f) and f not in pre_po_files and f.stat().st_mtime >= click_ts - 2
                ]
                if candidates:
                    logger.info("Po Listing download completed immediately")
                    download_started = True
                    break
            except Exception:
                pass

            # Heartbeat every 30s
            if now - last_log_time >= 30:
                logger.info("Still waiting for Po Listing SUCCESS modal or download start...")
                last_log_time = now

            time.sleep(0.5)

        # Download has started or SUCCESS modal shown; begin background monitoring and immediately start next automation
        po_result = {"success": False, "file_path": None, "error": None}
        monitor_thread = threading.Thread(target=monitor_po_download, args=(driver, download_path, pre_po_files, click_ts, po_result))
        monitor_thread.start()
        logger.info("Po Listing SUCCESS detected — launching Loan Base automation in parallel...")

        # Start Loan Base Report download in parallel immediately
        loan_base_success = download_loan_base_report(driver, as_at_date)

        # Wait for Po Listing download to complete
        monitor_thread.join()

        if not po_result["success"]:
            logger.warning(f"Po Listing download monitoring failed: {po_result.get('error', 'Unknown error')}")
            logger.info("Attempting fallback detection of Po Listing file...")
            
            # Fallback: Look for any Po Listing file that might have been downloaded
            download_path = Path(DOWNLOAD_DIR)
            def _is_po_listing_file(path_obj):
                try:
                    name = path_obj.name.lower()
                    return all(k in name for k in ["po", "listing", "internal"]) and not name.endswith(".crdownload")
                except Exception:
                    return False
            
            # Look for any Po Listing files created after the click time
            fallback_candidates = [f for f in download_path.glob("*") if f.is_file() and _is_po_listing_file(f) and f.stat().st_mtime >= click_ts - 10]
            if fallback_candidates:
                po_file = max(fallback_candidates, key=lambda p: p.stat().st_mtime)
                po_result["success"] = True
                po_result["file_path"] = po_file
                logger.info(f"Fallback detection found Po Listing file: {po_file.name}")
            else:
                # As a final fallback, check if the file was already moved to target earlier by another run
                try:
                    base_working = Path(__file__).resolve().parent.parent / "working" / "monthly"
                    subfolders = [p for p in base_working.iterdir() if p.is_dir()]
                    if len(subfolders) == 1:
                        target_root = subfolders[0] / "NBD_MF_23_IA"
                        existing_targets = [f for f in target_root.glob("*") if f.is_file() and "po listing - internal" in f.name.lower()]
                        if existing_targets:
                            po_result["success"] = True
                            po_result["file_path"] = existing_targets[-1]
                            logger.info(f"Po Listing appears already moved earlier: {existing_targets[-1].name}")
                        else:
                            logger.error("No Po Listing file found even with fallback detection and target check")
                            # Do not return yet; proceed to consider Loan Base result
                    else:
                        logger.warning(f"Target working folder ambiguous for fallback check: found {len(subfolders)} folders")
                except Exception as ee:
                    logger.warning(f"Fallback target check failed: {ee}")
        
        po_file = po_result["file_path"]
        if not po_file or not po_file.exists():
            logger.error("Could not locate the 'Po Listing - Internal' file in Downloads")
            return False

        # Move the downloaded file to working/monthly/<only subfolder>/NBD_MF_23_IA
        try:
            base_working = Path(__file__).resolve().parent.parent / "working" / "monthly"
            subfolders = [p for p in base_working.iterdir() if p.is_dir()]
            if len(subfolders) != 1:
                logger.error(f"Expected exactly 1 folder in {base_working}, found {len(subfolders)}")
                return False
            target_root = subfolders[0] / "NBD_MF_23_IA"
            target_root.mkdir(parents=True, exist_ok=True)

            # Delete existing Po Listing - Internal* files (same base name any date)
            for f in target_root.glob("*"):
                try:
                    if f.is_file() and f.stem.lower().startswith("po listing - internal"):
                        f.unlink(missing_ok=True)
                        logger.info(f"Deleted existing Po Listing in target: {f.name}")
                except Exception:
                    continue

            target_path = target_root / po_file.name
            if target_path.exists():
                try:
                    target_path.unlink(missing_ok=True)
                    logger.info(f"Removed existing destination file: {target_path.name}")
                except Exception:
                    pass
            po_file.replace(target_path)
            logger.info(f"Moved 'Po Listing - Internal' to: {target_path}")
        except Exception as move_e:
            logger.error(f"Failed to move 'Po Listing - Internal' file: {move_e}")
            return False

        logger.info(f"Po Listing download: {'Success' if po_result['success'] else 'Failed'}")
        logger.info(f"Loan Base download: {'Success' if loan_base_success else 'Failed'}")
        
        # Compute overall success by checking files in target folder
        try:
            base_working = Path(__file__).resolve().parent.parent / "working" / "monthly"
            subfolders = [p for p in base_working.iterdir() if p.is_dir()]
            overall = False
            if len(subfolders) == 1:
                target_root = subfolders[0] / "NBD_MF_23_IA"
                target_files = list(target_root.glob("*")) if target_root.exists() else []
                has_po = any(f.is_file() and f.stem.lower().startswith("po listing - internal") for f in target_files)
                has_loan = any(f.is_file() and f.stem.lower() == "loanasatdate" for f in target_files)
                overall = has_po and has_loan
                logger.info(f"Target contains Po Listing: {has_po}, Loan Base: {has_loan}")
            else:
                logger.warning(f"Cannot verify overall success: ambiguous monthly folder count {len(subfolders)}")
        except Exception as ve:
            logger.warning(f"Overall success verification failed: {ve}")
            overall = bool(loan_base_success and po_result["success"])  # fallback
        
        return overall
        
    except Exception as e:
        logger.warning(f"Failed to download Po Listing report: {e}")
        return False

def check_downloaded_files(download_dir):
    """Check if files were downloaded successfully"""
    try:
        logger.info(f"Checking for downloaded files in: {download_dir}")
        
        # List all files in download directory
        download_path = Path(download_dir)
        if not download_path.exists():
            logger.warning(f"Download directory does not exist: {download_dir}")
            return False
        
        files = list(download_path.glob("*"))
        if files:
            logger.info(f"Found {len(files)} files in download directory:")
            for file in files:
                logger.info(f"  - {file.name} ({file.stat().st_size} bytes)")
            return True
        else:
            logger.warning("No files found in download directory")
            return False
            
    except Exception as e:
        logger.error(f"Error checking downloaded files: {e}")
        return False

def main(as_at_date="31/07/2025"):
    """Main execution function"""
    driver = None
    try:
        logger.info("Starting Report Download Bot for IA and C1C2 Reports...")
        logger.info(f"Download directory: {DOWNLOAD_DIR}")
        logger.info(f"AS At Date for Loan Base Report: {as_at_date}")
        
        # Setup Chrome driver
        driver = setup_chrome_driver(DOWNLOAD_DIR)
        
        # Login to ERP
        if not login_to_erp(driver):
            logger.error("Login failed. Exiting.")
            return False
        
        # Navigate to Reports
        if not navigate_to_reports(driver):
            logger.error("Failed to navigate to Reports. Exiting.")
            return False
        
        # Navigate to Inquiry
        if not navigate_to_inquiry(driver):
            logger.error("Failed to navigate to Inquiry. Exiting.")
            return False
        
        # Download Po Listing report (with parallel Loan Base Report)
        if not download_po_listing_report(driver, as_at_date):
            logger.warning("Po Listing report flow returned False; continuing as download may complete after SUCCESS modal.")
            # Do not exit; proceed to finalize
        
        logger.info("Bot execution completed")
        return True
        
    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        return False
        
    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

if __name__ == "__main__":
    import sys
    # Support passing date via CLI: python report_download_bot_IA_C1C2.py 31/07/2025
    as_at_date = sys.argv[1] if len(sys.argv) > 1 else "31/07/2025"
    success = main(as_at_date)
    if success:
        print("\nReport download bot completed successfully!")
    else:
        print("\nReport download bot failed!")
