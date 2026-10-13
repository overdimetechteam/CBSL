import os
import re
import shutil
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
from datetime import datetime
import openpyxl

# Load environment variables
try:
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _env_path = os.path.join(root_dir, '.env')
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

# Where the downloaded LoanAsAtDate file gets transferred to once the run
# completes: outputs/NBD-QF-23-C3-C10-SF-FI/{latest-completed-date}/Input/
OUTPUTS_DIR = Path(root_dir) / "outputs"
QF_C3_C10_DIR = OUTPUTS_DIR / "NBD-QF-23-C3-C10-SF-FI"

os.makedirs(LOGS_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, 'loan_base_report_transfer.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def wait_for_loader(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "jqueryLoaderDiv"))
        )
        time.sleep(0.5)
    except:
        pass


def load_credentials():
    try:
        secret_key = os.environ.get('FERNET_KEY', '').strip()
        enc_username = os.getenv('ENC_USERNAME')
        enc_password = os.getenv('ENC_PASSWORD')

        if not secret_key:
            raise ValueError("FERNET_KEY environment variable not set")
        if not enc_username or not enc_password:
            raise ValueError("ENC_USERNAME and ENC_PASSWORD must be set in environment variables")

        logger.info("Found encrypted credentials, attempting to decrypt...")

        fernet = Fernet(secret_key.encode())
        username = fernet.decrypt(enc_username.encode()).decode()
        password = fernet.decrypt(enc_password.encode()).decode()

        logger.info("Successfully decrypted credentials")
        return username, password

    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        raise


def setup_driver(download_dir=None):
    try:
        import subprocess
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'], capture_output=True, timeout=10)
            subprocess.run(['taskkill', '/F', '/IM', 'chromedriver.exe'], capture_output=True, timeout=10)
            time.sleep(2)
            logger.info("Closed existing Chrome/ChromeDriver processes")
        except Exception as e:
            logger.warning(f"Could not kill existing processes: {e}")

        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        if download_dir:
            prefs = {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            options.add_experimental_option("prefs", prefs)

        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(os.getenv('USERNAME', '')),
        ]

        for path in chrome_paths:
            if os.path.exists(path):
                options.binary_location = path
                logger.info(f"Found Chrome at: {path}")
                break
        else:
            raise Exception("Chrome browser not found.")

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logger.info("Chrome driver started successfully")
        return driver

    except Exception as e:
        logger.error(f"Failed to setup Chrome driver: {e}")
        raise


def login_to_scienter(driver, username, password):
    try:
        logger.info("Starting Scienter login process...")
        driver.get(SCIENTER_LOGIN_URL)
        time.sleep(5)

        username_input = None
        for selector in ["input[formcontrolname='username']", "input[placeholder='Username']",
                         "input[type='text']", "input[name='username']"]:
            try:
                username_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                break
            except TimeoutException:
                continue

        if not username_input:
            raise Exception("Could not find username input field")
        username_input.clear()
        username_input.send_keys(username)

        password_input = None
        for selector in ["input[formcontrolname='password']", "input[placeholder='Password']",
                         "input[type='password']", "input[name='password']"]:
            try:
                password_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                break
            except TimeoutException:
                continue

        if not password_input:
            raise Exception("Could not find password input field")
        password_input.clear()
        password_input.send_keys(password)

        login_button = None
        for selector in ["button[type='submit']", "input[type='submit']",
                         "input[value='Login']", "input[value='Sign In']"]:
            try:
                login_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                break
            except TimeoutException:
                continue

        if not login_button:
            password_input.send_keys(Keys.RETURN)
        else:
            login_button.click()

        time.sleep(10)
        current_url = driver.current_url
        logger.info(f"Current URL after login: {current_url}")

        if "login" in current_url.lower() or "signin" in current_url.lower():
            logger.warning("Still on login page - login might have failed")
            return False

        logger.info("Login successful")
        return True

    except Exception as e:
        logger.error(f"Login failed: {e}")
        return False


def click_fixed_deposit(driver):
    try:
        logger.info("Clicking Fixed Deposit button...")
        time.sleep(2)
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


# ==========================
# NAVIGATE TO DEPOSIT BASE
# ==========================
def navigate_to_deposit_base(driver):
    """Navigate REPORTS > LOAN QUERY INFO > LOAN BASE REPORT.
       Branch defaults to ALL — only clicks dropdown if not already ALL."""
    try:
        logger.info("Navigating to Deposit Base page...")
        time.sleep(3)

        # Step 1: Click REPORTS menu
        reports_menu = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'REPORTS')]"))
        )
        reports_menu.click()
        logger.info("Clicked REPORTS menu")
        time.sleep(2)

        # Step 2: Click LOAN QUERY INFO submenu
        fd_query_info = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'LOAN QUERY INFO')]"))
        )
        fd_query_info.click()
        logger.info("Clicked LOAN QUERY INFO submenu")
        time.sleep(2)

        # Step 3: Click LOAN BASE REPORT
        deposit_base = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'LOAN BASE REPORT')]"))
        )
        deposit_base.click()
        logger.info("Clicked LOAN BASE REPORT")
        time.sleep(5)

        # Step 4: Check Branch — already ALL by default, only act if not
        branch_value = driver.execute_script(
            "var el = document.getElementById('ctl00_ContentPlaceHolder1_txtBranch_Input');"
            "return el ? el.value : '';"
        )
        logger.info(f"Branch field current value: '{branch_value}'")

        if branch_value.strip().upper() != "ALL":
            logger.info("Branch not ALL — selecting ALL from dropdown...")
            try:
                combo_arrow = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_txtBranch_Arrow"))
                )
                driver.execute_script("arguments[0].click();", combo_arrow)
                time.sleep(1)

                all_option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH,
                        "//div[@id='ctl00_ContentPlaceHolder1_txtBranch_DropDown']"
                        "//li[contains(@class,'rcbItem') and normalize-space(text())='ALL']"
                    ))
                )
                driver.execute_script("arguments[0].click();", all_option)
                logger.info("Selected ALL in branch dropdown")
                time.sleep(1)

            except Exception as e:
                logger.warning(f"Dropdown click failed — forcing via JS: {e}")
                driver.execute_script("""
                    var inp = document.getElementById('ctl00_ContentPlaceHolder1_txtBranch_Input');
                    var state = document.getElementById('ctl00_ContentPlaceHolder1_txtBranch_ClientState');
                    if (inp) inp.value = 'ALL';
                    if (state) state.value = '{"logEntries":[],"value":"ALL","text":"ALL","enabled":true}';
                """)
                logger.info("Branch forced to ALL via JS")
        else:
            logger.info("Branch already ALL — no action needed")

        logger.info("navigate_to_deposit_base completed")
        return True

    except Exception as e:
        logger.error(f"Failed to navigate to Deposit Base: {e}")
        return False


# ==========================
# ENTER AS AT DATE & CLICK DISPLAY
# ==========================
def enter_as_at_date_and_display(driver, wait, date_str: str = None):
    """Enter the given date (or today if not provided) in As At Date field and click Display"""

    if date_str:
        for _fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                today = datetime.strptime(date_str.strip(), _fmt)
                break
            except ValueError:
                continue
        else:
            raise ValueError(f"Unrecognised date format: {date_str!r}")
    else:
        today = datetime.today()
    today_str = today.strftime("%d/%m/%Y")          # e.g. 01/03/2026
    today_day = str(today.day)                       # e.g. "1" — no leading zero, Windows safe
    logger.info(f"Entering As At Date: {today_str}")

    try:
        # ── Step 1: Click calendar image button to open popup ──
        try:
            calendar_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_ImageButtonAsATDate"))
            )
            calendar_btn.click()
            logger.info("Opened calendar popup via image button")
            time.sleep(1.5)

            # Try clicking today's date in calendar popup
            try:
                today_link = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.XPATH,
                        f"//table[contains(@id,'Calendar') or contains(@class,'ajax__calendar')]"
                        f"//td/a[text()='{today_day}']"
                    ))
                )
                today_link.click()
                logger.info(f"Clicked today ({today_day}) in calendar popup")
                time.sleep(1)
            except Exception:
                today_link = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.XPATH,
                        f"//div[contains(@id,'Calendar') or contains(@class,'ajax__calendar')]"
                        f"//a[text()='{today_day}']"
                    ))
                )
                today_link.click()
                logger.info(f"Clicked today ({today_day}) via fallback calendar")
                time.sleep(1)

        except Exception as e:
            logger.warning(f"Calendar popup click failed: {e} — typing date directly")

        # ── Step 2: Always set field value directly as safety net ──
        as_at_input = wait.until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtASAtDate"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", as_at_input)
        driver.execute_script("arguments[0].removeAttribute('readonly');", as_at_input)
        time.sleep(0.3)

        # Set value via JS only (avoids send_keys corrupting the field)
        driver.execute_script(f"arguments[0].value = '{today_str}';", as_at_input)
        driver.execute_script("""
            arguments[0].dispatchEvent(new Event('input',  {bubbles:true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
            arguments[0].dispatchEvent(new Event('blur',   {bubbles:true}));
        """, as_at_input)
        time.sleep(0.5)

        # ── Step 3: Verify ──
        actual_value = driver.execute_script("return arguments[0].value;", as_at_input)
        logger.info(f"As At Date field value: '{actual_value}' (expected: '{today_str}')")

        if actual_value != today_str:
            logger.warning("⚠️ Date mismatch — forcing value via JS again")
            driver.execute_script(f"arguments[0].value = '{today_str}';", as_at_input)
            driver.execute_script("""
                arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
                arguments[0].dispatchEvent(new Event('blur',   {bubbles:true}));
            """, as_at_input)
            time.sleep(0.3)

        # ── Step 4: Click Display ──
        display_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_Display"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", display_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", display_btn)
        wait_for_loader(driver)
        logger.info("Display button clicked successfully")

    except Exception as e:
        logger.error(f"enter_as_at_date_and_display failed: {e}")
        raise

# SCROLL REPORT (VERTICAL + HORIZONTAL)
# ==========================
def scroll_report(driver):
    """Scroll the Scienter report - handles iframe and direct page"""
    logger.info("Scrolling report vertically and horizontally...")

    try:
        time.sleep(3)

        # ── Step 1: Switch into iframe if report is inside one ──
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        logger.info(f"Found {len(iframes)} iframe(s) on page")

        switched_to_frame = False
        for i, frame in enumerate(iframes):
            try:
                src = frame.get_attribute("src") or ""
                name = frame.get_attribute("name") or ""
                fid = frame.get_attribute("id") or ""
                logger.info(f"  iframe[{i}] id='{fid}' name='{name}' src='{src[:80]}'")

                # Switch into frame and check for oReportDiv
                driver.switch_to.frame(frame)
                time.sleep(1)

                report_divs = driver.find_elements(By.ID, "oReportDiv")
                if report_divs:
                    logger.info(f"Found oReportDiv inside iframe[{i}]")
                    switched_to_frame = True
                    break
                else:
                    # Not in this frame, switch back and try next
                    driver.switch_to.default_content()

            except Exception as e:
                logger.warning(f"  iframe[{i}] switch failed: {e}")
                driver.switch_to.default_content()

        if not switched_to_frame:
            logger.warning("oReportDiv not found in any iframe — trying default content")
            driver.switch_to.default_content()

        # ── Step 2: Get scroll dimensions ──
        scroll_width = driver.execute_script(
            "var el = document.getElementById('oReportDiv'); "
            "return el ? el.scrollWidth : document.body.scrollWidth;"
        )
        scroll_height = driver.execute_script(
            "var el = document.getElementById('oReportDiv'); "
            "return el ? el.scrollHeight : document.body.scrollHeight;"
        )
        logger.info(f"Scroll dimensions — Width: {scroll_width}px | Height: {scroll_height}px")

        # ── Step 3: Scroll VERTICALLY top to bottom ──
        v_pos = 0
        v_step = 400
        while v_pos < scroll_height:
            v_pos += v_step
            driver.execute_script("""
                var el = document.getElementById('oReportDiv');
                if (el) { el.scrollTop = arguments[0]; }
                else { window.scrollTo(0, arguments[0]); }
            """, v_pos)
            time.sleep(0.3)
        logger.info(f"Scrolled vertically to bottom ({scroll_height}px)")
        time.sleep(1)

        # ── Step 4: Scroll HORIZONTALLY left to right ──
        h_pos = 0
        h_step = 400
        while h_pos < scroll_width:
            h_pos += h_step
            driver.execute_script("""
                var el = document.getElementById('oReportDiv');
                if (el) { el.scrollLeft = arguments[0]; }
                else { window.scrollTo(arguments[0], 0); }
            """, h_pos)
            time.sleep(0.3)
        logger.info(f"Scrolled horizontally to right ({scroll_width}px)")
        time.sleep(1)

        # ── Step 5: Return to top-left ──
        driver.execute_script("""
            var el = document.getElementById('oReportDiv');
            if (el) { el.scrollTop = 0; el.scrollLeft = 0; }
            else { window.scrollTo(0, 0); }
        """)
        logger.info("Scrolled back to top-left")

        # ── Step 6: Switch back to main content ──
        if switched_to_frame:
            driver.switch_to.default_content()
            logger.info("Switched back to default content")

        logger.info("Report scroll completed successfully")

    except Exception as e:
        logger.error(f"Scroll failed: {e}")
        try:
            driver.switch_to.default_content()
        except:
            pass

def export_report(driver, wait):
    """Select EXCELOPENXML from dropdown and click Export."""
    # Select EXCELOPENXML from dropdown
    dropdown = wait.until(EC.presence_of_element_located((
        By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"
    )))
    select = Select(dropdown)
    select.select_by_value("EXCELOPENXML")

    # Click Export button
    export_btn = wait.until(EC.element_to_be_clickable((
        By.ID, "ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"
    )))
    export_btn.click()
    logger.info("Export button clicked successfully")

def get_capital_balance_sum(file_path):
    """
    Reads the Excel file and returns the FINAL TOTAL of Capital Balance (Column P).
    The last row contains the actual grand total value.

    Args:
        file_path (str): Path to the downloaded Excel file.

    Returns:
        float: Final Capital Balance total from the last summary row.
    """
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    last_value = None

    # Column P = index 16 (1-based), scan all rows to find last numeric value
    for row in ws.iter_rows(min_row=5, min_col=17, max_col=17, values_only=True):
        value = row[0]
        if isinstance(value, (int, float)):
            last_value = value  # Keep overwriting — last one is the grand total

    return last_value


def _find_date_folder_with_variations(parent_folder: Path, date_str: str):
    """Find the latest variation of a date folder (e.g. 30-09-2025, 30-09-2025(1)).
    Mirrors find_date_folder_with_variations() in cbsl.py so this standalone
    script doesn't need to import the whole Flask app just for this lookup."""
    if not parent_folder.exists():
        return None

    candidates = []
    date_variants = {date_str}
    try:
        dt = datetime.strptime(date_str, "%d/%m/%Y")
        date_variants.add(dt.strftime("%d-%m-%Y"))
    except ValueError:
        try:
            dt = datetime.strptime(date_str, "%d-%m-%Y")
            date_variants.add(dt.strftime("%d/%m/%Y"))
        except ValueError:
            pass

    for child in parent_folder.iterdir():
        if not child.is_dir():
            continue
        folder_name = child.name
        if any(folder_name.startswith(v) for v in date_variants):
            version_match = re.search(r'\((\d+)\)$', folder_name)
            version = int(version_match.group(1)) if version_match else 0
            candidates.append((version, child))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def transfer_downloaded_file_to_input(file_path):
    """
    Copy the downloaded LoanAsAtDate file into:
      outputs/NBD-QF-23-C3-C10-SF-FI/{latest-completed-date}/Input/
    replacing any existing file with a similar name (same 'LoanAsAtDate*'
    stem and extension). The target date folder is the one matching the
    LAST date recorded in outputs/NBD-QF-23-C3-C10-SF-FI/completed_dates.txt.

    Returns the destination Path on success, or None if the transfer could
    not be completed (logged, but never raises — this must not break the
    Loan Base Report run itself).
    """
    try:
        if not QF_C3_C10_DIR.exists():
            logger.warning(f"Transfer skipped: base folder does not exist: {QF_C3_C10_DIR}")
            return None

        completed_dates_file = QF_C3_C10_DIR / "completed_dates.txt"
        if not completed_dates_file.exists():
            logger.warning(f"Transfer skipped: completed_dates.txt not found in {QF_C3_C10_DIR}")
            return None

        with open(completed_dates_file, 'r') as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]

        if not lines:
            logger.warning(f"Transfer skipped: completed_dates.txt is empty in {QF_C3_C10_DIR}")
            return None

        last_date_str = lines[-1]
        logger.info(f"Last completed date from completed_dates.txt: {last_date_str}")

        date_folder = _find_date_folder_with_variations(QF_C3_C10_DIR, last_date_str)
        if not date_folder:
            logger.warning(f"Transfer skipped: no date folder matching '{last_date_str}' found in {QF_C3_C10_DIR}")
            return None

        target_dir = date_folder / "Input"
        target_dir.mkdir(parents=True, exist_ok=True)

        source_path = Path(file_path)

        # Remove any existing file with a similar name before copying the new one
        similar_name_pattern = f"LoanAsAtDate*{source_path.suffix}"
        for old_file in target_dir.glob(similar_name_pattern):
            try:
                old_file.unlink()
                logger.info(f"Removed old file in Input folder: {old_file.name}")
            except Exception as e:
                logger.warning(f"Failed to remove old file {old_file.name}: {e}")

        dest_file = target_dir / source_path.name
        shutil.copy2(source_path, dest_file)
        logger.info(f"Transferred downloaded file to: {dest_file}")
        return dest_file

    except Exception as e:
        logger.error(f"Failed to transfer downloaded file to Input folder: {e}")
        return None


def get_total_capital_balance(driver, download_dir):
    """Export the report to Excel, wait for the download, transfer it into the
    QF C3-C10 Input folder, and return total capital balance."""
    wait = WebDriverWait(driver, 30)
    try:
        export_report(driver, wait)

        file_path = None
        for _ in range(30):
            time.sleep(1)
            matches = glob.glob(os.path.join(download_dir, "LoanAsAtDate*.xlsx"))
            if matches:
                file_path = max(matches, key=os.path.getmtime)
                break

        if not file_path:
            logger.error("Downloaded Excel file not found in Downloads folder")
            return None

        logger.info(f"Found downloaded file: {file_path}")
        time.sleep(2)
        total = get_capital_balance_sum(file_path)
        logger.info(f"Capital Balance Total: {total:,.2f}")

        # End-of-process transfer into the QF C3-C10 Input folder
        transfer_downloaded_file_to_input(file_path)

        return total

    except Exception as e:
        logger.error(f"get_total_capital_balance failed: {e}")
        return None


def _to_ddmmyyyy(date_str: str) -> str:
    """Normalise any common date format to DD/MM/YYYY."""
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    raise ValueError(f"Unrecognised date format: {date_str!r}")


def run(date_str: str) -> float:
    """
    Run the Loan Base Report for the given date and return the capital balance.

    Args:
        date_str: Date in any common format (DD/MM/YYYY, M/D/YYYY, YYYY-MM-DD, etc.)

    Returns:
        Capital balance as a float, or None on failure
    """
    date_str = _to_ddmmyyyy(date_str)
    driver = None
    try:
        username, password = load_credentials()
        download_dir = str(Path.home() / "Downloads")
        driver = setup_driver(download_dir=download_dir)

        if not login_to_scienter(driver, username, password):
            logger.error("Login failed.")
            return None

        if not click_fixed_deposit(driver):
            logger.error("Failed to click Fixed Deposit.")
            return None

        if not navigate_to_deposit_base(driver):
            logger.error("Failed to navigate to Deposit Base.")
            return None

        wait = WebDriverWait(driver, 15)
        enter_as_at_date_and_display(driver, wait, date_str=date_str)
        scroll_report(driver)

        capital_balance = get_total_capital_balance(driver, download_dir)

        if capital_balance is not None:
            logger.info(f"Capital Balance: {capital_balance:,.2f}")

        return capital_balance

    except Exception as e:
        logger.error(f"Loan Base Report run failed: {e}")
        return None

    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")


def download_loan_base_report_and_trf():
    import sys

    if len(sys.argv) > 1:
        try:
            date_arg = _to_ddmmyyyy(sys.argv[1])
        except ValueError:
            print("Usage: python Loan_Base_Report_transfer.py <date>  (DD/MM/YYYY, YYYY-MM-DD, or M/D/YYYY)")
            sys.exit(1)
    else:
        date_arg = datetime.now().strftime("%d/%m/%Y")

    logger.info(f"Starting Loan Base Report for date: {date_arg}")

    driver = None
    try:
        username, password = load_credentials()
        download_dir = str(Path.home() / "Downloads")
        matches = glob.glob(os.path.join(download_dir, "LoanAsAtDate*.xlsx"))
        for f in matches:
            os.remove(f)
        driver = setup_driver(download_dir=download_dir)

        if not login_to_scienter(driver, username, password):
            logger.error("Login failed. Exiting.")
            sys.exit(1)

        if not click_fixed_deposit(driver):
            logger.error("Failed to click Fixed Deposit. Exiting.")
            sys.exit(1)

        if not navigate_to_deposit_base(driver):
            logger.error("Failed to navigate to Deposit Base. Exiting.")
            sys.exit(1)

        wait = WebDriverWait(driver, 15)
        enter_as_at_date_and_display(driver, wait)
        scroll_report(driver)

        capital_balance = get_total_capital_balance(driver, download_dir)


        if capital_balance is not None:
            print(f"\n{'='*45}")
            print(f"  Report     : Loan Base As At Report")
            print(f"  Date       : {datetime.today().strftime('%d/%m/%Y')}")
            print(f"  Capital Bal: {capital_balance:,.2f}")
            print(f"{'='*45}")

        logger.info("All steps completed successfully.")
        time.sleep(10)

    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        sys.exit(1)

    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")



