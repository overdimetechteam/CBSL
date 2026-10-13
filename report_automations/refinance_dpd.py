import os
import sys
import glob
import shutil
import time
from datetime import datetime, date

# Force UTF-8 output on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


class _TeeWriter:
    def __init__(self, original, log_file_path):
        self._orig = original
        self._fh = open(log_file_path, "a", encoding="utf-8", errors="replace")
    def write(self, data):
        self._orig.write(data)
        self._fh.write(data)
        self._fh.flush()
    def flush(self):
        self._orig.flush()
        self._fh.flush()
    def fileno(self):
        return self._orig.fileno()


def _setup_log_tee():
    log_dir = os.environ.get("IFRS_LGD_LOG_DIR", "").strip()
    if not log_dir:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "logs", f"run_{ts}"
        )
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "refinance_dpd.log")
    sys.stdout = _TeeWriter(sys.stdout, log_file)
    sys.stderr = _TeeWriter(sys.stderr, log_file)
    print(f"Log file: {log_file}")


_setup_log_tee()

from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from cryptography.fernet import Fernet
from selenium.webdriver.common.keys import Keys

# ==========================
# PATH SETUP
# ==========================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
REPORTS_DIR = os.path.join(PROJECT_ROOT, "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
_SCRIPT_ROOT = os.path.dirname(CURRENT_DIR)


def _find_lgd_working_folder() -> str:
    """Scan working/LGD/ and return <latest-date-folder>/LGD/."""
    base = os.path.join(_SCRIPT_ROOT, "working", "LGD")
    if os.path.isdir(base):
        subfolders = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d))]
        if subfolders:
            subfolders.sort(key=lambda d: os.path.getmtime(os.path.join(base, d)), reverse=True)
            return os.path.join(base, subfolders[0], "LGD")
    return base


_lgd_base  = os.environ.get("LGD_BASE_FOLDER", "").strip()
INPUTS_DIR = os.path.join(_lgd_base, "LGD") if _lgd_base else _find_lgd_working_folder()
os.makedirs(INPUTS_DIR, exist_ok=True)

# ==========================
# LOAD .ENV (ENCRYPTED)
# ==========================
ENV_PATH = r"C:\CBSL\Script\.env"
load_dotenv(dotenv_path=ENV_PATH, override=True)

FERNET_KEY         = os.environ.get("FERNET_KEY", "").strip()
ENCRYPTED_USERNAME = os.getenv("ENC_USERNAME", "").strip()
ENCRYPTED_PASSWORD = os.getenv("ENC_PASSWORD", "").strip()

if not FERNET_KEY:
    raise Exception(" FERNET_KEY not found in system environment variables")
if not ENCRYPTED_USERNAME or not ENCRYPTED_PASSWORD:
    raise Exception(" ENC_USERNAME or ENC_PASSWORD not found in .env file")

# ==========================
# DECRYPT CREDENTIALS
# ==========================
try:
    fernet = Fernet(FERNET_KEY.encode())
    USERNAME = fernet.decrypt(ENCRYPTED_USERNAME.encode()).decode()
    PASSWORD = fernet.decrypt(ENCRYPTED_PASSWORD.encode()).decode()
    print("✓ Credentials decrypted successfully")
except Exception as e:
    raise Exception(f" Credential decryption failed: {e}")

# ==========================
# HELPER
# ==========================
def wait_for_loader(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "jqueryLoaderDiv"))
        )
        time.sleep(0.5)
    except:
        pass

# ==========================
# LOGIN
# ==========================
def erp_login():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    })

    print("Starting Chrome...")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    wait = WebDriverWait(driver, 30)

    try:
        print("Opening ERP login page...")
        driver.get("https://erp.assetline.lk/")
        wait_for_loader(driver)

        print("Entering credentials...")
        wait.until(EC.presence_of_element_located((By.ID, "strUserName"))).send_keys(USERNAME)
        wait.until(EC.presence_of_element_located((By.ID, "strPassword"))).send_keys(PASSWORD)

        print("Clicking sign in...")
        wait.until(EC.element_to_be_clickable((By.ID, "idSignIn"))).click()
        wait_for_loader(driver)

        # print("Clicking select button...")
        # wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.blue.btn-circle"))).click()
        # wait_for_loader(driver)

        print("✓ Login successful!")
        return driver, wait

    except Exception as e:
        driver.save_screenshot("login_error.png")
        driver.quit()
        raise Exception(f" ERP login failed: {e}")

# ==========================
# NAVIGATE TO REPORT PAGE
# ==========================
def click_finance(driver, wait):
    print("Navigating to Report page...")
    driver.get("https://erp.assetline.lk/Application/Home/REPORT")
    wait_for_loader(driver)
    time.sleep(2)
    print("✓ Report page loaded:", driver.current_url)

# ==========================
# CLICK FINANCE TILE
# ==========================
def click_finance_tile(driver, wait):
    print("Clicking on Finance Tile")

    finance_element = driver.find_element(
        By.XPATH,
        "//*[@id='iddivAppContent']/div[1]/div/div[2]/div[1]/div/div[1]/div/div/div"
    )

    finance_element.click()
    print("Clicked Finance tile..")

# click DPD tile
    dpd_element = driver.find_element(
        By.XPATH,
        "//*[@id='iddivAppContent']/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div/div/div/div"
    )

    dpd_element.click()
    print("Clicked DPD tile..")

# ==========================
# CLICK INQUIRY BUTTON
# ==========================
def click_inquiry(driver, wait):
    print("Clicking Inquiry button...")

    time.sleep(2)

    try:
        inquiry_btn = wait.until(
            EC.presence_of_element_located((By.ID, "btn_04227_Inquiry"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inquiry_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", inquiry_btn)
        wait_for_loader(driver)
        print("✓ Inquiry button clicked successfully")

    except Exception as e:
        print(f" Inquiry button not found: {e}")

# ==========================
# SET ACTIVATION DATE FILTER
# ==========================
def set_activation_date_filter(driver, wait, end_date_str: str):
    """
    Select BETWEEN operator and enter date range for NEW_CON_NO_ACTIVATION_DATE.

    Args:
        end_date_str: End date in DD/MM/YYYY format (e.g. "31/08/2025")
                      Start date is computed as the 1st day of the month 3 months before.
    """
    # Compute start date: 3 months before end date, on the 1st of that month
    end_dt = datetime.strptime(end_date_str, "%d/%m/%Y")
    start_dt = (end_dt - relativedelta(months=2)).replace(day=1)
    start_date_str = start_dt.strftime("%d/%m/%Y")
    date_range_value = f"'{start_date_str}','{end_date_str}'"

    print(f"Setting activation date filter: {date_range_value}")

    try:
        # Select BETWEEN from the operator dropdown
        operator_select_elem = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//tr[td[normalize-space()='NEW_CON_NO_ACTIVATION_DATE']]//select[@name='RPTOperator']"
            ))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", operator_select_elem)
        time.sleep(0.3)
        Select(operator_select_elem).select_by_value("BETWEEN")
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            operator_select_elem
        )
        wait_for_loader(driver)
        print("✓ Operator set to BETWEEN")

        # Enter the date range into the criteria input
        criteria_input = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//tr[td[normalize-space()='NEW_CON_NO_ACTIVATION_DATE']]//input[contains(@class,'RPTCriteria')]"
            ))
        )
        driver.execute_script("arguments[0].removeAttribute('readonly');", criteria_input)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", criteria_input)
        time.sleep(0.3)

        criteria_input.click()
        criteria_input.send_keys(Keys.CONTROL + "a")
        criteria_input.send_keys(Keys.DELETE)
        driver.execute_script(f"arguments[0].value = '';", criteria_input)
        criteria_input.send_keys(date_range_value)
        driver.execute_script("""
            arguments[0].dispatchEvent(new Event('input',  {bubbles:true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
            arguments[0].dispatchEvent(new Event('blur',   {bubbles:true}));
        """, criteria_input)

        time.sleep(0.3)
        actual = criteria_input.get_attribute("value")
        print(f"✓ Activation date range entered | Field shows: {actual}")

    except Exception as e:
        print(f" Activation date filter failed: {e}")

# ==========================
# GENERATE REPORT & DOWNLOAD
# ==========================
def click_generate_report(driver, wait):
    print("Clicking Generate Report button...")

    try:
        generate_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='btn_04227_GenerateReport']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", generate_btn)
        time.sleep(0.3)
        generate_btn.click()
        print("✓ Generate Report button clicked")

        # Wait for the modal to appear, indicating download has started
        print("Waiting for download confirmation modal...")
        wait.until(EC.presence_of_element_located((By.ID, "ModalDescriptionText")))
        print("✓ Modal appeared — download started")

        # Wait for the file to appear in Downloads (up to 120s)
        print("Waiting for file to arrive in Downloads folder...")
        downloaded_path = None
        for _ in range(120):
            matches = glob.glob(os.path.join(DOWNLOAD_DIR, "Refinance DPD Report *.xlsx"))
            # Ignore Chrome's .crdownload temp files
            complete = [f for f in matches if not f.endswith(".crdownload")]
            if complete:
                downloaded_path = complete[0]
                break
            time.sleep(1)

        if not downloaded_path:
            print(" Download timed out — file not found in Downloads folder")
            return

        print(f"✓ File downloaded: {downloaded_path}")

        # Remove any existing "Refinance DPD Report *.xlsx" in inputs dir
        existing = glob.glob(os.path.join(INPUTS_DIR, "Refinance DPD Report *.xlsx"))
        for f in existing:
            os.remove(f)
            print(f"✓ Removed existing file: {f}")

        # Move downloaded file to inputs dir
        dest = os.path.join(INPUTS_DIR)
        shutil.move(downloaded_path, dest)
        print(f"✓ File moved to: {dest}")

    except Exception as e:
        print(f" Generate Report failed: {e}")


# ==========================
# ENTER DATE RANGE
# ==========================
def enter_date_range(driver, wait, from_date, to_date):
    print(f"Entering date range: {from_date} to {to_date}...")

    time.sleep(2)

    try:
        date_inputs = driver.find_elements(
            By.XPATH,
            "//span[contains(@class,'glyphicon-calendar')]/ancestor::div[contains(@class,'input-group')]//input"
        )

        print(f"Found {len(date_inputs)} date input fields")

        if len(date_inputs) >= 2:

            for idx, (inp, val) in enumerate([(date_inputs[0], from_date), (date_inputs[1], to_date)]):
                label = "From" if idx == 0 else "To"

                # Remove readonly, clear existing value
                driver.execute_script("arguments[0].removeAttribute('readonly');", inp)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
                time.sleep(0.3)

                # Triple-click to select all then type
                inp.click()
                time.sleep(0.3)
                inp.send_keys(Keys.CONTROL + "a")
                time.sleep(0.2)
                inp.send_keys(Keys.DELETE)
                time.sleep(0.2)

                # Set value via JS first
                driver.execute_script(f"arguments[0].value = '{val}';", inp)

                # Then send_keys to simulate real typing
                inp.send_keys(val)

                # Fire all necessary events
                driver.execute_script("""
                    arguments[0].dispatchEvent(new Event('input',  {bubbles:true}));
                    arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
                    arguments[0].dispatchEvent(new Event('blur',   {bubbles:true}));
                """, inp)

                time.sleep(0.5)
                actual = inp.get_attribute('value')
                print(f"✓ {label} date entered: {val} | Field shows: {actual}")

        else:
            print(f" Expected 2 date fields but found {len(date_inputs)}")

        wait_for_loader(driver)

    except Exception as e:
        print(f" Date entry failed: {e}")


def run(date_str: str) -> str:

    driver, wait = erp_login()
    try:
        click_finance(driver, wait)
        click_finance_tile(driver, wait)
        click_inquiry(driver, wait)
        set_activation_date_filter(driver, wait, end_date_str=date_str)
        click_generate_report(driver, wait)
        
    except Exception as e:
        print(f" Error during report generation: {e}")
    finally:
        driver.quit()
        print("✓ GL CodeBase Report completed")


# ==========================
# MAIN
# ==========================
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python refinance_dpd.py DD/MM/YYYY")
        sys.exit(1)
    run(sys.argv[1])