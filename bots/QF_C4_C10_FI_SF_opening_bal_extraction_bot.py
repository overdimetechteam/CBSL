import os
import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from cryptography.fernet import Fernet
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# ---------------- Load environment ----------------
_root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_env_path = os.path.join(_root_dir, '.env')
if os.path.exists(_env_path):
    load_dotenv(_env_path)
else:
    load_dotenv()

BASE_URL = "https://erp.assetline.lk/"

# ---------------- Logging Setup ----------------
LOGS_DIR = "logs"
os.makedirs(LOGS_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, 'selenium_login.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ---------------- Credential Decryption ----------------
def load_credentials():
    try:
        enc_username = os.getenv('ENC_USERNAME')
        enc_password = os.getenv('ENC_PASSWORD')
        if not enc_username or not enc_password:
            raise ValueError("ENC_USERNAME and ENC_PASSWORD must be set in environment variables")

        key_str = os.environ.get("FERNET_KEY", "").strip()
        if not key_str:
            raise RuntimeError("FERNET_KEY environment variable not set")
        fernet = Fernet(key_str.encode())
        username = fernet.decrypt(enc_username.encode()).decode()
        password = fernet.decrypt(enc_password.encode()).decode()

        logger.info("Credentials loaded successfully")
        return username, password
    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        raise

# ---------------- Selenium Login ----------------
def selenium_login():
    logger.info("Starting Selenium authentication...")

    options = Options()
    options.add_argument("--disable-notifications")
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    # Locate Chrome binary
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.join(os.getenv('USERPROFILE', ''), r"AppData\Local\Google\Chrome\Application\chrome.exe")
    ]
    for path in chrome_paths:
        if os.path.exists(path):
            options.binary_location = path
            logger.info(f"Found Chrome at: {path}")
            break
    else:
        raise Exception("Chrome browser not found. Please install Google Chrome.")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    logger.info("Chrome driver started successfully")

    try:
        driver.get(BASE_URL)
        time.sleep(2)

        username, password = load_credentials()

        # Fill username
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "strUserName"))).send_keys(username)
        logger.info("Username entered")

        # Fill password
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "strPassword"))).send_keys(password)
        logger.info("Password entered")

        # Click login button
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSignIn"))).click()
        logger.info("Sign in button clicked")
        time.sleep(3)

        # Handle potential secondary SELECT button
        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='idSignIn' and @type='submit']"))
            ).click()
            logger.info("SELECT button clicked")
        except:
            logger.info("No additional SELECT button found")

        logger.info("Login successful")
        return driver

    except Exception as e:
        logger.error(f"Login failed: {e}")
        driver.quit()
        raise

def get_opening_balance(driver, subledger_code: str, from_date: str):
    """
    Navigate through Finance > GL Setup > GL Account View, perform search, and return Opening Balance.
    
    Args:
        driver: Selenium WebDriver instance (already logged in)
        subledger_code: str, e.g., "1846"
        from_date: str, format "dd/mm/yyyy", e.g., "31/07/2025"
        to_date: str, format "dd/mm/yyyy", e.g., "31/07/2025"
    
    Returns:
        opening_balance: str
    """
    try:
        logger.info("Refreshing page after login...")
        driver.refresh()
        time.sleep(3)

        # --- Click Finance nav link ---
        finance_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@title,'Finance')]"))
        )
        finance_link.click()
        logger.info("Finance menu clicked")
        time.sleep(2)

        # --- Switch to the new tab opened by Finance ---
        driver.switch_to.window(driver.window_handles[-1])
        logger.info("Switched to Finance tab")

        # --- Click GL Setup tile ---
        gl_setup_tile = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'metro-tile-body') and div[contains(text(),'GL Setup')]]"))
        )
        gl_setup_tile.click()
        logger.info("GL Setup tile clicked")
        time.sleep(2)

        # --- Click GL Account View tile ---
        gl_account_view_tile = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'metro-tile-body') and div[contains(text(),'GL Account View')]]"))
        )
        gl_account_view_tile.click()
        logger.info("GL Account View tile clicked")
        time.sleep(2)

        # --- Click Inquiry button ---
        inquiry_btn = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "btn_02366_Inquiry"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(inquiry_btn).click().perform()
        logger.info("Inquiry button clicked successfully")

        # --- Input Sub Ledger Code ---`
        input_id = "strSubLedgerCode_FINGLAccountView"
        driver.execute_script(
            "document.getElementById(arguments[0]).value = arguments[1];",
            input_id,
            subledger_code
        )
        logger.info(f"Sub Ledger Code entered via JS: {subledger_code}")

        # From Date
        from_date_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "dtmTransactionDate_FINGLAccountView"))
        )
        from_date_input.click()
        from_date_input.send_keys(Keys.CONTROL + "a")  # Select all existing text
        from_date_input.send_keys(Keys.DELETE)
        from_date_input.send_keys(from_date)
        logger.info(f"From date entered: {from_date}")

        # To Date
        to_date_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "dtmTransactionToDate_FINGLAccountView"))
        )
        to_date_input.click()
        to_date_input.send_keys(Keys.CONTROL + "a")
        to_date_input.send_keys(Keys.DELETE)
        to_date_input.send_keys(from_date)
        logger.info(f"To date entered: {from_date}")

        # --- Click Search ---
        search_btn = driver.find_element(By.ID, "btn_02366_BtnSearch_FINGLAccountView")
        search_btn.click()
        logger.info("Search button clicked")
        time.sleep(3)  # Wait for results to load

        # --- Get Opening Balance ---
        opening_balance_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "intOpenningBalance_FINGLAccountView"))
        )
        opening_balance = opening_balance_input.get_attribute("value")
        logger.info(f"Opening Balance found: {opening_balance}")

        print(f"Opening Balance: {opening_balance}")
        return opening_balance

    except Exception as e:
        logger.error(f"Failed to get Opening Balance: {e}")
        raise

# ---------------- Main Execution ----------------
if __name__ == "__main__":
    driver = selenium_login()
    try:
        opening_balance = get_opening_balance(
            driver,
            subledger_code="1846",
            from_date="31/07/2025"
        )
    finally:
        driver.quit()