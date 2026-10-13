import os
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from cryptography.fernet import Fernet

# ==========================
# PATH SETUP
# ==========================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
REPORTS_DIR = os.path.join(PROJECT_ROOT, "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

print("Current:", CURRENT_DIR)
print("Root:", PROJECT_ROOT)
print("Reports:", REPORTS_DIR)

# ==========================
# LOAD .ENV (ENCRYPTED)
# FERNET_KEY  — system environment variable
# ENC_USERNAME / ENC_PASSWORD — from .env file
# ==========================
ENV_PATH = os.path.join(PROJECT_ROOT, ".env")
load_dotenv(dotenv_path=ENV_PATH, override=True)

FERNET_KEY         = os.environ.get("FERNET_KEY", "").strip()
ENCRYPTED_USERNAME = os.getenv("ENC_USERNAME", "").strip()
ENCRYPTED_PASSWORD = os.getenv("ENC_PASSWORD", "").strip()

if not FERNET_KEY:
    raise Exception("FERNET_KEY not found in system environment variables")
if not ENCRYPTED_USERNAME or not ENCRYPTED_PASSWORD:
    raise Exception("ENC_USERNAME or ENC_PASSWORD not found in .env file")

# ==========================
# DECRYPT CREDENTIALS
# ==========================
try:
    fernet = Fernet(FERNET_KEY.encode())
    USERNAME = fernet.decrypt(ENCRYPTED_USERNAME.encode()).decode()
    PASSWORD = fernet.decrypt(ENCRYPTED_PASSWORD.encode()).decode()
    print("✓ Username & password decrypted successfully")
except Exception as e:
    raise Exception(f"Credential decryption failed: {e}")

# ==========================
# HELPER FUNCTION
# ==========================
def wait_for_loader(driver, timeout=30):
    """Wait for ERP loader to disappear"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "jqueryLoaderDiv"))
        )
        time.sleep(0.5)
    except:
        pass

# ==========================
# ERP LOGIN FUNCTION
# ==========================
LOGIN_RETRIES = 3          # retries after the first attempt (4 attempts in total)
LOGIN_RETRY_DELAY = 10     # seconds; multiplied by the retry number


def erp_login():
    """Log in to the ERP, retrying up to LOGIN_RETRIES times on failure (e.g.
    net::ERR_CONNECTION_TIMED_OUT). Every attempt starts a fresh browser, since
    a failed one is closed by _erp_login_once()."""
    last_error = None
    for attempt in range(LOGIN_RETRIES + 1):
        try:
            return _erp_login_once()
        except Exception as e:
            last_error = e
            if attempt < LOGIN_RETRIES:
                delay = LOGIN_RETRY_DELAY * (attempt + 1)
                print(f"\n✗ Login attempt {attempt + 1}/{LOGIN_RETRIES + 1} failed: {e}")
                print(f"  Retrying in {delay}s (retry {attempt + 1}/{LOGIN_RETRIES})...")
                time.sleep(delay)
    raise Exception(f"ERP login failed after {LOGIN_RETRIES + 1} attempts: {last_error}")


def _erp_login_once():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")

    print("\nStarting Chrome...")
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
        # wait.until(
        #     EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.blue.btn-circle"))
        # ).click()
        # wait_for_loader(driver)

        print("✓ Login successful!")
        return driver, wait

    except Exception as e:
        driver.save_screenshot("login_error.png")
        driver.quit()
        raise Exception(f"ERP login failed: {e}")

# ==========================
# RUN DIRECTLY
# ==========================
if __name__ == "__main__":
    driver, wait = erp_login()
    print("\nKeeping browser open for 10 seconds...")
    time.sleep(10)
    driver.quit()
    print("✓ Script completed successfully")
