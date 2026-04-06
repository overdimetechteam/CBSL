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
# ==========================
ENV_PATH = os.path.join(CURRENT_DIR, ".env")
load_dotenv(dotenv_path=ENV_PATH, override=True)

SECRET_KEY = os.getenv("ERP_SECRET_KEY")
ENCRYPTED_USERNAME = os.getenv("ERP_USERNAME_ENC")
ENCRYPTED_PASSWORD = os.getenv("ERP_PASSWORD_ENC")

if not all([SECRET_KEY, ENCRYPTED_USERNAME, ENCRYPTED_PASSWORD]):
    raise Exception("❌ Missing encrypted credentials in .env file")

# ==========================
# DECRYPT CREDENTIALS
# ==========================
try:
    fernet = Fernet(SECRET_KEY.encode())
    USERNAME = fernet.decrypt(ENCRYPTED_USERNAME.encode()).decode()
    PASSWORD = fernet.decrypt(ENCRYPTED_PASSWORD.encode()).decode()
    print("✓ Username & password decrypted successfully")
except Exception as e:
    raise Exception(f"❌ Credential decryption failed: {e}")

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
def erp_login():
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

        print("Clicking select button...")
        wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.blue.btn-circle"))
        ).click()
        wait_for_loader(driver)

        print("✓ Login successful!")
        return driver, wait

    except Exception as e:
        driver.save_screenshot("login_error.png")
        driver.quit()
        raise Exception(f"❌ ERP login failed: {e}")

# ==========================
# RUN DIRECTLY
# ==========================
if __name__ == "__main__":
    driver, wait = erp_login()
    print("\nKeeping browser open for 10 seconds...")
    time.sleep(10)
    driver.quit()
    print("✓ Script completed successfully")
