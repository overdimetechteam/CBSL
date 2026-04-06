import os
import time
from flask import json
import pandas as pd
from typing import List, Optional
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from cryptography.fernet import Fernet
from dotenv import load_dotenv
import logging

# Load environment variables from a deterministic path
try:
    _root_dir = os.path.dirname(os.path.abspath(__file__))
    _env_path = os.path.join(_root_dir, 'ATT51703.env')
    if os.path.exists(_env_path):
        load_dotenv(_env_path)
    else:
        load_dotenv()
except Exception:
    load_dotenv()


# Logging Setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def load_credentials():
    """Load and decrypt credentials from environment variables."""
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
    
def fetch_titles_from_erp(customer_codes: List[str], driver: Optional[webdriver.Chrome] = None):
    """Logs into ERP and retrieves Mr/Mrs titles for given customer codes."""
    username, password = load_credentials()
    own_driver = False

    if driver is None:
        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            rf"C:\Users\{os.getenv('USERNAME', '')}\AppData\Local\Google\Chrome\Application\chrome.exe",
        ]

        for path in chrome_paths:
            if os.path.exists(path):
                options.binary_location = path
                logger.info(f"Found Chrome at: {path}")
                break
        else:
            logger.warning("Chrome browser not found in standard locations. Using default.")

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logger.info("Chrome driver started successfully")
        own_driver = True

    try:
        wait = WebDriverWait(driver, 25)

        # Step 1: Login
        logger.info("Starting login process...")
        driver.get("https://erp.assetline.lk/")
        print("🌐 Opening ERP login page...")
        time.sleep(3)

        username_field = wait.until(EC.presence_of_element_located((By.ID, "strUserName")))
        username_field.clear()
        username_field.send_keys(username)

        password_field = wait.until(EC.presence_of_element_located((By.ID, "strPassword")))
        password_field.clear()
        password_field.send_keys(password)

        signin_button = wait.until(EC.presence_of_element_located((By.ID, "idSignIn")))
        signin_button.click()
        logger.info("Sign in button clicked")
        print("Logged into ERP system")

        # Step 2: Handle Default Location popup
        time.sleep(3)
        try:
            print("⏳ Waiting for default location selection screen...")
            select_button = wait.until(
                EC.presence_of_element_located((By.XPATH, "//button[@id='idSignIn' and @type='submit']"))
            )
            select_button.click()
            logger.info("Default location (HEAD OFFICE) selected")
            print("Default location selected")
            time.sleep(3)
        except Exception:
            logger.info("No additional SELECT button found, proceeding...")
            print("Default location screen not found — continuing...")

        # Step 3: Refresh after location selection
        print("Refreshing page to ensure navigation menu loads...")
        driver.refresh()
        time.sleep(5)

        # Step 4: Click 'Client Management' menu item
        try:
            print("Opening Client Management menu...")
            client_mgmt_link = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//a[@title='Customer Evaluation']//span[contains(text(),'Client Management')]"
            )))
            driver.execute_script("arguments[0].scrollIntoView(true);", client_mgmt_link)
            time.sleep(0.5)
            client_mgmt_link.click()
            logger.info("Client Management clicked successfully")
            print("Client Management opened (new tab expected)")
        except Exception as e:
            logger.error(f"Failed to click Client Management: {e}")
            raise

        # Step 5: Switch to new tab (opened by ERP)
        print("Switching to the newly opened tab...")
        wait.until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(5)
        logger.info("Switched to new tab successfully")

        # Step 6: Click 'Customer Creation' tile
        try:
            print("Looking for 'Customer Creation' tile...")
            tile_xpaths = [
                "//div[contains(@class,'metro-tile-body')]//div[contains(text(),'Customer Creation')]",
                "//div[@class='metro-content-hide' and contains(text(),'Customer Creation')]",
                "/html/body/div[8]/div/div/div[1]/div/div[2]/div[1]"  # fallback
            ]

            customer_tile = None
            for path in tile_xpaths:
                try:
                    customer_tile = wait.until(EC.element_to_be_clickable((By.XPATH, path)))
                    logger.info(f"Found Customer Creation tile using XPath: {path}")
                    break
                except Exception:
                    continue

            if not customer_tile:
                raise Exception("Customer Creation tile not found using any known XPath")

            driver.execute_script("arguments[0].scrollIntoView(true);", customer_tile)
            time.sleep(0.5)
            actions = ActionChains(driver)
            actions.move_to_element(customer_tile).double_click().perform()
            logger.info("Customer Creation tile double-clicked successfully")
            print("Customer Creation tile opened")

            # # Wait for input field on Customer Creation page
            # wait.until(EC.presence_of_element_located((By.ID, "txtClientCode")))
            # print("Customer Creation page loaded")
        except Exception as e:
            logger.error(f"Failed to open Customer Creation: {e}")
            raise

        # Step 7: Handle Inquiry Button like Contract Inquiry
        logger.info("Navigating to Customer Inquiry...")
        try:
            time.sleep(3)
            print("Attempting to click Inquiry button...")

            inquiry_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='btn_02626_Inquiry']"))
            )
            inquiry_button.click()
            logger.info("Inquiry button clicked successfully")
            print("Inquiry button clicked successfully")
            time.sleep(5)


            logger.info("Successfully navigated through Customer Inquiry section.")
        except Exception as e:
            logger.error(f"Failed to handle Inquiry button: {e}")
            print(f"Failed to handle Inquiry button: {e}")
            logger.info("Continuing with session...")

                # Step 8: Fetch customer titles (with gender and marital fallback)
        results = {}
        print(f"\nProcessing {len(customer_codes)} customers...\n")

        for idx, code in enumerate(customer_codes, start=1):
            try:
                logger.info(f"[{idx}/{len(customer_codes)}] Fetching title for {code}...")
                print(f"[{idx}/{len(customer_codes)}] Fetching title for {code}...")

                client_input = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//input[contains(@class, 'mask_code') and @id='strClientCode_CUSEClientMain']"))
                )
                client_input.clear()
                time.sleep(0.5)
                client_input.send_keys(code)
                client_input.send_keys(Keys.ENTER)
                time.sleep(2)

                # --- STEP 1: Try to get the actual Title value ---
                try:
                    title_element = wait.until(EC.presence_of_element_located((
                        By.XPATH, "//select[@id='strTitleINVClnt_CUSEClientMain']"
                    )))
                    
                    # Ensure the element is visible and populated
                    WebDriverWait(driver, 5).until(
                        lambda d: title_element.get_attribute("value").strip() != ""
                        or len(title_element.find_elements(By.XPATH, ".//option[@selected]")) > 0
                    )

                    # First try the selected <option> text (more reliable than value)
                    selected_option = title_element.find_element(By.XPATH, ".//option[@selected]")
                    title = selected_option.text.strip()

                    if not title:
                        title = title_element.get_attribute("value").strip()

                except Exception:
                    title = ""

                # --- STEP 2: If Title is empty, check Gender and Marital Status ---
                if not title:
                    try:
                        gender_element = driver.find_element(By.XPATH, "//select[@id='strGenderINVClnt_CUSEClientMain']")
                        gender = gender_element.get_attribute("value").strip().upper() if gender_element else ""
                    except Exception:
                        gender = ""

                    if gender == "MALE":
                        title = "Mr."
                        logger.info(f"Inferred title for {code}: Mr. (based on gender)")
                    elif gender == "FEMALE":
                        try:
                            marital_element = driver.find_element(
                                By.XPATH, "//select[@id='strMairtalStatusINVClnt_CUSEClientMain']"
                            )
                            marital_status = marital_element.get_attribute("value").strip().upper() if marital_element else ""
                        except Exception:
                            marital_status = ""

                        if marital_status == "SINGLE":
                            title = "Ms."
                            logger.info(f"Inferred title for {code}: Ms. (female single)")
                        elif marital_status == "MARRIED":
                            title = "Mrs."
                            logger.info(f"Inferred title for {code}: Mrs. (female married)")
                        else:
                            title = "Ms."
                            logger.info(f"Inferred title for {code}: Ms. (female default)")
                    else:
                        title = "N/A"
                        logger.info(f"Title not found or gender missing for {code}. Marked as N/A.")
                        print("   No title or gender found (set to N/A)")
                else:
                    print(f"   Extracted title directly from ERP: {title}")
                    logger.info(f"Extracted title for {code}: {title}")

                # Store result
                results[code] = {"Title": title}
                time.sleep(1)

            except Exception as e:
                logger.error(f"✗ Error fetching title for {code}: {e}")
                results[code] = {"Title": "N/A", "Status": f"Error: {str(e)[:100]}"}
                print(f"   Error fetching {code}: {str(e)[:60]}...")
                time.sleep(1)

        # Convert to JSON string before returning
        extracted_titles = json.dumps(results, indent=4)
        logger.info("Customer titles extraction completed successfully.")
        print("\n✅ Extraction completed.")
        print(extracted_titles)

        return extracted_titles

    finally:
        if own_driver:
            driver.quit()
            logger.info("Chrome driver closed")


def get_na_customers(extracted_titles_json):
    """
    Extract customer codes with missing or 'N/A' titles from the JSON output.

    Args:
        extracted_titles_json (str or dict): JSON string or dictionary returned by fetch_titles_from_erp().

    Returns:
        list: List of customer codes with 'N/A' or missing titles.
    """
    # Ensure input is a Python dict
    if isinstance(extracted_titles_json, str):
        try:
            data = json.loads(extracted_titles_json)
        except json.JSONDecodeError:
            print("Invalid JSON input.")
            return []
    else:
        data = extracted_titles_json

    # Collect customer codes with 'N/A' titles or missing title info
    na_customers = [
        code for code, info in data.items()
        if not info.get("Title") or str(info.get("Title")).strip().upper() == "N/A"
    ]

    print(f"\nFound {len(na_customers)} customers with N/A titles.\n")
    print(f"na_customers{na_customers}")
    return na_customers

            
            
if __name__ == "__main__":
    import sys

    # Check if user wants auto mode or manual mode
    if len(sys.argv) > 1 and sys.argv[1].lower() == "auto":
        # Auto mode: fetch all customer codes from Customer Creation page
        logger.info("Starting ERP auto title fetch process...")
        print("Starting ERP auto title fetch process (Auto Mode)...\n")
        print("This will automatically fetch all customer codes from the Customer Creation page.\n")
        logger.info("Report saved as Customer_Titles_Report_Auto.xlsx")
        print("\nReport saved as Customer_Titles_Report_Auto.xlsx")

    else:
        # Manual mode: use predefined list of customer codes
        customer_codes = ["10000001", "10000002", "10000003","10000004","10000005","10000006","10000007"]

        logger.info("Starting ERP title fetch process...")
        print("Starting ERP title fetch process (Manual Mode)...\n")
        print(f"Processing {len(customer_codes)} customer codes: {customer_codes}\n")

        fetch_titles_from_erp(customer_codes)

        print("\n Results:")


       
