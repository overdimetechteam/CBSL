import os
import time
from pathlib import Path
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from cryptography.fernet import Fernet


# ==========================
# PATH SETUP
# ==========================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
REPORTS_DIR = os.path.join(PROJECT_ROOT, "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

# ==========================
# LOAD CREDENTIALS (lazy)
# ==========================
ENV_PATH = os.path.join(PROJECT_ROOT, ".env")

def _load_credentials():
    """Load .env and decrypt credentials at call time."""
    load_dotenv(dotenv_path=ENV_PATH, override=True)

    secret_key = os.environ.get("FERNET_KEY", "").strip()
    enc_username = os.getenv("ENC_USERNAME")
    enc_password = os.getenv("ENC_PASSWORD")

    if not secret_key:
        raise Exception("❌ FERNET_KEY environment variable not set")
    if not enc_username or not enc_password:
        raise Exception("❌ Missing ENC_USERNAME or ENC_PASSWORD in .env file")

    fernet = Fernet(secret_key.encode())
    username = fernet.decrypt(enc_username.encode()).decode()
    password = fernet.decrypt(enc_password.encode()).decode()
    print("✓ Credentials decrypted successfully")
    return username, password

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
    USERNAME, PASSWORD = _load_credentials()

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")

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

        print("Clicking select button...")
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.blue.btn-circle"))).click()
        wait_for_loader(driver)

        print("✓ Login successful!")
        return driver, wait

    except Exception as e:
        driver.save_screenshot("login_error.png")
        driver.quit()
        raise Exception(f"❌ ERP login failed: {e}")

# ==========================
# NAVIGATE TO FINANCE PAGE
# ==========================
def click_finance(driver, wait):
    print("Navigating to Finance page...")
    driver.get("https://erp.assetline.lk/Application/Home/FINANCE")
    wait_for_loader(driver)
    time.sleep(3)
    print("✓ Finance page loaded:", driver.current_url)

# ==========================
# CLICK GL ACCOUNT VIEW
# ==========================
def click_gl_account_view(driver, wait):
    print("Opening GL Account View...")

    # DEBUG: Find all elements containing 'GL Account View'
    all_elements = driver.find_elements(
        By.XPATH,
        "//*[contains(text(),'GL Account View') or contains(@innerHTML,'GL Account View')]"
    )
    print(f"Elements containing 'GL Account View': {len(all_elements)}")
    for i, el in enumerate(all_elements):
        print(f"  [{i}] tag={el.tag_name} | class='{el.get_attribute('class')}' | displayed={el.is_displayed()} | text='{el.text}'")

    # Try clicking each match via JS
    for el in all_elements:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", el)
            wait_for_loader(driver)
            time.sleep(2)
            print("✓ Clicked GL Account View element")
            print("URL after click:", driver.current_url)
            print("Page title after click:", driver.title)
            return
        except Exception as e:
            print(f"  Click failed on [{el.tag_name}]: {e}")

    print("❌ Could not find or click GL Account View")

# ==========================
# CLICK INQUIRY BUTTON
# ==========================
def click_inquiry(driver, wait):
    print("Clicking Inquiry button...")

    time.sleep(2)

    try:
        inquiry_btn = wait.until(
            EC.presence_of_element_located((By.ID, "btn_02366_Inquiry"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inquiry_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", inquiry_btn)
        wait_for_loader(driver)
        print("✓ Inquiry button clicked successfully")

    except Exception as e:
        print(f"❌ Inquiry button not found: {e}")

# ==========================
# ENTER GL CODE
# ==========================
def enter_gl_code(driver, wait, gl_code="1173"):
    print(f"Entering GL code: {gl_code}...")

    try:
        gl_input = wait.until(
            EC.presence_of_element_located((By.ID, "strSubLedgerCode_FINGLAccountView"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", gl_input)
        time.sleep(0.5)

        # Clear and enter GL code
        gl_input.clear()
        gl_input.send_keys(gl_code)

        # Trigger change events so ERP registers the value
        driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", gl_input)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', {bubbles:true}));", gl_input)

        wait_for_loader(driver)
        print(f"✓ GL code entered: {gl_code}")

    except Exception as e:
        print(f"❌ GL code entry failed: {e}")


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

                driver.execute_script("arguments[0].removeAttribute('readonly');", inp)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
                time.sleep(0.3)

                # Set value purely via JS — no send_keys to avoid double-entry
                driver.execute_script(
                    "arguments[0].value = '';"
                    "arguments[0].value = arguments[1];"
                    "arguments[0].dispatchEvent(new Event('input',  {bubbles:true}));"
                    "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));"
                    "arguments[0].dispatchEvent(new Event('blur',   {bubbles:true}));",
                    inp, val
                )

                time.sleep(0.5)
                actual = inp.get_attribute('value')
                print(f"✓ {label} date entered: {val} | Field shows: {actual}")

        else:
            print(f"❌ Expected 2 date fields but found {len(date_inputs)}")

        wait_for_loader(driver)

    except Exception as e:
        print(f"❌ Date entry failed: {e}")

 # ==========================
# CLICK SEARCH BUTTON
# ==========================
def click_search(driver, wait):
    print("Clicking Search button...")

    try:
        search_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "btn_02366_BtnSearch_FINGLAccountView"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", search_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", search_btn)
        wait_for_loader(driver)
        print("✓ Search button clicked successfully")

    except Exception as e:
        print(f"❌ Search button not found: {e}")

        # ==========================
# SCROLL PAGE DOWN
# ==========================
def scroll_page(driver, wait):
    print("Scrolling page down...")

    try:
        time.sleep(2)  # wait for search results to load

        # Scroll down smoothly to bottom
        driver.execute_script("""
            window.scrollTo({
                top: document.body.scrollHeight,
                behavior: 'smooth'
            });
        """)
        time.sleep(1)
        print("✓ Scrolled to bottom successfully")

    except Exception as e:
        print(f"❌ Scroll failed: {e}")

# ==========================
# GET CLOSING BALANCE
# ==========================
def get_closing_balance(driver, wait):
    print("Reading Closing Balance...")

    try:
        time.sleep(2)  # wait for results to fully render

        closing_balance_input = wait.until(
            EC.presence_of_element_located((By.ID, "intClossingBalance_FINGLAccountView"))
        )

        # Field is disabled — use JS to read value
        value = driver.execute_script("return arguments[0].value;", closing_balance_input)

        if value and value.strip():
            print(f"✓ Closing Balance: {value.strip()}")
            return value.strip()
        else:
            print("⚠️ Closing Balance field is empty")
            return None

    except Exception as e:
        print(f"❌ Could not read Closing Balance: {e}")
        return None
# ==========================
# DOWNLOAD EXCEL
# ==========================
def _find_latest_outputs_folder(outputs_subdir: "Path") -> "Path | None":
    """Return the highest-version date folder from outputs_subdir using completed_dates.txt."""
    import re as _re
    from pathlib import Path as _Path

    dates_file = _Path(outputs_subdir) / "completed_dates.txt"
    if not dates_file.exists():
        print(f"❌ completed_dates.txt not found in {outputs_subdir}")
        return None

    lines = [l.strip() for l in dates_file.read_text().splitlines() if l.strip()]
    if not lines:
        print(f"❌ completed_dates.txt is empty in {outputs_subdir}")
        return None

    last_date = lines[-1]
    candidates = []
    for child in _Path(outputs_subdir).iterdir():
        if not child.is_dir():
            continue
        if child.name == last_date or child.name.startswith(f"{last_date}("):
            m = _re.search(r'\((\d+)\)$', child.name)
            candidates.append((int(m.group(1)) if m else 0, child))

    if not candidates:
        print(f"❌ No outputs folder found for date '{last_date}' in {outputs_subdir}")
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def download_excel(driver, wait, download_dir: str, dest_outputs_name: str = "NBD_WF_18_ECS") -> "Path | None":
    """Click Excel Download button, wait for file, move to the latest date folder
    under outputs/<dest_outputs_name>/.

    Args:
        download_dir:       Chrome's download directory (usually ~/Downloads).
        dest_outputs_name:  Name of the subfolder under outputs/ to move the file into.
                            Defaults to "NBD_WF_18_ECS" to preserve existing ECS behaviour.
    """
    print("Waiting for Excel Download button...")
    try:
        excel_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='btn_02366_ExcelDownload']")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", excel_btn)
        time.sleep(0.5)

        download_path = Path(download_dir)
        pre_files = set(f for f in download_path.glob("*") if f.is_file())

        driver.execute_script("arguments[0].click();", excel_btn)
        print("✓ Excel Download button clicked — waiting for file...")

        # Wait for new xlsx file
        timeout = 120
        start = time.time()
        downloaded = None
        while time.time() - start < timeout:
            current = set(f for f in download_path.glob("*") if f.is_file())
            new = current - pre_files
            for f in new:
                if f.suffix == ".xlsx" and not f.name.endswith(".crdownload"):
                    s1 = f.stat().st_size
                    time.sleep(1)
                    s2 = f.stat().st_size
                    if s1 == s2 and s1 > 0:
                        downloaded = f
                        break
            if downloaded:
                break
            time.sleep(1)

        if not downloaded:
            print("❌ Excel file not downloaded within timeout")
            return None

        print(f"✓ Downloaded: {downloaded.name}")

        # Resolve destination outputs folder
        root = Path(__file__).resolve().parent.parent
        outputs_subdir = root / "outputs" / dest_outputs_name
        dest_folder = _find_latest_outputs_folder(outputs_subdir)
        if not dest_folder:
            print(f"❌ Could not resolve destination folder under outputs/{dest_outputs_name}")
            return downloaded

        # Remove existing *GL Account Details.xlsx before moving
        for old in dest_folder.glob("*GL Account Details.xlsx"):
            try:
                old.unlink()
                print(f"Removed existing: {old.name}")
            except Exception as e:
                print(f"Warning: could not remove {old.name}: {e}")

        dest = dest_folder / downloaded.name
        downloaded.replace(dest)
        print(f"✓ Moved to: {dest}")
        return dest

    except Exception as e:
        print(f"❌ Excel download failed: {e}")
        return None


# ==========================
# RUN (callable from other scripts)
# ==========================
def run(start_date_str: str, end_date_str: str = None, gl_code: str = "1173",
        dest_outputs_name: str = "NBD_WF_18_ECS") -> None:
    """
    Run the GL CodeBase report for the given date range.

    Args:
        start_date_str:     Start date in DD/MM/YYYY format.
        end_date_str:       End date in DD/MM/YYYY format. Defaults to start_date_str.
        gl_code:            GL account code to query (default "1173" for ECS).
        dest_outputs_name:  Subfolder under outputs/ to move the downloaded file into
                            (default "NBD_WF_18_ECS"; use "Supporting Schedules" for that report).
    """
    if end_date_str is None:
        end_date_str = start_date_str

    download_dir = str(Path.home() / "Downloads")

    driver, wait = erp_login()
    try:
        click_finance(driver, wait)
        click_gl_account_view(driver, wait)
        click_inquiry(driver, wait)
        enter_gl_code(driver, wait, gl_code=gl_code)
        enter_date_range(driver, wait, from_date=start_date_str, to_date=end_date_str)

        # Click Search
        search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='btn_02366_BtnSearch_FINGLAccountView']")))
        driver.execute_script("arguments[0].click();", search_btn)
        wait_for_loader(driver)
        print("✓ Search clicked")

        download_excel(driver, wait, download_dir, dest_outputs_name=dest_outputs_name)

    finally:
        driver.quit()
        print("✓ GL CodeBase Report completed")


# ==========================
# MAIN
# ==========================
if __name__ == "__main__":
    import sys
    start    = sys.argv[1] if len(sys.argv) > 1 else "01/12/2025"
    end      = sys.argv[2] if len(sys.argv) > 2 else start
    gl       = sys.argv[3] if len(sys.argv) > 3 else "1173"
    dest_out = sys.argv[4] if len(sys.argv) > 4 else "NBD_WF_18_ECS"
    run(start, end, gl_code=gl, dest_outputs_name=dest_out)