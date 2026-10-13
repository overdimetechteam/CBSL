import os
import sys
import glob
import shutil
import time
from datetime import datetime, date
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
from erp_login import erp_login
# ==========================
# PATH SETUP
# ==========================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
REPORTS_DIR = os.path.join(PROJECT_ROOT, "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
INPUTS_DIR = r"C:\Imparement\inputs"
os.makedirs(INPUTS_DIR, exist_ok=True)


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
        print(f"❌ Inquiry button not found: {e}")

# ==========================
# SET ACTIVATION DATE FILTER
# ==========================
# Report window (DD/MM/YYYY). The window is the 3 months ending on REFINANCE_END_DATE
# (start = 1st day of the month 2 months before the end date), so for 30/09/2025 it is
# 01/07/2025 to 30/09/2025. Set REFINANCE_START_DATE (or the REFINANCE_START_DATE
# environment variable / run() argument) only to force a different start.
REFINANCE_START_DATE = None
REFINANCE_END_DATE = "30/09/2025"


def set_activation_date_filter(driver, wait, end_date_str: str, start_date_str: str = None):
    """
    Select BETWEEN operator and enter date range for NEW_CON_NO_ACTIVATION_DATE.

    Args:
        end_date_str:   End date in DD/MM/YYYY format (e.g. "30/09/2025").
        start_date_str: Start date in DD/MM/YYYY format. If omitted, it is the
                        1st day of the month 2 months before the end date.
    """
    if not start_date_str:
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
        print(f"❌ Activation date filter failed: {e}")

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

        # Wait for modal — use a long timeout as report generation can take a while
        print("Waiting for download confirmation modal...")
        modal_wait = WebDriverWait(driver, 300)
        modal_wait.until(
            EC.text_to_be_present_in_element(
                (By.ID, "ModalDescriptionText"), "Report Successfully Generated"
            )
        )
        print("✓ Modal appeared — download started")

        # Poll Downloads for the file (up to 300s), staying alive so Chrome doesn't close
        print("Waiting for file to arrive in Downloads folder...")
        downloaded_path = None
        for _ in range(300):
            matches = glob.glob(os.path.join(DOWNLOAD_DIR, "Refinance DPD Report *.xlsx"))
            complete = [f for f in matches if not f.endswith(".crdownload")]
            if complete:
                downloaded_path = max(complete, key=os.path.getmtime)
                break
            time.sleep(1)

        if not downloaded_path:
            print("❌ Download timed out — file not found in Downloads folder")
            return False

        print(f"✓ File downloaded: {downloaded_path}")
        return True

    except Exception as e:
        print(f"❌ Generate Report failed: {e}")
        # Show what the ERP was actually saying when the wait ran out.
        try:
            print(f"   ERP modal text at failure: {driver.find_element(By.ID, 'ModalDescriptionText').text!r}")
        except Exception:
            print("   ERP modal text at failure: (no modal found)")
        try:
            driver.save_screenshot("refinance_generate_error.png")
            print("   Screenshot saved: refinance_generate_error.png")
        except Exception:
            pass
        return False


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
            print(f"❌ Expected 2 date fields but found {len(date_inputs)}")

        wait_for_loader(driver)

    except Exception as e:
        print(f"❌ Date entry failed: {e}")


def _norm_date(value: str) -> str:
    """Normalise to strictly zero-padded DD/MM/YYYY regardless of what was passed in."""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%-d/%-m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return value


def run(date_str: str = None, start_str: str = None) -> bool:
    date_str = _norm_date((date_str or os.environ.get("REFINANCE_END_DATE") or REFINANCE_END_DATE).strip())
    start_str = (start_str or os.environ.get("REFINANCE_START_DATE") or REFINANCE_START_DATE or "").strip()
    start_str = _norm_date(start_str) if start_str else None   # None = 3 months ending on date_str
    print(f"Refinance DPD Report window: {start_str or '3 months'} to {date_str}")

    driver, wait = erp_login()
    result = False
    try:
        click_finance(driver, wait)
        click_finance_tile(driver, wait)
        click_inquiry(driver, wait)
        set_activation_date_filter(driver, wait, end_date_str=date_str, start_date_str=start_str)
        result = click_generate_report(driver, wait) or False
    except Exception as e:
        print(f"❌ Error during report generation: {e}")
    driver.quit()
    print("✓ Refinance DPD download completed")
    return result
    


# ==========================
# MAIN
# ==========================
if __name__ == "__main__":
    # Usage: python LGD_refinance_dpd.py [END DD/MM/YYYY] [START DD/MM/YYYY]
    run(sys.argv[1] if len(sys.argv) > 1 else None,
        sys.argv[2] if len(sys.argv) > 2 else None)