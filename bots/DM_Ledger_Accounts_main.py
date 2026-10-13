try:
    from bots.erp_login import erp_login  # when imported from project root (e.g. scheduler.py)
except ImportError:
    from erp_login import erp_login       # when run directly inside bots/
import glob as _glob
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
from selenium.webdriver.support.ui import Select

_DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

# Report window sent to the ERP: 'dd/mm/yyyy'. The start stays at the beginning of
# the Write Off Working period; the end is the month-end being run. Override the
# end date with the LEDGER_END_DATE environment variable, a command-line argument,
# or Download_Ledger_Accounts_report(end_date="dd/mm/yyyy").
LEDGER_START_DATE = "01/04/2025"
LEDGER_END_DATE = "30/09/2025"

def Download_Ledger_Accounts_report(end_date=None):
    end_date = (end_date or os.environ.get("LEDGER_END_DATE") or LEDGER_END_DATE).strip()
    print(f"Ledger Details Report window: {LEDGER_START_DATE} to {end_date}")
    driver, wait = erp_login()
    time.sleep(5)
    # Click on Reports
    driver.get("https://erp.assetline.lk/Application/Home/REPORT")
    time.sleep(10)

    print("Clicked on Reports")
    # Switch to the new tab that opened after the click
    driver.switch_to.window(driver.window_handles[-1])
    # Wait for the new tab's page to fully load before interacting
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(5)
    print("Switched to new tab")
    # Save screenshot after switching tab
    driver.save_screenshot("screenshot_new_tab.png")
    wait = WebDriverWait(driver, 30)
    print("Screenshot saved")

    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[1]/span/a"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[1]/input"))).send_keys('Ledger Details Report - Internal')

    print("Searched for Ledger Details Report - Internal")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[2]/div[2]/div[1]/ul/li/a"))).click()
    time.sleep(5)
    print("Clicked on Ledger Details Report - Internal")
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/div[1]/div/div[1]"))).click()
    print("Clicked on Download button")
    time.sleep(5)

    dropdown = Select(driver.find_element("xpath", "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[1]/td[3]/select"))
    dropdown.select_by_visible_text("IN")
    print("Selected IN from dropdown")
    time.sleep(5)
    coa =wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[1]/td[4]/input")))
    coa.clear()
    coa.send_keys("'1562','1522','0130'")
    print("Entered values in input field")
    time.sleep(5)
    date_range = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[7]/td[4]/input")))
    date_range.clear()
    date_range.send_keys(f"'{LEDGER_START_DATE}','{end_date}'")
    print("Entered date range in input field")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[1]/div[5]/fieldset/div/div/button"))).click()
    print("Waiting for Ledger Details Report download to complete...")
    downloaded_path = None
    for _ in range(300):
        matches = _glob.glob(os.path.join(_DOWNLOAD_DIR, "Ledger Details Report - Internal*.xlsx"))
        complete = [f for f in matches if not f.endswith(".crdownload")]
        if complete:
            downloaded_path = complete[0]
            break
        time.sleep(1)
    if not downloaded_path:
        print(" Download timed out — file not found in Downloads folder")
        driver.quit()
        return False
    print(f"✓ File downloaded: {downloaded_path}")
    driver.quit()
    return True

if __name__ == "__main__":
    import sys
    Download_Ledger_Accounts_report(sys.argv[1] if len(sys.argv) > 1 else None)