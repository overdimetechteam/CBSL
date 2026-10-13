import re
import shutil
import sys
import os
import time
import glob
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# -------------------------------------------------
# PATH SETUP
# -------------------------------------------------
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
sys.path.insert(0, PROJECT_ROOT)

try:
    from bots.erp_login import wait_for_loader, USERNAME, PASSWORD  # from project root
except ImportError:
    from erp_login import wait_for_loader, USERNAME, PASSWORD       # run directly inside bots/


def erp_login_closure_report():

    print("Username and password decrypted successfully")

    # -------------------------------------------------
    # CHROME SETUP
    # -------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")

    driver = webdriver.Chrome(
        service=Service(),
        options=chrome_options
    )

    wait = WebDriverWait(driver, 30)

    try:
        # -------------------------------------------------
        # LOGIN
        # -------------------------------------------------
        print("Opening ERP...")
        driver.get("https://erp.assetline.lk/")
        wait_for_loader(driver)

        wait.until(EC.presence_of_element_located((By.ID, "strUserName"))).send_keys(USERNAME)
        wait.until(EC.presence_of_element_located((By.ID, "strPassword"))).send_keys(PASSWORD)

        wait.until(EC.element_to_be_clickable((By.ID, "idSignIn"))).click()
        wait_for_loader(driver)

        # wait.until(EC.element_to_be_clickable(
        #     (By.CSS_SELECTOR, "button.btn.blue.btn-circle"))
        # ).click()
        # wait_for_loader(driver)

        print(" Login successful")

        # -------------------------------------------------
        # OPEN REPORT WINDOW
        # -------------------------------------------------
        main_tab = driver.current_window_handle

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[normalize-space()='Report']"
        ))).click()

        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(
            [h for h in driver.window_handles if h != main_tab][0]
        )

        print(" Switched to Reports window")

        # -------------------------------------------------
        # CLICK REPORTS TILE
        # -------------------------------------------------
        reports_tile = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//div[contains(@class,'metro-tile') and .//div[normalize-space()='Reports']]"
        )))
        reports_tile.click()
        wait_for_loader(driver)

        # -------------------------------------------------
        # CLICK "Legal Bal New" (without Internal)
        # -------------------------------------------------
        element = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "/html/body/div[8]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div[2]/div/div/div/i"
            ))
        )

        driver.execute_script("arguments[0].click();", element)

        # -------------------------------------------------
        # CLICK INQUIRY
        # -------------------------------------------------
        inquiry_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//span[normalize-space()='Inquiry']"
        )))
        inquiry_btn.click()
        wait_for_loader(driver)

        print(" Inquiry opened")

        # -------------------------------------------------
        # GENERATE REPORT
        # -------------------------------------------------
        generate_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[contains(@id,'GenerateReport')]"
        )))



        # -------------------------------------------------
        # SNAPSHOT existing files BEFORE generating
        # -------------------------------------------------
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")

        existing_files = set(glob.glob(os.path.join(downloads_dir, "*.xlsx")) +
                             glob.glob(os.path.join(downloads_dir, "*.xls")))

        print(f" Existing files before generate: {len(existing_files)}")

        generate_btn.click()
        print(" Generating report...")

        # -------------------------------------------------
        # WAIT FOR A BRAND NEW FILE to appear
        # -------------------------------------------------
        timeout = time.time() + 90
        downloaded_file = None

        while time.time() < timeout:
            current_files = set(glob.glob(os.path.join(downloads_dir, "*.xlsx")) +
                                glob.glob(os.path.join(downloads_dir, "*.xls")))

            new_files = [
                f for f in (current_files - existing_files)
                if not os.path.basename(f).startswith("~$")
                and not f.endswith(".crdownload")
            ]

            if new_files:
                downloaded_file = max(new_files, key=os.path.getctime)
                break

            time.sleep(2)



        # -------------------------------------------------
        # CONFIRM DOWNLOAD & MOVE TO outputs/LGD/.../firstDefault/
        # -------------------------------------------------
        if not downloaded_file:
            raise RuntimeError("No valid file found within timeout")

        print(f"Downloaded: {os.path.basename(downloaded_file)}")

        # Resolve outputs/LGD/<latest-date-folder>/firstDefault/
        outputs_lgd = Path(PROJECT_ROOT) / "outputs" / "LGD"
        dates_file  = outputs_lgd / "completed_dates.txt"
        lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
        last_date = lines[-1]

        date_variants = {last_date}
        for fmt_in, fmt_out in [("%d-%m-%Y", "%d/%m/%Y"), ("%d/%m/%Y", "%d-%m-%Y")]:
            try:
                from datetime import datetime as _dt
                date_variants.add(_dt.strptime(last_date, fmt_in).strftime(fmt_out))
            except ValueError:
                pass

        candidates = []
        for child in outputs_lgd.iterdir():
            if not child.is_dir():
                continue
            if any(child.name.startswith(v) for v in date_variants):
                m = re.search(r"\((\d+)\)$", child.name)
                candidates.append((int(m.group(1)) if m else 0, child))

        if not candidates:
            raise RuntimeError(f"No date folder found for '{last_date}' in {outputs_lgd}")

        candidates.sort(key=lambda x: x[0], reverse=True)
        dest_folder = candidates[0][1] / "firstDefault"
        dest_folder.mkdir(parents=True, exist_ok=True)

        # Remove existing files matching the pattern before moving
        for old in dest_folder.glob("Legal Bal New  - Internal *.xlsx"):
            old.unlink()
            print(f"  Removed existing: {old.name}")

        dest_path = dest_folder / os.path.basename(downloaded_file)
        shutil.move(downloaded_file, str(dest_path))
        print(f"Moved to: {dest_path}")

        return driver

    except Exception as e:
        try:
            driver.save_screenshot("error.png")
        except:
            print(" Browser already closed")

        driver.quit()
        raise Exception(f" Automation failed: {e}")


# -------------------------------------------------
# MAIN
# -------------------------------------------------
if __name__ == "__main__":
    print("=" * 60)
    print("Legal Bal New Report")
    print("=" * 60)

    driver = erp_login_closure_report()

    print("\n Process completed")
    time.sleep(5)
    driver.quit()