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


def _find_lgd_firstdefault_folder() -> Path | None:
    """Resolve outputs/LGD/<latest-date-folder>/firstDefault/ from completed_dates.txt."""
    outputs_lgd = Path(PROJECT_ROOT) / "outputs" / "LGD"
    dates_file  = outputs_lgd / "completed_dates.txt"
    if not dates_file.exists():
        return None
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    if not lines:
        return None
    last_date = lines[-1]

    date_variants = {last_date}
    from datetime import datetime
    for fmt_in, fmt_out in [("%d-%m-%Y", "%d/%m/%Y"), ("%d/%m/%Y", "%d-%m-%Y")]:
        try:
            date_variants.add(datetime.strptime(last_date, fmt_in).strftime(fmt_out))
        except ValueError:
            pass

    candidates: list[tuple[int, Path]] = []
    for child in outputs_lgd.iterdir():
        if not child.is_dir():
            continue
        if any(child.name.startswith(v) for v in date_variants):
            m = re.search(r"\((\d+)\)$", child.name)
            candidates.append((int(m.group(1)) if m else 0, child))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    dest = candidates[0][1] / "firstDefault"
    dest.mkdir(parents=True, exist_ok=True)
    return dest


def erp_all_contracts_sector_report():

    # Credentials already decrypted by erp_login at import time
    print(" Credentials loaded")

    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    pre_files = set(glob.glob(os.path.join(downloads_dir, "*.xlsx")) +
                    glob.glob(os.path.join(downloads_dir, "*.xls")))

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory":   downloads_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade":   True,
        "safebrowsing.enabled":         True,
    })

    driver = webdriver.Chrome(service=Service(), options=chrome_options)
    wait   = WebDriverWait(driver, 30)

    # -------------------------------------------------
    # HELPERS
    # -------------------------------------------------
    def safe_click(el):
        try:
            el.click()
        except:
            driver.execute_script("arguments[0].click();", el)

    def switch_into_iframe():
        driver.switch_to.default_content()
        time.sleep(1)
        for f in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.frame(f)
                return True
            except:
                driver.switch_to.default_content()
        return False

    def check_row(field_name):
        """Tick the checkbox in the row whose second <td> exactly matches field_name."""
        print(f"  Selecting: {field_name}")
        row = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//tr[td[normalize-space()='{field_name}']]"
        )))
        cb = row.find_element(By.XPATH, ".//td[1]//input[@type='checkbox']")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
        time.sleep(0.3)
        if not cb.is_selected():
            try:
                cb.click()
            except:
                driver.execute_script("arguments[0].click();", cb)
        print(f"   Selected: {field_name}")

    try:
        # -------------------------------------------------
        # LOGIN
        # -------------------------------------------------
        driver.get("https://erp.assetline.lk/")
        wait_for_loader(driver)

        wait.until(EC.presence_of_element_located((By.ID, "strUserName"))).send_keys(USERNAME)
        wait.until(EC.presence_of_element_located((By.ID, "strPassword"))).send_keys(PASSWORD)
        safe_click(wait.until(EC.element_to_be_clickable((By.ID, "idSignIn"))))
        wait_for_loader(driver)

        # safe_click(wait.until(EC.element_to_be_clickable((
        #     By.CSS_SELECTOR, "button.btn.blue.btn-circle"
        # ))))
        # wait_for_loader(driver)
        print(" Login success")

        # -------------------------------------------------
        # OPEN REPORT WINDOW
        # -------------------------------------------------
        main_tab = driver.current_window_handle
        safe_click(wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[normalize-space()='Report']"
        ))))
        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window([h for h in driver.window_handles if h != main_tab][0])
        wait_for_loader(driver)
        switch_into_iframe()

        # -------------------------------------------------
        # NAVIGATION → Reports → All Contracts Sector Report → Inquiry
        # -------------------------------------------------
        safe_click(wait.until(EC.element_to_be_clickable((By.XPATH, "//div[text()='Reports']"))))
        wait_for_loader(driver)

        safe_click(wait.until(EC.element_to_be_clickable((
            By.XPATH, "//div[contains(text(),'All Contracts Sector Report')]"
        ))))
        wait_for_loader(driver)

        safe_click(wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Inquiry']"))))
        wait_for_loader(driver)

        # -------------------------------------------------
        # FILTER PANEL
        # -------------------------------------------------
        switch_into_iframe()

        # Click the "Select All" toggle (bootstrap-switch ON handle)
        safe_click(wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//span[contains(@class,'bootstrap-switch-handle-on') and normalize-space()='Select All']"
        ))))
        print(" Select All clicked")

        # Wait for the filter table rows to appear
        wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[td]")))
        time.sleep(1)

        # Tick CONTRACT_NO and INTEREST_RATE
        check_row("CONTRACT_NO")
        check_row("INTEREST_RATE")
        print(" Columns selected")

        # -------------------------------------------------
        # GENERATE REPORT
        # -------------------------------------------------
        generate_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(@id,'GenerateReport')]"
        )))
        driver.execute_script("arguments[0].click();", generate_btn)
        print(" Generating report...")

        # Wait for the SUCCESS modal to appear
        wait.until(EC.presence_of_element_located((
            By.XPATH, "//h4[@id='ModalHeaderText' and normalize-space()='SUCCESS !']"
        )))
        print(" SUCCESS modal appeared — download triggered")

        # -------------------------------------------------
        # WAIT FOR NEW EXCEL FILE IN DOWNLOADS
        # -------------------------------------------------
        print(" Waiting for download...")
        excel_file = None
        deadline = time.time() + 300  # 5-minute hard cap
        while time.time() < deadline:
            candidates = [
                f for f in glob.glob(os.path.join(downloads_dir, "*.xlsx"))
                + glob.glob(os.path.join(downloads_dir, "*.xls"))
                if f not in pre_files
                and not os.path.basename(f).startswith("~$")
                and not f.endswith(".crdownload")
                and not f.endswith(".tmp")
            ]
            if candidates:
                excel_file = max(candidates, key=os.path.getctime)
                break
            time.sleep(2)

        if not excel_file:
            raise RuntimeError("Download timed out — no new Excel file found in Downloads")

        print(f" Downloaded: {os.path.basename(excel_file)}")

        # -------------------------------------------------
        # MOVE TO outputs/LGD/<date>/firstDefault/
        # -------------------------------------------------
        dest_folder = _find_lgd_firstdefault_folder()
        if not dest_folder:
            raise RuntimeError("Could not resolve outputs/LGD/<date>/firstDefault/ folder")

        # Remove any stale copy matching the glob pattern
        for old in dest_folder.glob("All Contracts Sector Report - Internal *.xlsx"):
            old.unlink()
            print(f"  Removed existing: {old.name}")

        dest_path = dest_folder / os.path.basename(excel_file)
        shutil.move(excel_file, str(dest_path))
        print(f" Moved to: {dest_path}")

        return True

    except Exception as e:
        try:
            driver.save_screenshot("error.png")
        except:
            pass
        raise Exception(f"Failed: {e}")
    finally:
        try:
            driver.quit()
        except:
            pass


# -------------------------------------------------
# MAIN
# -------------------------------------------------
if __name__ == "__main__":
    ok = erp_all_contracts_sector_report()
    print(" DONE" if ok else "✗ FAILED")