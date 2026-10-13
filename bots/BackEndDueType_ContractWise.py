import sys
import os
import time
import glob
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# -------------------------------------------------
# PATH SETUP
# -------------------------------------------------
CURRENT_DIR  = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
sys.path.insert(0, PROJECT_ROOT)

try:
    from bots.erp_login import wait_for_loader, USERNAME, PASSWORD  # when imported from project root
except ImportError:
    from erp_login import wait_for_loader, USERNAME, PASSWORD       # when run directly inside bots/

# -------------------------------------------------
# REPORT NAME
# -------------------------------------------------
REPORT_NAME = "Backend Due Types - Contract wise Report - Internal"

# -------------------------------------------------
# MAIN FUNCTION
# -------------------------------------------------
# Report window (YYYY-MM-DD). write_off.py rebuilds the 0130-GL_Working sheet
# entirely from this export, so it must be cumulative: a month-to-date export would
# wipe the earlier months. The start is the first transaction date already in
# the Write Off Working; the end is the month-end being run. Override with the
# BACKEND_START_DATE / BACKEND_END_DATE environment variables or the arguments.
BACKEND_START_DATE = "2025-03-31"
BACKEND_END_DATE   = "2025-09-30"


def erp_all_contracts_sector_report(from_date: str = None, to_date: str = None):
    """
    Run the Backend Due Types bot.

    Args:
        from_date: First date in YYYY-MM-DD format. Defaults to BACKEND_START_DATE.
        to_date:   Last date in YYYY-MM-DD format. Defaults to BACKEND_END_DATE.
    """
    if from_date is None:
        from_date = os.environ.get("BACKEND_START_DATE") or BACKEND_START_DATE
    if to_date is None:
        to_date = os.environ.get("BACKEND_END_DATE") or BACKEND_END_DATE

    FROM_DATE = datetime.strptime(from_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    TO_DATE   = datetime.strptime(to_date,   "%Y-%m-%d").strftime("%d/%m/%Y")
    # Combined criteria string for BETWEEN operator: '01/04/2025','30/04/2025'
    COMBINED_DATES = f"'{FROM_DATE}','{TO_DATE}'"

    print(f"FROM          : {FROM_DATE}")
    print(f"TO            : {TO_DATE}")
    print(f"CRITERIA VALUE: {COMBINED_DATES}")

    driver = None

    try:
        # Credentials are decrypted by erp_login at import time
        print("✓ Credentials loaded")

        # -------------------------------------------------
        # DOWNLOAD LOCATION
        # -------------------------------------------------
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")

        prefs = {
            "download.default_directory":   downloads_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade":   True,
            "safebrowsing.enabled":         True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

        wait = WebDriverWait(driver, 30)

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
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            for f in frames:
                try:
                    driver.switch_to.frame(f)
                    return True
                except:
                    driver.switch_to.default_content()
            return False

        def select_checkbox(field_name):
            print(f"➡ Selecting: {field_name}")

            row = wait.until(EC.presence_of_element_located((
                By.XPATH,
                f"//tr[td[contains(text(),'{field_name}')]]"
            )))

            checkbox = row.find_element(
                By.XPATH, ".//td[1]//input[@type='checkbox']"
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", checkbox
            )

            if not checkbox.is_selected():
                try:
                    checkbox.click()
                except:
                    driver.execute_script("arguments[0].click();", checkbox)

            print(f"Selected: {field_name}")

        def fill_criteria_input(xpath, value):
            """
            Type value character by character into criteria input.
            Handles special chars like single quotes and commas correctly.
            Falls back to JS direct set if send_keys drops any character.
            """
            inp = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", inp
            )
            driver.execute_script("arguments[0].focus();", inp)

            # Clear completely
            driver.execute_script("arguments[0].value = '';", inp)
            inp.send_keys(Keys.CONTROL + "a")
            inp.send_keys(Keys.DELETE)
            time.sleep(0.3)

            # Type character by character so quotes and commas register
            for char in value:
                inp.send_keys(char)
                time.sleep(0.05)

            time.sleep(0.3)

            # Verify actual value in field
            actual = driver.execute_script("return arguments[0].value;", inp)
            print(f"    Field value: {actual}")

            if actual != value:
                print(f"   Mismatch — using JS fallback to force value")
                driver.execute_script(
                    "arguments[0].value = arguments[1];", inp, value
                )
                driver.execute_script("""
                    var el = arguments[0];
                    el.dispatchEvent(new Event('input',  { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
                """, inp)
                actual = driver.execute_script("return arguments[0].value;", inp)
                print(f"    Field value after JS fallback: {actual}")

        # -------------------------------------------------
        # SEARCH + CLICK REPORT HELPER
        # -------------------------------------------------
        def search_and_click_report():

            XPATHS_TO_TRY = [
                f"//div[contains(text(),'{REPORT_NAME}')]",
                f"//li[contains(text(),'{REPORT_NAME}')]",
                f"//span[contains(text(),'{REPORT_NAME}')]",
                f"//td[contains(text(),'{REPORT_NAME}')]",
                f"//a[contains(text(),'{REPORT_NAME}')]",
                "//div[contains(text(),'Backend Due Types')]",
                "//li[contains(text(),'Backend Due Types')]",
                "//span[contains(text(),'Backend Due Types')]",
                "//td[contains(text(),'Backend Due Types')]",
                "//a[contains(text(),'Backend Due Types')]",
            ]

            # Step 1 — Click magnifier icon
            try:
                magnifier = wait.until(EC.presence_of_element_located((
                    By.CSS_SELECTOR, "i.icon-magnifier"
                )))
                driver.execute_script("arguments[0].click();", magnifier)
                time.sleep(1.5)
                print(" Magnifier icon clicked")
            except Exception as e:
                print(f"Magnifier click skipped: {e}")

            # Step 2 — Locate search input
            search_box = wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR, "input.form-control[placeholder='Search']"
            )))

            # Step 3 — Clear via JS + fire events
            driver.execute_script("arguments[0].focus();", search_box)
            driver.execute_script("arguments[0].value = '';", search_box)
            driver.execute_script("""
                var el = arguments[0];
                el.dispatchEvent(new Event('input',  { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
            """, search_box)
            time.sleep(0.5)

            # Step 4 — Type search term
            driver.execute_script("arguments[0].value = '';", search_box)
            search_box.send_keys("Backend Due Types")
            time.sleep(2.5)

            print(" Typed search term — scanning for results...")

            # Strategy A — inside current iframe
            for xpath in XPATHS_TO_TRY:
                try:
                    el = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", el
                    )
                    driver.execute_script("arguments[0].click();", el)
                    print(f"Report clicked (iframe) via: {xpath}")
                    return True
                except:
                    continue

            # Strategy B — switch to default content and retry
            print("Not found in iframe — trying default content...")
            driver.switch_to.default_content()
            time.sleep(1)

            for xpath in XPATHS_TO_TRY:
                try:
                    el = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", el
                    )
                    driver.execute_script("arguments[0].click();", el)
                    print(f"Report clicked (default content) via: {xpath}")
                    return True
                except:
                    continue

            # Strategy C — debug dump + screenshot
            print("Could not click report. Dumping visible elements...")
            try:
                candidates = driver.find_elements(
                    By.XPATH,
                    "//*[contains(@class,'dropdown') or contains(@class,'search-result') "
                    "or contains(@class,'list') or contains(@class,'result')]"
                )
                for c in candidates:
                    txt = c.text.strip()
                    if txt:
                        print(
                            f"    [{c.tag_name}.{c.get_attribute('class')}]: "
                            f"{txt[:200]}"
                        )
            except:
                pass

            driver.save_screenshot("search_debug.png")
            print("Screenshot saved: search_debug.png")
            return False

        # -------------------------------------------------
        # LOGIN
        # -------------------------------------------------
        driver.get("https://erp.assetline.lk/")
        wait_for_loader(driver)

        wait.until(
            EC.presence_of_element_located((By.ID, "strUserName"))
        ).send_keys(USERNAME)

        wait.until(
            EC.presence_of_element_located((By.ID, "strPassword"))
        ).send_keys(PASSWORD)

        safe_click(wait.until(EC.element_to_be_clickable((By.ID, "idSignIn"))))
        wait_for_loader(driver)

        # The ERP no longer shows this extra "select" button after sign-in (erp_login.py
        # and the other bots have the click commented out). Click it only if it appears,
        # instead of waiting 30s and failing when it does not.
        try:
            safe_click(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                By.CSS_SELECTOR, "button.btn.blue.btn-circle"
            ))))
            wait_for_loader(driver)
        except Exception:
            print("No post-login select button - continuing")

        print("Login success")

        # -------------------------------------------------
        # OPEN REPORT WINDOW
        # -------------------------------------------------
        main_tab = driver.current_window_handle

        safe_click(wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[normalize-space()='Report']"
        ))))

        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(
            [h for h in driver.window_handles if h != main_tab][0]
        )
        wait_for_loader(driver)

        switch_into_iframe()

        # -------------------------------------------------
        # NAVIGATION — click Reports menu
        # -------------------------------------------------
        safe_click(wait.until(EC.element_to_be_clickable((
            By.XPATH, "//div[text()='Reports']"
        ))))
        wait_for_loader(driver)

        # -------------------------------------------------
        # SEARCH AND CLICK REPORT
        # -------------------------------------------------
        clicked = search_and_click_report()

        if not clicked:
            raise Exception(
                f"Could not find and click report: '{REPORT_NAME}'"
            )

        wait_for_loader(driver)

        # Re-enter iframe after report selection
        switch_into_iframe()

        # -------------------------------------------------
        # INQUIRY
        # -------------------------------------------------
        safe_click(wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[text()='Inquiry']"
        ))))
        wait_for_loader(driver)

        switch_into_iframe()

        # -------------------------------------------------
        # SELECT ALL
        # -------------------------------------------------
        safe_click(wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[normalize-space()='Select All']"
        ))))

        time.sleep(2)

        wait.until(EC.presence_of_element_located((
            By.XPATH, "//table/tbody/tr"
        )))

        # The 'Select All' click does not always take effect (the 05-Oct export came back
        # with only 2 columns, SLT_SLCCODE and SLT_TRNDATE, which write_off.py cannot use).
        # Fallback: click every column checkbox that is still unticked.
        BOX_XPATH = "//table/tbody/tr/td[1]//input[@type='checkbox']"

        def tick_count():
            bxs = driver.find_elements(By.XPATH, BOX_XPATH)
            return sum(1 for bx in bxs if bx.is_selected()), len(bxs)

        ticked, total = tick_count()
        print(f"Columns ticked after Select All: {ticked}/{total}")
        if ticked < total or total == 0:
            for bx in driver.find_elements(By.XPATH, BOX_XPATH):
                if bx.is_selected():
                    continue
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", bx)
                    bx.click()
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click();", bx)
                    except Exception:
                        pass
            ticked, total = tick_count()
            print(f"Columns ticked after clicking each box: {ticked}/{total}")

        # Last resort: click each field by name (the 17 columns write_off.py reads).
        if total == 0 or ticked < total:
            FIELDS = [
                "SLT_SLCCODE", "SLT_CONNO", "SLT_TRNNO", "SLT_TRNDATE", "SLT_TRNAMOUNT",
                "SLT_LASTUPDATE", "SLT_LASTUPTIME", "SLT_LASTUPUSERID", "SLST_TRNDRCR",
                "SLT_DESC", "SLT_CLMCODE", "CLM_NAME", "CLM_CLTITLE", "CLM_INITIALS",
                "CLM_INITIALSFULL", "DRAMT", "CRAMT",
            ]
            failed = []
            for field in FIELDS:
                try:
                    # exact column-name match first (CLM_INITIALS is also inside CLM_INITIALSFULL)
                    rows = driver.find_elements(By.XPATH, f"//tr[td[normalize-space()='{field}']]")
                    if not rows:
                        rows = driver.find_elements(By.XPATH, f"//tr[td[contains(text(),'{field}')]]")
                    cb = rows[0].find_element(By.XPATH, ".//td[1]//input[@type='checkbox']")
                    if not cb.is_selected():
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
                        try:
                            cb.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", cb)
                    if not cb.is_selected():
                        failed.append(field)
                except Exception:
                    failed.append(field)
            print(f"Named-field fallback done; could not tick: {failed or 'none'}")
            ticked, total = tick_count()
            print(f"Columns ticked after named fallback: {ticked}/{total}")
            required = {"SLT_CONNO", "SLT_TRNNO", "SLT_TRNDATE", "SLT_DESC", "CLM_NAME",
                        "CLM_INITIALS", "DRAMT", "CRAMT"}
            missing = sorted(required & set(failed))
            if missing:
                raise RuntimeError(f"Required report columns could not be selected: {missing}")

        # -------------------------------------------------
        # FILTER — SLT_SLCCODE
        # -------------------------------------------------
        select_checkbox("SLT_SLCCODE")

        operator = wait.until(EC.presence_of_element_located((
            By.XPATH, "//tr[1]/td[3]//select"
        )))
        Select(operator).select_by_value("=")

        criteria = wait.until(EC.presence_of_element_located((
            By.XPATH, "//input[contains(@class,'RPTCriteria')]"
        )))
        criteria.clear()
        criteria.send_keys("0130")

        print("Code filter applied")

        # -------------------------------------------------
        # FILTER — SLT_TRNDATE
        # '01/04/2026','28/04/2026' → written into FIRST criteria input as one string
        # -------------------------------------------------
        select_checkbox("SLT_TRNDATE")

        date_operator = wait.until(EC.presence_of_element_located((
            By.XPATH,
            "//tr[td[contains(text(),'SLT_TRNDATE')]]//select"
        )))
        Select(date_operator).select_by_value("BETWEEN")

        time.sleep(0.5)

        print(f"Writing combined date criteria: {COMBINED_DATES}")

        fill_criteria_input(
            "//tr[td[contains(text(),'SLT_TRNDATE')]]"
            "//input[contains(@class,'RPTCriteria')][1]",
            COMBINED_DATES      # → '01/04/2026','28/04/2026'
        )

        print(f"Date criteria applied: {COMBINED_DATES}")

        # -------------------------------------------------
        # GENERATE REPORT
        # -------------------------------------------------
        generate_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(@id,'GenerateReport')]"
        )))
        generate_clicked_at = time.time()
        driver.execute_script("arguments[0].click();", generate_btn)

        print("Generating report...")

        # -------------------------------------------------
        # WAIT FOR DOWNLOAD
        # -------------------------------------------------
        # Only accept a finished "Backend Due Types*.xlsx" that arrived AFTER Generate
        # was clicked. (Accepting any Excel file in Downloads returned instantly on an
        # old file, before this report had finished, and a short timeout gave up on the
        # larger cumulative export.)
        timeout    = time.time() + 600
        excel_file = None

        while time.time() < timeout:
            fresh = [
                f for f in glob.glob(os.path.join(downloads_dir, "Backend Due Types*.xlsx"))
                if not os.path.basename(f).startswith("~$")
                and os.path.getmtime(f) >= generate_clicked_at - 5
            ]
            still_downloading = glob.glob(os.path.join(downloads_dir, "*.crdownload"))

            if fresh and not still_downloading:
                excel_file = max(fresh, key=os.path.getmtime)
                break

            time.sleep(2)

        if not excel_file:
            raise RuntimeError("Backend Due Types report was generated but no new file "
                               "arrived in Downloads within 600s")

        print("REPORT DOWNLOADED:")
        print(excel_file)

        return driver

    except Exception as e:
        print(f"ERROR: {e}")

        try:
            if driver:
                driver.save_screenshot("error.png")
                print("Screenshot saved: error.png")
        except:
            pass

        try:
            if driver:
                driver.quit()
        except:
            pass

        raise


# -------------------------------------------------
# RUN
# -------------------------------------------------
if __name__ == "__main__":
    # python BackEndDueType_ContractWise.py [FROM_DATE [TO_DATE]]
    # FROM_DATE defaults to 1st of current month, TO_DATE defaults to today.
    _from = sys.argv[1] if len(sys.argv) > 1 else None
    _to   = sys.argv[2] if len(sys.argv) > 2 else None
    driver = erp_all_contracts_sector_report(_from, _to)
    print("DONE")
    time.sleep(5)
    driver.quit()