"""
Download "Reschedule Contract Details - Internal" from the Assetline ERP.

Flow: login -> (premise Select) -> Report -> Other Reports
      -> Reschedule Contract Details - Internal -> Inquiry -> Generate Report
      -> wait for the new file in Downloads -> copy it into the given
      destination folder (replacing any older copy there).

Used by report_automations/process_summary.py to refresh the reschedule list
that feeds the 'MOI List -31 Jan' sheet of the Quarter Classification report.
Messages go through the `logging` module, so when called from process_summary
they land in process_summary.log.

Run directly:
    python bots/reschedule_contract_details_bot.py --dest "<folder>"
"""
import os
import sys
import time
import glob
import shutil
import logging
import argparse
from pathlib import Path

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
sys.path.insert(0, PROJECT_ROOT)

try:
    from bots.erp_login import erp_login, wait_for_loader
except ImportError:
    from erp_login import erp_login, wait_for_loader

logger = logging.getLogger(__name__)

REPORT_NAME = "Reschedule Contract Details - Internal"
FILE_PATTERN = f"{REPORT_NAME}*.xlsx"
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
DOWNLOAD_TIMEOUT = 300  # seconds to wait for the file to land in Downloads

# ERP element locators
OTHER_REPORTS_TILE = ("//div[contains(@class,'metro-tile-body') and "
                      ".//div[normalize-space()='Other Reports']]")
RESCHEDULE_TILE = ("//div[contains(@class,'metro-tile-body') and "
                   f".//div[normalize-space()='{REPORT_NAME}']]")
INQUIRY_BTN_ID = "btn_01871_Inquiry"
GENERATE_BTN_ID = "btn_01871_GenerateReport"


def _downloaded_files():
    """Report files currently in Downloads, as {path: modified time}."""
    return {
        f: os.path.getmtime(f) for f in glob.glob(os.path.join(DOWNLOADS_DIR, FILE_PATTERN))
        if not os.path.basename(f).startswith("~$") and os.path.isfile(f)
    }


def _allow_downloads(driver):
    """Make Chrome save downloads straight into DOWNLOADS_DIR without a prompt.
    erp_login() starts Chrome with no download preferences, so set them here
    via DevTools (applies to every tab/window, incl. the Reports window)."""
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": DOWNLOADS_DIR,
            "eventsEnabled": True,
        })
        logger.info(f"Chrome downloads set to: {DOWNLOADS_DIR}")
    except Exception as e:
        logger.warning(f"Could not set Chrome download behaviour ({e}); using Chrome defaults")


def _close_success_dialog(driver, timeout=60):
    """Wait for the ERP's 'Report Successfully Generated' dialog and close it."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        for btn in driver.find_elements(By.XPATH, "//button[normalize-space()='CLOSE' or normalize-space()='Close']"):
            try:
                if btn.is_displayed():
                    logger.info("ERP: report generated - closing success dialog")
                    driver.execute_script("arguments[0].click();", btn)
                    return True
            except Exception:
                continue
        time.sleep(1)
    logger.info("No success dialog seen - continuing to wait for the download")
    return False


def _wait_for_download(before, timeout=DOWNLOAD_TIMEOUT):
    """Return the report file that appeared after Generate Report was clicked:
    a file not in `before` (or rewritten since), with a stable non-zero size.
    `before` is _downloaded_files() taken before the click.

    Chrome downloads into a .crdownload file and only renames it to the real
    name once complete, so the report's own name appearing means it finished.
    Other .crdownload files in Downloads are ignored - stale ones left by
    other downloads would otherwise block this forever."""
    deadline = time.time() + timeout
    last_heartbeat = time.time()
    while time.time() < deadline:
        current = _downloaded_files()
        new_files = [f for f, mtime in current.items()
                     if f not in before or mtime > before[f]]
        if new_files:
            newest = max(new_files, key=os.path.getmtime)
            size = os.path.getsize(newest)
            time.sleep(2)
            if size > 0 and os.path.getsize(newest) == size:
                return newest
        if time.time() - last_heartbeat >= 30:
            logger.info(f"Still waiting for the download ({int(deadline - time.time())}s left)")
            last_heartbeat = time.time()
        time.sleep(2)
    return None


def _select_premises(driver, timeout=15):
    """Click SELECT on the premise selection screen shown after sign-in.
    Skipped if the screen doesn't appear (the shared erp_login() doesn't do this)."""
    locators = [
        (By.CSS_SELECTOR, "button.btn.blue.btn-circle"),
        (By.XPATH, "//button[translate(normalize-space(), 'select', 'SELECT')='SELECT']"),
    ]
    deadline = time.time() + timeout
    while time.time() < deadline:
        for locator in locators:
            for btn in driver.find_elements(*locator):
                if btn.is_displayed() and btn.is_enabled():
                    logger.info("Clicking Select on premise selection...")
                    try:
                        btn.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", btn)
                    wait_for_loader(driver)
                    time.sleep(2)
                    return True
        time.sleep(1)
    logger.info("Premise selection screen not shown - continuing")
    return False


def _click(driver, wait, xpath=None, element_id=None, description=""):
    locator = (By.ID, element_id) if element_id else (By.XPATH, xpath)
    logger.info(f"Clicking {description}...")
    el = wait.until(EC.element_to_be_clickable(locator))
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    wait_for_loader(driver)


def _copy_into(downloaded, dest_dir):
    """Copy the downloaded file into dest_dir, replacing older copies there,
    then remove it from Downloads (best effort - e.g. it may be open in Excel).
    If dest_dir is Downloads itself, the downloaded file is kept where it is."""
    downloaded = Path(downloaded).resolve()
    dest = (dest_dir / downloaded.name).resolve()

    for old in dest_dir.glob(FILE_PATTERN):
        if old.name.startswith("~$") or old.resolve() == downloaded:
            continue
        try:
            old.unlink()
            logger.info(f"Removed old: {old}")
        except OSError as e:
            logger.warning(f"Could not remove old file {old} ({e}); the newest file will still be used")

    if dest == downloaded:
        logger.info(f"Destination is the Downloads folder - keeping: {dest}")
        return dest

    shutil.copy2(downloaded, dest)
    logger.info(f"Copied to: {dest}")
    try:
        os.remove(downloaded)
    except OSError as e:
        logger.warning(f"Copied, but could not remove {downloaded} from Downloads ({e})")
    return dest


def download_reschedule_contract_details(dest_dir):
    """Download the report from the ERP and copy it into `dest_dir`.

    Older "Reschedule Contract Details - Internal*.xlsx" files in `dest_dir`
    are removed so only the fresh one is picked up. Returns the new file's
    Path. Raises on any failure (caller decides whether to continue).
    """
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"Downloading '{REPORT_NAME}' from the ERP into {dest_dir}")

    driver, wait = erp_login()
    try:
        _allow_downloads(driver)
        _select_premises(driver)

        main_tab = driver.current_window_handle
        _click(driver, wait,
               xpath="//li[contains(@class,'nav-item')]//span[normalize-space()='Report']",
               description="Report")

        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window([h for h in driver.window_handles if h != main_tab][0])
        logger.info("Switched to Reports window")
        time.sleep(2)

        _click(driver, wait, xpath=OTHER_REPORTS_TILE, description="Other Reports")
        _click(driver, wait, xpath=RESCHEDULE_TILE, description=REPORT_NAME)
        _click(driver, wait, element_id=INQUIRY_BTN_ID, description="Inquiry")

        generate_btn = wait.until(EC.element_to_be_clickable((By.ID, GENERATE_BTN_ID)))
        before = _downloaded_files()
        generate_btn.click()
        logger.info("Generate Report clicked - waiting for download...")

        _close_success_dialog(driver)
        downloaded = _wait_for_download(before)
        if not downloaded:
            raise RuntimeError(f"No new '{FILE_PATTERN}' appeared in {DOWNLOADS_DIR} "
                               f"within {DOWNLOAD_TIMEOUT}s after Generate Report")
        logger.info(f"Downloaded: {downloaded}")

        return _copy_into(downloaded, dest_dir)

    except Exception:
        try:
            driver.save_screenshot(os.path.join(CURRENT_DIR, "reschedule_report_error.png"))
        except Exception:
            pass
        raise
    finally:
        driver.quit()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s',
                        datefmt='%H:%M:%S', handlers=[logging.StreamHandler(sys.stdout)])
    parser = argparse.ArgumentParser(description=f"Download {REPORT_NAME} from the ERP")
    parser.add_argument("--dest", required=True, help="Folder to copy the downloaded report into")
    args = parser.parse_args()
    path = download_reschedule_contract_details(args.dest)
    logger.info(f"Done: {path}")
