import re
import sys
import os
import time
import logging
from datetime import datetime
from pathlib import Path

# Force UTF-8 output on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def _setup_logger() -> logging.Logger:
    log_dir = os.environ.get("IFRS_LGD_LOG_DIR", "").strip()
    if not log_dir:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "logs", f"run_{ts}"
        )
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "Closure_Category_Wise_Report.log")

    logger = logging.getLogger("ClosureBot")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    logger.info(f"Logger initialized. Log file: {log_file}")
    return logger


logger = _setup_logger()
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# -------------------------------------------------
# PATH SETUP
# -------------------------------------------------
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
sys.path.insert(0, PROJECT_ROOT)

try:
    from bots.erp_login import wait_for_loader, USERNAME, PASSWORD
except ImportError:
    from erp_login import wait_for_loader, USERNAME, PASSWORD


def _find_lgd_firstdefault_folder() -> Path:
    """
    Resolve the firstDefault folder to download into.

    Prefers LGD_BASE_FOLDER (set by ifrs.py to the already-prepared working
    copy for this run) so the download lands where FDDW.py/write_off.py/
    LGD_computation.py actually read from. Without this, the bot would keep
    downloading into outputs/LGD/<date>/firstDefault/ — which the working
    copy is snapshotted from *before* this bot runs, so a fresh download
    there would never reach this run's working folder, only the next one's.

    Falls back to outputs/LGD/<latest-date-folder>/firstDefault/ (from
    completed_dates.txt) for standalone/manual runs outside the portal.
    """
    lgd_base = os.environ.get("LGD_BASE_FOLDER", "").strip()
    if lgd_base:
        dest = Path(lgd_base) / "firstDefault"
        dest.mkdir(parents=True, exist_ok=True)
        return dest

    outputs_lgd = Path(PROJECT_ROOT) / "outputs" / "LGD"
    dates_file  = outputs_lgd / "completed_dates.txt"
    lines = [l.strip() for l in dates_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    last_date = lines[-1]

    date_variants = {last_date}
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
        raise RuntimeError(f"No date folder found for '{last_date}' in {outputs_lgd}")
    candidates.sort(key=lambda x: x[0], reverse=True)
    dest = candidates[0][1] / "firstDefault"
    dest.mkdir(parents=True, exist_ok=True)
    return dest


def erp_login_closure_report():
    """Login to ERP and navigate to Closure Category Wise Report - Internal."""

    logger.info("Credentials loaded")

    # -------------------------------------------------
    # Resolve destination up front so Chrome can download straight into it
    # -------------------------------------------------
    dest_folder = _find_lgd_firstdefault_folder()
    logger.info(f"Destination folder: {dest_folder}")
    existing_before = {
        f.name for f in dest_folder.glob("Closure Category Wise Report - Internal*.xlsx")
    }

    # -------------------------------------------------
    # Chrome setup
    # -------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": str(dest_folder),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })

    logger.info("Starting Chrome...")
    driver = webdriver.Chrome(
        service=Service(),
        options=chrome_options
    )

    wait = WebDriverWait(driver, 30)

    try:
        # -----------------------------
        # LOGIN
        # -----------------------------
        logger.info("Opening ERP login page...")
        driver.get("https://erp.assetline.lk/")
        wait_for_loader(driver)

        logger.info("Entering credentials...")
        wait.until(EC.presence_of_element_located((By.ID, "strUserName"))).send_keys(USERNAME)
        wait.until(EC.presence_of_element_located((By.ID, "strPassword"))).send_keys(PASSWORD)

        logger.info("Clicking sign in...")
        wait.until(EC.element_to_be_clickable((By.ID, "idSignIn"))).click()
        wait_for_loader(driver)

        logger.info("Checking for premise select screen...")
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
                (By.XPATH, "//*[@id='idSignIn']"))
            ).click()
            wait_for_loader(driver)
            logger.info("Premise select screen handled.")
        except TimeoutException:
            logger.info("No premise select screen shown — continuing.")

        logger.info("Login successful")

        # -----------------------------
        # NAVIGATE TO REPORTS
        # -----------------------------
        logger.info("Navigating to Reports...")
        main_tab = driver.current_window_handle

        report_btn = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//li[contains(@class,'nav-item')]//span[normalize-space()='Report']"
            ))
        )
        report_btn.click()

        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(
            [h for h in driver.window_handles if h != main_tab][0]
        )

        logger.info("Switched to Reports window")
        time.sleep(2)

        reports_tile = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[contains(@class,'metro-tile') and .//div[normalize-space()='Reports']]"
            ))
        )
        reports_tile.click()
        wait_for_loader(driver)

        closure_tile = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[contains(@class,'metro-tile') and "
                ".//div[normalize-space()='Closure Category Wise Report - Internal']]"
            ))
        )
        closure_tile.click()
        wait_for_loader(driver)

        inquiry_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "btn_01956_Inquiry"))
        )
        inquiry_btn.click()
        wait_for_loader(driver)

        logger.info("Selecting report category checkboxes...")
        TOTAL_CATEGORY_ROWS = 29
        for row in range(1, TOTAL_CATEGORY_ROWS + 1):
            checkbox = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    f"//*[@id='grid_01956_ReportDetailsRPT']/tbody/tr[{row}]/td[1]/input"
                ))
            )
            if not checkbox.is_selected():
                checkbox.click()
        logger.info(f"Selected {TOTAL_CATEGORY_ROWS} report category checkboxes")

        generate_btn = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "btn_01956_GenerateReport")
            )
        )
        generate_btn.click()
        logger.info("Report generation started")

        # -----------------------------
        # WAIT FOR THE NEW FILE TO FINISH DOWNLOADING (straight into dest_folder)
        # -----------------------------
        DOWNLOAD_TIMEOUT = 600   # seconds (large reports with all categories can take a while to generate)
        POLL_INTERVAL    = 1.0   # seconds
        logger.info(f"Waiting for new report file in: {dest_folder} (timeout {DOWNLOAD_TIMEOUT}s)")

        downloaded_path = None
        waited = 0.0
        while waited < DOWNLOAD_TIMEOUT:
            candidates = [
                f for f in dest_folder.glob("Closure Category Wise Report - Internal*.xlsx")
                if f.name not in existing_before and not f.name.startswith("~$")
            ]
            in_progress = list(dest_folder.glob("Closure Category Wise Report - Internal*.crdownload"))
            if candidates and not in_progress:
                downloaded_path = max(candidates, key=lambda f: f.stat().st_mtime)
                break
            time.sleep(POLL_INTERVAL)
            waited += POLL_INTERVAL

        if downloaded_path:
            logger.info(f"New report downloaded: {downloaded_path.name} (waited {waited:.0f}s)")
            for old_name in existing_before:
                old_path = dest_folder / old_name
                try:
                    old_path.unlink()
                    logger.info(f"Removed old: {old_name}")
                except OSError as exc:
                    logger.warning(f"Could not remove old file {old_name}: {exc}")
        else:
            logger.warning(
                f"Closure Category Wise Report not found in {dest_folder} "
                f"after waiting {DOWNLOAD_TIMEOUT}s — existing file preserved"
            )

        return driver

    except Exception as e:
        driver.save_screenshot("closure_report_error.png")
        logger.error(f"Automation failed: {e}")
        driver.quit()
        raise Exception(f"Automation failed: {e}")


# -------------------------------------------------
# MAIN
# -------------------------------------------------
if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("Closure Category Wise Report")
    logger.info("=" * 60)

    driver = erp_login_closure_report()

    logger.info("Process completed successfully")
    time.sleep(10)

    driver.quit()
    logger.info("Browser closed")
