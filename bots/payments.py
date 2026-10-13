
import os
import re
import sys
import time
from pathlib import Path

from dotenv import load_dotenv
from cryptography.fernet import Fernet
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

CE_ROOT = Path(__file__).resolve().parent
ENV_PATH = CE_ROOT.parent / ".env"

SUB_LEDGER_CODE = "1173"
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
DOWNLOAD_DIR = Path.home() / "Downloads"

load_dotenv(ENV_PATH)


class ERPScraper:
    def __init__(self, headless: bool = False, download_dir: Path = DOWNLOAD_DIR):
        self.download_dir = download_dir
        self.chrome_opts = Options()
        if headless:
            self.chrome_opts.add_argument("--headless=new")
        self.chrome_opts.add_argument("--disable-blink-features=AutomationControlled")
        self.chrome_opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.chrome_opts.add_experimental_option("useAutomationExtension", False)
        self.chrome_opts.add_experimental_option("prefs", {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })
        self.driver = None

    def decrypt_credentials(self) -> tuple[str, str]:
        key_str = os.getenv("FERNET_KEY", "").strip()
        enc_username = os.getenv("ENC_USERNAME", "").strip()
        enc_password = os.getenv("ENC_PASSWORD", "").strip()

        if not key_str:
            raise ValueError(f"FERNET_KEY not found in {ENV_PATH}")
        if not enc_username or not enc_password:
            raise ValueError(f"ENC_USERNAME / ENC_PASSWORD not found in {ENV_PATH}")

        print("  Decrypting credentials...")
        fernet = Fernet(key_str.encode())
        username = fernet.decrypt(enc_username.encode()).decode()
        password = fernet.decrypt(enc_password.encode()).decode()
        print("  Credentials decrypted.")
        return username, password

    def login_to_erp(self, username: str, password: str) -> bool:
        try:
            self.driver = webdriver.Chrome(options=self.chrome_opts)
            self.driver.execute_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )

            print("  Opening ERP website...")
            self.driver.get("https://erp.assetline.lk")

            wait = WebDriverWait(self.driver, 15)

            print("  Entering username...")
            field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="strUserName"]')))
            field.clear()
            field.send_keys(username)

            print("  Entering password...")
            field = self.driver.find_element(By.XPATH, '//*[@id="strPassword"]')
            field.clear()
            field.send_keys(password)

            print("  Clicking Sign In...")
            self.driver.find_element(By.XPATH, '//*[@id="idSignIn"]').click()
            time.sleep(3)

            print("  Logged in!")
            return True

        except Exception as exc:
            print(f"  Login failed: {exc}")
            return False

    def navigate_to_gl_inquiry(self):
        wait = WebDriverWait(self.driver, 15)

        print("  Navigating to FINANCE module...")
        self.driver.get("https://erp.assetline.lk/Application/Home/FINANCE")

        print("  Opening tile...")
        tile = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="iddivAppContent"]/div[1]/div/div[2]/div[1]/div/div[2]/div[1]/div/div/div')
        ))
        tile.click()

        print("  Opening sub-tile...")
        sub_tile = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="iddivAppContent"]/div[1]/div/div[2]/div[2]/div/div/div/div/div')
        ))
        sub_tile.click()

        print("  Clicking Inquiry...")
        inquiry_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_02366_Inquiry"]')))
        inquiry_btn.click()

    def search_gl_account(self, from_date: str, to_date: str):
        wait = WebDriverWait(self.driver, 15)

        print(f"  Entering sub ledger code {SUB_LEDGER_CODE}...")
        sub_ledger_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="strSubLedgerCode_FINGLAccountView"]')
        ))
        sub_ledger_field.click()
        sub_ledger_field.send_keys(SUB_LEDGER_CODE)

        print(f"  Entering from-date {from_date}...")
        from_date_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="dtmTransactionDate_FINGLAccountView"]')
        ))
        from_date_field.click()
        from_date_field.send_keys(Keys.CONTROL + "a")
        from_date_field.send_keys(Keys.DELETE)
        from_date_field.send_keys(from_date)
        from_date_field.send_keys(Keys.RETURN)

        print(f"  Entering to-date {to_date}...")
        to_date_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="dtmTransactionToDate_FINGLAccountView"]')
        ))
        to_date_field.click()
        to_date_field.send_keys(Keys.CONTROL + "a")
        to_date_field.send_keys(Keys.DELETE)
        to_date_field.send_keys(to_date)
        to_date_field.send_keys(Keys.RETURN)

        print("  Clicking Search...")
        search_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="btn_02366_BtnSearch_FINGLAccountView"]')
        ))
        search_btn.click()

    def download_excel(self, timeout: int = 120):
        print("  Waiting for Excel Download button to appear...")
        wait = WebDriverWait(self.driver, 60)
        download_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_02366_ExcelDownload"]')))

        existing_files = set(self.download_dir.glob("*"))

        print("  Clicking Excel Download...")
        download_btn.click()

        print("  Waiting for download to finish...")
        end_time = time.time() + timeout
        while time.time() < end_time:
            current_files = set(self.download_dir.glob("*"))
            new_files = current_files - existing_files
            in_progress = [f for f in new_files if f.suffix in (".crdownload", ".tmp")]
            finished = [f for f in new_files if f.suffix not in (".crdownload", ".tmp")]

            if finished and not in_progress:
                print(f"  Download finished: {finished[0].name}")
                return finished[0]

            time.sleep(1)

        print("  Download did not complete within timeout.")
        return None

    def close(self):
        if self.driver:
            self.driver.quit()


def main():
    if len(sys.argv) != 3 or not DATE_RE.match(sys.argv[1]) or not DATE_RE.match(sys.argv[2]):
        print("  Usage: python erp_scraper.py dd/mm/yyyy dd/mm/yyyy")
        return

    from_date, to_date = sys.argv[1], sys.argv[2]

    scraper = ERPScraper(headless=False)

    try:
        username, password = scraper.decrypt_credentials()
    except ValueError as exc:
        print(f"  {exc}")
        return

    if not scraper.login_to_erp(username, password):
        return

    scraper.navigate_to_gl_inquiry()
    scraper.search_gl_account(from_date, to_date)
    scraper.download_excel()

    scraper.close()


if __name__ == "__main__":
    main()
