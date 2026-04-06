from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from cryptography.fernet import Fernet
from datetime import datetime, timedelta
import time
import os
import sys

# Configuration
LOGIN_URL = "https://efn.assetline.lk:1000/#/"
REPORT_URL = "https://efn.assetline.lk:1009/Report/MovementOfDailyDepositBalance.aspx"
INTEREST_PAYABLE_URL = "https://efn.assetline.lk:1009/Report/InterestPayableAsAtDate.aspx"
NEW_DEPOSIT_BASE_URL = "https://efn.assetline.lk:1009/NewDepositBase.aspx"
ERP_URL = "https://erp.assetline.lk"


class EFNBot:
    """
    EFN Report Automation Bot
    """

    def __init__(self):
        self.driver = None
        self.wait = None

    def calculate_date_range(self, monday_date_str):
        """
        Calculate the date range from Sunday to Sunday based on the Monday date.

        Args:
            monday_date_str: Date string in format YYYY/MM/DD (should be a Monday)

        Returns:
            tuple: (from_date, to_date) in YYYY/MM/DD format

        Example:
            If monday_date_str = "2025/10/13" (Monday)
            Returns: ("2025/10/05", "2025/10/12")
            - from_date: Previous Sunday (8 days before Monday)
            - to_date: Sunday just before Monday (1 day before Monday)
        """
        try:
            # Parse the Monday date
            monday_date = datetime.strptime(monday_date_str, "%Y/%m/%d")

            # Calculate the previous Sunday (1 day before Monday)
            to_date = monday_date - timedelta(days=1)

            # Calculate the Sunday before that (8 days before Monday = 7 days before to_date)
            from_date = to_date - timedelta(days=7)

            # Format dates back to YYYY/MM/DD
            from_date_str = from_date.strftime("%Y/%m/%d")
            to_date_str = to_date.strftime("%Y/%m/%d")

            print(f"Calculated date range:")
            print(f"  Monday argument: {monday_date_str} ({monday_date.strftime('%A')})")
            print(f"  From date: {from_date_str} ({from_date.strftime('%A')})")
            print(f"  To date: {to_date_str} ({to_date.strftime('%A')})")

            return from_date_str, to_date_str

        except ValueError as e:
            print(f"Error parsing date. Please use format YYYY/MM/DD: {e}")
            raise

    def get_date_list(self, from_date_str, to_date_str):
        """
        Generate a list of all dates between from_date and to_date (inclusive)

        Args:
            from_date_str: Start date in format YYYY/MM/DD
            to_date_str: End date in format YYYY/MM/DD

        Returns:
            list: List of date strings in DD/MM/YYYY format
        """
        from_date = datetime.strptime(from_date_str, "%Y/%m/%d")
        to_date = datetime.strptime(to_date_str, "%Y/%m/%d")

        date_list = []
        current_date = from_date

        while current_date <= to_date:
            # Format as DD/MM/YYYY for the report
            date_list.append(current_date.strftime("%d/%m/%Y"))
            current_date += timedelta(days=1)

        return date_list

    def get_last_date(self, from_date_str, to_date_str):
        """
        Get the last date (day 8) from the date range

        Args:
            from_date_str: Start date in format YYYY/MM/DD
            to_date_str: End date in format YYYY/MM/DD

        Returns:
            str: Last date in DD/MM/YYYY format
        """
        to_date = datetime.strptime(to_date_str, "%Y/%m/%d")
        return to_date.strftime("%d/%m/%Y")

    def decrypt_credentials(self, enc_username, enc_password):
        """
        Decrypt encrypted username and password
        """
        try:
            if not enc_username or not enc_password:
                raise ValueError("ENC_USERNAME and ENC_PASSWORD must be set in environment variables")

            print("Found encrypted credentials, attempting to decrypt...")

            key_str = os.environ.get("FERNET_KEY", "").strip()
            if not key_str:
                raise RuntimeError("FERNET_KEY environment variable not set")

            fernet = Fernet(key_str.encode())

            # Decrypt credentials
            username = fernet.decrypt(enc_username.encode()).decode()
            password = fernet.decrypt(enc_password.encode()).decode()

            print("Successfully decrypted credentials")
            return username, password

        except Exception as e:
            print(f"Failed to decrypt credentials: {e}")
            raise

    def initialize_browser(self):
        """
        Initialize the WebDriver
        """
        print("Initializing browser...")
        self.driver = webdriver.Chrome()  # or webdriver.Firefox(), webdriver.Edge(), etc.
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 20)
        print("Browser initialized successfully")

    def login(self, username, password):
        """
        Login to the EFN portal
        """
        try:
            # Step 1: Navigate to login page
            print("Navigating to login page...")
            self.driver.get(LOGIN_URL)

            # Step 2: Enter username
            print("Entering username...")
            username_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-loginnew/app-login/div/div[1]/div/form/div[3]/div/input"))
            )
            username_field.clear()
            username_field.send_keys(username)

            # Step 3: Enter password
            print("Entering password...")
            password_field = self.driver.find_element(By.XPATH, "/html/body/app-root/app-loginnew/app-login/div/div[1]/div/form/div[4]/div/input")
            password_field.clear()
            password_field.send_keys(password)

            # Step 4: Click login button
            print("Clicking login button...")
            login_button = self.driver.find_element(By.XPATH, "/html/body/app-root/app-loginnew/app-login/div/div[1]/div/form/button")
            login_button.click()

            print("Login successful")
            return True

        except Exception as e:
            print(f"Login failed: {e}")
            raise

    def navigate_to_report(self):
        """
        Navigate to the report page from dashboard
        """
        try:
            # Step 5: Wait and click the tile
            print("Waiting for dashboard and clicking tile...")
            tile = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main-layout/app-portal/app-megamenu/div/div/div/div/ul/mat-grid-list/div/div[1]/mat-grid-tile"))
            )
            tile.click()

            # Step 6: Navigate to the report URL
            print("Navigating to report page...")
            time.sleep(2)  # Give time for any transitions
            self.driver.get(REPORT_URL)

            print("Successfully navigated to report page")
            return True

        except Exception as e:
            print(f"Failed to navigate to report: {e}")
            raise

    def generate_report(self, from_date, to_date):
        """
        Generate the report with specified date range
        """
        try:
            # Step 7: Enter FROM date
            print(f"Entering FROM date: {from_date}")
            from_date_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtFromDate"]'))
            )
            from_date_field.clear()
            from_date_field.send_keys(from_date)

            # Step 8: Enter TO date
            print(f"Entering TO date: {to_date}")
            to_date_field = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtToDate"]')
            to_date_field.clear()
            to_date_field.send_keys(to_date)

            # Step 9: Click Display button
            print("Clicking Display button...")
            display_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnDisplay"]')
            display_button.click()

            # Step 10: Wait for report viewer and select dropdown option
            print("Selecting report format (option 8)...")
            time.sleep(3)  # Wait for report to load
            dropdown1 = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"]'))
            )
            select1 = Select(dropdown1)
            select1.select_by_index(7)  # Index 7 corresponds to option 8 (0-indexed)

            # Step 11: Click the first action link
            print("Clicking action link...")
            time.sleep(1)
            action_link1 = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"]')
            action_link1.click()

            print("Initial report generated successfully")
            return True

        except Exception as e:
            print(f"Failed to generate report: {e}")
            raise

    def generate_detail_report(self):
        """
        Generate the detailed report
        """
        try:
            # Step 12: Click Display Detail button
            print("Clicking Display Detail button...")
            time.sleep(2)
            display_detail_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnDisplayDetail"]')
            display_detail_button.click()

            # Step 13: Select dropdown option 8 again
            print("Selecting report format again (option 8)...")
            time.sleep(3)  # Wait for report to load
            dropdown2 = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"]'))
            )
            select2 = Select(dropdown2)
            select2.select_by_index(7)  # Index 7 corresponds to option 8 (0-indexed)

            # Step 14: Click the final action link
            print("Clicking final action link...")
            time.sleep(1)
            action_link2 = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"]')
            action_link2.click()

            print("Detailed report generated successfully")
            return True

        except Exception as e:
            print(f"Failed to generate detailed report: {e}")
            raise

    def download_daily_interest_reports(self, date_list):
        """
        Download Interest Payable reports for each day in the date range

        Args:
            date_list: List of dates in DD/MM/YYYY format
        """
        try:
            print("\n" + "=" * 60)
            print("Starting Daily Interest Payable Reports Download")
            print(f"Total reports to download: {len(date_list)}")
            print("=" * 60 + "\n")

            for index, date_str in enumerate(date_list, 1):
                print(f"\n--- Processing Report {index}/{len(date_list)} ---")
                print(f"Date: {date_str}")

                # Navigate to Interest Payable report page
                print("Navigating to Interest Payable report page...")
                self.driver.get(INTEREST_PAYABLE_URL)
                time.sleep(2)  # Wait for page to load

                # Enter the date
                print(f"Entering date: {date_str}")
                date_field = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtAsAtDate"]'))
                )
                date_field.clear()
                date_field.send_keys(date_str)

                # Click Display button
                print("Clicking Display button...")
                display_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnDisplay"]')
                display_button.click()

                # Wait for report to load and select dropdown option 8
                print("Selecting report format (option 8 - Excel)...")
                time.sleep(3)  # Wait for report viewer to load
                dropdown = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"]'))
                )
                select_dropdown = Select(dropdown)
                select_dropdown.select_by_index(2)  # Index 2 corresponds to option 3 (0-indexed)

                # Click the export link to download
                print("Clicking export link to download report...")
                time.sleep(1)
                export_link = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"]')
                export_link.click()

                # Wait for download to initiate
                time.sleep(3)  # Give time for download to start

                print(f"✓ Report {index}/{len(date_list)} downloaded successfully for {date_str}")

            print("\n" + "=" * 60)
            print(f"All {len(date_list)} daily reports downloaded successfully!")
            print("=" * 60 + "\n")
            return True

        except Exception as e:
            print(f"\nFailed to download daily reports: {e}")
            import traceback
            traceback.print_exc()
            raise

    def download_new_deposit_base_reports(self, date_list):
        """
        Download New Deposit Base reports for each day in the date range (Days 1-8)

        Args:
            date_list: List of dates in DD/MM/YYYY format
        """
        try:
            print("\n" + "=" * 60)
            print("Starting New Deposit Base Reports Download")
            print(f"Total reports to download: {len(date_list)}")
            print("=" * 60 + "\n")

            for index, date_str in enumerate(date_list, 1):
                print(f"\n--- Processing Report {index}/{len(date_list)} ---")
                print(f"Date: {date_str}")

                # Navigate to New Deposit Base report page
                print("Navigating to New Deposit Base report page...")
                self.driver.get(NEW_DEPOSIT_BASE_URL)
                time.sleep(2)  # Wait for page to load

                # Enter the date
                print(f"Entering date: {date_str}")
                date_field = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtDate"]'))
                )
                date_field.clear()
                date_field.send_keys(date_str)

                # Click Generate button
                print("Clicking Generate button...")
                gen_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnGen"]')
                gen_button.click()

                # Wait 3 seconds for report to load
                print("Waiting for report to generate...")
                time.sleep(3)

                # Select dropdown option 3
                print("Selecting report format (option 3 - Excel)...")
                dropdown = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl00"]'))
                )
                select_dropdown = Select(dropdown)
                select_dropdown.select_by_index(2)  # Index 2 corresponds to option 3 (0-indexed)

                # Click the export link to download
                print("Clicking export link to download report...")
                time.sleep(1)
                export_link = self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ReportViewer1_ctl01_ctl05_ctl01"]')
                export_link.click()

                # Wait for download to initiate
                time.sleep(3)  # Give time for download to start

                print(f"✓ Report {index}/{len(date_list)} downloaded successfully for {date_str}")

            print("\n" + "=" * 60)
            print(f"All {len(date_list)} New Deposit Base reports downloaded successfully!")
            print("=" * 60 + "\n")
            return True

        except Exception as e:
            print(f"\nFailed to download New Deposit Base reports: {e}")
            import traceback
            traceback.print_exc()
            raise

    def select_premises(self):
        """
        Just click the SELECT button to proceed
        """
        try:
            wait = WebDriverWait(self.driver, 20)
            
            # Wait a bit for the page/modal to load
            time.sleep(3)
            
            # Just click the Select button directly
            print("Clicking Select button...")
            select_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/form/div/button')))
            select_button.click()
            
            # Wait for page to load after selection
            time.sleep(3)
            
            # Refresh the page
            print("Refreshing the page...")
            self.driver.refresh()
            
            # Wait for page to fully reload
            print("Waiting for page to fully reload...")
            time.sleep(3)
            
            print("Premises selection completed!")
            return True
            
        except Exception as e:
            print(f"Error clicking Select button: {str(e)}")
            return False

    def download_erp_report(self, username, password, last_date):
        """
        Login to ERP system and download report for the last date (day 8)
        WITH PREMISES SELECTION STEP

        Args:
            username: Decrypted username
            password: Decrypted password
            last_date: Last date in DD/MM/YYYY format (day 8)
        """
        try:
            print("\n" + "=" * 60)
            print("Starting ERP Report Download")
            print(f"Date: {last_date} (Day 8)")
            print("=" * 60 + "\n")

            # Step 1: Navigate to ERP login page
            print("Step 1: Navigating to ERP login page...")
            self.driver.get(ERP_URL)
            time.sleep(2)

            # Step 2: Enter username
            print("Step 2: Entering username...")
            username_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="strUserName"]'))
            )
            username_field.clear()
            username_field.send_keys(username)

            # Step 3: Enter password
            print("Step 3: Entering password...")
            password_field = self.driver.find_element(By.XPATH, '//*[@id="strPassword"]')
            password_field.clear()
            password_field.send_keys(password)

            # Step 4: Click Sign In button
            print("Step 4: Clicking Sign In button...")
            signin_button = self.driver.find_element(By.XPATH, '//*[@id="idSignIn"]')
            signin_button.click()

            # Step 5: Wait 3 seconds
            print("Step 5: Waiting 3 seconds...")
            time.sleep(3)

            # Step 6: Check if redirected to ExternalLoginView
            current_url = self.driver.current_url
            print(f"Step 6: Current URL: {current_url}")

            if "ExternalLogin/ExternalLoginView" in current_url:
                print("Step 6: Detected ExternalLoginView page")
                print("Step 7: Clicking Sign In button on ExternalLoginView page...")
                signin_button_external = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="idSignIn"]'))
                )
                signin_button_external.click()
                
                # Step 8: Wait 3 seconds after clicking
                print("Step 8: Waiting 3 seconds after Sign In...")
                time.sleep(3)
            else:
                print("Step 6: Not on ExternalLoginView, skipping secondary Sign In")
                time.sleep(3)

            # NEW STEP: Select premises (branch selection)
            print("Step 7: Selecting premises (Head Branch)...")
            if not self.select_premises():
                print("Warning: Premises selection may have failed, but continuing...")

            # Step 9: Navigate to FINANCE page
            print("Step 9: Navigating to FINANCE page...")
            finance_url = "https://erp.assetline.lk/Application/Home/FINANCE"
            self.driver.get(finance_url)
            time.sleep(3)

            # Step 10: Click the tile
            print("Step 10: Clicking the finance tile...")
            tile = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="iddivAppContent"]/div[1]/div/div[1]/div[2]/div/div[3]'))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", tile)
            time.sleep(1)
            tile.click()

            # Step 11: Click Inquiry button
            print("Step 11: Clicking Inquiry button...")
            time.sleep(2)
            inquiry_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_03923_Inquiry"]'))
            )
            inquiry_button.click()

            # Step 12: Enter date - MANUALLY TYPE CHARACTER BY CHARACTER
            print(f"Step 12: Manually typing date: {last_date}")
            time.sleep(2)
            
            # Click the date field to activate it
            date_field = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="dtmDate_FINTB_DetailReport"]'))
            )
            date_field.click()
            time.sleep(1)
            
            # Select all existing text and delete it
            date_field.send_keys(Keys.CONTROL + "a")
            time.sleep(0.5)
            date_field.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)
            
            # Type the date character by character (DD/MM/YYYY format)
            print(f"Typing date character by character: {last_date}")
            for char in last_date:
                date_field.send_keys(char)
                time.sleep(0.15)  # Small delay between each character
            
            # Press Enter to confirm the date
            time.sleep(0.5)
            date_field.send_keys(Keys.ENTER)
            time.sleep(1.5)
            
            print(f"Date entered and confirmed: {last_date}")

            # Step 13: Click Download Excel Report button
            print("Step 13: Clicking Download Excel Report button...")
            time.sleep(1)
            download_button = self.driver.find_element(By.XPATH, '//*[@id="btn_03923_DownloadExcelReport_FINTB_DetailReport"]')
            download_button.click()

            # Step 14: Wait for download to start and complete
            print("Step 14: Waiting for report to generate and download...")
            print("This may take 5-10 minutes. Please be patient...")
            
            # Get the downloads folder path
            download_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            
            # Get list of files before download
            print("Monitoring download folder for new files...")
            files_before = set(os.listdir(download_folder))
            
            # Wait for download to start (check every 10 seconds for up to 10 minutes)
            download_started = False
            max_wait_start = 600  # 10 minutes in seconds
            check_interval = 10  # Check every 10 seconds
            elapsed_time = 0
            
            print("Waiting for download to start...")
            while elapsed_time < max_wait_start and not download_started:
                time.sleep(check_interval)
                elapsed_time += check_interval
                
                files_after = set(os.listdir(download_folder))
                new_files = files_after - files_before
                
                # Check if any new files appeared (including .crdownload for Chrome or .tmp for partial downloads)
                if new_files:
                    print(f"  [{elapsed_time}s] Download detected! New files: {new_files}")
                    download_started = True
                    break
                else:
                    print(f"  [{elapsed_time}s] Still waiting for download to start...")
            
            if not download_started:
                print("Warning: No new files detected after 10 minutes. Download may have failed.")
                print("Please check manually.")
            else:
                # Wait for download to complete (no .crdownload or .tmp files)
                print("Waiting for download to complete...")
                max_wait_complete = 600  # Additional 10 minutes for download to complete
                elapsed_time = 0
                
                while elapsed_time < max_wait_complete:
                    time.sleep(check_interval)
                    elapsed_time += check_interval
                    
                    files_current = os.listdir(download_folder)
                    
                    # Check if there are any incomplete download files
                    incomplete_files = [f for f in files_current if f.endswith('.crdownload') or f.endswith('.tmp') or f.endswith('.part')]
                    
                    if not incomplete_files:
                        print(f"  [{elapsed_time}s] Download completed successfully!")
                        break
                    else:
                        print(f"  [{elapsed_time}s] Download in progress... ({len(incomplete_files)} partial file(s))")
                
                if incomplete_files:
                    print("Warning: Download may still be in progress after waiting.")
                else:
                    # Show the downloaded file
                    files_final = set(os.listdir(download_folder))
                    downloaded_files = files_final - files_before
                    actual_files = [f for f in downloaded_files if not f.endswith('.crdownload') and not f.endswith('.tmp') and not f.endswith('.part')]
                    
                    if actual_files:
                        print(f"\n✓ Successfully downloaded: {actual_files}")

            print("\n" + "=" * 60)
            print("ERP Report download process completed!")
            print("=" * 60 + "\n")
            return True

        except Exception as e:
            print(f"\nFailed to download ERP report: {e}")
            import traceback
            traceback.print_exc()

            # Save screenshot on error
            try:
                error_screenshot = f"erp_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                self.driver.save_screenshot(error_screenshot)
                print(f"Error screenshot saved to: {error_screenshot}")
            except:
                pass

            raise

    def run_full_automation(self, enc_username, enc_password, from_date, to_date):
        """
        Run the full automation process
        """
        try:
            # Decrypt credentials
            username, password = self.decrypt_credentials(enc_username, enc_password)

            # Initialize browser
            self.initialize_browser()

            # ========== PHASE 1: EFN Portal - Movement Reports ==========
            print("\n" + "=" * 60)
            print("PHASE 1: EFN Portal - Movement of Daily Deposit Balance")
            print("=" * 60)

            # Login
            self.login(username, password)

            # Navigate to report
            self.navigate_to_report()

            # Generate initial report
            self.generate_report(from_date, to_date)

            # Generate detailed report
            self.generate_detail_report()

            print("\n" + "=" * 60)
            print("Phase 1 completed successfully!")
            print("=" * 60)

            # ========== PHASE 2: Daily Interest Payable Reports ==========
            print("\n" + "=" * 60)
            print("PHASE 2: Daily Interest Payable Reports")
            print("=" * 60)

            # Generate list of dates for daily reports
            date_list = self.get_date_list(from_date, to_date)
            print(f"\nPrepared {len(date_list)} dates for daily report download:")
            for i, date in enumerate(date_list, 1):
                print(f"  {i}. {date}")

            # Download daily interest reports
            self.download_daily_interest_reports(date_list)

            print("\n" + "=" * 60)
            print("Phase 2 completed successfully!")
            print("=" * 60)

            # ========== PHASE 3: New Deposit Base Reports ==========
            print("\n" + "=" * 60)
            print("PHASE 3: New Deposit Base Reports (Days 1-8)")
            print("=" * 60)

            print(f"\nUsing same {len(date_list)} dates for New Deposit Base reports:")
            for i, date in enumerate(date_list, 1):
                print(f"  {i}. {date}")

            # Download New Deposit Base reports
            self.download_new_deposit_base_reports(date_list)

            print("\n" + "=" * 60)
            print("Phase 3 completed successfully!")
            print("=" * 60)

            # ========== PHASE 4: ERP Report ==========
            print("\n" + "=" * 60)
            print("PHASE 4: ERP System Report")
            print("=" * 60)

            # Get the last date (day 8)
            last_date = self.get_last_date(from_date, to_date)
            print(f"Using date: {last_date} (Day 8 - Last Sunday)")

            # Download ERP report
            self.download_erp_report(username, password, last_date)

            print("\n" + "=" * 60)
            print("Phase 4 completed successfully!")
            print("=" * 60)

            # ========== COMPLETION ==========
            print("\n" + "=" * 70)
            print("🎉 FULL AUTOMATION COMPLETED SUCCESSFULLY! 🎉")
            print("=" * 70)
            print("\nSummary:")
            print(f"  ✓ Movement of Daily Deposit Balance Report")
            print(f"  ✓ {len(date_list)} Daily Interest Payable Reports")
            print(f"  ✓ {len(date_list)} New Deposit Base Reports")
            print(f"  ✓ ERP Detail Report for {last_date}")
            print(f"\nTotal Reports Downloaded: {2 * len(date_list) + 2}")
            print("=" * 70 + "\n")

            time.sleep(5)  # Keep browser open to view results

            return True

        except Exception as e:
            print(f"\n{'=' * 60}")
            print(f"An error occurred: {str(e)}")
            print("=" * 60)
            import traceback
            traceback.print_exc()
            return False

    def close_browser(self):
        """
        Close the browser
        """
        if self.driver:
            print("Closing browser...")
            self.driver.quit()
            print("Browser closed")


def main():
    """
    Main function to run the bot

    Usage:
        python efn_bot.py YYYY/MM/DD

        Where YYYY/MM/DD is the Monday date.
        The script will automatically calculate:
        - FROM date: Sunday 8 days before the Monday
        - TO date: Sunday 1 day before the Monday
        
        Then it will:
        1. Download Movement of Daily Deposit Balance reports
        2. Download daily Interest Payable reports for each day (8 reports)
        3. Download New Deposit Base reports for each day (8 reports)
        4. Download ERP Detail Report for the last date (day 8)

    Example:
        python efn_bot.py 2025/10/13
        This will use date range: 2025/10/05 to 2025/10/12
        Download 8 Interest reports + 8 Deposit Base reports + 1 ERP report = 17 reports total
    """
    # Check for command-line argument
    if len(sys.argv) != 2:
        print("=" * 60)
        print("Error: Missing date argument")
        print("=" * 60)
        print("\nUsage:")
        print("  python efn_bot.py YYYY/MM/DD")
        print("\nExample:")
        print("  python efn_bot.py 2025/10/13")
        print("\nThe date should be a Monday.")
        print("The script will calculate the date range and download all required reports.")
        print("=" * 60)
        sys.exit(1)

    monday_date = sys.argv[1]

    # Load environment variables from .env file if using python-dotenv
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        print("python-dotenv not installed. Make sure to set environment variables manually.")

    # Get encrypted credentials from environment variables
    ENC_USERNAME = os.getenv('ENC_USERNAME')
    ENC_PASSWORD = os.getenv('ENC_PASSWORD')

    if not ENC_USERNAME or not ENC_PASSWORD:
        print("=" * 60)
        print("Error: ENC_USERNAME and ENC_PASSWORD environment variables not set")
        print("Please set them in your .env file or environment")
        print("=" * 60)
        return

    # Create bot instance
    bot = EFNBot()

    try:
        print("\n" + "=" * 60)
        print("EFN Bot - Automated Report Generation")
        print("=" * 60 + "\n")

        # Calculate date range from the Monday argument
        from_date, to_date = bot.calculate_date_range(monday_date)

        print("\n" + "=" * 60)
        print("Starting automation...")
        print("=" * 60 + "\n")

        # Run automation with calculated dates
        success = bot.run_full_automation(ENC_USERNAME, ENC_PASSWORD, from_date, to_date)

        if success:
            print("\n" + "=" * 60)
            print("Bot execution completed successfully")
            print("=" * 60)
        else:
            print("\n" + "=" * 60)
            print("Bot execution failed")
            print("=" * 60)

    except Exception as e:
        print(f"\nFatal error: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Close browser after completion
        bot.close_browser()
        print("\nBrowser closed. Automation complete!")


if __name__ == "__main__":
    main()