from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
import os
import subprocess
from dotenv import load_dotenv
from cryptography.fernet import Fernet

class ERPLoginBot:
    def __init__(self, download_folder=None):
        # Close all existing Chrome browser instances to prevent crashes
        self._close_all_browsers()

        # Set download folder
        if download_folder is None:
            download_folder = os.path.join(os.getcwd(), "downloads")
        
        # Create download folder if it doesn't exist
        os.makedirs(download_folder, exist_ok=True)
        self.download_folder = os.path.abspath(download_folder)
        
        # Configure Chrome options
        self.chrome_options = Options()
        # Uncomment the next line to run headless (without opening browser window)
        # self.chrome_options.add_argument("--headless")
        self.chrome_options.add_argument("--no-sandbox")
        self.chrome_options.add_argument("--disable-dev-shm-usage")
        self.chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Set download preferences
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        self.chrome_options.add_experimental_option("prefs", prefs)
        
        self.driver = None
        print(f"Download folder set to: {self.download_folder}")

    def _close_all_browsers(self):
        """
        Close all existing Chrome and ChromeDriver instances to prevent crashes
        """
        try:
            print("Closing all existing browser instances...")

            # Kill Chrome browser processes
            subprocess.run(
                ["taskkill", "/F", "/IM", "chrome.exe"],
                capture_output=True,
                timeout=10
            )

            # Kill ChromeDriver processes
            subprocess.run(
                ["taskkill", "/F", "/IM", "chromedriver.exe"],
                capture_output=True,
                timeout=10
            )

            # Give processes time to fully terminate
            time.sleep(1)
            print("Browser cleanup completed")

        except subprocess.TimeoutExpired:
            print("Browser cleanup timed out, continuing anyway...")
        except Exception as e:
            print(f"Browser cleanup warning (non-fatal): {e}")

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
                raise ValueError("FERNET_KEY environment variable not set")

            key = key_str.encode()
            fernet = Fernet(key)
            
            # Decrypt credentials
            username = fernet.decrypt(enc_username.encode()).decode()
            password = fernet.decrypt(enc_password.encode()).decode()
            
            print("Successfully decrypted credentials")
            return username, password
            
        except Exception as e:
            print(f"Failed to decrypt credentials: {e}")
            raise
        
    def login_to_erp(self, username, password):
        """
        Login to ERP system
        """
        try:
            # Initialize Chrome driver
            self.driver = webdriver.Chrome(options=self.chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            # Navigate to ERP site
            print("Opening ERP website...")
            self.driver.get("https://erp.assetline.lk")
            
            # Wait for page to load
            wait = WebDriverWait(self.driver, 15)
            
            # Wait for username field and enter username
            print("Entering username...")
            username_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="strUserName"]')))
            username_field.clear()
            username_field.send_keys(username)
            
            # Enter password
            print("Entering password...")
            password_field = self.driver.find_element(By.XPATH, '//*[@id="strPassword"]')
            password_field.clear()
            password_field.send_keys(password)
            
            # Click sign in button
            print("Clicking sign in...")
            sign_in_button = self.driver.find_element(By.XPATH, '//*[@id="idSignIn"]')
            sign_in_button.click()
            
            # Wait for login to complete
            time.sleep(3)
            
            print("Login successful!")
            return True
                
        except Exception as e:
            print(f"Error during login: {str(e)}")
            return False
    
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
    
    def wait_for_download_completion(self, timeout=500):
        """
        Wait for TB-COA file to be downloaded
        Looks for files matching pattern: YYYY_MM_DD_HH_MM_Report.xlsx

        Args:
            timeout (int): Maximum time to wait in seconds (default: 120)

        Returns:
            bool: True if file downloaded, False if timeout
        """
        try:
            print(f"Monitoring download folder: {self.download_folder}")
            print("Looking for TB-COA Report file (pattern: YYYY_MM_DD_HH_MM_Report.xlsx)...")

            # Get list of existing files before download
            existing_files = set(os.listdir(self.download_folder))

            start_time = time.time()
            check_interval = 1  # Check every 1 second

            while time.time() - start_time < timeout:
                # Get current files in download folder
                current_files = set(os.listdir(self.download_folder))

                # Find new files
                new_files = current_files - existing_files

                for filename in new_files:
                    # Check if it matches TB-COA report pattern
                    # Pattern: YYYY_MM_DD_HH_MM_Report.xlsx or similar
                    if 'Report' in filename and filename.endswith('.xlsx'):
                        # Check if file is not a temporary download file
                        if not filename.endswith('.tmp') and not filename.endswith('.crdownload'):
                            file_path = os.path.join(self.download_folder, filename)

                            # Check if file is fully downloaded (size is stable)
                            try:
                                size1 = os.path.getsize(file_path)
                                time.sleep(1)
                                size2 = os.path.getsize(file_path)

                                if size1 == size2 and size1 > 0:
                                    print(f"✓ Downloaded: {filename}")
                                    print(f"  File size: {size1:,} bytes")
                                    print(f"  Location: {file_path}")
                                    return True
                            except:
                                # File might still be downloading
                                pass

                # Show progress indicator
                elapsed = int(time.time() - start_time)
                if elapsed % 5 == 0:  # Print every 5 seconds
                    print(f"  Waiting... ({elapsed}s / {timeout}s)")

                time.sleep(check_interval)

            # Timeout reached
            elapsed = int(time.time() - start_time)
            print(f"✗ Timeout reached after {elapsed} seconds")
            print(f"  No TB-COA Report file detected in: {self.download_folder}")

            return False

        except Exception as e:
            print(f"Error waiting for download: {str(e)}")
            return False

    def navigate_to_tb_report(self):
        """
        Navigate to TB-COA report section
        """
        try:
            wait = WebDriverWait(self.driver, 15)
            
            # Navigate directly to Finance page
            print("Navigating to Finance page...")
            self.driver.get("https://erp.assetline.lk/Application/Home/FINANCE")
            
            # Wait for Finance page to load
            print("Waiting for Finance page to load...")
            time.sleep(5)
            
            # Step 6: Navigate to specific section
            print("Navigating to Management Account section...")
            section = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="iddivAppContent"]/div[1]/div/div[1]/div[2]/div/div[1]/div')))
            section.click()
            time.sleep(2)
            
            # Step 7: Select TB-COA inquiry
            print("Opening TB-COA inquiry...")
            tb_inquiry = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_03832_Inquiry"]')))
            tb_inquiry.click()
            time.sleep(2)
            
            print("Successfully navigated to TB-COA report!")
            return True
            
        except Exception as e:
            print(f"Error navigating to TB report: {str(e)}")
            return False
    
    def generate_tb_report(self, report_date="28/09/2025"):
        """
        Generate and download TB report
        Args:
            report_date (str): Date in DD/MM/YYYY format
        """
        try:
            wait = WebDriverWait(self.driver, 15)

            # Step 8: Select date field and enter date using the correct XPath
            print(f"Entering report date: {report_date}...")

            # Wait for the date input container to be present
            date_container = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="plt_03832_Main"]/div[8]/div[2]/div/div')))

            # Additional wait to ensure page is fully loaded and no overlays
            time.sleep(2)

            # Try to find the actual input field within the container
            # It could be a direct input or nested within the div
            try:
                # Try finding input field directly within the container
                date_field = date_container.find_element(By.TAG_NAME, 'input')
            except:
                # If no input found, use the container itself
                date_field = date_container

            print(f"Found date field, entering date: {report_date}")

            # Scroll the element into view to avoid click interception
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", date_field)
            time.sleep(0.5)

            # Try clicking with JavaScript first to avoid interception
            try:
                self.driver.execute_script("arguments[0].click();", date_field)
                print("Clicked date field using JavaScript")
                time.sleep(0.5)
            except Exception as e:
                print(f"JavaScript click failed: {e}")
                # Fallback to regular click
                try:
                    date_field.click()
                    print("Clicked date field using Selenium")
                except Exception as click_error:
                    print(f"Regular click also failed: {click_error}")
                    print("Attempting to focus field using Tab navigation...")
                    # Try navigating to the field using Tab
                    self.driver.switch_to.active_element.send_keys(Keys.TAB)
                    time.sleep(0.3)

            time.sleep(1)

            # Try JavaScript method first (most reliable for automated scripts)
            try:
                print("Attempting to set date using JavaScript...")
                self.driver.execute_script(f"arguments[0].value = '{report_date}';", date_field)
                self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", date_field)
                self.driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", date_field)
                print("Date set using JavaScript")
                time.sleep(1)
            except Exception as js_error:
                print(f"JavaScript method failed: {js_error}")
                print("Falling back to keyboard input...")

                # Clear any existing value
                try:
                    date_field.clear()
                    time.sleep(0.5)
                except:
                    # If clear doesn't work, try selecting all and deleting
                    date_field.send_keys(Keys.CONTROL + "a")
                    time.sleep(0.3)
                    date_field.send_keys(Keys.DELETE)
                    time.sleep(0.5)

                # Enter the date
                date_field.send_keys(report_date)
                time.sleep(1)

                # Press Tab to confirm
                date_field.send_keys(Keys.TAB)
                time.sleep(0.5)

            print(f"✓ Date {report_date} entered successfully")
            time.sleep(2)
            
            # Step 9: Click download button
            print("Clicking Generate Report button...")
            download_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_03832_DownloadExcelReport_FINTB_COA"]')))
            download_button.click()

            print("\n" + "="*60)
            print("Report download initiated!")
            print("Waiting for file to be downloaded...")
            print("="*60 + "\n")

            # Wait for the TB-COA file to appear in the download folder
            downloaded = self.wait_for_download_completion()

            if downloaded:
                print("✓ TB-COA file downloaded successfully!")
                print("Closing browser...")
                time.sleep(2)  # Small delay before closing
                return True
            else:
                print("✗ Download timeout - file not detected")
                return False
            
        except Exception as e:
            print(f"Error generating TB report: {str(e)}")
            return False
    
    def run_full_automation(self, enc_username, enc_password, report_date="28/09/2025"):
        """
        Run the complete automation workflow
        """
        try:
            # Step 1: Decrypt credentials
            username, password = self.decrypt_credentials(enc_username, enc_password)
            
            # Step 2: Login
            if not self.login_to_erp(username, password):
                print("Login failed!")
                return False
            
            # Step 3: Click Select button only
            if not self.select_premises():
                print("Premises selection failed!")
                return False
            
            # Step 4: Navigate to TB report
            if not self.navigate_to_tb_report():
                print("Navigation to TB report failed!")
                return False
            
            # Step 5: Generate and download report
            if not self.generate_tb_report(report_date):
                print("Report generation failed!")
                return False
            
            print("\n" + "="*60)
            print("SUCCESS: TB Report automation completed!")
            print("="*60)
            return True
            
        except Exception as e:
            print(f"Error in automation workflow: {str(e)}")
            return False
            
    def close_browser(self):
        """
        Close the browser
        """
        if self.driver:
            try:
                print("Closing browser...")
                self.driver.quit()
            except:
                pass

def run_tb_download(input_date, download_folder):
    """
    Run TB report download with specified date and download folder
    Args:
        input_date (datetime or str): Date for the report (DD/MM/YYYY format or datetime object)
        download_folder (str): Path to download folder
    Returns:
        bool: True if successful, False otherwise
    """
    # Load environment variables from .env file in parent directory
    env_path = os.path.join(os.path.dirname(__file__), '..', '.env')
    load_dotenv(dotenv_path=env_path)

    # Get encrypted credentials from environment variables
    ENC_USERNAME = os.getenv('ENC_USERNAME')
    ENC_PASSWORD = os.getenv('ENC_PASSWORD')

    # Check if credentials were loaded
    if not ENC_USERNAME or not ENC_PASSWORD:
        print("ERROR: ENC_USERNAME or ENC_PASSWORD not found in environment variables")
        return False

    # Convert datetime to DD/MM/YYYY format if needed
    if hasattr(input_date, 'strftime'):
        report_date = input_date.strftime("%d/%m/%Y")
    else:
        report_date = input_date

    print(f"TB Report Date: {report_date}")
    print(f"Download Folder: {download_folder}")

    # Clear existing TB_DetailReport.xlsx files in download folder
    try:
        if os.path.exists(download_folder):
            for filename in os.listdir(download_folder):
                if "TB_DetailReport" in filename and filename.endswith('.xlsx'):
                    file_path = os.path.join(download_folder, filename)
                    os.remove(file_path)
                    print(f"Removed existing file: {filename}")
    except Exception as e:
        print(f"Warning: Could not clear existing files: {e}")

    # Create bot instance
    bot = ERPLoginBot(download_folder=download_folder)

    try:
        # Run full automation with encrypted credentials
        success = bot.run_full_automation(ENC_USERNAME, ENC_PASSWORD, report_date)

        if success:
            print("\nERP TB report automation completed successfully!")
        else:
            print("\nERP TB report automation failed!")

        return success

    except Exception as e:
        print(f"Script error: {str(e)}")
        return False

    finally:
        # Always close browser (if not already closed)
        bot.close_browser()

def main():
    """Legacy main function for backward compatibility"""
    import sys
    from datetime import datetime

    # Use current date if no args provided
    if len(sys.argv) > 1:
        # Try to parse date from command line
        try:
            input_date = datetime.strptime(sys.argv[1], "%m/%d/%Y")
        except:
            print("Usage: python tb_from_erp_bot.py [MM/DD/YYYY]")
            return
    else:
        input_date = datetime.now()

    # Use current directory downloads folder as default
    download_folder = os.path.join(os.getcwd(), "downloads")

    run_tb_download(input_date, download_folder)

if __name__ == "__main__":
    main()