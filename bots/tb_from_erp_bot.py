from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
import os
from dotenv import load_dotenv
from cryptography.fernet import Fernet

class ERPLoginBot:
    def __init__(self, download_folder=None):
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
    
    # def select_premises(self):
    #     """
    #     Just click the SELECT button to proceed
    #     """
    #     try:
    #         wait = WebDriverWait(self.driver, 20)
            
    #         # Wait a bit for the page/modal to load
    #         time.sleep(3)
            
    #         # Just click the Select button directly
    #         print("Clicking Select button...")
    #         select_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/form/div/button')))
    #         select_button.click()
            
    #         # Wait for page to load after selection
    #         time.sleep(3)
            
    #         # Refresh the page
    #         print("Refreshing the page...")
    #         self.driver.refresh()
            
    #         # Wait for page to fully reload
    #         print("Waiting for page to fully reload...")
    #         time.sleep(3)
            
    #         print("Premises selection completed!")
    #         return True
            
    #     except Exception as e:
    #         print(f"Error clicking Select button: {str(e)}")
    #         return False
    
    def wait_for_download_completion(self):
        """
        Wait for TB-COA file to be downloaded
        Looks for files matching pattern: YYYY_MM_DD_HH_MM_Report.xlsx

        Waits indefinitely until the file is detected and fully downloaded

        Returns:
            bool: True if file downloaded successfully
        """
        try:
            print(f"Monitoring download folder: {self.download_folder}")
            print("Looking for TB-COA Report file (pattern: YYYY_MM_DD_HH_MM_Report.xlsx)...")
            print("Waiting indefinitely until download is detected...")

            # Get list of existing files before download
            existing_files = set(os.listdir(self.download_folder))

            start_time = time.time()
            check_interval = 1  # Check every 1 second

            while True:  # Wait indefinitely
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
                if elapsed % 10 == 0:  # Print every 10 seconds
                    print(f"  Still waiting... ({elapsed}s elapsed)")

                time.sleep(check_interval)

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

            # Try to find the actual input field within the container
            # It could be a direct input or nested within the div
            try:
                # Try finding input field directly within the container
                date_field = date_container.find_element(By.TAG_NAME, 'input')
            except:
                # If no input found, use the container itself
                date_field = date_container

            print(f"Found date field, entering date: {report_date}")

            # Click the field to focus it
            date_field.click()
            time.sleep(1)

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

            # Press Enter or Tab to confirm
            date_field.send_keys(Keys.TAB)
            time.sleep(0.5)

            # Also try using JavaScript to set the value as backup
            try:
                self.driver.execute_script(f"arguments[0].value = '{report_date}';", date_field)
                self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", date_field)
                self.driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", date_field)
            except:
                print("JavaScript backup method skipped")

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
            
            # # Step 3: Click Select button only
            # if not self.select_premises():
            #     print("Premises selection failed!")
            #     return False
            
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

def main():
    # Load environment variables from .env file in parent directory
    env_path = os.path.join(os.path.dirname(__file__), '..', '.env')
    load_dotenv(dotenv_path=env_path)
    
    # Debug: Print if .env file was found
    print(f"Looking for .env file at: {os.path.abspath(env_path)}")
    
    # Get encrypted credentials from environment variables
    ENC_USERNAME = os.getenv('ENC_USERNAME')
    ENC_PASSWORD = os.getenv('ENC_PASSWORD')
    
    # Debug: Check if credentials were loaded
    if ENC_USERNAME:
        print(f"ENC_USERNAME loaded: {ENC_USERNAME[:20]}...")
    else:
        print("WARNING: ENC_USERNAME not found in environment variables")
    
    if ENC_PASSWORD:
        print(f"ENC_PASSWORD loaded: {ENC_PASSWORD[:20]}...")
    else:
        print("WARNING: ENC_PASSWORD not found in environment variables")
    
    REPORT_DATE = "28/09/2025"  # DD/MM/YYYY format
    
    # Set download folder to the project's Input folder
    DOWNLOAD_FOLDER = r"C:\CBSL\Script\working\weekly\07-06-2025\NBD_MF_04_LA\Input"
    
    # Create bot instance
    bot = ERPLoginBot(download_folder=DOWNLOAD_FOLDER)
    
    try:
        # Run full automation with encrypted credentials
        success = bot.run_full_automation(ENC_USERNAME, ENC_PASSWORD, REPORT_DATE)
        
        if success:
            print("\nERP TB report automation completed successfully!")
        else:
            print("\nERP TB report automation failed!")
            
    except Exception as e:
        print(f"Script error: {str(e)}")
        
    finally:
        # Always close browser (if not already closed)
        bot.close_browser()

if __name__ == "__main__":
    main()