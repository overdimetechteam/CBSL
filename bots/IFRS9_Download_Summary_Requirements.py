from erp_login import erp_login
import os
import sys
import time
from datetime import datetime
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


def download_backendDueTypes(driver, wait, date1, date2):

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
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[1]/input"))).send_keys('Backend Due Types - Contract wise Report - Internal')

    print("Searched for Backend Due Types - Contract wise Report - Internal")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[2]/div[2]/div[1]/ul/li/a"))).click()
    time.sleep(5)
    print("Clicked on Backend Due Types - Contract wise Report - Internal")
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/div[1]/div/div[1]"))).click()
    print("Clicked on Download button")
    time.sleep(5)

    dropdown = Select(driver.find_element("xpath", "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[1]/td[3]/select"))
    dropdown.select_by_visible_text("=")
    print("Selected = from dropdown")
    time.sleep(5)
    coa =wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[1]/td[4]/input")))
    coa.clear()
    coa.send_keys("0189")
    print("Entered values in input field")
    time.sleep(5)
    date_range = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[4]/td[4]/input")))
    date_range.clear()
    date_range.send_keys(f"'{date1}','{date2}'")
    print("Entered date range in input field")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[1]/div[5]/fieldset/div/div/button"))).click()
    time.sleep(20) 
    print("Download should be complete now")

def download_allContractSector(driver, wait):
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
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[1]/input"))).send_keys('All Contracts Sector Report - Internal')

    print("Searched for All Contracts Sector Report - Internal")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[2]/div[2]/div[1]/ul/li/a"))).click()
    time.sleep(5)
    print("Clicked on All Contracts Sector Report - Internal")
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/div[1]/div/div[1]"))).click()
    print("Clicked on Download button")
    time.sleep(5)
    # Unselect All 
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[1]/div[2]/div/div/div/span[1]"))).click()
    print("Unselected All options")
    # Select Contract Sector
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[1]"))).click()
    print("Selected Contract Sector")
    #select Equipment Sector
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[7]/td[1]"))).click()
    print("Selected Equipment Sector")
    #Select Equipment Description
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[2]/div/div/div[2]/div[2]/table/tbody/tr[9]/td[1]"))).click()
    print("Selected Equipment Description")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[1]/div[5]/fieldset/div/div/button"))).click()
    time.sleep(20) 
    print("Download should be complete now")


def download_Current_DPD_REport(driver, wait):
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
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[1]/input"))).send_keys('Current DPD Report New')

    print("Searched for Current DPD Report New")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[2]/div[2]/div[1]/ul/li/a"))).click()
    time.sleep(5)
    print("Clicked on Current DPD Report New")
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/div[1]/div/div[1]"))).click()
    print("Clicked on Download button")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[1]/div[5]/fieldset/div/div/button"))).click()
    time.sleep(20) 
    print("Download should be complete now")


def download_Month_DPD_REport(driver, wait):
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
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[1]/input"))).send_keys('Month DPD Report New')

    print("Searched for Month DPD Report New")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[2]/div[2]/div[2]/div[1]/ul/li/a"))).click()
    time.sleep(5)
    print("Clicked on Month DPD Report New")
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/div[1]/div/div[1]"))).click()
    print("Clicked on Download button")
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div/div[1]/div[2]/form/div/div/div[2]/div[1]/div[5]/fieldset/div/div/button"))).click()
    time.sleep(20) 
    print("Download should be complete now")


def run_main(date2):
    date1 = datetime.strptime(date2, "%d/%m/%Y").replace(day=1).strftime("%d/%m/%Y")
    driver, wait = erp_login()
    download_backendDueTypes(driver, wait, date1, date2)
    driver.quit()
    driver, wait = erp_login()
    download_allContractSector(driver, wait)
    driver.quit()
    driver, wait = erp_login()
    download_Current_DPD_REport(driver, wait)
    driver.quit()
    driver, wait = erp_login()
    download_Month_DPD_REport(driver, wait)
    time.sleep(10)
    return True

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python Download_LedgerAccounts.py <date2>")
        sys.exit(1)
    date2 = sys.argv[1]
    run_main(date2)
    driver.quit()
    print("✓ Script completed successfully")