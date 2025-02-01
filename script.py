import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

def setup_driver():
    chrome_options = Options()
    # For debugging, you can comment out the headless option:
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # Disable extensions that might trigger TensorFlow errors
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-software-rasterizer")
    
    # Use a user-agent to avoid detection (if needed)
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36")
    
    # Ensure PDFs download instead of opening in Chrome's built-in viewer
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd(),
        "plugins.always_open_pdf_externally": True,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    })
    
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def wait_and_find_element(driver, by, value, timeout=30):
    """
    Wait for an element to be visible on the page and return it.
    Using 'visibility_of_element_located' helps ensure that the element
    is not only present in the DOM but is also visible.
    """
    try:
        return WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((by, value))
        )
    except TimeoutException:
        print(f"❌ Timeout: Element {value} not found.")
        return None

def extract_seat_mother_info_from_excel(excel_path, processed_entries):
    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        return []
    
    try:
        df = pd.read_excel(excel_path, header=1)
        if 'Seat No' not in df.columns or 'Mother Name' not in df.columns:
            print("❌ Excel file must contain 'Seat No' and 'Mother Name' columns.")
            return []
        seat_mother_list = df[['Seat No', 'Mother Name']].dropna().values.tolist()
        # Skip already processed entries
        seat_mother_list = [entry for entry in seat_mother_list if entry[0] not in processed_entries]
        return seat_mother_list
    except Exception as e:
        print(f"❌ Error reading Excel file: {str(e)}")
        return []

def download_results(driver, seat_mother_list):
    for seat_no, mother_name in seat_mother_list:
        try:
            print(f"🔍 Processing {seat_no}")
            
            # Wait for input fields using a longer timeout
            seat_input = wait_and_find_element(driver, By.ID, "SeatNo", timeout=30)
            mother_input = wait_and_find_element(driver, By.ID, "MotherName", timeout=30)
            check_button = wait_and_find_element(driver, By.XPATH, "//*[@id='btn']", timeout=30)
            
            if not seat_input or not mother_input or not check_button:
                print(f"❌ Required input fields not found for {seat_no}. Skipping...")
                continue

            # Clear and enter values
            seat_input.clear()
            seat_input.send_keys(str(seat_no))
            time.sleep(2)
            mother_input.clear()
            mother_input.send_keys(str(mother_name))
            time.sleep(2)
            
            # Click the check button
            check_button.click()
            time.sleep(5)
            
            # Wait for the page URL to change to include "ViewResult1"
            WebDriverWait(driver, 30).until(EC.url_contains("ViewResult1"))
            
            # Wait for the download button (PDF link) to become visible
            download_button = wait_and_find_element(driver, By.XPATH, "//*[contains(@href, '.pdf')]", timeout=30)
            if download_button:
                download_button.click()
                print(f"✅ Downloaded result for {seat_no}")
                time.sleep(10)  # Wait for the download to complete
            else:
                print(f"❌ Download button not found for {seat_no}. Skipping.")
            
        except WebDriverException as e:
            print(f"❌ WebDriver error processing {seat_no}: {str(e)}")
            print("Current URL:", driver.current_url)
            # Save a screenshot for debugging
            driver.save_screenshot(f"error_{seat_no}.png")
            # Optionally restart the driver if the session is invalid
            if 'invalid session id' in str(e):
                driver.quit()
                driver = setup_driver()
            continue
        except Exception as e:
            print(f"❌ Error processing {seat_no}: {str(e)}")
            print("Current URL:", driver.current_url)
            driver.save_screenshot(f"error_{seat_no}.png")
            continue

def main():
    driver = None
    processed_entries = set()  # Set to store processed seat numbers

    try:
        # Read previously processed entries (if any)
        if os.path.exists("processed_entries.txt"):
            with open("processed_entries.txt", "r") as f:
                processed_entries = set(f.read().splitlines())
        
        print("🚀 Starting browser...")
        driver = setup_driver()
        
        # Open Results Page
        URL = "https://onlineresults.unipune.ac.in/Result/Dashboard/Default"
        print("🌐 Loading results page...")
        driver.get(URL)
        time.sleep(5)
        
        # Click Topmost "Go for Result" Button
        print("🔍 Looking for result link...")
        topmost_result = wait_and_find_element(driver, By.XPATH, "//*[@id='tblRVList']/tbody/tr[2]/td[4]/a/input", timeout=30)
        if topmost_result:
            topmost_result.click()
            time.sleep(5)
        else:
            print("❌ Could not find the 'Go for Result' button.")
            return
        
        # Read Seat No & Mother Name from Excel
        print("📂 Reading input Excel file...")
        seat_mother_list = extract_seat_mother_info_from_excel("input.xlsx", processed_entries)
        
        # Start processing from the 115th entry
        seat_mother_list = seat_mother_list[177:]
        
        if seat_mother_list:
            print(f"📊 Found {len(seat_mother_list)} entries to process (starting from 115th)")
            download_results(driver, seat_mother_list)
            
            # After processing, update the processed entries file
            with open("processed_entries.txt", "a") as f:
                for seat_no, _ in seat_mother_list:
                    f.write(f"{seat_no}\n")
        else:
            print("❌ No seat/mother information found in input.xlsx")
    
    except Exception as e:
        print(f"❌ Main execution error: {str(e)}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
