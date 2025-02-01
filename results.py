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
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd(),
        "plugins.always_open_pdf_externally": True
    })
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def wait_and_find_element(driver, by, value, timeout=20):
    try:
        return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    except TimeoutException:
        print(f"Timeout: Element {value} not found.")
        return None

def extract_seat_mother_info_from_excel(excel_path):
    if not os.path.exists(excel_path):
        print(f"Excel file not found: {excel_path}")
        return []
    
    try:
        # Read Excel file, skipping the first row
        df = pd.read_excel(excel_path, header=1)

        # Ensure necessary columns exist
        if 'Seat No' not in df.columns or 'Mother Name' not in df.columns:
            print("Excel file must contain 'Seat No' and 'Mother Name' columns.")
            return []

        # Extract all rows, removing any with NaN values
        seat_mother_list = df[['Seat No', 'Mother Name']].dropna().values.tolist()
        return seat_mother_list

    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return []

def download_results(driver, seat_mother_list):
    for seat_no, mother_name in seat_mother_list:
        try:
            print(f"Processing {seat_no}")

            # Navigate back to the input page for each new entry
            try:
                back_button = wait_and_find_element(driver, By.XPATH, "//input[@value='Back']")
                if back_button:
                    back_button.click()
                    time.sleep(3)
            except:
                # If we can't find back button, we might already be on the input page
                pass

            # Wait for input fields
            seat_input = wait_and_find_element(driver, By.ID, "SeatNo")
            mother_input = wait_and_find_element(driver, By.ID, "MotherName")
            check_button = wait_and_find_element(driver, By.XPATH, "//*[@id='btn']")

            if not seat_input or not mother_input or not check_button:
                print(f"Required input fields not found for {seat_no}. Skipping...")
                continue

            # Clear and enter values
            seat_input.clear()
            seat_input.send_keys(str(seat_no))
            time.sleep(2)
            mother_input.clear()
            mother_input.send_keys(str(mother_name))
            time.sleep(2)

            check_button.click()
            time.sleep(5)

            # Wait for the page URL to change to ViewResult1
            WebDriverWait(driver, 10).until(EC.url_contains("ViewResult1"))

            # Wait for the download button
            download_button = wait_and_find_element(driver, By.XPATH, "//*[contains(@href, '.pdf')]", timeout=20)
            
            if download_button:
                download_button.click()
                print(f"Downloaded result for {seat_no}")
                time.sleep(10)  # Wait for download to complete
            else:
                print(f"Download button not found for {seat_no}.")

        except Exception as e:
            print(f"Error processing {seat_no}: {str(e)}")
            continue

def main():
    driver = None
    try:
        print("Starting browser...")
        driver = setup_driver()

        # Open Results Page
        URL = "https://onlineresults.unipune.ac.in/Result/Dashboard/Default"
        print("Loading results page...")
        driver.get(URL)
        time.sleep(5)

        # Click Topmost "Go for Result" Button
        print("Looking for result link...")
        topmost_result = wait_and_find_element(driver, By.XPATH, "//*[@id='tblRVList']/tbody/tr[1]/td[4]/a/input")
        if topmost_result:
            topmost_result.click()
            time.sleep(5)
        else:
            print("Could not find the 'Go for Result' button.")
            return

        # Read Seat No & Mother Name from Excel
        print("Reading input Excel file...")
        seat_mother_list = extract_seat_mother_info_from_excel("input.xlsx")
        
        seat_mother_list = seat_mother_list[93:]

        if seat_mother_list:
            print(f"Found {len(seat_mother_list)} entries to process")
            download_results(driver, seat_mother_list)
        else:
            print("No seat/mother information found in input.xlsx")

    except Exception as e:
        print(f"Main execution error: {str(e)}")

    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()