import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
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
        df = pd.read_excel(excel_path, header=1)
        if 'Seat No' not in df.columns or 'Mother Name' not in df.columns:
            print("Excel file must contain 'Seat No' and 'Mother Name' columns.")
            return []
        return df[['Seat No', 'Mother Name']].dropna().values.tolist()
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return []

def process_entry(driver, input_form_url, seat_no, mother_name):
    try:
        # Create new tab for this entry
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        driver.get(input_form_url)
        time.sleep(3)

        # Fill form
        seat_input = wait_and_find_element(driver, By.ID, "SeatNo")
        mother_input = wait_and_find_element(driver, By.ID, "MotherName")
        check_button = wait_and_find_element(driver, By.XPATH, "//*[@id='btn']")

        if not all([seat_input, mother_input, check_button]):
            print(f"Form elements not found for {seat_no}")
            return False

        seat_input.clear()
        seat_input.send_keys(str(seat_no))
        mother_input.clear()
        mother_input.send_keys(str(mother_name))
        check_button.click()
        time.sleep(3)

        # Handle result page
        WebDriverWait(driver, 10).until(EC.url_contains("ViewResult1"))
        
        # Trigger download
        if download_button := wait_and_find_element(driver, By.XPATH, "//*[contains(@href, '.pdf')]"):
            main_window = driver.current_window_handle
            download_button.click()
            time.sleep(5)  # Allow download to initiate

            # Close any new tabs opened by PDF viewer
            for handle in driver.window_handles:
                if handle != main_window:
                    driver.switch_to.window(handle)
                    driver.close()
            driver.switch_to.window(main_window)
            
            print(f"Successfully processed {seat_no}")
            return True
        return False
    except Exception as e:
        print(f"Error in process_entry for {seat_no}: {str(e)}")
        return False
    finally:
        # Always close the processing tab
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

def main():
    driver = None
    try:
        driver = setup_driver()
        driver.get("https://onlineresults.unipune.ac.in/Result/Dashboard/Default")
        time.sleep(3)

        # Get to initial input form
        if go_button := wait_and_find_element(driver, By.XPATH, "//*[@id='tblRVList']/tbody/tr[1]/td[4]/a/input"):
            go_button.click()
            time.sleep(3)
            input_form_url = driver.current_url
            print(f"Input form URL: {input_form_url}")
        else:
            print("Initial form not found")
            return

        # Process entries
        entries = extract_seat_mother_info_from_excel("input.xlsx")[155:]
        print(f"Found {len(entries)} entries to process")

        for idx, (seat_no, mother_name) in enumerate(entries, 1):
            print(f"Processing entry {idx}/{len(entries)}: {seat_no}")
            process_entry(driver, input_form_url, seat_no, mother_name)
            time.sleep(2)  # Brief pause between entries

    except Exception as e:
        print(f"Main error: {str(e)}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()