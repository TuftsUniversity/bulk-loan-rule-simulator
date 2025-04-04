import pandas as pd
import os
import glob
import threading
import requests
import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import secrets_local
import sys
# Add the full path to "scripts" based on current script location
current_dir = os.path.dirname(os.path.abspath(__file__))
scripts_dir = os.path.join(current_dir, "scripts")

if scripts_dir not in sys.path:
    sys.path.append(scripts_dir)
from functions import *



# --- Constants & Configurations ---
INPUT_DIR_ITEM_POLICY = "./input/Item Policies and Locations"
INPUT_DIR_USER_GROUP = "./input/User Groups"
OUTPUT_FILE = "Bulk_Checkout_Request_Results.xlsx"
OUTPUT_DIR = "Output"
BUFFER_WRITE_INTERVAL = 10  # Buffer size for writing to Excel

N = 3
# Shared index tracker
row_index_lock = threading.Lock()
row_index = 0
# --- Load Input Data ---
def load_first_excel(directory):
    files = glob.glob(os.path.join(directory, "*.xlsx"))
    if not files:
        print(f"No Excel files found in {directory}. Exiting.")
        exit()
    return pd.read_excel(files[0], dtype="str", engine="openpyxl")


item_policy_data = load_first_excel(INPUT_DIR_ITEM_POLICY)
user_group_data = load_first_excel(INPUT_DIR_USER_GROUP)

def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    return driver

def worker_thread(thread_id, combined_df):
    global row_index
    driver = init_driver()
    buffer = []
    current_user_id = None
    driver.get(secrets_local.alma_base_url)
    login(driver, secrets_local.username, secrets_local.password)
    time.sleep(20)
    try:
        modal = driver.find_element(
            By.XPATH, "//div[@id='onetrust-close-btn-container']//button"
        )
        print("GDPR modal detected. Attempting to close it.")
        modal = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@id='onetrust-close-btn-container']//button")
            )
        )
        modal.click()
        print("GDPR modal closed.")
    except TimeoutException:
        print("No GDPR modal detected.")

    except:
        print("No GDPR modal")

    
    driver.get("https://tufts.alma.exlibrisgroup.com/ng/page;u=%2Fful%2Faction%2FpageAction.do%3FxmlFileName%3Dtou.fulfillment_configuration_utility.xml&pageViewMode%3DEdit&operation%3DLOAD&backUrl%3D%2Fful%2Faction%2Fmenu.do%3F&pageBean.selectedTab%3DtouType.loan&pageBean.touType%3DLoan&pageBean.displayDueDate%3Dtrue&pageBean.displayReturnDate%3Dtrue&pageBean.currentUrl%3DxmlFileName%253Dtou.fulfillment_configuration_utility.xml%2526pageViewMode%253DEdit%2526operation%253DLOAD%2526backUrl%253D%252Fful%252Faction%252Fmenu.do%253F%2526pageBean.selectedTab%253DtouType.loan%2526pageBean.touType%253DLoan%2526pageBean.displayDueDate%253Dtrue%2526pageBean.displayReturnDate%253Dtrue%2526resetPaginationContext%253Dtrue%2526showBackButton%253Dfalse&pageBean.navigationBackUrl%3D..%252Faction%252Fhome.do&resetPaginationContext%3Dtrue&showBackButton%3Dfalse&menuKey%3Dcom.exlibris.dps.adm.general.menu.initial.Fulfillment.FulfillmentHeader.FulConfigurationUtility")

    while True:
        with row_index_lock:
            if row_index >= len(combined_df):
                break
            row = combined_df.iloc[row_index]
            row_index += 1

        user_id = row["Primary Identifier"]

        # Change user if needed
        if user_id != current_user_id:
            try:
                user_menu = safe_find_element(driver, By.ID, "PICKUP_ID_pageBeandisplayNameOfUserOrUserIdendifier")
                user_id = row["Primary Identifier"].strip()
                user_group = row["User Group"].strip()

                user_menu.click()

                modal = safe_find_element(driver, By.CLASS_NAME, "modal")
                driver.switch_to.frame(driver.find_element(By.ID, "iframePopupIframe"))

                search_button = safe_find_element(driver, By.ID, "simpleSearchIndexButton")
                search_button.click()

                time.sleep(2)
                primary_identifier_link = safe_find_element(driver, By.XPATH, "//li[@id='TOP_NAV_Search_index_HFrUser.user_name']//a[text()='Primary identifier']")
                primary_identifier_link.click()

                input_field = safe_find_element(driver, By.ID, "ALMA_MENU_TOP_NAV_Search_Text")
                input_field.send_keys(user_id)

                search_button = safe_find_element(driver, By.ID, "simpleSearchBtn")
                search_button.click()

                row_el = safe_find_element(driver, By.XPATH, "//table[@id='TABLE_DATA_userList']/tbody/tr")
                row_el.click()
                time.sleep(4)
                current_user_id = user_id
                print(driver.page_source)
            except Exception as e:
                print(f"Error switching user: {e}")
                continue

        

        
        try:
            barcode = row["Barcode"]
            item_policy = row["Item Policy"]
            location = (
                row["Temporary Location Name"]
                if row["Temporary Physical Location In Use"] == "Yes"
                else row["Location Name"]
            )

            print(f"Processing item {barcode} - {item_policy} - {location}")

            # Enter barcode (Refind element before interaction)
            
            item_field = safe_find_element(driver, By.XPATH, "//input[@id='pageBeanbarcode']")
            send_keys_with_retry(driver, By.XPATH, "//input[@id='pageBeanbarcode']", barcode)

            item_field = safe_find_element(driver, By.XPATH, "//input[@id='pageBeanbarcode']")
            send_keys_with_retry(driver, By.XPATH, "//input[@id='pageBeanbarcode']", barcode)
            click_element_with_retry(driver, By.ID, "cbuttonok")

            loan_result = get_table_html_with_retry(driver, By.ID, "TABLE_DATA_policiesList", "loan")
            request_result = get_table_html_with_retry(driver, By.ID, "TABLE_DATA_policiesList", "request")
            loan_fulfillment_rule_name = loan_result[0]
            loan_tou_name = loan_result[1]
            loan_dict = loan_result[2]
            request_policy_list = get_table_html_with_retry(driver, By.ID, "TABLE_DATA_policiesList", "request")
            # --- Extract Request Tab Data ---

            request_fulfillment_rule_name = request_policy_list[0]
            request_tou_name = request_policy_list[1]
            request_dict = request_policy_list[2]
            fulfillment_unit_name = loan_result[3]


            row_dict = {
                    "User ID": user_id,
                    "User Group": user_group,
                    "Barcode": barcode,
                    "Item Policy": item_policy,
                    "Location": location,
                    "Fulfillment Unit Name": fulfillment_unit_name,
                    "Fulfillment Rule (Loan)": loan_fulfillment_rule_name,
                    "TOU (Loan)": loan_tou_name,
                    "Fulfillment Rule (Request)": request_fulfillment_rule_name,
                    "TOU (Request)": request_tou_name,
                    
                }
            row_dict.update(loan_dict)
            row_dict.update(request_dict)
            buffer.append(row_dict)
            buffer.append(row_dict)
        except Exception as e:
            print(f"Thread-{thread_id} failed to process item {barcode}: {e}")

        if len(buffer) >= BUFFER_WRITE_INTERVAL:
            write_buffer_to_excel(buffer, thread_id, OUTPUT_DIR)
            buffer.clear()

    # Final write
    if buffer:
        write_buffer_to_excel(buffer, thread_id, OUTPUT_DIR)

    driver.quit()
    print(f"Thread-{thread_id} finished.")
def cross_join(df1, df2):
    df1['key'] = 1
    df2['key'] = 1
    result = pd.merge(df1, df2, on='key').drop('key', axis=1)
    return result
def main():
    

    combined_df = cross_join(user_group_data, item_policy_data)
    combined_df = combined_df.sort_values(by=["Primary Identifier", "Location Name", "Item Policy"])

    threads = []
    
    for i in range(N):
        t = threading.Thread(target=worker_thread, args=(i, combined_df))
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    print("All threads complete.")

if __name__ == "__main__":
    main()
