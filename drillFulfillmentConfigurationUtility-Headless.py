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
import secrets_local
import numpy as np
import sys

sys.path.append("scripts/")
from functions import *
import time

# --- Constants & Configurations ---
INPUT_DIR_ITEM_POLICY = "./input/Item Policies and Locations"
INPUT_DIR_USER_GROUP = "./input/User Groups"
OUTPUT_FILE = "Bulk_Checkout_Request_Results.xlsx"
OUTPUT_DIR = "./Output"
BATCH_SIZE = 10  # Buffer size for writing to Excel


# --- Load Input Data ---
def load_first_excel(directory):
    files = glob.glob(os.path.join(directory, "*.xlsx"))
    if not files:
        print(f"No Excel files found in {directory}. Exiting.")
        exit()
    return pd.read_excel(files[0], dtype="str", engine="openpyxl")


item_policy_data = load_first_excel(INPUT_DIR_ITEM_POLICY)
user_group_data = load_first_excel(INPUT_DIR_USER_GROUP)

# --- Setup Selenium WebDriver ---
chrome_options = Options()
chrome_options.add_argument("--headless")
def process_user_group(thread_id, user_group_data, item_policy_data):
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    service = Service(ChromeDriverManager().install())
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.get(secrets_local.alma_base_url)

# --- Load Input Data ---
# --- Login ---
    login(driver, secrets_local.username, secrets_local.password)


    time.sleep(20)
    # --- Close GDPR Modal If Present ---
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

    # Locate and click the close button for the modal
    # close_button = modal.find_element(By.CLASS_NAME, "close")  # Adjust selector as needed

    # Wait for the modal to disappear
    # WebDriverWait(driver, 10).until(
    #    ec.invisibility_of_element((By.CLASS_NAME, "modal"))
    # except Exception as e:
    # print(f"Modal handling failed or modal not present: {e}")


    # Navigate to the Fulfillment Checkout page
    driver.get(
        "https://tufts.alma.exlibrisgroup.com/ng/page;u=%2Fful%2Faction%2FpageAction.do%3FxmlFileName%3Dtou.fulfillment_configuration_utility.xml&pageViewMode%3DEdit&operation%3DLOAD&backUrl%3D%2Fful%2Faction%2Fmenu.do%3F&pageBean.selectedTab%3DtouType.loan&pageBean.touType%3DLoan&pageBean.displayDueDate%3Dtrue&pageBean.displayReturnDate%3Dtrue&pageBean.currentUrl%3DxmlFileName%253Dtou.fulfillment_configuration_utility.xml%2526pageViewMode%253DEdit%2526operation%253DLOAD%2526backUrl%253D%252Fful%252Faction%252Fmenu.do%253F%2526pageBean.selectedTab%253DtouType.loan%2526pageBean.touType%253DLoan%2526pageBean.displayDueDate%253Dtrue%2526pageBean.displayReturnDate%253Dtrue%2526resetPaginationContext%253Dtrue%2526showBackButton%253Dfalse&pageBean.navigationBackUrl%3D..%252Faction%252Fhome.do&resetPaginationContext%3Dtrue&showBackButton%3Dfalse&menuKey%3Dcom.exlibris.dps.adm.general.menu.initial.Fulfillment.FulfillmentHeader.FulConfigurationUtility"
    )
    # --- Process Users ---
    session = requests.Session()  # Reuse session for efficiency
    buffer = []

    for _, user_row in user_group_data.iterrows():
        user_id = user_row["Primary Identifier"].strip()
        user_group = user_row["User Group"].strip()

        print(f"Processing user: {user_id}")

        # Fetch user details
        user_record = session.get(
            f"{secrets_local.alma_sandbox_user_url}/{user_id}?apikey={secrets_local.alma_sandbox_user_apikey}&format=json"
        ).json()

        # Locate and open user menu
        user_menu = safe_find_element(
            driver, By.ID, "PICKUP_ID_pageBeandisplayNameOfUserOrUserIdendifier"
        )
        user_menu.click()

        # Select user
        modal = safe_find_element(driver, By.CLASS_NAME, "modal")
        driver.switch_to.frame(driver.find_element(By.ID, "iframePopupIframe"))

        search_button = safe_find_element(driver, By.ID, "simpleSearchIndexButton")
        search_button.click()

        time.sleep(2)
        primary_identifier_link = safe_find_element(
            driver,
            By.XPATH,
            "//li[@id='TOP_NAV_Search_index_HFrUser.user_name']//a[text()='Primary identifier']",
        )
        primary_identifier_link.click()

        input_field = safe_find_element(driver, By.ID, "ALMA_MENU_TOP_NAV_Search_Text")
        input_field.send_keys(user_id)

        search_button = safe_find_element(driver, By.ID, "simpleSearchBtn")
        search_button.click()

        row = safe_find_element(
            driver, By.XPATH, "//table[@id='TABLE_DATA_userList']/tbody/tr"
        )
        row.click()
        time.sleep(4)
        print(driver.page_source)
        print("Switched back to main page")
        # --- Process Items for User ---
        for _, item_row in item_policy_data.iterrows():
            barcode = item_row["Barcode"]
            item_policy = item_row["Item Policy"]
            location = (
                item_row["Temporary Location Name"]
                if item_row["Temporary Physical Location In Use"] == "Yes"
                else item_row["Location Name"]
            )

            print(f"Processing item {barcode} - {item_policy} - {location}")

            # Enter barcode (Refind element before interaction)
            item_field = safe_find_element(
                driver, By.XPATH, "//input[@id='pageBeanbarcode']"
            )

            send_keys_with_retry(
                driver, By.XPATH, "//input[@id='pageBeanbarcode']", barcode
            )

            # Click OK button (Refind each time)
            ok_button = safe_find_element(driver, By.ID, "cbuttonok")
            click_element_with_retry(driver, By.ID, "cbuttonok")

            # Extract policy table
            loan_policy_list = get_table_html_with_retry(
                driver, By.ID, "TABLE_DATA_policiesList", "loan"
            )

            loan_fulfillment_rule_name = loan_policy_list[0]
            loan_tou_name = loan_policy_list[1]
            loan_dict = loan_policy_list[2]
            fulfillment_unit_name = loan_policy_list[3]


            request_policy_list = get_table_html_with_retry(driver, By.ID, "TABLE_DATA_policiesList", "request")
            # --- Extract Request Tab Data ---

            request_fulfillment_rule_name = request_policy_list[0]
            request_tou_name = request_policy_list[1]
            request_dict = request_policy_list[2]
            # request_tab = safe_find_element(driver, By.ID, "A_NAV_LINK_touTyperequest_span")
            # click_element_with_retry(driver, By.ID, "A_NAV_LINK_touTyperequest_span")

            # fulfillment_rule_request = safe_find_element_text(
            #     driver,
            #     By.XPATH,
            #     "//div[contains(@class, 'row ') and .//span[contains(text(), 'Fulfillment Unit Rule')]]//a",
            # )

            # fulfillment_unit_name = safe_find_element_text(
            #     driver,
            #     By.XPATH,
            #     "//div[contains(@class, 'row ') and .//span[contains(text(), 'Fulfillment Unit Name')]]//a",
            # )

            # tou_request_data = safe_find_element_text(
            #     driver,
            #     By.XPATH,
            #     "//div[contains(@class, 'row ') and .//span[contains(text(), 'Terms Of Use Name')]]//a",
            # )

            # # Extract Request Tab Data

            # # Extract the policy table for Request
            # request_policy_table_html = get_table_html_with_retry(
            #     driver, By.ID, "TABLE_DATA_policiesList"
            # )
            # request_policy_df = pd.read_html(request_policy_table_html)[0]
            # Store results

            

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

            output_path = os.path.join(OUTPUT_DIR, f"output_thread_{thread_id}.xlsx")
            append_to_excel(output_path, buffer)

            # append_to_excel(
            #     OUTPUT_FILE, buffer,
            # )  # Do not reassign buffer
            buffer.clear()

# --- Cleanup ---
    driver.quit()
    print("Processing completed.")
# --- Multithreading Setup ---
def main():
    N = 4  # Number of browsers to spawn
    thread_list = []
    user_group_chunks = np.array_split(user_group_data, N)  # Split data into N chunks


# --- Multithreading Setup ---
def main():
    N = 4  # Number of browsers to spawn
    thread_list = []
    user_group_chunks = np.array_split(user_group_data, N)  # Split data into N chunks

    
    # Start threads
    for i in range(N):
        t = threading.Thread(
            name=f"Thread-{i}",
            target=process_user_group,
            args=(i, user_group_chunks[i], item_policy_data),
        )
        t.start()
        print(f"{t.name} started!")
        thread_list.append(t)

    # Wait for all threads to complete
    for thread in thread_list:
        thread.join()

    # Merge Excel files
    merge_excel_files(N)
    print("Test completed!")
# --- Worker Thread ---
# def worker(queue, thread_id):
#     while not queue.empty():
#         user_group, item = queue.get()
#         try:
#             process(user_group, item, thread_id)
#         finally:
#             queue.task_done()

def merge_excel_files(num_threads):
    """Merge Excel files from all threads into a single file."""
    all_data = []
    
    for i in range(num_threads):
        file_path = f"{OUTPUT_DIR}/output_thread_{i}.xlsx"
        if os.path.exists(file_path):
            print(f"✅ Including: {file_path}")
            df = pd.read_excel(file_path, engine="openpyxl")
            all_data.append(df)
        else:
            print(f"⚠️ File not found: {file_path}")

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
        # Remove old file to ensure it's not partially overwritten
        if os.path.exists(OUTPUT_FILE):
            os.remove(OUTPUT_FILE)

        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"📁 Merged {len(all_data)} files. Final output: {OUTPUT_FILE} ({os.path.getsize(OUTPUT_FILE)/1024:.2f} KB)")
    else:
        print("❌ No files to merge.")

if __name__ == "__main__":
    main()


def merge_excel_files(num_threads):
    """Merge Excel files from all threads into a single file."""
    all_data = []
    for i in range(num_threads):
        file_path = f"{OUTPUT_DIR}/output_thread_{i}.xlsx"
        if os.path.exists(file_path):
            df = pd.read_excel(file_path, engine="openpyxl")
            all_data.append(df)

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        final_df.to_excel(OUTPUT_FILE, index=False)


# --- Multithreading Setup ---
# def main():
#     N = 4  # Number of browsers to spawn
#     thread_list = []
#     user_group_chunks = np.array_split(user_group_data, N)  # Split data into N chunks

#     # Start threads
#     for i in range(N):
#         t = threading.Thread(
#             name=f"Thread-{i}",
#             target=process_user_group,
#             args=(i, user_group_chunks[i], item_policy_data),
#         )
#         t.start()
#         print(f"{t.name} started!")
#         thread_list.append(t)

#     # Wait for all threads to complete
#     for thread in thread_list:
#         thread.join()

#     # Merge Excel files
#     merge_excel_files(N)
#     print("Test completed!")

def merge_excel_files(num_threads):
    """Merge Excel files from all threads into a single file."""
    all_data = []
    for i in range(num_threads):
        file_path = f"{OUTPUT_DIR}/output_thread_{i}.xlsx"
        if os.path.exists(file_path):
            df = pd.read_excel(file_path, engine="openpyxl")
            all_data.append(df)

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        final_df.to_excel(OUTPUT_FILE, index=False)

