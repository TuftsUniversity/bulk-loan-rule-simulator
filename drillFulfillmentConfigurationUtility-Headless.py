import pandas as pd
import os
import glob
import requests
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
from functions import login, safe_find_element, append_to_excel

# --- Constants & Configurations ---
INPUT_DIR_ITEM_POLICY = './input/Item Policies and Locations'
INPUT_DIR_USER_GROUP = './input/User Groups'
OUTPUT_FILE = 'Bulk_Checkout_Request_Results.xlsx'
BATCH_SIZE = 10  # Buffer size for writing to Excel

# --- Load Input Data ---
def load_first_excel(directory):
    files = glob.glob(os.path.join(directory, '*.xlsx'))
    if not files:
        print(f"No Excel files found in {directory}. Exiting.")
        exit()
    return pd.read_excel(files[0], dtype="str", engine="openpyxl")

item_policy_data = load_first_excel(INPUT_DIR_ITEM_POLICY)
user_group_data = load_first_excel(INPUT_DIR_USER_GROUP)

# --- Setup Selenium WebDriver ---
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get(secrets_local.alma_base_url)

# --- Login ---
login(driver, secrets_local.username, secrets_local.password)

# --- Close GDPR Modal If Present ---
try:
    modal = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='onetrust-close-btn-container']//button")))
    modal.click()
    print("GDPR modal closed.")
except TimeoutException:
    print("No GDPR modal detected.")

# --- Process Users ---
session = requests.Session()  # Reuse session for efficiency
buffer = []

for _, user_row in user_group_data.iterrows():
    user_id = user_row['Primary Identifier'].strip()
    user_group = user_row['User Group'].strip()

    print(f"Processing user: {user_id}")

    # Fetch user details
    user_record = session.get(f"{secrets_local.alma_sandbox_user_url}/{user_id}?apikey={secrets_local.alma_sandbox_user_apikey}&format=json").json()

    # Locate and open user menu
    user_menu = safe_find_element(driver, By.ID, "PICKUP_ID_pageBeandisplayNameOfUserOrUserIdendifier")
    user_menu.click()

    # Select user
    modal = safe_find_element(driver, By.CLASS_NAME, "modal")
    driver.switch_to.frame(driver.find_element(By.ID, "iframePopupIframe"))

    search_button = safe_find_element(driver, By.ID, "simpleSearchIndexButton")
    search_button.click()

    primary_identifier_link = safe_find_element(driver, By.XPATH, "//li[@id='TOP_NAV_Search_index_HFrUser.user_name']//a[text()='Primary identifier']")
    primary_identifier_link.click()

    input_field = safe_find_element(driver, By.ID, "ALMA_MENU_TOP_NAV_Search_Text")
    input_field.send_keys(user_id)

    search_button = safe_find_element(driver, By.ID, "simpleSearchBtn")
    search_button.click()

    row = safe_find_element(driver, By.XPATH, "//table[@id='TABLE_DATA_userList']/tbody/tr")
    row.click()

    # --- Process Items for User ---
    for _, item_row in item_policy_data.iterrows():
        barcode = item_row['Barcode']
        item_policy = item_row['Item Policy']
        location = item_row['Temporary Location Name'] if item_row['Temporary Physical Location In Use'] == "Yes" else item_row['Location Name']

        print(f"Processing item {barcode} - {item_policy} - {location}")

        # Enter barcode
        item_field = safe_find_element(driver, By.XPATH, "//input[@id='pageBeanbarcode']")
        item_field.clear()
        item_field.send_keys(barcode)

        # Click OK
        ok_button = safe_find_element(driver, By.ID, "cbuttonok")
        ok_button.click()

        # --- Extract Loan Policy Data ---
        loan_tab = safe_find_element(driver, By.ID, "A_NAV_LINK_touTypeloan_span")
        loan_tab.click()

        fulfillment_rule = safe_find_element(driver, By.XPATH, "//div[contains(@class, 'row ') and .//span[contains(text(), 'Fulfillment Unit Rule')]]//a").text
        tou_data = safe_find_element(driver, By.XPATH, "//div[contains(@class, 'row ') and .//span[contains(text(), 'Terms Of Use Name')]]//a").text

        # Store results
        buffer.append({
            "User ID": user_id,
            "User Group": user_group,
            "Barcode": barcode,
            "Item Policy": item_policy,
            "Location": location,
            "Fulfillment Rule (Loan)": fulfillment_rule,
            "TOU (Loan)": tou_data,
        })

        append_to_excel(OUTPUT_FILE, buffer, BATCH_SIZE)

# --- Cleanup ---
driver.quit()
print("Processing completed.")
