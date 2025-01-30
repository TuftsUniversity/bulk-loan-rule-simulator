import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as ec
from selenium import webdriver

from selenium.common.exceptions import StaleElementReferenceException

from webdriver_manager.chrome import ChromeDriverManager

import secrets_local
import sys
import os
import glob
import requests
from openpyxl import load_workbook
sys.path.append('scripts/')
from functions import *


# Define the directory where workbooks are located
item_policy_and_location_folder = './input/Item Policies and Locations'

user_group_folder = './input/User Groups'
# Find the first Excel workbook in the "Processing" folder
ip_l_workbooks = glob.glob(os.path.join(item_policy_and_location_folder, '*.xlsx'))

if not ip_l_workbooks:
    print("No Excel workbooks found in the 'Item Policy and Locations' folder.")
    exit()

# Select the first workbook found
item_policy_and_location_input_file = ip_l_workbooks[0]
print(f"Processing file: {item_policy_and_location_input_file}")
item_policy_and_location_data = pd.read_excel(item_policy_and_location_input_file, dtype="str", engine="openpyxl")

user_group_workbooks = glob.glob(os.path.join(user_group_folder, '*.xlsx'))

if not user_group_workbooks:
    print("No Excel workbooks found in the 'Item Policy and Locations' folder.")
    exit()

# Select the first workbook found
user_group_input_file = user_group_workbooks[0]
print(f"Processing file: {user_group_input_file}")
user_group_data = pd.read_excel(user_group_input_file, dtype="str", engine="openpyxl")
# Proper WebDriver setup
#options = webdriver.ChromeOptions()  # Create ChromeOptions object if needed

output_file = 'Bulk_Checkout_Request_Results.xlsx'
# Check if the output file exists
if os.path.exists(output_file):
    # Load existing results to avoid duplication
    existing_results = pd.read_excel(output_file, engine="openpyxl")
else:
    existing_results = pd.DataFrame()

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
service = Service(ChromeDriverManager().install())  # Use ChromeDriverManager for installation

# Initialize the WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get(secrets_local.alma_base_url_sandbox)

# Login to Alma
username = secrets_local.username
password = secrets_local.password
login(driver, username, password)

time.sleep(7)
#try:
# Wait for the modal to be present
#modal = WebDriverWait(driver, 10).until(
#    ec.visibility_of_element_located((By.XPATH, "//#div[@id='onetrust-constent-sdk']//button")
#))
page_source = str(driver.page_source)

#print("page source GDPR match")
#print(re.sub(r'(.{30}body_id_xml_file_.{30})', r'\1', page_source))

# file = open("test_file.html", "w+")
#
# file.write(str(driver.page_source.encode('utf-8')))
#
# file.close()
#
time.sleep(5)
#driver.switch_to.frame(driver.find_element(By.ID, "body_id_xml_file_"))

try:
    modal = driver.find_element(By.XPATH, "//div[@id='onetrust-close-btn-container']//button")
    print("GDPR modal detected. Attempting to close it.")
    modal.click()

except:

    print("No GDPR modal")

# Locate and click the close button for the modal
#close_button = modal.find_element(By.CLASS_NAME, "close")  # Adjust selector as needed

# Wait for the modal to disappear
#WebDriverWait(driver, 10).until(
#    ec.invisibility_of_element((By.CLASS_NAME, "modal"))
#except Exception as e:
    #print(f"Modal handling failed or modal not present: {e}")



# Navigate to the Fulfillment Checkout page
driver.get("https://tufts-psb.alma.exlibrisgroup.com/ng/page;u=%2Fful%2Faction%2FpageAction.do%3FxmlFileName%3Dtou.fulfillment_configuration_utility.xml&pageViewMode%3DEdit&operation%3DLOAD&backUrl%3D%2Fful%2Faction%2Fmenu.do%3F&pageBean.selectedTab%3DtouType.loan&pageBean.touType%3DLoan&pageBean.displayDueDate%3Dtrue&pageBean.displayReturnDate%3Dtrue&pageBean.currentUrl%3DxmlFileName%253Dtou.fulfillment_configuration_utility.xml%2526pageViewMode%253DEdit%2526operation%253DLOAD%2526backUrl%253D%252Fful%252Faction%252Fmenu.do%253F%2526pageBean.selectedTab%253DtouType.loan%2526pageBean.touType%253DLoan%2526pageBean.displayDueDate%253Dtrue%2526pageBean.displayReturnDate%253Dtrue%2526resetPaginationContext%253Dtrue%2526showBackButton%253Dfalse&pageBean.navigationBackUrl%3D..%252Faction%252Fhome.do&resetPaginationContext%3Dtrue&showBackButton%3Dfalse&menuKey%3Dcom.exlibris.dps.adm.general.menu.initial.Fulfillment.FulfillmentHeader.FulConfigurationUtility")

time.sleep(5)  # Ensure page loads fully
results = []

# Wait for the GDPR banner to appear and locate the dismiss button




# Output file path
output_file = 'Bulk_Checkout_Request_Results.xlsx'

# Check if the output file exists
if os.path.exists(output_file):
    # Load existing results to avoid duplication
    existing_results = pd.read_excel(output_file)
else:
    existing_results = pd.DataFrame()

# Buffer for incremental saving
buffer = []
buffer_size = 10  # Adjust this as needed to control frequency of saves

# Function to append data to Excel
def append_to_excel(file_path, data_frame):
    if os.path.exists(file_path):
        # Load the workbook and determine the last row
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            workbook = load_workbook(file_path)
            sheet = workbook.active
            startrow = sheet.max_row  # Get the last row with data
            data_frame.to_excel(writer, index=False, header=False, startrow=startrow)
    else:
        # Create a new file
        data_frame.to_excel(file_path, index=False)

# Iterate through the input rows
for _, row in user_group_data.iterrows():
    user_id = row['Primary Identifier'].strip()
    user_group = row['User Group'].strip()


    print(user_id)


    print(secrets_local.alma_sandbox_user_url + "/" + str(user_id) + "?" + secrets_local.alma_sandbox_user_apikey + "&format=json")

    user_record = requests.get(secrets_local.alma_sandbox_user_url + "/" + str(user_id) + "?apikey=" + secrets_local.alma_sandbox_user_apikey + "&format=json").json()

        # Click the first element
    user_menu = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.ID, "PICKUP_ID_pageBeandisplayNameOfUserOrUserIdendifier"))
    )

    user_menu.click()



    # Wait for the modal to appear
    modal = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.CLASS_NAME, "modal"))
    )
    print("Modal popup detected.")

    # Interact with elements in the modal
    # Example: Switch to iframe inside the modal
    driver.switch_to.frame(driver.find_element(By.ID, "iframePopupIframe"))
    print("Switched to iframe inside the modal.")

    search_index_button = WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.ID, "simpleSearchIndexButton"))
    )
    time.sleep(5)
    # Locate the <a> tag within the <li> element
    search_index_button.click()
    primary_identifier_link = driver.find_element(By.XPATH, "//li[@id='TOP_NAV_Search_index_HFrUser.user_name']//a[text()='Primary identifier']")
    #
    # # Click the link
    primary_identifier_link.click()

    input = driver.find_element(By.ID, "ALMA_MENU_TOP_NAV_Search_Text")

    input.send_keys(user_id)

    time.sleep(2)


    simple_search_button = WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.ID, "simpleSearchBtn"))
    )
    simple_search_button.click()

    time.sleep(2)

    row = WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, "//table[@id='TABLE_DATA_userList']/tbody/tr"))
    )

    row.click()



    time.sleep(2)


#print(user_record)


    print(driver.page_source)
    print("Switched back to main page")

    for _, row in item_policy_and_location_data.iterrows():
        barcode = row['Barcode']
        num_of_items = row['Num of Items (In Repository)']
        item_policy = row['Item Policy']
        if row['Temporary Physical Location In Use'] == "Yes":
            location = row['Temporary Location Name']
        else:
            location = row['Location Name']


        print(str(barcode) + "-" + item_policy + "-" + location )













        time.sleep(2)

        retry_count = 3  # Number of retries
        # for attempt in range(retry_count):
        #     try:
        item_field = driver.find_element(By.XPATH, "//input[@id='pageBeanbarcode']")

        item_field.clear()
        item_field.send_keys(barcode)




        ok_button = driver.find_element(By.ID, 'cbuttonok')

        ok_button.click()
            # except StaleElementReferenceException:
            #     print("Failed to find the item barcode button due to stale element reference.")
            #
        #     #     continue  # Skip to the next iteration if retries fail
        # else:
        #     print("Failed to find the item barcode button due to stale element reference")
        #     continue  # Skip to the next iteration if retries fail

        # Proceed to click the "OK" button
        # try:
        #     ok_button = WebDriverWait(driver, 10).until(
        #         ec.element_to_be_clickable((By.ID, "cbuttonok"))
        #     )
        #     ok_button.click()
        #     print("OK button clicked successfully.")
        # except Exception as e:
        #     print(f"Failed to click OK button: {e}")


        # fulfillment_rule_element = WebDriverWait(driver, 10).until(           ec.visibility_of_element_located((By.XPATH, "TABLE_DATA_policiesList")))

        time.sleep(2)

        # Switch to "Request" tab
        retry_count = 3  # Number of retries
        for attempt in range(retry_count):
            try:
                # Re-find the "Request" tab element each time
                loan_tab = driver.find_element(By.ID, 'A_NAV_LINK_touTypeloan_span')
                loan_tab.click()
                time.sleep(2)
                time.sleep(2)
                fulfillment_rule = driver.find_element(By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Fulfillment Unit Rule')]]//a").text


                tou_data = driver.find_element(By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Terms Of Use Name')]]//a").text

                # Extract the policy table for Loan
                loan_policy_table_html = driver.find_element(By.ID, 'TABLE_DATA_policiesList').get_attribute('outerHTML')
                loan_policy_df = pd.read_html(loan_policy_table_html)[0]





                # Exit loop if successful
                break
            except StaleElementReferenceException:
                print(f"Attempt {attempt + 1} of {retry_count} failed: Stale Element Reference. Retrying...")
                time.sleep(1)  # Small delay before retry
        else:
            print("Failed to process the Loan tab due to stale element reference.")
            continue  # Skip to the next iteration if retries fail






        # Switch to "Request" tab
        retry_count = 3  # Number of retries
        for attempt in range(retry_count):
            try:
                # Re-find the "Request" tab element each time
                request_tab = driver.find_element(By.ID, 'A_NAV_LINK_touTyperequest_span')
                request_tab.click()
                time.sleep(2)

                # Extract Request Tab Data
                fulfillment_unit_name = driver.find_element(
                    By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Fulfillment Unit Name')]]//a"
                ).text
                fulfillment_rule_request = driver.find_element(
                    By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Fulfillment Unit Rule')]]//a"
                ).text
                tou_request_data = driver.find_element(
                    By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Terms Of Use Name')]]//a"
                ).text

                # Extract the policy table for Request
                request_policy_table_html = driver.find_element(By.ID, 'TABLE_DATA_policiesList').get_attribute('outerHTML')
                request_policy_df = pd.read_html(request_policy_table_html)[0]

                # Exit loop if successful
                break
            except StaleElementReferenceException:
                print(f"Attempt {attempt + 1} of {retry_count} failed: Stale Element Reference. Retrying...")
                time.sleep(1)  # Small delay before retry
        else:
            print("Failed to process the Request tab due to stale element reference.")
            continue  # Skip to the next iteration if retries fail
        time.sleep(2)

        # # Extract Request Tab Data
        # fulfillment_unit_name = driver.find_element(
        #     By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Fulfillment Unit Name')]]//a"
        # ).text
        # fulfillment_rule_request = driver.find_element(
        #     By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Fulfillment Unit Rule')]]//a"
        # ).text
        # tou_request_data = driver.find_element(
        #     By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Terms Of Use Name')]]//a"
        # ).text
        #
        # # Extract the policy table for Request
        # request_policy_table_html = driver.find_element(By.ID, 'TABLE_DATA_policiesList').get_attribute('outerHTML')
        # request_policy_df = pd.read_html(request_policy_table_html)[0]

         # Append result to buffer
        buffer.append({
            "User ID": user_id,
            "User Group": user_group,
            "Barcode": barcode,
            "Item Policy": item_policy,
            "Number of Items": num_of_items,
            "Location": location,
            "Fulfillment Rule (Loan)": fulfillment_rule,
            "TOU (Loan)": tou_data,
            "Loan Policies": loan_policy_df.to_dict(orient='records'),
            "Fulfillment Unit Name (Request)": fulfillment_unit_name,
            "Fulfillment Rule (Request)": fulfillment_rule_request,
            "TOU (Request)": tou_request_data,
            "Request Policies": request_policy_df.to_dict(orient='records'),
        })

           # Save buffer to file if it reaches the specified size
        if len(buffer) >= buffer_size:
            buffer_df = pd.DataFrame(buffer)
            append_to_excel(output_file, buffer_df)
            print("Incremental results saved.")
            buffer.clear()


# Save any remaining data in the buffer
if buffer:
    buffer_df = pd.DataFrame(buffer)
    append_to_excel(output_file, buffer_df)
print("Processing completed. Incremental results saved.")
driver.quit()
