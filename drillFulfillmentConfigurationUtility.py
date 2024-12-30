import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as ec
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

import secrets_local
import sys
import os
import glob
import requests
sys.path.append('scripts/')
from functions import *
# Define the directory where workbooks are located
processing_folder = './Processing/'

# Find the first Excel workbook in the "Processing" folder
workbooks = glob.glob(os.path.join(processing_folder, '*.xlsx'))

if not workbooks:
    print("No Excel workbooks found in the 'Processing' folder.")
    exit()

# Select the first workbook found
input_file = workbooks[0]
print(f"Processing file: {input_file}")
data = pd.read_excel(input_file, dtype="str")

# Proper WebDriver setup
options = webdriver.ChromeOptions()  # Create ChromeOptions object if needed
service = Service(ChromeDriverManager().install())  # Use ChromeDriverManager for installation

# Initialize the WebDriver
driver = webdriver.Chrome(service=service, options=options)
driver.get(secrets_local.alma_base_url_sandbox)

# Login to Alma
username = secrets_local.username
password = secrets_local.password
login(driver, username, password)

time.sleep(20)
try:
    # Wait for the modal to be present
    modal = WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.CLASS_NAME, "modal"))
    )
    print("Modal detected. Attempting to close it.")

    # Locate and click the close button for the modal
    close_button = modal.find_element(By.CLASS_NAME, "close")  # Adjust selector as needed
    close_button.click()

    # Wait for the modal to disappear
    WebDriverWait(driver, 10).until(
        ec.invisibility_of_element((By.CLASS_NAME, "modal"))
    )
    print("Modal closed successfully.")
except Exception as e:
    print(f"Modal handling failed or modal not present: {e}")



# Navigate to the Fulfillment Checkout page
driver.get("https://tufts-psb.alma.exlibrisgroup.com/ng/page;u=%2Fful%2Faction%2FpageAction.do%3FxmlFileName%3Dtou.fulfillment_configuration_utility.xml&pageViewMode%3DEdit&operation%3DLOAD&backUrl%3D%2Fful%2Faction%2Fmenu.do%3F&pageBean.selectedTab%3DtouType.loan&pageBean.touType%3DLoan&pageBean.displayDueDate%3Dtrue&pageBean.displayReturnDate%3Dtrue&pageBean.currentUrl%3DxmlFileName%253Dtou.fulfillment_configuration_utility.xml%2526pageViewMode%253DEdit%2526operation%253DLOAD%2526backUrl%253D%252Fful%252Faction%252Fmenu.do%253F%2526pageBean.selectedTab%253DtouType.loan%2526pageBean.touType%253DLoan%2526pageBean.displayDueDate%253Dtrue%2526pageBean.displayReturnDate%253Dtrue%2526resetPaginationContext%253Dtrue%2526showBackButton%253Dfalse&pageBean.navigationBackUrl%3D..%252Faction%252Fhome.do&resetPaginationContext%3Dtrue&showBackButton%3Dfalse&menuKey%3Dcom.exlibris.dps.adm.general.menu.initial.Fulfillment.FulfillmentHeader.FulConfigurationUtility")

time.sleep(15)  # Ensure page loads fully
results = []

# Wait for the GDPR banner to appear and locate the dismiss button


# Iterate through the input rows
for _, row in data.iterrows():
    user_id = row['user_primary_identifier'].strip()
    barcode = row['item_barcode'].strip()

    print(secrets_local.alma_sandbox_user_url + "/" + str(user_id) + "?" + secrets_local.alma_sandbox_user_apikey + "&format=json")

    user_record = requests.get(secrets_local.alma_sandbox_user_url + "/" + str(user_id) + "?apikey=" + secrets_local.alma_sandbox_user_apikey + "&format=json").json()

    print(user_record)
    item_record = requests.get(secrets_local.alma_sandbox_item_url + barcode + "&apikey=" + secrets_local.alma_sandbox_bib_apikey + "&format=json").json()

    print(item_record)
    first_name = user_record['first_name']
    last_name = user_record['last_name']

    user_group = user_record['user_group']['desc']

    search_name = last_name + ",\u00A0" + first_name  + "\u00A0\u00A0-\u00A0" + user_group + "\u00A0\u00A0-\u00A0" + user_id

    print(search_name)


    search_item = barcode + "-" + item_record['bib_data']['title'] + "-" +  item_record['item_data']['library']['desc'] + "-" + item_record['item_data']['location']['desc']
    # Enter User ID
    # time.sleep(2)
    # # Click the first element
    # driver.find_element(By.ID, "PICKUP_ID_pageBeandisplayNameOfUserOrUserIdendifier").click()

    # Close the modal if it exists
    # try:
    #     # Wait for the modal to be present
    #     modal = WebDriverWait(driver, 10).until(
    #         ec.visibility_of_element_located((By.CLASS_NAME, "modal"))
    #     )
    #     print("Modal detected. Attempting to close it.")
    #
    #     # Check for a close button inside the modal
    #     try:
    #         close_button = modal.find_element(By.CLASS_NAME, "close")  # Adjust selector if necessary
    #         close_button.click()
    #         print("Modal close button clicked.")
    #     except Exception:
    #         print("Close button not found in modal. Attempting alternative dismissal.")
    #         # If no close button, attempt alternative methods
    #         driver.execute_script("arguments[0].style.display = 'none';", modal)
    #         print("Modal hidden via JavaScript.")
    #
    #     # Wait for the modal to disappear
    #     WebDriverWait(driver, 10).until(
    #         ec.invisibility_of_element((By.CLASS_NAME, "modal"))
    #     )
    #     print("Modal closed successfully.")
    # except Exception as e:
    #     print(f"Modal handling failed or modal not present: {e}")
    #
    # try:
    #     close_button = modal.find_element(By.CLASS_NAME, "close")  # Adjust selector if necessary
    #     close_button.click()
    #     print("Modal close button clicked.")
    # except Exception as e:
    #     print(f"Modal handling failed or modal not present: {e}")



    # Click the search button
    # try:
    #     search_button = driver.find_element(By.ID, "simpleSearchIndexButton")
    #     search_button.click()
    #     print("Search button clicked successfully.")
    # except Exception as e:
    #     print(f"Failed to click search button: {e}")
    # # Click the second element

    time.sleep(2)
    # Click the first element
    driver.find_element(By.ID, "PICKUP_ID_pageBeandisplayNameOfUserOrUserIdendifier").click()

    time.sleep(7)

    # Wait for the modal to appear
    modal = WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.CLASS_NAME, "modal"))
    )
    print("Modal popup detected.")

    # Interact with elements in the modal
    # Example: Switch to iframe inside the modal
    driver.switch_to.frame(driver.find_element(By.ID, "iframePopupIframe"))
    print("Switched to iframe inside the modal.")

    # try:
    # # Set the button's displayed value
    #     driver.execute_script("""
    #         const button = document.getElementById('simpleSearchIndexButton');
    #         const span = button.querySelector('#simpleSearchIndexDisplay');
    #         if (span) {
    #             span.textContent = 'Primary identifier'; // Set the desired value
    #             button.setAttribute('title', 'Primary identifier'); // Update the title attribute
    #         }
    #     """)
    #     print("Button value set to 'Primary identifier'.")
    # except Exception as e:
    #     print(f"Failed to set button value: {e}")
    # try:
    #     # Wait for the search button to become clickable
    #     search_button = WebDriverWait(driver, 10).until(
    #         ec.element_to_be_clickable((By.ID, "simpleSearchIndexButton"))
    #     )
    #     search_button.click()
    #     print("Search button clicked successfully.")
    # except Exception as e:
    #     print(f"Failed to click search button: {e}")

    time.sleep(2)
    # Locate the <a> tag within the <li> element
    driver.find_element(By.ID, "simpleSearchIndexButton").click()

    time.sleep(2)
    primary_identifier_link = driver.find_element(By.XPATH, "//li[@id='TOP_NAV_Search_index_HFrUser.user_name']//a[text()='Primary identifier']")
    #
    # # Click the link
    primary_identifier_link.click()

    input = driver.find_element(By.ID, "ALMA_MENU_TOP_NAV_Search_Text")

    input.send_keys(user_id)

    time.sleep(2)

    driver.find_element(By.ID, "simpleSearchBtn").click()

    time.sleep(2)

    driver.find_element(By.XPATH, "//table[@id='TABLE_DATA_userList']/tbody/tr").click()



    time.sleep(2)







    # user_field.send_keys(search_name)
    # time.sleep(2)




    # Wait for the dropdown to be visible
    # dropdown_item = WebDriverWait(driver, 10).until(
    #     ec.visibility_of_element_located((By.XPATH, "//ul[@id='pageBeandisplayNameOfUserOrUserIdendifier_list']/li/a"))
    # )

    # time.sleep(15)
    # dropdown_item = driver.find_element(By.XPATH, "//ul[@id='pageBeandisplayNameOfUserOrUserIdendifier_list']/li/a")
    # time.sleep(2)
    # dropdown_item.click()
    #
    # dropdown_item.time.sleep(5)
    # Enter Barcode


    time.sleep(8)

    print(driver.page_source)
    # driver.switch_to.frame(driver.find_element(By.ID, "body_id_xml_file_tou.fulfillment_configuration_utility.xml"))
    # print("Switched back to main page")

    time.sleep(2)
    item_field = driver.find_element(By.XPATH, "//input[@id='pageBeanbarcode']")

    item_field.send_keys(barcode)


    ok_button = driver.find_element(By.ID, 'cbuttonok')

    ok_button.click()
    # Proceed to click the "OK" button
    # try:
    #     ok_button = WebDriverWait(driver, 10).until(
    #         ec.element_to_be_clickable((By.ID, "cbuttonok"))
    #     )
    #     ok_button.click()
    #     print("OK button clicked successfully.")
    # except Exception as e:
    #     print(f"Failed to click OK button: {e}")

    time.sleep(10)


    fulfillment_rule = driver.find_element(By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Fulfillment Unit Rule')]]//a").text
    tou_data = driver.find_element(By.XPATH, "//div[@class='row ' and .//span[contains(text(), 'Terms Of Use Name')]]//a").text

    # Extract the policy table for Loan
    loan_policy_table_html = driver.find_element(By.ID, 'TABLE_DATA_policiesList').get_attribute('outerHTML')
    loan_policy_df = pd.read_html(loan_policy_table_html)[0]

    # Switch to "Request" tab
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

    # Save data for this transaction
    results.append({
        "User ID": user_id,
        "Barcode": barcode,
        "Fulfillment Rule (Loan)": fulfillment_rule,
        "TOU (Loan)": tou_data,
        "Loan Policies": loan_policy_df.to_dict(orient='records'),  # Save as JSON-like
        "Fulfillment Unit Name (Request)": fulfillment_unit_name,
        "Fulfillment Rule (Request)": fulfillment_rule_request,
        "TOU (Request)": tou_request_data,
        "Request Policies": request_policy_df.to_dict(orient='records'),  # Save as JSON-like
    })

# Save results to a DataFrame
results_df = pd.DataFrame(results)
results_df.to_excel('Bulk_Checkout_Request_Results.xlsx', index=False)

print("Checkout and Request process completed. Results saved.")
driver.quit()
