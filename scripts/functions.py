from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import pandas as pd
import os
from openpyxl import load_workbook

def login(driver, username, password):
    """Login to Alma"""
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
    username_field.send_keys(username)

    password_field = driver.find_element(By.ID, 'password')
    password_field.send_keys(password)

    password_field.submit()
    return driver

def safe_find_element(driver, by, value, retries=3):
    """Find element with retries for StaleElementReferenceException"""
    for attempt in range(retries):
        try:
            return WebDriverWait(driver, 10).until(EC.visibility_of_element_located((by, value)))
        except StaleElementReferenceException:
            print(f"Retry {attempt + 1}: Element stale, retrying...")
    raise Exception(f"Failed to locate element: {value}")

def append_to_excel(file_path, buffer, batch_size=10):
    """Append buffer data to Excel file in batches"""
    if len(buffer) >= batch_size:
        buffer_df = pd.DataFrame(buffer)
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            buffer_df.to_excel(writer, index=False, header=False)
        buffer.clear()
