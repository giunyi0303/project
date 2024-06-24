import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import os

def find_element_with_retry(driver, by, value, timeout=10, retries=3):
    for _ in range(retries):
        try:
            element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
            return element
        except TimeoutException:
            print(f"Timeout: Element not found {value}. Retrying...")
            time.sleep(2)
    raise TimeoutException(f"Failed to find element {value} after {retries} retries")

options = Options()

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get('https://admin.inhandplus.com/')
wait = WebDriverWait(driver, 10)

#이거 바꾸세요
watch_numbers = [
    "1810", "1898", "2380", 
    "1731", "2806", "1593", "2337"
]

try:
    username_field = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/main/div[2]/main/form/div/div[2]/div/input')
    password_field = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/main/div[2]/main/form/div/div[3]/div/input')
    
    
    login_button = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/main/div[2]/main/form/div/button')
    login_button.click()
    
    tap_button = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/div/div/div[2]/button[4]')
    tap_button.click()
    
    time.sleep(2)
    
    state_button = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/div/div/div[2]/div/div/div/div/div/button[5]')
    state_button.click()
    
    for watch_number in watch_numbers:
        watch_field = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/main/div[2]/div/form/div[2]/div/input')
        watch_field.send_keys(watch_number)
        
        selected_watch = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/main/div[2]/div/div/div[2]/div/div[2]/div[2]/div/div/div/div[2]/button')
        selected_watch.click()
        
        find_element_with_retry(driver, By.XPATH, '/html/body/div/div/main/div[2]/div/div[1]/div[2]/div/div[6]/button').click()
        
        select_location = find_element_with_retry(driver, By.XPATH, '/html/body/div[2]/div[3]/div/div/form/div[1]/div/div/input')
        select_location.send_keys('B-23') #이거 바꾸세요
        
        # Use JavaScript to scroll the dropdown into view
        driver.execute_script("arguments[0].scrollIntoView(true);", select_location)
        time.sleep(2)
        
       
        dropdown_option = find_element_with_retry(driver, By.XPATH, "//li[contains(text(), 'B-23')]") #이거 바꾸세요
        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_option)
        time.sleep(2)
        dropdown_option.click()
        find_element_with_retry(driver, By.XPATH, '/html/body/div[2]/div[3]/div/div/form/button').click()
        state_button = find_element_with_retry(driver, By.XPATH, '/html/body/div/div/div/div/div[2]/div/div/div/div/div/button[5]')
        state_button.click()

except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
    print(f"Exception encountered: {e}")
finally:
    driver.quit()
