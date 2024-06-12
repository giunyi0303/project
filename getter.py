import gspread
import numpy as np
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
from dotenv import load_dotenv
import os
from gspread_formatting import *

load_dotenv()
project_info = {
    'ADH_project': {'username': os.getenv('ADH_id'), 'password': os.getenv('ADH_pw')},
    'SPR_project': {'username': os.getenv('SPR_id'), 'password': os.getenv('SPR_pw')},
    'WKH_project': {'username': os.getenv('WKH_id'), 'password': os.getenv('WKH_pw')},
}

def dynamic_web_crawler(url, info):
    options = Options()
    options.add_argument('--headless') 

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        driver.get(url)
        wait = WebDriverWait(driver, 10)
        username_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="id"]')))
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
        
        username_field.send_keys(info['username'])
        password_field.send_keys(info['password'])

        login_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[1]/div/div[3]/div[1]/form/button')))
        login_button.click()
        
        adherence = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/div/div[2]/div[2]/div[1]/div[1]/div/p')))
        value = adherence.text
        return value

    finally:
        driver.quit()
        
        
def dynamic_web_crawler_SPR(url, info):
    options = Options()
    options.add_argument('--headless') 

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        driver.get(url)
        wait = WebDriverWait(driver, 20)
        username_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="id"]')))
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
        
        username_field.send_keys(info['username'])
        password_field.send_keys(info['password'])

        login_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[1]/div/div[3]/div[1]/form/button')))
        login_button.click()
        
        list_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="site-select"]')))
        
        def click_element(element):
            try:
                element.click()
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", element)
        
        click_element(list_button)
        
        list_menu = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="menu-"]/div[3]')))
        agency_list = list_menu.text.split('\n')

        result = []

        for i, agency in enumerate(agency_list):
            try:
                agency_element = wait.until(EC.presence_of_element_located((By.XPATH, f'//li[text()="{agency}"]')))
                click_element(agency_element)
                
                short_wait = WebDriverWait(driver, 5)
                try:
                    info_element = short_wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/div/div[2]/div[2]/div[1]/div[1]/div/p')))
                    info_text = info_element.text
                    value = info_text
                except TimeoutException:
                    info_text = "할당없음"
                    value = info_text
                    driver.back()
                    driver.refresh()
                    list_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="site-select"]')))
                    click_element(list_button)

                print(f"Agency: {agency}")
                print(f"Info extracted: {info_text}")

                result.append((agency, value))
                
                if i < len(agency_list) - 1:
                    list_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="site-select"]')))
                    click_element(list_button)
            except TimeoutException:
                print(f"Agency element not found for: {agency}")
                continue
        
        return result

    finally:
        driver.quit()


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
json_keyfile_name = os.getenv('json_keyfil_name')
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, scope)
gc = gspread.authorize(credentials)
sheet_url = os.getenv('json_key')
doc = gc.open_by_url(sheet_url)
url = os.getenv('url')

def format_percentage(worksheet, cell):
    format_cell_range(worksheet, cell, CellFormat(
        numberFormat=NumberFormat(type='PERCENT', pattern='0.00%')
    ))

for project, info in project_info.items():
    print(f"Logging in for project: {project}")
    if project == 'SPR_project':
        try:
            results = dynamic_web_crawler_SPR(url , info)
            if results is not None:
                for agency, value in results:
                    print(f"Agency: {agency}, Value: {value}")
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    worksheet = doc.worksheet(project)
                    
                    # 퍼센트 형식으로 변환
                    try:
                        percentage_value = float(value.strip('%')) / 100
                    except ValueError:
                        percentage_value = value  # 변환에 실패하면 원래 값을 유지
                    
                    row_data = [project, current_time, agency, percentage_value]
                    worksheet.append_row(row_data)
                    
                    # 형식 지정
                    last_row = len(worksheet.get_all_values())
                    cell = f'D{last_row}'
                    format_percentage(worksheet, cell)
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            print("Retrying...")
            results = dynamic_web_crawler_SPR(url , info)
            if results is not None:
                for agency, value in results:
                    print(f"Agency: {agency}, Value: {value}")
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    worksheet = doc.worksheet(project)
                    
                    try:
                        percentage_value = float(value.strip('%')) / 100
                    except ValueError:
                        percentage_value = value
                    
                    row_data = [project, current_time, agency, percentage_value]
                    worksheet.append_row(row_data)
                    
                    last_row = len(worksheet.get_all_values())
                    cell = f'D{last_row}'
                    format_percentage(worksheet, cell)
            else:
                print("No results obtained.")
    else:  
        try:
            value = dynamic_web_crawler(url, info)
            print(f"Value extracted: {value}")
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            worksheet = doc.worksheet(project)
            
            try:
                percentage_value = float(value.strip('%')) / 100
            except ValueError:
                percentage_value = value
            
            row_data = [project, current_time, percentage_value]
            worksheet.append_row(row_data)
            
            last_row = len(worksheet.get_all_values())
            cell = f'C{last_row}'
            format_percentage(worksheet, cell)
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            print("Retrying...")
            value = dynamic_web_crawler(url, info)
            print(f"Value extracted: {value}")
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            worksheet = doc.worksheet(project)
            
            try:
                percentage_value = float(value.strip('%')) / 100
            except ValueError:
                percentage_value = value
            
            row_data = [project, current_time, percentage_value]
            worksheet.append_row(row_data)
            
            last_row = len(worksheet.get_all_values())
            cell = f'C{last_row}'
            format_percentage(worksheet, cell)
