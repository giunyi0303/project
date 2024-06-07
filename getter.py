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

project_info = {
 # 아이디 비밀번호가 들어가야함 

}

def dynamic_web_crawler(url, info):
    # 크롬 옵션 창 열지 않기 
    options = Options()
    options.add_argument('--headless') 

    # instance 생성 
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        driver.get(url)

        # 열릴때까지 기다리기 
        wait = WebDriverWait(driver, 10)
        username_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="id"]')))
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
        
        # 로그인 값 입력하기 
        username_field.send_keys(info['username'])
        password_field.send_keys(info['password'])

        # 로그인 버튼 찾기 
        login_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[1]/div/div[3]/div[1]/form/button')))
        login_button.click()
        
        adherence = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/div/div[2]/div[2]/div[1]/div[1]/div/p')))
        value = adherence.text
        return value

    finally:
        driver.quit()

# google API 접속 
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
json_keyfile_name = 'test-project-424802-962f013a50f0.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, scope)
gc = gspread.authorize(credentials)
sheet_url = 'https://docs.google.com/spreadsheets/d/1Y5KmMPstYcEVlYH-FZ3q1P4X_yHGuvpCl5AwvJO0CX8/edit#gid=0'
doc = gc.open_by_url(sheet_url)
url = 'https://care.inhandplus.com/ko/login'
for project, info in project_info.items():
    print(f"Logging in for project: {project}")
    value = dynamic_web_crawler(url, info)
    print(f"Value extracted: {value}")
    
    # 현재 시간 가져오기
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 현재 시트에 데이터 추가
    worksheet = doc.worksheet(project)
    row_data = [project, current_time, value]
    worksheet.append_row(row_data)
