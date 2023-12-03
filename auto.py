'''
Steeleon e-sales data upload Automation(RPA)

programmed by Seongju Hong,
Hot Rolling Department, Pohang Works, POSCO

2023-11-10 : E-Sales, Excel
2023-11-13 : Excel, MES Added

'''


from selenium import webdriver
from selenium.webdriver.chrome.service import Service

from webdriver_manager.chrome import ChromeDriverManager
import time

from selenium.webdriver.common.by import By
import pyautogui
import pyperclip
import win32com.client

import pandas as pd
import datetime
import os

import schedule

def Auto_ChromeDriver():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

def auto_task():
    # execute Chrome WebDriver
    driver = webdriver.Chrome()

    # 홈페이지 열기
    driver.get('http://steel-n.com/')

    time.sleep(2)

    tabs = driver.window_handles
    # print(tabs)

    while len(tabs) != 1:
            driver.switch_to.window(tabs[1])
            driver.close()
            tabs = driver.window_handles

    driver.switch_to.window(tabs[0])

    time.sleep(1)

    esales_btn = driver.find_element(By.XPATH, '//*[@id="mainContainer"]/div[2]/ul/li[1]/a/img')
    esales_btn.click()

    time.sleep(1)

    kor_btn = driver.find_element(By.XPATH, '//*[@id="mainContainer"]/div[2]/ul/li[1]/div/div/ul/li[1]/span[2]/a')
    kor_btn.click()

    time.sleep(2)

    tabs_2 = driver.window_handles
    print(tabs_2)

    while len(tabs_2) != 1:
            driver.switch_to.window(tabs_2[1])
            driver.close()
            tabs_2 = driver.window_handles

    driver.switch_to.window(tabs_2[0])

    time.sleep(2)

    # Customer Code, ID, PW, CRT_PW
    customer_code = 'GDA81'
    user_id = 'CNC1'
    user_pw = 'Poscocnc2'
    sec_pw = 'Purchase@'

    # Input Log-In Data
    customer_code_input = driver.find_element(By.ID, 'customerNumberP')
    customer_code_input.click()
    customer_code_input.send_keys(customer_code)

    user_id_input = driver.find_element(By.ID, 'userIdP')
    user_id_input.click()
    user_id_input.send_keys(user_id)

    user_pw_input = driver.find_element(By.ID, 'passwordP')
    user_pw_input.click()
    user_pw_input.send_keys(user_pw)
    pyautogui.press('enter')

    time.sleep(4)

    # Get Cert

    tabs_3 = driver.window_handles
    driver.switch_to.window(tabs_3[1])

    cert_manage_btn = driver.find_element(By.ID, 'us-cert-manage-btn')
    cert_manage_btn.click()

    cert_get_btn = driver.find_element(By.ID, 'us-cert-manage-get-cert-btn')
    cert_get_btn.click()

    cert_import_btn = driver.find_element(By.ID, 'us-btn-import')
    cert_import_btn.click()

    time.sleep(1)

    pyautogui.write(r"C:\Users\pyram\OneDrive\Desktop")
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.write(r"cert.p12")
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.write(sec_pw)
    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(1)

    select_btn = driver.find_element(By.ID, 'us-storage-select-confirm-btn')
    select_btn.click()
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    close_btn = driver.find_element(By.ID, 'us-cert-manage-cls-btn')
    close_btn.click()
    time.sleep(1)

    pw_btn = driver.find_element(By.ID, 'us-pw-text')
    pw_btn.click()
    time.sleep(1)
    pw_btn.send_keys(sec_pw)
    pyautogui.press('enter')

    time.sleep(5)

    # E-sales screen

    # main screen shift
    driver.switch_to.window(tabs_3[0])
    ship_btn = driver.find_element(By.XPATH, '//*[@id="gnb"]/ul/li[3]/a')
    ship_btn.click()

    time.sleep(2)

    macro_btn = driver.find_element(By.XPATH, '//*[@id="gnb"]/ul/li[3]/div/ul/li[3]/ul/li[9]/a')
    macro_btn.click()

    time.sleep(1)
    # 506, 643 pos move 
    # pyautogui.moveTo(506, 643)
    pyautogui.moveTo(420, 751)
    pyautogui.click()
    time.sleep(1)

    # search_btn = driver.find_element(By.XPATH, '//*[@id="btnSearch"]/span')
    # search_btn.click()
    pyautogui.moveTo(497, 295)
    pyautogui.click()
    time.sleep(3)

    # download_btn = driver.find_element(By.XPATH, '//*[@id="btnExcel1"]/span')
    # download_btn.click()
    pyautogui.moveTo(617, 295)
    pyautogui.click()
    time.sleep(3)

    # Close Chrome
    driver.close()

    # Excel text Transformation

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(r"C:\Users\user\Downloads\c800600310xls01.xls")

    wb.SaveAs(r"C:\Users\user\Downloads\change", FileFormat = 51)
    excel.Quit()

    df = pd.read_excel(r"C:\Users\user\Downloads\change.xlsx")

    # Check the file capacity
    if len(df.index) != 0:
        # check column
        if len(df.columns) > 24:
            df.drop(labels=['냉연코일최대폭','냉연코일최소폭'], axis=1, inplace=True)
        # Filtering FFB
        df = df[df['주문품명코드'] == 'FFB']

        # datetime 분리
        df['제품종합판정일시'] = df['제품종합판정일시'].dt.strftime('%H:%M:%S')

        # Save File as Text File
        desktop_path = 'C:\Users\user\Desktop'
        filename = datetime.datetime.now().strftime("%Y%m%d")
        df.to_csv(desktop_path+filename+'.txt', sep='\t', index = False, header=False)

        # Desktop Screen
        pyautogui.keyDown('win')
        pyautogui.press('d')
        pyautogui.keyUp('win')
        
        time.sleep(3)
        
        # MES App Open
        pyautogui.moveTo(110, 519)
        pyautogui.doubleClick()

        time.sleep(5)

        # Login
        pyperclip.copy("PC107522")
        pyautogui.hotkey('ctrl','v')
        pyautogui.press('enter')

        time.sleep(2)

        # Bookmark
        pyautogui.moveTo(14, 326)
        pyautogui.click()

        time.sleep(1)

        # Bookmark
        pyautogui.moveTo(14, 326)
        pyautogui.click()

        time.sleep(1)

        # Quality
        pyautogui.moveTo(14, 326)
        pyautogui.click()

        time.sleep(1)

        # Information
        pyautogui.moveTo(14, 326)
        pyautogui.click()

        time.sleep(1)

        # Upload
        pyautogui.moveTo(14, 326)
        pyautogui.click()

        time.sleep(1)

        # Password
        pyautogui.moveTo(14, 326)
        pyautogui.click()

        time.sleep(1)

        # Upload
        pyautogui.moveTo(14, 326)
        pyautogui.click()

        time.sleep(1)

        # Input Filename
        filename = datetime.datetime.now().strftime("%Y%m%d")
        pyperclip.copy(filename)
        pyautogui.hotkey('ctrl','v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)

        # Save
        pyautogui.moveTo(1038, 229)
        pyautogui.click()
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')

        # Close
        pyautogui.moveTo(1038, 229)
        pyautogui.click()
        time.sleep(2)

        # MES Close
        pyautogui.moveTo(1038, 229)
        pyautogui.click()
        time.sleep(2)

    # Remove ex file
    os.remove(r"C:\Users\user\Downloads\change.xlsx")
    os.remove(r"C:\Users\user\Downloads\c800600310xls01.xls")


# Start Time every day
schedule.every().day.at("13:00").do(auto_task)

while True:
    schedule.run_pending()
    time.sleep(1)








    # 업로드하기


        