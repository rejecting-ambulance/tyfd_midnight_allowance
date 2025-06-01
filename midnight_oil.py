#網頁操作
from selenium import webdriver
from selenium.webdriver.support.ui import Select #選單
from selenium.webdriver.common.by import By #定位
from selenium.webdriver.support.ui import WebDriverWait #等待載入
from selenium.webdriver.support import expected_conditions as EC #等待載入
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
#excel操作
import openpyxl
#輔助
from datetime import datetime
import time
import sys
import os
import json

"""



"""


def add_zero(cc):   #1位補0
    if len(cc) < 2:
        cc = '0' + cc
    return cc


def alert_click(driver):    #警告處理
    try:
        time.sleep(0.2)
        alert = driver.switch_to.alert
        print(f'  {alert.text}')
        alert.accept()
    except:
        pass


def setup_chrome_driver():
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')  # 禁用 GPU
    options.add_argument('--disable-software-rasterizer')
    options.add_argument('--no-sandbox')
    options.add_argument('--log-level=3')  # 僅顯示致命錯誤，關掉 Info/Warning
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 關掉 DevTools 日誌
    
    driver = webdriver.Chrome(service=Service(), options=options)
    return driver


def get_exe_dir():  #取得 `.exe` 真正所在的目錄
    if getattr(sys, 'frozen', False):  # PyInstaller 打包後
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def load_accounts(config_path="config.json"):   #載入
    with open(config_path, "r", encoding="utf-8") as f: 
        return json.load(f)

def dropdown_by_value(id , value , driver, wait):   #一些瀏覽器操作區塊
    dropdown1 = wait.until(EC.visibility_of_element_located((By.ID, id)))   
    dropdown_sheet1 = Select(dropdown1)
    dropdown_sheet1.select_by_value(value)

def dropdown_by_index(id , value , driver, wait):
    dropdown1 = wait.until(EC.visibility_of_element_located((By.ID, id)))   
    dropdown_sheet1 = Select(dropdown1)
    dropdown_sheet1.select_by_index(value)

def dropdown_by_text(id , value , driver, wait):
    dropdown1 = wait.until(EC.visibility_of_element_located((By.ID, id)))   
    dropdown_sheet1 = Select(dropdown1)
    dropdown_sheet1.select_by_visible_text(value)

def click_by_id(id , driver, wait):
    button1 = driver.find_element(By.ID, id)
    button1.click()

def click_by_name(name , driver, wait):
    button1 = driver.find_element(By.NAME, name)
    button1.click()

def click_by_xpath(xpath , driver, wait):
    button1 = driver.find_element(By.XPATH, xpath)
    button1.click()

def str_line(show):     #大區段分隔線
    max_len = 50
    dash_len = int((max_len - len(show))/2)
    dash = ''
    for i in range(dash_len):
        dash = dash + '-'
    show = dash + show + dash + '\n'
    
    return show


def remove_duplicates(data):       #移除重複列
    seen = set()
    result = []
    for row in data:
        row_tuple = tuple(str(x) for x in row)
        if row_tuple not in seen:
            seen.add(row_tuple)
            result.append(row)
    return result

def insert_type(arr, new_value):
    if len(arr) < 9 or arr[8] in (None, ''):
        arr.append(new_value)
    else:
        # 拆解現有內容成 set 來比對，避免重複
        existing_values = set(arr[8].split(','))
        if new_value not in existing_values:
            arr[8] = f"{arr[8]},{new_value}"  # 逗號隔開
    return arr


def comapre_times(driver, wait, data, unit):      #爬蟲裡面的比對時間區段

    dropdown_by_value('_selYEAR',data['year'], driver, wait)
    dropdown_by_value('_selMONTH',data['month'], driver, wait)
    click_by_id('_btnQuery', driver, wait)    #點選查詢

    try:
        wait.until(EC.element_to_be_clickable((By.ID, '_btnQuery')))
        time.sleep(2)
    except TimeoutError as e:
        print(e)
        return 0

    # Collect all rows first
    table = driver.find_element(By.XPATH, '//*[@id="frm"]/table/tbody/tr[5]/td/table/tbody')
    rows = table.find_elements(By.TAG_NAME, 'tr')
    
    table_content = []   #爬到的內容
    wrong_array = []    #要被剃除的內容

    for row in rows:    
        
        cells = row.find_elements(By.TAG_NAME, 'td')
        if not cells:
            cells = row.find_elements(By.TAG_NAME, 'th')

        cell_values = [cell.text.strip() for cell in cells]
        cell_values.insert(0, unit)



        # Skip header or empty rows
        if '開始時間' in cell_values or len(cell_values) < 7:
            continue

        # check the background color of the third cell
        bg_color = row.value_of_css_property('background-color')
        if bg_color.startswith('rgba(255, 147, 147'):
            wrong_type = "案號相同"
            cell_values.append(wrong_type)
            wrong_array.append(cell_values)


        table_content.append(cell_values)

    if len(table_content) == 0:
        print('🚑查無資料')
         
    # Compare each row to the next row
    new_person = ''
    for i in range(len(table_content) - 1):
        current = table_content[i]
        next_row = table_content[i + 1]

        person_current = current[2]
        person_next = next_row[2]
        
        if new_person != person_current:
            new_person = person_current
            change = 1
            print('')
        else:
            change = 0

        # Only compare if same person
        if person_current != person_next and change :
            print(f"✅ {person_current} — row {i + 1} correct")
            continue
        elif person_current != person_next :
            continue

        try:
            try:
                start_next = datetime.strptime(next_row[4], "%Y/%m/%d %H:%M")
            except ValueError:
                print(f"⚠️ 時間格式錯誤：{next_row[4]}")
                continue
            try:
                end_current = datetime.strptime(current[5], "%Y/%m/%d %H:%M")
            except ValueError:
                print(f"⚠️ 時間格式錯誤：{current[5]}")
                continue
            
            if start_next <= end_current:
                print(f"⚠️ {person_current} —  rows {i + 1} and {i + 2} overlap")
                wrong_type = "時間重疊"
                table_content[i] = insert_type(table_content[i], wrong_type)
                table_content[i + 1] = insert_type(table_content[i + 1], wrong_type)                
                wrong_array.append(table_content[i])
                wrong_array.append(table_content[i + 1])
            elif start_next > end_current:
                print(f"✅ {person_current} — rows {i + 1} and {i + 2} correct")
        except Exception as e:
            print(f"⚠️ Error comparing rows {i + 1} and {i + 2}")
            print(f"   {current}")
            print(f"   {next_row}")
            print(f"   Error: {e}")
            wrong_type = "查詢錯誤"
            table_content[i] = insert_type(table_content[i], wrong_type)
            table_content[i + 1] = insert_type(table_content[i + 1], wrong_type)

            wrong_array.append(table_content[i])
            wrong_array.append(table_content[i + 1])

    print('\n')
    return wrong_array



def bug(data):
    print('\nWellcome to the fucking far kingddom - Shrek\n')
    #開啟Chrome瀏覽器、勤務系統
    #driver = webdriver.Chrome()
    driver = setup_chrome_driver()
    wait = WebDriverWait(driver, 10)  # 最長等待 10 秒

    driver.get('https://dutymgt.tyfd.gov.tw/tyfd119/login119')

    #登入操作
    username = driver.find_element(By.ID,"_txtUsername")
    password = driver.find_element(By.ID,"_txtPassword")
    username.send_keys(data['username'])
    password.send_keys(data['password'])

    click_by_name('login', driver, wait)  #點選登入
    try:
        wait.until(EC.presence_of_element_located((By.NAME, 'ehrFrame')))
        frameM = driver.find_element(By.NAME, 'ehrFrame')
    except:
        raise Exception('帳密錯誤，請確認config.json')

    print(str_line('登入成功'))
    
    #切換到選單Frame|#frameset是組合，不是Frame
    frameM = driver.find_element(By.NAME, 'ehrFrame')
    driver.switch_to.frame(frameM)
    frameL1 = driver.find_element(By.NAME, 'sidemenuFrame')
    driver.switch_to.frame(frameL1)
    frameL2 = driver.find_element(By.NAME, 'contentSidemenu')
    driver.switch_to.frame(frameL2)

    click_by_name('nodeIcon17', driver, wait)   #轉換左方選單
    click_by_xpath('//*[@id="item23"]/tbody/tr/td[2]/a/font', driver, wait)   #勤務基準表按鈕

    #轉換右方主要內容
    driver.switch_to.parent_frame()
    driver.switch_to.parent_frame()
    frameR1 = driver.find_element(By.NAME, 'contentFrame')
    driver.switch_to.frame(frameR1)
    frameR2 = driver.find_element(By.NAME, 'mainFrame')
    driver.switch_to.frame(frameR2)
    
    #查詢月份
    try:
        title = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'title')))
    except:
        raise TimeoutError('連線逾時，請關閉後重新操作')



    export_sheet = []
    dropdown_id = '_selDeptno'
    
    for i in range(len(Select(driver.find_element(By.ID, dropdown_id)).options)):
        # REFRESH the dropdown each loop
        dropdown_element = wait.until(EC.presence_of_element_located((By.ID, dropdown_id))        )
        dropdown = Select(dropdown_element)

        # REFRESH the option list each loop
        option = dropdown.options[i]
        value = option.get_attribute('value')
        text = option.text.strip()

        # Select the option
        dropdown.select_by_value(value)
        print(f"🔽 Selecting: {text} (value={value})")

        wrong = comapre_times(driver, wait, data, text)
        for row in wrong:
            export_sheet.append(row)

    
    clean_sheet = remove_duplicates(export_sheet)
    
    # Create a new workbook and sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exception"

    # Optional: Write a header if your ex_table has consistent structure
    header = ["單位","日期", "姓名", "勤務項目", "開始時間", "結束時間", "深夜勤務時數", "金額", "錯誤種類"]
    ws.append(header)

    # Write data rows
    for row in clean_sheet:
        ws.append(row)

    # Save Excel file
    output_path = os.path.join(os.getcwd(), f"深夜食堂 - {data['unit']}.xlsx")
    wb.save(output_path)

    print(f"✅ Excel exported successfully to {output_path}")

    driver.close()
    driver.quit()

    os.startfile(output_path)

    input('輸入任意鍵結束')

    

################################################主程式################################################
if __name__ == '__main__':

    config = load_accounts()
    bug(config)