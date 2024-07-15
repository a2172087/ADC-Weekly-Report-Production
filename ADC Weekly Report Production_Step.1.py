import sys
sys.stdout.reconfigure(encoding='utf-8')
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import glob
import shutil
from openpyxl.utils import get_column_letter
import datetime
from openpyxl.utils import column_index_from_string
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 新ChromeDriver的路徑
path_to_new_chromedriver = r'D:\Racky\公用相關\暫存檔案\chromedriver-win64\chromedriver.exe'

# 創建ChromeDriver的配置
chrome_options = webdriver.ChromeOptions()
# 在此添加其他選項，例如禁用圖片、使用無頭模式等
chrome_options.add_argument("--disable-images")  # 禁用圖片載入
chrome_options.add_argument("--headless")  # 使用無頭模式

# 創建新的Service物件，使用新的ChromeDriver
new_service = Service(path_to_new_chromedriver)

# 創建新的WebDriver實例，使用新的Service物件和chrome_options
driver = webdriver.Chrome(service=new_service, options=chrome_options)

url = "http://tstpas/TPAS/index.jsp"
user_id = "A005772"
user_pw = "A11111111111"

# 創建 WebDriver 實例 (以 Chrome 爲例)
s = Service(path_to_new_chromedriver)
driver = webdriver.Chrome(service=s)

# 打開網站
driver.get(url)

# 等待 mainFrame 加載完成
wait = WebDriverWait(driver, 10)
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "mainFrame")))

# 定位 User ID 和 User PW 輸入框 (根據網頁元素的name等定位)
user_id_input = driver.find_element(By.NAME, "userid")
user_pw_input = driver.find_element(By.NAME, "password")

# 輸入 User ID 和 User PW
user_id_input.send_keys(user_id)
user_pw_input.send_keys(user_pw)

# 提交登錄表單
login_form = driver.find_element(By.NAME, "form1")
login_form.submit()

# 如果需要, 切換回主窗口
driver.switch_to.default_content()

# 切換到包含 'Report' 的 frame
frame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//frame[@src="/TPAS/edu/edu_top.jsp"]')))
driver.switch_to.frame(frame)

# 點擊 'Report' 鏈接
xpath = '//a[@href="/REPORT/"]'
element_to_click = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xpath)))
element_to_click.click()

# 暂停 0 秒以观察网页操作
#time.sleep(100000)

# 切換到索引為0的frame（第一個frame）
driver.switch_to.frame(0)

# 點擊 'LOT INFO' 鏈接
xpath = '//td[@id="dm0m0i2tdT"]'
element_to_click = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xpath)))
element_to_click.click()

# 將焦點切換回父框架
driver.switch_to.parent_frame()

# 切換到包含 'AVI Inspection Report' 的 frame
frame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//frame[@src="welcome.jsp"]')))
driver.switch_to.frame(frame)

# 點擊 'AVI Inspection Report' 鏈接
xpath = '//td[@id="dm0m9i14tdT"]'
element_to_click = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xpath)))
element_to_click.click()

#等待 "stage" 的下拉框出現
stage_select = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "stage")))

#找到 "stage" 下拉框並選擇 "AVI" 選項
from selenium.webdriver.support.ui import Select

select_stage = Select(stage_select)
select_stage.select_by_visible_text("AVI")

# 計算當周的週四
today = datetime.date.today()
days_to_thursday = (3 - today.weekday()) % 7
this_week_thursday = today + datetime.timedelta(days=days_to_thursday)
this_week_thursday_str = this_week_thursday.strftime("%Y-%m-%d")

# 計算一週前的週四的日期
one_week_ago_thursday = this_week_thursday - datetime.timedelta(days=7)
one_week_ago_thursday_str = one_week_ago_thursday.strftime("%Y-%m-%d")

# 找到名爲 "begin_date" 的輸入框並設置新的日期值
begin_date_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "begin_date")))
begin_date_input.clear()
begin_date_input.send_keys(one_week_ago_thursday_str)

# 找到名爲 "end_date" 的輸入框並設置新的日期值
end_date_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "end_date")))
end_date_input.clear()
end_date_input.send_keys(this_week_thursday_str)

# 找到名爲 "Submit" 的按鈕並點擊
submit_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "Submit")))
submit_button.click()

# 暫停 5 秒以觀察網頁操作
time.sleep(5)

# 找到名爲 "EXCEL" 的按鈕並點擊
submit_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "EXCEL")))
submit_button.click()

# 等待文件下載完成 (根據文件大小和網絡速度調整等待時間)
time.sleep(5)

# 指定下載文件夾
download_folder = r"C:\Users\A005772\Downloads"

# 獲取下載文件夾中的最新文件
list_of_files = glob.glob(download_folder + r'\*')
latest_file = max(list_of_files, key=os.path.getctime)

# 構建新文件名
new_file_name = "TS AVI Output for PAD.xlsx"

# 將下載的文件重命名
shutil.move(latest_file, os.path.join(download_folder, new_file_name))

# 找到名爲 "EXCEL" 的按鈕並點擊
submit_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "EXCEL")))
submit_button.click()

# 等待文件下載完成 (根據文件大小和網絡速度調整等待時間)
time.sleep(5)

# 指定下載文件夾
download_folder = r"C:\Users\A005772\Downloads"

# 獲取下載文件夾中的最新文件
list_of_files = glob.glob(download_folder + r'\*')
latest_file = max(list_of_files, key=os.path.getctime)

# 構建新文件名
new_file_name = "TS AVI Output for Bump.xlsx"

# 將下載的文件重命名
shutil.move(latest_file, os.path.join(download_folder, new_file_name))

# 找到名爲 "adc_format" 的按鈕並點擊
submit_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "adc_format")))
submit_button.click()

# 找到名爲 "Submit" 的按鈕並點擊
submit_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "Submit")))
submit_button.click()

# 暫停 5 秒以觀察網頁操作
time.sleep(5)

# 找到名爲 "EXCEL" 的按鈕並點擊
submit_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "EXCEL")))
submit_button.click()

# 等待文件下載完成 (根據文件大小和網絡速度調整等待時間)
time.sleep(5)

# 指定下載文件夾
download_folder = r"C:\Users\A005772\Downloads"

# 獲取下載文件夾中的最新文件
list_of_files = glob.glob(download_folder + r'\*')
latest_file = max(list_of_files, key=os.path.getctime)

# 構建新文件名
new_file_name = "TS AVI ADC Output for Pad.xlsx"

# 將下載的文件重命名
shutil.move(latest_file, os.path.join(download_folder, new_file_name))

# 找到名爲 "EXCEL" 的按鈕並點擊
submit_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "EXCEL")))
submit_button.click()

# 等待文件下載完成 (根據文件大小和網絡速度調整等待時間)
time.sleep(5)

# 指定下載文件夾
download_folder = r"C:\Users\A005772\Downloads"

# 獲取下載文件夾中的最新文件
list_of_files = glob.glob(download_folder + r'\*')
latest_file = max(list_of_files, key=os.path.getctime)

# 構建新文件名
new_file_name = "TS AVI ADC Output for Bump.xlsx"

# 將下載的文件重命名
shutil.move(latest_file, os.path.join(download_folder, new_file_name))

# 暫停 5 秒以觀察網頁操作
time.sleep(5)

# 最後，關閉瀏覽器
driver.quit()