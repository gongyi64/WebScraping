# pySimpleGUI Version---NHK勤務表自動ダウンロードソフト　202311
#TEAMSへ自動ログインして、所望の年月のNHK勤務表をダウンロードし、xlsxファイル人変換する。

import PySimpleGUI as sg


sg.theme('Python')

layout =[[sg.Text('[NT勤務自動ダウンロード]',font = ('Noto Serif CJK JP',14))],

        [sg.Text('[九州沖縄TEAMSで月間勤務表出力し、★勤務確認などダウンロードフォルダに配置。] ',font = ('meiryo',10))],

        [sg.Text('出力したい年月（数字6桁）を入力',text_color='#FF0000',font =( 'meiryo,8')),sg.InputText(size = (10,2),key= '-YM-')],
        [sg.Text('[実行ボタンを押して入力。Windowを閉じる（右上×マーククリック）とスタート。] ',font = ('meiryo',10))],
        [sg.Button('実行', button_color=('red','#808080'),key = '-SUBMIT-')]]

window = sg.Window('勤務表制作APP',layout,size = (600,200))


while True:
   event,values = window.read()
   if event == '-SUBMIT-':

            num = values['-YM-']#チェックする年月を202211と6桁の数字で入力する。
            print(num)#6桁の年月



   if event == sg.WIN_CLOSED:
         break

window.close()
nen=num[:4]#頭4桁年数
tsuki = num[4:]#下2桁月数
# print(num)

# print(value)

# selenium 4

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
import os
import time
import datetime


from selenium.webdriver.support.ui import Select

from selenium.webdriver.common.by import By

time.sleep(3)


# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
# driver = webdriver.Chrome(ChromeDriverManager().install())
#


#最新のドライバーだとエラーが出るのでその対応策。https://qiita.com/hs2023/questions/ffab105c5692692624ab

options = webdriver.ChromeOptions()


print("========== 機材Teams　ログイン中========== ")
sg.popup_ok('機材TEAMSへログインします！',title = 'OK？')


driver_path = sg.popup_get_file('使用する最新chromedriverファイルを選択してください。')  # 使用するchromeのドライバーファイルを選択

# driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"

#2023_11_09 chromdriver　118→119　更新 2024\01\15 chromedriver　119→121更新20240204 外部ファイルから取得に変更

driver = webdriver.Chrome(service=ChromeService(driver_path))
#
# options.add_experimental_option("debuggerAddress", "127.0.0.1:9333")
# driver = webdriver.Chrome(executable_path=driver_path, options=options)



# driver = webdriver.Chrome(service,options=options)


# ページのタイトルを表示する

driver.get("https://gate.isso.nhk.or.jp/fw/dfw/lkteams/")

driver.maximize_window()

time.sleep(2)

form = driver.find_element(By.XPATH,'//*[@id="selector_form"]/div/div/div/button')
#
# time.sleep(1）

driver.implicitly_wait(2)

form.click()



form = driver.find_element(By.XPATH,'//*[@id="username"]')

form.send_keys('1010406239')

form = driver.find_element(By.XPATH,'//*[@id="uid_password"]')

form.send_keys('*kabu92772462')

form.submit()

time.sleep(1)

form = driver.find_element(By.XPATH,'//*[@id="btnClose"]')

form.click()

time.sleep(1)

form = driver.find_element(By.XPATH,'//*[@id="accordionSidebar"]/li[6]/a')

form.click()


dropdown = driver.find_element(By.ID,'page-top')

time.sleep(1)

driver.find_element(By.XPATH,'//*[@id="collapse5"]/div/a[5]').click()
# print(driver.page_source)

# print(driver.current_url)

# driver.get("https://gate.isso.nhk.or.jp/fw/dfw/lkteams/schedule/rs_print_monschedule_t1.php?action=init")

driver.implicitly_wait(10)

handle_array = driver.window_handles

print("handle_arrayの表示配列最初と次")
print(handle_array[0])
print(handle_array[1])
# print(handle_array[2])

driver.switch_to.window(handle_array[1])

time.sleep(5)


years = driver.find_element(By.CSS_SELECTOR,"#FormData > div.control.cfx > select:nth-child(8)")
years_select = Select(years)
# years.send_keys("2023")
years_select.select_by_value(nen)


months = driver.find_element(By.CSS_SELECTOR,"#FormData > div.control.cfx > select:nth-child(10)")
months_select = Select(months)
# months.send_keys("11月")
months_select.select_by_value(tsuki)



##FormData > div.control.cfx > select:nth-child(8)#年のセレクトCSS_Selector

##FormData > div.control.cfx > select:nth-child(10)#月のセレクトCSS_Selector



driver.find_element(By.CSS_SELECTOR,'#excelout').click()

# script = 'javascript:void(0);'
# form.driver.execute_script(script)

print("出力押した")

time.sleep(15)
driver.find_element(By.CSS_SELECTOR,'#close > a').click()
# driver.implicitly_wait(100) #ダウンロードフォルダへ格納　これを別フォルダへ移動させる。
dir_path = "C:\\Users\\406239\\OneDrive - (株)NHKテクノロジーズ\\デスクトップ\\ドキュメント\\Downloads"

# dir_path = 'C:\\Downloads'

# C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\ドキュメント\Downloads

files = os.listdir(dir_path)

print(files)#ダウンロードフォルダへ格納させたファイル名取得。このファイルで必要なものを抽出して所望のファルダへ移動させる。日付の後が大きいものが最新。

day = datetime.date.today()
today = str(day)[:4]+str(day)[5:7]

files_in = [s for s in files if today in s]#出力した月の中で最新のもの

print('ファイル名表示')
print(files_in)

newest_file = max(files_in)#最新ファイルの取得#出力した月の中で最新のもの
print('最新ファイル名表示')
print(newest_file)

file_name,ext = os.path.splitext(newest_file)
newest_file = str(newest_file)
print(file_name)
print(ext)

dt_now = datetime.datetime.now()
output_time = dt_now.strftime('%Y%m%d_%H%M')
print(type(output_time))
print(output_time)

# print('最新の勤務ファイル')

#fでformat変数、ｒで\\を\で表記可能。変数は、{}で囲む。文字は、””で囲む。formatで書くと、+は不要なのでカンタン。

oldpath = fr"C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\ドキュメント\Downloads\{newest_file}"

newpath = fr"C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_{num}_{output_time}.xls"


print(os.path.exists(oldpath))

os.rename(oldpath,newpath)

print(os.path.exists(newpath))

#xlsを一旦開いてから、xlsxで保存する。openpyexLを使用するため。変換は面倒そうなので、これがカンタン。

import xlwings as xw

path = fr'C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_{num}_{output_time}.xls'

wb = xw.Book(path)

path = fr'C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_{num}_{output_time}.xlsx'

wb.save(path)

wb.close()

#========================================ExcelのPDF化=========================

# from spire.xls import *
# from spire.xls.common import *
#
# workbook = Workbook()
# workbook.LoadFromFile(path)
#
# workbook.SaveToFile("output/ExcelをPDFに変換.pdf", FileFormat.PDF)
# workbook.Dispose()

#===================================================================================



