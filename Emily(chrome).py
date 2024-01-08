# pySimpleGUI Version---Emily 自動ログインツール　202312
#Emilyへ自動ログインして、各自のメニューにアクセスする。。


import PySimpleGUI as sg
import re

manNos = ('827861 相庭直史','406239 白川公一','380672 三角和浩','378035 岩田貴夫','805519 仲本祥子','880079 砂川航輝','806185 辻ひかる','880334   山城実咲','880518 古波蔵晃久','410993 黒岩英次','710463 山城徳松')

#manNos = ('827861 相庭直史　ap827861','406239 白川公一　ap406239','380672 三角和浩　ap380672','378035 岩田貴夫　ap378035','805519 仲本祥子   ap805519','880079 砂川航輝　ap880079','806185 辻ひかる　ap806185','880334   山城実咲 ap880334','880518 古波蔵晃久 ap880518','410993 黒岩英次   ap410993','710463 山城徳松   ap710463')

#manNos = ('827861','406239','380672','378035','805519','880079','806185','880334','880518','410993','710463')

sg.theme('Python')

layout =[[sg.Text('[NT_Emily_自動操作ソフト]',font = ('Noto Serif CJK JP',14))],

         [sg.Text('[Emilyに別ブラウザでログインしたいときに、、、。] ',font = ('meiryo',10))],

         [sg.Listbox(manNos,size =(25,len(manNos)),key='-MN-')],

         [sg.Text('Text', key = '-text1-')],

         [sg.Button('実行', button_color=('red','#808080'),key = '-SUBMIT-')]]

window = sg.Window('Emily_APP',layout,size = (750,350))



while True:
    event,values = window.read()

    if event == sg.WIN_CLOSED:
        break

    elif event == '-SUBMIT-':
        window['-text1-'].update(values['-MN-'][0])
        input_eplyNo = values['-MN-'][0]


window.close()



# input_eplyNo = values['-MN-'][0]

print(input_eplyNo)#所得したのは、リスト型

eplyNo = re.findall(r'\d+', input_eplyNo)#名前を除去して社員番号のみにして、ログインに使用する。

# eplyNo = eplyNo[:7]

print(eplyNo[0])#リストの要素を文字列として取得。この場合は、要素１つなので[0]

#eplypwd = re.findall(r'^[a-zA-Z0-9]{7}$, input_eplyNo)#名前とマンナンバーを除去してPWDのみにして、ログインに使用する。7桁の英数字想定。ｓ+マンナンバーなど。

# selenium 4

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
import os
import time
import datetime
import sys
import signal
from selenium.webdriver.common.action_chains import ActionChains

import pyexcel as p
import glob

from selenium.webdriver.support.ui import Select

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.common.by import By

from selenium.webdriver.common.keys import Keys

time.sleep(3)
#from webdriver_manager.chrome import ChromeDriverManager

#driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

#最新のドライバーだとエラーが出るのでその対応策。https://qiita.com/hs2023/questions/ffab105c5692692624ab

options = webdriver.ChromeOptions()


print("========== Emily　ログイン中========== ")
sg.popup_ok('Emilyへログインします！',title = 'OK？')


# service = Service(driver_path)
# service = Service(executable_path=driver_path)# 2) executable_pathを指定
# driver = webdriver.Chrome(service=service)# 3) serviceを渡す

# 起動時にオプションをつける。（ポート指定により、起動済みのブラウザのドライバーを取得）
# driver_path = "C:\\Users\\406239\\AppData\\Local\\Programs\\Python\\Python39\\chromedriver_binary\\chromedriver.exe"

driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"

#2023_11_09 chromdriver　118→119　更新

driver = webdriver.Chrome(service=ChromeService(driver_path))
# options.add_experimental_option("detach",True)#ドライバプロセスの終了後も開いたままにする。使い勝手よくなさそうなのでやめ。

options.add_experimental_option("debuggerAddress", "127.0.0.1:9333")
# driver = webdriver.Chrome(executable_path=driver_path, options=options)



# driver = webdriver.Chrome(service,options=options)


# ページのタイトルを表示する

driver.get("https://test9.emily.nhk-tech.co.jp/GRANDIT/CM_AC_03_S01.aspx")

driver.maximize_window()

time.sleep(2)
#
# form = driver.find_element(By.XPATH,'//*[@id="selector_form"]/div/div/div/button')
# #
# # time.sleep(1）

driver.implicitly_wait(2)

# form.click()



form = driver.find_element(By.XPATH,'//*[@id="LoginAccountText"]')

form.send_keys(eplyNo)#最初に取得したマンナンバーを入力

form = driver.find_element(By.XPATH,'//*[@id="PassText"]')

form.send_keys('09-Hdo9QDw ')


time.sleep(2)

form = driver.find_element(By.XPATH,'//*[@id="LoginButton"]/span')

form.click()


time.sleep(3)




#ログインまでは、出来ているこの後は、まだ。メニュー画面にログインまで。#c_12

handle_array = driver.window_handles

print("handle_arrayの表示配列最初と次")#windowshandleは2つ
print(handle_array[0])
print(handle_array[1])
# print(handle_array[2])


driver.switch_to.window(handle_array[1])

print(driver.current_url)

time.sleep(2)


# 要素を特定する


#elem = driver.find_elements(By.XPATH,'//*[@id="__pageIframe_01703644264780"]')

# elem = driver.find_elements(By.CSS_SELECTOR,'#__pageIframe_01703644347699')
#iFrameをDevelopersで検索してそのiFrameのXPATHをさがす


#
driver.switch_to.frame(0)#iFrameの最初に切り替え。１つしかないが、classが毎回変わるので、1番目（０）のiFrameに切り替えるということにした。


# but = driver.find_elements(By.CLASS_NAME,'hd_bg_4')#あった

dropdown = driver.find_element(By.CSS_SELECTOR,'#MenuDropDownList')
print('釦を取得できたらあり')
print(dropdown)

select = Select(dropdown)

select.select_by_index(1)

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。



time.sleep(3)

#M技管理メニュー操作


driver.find_element(By.XPATH,'//*[@id="c_12"]').click()#要員実績をクリック

driver.find_element(By.XPATH,'//*[@id="Form1"]/table/tbody/tr/td/table[2]/tbody/tr/td[2]/div[3]/table/tbody/tr/td[2]/div/a[3]').click()#225_要員実績入力をクリック

time.sleep(3)

handle_array = driver.window_handles

print("handle_arrayの表示配列最初と次")#windowshandleは2つ
print(handle_array[0])
print(handle_array[1])
# print(handle_array[2])


driver.switch_to.window(handle_array[1])

#案件番号の個別入力ルーチン（テスト後生かす予定）
#--------------------------------------------------------------------


# value = sg.popup_get_file('個別案件入力で作成した個人の案件番号を入力してください。')  # 使用する出力したの勤務チェック用のファイルを選択
#
# sg.theme('SystemDefault')
#
# layout = [[sg.Text('Emily・[個人要員案件番号の入力]', font=('Noto Serif CJK JP', 10))],
#           [sg.Text('個別案件で作成した案件番号を入力', font=('meiryo', 10))],
#           # [sg.Text('年月を入力',text_color='#FF0000',font =( 'meiryo,8')),sg.InputText(size = (10,2),key= '-YM-')],
#           [sg.Text('案件番号', text_color='#FF0000', font='meiryo,8'),
#            sg.Combo(['月前半', '月後半'], size=(10, 2), key='-UPLW-')],
#           [sg.Button('入力', button_color=('red', '#808080'), key='-SUBMIT-'),
#            sg.Text('入力ボタンを押した後,Windowを閉じてください。', font=('Noto Serif CJK JP', 10))]]
#
# window = sg.Window('基本業務入力APP', layout, size=(500, 150))
#
# while True:
#     event, values = window.read()
#     if event == '-SUBMIT-':
#         # num = values['-YM-']
#         # print(num)
#         num1 = values['-UPLW-']
#         print(num1)
#
#     if event == sg.WIN_CLOSED:
#         break
#
# window.close()


#--------------------------------------------------------------------

#os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。
time.sleep(2)

driver.switch_to.frame(1)#iFrameの最初に切り替え。２つあるが、２番目（１）のiFrameに切り替える。

kobetsu_No = '2024001176'#今は直接入力。こはぐら5月

form = driver.find_element(By.XPATH,'//*[@id="ProposalNoText"]')

# print(form)

form.send_keys(kobetsu_No)#案件番号を入力

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。
#
#
driver.find_element(By.XPATH,'//*[@id="TargetMonthBox"]').clear()

driver.find_element(By.XPATH,'//*[@id="TargetMonthBox"]').send_keys('202405')#対象年月を変更。この場合は、202305に。

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

time.sleep(3)

# os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。

for i in range(1,3):

    dropdown1 = driver.find_element(By.XPATH,'//*[@id="NewOpeDtlCodeDrop"]')#勤務内容選択　休日

    select = Select(dropdown1)

    select.select_by_index(len(select.options)-1)



# dropdown2 = driver.find_element(By.ID,'#NewDutyCodeDrop')#担務　担当

# dropdown2 = driver.find_element(By.XPATH,'//*[@id="NewDutyCodeDrop"]')
#
# print(dropdown2)
#
# select = Select(dropdown2)
#
# select.select_by_index(1)
#
# driver.implicitly_wait(10)
#
# driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

# select.select_by_index(len(select.options)-2)
# select.select_by_visible_text('担当')

    time.sleep(1)

    driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。
#
#
    driver.find_element(By.XPATH,'//*[@id="NewEmpCodeText"]').send_keys(eplyNo[0])##担当者　マンナンバー

    driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

    time.sleep(1)

    driver.find_element(By.XPATH,'//*[@id="NewDisplayOrderText"]').send_keys('i')#順

    driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

    time.sleep(3)

#driver.find_element(By.XPATH,'//*[@id="NewWorkDateBox"]"]').send_keys('2024/05/'{i+1})#実施年月日

    driver.find_element(By.CSS_SELECTOR,'#NewWorkDateBox').send_keys('2024/05/'+str(i).zfill(2))#実施年月日


    time.sleep(3)

    dropdown2 = driver.find_element(By.XPATH,'//*[@id="NewDutyCodeDrop"]')#担務入力なぜか先に入力すると入らないので、最後に。

    print(dropdown2)

    select = Select(dropdown2)

    select.select_by_index(1)

    driver.implicitly_wait(10)

    driver.find_element(By.XPATH,'//*[@id="RegistButton"]/span').click()#登録ボタン


time.sleep(3)







#登録ボタン





os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。



#
# print("handle_arrayの表示配列最初と次")
# print(handle_array[0])
# print(handle_array[1])
# # print(handle_array[2])
#
# driver.switch_to.window(handle_array[1])
#
# time.sleep(5)
#
#
# years = driver.find_element(By.CSS_SELECTOR,"#FormData > div.control.cfx > select:nth-child(8)")
# years_select = Select(years)
# # years.send_keys("2023")
# years_select.select_by_value(nen)
#
#
# months = driver.find_element(By.CSS_SELECTOR,"#FormData > div.control.cfx > select:nth-child(10)")
# months_select = Select(months)
# # months.send_keys("11月")
# months_select.select_by_value(num)
#
#
#
# ##FormData > div.control.cfx > select:nth-child(8)#年のセレクトCSS_Selector
#
# ##FormData > div.control.cfx > select:nth-child(10)#月のセレクトCSS_Selector
#
#
#
# driver.find_element(By.CSS_SELECTOR,'#excelout').click()
#
# # script = 'javascript:void(0);'
# # form.driver.execute_script(script)
#
# print("出力押した")
#
# time.sleep(15)
# driver.find_element(By.CSS_SELECTOR,'#close > a').click()
# # driver.implicitly_wait(100) #ダウンロードフォルダへ格納　これを別フォルダへ移動させる。
#
# dir_path = "C:\\Users\\406239\\OneDrive - (株)NHKテクノロジーズ\\デスクトップ\\ドキュメント\\Downloads"
#
# # C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\ドキュメント\Downloads
#
# files = os.listdir(dir_path)
#
# print(files)#ダウンロードフォルダへ格納させたファイル名取得。このファイルで必要なものを抽出して所望のファルダへ移動させる。日付の後が大きいものが最新。
#
# files_in = [s for s in files if '202311' in s]#出力した月の中で最新のもの
#
# print(files_in)
#
# newest_file = max(files_in)#最新ファイルの取得#出力した月の中で最新のもの
#
# print(newest_file)
#
# file_name,ext = os.path.splitext(newest_file)
# newest_file = str(newest_file)
# print(file_name)
# print(ext)
#
# dt_now = datetime.datetime.now()
# output_time = dt_now.strftime('%Y%m%d_%H%M')
# print(type(output_time))
# print(output_time)
#
# # print('最新の勤務ファイル')
#
# #fでformat変数、ｒで\\を\で表記可能。変数は、{}で囲む。文字は、””で囲む。formatで書くと、+は不要なのでカンタン。
#
# oldpath = fr"C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\ドキュメント\Downloads\{newest_file}"
#
# newpath = fr"C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_202312_{output_time}.xls"
#
#
# print(os.path.exists(oldpath))
#
# os.rename(oldpath,newpath)
#
# print(os.path.exists(newpath))
#
# #xlsを一旦開いてから、xlsxで保存する。openpyexLを使用するため。変換は面倒そうなので、これがカンタン。
#
# import xlwings as xw
#
# path = fr'C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_202312_{output_time}.xls'
#
# wb = xw.Book(path)
#
# path = fr'C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_202312_{output_time}.xlsx'
#
# wb.save(path)
#
# wb.close()
#
#
#
#
# # xls_path = r"C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル"
#
# # os.path("xls_path")
#
# # def convert_xls_to_xlsx():
# #     it = glob.glob("*.xls")
# #     for xls in it:
# #         xlsx = "{}".format(xls) + "x"
# #         print(xlsx)
# #         p.save_book_as(file_name='{}'.format(xls), dest_file_name='{}'.format(xlsx))
# #
# # print(sys.argv[0])
# #
# # print(os.listdir(xls_path))
# #
# # print(os.path.isdir(xls_path))
# #
# # convert_xls_to_xlsx()
# #
# print(os.listdir(xls_path))
