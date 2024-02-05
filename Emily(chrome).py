# pySimpleGUI Version---Emily 自動ログインツール　202312
#Emilyへ自動ログインして、各自のメニューにアクセスする。要員実績を自動で入力する（全部休日で埋める）ログインは、各自マンナンバーで。


import PySimpleGUI as sg
import pandas as pd
import re
import openpyxl

file_name = sg.popup_get_file('社員番号、氏名、パスワードの読み込みに使用するファイルを選択してください。')  # 使用する出力したの勤務チェック用のファイルを選択

df = pd.read_excel(file_name,sheet_name = 'Pass')#sheet_name ＝　Pass　に、pwd　を保存している。
manNos = []
for i in range(len(df['氏名'])):
    manNos.append(str(df['社員番号'][i])+' '+df['氏名'][i])

print(manNos)

#manNos = ('827861 相庭直史','406239 白川公一','380672 三角和浩','378035 岩田貴夫','805519 仲本祥子','880079 砂川航輝','806185 辻ひかる','880334   山城実咲','880518 古波蔵晃久','410993 黒岩英次','710463 山城徳松')

#manNos = ('827861 相庭直史　ap827861','406239 白川公一　ap406239','380672 三角和浩　ap380672','378035 岩田貴夫　ap378035','805519 仲本祥子   ap805519','880079 砂川航輝　ap880079','806185 辻ひかる　ap806185','880334   山城実咲 ap880334','880518 古波蔵晃久 ap880518','410993 黒岩英次   ap410993','710463 山城徳松   ap710463')

#manNos = ('827861','406239','380672','378035','805519','880079','806185','880334','880518','410993','710463')

sg.theme('Python')

layout =[[sg.Text('[NT_Emily_自動操作ソフト]',font = ('Noto Serif CJK JP',14))],

         [sg.Text('[Emilyで要員実績の箱を自動作成。要員作成する各自Noでログインしてください。] ',font = ('meiryo',10))],

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
eplyName = re.sub(r"[0-9]+", "", input_eplyNo)#社員番号削除して氏名のみに。

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

import calendar
from selenium.webdriver.support.ui import Select


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

# driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"

driver_path = sg.popup_get_file('使用する最新chromedriverファイルを選択してください。')  # 使用するchromeのドライバーファイルを選択

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
#====================pwdを外ファイルからゲットルーチンここから。20230118実装

#Emily_Pass.xlsxがパスワード保管ファイル

# file_name = 'c:/Users/406239/OneDrive - (株)NHKテクノロジーズ/デスクトップ/★勤務確認などのダウンロードデータ★/Emily_Files/Emily_Pass.xlsx'


df = pd.read_excel(file_name)

print(eplyNo[0])

eplyNo = int(eplyNo[0])#eplyNoは、リストなので、0番目を抽出し、intに変更。strだとエラー。

print (df[df['社員番号'] == eplyNo])#ログインする人の番号が含まれるdfを抽出


df_login = df[df['社員番号'] == eplyNo]


pwd = df_login.iat[0,2]#そのパスワードのみを抽出

print(pwd)

#====================================================================

form = driver.find_element(By.XPATH,'//*[@id="LoginAccountText"]')

form.send_keys(eplyNo)#最初に取得したマンナンバーを入力

form = driver.find_element(By.XPATH,'//*[@id="PassText"]')

form.send_keys(pwd)


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

anken_file_name = sg.popup_get_file('案件番号の書き出し読み出しに使用するファイルを選択してください。')  # 案件番号を保存、読みだすExcelファイルを選択

#------------案件番号取得用年月入力=====================

sg.theme('SystemDefault')

layout = [[sg.Text('要員実績を入力する年月を入れてください。年+月6桁',text_color='#FF0000',font =( 'meiryo,6')),sg.InputText(size = (10,2),key= '-YM-')],
          # [sg.Text('誰の案件？',text_color='#FF0000',font =( 'meiryo,8')),sg.InputText(size = (10,2),key= '-NM-')],
          # [sg.Listbox(manNos,size =(25,len(manNos)),key='-NM-')],
          [sg.Text('Text', key = '-text1-')],
          [sg.Button('入力', button_color=('red', '#808080'), key='-SUBMIT-'),
           sg.Text('入力ボタンを押した後,Windowを閉じてください。\n入力した年月でログインした人の要員実績の箱が自動作成されます。', font=('Noto Serif CJK JP', 10))]]

window = sg.Window('要員実績箱作成ツール', layout, size=(500, 100))

while True:
    event,values = window.read()

    if event == sg.WIN_CLOSED:
        break

    elif event == '-SUBMIT-':
        ym = values['-YM-']
        window['-text1-'].update(values['-YM-'])
        # input_eplyNo = values['-NM-'][0]
        # eplyName = re.sub(r"[0-9]+", "", input_eplyNo)#社員番号と名前から社員番号削除してフルネームのみに。


window.close()

#==========================案件番号取得ルーチン　20230123実装===============

# file_name = pd.ExcelFile( r'c:/Users/406239/OneDrive - (株)NHKテクノロジーズ/デスクトップ/★勤務確認などのダウンロードデータ★/Emily_Files/Emily_anken.xlsx')

file_name = pd.ExcelFile(anken_file_name)

# wb = openpyxl.load_workbook('c:/Users/406239/OneDrive - (株)NHKテクノロジーズ/デスクトップ/★勤務確認などのダウンロードデータ★/Emily_Files/Emily_anken.xlsx')

wb = openpyxl.load_workbook(anken_file_name)

target_name = str(ym) + '案件番号'

print(file_name)
print(wb)

check = False

for ws in wb.worksheets:  # Emily_anken/xlsxに、所望のシートが存在するか判定。
    if ws.title == target_name:
        check = True

if check == True:
    print(target_name + 'は、存在します。')
    sg.popup_ok('入力した年月の案件番号sheetは存在するので引き続き要員実績を作成します。')  # あれば、読みだしてｄｆに。

    df = pd.read_excel(file_name, sheet_name=ym + '案件番号', dtype=str)

    print(eplyName)

    print(df[df['name'] == str(ym) + eplyName])  # ログインする人の番号（年月+氏名）が含まれるdfを抽出

    if df[df['name'] == str(ym) + eplyName].empty:
        sg.popup_ok(eplyName + 'さんの案件番号は存在しませんので、作成してください')
    else:

        df_anken = df[df['name'] == str(ym) + eplyName]

        kobetsu_No = df_anken.iat[0, 2]  # その案件番号のみを抽出
        print(kobetsu_No)

else:
    print(target_name + 'は、存在しません。')  # なければ、中断。

    sg.popup_ok('入力した年月の案件番号は存在しないので、処理を中断します。\n作成してください。')

#--------------------------------------------------------------------

#os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。
time.sleep(2)

driver.switch_to.frame(1)#iFrameの最初に切り替え。２つあるが、２番目（１）のiFrameに切り替える。

# kobetsu_No = '2024001528'#今は直接入力。なかもと5月

taishou_mon = ym

form = driver.find_element(By.XPATH,'//*[@id="ProposalNoText"]')

# print(form)

form.send_keys(kobetsu_No)#案件番号を入力

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。
#
#
driver.find_element(By.XPATH,'//*[@id="TargetMonthBox"]').clear()

driver.find_element(By.XPATH,'//*[@id="TargetMonthBox"]').send_keys(taishou_mon)#対象年月を変更。

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

time.sleep(3)

# os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。

taishou_mon = str(ym)#スライス処理のためSTR化
taishou_year = taishou_mon[:4]#年のみ取り出し
taishou_month = taishou_mon[-2:]#月のみ取り出し
taishou_year = int(taishou_year)#calendarモジュール使用のため、INT化
taishou_month = int(taishou_month)#calendarモジュール使用のため、INT化

nichi = calendar.monthrange(taishou_year,taishou_month)[1]#対象の月の日数判定



for i in range(1,nichi+1):

    dropdown1 = driver.find_element(By.XPATH,'//*[@id="NewOpeDtlCodeDrop"]')#勤務内容選択　休日

    select = Select(dropdown1)

    select.select_by_index(len(select.options)-1)




    time.sleep(1)

    driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

    driver.find_element(By.XPATH,'//*[@id="NewEmpCodeText"]').send_keys(str(eplyNo))##担当者　マンナンバー

    driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

    time.sleep(1)

    driver.find_element(By.XPATH,'//*[@id="NewDisplayOrderText"]').send_keys(i)#順

    driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

    time.sleep(3)

#driver.find_element(By.XPATH,'//*[@id="NewWorkDateBox"]"]').send_keys('2024/05/'{i+1})#実施年月日

    driver.find_element(By.CSS_SELECTOR,'#NewWorkDateBox').send_keys(str(taishou_year)+'/'+str(taishou_month)+'/'+str(i).zfill(2))#実施年月日


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


