# pySimpleGUI Version---Emily 自動ログインツール　202401
#Emilyへ自動ログインして、各自のメニューにアクセスする。案件制作を自動で。Ver.作成中。



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
eplyName = re.sub(r"[0-9]+", "", input_eplyNo)#社員番号削除して氏名のみに。
# eplyNo = eplyNo[:7]

print(eplyNo[0])#リストの要素を文字列として取得。この場合は、要素１つなので[0]
print(eplyName[0])

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
import calendar
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
sg.popup_ok('Emilyへログイン！',title = 'LOGIN')


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

# ページのタイトルを表示する

driver.get("https://test9.emily.nhk-tech.co.jp/GRANDIT/CM_AC_03_S01.aspx")

driver.maximize_window()

time.sleep(2)

driver.implicitly_wait(2)



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


dropdown = driver.find_element(By.CSS_SELECTOR,'#MenuDropDownList')
print('釦を取得できたらあり')
print(dropdown)

select = Select(dropdown)

select.select_by_index(1)

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。



time.sleep(3)

#M技管理メニュー操作


driver.find_element(By.XPATH,'//*[@id="c_11"]').click()#個別案件/要員をクリック

time.sleep(3)

driver.find_element(By.XPATH,'//*[@id="Form1"]/table/tbody/tr/td/table[2]/tbody/tr/td[2]/div[3]/table/tbody/tr/td[1]/div/a[2]').click()#182_個別案件入力をクリック

time.sleep(3)

handle_array = driver.window_handles

print("handle_arrayの表示配列最初と次")#windowshandleは2つ
print(handle_array[0])
print(handle_array[1])
# print(handle_array[2])


driver.switch_to.window(handle_array[1])


sg.theme('SystemDefault')

layout = [[sg.Text('年月を入力',text_color='#FF0000',font =( 'meiryo,8')),sg.InputText(size = (10,2),key= '-YM-')],
          # [sg.Text('誰の案件？',text_color='#FF0000',font =( 'meiryo,8')),sg.InputText(size = (10,2),key= '-NM-')],
          [sg.Listbox(manNos,size =(25,len(manNos)),key='-NM-')],
          [sg.Text('Text', key = '-text1-')],
          [sg.Button('入力', button_color=('red', '#808080'), key='-SUBMIT-'),
           sg.Text('入力ボタンを押した後,Windowを閉じてください。', font=('Noto Serif CJK JP', 10))]]

window = sg.Window('案件自動作成ツール', layout, size=(500, 300))

while True:
    event,values = window.read()

    if event == sg.WIN_CLOSED:
        break

    elif event == '-SUBMIT-':
        ym = values['-YM-']
        window['-text1-'].update(values['-NM-'][0])
        input_eplyNo = values['-NM-'][0]
        eplyName = re.sub(r"[0-9]+", "", input_eplyNo)#社員番号と名前から社員番号削除してフルネームのみに。


window.close()


# while True:
#     event, values = window.read()
#     if event == '-SUBMIT-':
#         ym = values['-YM-']
#         nm = values['-NM-'][0]
#         print(ym)
#         print(nm)
#         # num1 = values['-YM-']
#         # print(num1)
#
#     if event == sg.WIN_CLOSED:
#         break
#
# window.close()



#案件番号の個別入力ルーチン（テスト後生かす予定）
#--------------------------------------------------------------------




# sg.theme('SystemDefault')
#
# layout = [[sg.Text('年月を入力',text_color='#FF0000',font =( 'meiryo,8')),sg.InputText(size = (10,2),key= '-YM-')],
#           [sg.Button('入力', button_color=('red', '#808080'), key='-SUBMIT-'),
#            sg.Text('入力ボタンを押した後,Windowを閉じてください。', font=('Noto Serif CJK JP', 10))]]
#
# window = sg.Window('基本業務入力APP', layout, size=(500, 150))
#
# while True:
#     event, values = window.read()
#     if event == '-SUBMIT-':
#         num = values['-SUBMIT-']
#         # print(num)
#         num1 = values['-SBMIT-']
#         print(num1)
#
#     if event == sg.WIN_CLOSED:
#         break
#
# window.close()


#--------------------------------------------------------------------

#os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。
time.sleep(2)

driver.switch_to.frame(1)#iFrameの最初に切り替え。２つあるが、2番目（1）のiFrameに切り替える。

form = driver.find_element(By.XPATH,'//*[@id="ProgramBusinessCodeText"]')

form.send_keys('2006101343')#番組業務番号の入力

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)

time.sleep(3)
dropdown = driver.find_element(By.XPATH, '//*[@id="MediaBusinessClassCodeList"]')  # 業務実施区分　沖縄　リスト92番目

select = Select(dropdown)

select.select_by_index(92)#沖縄事業所をドロップダウンで選択→変更してよいかと言われるのでアラート解除を下に追加。

handle_array = driver.window_handles

# print("別ページに切り替えた後のhandle_arrayの表示配列最初と次")#windowshandleは2つ結局かわらす。
# print(handle_array[0])
# print(handle_array[1])
driver.find_element(By.XPATH,'/html/body').click()#エンターを押して、次メニューに更新

driver.switch_to.window(handle_array[1])

alert = driver.switch_to.alert
print(alert.text)
alert.accept()

driver.find_element(By.XPATH,'/html/body').click()#エンターを押して、次メニューに更新

driver.switch_to.frame(1)#iFrameの最初に切り替え。２つあるが、2番目（1）のiFrameに切り替える。

monthday = str(calendar.monthrange(2024,5)[1])

# print(ym)

print(monthday)

# monthday = calendar.monthrange(str(ym[:4]),str(ym[4:]))[1]

ymd_s = str(ym[:4])+'/'+str(ym[4:])+'/'+'01'

ymd_l = str(ym[:4])+'/'+str(ym[4:])+'/'+monthday

print(ymd_s)

print(ymd_l)

form = driver.find_element(By.XPATH,'//*[@id="BaseDateText"]')

form.clear()

form.send_keys(ymd_s)


driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)

form = driver.find_element(By.XPATH,'//*[@id="ProgramDeptText"]')

form.clear()

time.sleep(3)

driver.find_element(By.XPATH,'//*[@id="ProgramDeptText"]').send_keys('557030')#実施担当部門沖縄　557030の入力。初期値は発注の福岡553010になっているため
driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)

time.sleep(3)


# form = driver.find_element(By.XPATH,'//*[@id="ProgramBusinessCodeText"]')
#
# form.send_keys('2006101343')#番組業務番号の入力
#
# driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)


Project1_No = 'M3P0000095-0H'#要員費

Project2_No = 'M3P0000095-0I'#諸経費

taishou_mon = 202405

form = driver.find_element(By.XPATH,'//*[@id="ProjCodeText1"]')

# print(form)

form.send_keys(Project1_No)#Project1 要員費番号を入力

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)

time.sleep(2)

form = driver.find_element(By.XPATH,'//*[@id="ProjCodeText2"]')

# print(form)

form.send_keys(Project2_No)#物品費番号を入力


driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。
#
time.sleep(2)

# os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。


time.sleep(3)

#driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。
driver.find_element(By.XPATH,'/html/body').click()#エンターを押して、次メニューに更新。



driver.implicitly_wait(2)

#dropdown3 = driver.find_element(By.CSS_SELECTOR, '#AcceptFormDtl1List')  # 受注形態詳細１　沖縄事業所入力　リスト22番目
dropdown3 = driver.find_element(By.XPATH,'//*[@id="AcceptFormDtl1List"]')
select = Select(dropdown3)

select.select_by_index(22)#沖縄事業所をドロップダウンで選択→プロジェクト入力で自動入力されるのでそのまま

driver.find_element(By.XPATH,'/html/body').click()#エンターを押して、次メニューに更新。

#dropdown4 = driver.find_element(By.CSS_SELECTOR, '#AcceptFormDtl2List')  # N/A入力
dropdown4 = driver.find_element(By.XPATH,'//*[@id="AcceptFormDtl2List"]')
select = Select(dropdown4)

select.select_by_index(1)

driver.find_element(By.XPATH,'/html/body').click()#エンターを押して、次メニューに更新。

dropdown = driver.find_element(By.XPATH, '//*[@id="StationInoutTypeList"]')  # 福岡　NHK局内選択

select = Select(dropdown)

select.select_by_index(13)







driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して

time.sleep(3)

dropdown = driver.find_element(By.XPATH, '//*[@id="ResourceList"]')  # リソース未定なし

select = Select(dropdown)

select.select_by_index(1)


driver.find_element(By.XPATH,'/html/body').click()#エンターを押して、次メニューに更新。

# os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。


driver.find_element(By.XPATH,'//*[@id="TabItem2"]/span').click()
#driver.find_element(By.CSS_SELECTOR,'#TabItem2 > span').click()

time.sleep(3)

handle_array = driver.window_handles

print("別ページに切り替えた後のhandle_arrayの表示配列最初と次")#windowshandleは2つ結局かわらす。
print(handle_array[0])
print(handle_array[1])


driver.switch_to.window(handle_array[1])

driver.switch_to.frame(1)#iFrameの最初に切り替え。２つあるが、2番目（1）のiFrameに切り替える。

driver.find_element(By.XPATH,'//*[@id="SubtitleText"]').send_keys(ym[:4]+'年'+ym[4:]+'月 '+eplyName)#副題


driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

time.sleep(2)

dropdown =driver.find_element(By.XPATH,'//*[@id="InputTable"]/tbody/tr/td[1]')#作業詳細

select = Select(dropdown)

select.select_by_index(-1)

time.sleep(2)




driver.find_element(By.XPATH,'//*[@id="InpSttTimeText"]').send_keys('00:00')#開始日時　時

time.sleep(1)

driver.find_element(By.XPATH,'//*[@id="InpEndDateText"]').send_keys(ymd_l)#終了日時　日


driver.find_element(By.XPATH,'//*[@id="InpEndTimeText"]').send_keys('00:00')#終了日時　


driver.find_element(By.XPATH,'//*[@id="InpOpeDtlCntText1"]').send_keys('1')#担当　担当者数　１を入力


driver.find_element(By.XPATH,'//*[@id="InpSttDateText"]').send_keys(ymd_s)#開始日時　日

time.sleep(3)

driver.find_element(By.XPATH,'/html/body').send_keys(Keys.ENTER)#エンターを押して、次メニューに更新。

# driver.find_element(By.XPATH,'//*[@id="RegistButton"]/span').click()#登録ボタン

driver.find_element(By.CSS_SELECTOR,'#RegistButton > span').click()


# os.kill(driver.service.process.pid,signal.SIGTERM)#ブラウザが閉じるのを止める。開きっぱなしにする。

driver.find_element(By.XPATH,'//*[@id="UpdateButton"]/span').click()#更新ボタン


handle_array = driver.window_handles

# print("別ページに切り替えた後のhandle_arrayの表示配列最初と次")#windowshandleは2つ結局かわらす。
# print(handle_array[0])
# print(handle_array[1])

driver.switch_to.window(handle_array[1])

alert = driver.switch_to.alert
print(alert.text)
alert.accept()



driver.switch_to.frame(1)#iFrameの最初に切り替え。２つあるが、2番目（1）のiFrameに切り替える。
anken_No = driver.find_element(By.XPATH,'//*[@id="ProposalNoText"]').get_attribute("value")

print(anken_No)
#
# //*[@id="ProposalNoText"]#案件番号のXPATH　　この中のvalueが案件番号

anken_data = [{'name':str(ym)+eplyName,'anken_No':str(anken_No)}]
#
# anken_data.append(ym+eplyName,anken_No)
#
print(anken_data)

sg.popup_ok(str(ym)+eplyName+'の案件番号'+anken_No,title = '案件番号')
#
# with open ('anken_data_list','w') as file:
#     for item in anken_data:
#         line = f'{item['name']:<10}}{item['anken_No']:010d}\n
#         file.write(line)

# t
