# # selenium 4


#
# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# selenium 4

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
#import os
import time

from selenium.webdriver.common.by import By

import pandas as pd


import sys

import PySimpleGUI as sg

# value = sg.popup_get_file('TEAMSで勤務表をダウンロードします。')#使用するダウンロード済みの勤務表元ファイルを選択


sg.theme('Python')

layout =[[sg.Text('[NT勤務自動ダウンロード]',font = ('Noto Serif CJK JP',14))],

        [sg.Text('[G-smartから勤務実績を取り込みます。] ',font = ('meiryo',10))],

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
mon = num[4:]#下2桁月数

# from pyvirtualdisplay import Display

from webdriver_manager.chrome import ChromeDriverManager

# from selenium.webdriver.common.by import className


# from webdriver_manager.chrome import ChromeDriverManager
#
# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

#最新のドライバーだとエラーが出るのでその対応策。https://qiita.com/hs2023/questions/ffab105c5692692624ab


options = webdriver.ChromeOptions()


print("========== G-smartを開く ========== ")


# service = Service(driver_path)
# service = Service(executable_path=driver_path)# 2) executable_pathを指定
# driver = webdriver.Chrome(service=service)# 3) serviceを渡す

# 起動時にオプションをつける。（ポート指定により、起動済みのブラウザのドライバーを取得）
# driver_path = "C:\\Users\\406239\\AppData\\Local\\Programs\\Python\\Python39\\chromedriver_binary\\chromedriver.exe"

driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"
driver = webdriver.Chrome(service=ChromeService(driver_path))




# G-smartのページを開いて、ログインする

driver.get("http://gate.isso.nhk.or.jp/fw/dfw/dmn/dm/dmenu")


driver.maximize_window()

time.sleep(1)
form = driver.find_element(By.XPATH,'//*[@id="kigyou_code"]')

form.send_keys('1010')

form = driver.find_element(By.XPATH,'//*[@id="syokuin_code2"]')

form.send_keys('406239')

form = driver.find_element(By.XPATH,'//*[@id="password2"]')

form.send_keys('*kabu92772462')

time.sleep(2)

form = driver.find_element(By.XPATH,'//*[@id="main"]/table/tbody/tr[5]/td[2]/input')

form.click()


print("========== G-smartにログイン完了・統合認証TOPページ ========== ") #この上まではOK。




driver.implicitly_wait(30)

print(driver.current_url)


driver.implicitly_wait(10)

#console_log = document.getElementsByClassName('accordionBox','boxTitle4')[0].style.display = 'background: url(&quot;./img/plus.png&quot;) right center no-repeat;'

# driver.get("http://gate.nhk.or.jp/fw/dfw/noccont/syuchi/Syuchi.asp")

frame_1 = driver.find_element(By.XPATH,"/html/frameset/frameset/frame[1]")

driver.switch_to.frame(frame_1)

stats = driver.find_elements(By.ID,"backendmenu")

print(stats)



print('総務系のメインメニュー押した？')

#driver.execute_script("document.getElementsByClassName('accordionBox','boxTitle4')[0].style.display = 'background: url(&quot;./img/plus.png&quot;) right center no-repeat;'")

# driver.execute_script("document.getElementsByClassName('boxTitle4')[0].style.display = 'block';")



time.sleep(2)




# <h3 class="boxTitle4" style="background: url(&quot;./img/plus.png&quot;) right center no-repeat;"><img src="/fw/dfw/dmn/img/somu.png">&nbsp;総務・経理・人材育成・部局システム</h3>





driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/h3').click()  #総務経理関係のメインメニュークリック

# driver.implicitly_wait(30)
time.sleep(3)

driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/p/a[1]/img').click()

time.sleep(3)

# driver.switch_to.default_content


print('G-smart押したあと')

driver.implicitly_wait(10)

handle_array = driver.window_handles

print("handle_arrayの表示配列最初と次toその次")
print(handle_array[0])
print(handle_array[1])

driver.switch_to.window(handle_array[1])

# frame_2 = driver.find_element(By.CSS_SELECTOR,'#_pprIFrame')
#
# driver.switch_to.frame(frame_2)








time.sleep(10)

handle_array = driver.window_handles

print("handle_arrayの表示配列最初と次toその次")
print(handle_array[0])
print(handle_array[1])

driver.switch_to.window(handle_array[1])


#管理者メニューのXpath

CSS_Select = driver.find_elements(By.CSS_SELECTOR,'#AppsNavLink')[4].click()#一番上の選択はできる
#4番目を押したいのでどうすれば？find_elementsとし、リストの要素を取得して解決






#<a id="AppsNavLink" href="http://nhksmrt.isso.nhk.or.jp/OA_HTML/OA.jsp?OAFunc=OAHOMEPAGE&amp;akRegionApplicationId=0&amp;navRespId=78937&amp;navRespAppId=20016&amp;navSecGrpId=0&amp;transactionid=373107157&amp;oapc=2&amp;oas=4qmIeywfLoeL4q8iPEG2ew.." class="xh">1D2C出退勤_101_セルフ_管理者</a>

#driver.execute_script('javascript:http://nhksmrt.isso.nhk.or.jp/OA_HTML/OA.jsp?OAFunc=OAHOMEPAGE&amp;akRegionApplicationId=0&amp;navRespId=78937&amp;navRespAppId=20016&amp;navSecGrpId=0&amp;transactionid=373107157&amp;oapc=2&amp;oas=4qmIeywfLoeL4q8iPEG2ew..')

#driver.find_element(By.XPATH,'//*[@id="AppsNavLink"] and [contains(text()="1D2C出退勤_101_セルフ_管理者")]').click()

# driver.find_element(By.CLASS_NAME('x4n')).click()

#driver.find_element(By.CLASS_NAME,'x4n' and contains(text()="1D2C出退勤_101_セルフ_管理者")).click()

time.sleep(5)

driver.find_element(By.ID,'N122').click()


time.sleep(5)


#氏名と社員番号の辞書を作成し、入力に使用する。0:専門職、一般職　1:経営職　2:スタッフ

eplyNo_name_dict = {406239:'白川　公一',380672:'三角　和浩',378035:'岩田　貴夫',805519:'仲本　祥子',880079:'砂川　航輝',806185:'辻　ひかる',880334:'山城　実咲',880518:'古波蔵　晃久'}#,360986:'山城　徳松'

eply_dict = {406239:'1',380672:'0',378035:'0',805519:'0',880079:'0',806185:'0',880334:'0',880518:'0'}#,360986:'2'


print(list(eplyNo_name_dict.keys()))

print('keyのデータ数')

print(len(list(eplyNo_name_dict.keys())))

#for分で全員分のデータを取り込み、Excelへ。作成中。/html/body/form/table[3]/tbody/tr[2]/td/input（名前のXpath）

# for i in range(len(list(eplyNo_name_dict.keys()))):

for i in range(0,8):#全員分（出向、スタッフ除く）の勤務取得ルーチン。

    input_eplyNo = list(eplyNo_name_dict.keys())[i]

    # input_eplyNo = list(eplyNo_name_dict.keys())[0]

    print (input_eplyNo)


    val_s = eply_dict[input_eplyNo]

    print(val_s)

    name = eplyNo_name_dict[input_eplyNo]

    print (name)
    print(input_eplyNo)

    driver.find_element(By.NAME, 'eplyNo').clear()  # 予め値が入っているので一旦クリア

    driver.find_element(By.NAME,'eplyNo').send_keys(input_eplyNo)

    driver.find_element(By.NAME,'dispYmY').clear()#予め値が入っているので一旦クリア

    driver.find_element(By.NAME,'dispYmY').send_keys(nen)

    driver.find_element(By.NAME,'dispYmM').clear()#予め値が入っているので一旦クリア

    driver.find_element(By.NAME,'dispYmM').send_keys(mon)

    if val_s == '0':

        driver.find_element(By.XPATH,'/html/body/form/table[2]/tbody/tr[3]/td[2]/input[1]').click()#一般・専門職押す

    if val_s == '1':

        driver.find_element(By.XPATH,'/html/body/form/table[2]/tbody/tr[3]/td[2]/input[2]').click()#経営職押す

    else:
        pass

# driver.find_element(By.XPATH,'/html/body/form/table[2]/tbody/tr[3]/td[2]/input[1]').click()#一般職・専門職押す

    time.sleep(2)

    driver.find_element(By.XPATH,'/html/body/form/table[4]/tbody/tr[2]/td[1]/input').click()#表示ボタン押す

#################################以上G-Smart画面へのログイン・表示########################

#################################勤務実績の取り込み########################



# handle_array = driver.window_handles
#
# print("handle_arrayの表示配列最初と次toその次")
# print(handle_array[0])
# print(handle_array[1])#window handleは、同じでした。
#
# driver.switch_to.window(handle_array[1])

    element = driver.find_element(By.XPATH,'/html/body/table[2]',)

    trs = element.find_elements(By.TAG_NAME,'tr')

# print(trs)

    time.sleep(10)
# df_kinmu = pd.DataFrame()


    df = []

    for i in range(0,len(trs)):#経営職では、1からにしていた。後日要チェック
        tds = trs[i].find_elements(By.TAG_NAME, "td")

        line = ""


        for j in range(0,len(tds)):
            if j < len(tds)-1:
                line += "%s\t" % (tds[j].text)#"\t";tab
            else:
                line += "%s" % (tds[j].text)

        print(line+"\r\n")
    # line.split('\t')
        df.append(line)#df=df.append~と書いたら値はNone！

    print(df)

    df_kinmu = pd.Series(df)

    print('Seriesです')

    print(df_kinmu)

# print(df_kinmu.columns.tolist())#Seriesにtolist()はだめ。

    print(type(df_kinmu))#Series

# ''.join(df_kinmu.splitlines())#Seriesではだめ

    df_org = df_kinmu.str.split('\t',expand = True)

# ''.join(df_org[0].splitlines())#Seriesではだめ\nを消したい

    print(df_org)

    print(type(df_org))#Dataframe

    print('一番上の行')

    print(df_org.loc[0])

    if val_s == '0':
#一般職のデータ整列用に一旦コメントアウトしている以下

#一般・専門職の勤務票処理=================================

#結合セルによるデータずれの補正(1行目）。2行目と統合させ1行に変更。ひとつひとつやるしかなさそう。

        df_org.loc[0].replace('所 定','所定出退勤',inplace = True)

        df_org.loc[0].replace('実  績','出勤-退勤',inplace = True)

        df_org.loc[0].replace('時間外','仮眠',inplace = True)

        df_org.loc[0].replace('前日時間外\n時間外合計','時間外',inplace = True)

        df_org.loc[0].replace('深夜','前日時間外\n時間外合計',inplace = True)

        df_org.loc[0].replace('休 日','深夜',inplace = True)

        df_org.loc[0].replace('事 由','時間外',inplace = True)

        df_org.loc[0].iat[9] = ''

        df_org.loc[0].iat[10] = '休日A'

        df_org.loc[0].iat[11] = '休日B'

        df_org.loc[0].iat[12] = '事　由'

        df_org.loc[0].iat[13] = '備　考'

        df_org.loc[0].iat[14] = '日帰100km'

        df_org.loc[0].iat[15] = '日帰８H'

        df_org.loc[0].iat[16] = '宿泊A'

        df_org.loc[0].iat[17] = '宿泊B'

        df_org.loc[0].iat[18] = '宿泊c'

        df_org.loc[0].iat[19] = '早朝'

        df_org.loc[0].iat[20] = '休張'

        df_org.loc[0].iat[21] = '休日A'

        df_org.loc[0].iat[22] = '休日B'

        df_org.loc[0].iat[23] = '休日C'

        df_org.loc[0].iat[24] = '緊急'

        df_org.loc[0].iat[25] = '航空'

        df_org.loc[0].iat[26] = '水中'

        df_org.loc[0].iat[27] = '海上'

        df_org.loc[0].iat[28] = '高所'

        df_org.loc[0].iat[29] = '高圧'

        df_org.loc[0].iat[30] = '高地'

        df_org.loc[0].iat[31] = '地下'

        df_org.loc[0].iat[32] = '石綿'

        df_org.loc[0].iat[33] = '放射線'







        print('インサート後？')

        print(df_org.loc[0])

        print('次の行')

        print(df_org.loc[1])

# df_org.loc[1].shift(3)
#


#0行目をcolumnにするコード　（以下）
        df_org.columns = df_org.iloc[0]

        df_org = df_org.drop(df_org.index[0])

        df_org.reset_index(drop=True,inplace=True)

        df_org = df_org.drop([0])#1行目（カラムにした行）削除


        print('data 加工後')

        print(df_org.columns.tolist())


        print(df_org)

        with pd.ExcelWriter('c:/Users/406239/OneDrive - (株)NHKテクノロジーズ/デスクトップ/★勤務確認などのダウンロードデータ★/検証中/{}_smart_kinmu.xlsx'.format(num), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_org.to_excel(writer, sheet_name='{}_{}'.format(num, name))
        # df_org.to_excel('c:/Users/406239/OneDrive - (株)NHKテクノロジーズ/デスクトップ/★勤務確認などのダウンロードデータ★/検証中/{}_smart_kinmu.xlsx'.format(num),sheet_name='{}_{}'.format(num,name))

        print("handle_arrayの表示配列最初と次toその次")
        print(handle_array[0])
        print(handle_array[1])

        driver.switch_to.window(handle_array[1])

        driver.find_element(By.XPATH, '/html/body/table[4]/tbody/tr[2]/td[4]/input').click()  # 終了押す

    elif val_s == '1':
    # 一般・専門職の勤務票処理=================================

    # 経営職職の勤務票処理=================================

    # 結合セルによるデータずれの補正(1行目）。2行目と統合させ1行に変更。ひとつひとつやるしかなさそう。

        df_org.loc[0].replace('実  績', '出勤-退勤', inplace=True)

        df_org.loc[0].replace('事  由', '睡眠等', inplace=True)

        df_org.loc[0].replace('備  考', '事  由', inplace=True)

        df_org.loc[0].replace('深　夜\n時間数', '備  考', inplace=True)

        df_org.loc[0].replace('休日', '深　夜\n時間数', inplace=True)

        df_org.loc[0].replace('宿泊', '休日A', inplace=True)

        df_org.loc[0].replace('日帰', '休日B', inplace=True)

        df_org.loc[0].replace('緊\n急', '休日C', inplace=True)

        df_org.loc[0].replace('休\n張', '宿泊A', inplace=True)

        df_org.loc[0].iat[12] = '宿泊B'

        df_org.loc[0].iat[13] = '宿泊C'

        df_org.loc[0].iat[14] = '日帰100km'

        df_org.loc[0].iat[15] = '日帰８H'

        df_org.loc[0].iat[16] = '緊\n急'

        df_org.loc[0].iat[17] = '休\n張'

        print('インサート後？')

        print(df_org.loc[0])

        print('次の行')

        print(df_org.loc[1])

    # df_org.loc[1].shift(3)
    #

    # 0行目をcolumnにするコード　（以下）
        df_org.columns = df_org.iloc[0]

        df_org = df_org.drop(df_org.index[0])

        df_org.reset_index(drop=True, inplace=True)

        df_org = df_org.drop([0,1])  # 2行名削除

        print('data 加工後')

        print(df_org.columns.tolist())

        print(df_org)
    #
    # #  経営職の勤務票処理=================================

        with pd.ExcelWriter('c:/Users/406239/OneDrive - (株)NHKテクノロジーズ/デスクトップ/★勤務確認などのダウンロードデータ★/検証中/{}_smart_kinmu.xlsx'.format(num),engine='openpyxl', mode ='a',if_sheet_exists ='replace') as writer:
            df_org.to_excel(writer, sheet_name='{}_{}'.format(num,name))
    # df_org.to_excel('c:/Users/406239/OneDrive - (株)NHKテクノロジーズ/デスクトップ/★勤務確認などのダウンロードデータ★/検証中/{}_smart_kinmu.xlsx'.format(num),sheet_name='{}_{}'.format(num,name))

        print("handle_arrayの表示配列最初と次toその次")
        print(handle_array[0])
        print(handle_array[1])

        driver.switch_to.window(handle_array[1])

        driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td[4]/input').click()  # 終了押す


    else:
        pass































