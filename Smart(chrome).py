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


driver.find_element(By.NAME,'eplyNo').send_keys('406239')

driver.find_element(By.NAME,'dispYmY').clear()#予め値が入っているので一旦クリア

driver.find_element(By.NAME,'dispYmY').send_keys('2023')

driver.find_element(By.NAME,'dispYmM').clear()#予め値が入っているので一旦クリア

driver.find_element(By.NAME,'dispYmM').send_keys('11')

# driver.find_element(By.XPATH,'/html/body/form/table[2]/tbody/tr[3]/td[2]/input[1]').click()#一般・専門職押す

driver.find_element(By.XPATH,'/html/body/form/table[2]/tbody/tr[3]/td[2]/input[2]').click()#経営職押す

time.sleep(2)

driver.find_element(By.XPATH,'/html/body/form/table[4]/tbody/tr[2]/td[1]/input').click()#表示ボタン押す

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

for i in range(1,len(trs)):
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

#結合セルによるデータずれの補正。ひとつひとつやるしかなさそう。

df_org.loc[0].replace('事  由','実  績',inplace = True)

df_org.loc[0].replace('備  考','事  由',inplace = True)

df_org.loc[0].replace('深　夜\n時間数','備  考',inplace = True)

df_org.loc[0].replace('休日','深　夜\n時間数',inplace = True)

df_org.loc[0].replace('宿泊','休日',inplace = True)

df_org.loc[0].replace('日帰','休日',inplace = True)

df_org.loc[0].replace('緊\n急','休日',inplace = True)

df_org.loc[0].replace('休\n張','宿泊',inplace = True)

df_org.loc[0].iat[12] = '宿泊'

df_org.loc[0].iat[13] = '宿泊'

df_org.loc[0].iat[14] = '日帰'

df_org.loc[0].iat[15] = '日帰'

df_org.loc[0].iat[16] = '緊\n急'

df_org.loc[0].iat[17] = '休\n張'


print('インサート後？')

print(df_org.loc[0])

print('次の行')

print(df_org.loc[1])

print(df_org.loc[1].shift(2))


#0行目をcolumnにするコード　（以下）
df_org.columns = df_org.iloc[0]

df_org = df_org.drop(df_org.index[0])

df_org.reset_index(drop=True,inplace=True)


print('data 加工後')

print(df_org.columns.tolist())

# columns_smart = df_org.columns.tolist()
#


print(df_org)


df_org.to_excel('smart_kinmu.xlsx') #smartの勤務表











# df[0].replace('\n','',regex = True)


























