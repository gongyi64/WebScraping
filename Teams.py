
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.chrome.service import Service
# # from selenium.webdriver.chrome.options import Options


# selenium 4

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
#import os
import time

from selenium.webdriver.support.select import Select

from selenium.webdriver.support.ui import Select

from selenium.webdriver.common.by import By

# from selenium.webdriver.common.action_chains import ActionChains
#
# print("No.1:ActionChain import OK")

time.sleep(3)
#from webdriver_manager.chrome import ChromeDriverManager

#driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

#最新のドライバーだとエラーが出るのでその対応策。https://qiita.com/hs2023/questions/ffab105c5692692624ab

# import requests
# res = requests.get('https://chromedriver.storage.googleapis.com/LATEST_RELEASE')
# driver = webdriver.Chrome(ChromeDriverManager(res.text).install())



# from os.path import join
# #
# root = join(__file__, "..")

# webdriverオブジェクトを作る（ブラウザが開く）
# driver_path = join(root, "chromedriver.exe")


options = webdriver.ChromeOptions()


print("========== 機材Teams　ログイン中========== ")


# service = Service(driver_path)
# service = Service(executable_path=driver_path)# 2) executable_pathを指定
# driver = webdriver.Chrome(service=service)# 3) serviceを渡す

# 起動時にオプションをつける。（ポート指定により、起動済みのブラウザのドライバーを取得）
# driver_path = "C:\\Users\\406239\\AppData\\Local\\Programs\\Python\\Python39\\chromedriver_binary\\chromedriver.exe"

driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"

#2023_11_09 chromdriver　118→119　更新

driver = webdriver.Chrome(service=ChromeService(driver_path))

options.add_experimental_option("debuggerAddress", "127.0.0.1:9333")
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
#
# select = Select(dropdown)
#
time.sleep(1)

# from selenium.webdriver.support.ui import WebDriverWait

driver.find_element(By.XPATH,'//*[@id="collapse5"]/div/a[5]').click()
print(driver.page_source)

print(driver.current_url)

# driver.get("https://gate.isso.nhk.or.jp/fw/dfw/lkteams/schedule/rs_print_monschedule_t1.php?action=init")

driver.implicitly_wait(10)

handle_array = driver.window_handles

print("handle_arrayの表示配列最初と次")
print(handle_array[0])
print(handle_array[1])
# print(handle_array[2])

driver.switch_to.window(handle_array[1])

time.sleep(10)
# driver.find_element(By.XPATH,'//*[@id="excelout"]"]"]').click()

# driver.find_element(By.XPATH,'//*[@id="close"]/a').click()

driver.find_element(By.CSS_SELECTOR,'#excelout').click()　#この画面で出力する月を選択させる。

# script = 'javascript:void(0);'
# form.driver.execute_script(script)

print("出力押した")

time.sleep(100)

driver.implicitly_wait(100) #ダウンロードフォルダへ格納　これを別フォルダへ移動させる。

