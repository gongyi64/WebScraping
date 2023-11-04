
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.chrome.service import Service
# # from selenium.webdriver.chrome.options import Options


# selenium 4

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
#import os
import time

from selenium.webdriver.common.by import By


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


print("========== kekka ========== 動いてる？")


# service = Service(driver_path)
# service = Service(executable_path=driver_path)# 2) executable_pathを指定
# driver = webdriver.Chrome(service=service)# 3) serviceを渡す

# 起動時にオプションをつける。（ポート指定により、起動済みのブラウザのドライバーを取得）
# driver_path = "C:\\Users\\406239\\AppData\\Local\\Programs\\Python\\Python39\\chromedriver_binary\\chromedriver.exe"

driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"
driver = webdriver.Chrome(service=ChromeService(driver_path))
options.add_experimental_option("debuggerAddress", "127.0.0.1:9333")
# driver = webdriver.Chrome(executable_path=driver_path, options=options)



# driver = webdriver.Chrome(service,options=options)


# ページのタイトルを表示する

driver.get("https://gate.isso.nhk.or.jp/fw/dfw/lkteams/")

driver.maximize_window()

time.sleep(1)
form = driver.find_element(By.XPATH,'//*[@id="selector_form"]/div/div/div/button')

form.click()

time.sleep(1)

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

form = driver.find_element(By.XPATH,'//*[@id="collapse5"]/div/a[5]')

#time.sleep(10)

form.click()

time.sleep(10)

#form = driver.find_element(By.XPATH,'//*[@id="excelout"]')

form.click()

time.sleep(20)

form = driver.find_element(By.XPATH,'//*[@id="close"]/a')

form.click()

# get_title = driver.title
# print(get_title)
# print("========== kekka ========== 動いてる？")
# print(driver.page_source)


#
# driver.quit()


# driver.get('https://www.insource.co.jp/bup/middle-schedule.html')

# driver.maximize_window()
#
# time.sleep(2)
#
# form = driver.find_element(By.XPATH,'//*[@id="search"]/div/input')
#
# form.send_keys('Python')
#
# form.submit()