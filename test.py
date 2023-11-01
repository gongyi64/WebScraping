
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.chrome.service import Service
# # from selenium.webdriver.chrome.options import Options


# selenium 4

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
import os
import time

options = webdriver.ChromeOptions()

#from webdriver_manager.chrome import ChromeDriverManager

#driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

#最新のドライバーだとエラーが出るのでその対応策。https://qiita.com/hs2023/questions/ffab105c5692692624ab

# import requests
# res = requests.get('https://chromedriver.storage.googleapis.com/LATEST_RELEASE')
# driver = webdriver.Chrome(ChromeDriverManager(res.text).install())



# from os.path import join
#
# root = join(__file__, "..")

# webdriverオブジェクトを作る（ブラウザが開く）
# driver_path=join(root, "chromedriver.exe")

#driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"

driver_path = "C:\\Users\\406239\\AppData\\Local\\Programs\\Python\\Python39\\chromedriver_binary\\chromedriver.exe"

print("========== kekka ========== 動いてる？")


# service = Service(driver_path)
# service = Service(executable_path=driver_path)# 2) executable_pathを指定
# driver = webdriver.Chrome(service=service)# 3) serviceを渡す

# 起動時にオプションをつける。（ポート指定により、起動済みのブラウザのドライバーを取得）
# options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
# driver = webdriver.Chrome(executable_path=driver_path, options=options)

driver = webdriver.Chrome(service=ChromeService(driver_path))

# driver = webdriver.Chrome(service,options=options)


# ページのタイトルを表示する

driver.get("http://www.yahoo.co.jp/")
get_title = driver.title
print(get_title)
print("========== kekka ========== 動いてる？")
print(driver.page_source)

time.sleep(5)

driver.quit()


