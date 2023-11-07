
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


# options = webdriver.ChromeOptions()


# print("========== kekka ========== 動いてる？")


# service = Service(driver_path)
# service = Service(executable_path=driver_path)# 2) executable_pathを指定
# driver = webdriver.Chrome(service=service)# 3) serviceを渡す

# 起動時にオプションをつける。（ポート指定により、起動済みのブラウザのドライバーを取得）
# driver_path = "C:\\Users\\406239\\AppData\\Local\\Programs\\Python\\Python39\\chromedriver_binary\\chromedriver.exe"

# driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"
# driver = webdriver.Chrome(service=ChromeService(driver_path))
# options.add_experimental_option("debuggerAddress", "127.0.0.1:9333")
# driver = webdriver.Chrome(executable_path=driver_path, options=options)



# driver = webdriver.Chrome(service,options=options)


# ページのタイトルを表示する
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.chrome.service import Service
# # from selenium.webdriver.chrome.options import Options


# selenium 4

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
#import os
import time

from selenium.webdriver.common.by import By

# from selenium.webdriver.common.by import className


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


print("========== G-smartを開く ========== ")


# service = Service(driver_path)
# service = Service(executable_path=driver_path)# 2) executable_pathを指定
# driver = webdriver.Chrome(service=service)# 3) serviceを渡す

# 起動時にオプションをつける。（ポート指定により、起動済みのブラウザのドライバーを取得）
# driver_path = "C:\\Users\\406239\\AppData\\Local\\Programs\\Python\\Python39\\chromedriver_binary\\chromedriver.exe"

driver_path = "C:\\Users\\406239\\PycharmProjects\\pythonProject1\\chromedriver_binary\\chromedriver.exe"
driver = webdriver.Chrome(service=ChromeService(driver_path))
# options.add_experimental_option("debuggerAddress", "127.0.0.1:9333")
# driver = webdriver.Chrome(executable_path=driver_path, options=options)



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


# <h3 class="boxTitle4" style="background: url(&quot;./img/plus.png&quot;) right center no-repeat;"><img src="/fw/dfw/dmn/img/somu.png">&nbsp;総務・経理・人材育成・部局システム</h3>

# frame_1 = driver.find_element(By.CLASS_NAME,"accordionBox")

frame_1 = driver.find_element(By.CSS_SELECTOR,"#backendmenu > div:nth-child(5) > h3")

# frame_1 = driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/h3')

# frame_1 = driver.find_element(By.CSS_SELECTOR,"#backendmenu > div:nth-child(5) > p > a:nth-child(1) > img")

# driver.switch_to.frame(frame_1)

print('CSS Selector')

time.sleep(1)

frame_1.click()

time.sleep(2)

frame_2 = driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/p/a[1]/img')

driver.switch_to.frame(frame_2)

form = driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/p/a[1]/img')

form.click()

driver.switch_to.default_content()

form = driver.find_element(By.XPATH,'//*[@id="collapse5"]/div/a[5]')

#time.sleep(10)

form.click()








