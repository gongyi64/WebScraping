

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

# handle_array = driver.window_handles
#
#
#
# print("handle_arrayの表示配列最初と次")
# print(handle_array[0])
# print(handle_array[1])
#
# driver.switch_to.window(handle_array[1])
#
# time.sleep(10)


# <h3 class="boxTitle4" style="background: url(&quot;./img/plus.png&quot;) right center no-repeat;"><img src="/fw/dfw/dmn/img/somu.png">&nbsp;総務・経理・人材育成・部局システム</h3>





driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/h3').click()  #総務経理関係のメインメニュークリック

# driver.implicitly_wait(30)
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/p/a[1]/img').click()

time.sleep(3)

# driver.switch_to.default_content


print('G-smart押した？')

driver.implicitly_wait(30)



# time.sleep(2)

# frame_2 = driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/p/a[1]/img')

# driver.switch_to.frame(frame_2)

# form = driver.find_element(By.XPATH,'//*[@id="backendmenu"]/div[4]/p/a[1]/img')
#
# form.click()
#
# driver.switch_to.default_content()

# form = driver.find_element(By.XPATH,'//*[@id="collapse5"]/div/a[5]')
#
# #time.sleep(10)
#
# form.click()








