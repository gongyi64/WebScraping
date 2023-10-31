
from selenium import webdriver
import time

from selenium.webdriver.common.by import By

from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.action_chains import ActionChains

driver = webdriver.Chrome()

#Googleのブラウザを開く
# driver.get('https://nhktech.sharepoint.com/sites/Portal001/SitePages/top_v2.aspx/')

driver.get('https://gate.isso.nhk.or.jp/fw/dfw/dmn/dm/dmenu/')
time.sleep(2)

driver.maximize_window()

time.sleep(2)

form = driver.find_element(By.XPATH,'//*[@id="kigyou_code"]')

form.send_keys('1010')

form.send_keys(Keys.TAB)

form =driver.find_element(By.XPATH,'//*[@id="syokuin_code2"]')

form.send_keys('406239')


time.sleep(2)

form.send_keys(Keys.TAB)

time.sleep(4)

# target_element = driver.find_element(By.XPATH,'//*[@id="syokuin_code2"]')
#
# actions.click(target_element)

# actions.move_to_element(target_element)

# actions.perform()

# form =driver.find_element(By.XPATH,'//*[@id="syokuin_code2"]')
# #
# form.send_keys('406239').
#
# target_element = driver.find_element(By.XPATH,'//*[@id="password2"]')

# actions = ActionChains(driver)
#
# actions.move_to_element(target_element).click()

# actions.perform()
#
form =driver.find_element(By.XPATH,'//*[@id="password2"]')
#
form.send_keys('*kabu92772462')
#
# driver.find_element(By.XPATH,'//*[@id="main"]/table/tbody/tr[5]/td[2]/input').submit()

time.sleep(5)

# driver.quit()
