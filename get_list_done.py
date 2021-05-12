import time
import sys
from selenium import webdriver

reload(sys)
sys.setdefaultencoding('utf-8')

driver = webdriver.Chrome('./bin/chromedriver')  # Optional argument, if not specified will search path.
driver.get('https://oa.99cloud.com.cn/');

login_username = driver.find_element_by_id('login_username')
login_username.send_keys('')

login_password = driver.find_element_by_id('login_password')
login_password.send_keys('')

login_button = driver.find_element_by_id('login_button')
login_button.submit()

js="document.getElementsByClassName('lev2 lev')[1].style.display='block'"
driver.execute_script(js)
driver.find_element_by_id('F01_listDone').click()

iframe = driver.find_element_by_id('mainIframe')
driver.switch_to_frame(iframe)

rows = driver.find_elements_by_css_selector('#list tr')
print len(rows)

time.sleep(3) # Let the user actually see something!
driver.quit()
