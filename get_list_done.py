import time
from selenium import webdriver

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

time.sleep(5) # Let the user actually see something!
driver.quit()
