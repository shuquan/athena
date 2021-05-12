import time
import sys
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

iframe = driver.find_element_by_id('mainIframe')
driver.switch_to.frame(iframe)

rows = driver.find_elements_by_css_selector('#list tr')
print(len(rows))
for row in rows:
    tds = row.find_elements_by_tag_name('td')
    subject = tds[1].find_element_by_tag_name('div').get_attribute('title')
    start_member_name = tds[2].find_element_by_tag_name('div').get_attribute('title')
    pre_approver_name = tds[3].find_element_by_tag_name('div').get_attribute('title')
    start_date = tds[4].find_element_by_tag_name('div').get_attribute('title')
    deal_time = tds[5].find_element_by_tag_name('div').get_attribute('title')
    current_nodes_info = tds[6].find_element_by_tag_name('div').get_attribute('title')
    str = "[%s],[%s],[%s],[%s],[%s],[%s]" %(subject,start_member_name, pre_approver_name,start_date,deal_time,current_nodes_info)
    print(str)

page_next_button = driver.find_element_by_class_name('pageNext')
page_next_button.click()

print('------------next page------------')

rows = driver.find_elements_by_css_selector('#list tr')
print(len(rows))

driver.quit()
