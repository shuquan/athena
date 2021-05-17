import time
import pandas as pd
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

total_number = 0
total_records = []
total_travels = []
total_procurement = []
total_contracts = []
total_weekly_reports = []
others = []
while True:
    rows = driver.find_elements_by_css_selector('#list tr')
    total_number = total_number + len(rows)

    # Just for testing limit to 100 records
    if total_number > 100:
        break

    for row in rows:
        tds = row.find_elements_by_tag_name('td')
        subject = tds[1].find_element_by_tag_name('div').get_attribute('title')
        start_member_name = tds[2].find_element_by_tag_name('div').get_attribute('title')
        pre_approver_name = tds[3].find_element_by_tag_name('div').get_attribute('title')
        start_date = tds[4].find_element_by_tag_name('div').get_attribute('title')
        deal_time = tds[5].find_element_by_tag_name('div').get_attribute('title')
        current_nodes_info = tds[6].find_element_by_tag_name('div').get_attribute('title')
        record = {'标题': subject, '发起人':start_member_name, '上一处理人':pre_approver_name, '发起时间':start_date, '处理时间':deal_time, '当前代办人':current_nodes_info}
        total_records.append(record)

        if '出差单' in subject:
            total_travels.append(record)
        elif '采购' in subject:
            total_procurement.append(record)
        elif '合同' in subject:
            total_contracts.append(record)
        elif '周报' in subject:
            total_weekly_reports.append(record)


    if len(rows) < 20:
        break

    page_next_button = driver.find_element_by_class_name('pageNext')
    page_next_button.click()
    time.sleep(2)

df = pd.DataFrame(total_records)
df.to_excel('总表.xlsx', sheet_name='总表')

df = pd.DataFrame(total_travels)
df.to_excel('出差单.xlsx', sheet_name='出差单')

driver.quit()
