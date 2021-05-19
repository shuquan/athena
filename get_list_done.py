import argparse
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def login():
    parser = argparse.ArgumentParser()
    parser.add_argument('url')
    parser.add_argument('username')
    parser.add_argument('password')
    args = parser.parse_args()

    driver = webdriver.Chrome('./bin/chromedriver')
    driver.get(args.url);

    login_username = driver.find_element_by_id('login_username')
    login_username.send_keys(args.username)

    login_password = driver.find_element_by_id('login_password')
    login_password.send_keys(args.password)

    login_button = driver.find_element_by_id('login_button')
    login_button.submit()

    # Move to List Done
    js="document.getElementsByClassName('lev2 lev')[1].style.display='block'"
    driver.execute_script(js)
    driver.find_element_by_id('F01_listDone').click()

    iframe = driver.find_element_by_id('mainIframe')
    driver.switch_to.frame(iframe)

    return driver

def travels_handler(total_travels, record, driver):
    print(driver.title)
    total_travels.append(record)

def procurement_handler(total_procurement, record, driver):
    print(driver.title)
    total_procurement.append(record)

def contracts_handler(total_contracts, record, driver):
    print(driver.title)
    total_contracts.append(record)

def weekly_reports_handler(total_weekly_reports, record, driver):
    print(driver.title)
    iframe = driver.find_element_by_id('zwIframe')
    driver.switch_to.frame(iframe)
    record_detail = {}
    rows = driver.find_elements_by_css_selector('.is-detailshover')
    for row in rows:
        tds = row.find_elements_by_css_selector('section [class*="browse"]')
        record_detail['年份'] = tds[0].get_attribute('textContent')
        record_detail['月份'] = tds[1].get_attribute('textContent')
        record_detail['周'] = tds[2].get_attribute('textContent')
        record_detail['客户名称'] = tds[3].get_attribute('textContent')
        record_detail['大类'] = tds[4].get_attribute('textContent')
        record_detail['中类'] = tds[5].get_attribute('textContent')
        record_detail['小类'] = tds[6].get_attribute('textContent')
        record_detail['耗时'] = tds[7].get_attribute('textContent')
        record_detail['加班'] = tds[8].get_attribute('textContent')
        record_detail['具体工作描述'] = tds[9].get_attribute('textContent')

    print(record_detail)
    total_weekly_reports.append(record)

def others_handler(others, record, driver):
    print(driver.title)
    others.append(record)

def to_excel(total_records, total_travels, total_procurement, total_contracts, total_weekly_reports, others):
    df = pd.DataFrame(total_records)
    df.to_excel('总表.xlsx', sheet_name='总表')

    df = pd.DataFrame(total_travels)
    df.to_excel('出差单.xlsx', sheet_name='出差单')

    df = pd.DataFrame(total_procurement)
    df.to_excel('采购单.xlsx', sheet_name='采购单')

    df = pd.DataFrame(total_contracts)
    df.to_excel('合同单.xlsx', sheet_name='合同单')

    df = pd.DataFrame(total_weekly_reports)
    df.to_excel('周报.xlsx', sheet_name='周报')

    df = pd.DataFrame(others)
    df.to_excel('未归类.xlsx', sheet_name='未归类')

def main():
    driver = login()

    total_number = 0
    total_records = []
    total_travels = []
    total_procurement = []
    total_contracts = []
    total_weekly_reports = []
    others = []

    current_window_handle = driver.current_window_handle
    # Get each record from List Done and handle it.
    while True:
        rows = driver.find_elements_by_css_selector('#list tr')
        total_number = total_number + len(rows)

        # Just for testing limit to 20 records
        if total_number > 20:
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

            # Click each row and Open new tab to fetch more data
            webdriver.ActionChains(driver).move_to_element(row).click(row).perform()
            WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
            window_handles = driver.window_handles
            driver.switch_to.window(window_handles[1])
            if '出差单' in subject:
                travels_handler(total_travels, record, driver)
            elif '采购' in subject:
                procurement_handler(total_procurement, record, driver)
            elif '合同' in subject:
                contracts_handler(total_contracts, record, driver)
            elif '周报' in subject:
                weekly_reports_handler(total_weekly_reports, record, driver)
            else:
                others_handler(others, record, driver)

            driver.close()
            driver.switch_to.window(current_window_handle)
            # Move to List Done
            iframe = driver.find_element_by_id('mainIframe')
            driver.switch_to.frame(iframe)



        if len(rows) < 20:
            break

        page_next_button = driver.find_element_by_class_name('pageNext')
        page_next_button.click()
        time.sleep(1)

    driver.quit()

    to_excel(total_records, total_travels, total_procurement, total_contracts, total_weekly_reports, others)

    print(len(total_travels))
    print(len(total_procurement))
    print(len(total_contracts))
    print(len(total_weekly_reports))
    print(len(others))
    print(len(total_records))

if __name__ == '__main__':
    main()
