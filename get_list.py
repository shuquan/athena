import argparse
import logging
import time
import pandas as pd
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException

from athena.objects import objects

parser = argparse.ArgumentParser()
parser.add_argument('url')
parser.add_argument('username')
parser.add_argument('password')
parser.add_argument('--list')
parser.add_argument('--week')
parser.add_argument('--num', nargs='?', const=1, type=int, default=40)
args = parser.parse_args()

def logger():
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    DATE_FORMAT = "%m/%d/%Y %H:%M:%S %p"
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, datefmt=DATE_FORMAT)


def login():
    chrome_options=Options()
    chrome_options.add_argument('--headless')
    driver = webdriver.Chrome('./bin/chromedriver', options=chrome_options)
    driver.get(args.url);

    login_username = driver.find_element_by_id('login_username')
    login_username.send_keys(args.username)

    login_password = driver.find_element_by_id('login_password')
    login_password.send_keys(args.password)

    login_button = driver.find_element_by_id('login_button')
    login_button.submit()

    js="document.getElementsByClassName('lev2 lev')[1].style.display='block'"
    driver.execute_script(js)

    if args.list == '待办事项':
        # Move to List Pending
        driver.find_element_by_id('F01_listPending').click()
    else:
        # Move to List Done
        driver.find_element_by_id('F01_listDone').click()

    iframe = driver.find_element_by_id('mainIframe')
    driver.switch_to.frame(iframe)

    return driver

def travels_handler(total_travels, record, driver):
    logging.info('[出差]%s' % driver.title)
    total_travels.append(record)

def procurement_handler(total_procurement, record, driver):
    logging.info('[采购]%s' % driver.title)
    total_procurement.append(record)

def contracts_handler(total_contracts, record, driver):
    logging.info('[合同]%s' % driver.title)
    iframe = driver.find_element_by_id('zwIframe')
    driver.switch_to.frame(iframe)
    while driver.execute_script("return document.readyState") != "complete" :
        time.sleep(1)
        logging.info('[合同]%s：等待iframe加载' % driver.title)
    rows = driver.find_elements_by_css_selector('section [class*="browse"]')
    while len(rows) != 52:
        time.sleep(1)
        logging.info('[合同]%s：需要加载52条，已经加载%s条，等待' % (driver.title,len(rows)))
        rows = driver.find_elements_by_css_selector('section [class*="browse"]')
    logging.info('[合同]%s：需要加载52条，已经加载%s条' % (driver.title,len(rows)))
    record['合同编号'] = rows[2].get_attribute('textContent')
    record['申请人'] = rows[3].get_attribute('textContent')
    record['申请部门'] = rows[4].get_attribute('textContent')
    record['申请日期'] = rows[5].get_attribute('textContent')
    record['客户单位名称'] = rows[6].get_attribute('textContent')
    record['履行地点'] = rows[7].get_attribute('textContent')
    record['区域'] = rows[10].get_attribute('textContent')
    record['合同名称'] = rows[16].get_attribute('textContent')
    record['合同价'] = rows[17].get_attribute('textContent')
    record['硬件金额'] = rows[19].get_attribute('textContent')
    record['成本'] = rows[20].get_attribute('textContent')
    record['合同内容'] = rows[21].get_attribute('textContent')
    record['服务期限'] = [rows[22].get_attribute('textContent'),rows[23].get_attribute('textContent')]
    record['收款方式'] = rows[24].get_attribute('textContent')
    record['备注'] = rows[25].get_attribute('textContent')
    record['项目编号'] = rows[27].get_attribute('textContent')
    record['项目名称'] = rows[28].get_attribute('textContent')
    record['软件许可数'] = rows[29].get_attribute('textContent')
    record['软件许可数'] = rows[29].get_attribute('textContent')
    record['纯服务合同'] = rows[30].get_attribute('textContent')
    record['运维服务'] = rows[31].get_attribute('textContent')
    record['产品'] = rows[32].get_attribute('textContent')
    record['产品分类'] = [rows[33].get_attribute('textContent'), rows[34].get_attribute('textContent')]
    record['业务来源'] = rows[35].get_attribute('textContent')
    record['业务需求'] = rows[36].get_attribute('textContent')
    record['行业类型'] = rows[37].get_attribute('textContent')
    record['所属大区'] = rows[38].get_attribute('textContent')
    record['签署收款金额'] = rows[39].get_attribute('textContent')
    record['到货收款金额'] = rows[40].get_attribute('textContent')
    record['初验收款金额'] = rows[41].get_attribute('textContent')
    record['终验收款金额'] = rows[42].get_attribute('textContent')
    record['保内尾款金额'] = rows[43].get_attribute('textContent')
    record['其他'] = rows[44].get_attribute('textContent')
    record['合同约定付款节点'] = rows[45].get_attribute('textContent')
    record['收款节点总计'] = rows[46].get_attribute('textContent')
    record['比例合计'] = rows[47].get_attribute('textContent')
    record['盖章人'] = rows[49].get_attribute('textContent')
    record['盖章状态'] = rows[50].get_attribute('textContent')
    record['归档确认'] = rows[51].get_attribute('textContent')
    total_contracts.append(record)

def income_handler(total_income, record, driver):
    logging.info('[收入]%s' % driver.title)
    iframe = driver.find_element_by_id('zwIframe')
    driver.switch_to.frame(iframe)
    while driver.execute_script("return document.readyState") != "complete" :
        time.sleep(1)
        logging.info('[收入]%s：等待iframe加载' % driver.title)
    rows = driver.find_elements_by_css_selector('section [class*="browse"]')
    while True:
        # 41 Only one income record; 51 means two records; 62 means three records
        if len(rows) >= 40:break
        time.sleep(1)
        logging.info('[收入]%s：需要加载40条，已经加载%s条，等待' % (driver.title,len(rows)))
        rows = driver.find_elements_by_css_selector('section [class*="browse"]')
    logging.info('[收入]%s：需要加载40条，已经加载%s条' % (driver.title,len(rows)))
    record['编号'] = rows[2].get_attribute('textContent')
    record['开票人'] = rows[3].get_attribute('textContent')
    record['申请部门'] = rows[4].get_attribute('textContent')
    record['申请日期'] = rows[5].get_attribute('textContent')
    record['合同号'] = rows[6].get_attribute('textContent')
    record['合同价'] = rows[7].get_attribute('textContent')
    record['合同名称'] = rows[8].get_attribute('textContent')
    record['客户单位名称'] = rows[9].get_attribute('textContent')
    record['银行账号'] = rows[10].get_attribute('textContent')
    record['开户银行'] = rows[11].get_attribute('textContent')
    record['税号'] = rows[12].get_attribute('textContent')
    record['地址'] = rows[13].get_attribute('textContent')
    record['电话'] = rows[14].get_attribute('textContent')
    record['业务类型'] = rows[15].get_attribute('textContent')
    record['来源类型'] = rows[16].get_attribute('textContent')
    record['备注说明'] = rows[17].get_attribute('textContent')
    record['收款类型'] = rows[19].get_attribute('textContent')
    record['开票类型'] = rows[20].get_attribute('textContent')
    record['应开票金额'] = rows[21].get_attribute('textContent')
    record['已开票金额'] = rows[22].get_attribute('textContent')
    record['本次开票金额'] = rows[23].get_attribute('textContent')
    record['发票类型'] = rows[24].get_attribute('textContent')
    record['税率'] = rows[25].get_attribute('textContent')
    record['税额'] = rows[26].get_attribute('textContent')
    record['金额(去税)'] = rows[27].get_attribute('textContent')
    record['发票号码'] = rows[28].get_attribute('textContent')
    record['备注'] = rows[29].get_attribute('textContent')
    total_income.append(record)

def weekly_reports_handler(total_weekly_reports, record, driver):
    logging.info('[周报]%s' % driver.title)
    week = args.week

    iframe = driver.find_element_by_id('zwIframe')
    driver.switch_to.frame(iframe)
    time.sleep(1)
    while driver.execute_script("return document.readyState") != "complete" :
        time.sleep(1)
        logging.info('[周报]%s：等待iframe加载' % driver.title)
    rows = driver.find_elements_by_css_selector('.is-detailshover')
    for row in rows:
       # try:
       #     tds = row.find_elements_by_css_selector('section [class*="browse"]')
       # except StaleElementReferenceException:
       #     time.sleep(1)
       #     tds = row.find_elements_by_css_selector('section [class*="browse"]')

        tds = row.find_elements_by_css_selector('section [class*="browse"]')

        report = {}
        report['发起人'] = record['发起人']
        report['年份'] = tds[0].get_attribute('textContent')
        report['月份'] = tds[1].get_attribute('textContent')
        report['周'] = tds[2].get_attribute('textContent')
        report['客户名称'] = tds[3].get_attribute('textContent')
        report['销售人员'] = tds[4].get_attribute('textContent')
        report['大类'] = tds[5].get_attribute('textContent')
        report['中类'] = tds[6].get_attribute('textContent')
        report['小类'] = tds[7].get_attribute('textContent')
        report['耗时'] = tds[8].get_attribute('textContent')
        report['加班'] = tds[9].get_attribute('textContent')
        report['下周计划'] = tds[10].get_attribute('textContent')
        report['具体工作描述'] = tds[11].get_attribute('textContent')

        if report['耗时']:
            report['耗时'] = float(report['耗时'])
        else:
            report['耗时'] = 0

        if report['加班']:
            report['加班'] = float(report['加班'])
        else:
            report['加班'] = 0

        if week is None:
            total_weekly_reports.append(report)
        elif report['周'] == week:
            total_weekly_reports.append(report)
        else:
            logging.info('SKIP [周报]%s [%s]' % (driver.title,report))

def others_handler(others, record, driver):
    logging.info('[其他]%s' % driver.title)
    others.append(record)

def to_excel(total_records, total_travels, total_procurement, total_contracts, total_income, total_weekly_reports, others):
    df = pd.DataFrame(total_records)
    df.to_excel('data/总表.xlsx', sheet_name='总表', engine='xlsxwriter')

    df = pd.DataFrame(total_travels)
    df.to_excel('data/出差单.xlsx', sheet_name='出差单', engine='xlsxwriter')

    df = pd.DataFrame(total_procurement)
    df.to_excel('data/采购单.xlsx', sheet_name='采购单', engine='xlsxwriter')

    df = pd.DataFrame(total_contracts)
    df.to_excel('data/合同单.xlsx', sheet_name='合同单', engine='xlsxwriter')

    df = pd.DataFrame(total_income)
    df.to_excel('data/收入单.xlsx', sheet_name='收入单', engine='xlsxwriter')

    df = pd.DataFrame(total_weekly_reports)
    df.to_excel('data/周报.xlsx', sheet_name='周报', engine='xlsxwriter')

    df = pd.DataFrame(others)
    df.to_excel('data/未归类.xlsx', sheet_name='未归类', engine='xlsxwriter')

def to_report(total_weekly_reports):
    # 售前支持 - 技术方案
    # 售前支持 - 技术交流
    # 售前支持 - PoC
    # 售前支持 - 投标工作
    # 售前支持 - 市场推广
    # 项目交付 - 项目管理

    # 内部工作 - 提供培训
    # 内部工作 - 参与培训
    # 内部工作 - 技术突破
    # 内部工作 - 证书过审
    # 内部工作 - 控标点过审
    # 内部工作 - 合同过审
    # 内部工作 - 团队管理
    # 内部工作 - 离职交接

    # 行政事务

    # 休假

    project_list = objects.ProjectList()

    for report in total_weekly_reports:
        p = objects.Project()
        p.create(report)
        project_list.insert(p)

    project_list.to_excel()
    project_list.to_report()


def main():

    logger()

    driver = login()

    total_number = 0
    total_records = []
    total_travels = []
    total_procurement = []
    total_contracts = []
    total_income = []
    total_weekly_reports = []
    others = []

    current_window_handle = driver.current_window_handle
    # Get each record from List Done and handle it.
    while True:
        rows = driver.find_elements_by_css_selector('#list tr')
        total_number = total_number + len(rows)

        # Limit to 40 records by default if not set limit number
        if total_number > args.num:
            break

        for row in rows:
           # try:
           #     tds = row.find_elements_by_tag_name('td')
           # except StaleElementReferenceException:
           #     time.sleep(1)
           #     tds = row.find_elements_by_tag_name('td')

            tds = row.find_elements_by_tag_name('td')

            subject = tds[1].find_element_by_tag_name('div').get_attribute('title')
            start_member_name = tds[2].find_element_by_tag_name('div').get_attribute('title')
            pre_approver_name = tds[3].find_element_by_tag_name('div').get_attribute('title')
            start_date = tds[4].find_element_by_tag_name('div').get_attribute('title')
            deal_time = tds[5].find_element_by_tag_name('div').get_attribute('title')
            current_nodes_info = tds[6].find_element_by_tag_name('div').get_attribute('title')
            record = {'标题': subject, '发起人':start_member_name, '上一处理人':pre_approver_name, '发起时间':start_date, '处理时间':deal_time, '当前代办人':current_nodes_info}
            total_records.append(record)

            # Check the message setting and ignore it
           # msgsettingmenu = driver.find_elements_by_class_name('msgSettingMenu')
           # if len(msgsettingmenu) > 1:
           #     msgsettingmenu[1].click()
            # Click each row and Open new tab to fetch more data
           # try:
           #     webdriver.ActionChains(driver).move_to_element(row).click(row).perform()
           # except StaleElementReferenceException:
           #     time.sleep(1)
           #     webdriver.ActionChains(driver).move_to_element(row).click(row).perform()

            webdriver.ActionChains(driver).move_to_element(row).click(row).perform()
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            window_handles = driver.window_handles
            driver.switch_to.window(window_handles[1])
            if '出差单' in subject:
                travels_handler(total_travels, record, driver)
            elif '销售合同' in subject:
                contracts_handler(total_contracts, record, driver)
            elif '销售开票' in subject:
                income_handler(total_income, record, driver)
            elif '周报' in subject:
                weekly_reports_handler(total_weekly_reports, record, driver)
            elif '采购' in subject:
                procurement_handler(total_procurement, record, driver)
            else:
                others_handler(others, record, driver)

            driver.close()
            driver.switch_to.window(current_window_handle)
            # Move to List Done
            iframe = driver.find_element_by_id('mainIframe')
            driver.switch_to.frame(iframe)

        if len(rows) < 20:
            break

        input_page = driver.find_elements_by_class_name('common_over_page_txtbox')[1]
        page_num = int(input_page.get_attribute('value'))
        logging.info('[页码]第%d页结束' % page_num)
        next_page_num = page_num + 1
        input_page.clear()
        input_page.send_keys(next_page_num)

        current_page_num = int(driver.find_elements_by_class_name('common_over_page_txtbox')[1].get_attribute('value'))
        logging.info('[页码]进入第%d页' % current_page_num)

        page_go = driver.find_element_by_class_name('common_over_page_go')
        webdriver.ActionChains(driver).move_to_element(page_go).click(page_go).perform()

        #page_next_button = driver.find_element_by_class_name('pageNext')
        #webdriver.ActionChains(driver).move_to_element(page_next_button).click(page_next_button).perform()
        time.sleep(1)

    driver.quit()

    to_excel(total_records, total_travels, total_procurement, total_contracts, total_income, total_weekly_reports, others)
    to_report(total_weekly_reports)


if __name__ == '__main__':
    main()
