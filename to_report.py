import argparse
import logging
import time
import pandas as pd
import jinja2
import time
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams['font.sans-serif'] = [u'Songti SC']
plt.rcParams['axes.unicode_minus'] = False

parser = argparse.ArgumentParser()
parser.add_argument('--week')
args = parser.parse_args()

def logger():
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    DATE_FORMAT = "%m/%d/%Y %H:%M:%S %p"
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, datefmt=DATE_FORMAT)

def to_report(report):
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

    all_presales = report.groupby([u'发起人'])[u'耗时'].sum().sort_values(ascending=True)
    all_presales.index.name = None
    all_presales.plot(kind='barh',figsize=(10,10),fontsize='15')
    plt.savefig('all_presales.svg')

    all_tasks = report.groupby(u'中类')[u'耗时'].sum().sort_values(ascending=True)
    all_tasks.index.name = None
    all_tasks.plot(kind='barh',figsize=(10,10),fontsize='15')
    plt.savefig('all_tasks.svg')

    all_projects = report.groupby(u'客户名称')[u'耗时'].sum().sort_values(ascending=True)
    all_projects.index.name=None
    all_projects.plot(kind='barh',figsize=(15,15),fontsize='15')
    plt.savefig('all_projects.svg')

    # Presales/Task Report
    presale_by_task_report = pd.DataFrame(index=all_presales.index, columns=all_tasks.index)
    presale_by_task_report=presale_by_task_report.fillna(0)
    presale_by_task_report.index.name=None

    for i in all_presales.index:
        presale_tasks = report[report.发起人 == i].groupby([u'中类'])[u'耗时'].sum()
        for j in all_tasks.index:
            if j in presale_tasks.index:
                presale_by_task_report[j][i]=presale_tasks[j]

    ax=presale_by_task_report.plot.barh(stacked=True,figsize=(20,20),fontsize='15');
    for i, v in enumerate(all_presales):
        ax.text(v+1, i, str(v), color='black', fontweight='bold', fontsize=13)
    plt.savefig('presale_by_task.svg')

    # Project/Task Report
    project_by_task_report = pd.DataFrame(index=all_projects.index, columns=all_tasks.index)
    project_by_task_report=project_by_task_report.fillna(0)
    project_by_task_report.index.name=None
    for i in all_projects.index:
        project_tasks = report[report.客户名称 == i].groupby([u'中类'])[u'耗时'].sum()
        for j in all_tasks.index:
            if j in project_tasks.index:
                project_by_task_report[j][i]=project_tasks[j]
    ax=project_by_task_report.plot.barh(stacked=True,figsize=(80,80),fontsize='50');
    ax.legend(fontsize=35)
    for i, v in enumerate(all_projects):
        ax.text(v, i, str(v), color='black', fontweight='bold', fontsize=45)
    plt.savefig('project_by_task.svg')

    # Project vs. Presales by Week Report
    presales_count=report.groupby([u'周'])[u'发起人'].apply(lambda x:len(set(x)))
    projects_count=report.groupby([u'周'])[u'客户名称'].apply(lambda x:len(set(x)))
    presale_vs_project_by_week_report = pd.DataFrame(index=presales_count.index, columns=['售前人数', '项目个数', '人均支持项目数'])
    presale_vs_project_by_week_report=presale_vs_project_by_week_report.fillna(0)
    presale_vs_project_by_week_report.index.name=None
    presale_vs_project_by_week_report[u'售前人数']=presales_count
    presale_vs_project_by_week_report[u'项目个数']=projects_count
    presale_vs_project_by_week_report[u'人均支持项目数']=np.around(presale_vs_project_by_week_report[u'项目个数']/presale_vs_project_by_week_report[u'售前人数'], decimals=1)


    # Echarts Template handling
    env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath=''))
    template_echarts = env.get_template('template_echarts.html')
    presale_by_task_chart_data = {
        'legend': list(presale_by_task_report.columns.values),
        'y_data': list(presale_by_task_report.index.values),
        'series': []
    }

    for (column_name, column_data) in presale_by_task_report.iteritems():
        presale_by_task_chart_data['series'].append({
            'name':column_name,
            'type': 'bar',
            'stack': 'total',
            'label': {
                'show': 'true'
            },
            'emphasis': {
                'focus': 'series'
            },
            'data':list(column_data)
        })

    project_by_task_chart_data = {
         'legend': list(project_by_task_report.columns.values),
         'y_data': list(project_by_task_report.index.values),
         'series': []
     }

    for (column_name, column_data) in project_by_task_report.iteritems():
         project_by_task_chart_data['series'].append({
             'name':column_name,
             'type': 'bar',
             'stack': 'total',
             'label': {
                 'show': 'true'
             },
             'emphasis': {
                 'focus': 'series'
             },
             'data':list(column_data)
         })

    presale_vs_project_by_week_chart_data = {
        'legend': list(presale_vs_project_by_week_report.columns.values),
        'x_data': list(presale_vs_project_by_week_report.index.values),
        'series': []
    }

    for (column_name, column_data) in presale_vs_project_by_week_report.iteritems():
        if column_name == '人均支持项目数':
            presale_vs_project_by_week_chart_data['series'].append({
                'name':column_name,
                'type': 'line',
                'yAxisIndex': '1',
                'label': {
                    'show': 'true'
                },
                'data':list(column_data)
        })
        else:
            presale_vs_project_by_week_chart_data['series'].append({
                'name':column_name,
                'type': 'bar',
                 'label': {
                     'show': 'true'
                 },
                'data':list(column_data)
            })

    html = template_echarts.render(
        presale_by_task_chart_data=presale_by_task_chart_data,
        project_by_task_chart_data=project_by_task_chart_data,
        presale_vs_project_by_week_chart_data=presale_vs_project_by_week_chart_data
    )

    # Write the Echarts HTML file
    with open('report_echarts.html', 'w') as f:
        f.write(html)


def main():
    logger()

    df = pd.DataFrame()
    total_weekly_reports = pd.read_excel("data/周报汇总2021H1.xlsx", sheet_name="周报")
    logging.info('Read Excel')

    to_report(total_weekly_reports)
    logging.info('Generate Report')

if __name__ == '__main__':
    main()
