#import datetime
import jinja2
import pandas as pd
import time
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = [u'Songti SC']
plt.rcParams['axes.unicode_minus'] = False

from enum import Enum

class Phase(Enum):
     # 项目阶段：技术交流、技术方案、PoC、招投标、合同签署、实施交付、维护、关闭
     TECH_COMMUNICATION = '技术交流'
     TECH_SOLUTION = '技术方案'
     POC = 'POC'
     BIDING = '招投标'
     CONTRACT = '合同签署'
     DELIVER = '实施交付'

class Status(Enum):
    # 项目状态：丢失、中标、加强、推迟、稳固、终止、跟进、风险
    LOST = '丢失'
    WIN = '中标'
    RISK = '风险'
    FOLLOW = '跟进'

class ProjectList():
    def __init__(self, *args, **kwargs):
        self.project_list = {}

    def insert(self, project):
        if project.project_name in self.project_list:
            p = self.project_list[project.project_name]

            if p.start_date > p.update_date:
                p.start_date = p.update_date

            p.sales = p.sales | project.sales
            p.pre_sales = p.pre_sales | project.pre_sales
            p.history = p.history + project.history
            p.hours = p.hours + project.hours

            p.history = sorted(p.history, key = lambda x:x['WW'], reverse=True)

            p.phase = p.history[0]['Detail'][0]
        else:
            self.project_list[project.project_name] = project

    def check(self):
        pass


    def to_excel(self):
        output = []
        for p in self.project_list:
            tmp = {}
            tmp['开始日期'] = self.project_list[p].start_date
            tmp['项目名称'] = self.project_list[p].project_name
            tmp['销售'] = self.project_list[p].sales
            tmp['售前'] = self.project_list[p].pre_sales
            tmp['跟进记录'] = self.project_list[p].history

            # TODO Use default status
            tmp['状态'] = self.project_list[p].status.value
            tmp['阶段'] = self.project_list[p].phase.value
            tmp['耗时'] = self.project_list[p].hours
            output.append(tmp)

        df = pd.DataFrame(output)
        df.to_excel('跟踪表.xlsx', sheet_name='跟踪表')

    def to_report(self):

        # In Python, weeks starts from 0 while we start from 1.
        current_work_week = int(time.strftime('%W')) + 1

        biding_table = []
        poc_table = []
        tech_communication_table = []
        tech_solution_table = []

        for p in self.project_list:
            tmp = {}
            tmp['项目名称'] = self.project_list[p].project_name
            tmp['销售'] = ','.join(self.project_list[p].sales)
            tmp['售前'] = ','.join(self.project_list[p].pre_sales)
            tmp['耗时'] = self.project_list[p].hours
            tmp['跟进记录'] = self.project_list[p].history[0]['Detail'][1]
            tmp['后续工作'] = self.project_list[p].history[0]['Detail'][2]

            if self.project_list[p].phase is Phase.BIDING:
                biding_table.append(tmp)
            elif self.project_list[p].phase is Phase.POC:
                poc_table.append(tmp)
            elif self.project_list[p].phase is Phase.TECH_COMMUNICATION:
                tech_communication_table.append(tmp)
            elif self.project_list[p].phase is Phase.TECH_SOLUTION:
                tech_solution_table.append(tmp)

        df = pd.DataFrame(biding_table)
        biding_table_html = df.to_html()

        df = pd.DataFrame(poc_table)
        poc_table_html = df.to_html()

        df = pd.DataFrame(tech_communication_table)
        tech_communication_table_html = df.to_html()

        df = pd.DataFrame(tech_solution_table)
        tech_solution_table_html = df.to_html()

        # Template handling
        env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath=''))
        template = env.get_template('template.html')
        html = template.render(work_week=current_work_week, biding_table=biding_table_html, poc_table=poc_table_html,
                tech_communication_table=tech_communication_table_html, tech_solution_table=tech_solution_table_html )

        # Write the HTML file
        html = html.replace(r"\n", "<br/>")
        with open('report.html', 'w') as f:
            f.write(html)

        # Plot
        df = pd.DataFrame()
        report = pd.read_excel("周报.xlsx", sheet_name="周报")

        all_presales = report.groupby([u'发起人'])[u'耗时'].sum()
        all_presales.index.name = None
        all_presales.plot(kind='barh',figsize=(10,10),fontsize='15')
        plt.savefig('all_presales.svg')

        all_tasks = report.groupby(u'中类')[u'耗时'].sum()
        all_tasks.index.name = None
        all_tasks.plot(kind='barh',figsize=(10,10),fontsize='15')
        plt.savefig('all_tasks.svg')

        final_report = pd.DataFrame(index=all_presales.index, columns=all_tasks.index)
        final_report=final_report.fillna(0)
        final_report.index.name=None

        for i in all_presales.index:
            presale_tasks = report[report.发起人 == i].groupby([u'中类'])[u'耗时'].sum()
            for j in all_tasks.index:
                if j in presale_tasks.index:
                    final_report[j][i]=presale_tasks[j]

        ax=final_report.plot.barh(stacked=True,figsize=(20,20),fontsize='15');
        for i, v in enumerate(all_presales):
            ax.text(v+1, i, str(v), color='black', fontweight='bold', fontsize=13)
        plt.savefig('plot.svg')

        # Echarts Template handling
        env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath=''))
        template_echarts = env.get_template('template_echarts.html')
        chart_data = {
            'legend': list(final_report.columns.values),
            'y_data': list(final_report.index.values),
            'series': []
        }

        for (column_name, column_data) in final_report.iteritems():
            chart_data['series'].append({
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
        html = template_echarts.render(chart_data=chart_data)

        # Write the Echarts HTML file
        with open('report_echarts.html', 'w') as f:
            f.write(html)

class Project():
    # Version 1.0: Initial version
    VERSION = '1.0'

    def __init__(self, *args, **kwargs):

        # 开始日期
        self.start_date = None

        # 结束日期
        self.end_date = None

        # 更新日期
        self.update_date = None

        # 跟进历史
        self.history = []

        # 项目名称
        self.project_name = None

        # 销售
        self.sales = set()

        # 售前
        self.pre_sales = set()

        # 项目阶段：技术交流、技术方案、PoC、招投标、合同签署、实施交付、维护、关闭
        self.phase = Phase.TECH_COMMUNICATION

        # 项目状态：丢失、中标、加强、推迟、稳固、终止、跟进、风险
        self.status = Status.FOLLOW

        # 耗时
        self.hours = 0

    def create(self, report):

        self.sales.add(report['销售人员'])
        self.pre_sales.add(report['发起人'])
        self.project_name = report['客户名称']
        self.hours = report['耗时']
        self.update_date = int(report['周'])

# TODO Use work week as start date
#        st = '%(y)s-%(m)s' % {'y':report['年份'], 'm':report['月份']}
#        self.start_date = datetime.datetime.strptime(st, '%Y-%m')
#        self.start_date = st
        self.start_date = self.update_date

# TODO Use the end of this year as end date
        st = '%(y)s-12' % {'y':report['年份']}
#       self.end_date = datetime.datetime.strptime(st, '%Y-%m')
        self.end_date = st

        if '技术交流' in report['中类']:
            self.phase = Phase.TECH_COMMUNICATION
        elif '技术方案' in report['中类']:
            self.phase = Phase.TECH_SOLUTION
        elif 'POC' in report['中类']:
            self.phase = Phase.POC
        elif '投标工作' in report['中类']:
            self.phase = Phase.BIDING
        elif '合同审核' in report['中类']:
            self.phase = Phase.CONTRACT
        elif '项目管理' in report['中类']:
            self.phase = Phase.DELIVER

        self.history.append({'WW':self.start_date, 'Detail':[self.phase, report['具体工作描述'], report['下周计划'], self.hours, self.pre_sales]})

    def update():
        pass
