import datetime
import pandas as pd
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
            p.sales = p.sales | project.sales
            p.pre_sales = p.pre_sales | project.pre_sales
        else:
            self.project_list[project.project_name] = project

    def to_excel(self):
        output = []
        for p in self.project_list:
            tmp = {}
            tmp['开始日期'] = self.project_list[p].start_date
            tmp['项目名称'] = self.project_list[p].project_name
            tmp['销售'] = self.project_list[p].sales
            tmp['售前'] = self.project_list[p].pre_sales
            output.append(tmp)

        df = pd.DataFrame(output)
        df.to_excel('跟踪表.xlsx', sheet_name='跟踪表')

class Project():
    # Version 1.0: Initial version
    VERSION = '1.0'

    def __init__(self, *args, **kwargs):

        # 开始日期
        self.start_date = None

        # 结束日期
        self.end_date = None

        # 更新日期 TODO

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

    def create(self, report):

        self.pre_sales.add(report['发起人'])
        self.project_name = report['客户名称']

        st = '%(y)s-%(m)s' % {'y':report['年份'], 'm':report['月份']}
        self.start_date = datetime.datetime.strptime(st, '%Y-%m')

        st = '%(y)s-12' % {'y':report['年份']}
        self.end_date = datetime.datetime.strptime(st, '%Y-%m')

    def update():
        pass
