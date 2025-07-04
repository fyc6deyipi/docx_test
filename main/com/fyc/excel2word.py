import datetime as dt
from docxtpl import DocxTemplate,InlineImage
from datetime import datetime, timedelta
import pandas as pd



class excel2word:
    my_dict = {}
    def __init__(self,excel_url,word_url,write_url):
        self.excel_url=excel_url
        self.word_url=word_url
        self.write_url=write_url
        self.data = pd.read_excel(self.excel_url, sheet_name='Sheet1', header=0)
    def get_last_friday(self, param='0'):
        if dt.datetime.today().weekday() >= 4:
            last_friday = datetime.today() - timedelta(days=(datetime.today().weekday() + 3) % 7) - timedelta(7)
            last_last_friday = last_friday - timedelta(7)
        else:
            # 直接计算上周五
            last_friday = datetime.today() - timedelta(days=(datetime.today().weekday() + 3) % 7)

            # 直接计算上上周五
            last_last_friday = last_friday - timedelta(7)

        if (param == -1):
            return int(last_last_friday.strftime('%Y%m%d'))
        return int(last_friday.strftime('%Y%m%d'))
    def get_condition(self,sys,week=0):
        condition=''
        if sys=='sys_19':
            condition = (self.data['ds'] == self.get_last_friday(week)) & (self.data['sys_status'] == '已纳管已盘点') & (self.data['is_core_sys'] == '是')
        elif sys=='sys_93':
            condition = (self.data['ds'] == self.get_last_friday(week)) & (self.data['sys_status'].str.contains('已纳管'))
        elif sys=='sys_all':
            condition= (self.data['ds'] == self.get_last_friday(week))
        else:
            print('注意系统范围')
        return condition
    def read_excel_part1(self):
        # 1源端系统全量纳管和动态监测
        # a_源端系统1
        self.my_dict['a_源端系统'] = 200
        # a_纳管源端系统
        condition = (self.data['ds'] == self.get_last_friday()) & (self.data['sys_status'].str.contains('已纳管'))
        a_纳管源端系统 = self.data[condition]
        self.my_dict['a_纳管源端系统'] = len(a_纳管源端系统)
        # a_纳管率
        self.my_dict['a_纳管率'] = round(self.my_dict['a_纳管源端系统'] / self.my_dict['a_源端系统'] * 100, 2)
        # 2重点针对
        # a_纳管源端系统表数量
        a_纳管源端系统表数量 = a_纳管源端系统['jczb001'].sum() / 10000
        self.my_dict['a_纳管源端系统表数量'] = round(a_纳管源端系统表数量, 2)
        # a_技术元数据合格率 jczb008
        a_技术元数据合格率 = a_纳管源端系统['jczb008'].min()
        self.my_dict['a_技术元数据合格率'] = round(a_技术元数据合格率, 2)
        # a_技术元数据合格率低于99 jczb008
        a_技术元数据合格率低于99 = a_纳管源端系统[a_纳管源端系统['jczb008'] < 99]
        self.my_dict['a_技术元数据合格率低于99'] = len(a_技术元数据合格率低于99)

        # 3系统业务梳理
        # a_已盘点核心系统
        condition = (a_纳管源端系统['sys_status'] == '已纳管已盘点') & (a_纳管源端系统['is_core_sys'] == '是')
        a_已盘点核心系统 = a_纳管源端系统[condition]
        self.my_dict['a_已盘点核心系统'] = len(a_已盘点核心系统)
        # a_已盘点系统功能模块 jczb018
        a_已盘点核心系统功能模块 = a_已盘点核心系统['jczb018'].sum()
        self.my_dict['a_已盘点核心系统功能模块'] = int(a_已盘点核心系统功能模块)
        # a_已盘点核心系统核心表 jczb019
        a_已盘点核心系统核心表 = a_已盘点核心系统['jczb019'].sum()
        self.my_dict['a_已盘点核心系统核心表'] = int(a_已盘点核心系统核心表)
        # a_已盘点核心系统主数据表 jczb020
        a_已盘点核心系统主数据表 = a_已盘点核心系统['jczb020'].sum()
        self.my_dict['a_已盘点核心系统主数据表'] = int(a_已盘点核心系统主数据表)
        # a_已盘点核心系统核心业务数据表 jczb021
        a_已盘点核心系统核心业务数据表 = a_已盘点核心系统['jczb021'].sum()
        self.my_dict['a_已盘点核心系统核心业务数据表'] = int(a_已盘点核心系统核心业务数据表)
        # a_已盘点核心系统核心表占比
        self.my_dict['a_已盘点核心系统核心表占比'] = round(self.my_dict['a_已盘点核心系统核心表'] / a_已盘点核心系统['jczb001'].sum() * 100, 2)

        # 4核心数据接入
        # a_接入中台表 jczb023
        a_接入中台表 = a_已盘点核心系统['jczb023'].sum()
        self.my_dict['a_接入中台表'] = int(a_接入中台表)
        # a_接入核心表 jczb026
        a_接入核心表 = a_已盘点核心系统['jczb026'].sum()
        self.my_dict['a_接入核心表'] = int(a_接入核心表)
        # a_按需接入表 jczb031
        a_按需接入表 = a_已盘点核心系统['jczb031'].sum()
        self.my_dict['a_按需接入表'] = int(a_按需接入表)
        for key, value in self.my_dict.items():
            print(f"{key}: {value}")
        # a_接入核心表占比
        self.my_dict['a_接入核心表占比'] = round(self.my_dict['a_接入核心表'] / self.my_dict['a_接入中台表'] * 100, 2)
        # a_质量校核表 jczb043
        a_质量校核表 = a_已盘点核心系统['jczb043'].sum()
        self.my_dict['a_质量校核表'] = int(a_质量校核表)
        # a_质量校核率
        self.my_dict['a_质量校核率'] = round(self.my_dict['a_质量校核表'] / self.my_dict['a_接入中台表'] * 100, 2)
        # a_技术质量合格表 jczb054
        a_技术质量合格表 = a_已盘点核心系统['jczb054'].sum()
        self.my_dict['a_技术质量合格表'] = int(a_技术质量合格表)
        # a_技术质量合格率
        self.my_dict['a_技术质量合格率'] = round(self.my_dict['a_技术质量合格表'] / self.my_dict['a_质量校核表'] * 100,4)
        # a_业务质量合格表 jczb093
        a_业务质量合格表 = a_已盘点核心系统['jczb093'].sum()
        self.my_dict['a_业务质量合格表'] = int(a_业务质量合格表)
        # a_业务质量合格率
        self.my_dict['a_业务质量合格率'] = round(self.my_dict['a_业务质量合格表'] / self.my_dict['a_质量校核表'] * 100,4)

        # 5 数据共享和服务
        # a_共享层表数量 jczb066
        # a_共享层表数量=df['jczb066'].sum()
        # self.my_dict['a_共享层表数量'] = int(a_共享层表数量)

        # 6 元数据和血缘信息维护
        # a_技术元数据质量合格表 jczb007
        a_技术元数据质量合格表 = a_已盘点核心系统['jczb007'].sum()
        self.my_dict['a_技术元数据质量合格表'] = int(a_技术元数据质量合格表)
        # a_技术元数据质量合格率 jczb008   纳管表数量（表） jczb089
        a_纳管表数量 = a_已盘点核心系统['jczb089'].sum()
        self.my_dict['a_技术元数据质量合格率'] = round(self.my_dict['a_技术元数据质量合格表'] / a_纳管表数量 * 100, 2)
        # a_业务元数据质量合格表 jczb009
        a_业务元数据质量合格表 = a_已盘点核心系统['jczb009'].sum()
        self.my_dict['a_业务元数据质量合格表'] = int(a_业务元数据质量合格表)
        # a_业务元数据质量合格率
        self.my_dict['a_业务元数据质量合格率'] = round(self.my_dict['a_业务元数据质量合格表'] / a_纳管表数量 * 100, 2)
        #  a_管理元数据质量标签合格表 管理元数据质量标签合格表数量  jczb011  需人工配置质量规则校核表数量  jczb033
        a_管理元数据质量标签合格表 = a_已盘点核心系统['jczb011'].sum()
        a_需人工配置质量规则校核表 = a_已盘点核心系统['jczb033'].sum()
        self.my_dict['a_管理元数据质量标签维护率'] = round(a_管理元数据质量标签合格表 / a_需人工配置质量规则校核表 * 100, 2)
        # a_93接入贴源层表
        a_93接入贴源层表 = a_纳管源端系统['jczb023'].sum()
        self.my_dict['a_93接入贴源层表'] = int(a_93接入贴源层表)
        # a_93源端至共享层有血缘的源端表数量 jczb078
        a_93源端至共享层有血缘的源端表数量 = a_纳管源端系统['jczb078'].sum()
        self.my_dict['a_93源端至共享层有血缘的源端表数量'] = int(a_93源端至共享层有血缘的源端表数量)
        # a_93源端表使用率
        self.my_dict['a_93源端表使用率'] = round(self.my_dict['a_93源端至共享层有血缘的源端表数量'] / self.my_dict['a_93接入贴源层表'] * 100, 2)
        # a_93共享层一级系统表 jczb076
        a_93共享层一级系统表 = a_纳管源端系统['jczb076'].sum()
        self.my_dict['a_93共享层一级系统表'] = int(a_93共享层一级系统表)
        # a_93共享层表血缘覆盖率
        self.my_dict['a_93共享层表血缘覆盖率'] = round(self.my_dict['a_93源端至共享层有血缘的源端表数量'] / a_93共享层一级系统表 * 100, 2)

    def read_excel_part2_1(self):

        # condition = self.get_condition(sys='sys_93',week=0)
        data93_l = self.data[self.data['systemname'] == 'sys_93'].copy()[['jczb001','jczb008','ds']]
        data93_l.rename(columns={'jczb001':'源端数据表数量','jczb008':'技术元数据质量合格率'}, inplace=True)
        data93_l.set_index('ds', inplace=True)
        data93_l = data93_l.transpose()
        data93_l.reset_index(drop=False, inplace=True)
        data93_l=data93_l.assign(
            value1= lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])
                    ,(lambda s: s.str.contains('数量'),round((data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])/data93_l[self.get_last_friday(-1)]*100,2))
                ]
            )
        )

        data19_l = self.data[self.data['systemname'] == 'sys_19'].copy()[['jczb001', 'jczb008', 'ds']]
        data19_l.rename(columns={'jczb001': '已盘点源端数据表数量', 'jczb008': '已盘点技术元数据质量合格率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                     data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)])
                    , (lambda s: s.str.contains('数量'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        data = pd.concat([data93_l, data19_l], ignore_index=True)
        # print(data)
        for x in range (4):
            for y in range (4):
                self.my_dict['b1_t_'+str(x)+'_'+str(y)] = data.iloc[x,y]

        condition = self.get_condition('sys_19')
        data19 = self.data[condition][['systemname', 'jczb008']]
        sort_values = data19[data19['jczb008'] < 90].sort_values(by=['jczb008'], ascending=False)
        for x in range (len(sort_values)):
            print(sort_values.iloc[x,0])
        print(data19)
        print(sort_values)
        

    def read_excel_part2_2(self):

        data93_l = self.data[self.data['systemname'] == 'sys_93'].copy()[['jczb001','jczb010','ds']]
        data93_l.rename(columns={'jczb001':'源端数据表数量','jczb010':'业务元数据质量合格率'}, inplace=True)
        data93_l.set_index('ds', inplace=True)
        data93_l = data93_l.transpose()
        data93_l.reset_index(drop=False, inplace=True)
        data93_l=data93_l.assign(
            value1= lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])
                    ,(lambda s: s.str.contains('数量'),round((data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])/data93_l[self.get_last_friday(-1)]*100,2))
                ]
            )
        )

        data19_l = self.data[self.data['systemname'] == 'sys_19'].copy()[['jczb001', 'jczb010', 'ds']]
        data19_l.rename(columns={'jczb001': '已盘点源端数据表数量', 'jczb010': '已盘点业务元数据质量合格率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                     data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)])
                    , (lambda s: s.str.contains('数量'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        data = pd.concat([data93_l, data19_l], ignore_index=True)
        print(data)
        for x in range (4):
            for y in range (4):
                self.my_dict['b2_t_'+str(x)+'_'+str(y)] = data.iloc[x,y]

    def read_excel_part2_3(self):

        data93_l = self.data[self.data['systemname'] == 'sys_93'].copy()[['jczb001','jczb012','jczb013','ds']]
        data93_l.rename(columns={'jczb001':'源端数据表数量','jczb012':'源端系统管理元数据质量标签维护率','jczb013':'管理元数据负面清单维护率'}, inplace=True)
        data93_l.set_index('ds', inplace=True)
        data93_l = data93_l.transpose()
        data93_l.reset_index(drop=False, inplace=True)
        data93_l=data93_l.assign(
            value1= lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])
                    ,(lambda s: s.str.contains('数量'),round((data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])/data93_l[self.get_last_friday(-1)]*100,2))
                ]
            )
        )
        data19_l = self.data[self.data['systemname'] == 'sys_19'].copy()[['jczb001', 'jczb012', 'jczb013', 'ds']]
        data19_l.rename(columns={'jczb001': '已盘点源端系统数据表数量', 'jczb012': '已盘点源端系统管理元数据质量标签维护率',
                                 'jczb013': '已盘点管理元数据负面清单维护率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                     data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)])
                    , (lambda s: s.str.contains('数量'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        data = pd.concat([data93_l, data19_l], ignore_index=True)
        print(data)
        for x in range (6):
            for y in range (4):
                self.my_dict['b3_t_'+str(x)+'_'+str(y)] = data.iloc[x,y]





    def run (self):
        # self.read_excel_part1()
        self.read_excel_part2_1()
        # self.read_excel_part2_2()
        # self.read_excel_part2_3()

    def sout_dict(self):
        for key, value in self.my_dict.items():
            print(f"{key}: {value}")
    def write(self):
        doc = DocxTemplate(self.word_url)
        doc.render(self.my_dict)
        doc.save(self.write_url)
