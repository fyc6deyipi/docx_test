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
    def get_last_friday(self, param=0):
        if dt.datetime.today().weekday() >= 4:
            last_friday = datetime.today() - timedelta(days=(datetime.today().weekday() + 3) % 7) - timedelta(7)
            last_last_friday = last_friday - timedelta(7)
        else:
            # 直接计算上周五
            last_friday = datetime.today() - timedelta(days=(datetime.today().weekday() + 3) % 7)

            # 直接计算上上周五
            last_last_friday = last_friday - timedelta(7)

        if param == -1:
            return int(last_last_friday.strftime('%Y%m%d'))
        return int(last_friday.strftime('%Y%m%d'))
    def get_condition(self,sys,week=0):
        condition=''
        if sys=='已纳管核心19':
            condition = (self.data['ds'] == self.get_last_friday(week)) & (self.data['sys_status'] == '已纳管已盘点') & (self.data['is_core_sys'] == '是')& (self.data['sys_code'] != 99999999)
        elif sys=='已纳管93':
            condition = (self.data['ds'] == self.get_last_friday(week)) & (self.data['sys_status'].str.contains('已纳管')) & (self.data['sys_code']  != 99999999)
        elif sys=='sys_all':
            condition= (self.data['ds'] == self.get_last_friday(week))
        else:
            print('注意系统范围')
        return condition
    def read_excel_part1(self):
        list_ng= self.data[self.get_condition('已纳管93')]
        condition = (self.data['ds'] == self.get_last_friday()) & (self.data['systemname'] == '已纳管93')
        line_ng = self.data[condition].iloc[0]
        # 1源端系统全量纳管和动态监测
        # a_源端系统1
        self.my_dict['a_源端系统'] = 1089
        # a_纳管源端系统
        self.my_dict['a_纳管源端系统'] = len(list_ng)
        # a_纳管率
        self.my_dict['a_纳管率'] = round(self.my_dict['a_纳管源端系统'] / self.my_dict['a_源端系统'] * 100, 2)
        # 2重点针对
        # a_纳管源端系统表数量
        self.my_dict['a_纳管源端系统表数量'] = round(line_ng['jczb001']/10000,2)
        # a_技术元数据合格率 jczb008
        self.my_dict['a_技术元数据合格率'] = line_ng['jczb008']
        # a_技术元数据合格率低于99 jczb008
        a_技术元数据合格率低于99 = list_ng[list_ng['jczb008'] < 99]
        self.my_dict['a_技术元数据合格率低于99'] = len(a_技术元数据合格率低于99)

        list_hx = self.data[self.get_condition('已纳管核心19')]
        condition = (self.data['ds'] == self.get_last_friday()) & (self.data['systemname'] == '已纳管核心19')
        line_hx = self.data[condition].iloc[0]
        # 3系统业务梳理
        # a_已盘点核心系统
        self.my_dict['a_已盘点核心系统'] = len(list_hx)
        # a_已盘点系统功能模块 jczb018
        self.my_dict['a_已盘点核心系统功能模块'] = line_hx['jczb018']
        # a_已盘点核心系统核心表 jczb019
        self.my_dict['a_已盘点核心系统核心表'] = line_hx['jczb019']
        # a_已盘点核心系统主数据表 jczb020
        self.my_dict['a_已盘点核心系统主数据表'] = line_hx['jczb020']
        # a_已盘点核心系统核心业务数据表 jczb021
        self.my_dict['a_已盘点核心系统核心业务数据表'] = line_hx['jczb021']
        # a_已盘点核心系统核心表占比
        self.my_dict['a_已盘点核心系统核心表占比'] = round(self.my_dict['a_已盘点核心系统核心表'] / line_hx['jczb001'].sum() * 100, 2)
        # a_核心范围不足系统
        list_hx=list_hx.copy()
        list_hx['lv']=round(list_hx['jczb019']/list_hx['jczb001'] * 100, 2)
        sort_values = list_hx[list_hx['lv'] < 20].sort_values(by=['lv'], ascending=True)[['systemname','lv']]
        self.my_dict['a_核心范围不足系统'] = sort_values.iloc[0,0]+'、'+sort_values.iloc[1,0]+'、'+sort_values.iloc[2,0]
        self.my_dict['a_核心范围不足系统数'] = len(sort_values)


        # 4核心数据接入
        # a_接入中台表 jczb023
        self.my_dict['a_接入中台表'] = line_hx['jczb023']
        # a_接入核心表 jczb026
        self.my_dict['a_接入核心表'] = line_hx['jczb026']
        # a_按需接入表 jczb031
        self.my_dict['a_按需接入表'] = line_hx['jczb031']
        # a_接入核心表占比
        self.my_dict['a_接入核心表占比'] = line_hx['jczb030']
        # a_质量校核表 jczb043
        self.my_dict['a_质量校核表'] = line_hx['jczb043']
        # a_质量校核率 jczb045
        self.my_dict['a_质量校核率'] = line_hx['jczb045']
        tmp=list_hx[list_hx['jczb045']<100][['systemname']]
        if len(tmp) == 0:
            self.my_dict['a_质量校核ms'] = ''
        elif len(tmp) == 1:
            self.my_dict['a_质量校核ms'] = '，其中，' + tmp.iloc[0, 0] + '未完成校核'
        elif len(tmp) > 1:
            self.my_dict['a_质量校核ms'] = '，其中，'+sort_values.iloc[0,0]+'、'+sort_values.iloc[1,0] + '未完成校核'

        # a_技术质量合格表 jczb054
        self.my_dict['a_技术质量合格表'] = line_hx['jczb054']
        # a_技术质量合格率
        self.my_dict['a_技术质量合格率'] = line_hx['jczb055']
        # tmp = list_hx[list_hx['jczb055'] < 90][['systemname','jczb055']]
        # print(len(tmp))
        # if len(tmp) == 0:
        #     self.my_dict['a_技术质量校核ms'] = ''
        # elif len(tmp) == 1:
        #     self.my_dict['a_技术质量校核ms'] = '，其中，' + tmp.iloc[0, 0] + '未完成校核'
        # elif len(tmp) > 1:
        #     self.my_dict['a_技术质量校核ms'] = '，其中，' + sort_values.iloc[0, 0] + '、' + sort_values.iloc[1, 0] + '未完成校核'

        # a_业务质量合格表 jczb093
        self.my_dict['a_业务质量合格表'] = line_hx['jczb093']
        # a_业务质量合格率 jczb094
        self.my_dict['a_业务质量合格率'] =line_hx['jczb094']



        # 5 数据共享和服务
        # a_共享层表数量 jczb066
        # a_共享层表数量=df['jczb066'].sum()
        # self.my_dict['a_共享层表数量'] = int(a_共享层表数量)

        # 6 元数据和血缘信息维护
        # a_技术元数据质量合格表 jczb007
        self.my_dict['a_技术元数据质量合格表'] = line_hx['jczb007']
        # a_技术元数据质量合格率 jczb008
        self.my_dict['a_技术元数据质量合格率'] = line_hx['jczb008']
        # a_业务元数据质量合格表 jczb009
        self.my_dict['a_业务元数据质量合格表'] = line_hx['jczb009']
        # a_业务元数据质量合格率 10
        self.my_dict['a_业务元数据质量合格率'] = line_hx['jczb010']
        #  a_管理元数据质量标签维护率 jczb012
        self.my_dict['a_管理元数据质量标签维护率'] = line_hx['jczb012']
        #  a_管理元数据负面清单维护率 jczb012
        self.my_dict['a_管理元数据负面清单维护率'] = line_hx['jczb013']


        # a_93接入贴源层表 23
        self.my_dict['a_93接入贴源层表'] = line_ng['jczb023']
        # a_93源端至共享层有血缘的源端表数量 jczb078
        self.my_dict['a_93源端至共享层有血缘的源端表数量'] = line_ng['jczb078']
        # a_93源端表使用率
        self.my_dict['a_93源端表使用率'] = line_ng['jczb079']
        # a_93共享层一级系统表 jczb076
        self.my_dict['a_93共享层一级系统表'] = line_ng['jczb078']

    def read_excel_part2_1(self):

        # condition = self.get_condition(sys='已纳管93',week=0)
        data93_l = self.data[self.data['systemname'] == '已纳管93'].copy()[['jczb001','jczb008','ds']]
        data93_l.rename(columns={'jczb001':'源端数据表数量','jczb008':'技术元数据质量合格率'}, inplace=True)
        data93_l.set_index('ds', inplace=True)
        data93_l = data93_l.transpose()
        data93_l.reset_index(drop=False, inplace=True)
        data93_l=data93_l.assign(
            value1= lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),round(data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)],2))
                    ,(lambda s: s.str.contains('数量'),round((data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])/data93_l[self.get_last_friday(-1)]*100,2))
                ]
            )
        )

        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[['jczb001', 'jczb008', 'ds']]
        data19_l.rename(columns={'jczb001': '已盘点源端数据表数量', 'jczb008': '已盘点技术元数据质量合格率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                     round(data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)],2))
                    , (lambda s: s.str.contains('数量'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        data = pd.concat([data93_l, data19_l], ignore_index=True)

        for x in range (4):
            for y in range (4):
                self.my_dict['b1_t_'+str(x)+'_'+str(y)] = data.iloc[x,y]

        condition = self.get_condition('已纳管核心19')
        data19 = self.data[condition][['systemname', 'jczb008']]
        sort_values = data19[data19['jczb008'] < 90].sort_values(by=['jczb008'], ascending=False)
        length=len(sort_values)
        if length == 0 :
            self.my_dict['b1_ms']=''
        elif len(sort_values) == 1 :
            self.my_dict['b1_ms'] = '其中，'+ sort_values.iloc[x, 0] +'系统技术元数据维护情况较差，合格率低于90%，不合格原因主要为字段名不完整。'
        elif len(sort_values) > 1 :
            self.my_dict['b1_ms']='其中，'+sort_values.iloc[0,0]+'、'+sort_values.iloc[1,0]+'等系统技术元数据维护情况较差，合格率低于90%，不合格原因主要为字段名不完整。'

    def read_excel_part2_2(self):

        data93_l = self.data[self.data['systemname'] == '已纳管93'].copy()[['jczb001','jczb010','ds']]
        data93_l.rename(columns={'jczb001':'源端数据表数量','jczb010':'业务元数据质量合格率'}, inplace=True)
        data93_l.set_index('ds', inplace=True)
        data93_l = data93_l.transpose()
        data93_l.reset_index(drop=False, inplace=True)
        data93_l=data93_l.assign(
            value1= lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),round(data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)],2))
                    ,(lambda s: s.str.contains('数量'),round((data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])/data93_l[self.get_last_friday(-1)]*100,2))
                ]
            )
        )

        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[['jczb001', 'jczb010', 'ds']]
        data19_l.rename(columns={'jczb001': '已盘点源端数据表数量', 'jczb010': '已盘点业务元数据质量合格率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                     round(data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)],2))
                    , (lambda s: s.str.contains('数量'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        data = pd.concat([data93_l, data19_l], ignore_index=True)

        for x in range (4):
            for y in range (4):
                self.my_dict['b2_t_'+str(x)+'_'+str(y)] = data.iloc[x,y]

        condition = self.get_condition('已纳管核心19')
        data19 = self.data[condition][['systemname', 'jczb010']]
        sort_values = data19[data19['jczb010'] < 90].sort_values(by=['jczb010'], ascending=False)
        length = len(sort_values)
        if length == 0:
            self.my_dict['b2_ms'] = ''
        elif len(sort_values) == 1:
            self.my_dict['b2_ms'] = '其中，' + sort_values.iloc[x, 0] + '系统业务元数据维护情况较差，合格率低于90%，不合格原因主要为XXX。'
        elif len(sort_values) > 1:
            self.my_dict['b2_ms'] = '其中，' + sort_values.iloc[0, 0] + '、' + sort_values.iloc[1, 0] + '等系统业务元数据维护情况较差，合格率低于90%，不合格原因主要为字段名不完整XXX。'

    def read_excel_part2_3(self):

        data93_l = self.data[self.data['systemname'] == '已纳管93'].copy()[['jczb001','jczb012','jczb013','ds']]
        data93_l.rename(columns={'jczb001':'源端数据表数量','jczb012':'源端系统管理元数据质量标签维护率','jczb013':'管理元数据负面清单维护率'}, inplace=True)
        data93_l.set_index('ds', inplace=True)
        data93_l = data93_l.transpose()
        data93_l.reset_index(drop=False, inplace=True)
        data93_l=data93_l.assign(
            value1= lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),round(data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)],2))
                    ,(lambda s: s.str.contains('数量'),round((data93_l[self.get_last_friday()]-data93_l[self.get_last_friday(-1)])/data93_l[self.get_last_friday(-1)]*100,2))
                ]
            )
        )
        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[['jczb001', 'jczb012', 'jczb013', 'ds']]
        data19_l.rename(columns={'jczb001': '已盘点源端系统数据表数量', 'jczb012': '已盘点源端系统管理元数据质量标签维护率',
                                 'jczb013': '已盘点管理元数据负面清单维护率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                     round(data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)],2))
                    , (lambda s: s.str.contains('数量'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        data = pd.concat([data93_l, data19_l], ignore_index=True)
        for x in range (6):
            for y in range (4):
                self.my_dict['b3_t_'+str(x)+'_'+str(y)] = data.iloc[x,y]

        condition = self.get_condition('已纳管核心19')
        data19 = self.data[condition][['systemname', 'jczb010']]
        sort_values = data19[data19['jczb010'] < 90].sort_values(by=['jczb010'], ascending=False)
        length = len(sort_values)
        if length == 0:
            self.my_dict['b3_ms'] = ''
        elif len(sort_values) == 1:
            self.my_dict['b3_ms'] = '其中，' + sort_values.iloc[x, 0] + '系统管理元数据质量标签维护情况较差，合格率低于90%，不合格原因主要为XXX。'
        elif len(sort_values) > 1:
            self.my_dict['b3_ms'] = '其中，' + sort_values.iloc[0, 0] + '、' + sort_values.iloc[1, 0] + '等系统管理元数据质量标签维护情况较差，合格率低于90%，不合格原因主要为字段名不完整XXX。'

    def read_excel_part2_4(self):
        condition = self.get_condition('已纳管核心19',-1)
        data_l = self.data[condition]
        condition = self.get_condition('已纳管核心19')
        data= self.data[condition]
        merge = pd.merge(data, data_l, on='systemname', how='outer')
        add_sys = merge[merge['ds_y'].isnull()].copy()
        self.my_dict['b4_新增纳管系统数'] = len(add_sys)
        if len(add_sys) ==0:
            self.my_dict['b4_新增纳管系统'] = ''
        if len(add_sys) ==1:
            self.my_dict['b4_新增纳管系统'] = add_sys
        change_sys = data[(data['jczb003'] > 0) | (data['jczb004'] > 0) | (data['jczb005'] > 0)][['jczb003', 'jczb004', 'jczb005']]
        self.my_dict['b4_元数据变化系统'] = len(change_sys)
        change_sum = change_sys.sum()
        self.my_dict['b4_新增纳管表'] = change_sum['jczb003']
        self.my_dict['b4_删除纳管表'] = change_sum['jczb004']
        self.my_dict['b4_修改纳管表'] = change_sum['jczb005']


        data_19_l = data[['systemname','jczb001', 'jczb008', 'ds']]
        condition = self.get_condition('已纳管核心19',-1)
        data_19_ll = self.data[condition].copy()[['systemname', 'jczb001', 'jczb008', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[['systemname','jczb001_x', 'jczb008_y','jczb008_x']]
        merge['lv'] = round(merge['jczb008_x']-merge['jczb008_y'],2)
        sort_value = merge.sort_values(by=['lv','jczb001_x'], ascending=False)


        tmp = sort_value[sort_value['jczb008_x'] < 100]
        if len(tmp) == 0:
            self.my_dict['b4_技术元数据质量info2'] ='所有系统技术元数据质量合格率都已达到100%，且保持良好。'
        elif len(tmp) == 1:
            self.my_dict['b4_技术元数据质量info2'] = '其中，除'+tmp.iloc[0,0]+'系统以外,其余系统技术元数据质量合格率都已达到100%，且保持良好。'
        elif len(tmp) == 2:
            self.my_dict['b4_技术元数据质量info2'] = '其中，除'+tmp.iloc[0,0]+'、'+tmp.iloc[1,0]+'系统以外,其余系统技术元数据质量合格率都已达到100%，且保持良好。'
        elif len(tmp) > 2:
            self.my_dict['b4_技术元数据质量info2'] = '其中，除'+tmp.iloc[0,0]+'、'+tmp.iloc[1,0]+'等系统以外,其余系统技术元数据质量合格率都已达到100%，且保持良好。'

        tmp = sort_value[sort_value['lv'] != 0]
        if len(tmp) == 0:
            self.my_dict['b4_技术元数据质量info1'] =''
        elif len(tmp) == 1:
            self.my_dict['b4_技术元数据质量info1'] = tmp.iloc[0,0]+'系统的元数据质量合格率较上周有明显提升，提升为'+str(tmp.iloc[0,4])+'%;'
        elif len(tmp) > 1:
            self.my_dict['b4_技术元数据质量info1'] = tmp.iloc[0,0]+'、'+tmp.iloc[1,0]+'系统的元数据质量合格率较上周有明显提升，分别提升'+str(tmp.iloc[0,4])+'%、'+str(tmp.iloc[1,4])+'%;'


        data19_l = self.data[(self.data['systemname'] == '已纳管核心19')&(self.data['ds'] == self.get_last_friday())][['systemname','jczb001', 'jczb008', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19')&(self.data['ds'] == self.get_last_friday(-1))][['systemname','jczb001', 'jczb008', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[['systemname', 'jczb001_x', 'jczb008_y', 'jczb008_x']]
        merge['lv'] = round(merge['jczb008_x']-merge['jczb008_y'],2)

        data = pd.concat([sort_value, merge], axis=0, sort=False)
        for x in range(20):
            for y in range(5):
                self.my_dict['b4_t_' + str(x) + '_' + str(y)] = data.iloc[x,y]


        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb001', 'jczb010', 'ds']]
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb001', 'jczb010', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb001_x', 'jczb010_y', 'jczb010_x']]
        merge['lv'] = round(merge['jczb010_x'] - merge['jczb010_y'],2)
        sort_value = merge.sort_values(by=['lv', 'jczb001_x'], ascending=False)

        tmp = sort_value[sort_value['jczb010_x'] < 60]
        self.my_dict['b4_业务元数据质量info1'] = tmp.iloc[0, 0] + '、' + tmp.iloc[1, 0]  + '、' + tmp.iloc[2, 0] + '等'

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb001', 'jczb010', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb001', 'jczb010', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb001_x', 'jczb010_y', 'jczb010_x']]
        merge['lv'] =round( merge['jczb010_x'] - merge['jczb010_y'],2)
        data1 = pd.concat([sort_value, merge], axis=0, sort=False)
        for x in range(20):
            for y in range(5):
                self.my_dict['e1_2_t1_' + str(x) + '_' + str(y)] = data1.iloc[x, y]

    def read_excel_part2_5_1(self):
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname','jczb001', 'jczb079', 'ds']]
        condition = self.get_condition('已纳管核心19',-1)
        data_19_ll = self.data[condition].copy()[['systemname', 'jczb001', 'jczb079', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[['systemname','jczb001_x', 'jczb079_y','jczb079_x']]
        merge['lv'] = round(merge['jczb079_x']-merge['jczb079_y'],2)
        sort_value = merge.sort_values(by=['lv','jczb001_x'], ascending=False)

        tmp = sort_value[sort_value['lv'] != 0]
        if len(tmp) == 0:
            self.my_dict['b5_源端表使用率info1'] =''
        elif len(tmp) == 1:
            self.my_dict['b5_源端表使用率info1'] = tmp.iloc[0,0]+'系统的源端表使用率较上周有明显提升，提升为'+str(tmp.iloc[0,4])+'%;'
        elif len(tmp) > 1:
            self.my_dict['b5_源端表使用率info1'] = tmp.iloc[0,0]+'、'+tmp.iloc[1,0]+'系统的源端表使用率较上周有明显提升，分别提升'+str(tmp.iloc[0,4])+'%、'+str(tmp.iloc[1,4])+'%;'

        tmp = sort_value[sort_value['jczb079_x'] < 80]
        if len(tmp) == 0:
            self.my_dict['b5_源端表使用率info2'] = '所有系统技术元数据质量合格率都已达到80%，且保持良好。'
        elif len(tmp) == 1:
            self.my_dict['b5_源端表使用率info2'] = '其中，除' + tmp.iloc[0, 0] + '系统以外,其余系统技术元数据质量合格率都已达到80%，且保持良好。'
        elif len(tmp) == 2:
            self.my_dict['b5_源端表使用率info2'] = '其中，除' + tmp.iloc[0, 0] + '、' + tmp.iloc[
                1, 0] + '系统以外,其余系统技术元数据质量合格率都已达到80%，且保持良好。'
        elif len(tmp) > 2:
            self.my_dict['b5_源端表使用率info2'] = '其中，除' + tmp.iloc[0, 0] + '、' + tmp.iloc[
                1, 0] + '等系统以外,其余系统技术元数据质量合格率都已达到80%，且保持良好。'

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19')&(self.data['ds'] == self.get_last_friday())][['systemname','jczb001', 'jczb079', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19')&(self.data['ds'] == self.get_last_friday(-1))][['systemname','jczb001', 'jczb079', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[['systemname', 'jczb001_x', 'jczb079_y', 'jczb079_x']]
        merge['lv'] = round(merge['jczb079_x']-merge['jczb079_y'],2)

        data = pd.concat([sort_value, merge], axis=0, sort=False)

        for x in range(20):
            for y in range(5):
                self.my_dict['b5_t1_' + str(x) + '_' + str(y)] = data.iloc[x, y]

    def read_excel_part3_1(self):
        # jczb045 质量校核率
        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[
            ['jczb043', 'jczb045', 'jczb054', 'jczb055', 'ds']]
        data19_l.rename(columns={'jczb043': '覆盖质量校核表数量', 'jczb045': '质量校核率', 'jczb054': '技术质量合格表',
                                 'jczb055': '技术质量合格率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                     round(data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)],2))
                    , (lambda s: s.str.contains('表'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        for x in range(4):
            for y in range(4):
                self.my_dict['c1_t_' + str(x) + '_' + str(y)] = data19_l.iloc[x, y]
        lv_change = self.my_dict['c1_t_1_3']
        if lv_change > 0:
            self.my_dict['c1_质量校核率变化'] = '增加'+str(lv_change)+'%'
        elif lv_change < 0:
            self.my_dict['c1_质量校核率变化'] = '减少'+str(abs(lv_change)) +'%'
        else:
            self.my_dict['c1_质量校核率变化'] = '不变'

    def read_excel_part3_1_1(self):
        # 表7.数据表质量校核情况监测
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb043', 'jczb045','ds']]
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb043', 'jczb045', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[['systemname', 'jczb043_x', 'jczb045_y', 'jczb045_x']]
        merge['lv'] = round(merge['jczb045_x'] - merge['jczb045_y'],2)
        sort_value = merge.sort_values(by=['lv', 'jczb043_x'], ascending=False)

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][['systemname', 'jczb043', 'jczb045', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][['systemname', 'jczb043', 'jczb045', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[['systemname', 'jczb043_x', 'jczb045_y', 'jczb045_x']]
        merge['lv'] = round(merge['jczb045_x'] - merge['jczb045_y'],2)

        data = pd.concat([sort_value, merge], axis=0, sort=False)
        for x in range(20):
            for y in range(5):
                self.my_dict['c1_1_t1_' + str(x) + '_' + str(y)] = data.iloc[x, y]



        # 表8.技术校核覆盖率情况监测
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb054', 'jczb023', 'ds']]
        data_19_l['lv']  = round(data_19_l['jczb054']/data_19_l['jczb023']*100, 2)
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb054', 'jczb023', 'ds']]
        data_19_ll['lv']  = round(data_19_ll['jczb054']/data_19_ll['jczb023']*100, 2)
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb054_x', 'lv_y', 'lv_x']]
        merge['lv_sub'] = round(merge['lv_x'] - merge['lv_y'],2)
        sort_value = merge.sort_values(by=['lv_sub', 'jczb054_x'], ascending=False)
        tmp = sort_value[sort_value['lv_sub'] > 0]
        if len(tmp) == 0:
            self.my_dict['c1_技术校核覆盖率info1'] = ''
        elif len(tmp) == 1:
            self.my_dict['c1_技术校核覆盖率info1'] = tmp.iloc[
                                                         0, 0] + '系统的技术校核覆盖率较上周有明显提升，提升为' + str(
                tmp.iloc[0, 4]) + '%;'
        elif len(tmp) > 1:
            self.my_dict['c1_技术校核覆盖率info1'] = tmp.iloc[0, 0] + '、' + tmp.iloc[
                1, 0] + '系统的源技术校核覆盖率较上周有明显提升，分别提升' + str(tmp.iloc[0, 4]) + '%、' + str(
                tmp.iloc[1, 4]) + '%;'

        tmp = sort_value[sort_value['lv_x'] < 60]
        if len(tmp) == 0:
            self.my_dict['c1_技术校核覆盖率info2'] = '所有系统技术校核覆盖率都已达到60%，且保持良好。'
        elif len(tmp) == 1:
            self.my_dict['c1_技术校核覆盖率info2'] = tmp.iloc[0, 0]
        elif len(tmp) == 2:
            self.my_dict['c1_技术校核覆盖率info2'] = tmp.iloc[0, 0] + '、' + tmp.iloc[1, 0]
        elif len(tmp) > 2:
            self.my_dict['c1_技术校核覆盖率info2'] = tmp.iloc[0, 0] + '、' + tmp.iloc[1, 0] + '、' + tmp.iloc[2, 0] + '等'


        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb054', 'jczb023', 'ds']]
        data19_l['lv'] = round(data19_l['jczb054'] / data19_l['jczb023'] * 100, 2)
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb054', 'jczb023', 'ds']]
        data19_ll['lv'] = round(data19_ll['jczb054'] / data19_ll['jczb023'] * 100, 2)


        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb054_x', 'lv_y', 'lv_x']]
        merge['lv_sub'] = round(merge['lv_x'] - merge['lv_y'],2)
        data = pd.concat([sort_value, merge], axis=0, sort=False)
        for x in range(20):
            for y in range(5):
                self.my_dict['c1_1_t2_' + str(x) + '_' + str(y)] = data.iloc[x, y]
        # 表9.技术校核覆盖率下的观察指标情况监测
        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[
            ['jczb043', 'jczb033', 'jczb036','jczb037','jczb038', 'ds']]
        data19_l['无需校核占比']=round((data19_l['jczb043']-data19_l['jczb033'])/data19_l['jczb043']*100,2)
        data19_l['主键唯一校核覆盖率']=round(data19_l['jczb036']/data19_l['jczb043']*100,2)
        data19_l['必填字段校核覆盖率']=round(data19_l['jczb037']/data19_l['jczb043']*100,2)
        data19_l['码值合规校核覆盖率']=round(data19_l['jczb038']/data19_l['jczb043']*100,2)
        data19_l.rename(columns={'jczb043': '技术校核表数量（张）', 'jczb036': '主键唯一校核表数量（张）', 'jczb037': '必填字段完整校核表数量（张）',
                                 'jczb038': '码值合规校核表数量（张）'}, inplace=True)

        data19_l=data19_l[['无需校核占比','主键唯一校核覆盖率','必填字段校核覆盖率','码值合规校核覆盖率','技术校核表数量（张）','主键唯一校核表数量（张）','必填字段完整校核表数量（张）','码值合规校核表数量（张）','ds']]
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l.reset_index(drop=False, inplace=True)
        data19_l = data19_l.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率') | s.str.contains('比'),
                     round(data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)], 2))
                    , (lambda s: s.str.contains('表'), round(
                    (data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )

        tmp = data19_l[data19_l['value1'] < 0]
        self.my_dict['c1_技术校核覆盖率info4'] =''
        for x in range(len(tmp)):
            self.my_dict['c1_技术校核覆盖率info4']=self.my_dict['c1_技术校核覆盖率info4'] +tmp.iloc[x,0]+'较上周环比下降'+str(abs(tmp.iloc[x,3]))+'%,'
        tmp = data19_l[data19_l['value1'] > 0]
        self.my_dict['c1_技术校核覆盖率info3'] = ''
        for x in range(len(tmp)):
            self.my_dict['c1_技术校核覆盖率info3'] = self.my_dict['c1_技术校核覆盖率info3'] + tmp.iloc[x, 0] + '较上周环比提升' + str(tmp.iloc[x, 3]) + '%,'

        for x in range(8):
            for y in range(4):
                self.my_dict['c1_1_t3_' + str(x) + '_' + str(y)] = data19_l.iloc[x, y]

        # 表10. 业务校核覆盖率情况监测
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb092', 'jczb023', 'ds']]
        data_19_l['lv'] = round(data_19_l['jczb092'] / data_19_l['jczb023'] * 100, 2)
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb092', 'jczb023', 'ds']]
        data_19_ll['lv'] = round(data_19_ll['jczb092'] / data_19_ll['jczb023'] * 100, 2)
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb092_x', 'lv_y', 'lv_x']]
        merge['lv_sub'] = merge['lv_x'] - merge['lv_y']
        sort_value = merge.sort_values(by=['lv_sub', 'jczb092_x'], ascending=False)

        tmp = sort_value[sort_value['lv_sub'] > 0]
        if len(tmp) == 0:
            self.my_dict['c1_业务校核覆盖率info1'] = ''
        elif len(tmp) == 1:
            self.my_dict['c1_业务校核覆盖率info1'] = tmp.iloc[0, 0] + '系统的业务校核覆盖率较上周略有提升，提升为' + str(
                tmp.iloc[0, 4]) + '%;'
        elif len(tmp) > 1:
            self.my_dict['c1_业务校核覆盖率info1'] = tmp.iloc[0, 0] + '、' + tmp.iloc[
                1, 0] + '系统的业务校核覆盖率较上周略有提升，分别提升' + str(tmp.iloc[0, 4]) + '%、' + str(
                tmp.iloc[1, 4]) + '%;'

        tmp = sort_value[sort_value['lv_sub'] < 60]
        if len(tmp) == 0:
            self.my_dict['c1_业务校核覆盖率info2'] = '所有系统业务校核覆盖率都已达到60%，且保持良好。'
        elif len(tmp) == 1:
            self.my_dict['c1_业务校核覆盖率info2'] = tmp.iloc[0, 0]
        elif len(tmp) == 2:
            self.my_dict['c1_业务校核覆盖率info2'] = tmp.iloc[0, 0] + '、' + tmp.iloc[1, 0]
        elif len(tmp) > 2:
            self.my_dict['c1_业务校核覆盖率info2'] = tmp.iloc[0, 0] + '、' + tmp.iloc[1, 0] + '、' + tmp.iloc[1, 0] + '等'

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb092', 'jczb023', 'ds']]
        data19_l['lv'] = round(data19_l['jczb092'] / data19_l['jczb023'] * 100, 2)
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb092', 'jczb023', 'ds']]
        data19_ll['lv'] = round(data19_ll['jczb092'] / data19_ll['jczb023'] * 100, 2)
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb092_x', 'lv_y', 'lv_x']]
        merge['lv_sub'] = merge['lv_x'] - merge['lv_y']
        data = pd.concat([sort_value, merge], axis=0, sort=False)
        for x in range(20):
            for y in range(5):
                self.my_dict['c1_1_t4_' + str(x) + '_' + str(y)] = data.iloc[x, y]

        # 表11.业务校核覆盖率下的观察指标情况监测
        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[
            ['jczb093', 'jczb095', 'jczb096', 'jczb097', 'jczb098', 'jczb099', 'ds']]
        data19_l.rename(columns={'jczb093': '业务校核表数量（张）', 'jczb095': '校核完整性表数量（张）', 'jczb096': '校核一致性表数量（张）',
                                 'jczb097': '校核准确性表数量（张）','jczb098': '校核有效性表数量（张）','jczb099': '校核唯一性表数量（张）'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l['lv']=round((data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[self.get_last_friday(-1)] * 100, 2)
        data19_l.reset_index(drop=False, inplace=True)
        for x in range(6):
            for y in range(4):
                self.my_dict['c1_1_t5_' + str(x) + '_' + str(y)] = data19_l.iloc[x, y]
        tmp = data19_l[data19_l['lv'] < 0]
        self.my_dict['c1_业务校核覆盖率info4'] = ''
        for x in range(len(tmp)):
            self.my_dict['c1_业务校核覆盖率info4'] = self.my_dict['c1_业务校核覆盖率info4'] + tmp.iloc[
                x, 0] + '较上周环比下降' + str(abs(tmp.iloc[x, 3])) + '%,'
        tmp = data19_l[data19_l['lv'] > 0]
        self.my_dict['c1_业务校核覆盖率info3'] = ''
        for x in range(len(tmp)):
            self.my_dict['c1_业务校核覆盖率info3'] = self.my_dict['c1_业务校核覆盖率info3'] + tmp.iloc[
                x, 0] + '较上周环比提升' + str(tmp.iloc[x, 3]) + '%,'

        print(self.my_dict['c1_业务校核覆盖率info3'])

        print(self.my_dict['c1_业务校核覆盖率info4'])

    def read_excel_part3_2(self):

        # 表13.数据表技术校核质量合格情况监测
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb054', 'jczb055', 'ds']]
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb054', 'jczb055', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb054_x', 'jczb055_y', 'jczb055_x']]
        merge['lv'] = merge['jczb055_x'] - merge['jczb055_y']
        sort_value = merge.sort_values(by=['lv', 'jczb054_x'], ascending=False)
        sort_value2 = merge.sort_values(by=['jczb054_x'], ascending=False)
        sort_len=len(sort_value2)
        self.my_dict['c2_技术质量合格率较低']=sort_value2.iloc[sort_len-1,0]+'、'+sort_value2.iloc[sort_len-2,0]+'、'+sort_value2.iloc[sort_len-3,0]+'、'+sort_value2.iloc[sort_len-4,0]+'、'+sort_value2.iloc[sort_len-5,0]

        tmp = sort_value[sort_value['lv'] > 0]
        if len(tmp) == 0:
            self.my_dict['c2_技术质量合格率info1'] = ''
        elif len(tmp) == 1:
            self.my_dict['c2_技术质量合格率info1'] = tmp.iloc[0, 0] + '系统的技术质量合格率较上周有明显提升，提升为' + str(
                tmp.iloc[0, 4]) + '%;'
        elif len(tmp) > 1:
            self.my_dict['c2_技术质量合格率info1'] = tmp.iloc[0, 0] + '、' + tmp.iloc[
                1, 0] + '系统的技术质量合格率较上周有明显提升，分别提升' + str(tmp.iloc[0, 4]) + '%、' + str(
                tmp.iloc[1, 4]) + '%;'

        tmp = sort_value[sort_value['lv'] < 0]
        if len(tmp) == 0:
            self.my_dict['c2_技术质量合格率info2'] = ''
        elif len(tmp) == 1:
            self.my_dict['c2_技术质量合格率info2'] = tmp.iloc[0, 0]+'系统的技术质量合格率较上周下降，原因为xx'
        elif len(tmp) == 2:
            self.my_dict['c2_技术质量合格率info2'] = tmp.iloc[0, 0] + '、' + tmp.iloc[1, 0]+'系统的技术质量合格率较上周下降，原因为xx'
        elif len(tmp) > 2:
            self.my_dict['c2_技术质量合格率info2'] = tmp.iloc[0, 0] + '、' + tmp.iloc[1, 0] + '、' + tmp.iloc[2, 0] + '等系统的技术质量合格率较上周下降，原因为xx'

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb054', 'jczb055', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb054', 'jczb055', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb054_x', 'jczb055_y', 'jczb055_x']]
        merge['lv'] = merge['jczb055_x'] - merge['jczb055_y']

        data = pd.concat([sort_value, merge], axis=0, sort=False)
        for x in range(20):
            for y in range(5):
                self.my_dict['c2_t2_' + str(x) + '_' + str(y)] = data.iloc[x, y]

        # 表14.技术校核覆盖率下的观察指标情况监测

        # 表15.数据表业务校核质量合格情况监测
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb093', 'jczb094', 'ds']]
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb093', 'jczb094', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb093_x', 'jczb094_y', 'jczb094_x']]
        merge['lv'] = merge['jczb094_x'] - merge['jczb094_y']
        sort_value = merge.sort_values(by=['lv', 'jczb093_x'], ascending=False)

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb093', 'jczb094', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb093', 'jczb094', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb093_x', 'jczb094_y', 'jczb094_x']]
        merge['lv'] = merge['jczb094_x'] - merge['jczb094_y']

        data = pd.concat([sort_value, merge], axis=0, sort=False)
        for x in range(20):
            for y in range(5):
                self.my_dict['c2_t4_' + str(x) + '_' + str(y)] = data.iloc[x, y]
        # 表16.技术校核覆盖率下的观察指标情况监测
        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[
            ['jczb093', 'jczb095', 'jczb096', 'jczb097', 'jczb098', 'jczb100', 'jczb101', 'jczb102', 'jczb103', 'ds']]
        data19_l.rename(columns={'jczb093': '业务校核质量合格表数量（张）', 'jczb095': '完整性合格表数量（张）',
                                 'jczb096': '一致性合格表数量（张）',
                                 'jczb097': '准确性合格表数量（张）', 'jczb098': '有效性合格表数量（张）',
                                 'jczb100': '完整性平均值','jczb101': '一致性平均值', 'jczb102': '准确性平均值', 'jczb103': '有效性平均值'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l['lv'] = round((data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]) / data19_l[
            self.get_last_friday(-1)] * 100, 2)
        data19_l.reset_index(drop=False, inplace=True)
        for x in range(9):
            for y in range(4):
                self.my_dict['c2_t5_' + str(x) + '_' + str(y)] = data19_l.iloc[x, y]

    def read_excel_part4_1(self):
        data19_l = self.data[(self.data['systemname'] == '已纳管核心19')&(self.data['ds'] == self.get_last_friday())]
        self.my_dict['d1_接入中台表数量'] = data19_l['jczb023'].iloc[0]
        self.my_dict['d1_接入核心表数量'] = data19_l['jczb026'].iloc[0]
        data19 = self.data[self.data['systemname'] == '已纳管核心19'].copy()[
            ['jczb072', 'jczb073', 'jczb063', 'jczb065', 'jczb066', 'jczb067', 'jczb070', 'jczb071', 'jczb073',
             'jczb075', 'ds']]
        data19.rename(columns={'jczb072': '共享表使用数量(贴源层)', 'jczb073': '共享使用率(贴源层)',
                               'jczb063': '核心表使用率',
                               'jczb065': '按需表使用率', 'jczb066': '共享层表数量',
                               'jczb067': '共享层核心表数量', 'jczb070': '共享层授权表数量',
                               'jczb071': '共享层核心表授权数量',
                               'jczb073': '共享层表使用率', 'jczb075': '共享层核心表使用率'}, inplace=True)
        data19.set_index('ds', inplace=True)
        data19 = data19.transpose()
        data19.reset_index(drop=False, inplace=True)
        data19 = data19.assign(
            value1=lambda df: df['index'].case_when(
                [
                    (lambda s: s.str.contains('率'),
                    round( data19[self.get_last_friday()] - data19[self.get_last_friday(-1)],2))
                    , (lambda s: s.str.contains('数量'), round(
                    (data19[self.get_last_friday()] - data19[self.get_last_friday(-1)]) / data19[
                        self.get_last_friday(-1)] * 100, 2))
                ]
            )
        )
        for x in range(10):
            for y in range(4):
                self.my_dict['d1_t1_' + str(x) + '_' + str(y)] = data19.iloc[x, y]

    def read_excel_part4_4(self):
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb067', 'jczb075', 'ds']]
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb067', 'jczb075', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb067_x', 'jczb075_y', 'jczb075_x']]
        merge['lv'] = merge['jczb075_x'] - merge['jczb075_y']
        sort_value = merge.sort_values(by=['lv', 'jczb067_x'], ascending=False)

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb067', 'jczb075', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb067', 'jczb075', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb067_x', 'jczb075_y', 'jczb075_x']]
        merge['lv'] = merge['jczb075_x'] - merge['jczb075_y']

        data = pd.concat([sort_value, merge], axis=0, sort=False)

        for x in range(20):
            for y in range(5):
                self.my_dict['d4_t1_' + str(x) + '_' + str(y)] = data.iloc[x, y]

    def read_excel_part5_1(self):
        data19_l = self.data[self.data['systemname'] == '已纳管核心19'].copy()[
            ['jczb008', 'jczb010', 'jczb012', 'jczb013',  'ds']]
        data19_l.rename(columns={'jczb008': '技术元数据合格率', 'jczb010': '业务元数据合格率',
                                 'jczb012': '管理元数据-质量标签维护率',
                                 'jczb013': '管理元数据-负面清单维护率'}, inplace=True)
        data19_l.set_index('ds', inplace=True)
        data19_l = data19_l.transpose()
        data19_l['lv'] = data19_l[self.get_last_friday()] - data19_l[self.get_last_friday(-1)]
        data19_l = data19_l[[self.get_last_friday(-1), self.get_last_friday(), 'lv']]
        data19_l.reset_index(drop=False, inplace=True)
        for x in range(4):
            for y in range(4):
                self.my_dict['e1_t1_' + str(x) + '_' + str(y)] = data19_l.iloc[x, y]

    def read_excel_part5_1_1(self):
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb001', 'jczb008', 'ds']]
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb001', 'jczb008', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb001_x', 'jczb008_y', 'jczb008_x']]
        merge['lv'] = merge['jczb008_x'] - merge['jczb008_y']
        sort_value = merge.sort_values(by=['lv', 'jczb001_x'], ascending=False)

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb001', 'jczb008', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb001', 'jczb008', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb001_x', 'jczb008_y', 'jczb008_x']]
        merge['lv'] = merge['jczb008_x'] - merge['jczb008_y']
        data = pd.concat([sort_value, merge], axis=0, sort=False)


        for x in range(20):
            for y in range(5):
                self.my_dict['e1_1_t1_' + str(x) + '_' + str(y)] = data.iloc[x, y]

    def read_excel_part5_1_2(self):
        condition = self.get_condition('已纳管核心19')
        data_19_l = self.data[condition][['systemname', 'jczb001', 'jczb010', 'ds']]
        condition = self.get_condition('已纳管核心19', -1)
        data_19_ll = self.data[condition][['systemname', 'jczb001', 'jczb010', 'ds']]
        merge = pd.merge(data_19_l, data_19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb001_x', 'jczb010_y', 'jczb010_x']]
        merge['lv'] = merge['jczb010_x'] - merge['jczb010_y']
        sort_value = merge.sort_values(by=['lv', 'jczb001_x'], ascending=False)

        data19_l = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday())][
            ['systemname', 'jczb001', 'jczb010', 'ds']]
        data19_ll = self.data[(self.data['systemname'] == '已纳管核心19') & (self.data['ds'] == self.get_last_friday(-1))][
            ['systemname', 'jczb001', 'jczb010', 'ds']]
        merge = pd.merge(data19_l, data19_ll, how='left', on=['systemname'])[
            ['systemname', 'jczb001_x', 'jczb010_y', 'jczb010_x']]
        merge['lv'] = merge['jczb010_x'] - merge['jczb010_y']
        data = pd.concat([sort_value, merge], axis=0, sort=False)

        for x in range(20):
            for y in range(5):
                self.my_dict['e1_2_t1_' + str(x) + '_' + str(y)] = data.iloc[x, y]

    def run (self):
        self.read_excel_part1()
        self.read_excel_part2_1()
        self.read_excel_part2_2()
        self.read_excel_part2_3()
        self.read_excel_part2_4()
        self.read_excel_part2_5_1()
        self.read_excel_part3_1()
        self.read_excel_part3_1_1()
        self.read_excel_part3_2()
        self.read_excel_part4_1()
        self.read_excel_part4_4()
        self.read_excel_part5_1()
        self.read_excel_part5_1_1()
        self.read_excel_part5_1_2()

    def sout_dict(self):
        for key, value in self.my_dict.items():
            try:
                self.my_dict[key] = round(float(value), 2)
            except ValueError:
                self.my_dict[key] = value
        for key, value in self.my_dict.items():
            print(f"{key}: {value}")
    def write(self):
        doc = DocxTemplate(self.word_url)
        doc.render(self.my_dict)
        doc.save(self.write_url)
