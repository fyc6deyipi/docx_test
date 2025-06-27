from datetime import datetime, timedelta
import pandas as pd

from src.main.com.fyc.wrrite_docx import write_docx

my_dict={}

def sout_dict(my_dict):
    for key, value in my_dict.items():
        print(f"{key}: {value}")

def get_last_friday(param='0'):
    # 直接计算上周五
    last_friday = datetime.today() - timedelta(days=(datetime.today().weekday() + 3) % 7)

    # 直接计算上上周五
    last_last_friday = last_friday - timedelta(7)

    if(param == -1):
        return int(last_last_friday.strftime('%Y%m%d'))
    return int(last_friday.strftime('%Y%m%d'))


def read_excel(my_dict):
    # df = pd.read_excel('C:\\Users\\ycf\\Desktop\\write_docx\\data.xlsx', sheet_name='Sheet1',header=0)
    df = pd.read_excel('C:\\Users\\Administrator\\Desktop\\write_docx\\data.xlsx', sheet_name='Sheet1',header=0)


    # 1源端系统全量纳管和动态监测
    # a_源端系统1
    my_dict['a_源端系统']=200
    # a_纳管源端系统
    condition = (df['ds'] == get_last_friday() ) & (df['sys_status'].str.contains('已纳管'))
    a_纳管源端系统 = df[condition]
    my_dict['a_纳管源端系统']=len(a_纳管源端系统)
    # a_纳管率
    my_dict['a_纳管率']=round(my_dict['a_纳管源端系统']/my_dict['a_源端系统']*100,2)


    # 2重点针对
    #a_纳管源端系统表数量
    a_纳管源端系统表数量 = a_纳管源端系统['jczb001'].sum()/10000
    my_dict['a_纳管源端系统表数量'] = round(a_纳管源端系统表数量,2)
    #a_技术元数据合格率 jczb008
    a_技术元数据合格率 = a_纳管源端系统['jczb008'].min()
    my_dict['a_技术元数据合格率'] = round(a_技术元数据合格率,2)
    # a_技术元数据合格率低于99 jczb008
    a_技术元数据合格率低于99 = a_纳管源端系统[a_纳管源端系统['jczb008'] < 99]
    my_dict['a_技术元数据合格率低于99'] = len(a_技术元数据合格率低于99)

    # 3系统业务梳理
    #a_已盘点核心系统
    condition = (a_纳管源端系统['sys_status'] == '已纳管已盘点')&(a_纳管源端系统['is_core_sys'] == '是')
    a_已盘点核心系统 = a_纳管源端系统[condition]
    my_dict['a_已盘点核心系统']=len(a_已盘点核心系统)
    #a_已盘点系统功能模块 jczb018
    a_已盘点核心系统功能模块 = a_已盘点核心系统['jczb018'].sum()
    my_dict['a_已盘点核心系统功能模块'] = int(a_已盘点核心系统功能模块)
    # a_已盘点核心系统核心表 jczb019
    a_已盘点核心系统核心表 = a_已盘点核心系统['jczb019'].sum()
    my_dict['a_已盘点核心系统核心表'] = int(a_已盘点核心系统核心表)
    # a_已盘点核心系统主数据表 jczb020
    a_已盘点核心系统主数据表 = a_已盘点核心系统['jczb020'].sum()
    my_dict['a_已盘点核心系统主数据表'] = int(a_已盘点核心系统主数据表)
    # a_已盘点核心系统核心业务数据表 jczb021
    a_已盘点核心系统核心业务数据表 = a_已盘点核心系统['jczb021'].sum()
    my_dict['a_已盘点核心系统核心业务数据表'] = int(a_已盘点核心系统核心业务数据表)
    #a_已盘点核心系统核心表占比
    my_dict['a_已盘点核心系统核心表占比']=round(my_dict['a_已盘点核心系统核心表']/a_已盘点核心系统['jczb001'].sum()*100,2)


    # 4核心数据接入
    # a_接入中台表 jczb023
    a_接入中台表 = a_已盘点核心系统['jczb023'].sum()
    my_dict['a_接入中台表'] = int(a_接入中台表)
    # a_接入核心表 jczb026
    a_接入核心表 = a_已盘点核心系统['jczb026'].sum()
    my_dict['a_接入核心表'] = int(a_接入核心表)
    # a_按需接入表 jczb031
    a_按需接入表 = a_已盘点核心系统['jczb031'].sum()
    my_dict['a_按需接入表'] = int(a_按需接入表)
    # a_接入核心表占比
    my_dict['a_接入核心表占比'] = round(my_dict['a_接入核心表']/my_dict['a_接入中台表']*100,2)
    # a_质量校核表 jczb043
    a_质量校核表 = a_已盘点核心系统['jczb043'].sum()
    my_dict['a_质量校核表'] = int(a_质量校核表)
    # a_质量校核率
    my_dict['a_质量校核率'] = round(my_dict['a_质量校核表'] / my_dict['a_接入中台表'] * 100, 2)
    # a_技术质量合格表 jczb054
    a_技术质量合格表 = a_已盘点核心系统['jczb054'].sum()
    my_dict['a_技术质量合格表'] = int(a_技术质量合格表)
    # a_技术质量合格率
    my_dict['a_技术质量合格率'] = round(my_dict['a_技术质量合格表'] / my_dict['a_质量校核表'] * 100, 4)
    # a_业务质量合格表 jczb093
    a_业务质量合格表 = a_已盘点核心系统['jczb093'].sum()
    my_dict['a_业务质量合格表'] = int(a_业务质量合格表)
    # a_业务质量合格率
    my_dict['a_业务质量合格率'] = round(my_dict['a_业务质量合格表'] / my_dict['a_质量校核表'] * 100, 4)


    #5 数据共享和服务
    #a_共享层表数量 jczb066
    # a_共享层表数量=df['jczb066'].sum()
    # my_dict['a_共享层表数量'] = int(a_共享层表数量)


    #6 元数据和血缘信息维护
    #a_技术元数据质量合格表 jczb007
    a_技术元数据质量合格表 = a_已盘点核心系统['jczb007'].sum()
    my_dict['a_技术元数据质量合格表'] = int(a_技术元数据质量合格表)
    #a_技术元数据质量合格率 jczb008   纳管表数量（表） jczb089
    a_纳管表数量 = a_已盘点核心系统['jczb089'].sum()
    my_dict['a_技术元数据质量合格率'] = round(my_dict['a_技术元数据质量合格表'] / a_纳管表数量 * 100, 2)
    # a_业务元数据质量合格表 jczb009
    a_业务元数据质量合格表 = a_已盘点核心系统['jczb009'].sum()
    my_dict['a_业务元数据质量合格表'] = int(a_业务元数据质量合格表)
    # a_业务元数据质量合格率
    my_dict['a_业务元数据质量合格率'] = round(my_dict['a_业务元数据质量合格表'] / a_纳管表数量 * 100, 2)
    #  a_管理元数据质量标签合格表 管理元数据质量标签合格表数量  jczb011  需人工配置质量规则校核表数量  jczb033
    a_管理元数据质量标签合格表 =  a_已盘点核心系统['jczb011'].sum()
    a_需人工配置质量规则校核表 =  a_已盘点核心系统['jczb033'].sum()
    my_dict['a_管理元数据质量标签维护率'] = round(a_管理元数据质量标签合格表/a_需人工配置质量规则校核表 * 100, 2)
    #a_93接入贴源层表
    a_93接入贴源层表 = a_纳管源端系统['jczb023'].sum()
    my_dict['a_93接入贴源层表'] = int(a_93接入贴源层表)
    #a_93源端至共享层有血缘的源端表数量 jczb078
    a_93源端至共享层有血缘的源端表数量 = a_纳管源端系统['jczb078'].sum()
    my_dict['a_93源端至共享层有血缘的源端表数量'] = int(a_93源端至共享层有血缘的源端表数量)
    #a_93源端表使用率
    my_dict['a_93源端表使用率'] = round(my_dict['a_93源端至共享层有血缘的源端表数量']/my_dict['a_93接入贴源层表']*100,2)
    #a_93共享层一级系统表 jczb076
    a_93共享层一级系统表 = a_纳管源端系统['jczb076'].sum()
    my_dict['a_93共享层一级系统表'] = int(a_93共享层一级系统表)
    #a_93共享层表血缘覆盖率
    my_dict['a_93共享层表血缘覆盖率'] = round(my_dict['a_93源端至共享层有血缘的源端表数量']/a_93共享层一级系统表*100,2)



    sout_dict(my_dict)
    return my_dict



write_docx(read_excel(my_dict))
# read_excel(my_dict)