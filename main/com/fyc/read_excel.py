from datetime import datetime, timedelta
import pandas as pd

from src.main.com.fyc.wrrite_docx import write_docx

my_dict={}

def get_last_friday(param='0'):
    # 直接计算上周五
    last_friday = datetime.today() - timedelta(days=(datetime.today().weekday() + 3) % 7)

    # 直接计算上上周五
    last_last_friday = last_friday - timedelta(7)

    if(param == -1):
        return int(last_last_friday.strftime('%Y%m%d'))
    return int(last_friday.strftime('%Y%m%d'))


def read_excel():
    # df = pd.read_excel('C:\\Users\\ycf\\Desktop\\write_docx\\data.xlsx', sheet_name='Sheet1',header=0)
    df = pd.read_excel('C:\\Users\\Administrator\\Desktop\\write_docx\\data.xlsx', sheet_name='Sheet1',header=0)
    # a_源端系统
    my_dict['a_源端系统']=200
    # a_纳管源端系统
    condition = (df['ds'] == get_last_friday() ) & (df['sys_status'].str.contains('已纳管'))
    tmp = len(df[condition])
    my_dict['a_纳管源端系统']=tmp
    # a_纳管率
    my_dict['a_纳管率']=round(tmp/my_dict['a_源端系统']*100,2)
    print(my_dict)
    write_docx(my_dict)

read_excel()