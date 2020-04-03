# -*- coding: utf-8 -*-
"""
更新Ave.
"""

import os
import time
import pandas as pd
import xlwings as xw
from datetime import datetime, timedelta, date


ADDR = r'H:\SZ_数据\Download\Ave'

now = lambda : time.perf_counter()

dat = lambda x: datetime.strptime(str(date.today() - timedelta(x)), '%Y-%m-%d')

def connect():
    from sqlalchemy import create_engine
    ss = "mssql+pymssql://%s:%s@%s:%s/%s"
    try:
        engine = create_engine(ss % loginInfo())
    except Exception as e:
        print("连接失败: %s" % e)
        raise
    else:
        return engine

def loginInfo(section='SQL Server'):
    from configparser import ConfigParser
    conf = ConfigParser()
    path = r'C:\User\chen.huaiyu\Chinasearch\c.s.conf'
    conf.read(path)
    if os.path.exists(path):
        usr = conf.get(section, 'accountname')
        pw = conf.get(section, 'password')
        ip = conf.get(section, 'ip')
        port = conf.get(section, 'port')
        db = conf.get(section, 'dbname')
        return usr, pw, ip, port, db
    else:
        raise FileNotFoundError("文件不存在：%s" % path)

def getQ(days=1):
    '''
    根据日期转换为季度

    Parameters
    ----------
    days : TYPE, optional
        DESCRIPTION. The default is 1.

    Returns
    -------
    str
        季度
    str
        英文表述月度区间

    '''
    # Default: Yesterday
    m = dat(days).month
    if m in (1, 2, 3):
        return 'Q1', 'Jan to Mar'
    elif m in (4, 5, 6):
        return 'Q2', 'Apr to Jun'
    elif m in (7, 8, 9):
        return 'Q3', 'Jul to Sep'
    else:
        return 'Q4', 'Oct to Dec'

def getPath(days=1):
    '''
    获取文件路径
    
    Parameters
    ----------
    days : TYPE, optional
        DESCRIPTION. The default is 1.

    Returns
    -------
    path1 : TYPE
        目标文件绝对路径

    '''
    # Default: Yesterday
    str1 = 'Ave.workday&weekday'
    str2 = '.xlsx'
    name = (str1 + getQ()[0] + '(' + str(dat(days).year) + ' ' + 
            getQ()[1] + ')' + dat(days).strftime('%Y.%m.%d') + str2)
    path1 = os.path.join(ADDR, name)
    if os.path.isfile(path1):
        return path1
    else:
        # 查看往前推2-14天的文件是否存在
        for n in range(2, 15):
            name = (str1 + getQ(n)[0] + '(' + str(dat(n).year) + 
                    ' ' + getQ(n)[1] + ')' + 
                    dat(n).strftime('%Y.%m.%d') + str2)
            path2 = os.path.join(ADDR, name)
            if os.path.isfile(path2):
                os.rename(path2, path1)
                return path1

def col(conn):
    sql = ''' select * from information_schema.columns
                where table_name=%s
        '''
    cols = [i[3] for i in conn.execute(sql).fetchall()]
    return cols

def upBasicInfo(conn, sht):
    sql = 'select * from basicInfo'
    data = list(map(lambda x: list(x), conn.execute(sql).fetchall()))
    df = pd.DataFrame(data, columns=col(conn))
    df.sort_values(by='Id', inplace=True)
    # 写入excel
    sht['A3'].value = df.vlaues
    
def upSpending(conn, sht, datSt, datEnd, class_):
    sql = '''select Id, b.用户名, s.日期, s.sum_ 
            from basicInfo b
            left join 
                (select 日期, 用户名, 类别, sum(金额) as sum_
                    from 消费 
                    where 日期 between %s and %s
                      and 类别=%s
                    group by 日期, 用户名, 类别
                ) s
              on b.用户名=s.用户名
            ''' % (datSt, datEnd, class_)
    data = list(map(lambda x: list(), conn.execute(sql).fetchall()))
    cols = ['Id', '用户名', '日期', '金额']
    df = pd.DataFrame(data, columns=cols)
    df.sort_values(by='Id', inplace=True)
    df.fillna(0, inplace=True)
    df = df.pivot_table(values=['金额'], columns=['日期'], index=['Id', '用户名'])
    # 写入excel
    lis = sht['A2:FA2'].value
    for dat in pd.date_range(datSt, datEnd):
        n = lis.index(dat)
        sht[2, n].options(transpose=True).value = df[dat].values

def getSpending():
    
    pass

def fillFormula(sht, days, rows):
    from xlwings import constants
    # 更新新户后 行数
    rows2 = sht['A1'].current_region.rows.count
    # 汇总区域
    rng1 = sht['AE' + str(rows) + ':BD' + str(rows)]
    rng2 = sht['AE' + str(rows) + ':BD' + str(rows2)]
    rng1.api.AutoFill(rng2.api, constants.AutoFillType.xlFillCopy)
    # 预估区
    rng3 = sht['FB' + str(rows) + ':PB' + str(rows)]
    rng4 = sht['FB' + str(rows) + ':PB' + str(rows2)]
    rng3.api.AutoFill(rng4.api, constants.AutoFillType.xlFillCopy)
    # 消费区
    lis = sht['A2:FA2'].value
    cols = lis.index(dat(days))
    sht[rows+1:rows2, 26:cols].value = 0
    

def forecast():
    pass

def toZero():
    pass

def main():
    '''
    主函数

    Returns
    -------
    None.

    '''
    print("默认：至昨天")
    days = input("从多少天以前？(比如：1)")
    wb = xw.Book(getPath())
    with connect().begin() as conn:
        conn.execute('select 1')
        for shtName in ['搜索', '其他新产品', '原生广告']:
            sht = wb.sheets[shtName]
            # 原表行列
            rows = sht['A1'].current_region.rows.count
            # 更新基本信息 
            upBasicInfo(conn, sht)
            # 更新消费
            upSpending(conn, sht, dat(days), dat(1), shtName)
            # 填充公式列
            fillFormula(sht, days, rows)
            
            
    # 预估均值公式
    forecast()
    # 近n日消费为0，预估均值 -> 0
    toZero()

if __name__ == '__main__':
    st = now()
    print('默认昨日')
    main()
    print(now() - st)
    