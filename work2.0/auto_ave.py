# -*- coding: utf-8 -*-
"""
更新Ave.
"""

import os
import time
import pandas as pd
import xlwings as xw
from xlwings import constants
from datetime import datetime, timedelta, date


ADDR = r'H:\SZ_数据\Download\Ave'

now = lambda : time.perf_counter()

dat = lambda x: datetime.strptime(str(date.today() - timedelta(x)), '%Y-%m-%d')

def connect():
    '''
    连接 MSSQL 数据库
    
    Returns
    -------
    engine

    '''
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
    '''
    从配置文件中获取数据库登陆信息

    Parameters
    ----------
    section : The default is 'SQL Server'.

    Raises
    ------
    FileNotFoundError
        DESCRIPTION.

    Returns
    -------
    usr : 用户名
    pw : 密码
    ip : 服务器ip
    port : 端口
    db : 访问数据库.

    '''
    from configparser import ConfigParser
    conf = ConfigParser()
    path = r'H:\SZ_数据\Python\c.s.conf'
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
    根据日期转换为对应季度

    Parameters
    ----------
    days : The default is 1.即昨日

    Returns
    -------
    (季度，英文表述月度区间)

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
    获取xlsx文件路径
    
    Parameters
    ----------
    days : The default is 1.即昨日

    Returns
    -------
    path1 : xlsx文件绝对路径

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

def col(conn, tableName):
    '''
    从数据库获取对应数据表字段序列

    Parameters
    ----------
    conn : 连接数据库。
    tableName : 数据库中表名

    Returns
    -------
    cols : 字段列表

    '''
    sql = ''' select * from information_schema.columns
                where table_name='%s'
        ''' % tableName
    cols = [i[3] for i in conn.execute(sql).fetchall()]
    return cols

def upBasicInfo(conn, sht):
    '''
    从数据库获取信息，并更新xlsx中的基本信息

    Parameters
    ----------
    conn : 数据库连接
    sht : xlsx 中指定更新的工作表

    Returns
    -------
    None.

    '''
    sql = 'select * from basicInfo'
    data = list(map(lambda x: list(x), conn.execute(sql).fetchall()))
    df = pd.DataFrame(data, columns=col(conn, 'basicInfo'))
    df.sort_values(by='Id', inplace=True)
    df.drop(columns=['Id'], inplace=True)
    # 写入excel
    sht['A3'].value = df.values
    
def upSpending(conn, sht, datSt, datEnd, class_):
    '''
    更新 xlsx 中指定的工作表

    Parameters
    ----------
    conn : 数据库
    sht : xlsx 中指定的工作表
    datSt : 待更新的消费的起始日
    datEnd : 待更新的消费的终止日
    class_ : 消费产品类型

    Returns
    -------
    None.

    '''
    with connect().begin() as conn:
        sql = '''select Id, b.用户名, s.日期, s.sum_ 
                from basicInfo b
                left join 
                    (select 日期, 用户名, 类别, sum(金额) as sum_
                        from 消费 
                        where 日期 between '%s' and '%s'
                          and 类别='%s'
                        group by 日期, 用户名, 类别
                    ) s
                  on b.用户名=s.用户名
                ''' % (datSt.strftime('%Y%m%d'), datEnd.strftime('%Y%m%d')
                , class_)
        data = list(map(lambda x: list(x), conn.execute(sql).fetchall()))
        cols = ['Id', '用户名', '日期', '金额']
        df = pd.DataFrame(data, columns=cols)
        df.fillna(0, inplace=True)
        df = df.pivot_table(values=['金额'], columns=['日期']
                            , index=['Id', '用户名'])
        df.columns = df.columns.get_level_values(1)
        df.reset_index(inplace=True)
        df.sort_values(by='Id', inplace=True)
        df.fillna(0, inplace=True)
    # 写入excel
    lis = sht['A2:FA2'].value
    for dat in pd.date_range(datSt, datEnd):
        n = lis.index(dat)
        datStr = dat.strftime('%Y%m%d')
        if datStr in df.columns:
            sht[2, n].options(transpose=True).value = df[datStr].values

def fillFormula(sht, days, rows, rows2):
    '''
    填充 xlsx 中对应的空白区域，包括相应的公式、及补0操作

    Parameters
    ----------
    sht : 指定的工作表
    days : 指定更新的日期
    rows : 更新前的数据行数
    rows2 : 更新后的数据行数

    Returns
    -------
    None.

    '''
    if rows2 > rows:
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
        sht[rows:rows2, 56:cols].value = 0
        # Borders
        rng5 = sht['A' + str(rows) + ':PB' + str(rows2)]
        for i in range(7, 13):
            rng5.api.Borders(i).LineStyle = 1
            rng5.current_region.api.Borders(i).weight = 2
    

def forecast(wb, sht, rows, daysWo, daysWe):
    '''
    构造预估均值计算公式

    Parameters
    ----------
    wb : xlsx文件
    sht : 指定工作表
    daysWo : 工作日预估均值天数：默认 3日,int
    daysWe : 节假日预估均值天数：默认 1日,int

    Returns
    -------
    None

    '''
    def getAvg(lis, days):
        '''
        构造 xlsx 预估均值计算公式
        
        Parameters
        ----------
        lis: 工作日、周末日期列表
        days: 预估均值使用天数，如工作日3天，周末1天
        
        Returns
        -------
        预估均值公式，如 '=average(a3,b3,c3)'
        '''
        dat = [str_(sht, datLis.index(d), 1) for d in lis[:days]]
        avg = '=average('
        for s in dat:
            avg += s
            if s == dat[-1]:
                avg += ')'
            else:
                avg += ','
        return avg
    
    # 拆分工作日、周末 （除节假日占用工作日）
    lis_wo, lis_we = work_week_days(wb)
    #
    # 均值
    #
    datLis = sht['A2:FA2'].value  # 表头
    # 工作日
    sht['MK3'].formula = getAvg(lis_wo, daysWo)
    # 节假日
    sht['ML3'].formula = getAvg(lis_we, daysWe)
    # 填充
    rows = sht['A1'].current_region.rows.count  # 原表后行数
    rng = sht['MK3' + ':ML' + str(rows)]
    sht['MK3:ML3'].api.AutoFill(rng.api, constants.AutoFillType.xlFillCopy)

def str_(sht, col, n):
    '''
    提取单元格名称

    Parameters
    ----------
    sht : 指定工作表
    col : 指定列号，从0开始，int
    n : 计算行号，从 xlsx 表头第2行开始

    Returns
    -------
    单元格名称，如'A1'

    '''
    from functools import reduce
    # 合并字符串
    toStr = lambda x, y: x + str(int(y) + n)
    # 提取单元格名称
    getCell = str(sht[1, col])[str(sht[1, col]).index('!') + 2 : -1].split('$')
    return reduce(toStr, getCell)
    
def work_week_days(wb):
    '''
    拆分工作日、节假日

    Parameters
    ----------
    wb : xlsx

    Returns
    -------
    lis_wo : 工作日序列，除节假日占用
    lis_we : 周末序列

    '''
    # 节假日
    shtDat = wb.sheets['Date List']
    Holiday = shtDat['A2:A27'].value
    # 工作日、节假日拆分
    lis_we, lis_wo = [], []
    for i in range(1, 15):
        if dat(i).weekday() in (5, 6):
            lis_we.append(dat(i))
        elif dat(i) in Holiday:
            #lis_we.append(dat(i))  # 节假日 占有工作日
            pass
        else:
            lis_wo.append(dat(i))
    return lis_wo, lis_we

def toZero(wb, sht, rows2):
    '''
    将近 n 日工作日消费为0的预估均值变0
    占用 95% 的时间 

    Parameters
    ----------
    wb : xlsx
    sht : 指定工作表
    rows2 : 更新后的行数

    Returns
    -------
    None

    '''
    # 返回指定日期所在xlsx中位置
    def col(header, n):
        return header.index(lis_wo[n])
    
    # 获取 近5个工作日
    lis_wo, _ = work_week_days(wb)
    # 获取 表头
    header = sht['A2:FA2'].value
    #
    # 近2个工作日为0，则预估均值为0
    #
    rngVal = (str_(sht, col(header, 0), 1) + ':' 
              + str_(sht, col(header, 1), rows2-2))
    df = pd.DataFrame(sht[rngVal].value)
    df['sum_'] = df[0] + df[1]
    for r in  df[df['sum_'] == 0].index:
        sht['MK' + str(r+3)].value = 0
        sht['ML' + str(r+3)].value = 0

def main():
    '''
    主函数

    Returns
    -------
    None.

    '''
    print("默认：至昨天")
    days = int(input("从多少天以前开始至昨天？(比如：1)"))
    workDays = eval(input('均值多少个工作日？'))
    weekDays = eval(input('均值多少个周末日？'))
    if (not isinstance(days, int) and not isinstance(workDays, int) 
        and not isinstance(weekDays, int)):
        raise TypeError('请输入整数')
    wb = xw.Book(getPath())
    tableList = ['搜索点击', '新产品', '自主投放']
    with connect().begin() as conn:
        conn.execute('select 1')
        for n, shtName in enumerate(['搜索', '其他新产品', '原生广告']):
            sht = wb.sheets[shtName]
            rows = sht['A1'].current_region.rows.count  # 更新前行数
            # 更新基本信息 
            upBasicInfo(conn, sht)
            rows2 = sht['A1'].current_region.rows.count  # 原表后行数
            # 更新消费
            upSpending(conn, sht, dat(days), dat(1), tableList[n])
            # 填充公式列
            fillFormula(sht, days, rows, rows2)
            # 预估均值公式
            forecast(wb, sht, rows2, workDays, weekDays)
            # 近n日消费为0，预估均值 -> 0
            toZero(wb, sht, rows2)
    wb.app.calculate()
    wb.save()

if __name__ == '__main__':
    st = now()
    main()
    print('The program ends in %s min' % ((now() - st)/60))
    