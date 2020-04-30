# -*- conding: utf-8 _*_
'''

图纸  (5w2h、模型、指标、项目管理)
素材  -- 业务
工具  -- 工具
流程、方法 -- 分析


- 分析@知识@技能
- auto_work

'''

import os
import time
import functools
import pandas as pd
from auto_ave import getPath
from datetime import datetime, timedelta, date


now = lambda : time.perf_counter()
dat = lambda n: date.today() - timedelta(n)
dat_ = lambda n: date.today() + timedelta(n)

D = dat(3)

def log(func):
    @functools.wraps(func)
    def wrapper(*args, **kw):
        print("Call %s():" % func.__name__)
        return func(*args, **kw)
    return wrapper

@log
def readAve(shtName):
    path = getPath()
    df = pd.read_excel(path, sheet_name=shtName, skiprows=1)
    dataCleaning(df)
    return df.loc[:, :'Notes'], df.loc[:, 'Notes':]

def getQ(dat):
    # 返回给定日期所在季度
    # 目的：剔除非本季度的消费
    if isinstance(dat, date):
        m = dat.month
        if m in (1, 2, 3):
            return 'Q1'
        elif m in (4, 5, 6):
            return 'Q2'
        elif m in (7, 8, 9):
            return 'Q3'
        else:
            return 'Q4'
    else:
        raise

def dataCleaning(df):
    
    def dropBracket(lis):
        # 去除字段中的 圆括号 ()
        return tuple(map(
            lambda s: s.replace('(', ' ').replace(')', '')
            , lis))
    
    def dropDat(df, col):
        # 删除非本季度消费字段
        return df.drop(columns=col, inplace=True)
    
    def transferDat():
        # '字段日期格式转换： datetime/str -> %Y%m%d'
        # 剔除非本季度的消费
        newCols = []
        for col in df.columns:
            n = 1
            try:
                if getQ(col) == getQ(D):
                    col_ = col.strftime('%Y%m%d')
                    newCols.append(col_)
                else:
                    # 删除非本季度消费
                    dropDat(df, col)
            except:
                try:
                    col_ = datetime.strptime(col, '%Y-%m-%d %H:%M:%S.%f'
                                             ).strftime('%Y%m%d')
                    if col_ in newCols:
                        col_1 = col_ + '_' + str(n)
                        n += 1
                        if col_1 in newCols:
                            col_2 = col_ + '_' + str(n)
                            newCols.append(col_2)
                        else:
                            newCols.append(col_1)
                    else:
                        newCols.append(col_)
                except ValueError:
                    newCols.append(col)
        return dropBracket(newCols)
    
    def consistentType(df):
        # 统一数据类型为 datetime.datetime
        # datetime.time(0,0) -> None
        df['首次消费日_1'] = df['首次消费日']
        df['Campaign Start Date'] = df['首次消费日']
        df['Campaign Start Date_1'] = df['首次消费日']
        
    def dropZeroSpending(df):
        # 当前年度 2020
        # 前一年度 2019
        index = df[df['2019YTD'] + df['2020YTD'] == 0].index
        df.drop(index=index, inplace=True)
    
    consistentType(df)
    #dropZeroSpending(df)  # 降低时间有限；且影响 output结果带来不确定
    # 增加标识列,拆为两部分上载，便于再次合并
    df['用户名1'] = df['用户名']
    # 日期转换
    df.columns = transferDat()

@log
def upLoad(getTable, shtName):
    def dropNotes(df):
        df.drop(columns='Notes', inplace=True)  # 多一列 Notes
        
    def getDate(ver):
        return D.strftime('%Y%m%d') + ver
        
    def getName():
        if '搜' in shtName:
            return 'P4P_' + getDate('_1'), 'P4P_' + getDate('_2')
        elif '新' in shtName:
            return 'NP_' + getDate('_1'), 'NP_' + getDate('_2')
        elif '原' in shtName:
            return 'Infeeds_' + getDate('_1'), 'Infeeds_' + getDate('_2')
    
    t, t1 = getTable(shtName)
    table1, table2 = getName()
    dropNotes(t)
    # 写入
    with connect().begin() as en:
        t.to_sql(table1, con=en, if_exists='replace', index=False
                 , chunksize=1000)
        t1.to_sql(table2, con=en, if_exists='replace', index=False
                  , chunksize=1000)

def connect():
    def loginInfo(section):
        from configparser import ConfigParser
        path = r'H:\SZ_数据\Python\c.s.conf'
        #path = r'C:\users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = ConfigParser()
        conf.read(path)
        return (conf.get(section, 'acc'), conf.get(section, 'pw')
                , conf.get(section, 'ip'), conf.get(section, 'port')
                , conf.get(section, 'dbname'))
    
    from sqlalchemy import create_engine
    try:
        engine = create_engine(
            'mssql+pymssql://%s:%s@%s:%s/%s' % loginInfo('Output'))
    except Exception as e:
        print('连接失败： %s' % e)
        raise
    else:
        return engine

@log
def output(version):
    
    def outputFilList(version):
        '''
        构建 output 文件列表
        以模版为基础新建文件夹、复制文件、修改名称
    
        Returns
        -------
        None.
    
        '''
        def sFilName():
            for d in range(2, 60):
                filName = dat(d).strftime('%Y%m%d v1')
                if os.path.isdir(os.path.join(path, filName)):
                    return filName, dat(d).year, getQ(dat(d))
        
        def source():
            return os.path.join(path, sFilName()[0])
        
        def tFilName():
            return D.strftime('%Y%m%d ' + version)
        
        def target():
            return os.path.join(path, tFilName())
        
        def cpDir():
            import shutil
            lis = []
            sName, sYear, sQuar = sFilName()
            sourceDir = source()
            if not os.path.exists(target()):
                shutil.copytree(sourceDir, target())    
            for f in os.listdir(sourceDir):
                s = os.path.join(target(), f)
                # modify new name
                f = f.replace(sName, tFilName())
                f = f.replace(str(sYear), str(D.year))   # year
                f = f.replace(sQuar, getQ(D))   # quarter
                t = os.path.join(target(), f)
                lis.append(t)
                os.rename(s, t)
            return lis
        
        path = r'H:\SZ_数据\Output'
        return cpDir()
    
    def data(conn, sql):
        return list(map(lambda x: list(x), conn.execute(sql).fetchall()))
    
    def shtNameLis(table):
        strDat = D.strftime('%Y%m%d')
        return (table + strDat, table+strDat+'_1', table+strDat+'_2')
    
    def getRows(sht, row, col):
        # 获取行数
        return sht[row, col].current_region.rows.count
    
    def getCols(sht, row, col):
        # 获取列数
        return sht[row, col].current_region.columns.count
    
    def getRng(sht, row, col):
        # 数据首行区域 & 全部区域
        cntRow = getRows(sht, row, col)
        cntCol = getCols(sht, row, col)
        return sht[row, col:col+cntCol], sht[row:row+cntRow, col:col+cntCol]
    
    def getRng2(sht, newsht, row, col):
        # 外部表与当前表对应区域
        cntRow = getRows(sht, row, col)
        cntCol = getCols(sht, row, col)
        if getRows(newsht, row, col) < cntRow:
            return newsht[row, col:col+cntCol], newsht[row:row+cntRow-1
                                                       , col:col+cntCol]
        
    def getRng3(sht, row, col1, col2):
        # 获取指定区域
        cntRow = getRows(sht, row, col1)
        return sht[row, col1:col2], sht[row:row+cntRow-3, col1:col2]
    
    def clear2(sht, row, col1, col2):
        # 清空 指定区域
        cntRow = getRows(sht, row, col1)
        rng = sht[row:row+cntRow, col1:col2]
        rng.clear()
    
    def clear(sht, row, col):
        # 清空 所有区域
        _, rng = getRng(sht, row, col)
        rng.clear()
        
    def write(sht, row, col, conn, sql):
        # 指定位置写入数据
        sht[row, col].value = data(conn, sql)
        
    def sqlSpendForecast(lis):
        return "EXEC dbo.pr_spendForecast '%s', '%s', '%s'" % lis
    
    def sqlSpendForecastV(lis):
        return "EXEC dbo.pr_spendForecastV1 '%s', '%s' " % lis
    
    def sqlCashForecast(lis):
        return ''' EXEC dbo.pr_cashSpendForecast '%s', '%s', '%s', '%s', '%s'
                , '%s', '%s' ''' % lis
    
    def sqlHandlingFee(lis):
        return ''' EXEC dbo.pr_handlingFee '%s', '%s', '%s', '%s', '%s', '%s'
                , '%s' ''' % lis
    
    def sqlHandlingFeeDetails(t):
        return ''' EXEC dbo.pr_handlingFeeDetails %s ''' % t
    
    def sqlSalesForecast(lis):
        return ''' EXEC dbo.pr_salesForecast '%s', '%s' ''' % lis
    
    def sqlGPRatio(lis):
        return ''' EXEC dbo.pr_GPRatio '%s', '%s' ''' % lis
    
    def sqlSalesTracking(lis):
        return ''' EXEC dbo.pr_salesTracking '%s', '%s', '%s', '%s'
                , '%s' ''' % lis
    
    def sqlAMTracking(lis):
        return ''' EXEC dbo.pr_AMTracking '%s', '%s', '%s', '%s', '%s'
                , '%s' ''' % lis
    
    def sqlCreateFAF(lis):
        return ''' EXEC dbo.pr_createFAF '%s', '%s', '%s', '%s', '%s', '%s'
                , '%s' ''' % lis
    
    def sqlFAF(lis):
        return ''' EXEC dbo.pr_FAF '%s', '%s', '%s', '%s', '%s', '%s' ''' % lis
    
    def getMonth():
        q = getQ(D)
        if q == 'Q1':
            return 'Jan', 'Feb', 'Mar'
        elif q == 'Q2':
            return 'Apr', 'May', 'Jun'
        elif q == 'Q3':
            return 'Jul', 'Aug', 'Sep'
        else:
            return 'Oct', 'Nov', 'Dec'
    
    def index(sht, newsht, row, col):
        # 更新首列：用户名 & 端口
        cntRow = getRows(sht, row, col)
        newsht[row, col].options(transpose=True).value = sht[1:cntRow
                                                             , col].value
    
    def fillFormula(sht, newsht, row, col):
        # 填充公式
        if getRng2(sht, newsht, row, col):
            rng0, rng1 = getRng2(sht, newsht, row, col)
            rng0.api.AutoFill(rng1.api, constants.AutoFillType.xlFillCopy)
    
    def fillFormula2(sht, row, col1, col2):
        # 填充指定区域
        rng0, rng1 = getRng3(sht, row, col1, col2)
        rng0.api.AutoFill(rng1.api, constants.AutoFillType.xlFillCopy)
    
    import xlwings as xw
    from xlwings import constants
    #
    with connect().begin() as conn:
        # 合并
        conn.execute("EXEC newAve '%s', '%s', '%s'" % shtNameLis('P4P_'))
        conn.execute("EXEC newAve '%s', '%s', '%s'" % shtNameLis('NP_'))
        conn.execute("EXEC newAve '%s', '%s', '%s'" % shtNameLis('Infeeds_'))
        # 写入excel
        filLis = outputFilList('v1')
        p4p, np, inf = [shtNameLis(i)[0] for i in ['P4P_', 'NP_', 'Infeeds_']]
        year = D.strftime('%Y')
        q = getQ(D)
        m1, m2, m3 = getMonth()
        for path in filLis:
            # 打开
            wb = xw.Book(path)
            if 'Spending Forecast' in path and '_v1' not in path:
                # 清除表中数据
                # 更新数据
                #
                sht1 = wb.sheets['P4P']
                clear(sht1, 1, 0)
                write(sht1, 1, 0, conn, sqlSpendForecast((p4p, '端口', year)))
                clear(sht1, 1, 14)
                write(sht1, 1, 14, conn, sqlSpendForecast((p4p, '用户名'
                                                           , year)))
                #
                sht = wb.sheets['NP']
                clear(sht, 1, 0)
                write(sht, 1, 0, conn, sqlSpendForecast((np, '端口', year)))
                clear(sht, 1, 14)
                write(sht, 1, 14, conn, sqlSpendForecast((np, '用户名', year)))
                #
                sht = wb.sheets['Infeeds']
                clear(sht, 1, 0)
                write(sht, 1, 0, conn, sqlSpendForecast((inf, '端口', year)))
                clear(sht, 1, 14)
                write(sht, 1, 14, conn, sqlSpendForecast((inf, '用户名'
                                                          , year)))
                #
                sht = wb.sheets['Region']
                write(sht, 2, 0, conn, sqlSpendForecast((p4p, '区域', year)))
                write(sht, 11, 0, conn, sqlSpendForecast((np, '区域', year)))
                write(sht, 20, 0, conn, sqlSpendForecast((inf, '区域', year)))
                #
                sht = wb.sheets['All']
                fillFormula(sht1, sht, 1, 1)
                fillFormula(sht1, sht, 1, 15)
                index(sht1, sht, 1, 0)
                index(sht1, sht, 1, 14)
                #
                sht = wb.sheets['Infeeds(35%)']
                fillFormula(sht1, sht, 1, 1)
                fillFormula(sht1, sht, 1, 15)
                index(sht1, sht, 1, 0)
                index(sht1, sht, 1, 14)
                # 保存、关闭
                wb.save()
                wb.close()
            elif 'Spending Forecast' in path and '_v1' in path:
                #
                sht1 = wb.sheets['P4P']
                write(sht1, 1, 0, conn, sqlSpendForecastV((p4p, year)))
                #
                sht = wb.sheets['NP']
                write(sht, 1, 0, conn, sqlSpendForecastV((np, year)))
                #
                sht = wb.sheets['Infeeds']
                write(sht, 1, 0, conn, sqlSpendForecastV((inf, year)))
                #
                sht = wb.sheets['All']
                clear2(sht, 3, 0, 18)
                fillFormula(sht1, sht, 1, 0)
                #
                wb.save()
                wb.close()
            elif 'Cash' in path:
                #
                sht = wb.sheets['Region']
                write(sht, 2, 0, conn, sqlCashForecast((p4p, '区域'
                                                        , year, q
                                                        , m1, m2, m3)))
                write(sht, 12, 0, conn, sqlCashForecast((np, '区域'
                                                         , year, q
                                                         , m1, m2, m3)))
                write(sht, 22, 0, conn, sqlCashForecast((inf, '区域'
                                                         , year, q
                                                         , m1, m2, m3)))
                #
                sht = wb.sheets['Finance Region']
                write(sht, 2, 0, conn, sqlCashForecast((p4p, '财务'
                                                        , year, q
                                                        , m1, m2, m3)))
                write(sht, 18, 0, conn, sqlCashForecast((np, '财务'
                                                         , year, q
                                                         , m1, m2, m3)))
                write(sht, 34, 0, conn, sqlCashForecast((inf, '财务'
                                                         , year, q
                                                         , m1, m2, m3)))
                # 
                wb.save()
                wb.close()
            elif 'Handling' in path and 'Details' not in path:
                #
                sht = wb.sheets['AM']
                write(sht, 2, 0, conn, sqlHandlingFee((p4p, 'AM'
                                                       , year, q
                                                       , m1, m2, m3)))
                write(sht, 19, 0, conn, sqlHandlingFee((np, 'AM'
                                                        , year, q
                                                        , m1, m2, m3)))
                write(sht, 36, 0, conn, sqlHandlingFee((inf, 'AM'
                                                        , year, q
                                                        , m1, m2, m3)))
                #
                sht = wb.sheets['Sales']
                write(sht, 2, 0, conn, sqlHandlingFee((p4p, 'Sales'
                                                       , year, q
                                                       , m1, m2, m3)))
                write(sht, 25, 0, conn, sqlHandlingFee((np, 'Sales'
                                                        , year, q
                                                        , m1, m2, m3)))
                write(sht, 49, 0, conn, sqlHandlingFee((inf, 'Sales'
                                                        , year, q
                                                        , m1, m2, m3)))
                #
                wb.save()
                wb.close()
            elif 'Handling' in path and 'Details' in path:
                #
                sht = wb.sheets['P4P']
                clear(sht, 1, 0)
                write(sht, 1, 0, conn, sqlHandlingFeeDetails(p4p))
                #
                sht = wb.sheets['NP']
                clear(sht, 1, 0)
                write(sht, 1, 0, conn, sqlHandlingFeeDetails(np))
                #
                sht = wb.sheets['Infeeds']
                clear(sht, 1, 0)
                write(sht, 1, 0, conn, sqlHandlingFeeDetails(inf))
                #
                wb.save()
                wb.close()
            elif 'Sales Forecast' in path:
                #
                sht = wb.sheets['Data']
                clear2(sht, 3, 0, 12)
                clear2(sht, 4, 12, 18)
                write(sht, 3, 0, conn, sqlSalesForecast((p4p, q)))
                write(sht, 3, 4, conn, sqlSalesForecast((np, q)))
                write(sht, 3, 7, conn, sqlSalesForecast((inf, q)))
                fillFormula2(sht, 3, 12, 18)
                #
                wb.save()
                wb.close()
            elif 'GP' in path:
                #
                sht = wb.sheets[0]
                write(sht, 2, 0, conn, sqlGPRatio((p4p, q)))
                write(sht, 13, 0, conn, sqlGPRatio((np, q)))
                write(sht, 24, 0, conn, sqlGPRatio((inf, q)))
                #
                wb.save()
                wb.close()
            elif 'Sales Tracking' in path:
                #
                sht = wb.sheets[0]
                write(sht, 2, 0, conn, sqlSalesTracking((p4p, q, m1, m2, m3)))
                write(sht, 27, 0, conn, sqlSalesTracking((np, q, m1, m2, m3)))
                write(sht, 52, 0, conn, sqlSalesTracking((inf, q
                                                          , m1, m2, m3)))
                # 
                wb.save()
                wb.close()
            elif 'AM Tracking' in path:
                #
                sht = wb.sheets[0]
                #
                write(sht, 2, 0, conn
                      , sqlAMTracking((p4p, year, q, m1, m2, m3)))
                write(sht, 17, 0
                      , conn, sqlAMTracking((np, year, q, m1, m2, m3)))
                write(sht, 32, 0
                      , conn, sqlAMTracking((inf, year, q, m1, m2, m3)))
                #
                wb.save()
                wb.close()
            elif 'FAF' in path:
                # createFAF
                conn.execute(sqlCreateFAF((p4p, np, inf, year, m1, m2, m3)))
                # select
                # GP
                sht = wb.sheets['Weekly GP Report']
                write(sht, 2, 0, conn, sqlFAF(('region', year
                                               , q, m1, m2, m3)))
                # KPI
                sht = wb.sheets['KPI']
                write(sht, 2, 0, conn, sqlFAF(('am', year, q, m1, m2, m3)))
                #
                wb.save()
                wb.close()  

@log
def main():
    '''
    主函数

    Returns
    -------
    None.

    '''
    for shtName in ('搜索', '其他新产品', '原生广告'):
        #
        try:
            upLoad(readAve, shtName)
            # 版本
            output(input('版本：v1?'))
        except UnicodeDecodeError as e:
            print("局域网连接异常，刷新文件夹: %s" % e)
            print('Runtime %.3f min' % ((now() - st)/60))
            raise
        except ValueError as e:
            print('Runtime %.3f min' % ((now() - st)/60))
            if 'Invalid file path' in e:
                print('局域网连接异常，打开部门内共享文件夹，刷新一下')
            raise
        except Exception as e:
            print(e)
            print('Runtime %.3f min' % ((now() - st)/60))
            raise
    

if __name__ == '__main__':
    
    st = now()
    
    output('v1')
    #main()
    
    print('Runtime %.3f min' % ((now() - st)/60))