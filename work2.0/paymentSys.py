# -*- coding: utf-8 -*-
"""
Created on Fri Jan 11 11:30:32 2019

1.读取P4P,加款端口，Master系统账户
2.筛选：2019有消费、除端口加款、除已有账户
3.结构转换、写入excel

@author: chen.huaiyu
"""
import os
import time
import functools
import configparser
import pandas as pd
import xlwings as xw
from datetime import date, datetime
from sqlalchemy import create_engine

now = lambda : time.perf_counter()

def costTime(func):
    @functools.wraps(func)
    def wrapper():
        print('{}() start: {}'.format(func.__name__, datetime.now()))
        st = now()
        func()
        print('Runtime {}s'.format(round(now() - st, 3)))
    return wrapper
    
def connect():
    '连接数据库'
    def login():
        CONF = r'c:\users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        item = 'SQL Server'
        if os.path.exists(CONF):
            conf.read(CONF)
            host = conf.get(item, 'ip')
            port = conf.get(item, 'port')
            dbname = conf.get(item, 'dbname')
            return host, port, dbname
        else:
            return 'NotFoundFile {}'.format(CONF)
    try:
        engine = create_engine('mssql+pymssql://{}:{}/{}'.format(login()[0], 
                               login()[1], login()[2]))
    except Exception as e:
        print('connection failed: {}'.format(e))
    else:
        print('connection success')
        return engine

def handling(conn):
    '获取数据、处理数据得到最终结果'
    def col():
        lis = ['端口','用户名','SG','财务做账区域','付款方式','销售','AM',
               '操作','销售邮箱','AM邮箱','OP邮箱','客户','网站名称','广告主',
               '客户地址','URL','信誉成长值','币种','联络人','联络人邮箱',
               '联络人电话','开户日期','收取年服务费时间','年费',
               '续费返点(p4p)','管理费(p4p','续费返点(inf)','管理费(inf)',
               '账期/预付']
        return lis

    def data():
        sql = ''' SELECT b.端口
                        , b.用户名
                        , '-' AS 'SG'
                        , 财务做账区域
                        , '-' AS 付款方式
                        , b.销售
                        , AM
                        , 操作
                        , '-' AS '销售邮箱'
                        , '-' AS 'AM邮箱'
                        , '-' AS 'OP邮箱'
                        , b.客户
                        , b.网站名称
                        , 广告主
                        , '-' AS '客户地址'
                        , URL
                        , 信誉成长值
                        , '-' AS '币种'
                        , '-' AS '联络人'
                        , '-' AS '联络人邮箱'
                        , '-' AS '联络人电话'
                        , 开户日期
                        , 收取年服务费时间
                        , kh.年费
                        , '-' AS '续费返点(p4p)'
                        , '-' AS '管理费(p4p'
                        , '-' AS '续费返点(inf)'
                        , '-' AS '管理费(inf)'
                        , kh.[账期/预付]
                    FROM dbo.basicInfo b
					  LEFT JOIN 开户申请表 kh
					  ON b.用户名 = kh.用户名
                    WHERE 
                      NOT EXISTS
                        (SELECT *
                          FROM dbo.payment u
                          WHERE b.用户名 = u.用户名)
                      AND NOT EXISTS
                        (SELECT *
                          FROM dbo.channel p
                          WHERE p.加款端口 = 'Y'
                            AND p.端口 = b.端口)
                      AND EXISTS
                        (SELECT *
                          FROM dbo.消费 sp
                          WHERE sp.日期 >= '20190101'
                            AND sp.类别 = '总点击'
                            AND sp.金额 > 0
                            AND sp.用户名 = b.用户名)
                '''
        data = conn.execute(sql).fetchall()
        return data
    
    return col, data

def toEx(col, data, conn):
    '将最终结果写入excel'
    @costTime
    def updatePayment():
        # '新增的户插入payment'
        sql = ''' INSERT INTO dbo.payment (用户名)
                    VALUES ('{}')
            '''
        for i in data():
            conn.execute(sql.format(i[1]))
    
    def getMonthWeek():
        '获取当月的第几周'
        firstDay = date(today.year, today.month, 1)
        w1 = firstDay.isocalendar()[1]
        w2 = today.isocalendar()[1]
        return str(w2 - w1 + 1)
    
    def getPath():
        PATH = r'D:\陈怀玉\工作\周工作\Payment System'
        name = '系统-_' + today.strftime('%Y%m') + '_Master-' + getMonthWeek() + '.xlsm'
        if os.path.exists(os.path.join(PATH, name)):
            return os.path.join(PATH, name)
        else:
            print('NotFoundFil: {}'.format(name))
    
    today = date.today()
    try:
        wb = xw.Book(getPath())
        wb.app.calculation = 'manual'
        print(wb.sheets)
        for shtName in ['Master']:
            sht = wb.sheets[shtName]
            nums = sht['A1'].current_region.rows.count
            rng = sht['A1'].offset(nums+1, 0)
            rng.value = pd.DataFrame(data(), columns=col()).values
        wb.save()
    except Exception as e:
        print('toEx Error: {}'.format(e))
    else:
        updatePayment()

@costTime
def main():
    '主程序'
    engine = connect()
    with engine.begin() as conn:
        col, data = handling(conn)
        toEx(col, data, conn)

if __name__ == '__main__':
    main()

