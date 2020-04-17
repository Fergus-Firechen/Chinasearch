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
import pandas as pd
import xlwings as xw
import configparser
from datetime import date
from sqlalchemy import create_engine







path1 = [r'H:\SZ_数据\Input\P4P 消费报告2019.10...xlsx', 
        r'D:\陈怀玉\工作\周工作\Payment System\Channel.csv',
        r'D:\陈怀玉\工作\周工作\Payment System\系統-_201910_Master-5.xlsm']

path = r'D:\陈怀玉\工作\周工作\Payment System'


def get_month_week():
    '获取当月的第几周'
    firstDay = date(date.today().year, date.today().month, 1)
    w1 = firstDay.isocalendar()[1]
    w2 = date.today().isocalendar()[1]
    return str(w2 - w1 + 1)


name = ('系统-_' + date.today().strftime('%Y%m') + '_Master-' + 
        get_month_week() + '.xlsm')

i = '19年已消费'



def connect():
    def login():
        CONF = r'c:\users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        if os.path.exists(CONF):
            conf.read(CONF)
            host = conf.get('SQL Server', 'ip')
            port = conf.get('SQL Server', 'port')
            dbname = conf.get('SQL Server', 'dbname')
            return host, port, dbname
        else:
            return 'tips:检查配置文件'
    try:
        engine = create_engine('mssql:pymssql://{}:{}/{}'.format(login()))
    except Exception as e:
        print('connection failed: {}'.format(e))
    else:
        print('connection success')
        return engine
 
def col(args):
    '获取表列名'
    sql = ''' select * from information_schema.columns where table_name='{}'
        '''.format(args)
    col = [i[3] for i in engine.execute(sql).fetchall()]
    return col

def data(args):
    '获取表数据'
    sql = ''' select * from '{}'
        '''.format()
    data = engine.execute(sql).fetchall()
    return data

def spending():
    '有消费的账户列表'
    sql = ''' select distinct 用户名 from 消费
        '''
    lis = engine.execute(sql).fetchall()
    return lis

def master():
    '获取Master数据'
    if os.path.exists(os.path.join(path, name)):
        wb = xw.Book(os.path.join(path, name))
        sht = wb.sheets['Master']
        rowCnt = sht['A1'].current_region.rows.count
        
        lis = sht['B2:B'+str(rowCnt)].value
        
        
        # xlrd测试
        import xlrd
        wb1 = xlrd.open_workbook(os.path.join(path, name))
        print(wb1.sheet_names())
        sht = wb1.sheets['Master']
        print(sht.cell_value(1, 1))
        
        from xlrd import open_workbook
        if os.path.exists(os.path.join(path, name)):
            with open_workbook(os.path.join(path, name)) as Book:
                print(Book.sheet_names())
                sht = Book.sheets()[1]
                print(sht.nrows, sht.ncols)
                print(sht.row_values(0))
                print(sht.col_values(0))
                print(sht.cell(0, 0).value)
                print(sht.row(0)[0].value)
                print(sht.col(0)[0].vlaue)
                
        else:
            print('NotFoundFile:{}'.format(os.path.join(path, name)))
        
        
    else:
        print('{} 文件不存在'.format(os.path.join(path, name)))
        
    


def cost_time(func):
    '耗时跟踪'
    @functools.wraps(func)
    def wrapper(*args):
        print('%s() start:' %func.__name__)
        start_time = time.time()
        func(*args)
        stop_time = time.time()
        print('\a cost time: %.2f min' % ((stop_time-start_time)/60))
    return wrapper

@cost_time
def run():
    '测试'
    pass


@cost_time
def read_file():
    global p4p, channel, master
    p4p = pd.read_excel(path[0], sheet_name='P4P消费')
    channel = pd.read_csv(path[1], engine='python')
    master = pd.read_excel(path[2], sheet_name='Master')

@cost_time
def main():
    
    '数据筛选；结构转换；写入；'
    global p4p, channel, master
    
    # 数据清洗
    p4p = p4p.iloc[8:, :]
    p4p.fillna('-', inplace=True)
    # 信誉成长值 = 二级行业
    master.rename(columns={'財務加款的端口':'端口', '客戶用户名':'用户名',
                           '广告主URL':'URL', '行业分类':'信誉成长值', 
                           '查賬財務郵箱地址':'财务做账区域'}, inplace=True)
    column = list(master.columns)
    master = master[['用户名', '付款方式']]
    master.fillna('-', inplace=True)
    
    # 1.2019年已消费 > 0
    p4p = p4p[p4p[i] > 0]
    
    # 2.P4P & 端口 del端口加款账户
    for j in channel['端口'].tolist():
        p4p.drop(axis=0, index=p4p[p4p['端口'] == j].index, inplace=True)
    del channel
    
    # 3.merge_1 & master del已有账户
    merge_1 = pd.merge(p4p, master, how='left', on='用户名')
    merge_1 = merge_1[merge_1['付款方式'].isna()]
    del master
    
    # 4.结构转换
    merge_1 = merge_1.reindex(columns=column)
    merge_1.fillna('-', inplace=True)
    
    # 5.写入
    wb = xw.Book(path[2])
    wb.app.calculation = 'manual'
    wb.app.visible = True
    wb.app.display_alerts = False
    wb.app.screen_updating = False
    sht = wb.sheets['Master']
    sht[sht[0, 0].current_region.rows.count, 0].color = (255, 255, 204)  #ffffcc
    sht[sht[0, 0].current_region.rows.count, 0].value = merge_1.values
    wb.save()
    wb.app.screen_updating = True
    #wb.close()
    
if __name__ == '__main__':
    
    engine = connect()
    run()
    read_file()
    main()
    pass
    
    