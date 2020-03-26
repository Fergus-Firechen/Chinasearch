# -*- coding: utf-8 -*-
"""
Created on Fri Feb 28 15:29:57 2020

@author: chen.huaiyu
"""

import os
import time
import configparser
import pandas as pd
from sqlalchemy import create_engine

now = lambda : time.perf_counter()

#PATH = r'c:/users/chen.huaiyu/desktop/行业变更.xlsx'
PATH = r'H:\SZ_数据\Input\P4P 消费报告2020.03...xlsx'
input("Check the address of file: {}".format(PATH))

def func_modify_industry():
    '临时：行业变更'
    
    def connectDB():
        ''' 连接Account Management
        '''
        def loginAccount():
            CONF = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
            conf = configparser.ConfigParser()
            if os.path.exists(CONF):
                conf.read(CONF)
                host = conf.get('SQL Server', 'ip')
                port = conf.get('SQL Server', 'port')
                dbname = conf.get('SQL Server', 'dbname')
                return host, port, dbname
        try:
            engine = create_engine(
                    "mssql+pymssql://@%s:%s/%s" % loginAccount())
        except:
            print('SQL Server 连接失败')
        else:
            print('SQL Server 连接成功')
            return engine
    
    def col(tableName):
        sql = ''' select * from information_schema.columns where table_name = '{}'
            '''.format(tableName)
        col = [i[3] for i in engine.execute(sql).fetchall()]
        return col
    
    def data(tableName):
        sql = "select * from {}".format(tableName)
        dat = engine.execute(sql)
        return dat
    
    def update(tableName, Ind1, Ind2, userName):
        print('update start.')
        sql = ''' update {} set Industry='{}', 信誉成长值='{}'
            where 用户名='{}'
            '''.format(tableName, Ind1, Ind2, userName)
        engine.execute(sql)
    
    def select(user):
        sql = ''' select 用户名, Industry, 信誉成长值 from basicInfo where 用户名='{}'
            '''.format(user)
        data = engine.execute(sql).fetchall()
        return data
    
    def backUp(tableName):
        sql = ''' if exists(select * from information_schema.columns where table_name='basicInfoBackUp')
                    drop table basicInfoBackUp
                    
                  select *
                      into basicInfoBackUp
                  from {}
                '''.format(tableName)
        engine.execute(sql)
        print('备份表名： basicInfoBackUp')
    
    st = now()
    engine = connectDB()
    
    if os.path.exists(PATH):
        df = pd.read_excel(PATH, sheet_name='P4P消费')
    print('Read file: {}'.format(now() - st))
    
    backUp('basicInfo')
    time.sleep(3)
    for user in df['用户名'].tolist():
        print(user)
        try:
            Ind1 = df.loc[df['用户名'] == user, '一级行业'].values[0]
            Ind2 = df.loc[df['用户名'] == user, '二级行业'].values[0]
            print(user, Ind1, Ind2)
        except Exception as e:
            print('\n-Error 1:{}: {}'.format(e, user))
        else:
            update('basicInfo', Ind1, Ind2, user)
            print(select(user))
    
    print(now() - st)

if __name__ == '__main__':
    
    func_modify_industry()