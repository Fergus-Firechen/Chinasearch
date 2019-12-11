# -*- coding: utf-8 -*-
"""
Created on Wed Jul 24 13:49:02 2019

@author: chen.huaiyu
"""


import pandas as pd
from sqlalchemy import create_engine
import configparser
import os


def connectDB():
    def login():
        CONF = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        if os.path.exists(CONF):
            conf.read(CONF)
            accountName = conf.get('SQL Server', 'accountname')
            pw = conf.get('SQL Server', 'password')
            host = conf.get('SQL Server', 'ip')
            port = conf.get('SQL Server', 'port')
            dbname = conf.get('SQL Server', 'dbname')
            return accountName, pw, host, port, dbname
        else:
            print('SQL Server配置文件不存在')
            
    s = "mssql+pymssql://%s:%s@%s:%s/%s"
    engine = create_engine(s % login())
    if engine.execute('select 1'):
        print("SQL Server连接成功")
    else:
        raise
    return engine

def get_path_ka():
    PATH = r'C:\Users\chen.huaiyu\Downloads'
    name = r'代理商用户报表_订单明细日粒度下载_' + date_st + '-' + date_ed + '.csv'
    path = os.path.join(PATH, name)
    return path

def read_ka():
    if os.path.exists(get_path_ka()):
        path = get_path_ka()
        print(path)
        ka = pd.read_csv(path, encoding='gbk', engine='python')
        return ka
    else:
        print('NotFoundFile: {}'.format(get_path_ka))

def handling_ka():
    ka = read_ka()
    if isinstance(ka, pd.DataFrame):
        ka.to_sql('ka_basicInfo', con=engine, if_exists='replace', index=False)
        #
        ka.drop(index=ka[ka['合同号'] == 'A17KA1289'].index, inplace=True)
        ka.drop(index=ka[ka['广告主名称'] == '草莓有限公司'].index, inplace=True)
        #
        ka.rename(columns={'发生日期':'日期', '收入金额':'金额', '合同号':'用户名'}, inplace=True)
        ka['类别'] = 'KA'
        #
        ka_1 = ka[['日期', '金额', '用户名', '类别']]
        ka_1['日期'] = pd.to_datetime(ka_1['日期'])
        ka_1['日期'] = ka_1['日期'].apply(lambda x: x.strftime('%Y%m%d'))
        ka_1['金额'] = ka_1['金额'].str.replace(',', '')
        #
        ka_1['金额'] = ka_1['金额'].apply(lambda x: eval(x))
        #
        ka_1.to_sql(table_name, con=engine, if_exists='append', index=False, chunksize=1000)

def drop_db():
    sql = ''' DELETE FROM [Account Management].dbo.{}
                WHERE 类别 = 'KA'
                  AND 日期 BETWEEN '{}' AND '{}'
        '''
    with engine.begin() as conn:
        conn.execute(sql.format(table_name, date_st, date_ed))


if __name__ == '__main__':

    print('Note: ka_to_db')
    date_st = input('输入开始日期(格式如:20190101):')
    date_ed = input('输入截止日期(格式如:20190101):')
    table_name = input('消费/现金？(默认消费)')
    if table_name == '':
        table_name = '消费'
    else:
        table_name = '现金'
    if date_st and date_ed:
        engine = connectDB()
        drop_db()
        handling_ka()

