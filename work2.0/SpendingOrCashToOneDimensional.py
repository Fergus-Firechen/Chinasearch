# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 11:22:45 2019

- 跨月数据由于百通在不同表中，无法一次性处理
- 不用与basicInfo对应，只与消费有关

@author: chen.huaiyu
"""
import os
import time
import functools
import configparser
import pandas as pd
import xlwings as xw
from datetime import datetime
from sqlalchemy import create_engine



class array(object):
    '''构造序列：日期，icrm字段
    '''
    def __init__(self, startDate, endDate, target):
        if (isinstance(strToDate(startDate), datetime) and 
            isinstance(strToDate(endDate), datetime)):
            self.startDate = startDate
            self.endDate = endDate
        if target in ['消费', '现金']:
            self.target = target
        else:
            raise "只能是'消费'/'现金'"
        
    def date(self):
        # 生成：目标日期区间
        date = [i.strftime('%Y%m%d') for i in pd.date_range(
                start=self.startDate, end=self.endDate)]
        return date
    
    def field(self):
        # 字段集 产品名+日期
        col = []
        for i in range(6):
            col.append(list(map(lambda x: self.print_col()[i] + self.target + x, self.date())))
        return col
    
    def print_date(self):
        return datetime.strptime(self.startDate, '%Y-%m-%d')
    
    def print_col(self):
        return ['总点击', '搜索点击', '自主投放', '新产品', '百通', '无线搜索点击']
    
    def print_path(self):
        path0 = os.getcwd()
        path = r"c:\users\chen.huaiyu\downloads"
        # 统一文件名
        os.chdir(path)
        for i in filter(lambda x: '~' in x, os.listdir()):
            os.rename(i, i.replace('~', '_'))
        # 生成访问路径
        if self.target == '消费':
            path = os.path.join(path, "p4p %s_%s.csv" % (self.date()[0], self.date()[-1]))
        elif self.target == '现金':
            path = os.path.join(path, "cash %s_%s.csv" % (self.date()[0], self.date()[-1]))
        os.chdir(path0)
        return path
    
def strToDate(strDate):
    ''' 日期检查
    '''
    try:
        return datetime.strptime(strDate, '%Y-%m-%d')
    except Exception as e:
        print('请输入正确的日期格式(YYYY-mm-dd: {}'.format(e))
                                                           
def cost_time(func):
    '耗时'
    @functools.wraps(func)
    def wrapper(*args):
        print('%s() start:' %func.__name__)
        start_time = time.time()
        func(*args)
        end_time = time.time()
        print('\a\a cost time: %.2f min' % ((end_time-start_time)/60))
    return wrapper

@cost_time
def run():
    '测试'
    pass

def dfFromDB(tableName):
    sql = "select * from information_schema.columns where table_name='basicInfo'"
    col = [i[3] for i in engine.execute(sql)]
    data = engine.execute('select * from basicInfo')
    df = pd.DataFrame(data, columns=col)
    return df

@cost_time
def read_file(obj):
    global df1, df3  #df2, 
    '==icrm 消费/现金csv=='
    df1 = pd.read_csv(a.print_path(), engine='python', encoding='gbk')
    df1.rename(columns={'账户名称':'用户名'}, inplace=True)
    df1['用户名'] = df1['用户名'].astype(str)
    df1.drop(columns='账户ID', inplace=True)
    '==百通=='
    wb = xw.Book(r'H:\SZ_数据\Input\每日百度消费.xlsx')
    sheet = 'P4P消费'+str(a.print_date().month)+'月'
    sht = wb.sheets[sheet]
    cnt = sht['A39'].current_region.rows.count - 2
    df3 = pd.DataFrame(sht[38:38+cnt, :33].value)
    #df3 = pd.read_excel(r'H:\SZ_数据\Input\每日百度消费.xlsx', sheet_name=sheet).iloc[38:cnt+38,:]
    '结构整理'
    df3.iloc[1, 0] = '用户名'
    df3.columns = df3.iloc[1, :].tolist()
    df3.drop(index=[0, 1], inplace=True)
    df3.drop(columns='总计', inplace=True)
    df3.set_index('用户名', drop=True, inplace=True)
    df3.dropna(inplace=True, how='all', axis=1)
    df3.columns = [i.strftime('%Y%m%d') for i in df3.columns]
    df3 = df3.loc[:, a.date()[0]:a.date()[-1]]
    df3.columns = a.field()[4]
    if obj == '现金':
        df3 = df3.applymap(lambda x: x/1.22)
    else:
        pass
    
def upload_ka(start, stop):
    
    def strf(arg):
        s = strToDate(arg)
        return s.strftime('%Y%m%d')
    
    path = r"c:\users\chen.huaiyu\downloads"
    fil = '代理商用户报表_订单明细日粒度下载_%s-%s.csv' % (strf(start), strf(stop))
    ka = pd.read_csv(os.path.join(path, fil), encoding='GBK')

    # 筛选符合条件的数据
    ka.drop(index=ka[ka['合同号'] == 'A17KA1289'].index, inplace=True)
    ka.drop(index=ka[ka['广告主名称'] == '草莓有限公司'].index, inplace=True)

    # 调整
    ka.rename(columns={'发生日期':'日期', '收入金额':'金额', '合同号':'用户名'}, inplace=True)
    ka['类别'] = 'KA'
    ka_1 = ka[['日期', '金额', '用户名', '类别']]
    ka_1['日期'] = pd.to_datetime(ka_1['日期'])
    ka_1['日期'] = ka_1['日期'].apply(lambda x: x.strftime('%Y%m%d'))
    ka_1['金额'] = ka_1['金额'].str.replace(',', '')
    ka_1['金额'] = ka_1['金额'].apply(lambda x: eval(x))
    ka_1.to_sql('消费', con=engine, if_exists='append', index=False, chunksize=1000)
    

@cost_time
def main():
    global df1
    # 元字段生成
    col_all, col_p4p, col_inf, col_np, col_bt, col_mo = a.field()
    '新产品计算'
    for i in range(len(col_all)):
        df1[col_np[i]] = df1[col_all[i]] - df1[col_p4p[i]] - df1[col_inf[i]]
    '+百通'
    df1.set_index('用户名', drop=True, inplace=True)
    for i in range(len(col_all)):
        print(i)
        for j in df3.index:
            print(j)
            df1.loc[j, col_all[i]] = df1.loc[j, col_all[i]] + df3.loc[j, 
                                       col_bt[i]]
            df1.loc[j, col_np[i]] = df1.loc[j, col_np[i]] + df3.loc[j, 
                                       col_bt[i]]
    '消费文件'
    df1.reset_index(inplace=True)
    df1 = pd.merge(df1, df3, how='left', on='用户名')
    df1.fillna(0, inplace=True)
    '转换为一维表'
    column = ['日期', '用户名', '类别', '金额']
    df_1 = pd.DataFrame(columns=column)
    col_list = [col_all, col_p4p, col_inf, col_np, col_bt, col_mo]
    for m,j in enumerate(a.date()):
        df_d = df1[df1[col_all[m]] > 0]
        if df_d.shape[0] > 10:
            for i in df_d['用户名'].tolist():
                for n, k in enumerate(a.print_col()):
                    df_2 = pd.DataFrame([[j, i, k, df_d.loc[df_d['用户名'] == i, 
                           col_list[n][m]].values[0]]],columns=column)
                    df_1 = df_1.append(df_2, ignore_index=True)
        else:
            continue
    '输出'
    df_1.to_sql(a.target, con=engine, if_exists='append', index=False, chunksize=100)
    
def connectDB():
    def login():
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
                "mssql+pymssql://@%s:%s/%s" % login())
    except Exception as e:
        print('连接失败 %s' % e)
    else:
        print('连接成功')
        return engine

def get_latest(folder):
    '获取文件夹中最新文件  暂时无用'
    files = [os.path.join(folder, f) for f in os.listdir(folder)]
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return files[0]

if __name__ == '__main__':
    try:
        # 连接 DB
        engine = connectDB()
        # 创造一个序列实列，以便生成所需要的各种列表
        star = input('输入起始日期(2019-01-01):')
        stop = input('输入终止日期(2019-01-01):')
        val = input('消费/现金？(默认消费):')
        if val == '':
            val = '消费'
        print(val)
        a = array(star, stop, val)
        if os.path.exists(a.print_path()):
            pass
        else:
            raise
    except FileNotFoundError:
        print('消费/现金文件不存在，检查文件名。')
    except Exception as e:
        print(e)
    else:
        print('SQL Server连接成功')
        run()
        read_file(a.target)
        main()
        upload_ka(star, stop)
    finally:
        print('程序结束')
    
    

    