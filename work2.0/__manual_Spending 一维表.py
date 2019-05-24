# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 11:22:45 2019

- 跨月数据由于百通在不同表中，无法一次性处理

@author: chen.huaiyu
"""
from datetime import datetime, timedelta
from sqlalchemy import create_engine
import time, functools, os
import pandas as pd


class array(object):
    
    def __init__(self, beforeDays, continueDays, target):
        if isinstance(beforeDays, int) and isinstance(continueDays, int):
            self.beforeDays = beforeDays
            self.continueDays = continueDays
        else:
            raise "只能是整数"
        if target in ['消费', '现金']:
            self.target = target
        else:
            raise "只能是'消费'/'现金'"
        
    def date(self):
        # 生成：目标日期区间
        date_str = (datetime.today() - timedelta(self.beforeDays)
                    ).strftime("%Y-%m-%d")
        date = [i.strftime('%Y%m%d') for i in pd.date_range(
                start=date_str, periods=self.continueDays)]
        return date
    
    def field(self):
        # 字段集
        col = []
        for i in range(6):
            col.append(list(map(lambda x: self.print_col()[i] + self.target + x, self.date())))
        return col
    
    def print_date(self):
        return datetime.today() - timedelta(self.beforeDays)
    
    def print_col(self):
        return ['总点击', '搜索点击', '自主投放', '新产品', '百通', '无线搜索点击']
    
    def print_path(self):
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
        return path
                                                               
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

@cost_time
def read_file():
    global df1, df2, df3
    '==icrm 消费csv=='
    df1 = pd.read_csv(a.print_path(), engine='python', encoding='gbk')
    df1.rename(columns={'账户名称':'用户名'}, inplace=True)
    df1['用户名'] = df1['用户名'].astype(str)
    df1.drop(columns='账户ID', inplace=True)
    '==基本信息=='
    df2 = pd.read_excel(path[1], sheet_name='基本信息')
    df2['用户名'] = df2['用户名'].astype(str)
    df2 = df2.loc[8:, ['区域', '用户名', '广告主']]
    '==百通=='
    sheet = 'P4P消费'+str(a.print_date().month)+'月'
    df3 = pd.read_excel(path[2], sheet_name=sheet).iloc[38:52,:]
    '结构整理'
    df3.iloc[0, 0] = '用户名'
    df3.columns = df3.iloc[0, :].tolist()
    df3.drop(index=[38], inplace=True)
    df3.drop(columns='总计', inplace=True)
    df3.set_index('用户名', drop=True, inplace=True)
    df3.columns = [i.strftime('%Y%m%d') for i in df3.columns]
    df3 = df3.loc[:, a.date()[0]:a.date()[-1]]
    df3.columns = a.field()[4]
    
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
        for j in df3.index:
            df1.loc[j, col_all[i]] = df1.loc[j, col_all[i]] + df3.loc[j, 
                                       col_bt[i]]
            df1.loc[j, col_np[i]] = df1.loc[j, col_np[i]] + df3.loc[j, 
                                       col_bt[i]]
    '消费文件'
    df1.reset_index(inplace=True)
    df1 = pd.merge(df1, df3, how='left', on='用户名')
    df1.fillna(0, inplace=True)
    '输出'
    df1.to_csv(r'c:\users\chen.huaiyu\Desktop\p4p 消费报告.csv', encoding='GBK')
    df = pd.merge(df2, df1, on='用户名', how='left')
    '转换为一维表'
    column = ['日期', '用户名', '类别', '金额']
    df_1 = pd.DataFrame(columns=column)
    col_list = [col_all, col_p4p, col_inf, col_np, col_bt, col_mo]
    for m,j in enumerate(a.date()):
        df_d = df[df[col_all[m]] > 0]
        if df_d.shape[0] > 10:
            for i in df_d['用户名'].tolist():
                for n, k in enumerate(a.print_col()):
                    df_2 = pd.DataFrame([[j, i, k, df_d.loc[df_d['用户名'] == i, 
                           col_list[n][m]].values[0]]],columns=column)
                    df_1 = df_1.append(df_2, ignore_index=True)
        else:
            continue
    '输出'
    df_1.to_sql(a.target, con=engine, if_exists='append', index=False)
    
if __name__ == '__main__':
    
    try:
        engine = create_engine("mssql+pyodbc://SQL Server")
    except:
        raise "SQL Server连接失败"
    else:
        print('SQL Server连接成功')
        
        path = ['',  ## 改
            r'H:\SZ_数据\基本信息拆解.xlsx',
            r'H:\SZ_数据\Input\每日百度消费.xlsx']
        
        run()
        
        # 创造一个序列实列，以便生成所需要的各种列表
        a = array(23, 22, '现金')
        
        read_file()
    
        main()
    finally:
        print('程序结束')
    
    

    