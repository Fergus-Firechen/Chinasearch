# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 11:22:45 2019

- 跨月数据由于百通在不同表中，无法一次性处理

@author: chen.huaiyu
"""

import time, functools, datetime
import pandas as pd
import xlwings as xw

year = (datetime.date.today()-datetime.timedelta(1)).year
month = (datetime.date.today()-datetime.timedelta(1)).month

path = [r'C:\Users\chen.huaiyu\Downloads\p4p 20190519_20190521.csv',  ## 改
        r'H:\SZ_数据\基本信息拆解.xlsx',
        r'H:\SZ_数据\Input\每日百度消费.xlsx']

'日期序列'
date = pd.date_range(start='2019-05-19', end='2019-05-21')  ## 改
date = [i.strftime('%Y%m%d') for i in date]

'消费字段序列'
col1 = ['总点击消费', '搜索点击消费', '自主投放消费', '新产品消费', '百通消费', 
        '无线搜索点击消费']

'字段集'
col_all = list(map(lambda x:col1[0]+x, date))
col_p4p = list(map(lambda x:col1[1]+x, date))
col_inf = list(map(lambda x:col1[2]+x, date))
col_np = list(map(lambda x:col1[3]+x, date))
col_bt = list(map(lambda x:col1[4]+x, date))
col_mo = list(map(lambda x:col1[5]+x, date))


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
    df1 = pd.read_csv(path[0], engine='python', encoding='gbk')
    df1.rename(columns={'账户名称':'用户名'}, inplace=True)
    df1['用户名'] = df1['用户名'].astype(str)
    df1.drop(columns='账户ID', inplace=True)
    '==基本信息=='
    df2 = pd.read_excel(path[1], sheet_name='基本信息')
    df2['用户名'] = df2['用户名'].astype(str)
    df2 = df2.loc[8:, ['区域', '用户名', '广告主']]
    '==百通=='
    sheet = 'P4P消费'+str(month)+'月'
    df3 = pd.read_excel(path[2], sheet_name=sheet).iloc[38:52,:]
    '结构整理'
    df3.iloc[0, 0] = '用户名'
    df3.columns = df3.iloc[0, :].tolist()
    df3.drop(index=[38], inplace=True)
    df3.drop(columns='总计', inplace=True)
    df3.set_index('用户名', drop=True, inplace=True)
    df3.columns = [i.strftime('%Y%m%d') for i in df3.columns]
    df3 = df3.loc[:, date[0]:date[-1]]
    df3.columns = col_bt
    
@cost_time
def main():
    global df1
    '新产品计算'
    for i in range(len(date)):
        df1[col_np[i]] = df1[col_all[i]] - df1[col_p4p[i]] - df1[col_inf[i]]
    '+百通'
    df1.set_index('用户名', drop=True, inplace=True)
    for i in range(len(date)):
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
    for m,j in enumerate(date):
        df_d = df[df[col_all[m]] > 0]
        if df_d.shape[0] > 10:
            for i in df_d['用户名'].tolist():
                for n, k in enumerate(col1):
                    df_2 = pd.DataFrame([[j, i, k, df_d.loc[df_d['用户名'] == i, 
                           col_list[n][m]].values[0]]],columns=column)
                    df_1 = df_1.append(df_2, ignore_index=True)
        else:
            continue
    '输出'
    wb = xw.Book(path[1])
    wb.app.visible = False
    sht = wb.sheets['Spending']
    rng = sht[0, 0].current_region
    row = rng.rows.count
    sht[row, 0].color = (255, 255, 0)
    sht[row, 0].value = df_1.values
    sht[0, 0].current_region.column_width = 12
    wb.save()
    wb.close()
        
    
if __name__ == '__main__':
    
    run()
    
    read_file()
    
    main()