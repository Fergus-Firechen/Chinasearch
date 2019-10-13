# -*- coding: utf-8 -*-
"""
Created on Fri Jan 11 11:30:32 2019

1.读取P4P,加款端口，Master系统账户
2.筛选：2019有消费、除端口加款、除已有账户
3.结构转换、写入excel

@author: chen.huaiyu
"""
import time, functools
import pandas as pd
import xlwings as xw

path = [r'H:\SZ_数据\Input\P4P 消费报告2019.10...xlsx', 
        r'D:\陈怀玉\工作\周工作\Payment System\Channel.csv',
        r'D:\陈怀玉\工作\周工作\Payment System\系統-_201910_Master-1.xlsm']

i = '19年已消费'

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
    
    run()
    read_file()
    main()
    pass
    
    