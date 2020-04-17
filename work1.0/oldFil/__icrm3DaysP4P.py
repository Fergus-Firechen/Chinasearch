# -*- coding: utf-8 -*-
"""
Created on Tue Oct 30 11:27:38 2018
# 处理icrm下载的p4p 如周一，需连续调整多日数据顺序；
@author: chen.huaiyu
"""
import pandas as pd
import time, datetime


start = time.clock()
dateToday = datetime.datetime.today()
dateToday = dateToday.replace(dateToday.year, dateToday.month, dateToday.day, 0, 0, 0, 0)
date = datetime.datetime.strftime(dateToday - datetime.timedelta(1), '%Y%m%d')
print('近3天')
startDay = datetime.datetime.strftime(datetime.datetime.today() - datetime.timedelta(3), '%Y%m%d')
# icrm = pd.read_csv(r'C:\Users\chen.huaiyu\Downloads\fq (1).csv', engine='python', encoding='gbk')
icrm = pd.read_csv(r'C:\Users\chen.huaiyu\Downloads\p4p ' + startDay + '_' + date + '.csv', encoding='gbk')

icrm.rename(columns={'账户名称':'用户名'}, inplace=True)
# 直接使用已运行后的变量
try:
    xiaofei1 = pd.DataFrame(basicMessage.loc[:, '用户名'])
except:
    xiaofei1 = pd.read_excel(r'H:\SZ_数据\Input\P4P 消费报告2019.' + str(datetime.datetime.today().month) + '...xlsx', sheet_name=2, usecols=[9])
else:
    print('basicMessage正常。')
xiaofei1['用户名'] = xiaofei1['用户名'].astype(str)
xiaoFeiName = xiaofei1.iloc[2:,:]
merge1 = pd.merge(xiaoFeiName, icrm, how='left', on='用户名')

# =============================================================================
# merge1.to_excel(r'C:\Users\chen.huaiyu\Desktop\Output\p4p数据1.xlsx')
# =============================================================================

# 计算新产品消费
merge1['新'+startDay] = merge1['总点击消费'+startDay]-merge1['搜索点击消费'+startDay]-merge1['自主投放消费'+startDay]
merge1['新'+str(int(startDay)+1)] = merge1['总点击消费'+str(int(startDay)+1)]-merge1['搜索点击消费'+str(int(startDay)+1)]-merge1['自主投放消费'+str(int(startDay)+1)]
merge1['新'+str(int(startDay)+2)] = merge1['总点击消费'+str(int(startDay)+2)]-merge1['搜索点击消费'+str(int(startDay)+2)]-merge1['自主投放消费'+str(int(startDay)+2)]

# 计入百通消费
dataBT = pd.read_excel(r'H:\SZ_数据\Input\每日百度消费.xlsx', 
                           sheet_name='P4P消费'+str(dateToday.month)+'月'
                           ).iloc[38:52, :]
dataBT.iloc[0, 0] = '用户名'
dataBT.iloc[-1, 0] = 'Total'
dataBT1 = dataBT.T.set_index(38, drop=True)
dataBT1 = dataBT1.T
# =============================================================================
# dataBT.set_index('用户名', inplace=True)
# =============================================================================

# 取百通近3天的消费数据
last3Days = dataBT1.loc[:, dateToday-datetime.timedelta(3):dateToday-datetime.timedelta(1)]
last3Days.columns = ['百通'+datetime.datetime.strftime(i, '%Y%m%d') for i in last3Days.columns]
last3Days['用户名'] = list(dataBT1['用户名'].values)

# 合并百通消费数据
# =============================================================================
# merge1.set_index('用户名', inplace=True)
# =============================================================================
mergeBaiTong = pd.merge(merge1, last3Days, how='left', on='用户名')
mergeBaiTong.fillna(0, inplace=True)

# +百通消费
mergeBaiTong['总_'+startDay] = mergeBaiTong['总点击消费'+startDay] + mergeBaiTong['百通'+startDay]
mergeBaiTong['总_'+str(int(startDay)+1)] = mergeBaiTong['总点击消费'+str(int(startDay)+1)] + mergeBaiTong['百通'+str(int(startDay)+1)]
mergeBaiTong['总_'+str(int(startDay)+2)] = mergeBaiTong['总点击消费'+str(int(startDay)+2)] + mergeBaiTong['百通'+str(int(startDay)+2)]

mergeBaiTong['新_'+startDay] = mergeBaiTong['新'+startDay] + mergeBaiTong['百通'+startDay]
mergeBaiTong['新_'+str(int(startDay)+1)] = mergeBaiTong['新'+str(int(startDay)+1)] + mergeBaiTong['百通'+str(int(startDay)+1)]
mergeBaiTong['新_'+str(int(startDay)+2)] = mergeBaiTong['新'+str(int(startDay)+2)] + mergeBaiTong['百通'+str(int(startDay)+2)]

mergeBaiTong.to_excel(r'C:\Users\chen.huaiyu\Desktop\Output\p4p数据1.xlsx', freeze_panes=(1,0))

import xlwings as xw
wb = xw.Book(r'C:\Users\chen.huaiyu\Desktop\Output\p4p数据1.xlsx')
sht = wb.sheets[0]
sht.autofit()
wb.app.calculation = 'manual'
wb.save()
print('耗时：{:.3f}min'.format((time.clock()-start)/60))