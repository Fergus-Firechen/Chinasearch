# -*- coding: utf-8 -*-
"""
Created on Wed Oct 31 14:51:48 2018


- 准备阶段将所有的变量全部前置处理；处理完后一次性跑完；


@author: Fergus
"""

import datetime, time
import xlwings as xw
import win32com.client as win32
import pandas as pd
from xlwings import constants

start = time.perf_counter()
today = datetime.datetime.today()

# 输入
# 1.N:前1日=1；前2日=0；前3日=-1...依次类推
# 
N = 2
# 修改 AVG.日期
Yesterday = today-datetime.timedelta(N)

# 直接引用邮件中文件
wb1 = xw.books('P4P 消费报告2019.4.15.xlsx')

sht1 = wb1.sheets['P4P消费']
sht20 = wb1.sheets['搜索点击消费']
sht30 = wb1.sheets['新产品消费（除原生广告）']
sht40 = wb1.sheets['原生广告']



# 1 读取值
# 1.1 基本信息区域
row10 = 10
row11 = sht1[0, 0].current_region.rows.count
range10 = sht1['A' + str(row10) + ':AC' + str(row11)]  # 消费报告基本信息区域
# 1.2 数据区域
# 1.2.1 首列
date10 = today - datetime.timedelta(N)  # 改1.1
date10 = date10.replace(date10.year, date10.month, date10.day, 0, 0, 0, 0)
date11 = date10.replace(date10.year, date10.month, 1)  # 本月第一天
column11 = sht1[0, :].value.index(date11)
# 1.2.2 末列 
date12 = date10 - datetime.timedelta(N)
date13 = date12 - datetime.timedelta(1)
column12 = sht1[0, :].value.index(date10)

# 1.2.3 取值 前天
'''
range20 = sht20[9:row11, column11:column12]
range30 = sht30[9:row11, column11:column12]
range40 = sht40[9:row11, column11:column12]
'''

#2 拿取月底最后一天的值  11.30
range20 = sht20[9:row11, column11:column12 + 1]
range30 = sht30[9:row11, column11:column12 + 1]
range40 = sht40[9:row11, column11:column12 + 1]


# 繁
wb2 = xw.Book(r'C:\Users\chen.huaiyu\Downloads\Ave.workday&weekdayQ2(2019 Apr_Jun)2019.04.15.xlsx')

# 简 wb2 = xw.Book(r'C:\Users\chen.huaiyu\Downloads\Ave.workday&weekdayQ4- 2018.12.23(simplified)-v1.xlsx')
# wb2 = xw.books('Ave.workday&weekdayQ4- 2018.12.23(simplified)')

sht2 = wb2.sheets['搜索']
sht3 = wb2.sheets['其他新产品']
sht4 = wb2.sheets['原生广告']
sht5 = wb2.sheets['Date List']

# 写入前总行数
row200 = sht2[1, 0].current_region.rows.count - 1  # 写入前
print('写入前账户数：%s' %row200)
sht2[row200+1, 0].color = (255, 255, 0)
column20 = sht2[1, :].value.index(date11)  # 当月首日列      运1.2
column200 = 29  # AD列值
column21 = sht2[1, :].value.index(date10)  # 当月消费日末列

# 2.1 基本信息位置
sht2[2, 0].value = range10.value
sht3[2, 0].value = range10.value
sht4[2, 0].value = range10.value

sht2[2, 0].color = (255, 255, 0)
sht3[2, 0].color = (126, 126, 165)
sht4[2, 0].color = (126, 126, 165)
# 2.2 数据填充
# 2.2.1 位置坐标
row20 = 2
row201 = sht2[1, 0].current_region.rows.count  # 写入后
# 2.2.2 赋值
sht2[row20, column20].value = range20.value
sht3[row20, column20].value = range30.value
sht4[row20, column20].value = range40.value

sht2[row20, column20].color = (255, 255, 0)
sht3[row20, column20].color = (255, 255, 0)
sht4[row20, column20].color = (255, 255, 0)

# 3 空格自动向下填充
# 3.1 区域
# 3.1.1 写入前，计划填充列
range220 = sht2[row200, column200:column20]
range230 = sht3[row200, column200:column20]
range240 = sht4[row200, column200:column20]
# 3.1.2 写入后，被填充列
range221 = sht2[row200:row201, column200:column20]
range231 = sht3[row200:row201, column200:column20]
range241 = sht4[row200:row201, column200:column20]
# 3.2 前值填充
if range220 == range221:
    pass
else:
    range220.api.AutoFill(range221.api, constants.AutoFillType.xlFillCopy)
    range230.api.AutoFill(range231.api, constants.AutoFillType.xlFillCopy)
    range240.api.AutoFill(range241.api, constants.AutoFillType.xlFillCopy)
    # 3.3 后值填充
    range222 = sht2[row200, column21 + 1:]
    range232 = sht3[row200, column21 + 1:]
    range242 = sht4[row200, column21 + 1:]
    
    range223 = sht2[row200:row201, column21 + 1:]
    range233 = sht3[row200:row201, column21 + 1:]
    range243 = sht4[row200:row201, column21 + 1:]
    
    range222.api.AutoFill(range223.api, constants.AutoFillType.xlFillCopy)
    range232.api.AutoFill(range233.api, constants.AutoFillType.xlFillCopy)
    range242.api.AutoFill(range243.api, constants.AutoFillType.xlFillCopy)

star1 = (time.perf_counter() - start)/60

# 加边框
# 军朗填充为绿色
for i in [sht2, sht3, sht4]:
    for j in range(7, 13):
        i[row200:row201, :].api.Borders(j).LineStyle = 1
        i[row200:row201, :].current_region.api.Borders(j).weight = 2
    for j in i['S' + str(row200) + ':S' + str(row201)]:
        if j.value == '北京军朗广告有限公司':  # 标识军朗账户 s列
            rng = j.get_address().replace('S', 'I')
            i[rng].color = (146, 208, 80)
wb2.save()

# 最近2日列
co = 'bv'
lu = 'bs'


'''  均值 '''
for i in [sht2, sht3, sht4]:
    i['MD3'].formula = '=average(bk3,bo3:bs3,bv3)'  # 工；繁版；周三；改！
    i['ME3'].formula = '=average(bm3:bn3,bt3:bu3)'
    # '''
    i['MD3:ME3'].api.AutoFill(i['MD3:ME' + str(row201)].api, constants.AutoFillType.xlFillCopy)
    print('耗时：{:3f}'.format((time.clock() - start)/60))
    
wb2.save()
