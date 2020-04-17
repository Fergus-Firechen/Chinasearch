# -*- coding: utf-8 -*-
"""
Created on Wed Oct 31 14:51:48 2018


- 准备阶段将所有的变量全部前置处理；处理完后一次性跑完；


@author: Fergus
"""
import os
import time
import datetime
import xlwings as xw
#import win32com.client as win32
#import pandas as pd
from xlwings import constants

start = time.perf_counter()
today = datetime.datetime.today()

# 输入
# 1.N:前1日=1；前2日=0；前3日=-1...依次类推
# 
N = 2  # 改2
# 修改 AVG.日期
yes = today-datetime.timedelta(N)

# 直接引用邮件中文件
path = r'H:\SZ_数据\Input'
name = 'P4P 消费报告' + str(yes.year) + '.' + str(yes.month) + '..F.xlsx'
wb1 = xw.Book(os.path.join(path, name))
#wb1 = xw.books(r'P4P 消费报告2019.10.15')

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


# 繁  # 改 1
path = r'C:\Users\chen.huaiyu\Downloads'
#name = 'Ave.workday&weekdayQ4(2019 Oct to Dec)2019.11.5' + '.xlsx'
name = ('Ave.workday&weekdayQ4(2019 OCT to Dec)2019' + '.' + yes.strftime('%m.%d') + '.xlsx')
wb2 = xw.Book(os.path.join(path, name))
wb2.app.calculation = 'manual'

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

sht2[row20, column20].color = (255, 255, 100)
sht3[row20, column20].color = (255, 255, 100)
sht4[row20, column20].color = (255, 255, 100)

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
        if j.value in ['北京军朗广告有限公司', '企游科技有限公司']:  # 标识军朗账户 s列
            rng = j.get_address().replace('S', 'I')
            i[rng].color = (146, 208, 80)
# wb2.save()
#wb2.app.calculate()
wb2.save()
print('耗时：{:3f}'.format((time.clock() - start)/60))

# 均值
def avg():
    str_avg_work = '=average(en3:eq3,et3:eu3,ex3)'
    str_avg_week = '=average(er3:es3,ey3:ez3)'
    
    
    '''  均值 '''
    for i in [sht2, sht3, sht4]:
    # =============================================================================
    #     # 月度总消费
    #     i['MJ3'].value = 0
    #     i['MK3'].value = 0
    # =============================================================================
        i['MN3'].formula = str_avg_work  # 工；繁版；周三；改！
        i['MO3'].formula = str_avg_week
        # '''
        i['MN3:MO3'].api.AutoFill(i['MN3:MO' + str(row201)].api, constants.AutoFillType.xlFillCopy)
        print('耗时：{:3f}'.format((time.clock() - start)/60))
    
    # avg SaudiCommissionForTourism&NationalHeritage用  hkd-sauditourism-1909  日均 *0.6   max(1,200,000)
    sht2['MN12364'].value = 0  # .formula = str_avg_work.replace('3', str(12364)) + '*0.05' # "=(SUM($BB$12364:$BD$12364)/SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))*(1200000-SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))/SUM('Date List'!$H$14:$H$15)"
    sht2['MO12364'].value = 0
    # sht2['MO12364'].formula = str_avg_week.replace('3', str(12364)) + '*0.1' # "=(SUM($BB$12364:$BD$12364)/SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))*(1200000-SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))/SUM('Date List'!$H$14:$H$15)"
    sht3['MN12364'].value = 0  # .formula = str_avg_work.replace('3', str(12364)) + '*0.05' # "=(SUM($BB$12364:$BD$12364)/SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))*(1200000-SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))/SUM('Date List'!$H$14:$H$15)"
    # sht3['MO12364'].formula = str_avg_week.replace('3', str(12364)) + '*0.1' # "=(SUM($BB$12364:$BD$12364)/SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))*(1200000-SUM($BB$12364:$BD$12364,原生广告!$BB$12364:$BD$12364))/SUM('Date List'!$H$14:$H$15)"
    sht3['MO12364'].value = 0
    sht4['MN12364'].value = 0 # = "=(SUM($BB$12364:$BD$12364)/SUM($BB$12364:$BD$12364,搜索!$BB$12364:$BD$12364))*(1200000-SUM($BB$12364:$BD$12364,搜索!$BB$12364:$BD$12364))/SUM('Date List'!$H$14:$H$15)"
    sht4['MO12364'].value = 0 # formula = "=(SUM($BB$12364:$BD$12364)/SUM($BB$12364:$BD$12364,搜索!$BB$12364:$BD$12364))*(1200000-SUM($BB$12364:$BD$12364,搜索!$BB$12364:$BD$12364))/SUM('Date List'!$H$14:$H$15)"
    
    # VenetianCotaiLimited
    for n in [9160, 9161, 9162, 9163, 9416, 10900]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.2' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.2' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.2' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.2' 
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0
    
    # 万洲金业集团有限公司
    for n in [12195, 12196, 12444]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        
    # BurberryAsiaLimited
    for n in [2157, 12316]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0 
    
    
    # 潘多拉珠宝亚太有限公司
    for n in [6284, 7441, 8093]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0 
    
    # THENET-A-PORTERGROUPLIMITED
    # 
    for n in [8412, 8464, 8785, 9048, 9231, 9402, 9834, 10324]:
# =============================================================================
#         sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.6' 
#         sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.6' 
#         sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.6' 
#         sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.6' 
# =============================================================================
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0 
    
    # MALAYSIAAIRLINESBERHAD
    for n in [7459, 9432]:
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0 
        sht2['MN' + str(n)].value = 0
        sht2['MO' + str(n)].value = 0 
        sht3['MN' + str(n)].value = 0
        sht3['MO' + str(n)].value = 0 
    
    # TOURISMANDEVENTSQUEENSLAND
    for n in [10367, 10377]:
# =============================================================================
#         sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.3' 
#         sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.3' 
#         sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.3' 
#         sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.3' 
# =============================================================================
        sht2['MN' + str(n)].value = 0
        sht2['MO' + str(n)].value = 0 
        sht3['MN' + str(n)].value = 0
        sht3['MO' + str(n)].value = 0 
        sht4['MN' + str(n)].value = 0  # .formula = str_avg_work.replace('3', str(n)) + '*0.1' 
        sht4['MO' + str(n)].value = 0  # .formula = str_avg_week.replace('3', str(n)) + '*0.1' 
    
    # AdobeSystemsHongKongLimited
# =============================================================================
#     for n in [9613, 12446]:
#         sht4['MN' + str(n)].value = 0
#         sht4['MO' + str(n)].value = 0 
#         sht2['MN' + str(n)].value = 0
#         sht2['MO' + str(n)].value = 0 
#         sht3['MN' + str(n)].value = 0
#         sht3['MO' + str(n)].value = 0 
# =============================================================================
    
    # UNIVERSITYOFSYDNEY
# =============================================================================
#     for n in [9617]:
#         sht4['MN' + str(n)].value = 0
#         sht4['MO' + str(n)].value = 0 
#         sht2['MN' + str(n)].value = 0
#         sht2['MO' + str(n)].value = 0 
#         sht3['MN' + str(n)].value = 0
#         sht3['MO' + str(n)].value = 0 
# =============================================================================
        
    # NaeHongKongLimited
# =============================================================================
#     for n in [10053, 10253, 10503, 10504, 10505, 10506, 12405, 12406, 12407, 12408]:
#         sht4['MN' + str(n)].value = 0
#         sht4['MO' + str(n)].value = 0 
# =============================================================================
    
    # StudentUniverse.comInc
# =============================================================================
#     for n in [6273, 10092]:
#         sht4['MN' + str(n)].value = 0
#         sht4['MO' + str(n)].value = 0 
# =============================================================================
    
    # VIAGOGOAG 1  11488, 
    # LenzingAktiengesellschaft 1
    # AutismPartnershipLimited 1  12108,
    # MichaelPageInternational(hongkong)limited 1
# =============================================================================
#     for n in [12035, 18]:
#         sht4['MN' + str(n)].value = 0
#         sht4['MO' + str(n)].value = 0 
# =============================================================================
    
    # Etoro(Uk)Limited 1
    # 万代南梦宫娱乐香港有限公司 1 
    # AkamaiTechnologies.Inc 1
    for n in [12350, 10211, 9540]:
        sht2['MN' + str(n)].value = 0
        sht2['MO' + str(n)].value = 0 
        sht3['MN' + str(n)].value = 0
        sht3['MO' + str(n)].value = 0 
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0 
    
    # 香港鑫圣金业集团有限公司
    for n in [2342, 2777, 6348, 6415, 6416, 6417, 6418]:
# =============================================================================
#         sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n))
#         sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n))
#         sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n))
#         sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n))
# =============================================================================
        sht2['MN' + str(n)].value = 0
        sht2['MO' + str(n)].value = 0
        sht3['MN' + str(n)].value = 0
        sht3['MO' + str(n)].value = 0 
        sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.8' 
        sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.8' 
    for n in [6418]:
        sht2['MN' + str(n)].value = 17000
        sht2['MO' + str(n)].value = 12000
    
    # Worldfirst
    for n in [7900, 9916, 11051, 12230]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5'

    
    # ACCEL(HK)COMPANYLIMITED
    for n in [10082]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
    
    # 香港金盛贵金属有限公司
    for n in [8541, 8763, 8926, 10102, 10103, 10104, 10105]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
        sht4['MN' + str(n)].value = 0 
        sht4['MO' + str(n)].value = 0
    
    # 国际移动娱乐有限公司
    for n in [9941, 10019, 10089, 10603, 10641, 10694, 11181, 11182, 11214, 11215, 12207, 12329]:        
# =============================================================================
#         sht2['MN' + str(n)].value = 0 
#         sht2['MO' + str(n)].value = 0        
#         sht3['MN' + str(n)].value = 0 
#         sht3['MO' + str(n)].value = 0        
#         sht4['MN' + str(n)].value = 0
#         sht4['MO' + str(n)].value = 0
# =============================================================================
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0
        # sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
        # sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
    
    # 瑞丰德永国际商务（中国）有限公司
    for n in [281, 6406, 9747, 9748, 12219]:      
        sht2['MN' + str(n)].value = 0 
        sht2['MO' + str(n)].value = 0        
        sht3['MN' + str(n)].value = 0 
        sht3['MO' + str(n)].value = 0        
        sht4['MN' + str(n)].value = 0 
        sht4['MO' + str(n)].value = 0
    for n in [12219]: 
        sht2['MN' + str(n)].value = 17000
        sht2['MO' + str(n)].value = 2000    
        sht3['MN' + str(n)].value = 0 
        sht3['MO' + str(n)].value = 0        
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0
# =============================================================================
#         sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
#         sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
#         sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
#         sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
#         sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.7' 
#         sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.7' 
# =============================================================================

    # 金荣中国
    # 客户预算砍半
    for n in [2329, 2458, 7941, 7942]:  
        sht2['MN' + str(n)].value = 0 
        sht2['MO' + str(n)].value = 0        
        sht3['MN' + str(n)].value = 0 
        sht3['MO' + str(n)].value = 0        
        sht4['MN' + str(n)].value = 0 
        sht4['MO' + str(n)].value = 0
# =============================================================================
#         sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
#         sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
#         sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
#         sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
#         sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
#         sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
# =============================================================================

    # 兴业证劵有限公司
# =============================================================================
#     for n in [3563]: # 3245, 3246, 3540, 3808, 
#         sht2['MN' + str(n)].value = 7000 
#         sht2['MO' + str(n)].value = 7000     
#         sht3['MN' + str(n)].value = 0 
#         sht3['MO' + str(n)].value = 0
#         sht4['MN' + str(n)].value = 0 
#         sht4['MO' + str(n)].value = 0
# =============================================================================
    
    # GlobalKapitalHoldingsLtd.  2
    # InterechoLimited 1
    # AtGlobalMarkets(Uk)Limited 1
    # hkd-xuanhot-1902 1
    # hkd-huitoutiao-1810 1
    for n in [12445, 12482, 12478, 12487, 11873, 10936]: # 3245, 3246, 3540, 3808, 
        sht2['MN' + str(n)].value = 0
        sht2['MO' + str(n)].value = 0     
        sht3['MN' + str(n)].value = 0 
        sht3['MO' + str(n)].value = 0
        sht4['MN' + str(n)].value = 0 
        sht4['MO' + str(n)].value = 0
    
    # 艾德金业有限公司
    # GroupeKedgeBusinessSchool 1
    for n in [12100, 12428, 10091]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.6' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.6' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.6' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.6' 
        sht4['MN' + str(n)].value = 0
        sht4['MO' + str(n)].value = 0 
        
        
    # SONDERCLOUDLIMITED
    for n in [8603, 10385]:
        sht2['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht2['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht3['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht3['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        sht4['MN' + str(n)].formula = str_avg_work.replace('3', str(n)) + '*0.5' 
        sht4['MO' + str(n)].formula = str_avg_week.replace('3', str(n)) + '*0.5' 
        
    # 忽略
    # 财团法人台湾贸易中心
    # 金道贵金属有限公司
    
    # hkd-marriottfliggy-1810
    

# =============================================================================
#     
# avg()
# wb2.save()
# =============================================================================
print('耗时：{:3f}'.format((time.clock() - start)/60))
