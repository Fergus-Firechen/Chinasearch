# -*- coding: utf-8 -*-
"""
Created on Thu Jan  3 19:12:56 2019

@author: chen.huaiyu
"""

''' 手动导入数据表  '''

import pyodbc, os, time
import xlwings as xw
from xlwings import constants

star = time.clock()

DATE = '20190415'  # 改1
target = r'C:\Users\chen.huaiyu\Desktop\Output\SQL Server' + '\\' + DATE
server_name = 'SZ-CS-0038LT\SQLEXPRESS'
db_name = 'CSA_2019Q2'
db_name_table_1 = 'P4P_20181211_1'.replace('20181211', DATE)
db_name_talbe_2 = 'P4P_20181211_2'.replace('20181211', DATE)

cnxn = pyodbc.connect('Driver={SQL Server};'
                      'Server='+server_name+';'
                      'Database='+db_name+';'
                      'Trusted_Connection=yes;'
                      )
cursor = cnxn.cursor()


# Merge P4P & NP & Infeeds
cursor.execute('''
                select a.*, b.* into P4P_20190402 from P4P_20190402_1 a inner join P4P_20190402_2 b on a.用户名=b.用户名1
                alter table P4P_20190402 drop column 用户名1
                
                alter table P4P_20190402 add Forex varchar
                update P4P_20190402
                set Forex = b.Forex
                from P4P_20190402 a inner join Forex1 b on a.用户名=b.用户名1
                
                
                
                select a.*, b.* into NP_20190402 from NP_20190402_1 a inner join NP_20190402_2 b on a.用户名=b.用户名1
                alter table NP_20190402 drop column 用户名1
                
                alter table NP_20190402 add Forex varchar
                update NP_20190402
                set Forex = b.Forex
                from NP_20190402 a inner join Forex1 b on a.用户名=b.用户名1
                
                
                
                select a.*, b.* into Infeeds_20190402 from Infeeds_20190402_1 a inner join Infeeds_20190402_2 b on a.用户名=b.用户名1
                alter table Infeeds_20190402 drop column 用户名1
                
                alter table Infeeds_20190402 add Forex varchar
                update Infeeds_20190402
                set Forex = b.Forex
                from Infeeds_20190402 a inner join Forex1 b on a.用户名=b.用户名1
               '''.replace('20190402', DATE))
cnxn.commit()

# try:
# Excel File:Spending Forecast
# Pivot Table(端口)
# P4P

print('注1：\nSpending Forecast')
wb = xw.Book(target+'\\'+os.listdir(target)[0])
cursor.execute('''
               select 端口, sum([2019Jan]) as Jan_19,
                sum([2019Feb]) as Feb_19,
                sum([2019Mar]) as Mar_19,
                sum([Apr Spending Forecast]) as Apr_19,
                sum([May Spending Forecast]) as May_19,
                sum([Jun Spending Forecast]) as Jun_19,
                sum([Jul Spending Forecast]) as Jul_19,
                sum([Aug Spending Forecast]) as Aug_19,
                sum([Sep Spending Forecast]) as Sep_19,
                sum([Oct Spending Forecast]) as Oct_19,
                sum([Nov Spending Forecast]) as Nov_19,
                sum([Dec Spending Forecast]) as Dec_19
                from P4P_20190402
                where 端口 not like '%wrong%'
                group by 端口
                order by 端口
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht = wb.sheets['P4P']
sht[1, 0].value = list(map(lambda x:list(x), item))
print('Spending Forecast P4P Port Count: %s' % sht[1, 0].current_region.rows.count)

# NP
cursor.execute('''
select 端口, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from NP_20190402
where 端口 not like '%wrong%'
group by 端口
order by 端口
              '''.replace('20190402', DATE))
item = cursor.fetchall()
sht = wb.sheets['NP']
sht[1, 0].value = list(map(lambda x:list(x), item))
print('Spending Forecast NP Port Count: %s' % sht[1, 0].current_region.rows.count)

# Infeeds
cursor.execute('''
select 端口, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from Infeeds_20190402
where 端口 not like '%wrong%'
group by 端口
order by 端口
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht = wb.sheets['Infeeds']
sht[1, 0].value = list(map(lambda x:list(x), item))
print('Spending Forecast Infeeds Port Count: %s' % sht[1, 0].current_region.rows.count)

# 填充 All
sht1 = wb.sheets['All']
rows100 = sht1['A1'].current_region.rows.count
rows101 = sht[1, 0].current_region.rows.count
sht1['A2'].options(transpose=True).value = sht['A2:A'+str(rows101)].value
if rows101 > rows100:
    range100 = sht1['B'+str(rows100)+':M'+str(rows100)]
    range101 = sht1['B'+str(rows100)+':M'+str(rows101)]
    range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)

# Infeeds(35%)
sht1 = wb.sheets['Infeeds(35%)']
rows100 = sht1['A1'].current_region.rows.count
sht1[1, 0].options(transpose=True).value = sht[1:rows101,0].value
if rows101 > rows100:
    range100 = sht1['B'+str(rows100)+':M'+str(rows100)]
    range101 = sht1['B'+str(rows100)+':M'+str(rows101)]
    range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)
    
# Infeeds(27%)
sht1 = wb.sheets['Infeeds(27%)']
rows100 = sht1['A1'].current_region.rows.count
sht1['A2'].options(transpose=True).value = sht[1:rows101, 0].value
if rows100 < rows101:
    range100 = sht1['B'+str(rows100)+':M'+str(rows100)]
    range101 = sht1['B'+str(rows100)+':M'+str(rows101)]
    range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)

    
# Pivot Table (用户名)
# P4P
cursor.execute('''
select 用户名, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from P4P_20190402
where 端口 not like '%wrong%'
group by 用户名
order by 用户名
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht = wb.sheets['P4P']
sht['O2'].value = list(map(lambda x:list(x), item))
print('Spending Forecast P4P User Count: %s' % sht['O1'].current_region.rows.count)

# NP
cursor.execute('''
select 用户名, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from NP_20190402
where 端口 not like '%wrong%'
group by 用户名
order by 用户名

               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht = wb.sheets['NP']
sht['O2'].value = list(map(lambda x:list(x), item))
print('Spending Forecast NP User Count: %s' % sht['O1'].current_region.rows.count)


# Infeeds
cursor.execute('''
select 用户名, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from Infeeds_20190402
where 端口 not like '%wrong%'
group by 用户名
order by 用户名
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht = wb.sheets['Infeeds']
sht['O2'].value = list(map(lambda x:list(x), item))
print('Spending Forecast Infeeds User Count: %s' % sht['O1'].current_region.rows.count)


# ALL
sht1 = wb.sheets['All']
rows100 = sht1['O1'].current_region.rows.count
rows101 = sht['O1'].current_region.rows.count
sht1['O2'].options(transpose=True).value = sht['O2:O'+str(rows101)].value
if rows101 > rows100:
    range100 = sht1['P'+str(rows100)+':AA'+str(rows100)]
    range101 = sht1['P'+str(rows100)+':AA'+str(rows101)]
    range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)

# Infeeds(35%)
sht1 = wb.sheets['Infeeds(35%)']
rows100 = sht1['O1'].current_region.rows.count  # 计数：用户名
sht1['O2'].options(transpose=True).value = sht['O2:O'+str(rows101)].value  # 更新用户名
if rows100 < rows101:  # 如用户名有新增；则向下填充公式
    range100 = sht1['P'+str(rows100)+':AA'+str(rows100)]
    range101 = sht1['P'+str(rows100)+':AA'+str(rows101)]
    range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)
# Infeeds(27%)
sht1 = wb.sheets['Infeeds(27%)']
rows100 = sht1['O1'].current_region.rows.count
sht1['O2'].options(transpose=True).value = sht['O2:O'+str(rows101)].value
if rows101 > rows100:
    range100 = sht1['P'+str(rows100)+':AA'+str(rows100)]
    range101 = sht1['P'+str(rows100)+':AA'+str(rows101)]
    range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)


# Pivot Table (区域)
# P4P
cursor.execute('''
select 区域, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from P4P_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht = wb.sheets['Region']
sht['A3'].value = list(map(lambda x:list(x), item))
# NP
cursor.execute('''
select 区域, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from NP_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A12'].value = list(map(lambda x:list(x), item))
# Infeeds
cursor.execute('''
select 区域, sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from Infeeds_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A21'].value = list(map(lambda x:list(x), item))
wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File: Spending Forecast _v1
# P4P
print('注2：\nSpending Forecast _v1')
wb = xw.Book(target+'\\'+os.listdir(target)[1])
sht = wb.sheets['P4P']
cursor.execute('''
select AM, 端口, 用户名, 广告主, sum([Ave# Daily Workday]) as [avg_daily_workday], sum([Ave# Daily Holiday]) as [avg_daily_holiday],
sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from P4P_20190402
where 端口 not like '%wrong%'
group by AM, 端口, 用户名, 广告主
order by AM, 端口, 用户名, 广告主
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A2'].value = list(map(list, item))
print('sht P4P rows counts:%s' %sht['A1'].current_region.rows.count)

# NP
sht = wb.sheets['NP']
cursor.execute('''
select AM, 端口, 用户名, 广告主, sum([Ave# Daily Workday]) as [avg_daily_workday], sum([Ave# Daily Holiday]) as [avg_daily_holiday],
sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from NP_20190402
where 端口 not like '%wrong%'
group by AM, 端口, 用户名, 广告主
order by AM, 端口, 用户名, 广告主
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A2'].value = list(map(list, item))
print('sht NP rows counts:%s' %sht['A1'].current_region.rows.count)

# Infeeds
sht = wb.sheets['Infeeds']
cursor.execute('''
select AM, 端口, 用户名, 广告主, sum([Ave# Daily Workday]) as [avg_daily_workday], sum([Ave# Daily Holiday]) as [avg_daily_holiday],
sum([2019Jan]) as Jan_19,
sum([2019Feb]) as Feb_19,
sum([2019Mar]) as Mar_19,
sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Jul Spending Forecast]) as Jul_19,
sum([Aug Spending Forecast]) as Aug_19,
sum([Sep Spending Forecast]) as Sep_19,
sum([Oct Spending Forecast]) as Oct_19,
sum([Nov Spending Forecast]) as Nov_19,
sum([Dec Spending Forecast]) as Dec_19
from Infeeds_20190402
where 端口 not like '%wrong%'
group by AM, 端口, 用户名, 广告主
order by AM, 端口, 用户名, 广告主
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A2'].value = list(map(list, item))
print('Spending Forcast v1 Infeeds counts: %s' % sht['A1'].current_region.rows.count)


# ALL
rows101 = sht['A1'].current_region.rows.count
sht = wb.sheets['All']
rows100 = sht['A1'].current_region.rows.count
if rows100 < rows101:
    range100 = sht['A'+str(rows100)+':R'+str(rows100)]
    range101 = sht['A'+str(rows100)+':R'+str(rows101)]
    range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File: 2018Q4 Cash Spend Forecast
wb = xw.Book(target+'\\'+os.listdir(target)[3])
sht = wb.sheets['Region']
# P4P
sql = '''
select 区域, sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD
from P4P_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
   '''.replace('20190402', DATE)
   
cursor.execute(sql)
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))

# NP
cursor.execute('''
select 区域, sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD
from NP_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
               '''.replace('20190402', DATE))             
item = cursor.fetchall()
sht['A12'].value = list(map(list, item))

# Infeeds
cursor.execute('''
select 区域, sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD
from Infeeds_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A21'].value = list(map(list, item))

# Finance Region
sht = wb.sheets['Finance Region']
# P4P
cursor.execute('''
select 财务做账区域, 区域, sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD
from P4P_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 财务做账区域, 区域
order by 财务做账区域, 区域
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))

# NP
cursor.execute('''
select 财务做账区域, 区域, sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD
from NP_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 财务做账区域, 区域
order by 财务做账区域, 区域
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A18'].value = list(map(list, item))

# Infeeds
cursor.execute('''
select 财务做账区域, 区域, sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD
from Infeeds_20190402
where 端口 not like '%wrong%' and 区域 <> '-'
group by 财务做账区域, 区域
order by 财务做账区域, 区域

               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A33'].value = list(map(list, item))

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File:Handling Fee
# AM Handling Fee
wb = xw.Book(target+'\\'+os.listdir(target)[5])
sht = wb.sheets['AM']
# P4P
cursor.execute('''
select a.AM, sum([Q2 Total Spending Forecast]) as Q2_Total_Spending_Forecast, 
sum([Apr_FX GAIN(RMB)]) as Apr_FX_Gain_RMB, sum([May_FX GAIN(RMB)]) as May_FX_Gain_RMB, sum([Jun_FX GAIN(RMB)]) as Jun_FX_Gain_RMB,
sum([FX GAIN(RMB)]) as FX_Gain_RMB, sum([FX GAIN(RMB)])/0.18 as FX_Gain_Spending_RMB, 
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD,
sum([QTD FX GAIN (RMB)]) as QTD_FX_Gain_RMB, sum([QTD FX GAIN (RMB)])/0.18 as QTD_FX_Gain_Spending_RMB
from P4P_20190402 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Group, a.AM
order by b.AM_Group, a.AM
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))
print('注意：AM人员变更，需同步修改表单结构；')

# NP
cursor.execute('''
select a.AM, sum([Q2 Total Spending Forecast]) as Q2_Total_Spending_Forecast, 
sum([Apr_FX GAIN(RMB)]) as Apr_FX_Gain_RMB, sum([May_FX GAIN(RMB)]) as May_FX_Gain_RMB, sum([Jun_FX GAIN(RMB)]) as Jun_FX_Gain_RMB,
sum([FX GAIN(RMB)]) as FX_Gain_RMB, sum([FX GAIN(RMB)])/0.18 as FX_Gain_Spending_RMB, 
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD,
sum([QTD FX GAIN (RMB)]) as QTD_FX_Gain_RMB, sum([QTD FX GAIN (RMB)])/0.18 as QTD_FX_Gain_Spending_RMB
from NP_20190402 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Group, a.AM
order by b.AM_Group, a.AM
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A21'].value = list(map(list, item))

# Infeeds
cursor.execute('''
select a.AM, sum([Q2 Total Spending Forecast]) as Q2_Total_Spending_Forecast, 
sum([Apr_FX GAIN(RMB)]) as Apr_FX_Gain_RMB, sum([May_FX GAIN(RMB)]) as May_FX_Gain_RMB, sum([Jun_FX GAIN(RMB)]) as Jun_FX_Gain_RMB,
sum([FX GAIN(RMB)]) as FX_Gain_RMB, sum([FX GAIN(RMB)])/0.18 as FX_Gain_Spending_RMB, 
sum([2019Apr]+[2019May]+[2019Jun]) as Q2_Total_Spending_QTD,
sum([QTD FX GAIN (RMB)]) as QTD_FX_Gain_RMB, sum([QTD FX GAIN (RMB)])/0.18 as QTD_FX_Gain_Spending_RMB
from Infeeds_20190402 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Group, a.AM
order by b.AM_Group, a.AM
               '''.replace('20190402', DATE))
item = cursor.fetchall()
sht['A39'].value = list(map(list, item))

# Sales
# P4P
sht = wb.sheets['Sales']
sql = '''
select a.NB, a.销售, sum([Q2 Eligible Spending Forecast]) as Q2_Eligible_Spending_Forecast, sum([FX GAIN(RMB)_Sales]) as FX_Gain_RMB_Sales,
sum([FX GAIN(RMB)_Sales])/0.18 as FX_Gain_Spending_RMB_Sales, sum([Eligible Spending(QTD)]) as Q2_Eligible_Spending_QTD, 
sum([QTD FX GAIN (RMB)_Sales]) as QTD_FX_Gain_RMB_Sales, sum([QTD FX GAIN (RMB)_Sales])/0.18 as QTD_FX_Gain_Spending_RMB_Sales
from P4P_20190402 a inner join [CSA_HK_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%' and NB <> '2017&2018EB'
group by a.NB, a.销售
order by a.NB, a.销售
       '''
cursor.execute(sql.replace('20190402', DATE))
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))

# NP
cursor.execute(sql.replace('P4P_20190402', 'NP_'+DATE))
item = cursor.fetchall()
sht['A17'].value = list(map(list, item))
# Infeeds
cursor.execute(sql.replace('P4P_20190402', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A31'].value = list(map(list, item))

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File: Handling Fee Details
# p4p
wb = xw.Book(target+'\\'+os.listdir(target)[4])
sht = wb.sheets['P4P']
sql = '''
select * from P4P_20190402 where [FX GAIN(RMB)] > 0 and 端口 not like '%wrong%'
order by [FX GAIN(RMB)] desc
    '''
# P4P
cursor.execute(sql.replace('20190402', DATE))
item = cursor.fetchall()
sht['A2'].value = list(map(list, item))

# NP
cursor.execute(sql.replace('P4P_20190402', 'NP_'+DATE))
item = cursor.fetchall()
sht = wb.sheets['NP']
sht['A2'].value = list(map(list, item))

# Infeeds
cursor.execute(sql.replace('P4P_20190402', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht = wb.sheets['Infeeds']
sht['A2'].value = list(map(list, item))

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File:Sales Forecast
wb = xw.Book(target+'\\'+os.listdir(target)[6])
sht = wb.sheets['Data']
rows100 = sht['A1'].current_region.rows.count
sql = '''
select a.销售, [NB Month], sum([Q2 Eligible Spending Forecast]) as Q2_Eligible_Spending_Forecast, sum([Q2 Eligible GP Forecast]) as Q2_Eligible_GP_Forecast
from P4P_20190402 a inner join [CSA_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%'
group by a.销售, [NB Month]
order by a.销售, [NB Month]
    '''
# P4P
cursor.execute(sql.replace('20190402', DATE))
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))
# NP
cursor.execute(sql.replace('P4P_20190402', 'NP_'+DATE))
item = cursor.fetchall()
sht['E3'].value = list(map(list, item))
# Infeeds
cursor.execute(sql.replace('P4P_20190402', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['I3'].value = list(map(list, item))
# All
sht['M3'].value = sht['I3:J'+str(len(item)+2)].value
rows101 = sht['A1'].current_region.rows.count
if rows101 > rows100:
    rng100 = sht['O'+str(rows100)+':P'+str(rows100)]
    rng101 = sht['O'+str(rows100)+':P'+str(rows101)]
    rng100.api.AutoFill(rng101.api, constants.AutoFillType.xlFillCopy)

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# GP Ratio
# 1.修改：2019Q1
wb = xw.Book(target+'\\'+os.listdir(target)[7])
sht = wb.sheets['Sheet1']
sql = '''
select a.销售, sum([Q2 Eligible Spending Forecast]) as Q2_Eligible_Spending_Forecast, 
sum([Q2 Eligible GP Forecast]) as Q2_Eligible_GP_Forecast
from P4P_20190402 a inner join [CSA_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%' and a.NB='2019Q2'
group by a.销售
order by a.销售
    '''
# P4P
cursor.execute(sql.replace('20190402', DATE))
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))
# NP
cursor.execute(sql.replace('P4P_20190402', 'NP_'+DATE))
item = cursor.fetchall()
sht['A13'].value = list(map(list, item))
# Infeeds
cursor.execute(sql.replace('P4P_20190402', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A23'].value = list(map(list, item))
# 汇总
sht['E3'].options(transpose=True).value = [i[0] for i in list(map(list, item))]

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File: Sales Tracking Report
# 销售
# 1.修改where
wb = xw.Book(target+'\\'+os.listdir(target)[8])
sht = wb.sheets[0]
sql = '''
select a.NB, a.销售, sum([Apr Eligible Spending]) as Apr_Eligible_Spending, sum([May Eligible Spending]) as May_Eligible_Spending, 
sum([Jun Eligible Spending]) as Jun_Eligible_Spending, sum([Apr Eligible Spending Forecast]) as Apr_Eligible_Spending_Forecast, 
sum([May Eligible Spending Forecast]) as May_Eligible_Spending_Forecast, sum([Jun Eligible Spending Forecast]) as Jun_Eligible_Spending_Forecast, 
sum([Q2 Eligible Spending Forecast]) as Q2_Eligible_Spending_Forecast, sum([Eligible Spending(QTD)]) as Q2_Eligible_Spending_QTD
from P4P_20190402 a inner join [CSA_HK_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%' and a.NB not like '2017&2018EB' and b.Sales not like '%Eric%'
group by a.NB, a.销售
order by a.NB, a.销售
    '''
# P4P
cursor.execute(sql.replace('20190402', DATE))
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))
# NP
cursor.execute(sql.replace('P4P_20190402', 'NP_'+DATE))
item = cursor.fetchall()
sht['A14'].value = list(map(list, item))
# Infeeds
cursor.execute(sql.replace('P4P_20190402', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A25'].value = list(map(list, item))

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File: AM Tracking Report
# AM
wb = xw.Book(target+'\\'+os.listdir(target)[2])
sht = wb.sheets[0]
sql = '''
select b.AM_Region, b.AM_Group,b.AM, sum([Apr Spending Forecast]) as Apr_19,
sum([May Spending Forecast]) as May_19,
sum([Jun Spending Forecast]) as Jun_19,
sum([Total Spending]) as Q2_Total_Spending_QTD
from P4P_20190402 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Region, b.AM_Group,b.AM
order by b.AM_Region, b.AM_Group,b.AM
    '''
# P4P
cursor.execute(sql.replace('20190402', DATE))
item = cursor.fetchall()
sht['A3'].value = list(map(list, item))
# NP
cursor.execute(sql.replace('P4P_20190402', 'NP_'+DATE))
item = cursor.fetchall()
sht['A19'].value = list(map(list, item))
# Infeeds
# n1 = 0.03227055633443914
# n2 = 0.0789313904068002
sql = '''
select b.AM_Region, b.AM_Group,b.AM, sum([Apr Spending Forecast])+sum([Apr Spending Forecast])*0.0789313904068002/0.18 as Apr_19,
sum([May Spending Forecast])+sum([May Spending Forecast])*0.0789313904068002/0.18 as May_19,
sum([Jun Spending Forecast])+sum([Jun Spending Forecast])*0.0789313904068002/0.18 as Jun_19,
sum([Total Spending])+sum([Total Spending])*0.0789313904068002/0.18 as Q2_Total_Spending_QTD
from Infeeds_20190402 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Region, b.AM_Group,b.AM
order by b.AM_Region, b.AM_Group,b.AM
    '''
cursor.execute(sql.replace('20190402', DATE))
item = cursor.fetchall()
sht['A35'].value = list(map(list, item))


sht['I3'].value = sht['A3:C'+str(len(item)+2)].value

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# drop P4P & NP & Infeeds
sql = '''
DROP TABLE P4P_20190402
DROP TABLE NP_20190402
DROP TABLE Infeeds_20190402
'''
cursor.execute(sql.replace('20190402', DATE))
cnxn.commit()


print('\a耗时：%s' % str((time.clock()-star)/60))