# -*- coding: utf-8 -*-
"""
Created on Thu Jan  3 19:12:56 2019

@author: chen.huaiyu
"""

''' 手动导入数据表  '''

import pyodbc
import os
import time
import xlwings as xw
from xlwings import constants
from datetime import date, timedelta

star = time.clock()

def dat():
	n = input('Default Yesterday?：')
	if n == '':
		n = '1'
	return (date.today() - timedelta(eval(n))).strftime('%Y%m%d')  # 改 

DATE = dat()
version = input("Version: ( v1? )")
print('DATE = {}'.format(DATE))
target = r'C:\Users\chen.huaiyu\Desktop\Output\SQL Server' + '\\' + DATE + ' ' + version
server_name = 'SZ-CS-0038LT\SQLEXPRESS'
db_name = 'Output'
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
                if object_id('P4P_20191002') is not null
                    drop table P4P_20191002
                select a.*, b.* into P4P_20191002 from P4P_20191002_1 a inner join P4P_20191002_2 b on a.用户名=b.用户名1
                alter table P4P_20191002 drop column 用户名1
                
                alter table P4P_20191002 add Forex varchar
                update P4P_20191002
                set Forex = b.Forex
                from P4P_20191002 a inner join Forex1 b on a.用户名=b.用户名1
                
                
                
                if object_id('NP_20191002') is not null
                    drop table NP_20191002
                select a.*, b.* into NP_20191002 from NP_20191002_1 a inner join NP_20191002_2 b on a.用户名=b.用户名1
                alter table NP_20191002 drop column 用户名1
                
                alter table NP_20191002 add Forex varchar
                update NP_20191002
                set Forex = b.Forex
                from NP_20191002 a inner join Forex1 b on a.用户名=b.用户名1
                
                
                
                if object_id('Infeeds_20191002') is not null
                    drop table Infeeds_20191002
                select a.*, b.* into Infeeds_20191002 from Infeeds_20191002_1 a inner join Infeeds_20191002_2 b on a.用户名=b.用户名1
                alter table Infeeds_20191002 drop column 用户名1
                
                alter table Infeeds_20191002 add Forex varchar
                update Infeeds_20191002
                set Forex = b.Forex
                from Infeeds_20191002 a inner join Forex1 b on a.用户名=b.用户名1
               '''.replace('20191002', DATE))
cnxn.commit()

# 修改 
# 1. 3月AVG都只需要 直接把cny-jrjr-2 地區改去 SZ
# 2. AM 轉黄希腾
sql = ''' update {}
    set 区域='SZ', AM='黄希腾', Industry='金融服务'
    where 用户名 = 'cny-jrjr-2'
    '''
for i in ['P4P_', 'NP_', 'Infeeds_']:
    i += DATE
    cursor.execute(sql.format(i))
cnxn.commit()


# try:
# Excel File:Spending Forecast
# Pivot Table(端口)
# P4P

print('注1：\nSpending Forecast')
wb = xw.Book(target+'\\'+os.listdir(target)[0])
sql = '''
select 端口, 
	sum([Jan Spending Forecast]) as Jan_20,
	sum([Feb Spending Forecast]) as Feb_20,
	sum([Mar Spending Forecast]) as Mar_20,
	sum([Apr Spending Forecast]) as Apr_20,
	sum([May Spending Forecast]) as May_20,
	sum([Jun Spending Forecast]) as Jun_20,
	sum([2020Jul]) as Jul_20,
	sum([2020Aug]) as Aug_20,
	sum([2020Sep]) as Sep_20,
	sum([2020Oct]) as Oct_20,
	sum([2020Nov]) as Nov_20,
	sum([2020Dec]) as Dec_20
from P4P_20191002
where 端口 not like '%wrong%'
group by 端口
order by 端口
'''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
print(len(item))
sht = wb.sheets['P4P']
sht[1, 0].value = list(map(lambda x:list(x), item))
print('Spending Forecast P4P Port Count: %s' % sht[1, 0].current_region.rows.count)

# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht = wb.sheets['NP']
sht[1, 0].value = list(map(lambda x:list(x), item))
print('Spending Forecast NP Port Count: %s' % sht[1, 0].current_region.rows.count)

# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
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
    
# =============================================================================
# # Infeeds(27%)
# sht1 = wb.sheets['Infeeds(27%)']
# rows100 = sht1['A1'].current_region.rows.count
# sht1['A2'].options(transpose=True).value = sht[1:rows101, 0].value
# if rows100 < rows101:
#     range100 = sht1['B'+str(rows100)+':M'+str(rows100)]
#     range101 = sht1['B'+str(rows100)+':M'+str(rows101)]
#     range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)
# 
# =============================================================================
    
# Pivot Table (用户名)
# P4P
sql = '''
select 用户名, 
	sum([Jan Spending Forecast]) as Jan_20,
	sum([Feb Spending Forecast]) as Feb_20,
	sum([Mar Spending Forecast]) as Mar_20,
	sum([Apr Spending Forecast]) as Apr_20,
	sum([May Spending Forecast]) as May_20,
	sum([Jun Spending Forecast]) as Jun_20,
	sum([2020Jul]) as Jul_20,
	sum([2020Aug]) as Aug_20,
	sum([2020Sep]) as Sep_20,
	sum([2020Oct]) as Oct_20,
	sum([2020Nov]) as Nov_20,
	sum([2020Dec]) as Dec_20
from P4P_20191002
where 端口 not like '%wrong%'
group by 用户名
order by 用户名
'''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
sht = wb.sheets['P4P']
sht['O2'].value = list(map(lambda x:list(x), item))
print('Spending Forecast P4P User Count: %s' % sht['O1'].current_region.rows.count)

# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht = wb.sheets['NP']
sht['O2'].value = list(map(lambda x:list(x), item))
print('Spending Forecast NP User Count: %s' % sht['O1'].current_region.rows.count)


# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
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
# =============================================================================
# # Infeeds(27%)
# sht1 = wb.sheets['Infeeds(27%)']
# rows100 = sht1['O1'].current_region.rows.count
# sht1['O2'].options(transpose=True).value = sht['O2:O'+str(rows101)].value
# if rows101 > rows100:
#     range100 = sht1['P'+str(rows100)+':AA'+str(rows100)]
#     range101 = sht1['P'+str(rows100)+':AA'+str(rows101)]
#     range100.api.AutoFill(range101.api, constants.AutoFillType.xlFillCopy)
# =============================================================================


# Pivot Table (区域)
# P4P
sql = '''
select 区域, sum([Jan Spending Forecast]) as Jan_20,
	sum([Feb Spending Forecast]) as Feb_20,
	sum([Mar Spending Forecast]) as Mar_20,
	sum([Apr Spending Forecast]) as Apr_20,
	sum([May Spending Forecast]) as May_20,
	sum([Jun Spending Forecast]) as Jun_20,
	sum([2020Jul]) as Jul_20,
	sum([2020Aug]) as Aug_20,
	sum([2020Sep]) as Sep_20,
	sum([2020Oct]) as Oct_20,
	sum([2020Nov]) as Nov_20,
	sum([2020Dec]) as Dec_20
from P4P_20191002
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
'''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
sht = wb.sheets['Region']
sht['A3'].value = list(map(lambda x:list(x), item))
# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A12'].value = list(map(lambda x:list(x), item))
# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
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
sql = '''
select AM, 
    端口, 
    用户名, 
    广告主, 
	sum([Ave# Daily Workday]) as avg_daily_workday,
    sum([Ave# Daily Holiday]) as avg_daily_holiday,
	sum([Jan Spending Forecast]) as Jan_20,
	sum([Feb Spending Forecast]) as Feb_20,
	sum([Mar Spending Forecast]) as Mar_20,
	sum([Apr Spending Forecast]) as Apr_20,
	sum([May Spending Forecast]) as May_20,
	sum([Jun Spending Forecast]) as Jun_20,
	sum([2020Jul]) as Jul_20,
	sum([2020Aug]) as Aug_20,
	sum([2020Sep]) as Sep_20,
	sum([2020Oct]) as Oct_20,
	sum([2020Nov]) as Nov_20,
	sum([2020Dec]) as Dec_20
from P4P_20191002
where 端口 not like '%wrong%'
group by AM, 端口, 用户名, 广告主
order by AM, 端口, 用户名, 广告主
'''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
sht['A2'].value = list(map(list, item))
print('sht P4P rows counts:%s' %sht['A1'].current_region.rows.count)

# NP
sht = wb.sheets['NP']
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A2'].value = list(map(list, item))
print('sht NP rows counts:%s' %sht['A1'].current_region.rows.count)

# Infeeds
sht = wb.sheets['Infeeds']
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
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
select 区域, 
	sum([Jan Spending Forecast]) as Jan_20,
	sum([Feb Spending Forecast]) as Feb_20,
	sum([Mar Spending Forecast]) as Mar_20,
    sum([2020Jan]+[2020Feb]+[2020Mar]) as Q1_Total_Spending_QTD
from P4P_20191002
where 端口 not like '%wrong%' and 区域 <> '-'
group by 区域
order by 区域
   '''
   
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
Ex_row = sht['A1'].current_region.rows.count - 3
print('SQL Server %s 行，Excel %s 行' % (len(item), Ex_row))
sht['A3'].value = list(map(list, item))

# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))             
item = cursor.fetchall()
sht['A13'].value = list(map(list, item))

# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A23'].value = list(map(list, item))

# Finance Region
sht = wb.sheets['Finance Region']
# P4P
sql = '''
select 财务做账区域, 
	区域, 
	sum([Jan Spending Forecast]) as Jan_20,
	sum([Feb Spending Forecast]) as Feb_20,
	sum([Mar Spending Forecast]) as Mar_20,
	sum([2020Jan]+[2020Feb]+[2020Mar]) as Q1_Total_Spending_QTD
from P4P_20191002
where 端口 not like '%wrong%' and 区域 <> '-'
group by 财务做账区域, 区域
order by 财务做账区域, 区域
'''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
Ex_row = sht['A1'].current_region.rows.count - 3
print('SQL Server %s 行，Excel %s 行' % (len(item), Ex_row))
sht['A3'].value = list(map(list, item))

# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A19'].value = list(map(list, item))

# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A35'].value = list(map(list, item))

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File:Handling Fee
# AM Handling Fee
wb = xw.Book(target+'\\'+os.listdir(target)[5])
sht = wb.sheets['AM']
# P4P
sql = '''
select a.AM, 
	sum([Q1 Total Spending Forecast]) as Q1_Total_Spending_Forecast, 
	sum([Jan_FX GAIN(RMB)]) as Jan_FX_Gain_RMB, 
	sum([Feb_FX GAIN(RMB)]) as Feb_FX_Gain_RMB, 
	sum([Mar_FX GAIN(RMB)]) as Mar_FX_Gain_RMB,
	sum([FX GAIN(RMB)]) as FX_Gain_RMB, 
	sum([FX GAIN(RMB)])/0.18 as FX_Gain_Spending_RMB, 
	sum([2020Jan]+[2020Feb]+[2020Mar]) as Q1_Total_Spending_QTD,
	sum([QTD FX GAIN (RMB)]) as QTD_FX_Gain_RMB, 
	sum([QTD FX GAIN (RMB)])/0.18 as QTD_FX_Gain_Spending_RMB
from P4P_20191002 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Group, a.AM
order by b.AM_Group, a.AM
'''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
len1 = len(item)
print(len1)
r = sht['A1'].current_region.rows.count
if len1 > r-4:
    print('Am增加，请插入%s行' % str(len1-(r-4)))
    sht[r-2, :].api.insert
    sht[r-2, :].offset(r+1).api.insert
    sht[r-2, :].offset(2*r+3).api.insert
    sht[r-2, :].offset(3*r+5).api.insert
    input('Tips:检查Excel，数据定位')
elif len1 < r-4:
    print('AM减少,请删除%s行' % str(r-4-len1))
else:
    print('AM行列正常')

sht['A3'].value = list(map(list, item))
print('注意：AM人员变更，需同步修改表单结构；')

# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A20'].value = list(map(list, item))

# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A37'].value = list(map(list, item))

# Sales
# P4P
sht = wb.sheets['Sales']
sql = '''
select a.NB, 
	a.销售, 
	sum([Q1 Eligible Spending Forecast]) as Q1_Eligible_Spending_Forecast, 
	sum([FX GAIN(RMB)_Sales]) as FX_Gain_RMB_Sales,
	sum([FX GAIN(RMB)_Sales])/0.18 as FX_Gain_Spending_RMB_Sales, 
	sum([Eligible Spending(QTD)]) as Q1_Eligible_Spending_QTD, 
	sum([QTD FX GAIN (RMB)_Sales]) as QTD_FX_Gain_RMB_Sales, 
	sum([QTD FX GAIN (RMB)_Sales])/0.18 as QTD_FX_Gain_Spending_RMB_Sales
from P4P_20191002 a inner join [CSA_HK_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%' 
	and NB <> '2017&2018EB'
group by a.NB, a.销售
order by a.NB, a.销售
       '''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
print('SQL Server数量：', len(item))
print('表格', sht['A1'].current_region.rows.count - 5)
sht['A3'].value = list(map(list, item))

# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A26'].value = list(map(list, item))
# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A49'].value = list(map(list, item))

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File: Handling Fee Details
# p4p
wb = xw.Book(target+'\\'+os.listdir(target)[4])
sql_exist = '''
    IF (EXISTS( SELECT * 
               FROM sysobjects 
               WHERE name='test_temp'))
        DROP TABLE test_temp
    '''
sql_test_temp = '''
select * 
    into test_temp
from P4P_20191002 
where [FX GAIN(RMB)] > 0 
	and 端口 not like '%wrong%'
order by [FX GAIN(RMB)] desc
    '''
for sheet in ['P4P', 'NP', 'Infeeds']:
    sht = wb.sheets[sheet]
    time.sleep(5)
    sql0 = sql_test_temp.replace('20191002', DATE)
    sql0 = sql0.replace('P4P', sheet)
    cursor.execute(sql_exist)
    cursor.execute(sql0)
    sql1 = '''
        SELECT * FROM test_temp
        '''
    item = cursor.execute(sql1).fetchall()
    sql_ = '''
        SELECT * FROM information_schema.columns
        WHERE table_name='test_temp'
        '''
    title = [i[3] for i in cursor.execute(sql_).fetchall()]
    ex_row = sht['A1'].current_region.rows.count
    print('写入数据：%s行；原表中:%s行' % (len(item), ex_row))
    sht[:, :].clear()
    sht['A1'].value = title
    data = list(map(list, item))
    
    # test
    import pandas as pd
    df = pd.DataFrame(data, columns=title)
    
    
    sht['A2'].options(expend='table').value = df.values
wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File:Sales Forecast
wb = xw.Book(target+'\\'+os.listdir(target)[6])
sht = wb.sheets['Data']
rows100 = sht['A1'].current_region.rows.count
sht['A3:L'+str(rows100)].clear()
sql_exist = '''
    IF EXISTS(SELECT * FROM sysobjects WHERE name='test_temp')
        DROP TABLE test_temp
    '''
sql_test_temp = '''
select a.销售
	, [NB Month]
	, sum([Q1 Eligible Spending Forecast]) as Q1_Eligible_Spending_Forecast
	, sum([Q1 Eligible GP Forecast]) as Q1_Eligible_GP_Forecast
  into test_temp
from P4P_20191002 a inner join [CSA_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%'
group by a.销售, [NB Month]
order by a.销售, [NB Month]
    '''
sql_temp = '''
    SELECT * FROM test_temp
    '''
for sheet in ['P4P', 'NP', 'Infeeds']:
    cursor.execute(sql_exist)
    sql0 = sql_test_temp.replace('20191002', DATE).replace('P4P', sheet)
    cursor.execute(sql0)
    title = [i[3] for i in cursor.execute(
                    '''SELECT * 
                    FROM information_schema.columns 
                    WHERE table_name='test_temp'
                    ''')]
    item = cursor.execute(sql_temp).fetchall()
    if sheet == 'P4P':
        sht['A3'].value = title
        sht['A4'].value = list(map(list, item))
    elif sheet == 'NP':
        sht['E3'].value = title
        sht['E4'].value = list(map(list, item))
    else:
        sht['I3'].value = title
        sht['I4'].value = list(map(list, item))
        
# All
# sht['O3'].value = sht['I3:J'+str(len(item)+2)].value
rows101 = sht['A1'].current_region.rows.count
if rows101 > rows100:
    rng100 = sht['M'+str(rows100)+':R'+str(rows100)]
    rng101 = sht['M'+str(rows100)+':R'+str(rows101)]
    rng100.api.AutoFill(rng101.api, constants.AutoFillType.xlFillCopy)

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# GP Ratio
# 1.修改：2019Q1
wb = xw.Book(target+'\\'+os.listdir(target)[7])
sht = wb.sheets['Sheet1']
sql = '''
select a.销售
	, sum([Q1 Eligible Spending Forecast]) as Q1_Eligible_Spending_Forecast
	, sum([Q1 Eligible GP Forecast]) as Q1_Eligible_GP_Forecast
from P4P_20191002 a inner join [CSA_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%' and a.NB='2019Q4'
group by a.销售
order by a.销售
    '''
# P4P
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
Cnt = len(item)
row = sht['A1'].current_region.rows.count - 2
print('SQL Server 行数：%s, Excel行数：%s' % (Cnt, row))
sht['A3:C'+str(row)].clear()
sht['A3'].value = list(map(list, item))
# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A14'].value = list(map(list, item))
# Infeeds
cursor.execute(sql.replace('P4P_20191002', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A25'].value = list(map(list, item))
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
select a.NB
	, a.销售
	, sum([Jan Eligible Spending]) as Jan_Eligible_Spending
	, sum([Feb Eligible Spending]) as Feb_Eligible_Spending
	, sum([Mar Eligible Spending]) as Mar_Eligible_Spending
	, sum([Jan Eligible Spending Forecast]) as Jan_Eligible_Spending_Forecast
	, sum([Feb Eligible Spending Forecast]) as Feb_Eligible_Spending_Forecast
	, sum([Mar Eligible Spending Forecast]) as Mar_Eligible_Spending_Forecast
    , sum([Q1 Eligible Spending Forecast]) as Q1_Eligible_Spending_Forecast
    , sum([Eligible Spending(QTD)]) as Q1_Eligible_Spending_QTD
from P4P_20191002 a 
    inner join [CSA_HK_Sales] b 
        on a.销售=b.Sales
where 端口 not like '%wrong%' 
    and a.NB not like '2017&2018EB'
group by a.NB, a.销售
order by a.NB, a.销售
    '''
# P4P
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
ex_row = sht['A3'].current_region.rows.count
print('SQL Server输出:%s, Excel行数:%s' % (len(item), ex_row))
sht['A3'].value = list(map(list, item))
# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A28'].value = list(map(list, item))
# Infeeds
# sql
sql = '''
select a.NB
	, a.销售
	, sum([Jan Eligible Spending]+[Jan Eligible Spending]*0.0789313904068002/0.18) as Jan_Eligible_Spending
	, sum([Feb Eligible Spending]+[Feb Eligible Spending]*0.0789313904068002/0.18) as Feb_Eligible_Spending
	, sum([Mar Eligible Spending]+[Mar Eligible Spending]*0.0789313904068002/0.18) as Mar_Eligible_Spending
	, sum([Jan Eligible Spending Forecast]+[Jan Eligible Spending Forecast]*0.0789313904068002/0.18) as Jan_Eligible_Spending_Forecast
	, sum([Feb Eligible Spending Forecast]+[Feb Eligible Spending Forecast]*0.0789313904068002/0.18) as Feb_Eligible_Spending_Forecast
	, sum([Mar Eligible Spending Forecast]+[Mar Eligible Spending Forecast]*0.0789313904068002/0.18) as Mar_Eligible_Spending_Forecast
	, sum([Q1 Eligible Spending Forecast]+[Q1 Eligible Spending Forecast]*0.0789313904068002/0.18) as Q1_Eligible_Spending_Forecast
	, sum([Eligible Spending(QTD)]+[Eligible Spending(QTD)]*0.0789313904068002/0.18) as Q1_Eligible_Spending_QTD
from Infeeds_20191002 a inner join [CSA_HK_Sales] b on a.销售=b.Sales
where 端口 not like '%wrong%' and a.NB not like '2017&2018EB'
group by a.NB, a.销售
order by a.NB, a.销售
'''
cursor.execute(sql.replace('Infeeds_20191002', 'Infeeds_'+DATE))
item = cursor.fetchall()
sht['A53'].value = list(map(list, item))

wb.app.calculation = 'automatic'
wb.save()
wb.close()


# Excel File: AM Tracking Report
# AM
wb = xw.Book(target+'\\'+os.listdir(target)[2])
sht = wb.sheets[0]
sql = '''
select b.AM_Region
	, b.AM_Group
	, b.AM
	, sum([Jan Spending Forecast]) as Jan_19
	, sum([Feb Spending Forecast]) as Feb_19
	, sum([Mar Spending Forecast]) as Mar_19
	, sum([Total Spending]) as Q1_Total_Spending_QTD
from P4P_20191002 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Region, b.AM_Group,b.AM
order by b.AM_Region, b.AM_Group,b.AM
    '''
# P4P
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
Ex_row = sht['A1'].current_region.rows.count - 2
print('SQL Server输出 %s 行,Excel原有 %s 行' % (len(item), Ex_row))
sht['A3'].value = list(map(list, item))
# NP
cursor.execute(sql.replace('P4P_20191002', 'NP_'+DATE))
item = cursor.fetchall()
sht['A18'].value = list(map(list, item))
# Infeeds
# n1 = 0.03227055633443914
# n2 = 0.0789313904068002
sql = '''
select b.AM_Region
	, b.AM_Group
	,b.AM
	, sum([Jan Spending Forecast])+sum([Jan Spending Forecast])*0.0789313904068002/0.18 as Jan_19
	, sum([Feb Spending Forecast])+sum([Feb Spending Forecast])*0.0789313904068002/0.18 as Feb_19
	, sum([Mar Spending Forecast])+sum([Mar Spending Forecast])*0.0789313904068002/0.18 as Mar_19
	, sum([Total Spending])+sum([Total Spending])*0.0789313904068002/0.18 as Q1_Total_Spending_QTD
from Infeeds_20191002 a inner join [CSA_AM] b on a.AM=b.AM
where 端口 not like '%wrong%'
group by b.AM_Region, b.AM_Group,b.AM
order by b.AM_Region, b.AM_Group,b.AM
    '''
cursor.execute(sql.replace('20191002', DATE))
item = cursor.fetchall()
sht['A33'].value = list(map(list, item))


sht['I3'].value = sht['A3:C'+str(len(item)+2)].value

wb.app.calculation = 'automatic'
wb.save()
wb.close()

# =============================================================================
# 
# # drop P4P & NP & Infeeds
# sql = '''
# DROP TABLE P4P_20191002
# DROP TABLE NP_20191002
# DROP TABLE Infeeds_20191002
# '''
# cursor.execute(sql.replace('20191002', DATE))
# cnxn.commit()
# 
# =============================================================================

print('\a耗时：%s' % str((time.clock()-star)/60))