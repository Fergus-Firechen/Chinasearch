# -*- coding: utf-8 -*-
"""
Created on Mon Dec 24 10:19:36 2018

@author: chen.huaiyu
"""

'''   分割：准备  '''

# 批量修改文件名
import os
import shutil
import xlwings as xw

date_file = '20191008'
DATE = '20191015'  # 改1

# eg.规范命命：ave.workday&weekdayq4(2018 oct_dec)2019.01.02_v3
# wb2 = xw.Book(r'C:\Users\chen.huaiyu\Downloads\Ave.workday&weekdayQ4(2018 Oct_Dec) ' + DATE_ + 'v1.xlsx')


# 1. 新建文件夹
SOURCE = r'C:\Users\chen.huaiyu\Desktop\Output\SQL Server'+'\\'+date_file  # 改2
target = r'C:\Users\chen.huaiyu\Desktop\Output\SQL Server' + '\\' + DATE
if os.path.exists(target) == False:
    os.makedirs(target)
    # 2. 复制文件
    #2.1 获取文件目录
    fileList = os.listdir(SOURCE)
    for i in fileList:
        try:
            shutil.copy(SOURCE + '\\' + i, target + '\\' + i)
        except IOError as e:
            print('Unable to copy file. %s' % e)
        except:
            print('Unexcepted error')
    
    # 3. 修改文件名
    for i in os.listdir(target):
        oldName = target + '\\' + i  # 绝对地址！
        newName = target + '\\' + i.replace(date_file, DATE)  # 改3
        os.rename(oldName, newName)

# 分割文件
# 准备
Source_input = r'C:\Users\chen.huaiyu\Desktop\Input\SQL server' + '\\'+date_file
Target_input = r'C:\Users\chen.huaiyu\Desktop\Input\SQL server' + '\\' + DATE
if os.path.exists(Target_input) == False:  
    os.makedirs(Target_input)
    # os.remove(Target_input)
    fileInputList = os.listdir(Source_input)
    for i in fileInputList:
        shutil.copy(Source_input + '\\' + i, Target_input + '\\' + i)
    for n,i in enumerate(os.listdir(Target_input)):
        oldName = Target_input + '\\' + i
        newName = Target_input + '\\' + i.replace(date_file, DATE)
        os.rename(oldName, newName)
# =============================================================================
# # 修改 Q3 -> Q4
# for i in os.listdir(target):
#     print(i)
#     oldName = target + '\\' + i
#     newName = target + '\\' + i.replace('Q3', 'Q4')
#     os.rename(oldName, newName)
# =============================================================================


''' 处理  
1.观察AVG & 拆分模版区别（特别是元数据；字段）
'''


# 数据处理（2019.01.16
DATE_ = DATE[:4] + '.' + DATE[4:6] + '.' + DATE[6:]
wb2 = xw.Book(r'C:\Users\chen.huaiyu\Downloads\Ave.workday&weekdayQ4(2019 OCT to Nov )' + DATE_ + '.xlsx')  # 改4
sht2 = wb2.sheets['搜索']
sht3 = wb2.sheets['其他新产品']
sht4 = wb2.sheets['原生广告']
# 行 & 列数
rows_2 = sht2[0, 0].current_region.rows.count
columns_2 = sht2[0, 0].current_region.columns.count

wb3 = xw.Book(r'C:\Users\chen.huaiyu\Desktop\Input\SQL server' + '\\' + DATE + '\\P4P_' + DATE + '.xlsx')
sht32 = wb3.sheets['P4P1']
sht33 = wb3.sheets['P4P2']

wb4 = xw.Book(r'C:\Users\chen.huaiyu\Desktop\Input\SQL server' + '\\' + DATE + '\\NP_' + DATE + '.xlsx')
sht42 = wb4.sheets['NP1']
sht43 = wb4.sheets['NP2']

wb5 = xw.Book(r'C:\Users\chen.huaiyu\Desktop\Input\SQL server' + '\\' + DATE + '\\Infeeds_' + DATE + '.xlsx')
sht52 = wb5.sheets['Infeeds1']
sht53 = wb5.sheets['Infeeds2']


# =============================================================================
# def CopyExcel(sht):
#     sht31['A2:A' + str(rows_2)].color = (162, 163, 165)
#     sht31['JF2:JF' + str(rows_2)].color = (162, 163, 165)
#     sht31['A2'].value = sht['A3:JE' + str(rows_2)].value
#     sht31['JF2'].value = sht['JF3:OL' + str(rows_2)].value
#     sht32[1, 0].value = sht['A3:JE' + str(rows_2)].value
#     sht33[1, 0].value = sht['J3:J' + str(rows_2)].value
#     sht33[1, 1].value = sht['JF3:OL' + str(rows_2)].value
# =============================================================================


j1 = 'A'  # 搜索 & P4P
j2 = 'BD'  # BE
j3 = 'BO'
j4 = 'IW'  # 
j5 = 'IX'  # 2
j6 = 'PE'  # 

k1 = 'BE'  # P4P1

## P4P
# 分拆为四份 A:AU
i, l1, l2 = sht2, sht32, sht33

l1['A2'].value = i[j1+'2:'+j2 + str(rows_2//4)].value
l1['A'+str(rows_2//4+1)].value = i[j1+str(rows_2//4+1)+':'+j2+str(rows_2//2)].value
l1['A'+str(rows_2//2+1)].value = i[j1+str(rows_2//2+1)+':'+j2+str(3*rows_2//4)].value
l1['A'+str(3*rows_2//4+1)].value = i[j1+str(3*rows_2//4+1)+':'+j2+str(rows_2)].value


l1[k1+'2'].value = i[j3+'2:'+j4 + str(rows_2//4)].value
l1[k1+str(rows_2//4+1)].value = i[j3+str(rows_2//4+1)+':'+j4+str(rows_2//2)].value
l1[k1+str(rows_2//2+1)].value = i[j3+str(rows_2//2+1)+':'+j4+str(3*rows_2//4)].value
l1[k1+str(3*rows_2//4+1)].value = i[j3+str(3*rows_2//4+1)+':'+j4+str(rows_2)].value

# 用户名1
l2['A2'].options(transpose=True).value = i['J2:J'+str(rows_2)].value
# 数据
l2['B2'].value = i[j5+'2:'+j6 + str(rows_2//4)].value
l2['B'+str(rows_2//4+1)].value = i[j5+str(rows_2//4+1)+':'+j6+str(rows_2//2)].value
l2['B'+str(rows_2//2+1)].value = i[j5+str(rows_2//2+1)+':'+j6+str(3*rows_2//4)].value
l2['B'+str(3*rows_2//4+1)].value = i[j5+str(3*rows_2//4+1)+':'+j6+str(rows_2)].value

'''  查 '''

## NP
# 分拆为四份 A:AU
i, l1, l2 = sht3, sht42, sht43

l1['A2'].value = i[j1+'2:'+j2 + str(rows_2//4)].value
l1['A'+str(rows_2//4+1)].value = i[j1+str(rows_2//4+1)+':'+j2+str(rows_2//2)].value
l1['A'+str(rows_2//2+1)].value = i[j1+str(rows_2//2+1)+':'+j2+str(3*rows_2//4)].value
l1['A'+str(3*rows_2//4+1)].value = i[j1+str(3*rows_2//4+1)+':'+j2+str(rows_2)].value


l1[k1+'2'].value = i[j3+'2:'+j4 + str(rows_2//4)].value
l1[k1+str(rows_2//4+1)].value = i[j3+str(rows_2//4+1)+':'+j4+str(rows_2//2)].value
l1[k1+str(rows_2//2+1)].value = i[j3+str(rows_2//2+1)+':'+j4+str(3*rows_2//4)].value
l1[k1+str(3*rows_2//4+1)].value = i[j3+str(3*rows_2//4+1)+':'+j4+str(rows_2)].value

# 用户名1
l2['A2'].options(transpose=True).value = i['J2:J'+str(rows_2)].value
# 数据
l2['B2'].value = i[j5+'2:'+j6 + str(rows_2//4)].value
l2['B'+str(rows_2//4+1)].value = i[j5+str(rows_2//4+1)+':'+j6+str(rows_2//2)].value
l2['B'+str(rows_2//2+1)].value = i[j5+str(rows_2//2+1)+':'+j6+str(3*rows_2//4)].value
l2['B'+str(3*rows_2//4+1)].value = i[j5+str(3*rows_2//4+1)+':'+j6+str(rows_2)].value

## Infeeds
# 分拆为四份 A:AU
i, l1, l2 = sht4, sht52, sht53

l1['A2'].value = i[j1+'2:'+j2 + str(rows_2//4)].value
l1['A'+str(rows_2//4+1)].value = i[j1+str(rows_2//4+1)+':'+j2+str(rows_2//2)].value
l1['A'+str(rows_2//2+1)].value = i[j1+str(rows_2//2+1)+':'+j2+str(3*rows_2//4)].value
l1['A'+str(3*rows_2//4+1)].value = i[j1+str(3*rows_2//4+1)+':'+j2+str(rows_2)].value


l1[k1+'2'].value = i[j3+'2:'+j4 + str(rows_2//4)].value
l1[k1+str(rows_2//4+1)].value = i[j3+str(rows_2//4+1)+':'+j4+str(rows_2//2)].value
l1[k1+str(rows_2//2+1)].value = i[j3+str(rows_2//2+1)+':'+j4+str(3*rows_2//4)].value
l1[k1+str(3*rows_2//4+1)].value = i[j3+str(3*rows_2//4+1)+':'+j4+str(rows_2)].value

# 用户名1
l2['A2'].options(transpose=True).value = i['J2:J'+str(rows_2)].value
# 数据
l2['B2'].value = i[j5+'2:'+j6 + str(rows_2//4)].value
l2['B'+str(rows_2//4+1)].value = i[j5+str(rows_2//4+1)+':'+j6+str(rows_2//2)].value
l2['B'+str(rows_2//2+1)].value = i[j5+str(rows_2//2+1)+':'+j6+str(3*rows_2//4)].value
l2['B'+str(3*rows_2//4+1)].value = i[j5+str(3*rows_2//4+1)+':'+j6+str(rows_2)].value
  
# =============================================================================
# wb3.save()
# wb4.save()
# wb5.save()
# =============================================================================

# 删除第二行
input('\a表单检查：')
for i in [sht32, sht33, sht42, sht43, sht52, sht53]:
    i[1,:].api.Delete()

wb3.save()
wb4.save()
wb5.save()

wb3.close()
wb4.close()
wb5.close()



'''   分割   '''


