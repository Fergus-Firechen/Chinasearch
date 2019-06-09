# -*- coding: utf-8 -*-
"""
Created on Thu Aug 23 15:55:02 2018
日环比
# 2018/9/15 增环比日期选择：m=2, 12日环比11日
# 跨月处理
# 一次性读入；设置为全局变量；空间换时间:9.38min==>3.856min;
# 调整：新产品（含原生）

2018/9/29 修正
# Q1: 广告主对应AM与excle排序后不一致？因采取不稳定算法排序，取到的AM不一定；
# A1：改为稳定排序mergesort; 先对index排序,后对YTD排序；
# Q2: summary日P4P消费与P4P消费报告不一致？因数据源采用广告主列修正AM后的，少量数据加总到其它AM上；
# A2：修改数据源；am与ad_index的区别是什么？AM; 主导环比summary的是什么？AM; 关键因素发生了变化！ 修正回原am中的AM；

@author: chen.huaiyu
"""

import pandas as pd
import datetime, time, os
from win32com.client import Dispatch

# 近5日工作日日期序列生成
def nearly5Workdays(n=7, m=0):
    date1 = datetime.datetime.today() - datetime.timedelta(m)  # 调整运行时间
    date2 = date1 - datetime.timedelta(n)  # timedelta(i)表示时间差对象
    date3 = []
    [date3.append(i) for i in pd.date_range(date2, periods=n, normalize=True)
    if i.weekday() not in [5, 6]]                    # 条件判断出错！？
    return date3

# 输出文件命名
def giveName():
    global date1
    path = ('C:\\Users\\chen.huaiyu\\Desktop\\Output\\' + str(date1[-1].day) + 
            '日环比' + str(date1[-2].day) + '日 v1.xlsx')
    return path

# 源数源地址
def sourceData():
    global date1
    # date3 = date1.year  # 是否跨年？  待判定
    date2 = date1[-1].year
    date3 = date1[0].year
    if date1[0].month == date1[-1].month:   # 是否跨月
        path = (r'H:\SZ_数据\Input\P4P 消费报告' + str(date2) + '.'
                    + str(date1[-1].month) + '...xlsx')
        path1 = ''
    else:
        path = (r'H:\SZ_数据\Input\P4P 消费报告' + str(date2) + '.'
                    + str(date1[-1].month) + '...xlsx')
        path1 = (r'H:\SZ_数据\Input\P4P 消费报告' + str(date3) + '.'
                    + str(date1[0].month) + '...xlsx')
    return path, path1

# P4P消费转移
def P4P():
    # 数字读入后被默认为nan，写入excel为数字格式，非文本格式；调整如下：
    p4p = pd.read_excel(sourceData()[0], sheet_name=2, header=0, converters={u'广告主':str})  # 默认：sheet_name=2
    p4p['AM'] = p4p['AM'].str.title()
    p4p['广告主'] = p4p['广告主'].str.upper()
    return p4p

# 指定列提取:搜索(sheetN=3)
def souSuo(sheetN=3):
    global date1
    date = datetime.datetime.today()
    if sourceData()[1] == '':
        s1 = pd.read_excel(sourceData()[0], sheet_name=sheetN)
        s5 = s1.loc[:, ['用户名'] + date1].iloc[8:, :]
    else:
        # 涉及跨月，筛选当月工作日
        j = 0
        for i in date1:
            if i.month == date.month:
                j += 1
            else:
                pass
        s1 = pd.read_excel(sourceData()[0], sheet_name=sheetN)  # 当月
        s2 = pd.read_excel(sourceData()[1], sheet_name=sheetN)  # 上个月
        # s2 = s1.reindex(columns=[['用户名'] + nearly5Workdays()]) 错了！
        # usecols = [['用户名'] + nearly5Workdays()]                亦错！
        s3 = s1.loc[:, ['用户名'] + date1[-j:]].iloc[8:, :]
        # s3.set_index('用户名', drop=True, inplace=True)
        s4 = s2.loc[:, ['用户名'] + date1[:-j]].iloc[8:, :]
        # s4.set_index('用户名', drop=True, inplace=True)
        # s5 = pd.concat((s4, s3), axis=1)
        s5 = pd.merge(s4, s3, how='outer', on='用户名')
        # s5.reset_index(inplace=True)
        s5.fillna(value=0, inplace=True)
    s5['均值'] = s5.mean(1)  # mean() 列求均值；mean(1) 行求均值
    s5.reset_index(inplace=True)
    s5.fillna(value=0, inplace=True)
    return s5

# 指定提取：新产品
def xinChanPin_infeeds():
    global date1
    x3 = souSuo(4)  # 新产品，默认值：4
    x4 = souSuo(5)  # 原生，默认值：5
    x5 = pd.merge(x3, x4, on='用户名')
    x5[date1[0]] = x5.loc[:, str(date1[0]) + '_x'] + x5.loc[:, str(date1[0]) + '_y']
    x5[date1[1]] = x5.loc[:, str(date1[1]) + '_x'] + x5.loc[:, str(date1[1]) + '_y']
    x5[date1[2]] = x5.loc[:, str(date1[2]) + '_x'] + x5.loc[:, str(date1[2]) + '_y']
    x5[date1[3]] = x5.loc[:, str(date1[3]) + '_x'] + x5.loc[:, str(date1[3]) + '_y']
    x5[date1[4]] = x5.loc[:, str(date1[4]) + '_x'] + x5.loc[:, str(date1[4]) + '_y']
    x6 = x5.loc[:, ['用户名'] + date1]
    x6['均值'] = x6.mean(1)
    return x6

# 账户消费环比
def ringRatio():
    global p4p, sou_suo, xin_chan_pin, date1
    # 格式统一：AM.title; ad.upper
    # list_1 = []
    # list_2 = []
    rR = p4p[["AM", "用户名", "URL", "广告主", "客户", "端口"]][8:]
    rR.reset_index(drop=True, inplace=True)  # 临时
    sou_suo_2 = sou_suo.loc[:, [date1[-2], date1[-1], '用户名', '均值']]
    rR = pd.merge(rR, sou_suo_2, on='用户名')
    rR.rename(columns={date1[-2] : str(date1[-2].day) + '日搜索', 
                       date1[-1] : str(date1[-1].day) + '日搜索', 
                       '均值' : '近5工作日均值1'}, inplace=True)
    rR['搜索日环比'] = rR[str(date1[-1].day) + '日搜索'] - rR[str(date1[-2].day) + '日搜索']
    
    # 当除数为0时，结果致为0
    rR['比例1'] = rR['搜索日环比'] / rR[str(date1[-2].day) + '日搜索']  # 注意除数为0 
    rR.loc[rR[str(date1[-2].day) + '日搜索'] == 0, '比例1'] = 0
    
    rR['昨日VS前5工作日1'] = rR[str(date1[-1].day) + '日搜索'] - rR['近5工作日均值1']
    # 新产品
    xin_chan_pin_2 = xin_chan_pin.loc[:, [date1[-2], date1[-1], '用户名', '均值']]
    rR = pd.merge(rR, xin_chan_pin_2, on='用户名')
    rR.rename(columns={date1[-2] : str(date1[-2].day) + '日新产品', 
                       date1[-1] : str(date1[-1].day) + '日新产品', 
                       '均值' : '近5工作日均值2'}, inplace=True)
    rR['新产品日环比'] = rR[str(date1[-1].day) + '日新产品'] - rR[str(date1[-2].day) + '日新产品']
    
    # 当除数为0时，结果致为0
    rR['比例2'] = rR['新产品日环比'] / rR[str(date1[-2].day) + '日新产品']  # 除数为0
    rR.loc[rR[str(date1[-2].day) + '日新产品'] == 0, '比例2'] = 0
    
    rR.fillna(value=0, inplace=True)
    rR['昨日VS前5工作日2'] = rR[str(date1[-1].day) + '日新产品'] - rR['近5工作日均值2']
    # p4p
    rR['P4P日环比'] = rR['搜索日环比'] + rR['新产品日环比']
    rR['昨日P4P日均环比'] = rR['昨日VS前5工作日1'] + rR['昨日VS前5工作日2']
    rR = rR[['AM', '用户名', 'URL', '广告主', '客户', '端口', str(date1[-2].day) + '日搜索', 
             str(date1[-1].day) + '日搜索', '搜索日环比', '比例1', '近5工作日均值1', 
             '昨日VS前5工作日1', str(date1[-2].day) + '日新产品', str(date1[-1].day) + '日新产品', 
             '新产品日环比', '比例2', '近5工作日均值2', '昨日VS前5工作日2', 'P4P日环比', '昨日P4P日均环比']]
    return rR

# 广告主消费环比
def adRingRatio():
    global p4p, date1
    # 取AM,逻辑；18年已消费降序后取第一个AM
    p4pAll = p4p.loc[8:, ['用户名', str(date1[-1].year)[-2:] + '年已消费']].reset_index(drop=True)
    am = pd.merge(ringRatio(), p4pAll, on='用户名')
    # am['AM'] = am['AM'].str.title()
    #　am['广告主'] = am['广告主'].str.upper()
    am.sort_index(axis=0, ascending=True, inplace=True, kind='mergesort')
    am.sort_values(by=[str(date1[-1].year)[-2:]+'年已消费'], ascending=False, inplace=True, kind='mergesort')
    am.reset_index(drop=True, inplace=True)
    # 注：降序之后，set时变为无序，即降序无效了
    ad1 = list(set(am.loc[:, '广告主']))
    # 按降序后的索引广告主对应的第一个AM统一修改
    ad_index = am.set_index('广告主')
    for i in ad1:
        try:
            ad_index.loc[i, 'AM'] = ad_index.loc[i, 'AM'].iloc[0]
        except AttributeError:
            pass
    ad_index.drop(labels=[str(date1[-1].year)[-2:] + '年已消费'], axis=1, inplace=True)
    ad_index.reset_index(inplace=True)
    
    # 数据换算
    ad_index[str(date1[-2].day) + '日P4P消费'] = (ad_index[str(date1[-2].day) + '日搜索']
                                    + ad_index[str(date1[-2].day) + '日新产品'])
    ad_index[str(date1[-1].day) + '日P4P消费'] = (ad_index[str(date1[-1].day) + '日搜索']
                                        + ad_index[str(date1[-1].day) + '日新产品'])
    ad_index['近5工作日均P4P'] = ad_index['近5工作日均值1'] + ad_index['近5工作日均值2']
    ar2 = ad_index.loc[:, ['AM', '广告主', str(date1[-2].day) + '日P4P消费', 
                   str(date1[-1].day) + '日P4P消费', '近5工作日均P4P']]
    ar4 = ar2.groupby(by=['广告主', 'AM']).sum()
    ar4['环比增长'] = (ar4[str(date1[-1].day) + '日P4P消费'] - ar4[str(date1[-2].day) 
                      + '日P4P消费'])
    ar4['昨日环比日均'] = (ar4[str(date1[-2].day) + '日P4P消费'] - ar4['近5工作日均P4P'])
    ar4.reset_index(inplace=True)
    ar = ar4[['AM', '广告主', str(date1[-2].day) + '日P4P消费', str(date1[-1].day) + '日P4P消费', 
              '环比增长', '近5工作日均P4P', '昨日环比日均']]
    ar.sort_values(by=['环比增长'], ascending=True, inplace=True)
    return ad_index, ar, am


def agencyRingRatio():
    '代理商消费环比'
    global p4p, date1, ringRatio1
    
    p4pAll = p4p.loc[8:, ['用户名', 'channel']].reset_index(drop=True)
    ad_index = pd.merge(ringRatio1, p4pAll, on='用户名')
    ad_index['客户'] = ad_index['客户'].str.upper()
    
    # 数据换算
    ad_index[str(date1[-2].day) + '日P4P消费'] = (ad_index[str(date1[-2].day) + '日搜索']
                                    + ad_index[str(date1[-2].day) + '日新产品'])
    ad_index[str(date1[-1].day) + '日P4P消费'] = (ad_index[str(date1[-1].day) + '日搜索']
                                        + ad_index[str(date1[-1].day) + '日新产品'])
    ad_index['近5工作日均P4P'] = ad_index['近5工作日均值1'] + ad_index['近5工作日均值2']
    ar2 = ad_index.loc[:, ['客户', 'channel', str(date1[-2].day) + '日P4P消费', 
                   str(date1[-1].day) + '日P4P消费', '近5工作日均P4P']]
    ar4 = ar2.groupby(by=['客户', 'channel']).sum()
    ar4['环比增长'] = (ar4[str(date1[-1].day) + '日P4P消费'] - ar4[str(date1[-2].day) 
                      + '日P4P消费'])
    ar4['昨日环比日均'] = (ar4[str(date1[-2].day) + '日P4P消费'] - ar4['近5工作日均P4P'])
    ar4.reset_index(inplace=True)
    ag = ar4[['客户', 'channel', str(date1[-2].day) + '日P4P消费', str(date1[-1].day) + '日P4P消费', 
              '环比增长', '近5工作日均P4P', '昨日环比日均']]
    ag.sort_values(by=['环比增长'], ascending=True, inplace=True)
    return ag


# 环比汇总表
def ringSummary():
    global date1, ad
    D1 = str(date1[-2].day) + '日P4P消费' # 前日
    D2 = str(date1[-1].day) + '日P4P消费' # 昨日
    rs1 = ad[0].loc[:, ['AM', D1, D2, '近5工作日均P4P', 'P4P日环比']]
    rs1['AM'] = ad[2].loc[:, 'AM']
    rs1['环比'] = rs1[D2] - rs1[D1]
    rs1['昨日环比日均'] = rs1[D2] - rs1['近5工作日均P4P']
    list_1 = []
    list_2 = []
    for i in rs1.loc[:, 'AM']:
        if i in ['黄希腾', '赵宗州', '李裕玲']:
            list_1.append('深圳区')
            list_2.append('麦静施')
        elif i in ['鲁东栋', '陈宛欣']:
            list_1.append('深圳区')
            list_2.append('刘婷')
        elif i in ['Olivia', 'Tibby', 'Bruce']:
            list_1.append('香港区')
            list_2.append('Kendi组')
        elif i in ['Jacqueline',  'Estelle', 'Jessie']:
            list_1.append('香港区')
            list_2.append('Billy组')
        elif i in ['Stella', 'Yiwen']:
            list_1.append('SG区')
            list_2.append('SG')
        else:
            list_1.append('-')
            list_2.append('-')
    rs1.loc[:, '区域'] = list_1
    rs1.loc[:, '组长'] = list_2
    rs1 = rs1[rs1['区域'] != '-']
    rs = rs1.groupby(by=['区域', '组长', 'AM']).sum()
    # rs = rs1[[D1, D2, '环比', '近5工作日均P4P', '昨日环比日均']].groupby(rs1['AM']).sum()
    rs['下降账户数'] = rs1[rs1['P4P日环比'] < 0].groupby(by=['区域', '组长', 'AM']).size()
    rs['下降金额'] = rs1[rs1['P4P日环比'] < 0].loc[:, ['区域', '组长', 'AM', 'P4P日环比']].groupby(by=['区域', '组长', 'AM']).sum()
    # rs.dropna(axis=0, how='any', inplace=True)
    rs = rs[[str(date1[-2].day) + '日P4P消费', str(date1[-1].day) + '日P4P消费', 
             '环比', '近5工作日均P4P', '昨日环比日均', '下降账户数', '下降金额']]
    rs.fillna(value=0, inplace=True)
    return rs

def transferSummary():
    ring = ringSummary()   #  注： 小计后还需按区域合计
    ring1 = ring.reset_index(level=0).set_index(keys='区域', drop=True)  # 区域AM消费；区域求和
    ring2 = ring.reset_index().set_index(keys='区域', drop=True).drop_duplicates(subset='组长', keep='first')  # 辅助变量,各区域下组长
    arr0 = sorted(list(set(list(ring.index.get_level_values(level=0)))), key=list(ring.index.get_level_values(level=0)).index)  # 0级索引值；构建最终表格的MultiIndex
    arr1 = ring.index.get_level_values(level=1)
    arr3 = sorted(list(set(list(arr1))), key=list(arr1).index)  # 1级索引去重，顺序不变;小组求和；索引 
    arr4 = list(ring2.index)  # 区域对应组长列表；区域求和；索引
    count1 = pd.DataFrame(dict((j, list(arr1).count(j)) for j in arr3), index=['count']).T  # 统计各组AM人数
    n = 0
    data4 = pd.DataFrame()
    ring = ring.reset_index(level=2).set_index(keys='AM', drop=True)  # 各组AM消费
    for j in count1.values:
        data1 = ring.iloc[: j[0], :]
        ring.drop(index=ring.index[:j[0]], inplace=True)
        data2 = pd.DataFrame(data1.apply(lambda x: x.sum(), axis=0), columns=[count1.index[n]]).T  # 小组消费小计
        area = pd.DataFrame(ring1.loc[arr4[n], :].apply(lambda x: x.sum(), axis=0), columns=[arr4[n]]).T  # 区域消费小计
        data3 = data1.append(data2)
        data3 = data3.append(area)
        data4 = data4.append(data3)
        n += 1
    # 删除多余的区域汇总：索引去重，保留last
    data4.reset_index(inplace=True)
    data4.drop_duplicates(subset='index', keep='last', inplace=True)
    data4.set_index(keys='index', inplace=True)
    # 构造多重索引
    data5 = data4.copy()
    areaList = []
    amList = []
    for n in range(len(arr0)):
        class0 = list(data4.loc[:arr0[n], :].index)
        class1 = [arr0[n]] * len(class0)
        areaList.extend(class1)
        amList.extend(class0)
        data4.drop(index=class0, inplace=True)
    index = pd.MultiIndex.from_arrays([areaList, amList], names=['区域', 'AM'])
    data6 = pd.DataFrame(data5.values, index=index, columns=data5.columns)
    data6.reset_index(inplace=True)
    return data6

if __name__ == '__main__':
    start = time.perf_counter()
    print('请稍等...')
# =============================================================================
    date1 = nearly5Workdays()  # 默认值：n=7, m=0  最近7天，昨日数据为0
# =============================================================================
    p4p = P4P()
    sou_suo = souSuo()
    xin_chan_pin = xinChanPin_infeeds()
    ad = adRingRatio()
    
    # 日环比summary，账户消费环比
    summary = transferSummary()
    ringRatio1 = ringRatio()
    ad1 = ad[1]
    ag = agencyRingRatio()
    # 格式转换
    summary1 = pd.DataFrame()
    ringRatio2 = pd.DataFrame()
    ad2 = pd.DataFrame()
    ag1 = pd.DataFrame()
    print("\a耗时：{0:.3f}min".format((time.perf_counter() - start)/60))
    print('请稍等...')
    
    path = pd.ExcelWriter(giveName(), engine='xlsxwriter')
# =============================================================================
#     p4p.to_excel(path, sheet_name='P4P消费', index=False, freeze_panes=(1,10))
# =============================================================================
    '''
    sou_suo.to_excel(path, sheet_name='近5日搜索', index=False, 
                      freeze_panes=(1,0))
    xin_chan_pin.to_excel(path, sheet_name='近5日新产品(含原生)', index=False, 
                          freeze_panes=(1,0))
    '''
    summary1.to_excel(path, sheet_name='环比汇总表', freeze_panes=(1,0))
    ringRatio2.to_excel(path, sheet_name='账户消费环比', freeze_panes=(1,0))
    ad2.to_excel(path, sheet_name='广告主消费环比', freeze_panes=(1,0))
    ag1.to_excel(path, sheet_name='客户消费环比', freeze_panes=(1,0))
# =============================================================================
#     ringRatio().to_excel(path, sheet_name='账户消费环比', index=False, na_rep=0,
#                        freeze_panes=(1,0))
# =============================================================================
# =============================================================================
#     ad[1].to_excel(path, sheet_name='广告主消费环比', index=False, freeze_panes=(1,0))
# =============================================================================
    
    ### 【环比汇总表】格式设置
    workbook = path.book
    worksheet1 = path.sheets['环比汇总表']
    # 1.首行&列宽
    firstRow = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': '#F27F1B', 'bold': True, 'border': 1})
    worksheet1.set_row(0, 20)
    worksheet1.set_column(0, len(list(summary.columns))-1, 15)
    worksheet1.write_row(0, 0, list(summary.columns), firstRow)
    # 2.数据写入
    formatAm = workbook.add_format({'border': 1, 'num_format': '#,##0'})
    formatGroup = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#7799A6', 'num_format': '#,##0'})
    formatArea = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#F2B705', 'num_format': '#,##0'})
    format1 = workbook.add_format({'bg_color': '#FFC7CE'})  # 条件格式
# =============================================================================
#     # 注意AM变更,list中数需要修改
# =============================================================================
    for n, item in enumerate(summary.values):
        if n in [2, 6, 10, 15, 19]:  # ！！
            worksheet1.write_row(n+1, 0, item, formatGroup)
        elif n in [3, 11, 20]:  # ！！
            worksheet1.write_row(n+1, 0, item, formatArea)
        else:
            worksheet1.write_row(n+1, 0, item, formatAm)
            worksheet1.conditional_format('E'+str(n+2)+':'+'G'+str(n+2), {'type': 'cell', 
                                          'criteria': '<', 'value': 0, 'format': format1})  # 条件格式；实际对应的表单
    # 3.合并区域
    center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True})
    worksheet1.merge_range('A2:A5', 'SG区', center)
    centerBg1 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'bg_color': '#7799A6'})
    worksheet1.merge_range('A6:A13', '深圳区', centerBg1)
    centerBg2 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'bg_color': '#F2B705'})
    worksheet1.merge_range('A14:A22', '香港区', centerBg2)
    # 4.隐藏SG区
    for n in range(list(summary['区域']).count('SG区')):
        worksheet1.set_row(n + 1, None, None, {'hidden':True})  # 隐藏“SG区”
    
    
    ### 【账户消费环比】格式设置
    # Q1.inf无法写入excel?:
    # A1. to_excel( inf_rep=0)
    worksheet2 = path.sheets['账户消费环比']
    
    # 1.首行
    firstRow1 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color':'F2B705', 'bold': True, 'border': 1})
    worksheet2.write_row(0, 0, list(ringRatio1.columns)[0:9], firstRow)
    worksheet2.write_row(0, 9, list(ringRatio1.columns)[9:11], firstRow1)
    worksheet2.write_row(0, 11, list(ringRatio1.columns)[11:15], firstRow)
    worksheet2.write_row(0, 15, list(ringRatio1.columns)[15:17], firstRow1)
    worksheet2.write_row(0, 17, list(ringRatio1.columns)[17:], firstRow)
    
    # 2。列宽
    worksheet2.set_column(0, 19, 18)
    worksheet2.set_row(0, 20)
    
    # 3.写入数据
    formatCenter = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0', 'border': 1})
    formatLeft = workbook.add_format({'align': 'left', 'border': 1})
    for n, i in enumerate(ringRatio1.values.T):
        if n < 5:
            worksheet2.write_column(1, n, i, formatLeft)
        else:
            worksheet2.write_column(1, n, i, formatCenter)
    # 4. 条件格式
    worksheet2.conditional_format('H1:I'+str(ringRatio1.shape[0]), {'type': 
        'cell', 'criteria': '<', 'value': 0, 'format': format1})
    worksheet2.conditional_format('K1:K'+str(ringRatio1.shape[0]), {'type': 
        'cell', 'criteria': '<', 'value': 0, 'format': format1})
    worksheet2.conditional_format('N1:O'+str(ringRatio1.shape[0]), {'type': 
        'cell', 'criteria': '<', 'value': 0, 'format': format1})
    worksheet2.conditional_format('Q1:S'+str(ringRatio1.shape[0]), {'type': 
        'cell', 'criteria': '<', 'value': 0, 'format': format1})
    

    ### 广告主消费环比
    worksheet3 = path.sheets['广告主消费环比']
    # 1.首行
    worksheet3.write_row(0, 0, list(ad1.columns), firstRow)
    # 2.列宽
    worksheet3.set_column(0, ad1.shape[1], 18)
    # 3.写入数据
    for n, i in enumerate(ad1.values.T):
        if n < 2:
            worksheet3.write_column(1, n, i, formatLeft)
        else:
            worksheet3.write_column(1, n, i, formatCenter)
    # 4.条件格式
    worksheet3.conditional_format('E1:E'+str(ad1.shape[0]), {'type': 'cell', 
                                  'criteria': '<', 'value': 0, 'format': format1})
    worksheet3.conditional_format('G1:G'+str(ad1.shape[0]), {'type': 'cell', 
                                  'criteria': '<', 'value': 0, 'format': format1})
    
    
    ### 代理商消费环比
    worksheet3 = path.sheets['客户消费环比']
    # 1.首行
    worksheet3.write_row(0, 0, list(ag.columns), firstRow)
    # 2.列宽
    worksheet3.set_column(0, ag.shape[1], 18)
    # 3.写入数据
    for n, i in enumerate(ag.values.T):
        if n < 2:
            worksheet3.write_column(1, n, i, formatLeft)
        else:
            worksheet3.write_column(1, n, i, formatCenter)
    # 4.条件格式
    worksheet3.conditional_format('E1:E'+str(ag.shape[0]), {'type': 'cell', 
                                  'criteria': '<', 'value': 0, 'format': format1})
    worksheet3.conditional_format('G1:G'+str(ag.shape[0]), {'type': 'cell', 
                                  'criteria': '<', 'value': 0, 'format': format1})
    path.save()
    
    
    
    # 复制sheet
    excel = Dispatch('Excel.Application')
    wb1 = excel.Workbooks.Open(giveName())
    sht1 = wb1.Worksheets('环比汇总表')
    
    yesDateStr = (datetime.date.today()-datetime.timedelta(1)).strftime('%Y.%m.%d')
    fileName = 'P4P 消费报告' + yesDateStr + '.xlsx'
    filePath = r'c:\users\chen.huaiyu\desktop'
    path = os.path.join(filePath, fileName)
    wb2 = excel.Workbooks.Open(path)
    sht2 = wb2.Worksheets('P4P消费')
    
    sht2.Copy(sht1)
    wb1.Save()
    #wb1.Close()
    #wb2.Close()
    
    
    '''
    # 邮件加载附件发送
    from email.mime.text import MMIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import Header
    from email.utils import parseaddr, formataddr
    from email import encoders
    import smtplib
    
    from_addr = 'chenhuaiyu@baiduhk.com.hk'
    password = 'huai@2018'
    to_addr = 'chenhuaiyu@baiduhk.com.hk'
    smtp_server = 'smtp.office365.com'
    
    
    
    '''
    
    
    print("\a耗时：{0:.3f}min".format((time.perf_counter() - start)/60))
    print('注意：\n1.AM变更；')
    print('程序结束')

