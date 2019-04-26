# coding:utf-8
'''
1.增加：消费数据列行求和汇总 ——待
2.2018/11/20 +百通消费
3.2018/11/21 自动填充进入P4P报表
    > 基本信息
4.2018/11/27 繁转简
5.广告主【旧+新】
注意：广告主繁转简不全面，与excel自带繁转简有区别，影响HK BP
'''

import pandas as pd
import xlwings as xw
# import numpy as np
import datetime, time
from __MergeExcel import MergeIO
from __DaytoDayRatio_v3 import nearly5Workdays
from xlwings import constants


def updatePath():
    global date1
    path = 'C:\\Users\\chen.huaiyu\\Downloads\\消费报告 ' + date1 + '_' + date1 + '.csv'
    return path

def getLastWeekday():
    now = datetime.datetime.today()
    if now.isoweekday() == 1:
        dayStep = 3
    else:
        dayStep = 1
    lastWeekday = now - datetime.timedelta(dayStep)
    print('如遇节假日请手动辅助处理...')  # 待办
    return lastWeekday

def reportPath():
    # 如跨年
    date = getLastWeekday()
    path = (r'H:\SZ_数据\Input\P4P 消费报告' + str(date.year) + 
            '.' + str(date.month) + '...xlsx')
    return path

def baiTong():
    # 读取每日百度消费中百通消费数据，自动加总
    date = getLastWeekday()
    dataBT = pd.read_excel(r'H:\SZ_数据\Input\每日百度消费.xlsx', 
                           sheet_name='P4P消费'+str(date.month)+'月'
                           ).iloc[38:52, :]
    dataBT.iloc[0, 0] = '用户名'
    dataBT.iloc[-1, 0] = 'Total'
    dataBT1 = dataBT.T.set_index(38, drop=True)
    dataBT2 = dataBT1.T
    return dataBT2


if __name__ == '__main__':
    start = time.perf_counter()
    N = eval(input('（制作昨天消费报告输入1，前天的输入2，依次...）请输入：'))
    M = input('注意：季度AM调整（M）？')
    input('百通？')
    print('请稍等......')
    date = datetime.datetime.today() - datetime.timedelta(N)  # 1：前1天
    date1 = datetime.datetime.strftime(date, '%Y%m%d')
    date2 = date.replace(date.year, date.month, date.day, 0, 0, 0, 0)
    
    # 读入数据：更新数据+消费报告基本信息
    
    xiaoFei = pd.read_csv(updatePath(), encoding='gbk', engine='python', dtype={'公司名称':str})
    
    # 目标百通消费 baiTong.loc[:, ['用户名', date.replace(date.year, date.month, date.day, 0, 0, 0, 0)]]
    baiTong = baiTong()
    baiTong1 = baiTong.loc[:, ['用户名', date2]]
    
    # p4p = pd.read_excel(reportPath(), sheet_name=2, usecols='A:AC', converter={'广告主':str})  # '用户名':str, 无效？
    p4p = pd.read_excel(reportPath(), sheet_name=2, converter={'广告主':str})
    p4p['用户名'] = p4p['用户名'].astype(str)
    # 用户名统一：1852：001852；15：000015；220595：0220595；852009：0852009；
    p4p.loc[3815, '用户名'] = '000015'
    p4p.loc[1472, '用户名'] = '001852'
    p4p.loc[3820, '用户名'] = '0220595'
    p4p.loc[3827, '用户名'] = '00852009'
    print('读取文件用时：{:.3f}MIN'.format((time.perf_counter() - start)/60))
    invalidPort = ['baiduhk-p4p', 'baidu-CSHK', 'csa-cny-004', 'cny-004', 'cny-junlangkh', 'csa-cny-junlangkh']  # 开户端口
    # invalidUserNames = ['test-eee789']  # 无效用户名
    # 1.直接删除不统计的端口invalidPort
    # 2.删除无效影响统计的用户名invalidUserName
    # 3.在此计算新产品消费、首次消费日无意义、重命名（除用户名）
    # 
    # 去掉不统计的端口invalidPort
    # 去掉无效用户名invalidUserName
    xiaoFei.rename(columns={'账户名称':'用户名'}, inplace=True)
    icrmData = xiaoFei.copy()
    for i in invalidPort:
        print(i)
        icrmData = icrmData[icrmData['SF对应二级帐号'] != i]
    # 更新用户,与icrm数据合并处理，并更新相应数据
    newUserMerge = pd.merge(p4p, icrmData, how='outer', on='用户名').iloc[:,:29]
    icrmMerge = pd.merge(newUserMerge, xiaoFei, how='left', on='用户名').iloc[8:, :].reset_index()
    
    # 基础数据更新;
    updateCols = ['端口', 'URL',  '信誉成长值', '开户日期_x', '账户每日消费预算', 
                 '今日账户状态_x', '主体资质到期日_x', '加V缴费到期日', '网站名称_x',
                 '广告主', '账户ID']
    icrmMerge.drop(updateCols, axis=1, inplace=True)
    icrmMerge.rename(columns={'SF对应二级帐号':'端口', '网站URL':'URL', '合规信用值':'信誉成长值',
                              '搜索预算'+str(date1):'账户每日消费预算', '缴费到期日':'加V缴费到期日',
                              '公司名称':'广告主', '总点击消费'+str(date1):'总', '搜索点击消费'+str(date1):'搜索',
                              '自主投放消费'+str(date1):'原生', '无线搜索点击消费'+str(date1):'无线'}, 
                    inplace=True)
    icrmMerge['新产品'] = icrmMerge['总'] - icrmMerge['搜索'] - icrmMerge['原生']
    # icrmMerge['属性'] = '-'
    icrmMerge['BU'] = 'CSA'
    # icrmMerge['开户性质'] = '-'
    icrmMerge['下单方'] = '海外渠道'
    icrmMerge['销售'] = icrmMerge['销售'].str.title()
    icrmMerge['销售'] = icrmMerge['销售'].str.replace(' ', '')
    icrmMerge['AM'] = icrmMerge['AM'].str.title()
    icrmMerge['AM'] = icrmMerge['AM'].str.replace(' ', '')
    icrmMerge['开户日期'] = pd.to_datetime(icrmMerge['开户日期'])
    icrmMerge['主体资质到期日'] = pd.to_datetime(icrmMerge['主体资质到期日'])
    icrmMerge['加V缴费到期日'] = pd.to_datetime(icrmMerge['加V缴费到期日'])
    # icrm中广告主信息异常
    icrmMerge.loc[906, '广告主'] = 'AIRNEWZEALANDLIMITED'
    
    # 广告主格式统一
    icrmMerge['广告主'] = icrmMerge['广告主'].str.title()
    icrmMerge['广告主'] = icrmMerge['广告主'].str.replace(' ', '')
    icrmMerge.fillna('-', inplace=True)

    # 新消户基本信息补充:
    '''
    1.提取新户
    2.补充信息
    3.收取年服务费时间:开户日期 + 1年
    '''
    MergeIO()
    kaiHu = pd.read_excel(r'H:\SZ_数据\Input\开户进度总表.xlsx')
    kaiHu.drop_duplicates(subset='用户名', keep='first', inplace=True)
    newAccount = icrmMerge[icrmMerge['新旧客户'] == '-']
    # 删除待补充信息：余 开户日期(_x,_y)icrm中为_x;端口(_x,_y) _y判定财务做账区，_x为真实端口;
    columns = ['区域', '销售', 'AM', '操作', '新旧客户', '财务做账区域', '资质归属地', 
                    '公司总部', 'Region', 'Industry', 'channel', '客户', '收取年服务费时间']
    newAccount1 = newAccount.drop(columns, axis=1)
    
    oldAccount = icrmMerge[icrmMerge['新旧客户'] != '-'].drop(['index'], axis=1)
    newKaiHu = pd.merge(newAccount1, kaiHu, how='inner', on='用户名')
    
    
    
    ''' 军朗直接在端口下开申请开户  '''
    # 删除测试账户
    newAccount2 = newAccount.drop(index=newAccount[newAccount['用户名']=='test-eee789'].index)
    
    if newAccount2.size > 0:
        print('注:\n端口下开户。')
        
        # 筛查新申户
        newAccount2 = newAccount2.append(pd.DataFrame(newKaiHu['用户名']), sort=False)
        newAccount2.drop_duplicates(subset='用户名', keep='last', inplace=True)
        newAccount2 = newAccount2[newAccount2['BU'] == 'CSA']
        
        # 将遗漏的账户补充到新申户表
        wb = xw.Book(r'H:\SZ_数据\Input\io system.xlsx')
        wb.app.visible = False
        sht = wb.sheets(1)
        i = sht[0, 0].current_region.rows.count
        df = newAccount2.reindex(columns=sht['A1:H1'].value)
        sht[i, 0].value = df.values
        wb.save()
        wb.close()
        
        newAccount2.loc[(newAccount2['端口'] == 'baidu-Junlang') | (newAccount2['端口'] == 'cny-019')|(newAccount2['端口'] == 'csa-baidu-Junlanghk'), '销售'] = '顾凡凡'
        newAccount2.loc[(newAccount2['端口'] == 'baidu-Junlang') | (newAccount2['端口'] == 'cny-019')|(newAccount2['端口'] == 'csa-baidu-Junlanghk'), 'AM'] = '鲁东栋'
        newAccount2.loc[(newAccount2['端口'] == 'baidu-Junlang') | (newAccount2['端口'] == 'cny-019')|(newAccount2['端口'] == 'csa-baidu-Junlanghk'), 'Region'] = '香港'
        newAccount2.loc[(newAccount2['端口'] == 'baidu-Junlang') | (newAccount2['端口'] == 'cny-019')|(newAccount2['端口'] == 'csa-baidu-Junlanghk'), 'channel'] = '代理商'
        newAccount2.loc[(newAccount2['端口'] == 'baidu-Junlang') | (newAccount2['端口'] == 'cny-019')|(newAccount2['端口'] == 'csa-baidu-Junlanghk'), '客户'] = '北京军朗广告有限公司'
        
        newAccount2.rename(columns={'开户日期': '开户日期_x', '端口': '端口_x'}, inplace=True)
        
        newAccount_kaihu = newAccount2.reindex(columns=newKaiHu.columns)
        newAccount_kaihu['端口_y'] = 'cny'
        newAccount_kaihu['总部'] = '香港'
        newAccount_kaihu['资质归属地'] = '香港'
        newAccount_kaihu['行业'] = '游戏软件'
        
        # 补充开户进度总表
        data = newAccount_kaihu[['用户名', '端口_y', '销售', 'AM', '广告主', '行业', '开户日期_x']]
        data.rename(columns={'端口_y':'端口', '广告主':'广告主名称', '开户日期_x':'开户日期'}, inplace=True)
        kaiHu = kaiHu.append(data)
        kaiHu.fillna('-', inplace=True)
        kaiHu.to_excel(r'H:\SZ_数据\Input\开户进度总表.xlsx')
        
    else:
        newAccount_kaihu = pd.DataFrame()
    newKaiHu = newKaiHu.append(newAccount_kaihu, sort=True)
    
    
    #1 区域赋值
    if newKaiHu.size > 0:
        newKaiHu = newKaiHu[newKaiHu['开户日期_x'] != '-']
        flag = 1  # 有新户，标识
        newKaiHu['区域'] = '-'
        newKaiHu.loc[((newKaiHu['AM'] == '黄希腾') | (newKaiHu['AM'] == '鲁东栋') | (newKaiHu['AM'] == '陈宛欣') |
                 (newKaiHu['AM'] == '李裕岭') | (newKaiHu['AM'] == '赵宗州')), '区域'] = 'SZ'
        newKaiHu.loc[((newKaiHu['AM'] == 'Olivia') | (newKaiHu['AM'] == 'Tibby') | (newKaiHu['AM'] == 'Bruce') |
                 (newKaiHu['AM'] == 'Jessie') | (newKaiHu['AM'] == 'Jacqueline') | (newKaiHu['AM'] == 'Estelle')) 
                & (newKaiHu['channel'] == '代理商'), '区域'] = 'HK 4A'
        newKaiHu.loc[((newKaiHu['AM'] == 'Olivia') | (newKaiHu['AM'] == 'Tibby') | (newKaiHu['AM'] == 'Bruce') |
                 (newKaiHu['AM'] == 'Jessie') | (newKaiHu['AM'] == 'Jacqueline') | (newKaiHu['AM'] == 'Estelle'))
                & (newKaiHu['channel'] == '直接客户'), '区域'] = 'HK DS'
        #2.0 鲁东栋 & 陈宛欣账户分配按端口实施
# =============================================================================
#         newKaiHu.loc[((newKaiHu['端口_x'] == 'baidu-Junlang') | 
#                  (newKaiHu['端口_x'] == 'baidu-mpower') |
#                  (newKaiHu['端口_x'] == 'baidu-mpower-2') | 
#                  (newKaiHu['端口_x'] == 'baidu-mpower-3')|
#                  (newKaiHu['端口_x'] == 'cny-031') | 
#                  (newKaiHu['端口_x'] == 'cny-032') | 
#                  (newKaiHu['端口_x'] == 'cny-019') | 
#                  (newKaiHu['端口_x'] == 'csa-baidu-Junlanghk')), 'AM'] = '陈宛欣'
#         newKaiHu.loc[((newKaiHu['AM'] == '鲁东栋') | (newKaiHu['AM'] == '陈宛欣')) &
#                  ((newKaiHu['端口_x'] != 'baidu-Junlang') & (newKaiHu['端口_x'] != 'baidu-mpower') &
#                  (newKaiHu['端口_x'] != 'baidu-mpower-2') & (newKaiHu['端口_x'] != 'baidu-mpower-3') & 
#                  (newKaiHu['端口_x'] != 'cny-031') & (newKaiHu['端口_x'] != 'cny-032') & 
#                  (newKaiHu['端口_x'] != 'cny-019') & (newKaiHu['端口_x'] == 'csa-baidu-Junlanghk')), 'AM'] = '鲁东栋'
# =============================================================================
        #2 操作赋值（深圳有指定；香港无指定_随意即可）
        newKaiHu['操作'] = '-'
        newKaiHu.loc[(newKaiHu['AM'] == '黄希腾'), '操作'] = '陈熙香'
        newKaiHu.loc[(newKaiHu['AM'] == '鲁东栋') | (newKaiHu['AM'] == '陈宛欣'), '操作'] = '顾凡凡'
        'AM：据端口分配账户'
        newKaiHu.loc[((newKaiHu['AM'] == '赵宗州')|(newKaiHu['AM'] == '李裕玲'))
                    &((newKaiHu['端口_x'] == 'cny-005')|(newKaiHu['端口_x'] == 'csa-cny-008')
                    |(newKaiHu['端口_x'] == 'cny-018')|(newKaiHu['端口_x'] == 'cny-008')
                    |(newKaiHu['端口_x'] == 'baidu-boyaa')), 'AM'] = '赵宗州'
        newKaiHu.loc[((newKaiHu['AM'] == '赵宗州')|(newKaiHu['AM'] == '李裕玲'))
                    &((newKaiHu['端口_x'] != 'cny-005')|(newKaiHu['端口_x'] != 'csa-cny-008')
                    |(newKaiHu['端口_x'] != 'cny-018')|(newKaiHu['端口_x'] != 'cny-008')
                    |(newKaiHu['端口_x'] != 'baidu-boyaa')), 'AM'] = '李裕岭'
        newKaiHu.loc[(newKaiHu['AM'] == '赵宗州&李裕玲'), '操作'] = '卢铭坛'
        newKaiHu.loc[(newKaiHu['AM'] == 'Olivia') | (newKaiHu['AM'] == 'Tibby') | 
                (newKaiHu['AM'] == 'Bruce'), '操作'] = '李燕'
        newKaiHu.loc[(newKaiHu['AM'] == 'Jessie') | (newKaiHu['AM'] == 'Jacqueline')
                | (newKaiHu['AM'] == 'Estelle'), '操作'] = '卢雅洁'
        #3 新旧广告主
        # 繁转简后判定新旧
        from zhconv import convert
        seriesOldAD = oldAccount['广告主'].apply(lambda x: convert(x, 'zh-cn'))
        seriesNewAD = newKaiHu['广告主'].apply(lambda x: convert(x, 'zh-cn'))
        
        newOldAD = []
        newAD = []
        for i in seriesNewAD:
            if i in seriesOldAD.values:
                newOldAD.append('EB')
            else:
                if i in newAD:
                    newOldAD.append('EB')
                else:
                    newOldAD.append('NB')
                newAD.append(i)
        newKaiHu['新旧客户'] = newOldAD
        
        #4 财务作账区
        newKaiHu.loc[newKaiHu['端口_y'] == 'csa', '财务做账区域'] = 'HK'
        newKaiHu.loc[newKaiHu['端口_y'] == 'cny', '财务做账区域'] = 'SZ'
        
        #5 收取年服务费时间
        dateList = []
        for i in newKaiHu['开户日期_x']:
            
            j = datetime.datetime(i.year+1, i.month, i.day)
            dateList.append(j)
        newKaiHu.loc[:, '收取年服务费时间'] = dateList
        newKaiHu.rename(columns={'开户日期_x':'开户日期', '端口_x':'端口', 
                                 '行业':'Industry', '总部':'公司总部'}, 
                     inplace=True)
        
        
        ''' 新户 '''
        newAccountMessage = newKaiHu[oldAccount.columns]
        '军朗'
        newAccountMessage.loc[newAccountMessage['端口'] == 'csa-baidu-Junlanghk','财务做账区域'] = 'HK'
        
        
        ''' 旧广告主_简体+新广告主_简体 '''
        simplifiedAD = p4p.loc[8:, '广告主'].append(seriesNewAD, ignore_index=True)
        
        # P4P消费基本信息总表
        basicMessage = pd.concat((oldAccount, newAccountMessage), axis=0, join='outer', ignore_index=True, sort=True)
        basicMessage['广告主'] = simplifiedAD
        
    else:
        pass
        print('注：\n无新申请账户')
        flag = 0
        basicMessage = oldAccount.copy()
        simplifiedAD = p4p.loc[8:, '广告主'].reset_index(drop=True)
        basicMessage['广告主'] = simplifiedAD
        newAccountMessage = pd.DataFrame(columns=basicMessage.columns)
    '军朗'
    basicMessage.loc[basicMessage['客户'] == '北京军朗广告有限公司', 'AM'] = '陈宛欣'
    
    # + 百通消费    
    basic_BaiTong = pd.merge(basicMessage, baiTong1, on='用户名', how='left')
    basic_BaiTong.fillna(0, inplace=True)
    
    try:
        basic_BaiTong['新产品1'] = basic_BaiTong['新产品'] + basic_BaiTong[date2]
        basic_BaiTong['总1'] = basic_BaiTong['总'] + basic_BaiTong[date2]
    except:
        # 账户突然丢失处理
        Diushi_list = list(basic_BaiTong[basic_BaiTong['新产品'] == '-']['新产品'].index)
        for i in Diushi_list:
            for j in ['总', '搜索', '新产品', '无线', '原生']:
                basic_BaiTong.loc[i, j] = 0
        # 加百通
        basic_BaiTong['新产品1'] = basic_BaiTong['新产品'] + basic_BaiTong[date2]
        basic_BaiTong['总1'] = basic_BaiTong['总'] + basic_BaiTong[date2]
    
    # 提前消费 & 端口转移不及时
    if basic_BaiTong['总1'].sum() < xiaoFei['总点击消费'+str(date1)].sum():
        print('注意：账户提前消费or未及时转移端口!')
        pass
    else:
        print('消费正常，无遗漏')
    
    # 文档输出
    print('\a\a中间处理耗时：{:.3f}MIN'.format((time.perf_counter() - start)/60))
    path = pd.ExcelWriter("C:\\Users\\chen.huaiyu\\Desktop\\Output\\P4P消费报告1" 
                          + str(date1) + ".xlsx", engine='xlsxwriter')
    basic_BaiTong[['用户名', '新产品1', '总1', '搜索', '原生', '无线']].to_excel(path, 
                str(date1) + '消费数据', freeze_panes=(1,0), index=False)
    basicMessage[['属性', 'BU', '区域', '销售', 'AM', '操作', '开户性质', '新旧客户', 
               '端口', '用户名','财务做账区域', '资质归属地', '公司总部', 'Region',
               'Industry', 'URL', '信誉成长值','channel', '客户', '网站名称', 
               '广告主', '开户日期', '首次消费日', '收取年服务费时间', '下单方',
               '账户每日消费预算', '今日账户状态', '主体资质到期日', '加V缴费到期日'
               ]].to_excel(path, 'p4p', freeze_panes=(1,0), index=False)
    basicMessage[['属性', '区域', '销售', 'AM', '操作', '端口', '用户名', 'Region', 
               'Industry', 'URL', '信誉成长值', '网站名称', '广告主', '开户日期',
               '首次消费日']].to_excel(path, '无线', freeze_panes=(1,0), index=False)
    newAccountMessage[['属性', 'BU', '区域', '销售', 'AM', '操作', '开户性质', '新旧客户', 
               '端口', '用户名','财务做账区域', '资质归属地', '公司总部', 'Region',
               'Industry', 'URL', '信誉成长值','channel', '客户', '网站名称', 
               '广告主', '开户日期', '首次消费日', '收取年服务费时间', '下单方',
               '账户每日消费预算', '今日账户状态', '主体资质到期日', '加V缴费到期日'
               ]].to_excel(path, 'p4p新户', freeze_panes=(1,0), index=False)
    newAccountMessage[['属性', '区域', '销售', 'AM', '操作', '端口', '用户名', 'Region', 
               'Industry', 'URL', '信誉成长值', '网站名称', '广告主', '开户日期',
               '首次消费日']].to_excel(path, '无线新户', freeze_panes=(1,0), index=False)
    basicMessage[basicMessage['新旧客户'].isin(['NB'])].loc[:, ['广告主', '区域', 'AM', '首次消费日']
                ].to_excel(path, '广告主', freeze_panes=(1,0), index=False)
    
    ### 格式调整
    # 广告主加条件格式(重复duplicate)
    wb = path.book
    sht = path.sheets['广告主']
    format1 = wb.add_format({'bg_color': '#FFC7CE'})
    sht.conditional_format('A1:A'+str(basicMessage.shape[0]), {'type': 'duplicate',
                           'format': format1})
    path.save()
    
    wb_basic = xw.Book(r'H:\SZ_数据\基本信息拆解.xlsx')
    sht = wb_basic.sheets['基本信息']
    sht[0, 0].value = basic_BaiTong[['属性', 'BU', '区域', '销售', 'AM', '操作', '开户性质', '新旧客户', 
               '端口', '用户名','财务做账区域', '资质归属地', '公司总部', 'Region',
               'Industry', 'URL', '信誉成长值','channel', '客户', '网站名称', 
               '广告主', '开户日期', '首次消费日', '收取年服务费时间', '下单方',
               '账户每日消费预算', '今日账户状态', '主体资质到期日', '加V缴费到期日'
               ]]
    wb_basic.save()
    print('输出完成。')
    
    
    
    ### P4P报表操作
    #1.基本信息写入
    #2.消费数据写入
    #3.中间空格填充:
        # 2018年1月在第36列AJ列(columns中35)
    #4.
    #app = xw.App(visible=False, add_book=False)
    #wb1 = app.books.open(r'C:\Users\chen.huaiyu\Desktop\P4P 消费报告2018.11...xlsx')
    wb1 = xw.Book(r'H:\SZ_数据\Input\P4P 消费报告' + str(date.year) + '.' + str(date.month) + '...xlsx')
    wb1.app.calculation = 'manual'
    wb1.app.visible = True
    
    sht1 = wb1.sheets['P4P消费']
    sht1.range('A10').value = basicMessage[['属性', 'BU', '区域', '销售', 'AM', '操作', '开户性质', '新旧客户', 
               '端口', '用户名','财务做账区域', '资质归属地', '公司总部', 'Region',
               'Industry', 'URL', '信誉成长值','channel', '客户', '网站名称', 
               '广告主', '开户日期', '首次消费日', '收取年服务费时间', '下单方',
               '账户每日消费预算', '今日账户状态', '主体资质到期日', '加V缴费到期日'
               ]].values
    # 第10行
    sht1.range('A1').offset(9, list(p4p.columns).index(date2)).options(transpose=True).value = basic_BaiTong['总1'].values
    # 如无新消户增加，则不用执行下面的语句
    if flag:
        sht1.range('AJ1').offset(p4p.shape[0] + 1, 0).resize(newKaiHu.shape[0], list(p4p.columns).index(date2) - list(p4p.columns).index(str(date.year)+'年1月')).value = 0
        # 求和公式 总消费
        column1 = list(p4p.columns).index(str(date2.year) + '年' + str(date2.month) + '月')  # 月
        for i in range(newKaiHu.shape[0]):
            # 年
            rng1 = sht1.range('AI1').offset(p4p.shape[0] + i + 1)
            rng1.formula = '=sum(AJ' + str(rng1.row) + ':AU' + str(rng1.row) + ')'
            # 月
            rng11 = sht1.range('A1').offset(p4p.shape[0] + i + 1, column1)
            rng11.formula = '=sum(AV' + str(rng11.row) + ':BZ' + str(rng11.row) + ')'
        # 加边框
        #1 area:current_region
        for i in range(7, 13):
            sht1.range('A1').current_region.api.Borders(i).LineStyle = 1
    
    
    sht2 = wb1.sheets['搜索点击消费']
    sht2.range('A10').value = basicMessage[['属性', 'BU', '区域', '销售', 'AM', '操作', '开户性质', '新旧客户', 
               '端口', '用户名','财务做账区域', '资质归属地', '公司总部', 'Region',
               'Industry', 'URL', '信誉成长值','channel', '客户', '网站名称', 
               '广告主', '开户日期', '首次消费日', '收取年服务费时间', '下单方',
               '账户每日消费预算', '今日账户状态', '主体资质到期日', '加V缴费到期日'
               ]].values
    sht2.range('A1').offset(9, list(p4p.columns).index(date2)).options(transpose=True).value = basic_BaiTong['搜索'].values
    # 如无新消户增加，则不用执行下面的语句
    if flag:
        sht2.range('AJ1').offset(p4p.shape[0] + 1, 0).resize(newKaiHu.shape[0], list(p4p.columns).index(date2) - 35).value = 0
        # 求和公式填充
        for i in range(newKaiHu.shape[0]):
            # 年
            rng1 = sht2.range('AI1').offset(p4p.shape[0] + i + 1)
            rng1.formula = '=sum(AJ' + str(rng1.row) + ':AU' + str(rng1.row) + ')'
            # 月
            rng11 = sht2.range('A1').offset(p4p.shape[0] + i + 1, column1)
            rng11.formula = '=sum(AV' + str(rng11.row) + ':BZ' + str(rng11.row) + ')'
        # 加边框
        #1 area:current_region
        for i in range(7, 13):
            sht2.range('A1').current_region.api.Borders(i).LineStyle = 1
        
    
    sht3 = wb1.sheets['新产品消费（除原生广告）']
    sht3.range('A10').value = basicMessage[['属性', 'BU', '区域', '销售', 'AM', '操作', '开户性质', '新旧客户', 
               '端口', '用户名','财务做账区域', '资质归属地', '公司总部', 'Region',
               'Industry', 'URL', '信誉成长值','channel', '客户', '网站名称', 
               '广告主', '开户日期', '首次消费日', '收取年服务费时间', '下单方',
               '账户每日消费预算', '今日账户状态', '主体资质到期日', '加V缴费到期日'
               ]].values
    sht3.range('A1').offset(9, list(p4p.columns).index(date2)).options(transpose=True).value = basic_BaiTong['新产品1'].values
    # 如无新消户增加，则不用执行下面的语句
    if flag:
        sht3.range('AJ1').offset(p4p.shape[0] + 1, 0).resize(newKaiHu.shape[0], list(p4p.columns).index(date2) - 35).value = 0
        # 求和公式
        for i in range(newKaiHu.shape[0]):
            # 年
            rng1 = sht3.range('AI1').offset(p4p.shape[0] + i + 1)
            rng1.formula = '=sum(AJ' + str(rng1.row) + ':AU' + str(rng1.row) + ')'
            # 月
            rng11 = sht3.range('A1').offset(p4p.shape[0] + i + 1, column1)
            rng11.formula = '=sum(AV' + str(rng11.row) + ':BZ' + str(rng11.row) + ')'
        # 加边框
        #1 area:current_region
        for i in range(7, 13):
            sht3.range('A1').current_region.api.Borders(i).LineStyle = 1
    
    
    sht4 = wb1.sheets['原生广告']
    sht4.range('A10').value = basicMessage[['属性', 'BU', '区域', '销售', 'AM', '操作', '开户性质', '新旧客户', 
               '端口', '用户名','财务做账区域', '资质归属地', '公司总部', 'Region',
               'Industry', 'URL', '信誉成长值','channel', '客户', '网站名称', 
               '广告主', '开户日期', '首次消费日', '收取年服务费时间', '下单方',
               '账户每日消费预算', '今日账户状态', '主体资质到期日', '加V缴费到期日'
               ]].values
    sht4.range('A1').offset(9, list(p4p.columns).index(date2)).options(transpose=True).value = basic_BaiTong['原生'].values
    # 如无新消户增加，则不用执行下面的语句
    if flag:
        sht4.range('AJ1').offset(p4p.shape[0] + 1, 0).resize(newKaiHu.shape[0], list(p4p.columns).index(date2) - 35).value = 0
        # 求和
        for i in range(newKaiHu.shape[0]):
            # 年
            rng1 = sht4.range('AI1').offset(p4p.shape[0] + i + 1)
            rng1.formula = '=sum(AJ' + str(rng1.row) + ':AU' + str(rng1.row) + ')'
            # 月
            rng11 = sht4.range('A1').offset(p4p.shape[0] + i + 1, column1)
            rng11.formula = '=sum(AV' + str(rng11.row) + ':BZ' + str(rng11.row) + ')'
        # 加边框
        #1 area:current_region
        for i in range(7, 13):
            sht4.range('A1').current_region.api.Borders(i).LineStyle = 1
    
    
    sht5 = wb1.sheets['无线搜索点击消费']
    sht5.range('A10').value = basicMessage[['属性', '区域', '销售', 'AM', '操作', '端口', '用户名', 'Region', 
               'Industry', 'URL', '信誉成长值', '网站名称', '广告主', '开户日期',
               '首次消费日']].values
    columns_wuxian = ['属性', '区域', '销售', 'AM', '操作', '端口', '用户名', 'Region', 'Industry', 
                      'URL', '信誉成长值', '网站名称', '广告主', '开户日期', '首次消费日'] + list(p4p.columns[30:])
    sht5.range('A1').offset(9, columns_wuxian.index(date2)).options(transpose=True).value = basic_BaiTong['无线'].values
    # 如无新消户增加，则不用执行下面的语句
    if flag:
        sht5.range('U1').offset(p4p.shape[0] + 1, 0).resize(newKaiHu.shape[0], columns_wuxian.index(date2) - columns_wuxian.index(str(date.year)+'年1月')).value = 0
        # 求和
        column2 = columns_wuxian.index(str(date2.year) + '年' + str(date2.month) + '月')
        for i in range(newKaiHu.shape[0]):
            # 年
            rng1 = sht5.range('T1').offset(p4p.shape[0] + i + 1)
            rng1.formula = '=sum(U' + str(rng1.row) + ':AF' + str(rng1.row) + ')'
            # 月
            rng11 = sht5.range('A1').offset(p4p.shape[0] + i + 1, column2)
            rng11.formula = '=sum(AG' + str(rng11.row) + ':BK' + str(rng11.row) + ')'
        # 加边框
        #1 area:current_region
        for i in range(7, 13):
            sht5.range('A1').current_region.api.Borders(i).LineStyle = 1
    
    # 如无新消户增加，则不用执行下面的语句
    try:
        if len(newAD):
            sht7 = wb1.sheets['P4P广告主消费']
            column70 = sht7[4, 0].current_region.rows.count
            row70 = sht7[4, 0].current_region.columns.count
            sht7[column70, 0].value = newAccountMessage[newAccountMessage['新旧客户'].isin(['NB'])].loc[:, ['广告主', '区域', 'AM', '首次消费日']].values
            column71 = sht7[4, 0].current_region.rows.count
            sht7[column70 - 1, 10:row70].api.AutoFill(sht7[(column70 - 1):column71, 10:row70].api, constants.AutoFillType.xlFillCopy)
            for i in range(7, 13):
                sht7[4, 0].current_region.api.Borders(i).LineStyle = 1
    except:
        pass
    
    sht6 = wb1.sheets['每日消费走势']
    print('\a\a耗时：{:.3f}MIN'.format((time.perf_counter() - start)/60))
    flag6 = int(input('请完成手动计算 & 检查新消户信息，完成后回复(1)：'))
    if flag6 == 1:
        date3 = nearly5Workdays()
        sht6.range('C3').value = '环比增长额' + '\n' + str(date3[-1].day) + '日环比' + str(date3[-2].day) + '日'
        column3 = sht6.range('A3:AJ3').value.index(date3[-1])
        column4 = sht6.range('A3:AJ3').value.index(date3[-2])
        for i in range(17):
            print(i)
            sht6.range('C' + str(4 + i)).value = (sht6.range('A3').offset(i+1, column3).value) - (sht6.range('A3').offset(i+1, column4).value)
            sht6.range('D' + str(4 + i)).value = (sht6.range('C' + str(4 + i)).value)/(sht6.range('A3').offset(i+1, column4).value)

    wb1.save()
    # wb1.close()
    # app.quit()
    print('文件写入完毕。')
    print('\a\a运行结束，运行耗时：{:.3f}MIN'.format((time.perf_counter() - start)/60))
    print('2. 返点客户同步Brian.')
    print('3. 检查新旧广告主！！ \n4.广告主消费补充；\n5.军朗基本信息变更？')
    print('\n6.告知Chris军朗昨日消费：{}'.format(basic_BaiTong.loc[basic_BaiTong['客户'].isin(['北京军朗广告有限公司', '深圳麦斯动力广告有限公司']), ['客户', '总1']].groupby(['客户']).sum()))
    