# ecoding: utf-8

import os, time
import pandas as pd

def MergeExcel(PATH): 
    file = os.listdir(PATH)
    kaiHu = pd.read_excel(r'H:\SZ_数据\Input\开户进度总表.xlsx')
    for i in range(len(file)):
        kaiHu_0 = pd.read_excel(PATH + '\\' + file[i])
        kaiHu = pd.concat([kaiHu, kaiHu_0], axis=0, ignore_index=True, sort=True)
    kaiHu['用户名'] = kaiHu['用户名'].str.replace(' ', '')
    kaiHu.drop_duplicates('用户名', inplace=True, keep='last')
    kaiHu.fillna(value='-', inplace=True)
    # 销售&AM 提取英文名
    kaiHu['销售'] = pd.DataFrame((i.split(' ') for i in kaiHu['销售']), index=kaiHu.index).iloc[:, 0]
    kaiHu['销售'] = kaiHu['销售'].str.title()
    kaiHu['AM'] = pd.DataFrame((j.split(' ') for j in kaiHu['AM']), index=kaiHu.index).iloc[:, 0]
    kaiHu['AM'] = kaiHu['AM'].str.title()
    # 
    kaiHu.to_excel(r'H:\SZ_数据\Input\开户进度总表.xlsx', index=False, freeze_panes=(1,0))
    return kaiHu

def MergeIO():
    # ready！读取数据
    kaiHu = pd.read_excel(r'H:\SZ_数据\Input\开户进度总表.xlsx')
    ioSummary = pd.read_excel(r'H:\SZ_数据\Input\IO系統母版-3.07.xlsm', sheet_name=1)
    # &合并提取 Region,channel,客户
    kaiHu.set_index('用户名', inplace=True)
    ioSummary.set_index('用户名', inplace=True)
    MergeIO = pd.merge(kaiHu, ioSummary, left_index=True, right_index=True, how='left', sort=True)
    MergeIO.reset_index(inplace=True)
    MergeIO.rename(columns={'AM_x':'AM', '销售_x':'销售', 'channel_y':'channel', 'Region_y':'Region', '客户_y':'客户', 
                            '开户日期_x':'开户日期'}, inplace=True)
    kaiHuFinal = MergeIO[['AM', 'ID', 'Region', 'channel', '加款', '客户', '广告主名称', '开户', '开户日期', '总部', 
                          '截图', '推广URL', '时长（天）', '是否注册过', '未调户（上线）的客户统计', '生效日期', '用户名', 
                          '真实性验证', '端口', '行业', '账户加V', '资质上传', '资质归属地', '进度&备注', '通过验证日期', 
                          '通过验证时长', '销售', '预估月消费']]
    kaiHuFinal.fillna('-', inplace=True)
    kaiHuFinal.to_excel(r'H:\SZ_数据\Input\开户进度总表.xlsx', index=False, freeze_panes=(1,0))
    
def Check():
    p4p = input('1.最新的P4P消费报告(Input)? ：')
    master = input('2.最新的Master(Input)? :')
    if p4p == 'Y' and master == 'Y':
        return True

def Master():
    p4p = pd.read_excel(r'C:\Users\chen.huaiyu\Desktop\Input\p4p.xlsx', sheet_name=1, 
                        usecols='A:AC, AI').loc[8:, :]
    # 临时
    p4p = p4p.iloc[8:, :].reset_index()  # 临时
    master = pd.read_excel(r'C:\Users\chen.huaiyu\Desktop\Input\master.xlsm', sheet_name=1, usecols=1)
    # 1.已消费 >0
    data1 = p4p[p4p['18年已消费'] != 0]
    # 2.新用户名
    master.rename(columns={'客戶用户名':'用户名'}, inplace=True)
    merge1 = pd.merge(master, data1, on='用户名', how='right')
    merge1.fillna(value='-', inplace=True)
    merge2 = merge1[merge1['財務加款的端口'] == '-']
    # 3.除端口加款
    channel = pd.read_excel(r'c:\users\chen.huaiyu\Desktop\Input\channel.xlsx', sheet_name='master')
    for i in channel['端口加款'].tolist():
        merge2 = merge2[merge2['端口'] != i]
    # 4.结构调整
    columns = ['端口', '用户名', '财务做账区域', '付款方式', '销售', 'AM', '操作', '销售郵箱地址',
       'AM郵箱地址', 'OP郵箱地址', '客户', '网站名称', '广告主', '客户地址', 'URL', 'Industry',
       '付款币种', '联络人姓名', '联络人邮箱', '联络人电话', '开户日期', '收取年服务费时间',
       '收取年服务费（元）\n（不收取的填0）', '续费返点率（%）\n（不收取的填0）', '管理费率（%）\n（不收取的填0）',
       '预付/ 账期', '付款条款', '付款时间表', 'Unnamed: 28', 'Unnamed: 29', ]
    newMaster = merge2.reindex(columns=columns)
    # 5.信息补充
    # 5.1 财务
    newMaster.loc[newMaster['财务做账区域'] == 'SZ', '财务做账区域'] = 'Automation-SZ財務查賬'
    newMaster.loc[newMaster['财务做账区域'] == 'HK', '财务做账区域'] = 'Automation-HK財務查賬'
    # 5.2 AM
    newMaster.loc[newMaster['AM'] == '赵宗州&李裕玲', 'AM'] = '赵宗州 & 李裕玲'
    newMaster.loc[(newMaster['AM'] == 'Claire') | (newMaster['AM'] == 'Jacqueline') |
                  (newMaster['AM'] == 'Estelle'), 
                  'AM'] = 'Billy & Claire & Jacqueline & Estelle'
    newMaster.loc[(newMaster['AM'] == 'Kendi') | (newMaster['AM'] == 'Tibby') | 
                  (newMaster['AM'] == 'Bruce') | (newMaster['AM'] == 'Olivia'), 
                  'AM'] = 'Kendi & Cindy & Tibby & Bruce & Olivia'
    print('注意：\n1.与"操作"绑定的户；\n2.抄货币；\n3.抄年服务费；\n4.预付/账期.')
    # 5.3 货币、年服务费、预付/账期确认
    newMaster.to_excel(r'C:\Users\chen.huaiyu\Desktop\Output\newMaster.xlsx')

if __name__ == '__main__':
    start = time.clock()
    # Master表单处理
    if Check():
        Master()
    # 其它
    PATH = r'C:\Users\chen.huaiyu\Desktop\Output\20180906开户进度表'
    f = MergeExcel(PATH)
    f.to_excel(r'C:\Users\chen.huaiyu\Desktop\Output\开户进度总表\开户进度总表.xlsx', index=False, freeze_panes=(1,0))
    print('耗时：{}'.format(time.clock() - start))
