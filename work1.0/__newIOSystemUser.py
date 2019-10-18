# -*- coding: utf-8 -*-
"""
Created on Wed Sep 26 15:19:14 2018
# 1.
# 2.
@author: chen.huaiyu
"""
import os
import time
import datetime
import configparser
import pandas as pd
from sqlalchemy import create_engine


def connectDB():
    def login():
        CONF = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        if os.path.exists(CONF):
            conf.read(CONF)
            host = conf.get('SQL Server', 'ip')
            port = conf.get('SQL Server', 'port')
            dbname = conf.get('SQL Server', 'dbname')   
            return host, port, dbname
    try:
        ss = "mssql+pymssql://@%s:%s/%s"
        engine = create_engine(ss % login())
    except Exception as e:
        print('Login failed %s' % e)
    else:
        print('Successful')
        return engine

engine = connectDB()


def col_name():
    sql = ''' SELECT * 
        FROM information_schema.columns 
        WHERE table_name='开户申请表'
        '''
    col = [i[3] for i in engine.execute(sql).fetchall()]
    return col

def data():
    sql = ''' SELECT *
        FROM 开户申请表
        '''
    dt = engine.execute(sql).fetchall()
    return dt

def mergeExcel():
    '''
    1.
    2.
    '''
    print('测试：\n1.默认昨日；')
    # 昨日icrm消费报告地址
    date = datetime.datetime.strftime(datetime.datetime.today()-datetime.timedelta(1), '%Y%m%d')  # 昨日 默认1
    icrmPath = r'C:\Users\chen.huaiyu\Downloads\消费报告 ' + date + '_' + date + '.csv'
    # 数据读取
    ioSystem = pd.read_excel(r'H:\SZ_数据\Input\IO System.xlsx')  # 新申户
    
    
    # 去重
    ioSystem.drop_duplicates('用户名')
    icrm = pd.read_csv(icrmPath, encoding='gbk', engine='python')  # 前日icrm
    #account = pd.read_excel(r'H:\SZ_数据\Input\开户进度总表.xlsx')  # 开户进度总表
    account = pd.DataFrame(data(), columns=col_name())
    
    # 合并——新申户&icrm，提取网站名称、网站URl
    icrm.set_index('账户名称', drop=True, inplace=True)
    ioSystem.set_index('用户名', drop=True, inplace=True)
    icrmIO = pd.merge(icrm, ioSystem, left_index=True, right_index=True, how='right', sort=False)  # 右侧索引合并
    icrmIO1 = icrmIO[['网站名称', '客户', '网站URL', 'Region', 'channel', '付款币种', '年服务费', '预付/账期']]
    
    # &开户进度总表——合并
    account.set_index('用户名', drop=True, inplace=True)
    account.drop(['客户', 'Region', '渠道'], axis=1, inplace=True)
    account1 = pd.merge(icrmIO1, account, how='left', left_index=True, right_index=True, sort=False)
    account1.reset_index(inplace=True)
    account1.rename(columns={'index': '用户名'}, inplace=True)
    # 结构调整
    biaoTou = ['端口', '用户名', '查賬財務郵箱地址', '付款方式', '销售', 'AM', 
               '操作', '=', '==', '===', '客户', '网站名称', '广告主名称', 
               '*', '网站URL', '行业', '付款币种', '***', '****', '*****', 
               '开户日期', 'channel', 'Region', '-', '年服务费', '--', '---', 
               '预付/账期']
    account2 = account1.reindex(columns=biaoTou)
    account2.fillna(value='-', inplace=True)
    # 内容调整&补充
    # 香港
    account2.loc[account2['端口'] == 'csa', '端口'] = 'csa-cny-004'
    account2.loc[account2['端口'] == 'csa-cny-004', '查賬財務郵箱地址'] = 'Automation-HK財務查賬'
    account2.loc[account2['端口'] == 'csa-cny-004', '付款方式'] = 'China Search (Asia) Limited'
    account2.loc[(account2['AM'] == 'Billy') | (account2['AM'] == 'Jessie') | 
            (account2['AM'] == 'Estelle') | (account2['AM'] == 'Jacqueline'), 'AM'] = 'Billy & Jacqueline & Estelle & Jessie'
    account2.loc[account2['AM'] == 'Billy & Jacqueline & Estelle & Jessie', '操作'] = '吴景虹 & 卢雅洁'
    account2.loc[(account2['AM'] == 'Kendi') | (account2['AM'] == 'Cindy') | (account2['AM'] == 'Olivia') |
            (account2['AM'] == 'Tibby') | (account2['AM'] == 'Bruce'), 'AM'] = 'Kendi & Cindy & Tibby & Bruce & Olivia'
    account2.loc[account2['AM'] == 'Kendi & Cindy & Tibby & Bruce & Olivia', '操作'] = '董湘君 & 徐琳玲'
    account2.loc[account2['AM'] == 'Stella', '操作'] = '董湘君 & 徐琳玲'
    # 深圳
    account2.loc[account2['端口'] == 'cny', '端口'] = 'cny-004'
    account2.loc[account2['端口'] == 'cny-004', '查賬財務郵箱地址'] = 'Automation-SZ財務查賬'
    account2.loc[account2['端口'] == 'cny-004', '付款方式'] = '搜索亞洲科技(深圳)有限公司'
    account2.loc[account2['AM'] == '鲁东栋', '操作'] = '顾凡凡'
    account2.loc[account2['AM'] == '陈宛欣', '操作'] = '顾凡凡'
    account2.loc[account2['AM'] == '赵宗州', '操作'] = '卢铭坛'
    account2.loc[account2['AM'] == '黄希腾', '操作'] = '陈熙香'
    account2.loc[account2['AM'] == '李裕玲', '操作'] = '卢铭坛'
# =============================================================================
#     account2[account2['端口'] != '-'].to_excel(r'C:\Users\chen.huaiyu\Desktop\Output\IO客户新申户' + str(round(time.clock(), 3))
#                      + '.xlsx', index=False, freeze_panes=(1, 0))
# =============================================================================
    
    import xlwings as xw
    
    wb = xw.Book(r'H:\SZ_数据\Input\IO系統母版-3.07.xlsm')
    wb.app.visible = True
    wb.app.screen_updating = False
    sht = wb.sheets['IO']
    Row = sht.range('A1').current_region.rows.count + 1
    sht.range('A' + str(Row)).value = account2[account2['端口'] != '-'].values
    sht.range('A' + str(Row)).color = (162, 163, 165)  # RGB
    wb.save()
    wb.app.screen_updating = True
    #wb.close()
    

if __name__ == '__main__':
    start = time.clock()
    # ready!
    test1 = input('开户进度总表处理完毕(account)？(Y/N)')
    test2 = input('新增账户【客户】地址对么？(ioSystem)？(Y/N):')
    test3 = input('input-icrm消费数据下载了么(icrm)？(Y/N):')
    test4 = input('AM & 操作是否变更？(Y/N):')
    if test1 == 'Y' and test2 == 'Y' and test3 == 'Y' and test4 == 'Y':
        print('请稍等......')
        print('注意：文件有外链，不更新；')
        mergeExcel()
        print('邮件中抄【客户】、标识已通过验证账户并检查，复制到IO系统客户信息表')
        print('邮件中抄【客户】、标识已通过验证账户并检查，复制到IO系统客户信息表')
    else:
        print('请补充相关信息。')
        pass
    print('\a\a执行完毕，耗时： {:.3f}Min'.format((time.clock() - start)/60))
