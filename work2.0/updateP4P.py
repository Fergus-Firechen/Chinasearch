# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 17:16:26 2019

# 连接DB
# 
# 获取basicInfo，消费
# 写入excel
# 填充
# 无线待完善

@author: chen.huaiyu
"""

from sqlalchemy import create_engine
from datetime import datetime
from xlwings import constants
#import win32com.client as win32
import xlwings as xw
import pandas as pd
import configparser
import os, time


class Sqlserver(object):
    
    def __init__(self, tablename):
        ''' 
        '''
        if isinstance(tablename, str):
            self.__tablename = tablename
    
    @staticmethod
    def __connectDB():
        ''' 连接SQL Server
        '''
        def loginAccount():
            CONF = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
            conf = configparser.ConfigParser()
            if os.path.exists(CONF):
                conf.read(CONF)
                host = conf.get('SQL Server', 'ip')
                port = conf.get('SQL Server', 'port')
                dbname = conf.get('SQL Server', 'dbname')
                return host, port, dbname
        try:
            engine = create_engine(
                    "mssql+pymssql://@%s:%s/%s" % loginAccount())
        except:
            print('SQL Server 连接失败')
        else:
            print('SQL Server 连接成功')
            return engine
    
    def querySpending(self, category, date):
        ''' 表单查询: 种类、日期
        '''
        EN = self.__connectDB()
        # 参数检查
        if (EN.execute('select 1') and 
            category in ['总点击', '搜索点击', '新产品', '自主投放'] and 
            isinstance(date, str)):
            # 获取表头
            sql = '''
            select * from information_schema.columns where table_name=%s
            '''
            col = [i[3] for i in EN.execute(sql, self.__tablename).fetchall()]
            col.append('金额')
            
            # 获取表内容
            sql = '''
            select a.*, b.金额
            from [Account Management].[dbo].basicInfo a
            left join 
            (select * from [Account Management].[dbo].消费 
            where 类别=%s and 日期=%s ) b
            on a.用户名 = b.用户名
            order by Id
            '''
            data = EN.execute(sql, (category, date)).fetchall()
            
            # 返回DataFrame
            return pd.DataFrame(data, columns=col)
    
    
class Excel(object):
    
    def __init__(self, date):
        ''' 初始化
        '''
        if isinstance(date, str):
            self._date = date
        self.path = r'H:\SZ_数据\Input\P4P 消费报告'
    
    def excel_path(self):
        ''' Excel路径生成
        '''
        date_str = datetime.strptime(self._date, '%Y%m%d').strftime('%Y.%m..')
        return self.path + date_str + '.xlsx'
    
    def toExcel(self, df, sht):
        ''' 写入
        '''
        # 转换
        df_b = df.loc[:, '属性':'加V缴费到期日']  # 基本信息
        df_b.fillna(value='-', inplace=True)
        df_s = df.loc[:, '金额']  # 消费
        df_s.fillna(value=0, inplace=True)
        
        # 基本信息
        cntSta = sht['A1'].current_region.rows.count  # 行数：更新前
        sht[cntSta, 0].color = (255, 255, 0)  # 新增标识
        sht['A10'].value = df_b.values  # 更新
        cntEnd = sht['A1'].current_region.rows.count  # 行数：更新后
        
        # 填充
        date_date = datetime.strptime(self._date, '%Y%m%d')
        ll = [i.value for i in sht['A1'].current_region.rows[:1]][0]  # 字段
        sht['A1:A'+str(cntEnd)].rows.autofit()
        if cntEnd > cntSta:
            # 汇总列填充
            sht[cntSta-1, 34:47].api.AutoFill(sht[cntSta-1:cntEnd, 34:47].api,
               constants.AutoFillType.xlFillCopy)
            # 消费列填充
            sht[cntSta:cntEnd, 47:ll.index(date_date)].value = 0
        # 消费
        sht['A1'].offset(9, ll.index(date_date)).options(transpose=True
                                                         ).value = df_s.values
        # 格式
        for i in range(7, 13):
            sht[cntSta:cntEnd, :ll.index(date_date)+1].api.Borders(i).lineStyle = 1
        
        

def main(dateStr):
    # 实例化
    DB = Sqlserver('basicInfo')
    ex = Excel(dateStr)
    
    # 打开文件
    wb = xw.Book(ex.excel_path())
    shtList = ['P4P消费', '搜索点击消费', '新产品消费（除原生广告）', 
               '原生广告', '无线搜索点击消费']
    
    # 获取数据
    for n, i in enumerate(['总点击', '搜索点击', '新产品', '自主投放', '无线搜索点击']):
        if n in [0, 1, 2, 3]:
            continue
        global df
        print(n, i)
        df = DB.querySpending(i, dateStr)
# =============================================================================
#         sht = wb.sheets[shtList[n]]
# =============================================================================
        
# =============================================================================
#         # 写入
#         ex.toExcel(df, sht)
# =============================================================================
        
    # 保存
# =============================================================================
#     wb.save()
# =============================================================================
    # wb.close()
    print('\a\a程序结束')
    
if __name__ == '__main__':
    
    # 日期锁定
    star = time.perf_counter()
    dateStr = pd.date_range(start='20190604', periods=1)
    for i in dateStr:
        main(dateStr=i.strftime('%Y%m%d'))
    stop = time.perf_counter()
    print('耗时：%.3f s' % (stop - star))
    
    
    