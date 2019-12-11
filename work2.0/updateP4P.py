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
from datetime import datetime, timedelta
from xlwings import constants
import xlwings as xw
import pandas as pd
import configparser
import functools
import time
import os


now = lambda : time.perf_counter()

def cost_time(func):
    @functools.wraps(func)
    def wrapper(*args):
        st = now()
        print('{}() start:'.format(func.__name__))
        func(*args)
        ed = now()
        print('Run time {} min'.format(round((st - ed)/60, 3)))
    return wrapper
    

class Sqlserver(object):
    
    def __init__(self, tablename):
        ''' 
        '''
        if isinstance(tablename, str):
            self.__tablename = tablename
    
    @staticmethod
    def connectDB():
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
    
    def querySpending(self, engine, category, date):
        ''' 表单查询: 种类、日期
        '''
        # 参数检查
        if (engine.execute('select 1') and 
            category in ['总点击', '搜索点击', '新产品', '自主投放', 
                         '无线搜索点击'] and isinstance(date, str)):
            # 获取表头
            sql = '''
                select * from information_schema.columns where table_name=%s
                '''
            col = [i[3] for i in engine.execute(sql, self.__tablename).fetchall()]
            col.append('金额')
            
            # 获取表内容
            if obj == '现金':
                sql = '''
                    select a.*, b.金额
                    from [Account Management].[dbo].basicInfo a
                    left join 
                    (select * from [Account Management].[dbo].现金
                    where 类别=%s and 日期=%s ) b
                    on a.用户名 = b.用户名
                    order by Id
                    '''
            else:
                sql = '''
                    select a.*, b.金额
                    from [Account Management].[dbo].basicInfo AS a
                    left join 
                    (select * from [Account Management].[dbo].消费
                    where 类别=%s and 日期=%s ) AS b
                    on a.用户名 = b.用户名
                    order by Id
                    '''
            data = engine.execute(sql, (category, date)).fetchall()
            print(len(data))
            # 返回DataFrame
            df = pd.DataFrame(data, columns=col)
            
            return df
    
    
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
        date_str = datetime.strptime(self._date, '%Y%m%d').strftime('%Y.%m..F')
        if obj == '现金':
            path = self.path + date_str + 'cash.xlsx'
            if os.path.exists(path):
                return path
            else:
                print('NotFoundFil: {}'.format(path))
        else:
            path = self.path + date_str + '.xlsx'
            if os.path.exists(path):
                return path
            else:
                print('NotFoundFil: {}'.format(path))
    
    def toExcel(self, df, sht):
        ''' 写入
        '''
        # 转换
        try:
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
            
            # 消费
            sht['A1'].offset(9, ll.index(date_date)).options(transpose=True
                                                             ).value = df_s.values
               
            if cntEnd > cntSta:
                # 填充汇总列
                sht[cntSta-1, 34:47].api.AutoFill(sht[cntSta-1:cntEnd, 34:47].api,
                   constants.AutoFillType.xlFillCopy)
                # 填充消费列
                sht[cntSta:cntEnd, 47:ll.index(date_date)].value = 0
                # 填充格式
                for i in range(7, 13):
                    sht[cntSta-30:cntEnd, :ll.index(date_date)+1].api.Borders(i).lineStyle = 1
        except Exception as e:
            print('toExcel产生异常：{}'.format(e))
            
    def dailyRatio(self, sht, dateStr):
        '''日环比；不计节假日
        '''
        try:
            date = datetime.strptime(dateStr, '%Y%m%d')
            if date.weekday() == 6:  # 周日(1)
                date1 = date - timedelta(2)
                date2 = date - timedelta(3)
            elif date.weekday() == 5:  # 周六(7)
                date1 = date - timedelta(1)
                date2 = date - timedelta(2)
            elif date.weekday() == 0:  # 周一(2)
                date1 = date
                date2 = date - timedelta(3)
            elif date.weekday() in [1, 2, 3, 4]:  # 周二、三、四、五(3,4,5,6)
                date1 = date
                date2 = date - timedelta(1)
            # 赋值
            sht[2, 2].value = '环比增长额\n%s日环比%s日' % (date1.day, date2.day)
            cnt1 = sht['A3:AJ3'].value.index(date1)
            cnt2 = sht['A3:AJ3'].value.index(date2)
            for i in range(3, 19):
                sht[i, 2].value = sht[i, cnt1].value - sht[i, cnt2].value  # 环比增长额
                if sht[i, cnt2].value == 0:
                    sht[i, 3].value = 0
                else:
                    sht[i, 3].value = sht[i, 2].value/sht[i, cnt2].value  # 环比增长率
        except Exception as e:
            print('5 日环比异常：%s', e)
        else:
            print('5 日环比')


def main(dateStr):
    # 实例化
    DB = Sqlserver('basicInfo')
    engine = DB.connectDB()
    ex = Excel(dateStr)
    # 准备excel表格
    wb = xw.Book(ex.excel_path())
    wb.app.calculation = 'manual'
    shtList = ['P4P消费', '搜索点击消费', '新产品消费（除原生广告）', 
               '原生广告', '无线搜索点击消费']
    lis = ['总点击', '搜索点击', '新产品', '自主投放', '无线搜索点击']
    # 获取数据
    if all_mob == 'all' or all_mob == '':
        pass
    elif all_mob == 'mob':
        shtList = [shtList[-1]]
        lis = [lis[-1]]
    else:
        print('all_mob需输入:all(全部) or mob(无线)')
        lis = []
        shtList = []
    for n, i in enumerate(lis):
        print(n, i)
        df = DB.querySpending(engine, i, dateStr)
        sht = wb.sheets[shtList[n]]
        # 写入
        ex.toExcel(df, sht)
    wb.app.calculate()
    # 每日消费走势
    if periods == 1:
        sht = wb.sheets['每日消费走势']
        ex.dailyRatio(sht, dateStr)
        # 保存
        wb.save()
        # wb.close()
        print('\a程序结束')
    

if __name__ == '__main__':
    
    # 日期锁定
    print('Note: upEx')
    star = time.perf_counter()
    start = input('输入起始日期(20190101)：')
    periods = eval(input('输入持续日期数(1,2,...,n)：'))
    obj = input('消费/现金?(默认消费)：')
    all_mob = input("mob/all？(默认all)")
    if obj == '':
        obj = '消费'
        print(obj)
    # 输入参数检查
    if datetime.strptime(start, '%Y%m%d') and isinstance(periods, int):
        print(start, periods)
        dateStr = pd.date_range(start=start, periods=periods)
        for i in dateStr:
            print(i)
            main(dateStr=i.strftime('%Y%m%d'))
            periods -= 1
        stop = time.perf_counter()
        print('\a耗时：%.3f min' % ((stop - star)/60))
    pass
