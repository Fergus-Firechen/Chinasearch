# -*- coding: utf-8 -*-
'''
- 打开文件
- 获取信息
- 填充
	- 基本信息填充
	- 当有新户增加时，向下填充所有后续行
	- 消费填充

'''

import os
import time
import configparser
import pandas as pd
import xlwings as xw
from xlwings import constants
from datetime import datetime, timedelta
from sqlalchemy import create_engine


now = lambda : time.perf_counter()

class GetPath():
	def __init__(self, n_days_ago):
		self.n_days_ago = n_days_ago
		self.__PATH_1 = r'H:\SZ_数据\Input'
		self.__PATH_2 = r'C:\users\chen.huaiyu\downloads'

	def datetime_ago(self):
		return datetime.now() - timedelta(self.n_days_ago)

	def p4p_spending_report(self):
		name = 'P4P 消费报告' + self.datetime_ago().strftime('%Y.%m.') + '.F.xlsx'
		if os.path.exists(os.path.join(self.__PATH_1, name)):
			return os.path.join(self.__PATH_1, name)
		else:
			raise IOError('NotFoundFile:{}'.format(name))

	def get_q(self):
		m = self.datetime_ago().strftime('%m')
		if m in (1, 2, 3):
			return 'Q1'
		elif m in (4, 5, 6):
			return 'Q2'
		elif m in (7, 8, 9):
			return 'Q3'
		else:
			return 'Q4'

	def get_q_mon(self):
		if self.get_q() == 'Q1':
			return 'Jan to Mar'
		elif self.get_q() == 'Q2':
			return 'Apr to Jun'
		elif self.get_q() == 'Q3':
			return 'Jul to Sep'
		else:
			return 'Oct to Dec'

	def ave(self):
		path = ('Ave.workday&weekday' + self.get_q() + '(' + 
				str(self.datetime_ago().year) + ' ' + self.get_q_mon() + ')' 
				+ self.datetime_ago().strftime('%Y.%m.%d') + '.xlsx')

		if os.path.exists(os.path.join(self.__PATH_2, path)):
			return os.path.join(self.__PATH_2, path)
		else:
			raise IOError('NotFoundFile: {}'.format(path))


class SqlServer():
	def __init__(self, db):
		print('SqlServer._init__')
		self.__PATH = r'C:\users\chen.huaiyu\Chinasearch\c.s.conf'
		self.db = db

	def __login(self):
		conf = configparser.ConfigParser()
		if os.path.exists(self.__PATH):
			conf.read(self.__PATH)
			sa = conf.get(self.db, 'accountname')
			pw = conf.get(self.db, 'password')
			host = conf.get(self.db, 'ip')
			port = conf.get(self.db, 'port')
			dbname = conf.get(self.db, 'dbname')
			return sa, pw, host, port, dbname
		else:
			raise IOError('NotFoundFile:{}'.format(os.split(self.__PATH)[-1]))

	def connection(self):
		s = 'mssql+pymssql://%s:%s@%s:%s/%s'
		try:
			engine = create_engine(s % self.__login())
		except:
			raise
		else:
			return engine

	@staticmethod
	def col(tableName):
		'获取列字段'
		if isinstance(tableName, str):
			sql = '''select * from information_schema.columns 
                where table_name='{}'
				'''.format(tableName)
			return sql
		else:
			raise ValueError("Add '' ")

	@staticmethod
	def data_basicInfo():
		'读取基本信息'
		sql = ''' select * from basicInfo
			'''
		return sql

	@staticmethod
	def data_spending(date1, date2):
		'获取消费数据'
		if isinstance(date1, str) and isinstance(date2, str):
			sql = ''' select * from 消费 where 日期 between '{}' and '{}'
				'''.format(date1, date2)
			return sql

def handling_data(product):
    # 分别筛选出P4P、NP、INF，然后pivotz
    sp = spending[spending['类别'] == product]
    sp_pivot = pd.pivot_table(sp, index='用户名', columns='日期', values='金额')
    merge_basicInfo_sp = pd.merge(basicInfo, sp_pivot, on='用户名', how='left')
    merge_basicInfo_sp.fillna(0, inplace=True)
    return merge_basicInfo_sp

def write_to_excel(shtName):
    getPath = GetPath(eval(input('前1天？输入1')))  # 默认一天前
    wb = xw.Book(getPath.ave())
    sht = wb.sheets[shtName]
    # 基本信息
    basicInfo.drop(columns=['Id'], axis=1, inplace=True)
    sht['A3'].options(expand='table').value = basicInfo.fillna('-').values

    # 消费
    #据日期定位 - 填充 - 保存
    # 定位
    num_col = sht['A2'].current_region.columns.count
    print(sht[1, :num_col])

    cols = sht[1, :num_col]
    lis_cols_value = cols.value
    print(lis_cols_value)
    
    date_first = getPath.datetime_ago()
    date_first = date_first.replace(date_first.year, date_first.month, 1, 
                       0, 0, 0, 0)
    print(date_first)
    
    date_xy = lis_cols_value.index(date_first)
    print(date_xy)
    print(sht[1, date_xy])

    # 填充
    sht[2, date_xy].options(expand='table').value = basicInfo_and_sp.loc[:, date_first.strftime('%Y%m%d'):].values


if __name__ == '__main__':
    st = now()
    SQL = SqlServer('SQL Server')
    ENGINE = SQL.connection()
    
    with ENGINE.begin() as connection:
		# basicInfo
        data = connection.execute(SQL.data_basicInfo()).fetchall()
        col = [i[3] for i in connection.execute(SQL.col('basicInfo'))]
        basicInfo = pd.DataFrame(data, columns=col)
        
        # data_spending
        data = list(map(list, connection.execute(SQL.data_spending('20191001', '20191005'))))
        col = [i[3] for i in connection.execute(SQL.col('消费'))]
        spending = pd.DataFrame(data, columns=col)
        
        print(basicInfo.shape, spending.shape)

    # 分别筛选P4P、NP、INF，然后pivot
    for product in ['搜索点击', '新产品', '自主投放']:
        basicInfo_and_sp = handling_data(product)
        print(basicInfo_and_sp.head(3))

        # 写入excel
        if product == '搜索点击':
        	write_to_excel('搜索')
        # elif product == '新产品':
        # 	write_to_excel('其他新产品')
        # elif product == '自主投放':
        # 	write_to_excel('原生广告')
        else:
        	print('NotFoundSheet.')
    
    print(round((now() - st), 3))
    

