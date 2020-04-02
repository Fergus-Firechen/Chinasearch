# _*_ coding: utf-8 _*_

import os
import time
import pandas as pd
from sqlalchemy import create_engine
from configparser import ConfigParser


now = lambda : time.perf_counter()

def getAccountInfo(section):
	path = r'c:\users\chen.huaiyu\chinasearch\c.s.conf'
	conf = ConfigParser()
	if os.path.isfile(path):
		conf.read(path)
		acc = conf.get(section, 'accountname')
		pw = conf.get(section, 'password')
		ip = conf.get(section, 'ip')
		port = conf.get(section, 'port')
		db = conf.get(section, 'dbname')
		return acc, pw, ip, port, db

def connect():
	ss = "mssql+pymssql://%s:%s@%s:%s/%s"
	try:
		engine = create_engine(ss % getAccountInfo('SQL Server'))
	except Exception as e:
		print("连接失败")
		raise
	else:
		print("连接成功")
		return engine

def main():
	# 获取变更信息
	# 文件地址
	addr = r'H:\SZ_数据\Input\端口转移.xlsx'
	# 读取数据
	df1 = pd.read_excel(addr, sheet_name='客户')  # 客户
	df2 = pd.read_excel(addr, sheet_name='AM')  # AM

	# 更新 客户
	sql1 = "update basicInfo set 客户='{}' where 用户名='{}'"
	# 更新 AM
	sql2 = "update basicInfo set AM='{}' where AM='{}'"
	#
	with connect().begin() as conn:
		if len(df1) != 0:
			for user, master in df1.values:
				conn.execute(sql.format(master, user))
		if len(df2) != 0:
			for am, new_am in df2.values:
				conn.execute(sql.format(new_am, am))


if __name__ == '__main__':
	st = now()
	main()
	print("Runtime: {} min".format(round((now() - st)/60, 3)))