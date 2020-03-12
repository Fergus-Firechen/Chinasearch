#!/usr/bin/python
#_*_ coding:utf-8 _*_

'''
- 更新基本信息 & 消费
直接从SQL Server中读取数据，然后写入Avg表中

'''

from datetime import date, timedelta
import xlwings as xw
import time
import os

now = lambda : time.perf_counter()
dat = lambda : date.today() - timedelta(x)

def getQ(ed):
	m = dat(ed).month
	if m in (1, 2, 3):
		return "Q1"
	elif m in (4, 5, 6):
		return "Q2"
	elif m in (7, 8, 9):
		return "Q3"
	elif m in (10, 11, 12):
		return "Q4"

def getPath():
	pass

def main():
	path = getPath()

if __name__ == "__main__":
	st = now()

	main()
	
	tt = now() - st
	print("Runtime all: {} min {} s".format(tt//60, tt%60))
