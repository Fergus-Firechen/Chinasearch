#!/usr/bin/python
# _*_ conding: utf-8 _*_

from sqlalchemy import create_engine
import os

def connect_db():
	def login():
		import configparser
		CONF = r'c:\users\chen.huaiyu\chinasearch\c.s.conf'
		conf = configparser.ConfigParser()
		if os.path.exists(CONF):
			conf.read(CONF)
			s = 'SQL Server'
			host = conf.get(s, 'ip')
			port = conf.get(s, 'port')
			db = conf.get(s, 'dbname')
			return host, port, db
	try:
		s = 'mssql+pymssql://@%s:%s/%s'
		engine = create_engine(s % login())
	except Exception as e:
		print('Failed connect: {}'.format(e))
	else:
		return engine





if __name__ == '__main__':

	engine = connect_db()
	print(engine.execute('select 1').fetchall())
