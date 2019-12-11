# -*- coding: utf-8 -*-
"""
Created on Tue Sep 18 19:08:51 2018

- Windows 计划任务运行代码

@author: chen.huaiyu
"""

'自动抓取邮件获取：开户申请表'
from __mail_AccountApplicationTable import mainKH
from sqlalchemy import create_engine
import pandas as pd
import logging.config


# 账号密码 配置文件地址
path = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'

# 日志
PATH = r'C:\Users\chen.huaiyu\Chinasearch\logging.conf'
logging.config.fileConfig(PATH)
logger = logging.getLogger('chinaSearch')

# 连接SQL Server
engine = create_engine(r'mssql+pyodbc://SQL Server')
if engine.execute('select 1'):
    logger.info('SQL Server 连接正常')
    # 初始化
    # 人员信息表表头字段
    columns = [i[3] for i in engine.execute(
            "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='开户申请表'"
            ).fetchall()]

    #'据数据库中最近日期判定抓取日期'
    date_0 = engine.execute('''select top 1 日期 from 开户申请表 
                            ORDER BY 日期 DESC'''
                              ).fetchone()[0]
    
    # 删除DB中标识项
    engine.execute("DELETE FROM 开户申请表 WHERE 用户名 = '0.0'")
    df = pd.DataFrame(engine.execute('select * from 开户申请表 order by 日期'
                                         ).fetchall(), columns=columns)

    # 主程序
    mainKH(date_0, 1, path, )
else:
    logger.warning('SQL Server 连接失败')