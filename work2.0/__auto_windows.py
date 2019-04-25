# -*- coding: utf-8 -*-
"""
Created on Tue Sep 18 19:08:51 2018

- Windows 计划任务运行代码

@author: chen.huaiyu
"""

'自动抓取邮件获取：开户申请表'
from __mail_AccountApplicationTable import mainKH
from sqlalchemy import create_engine
from datetime import datetime
import pandas as pd
#import numpy as np


columns = ['日期', '合同原件是否已回', '是否赠送服务费', '是否开票', 
           '推广性质', '销售', '客服', '用户名', '端口', '行业',
           '渠道', '广告主总部', '资质归属地', '预估月消费', 
           'Region', '账期/预付', '服务费', '年费', '服务费币种', 
           '网站名称', '广告主名称', '广告主_简体', 'URL', '登记证编号', '生效日',
           '届满日期', '联系人', '电话', '客户', 'flag']
engine = create_engine('mssql+pyodbc://SQL Server')
'check'
df = pd.DataFrame(engine.execute('select * from 开户申请表 order by 日期 desc'
                                     ).fetchmany(2), columns=columns)
df.to_csv(r'c:\users\chen.huaiyu\Desktop\df.csv', encoding='GBK')


def func():
    '据数据库中最近日期判定抓取日期'
    date = engine.execute('''select top 1 日期 from 开户申请表 
                          order by 日期 desc'''
                          ).fetchone()
    n = (datetime.today() - date[0]).days
    return n


##1. 开户申请表
path = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
mainKH(func(), 1, path)

##2. '邮件据主题抓取：io system'
from __mail_io import mainIO
mainIO(func(), path)

