#!/usr/bin/python
# -*- coding:utf-8 -*-

from sqlalchemy import create_engine
import configparser
import os
import pandas as pd
import matplotlib.pyplot as plt

def connectDB(args):
    def login(args):
        CONF = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        if os.path.exists(CONF):
            conf.read(CONF)
            accountName = conf.get(args, 'accountname')
            pw = conf.get(args, 'password')
            host = conf.get(args, 'ip')
            port = conf.get(args, 'port')
            dbname = conf.get(args, 'dbname')
            return accountName, pw, host, port, dbname
        else:
            print('Output配置文件不存在!')
            
    s = "mssql+pymssql://%s:%s@%s:%s/%s"
    engine = create_engine(s % login(args))
    if engine.execute('select 1'):
        print('连接成功')
    else:
        raise
    return engine

def box(item, ax):
    plt.figure(figsize=(8,10))
    p = df1.loc[df1[item] > 0, item].plot(kind='box', return_type='dict', ax=ax)
    x = p['fliers'][0].get_xdata()
    print(x)
    y = p['fliers'][0].get_ydata()
    y.sort()
    print(y)
    for i in range(len(x)): 
      if i > 0:
        plt.annotate(y[i], xy = (x[i],y[i]), xytext=(x[i]+0.05, y[i]))
      else:
        plt.annotate(y[i], xy = (x[i],y[i]), xytext=(x[i]+0.08, y[i]))
    plt.title(item)

def hist(item, ax):
    df1.loc[df1[item] > 0, item].plot(kind='hist', title=item, bins=50, ax=ax)


if __name__ == '__main__':
    
    engine = connectDB('SQL Server')
    
    sql1 = ''' select * from 消费
                where 日期='20191128'
        '''
    sql2 = ''' select * from information_schema.columns
                where table_name='消费'
        '''
    col = [i[3] for i in engine.execute(sql2).fetchall()]
    data = engine.execute(sql1).fetchall()
    
    df = pd.DataFrame(data, columns=col)
    print(df.shape)
    
    df.groupby(by='类别').sum()
    df1 = df.pivot_table(values='金额', index='用户名', columns='类别', 
                         fill_value=0, aggfunc=sum)
    plt.rcParams['font.family'] = ['Microsoft YaHei']
    
    # 组图
    df1.plot(kind='box', subplots=True, figsize=(24,12))
    
    fig, ((ax1, ax2, ax3), (ax4, ax5, ax6)) = plt.subplots(2, 3)
    box('KA', ax1)
    box('搜索点击', ax2)
    box('新产品', ax3)
    box('自主投放', ax4)
    box('百通', ax5)
    
    
    import matplotlib.gridspec as gds
    plt.figure()
    G = gds.GridSpec(2, 2)
    ax1 = plt.subplot(G[0, 0])
    ax2 = plt.subplot(G[0, 1])
    ax3 = plt.subplot(G[1, 0])
    ax4 = plt.subplot(G[1, 1])

    hist('KA', ax1)
    hist('搜索点击', ax3)
    hist('新产品', ax4)
    hist('自主投放', ax2)
    plt.show()
    
    
    
    
    
    