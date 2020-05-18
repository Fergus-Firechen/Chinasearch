# -*- coding: utf-8 -*-
"""
Created on Wed Sep 26 15:19:14 2018

1.已上载账户,旧户  来自excel  --sql.table:iosystem
2.sql构造表结构 + 表信息
3.新增户输出至结果文件
@author: chen.huaiyu
"""
import os
import time
import configparser
import pandas as pd

now = lambda : time.perf_counter()

def get_path():
    path = r'C:\users\chen.huaiyu\desktop'
    name = 'io系統母版-3.07.xlsm'
    if os.path.exists(os.path.join(path, name)):
        return os.path.join(path, name)

def df():
    return pd.read_excel(get_path(), sheet_name='IO')

def connect_db():
    def login(sql='SQL Server'):
        PATH = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        if os.path.exists(PATH):
            conf.read(PATH)
            host = conf.get(sql, 'ip')
            port = conf.get(sql, 'port')
            db = conf.get(sql, 'dbname')
            return host, port, db
        else:
        	raise Exception('NotFoundFil:{}'.format(PATH))
    
    from sqlalchemy import create_engine
    ss = 'mssql+pymssql://@%s:%s/%s'
    try:
        engine = create_engine(ss % login())
    except Exception as e:
        engine = create_engine(ss % login('Note'))
    else:
        return engine
    
def sql():
    sql = ''' SELECT * FROM todo.v_iosys
        '''
    return sql

def col():
    sql = ''' SELECT * FROM INFORMATION_SCHEMA.columns where TABLE_NAME = 'v_iosys'
        '''
    return sql

def up_p4p():
    # 获取文件 “消费报告 ...”
    # 倒序排列，获取最新的文件
    # 
    PATH = r'H:\SZ_数据\Download'
    lis = [i for i in os.listdir(PATH) if '消费报告 ' in i and '.csv' in i]
    if len(lis) > 0:
        lis.sort(key=lambda x: os.stat(os.path.join(PATH, x)).st_mtime, reverse=True)
        print(lis)
        df = pd.read_csv(os.path.join(PATH, lis[0]), encoding='GBK', engine='python')
        print(df.shape, df.head())
        df.to_sql('icrm_p4p', con=connect_db(), if_exists='replace', index=False)
    else:
        print('NotFoundFil: 消费报告')        

def to_ex_path():
    from datetime import date, timedelta
    PATH = r'D:\陈怀玉\工作\周工作\IO系统客户信息'
    name = ('IO系统客户信息(' 
            + (date.today() - timedelta(7)).strftime('%m.%d')
            + '-' + date.today().strftime('%m.%d') + ').xlsx')
    return os.path.join(PATH, name), name

def update_user_name(df, conn):
    for i in df['用户名']:
        conn.execute("insert into IOSystem values('{}')".format(i))

def send_mail():
    import smtplib
    from email import encoders
    from email.header import Header
    from email.mime.text import MIMEText
    from email.utils import parseaddr, formataddr
    from email.mime.multipart import MIMEMultipart, MIMEBase
    
    def _format_addr(s):
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode(), addr))
    
    def login():
        PATH = r'c:\users\chen.huaiyu\chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        if os.path.exists(PATH):
            conf.read(PATH)
            smt = conf.get('mail_baidu', 'sender server')
            from_addr = conf.get('mail_baidu', 'email')
            pw = conf.get('mail_baidu', 'password')
            to_addr = conf.get('to_addr', 'newIOSys')
            return smt, from_addr, pw, to_addr
    
    msg = MIMEMultipart()
    msg['From'] = login()[1]
    msg['To'] = login()[3]
    msg['Subject'] = Header('IO系统客户信息', 'utf-8').encode()

    msg.attach(MIMEText(''' Dear all,

        附件为新增IO系统客户信息，请查阅。
如有任何疑问，可随时和我联系，谢谢。

祝好

陈怀玉 | Fergus
Data | Shenzhen
百度HI：astfire
地址：深圳市罗湖区南东路5002号信兴广场主楼地王大厦5903-06室 ''',
            'plain', 'utf-8'))
    
    with open(to_ex_path()[0], 'rb') as f:
        mime = MIMEBase('text', 'plain', filename=to_ex_path()[1])
        # 加上必要的头信息
        mime.add_header('Content-Disposition', 'attachment', filename=to_ex_path()[1])  # 1.内容传输编码;2.;3.文件名
        mime.add_header('Content-ID', '<0>')
        mime.add_header('X-Attachment-Id', '0')
        # add 附件
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        msg.attach(mime)
    
    server = smtplib.SMTP(login()[0], 25)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.set_debuglevel(1)
    server.login(login()[1], login()[2])
    try:
        server.sendmail(login()[1], login()[3].split(','), msg.as_string())
    except Exception as e:
        print('Failed send: {}'.format(e))
    else:
        print('Success send.')
    server.quit()
    

def main():
    ''' main '''
    try:
        st = now()
        # 准备ioSystem表，已上载文件
# =============================================================================
#     df()['用户名'].to_sql('IOSystem', con=connect_db(), 
#                if_exists='replace', index=False)
#     print('Runtime: {0:>10.3f}'.format(now() - st))
# =============================================================================

        '上载已完成'
        
        up_p4p()
        
        # sql： 完成结构创建、数据筛选 & 拼接
        # 补充信息完整
        # 1. sql直接计算、查询personInfo补充
        # 2. icrm补充：URL & 开户日期
        # 2.1 加载最新icrm：消费表
        #
        with connect_db().begin() as conn:
            columns = [i[3] for i in conn.execute(col()).fetchall()]
            df = pd.DataFrame(conn.execute(sql()).fetchall(), columns=columns)
        df.to_excel(to_ex_path()[0])
    except Exception as e:
        print('Failed execute: {}'.format(e))
    else:
        # 更新IOSystem
        with connect_db().begin() as conn:
            update_user_name(df, conn)
        send_mail()
    finally:
        print('Runtime: {:.3f}s.程序结束'.format(now() - st))

if __name__ == '__main__':
    
    main()
    #pass
    

