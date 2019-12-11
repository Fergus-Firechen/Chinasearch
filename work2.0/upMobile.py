# -*- coding: utf-8 -*-
"""
Created on Sat Nov  9 17:15:08 2019

- 更新无线
1.读取(参数检查：文件存在？)
2.处理(日期序列；结构转换；)
3.较对(转换前后总数)
4.连接DB
5.删库
6.上载
7.附加：装饰器提示

@author: chen.huaiyu
"""
import pandas as pd
import functools
import time
import os

now = lambda : time.perf_counter()

def log(text):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kw):
            print('Call {}(): {}'.format(func.__name__, text))
            return func(*args, **kw)
        return wrapper
    return decorator

def get_path(val, st, ed):
    path = r'c:\users\chen.huaiyu\downloads'
    if val == '现金':
        name = 'cash ' + st + '_' + ed + '.csv'
    else:
        name = 'p4p ' + st + '_' + ed + '.csv'
    if os.path.exists(os.path.join(path, name)):
        return os.path.join(path, name)
    else:
        raise NameError('NotFoundFil:{}'.format(name))

def read_fil(val, st, ed):
    return pd.read_csv(get_path(val, st, ed), encoding='GBK')

def handling(fil, st, ed):
    def get_date():
        return map(lambda x: x.strftime('%Y%m%d'), pd.date_range(st, ed))
    
    cols = ['日期', '用户名', '类别', '金额']
    global df
    df = pd.DataFrame(columns=cols)
    fil.rename(columns={'账户名称': '用户名'}, inplace=True)
    for dat in get_date():
        col = lambda x: '无线搜索点击消费' + dat
        dic2 = {col(dat): '金额'}
        df1 = fil.loc[fil[col(dat)] > 0, ['用户名', col(dat)]]
        df1.rename(columns=dic2, inplace=True)
        df1['日期'] = dat
        df1['类别'] = '无线搜索点击'
        df = df.append(df1, sort=False)
        check(dat, df1, df)
    return df
    
def check(dat, df1, df):
    su = df1['金额'].sum()
    piv = df.loc[df['日期'] == dat, '金额'].sum()
    if su == piv:
        pass
    else:
        raise ValueError('二维转一维数据转换出现错误: {}'.format(dat))

def connect_db():
    'connect DB'
    def login():
        import configparser
        CONF = r'c:\users\chen.huaiyu\chinasearch\c.s.conf'
        SQL = 'SQL Server'
        if os.path.exists(CONF):
            conf = configparser.ConfigParser()
            conf.read(CONF)
            ip = conf.get(SQL, 'ip')
            port = conf.get(SQL, 'port')
            db = conf.get(SQL, 'dbname')
            return ip, port, db
        else:
            raise
    ss = 'mssql+pymssql://@{}:{}/{}'.format(login()[0], login()[1], login()[2])
    try:
        from sqlalchemy import create_engine
        engine = create_engine(ss, echo=False)
    except Exception as e:
        print('Connect failed:{}'.format(e))
    else:
        print('Connect success.')
        return engine

def del_mob(val, st, ed):
    'del mobile data'
    sql = ''' exec todo.delete_mob {}, {}, {}
        '''
    return sql.format(val, st, ed)

@log('main')
def main():
    '主函数'
    print('upMobile,输入指定参数:')
    star = now()
    while True:
        val = input('表：消费/现金？')
        st = input('开始日期(格式:20190101)?')
        ed = input('结束日期(格式:20190101)?')
        if val and st and ed:
            try:
                df = handling(read_fil(val, st, ed), st, ed)
            except Exception as e:
                print(e)
            else:
                engine = connect_db()
                with engine.begin() as conn:
                    conn.execute(del_mob(val, st, ed))
                df.to_sql('消费', con=engine, if_exists='append', index=False)
            finally:
                print('Runtime {}s'.format(round(now() - star, 3)))
                break
        else:
            print('重新输入:')
            continue

if __name__ == '__main__':
    main()
