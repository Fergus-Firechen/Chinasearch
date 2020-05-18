# _*_ coding: utf-8 _*_
'''
- 历史数据上载
- Q
# pd.DataFrame([0, 1], columns=['a', 'b']
>>> pd.DataFrame([[0, 1]], columns=['a', 'b'])
- pd.ExcelWriter
with pd.ExcelWriter(os.path.join(PATH, 't' + dat + '.xlsx')) as writer:
    df_2.to_excel(writer, sheet_name='sht1')
- with engine.begin() as conn:
    conn.execute(sql)
/ 直至数据上载完后，断开才commit?

'''
import os
import time
import pandas as pd
from getPath import Doc
from getConfig import Conf
from sqlalchemy import create_engine


PATH = r'H:\SZ_数据\Download\15-18年 ICRM数据'
CAT = ('总点击', '搜索点击', '无线搜索点击', '自主投放', '新产品')
HEADER = ('日期', '用户名', '类别', '金额')

now = lambda: time.perf_counter()

def sp(key):
    # 获取文件
    doc = Doc(PATH)
    os.chdir(PATH)
    for f in doc.getSome(key):
        # 读取
        df = pd.read_csv(f, engine='python', encoding='GBK')
        df.rename(columns={'账户名称': '用户名'}, inplace=True)
        df.set_index('用户名', inplace=True, drop=True)
        # 获取时段
        with create_engine(getUrl()).begin() as conn:
            for dat in getDat(f):
                # 构造字段
                lis = getHeader(key, dat)
                # 筛选，去0
                df1 = df.loc[df[lis[0]] > 0, lis[:-1]]
                # 计算新产品
                df1[lis[-1]] = df1[lis[0]] - df1[lis[1]] - df1[lis[3]]
                # 转换为一维
                newDf = pd.DataFrame(columns=HEADER)
                newDf = toDimension(df1, newDf, dat, lis)
                # 上载
                print(dat, newDf.head(2))
                newDf.to_sql(key, con=conn, if_exists='append', index=False)
    
def getDat(f):
    # 由时期直接获取时段 如 '15Q1' --> '20150101' -- '20150331'
    # 日期时段
    #
    def QtoDateRange(Q):
        if 'Q1' in Q:
            return ['20' + Q[:2] + i for i in ('0101', '0331')]
        elif 'Q2' in Q:
            return ['20' + Q[:2] + i for i in ('0401', '0630')]
        elif 'Q3' in Q:
            return ['20' + Q[:2] + i for i in ('0701', '0930')]
        elif 'Q4' in Q:
            return ['20' + Q[:2] + i for i in ('1001', '1231')]
    
    st, ed = QtoDateRange(f.split('年')[0] + f.split('年')[1].split('消费')[0])
    return map(lambda x: x.strftime('%Y%m%d'), pd.date_range(st, ed))

def getHeader(key, dat):
    return [cat + key + dat for cat in CAT]

def toDimension(df, newDf, dat, lis):
    for acc in df.index:
        for n, cat in enumerate(CAT):
            df_1 = pd.DataFrame([[dat, acc, cat, df.loc[acc, lis[n]]]]
                                , columns=HEADER)
            newDf = newDf.append(df_1, sort=False)
    return newDf

def bt():
    pass

def getUrl():
    # 连接数据库
    conf = Conf(r'H:\SZ_数据\Python\c.s.conf')
    url = ('mssql+pymssql://%s:%s@%s:%s/%s' %
            conf.getInfo('SQL Server', 'accountname', 'password'
                         , 'ip', 'port', 'dbname'))
    return url

if __name__ == '__main__':
    st = now()
    sp('消费')
    sp('现金')
    print('Runtime: {:.3f} min'.format(((now() - st)/60)))