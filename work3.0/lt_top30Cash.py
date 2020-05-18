# _*_ coding: utf-8 _*_
'''
top30Cash

'''
import time
import pandas as pd
from db import getUrl
from sqlalchemy import create_engine
from datetime import date, timedelta

PATH = r'H:\SZ_数据\Python\c.s.conf'
URL = getUrl('SQL Server', 'acc', 'pw', 'ip', 'port', 'db')
ST_DAT, ED_DAT = eval(input('输入：QTD起始日,终止日；如(20200401,20200430):'))

# 时间
now = lambda: time.perf_counter()
dat = lambda n: (date.today() - timedelta(n)).strftime('%Y%m%d')
mDat = lambda n: (date.today() - timedelta(n)).strftime('%m.%d')
qDat = lambda n: (int((date.today() - timedelta(n)).strftime('%m'))-1)//3+1
yDat = lambda n: (date.today() - timedelta(n)).strftime('%Y')

def getData(sql, url):
    with create_engine(URL).begin() as conn:
        return map(lambda x: list(x), conn.execute(sql).fetchall())

def week(data):
    # 周粒度消费
    df = pd.DataFrame(list(data), columns=('用户名', '类别', '金额', '周'))
    df = df.pivot_table(values=['金额'], index=['用户名']
                        , columns=['类别', '周'])
    return df

def getP4P(df):
    # 上上、上、本
    datLis = df.columns.get_level_values(2).unique()
    # 产品大类
    clsLis = tuple(set(df.columns.get_level_values(1)))
    for d in datLis:
        df[('金额', 'P4P', d)] = (df[('金额', clsLis[0], d)]
                                    + df[('金额',clsLis[1], d)]
                                    + df[('金额', clsLis[2], d)])
    # 索引
    df.columns = [c + dat for c in df.columns.get_level_values(1).unique()
                  for dat in datLis]
    df.reset_index(inplace=True)

def merge(basicInfo, qtd, data):
    # basicInfo & qtd
    col1 = ['用户名', '广告主', '二级行业', '区域']
    col2 = ['用户名', 'QTD']
    df = pd.merge(pd.DataFrame(list(basicInfo), columns=col1)
                , pd.DataFrame(list(qtd), columns=col2)
                , how='left', on='用户名')
    # region -> hk
    df['区域'] = df['区域'].str.replace(r'^HK.+', 'HK')
    # week
    df = pd.merge(df, data, how='left', on='用户名')
    df.fillna(0, inplace=True)
    return df
    
def rank(df):
    # 上周排名
    df['sum'] = df['P4P上上周'] + df['P4P上周']
    df.sort_values('sum', inplace=True, ascending=False)
    df.reset_index(drop=True, inplace=True)
    df.index = [i+1 for i in df.index]
    df.index.name = '上周排名'
    df.reset_index(inplace=True)
    # 本周排名
    df['sum'] = df['P4P上周'] + df['P4P本周']
    df.sort_values('sum', inplace=True, ascending=False)
    df.reset_index(drop=True, inplace=True)
    df.index = [i+1 for i in df.index]
    df.index.name = '本周排名'
    df.drop(columns=['sum'] + list(filter(lambda x: '上上' in x, df.columns))
            , inplace=True)

def ringRatio(df):
    df['环比增长'] = df['P4P本周'] - df['P4P上周']
    df['环比增长率'] = df['环比增长'] / df['P4P上周']

def fmt(df):
    # Output
    path = r'H:\SZ_数据\Download\Top30Cash(' + mDat(14
                                                  ) +'_' + mDat(1) + ').xlsx'
    with pd.ExcelWriter(path) as writer:
        df[:30].to_excel(writer, startrow=2, freeze_panes=(3,0))
    # 修改表
    import xlwings as xw
    wb = xw.Book(path)
    sht = wb.sheets[0]
    cntRow = sht['A3'].current_region.rows.count
    cntCol = sht['A3'].current_region.columns.count
    # 标签
    sht[0, 0].value = 'P4P'
    for n, v in enumerate(sht[2, :cntCol].value):
        if '排名' in v:
            sht[1, n].clear()
        elif '上' in v:
            sht[1, n].value = mDat(14) + '-' + mDat(7)
        elif '本' in v:
            sht[1, n].value = mDat(7) + '-' + mDat(1)
        elif 'QTD' in v:
            sht[1, n].value = yDat(n) + 'Q' + str(qDat(n)) + '现金'
        elif '环比' in v:
            sht[1, n].value = 'P4P现金'
    # 边框
    for b in range(7, 13):
        sht[1:cntRow, :cntCol].current_region.api.Borders(b).weight = 2
    # 列宽
    sht[:, :cntCol].autofit()
    # 加粗
    sht[:3, :cntCol].api.Font.Bold = True
    # 数字格式
    sht[3:, :cntCol-1].api.NumberFormat = '#,##0'
    sht[3:, cntCol-1].api.NumberFormat = '0.0%'
    #
    wb.save()
    wb.close()
    return path
    
def main():
    # basicInfo
    sql1 = "SELECT 用户名, 广告主, 信誉成长值, 区域 FROM basicInfo"
    # QTD
    sql2 = "SELECT * FROM getCashSUM('%s', '%s')" % (ST_DAT, ED_DAT)
    # Nearly Three weeks Spending
    sql3 = '''SELECT * FROM getThrWeekCash(%s)''' % dat(21)
    # 
    basicInfo, qtd, data = (getData(sql1, URL), getData(sql2, URL)
                            , getData(sql3, URL))
    # week
    w = week(data)
    # calculate P4P
    getP4P(w)
    # merge
    df = merge(basicInfo, qtd, w)
    # groupby
    df = df.groupby(['广告主', '二级行业', '区域']).sum()
    df.reset_index(inplace=True)
    # rank
    rank(df)
    # ring ratio
    ringRatio(df)
    # fmt
    path = fmt(df)
    # 发送
    from sendEmail import sendEmail
    sendEmail('Top 30广告主现金', '    见附件。', [path])
    
if __name__ == '__main__':
    st = now()
    main()
    print('Runtime: {:.3f} min'.format((now()-st)/60))
    