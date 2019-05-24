# -*- coding: utf-8 -*-
"""
Created on Tue Mar 19 20:10:33 2019

# 新消户不存于端口表中
# DB取代文件“数据拆解表”
# df_kh 字段名统一？

@author: chen.huaiyu
"""

import time, functools, os
import pandas as pd
#import xlwings as xw
from sqlalchemy import create_engine
from datetime import datetime, timedelta

# 连接 DB
engine = create_engine(r'mssql+pyodbc://@SQL Server')
print(engine.execute('select 1'), '\nSQL Server 连接正常')

def path_date_str(n=1):
    '消费文件地址构造，默认昨日'
    yes_str = datetime.strftime(datetime.today() - timedelta(n), "%Y%m%d")
    path = os.chdir(r"c:\Users\chen.huaiyu\downloads")
    for i in filter(lambda x: yes_str in x, os.listdir()):
        os.rename(i, i.replace('~', '_' ))
    path = os.path.join(os.getcwd(), '消费报告 %s_%s.csv' % (yes_str, yes_str))
    return path

def cost_time(func):
    '耗时'
    @functools.wraps(func)
    def wrapper(*args):
        print('%s() start:' % func.__name__)
        start = time.time()
        func(*args)
        stop = time.time()
        print('\a\a cost time: %.2f min' % ((stop-start)/60))
    return wrapper

@cost_time
def run():
    '测试'
    print('测试正常')
    pass


def dff(df):
    # 2.统一时间列格式,将'-'变为None
    for i in ['开户日期', '首次消费日', '收取年服务费时间', '主体资质到期日', '加V缴费到期日']:
        df.loc[df[i] == '-', i] = None
        df[i] = pd.to_datetime(df[i])
    return df

def initBasicInfo(path):
    '''初始化
    从桌面获取  基本信息
    '''
    dic = {'用户名':str, '信誉成长值':str, '客户':str, '网站名称':str, '广告主':str}
    df = pd.read_excel(path, sheet_name='基本信息', converters=dic)
    # 1.修改序号列 为 “Id”
    df.rename(columns={df.columns[0]:'Id'}, inplace=True)
    # 2.统一时间
    df = dff(df)
    return df

def initBasicInfo2(df):
    # 复位ID
    df['Id'] = df_b.index.tolist()
    df = df.reindex(columns=col('basicInfo'))
    # 复位信誉值 ==>str
    df['信誉成长值'] = df['信誉成长值'].astype(str)
    df = dff(df)
    return df

def data(args):
    sql = "select * from %s" % args
    data = engine.execute(sql).fetchall()
    return data

def col(args):
    sql = "select * from information_schema.columns where table_name='%s'" % args
    col = [i[3] for i in engine.execute(sql).fetchall()]
    return col

def df(args):
    df = pd.DataFrame(data(args), columns=col(args))
    if 'Id' in col(args):
        df.drop(columns=['Id'], inplace=True)
    return df

@cost_time
def read_file():
    '读取数据'
    # 临时全局变量
    global df_b, df_i, df_kh, df_em, df_p
    # 文件地址
    path = path_date_str()
    # accountApplication
    df_kh = df('开户申请表')
    # basicInfo
    df_b = df('basicInfo')
    # channel
    df_p = df('channel')
    # personInfo
    df_employee = df('personInfo')
    # 字段
    df_normal = df('统一字段表')
    # 数据获取
    if os.path.exists(path):
        df_i = pd.read_csv(path, encoding='gbk', engine='python')
    else:
        raise 'Tips: path，路径文件不存在；'
    
    'icrm 转换'
    df_normal.set_index('第三方', drop=True, inplace=True)
    dic = dict(df_normal['标准'])
    df_i.rename(columns=dic, inplace=True)
    df_i['用户名'] = df_i['用户名'].astype(str)
    
    '基本信息 转换'
    lis_index = [3807, 1464, 3812, 3819]
    lis_user = ['000015', '001852', '0220595', '00852009']
    for n in range(len(lis_index)):
        df_b.loc[lis_index[n], '用户名'] = lis_user[n]
    '基本信息更新'
    lis1 = ['信誉成长值', 'URL', '加V缴费到期日', '端口', '网站名称', 
            '主体资质到期日', '今日账户状态', '开户日期']
    # 如无开户日期则不更新；
    if '开户日期' not in df_i.columns.tolist():
        lis1.remove('开户日期')
    df_b_i = df_b.drop(columns=lis1)
    df_b_i = pd.merge(df_b_i, df_i, on='用户名', how='left', suffixes=('', '_y'))
    df_b = df_b_i[df_b.columns]
    
    '开户申请表'
    df_kh.rename(columns=dic, inplace=True)
    
    '人员信息表'
    df_em = df_employee[['区域', '姓名', '跟随op']].copy()
    df_em.dropna(how='any', inplace=True)
    df_em.set_index('姓名', inplace=True)
    
    '端口'
    df_p.set_index('端口', inplace=True)
    # return df_b, df_i  ???调用返回 None  ——修改装饰器

def read2():
    '校验完后二次读取'
    pass


def new():
    '新消户筛选'
    global df_b, df_i
    # 开户/财务端口
    Port = df_p.loc[df_p['财务&开户'].notna(), '财务&开户'].index.tolist()
    acc = ['test-eee789']  # 测试账户
    
    '删除不统计端口'
    df_new = df_i.copy()
    df_new.drop(index=df_new[df_new['端口'].isin(Port)].index, inplace=True)
    df_new.drop(index=df_new[df_new['用户名'].isin(acc)].index, inplace=True)
    
    '求差集'
    df_new = df_new.append(df_b, sort=False)  # False 沉默警告，不排序
    df_new = df_new.append(df_b, sort=False)
    df_new.drop_duplicates('用户名', keep=False, inplace=True)  # 删除所有重复项
    df_new = df_new.reindex(columns=df_b.columns)
    df_new.fillna('-', inplace=True)
    return df_new

@cost_time
def new_b():
    '新消户数据补充'
    
    global df_em, df_new, df_b
    df_new = new()
    
    '如有新增消费账户，则补充相关信息；'
    if df_new.shape > (0, 29):
        '默认值'
        df_new.loc[:, 'BU'] = 'CSA'
        df_new.loc[:, '下单方'] = '海外渠道'
        '开户申请表'
        col1 = ['销售', 'AM', '资质归属地', '公司总部', 'Region', 
                'Industry', 'channel', '客户']
        for i in df_new['用户名'].tolist():
            for j in col1:
                try:
                    df_new.loc[df_new['用户名'] == i, 
                               j] = df_kh.loc[df_kh['用户名'] == i, j].values
                except Exception as e:
                    print('\a\a Error Row(141): 开户申请表无账户 %s\n%s' % (i,e))
                    continue
        '标准化'
        # 广告主
        df_new['广告主'] = df_new['广告主'].str.title()
        df_new['广告主'] = df_new['广告主'].str.replace(' ', '')
        # 端口分配
        ## '军朗'   使用端口判定
        lis1 = df_p.loc[df_p['客户'] == '北京军朗广告有限公司', 
                        '客户'].index.tolist()
        lis2 = ['顾凡凡', '陈宛欣', '香港', '香港', '香港', '代理商', '软件游戏']
        col3 = ['销售', 'AM', '资质归属地', '公司总部', 'Region', 'channel',
                'Industry']
        for n,i in enumerate(col3):
            df_new.loc[df_new['端口'].isin(lis1), i] = lis2[n]
        ## AM
        ### 获取端口列表,SZ账户按端口分配
        lis4 = df_p.loc[df_p['AM'].notna(), 'AM'].index.tolist()
        for i in set(df_new['端口']):
            print(i)
            if i in lis4:
                j = df_new.loc[df_new['端口'] == i, 'AM']
                # AM为'-'时忽略；影响后绪逻辑判断
                if j.values[0] == '-':
                    continue
                if df_em.loc[j.values[0], '区域'] == 'SZ':
                    j = df_p.loc[i, 'AM']
            else:
                continue
        
        '补充手段：据端口填充端口加款客户'
# =============================================================================
#         for i in set(df_new['端口']):
#             if i not in df_p.index:
#                 raise KeyError ("新端口;端口表中缺失，请补充")
#                 print(i)
#                 continue
#             elif pd.isna(df_p.loc[i, '客户']):
#                 print('跳过 %s;因：%s' %(i, df_p.loc[i, '客户']))
#                 continue
#             else:
#                 df_new.loc[df_new['端口'] == i, '客户'] = df_p.loc[i, '客户']
#                 print(df_p.loc[i, '客户'])
# =============================================================================
        
        # 新旧
        from zhconv import convert
        df_new['广告主'] = df_new['广告主'].apply(lambda x: convert(x, 'zh-cn'))
        print('Tips:Py对大小写敏感，旧广告主&新广告主=非')
        lis5 = []
        for i in df_new['广告主'].tolist():
            if i in df_b['广告主'].tolist():
                df_new.loc[df_new['广告主'] == i, '新旧客户'] = 'EB'
            else:
                if i in lis5:
                    df_new.loc[df_new['广告主'] == i, '新旧客户'] = 'EB'
                else:
                    df_new.loc[df_new['广告主'] == i, '新旧客户'] = 'NB'
                    lis5.append(i)
                
        # 区域 & 操作 & channel
        for i in df_em.index:
            df_new.loc[df_new['AM'] == i, '操作'] = df_em.loc[i, '跟随op']
            df_new.loc[df_new['AM'] == i, '区域'] = df_em.loc[i, '区域']
        ## 代理or直客
        # 如客户为空，即'-'，留白(channel;区域)
        df_y = df_new['客户'].str.title()
        df_y = df_y.str.replace(' ', '')
        df_x = df_y == df_new['广告主']
        df_new.fillna('-', inplace=True)
        df_new.loc[(df_x == False) & (df_new['客户'] != '-'), 
                   'channel'] = '代理商'
        df_new.loc[(df_x == True) & (df_new['客户'] != '-'),
                   'channel'] = '直接客户'
        
        ## HK调整
        df_new.loc[(df_new['区域'] == 'HK') & 
                    (df_new['channel'] == '代理商'), '区域'] = 'HK 4A'
        df_new.loc[(df_new['区域'] == 'HK') & 
                   (df_new['channel'] == '直接客户'), '区域'] = 'HK DS'
        # 财务做账区
        for i in df_new['端口'].tolist():
            df_new.loc[df_new['端口'] == i, 
                       '财务做账区域'] = df_p.loc[i, '财务做账区']
        # 收取年服务费时间
        df_b = df_b.append(df_new, ignore_index=True, sort=False)
        ## 如无开户日期则不更新；
        if '开户日期' not in df_i.columns.tolist():
            pass
        else:
            date = '收取年服务费时间'
            df_date = df_b.loc[(df_b[date] == '-') & (df_b['开户日期'] != '-'), :]
            df_date[date] = 0
            df_date[date] = pd.to_datetime(df_date[date])
            df_date['开户日期'] = pd.to_datetime(df_date['开户日期'])
            ## 收取年服务费时间 = 开户日期 + 1year
            after_a_year(df_date, date, '开户日期')
            df_b.loc[df_b['用户名'].isin(df_date['用户名'].tolist()), 
                     date] = df_date[date].apply(lambda x:str(x)).values
        # 结束,更新DB
        initBasicInfo2(df_b).to_sql('basicInfo1', con=engine, if_exists='replace', index=False)
    else:
        pass

def after_a_year(df, col1, col2):
    '+ 1年'
    ## 普通润平年处理;忽略世纪润年；
    try:
        i = timedelta(366)
        j = timedelta(365)
        if datetime.today().year % 4 == 0:
            df[col1] = df[col2].apply(lambda x:x+i)
        else:
            df[col1] = df[col2].apply(lambda x:x+j)
    except Exception as e:
        print('Tips row(173): %s' %e)

if __name__ == '__main__':
    
    #'保留测试账户，进行测试'  --已
    run()  # 测试
    # initBasicInfo()  # 初始化；从桌面读入基本信息，整理
    read_file()
    new_b()
    pass