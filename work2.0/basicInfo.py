# -*- coding: utf-8 -*-
"""
Created on Tue Mar 19 20:10:33 2019

# 新消户不存于端口表中
# DB取代文件“数据拆解表”
# df_kh 字段名统一？
# 20190528 增首次消费日
# Q basicInfo乱序：删除Id前修改 index=Id  -- 2019.6.14
# Q basicInfo乱序：从sqlserver读取basicInfo后，立即按Id进行排序  -- 2019.6.17
@author: chen.huaiyu
"""
import os
import time
import functools
import configparser
import pandas as pd
#import xlwings as xw
from sqlalchemy import create_engine
from datetime import datetime, timedelta


def path_date_str(n):
    '消费文件地址构造，默认昨日'
    yes_str = datetime.strftime(datetime.today() - timedelta(n), "%Y%m%d")
    print("默认昨日：{}".format(yes_str))
    path = os.chdir(r"H:\SZ_数据\Download")
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
    # 3.  2020-2-10
    #
    df.loc[df['加V缴费到期日'] == '0002-11-30 00:00:00', '加V缴费到期日'] = None
    # 2.统一时间列格式,将'-'变为None
    for i in ['开户日期', '首次消费日', '收取年服务费时间', '主体资质到期日', '加V缴费到期日']:
        try:
            df.loc[df[i] == '-', i] = None
            df[i] = pd.to_datetime(df[i])
        except ValueError as e:
            print('%s : %s' % (i, e))
    return df

def initBasicInfo(path):
    '''初始化
    从桌面获取  基本信息
    '''
    dic = {'用户名':str, '客户':str, '网站名称':str, '广告主':str}
    df = pd.read_excel(path, converters=dic)
    # 1.修改序号列 为 “Id”
    df.rename(columns={df.columns[0]:'Id'}, inplace=True)
    # 2.统一时间
    df = dff(df)
    return df

def initBasicInfo2(df):
    # 复位ID
    df['Id'] = df_b.index
    df = df.reindex(columns=col('basicInfo'))
    df = dff(df)
    # 排除用户名、端口、广告主 包含'\r\n'
    # 信誉成长值 == 二级行业
    ls = ['用户名', '端口', '广告主', 'URL', '网站名称', 'Industry', '信誉成长值']
    for i in ls:
        df[i] = df[i].str.replace('\r\n', '')
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
    df.sort_values('Id', ascending=True, inplace=True)
    if 'Id' in col(args):
        df.index = df['Id']
        df.drop(columns=['Id'], inplace=True)
    return df

@cost_time
def read_file(n):
    '读取数据'
    # 临时全局变量
    global df_b, df_i, df_kh, df_em, df_p
    # 文件地址
    path = path_date_str(n)
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
        raise 'Tips: df_i 路径文件不存在：%s' 
    
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
    lis1 = ['URL', '加V缴费到期日', '端口', '网站名称', 
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


def new(n):
    '''新消户筛选：必须有消费'''
    global df_b, df_i
    # 开户/财务端口
    # Port = df_p.loc[df_p['财务&开户'].notna(), '财务&开户'].index.tolist()
    acc = ['test-eee789']  # 测试账户
    
    # '删除不统计端口：财务端口 & 开户端口'
    df_new = df_i.copy()
    # df_new.drop(index=df_new[df_new['端口'].isin(Port)].index, inplace=True)
    df_new.drop(index=df_new[df_new['用户名'].isin(acc)].index, inplace=True)
    
    # '求差集:去除所有消费报告中已有的账户'
    df_new = df_new.append(df_b, sort=False)  # False 沉默警告，不排序
    df_new = df_new.append(df_b, sort=False)
    df_new.drop_duplicates('用户名', keep=False, inplace=True)  # 删除所有重复项
    df_new = df_new.reindex(columns=df_b.columns)
    df_new.fillna('-', inplace=True)
    
    # 筛选有消费的账户
    sql = "select distinct 用户名 from 消费"
    ls_username = [i[0] for i in engine.execute(sql).fetchall()]
    
    # 去除无消费的账户
    for i in df_new['用户名']:
        if i in ls_username:
            pass
        else:
            df_new.drop(index=df_new[df_new['用户名'] == i].index, inplace=True)
    
    return df_new

@cost_time
def new_b(n):
    '新消户数据补充'
    
    global df_em, df_new, df_b
    df_new = new(n)
    
    # 如有新增消费账户，则：
    # 1.补充相关信息
    # 2.更新数据库
    #
    if df_new.shape > (0, 29):
        '默认值'
        df_new.loc[:, 'BU'] = 'CSA'
        df_new.loc[:, '下单方'] = '海外渠道'
        '开户申请表'
        col1 = ['销售', 'AM', '资质归属地', '公司总部', 'Region', 'channel', '客户']
        for i in df_new['用户名'].tolist():
            for j in col1:
                try:
                    df_new.loc[df_new['用户名'] == i, 
                               j] = df_kh.loc[df_kh['用户名'] == i, j].values[0]
                except Exception as e:
                    print('\a\a Error Row(141): 开户申请表无账户 %s\n%s' % (i,e))
                    continue
        # 一级行业 & 二级行业   
        for i in df_new['用户名'].tolist():
            try:
                df_new.loc[df_new['用户名'] == i, 
                           'Industry'] = df_kh.loc[df_kh['用户名'] == i, '一级行业'].values[0]
                df_new.loc[df_new['用户名'] == i,
                           '信誉成长值'] = df_kh.loc[df_kh['用户名'] == i, '二级行业'].values[0]
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
        lis2 = ['顾凡凡', '陈宛欣', '香港', '香港', '香港', '代理商']
        col3 = ['销售', 'AM', '资质归属地', '公司总部', 'Region', 'channel']
        for n,i in enumerate(col3):
            df_new.loc[df_new['端口'].isin(lis1), i] = lis2[n]
        ## AM
        ### 获取端口列表,SZ账户按端口分配
        lis4 = df_p.loc[df_p['AM'].notna(), 'AM'].index.tolist()
        for i in set(df_new['端口']):
            print(i)
            if i in lis4:
                # AM 为空，跳过
                if df_new.loc[df_new['端口'] == i, 'AM'].shape == (0,):
                    continue
                # AM为'-'时忽略；影响后绪逻辑判断
                if df_new.loc[df_new['端口'] == i, 'AM'].values[0] == '-':
                    df_new.loc[df_new['端口'] == i, 'AM'] = df_p.loc[i, 'AM']
                    continue
                try:
                    if 'SZ' in df_em.loc[df_new.loc[df_new['端口'] == i, 'AM'].values[0], '区域'].tolist():
                        df_new.loc[df_new['端口'] == i, 'AM'] = df_p.loc[i, 'AM']
                except:
                    if 'SZ' == df_em.loc[df_new.loc[df_new['端口'] == i, 'AM'].values[0], '区域']:
                        df_new.loc[df_new['端口'] == i, 'AM'] = df_p.loc[i, 'AM']
            else:
                continue
        
        '补充手段：据端口填充端口加款客户'
        for i in set(df_new['端口']):
            if i not in df_p.index:
                print(i)
                raise KeyError ("新端口;端口表中缺失，请补充")
                continue
            elif pd.isna(df_p.loc[i, '客户']):
                print('跳过 %s;因：%s' %(i, df_p.loc[i, '客户']))
                continue
            else:
                df_new.loc[df_new['端口'] == i, '客户'] = df_p.loc[i, '客户']
                print(df_p.loc[i, '客户'])
        
        # 新旧
        from zhconv import convert
        df_new['广告主'] = df_new['广告主'].apply(lambda x: convert(x, 'zh-cn'))
        df_new['客户'] = df_new['客户'].apply(lambda x: convert(x, 'zh-cn'))
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
        df_em.drop(index = df_em[df_em['区域'] == 'SG'].index, inplace=True)
        for i in df_em.index:
            print(i)
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
        # 
        new_b = df_new['用户名'].tolist()
        if len(new_b) > 0:
            df_new = df_b[df_b['用户名'].isin(new_b)]
            df_new.to_sql('basicInfo', con=engine, 
	                      if_exists='append', index=False)
    else:
        print('无新消户')
    # 结束,准备更新DB
    df_b = initBasicInfo2(df_b)

@cost_time
def update_basicInfo():
    '基本信息更新'
    lis1 = ['URL', '加V缴费到期日', '端口', '网站名称', 
            '主体资质到期日', '今日账户状态', '开户日期', '用户名']
    # 3.0 筛选，仅对lis1中在icrm发生了更新了的账户
    # 3.1 筛选前更新全部数据  筛选后预估<<1000
    # 3.2 --重新从数据库中读取数据，并以更新列完全去重； --后以用户名保留首位去重
    # 3.3 -- 放弃：许多户除用户名其它都一样  -- 另数据量不大，消耗资源有限
    #
    df_db = df('basicInfo')
    df1 = df_b.append(df_db, sort=False)
    df1.drop_duplicates(lis1, keep=False, inplace=True)
    df1.drop_duplicates('用户名', keep='first', inplace=True)
    # 开始更新
    print('更新账户：{}个'.format(df1.shape[0]))
    try:
        for i in df1[lis1].values:
            sql_ = ''' UPDATE basicInfo
                SET URL='{}', 加V缴费到期日='{}', 端口='{}', 网站名称='{}', 主体资质到期日='{}', 今日账户状态='{}', 开户日期='{}' 
                WHERE 用户名='{}'
                '''
            # 1. py中NaT & sql中NULL不兼容
            if isinstance(i[1], type(pd.NaT)):
                i[1] = 'NULL'
                sql_ = sql_.replace("加V缴费到期日='{}'", "加V缴费到期日={}")
            if isinstance(i[4], type(pd.NaT)):
                i[4] = 'NULL'
                sql_ = sql_.replace("主体资质到期日='{}'", "主体资质到期日={}")
            if isinstance(i[6], type(pd.NaT)):
                i[6] = 'NULL'
                sql_ = sql_.replace("开户日期='{}'", "开户日期={}")
            # 2.0. py中，可以使用双引号"&单引号'之间的嵌套
            # 2.1. sql中，'一般均用两个''替代
            #
            if isinstance(i[3], str) and "'" in i[3]:
                    i[3] = i[3].replace("'", "''")
            sql = sql_.format(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7])
            engine.execute(sql)
    except Exception as e:
        print(e)
    else:
        print('The update is normal.')


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

def update_first_spend_date():
    '''更新首次消费日
    '''
    # 1.1 筛选 basicInfo中首次消费日为null的账户
    sql = "select 用户名 from basicInfo where 首次消费日 is null"
    lis1 = (i[0] for i in engine.execute(sql).fetchall())
    
    # 1.2 筛选 spending中 ‘总点击’：用户名,并去重
    # 1.3 筛除 1.1中不在1.2中的账户
    # 1.4 如存在，查询对应账户日期并更新basicInfo
    #
    # 消费表中用户名去重提取
    #
    sql1 = "select 用户名 from 消费 where 类别='总点击'"
    userName = set([i[0] for i in engine.execute(sql1).fetchall()])
    
    # 准备：
    # 查询spending 首次消费日
    #
    sql2 = "select 日期 from 消费 where 用户名=%s order by 日期"
    # 更新basicInfo2 首次消费日
    sql3 = "update basicInfo set 首次消费日=%s where 用户名=%s"
    for i in lis1:
        # 判断 无首消费 户是否在消费表，如不在，跳过
        if i not in userName:
            continue
        else:
            # 如在，查询最早日期，并更新basicInfo2
            # 查询，升序
            date_firstSpending = engine.execute(sql2, i).fetchone()[0]
            # ‘20190528’ ==> datetime.datetime
            date_date = datetime.strptime(date_firstSpending, '%Y%m%d')
            # 更新basicInfo
            engine.execute(sql3, str(date_date), i)
            
def connectDB():
    def login():
        CONF = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
        conf = configparser.ConfigParser()
        if os.path.exists(CONF):
            conf.read(CONF)
            host = conf.get('SQL Server', 'ip')
            port = conf.get('SQL Server', 'port')
            dbname = conf.get('SQL Server', 'dbname')
            return host, port, dbname
    try:
        engine = create_engine(
                'mssql+pymssql://@%s:%s/%s' % login())
    except Exception as e:
        raise Exception('连接成功 %s' % e)
    else:
        print('连接成功')
        return engine


if __name__ == '__main__':
    
    # 连接 DB
    engine = connectDB()
    
    #'保留测试账户，进行测试'  --已
    run()  # 测试
    # initBasicInfo()  # 初始化；从桌面读入基本信息，整理
    n = input('默认昨日(Enter)')  # 昨日=1
    if n == '':
    	n = 1
    else:
    	n = eval(n)
    read_file(n)
    new_b(n)
    update_first_spend_date()
    update_basicInfo()
    pass