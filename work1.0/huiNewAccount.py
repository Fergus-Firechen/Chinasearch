# -*- coding: utf-8 -*-
'''
# 2019.6.11
# 目标：读入目标数据，处理完后分词输出
# 思路：
    - .xlsx;.csv读取
    - icrm数据处理；hui数据处理；已开户账户准备
    - 对公司名称分词

'''

from zhconv import convert
from datetime import datetime
import pandas as pd
import os
import jieba
import time

now = lambda : time.perf_counter()

def readFile(path, shtname=0):
    '''读取 csv / xlsx
    '''
    if os.path.splitext(path)[1] == '.csv':
        return pd.read_csv(path, engine='python')
    elif os.path.splitext(path)[1] == '.xlsx':
        return pd.read_excel(path, sheet_name=shtname)

def dataHandlingTar(df):
    '''找出已找到的开户账户
    '''
    # 构建df.columns
    df.columns = df.loc[0, :].values
    df.drop(index=[0], inplace=True)
    # 去null, '-', '无'
    df.drop(index=df[df['开户日期'].isna() | (df['开户日期'] == '-') | 
            (df['开户日期'] == '无')].index, inplace=True)
    return df[df['开户日期'] >= datetime(2018, 1, 1)]


def dataHandlingIcrm(df, ls):
    # 筛选2018以后
    df = df[df['开户日期'] >= '2018-01-01']
    # 广告主 繁转简
    df['公司名称'] = df['公司名称'].apply(lambda x: convert(x, 'zh-cn'))
    # 删除已找到的
    for i in ls:
        df.drop(index=df[df['账户名称'] == i].index, inplace=True)
    return df

def dataHandlingHui(df, ls):
    # 除空、日期
    df.drop(columns=df.columns[-4:], inplace=True)
    df.dropna(axis=0, how='any', inplace=True)
    # 繁转简
    df = df.applymap(lambda x: convert(x, 'zh-cn'))
    # 多余抬头
    df.drop(index=df[df['公司名称'] == '公司名称'].index, inplace=True)
    # 除 公司名称 == 联系人
    df = pd.merge(df, df['联系方式'].str.split(r'(', expand=True), how='left', left_index=True, right_index=True)
    df.drop(index=df[df['公司名称'] == df[0]].index, inplace=True)
    # 增序号
    df['ID'] = df.index
    # 除已找到
    for i in ls:
        df.drop(index=df[df['公司名称'] == i].index, inplace=True)
    return df
    
def jieba_(ls):
    s = ','.join(ls)
    s = s.lower()
    words = jieba.lcut(s)
    dic = {}
    for word in words:
        dic[word] = dic.get(word, 0) + 1
    dic = sorted(dic.items(), key=lambda k: k[1], reverse=False)
    df = pd.DataFrame(dic, columns=['Word', 'Cnt'])
    return df[df['Cnt'] == 1]

def main():
    path = r'D:\陈怀玉\工作\月工作\新开户跟进-辉\账户新开跟进'
    icrm = '消费报告 20191017_20191017.csv'
    hui = '170517-191008.xlsx'
    target = '已开户公司确认 v21.xlsx'
    
    # 已消费户 df_t['icrm公司名称]
    df_t = readFile(os.path.join(path, target), shtname='对应')
    df_t = dataHandlingTar(df_t)
    df_t['icrm公司名称'] = df_t['icrm公司名称'].apply(lambda x: convert(x, 'zh-cn'))
    ls = df_t['icrm公司名称'].tolist()
    
    # 辉数清洗
    df_hui = readFile(os.path.join(path, hui))
    df_hui = dataHandlingHui(df_hui, ls)
    df_hui.to_csv(os.path.join(path, '辉简.csv'), encoding='utf-8-sig')
    df_hui = jieba_(df_hui['公司名称'].tolist())
    df_hui.to_csv(os.path.join(path, 'hui.csv'), encoding='utf-8-sig')
    
    # 处理：icrm
    df_icrm = readFile(os.path.join(path, icrm))
    df_icrm = dataHandlingIcrm(df_icrm, ls)
    df_icrm.to_csv(os.path.join(path, 'icrm简.csv'), encoding='utf-8-sig')
    df_icrm = jieba_(df_icrm['公司名称'].tolist())
    df_icrm.to_csv(os.path.join(path, 'icrm.csv'), encoding='utf-8-sig')

if __name__ == '__main__':
    start = now()
    main()
    print('运行结束,耗时: {:.3f} s'.format(now() - start))
    