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


def dataHandlingIcrm(df):
    # 筛选2018以后
    df = df[df['开户日期'] >= '2018-01-01']
    # 广告主 繁转简
    df['公司名称'] = df['公司名称'].apply(lambda x: convert(x, 'zh-cn'))
    # 
    
    
    pass

def dataHandlingHui():
    pass
    
def jieba():
    pass

def main():
    path = r'D:\陈怀玉\工作\月工作\新开户跟进-辉\账户新开跟进'
    icrm = '消费报告 20190610_20190610.csv'
    hui = '170517-190603.xlsx'
    target = '已开户公司确认 v16.xlsx'
    
    # 已消费户
    df_tar = readFile(os.path.join(path, target), shtname='对应')
    df_t = dataHandlingTar(df_tar)
    
    df_hui = readFile(os.path.join(path, hui))
    df2 = dataHandlingHui(df_hui)
    jieba(df2)
    
    df_icrm = readFile(os.path.join(path, icrm))
    df1 = dataHandlingIcrm(df_icrm)
    jieba(df1)

if __name__ == '__main__':
    
    pass