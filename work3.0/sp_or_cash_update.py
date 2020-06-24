# -*- coding: utf-8 -*-
"""
Spyder 编辑器

sp_or_cash_update.py
"""

from sqlalchemy import create_engine
from datetime import datetime
from db import getUrl
import pandas as pd
import time
import os


now = lambda : time.perf_counter()

class data:
    '''
    
    Input
    -----
    startTime, 数据起始日，格式如"20200601"
    endTime, 数据截止日，格式如"20200630"
    spOrCash, 上载数据类型选择，消费或现金
    
    Returns
    -------
    None.
    '''
    
    def __init__(self, startTime, endTime, spOrCash):
        
        self.path = r'H:\SZ_数据\Download\IDATE数据'
        self.startTime = startTime
        self.endTime = endTime
        self.spOrCash = spOrCash
        # 输入参数检查
        try:
            # 检查输入日期
            datetime.strptime(self.startTime, '%Y%m%d')
            datetime.strptime(self.endTime, '%Y%m%d')
            # 检查产品类型
            if self.spOrCash == '消费':
                pass
            elif self.spOrCash == '现金':
                pass
            else:
                raise
        except:
            raise ValueError("请按要求输入起始日期,如:20200601,20200630,消费")
            
    def getPath(self):
        '''
        1.下载的数据分为两种：即消费&现金，数据结构完全一致，后续处理完全一致；
        2.消费/现金数据最终各获得两个csv文件；
        3.下载文件命名规则：消费/现金->p4p_date1_date2.csv;无线_date1_date2.csv
        4.据文件下载路径拼接出绝对路径。

        Raises
        ------
        FileNotFoundError
            检查路径下文件是否存在。

        Returns
        -------
        p4p : TYPE
            文件1绝对路径
        mob : TYPE
            文件2绝对路径

        '''
        if self.spOrCash == '消费':
            p4p, mob = (os.path.join(self.path, 'p4p_' + self.startTime 
                                     + '_' + self.endTime + '.csv'),
                        os.path.join(self.path, '无线_' + self.startTime
                                     + '_' + self.endTime + '.csv'))
            if os.path.exists(p4p) and os.path.exists(mob):
                return p4p, mob
            else:
                raise FileNotFoundError("文件不存在：\n1)%s, \n2)%s" % p4p, mob)
        elif self.spOrCash == '现金':
            p4p, mob = (os.path.join(self.path, 'cash_' + self.startTime 
                                     + '_' + self.endTime + '.csv'),
                        os.path.join(self.path, '无线_' + self.startTime
                                     + '_' + self.endTime + '_cash.csv'))
            if os.path.exists(p4p) and os.path.exists(mob):
                return p4p, mob
            else:
                raise FileNotFoundError("路徑下文件不存在：\n1)%s, \n2)%s" 
                                        % p4p, mob)
            
    def readFil(self):
        '''
        分别讀取下载的两个文件數據
        
        Returns
        -------
        TYPE
            返回下载的两个文件的数据

        '''
        # 获取文件路径
        path_p4p, path_mob = self.getPath()
        # 获取数据
        df_p4p, df_mob = (pd.read_csv(path_p4p, encoding='utf-8'
                                      , engine='python')
                          , pd.read_csv(path_mob, encoding='utf-8'
                                        , engine='python'))
        self.clean(df_p4p, df_mob)
        return df_p4p, df_mob
    
    @staticmethod
    def clean(p4p, mob):
        '''
        剔除不需要的列

        Parameters
        ----------
        p4p : TYPE
            下载p4p的消费或现金数据
        mob : TYPE
            下载的无线的消费或现金数据

        Returns
        -------
        None.

        '''
        
        # 删除多余的列
        # p4p
        columns = [c for c in p4p.columns if '消费' not in c and '账户名称' != c]
        p4p.drop(columns=columns, inplace=True)
        # mob
        columns = [c for c in mob.columns if '消费' not in c and '账户名称' != c]
        mob.drop(columns=columns, inplace=True)
            
    def merge_(self):
        '''
        合并下载的两份数据文件

        Returns
        -------
        返回合并后的数据

        '''
        df_p4p, df_mob = self.readFil()
        try:
            df = df_p4p.merge(df_mob, on='账户名称', suffixes=('', '_')
                              , how='outer')
            df.fillna(0, inplace=True)
        except KeyError:
            print('数据编码异常，修改p4p和无线文件编码为"utf-8"')
        else:
            return df
    
    def oneDimension(self):
        '''
        将数据转换为一维表

        Returns
        -------
        df : TYPE
            返回转换后的一维表

        '''
        
        # 合并后的数据
        df = self.merge_()
        # 将账户名称换为索引
        df.set_index('账户名称', drop=True, inplace=True)
        # 转置为一维
        df = df.stack()
        # 重置索引列
        df = df.reset_index()
        # 分列
        ## 无线的日期中包含'_移动(含阿拉丁)',需剔除
        #
        df[['子类', '日期']] = df['level_1'].str.split('消费', expand=True)
        df.drop(columns=['level_1'], inplace=True)
        df['日期'] = df['日期'].apply(lambda x: x.replace('_移动(含阿拉丁)', ''))
        # 重命名
        dic = {'账户名称':'用户名', 0:'金额'}
        df.rename(columns=dic, inplace=True)
        return df
    
    def flag(self):
        '''
        将子类聚合为搜索、原生、新产品

        Returns
        -------
        df : TYPE
            返回标识聚合后的一维表

        '''
        
        df = self.oneDimension()
        df['类别'] = ''
        # 搜索点击 = 搜索点击 + 点击质量调整
        df.loc[(df['子类'] == '搜索点击') | (df['子类'] == '点击质量调整')
               , '类别'] = '搜索点击'
        # 自主投放 = 原生CPC + 原生CPM
        df.loc[(df['子类'] == '原生CPC') | (df['子类'] == '原生CPM')
               , '类别'] = '自主投放'
        # 无线搜索点击 = 凤巢
        df.loc[df['子类'] == '凤巢', '类别'] = '无线搜索点击'
        # 新产品 = 凤巢阿拉丁 + 品牌起跑线 + 品牌华表 + 图片推广 + 知识营销 + 百通
        df.loc[df['类别'] == '', '类别'] = '新产品'
        df = self.sum_(df)
        return df
    
    @staticmethod
    def sum_(df):
        '''
        换算总点击：= 搜索点击 + 新产品 + 原生

        Parameters
        ----------
        df : TYPE
            待换算数据

        Returns
        -------
        df : TYPE
            返回换算后的数据

        '''
        
        df = df.pivot_table('金额', columns='类别'
                            , index=['用户名','子类','日期'], aggfunc='sum')
        df.fillna(0, inplace=True)
        # 计算总点击
        df['总点击'] = df['搜索点击'] + df['新产品'] + df['自主投放']
        # 转换为一维表
        df = df.stack()
        df = df.reset_index()
        df.rename(columns={0:'金额'}, inplace=True)
        return df



if __name__ == '__main__':
    
    st = now()
    print('Idata system spending or cash data upload:', end='\n\n')
    # 
    try:
        startTime, endTime, spOrCash = input('请输入开始日期,结束日期,消费/现金：\
                                                  如："20200601,20200630,消费"\
                                                      ').split(',')
        # 实例化
        instance = data(startTime, endTime, spOrCash)
        # 数据合并、换算
        df = instance.flag()
        # 连接数据库上载
        with create_engine(getUrl('SQL Server', 'acc', 'pw', 'ip'
                                  , 'port', 'db')).begin() as conn:
            df.to_sql('消费_idata', con=conn, index=False, if_exists='append')
    except Exception as e:
        print(e)
    else:
        print('程序结束：%s. \n耗时%.3fs' % (e, now() - st))
        
        
        
        with open(r'') as f:
            print(f.read())