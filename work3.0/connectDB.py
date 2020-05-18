# _*_ coding: utf-8 _*_
'''

连接数据库
'''

import pandas as pd


data = {'dat':['20200101', '20200102', '20200103', '20200104'
         , '20200105', '20200106', '20200107', '20200108']
        , 'num': list(range(8))}
df = pd.DataFrame(data)
print(df)