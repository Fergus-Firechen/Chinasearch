# _*_ coding: utf-8 _*_
'''
1.获取路径下的文件
2.获取指定【关键词】筛选的文件

'''


import os

class Doc():
    def __init__(self, path):
        self._path = path
        
    def getAll(self):
        return os.listdir(self._path)
        
    def getSome(self, key):
        return [f for f in os.listdir(self._path) if key in f]
        
    def getRecent(self):
        lis = os.listdir(self._path)
        lis.sort(key=lambda x: 
            os.stat(os.path.join(self._path, x)).st_mtime, reverse=True)
        return lis
        
if __name__ == '__main__':

    doc = Doc(r'H:\SZ_数据\Download\15-18年 ICRM数据')
    print(doc.getAll())
    print()
    print(doc.getSome('消费'))
    print()
    print(doc.getSome('现金'))
    print()
    print(doc.getSome('百通'))
    print()
    print(doc.getRecent())