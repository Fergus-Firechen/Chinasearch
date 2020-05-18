# _*_ coding: utf-8 _*_

from sqlalchemy import create_engine
from getConfig import Conf

def getUrl(sec, acc, pw, ip, port, db):
    conf = Conf()
    url = ('mssql+pymssql://%s:%s@%s:%s/%s'
            % conf.getInfo(sec, acc, pw, ip, port, db))
    return url
    
if __name__ == '__main__':
    path = r'H:\SZ_数据\Python\c.s.conf'
    print(getUrl('Output', 'acc', 'pw', 'ip', 'port', 'db'))
    
    
    
    