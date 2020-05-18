# _*_ conding: utf-8 _*_
'''

从配置文件中读取数据
'''

from configparser import ConfigParser

class Conf():

    def __init__(self):
        self._path = r'H:\SZ_数据\Python\c.s.conf'

    def getInfo(self, sec, U, P, ip, port, db):
        fil = ConfigParser()
        fil.read(self._path)
        return (fil.get(sec, U), 
                fil.get(sec, P),
                fil.get(sec, ip),
                fil.get(sec, port),
                fil.get(sec, db))
    
    def getEmail(self, sec, smt, email, pw):
        fil = ConfigParser()
        fil.read(self._path)
        return (fil.get(sec, smt),
                fil.get(sec, email),
                fil.get(sec, pw))
                
    def getToEmail(self, sec, toEmail):
        fil = ConfigParser()
        fil.read(self._path)
        return fil.get(sec, toEmail)
        
if __name__ == '__main__':
    
    from sqlalchemy import create_engine

    conf = Conf()
    url = ("mssql+pymssql://%s:%s@%s:%s/%s" % 
           conf.getInfo('Output', 'acc', 'pw', 'ip', 'port', 'db'))
    with create_engine(url).begin() as conn:
        print(conn.execute('select 1').fetchall())


