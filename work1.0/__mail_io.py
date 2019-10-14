# -*- coding: utf-8 -*-
"""
Created on Thu Sep 20 12:57:12 2018
1.增：对销售&AM分列，以提取AM名字（统一使用第一个）
2.
@author: chen.huaiyu
"""

import poplib, time, os, re, datetime
import pandas as pd
import numpy as np
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from __MergeExcel import MergeExcel
import logging.config

listSubject = []
listSubject1 = []

PATH = r'C:\Users\chen.huaiyu\Desktop\Input\logging.conf'
logging.config.fileConfig(PATH)
logger = logging.getLogger('chinaSearch')

def decodeStr(s):
    value, charset = decode_header(s)[0]
    if charset:
        try:
            value = value.decode(charset)
        except UnicodeDecodeError:
            value = value.decode('gb18030')
    return value

def guessCharset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def getFileName(fileName):
    filePath, tempFileName = os.path.split(fileName)
    shotName, extension = os.path.splitext(tempFileName)
    return shotName

def getMailAttachment(server, from_addr, password, datestr, keywords):
    global newPath
    if os.path.exists(newPath) == False:
        os.makedirs(newPath)
    pop3Server = poplib.POP3_SSL(server)
    pop3Server.set_debuglevel(1)
    logger.info(pop3Server.getwelcome().decode('utf-8'))
    pop3Server.user(from_addr)
    pop3Server.pass_(password)
    message, size = pop3Server.stat()
    logger.info('Message:%s, Size:%s', message, size)
    num = len(pop3Server.list()[1])
    getFileSuccess = 0
    # 倒叙遍历邮件
    for i in range(num, 0, -1):
        resp, lines, octets = pop3Server.retr(i)
        msgContent = b'\r\n'.join(lines).decode('utf-8')
        msg = Parser().parsestr(msgContent)
        # 读取邮件时间
        date1 = time.strptime(msg.get('Date')[:24], '%a, %d %b %Y %H:%M:%S')
        date2 = time.strftime('%Y%m%d', date1)
        # 如日期不满足，跳出
        if date2 < datestr:
            break
        # 获取邮件标题和发件人
        for header in ['From', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header == 'Subject':
                    subject = decodeStr(value)
                    if ('开户进度' in subject) or ('开户及提前加款' in 
                       subject) or ('開戶及提前加款' in subject):
                        listSubject.append(subject)
                    logger.info('邮件标题:%s', subject)
                else:
                    hdr, addr = parseaddr(value)
                    name = decodeStr(hdr)
                    fromname = u'%s' % (name)
                    # fromaddr = u'%s' % (addr)
        logger.info('发件人%s: %s', i, name)
        logger.info('====================')
        if keywords in subject:
            for part in msg.walk():
                filename = part.get_filename()
                if filename:
                    data = part.get_payload(decode=True)
                    os.chdir(newPath)
                    f = open('%s%s.xlsx' % (date2, fromname), 'wb')
                    f.write(data)
                    f.close()
                    logger.info('文件下载成功！')
                    getFileSuccess  += 1
                else:
                    logger.info('匹配成功， 但无附件！\n')
                    pass
            logger.info('下载文件：%s个', getFileSuccess)
        else:
            logger.info('无匹配邮件！\n')
    pop3Server.quit()
    
def reDuplicates():
    global listSubject1
    for i in listSubject:
        listSubject1 += re.findall(
                r'''[a-zA-Z]{2,10}\-[\u4E00-\u9FA5a-zA-Z0-9\-]*[a-zA-Z0-9]''',
                                   i)
    listSubject1 = list(set(listSubject1))
    mailSubject1 = pd.DataFrame({'用户名':listSubject1, 
                                 'channel':np.ones(len(listSubject1)), 
                                '客户':np.ones(len(listSubject1)), 
                                'Region':np.ones(len(listSubject1)),
                                '付款币种':np.ones(len(listSubject1))})
    mailSubject2 = pd.read_excel(path1[0])
    mailSubject = mailSubject2.append(mailSubject1, sort=False)
    mailSubject.drop_duplicates(subset='用户名', keep='first', inplace=True)
    ioAccount = pd.read_excel(path1[1], sheet_name=1, usecols=[1])
    m1 = pd.concat((mailSubject, ioAccount), axis=0, sort=True, join='outer')
    m1.drop_duplicates(subset='用户名', keep='last', inplace=True)
    mailSubject = m1[m1['channel'].notnull()]
    mailSubject.to_excel(path1[0], index=False, freeze_panes=(1,0))  # 新申户文件输出
    
def mainIO(n, path):
    '主程序'
    try:
        start = time.clock()
        global path1
        path1 = [r'H:\SZ_数据\Input\io system.xlsx',
                r'H:\SZ_数据\Input\IO系統母版-3.07.xlsm',
                r'H:\SZ_数据\Output']
        for i in path1:
            if os.path.exists(i):
                pass
            else:
                try:
                    raise ValueError('地址不存在.')
                except ValueError as e:
                    logger.error(e)
                
        # 登陆信息：
        import configparser
        cf = configparser.ConfigParser()
        cf.read(path)
        
        keywords = '开户进度表'
        logger.info('%s 天前；', n)
        datestr = (datetime.datetime.today()-datetime.timedelta(n)).strftime('%Y%m%d')
        global newPath
        newPath = os.path.join(path1[2], datestr + keywords)
        logger.info('请稍候...')
        getMailAttachment(cf.get('mail_baidu', 'receiving server'), 
                          cf.get('mail_baidu', 'email'), 
                          cf.get('mail_baidu', 'password'), 
                          datestr, keywords)  # 文件获取
        reDuplicates()
        MergeExcel(newPath)
        logger.info('运行耗时：{:3f}min'.format((time.perf_counter() - start)/60))
        logger.info('获取：\n1.新申户；\n2.开户进度表总表,需手动修复；')
        logger.info('运行结束')
    except Exception as e:
        logger.error(e)
    pass

if __name__ == '__main__':
    # 账号密码 配置文件地址
    path = r'c:\users\chen.huaiyu\Desktop\Input\c.s.conf'
    mainIO(2, path)
    pass
        