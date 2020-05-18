# _*_ coding: utf-8 _*_
'''
Send email

'''

import smtplib
from getConfig import Conf
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def _format_addr(s):
    # 格式化邮件地址
    # Header 如是中文必须编码
    #
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def sendEmail(subject, message, files=None, to_addr='newIOSys'):
    # login info
    conf = Conf()
    smt, fro, pw = conf.getEmail('mail_baidu', 'sender server', 'email'
                                , 'password')
    to = conf.getToEmail('to_addr', to_addr)
    # header of email
    msg = MIMEMultipart()
    msg['From'] = fro
    msg['To'] = to
    msg['Subject'] = Header(subject, 'utf-8').encode()
    # 正文
    msg.attach(MIMEText(message, 'plain', 'utf-8'))
    # 附件
    import os
    if files != None:
        for i in range(len(files)):
            if os.path.isfile(files[i]):
                with open(files[i], 'rb') as f:
                    xl = MIMEApplication(f.read())
                    xl.add_header('Content-Disposition', 'attachment'
                                , filename=os.path.split(files[i])[-1])
                    msg.attach(xl)
    # 发送
    with smtplib.SMTP(smt, 25) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.set_debuglevel(1)
        smtp.login(fro, pw)
        try:
            smtp.sendmail(fro, to.split(','), msg.as_string())
        except Exception as e:
            print('Failed send: {}'.format(e))
        else:
            print('Success send.')

if __name__ == '__main__':
    sendEmail('from python ...', 'test')