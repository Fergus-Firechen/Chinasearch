# -*- coding: utf-8 -*-
"""
Created on Mon Apr  8 09:38:35 2019

# 2019.4.16 增日志模块
# 2019.4.17 增空行标识行；1）标识程序运行日期(check点);2)提升运行效率;
# 逻辑错误：
# 
@author: chen.huaiyu
"""

import poplib, time, datetime, re, functools, os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from sqlalchemy import create_engine
import pandas as pd
import logging.config


columns = ['日期', '合同原件是否已回', '是否赠送服务费', '是否开票', 
           '推广性质', '销售', '客服', '用户名', '端口', '行业',
           '渠道', '广告主总部', '资质归属地', '预估月消费', 
           'Region', '账期/预付', '服务费', '年费', '服务费币种', 
           '网站名称', '广告主名称', '广告主_简体', 'URL', '登记证编号', '生效日',
           '届满日期', '联系人', '电话', '客户', 'flag']

# 日志
PATH = r'C:\Users\chen.huaiyu\Chinasearch\logging.conf'
logging.config.fileConfig(PATH)
logger = logging.getLogger('chinaSearch')

# 连接SQL Server
engine = create_engine(r'mssql+pyodbc://SQL Server')
if engine.execute('select 1'):
    logger.info('SQL Server 连接正常')
    
# 删除DB中标识项
engine.execute("DELETE FROM 开户申请表 WHERE 用户名 = '0.0'")
df = pd.DataFrame(engine.execute('select * from 开户申请表 order by 日期'
                                     ).fetchall(), columns=columns)

def cost_time(func):
    '''耗时跟进'''
    @functools.wraps(func)
    def wrapper(*args):
        logger.info('%s() start:', func.__name__)
        start_time = time.time()
        func(*args)
        end_time = time.time()
        logger.info('cost time: %.2f min', ((end_time-start_time)/60))
    return wrapper
    
@cost_time
def run():
    '''测试'''
    pass
    
def decode_str(s):
    '''
    # 邮件的主题是经过编码后的str,先解码
    # decode_header返回list,像Cc,Bcc这样的字段可包含多个邮件地址，有多个元素
    # --此处只取了一个元素
    '''
    value, charset = decode_header(s)[0]  # 解码，不转换
    if charset:
        try:
            value = value.decode(charset)
        except UnicodeDecodeError:  # GB2312 < GBK < GB18030
            value = value.decode('GB18030')
    return value

def guess_charset(msg):
    '''
    # 邮件内容为str,先检测内容编码；否则非utf-8类型无法显示
    # get_charsets():
    # 1)返回包含消息中字符名称的列表；
    # 2)如果消息是多部分，则列表将包含有效负载中每个子部分的一个元素
    # -- 否则，它将是长度为1的列表
    '''
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def print_info(msg, indent=0):
    '''解析邮件内容
    # 只需一行即可将邮件内容解析为Message对象
    # msg = Parser().parsestr(msg_content)
    # Message对象本身可能是一个MIMEMultipart对象，
    # --即包含嵌套的其它MIMEBase对象，可能不止一层
    # 要递归打印出Message对象
    # indent用于缩进显示
    '''
    if indent == 0:
        for header in ['From', 'To', 'Subject', 'Date']:
            value = msg.get(header, '')
            if value:
                if header == 'Subject':
                    value = decode_str(value)
                    # '申明全局变量，为后续判定做准备'
                    global subject
                    subject = value
                elif header == 'Date':
                    value = msg.get(header, '')
                else:
                    hdr, addr = parseaddr(value)
                    name = decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
            logger.info('%s %s: %s', '  '*indent, header, value)
    # is_multipart():判定是否为EmailMessage对象
    # 如是有效载荷子的列表，返回True；否则False
    #
    if (msg.is_multipart()):
        parts = msg.get_payload()  # 获取内容
        for n, part in enumerate(parts):
            logger.info('%s part %s', '  '*indent, n)
            logger.info('%s--------', '  '*indent)
            print_info(part, indent+1)
    else:
        content_type = msg.get_content_type()  # 获取消息类型
        if content_type == 'text/plain' or content_type == 'text/html':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset:
                try:
                    content = content.decode(charset)
                except UnicodeDecodeError:
                    logger.info('转由gb18030解码')
                    content = content.decode('GB18030')
                start = content.find(r'合同原件是否已回')
                end = content.find(r'电子邮箱')
                dataCleaning(getForm(content[start:end]))
        else:
            logger.info('%s Attachment:%s', '  '*indent, content_type)

def getForm(infobox):
    '''截取邮件中表格部分'''
    # (?isu)意思是，搜索时包含回车、换行、汉字、空格
    p1 = re.compile(r'(?isu)<tr[^>]*>(.*?)</tr>')
    p2 = re.compile(r'(?isu)<td[^>]*>(.*?)</td>')
    p3 = re.compile(r'<[^>]*>')
    p4 = re.compile(r'\s+')

    dict_1 = {i:[None] for i in columns}
    a_list = ['flag']

    for row in p1.findall(infobox):
        if a_list[-1] == '电话':
            # '限制循环判定一次columns结束'
            break
        for col in p2.findall(row):
            if '公司名称' in col:
                # '表单未统一/自行调整了'
                a_list.append('广告主名称')
                continue
            elif '简体' in col:
                # '部分表单中：广告主名称（简体）'
                continue
            elif '账期' in col and '预付' in col:
                # '账期&预付 各有格式标签'
                a_list.append('账期/预付')
                col = '账期/预付'
            elif len(list(filter(lambda x: x in col, columns))) == 1:
                column = list(filter(lambda x: x in col, columns))[0]
                a_list.append(column)
                continue
            elif len(list(filter(lambda x: x in col, columns))) == 2:
                # '服务费 & 服务费货币 相互影响'
                column = list(filter(lambda x: x in col, columns))[1]
                a_list.append(column)
                continue
            elif col == '\r\n' or col == '\n':
                # '文档中存在大量的换行符'
                continue
            elif col:
                col = p3.sub('', col)
                col = p4.sub('', col)
                col = col.replace('&nbsp;', '')
                col = col.replace('&#43;', '')
                dict_1[a_list[-1]].append(col)
                if a_list[-1] == '电话':
                    # '限制循环判定一次columns结束'
                    break
            else:
                pass
    # '字典长度统一'
    i = max(len(dict_1[x]) for x in columns)
    l = list(filter(lambda x: len(dict_1[x]) < i, columns))
    [dict_1[k].append(None) for k in l for j in range(i-len(dict_1[k]))]
    return dict_1

@cost_time
def dataCleaning(dic):
    '''对邮件抓取到的数据进行清洗'''
    df1 = pd.DataFrame(dic)
    # '军朗 填充'
    try:
        if df1.loc[1, '端口'] == 'baidu-Junlang':
            # '多账户同时申请，特殊处理'
            if len(df1['用户名']) > 2:
                df1.loc[2:, '用户名'] = df1.loc[2:, '端口']
                df1.loc[2:, '端口'] = None
                df1.fillna(method='ffill', inplace=True)
        else:
            pass
    except:
        pass
    # '删除 用户名为空'
    # df1.dropna(axis=0, how='all', inplace=True)
    df1.drop(index=df1[df1['用户名'].isna()].index, inplace=True)
    # '空值 填充'
    df1.fillna('-', inplace=True)
    # '邮件日期 补充'
    global df
    df1['日期'] = date
    df1['flag'] = 'IO'
    for i in df1.columns[1:]:
        df1[i] = df1[i].apply(lambda x: str(x))
    # '合并 去重'
    df1 = normal(df1)
    df = df.append(df1, ignore_index=True, sort=False)
    df.drop_duplicates('用户名', keep='last', inplace=True)

def normal(df):
    '''默认格式
    '''
    df = df.applymap(lambda x: str(x))
    df['日期'] = pd.to_datetime(df['日期'])
    df.sort_values(by='日期', ascending=True, inplace=True)
    return df

def dfNull(dat=None):
    '''构造空行；录入程序运行日日期;提高运行效率
       - 读入前删除标识行
    '''
    import numpy as np
    dff = pd.DataFrame(np.zeros((1,len(columns))), columns=columns)
    if dat == None:
        dff['日期'] = datetime.datetime.now()-datetime.timedelta(hours=1)
    else:
        dff['日期'] = dat
    # 格式化
    dff = normal(dff)
    return dff

def restore(dat=None):
    '''异常恢复/增加标识行'''
    dfNull(dat).to_sql('开户申请表', con=engine, if_exists='append', index=False)

@cost_time
def mainKH(date_0, sec, path):
    '''
    据设置抓取时段，完成邮件抓取
    '''
    try:
        logger.info('Tips: catch the frequency %ss', sec)
        if os.path.exists(path):
            import configparser
            conf = configparser.ConfigParser()
            conf.read(path)
        else:
            raise FileExistsError('file c.s.conf is not exists')
        # 登陆、遍历邮件、解析
        server = poplib.POP3_SSL(conf.get('mail_baidu', 'receiving server'))
        logger.info(server.set_debuglevel(1))
        logger.info(server.getwelcome().decode('utf-8'))
        server.user(conf.get('mail_baidu', 'email'))
        server.pass_(conf.get('mail_baidu', 'password'))
        Message, Size = server.stat()
        logger.info('Message: %s Size: %s', Message, Size)
        resp, mails, octets = server.list()
        index = len(mails)
        for i in range(index, 0, -1):
            time.sleep(sec)
            try:
                resp, lines, octets = server.retr(i)
                msg_content = b'\r\n'.join(lines).decode('utf-8')
                msg = Parser().parsestr(msg_content)
                # '指定日期'
                global date
                date = msg.get('Date')
                date = datetime.datetime.strptime(date[:24], 
                                                  '%a, %d %b %Y %H:%M:%S')
                if date < date_0:
                    break
                sub = decode_str(msg.get('Subject'))
                if '开户进度' in sub:
                    print_info(msg)
                else:
                    logger.info('跳过，非目标文件 %s', sub)
                    continue
            except TimeoutError as e:
                logger.error('访问受限，连接超时 %s', e)
        server.quit()
        # '写入 SQL Server，替换写'
        # 后续变更为只增加新户
        #
        df.to_sql('开户申请表', con=engine, if_exists='replace', index=False)
        restore()
    except FileExistsError as e:
        # 复位
        restore(date_0)
        logger.error(e, exc_info=True)
    except KeyboardInterrupt:
        # 复位
        restore(date_0)
        logger.warning('KeyboardInterrupt', exc_info=True)
    except Exception as e:
        # 复位
        restore(date_0)
        logger.warning('Warning: %s', e, exc_info=True)


if __name__ == '__main__':
    
    # 账号密码 配置文件地址
    path = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
    
    #'据数据库中最近日期判定抓取日期'
    date_0 = engine.execute('''select top 1 日期 from 开户申请表'''
                              ).fetchone()[0]
    mainKH(date_0, 1, path)
        
    pass

