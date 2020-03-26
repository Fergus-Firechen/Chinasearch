# *-* coding:utf-8 _*_

'''
# 构建基表
# 筛选
'''
from sqlalchemy import create_engine
from configparser import ConfigParser
import pandas as pd
import time
import os

now = lambda : time.perf_counter()

def connectDB():
	# 连接数据库
	def getLoginInfo():
		CONF = r'C:\Users\chen.huaiyu\Chinasearch\c.s.conf'
		conf = ConfigParser()
		if os.path.exists(CONF):
			conf.read(CONF)
			section = 'SQL Server'
			sa = conf.get(section, 'accountname')
			pw = conf.get(section, 'password')
			host = conf.get(section, 'ip')
			port = conf.get(section, 'port')
			dbname = conf.get(section, 'dbname')
			return sa, pw, host, port, dbname
	try:
		engine = create_engine(
			"mssql+pymssql://%s:%s@%s:%s/%s" % getLoginInfo())
	except Exception as e:
		print("连接失败： {}".format(e))
	else:
		return engine

def basicInfo(conn):
	# 基本信息 + QTD
	sql = ''' select 区域, 用户名, AM, 广告主, 客户, channel 
		from basicInfo
		'''
	return conn.execute(sql).fetchall()

def qtd(conn):
	# QTD
	sql = ''' select * from QTD_Total
		'''
	return conn.execute(sql).fetchall()

def weekSpending(conn, dateStarBefore, dateEndBefore):
	# 周消费
	# eg. dateStarBefore = -15
	# eg. dateEndBefore = -1
	# class_ = 搜索点击、自主投放，新产品
	sql = ''' select *
		from 消费
	    where 日期 between DATEADD(DD, {}, GETDATE()) and DATEADD(DD, {}, GETDATE())
    	  and 金额 > 0
		  and 类别 in ('搜索点击', '自主投放', '新产品')
		'''.format(dateStarBefore, dateEndBefore)
	return conn.execute(sql).fetchall()

def transform(s, basicInfo, qtd, twoWeeks):
    ' 广告主'
    if s == 'ad':
        lis = ['广告主', '客户', 'AM']
    elif s == 'ag':
        lis = ['客户', 'channel']
    df = nearlyTwoWeeksSpend(basicInfo, qtd, twoWeeks, lis)
    lis1 = ['QTD', '上周P4P消费', '本周P4P消费', '上周搜索', '本周搜索', 
            '上周新产品', '本周新产品', '上周排名',  '环比P4P消费', 
            '环比P4P消费增长率']
    df = df.reindex(columns=lis + lis1)
    return df

def nearlyTwoWeeksSpend(basicInfo_, qtd, nearlyTwoWeeksSpending, lis):
    ''' 近两周 消费 
    1.近两周有消费
    2.
    '''
    # 基本信息 + qtd
    df = basicInfo_.merge(qtd, how='left', on='用户名')
    df.sort_values('QTD', inplace=True, ascending=False)
    if '广告主' in lis:
        for ad in df['广告主'].unique():
            df.loc[df['广告主'] == ad, ['客户']] = df.loc[df['广告主'] == ad, 
                   ['客户']].values[0]
            df.loc[df['广告主'] == ad, ['AM']] = df.loc[df['广告主'] == ad, 
                   ['AM']].values[0]
    elif '客户' in lis:
        for ag in df['客户'].unique():
            df.loc[df['客户'] == ag, ['channel']] = df.loc[df['客户'] == ag, 
                   ['channel']].values[0]
    # + 近两周消费
    df = df.merge(transWeek(nearlyTwoWeeksSpending, 6), how='left', on='用户名')
    df = df.merge(transWeek(nearlyTwoWeeksSpending, 13), how='left', on='用户名')
    df.fillna(0, inplace=True)
    # transform
    df['上周P4P消费'] = df['上周搜索'] + df['上周新产品']
    df['本周P4P消费'] = df['本周搜索'] + df['本周新产品']
    df['环比P4P消费'] = df['本周P4P消费'] - df['上周P4P消费']
    # delete sp = 0
    df = df.groupby(lis).sum()
    df.reset_index(inplace=True)
    df = df[(df['上周P4P消费'] + df['本周P4P消费']) > 0]
    # 环比增长率
    df['环比P4P消费增长率'] = df['环比P4P消费'] / df['上周P4P消费']
    df.loc[df['上周P4P消费'] == 0, '环比P4P消费增长率'] = 1
    # 上周消费排名
    df.sort_values('上周P4P消费', ascending=False, inplace=True)
    df.reset_index(inplace=True, drop=True)
    df['上周排名'] = list(map(lambda x: x+1, df.index))
    # 近两周消费排名
    df['flag'] = df['上周P4P消费'] + df['本周P4P消费']
    df.sort_values('flag', ascending=False, inplace=True)
    df.reset_index(inplace=True, drop=True)
    df.index = list(map(lambda x: x+1, df.index))
    df.drop(columns=['flag'], inplace=True)
    return df

# Top15
lis = ['上周P4P消费', '本周P4P消费', '环比P4P消费']
# Top15_1 总
Top15_1 = lambda x, y: ADSpendingRingRatio(x, y)[lis].sum()
# Top15_2 直客 top15
Top15_2 = lambda x, y: DSMaster(x, y)[['广告主'] + lis][:15]
# Top15_3 代理 top5
Top15_3 = lambda x, y: Agency(x, y)[['客户'] + lis][:5]

def GroupAM(basicInfo_, df):
    ' SZ AM 汇总 '
    data = ADSpendingRingRatio(basicInfo_, df)
    data = data.pivot_table(values=['上周P4P消费', '本周P4P消费', '环比P4P消费'],
                            index=['AM'], aggfunc=sum)
    return data
    
def Agency(basicInfo_, df):
    ' SZ 代理消费环比 '
    data = weeklyRingRatio(basicInfo_, df)
    lis1 = ['客户', '上周新产品', '本周新产品', '环比新产品', '上周搜索', 
           '本周搜索', '环比搜索']
    data = data.loc[data['channel'].isin(['代理商']), lis1]
    data['上周P4P消费'] = data['上周搜索'] + data['上周新产品']
    data['本周P4P消费'] = data['本周搜索'] + data['本周新产品']
    data['环比P4P消费'] = data['本周P4P消费'] - data['上周P4P消费']
    data = data.groupby('客户').sum()
    data.reset_index(inplace=True)
    lis = []
    for i in data['客户'].str.findall('[^a-zA-Z.,，]+$'):
        if len(i) == 0:
            lis.append('-')
        else:
            lis.append(i[0])
    data['客户'] = lis
    data.drop(index=data[data['客户'].isin(['-'])].index, inplace=True)
    data.sort_values('本周P4P消费', ascending=False, inplace=True)
    data.reset_index(inplace=True, drop=True)
    data.index = list(map(lambda x: x+1, data.index))
    return data

def DSMaster(basicInfo_, df):
    ' SZ 直客消费环比 '
    ds = ADSpendingRingRatio(basicInfo_, df)
    ds.sort_values('本周P4P消费', ascending=False, inplace=True)
    ds = ds[ds['channel'].isin(['直接客户'])]
    ds.reset_index(inplace=True, drop=True)
    ds.index = list(map(lambda x: x+1, ds.index))
    return ds

def ADSpendingRingRatio(basicInfo_, df):
    ' SZ 广告主消费环比 '
    # 深圳 AM
    lis1 = ['作废', '公司备用', '开户专用', '-']
    lis2 = ['AM', '广告主', 'channel', '用户名']
    basicInfo_ = basicInfo_.loc[basicInfo_['区域'].isin(['SZ']) &
                                ~basicInfo_['AM'].isin(lis1), lis2]
    # 周环费 & 环比
    week = basicInfo_.merge(transWeek(df, 6), how='left', on='用户名')
    week = week.merge(transWeek(df, 13), how='left', on='用户名')
    week.fillna(0, inplace=True)
    week['环比搜索'] = week['本周搜索'] - week['上周搜索']
    week['环比新产品'] = week['本周新产品'] - week['上周新产品']
    lis3 = ['AM', '广告主', 'channel', '上周新产品', '本周新产品', '环比新产品', 
            '上周搜索', '本周搜索', '环比搜索']
    week = week.reindex(columns=lis3)
    week['上周P4P消费'] = week['上周搜索'] + week['上周新产品']
    week['本周P4P消费'] = week['本周搜索'] + week['本周新产品']
    week['环比P4P消费'] = week['本周P4P消费'] - week['上周P4P消费']
    # 近两周 有消费
    # Result
    week = week[week['上周P4P消费'] + week['本周P4P消费'] > 0].groupby(['AM',
               '广告主', 'channel']).sum()
    week.reset_index(inplace=True)
    week.sort_values('环比P4P消费', inplace=True)
    return week

def weeklyRingRatio(basicInfo_, df):
    ' SZ 周环比 '
    # 深圳
    
    lis = ['作废', '公司备用', '开户专用', '-', '陈熙香', '顾凡凡']
    basicInfo_ = basicInfo_[basicInfo_['区域'].isin(['SZ']) &
                                ~basicInfo_['AM'].isin(lis)]
    # 搜索 近两周日消费明细
    dailySpending = df[df['类别'].isin(['搜索点击'])]
    dailySpending = dailySpending.pivot_table(
            values=['消费'], index=['用户名'], columns=['日期'])
    dailySpending.columns = dailySpending.columns.get_level_values(1)
    # 周消费 及 环比
    week = basicInfo_.merge(transWeek(df, 6), how='left', on='用户名')
    week = week.merge(transWeek(df, 13), how='left', on='用户名')
    week.fillna(0, inplace=True)
    week['环比搜索'] = week['本周搜索'] - week['上周搜索']
    week['环比新产品'] = week['本周新产品'] - week['上周新产品']
    # 有消费
    week = week[week['上周搜索'] + week['上周新产品'] + week['本周搜索'] +
                week['本周新产品'] > 0]
    lis2 = week.columns.tolist()[:6] + ['上周新产品', '本周新产品', '环比新产品', 
                              '上周搜索', '本周搜索', '环比搜索']
    week = week.reindex(columns=lis2)
    # Result
    week = week.merge(dailySpending, how='left', on='用户名')
    return week

def transWeek(df, num):
    dates = df['日期'].unique()
    dates.sort()
    # 筛选 汇总
    # 新产品 = np + inf
    if num == 6:
        df_W = df.loc[df['日期'] <= dates[num], ['用户名', '类别', '消费']
            ].groupby(['用户名', '类别']).sum()
        df_W = df_W.pivot_table(values=['消费'], index=['用户名'], columns=['类别'])
        df_W.reset_index(inplace=True)
        df_W.columns = ['用户名', '上周搜索', '上周新产品', '上周原生']
        df_W.fillna(0, inplace=True)
        df_W['上周新产品'] = df_W['上周新产品'] + df_W['上周原生']
        df_W.drop(columns=['上周原生'], inplace=True)
    elif num == 13:
        df_W = df.loc[df['日期'] > dates[num//2], ['用户名', '类别', '消费']
            ].groupby(['用户名', '类别']).sum()
        df_W = df_W.pivot_table(values=['消费'], index=['用户名'], columns=['类别'])
        df_W.reset_index(inplace=True)
        df_W.columns = ['用户名', '本周搜索', '本周新产品', '本周原生']
        df_W.fillna(0, inplace=True)
        df_W['本周新产品'] = df_W['本周新产品'] + df_W['本周原生']
        df_W.drop(columns=['本周原生'], inplace=True)
    return df_W
    
def getPath(df):
    ' 获取地址 '
    dates = df['日期'].unique()
    dates.sort()
    path = r'C:\Users\chen.huaiyu\Desktop\Output'
    name1 = 'P4P 消费周环比(' + dates[0] + '_' + dates[-1] + ').xlsx'
    name2 = '近两周代理商消费(' + dates[0] + '_' + dates[-1] + ').xlsx'
    name3 = '近两周广告主消费(' + dates[0] + '_' + dates[-1] + ').xlsx'
    name4 = 'Top 50广告主(' + dates[0] + '_' + dates[-1] + ').xlsx'
    # Result
    path1 = os.path.join(path, name1)
    path2 = os.path.join(path, name2)
    path3 = os.path.join(path, name3)
    path4 = os.path.join(path, name4)
    return path1, path2, path3, path4

def sendMail(subject, dat, message, fils):
    import smtplib
    from email import encoders
    from email.header import Header
    from email.mime.text import MIMEText
    from email.utils import parseaddr, formataddr
    from email.mime.multipart import MIMEMultipart, MIMEBase
    
    def _format_addr(s):
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode(), addr))
    
    def login():
        PATH = r'c:\users\chen.huaiyu\chinasearch\c.s.conf'
        conf = ConfigParser()
        if os.path.exists(PATH):
            conf.read(PATH)
            smt = conf.get('mail_baidu', 'sender server')
            from_addr = conf.get('mail_baidu', 'email')
            pw = conf.get('mail_baidu', 'password')
            to_addr = conf.get('newIOSystem', 'to_addr')
            return smt, from_addr, pw, to_addr
    
    msg = MIMEMultipart()
    msg['From'] = login()[1]
    msg['To'] = login()[3]
    msg['Subject'] = Header(subject, 'utf-8').encode()

    msg.attach(MIMEText(message, 'plain', 'utf-8'))
    for path in fils:
        print(path)
        
        with open(path, 'rb') as f:
            mime = MIMEBase('text', 'plain', filename=os.path.split(path)[-1])
            # 加上必要的头信息
            mime.add_header('Content-Disposition', 'attachment', filename=os.path.split(path)[-1])  # 1.内容传输编码;2.;3.文件名
            mime.add_header('Content-ID', '<0>')
            mime.add_header('X-Attachment-Id', '0')
            # add 附件
            mime.set_payload(f.read())
            encoders.encode_base64(mime)
            msg.attach(mime)
    
    server = smtplib.SMTP(login()[0], 25)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.set_debuglevel(1)
    server.login(login()[1], login()[2])
    try:
        server.sendmail(login()[1], login()[3].split(','), msg.as_string())
    except Exception as e:
        print('Failed send: {}'.format(e))
    else:
        print('Success send.')
    server.quit()


def main():
    # 主程序
    print("main")
    # star = int(input('起始日:(两周前：-15)'))
    # end = int(input('终止日:(昨日：-1)'))
    star = -15
    end = -1
    if end - star == 14:
        col1 = ['区域', '用户名', 'AM', '广告主', '客户', 'channel']
        col2 = ['日期', '用户名', '类别', '消费']
        # 连接数据库 获取数据
        with connectDB().begin() as conn:
            df_basicInfo = pd.DataFrame(basicInfo(conn), columns=col1)
            df_qtd = pd.DataFrame(qtd(conn), columns=['用户名', 'QTD'])
            df_two_weeks = pd.DataFrame(weekSpending(conn, star, end), columns=col2)
        # 格式统一
        df_basicInfo['客户'] = df_basicInfo['客户'].str.lower()
        df_basicInfo['客户'] = df_basicInfo['客户'].str.title()
        df_basicInfo['客户'] = df_basicInfo['客户'].str.replace(' ', '')
        df_basicInfo['广告主'] = df_basicInfo['广告主'].str.lower()
        df_basicInfo['广告主'] = df_basicInfo['广告主'].str.title()
        df_basicInfo['广告主'] = df_basicInfo['广告主'].str.replace(' ', '')
        # Output
        ## 近两周消费周环比
        # 1.近两周有消费
        # 2.SZ AM
        #
        p1, p2, p3, p4 = getPath(df_two_weeks)
        t1 = Top15_1(df_basicInfo, df_two_weeks)
        t2 = Top15_2(df_basicInfo, df_two_weeks)
        t2.index.name = '直客'
        t3 = Top15_3(df_basicInfo, df_two_weeks)
        t3.index.name = '代理'
        groupAM = GroupAM(df_basicInfo, df_two_weeks)
        adRingRatio = ADSpendingRingRatio(df_basicInfo, df_two_weeks)
        ds = DSMaster(df_basicInfo, df_two_weeks)
        ag = Agency(df_basicInfo, df_two_weeks)
        week = weeklyRingRatio(df_basicInfo, df_two_weeks)
        #
        ## P4P消费周环比
        #
        with pd.ExcelWriter(p1, engine='xlsxwriter') as path:
            t1.to_excel(path, sheet_name='Top 15', header=False)
            t2.to_excel(path, sheet_name='Top 15',startrow=6)
            t3.to_excel(path, sheet_name='Top 15', startrow=24)
            groupAM.to_excel(path, sheet_name='汇总')
            adRingRatio.to_excel(path, sheet_name='广告主消费环比', 
                                 freeze_panes=(1, 0), index=False)
            ds.to_excel(path, sheet_name='直客消费环比', freeze_panes=(1, 0), index=False)
            ag.to_excel(path, sheet_name='代理商消费环比', freeze_panes=(1,0), index=False)
            week.to_excel(path, sheet_name='周环比', freeze_panes=(1, 0), index=False)
            #
            # Top 15
            #
            wb = path.book
            sht = path.sheets['Top 15']
            fmt1 = wb.add_format({'num_format': '#,##0'})
            fmt2 = wb.add_format({'num_format': '0.00%'})
            fmt3 = wb.add_format({'num_format': '#,##0', 'bold': True, 'border': 1,
                                  'align': 'center'})
            fmt4 = wb.add_format({'bg_color': '#ffc7ce'})
            cond_fmt = {'type':'cell', 'criteria':'less than', 'value':0, 'format':fmt4}
            sht.set_column('A:A', 12)
            sht.set_column('B:B', 45, fmt1)
            sht.set_column('C:H', 12, fmt1)
            sht.set_column('I:I', 25, fmt1)
            ## 直客
            sht.write('F7', '占比本周消费', fmt3)
            sht.write('G7', '预估季度消费', fmt3)
            sht.write('H7', '浮动原因', fmt3)
            sht.write('I7', '预估季度消费(上周提供)', fmt3)
            sht.conditional_format('E8:E22', cond_fmt)
            for i in range(8, 23):
                sht.write('F' + str(i), '=D' + str(i) + '/B2', fmt2)
            ## 代理
            sht.write('F25', '占比本周消费', fmt3)
            sht.write('G25', '预估季度消费', fmt3)
            sht.write('H25', '浮动原因', fmt3)
            sht.write('I25', '预估季度消费(上周提供)', fmt3)
            sht.conditional_format('E26:E30', cond_fmt)
            for i in range(26, 31):
                sht.write('F' + str(i), '=D' + str(i) + '/B2', fmt2)
            #
            # 汇总
            #
            sht2 = path.sheets['汇总']
            sht2.set_column('B:D', 15, fmt1)
            sht2.write(groupAM.shape[0]+1, 0, '合计', fmt3)
            sht2.write(groupAM.shape[0]+1, 1, '=SUM(B2:B5)')
            sht2.write(groupAM.shape[0]+1, 2, '=SUM(C2:C5)')
            sht2.write(groupAM.shape[0]+1, 3, '=SUM(D2:D5)')
            #
            # 广告主消费环比
            #
            sht3 = path.sheets['广告主消费环比']
            sht3.set_column('A:L', 12, fmt1)
            sht3.set_column('B:B', 25)
            sht3.conditional_format('L2:L'+str(adRingRatio.shape[0]+1), cond_fmt)
            #
            # 直客消费环比
            #
            sht4 = path.sheets['直客消费环比']
            sht4.set_column('A:L', 12, fmt1)
            sht4.set_column('B:B', 25)
            sht4.conditional_format('L2:L'+str(ds.shape[0]+1), cond_fmt)
            #
            # 代理商消费环比
            #
            sht5 = path.sheets['代理商消费环比']
            sht5.set_column('B:J', 12, fmt1)
            sht5.set_column('A:A', 25)
            sht5.conditional_format('D1:D'+str(ag.shape[0]+1), cond_fmt)
            sht5.conditional_format('G1:G'+str(ag.shape[0]+1), cond_fmt)
            sht5.conditional_format('J1:J'+str(ag.shape[0]+1), cond_fmt)
            #
            # 周环比
            #
            sht6 = path.sheets['周环比']
            sht6.set_column('A:W', 12, fmt1)
            sht6.set_column('D:E', 25)
            sht6.conditional_format('I1:I'+str(week.shape[0]+1), cond_fmt)
            sht6.conditional_format('L1:L'+str(week.shape[0]+1), cond_fmt)
        #
        ## 近两周 广告主消费
        #
        ad = transform('ad', df_basicInfo, df_qtd, df_two_weeks)
        ad.index.name = '序号'
        with pd.ExcelWriter(p3, engine='xlsxwriter') as path1:
            ad.to_excel(path1, sheet_name='近两周广告主消费', freeze_panes=(1,0))
            # 格式调整
            wb = path1.book
            sht = path1.sheets['近两周广告主消费']
            fmt1 = wb.add_format({'num_format': '#,##0'})
            fmt2 = wb.add_format({'num_format': '0%'})
            fmt3 = wb.add_format({'num_format': '#,##0;[Red](#,##0);'})
            dic = {'type': 'icon_set',
                   'icon_style': '3_arrows',
                   'icons': [{'criteria': '>', 'type': 'number', 'value': 0},
                             {'criteris': '<', 'type': 'number', 'value':0}]}
            sht.set_column('B:C', 25)
            sht.set_column('D:L', 11, fmt1)
            sht.set_column('N:N', 19, fmt2)
            sht.set_column('M:M', 11, fmt3)
            sht.conditional_format('M2:M'+str(ad.shape[0]+1), dic)
        #
        ## Top 50
        #
        ad50 = ad[:50]
        with pd.ExcelWriter(p4, engine='xlsxwriter') as path3:
            ad50.to_excel(path3, sheet_name='Top50', freeze_panes=(1,0))
            # 格式调整
            wb = path3.book
            fmt1 = wb.add_format({'num_format': '#,##0'})
            fmt2 = wb.add_format({'num_format': '0%'})
            fmt3 = wb.add_format({'num_format': '#,##0;[Red](#,##0);'})
            dic = {'type': 'icon_set',
                   'icon_style': '3_arrows',
                   'icons': [{'criteria': '>', 'type': 'number', 'value': 0},
                             {'criteris': '<', 'type': 'number', 'value':0}]}
            sht = path3.sheets['Top50']
            sht.set_column('B:C', 25)
            sht.set_column('D:L', 11, fmt1)
            sht.set_column('N:N', 19, fmt2)
            sht.set_column('M:M', 11, fmt3)
            sht.conditional_format('M2:M'+str(ad50.shape[0]+1), dic)
        #
        ## 近两周 代理消费
        #
        ag = transform('ag', df_basicInfo, df_qtd, df_two_weeks)
        ag.index.name = '序号'
        with pd.ExcelWriter(p2, engine='xlsxwriter') as path2:
            ag.to_excel(path2, sheet_name='近两周代理消费', freeze_panes=(1,0))
            # 格式调整
            wb = path2.book
            sht = path2.sheets['近两周代理消费']
            fmt1 = wb.add_format({'num_format': '#,##0'})
            fmt2 = wb.add_format({'num_format': '0%'})
            fmt3 = wb.add_format({'num_format': '#,##0;[Red](#,##0);'})
            dic = {'type': 'icon_set', 
                   'icon_style': '3_arrows',
                   'icons': [{'criteria': '>', 'type': 'number', 'value': 0},
                             {'criteria': '<', 'type': 'number', 'value': 0}]}
            sht.set_column('B:B', 35)
            sht.set_column('C:K', 11, fmt1)
            sht.set_column('M:M', 19, fmt2)
            sht.set_column('L:L', 15, fmt3)
            sht.conditional_format('L2:L'+str(ag.shape[0]+1), dic)
        #
        ## 发送
        #
        dat = os.path.split(p1)[-1][-9:-5]
        note = '''
   
祝好
   
林婷
BP | Shen Zhen
TEL:(86)755 25020862-818 |Mobile：(86)13148704556
地址：深圳市罗湖区南东路5002号信兴广场主楼地王大厦5903-06室
            '''
        subject = 'P4P消费周环比(' + dat + ')'
        mes = '''Dear all,
       
    附件为: {}，请查阅。
如有任何疑问，可随时和我联系，谢谢。
            '''.format(subject) + note
        fils = [p1]
        sendMail(subject, dat, mes, fils)
        #
        ## 
        subject = '近两周广告主 || 代理商消费 & TOP 50广告主(' + dat + ')'
        mes = '''Dear all,
       
    附件为: {}，请查阅。
如有任何疑问，可随时和我联系，谢谢。
            '''.format(subject) + note
        fils = [p2, p3, p4]
        sendMail(subject, dat, mes, fils)


if __name__ == '__main__':
    try:
        st = now()
        main()
    except Exception as e:
        print('Error: {}'.format(e))
    finally:
        print('Run time: {:3f} min'.format((now() - st)/60))