#!/usr/bin/env python
#coding: utf-8
import string
import xlsxwriter
import MySQLdb
import time
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


try:
    res=[[0 for col in range(0)] for row in range(16)]
    ser=[]
    index=0
    ISOTIMEFORMAT='%Y-%m-%d %X'
    domainlist=['sic-ca.com','cssca.com','cwindow.net','syncapital.com','docmail.cn','nbvesen.cn','jctvcm.com']
    before7day=datetime.date.today()+datetime.timedelta(-7)
    today=datetime.date.today()
    etime=time.strftime( ISOTIMEFORMAT, time.localtime() )
    stime="%s 00:00:00" % before7day
    for domain in domainlist:
        search="select count(a1.num_login), sum(a1.num_login),s1.num_send, s2.num_delivery from (select count(a.user) num_login from access_logs a where a.user like concat('%%@', '%s') and STR_TO_DATE(a.In_time, '%%b %%e %%H:%%i:%%S %%Y')>'%s' and STR_TO_DATE(a.In_time, '%%b %%e %%H:%%i:%%S %%Y')<'%s' group by a.user) a1,(select count(s.newmsg_id) num_send from smtp_logs s where s.sender_domain='%s' and s.delivery_success='yes' and s.time_stamp>'%s' and s.time_stamp<'%s') s1,(select count(s.newmsg_id) num_delivery from smtp_logs s where s.delivery_domain='%s' and s.delivery_success='yes' and s.time_stamp>'%s' and s.time_stamp<'%s') s2;" % (domain,stime,etime,domain,stime,etime,domain,stime,etime)
        conn=MySQLdb.connect(host='192.168.80.5',user='xxxxxxxx',passwd='xxxxxxxxxxx',db='cwinstats',charset='utf8',port=3306)
        cur=conn.cursor()
        cur.execute(search)
        result=cur.fetchone()
        for i in result:
            res[index].append(i)
        index +=1
        cur.close()
        conn.close()
    
except MySQLdb.Error,e:
     print "Mysql Error %d: %s" % (e.args[0], e.args[1])


workbook = xlsxwriter.Workbook('/root/operation_data.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:G',15)
chart = workbook.add_chart({'type':'column'})
title = [u'业务名称','sic-ca.com','cssca.com','cwindow.net','syncapital.com','docmail.cn','nbvesen.cn','jctvcm.com']
buname = [u'活跃用户',u'登录次数',u'发信次数',u'收信次数']

data = [[0 for col in range(7)] for row in range(4)]
for i in range(0,len(res)):
    for j in range(0,len(res[i])):
        data[j][i]=res[i][j]


#for i in res:
#    for j in i:
#        ser.append(j)
#for i in ser:
#        data[index].append(i)
#    index +=1
    
#data = [
#        [150,152,158,149,155,145,148,123,123,123,123,123,123,123,123,123],
#        [89,88,95,93,98,100,99,123,123,123,123,123,123,123,123,123],
#        [201,200,198,175,170,198,195,123,123,123,123,123,123,123,123,123],
#        [75,77,78,78,74,70,79,123,123,123,123,123,123,123,123,123],
#]

format = workbook.add_format()
format.set_border(1)
format.set_align('center')

format_title = workbook.add_format()
format_title.set_border(1)
format_title.set_bg_color('#cccccc')
format_title.set_align('center')
format_title.set_bold()

format_ave = workbook.add_format()
format_ave.set_border(1)
format_ave.set_num_format('0.00')
format_ave.set_align('center')

worksheet.write_row('A1',title,format_title)
worksheet.write_column('A2',buname,format)
worksheet.write_row('B2',data[0],format)
worksheet.write_row('B3',data[1],format)
worksheet.write_row('B4',data[2],format)
worksheet.write_row('B5',data[3],format)

def chart_series(cur_row):
   # worksheet.write_formula('I'+cur_row,'=AVERAGE(B'+cur_row+':H'+cur_row+')',format_ave)
    chart.add_series({
        'categories':'=Sheet1!$B$1:$Q$1',
        'values':   '=Sheet1!$B$'+cur_row+':$Q$'+cur_row,
        'line':     {'color':'black'},
        'name': 'Sheet1!$A$'+cur_row,
            
    })

for row in range(2,6):
    chart_series(str(row))

chart.set_size({'width':900,'height':287})
chart.set_title({'name':u'邮件运营数据周报图表 %s~%s' % (before7day,today)})
chart.set_y_axis({'name':u'次数'})

worksheet.insert_chart('A8',chart)
workbook.close()

HOST = "mail.sic-ca.com"
SUBJECT = u'IDC运营数据周报'
Text = u'各位领导，大家好，附件里是本周IDC环境各域的运营数据详情，请查阅.'
TO1 = "6734957@qq.com,yongzhi.fu@sic-ca.com"
TO = string.splitfields(TO1, ",")
FROM = "operation@sic-ca.com"
msg = MIMEMultipart('related')
#msgtext = MIMEText(Text,format,'utf-8')
msgtext = MIMEText(Text,'plain','utf-8')
attach = MIMEText(open("/root/operation_data.xlsx","rb").read(), "base64", "utf-8")
attach["Content-Type"] = "application/octet-stream"
attach["Content-Disposition"] = "attachment; filename=\"运营数据周报.xlsx\""
msg.attach(attach)
msg.attach(msgtext)
msg['Subject'] = SUBJECT
msg['From'] = FROM
msg['To'] = TO1


#BODY = string.join((
#    "From: %s" % FROM,
#    "To: %s" % TO,
#    "Subject: %s" % SUBJECT,
#    "",
#    text
#    ),"\r\n")
try:
    server = smtplib.SMTP()
    server.connect(HOST,"25")
    server.starttls()
    server.login("operation@sic-ca.com","xxxxxxxxx")
    server.sendmail(FROM, TO,msg.as_string())
    server.quit()
    print "邮件发送成功！"
except Exception, e:
    print "失败："+str(e)
#HOST = "mail.sic-ca.com"
#SUBJECT = u'IDC运营数据周报'
#TO1 = "6734957@qq.com,yongzhi.fu@sic-ca.com"
#TO = string.splitfields(TO1, ",")
#FROM = "yongzhi.fu@sic-ca.com"
#Text = "各位领导，大家好，附件里是本周IDC环境各域的运营数据详情，请查阅。"
#msgtext = MIMEText(Text,'plain','utf-8')
#msg = MIMEMultipart('related')
#attach = MIMEText(open("operation_data.xlsx","rb").read(), "base64", "utf-8")
#attach["Content-Type"] = "application/octet-stream"
#attach["Content-Disposition"] = "attachment; filename=\"operation_data.xlsx\""
#msg.attach(attach)
#msg.attach(msgtext)
#msg['Subject'] = SUBJECT
#msg['From'] = FROM
#msg['To'] = TO1
#msg["Accept-Language"]="zh-CN"
#msg["Accept-Charset"]="ISO-8859-1,utf-8"
#BODY = string.join((
#    "From: %s" % FROM,
#    "To: %s" % TO,
#    "Subject: %s" % SUBJECT,
#    "",
#    text
#    ),"\r\n")
#try:
#    server = smtplib.SMTP()
#    server.connect(HOST,"25")
#    server.starttls()
#    server.sendmail(FROM, [TO],msg.as_string())
#    server.quit()
#    print "邮件发送成功！"
#except Exception, e:
#    print "失败："+str(e)
