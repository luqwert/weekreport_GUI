#!/usr/bin/env python  
# _*_ coding:utf-8 _*_  
# @Author  : lusheng



import email
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr
import time
import mimetypes
import globalmap as gl
# gl._init()

res = [{'name': '金总', 'mail': 'lusheng@sinometalsh.com'},
             ]
report_path = 'C:\\Users\\LUS\\Desktop\\各产品周市场分析.docx'

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((
        Header(name, 'utf-8').encode(),
        addr))

def send_report():
    mail_host="smtp.qq.com"  #设置服务器
    mail_user="228383562@qq.com"    #用户名
    mail_pass="waajnvtmdhiucbef"   #口令
    sender = '228383562@qq.com'
    receiversName = '个人邮箱'
    # receivers = 'lusheng1234@126.com'
    receivers = res[0]['mail']
    now = time.localtime()
    date = time.strftime("%Y-%m-%d", now)
    print(receiversName, receivers)


    text_msg = """
    <p> %s:</p>
    <p> 您好！附件为本周产品市场分析报告，请查收。</p>
    <p>这是周报推送测试... </p>
    """ % res[0]['name']

    main_msg = email.mime.multipart.MIMEMultipart()
    main_msg.attach(MIMEText(text_msg, 'html', 'utf-8'))


    # main_msg['From'] = Header(mail_user)
    # main_msg['To'] = Header(receivers)
    # main_msg['Subject'] = Header('各产品周市场分析报告')
    # main_msg['Date'] = email.utils.formatdate()

    main_msg['From'] = _format_addr(u'芦胜 <%s>' % mail_user)
    main_msg['To'] = _format_addr(u'金总 <%s>' % receivers)
    main_msg['Subject'] = Header(u'各产品周市场分析报告', 'utf-8').encode()


    with open(report_path, 'rb') as f:
        ctype, encoding = mimetypes.guess_type(report_path)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        file_msg = email.mime.base.MIMEBase(maintype, subtype)
        file_msg.set_payload(f.read())
        f.close()
        file_msg.add_header('Content-Disposition', 'attachment', filename=Header(u'各产品周市场分析.docx', 'utf-8').encode())
        email.encoders.encode_base64(file_msg)
        main_msg.attach(file_msg)

    try:
        smtpObj = smtplib.SMTP('smtp.qq.com', 587)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.login(mail_user, mail_pass)
        # print(smtpObj.login(mail_user, mail_pass))
        smtpObj.sendmail(sender, receivers, main_msg.as_string())
        smtpObj.quit()
        print("邮件发送成功")
        result2 = "报告已发送成功"
        gl.set_value('result2', result2)
    except smtplib.SMTPException:
        print("Error: 无法发送邮件")
        result2 = "报告没有发送成功"
        gl.set_value('result2', result2)


# send_report()
