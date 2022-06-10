# -*- codeing = urf-8 -*-
# @Time :2022/4/24
# @Author:LeiXin
# @File : index.py
# @SoftWore : PyCharm
# 十分感谢W01fh4cker大佬，本项目基于W01fh4cker开源项目nuist-classcube-auto-report
# 本程序的所用内容仅供学习交流，请下载后24小时内删除!
# 使用者的所有行为均与制作者无关!使用本程序代表默认上述观点！

import requests
# from apscheduler.schedulers.blocking import BlockingScheduler
import re
import smtplib
import socket
import os
import configparser
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.mime.text import MIMEText
from email.header import Header
from smtplib import SMTP_SSL
session = requests.session()
# 登陆的地址
login_url = 'http://banjimofang.com/student/login?ref=%2Fstudent'
# 填报的链接（需要修改）
add_url = 'http://banjimofang.com/student/course/43535/covid19d3'
ua = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36 Edg/100.0.1185.50'}



# 这里是你的班级魔方账号(直接填写，不用引号）
phonenumber =

# 这里是你的班级魔方密码（直接填写，可以加上引号)
# 忘记了在这里重置找回http://banjimofang.com/student/login?ref=%2Fstudent
password =

# 这里是你的发信人，告诉你程序成功了没
sender =""
# 这里是QQ邮箱的授权码，不明白可以百度
pw =""

# 这里就是你接收有没有成功的QQ邮箱了
receivers =""



def login():
    res = session.get(login_url, headers=ua)
    token = re.findall('\"_token\" value=\"(.+)\"', res.text)[0]
    login_data = {'_token': token, 'username': phonenumber, 'password': password}
    res = session.post(login_url, headers=ua, data=login_data)
    res = session.get(add_url, headers=ua)
    if '微信扫码快速登录' not in res.text:
        login_success = '[√]登录成功!'
        print(login_success)
    else:
        login_fail = '[×]登录失败。请自查原因。'
        print(login_fail)


def work():
    login()
    formdata ={'t1': 36.5, 't2': 36.7, 't3': 36.5}
    res = session.post(add_url, headers=ua, data=formdata)
    print("已经当做成功提交")
    print(res.text)
    if '提交成功' in res.text:
        write_success = '[√]填报成功!'
        print(write_success)
        send_msg_success()

    else:
        write_fail = '[×]填报失败。请自查原因，若确认无误，请联系我'
        print(write_fail)
        send_msg_fail()


def send_msg_success():
    server = "smtp.qq.com"
    port = 465
    machine_name = socket.gethostname()
    msg_root = MIMEMultipart('mixed')
    msg_root['From'] = Header(f'体温填报成功提醒！')
    msg_root['Subject'] = Header(f'Message from {machine_name}', 'utf-8')
    mail_msg = f"""
	<html>

<body>
<p>您好，您现在的体温填报成功啦！</p>
</body>
</html>"""
    msg_root.attach(MIMEText(mail_msg, 'html', 'utf-8'))
    smtp = smtplib.SMTP_SSL(server, port)
    smtp.login(sender, pw)
    smtp.sendmail(sender, receivers, msg_root.as_string())
    smtp.quit()



def send_msg_fail():
    server = "smtp.qq.com"
    port = 465
    machine_name = socket.gethostname()
    msg_root = MIMEMultipart('mixed')
    msg_root['From'] = Header(f'体温填报失败提醒！')
    msg_root['Subject'] = Header(f'Message from {machine_name}', 'utf-8')
    mail_msg = f"""
	<html>

<body>
<p>失败啦！看到这个消息，不出意外是凉凉了！手动填报吧！</p>
</body>
</html>"""
    msg_root.attach(MIMEText(mail_msg, 'html', 'utf-8'))
    smtp = smtplib.SMTP_SSL(server, port)
    smtp.login(sender, pw)
    smtp.sendmail(sender, receivers, msg_root.as_string())
    smtp.quit()


if __name__ == '__main__':
    try:

        work()
        print("su")
    except:
        pass

