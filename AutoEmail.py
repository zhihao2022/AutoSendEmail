import smtplib
# 负责构造文本
from email.mime.text import MIMEText
# 负责将多个对象集合起来
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.application import MIMEApplication
import os
import time

message_settings = {}

with open('信息设置.txt', 'r',encoding='utf-8') as f:
    Key = ''
    if_in = False
    for line in f.readlines():
        if line == '\n':
            continue
        line = line.strip('\n')
        if not if_in:
            Key = line
            if_in = True
        else:
            message_settings[Key] = line
            if_in = False

receivers = []

with open(message_settings['收件人信息(txt文件):'], 'r', encoding='utf-8') as f:
    if_continue = True
    for line in f.readlines():
        if line == '\n':
            continue
        if '开始' in line:
            if_continue = False
            continue
        if if_continue:
            continue
        line = line.strip('\n')
        tmp = line.split()
        name = ' '.join(tmp[:-1])
        email_line = tmp[-1]
        name_email = [name, email_line]
        receivers.append(name_email)

receivers_num = len(receivers)

Head = ''
words = ''

with open(message_settings['邮件文案(txt文件):'],'r',encoding='utf-8') as f:
    if_head = False
    if_words = False
    for line in f.readlines():
        if line == '标题\n':
            if_head = True
            continue
        if line == '正文\n':
            if_head = False
            if_words = True
            continue
        if if_head:
            Head += line
            continue
        if if_words:
            words += line

Head = Head.strip('\n')
words = words.strip('\n')

Keys = []

with open(message_settings['替换文件名称:'], 'r', encoding='utf-8') as f:
    for line in f.readlines():
        Keys.append(line.strip('\n'))
    if '' in Keys:
        Keys = Keys.remove('')

# SMTP服务器,这里使用qq邮箱
mail_host = "smtp.qq.com"
# 发件人邮箱
mail_sender = message_settings['发件人邮箱:']
# 邮箱授权码,注意这里不是邮箱密码,如何获取邮箱授权码,请看本文最后教程
mail_license = message_settings['邮箱STMP服务授权码:']
# 收件人邮箱，可以为多个收件人
mail_receivers = [rec[1] for rec in receivers]

receivers_str = [rec[0]+'<'+rec[1]+'>' for rec in receivers]

# # 构造附件
# print('正在打开附件...')
# atta = MIMEApplication(open(message_settings['附件名称:'], 'rb').read(), 'zip')
# # 设置附件信息
# atta.add_header('Content-Disposition', 'attachment', filename=message_settings['附件名称:'])
# print('附件打开成功')

# 创建SMTP对象
stp = smtplib.SMTP_SSL(mail_host)

# 设置发件人邮箱的域名和端口，端口地址为25
stp.connect(mail_host, 465)
# set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
# stp.set_debuglevel(1)
# 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码
stp.login(mail_sender, mail_license)

for i in range(receivers_num):
    mm = MIMEMultipart('related')

    # 邮件主题
    subject_content = Head
    # 设置发送者,注意严格遵守格式,里面邮箱为发件人邮箱
    mm["From"] = message_settings['发件人姓名:']+'<'+message_settings['发件人邮箱:']+'>'
    # 设置接受者,注意严格遵守格式,里面邮箱为接受者邮箱
    mm["To"] = receivers_str[i]


    # 设置邮件主题
    mm["Subject"] = Header(subject_content, 'utf-8')

    # 邮件正文内容
    # body_content = """你好，这是一个测试邮件！"""
    tmp = Keys[i*3]+'\n'+Keys[i*3+1]+'\n'+Keys[i * 3+2]
    body_content = words.replace('<ActivationCodes>', tmp)
    # 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
    message_text = MIMEText(body_content,"plain","utf-8")
    # 向MIMEMultipart对象中添加文本对象
    mm.attach(message_text)

    with open("Keys_log.txt",'a',encoding='utf-8') as f:
        f.write(tmp+'\n')

    #
    # # 添加附件到邮件信息当中去
    # mm.attach(atta)



    # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
    print('正在给%s发送邮件...'%receivers_str[i])
    t1 = time.time()
    with open('log.txt','a',encoding='utf-8') as f:
        f.write('正在给%s发送邮件...\n'%receivers_str[i])
    stp.sendmail(mail_sender, mail_receivers[i], mm.as_string())
    with open('log.txt','a',encoding='utf-8') as f:
        f.write("给%s的邮件发送成功\n\n"%receivers_str[i])
    print("给%s的邮件发送成功"%receivers_str[i])
    print('用时%s秒'%str(time.time() - t1))
    print('休息0.5秒')
    time.sleep(0.5)

# 关闭SMTP对象
stp.quit()

with open(message_settings["替换文件名称:"],'w',encoding='utf-8') as f:
    for i in range(receivers_num):
        f.write(Keys[i*3]+'\n'+Keys[i*3+1]+'\n'+Keys[i * 3+2]+'\n')
    f.write('-'*10+'这是分界线'+'-'*10+'\n')
    for i in range(receivers_num*3, len(Keys)):
        f.write(Keys[i]+'\n')

os.system('pause')

