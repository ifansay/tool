# -*- coding: utf-8 -*-
"""
Created on Mon Oct 16 22:09:20 2017

@title:mail self—send

@author: ifansay

@email: ifansay.chn@qq.com
"""
# \033[显示方式;字体色;背景色m......[\033[0m]
# error:red background
# waring:yellow background
# normal:write background

import os
import os.path
import uuid
import datetime
import time
import shutil
import smtplib
import winsound
import pyttsx3
import zipfile
import pythoncom
import configparser as cp

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email.utils import parseaddr, formataddr
from email import encoders
from lxml import etree


"""
part1
"""

# 配置文件路径
print('^_^欢迎使用智文邮件助手^_^\n')
version = '2.6.7'
update_date = '2019/04/14'

config = cp.RawConfigParser()
file_path, code_file = os.path.split(os.path.realpath(__file__))

config.read(file_path+'\\sf_config.ini', encoding='utf-8')
sf_config = config['DEFAULT']
path_project = sf_config['project']
path_file = sf_config['file']
path_filed = sf_config['filed']
path_mail = sf_config['mail']
path_mailed = sf_config['mailed']

# 读取收件人配置
f = open(path_project+'\\recipients.txt', 'r', encoding='utf-8')
addr_dict, j, error_recipients, error_email = {}, 1, [], []
for i in f:
    try:
        i = i.strip('\n').split('/')
        addr_dict[i[0], i[1]] = i[2]
        # ^[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$ email check
    except IndexError:
        print('❤第%d行%s错误' % (j, i))
    j += 1
f.close()

cate = [i[0] for i in addr_dict if i[1] == 'all']
for i in cate:
    check = [g[1] for g in addr_dict if g[0] == i]
    if len(check) > 1:
        print('\033[1;33;41merror:收件人类型%s错误\033[0m,请勿使用或修改后使用' % i)
        error_recipients.append(i)

# 读取附件列表
os.chdir(path_mail)
file_list = os.listdir('.')
file_dict, public_file, file_set = {}, [], set()
for file in file_list:
    if file not in ('desktop.ini', '.DS_Store',):
        try:
            city = file[0:file.index("❤")]
            file_set.add(file[file.index("❤")+1:])
            file_dict[city] = file_dict.get(city, [])+[file]
        except BaseException:
            public_file.append(file)
            file_set.add(file)
print("收件人为:\033[1;34;47m%s\033[0m" % set(file_dict), ",附件:\033[1;34;47m%s\033[0m" % file_set, end='')


"""
part2
"""


def public(public_file):
    if len(public_file) > 0:
        print(';其中\033[1;35;47m%s\033[0m作为公共附件发送' % public_file)
        for city in file_dict:
            file_dict[city] += [file for file in public_file]
        try:
            pythoncom.CoInitialize()
            engine = pyttsx3.init()
            engine.say('公共附件%s' % public_file)
            engine.runAndWait()
        except BaseException:
            pass


# 时间编号
def time_no():
    no = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
    return no


# 发件人获取(存在dict中)
def get_from(path):
    winsound.Beep(400, 600)  # 其中400表示声音大小，1000表示发生时长，1000为1秒
    input_name = input('请选择人员:')
    if not input_name:
        input_name = 'DEFAULT'
    config.read(path+'\\sender.ini', encoding='utf-8')
    sender_config = config[input_name]
    sender = []
    for i in ['name', 'address', 'password', 'smtp_server', 'smtp_port', 'sig_pic']:
        if i in sender_config:
            sender.append(sender_config[i])

    print('(*￣︶￣)欢迎%s,您的发件箱是:%s' % tuple(sender[:2]))
    try:
        pythoncom.CoInitialize()
        engine = pyttsx3.init()
        engine.say('hi%s' % sender[0])
        engine.runAndWait()
    except BaseException:
        pass
    return sender


# 收件人获取
def get_recipients(addr_dict, addr_in, cc_all):
    addr_city = {}
    for i in addr_in:
        if i not in [j[0] for j in addr_dict]:
            print('￣へ￣\033[1;31merror:收件人类型 %s 不存在' % i)
        elif i in [j[0] for j in addr_dict if j[1] == 'all']:
            if len([j for j in addr_dict if j[0] == i]) > 1:
                print('￣へ￣\033[1;31merror:收件人类型 %s 异常' % i)
            else:
                cc_all += addr_dict[i, 'all']+","
        else:
            for city in set(j[1] for j in addr_dict if j[0] == i):
                addr_city[city] = addr_city.get(city, '')+addr_dict.get((i, city), '')+","
    return addr_city, cc_all


# 登录邮箱
def login(address, password, smtp_server, smtp_port):
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    # server = smtplib.SMTP()
    # server.starttls()
    server.set_debuglevel(1)
    server.login(address, password)
    return server


# 添加附件
def add_attachment(file, i=0):
    with open(file, 'rb') as f:
        shotname, extension = os.path.splitext(file)
        mime = MIMEBase(shotname, extension, filename=file)
        mime.add_header('Content-Disposition', 'attachment', filename=Header(file, 'utf-8').encode())
        mime.add_header('Content-ID', '<%s>' % i)
        mime.add_header('X-Attachment-Id', 'i')
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        return mime


# 格式化
def format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, "utf-8").encode(), addr))


# 移动文件
def move(file, path):
    shotname, extension = os.path.splitext(file)
    shutil.move(file, path+"\\♠"+shotname+extension)


# 发送邮件
def send(sender, head, text, to, cc, file_list, path, server):
    receive = str(to+cc).split(',')
    msg = MIMEMultipart()
    if len(sender) == 6:
        msg.attach(add_attachment(sender[5], 99))
    list(map(lambda x: msg.attach(add_attachment(x)), file_list))
    msg.attach(MIMEText('%s' % text, 'html', 'utf-8'))  # plain text
    msg["from"] = format_addr("%s<%s>" % tuple(sender[:2]))
    msg["to"] = to
    msg["cc"] = cc
    msg["subject"] = Header("%s" % head, "utf-8").encode()
    server.sendmail(sender[1], receive, msg.as_string())
    for file in file_list:
        if '❤' in file or '未发送附件' in file:
            move(file, path)


# 压缩文件
def zip_file(path, path_done):
    os.chdir(path)
    file_list = os.listdir('.')
    file_city = [file for file in file_list if '❤' in file]
    if len(file_city) > 0:
        file_zip = '未发送附件%s.zip' % time_no()
        zip = zipfile.ZipFile(file_zip, 'w', zipfile.ZIP_DEFLATED)
        for file in file_list:
            zip.write(file)
            shutil.move(file, path_done+'\\'+file)
        zip.close()
        return file_zip

    if len(file_city) == 0:
        for file in file_list:
            move(file, path_done)


# 邮件内容输入
def message(mode='general'):
    cc_all, mimetext = '', ''
    ad = '<br /><br />由<a href="https://github.com/ifansay/tool">智文邮件助手</a>发送<br />ID:'
    if mode == 'general':
        header_in = input("请输入主题:\n")
        mimetext = input("请分段输入正文:\n")
        while mimetext != "":
            mimetext_in = input("请继续输入正文:\n")
            if not mimetext_in:
                break
            mimetext += "<br />"+mimetext_in
            mimetext = mimetext.replace('\n', '<br />')
        if mimetext == '':
            mimetext = '邮件自动推送.<br />'
        mimetext += ad+str(uuid.uuid1())

        to_in = input("请选择收件人类型(默认1):")
        if to_in == '':
            to_in = '1'

        cc_in = input("请选择抄送人类型:")
        while True:
            cc_all_in = input("请添加统一抄送人,多个用英文','隔开:")
            cc_all += cc_all_in+","
            if not cc_all_in:
                break

        to_add, cc_all = get_recipients(addr_dict, to_in, cc_all)
        cc_add, cc_all = get_recipients(addr_dict, cc_in, cc_all)
        confirm = input("输入'yes'确认发送:")
        return [header_in, mimetext], [to_add, cc_add, cc_all], confirm
    else:
        header_in = '未发送的附件,请处理!!!'+"◇"+time_no()
        mimetext = '附件发送失败,请处理.'+ad+str(uuid.uuid1())
        return [header_in, mimetext]


# 配置邮件内容
def struct_mail(header, mimetext, sender, to_add, cc_add, cc_all, city):
    head = city + "^_^" + header + "_M" + time_no()
    sig = ''
    if len(sender) == 6:
        sig = '<style type="text/css">p.margin {margin: -0.4cm 0cm -0.15cm 0cm}</style><p class="margin"><img src="cid:99" height="50"></p>'
    normal = mimetext+'<br />'+'--'*max(len(sender[1]),30)+sig+sender[0]+'<br /><a href="mailto:'+sender[1]+'">'+sender[1]
    html = etree.HTML(normal)
    text = etree.tostring(html).decode('utf-8')
    to_city, cc_city = to_add[city], cc_add.get(city, '') + cc_all
    return head, text, to_city, cc_city


# 发信主函数
def main(sender, info, recipients, path_done, path, file):
    try:
        server, time_start, unsent = login(*sender[1:5]), time.time(), set()
        for city in file:
            try:
                city_info = struct_mail(*info, sender, *recipients, city)
                send(sender, *city_info, file[city], path_done, server)
            except KeyError as e:
                print('KeyError', city, e)
                unsent.add(city)
            except FileNotFoundError as e:
                print('FileNotFoundError', e)

        zfile = zip_file(path, path_done)
        zip_info = message('fail')
        send(sender, *zip_info, sender[1], '', [zfile], path_done, server)
        time_end = time.time()
        print('用时 %.2f 秒' % (time_end-time_start))
        server.quit()
        if len(unsent) > 0:
            print('\033[1;31;47merror:%s发送失败\033[0m' % unsent)
        winsound.Beep(1000, 600)
    except TypeError as e:
        print('TypeError:', e)
    except NameError as e:
        print('NameError:', e)
    except smtplib.SMTPException as e:
        print('smtplib.SMTPException:', e)


"""
part3
"""

if __name__ == '__main__':
    pass

if not file_dict:
    print(';\033[1;35;43merror:空附件,不能发信\033[0m')
else:
    public(public_file)
    try:
        sender = get_from(path_project)
        info, recipients, confirm = message()
        if not recipients[0]:
            print('\n\033[1;35;43merror:无收件人,不能发信\033[0m')  # 如果只有抄送人,不发信
        elif confirm.lower() == "yes":
            main(sender, info, recipients, path_mailed, path_mail, file_dict)
    except KeyError as e:
        print('KeyError:', e)


winsound.Beep(1000, 300)  # 其中400表示声音大小，1000表示发生时长，1000为1秒
# input('\n\n请按任意键退出')
