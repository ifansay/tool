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
# 如果是含","的文件,是否拆分到多个组织呢

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
import string
import random
import configparser as cp

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email.utils import parseaddr, formataddr
from email import encoders
from lxml import etree


version = '3.1.1'
update_date = '2019/06/05'

config = cp.RawConfigParser()
file_path, code_file = os.path.split(os.path.realpath(__file__))

config.read(file_path+'\\sf_config.ini', encoding='utf-8')
sf_config = config['DEFAULT']
path_project = sf_config['project']
path_file = sf_config['file']
path_filed = sf_config['filed']
path_mail = sf_config['mail']
path_mailed = sf_config['mailed']
separator = sf_config['separator']

# 读取收件人配置
f = open(path_project+'\\recipients.txt', 'r', encoding='utf-8')
addr_dict, j, error_recipients, error_email = {}, 1, [], []
for i in f:
    if ';' not in i and i != '\n':
        try:
            i = i.strip('\n').split('/')
            addr_dict[i[0], i[1]] = i[2:]
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


ad = '由<a href="mailto:ifansay.chn@qq.com">ifansay</a>开发的<a href="https://github.com/ifansay/tool">智文邮件助手</a>发邮件</p>'


# 文件整理
def filePack(file_list, separator):
    file_dict, public_file, file_set, file_city = {}, set(), set(), {}
    # 按照组织打包的文件/公共文件/无组织文件/文件对应的组织
    for file in file_list:
        if file not in ('desktop.ini', '.DS_Store',):
            try:
                file_set.add(file[file.index(separator)+1:])
                city = file[0:file.index(separator)].split(',')
                # 有分隔符号的文件进入组织附件,否则算公共附件
                file_city[file] = file_city.get(file, set())
                # file_dict[city] = file_dict.get(city, [])+[file]
                # city_list = city
                for i in range(len(city)):
                    """
                    city的每个元素建立key,如[a, b, c]建立[a, b, c, 'a,b', 'a,b,c'] 5个key,
                    对应同一个file,如果后续某key增加新file,其他key不变
                    """
                    file_dict[city[i]] = file_dict.get(city[i], set())
                    file_dict[city[i]].add(file)
                    file_city[file].add(city[i])

                    com_city = ','.join(city[:i+1])
                    file_dict[com_city] = file_dict.get(com_city, set())
                    file_dict[com_city].add(file)
                    file_city[file].add(com_city)

            except BaseException:
                public_file.add(file)
                file_set.add(file)
    return file_dict, file_set, public_file, file_city
    # 组织对应附件/附件通名/公共附件/附件对应组织


def speak(characters):
    pythoncom.CoInitialize()
    engine = pyttsx3.init()
    engine.say(characters)
    engine.runAndWait()


# 公共附件
def public(file_dict, public_file):
    print(';其中\033[1;35;47m%s\033[0m作为公共附件发送' % public_file)
    for city in file_dict:
        file_dict[city] = file_dict[city] | public_file
    try:
        speak('公共附件%s' % public_file)
    except BaseException:
        pass


# 时间编号
def time_no():
    no = datetime.datetime.today().strftime('%y%m%d')
    return no


# 发件人获取(存在dict中)
def sender(path):
    winsound.Beep(400, 600)  # 其中400表示声音大小，1000表示发生时长，1000为1秒
    input_name = input('请选择人员:')
    if not input_name:
        input_name = 'DEFAULT'
    configs = cp.RawConfigParser()
    configs.read(path+'\\sender.ini', encoding='utf-8')
    sender_config = configs[input_name]
    sender, sig_extra = [], []
    index = ['name', 'address', 'password', 'smtp_server', 'smtp_port', 'sig_pic']
    for i in index:
        if i in sender_config:
            sender.append(sender_config[i])
    for j in sender_config:
        if j not in index:
            sig_extra.append(sender_config[j].replace('\n', '<br />'))
    print('(*￣︶￣)欢迎%s,您的发件箱是:%s' % tuple(sender[:2]))
    print('正在登录邮箱,loading...')
    try:
        speak('hi%s' % sender[0])
    except BaseException:
        pass
    return sender, sig_extra


# 收件人获取
def recipientsGet(addr_dict, addr_in, cc_all):
    addr_city = {}
    for i in addr_in:
        if i not in [j[0] for j in addr_dict]:
            print('￣へ￣\033[1;31merror:收件人类型 %s 不存在' % i)
        elif i in [j[0] for j in addr_dict if j[1] == 'all']:
            if len([j for j in addr_dict if j[0] == i]) > 1:
                print('￣へ￣\033[1;31merror:收件人类型 %s 异常' % i)
            else:
                cc_all += addr_dict[i, 'all'][0]+","
        else:
            for city in set(j[1] for j in addr_dict if j[0] == i):
                try:
                    a = addr_dict[(i, city)]
                except KeyError:
                    a = ['','']
                if city in addr_city:
                    addr_city[city][0] += ','+a[0]
                    addr_city[city][1] += ','+a[1]
                else:
                    addr_city[city] = a
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
def addrFormat(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, "utf-8").encode(), addr))


# 移动文件
def move(file, path):
    src = string.ascii_letters+string.digits
    randstr = ''.join(random.sample(src, 6))
    shotname, extension = os.path.splitext(file)
    shutil.move(file, path+'\\♣'+shotname+'_'+randstr+extension)


# 发送邮件
def send(sender, head, text, to, cc, file, path, server, file_city={}, to_set=set(), sent=set(), city=''):
    # 发件人/标题/正文/收件人/抄送/附件list/移动后附件路径/服务器/文件对应组织/收件人组织/已发送组织
    # 有逗号的附件,要保证使用此附件的组织都发送完成才能移动附件,即文件对应组织和收件人组织的交集是已发送组织子集
    receive = str(to+cc).split(',')
    msg = MIMEMultipart()
    if len(sender) >= 6:
        msg.attach(add_attachment(sender[5], 'pic99'))
    for f in file:
        i = 1
        msg.attach(add_attachment(f, 'f'+str(i)))
        i += 1
    # list(map(lambda x: msg.attach(add_attachment(x)), file))
    msg.attach(MIMEText(text, 'html', 'utf-8'))  # plain:text
    msg["from"] = addrFormat("%s<%s>" % tuple(sender[:2]))
    msg["to"] = to
    msg["cc"] = cc
    if '^_^_M' in head:
        head = list(file)[0]
    msg["subject"] = Header("%s" % head, "utf-8").encode()
    server.sendmail(sender[1], receive, msg.as_string())
    sent.add(city)
    for f in file:
        if 'UNSENT' in f or to_set & file_city[f] <= sent:
            # 单一附件应发送组织是附件对应组织(file_city[file])与收件人组织(to_add的key)的交集,如果此组织是已发送组织子集,则移动
            move(f, path)


# 压缩文件
def fileZip(path, path_done, separator):
    os.chdir(path)
    file_list = os.listdir('.')
    file_city = [file for file in file_list if separator in file]
    if len(file_city) > 0:
        file_zip = '!UNSENT!%s.zip' % time_no()
        zip = zipfile.ZipFile(file_zip, 'w', zipfile.ZIP_DEFLATED)
        for file in file_list:
            print(file)
            zip.write(file)
            move(file, path_done)  # 压缩后移动原文件
        zip.close()
        return file_zip, file_list

    if len(file_city) == 0:
        for file in file_list:  # 公共附件
            move(file, path_done)


def failinfo(file):
    if file:
        print('Friendly Tipsfollowing file:')
        list(map(lambda x: print(x), file))
        print('\033[1;31;47merror:not be sent\033[0m')


# 邮件内容输入
def infoGet(path='', addr_dict=None, mode='general'):
    cc_all, mimetext = '', ''
    if mode == 'general':
        header_in = input("请输入主题:\n")
        if path+'\\mailtext.html':
            fm = open(path_project+'\\mailtext.html', 'r', encoding='utf-8-sig')
            mimetext = ''.join(fm.readlines())
            print('检测到正文文件,正文为("%待替换姓名%"会替换为收件人姓名):')
            print(mimetext)
            fm.close()
        else:
            mimetext = input("请分段输入正文:\n")
            while mimetext != "":
                mimetext_in = input("请继续输入正文:\n")
                if not mimetext_in:
                    break
                mimetext += "<br />"+mimetext_in
            if mimetext == '':
                mimetext = '<p>自动发送.<br />'
        mimetext = mimetext.replace('\n', '<br />')
        mimetext = "<p>"+mimetext+"<br /><br />ID:"+str(uuid.uuid1())+"<br />"

        to_in = input("请选择收件人类型(默认1):")
        if to_in == '':
            to_in = '1'

        cc_in = input("请选择抄送人类型:")
        while True:
            cc_all_in = input("请添加统一抄送人,多个用英文','隔开:")
            cc_all += cc_all_in+","
            if not cc_all_in:
                break

        to_add, cc_all = recipientsGet(addr_dict, to_in, cc_all)
        cc_add, cc_all = recipientsGet(addr_dict, cc_in, cc_all)
        confirm = input("输入'yes'确认发送:")
        return [header_in, mimetext], [to_add, cc_add, cc_all], confirm
    else:
        header_in = '失败附件!'+time_no()
        mimetext = '附件发送失败,请处理.'+ad
        return [header_in, mimetext]


# 邮件内容配置
def mailStruct(header, mimetext, sender, to_add, cc_add, cc_all, city, sig_extra):
    head = city + '^_^' + header + '_M' + time_no()
    s, ex = '', ''
    if len(sender) == 6:
        s = '<img src="cid:pic99" height="50"><br />'
    if sig_extra:
        ex = '<br />'.join(sig_extra)+"<br /><br />"
    n = "<b>"+sender[0]+"</b><br />"
    m = '<i><a href="mailto:'+sender[1]+'">'+sender[1]+"</a></i><br />"
    selfinfo = "--"*max(len(sender[1]), 30)+"<br />"+s+n+m+ex
    mimetext = mimetext.replace('%待替换文本%', to_add[city][1])
    normal = mimetext+selfinfo+ad
    html = etree.HTML(normal)
    text = etree.tostring(html).decode('utf-8')
    to_city, cc_city = to_add[city][0], cc_add.get(city, [''])[0] + cc_all
    return head, text, to_city, cc_city


# 发信主函数
def main(sender, info, recipients, path_done, path, file, separator, file_city, server):
    try:
        time_start, sent = time.time(), set()
        to_set = recipients[0].keys()
        # 服务/起始时间/失败组织/成功组织
        for city in file:
            try:
                city_info = mailStruct(*info, sender, *recipients, city, sig_extra)
                send(sender, *city_info, file[city], path_done, server, file_city, to_set, sent, city)
            except KeyError as e:
                print('KeyError', city, e)
            except FileNotFoundError as e:
                print('FileNotFoundError', e)

        zfile, fail_file = fileZip(path, path_done, separator)
        zip_info = infoGet(mode='fail')
        send(sender[:-1], *zip_info, sender[1], '', [zfile], path_done, server)
        time_end = time.time()
        server.quit()
        print('用时 %.2f 秒' % (time_end-time_start))
        failinfo(fail_file)
        winsound.Beep(1000, 600)
    except TypeError as e:
        print('TypeError:', e)
    except NameError as e:
        print('NameError:', e)


os.chdir(path_mail)
file_list = os.listdir('.')

if __name__ == '__main__':
    pass

file_dict, file_set, public_file, file_city = filePack(file_list, separator)

if file_dict:
    print('^_^欢迎使用智文邮件助手^_^\n')
    print("收件人为:\033[1;34;47m%s\033[0m" % set(file_dict), ",附件:\033[1;34;47m%s\033[0m" % file_set, end='')
    if public_file:
        public(file_dict, public_file)
    try:
        sender, sig_extra = sender(path_project)
        server = login(*sender[1:5])
        info, recipients, confirm = infoGet(path_project, addr_dict)
        if not recipients[0]:
            print('\n\033[1;35;43merror:无收件人,不能发信\033[0m')  # 如果只有抄送人,不发信
        elif confirm.lower() == "yes":
            main(sender, info, recipients, path_mailed, path_mail, file_dict, separator, file_city, server)
    except KeyError as e:
        print('KeyError:', e)
    except smtplib.SMTPException as e:
        print('smtplib.SMTPException:', e)

    winsound.Beep(1000, 300)  # 其中400表示声音大小，1000表示发生时长，1000为1秒

# input('\npress any key to exit')
