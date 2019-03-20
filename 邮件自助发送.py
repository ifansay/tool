# -*- coding: utf-8 -*-
"""
Created on Mon Oct 16 22:09:20 2017

@title:mail self—send

@author: ifansay

@email: ifansay.chn@qq.com
"""

# 主要步骤
"partA-step1如果待拆分文件夹存在文件,拆分,保存到发件文件夹"
"part1读取配置"
"part2自定义函数"
"part3输入参数"
"part4调用函数,发送邮件"
"""定义三类的附件:1.私有附件.2.公共附件.3.汇总附件."""

# 下步计划
"""
1.GUI
2.超链接
2.图片附件
3.签名
4.异常抛出
5.class改写
"""

# 更新记录
"""版本号v1.32,更新日期2018-04-03
1.增加对汇总附件的支持"""

"""版本号v1.31,更新日期2018-03-23
1.增加无收件人报错功能
2.对分段正文的支持"""

"""版本号v1.3,更新日期2018-03-20
1.修改附加添加函数
2.修改发送邮件函数
3.增加公共附件发送功能
4.增加未发送附件返回发件人功能"""

"""版本号v1.22,更新日期2018-01-16
1.修复抄送人的bug"""

"""版本号v1.21,更新日期2018-01-15
1.读取csv文件有问题，暂时删除此功能"""

"""版本号v1.2,更新日期2018-01-12
1.文件拆分/发送后,清空文件到其他文件夹(未发送不清空)
2.允许各类型收件人设置不同的收件人集合
3.合并邮件标题和正文的编码逻辑"""

"""版本号v1.11,更新日期2018-01-11
1.可以设置多个统一收件人"""

"""版本号v1.1,更新日期2018-01-10
1.增加表格拆分功能
2.优化文件路径设置"""

"""版本号v1.0,更新日期2018-01-06
1.根据文本文件读取配置文件"""

"""版本号v0.2,更新日期2018-01-04
1.优化bug:抄送人可用"""

"""版本号v0.1,更新日期2017-12-20
1.创建文件,可用"""

# \033[显示方式;字体色;背景色m......[\033[0m]
# error:red background
# waring:yellow background
# normal:write background


import os,os.path
import uuid
import datetime,time
import shutil
import smtplib
import configparser as cp
import winsound
import pyttsx3
import zipfile

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email.utils import parseaddr,formataddr
from email import encoders


"""
part1:固定配置读取
公共附件何时移动
"""

# 配置文件路径
path = "/mailsend/"  # 全路径
path_original = path + "original_file/"  # 待拆分文件路径
path_original_done = path + 'original_file_done/'  # 已拆分文件保存路径
path_send = path + "send_file/"  # 待发送文件路径
path_send_done = path + 'send_file_done/'  # 已发送文件保存路径

# 读取收件人配置
f = open(path+'recipients.txt', 'r', encoding='utf-8')
addr_dict, j = {}, 1
for i in f:
    try:
        i = i.strip('\n').split('/')  # 将字符串格式化为list
        addr_dict[i[0], i[1]] = i[2]  # list to dict
    except:
        print('❤第%d行%s错误' % (j, i))
    j += 1
f.close()

cate = [i[0] for i in addr_dict if i[1] == 'all']
for i in cate:
    check = [g[1] for g in addr_dict if g[0] == i]
    if len(check) > 1:
        print('\033[1;33;41m错误:收件人类型%s错误\033[0m' % i)
        pass  # 如果一个类型为all,但是有其他组织,报错

# 读取附件列表
os.chdir(path_send)
file_list = os.listdir('.')
file_dict = {}; public_file = []; file_set = set()
for file in file_list:
    """按照一定规则切分文件名,否则作为公共附件添加"""
    if file not in ('desktop.ini', '.DS_Store'):
        try:
            city = file[0:file.index("❤")]  # 会截取第一个 '❤' 的字段
            file_set.add(file[file.index("❤")+1:])  # 截取"❤"
            file_dict[city] = file_dict.get(city, [])+[(path_send + file, file)]  # 含路径和文件名、文件组织等参数
        except:
            public_file.append(file)
            file_set.add(file)

print("您即将发送的组织为:\033[1;34;47m%s\033[0m" % set(file_dict), "附件:\033[1;34;47m%s\033[0m" % file_set, end="")

"处理公共附件"
if len(public_file) > 0:
    print(';其中\033[1;35;47m%s\033[0m作为公共附件发送' % public_file)
    engine = pyttsx3.init()
    engine.say('公共附件%s' % public_file)  # 公共附件播报
    engine.runAndWait()
    for city in file_dict:
        file_dict[city] += [(path_send + file, file) for file in public_file]


"""
part2:自定义函数
"""


# 定义时间编号
def time_no():
    no = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
    return no


# 发件人配置(存在dict中)
def get_from():
    winsound.Beep(400, 600)  # 其中400表示声音大小，1000表示发生时长，1000为1秒
    input_name = input('请选择人员:\n')
    config = cp.ConfigParser()
    config.read(path+'config.ini', encoding='utf-8')
    index = {'address', 'name', 'password', 'smtp_server', 'smtp_port'}
    from_dict = {}
    for i in index:
        try:
            from_dict[i] = config['DEFAULT']['%s' % str(input_name+i)]
        except:
            if i == 'name':
                print('\033[1;31m警告:输入人员不存在,使用默认邮箱\033[0m')
            from_dict[i] = config['DEFAULT']['%s' % i]
    try:
        engine = pyttsx3.init()
        engine.say('hi%s' % from_dict['name'])
        print('\n==(*￣︶￣)欢迎\033[1;34;47m%s\033[0m,您的发件箱是:\033[1;34;47m%s\033[0m;' % (from_dict['name'], from_dict['address']))
        engine.runAndWait()
    except:
        pass
    return from_dict


# 收件人设置
def get_recipients(addr_dict, addr_in, cc_addr_all):
    addr_dict_city = {}
    for i in addr_in:
        if i not in [j[0] for j in addr_dict.keys()]:
            print('￣へ￣\033[1;31m警告:收件人类型 %s 不存在,不予发送\033[0m￣へ￣' % i)
        if i in [j[0] for j in addr_dict.keys() if j[1] == 'all']:  # 如果是统一抄送人类型
            cc_addr_all += addr_dict[i, 'all']+","  # 把收件人添加到统一抄送人
        for city in set(j[1] for j in addr_dict.keys() if j[1] != 'all' and j[0] == i):
            # 校验all在此类型是否唯一
            # if len(set(j[1] for j in addr_dict.keys() if j[0] == i))>0:pass #
            addr_dict_city[city] = addr_dict_city.get(city, '')+addr_dict.get((i, city), '')+","
    return addr_dict_city, cc_addr_all


# 邮件内容配置
def input_message():
    cc_all = ''  # 初始化统一抄送人
    ad = "\n\n\nby Mail Self-Sending System.\nno:"  # 广告信息
    header_in = input("请输入主题:\n")

    mimetext = input("请分段输入正文:\n")
    while mimetext != "":
        mimetext_in = input("请分段输入正文:\n")
        mimetext += "\n"+mimetext_in
        if not mimetext_in:
            break
    if mimetext == '':
        mimetext = '❤此邮件自动推送,无需回复❤\n'

    to_in = input("请选择收件人(默认类型1):\n")  # 输入收件人选项
    if to_in == '':
        to_in = '1'

    cc_in = input("请选择抄送人:\n")
    while True:
        cc_all_in = input("请添加统一抄送人,多个用英文','隔开:\n")
        cc_all += cc_all_in+","
        if not cc_all_in:
            break

    # 配置收件人
    to_add, cc_all = get_recipients(addr_dict, to_in, cc_all)  # all-public receiver
    cc_add, cc_all = get_recipients(addr_dict, cc_in, cc_all)

    confirm_value = input("发送邮件请输入'yes'确认\n")

    return ad, header_in, mimetext, to_add, cc_add, cc_all, confirm_value


# 压缩文件
def zip_file(path_send, path_done):
    os.chdir(path_send)
    file_list = os.listdir('.')
    file_city = [file for file in file_list if '❤' in file]  # 有未发送的组织附件,则打包;分隔符❤
    if len(file_city) > 0:
        file_zip = '未发送附件%s.zip' % time_no()
        zip = zipfile.ZipFile(file_zip, 'w', zipfile.ZIP_DEFLATED)
        for file in file_list:
            zip.write(file)
            shutil.move(file, path_done+file)  # 将未发送附件打包并移动,保持原名
        zip.close()
        return file_zip  # 有未发送的组织附件,则返回打包

    if len(file_city) == 0:
        for file in file_list:  # 移动公共附件,文件重命名
            shotname, extension = os.path.splitext(file)  # 获取文件拓展名
            shutil.move(file, path_send_done+"♠"+shotname+extension)


# 添加附件
def add_attachment(path, file):
    with open(path, 'rb') as f:
        shotname, extension = os.path.splitext(file)  # 获取文件拓展名
        # 设置附件的MIME和文件名:
        mime = MIMEBase(shotname, extension, filename=file)
        # 加上必要的头信息:
        mime.add_header('Content-Disposition', 'attachment', filename=Header(file, 'utf-8').encode())
        mime.add_header('Content-ID', '<0>')
        mime.add_header('X-Attachment-Id', '0')
        # 把附件的内容读进来:
        mime.set_payload(f.read())
        # 用Base64编码:
        encoders.encode_base64(mime)
        # 添加到MIMEMultipart:
        return mime


# 格式化
def format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, "utf-8").encode(), addr))


# 登录邮箱
def login(smtp_server, smtp_port, from_add, password):
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)  # 默认端口25,阿里云邮SSL加密465
    # server = smtplib.SMTP()
    server.set_debuglevel(1)
    # server.connect(smtp_server, '25')
    # server.helo()
    # server.ehlo()
    # server.starttls()
    server.login(from_add, password)  # 登录服务器
    return server


# 发送邮件
def mail_send(from_add, name, to_add, file_list, header, mimetext, path_done, cc_add=""):
    # 发件人/收件人/抄送人/附件/收件人组织/邮件标题/邮件正文/广告//统一抄送
    receive = str(to_add+cc_add).split(',')  # 转化为list
    msg = MIMEMultipart()
    for file in file_list:
        msg.attach(add_attachment(*file))  # 这是什么动作
    msg.attach(MIMEText('%s' % mimetext, 'plain', 'utf-8'))
    msg["from"] = format_addr("%s<%s>" % (name, from_add))
    msg["to"] = to_add
    msg["cc"] = cc_add
    msg["subject"] = Header("%s" % header, "utf-8").encode()
    server.sendmail(from_add, receive, msg.as_string())
    for file in file_list:
        if '❤' in file[1] or '未发送附件' in file[0]:  # public文件不移动
            shotname, extension = os.path.splitext(file[1])  # 获取文件拓展名
            shutil.move(file[1], path_done+"♠"+shotname+extension)
    '''移动附件到已发送文件夹并重命名,公共附件不移动'''


# 定义发送回执
def send_message(unsent_city, unavailable_city):
    if len(unsent_city) > 0 and len(unavailable_city) > 0:
        print('\033[1;31;47m警告:%s无法发送;%s发送失败\033[0m' % (unavailable_city, unsent_city))
    elif len(unsent_city) > 0:
        print('\033[1;31;47m警告:%s发送失败\033[0m' % unsent_city)
    elif len(unavailable_city) > 0:
        print('\033[1;31;47m警告:%s无法发送\033[0m' % unavailable_city)


"""
part4:发送邮件
"""
# 如果收件空或附件空,停止发件;否则读取输入信息;如果确认是yes,登录MUA调用函数发送邮件.
if len(file_dict) == 0:
    print('\n\033[1;35;43m警告:空附件,不能发信\033[0m')
else:
    from_dict = get_from()
    ad, header_in, mimetext, to_add, cc_add, cc_all, confirm_value = input_message()
    if len(to_add) == 0:  # 必须有收件人,只有抄送人不行
        print('\n\033[1;35;43m警告:无收件人,不能发信\033[0m')
    elif confirm_value.lower() == "yes":
        unsent_city, unavailable_city, send_city = set(), set(), set()  # 记录异常数据,已发数据
        time_start = time.time()  # 初始时间
        server = login(from_dict['smtp_server'], from_dict['smtp_port'], from_dict['address'], from_dict['password'])  # 登录邮箱

        for city in file_dict:
            if city in to_add.keys():  # 如果在收件人列表中
                "如果不填写邮件标题，默认为附件的名称"
                header = city + "^_^" + header_in + "_M" + time_no()
                if header_in == '':
                    header = file_dict[city][0][-1]
                mimetext_all = mimetext+ad+str(uuid.uuid1())+"\n"+"--"*len(from_dict['address'])+"\n"+from_dict['name']+"\n"+from_dict['address']
                to_city,cc_city = to_add[city], cc_add.get(city, '') + cc_all
                try:
                    mail_send(from_dict['address'], from_dict['name'], to_city, file_dict[city], header,
                              mimetext_all, path_send_done, cc_city)  # 发送
                except:
                    unsent_city.add(city)
            else:
                print('\033[1;31;47m警告:%s不在收件人列表\033[0m' % city)
                unavailable_city.add(city)

        file_zip = zip_file(path_send, path_send_done)  # 压缩未发送附件并移动原文件
        fail_header = '未发送的附件,请处理!!!' + "◇" + time_no()
        fail_text = '❤%s您好,\n本邮件附件未能成功发送,请及时处理.%s %s' % (from_dict['name'], ad, uuid.uuid1())
        try:
            mail_send(from_dict['address'], from_dict['name'], from_dict['address'], [(path_send+file_zip, file_zip)], fail_header, fail_text, path_send_done)
        except:
            pass
        server.quit()  # 登出
        time_end = time.time()
        winsound.Beep(1000, 600)  # 其中400表示声音大小，1000表示发生时长，1000为1秒
        print('用时 %.2f 秒' % (time_end-time_start))
        send_message(unsent_city, unavailable_city)  # 打印失败组织
    elif confirm_value.lower() != "yes":
        print('\033[1;35;47m您取消了邮件发送\033[0m')
    else:
        print('\033[1;35;47m其他异常不能发送\033[0m')

"""
part5:退出
"""
winsound.Beep(1000, 300)  # 其中400表示声音大小，1000表示发生时长，1000为1秒
# input('\n\n请按任意键退出')
