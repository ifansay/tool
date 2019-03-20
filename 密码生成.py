# -*- coding: utf-8 -*-
"""
Created on Mon Mar 18 13:35:38 2019

@title: password generation

@author: 饭未眠

@email: ifansay.chn@qq.com
"""
import string
import random
import datetime
import pathlib


def passwordProduct(long):
    src = string.ascii_letters + string.digits + string.punctuation
    # 此处亦可以用正则表达式校验
    password = random.sample(src, sorted([5, long-3, 32])[1])  # 随机取n位,排序取中值
    password.extend(random.sample(string.digits, 1))  # 让密码中一定包含数字
    password.extend(random.sample(string.ascii_lowercase, 1))  # 让密码中一定包含小写字母
    password.extend(random.sample(string.ascii_uppercase, 1))  # 让密码中一定包含大写字母
    for i in range(10000):
        random.shuffle(password)  # 打乱列表顺序
    str_password = ''.join(password)  # 将列表转化为字符串
    return str_password


def longCheck(long):
    if long < 8:
        print('密码最短8位')
    elif long > 35:
        print('密码最长35位')


def passwordSave(tags, account, password, date):
    file = '%s密码.txt' % tags
    fileCheck = pathlib.Path(file).is_file()
    f = open(file, 'a+', encoding='utf-8')
    if not fileCheck:  # 确认文件是否存在的判断规则
        f.writelines('App&Web'.ljust(15, ' '))
        f.writelines('AccountNo'.ljust(35, ' '))
        f.writelines('Passwords'.ljust(40, ' '))
        f.writelines('CreatTime\n')
    f.writelines(tags.ljust(15, ' '))
    f.writelines(account.ljust(35, ' '))
    f.writelines(password.ljust(40, ' '))
    f.writelines(date)
    f.writelines('\n')
    f.close()


def infoInput():
    tags = input('请输入账号归属(最长10位,英文):')[:10]
    account = input('请输入账号(最长30位,英文):')[:30]
    passlong = int(input('请输入密码长度:'))
    return tags, account, passlong


def main():
    try:
        tags, account, long = infoInput()
        longCheck(long)
        password = passwordProduct(long)
        passwordSave(tags, account, password, datetime.datetime.today().strftime('%Y/%m/%d %H:%M:%S'))
        print(tags, "的密码为:", password, '\n密码已存入文件')
    except:
        print('输入有误')


main()
