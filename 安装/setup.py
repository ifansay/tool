#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @author: ifansay
# @email: ifansay.chn@qq.com

from oh_my_tuna import *
import sys,os
import platform
import configparser as cp

# 安装相应库
os.system('pip3 install xlrd')
os.system('pip3 install xlsxwriter')
os.system('pip3 install chardet')
os.system('pip3 install lxml')
# os.system('pip3 install lxml')
if platform.system()=='Windows':
    os.system('pip3 install winsound')
    os.system('pip3 install pyttsx3')
    os.system('pip3 install pythoncom')
    
# 创建相关文件夹
