#!/usr/bin/env python3﻿
# -*- coding: utf-8 -*-
"""
Created on Thu Mar  8 11:17:20 2018

@title: excel merge

@author: ifansay

@email: ifansay.chn@qq.com
"""

import os
import os.path
import time
import shutil
import xlrd
import xlsxwriter
import re
import chardet
#import winsound
import configparser as cp


config = cp.RawConfigParser()
file_path, code_file = os.path.split(os.path.realpath(__file__))

config.read(file_path+'//config.ini', encoding='utf-8')
sf_config = config['default']
path_project = sf_config['project']
path_file = sf_config['file']
path_filed = sf_config['filed']
path_mail = sf_config['mail']
path_mailed = sf_config['mailed']
separator = sf_config['separator']


# 定义时间编号
def time_no():
    return time.strftime("%Y%m%d", time.localtime())


# 读取csv
def csvRead(file, data):
    shotname, extension = os.path.splitext(file)
    f = open(file, 'rb')
    for line in f.readlines():
        try:
            encod_char = chardet.detect(line)
            encoding = encod_char['encoding']
            if encod_char['language'] == 'Chinese':
                t = line.decode(encoding)
            else:
                try:
                    t = line.decode('gb18030')
                except UnicodeDecodeError:
                    t = line.decode('utf-8')
        except TypeError:
            t = line.decode('gb18030')
        except UnicodeDecodeError:
            try:
                t = line.decode('gb18030')
            except UnicodeDecodeError:
                t = line.decode('utf-8')
        except UnicodeTranslateError:
            t = line.decode('gb18030')
        text = re.sub('[\r\n\t"]', '', t).split(',')
        if data:
            data['sheet'] += [text+[shotname]]
        else:
            data['sheet'] = [text+['source']]
    f.close()
    return data


# 格式化行文本
def textFormat(ctype, cell):
    if ctype == 2 and cell % 1 == 0:
        cell = int(cell)
    elif ctype == 3 and cell % 1 == 0:
        cell = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y/%m/%d')
    elif ctype == 3:
        cell = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y/%m/%d %H:%M:%S')
    elif ctype == 4:
        cell = True if cell == 1 else False
    return cell


def excelRead(file, data, mode):
    shotname, extension = os.path.splitext(file)
    book = xlrd.open_workbook(file)
    for sheet in book.sheets():
        if mode == '1':
            sheet_name = sheet.name
            source = shotname
        else:
            sheet_name = 'sheet'
            source = shotname+'/'+sheet.name

        for row in range(0, sheet.nrows):
            text = sheet.row_values(row)
            text_type = sheet.row_types(row)
            text = list(map(textFormat, text_type, text))
            if data:
                text.append(source)
            else:
                text.append('source')
            data[sheet_name] = data.get(sheet_name, []) + [text]
    return data


def excelWrite(data, file):
    #winsound.Beep(400, 600)
    print('正在写入',file)
    workbook = xlsxwriter.Workbook(str(file+"_C"+time_no()+'.xlsx'))
    top = {"font_name": u'微软雅黑',
           'border': 4,
           'border_color': '50616d',
           'font_size': 11,
           'bold': True,
           'bg_color': 'fffbf0'}
    oth = {"font_name": u'微软雅黑',
           'border': 4,
           'border_color': '50616d',
           'font_size': 10}
    for sheet, data_w in data.items():
        worksheet = workbook.add_worksheet(name=sheet)
        for row in range(len(data_w)):
            for col in range(len(data_w[row])):
                if row == 0:
                    worksheet.write(row, col, data_w[row][col], workbook.add_format(top))
                else:
                    worksheet.write(row, col, data_w[row][col], workbook.add_format(oth))
    workbook.close()


os.chdir(path_file)
file_list = os.listdir('.')

if file_list:
    data = {}
    print('^_^欢迎使用智文文件助手^_^\n')
    mode = input('请输入合并模式(1.同名sheet合并0.全部sheet合并,默认0模式):')
    nf = input('请输入合并后文件名称:\n')
    time_start = time.time()
    for of in file_list:
        shotname, extension = os.path.splitext(of)
        if extension in ['.xlsx', '.xls', '.csv']:
            print('正在处理', of)
            try:
                data = excelRead(of, data, mode)
            except BaseException:  # xlsxwriter.XLRDError
                data = csvRead(of, data)
            shutil.move(of, path_filed+'//'+shotname+"_CO"+time_no()+extension)
    if data:
        excelWrite(data, path_mail+'//'+nf)
        time_end = time.time()
        print('用时 %.2f 秒' % (time_end - time_start))
        # winsound.Beep(1000, 300)

# input('\npress any key exit')

