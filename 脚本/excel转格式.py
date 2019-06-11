# -*- coding: utf-8 -*-
"""
Created on Wed Sep  5 16:58:55 2018

@title: excel conversion

@author: ifansay

@email: ifansay.chn@qq.com
"""

import os
import os.path
import datetime
import time
import shutil
import xlrd
import xlsxwriter
import re
import chardet
import winsound
import csv
import configparser as cp


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


def time_no():
    no = datetime.datetime.today().strftime('%y%m%d')
    return no


def csvRead(file):
    data = []
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
        data.append(text)
    f.close()
    return data


# read_excel:读取excel
def excelRead(file):
    shotname, extension = os.path.splitext(file)
    book = xlrd.open_workbook(file)
    data = {}
    for sheet in book.sheets():
        for i in range(sheet.nrows):
            j = 0
            text = sheet.row_values(i)
            # type: 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
            text_type = sheet.row_types(i)
            for cell in text:
                ctype = text_type[j]
                if ctype == 2 and cell % 1 == 0:
                    text[j] = int(cell)
                elif ctype == 3 and cell % 1 == 0:
                    text[j] = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y-%m-%d')
                elif ctype == 3:
                    text[j] = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y-%m-%d %H:%M:%S')
                elif ctype == 4:
                    text[j] = True if cell == 1 else False
                j += 1
            data[sheet.name] = data.get(sheet.name, [])+[text]
    return data


def excelWrite(name, data, path):
    oth = {"font_name": u'微软雅黑',
           'border': 4,
           'border_color': '50616d',
           'font_size': 10
           }
    workbook = xlsxwriter.Workbook(str(path+'\\'+name+'.xlsx'))
    worksheet = workbook.add_worksheet(name=name[:30])
    for row in range(len(data)):
        for col in range(len(data[row])):
            worksheet.write(row, col, data[row][col], workbook.add_format(oth))
    workbook.close()


def csvWrite(name, data, path):
    with open(str(path+'\\'+name+'.csv'), "w", newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)


def shut(file, new_file, time_start):
    time_end = time.time()
    shutil.move(file, new_file)
    print('用时 %.2f 秒' % (time_end-time_start))


def main(file, path, path_filed, mode):
    shotname, extension = os.path.splitext(file)
    print(file)
    time_start = time.time()
    new_file = path_filed+'\\'+shotname+"_TO"+time_no()+extension
    if extension == '.csv' and mode == '1':
        data = csvRead(file)
        excelWrite(shotname, data, path)
        shut(file, new_file, time_start)
    elif extension in ['.xls', '.xlsx'] and mode == '0':
        data = excelRead(file)
        for sheet, data_w in data.items():
            name = str(shotname+"-"+sheet)
            if len(data) == 1:
                name = str(shotname)
            csvWrite(name, data_w, path)
        shut(file, new_file, time_start)


os.chdir(path_file)
file_list = os.listdir('.')

if file_list:
    print('^_^欢迎使用智文文件助手^_^\n')
    mode = input('请输入转化类型(1.csv转excel;0.excel转csv):')
    winsound.Beep(220, 800)
    for file in file_list:
        main(file, path_file, path_filed, mode)
    winsound.Beep(1000, 300)

# input('\npress any key to exit')
