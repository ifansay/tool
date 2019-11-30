# -*- coding: utf-8 -*-
"""
Created on Wed Feb 13 11:08:28 2019

@author: 饭未眠

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
import string
import random
import winsound
import configparser as cp

# 如果excel内容报错,拆分会陷入死循环 #
# comment
"""版本号v2.7,更新日期2018-12-13
1.增加对单一文件的多次拆分支持"""

"""版本号v2.6.1,更新日期2018-09-09
1.增加对数字、布尔值作为拆分列的支持"""

"""版本号v2.6,更新日期2018-08-25
1.增加对多列拆分的需求
2.增加对整数,日期,布尔值的支持
3.修改拆分逻辑,提升效率"""


config = cp.RawConfigParser()
file_path, code_file = os.path.split(os.path.realpath(__file__))

# 读取文件路径
config.read(file_path+'//config.ini', encoding='utf-8')
sf_config = config['default']
path_project = sf_config['project']
path_file = sf_config['file']
path_filed = sf_config['filed']
path_mail = sf_config['mail']
path_mailed = sf_config['mailed']
separator = sf_config['separator']


def time_no():
    return time.strftime("%Y%m%d", time.localtime())


def move(file, path):
    src = string.ascii_letters+string.digits
    randstr = ''.join(random.sample(src, 6))
    shotname, extension = os.path.splitext(file)
    shutil.move(file, path+"//"+shotname+'_SO'+'_'+randstr+extension)


def textSplit(text, split, data, title, sheet='sheet'):
    city = tuple([text[int(i)] for i in split])
    if city in data and sheet in data[city]:
        data[city][sheet].append(text)
    elif city in data and sheet not in data[city]:
        data[city][sheet] = [title, text]
    else:
        data[city] = {sheet: [title, text]}
    return data


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


def csvRead(file, split_no):
    data = {}
    n = 1
    f = open(file, 'rb')
    for line in f.readlines():
        '''
        try:
            reader = csv.DictReader(f)
            for row in reader:
                print(row)
        '''
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
        if n == 1:
            title = text
            print('%s拆分的列:' % file, [title[int(i)] for i in split_no])
        else:
            data = textSplit(text, split_no, data, title)
        n += 1
    f.close()
    return data


def excelRead(file, split_no):
    book = xlrd.open_workbook(file)
    data = {}
    for sheet in book.sheets():
        try:
            title = sheet.row_values(0)
            print('%s拆分的列:' % sheet.name, [title[int(i)] for i in split_no])
            for row in range(1, sheet.nrows):
                text = sheet.row_values(row)
                text_type = sheet.row_types(row)
                # get datatype,type: 0empty,1string, 2number, 3date, 4boolean, 5error
                text = list(map(textFormat, text_type, text))
                data = textSplit(text, split_no, data, title, sheet=sheet.name)
        except IndexError:
            print('数据错误', sheet.name)
    return data


def excelWrite(path, data, name, separator):  # data {city:{sheet:data}}
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
    for city, data_city in data.items():
        print(city)
        workbook = xlsxwriter.Workbook(str(path+'//%s'+separator+name+'.xlsx') % re.sub('[\:\/\|\<\>\*\"\?\r\n\t//"]', '', (','.join([str(i) for i in city]))))
        "文件名是第一层data的key值的逗号连接文本"
        for sheetname, data_w in data_city.items():
            worksheet = workbook.add_worksheet(name=sheetname)
            # 5.4 用行列表示法（行列索引都从0开始）写入数字
            for row in range(len(data_w)):
                for col in range(len(data_w[row])):
                    if row == 0:
                        worksheet.write(row, col, data_w[row][col], workbook.add_format(top))
                    else:
                        worksheet.write(row, col, data_w[row][col], workbook.add_format(oth))
        workbook.close()


def main(file, path_mail=path_mail, path_filed=path_filed, separator=separator):
    shotname, extension = os.path.splitext(file)
    winsound.Beep(220, 800)
    print('\n正在拆分 %s' % file, end='')
    time_id = time_no()
    try:
        while True:
            code = input('请输入要拆分的列(从0列开始,多列请用英文逗号分开):').split(',')
            time_start = time.time()
            try:
                if extension in ('.xlsx', '.xls'):
                    data = excelRead(file, code)
                elif extension in ('.csv'):
                    data = csvRead(file, code)
                print('loading...')
                excelWrite(path_mail, data, shotname+"_S"+time_id, separator)
                time_end = time.time()
                print('用时 %.2f 秒' % (time_end-time_start))
            except ValueError:
                break
        st = input('请对原文件重命名:')
        if st.strip(' '):
            shutil.copyfile(file, path_mail+'//'+st+separator+shotname+"_S"+time_id+extension)
        shutil.move(file, path_filed+'//'+shotname+"_SO"+time_id+extension)
        winsound.Beep(1000, 800)
    except FileNotFoundError:
        print('文件不存在', file)
    except UnboundLocalError:
        print('文件类型错误', file)
    except AttributeError:
        print('数据错误', file)


os.chdir(path_mail)
mail_list = [i for i in os.listdir('.') if i not in ['.DS_Store','desktop.ini']]
if __name__ == '__main__':
    print('\n^_^欢迎使用智文文件助手^_^\n')
    os.chdir(path_file)
    file_list = [i for i in os.listdir('.') if i not in ['.DS_Store','desktop.ini']]
    if file_list:
        confirm_value = 'no'
        if len(mail_list)>0:
            confirm_value = input('mail文件夹存在文件,是否继续.确认继续请输入"yes":')
        if confirm_value.lower() == "yes" or len(mail_list)==0:
            list(map(main, file_list))
        else:
            print('取消了操作')
    else:
        print('空文件,无需拆分')


# input('\npress any key to exit')
