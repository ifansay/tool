# -*- coding: utf-8 -*-
"""
Created on Wed Jul 12 16:52:22 2017

@title: excel split

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

# comment
"""版本号v2.6.2,更新日期2018-12-13
1.增加对单一文件的多次拆分支持"""

"""版本号v2.6.1,更新日期2018-09-09
1.增加对数字、布尔值作为拆分列的支持"""

"""版本号v2.6,更新日期2018-08-25
1.增加对多列拆分的需求
2.增加对整数,日期,布尔值的支持
3.修改拆分逻辑,提升效率"""


# 配置文件路径
path = '/mailsend/'  # 相对路径
path_original = path + "original_file/"  # 待拆分文件路径
path_original_done = path + 'original_file_done/'  # 已拆分文件保存路径
path_send = path + "send_file/"  # 待发送文件路径
path_send_done = path + 'send_file_done/'  # 已发送文件保存路径


# 定义时间编号
def time_to_str():
    no = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
    return no


# 定义拆分方法
def data_split(text, split, title, sheetname='sheet'):
    city = tuple([text[i] for i in split])
    if city in data and sheetname in data[city]:
        data[city][sheetname].append(text)
    elif city in data and sheetname not in data[city]:
        data[city][sheetname] = [title, text]
    else:
        data[city] = {sheetname: [title, text]}
    return data


# read_cav:读取csv
def csv_read(file, split):
    # 注意csv文件的编码问题
    data = {}
    n = 1
    f = open(file, 'rb')
    for line in f.readlines():
        try:
            code = chardet.detect(line)['encoding']
            if code in ('gbk', 'gb2312', 'gb18030', 'utf-8'):
                t = line.decode(code)  # 读取这一行的编码
            else:
                t = line.decode('gb2312')
        except:
            try:
                t = line.decode('gbk')  # 尝试gbk,因为gb2312和gbk会有错误
            except:
                try:
                    t = line.decode('utf-8')
                except:
                    continue
        text = re.sub('[\r\n\t"]', '', t).split(',')

        city = tuple([text[i] for i in split])
        if n == 1:
            title = text
        elif city in data:
            data[city]['sheet'].append(text)
        else:
            data[city] = {'sheet': [title, text]}
        n += 1
    f.close()
    return data


# read_excel:读取excel
def excel_read(file, split):
    book = xlrd.open_workbook(file)
    data = {}
    for sheet in book.sheets():
        try:
            title = sheet.row_values(0)
            for i in range(1, sheet.nrows):
                j = 0
                text = sheet.row_values(i)
                # type: 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
                text_type = sheet.row_types(i)

                for cell in text:
                    ctype = text_type[j]
                    if ctype == 2 and cell % 1 == 0:  # 如果是整形
                        text[j] = int(cell)
                    elif ctype == 3 and cell % 1 == 0:
                        text[j] = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y/%m/%d')
                    elif ctype == 3:
                        text[j] = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y/%m/%d %H:%M:%S')
                    elif ctype == 4:
                        text[j] = True if cell == 1 else False
                    j += 1

                city = tuple([text[i] for i in split])
                if city in data and sheet.name in data[city]:
                    data[city][sheet.name].append(text)
                elif city in data and sheet.name not in data[city]:
                    data[city][sheet.name] = [title, text]
                else:
                    data[city] = {sheet.name: [title, text]}

        except:
            if sheet.nrows == 0:
                print('工作表%s为空' % sheet.name)
            else:
                print('工作表%s数据异常' % sheet.name)
            continue
    return data


# write_excel:写入到excel中
def write_excel(path, data, name):  # 此处的data为dict
    "#4. 定义格式对象"
    top = {"font_name": u'微软雅黑',
           'border': 4,
           'border_color': '50616d',
           'font_size': 11,
           'bold': True,
           'bg_color': '#fffbf0'}
    oth = {"font_name": u'微软雅黑',
           'border': 4,
           'border_color': '50616d',
           'font_size': 10}
    for city, data_city in data.items():
        workbook = xlsxwriter.Workbook(str(path+'%s❤'+name+'.xlsx') %
                                       re.sub('[\ \:\/\|\<\>\*\"\?\\\\r\n\t"]', '', (','.join([str(i) for i in city]))))
        "1. 创建一个Excel文件"
        "2. 创建一个工作表sheet对象"
        for sheet, data_w in data_city.items():
            worksheet = workbook.add_worksheet(name=sheet)
            # 5.4 用行列表示法（行列索引都从0开始）写入数字
            for row in range(len(data_w)):
                for col in range(len(data_w[row])):
                    if row == 0:
                        worksheet.write(row, col, data_w[row][col], workbook.add_format(top))
                    else:
                        worksheet.write(row, col, data_w[row][col], workbook.add_format(oth))
        # 5.7 关闭并保存文件
        workbook.close()


# 定义转化字符
def cncode_str(a):
    try:
        split = a.strip(' ').split(',')
        split = [int(i) for i in split]
    except:
        print('\033[1;31;43拆分列错误')
    return split


# 定义拆分文件
def split_file(file, split):
    try:
        data = excel_read(file, split)
    except:
        try:
            data = csv_read(file, split)
        except:
            print('\033[1;31;43%s文件错误' % file)
            data = None
    return data


# 读取文件并拆分到excel
os.chdir(path_original)
file_list_original = os.listdir('.')

if len(file_list_original) > 0:
    for file in file_list_original:
        shotname, extension = os.path.splitext(file)  # 获取文件拓展名
        '''只读取表格文件,其他略过'''
        print("\n拆分 %s ,loading......" % file, end="")
        while True:
            winsound.Beep(220, 800)  # 其中400表示声音大小，1000表示发生时长，1000为1秒
            a = input('请输入要拆分的列号(从0列开始,多列请用英文逗号分开):==> ')
            if not a:
                break
            time_start = time.time()
            time_no = time_to_str()
            try:
                split = cncode_str(a)
                data = split_file(file, split)
                write_excel(path_send, data, shotname+"_S"+time_no)
                time_end = time.time()
                print('本次拆分用时 %.2f 秒' % (time_end-time_start))
            except:
                continue

        file_st = input('请对原文件重命名:==> ')
        if len(file_st.strip(' ')) > 0:
            shutil.copyfile(file, path_send+file_st+'❤'+shotname+"_S"+time_no+extension)
        # 复制文件重命名到发送文件夹
        shutil.move(file, path_original_done+shotname+"_SO"+time_no+extension)
        winsound.Beep(1000, 800)

else:
    print("\033[1;35m 空文件夹,无需拆分")
    winsound.Beep(1000, 300)

# input('\n\n请按任意键退出')
