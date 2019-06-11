# -*- coding: utf-8 -*-
"""
Created on Fri Mar 15 15:38:35 2019

@title: file

@author: 饭未眠

@email: ifansay.chn@qq.com
"""
import os, sys
import datetime
import time
import re
import shutil
import xlrd
import xlsxwriter
import winsound
import pathlib


# 改变日期格式
def transTime(timeStamp):
    dateArray = datetime.datetime.fromtimestamp(timeStamp)  # utc+8:00
    otherStyleTime = dateArray.strftime("%Y/%m/%d %H:%M:%S")
    return otherStyleTime


# 对文件大小格式化
def formatSize(size):
    kb = size / 1024
    if kb >= 1024:
        M = kb / 1024
        if M >= 1024:
            G = M / 1024
            return "%.2fG" % (G)
        else:
            return "%.2fM" % (M)
    else:
        return "%.2fkb" % (kb)


# 获取文件信息
def getFileInfo(file):
    shotname, extension = os.path.splitext(file)  # 获取文件拓展名
    size = formatSize(os.path.getsize(file))
    ctime = transTime(os.path.getctime(file))  # 创建时间
    atime = transTime(os.path.getatime(file))  # 访问时间
    mtime = transTime(os.path.getmtime(file))  # 修改时间
    return [file, shotname, extension, size, ctime, atime, mtime]


# 添加文件信息
def addInfo(file_list):
    info_dict = [['文件全名', '文件名', '类型', '大小', '创建时间', '最后访问时间', '最后修改时间', '标签']]
    file_list = [i for i in file_list if os.path.isfile(i)]
    for file in file_list:
        info_dict.append(getFileInfo(file))
    return info_dict


# 写入excel中
def writeExcel(path, data, file):  # 此处的data为dict
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet()
    for row in range(len(data)):
        for col in range(len(data[row])):
            worksheet.write(row, col, data[row][col])
        # 5.7 关闭并保存文件
    workbook.close()


# 读取excel
def readExcel(file):
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(0)
    data = []
    for i in range(1, sheet.nrows):
        text = sheet.row_values(i)
        data.append(text)
    return data


# 移动文件
def shutilFile(file_info, exc_char):
    for file in file_info:
        if file[0].find(exc_char) == -1:
            try:
                shutil.move(file[0], file[-1])  # 如果文件名称重复,重命名?
            except:
                new_path_file = pathlib.Path("%s/%s" % (file[-1], file[0]))
                old_path_file = pathlib.Path("%s" % file[0])
                if not old_path_file.is_file():
                    print('%s文件不存在,请检查文件是否重名或已被移动' % file[0])
                elif new_path_file.is_file():
                    print('%s/%s同名文件已存在,无法移动' % (file[-1], file[0]))
                else:
                    print('%s 归档失败,请检查文件是否重名或已被移动' % file[0])
                continue
                return file[0]


# 读取文件列表并保存
def obtainFile(path, codefile):
    file_list = [file for file in os.listdir('.') if file != codefile]
    file_info = addInfo(file_list)
    return file_list, file_info


# 创建文件夹
def creatPath(paths):
    for path in paths:
        path = re.sub('[\ \:\/\|\<\>\*\"\?\\\\r\n\t"]', '', path)
        isExists = os.path.exists(path)
        # 判断结果
        if not isExists:
            # 如果不存在则创建目录
            try:
                os.makedirs(path)
            except BaseException:
                print('%s 文件路径异常,请修改标签' % path)


arrange_file = '♠标签管理.xlsx'
# path = sys.path[0]
# path = "D:/test"
file_path, code_file = os.path.split(os.path.realpath(__file__))  # code_file
os.chdir(file_path)
file_list, file_info = obtainFile(file_path, code_file)
print('当前归档目录为:%s' % file_path)
# winsound.Beep(220, 800)
time_start = time.time()
if arrange_file in file_list:
    data = readExcel(arrange_file)
    tags = [str(i[-1]) for i in data]
    if len(set([i.strip() for i in tags if i.strip() != ''])) <= 1:
        print('请打开 %s 补充文件标签' % arrange_file)
    else:
        creatPath(tags)
        shutilFile(data, arrange_file)
else:
    shutil_mode = input('请输入归档模式:\n    1.一键归档\n    2.文件类型归档\n    3.文件标签归档\n')
    if shutil_mode in ('1', '2'):
        if shutil_mode == '1':
            tags = ['%s文件归档' % datetime.datetime.today().strftime('%Y%m%d')]
            file_info = [file+tags for file in file_info[1:]]
        elif shutil_mode == '2':
            tags = [file[2] for file in file_info[1:]]
            file_info = [file[:-1]+[file[2]] for file in file_info[1:]]
        exc_char = input('请输入不归档文件标记:').strip()
        if str.isspace(exc_char) or exc_char == '':
            exc_char = 'z'*1024
        creatPath(tags)
        shutilFile(file_info, exc_char)
    elif shutil_mode == '3':
        info_dict = addInfo(file_list)
        writeExcel(file_path, info_dict, arrange_file)
        print('请打开 %s 打标签(标签请以文本保存)' % arrange_file)
    else:
        print('\033[1;35;47m您取消了文件归档\033[0m')

winsound.Beep(1000, 800)
input('\npress any key to exit')
