# -*- coding: utf-8 -*-
"""
Created on Thu Mar  8 11:17:20 2018

@title: excel merge

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


# 配置文件路径
path = "/mailsend/"  # 全路径
path_original = path + "original_file/"  # 待拆分文件路径
path_original_done = path + 'original_file_done/'  # 已拆分文件保存路径
path_send = path + "send_file/"  # 待发送文件路径
path_send_done = path + 'send_file_done/'  # 已发送文件保存路径


# 定义时间编号
def time_no():
    no = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
    return no


def csv_read(data, file):
    # 注意csv文件的编码问题
    shotname,extension = os.path.splitext(file) #获取文件拓展名
    f= open(file,'rb')
    for line in f.readlines():
        try:
            code = chardet.detect(line)['encoding']
            if code in ('gbk','gb2312','gb18030','utf-8'):
                t=line.decode(code) #读取这一行的编码
            else:
                t=line.decode('gb2312')
        except :
            try:
                t=line.decode('gbk') #尝试gbk,因为gb2312和gbk会有错误
            except :
                try:
                    t=line.decode('utf-8')
                except:
                    continue

        text = re.sub('[\r\n\t"]','',t).split(',')

        try:
            if text != data['sheet'][0][0:-1]:
                data['sheet'] = data['sheet'] + [re.sub('[\r\n\t"]','',t).split(',')+ [shotname]]
        except:
            data['sheet'] = [re.sub('[\r\n\t"]','',t).split(',')+['source']]
    f.close()
    return data

#read_excel:读取excel
def excel_read(data,file,mode=0):
    shotname,extension = os.path.splitext(file) #获取文件拓展名
    book = xlrd.open_workbook(file)
    for sheet in book.sheets():
        data_file = [sheet.row_values(i) for i in range(0,sheet.nrows)]
        data_file[0].append('source')
        """
        j = 0;text_type = sheet.row_types(i)
        for cell in data_file[1]:
            ctype = text_type[j]
            if ctype == 2 and cell % 1 == 0:  # 如果是整形
                 text[j] = int(cell)
            elif ctype == 3 and cell % 1 ==0:
                text[j] = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y/%m/%d')
            elif ctype == 3:
                text[j] = xlrd.xldate.xldate_as_datetime(cell, 0).strftime('%Y/%m/%d %H:%M:%S')
            elif ctype == 4:
                text[j] = True if cell == 1 else False
            j += 1
        """
        if mode == '1' :
            sheet_name = sheet.name  #新表的sheet名
            source = shotname
        else:
            sheet_name = 'sheet'
            source = shotname+'/'+sheet.name

        #添加一列:数据来源,同名sheet合并取file名;全sheet合并取sheet名
        for line in data_file[1:]:
            line.append(source)

        try:
            data[sheet_name] = data[sheet_name] + data_file[1:]
        except:
            data[sheet_name] = data_file
        #读取excel,每个sheet数据为一个一个list,一个excel保存到dict
    return data

#write_excel:写入到excel中
def write_excel(name,data):
    workbook = xlsxwriter.Workbook(str(name+'.xlsx'))
    top = {"font_name":u'微软雅黑','border':4,'border_color':'50616d','font_size':11,'bold': True,'bg_color':'#fffbf0',}
    oth = {"font_name":u'微软雅黑','border':4,'border_color':'50616d','font_size':10}
    # 5.4 用行列表示法（行列索引都从0开始）写入数字
    for sheet , data_w in  data.items():
        worksheet = workbook.add_worksheet(name=sheet)
        for row in range(len(data_w)):
            for col in range(len(data_w[row])):
                if row==0:worksheet.write(row,col,data_w[row][col],workbook.add_format(top)) #row,col,*args
                else:worksheet.write(row,col,data_w[row][col],workbook.add_format(oth))
    workbook.close()

#读取文件并拆分到excel
os.chdir(path_original)
file_list_original = os.listdir('.')

if len(file_list_original)>0:
    data = {}
    name = input('请输入合并后文件名称:\n')
    mode = input('请输入合并模式(默认0模式):\n   1.同名sheet合并\n   0.全部sheet合并\n')
    time_start = time.time()
    for file in file_list_original:
        shotname,extension = os.path.splitext(file) #获取文件拓展名
        if extension in ['.csv','.xls','.xlsx']:
            try:
                if extension == '.csv':
                    data = csv_read(data,file)
                elif extension in {'.xls','.xlsx'}:
                    data = excel_read(data,file,mode)
                shutil.move(file,path_original_done+shotname+"_CO"+time_no()+extension)
            except:
                print('❤%s 未知异常' % file)
        else:
            print('❤%s 文件格式异常' % file)

    if len(data) > 0:
         write_excel(path_send+name+"_C"+time_no(),data)
         time_end=time.time()
         print('用时 %.2f 秒'% (time_end-time_start))

else:
    print("❤空文件夹,无需合并❤")
winsound.Beep(1000,300) #其中400表示声音大小，1000表示发生时长，1000为1秒
#input('\n\n请按任意键退出')
