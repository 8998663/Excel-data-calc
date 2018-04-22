#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
@Time    : 2018/4/21 下午3:56
@Author  : ldx
@contact : ldx.9@163.com
@File    : main.py
@Software: PyCharm
'''

'''
将后缀名为xls和xlsx的文件，与本代码放在同一个文件夹中，执行python main.py ；
本程序可将多个Excel文件中的指定列（符合一定条件的）拼接到一个excel中，并命名为1.xls
规则为，第5列中，值等于7的情况下，第3列和第7列复制到新的excel中。
'''


import xlrd
import xlwt
import os

def open_excel(file):
    try:
        data = xlrd.open_workbook(file)             #打开excel文件
        return data
    except Exception as e:
        print str(e)

def excel_table_byname(file_path,colnameindex=0):
    data = open_excel(file_path)                     #打开excel文件
    table = data.sheet_by_index(1)                  #打开第二个sheet
    nrows = table.nrows                             #获取总行数
    ncols = table.ncols                             #获取总列数
    row_result= []
    index_time = "Step_Time(s)"
    for rownum  in range(0,nrows):           #遍历每一行的内容
        row = table.row_values(rownum)          #根据行号获取行
        # print row
        if row[4] == 7:
            row_result.append(rownum)               #找到第4列中，值等于7的第一个行号
            break
        else:
            continue
    step_times = table.col_values(3,start_rowx=row_result[0],end_rowx=None)     #获取第3列，从值等于7的第一个行号到最后一行
    voltage = table.col_values(7,start_rowx=row_result[0],end_rowx=None)        ##获取第7列，从值等于7的第一个行号到最后一行
    return step_times,voltage


#获取某个目录下，后缀名是xls和xlsx的文件名称
def file_name(file_dir):
    L = []
    for dirpath, dirnames, filenames in os.walk(file_dir):
        for file in filenames:
            if ((os.path.splitext(file)[1] == '.xls') or (os.path.splitext(file)[1] == '.xlsx')):
                L.append(file)
    return L

        # return data

if __name__ == "__main__":
    path_dir = os.getcwd()          #获取程序当前路径
    print path_dir
    path = file_name(path_dir)      #获取当前目录下的文件名
    path.sort()                     #文件名排序
    print path
    step = []
    volta = []
    result = []
    # print len(path)
    for i in range(len(path)):
        step,volta = excel_table_byname(path[i])
        result.append(step)
        result.append(volta)

    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('result')
    for i in range(0,len(result)):
        for j in range(len(result[i])):
            worksheet.write(j, i, result[i][j])

    workbook.save('1.xls')




