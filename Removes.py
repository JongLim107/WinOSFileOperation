# coding=utf-8
# Python3

import os
import xlrd
from openpyxl import Workbook
from time import ctime

count = cnt = ct = 1
workbook = Workbook()
worksheet = workbook.worksheets[0]
xlsName = ['Calculate.xls', 'Calculate1.xls']


def save2xlsx(fullname, col):
    global count, cnt, ct
    index = count if col == 1 else cnt if col == 2 else ct
    worksheet.cell(row=index, column=col).value = fullname
    if(col == 1):
        count += 1
    elif(col == 2):
        cnt += 1
    else:
        ct += 1


def removeFile(fullname):
    try:
        os.remove(fullname)
        save2xlsx(fullname, 3)
    except:
        pass


def readFile(file1, file2):
    sheet1 = xlrd.open_workbook(file1).sheet_by_index(0)
    sheet2 = xlrd.open_workbook(file2).sheet_by_index(0)
    rows1 = sheet1.col_values(0)
    rows2 = sheet2.col_values(0)
    for i in range(sheet1.nrows):
        names = rows1[i].split('\\')
        n1 = names[len(names) - 1]
        n1 = n1.split('.')[0]  # for duplicate images
        save2xlsx(rows1[i], 1)
        for j in range(sheet2.nrows):
            names = rows2[j].split('\\')
            n2 = names[len(names) - 1]
            # if n1 == n2:  # for duplicate mp3
            if n1 in n2:  # for duplicate images
                save2xlsx(rows2[j], 2)
                removeFile(rows2[j])
                pass


if __name__ == "__main__":
    print('\n>>> start ' + ctime() + '\n')
    file1 = os.getcwd() + '\\' + xlsName[0]
    file2 = os.getcwd() + '\\' + xlsName[1]
    readFile(file1, file2)

    workbook.save(filename='D:\\PythonProject\\PythonSpider\\RemovesJPG2.xls')
    print('\n<<< end ' + ctime())
