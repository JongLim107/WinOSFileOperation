# coding=utf-8
# python3

import os
from openpyxl import Workbook
from time import ctime
# from korean import Noun

count = colus = cnt = 1
workbook = Workbook()
worksheet = workbook.worksheets[0]


def save2xlsx(fullname, col):
    global count, colus, cnt
    index = count if col == 1 else colus if col == 2 else cnt
    worksheet.cell(row=index, column=col).value = fullname
    if(col == 1):
        count += 1
    elif(col == 2):
        colus += 1
    else:
        cnt += 1


def searchFolder(dir, col):
    for file in os.listdir(dir):
        fullname = dir + '\\' + file
        if os.path.isdir(fullname):
            searchFolder(fullname, col)
        else:
            if '.mp3' in file.lower():
                save2xlsx(fullname, col)
            if '???' in file:
                print(fullname)
                if '.lrc' in file.lower():
                    try:
                        os.remove(fullname)
                        save2xlsx(fullname, 3)
                    except:
                        pass


if __name__ == "__main__":
    print('\n>>> start ' + ctime() + '\n')
    for adrr in ['E:', 'F:']:
        for file in os.listdir(adrr):
            dir = os.path.join(adrr, '\\' + file)
            if os.path.isdir(dir) and 'music' in dir.lower():
                searchFolder(dir, 1)
    searchFolder(os.getcwd(), 2)
    workbook.save(filename='D:\\Python\\PythonSpider\\Calculate2.xls')
    print('\n<<< end ' + ctime())
