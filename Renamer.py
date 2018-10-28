# coding=utf-8
# Python3

import os
import eyed3
from openpyxl import Workbook
from time import ctime

count = colus = 1
dictionary = ['(1)', '(2)', '（1）', '（2）', '_1', '_2', '- Copy']
workbook = Workbook()
worksheet = workbook.worksheets[0]


def save2xlsx(fullname, col):
    global count, colus
    index = count if col == 1 else colus
    worksheet.cell(row=index, column=col).value = fullname
    if(col == 1):
        count += 1
    else:
        colus += 1


def renameFile(fullname, newName):
    if os.path.exists(newName):
        try:
            os.remove(fullname)
            save2xlsx(fullname, 2)
        except:
            pass
    else:
        try:
            os.rename(fullname, newName)
            save2xlsx(fullname, 2)
        except:
            pass


def renameMp3(name):
    audiofile = eyed3.load(name)
    try:
        title = audiofile.tag.title
        print(title)
        for wd in dictionary:
            if wd in title:
                arr = title.split(wd)
                audiofile.tag.title = arr[0].strip()
                audiofile.tag.save()
                print(audiofile.tag.title)
                break
    except:
        pass


def checker(fullname, file):
    newName = None
    if '.mp3' in file.lower():
        save2xlsx(fullname, 1)
    for wd in dictionary:
        if wd in file:
            isMP3 = '.mp3' in file.lower()
            isLrc = '.lrc' in file.lower()
            arr = file.split(wd)
            newName = arr[0].strip() + \
                ('.mp3' if isMP3 else '.lrc' if isLrc else '')
            name_utf8 = fullname
            if isMP3:
                renameMp3(name_utf8)
            renameFile(fullname, newName)


def searchFolder(dir):
    for file in os.listdir(dir):
        fullname = dir + '\\' + file
        if os.path.isdir(fullname):
            searchFolder(fullname)
        else:
            checker(fullname, file)


if __name__ == "__main__":
    print('\n>>> start ' + ctime() + '\n')
    for adrr in ['D:', 'E:', 'F:']:
        for file in os.listdir(adrr):
            dir = os.path.join(adrr, '\\' + file)
            if os.path.isdir(dir) and 'music' in dir.lower():
                searchFolder(dir)
    workbook.save(filename='D:\\Python\\PythonSpider\\Renamer1.xls')
    print('\n<<< end ' + ctime())
