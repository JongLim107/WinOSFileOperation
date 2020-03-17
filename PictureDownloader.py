import distutils.dir_util
import urllib.request
import os
import requests
import sys
import time

from bs4 import BeautifulSoup
from xlrd import open_workbook
from xlutils.copy import copy

workDir = os.getcwd()+'/'
destPath = workDir + 'Download/'
newExcel = workDir+'PictureUrls_copy.xls'
srcExcel = 'PictureUrlsFinal.xls'
maxRetry = 4
dosum = False


def reName(n, new):
    for file in os.listdir(destPath):
        if file == n:
            try:
                os.rename(destPath+file, destPath+new)
            except:
                pass
            break


def readFile(file, rb):
    result = []
    if(file.endswith('.txt')):
        data = open(file, "r")
        for line in data:
            line = line.rstrip()
            urlsplit = line.split("/")
            result.append({
                'name': urlsplit[len(urlsplit) - 1].rstrip(),
                'url': line
            })
        return result
    else:
        sheet = rb.sheet_by_index(0)
        names = sheet.col_values(2)
        status = sheet.col_values(7)
        subName = sheet.col_values(3)
        for i in range(0, len(subName)):
            name = names[i].rstrip().replace(' ', '_')
            result.append({
                'name': name+'-'+subName[i],
                'state': status[i],
                'url': 'https://vhzc.hpb.gov.sg/vhz/phs/capi/cloud/getUserFoodImage.ashx?Service=hpbphs&name='+subName[i]+'&mime=image/jpeg'
            })
            # reName(name+'.jpeg', name+'-'+subName[i])
        return result


def checkFile(name):
    if os.path.exists(name):
        try:
            os.remove(name)
        except Exception as e:
            print(e)
            return True
    return False


def tryDownload(meal, cnt=1):
    name = meal['name']
    if meal['state'] != '':
        # Update: if last time is successful, just skip this onw
        return 7
    if os.path.exists(destPath+name) == True:
        print('*** '+name+' is exist already.')
        return 0
    while cnt < maxRetry:
        try:
            urllib.request.urlretrieve(meal['url'], destPath+name)
            print(cnt, '| ' + name)
            break
        except:
            cnt += 1
    return cnt


def containSub(subs, sub):
    # for i in range(len(subs)):
    for i in range(10000):
        if sub == subs[i]:
            return i
    return -1


def localFile2Excel():
    rb = open_workbook(workDir+srcExcel)
    wb = copy(rb)
    sh = wb.get_sheet(0)
    sheet = rb.sheet_by_index(0)
    subs = sheet.col_values(3)
    status = sheet.col_values(6)

    res = os.listdir(destPath)
    for name in res:
        arr = name.split('-')
        index = containSub(subs, arr[len(arr) - 1])
        if index != -1:
            sh.write(index, 7, 1)
            if status[index].startswith('All') == True:
                sh.write(index, 6, 'Succeed at: 8')
            if status[index].endswith('0') == True:
                sh.write(index, 6, 'Succeed at: 8')
        else:
            sh.write(index, 6, 'All Failed')
    wb.save(newExcel)
    return


if __name__ == "__main__":
    if dosum == True:
        localFile2Excel()
        exit()

    rb = open_workbook(workDir+srcExcel)
    workBook = copy(rb)
    newSheet = workBook.get_sheet(0)

    print('\n>>> start ' + time.ctime() + '\n')
    if os.path.exists(destPath) != True:
        distutils.dir_util.mkpath(destPath)
    if os.path.exists(newExcel):
        os.remove(newExcel)

    rows = readFile(srcExcel, rb)
    length = len(rows)
    for i in range(10000, 20001):
        value = tryDownload(rows[i])
        prin = False
        if value != maxRetry:
            prin = True
            # Update: add new flag to indicate sucessful or not, easy for summation
            newSheet.write(i, 7, 1)
        else:
            prin = i % 10 == 0

        if prin == True:
            print('>>> ' + str(round(i/length*100, 2)) + '%', i, value)

        if value != 7:
            res = 'All Failed' if value == maxRetry else (
                'Succeed at: ' + str(value))
            newSheet.write(i, 6, res)

        if i % 100 == 1:
            workBook.save(newExcel)
            time.sleep(5)
    workBook.save(newExcel)
