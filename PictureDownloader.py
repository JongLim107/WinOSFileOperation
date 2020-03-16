import distutils.dir_util
import urllib.request
import os
import requests
import sys

from bs4 import BeautifulSoup
from time import ctime
from xlrd import open_workbook
from xlutils.copy import copy

workDir = os.getcwd()+'/'
destPath = workDir + 'Download/'
newExcel = workDir+'PictureUrls_copy.xls'
srcName = 'PictureUrls.xls'
maxRetry = 5


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
        subName = sheet.col_values(3)
        for i in range(1, len(subName)):
            name = names[i].rstrip().replace(' ', '_')
            result.append({
                'name': name+'-'+subName[i],
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


def tryDownload(meal, url, cnt=1):
    if os.path.exists(destPath+meal) == True:
        # print('*** '+meal+' is exist already.')
        return 0
    while cnt <= maxRetry:
        try:
            # line ending chars throws 404 issue
            urllib.request.urlretrieve(url, destPath+meal)
            print(cnt, '| ' + meal)
            break
        except Exception as e:
            cnt += 1
    return cnt


if __name__ == "__main__":
    rb = open_workbook(workDir+srcName)
    workBook = copy(rb)
    newSheet = workBook.get_sheet(0)

    print('\n>>> start ' + ctime() + '\n')
    if os.path.exists(destPath) != True:
        distutils.dir_util.mkpath(destPath)
    if os.path.exists(newExcel):
        os.remove(newExcel)

    rows = readFile(srcName, rb)
    length = len(rows)
    for i in range(1, length):
        name = rows[i]['name']
        url = rows[i]['url']
        value = tryDownload(name, url)
        if value != 6 and value != 0:
            print('>>> ' + str(round(i/length*100, 2)) + '%', i, value)
        res = ('All Failed: ' if value == 6 else 'Succeed at: ') + str(value)
        newSheet.write(i+1, 6, res)
        if i % 100 == 1:
            workBook.save(newExcel)
    workBook.save(newExcel)
