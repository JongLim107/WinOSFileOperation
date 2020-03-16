import distutils.dir_util
import urllib.request
import os
import requests
import sys
from bs4 import BeautifulSoup

pathname = os.path.dirname(sys.argv[0])  # get current script path
destPath = os.path.abspath(pathname)+'/Download/'
print('Running script...\n' + destPath)


def checkFile(name):
    if os.path.exists(name):
        try:
            os.remove(name)
        except Exception as e:
            print(e)
            return True
    return False


with open("UrlList.txt", "r") as ins:  # read file for imagelinks
    for line in ins:
        line = line.rstrip()
        urlsplit = line.split("/")
        if os.path.exists(destPath) != True:
            print(distutils.dir_util.mkpath(destPath))
        try:
            print('>>> ' + line)
            page = requests.get(line)  # line ending chars throws 404 issue
            if page.status_code == 200:
                fileName = urlsplit[len(urlsplit) - 1].rstrip()
                if checkFile(fileName) == True:
                    raise 'File is exist already.'
                print('Url found and Saving file to ' + fileName)
                urllib.request.urlretrieve(line, destPath+fileName)
            else:
                print(page.status_code + " Url not found.")
        except Exception as e:
            print(e)
