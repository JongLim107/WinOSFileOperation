# coding=utf-8
# author:JongLim107
# python2

from time import ctime
import threading
import urllib2
import re
import os

print ">>> start %s" % ctime()

url = "http://dl.upload08.com/files/Serial/Friends/"
html = urllib2.urlopen(url).read()  # Open and Get raw data from website
urls = re.findall('href="(.*?)"', html)

path = os.path.abspath(os.curdir)  # Get currently path
# exactlyName = path+'\\'+name.encode('gbk')  # for chinese world


def graphsub(uri, name):
    print (">>>>> run %s" % ctime())
    html2 = urllib2.urlopen(uri).read()
    nameArr = re.findall('href="(.*?)"', html2)

    fileName = path + ('\\'+'%s.txt' % name)
    if os.path.exists(fileName):
        os.remove(fileName)

    File = open(fileName, 'w')
    for nstr in nameArr:
        if("../" in nstr):
            pass
        else:
            # print >> File, uri+nstr
            File.write(uri+nstr + "\n")
    File.close()


i = 1
threads = []
for uri in urls:
    if uri == "../":
        pass
    else:
        uri = url+uri+"1080p%20x265/"
        threads.append(threading.Thread(target=graphsub,
                                        args=(str(uri), 'Season%d' % i)))
        print (">>>>> add %s" % ctime())
        i += 1

if __name__ == '__main__':

    for t in threads:
        t.setDaemon(True)
        t.start()

    t.join()

    print (">>> stop! %s" % ctime())
