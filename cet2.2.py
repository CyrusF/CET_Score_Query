# -*- coding:UTF-8 -*-

import requests
import os,sys
import xlrd,xlwt
import urllib
import threading
import queue
import time

errorlist=["OK","你忘了Referer了","准考证号和姓名不匹配","你忘了Encode(GBK)了","傻逼Python又出错了","傻逼GBK编码又有BUG了","辣鸡成电网又断了"]

def query(i,zkzh,xm,jb,nj,yx,xb):
    u=requests.session()
    u.cookies.clear()
    try:
        url = "http://cet.99sushe.com/getscore"
        headers = {"Referer": "http://cet.99sushe.com"}
        postdata = {"id": zkzh, "name": xm[:2].encode('GBK')}
    except:
        return 5
    else:
        pass
    try:
        r = u.post(url, headers=headers, data=postdata).text.strip()
    except:
        return 6
    else:
        pass
    if (len(r)==1):
        return int(r)
    else:
        try:
            dictr = r.split(',')
            if (dictr[9].partition('--')[1]=='--'):
                dictr[9]='\\'
            if (dictr[10].partition('--')[1]=='--'):
                dictr[10]='\\'
            ressheet.write(i, 0, nj)
            ressheet.write(i, 1, yx)
            ressheet.write(i, 2, xm)
            ressheet.write(i, 3, xb)
            ressheet.write(i, 4, jb)
            ressheet.write(i, 5, dictr[5])
            ressheet.write(i, 6, dictr[2])
            ressheet.write(i, 7, dictr[3])
            ressheet.write(i, 8, dictr[4])
            ressheet.write(i, 9, dictr[10])
        except:
            return 4
        else:
            return 0

def query2(i,zkzh,xm,jb,nj,yx,xb):
    #如果GBK编码错误换位学信网的数据源
    u = requests.session()
    u.cookies.clear()
    try:
        url = "http://www.chsi.com.cn/cet/query?zkzh=%s&xm=%s" % (zkzh, urllib.request.quote(xm))
        headers = {"Referer": "http://www.chsi.com.cn/cet/",
                   "X-Forwarded-For": "127.0.0.1234"
                   }
        r = u.get(url, headers=headers).text
    except:
        return 6
    else:
        pass
    try:
        r = r.partition(zkzh)[2].partition("colorRed\">")[2]
        xx = r.partition("</span>")[0].strip()
        x1 = r.partition("力")[2].partition("<td>")[2].partition("</td>")[0].strip()
        x2 = r.partition("读")[2].partition("<td>")[2].partition("</td>")[0].strip()
        x3 = r.partition("译")[2].partition("<td>")[2].partition("</td>")[0].strip()
        x4 = r.partition("级")[2].partition("colspan=\"2\">")[2].partition("</td>")[0].strip()
        ressheet.write(i, 0, nj)
        ressheet.write(i, 1, yx)
        ressheet.write(i, 2, xm)
        ressheet.write(i, 3, xb)
        ressheet.write(i, 4, jb)
        ressheet.write(i, 5, xx)
        ressheet.write(i, 6, x1)
        ressheet.write(i, 7, x2)
        ressheet.write(i, 8, x3)
        ressheet.write(i, 9, x4)
    except:
        return 4
    else:
        return 0

def work():
    if not free.empty():
        x=free.get()
        global i
        i += 1
        #print(">>> Get %s for processing %d"%(x,i))
        global ok
        global error
        zkzh = table.cell(i, 5).value.strip()
        xm = table.cell(i, 6).value.strip()
        jb = table.cell(i, 1).value.strip()
        nj = table.cell(i, 12).value.strip()
        yx = table.cell(i, 11).value.strip()
        xb = table.cell(i, 7).value.strip()
        k = query(i, zkzh, xm, jb, nj, yx, xb)
        if (k == 0):
            print("[%05d] %s - %s!" % (i, xm, errorlist[k]))
            ok += 1
        else:
            print("[%05d] %s - %s!" % (i, xm, errorlist[k]))
            if (k == 5):
                k = query2(i, zkzh, xm, jb, nj, yx, xb)
            else:
                k = query(i, zkzh, xm, jb, nj, yx, xb)
            if (k == 0):
                print("Try Again [%05d] %s - %s!" % (i, xm, errorlist[k]))
                ok += 1
            else:
                print("Try Again [%05d] %s - %s!" % (i, xm, errorlist[k]))
                if (k == 5):
                    k = query2(i, zkzh, xm, jb, nj, yx, xb)
                else:
                    k = query(i, zkzh, xm, jb, nj, yx, xb)
                if (k == 0):
                    print("Try Again [%05d] %s - %s!" % (i, xm, errorlist[k]))
                    ok += 1
                else:
                    print("Try Again [%05d] %s - %s!" % (i, xm, errorlist[k]))
                    error += 1
        #print("Free %s"%(x))
        free.put(x)


print("准备查询")
os.system("chcp 437")
os.system("cls")

workbook=xlrd.open_workbook('16UESTC.xlsx')
table=workbook.sheet_by_name(u'Sheet1')
nrows = table.nrows
resworkbook=xlwt.Workbook()
ressheet=resworkbook.add_sheet(u'Sheet1')
type = sys.getfilesystemencoding()
i=0
ok=0
error=0
ressheet.write(i, 0, '年级')
ressheet.write(i, 1, '院系')
ressheet.write(i, 2, '姓名')
ressheet.write(i, 3, '性别')
ressheet.write(i, 4, '级别')
ressheet.write(i, 5, '总分')
ressheet.write(i, 6, '听力')
ressheet.write(i, 7, '阅读')
ressheet.write(i, 8, '写作翻译')
ressheet.write(i, 9, '口语')

print("获取到 %d 条学生信息"%(nrows-1))

nthread=1000

free=queue.LifoQueue(nthread)
for t in range(nthread):
    free.put("thread"+str(t))

while i<=nrows-2:
    t = threading.Thread(target=work)
    t.start()

time.sleep(2)
print("查询完毕，共计 %d 条结果"%(i))
print("其中 %d 条结果查询失败"%(error))
resworkbook.save('Result.xls')
os.system("pause")
