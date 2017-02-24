# -*- coding:UTF-8 -*-

import requests
import os
import sys
import urllib.request
import time

def query(nj,yx,bj,zh,zkzh,xm,jb,xff):
    u=requests.session()
    u.cookies.clear()
    url = "http://www.chsi.com.cn/cet/query?zkzh=%s&xm=%s" % (zkzh, urllib.request.quote(xm))
    headers = {"Referer": "http://www.chsi.com.cn/cet/",
               "X-Forwarded-For":"127.0.0."+str(xff)
               }
    r = u.get(url, headers=headers).text.partition(zkzh)[2].partition("colorRed\">")[2]
    xx=r.partition("</span>")[0].strip()
    x1=r.partition("力")[2].partition("<td>")[2].partition("</td>")[0].strip()
    x2 = r.partition("读")[2].partition("<td>")[2].partition("</td>")[0].strip()
    x3 = r.partition("译")[2].partition("<td>")[2].partition("</td>")[0].strip()
    x4 = r.partition("级")[2].partition("colspan=\"2\">")[2].partition("</td>")[0].strip()
    uestc={
        '00':'英才',
        '01':'通信',
        '02':'电工',
        '03':'微固',
        '04':'物理',
        '05':'光电',
        '06':'计科',
        '07':'自动',
        '08':'机电',
        '09':'生科',
        '10':'数学',
        '11':'经管',
        '12':'政管',
        '13':'外语',
        '14':'体育',
        '16':'马院',
        '17':'能源',
        '18':'资环',
        '19':'空天',
        '20':'格院',
        '21':'医学',
        '22':'信软',
        '23':'示范性软件学院',
        '24':'电子科学',
        '26':'通信抗干扰',
        '95':'其他',
        '97':'心理',
        '98':'军事',
        '99':'艺术',
    }
    for k in uestc:
        yx=yx.replace(k,uestc[k])
    jb=jb.partition("CET")[2]
    if (x4=="--"):
        x4='\\'
    if (xx==""):
        return 0
    else:
        res="%s - %s - %s - %s - %s - %s - %s - %s - %s - %s - %s\n"%(nj,yx,bj,xh,xm,jb,xx,x1,x2,x3,x4)
        data.write(res)
        return 1

print("准备查询")
os.system("chcp 437")
os.system("cls")

students=open('students_all.txt','r',encoding='utf-8')
data=open('data.txt','w',encoding='utf-8')
type = sys.getfilesystemencoding()
i=1
ok=0;
error=0;
xff=123;
while True:
    nj=students.readline().partition("\n")[0]
    yx=students.readline().partition("\n")[0]
    bj=students.readline().partition("\n")[0]
    xh=students.readline().partition("\n")[0]
    zkzh=students.readline().partition("\n")[0]
    xm=students.readline().partition("\n")[0]
    jb=students.readline().partition("\n")[0]
    if len(bj)==0:
        break
    else:
        if (query(nj,yx,bj,xh,zkzh,xm,jb,xff)==1):
            print("[%05d] %s - OK!" % (i,zkzh))
            ok+=1
        else:
            xff+=1
            print("[%05d] %s - ERROR!" % (i,zkzh))
            if (query(nj,yx,bj,xh,zkzh,xm,jb,xff)==1):
                print("Try Again [%03d] %s - OK!" % (i,zkzh))
                ok+=1
            else:
                xff+=1
                print("Try Again [%05d] %s - ERROR!" % (i,zkzh))
                if (query(nj,yx,bj,xh,zkzh,xm,jb,xff)==1):
                    print("Try Again [%05d] %s - OK!" % (i,zkzh))
                    ok+=1
                else:
                    xff+=1
                    print("Try Again [%05d] %s - ERROR!" % (i,zkzh))
                    error+=1
        i+=1
        '''if (i>100):
            break'''
        '''if (i%43==0):
            print("Pausing for auto-detection...")
            time.sleep(5)'''

print("查询完毕，共计 %d 条结果"%(i-1))
print("其中 %d 条结果查询失败"%(error))
#print("===%d completed and %d failed.==="%(ok,error))
students.close()
data.close()
os.system("pause")
