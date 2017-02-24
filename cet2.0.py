# -*- coding:UTF-8 -*-

import requests
import os,sys
import xlrd,xlwt

errorlist=["0","Referer Error","Query Error","Encoding Error","傻逼Python又出错了","傻逼GBK编码又有BUG了","辣鸡成电网又断了"]

def query(i,zkzh,xm,jb):
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
            ressheet.write((i-1)%5000+1, 0, xm)
            ressheet.write((i-1)%5000+1, 1, jb)
            ressheet.write((i-1)%5000+1, 2, dictr[5])
            ressheet.write((i-1)%5000+1, 3, dictr[2])
            ressheet.write((i-1)%5000+1, 4, dictr[3])
            ressheet.write((i-1)%5000+1, 5, dictr[4])
            ressheet.write((i-1)%5000+1, 6, dictr[10])
        except:
            return 4
        else:
            return 0

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
ressheet.write(i, 0, '姓名')
ressheet.write(i, 1, '级别')
ressheet.write(i, 2, '总分')
ressheet.write(i, 3, '听力')
ressheet.write(i, 4, '阅读')
ressheet.write(i, 5, '写作翻译')
ressheet.write(i, 6, '口语')
ok=0;
error=0;

for i in range(9999,nrows+1):
    zkzh=table.cell(i,5).value.strip()
    xm=table.cell(i,6).value.strip()
    jb=table.cell(i,1).value.strip()
    k=query(i,zkzh,xm,jb)
    if (k==0):
        print("[%05d] %s - OK!" % (i,xm))
        ok+=1
    else:
        print("[%05d] %s - %s!" % (i,xm,errorlist[k]))
        k=query(i,zkzh,xm,jb)
        if (k==0):
            print("Try Again [%03d] %s - OK!" % (i,xm))
            ok+=1
        else:
            print("Try Again [%05d] %s - %s!" % (i,xm,errorlist[k]))
            k=query(i,zkzh,xm,jb)
            if (k==0):
                print("Try Again [%05d] %s - OK!" % (i,xm))
                ok+=1
            else:
                print("Try Again [%05d] %s - %s!" % (i,xm,errorlist[k]))
                error+=1
    if (i%5000==0):
        resworkbook.save('Result'+str(int(i/5000))+'.xls')
        print('Saved Workbook %d'%(i/5000))

print("查询完毕，共计 %d 条结果"%(i))
print("其中 %d 条结果查询失败"%(error))
if (i%5000!=0):
    resworkbook.save('Result'+str(int(i/5000+1))+'.xls')
os.system("pause")
