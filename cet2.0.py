# -*- coding:UTF-8 -*-

import requests
import os,sys
import xlrd,xlwt

def query(i,zkzh,xm,jb):
    u=requests.session()
    u.cookies.clear()
    url = "http://cet.99sushe.com/getscore"
    headers = {"Referer": "http://cet.99sushe.com"}
    postdata = {"id": zkzh, "name": xm.encode('GBK')}
    r = u.post(url, headers=headers, data=postdata).text.strip()
    if (r==""):
        return 0
    else:
        dict = r.split(',')
        if (dict[9].partition('--')[1]=='--'):
            dict[9]='\\'
        if (dict[10].partition('--')[1]=='--'):
            dict[10]='\\'
        ressheet.write(i, 0, xm)
        ressheet.write(i, 1, jb)
        ressheet.write(i, 2, dict[5])
        ressheet.write(i, 3, dict[2])
        ressheet.write(i, 4, dict[3])
        ressheet.write(i, 5, dict[4])
        ressheet.write(i, 6, dict[10])
    return 1

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

for i in range(1,nrows):
    zkzh=table.cell(i,5).value
    xm=table.cell(i,6).value
    jb=table.cell(i,1).value
    if (query(i,zkzh,xm,jb)==1):
        print("[%05d] %s - OK!" % (i,zkzh))
        ok+=1
    else:
        print("[%05d] %s - ERROR!" % (i,zkzh))
        if (query(i,zkzh,xm,jb)==1):
            print("Try Again [%03d] %s - OK!" % (i,zkzh))
            ok+=1
        else:
            print("Try Again [%05d] %s - ERROR!" % (i,zkzh))
            if (query(i,zkzh,xm,jb)==1):
                print("Try Again [%05d] %s - OK!" % (i,zkzh))
                ok+=1
            else:
                print("Try Again [%05d] %s - ERROR!" % (i,zkzh))
                error+=1
    i+=1

print("查询完毕，共计 %d 条结果"%(i-1))
print("其中 %d 条结果查询失败"%(error))
resworkbook.save('Result.xls')
os.system("pause")
