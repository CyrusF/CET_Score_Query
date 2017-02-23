# CET_Score_Query

## Version 1.0

**手动将excel表格数据提取到input file，输出文件为“-”隔开的txt格式</br>**
用来批量查询CET成绩的python脚本，通常不会触发验证码。</br>
成绩发布后随着查询人数增多触发验证码几率会变大。</br>
如果已经触发验证码显示ERROR了，建议调小pause间隔换个ip再试一次。</br>
祝取得满意成绩</br>
使用Python3+运行，Windows cmd/PS如果乱码用`chcp 437`更改一下代码页再运行。

## Version 2.0

**自动读取xlsx格式的input file，输出文件为xls格式</br>**
更换了一个查询网站，几乎完全没有验证码问题了。</br>
只是这网站用GBK编码也是非常奇妙。</br>
运行环境要求和version 1.0相同，要求xlrd和xlwt库的支持。
