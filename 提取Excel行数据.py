
import os
import xlrd2 as xlrd
from xlutils.copy import copy
def base_dir(filename=None):
    return os.path.join(os.path.dirname(__file__),filename)
work = xlrd.open_workbook(base_dir("sj01.xls"))
old_content = copy(work)
ws = old_content.get_sheet(0)

#base_dir("sj01.xls")  第一个需要修改的地方，原始文件的地址sj01.xls、sj02.xls、、、，要和程序放在同一个文件夹下

'''
一共15个表
表1-14:07-20年数据，表15

表0 2007--行3
表1 2008--行4
表2 2009--行5
表3 2010--行6
表4 2011--行7
表5 2012--行8
表6 2013--行9
表7 2014--行10
表8 2015--行11
表9 2016--行12
表10 2017--行13
表11 2018--行14
表12 2019--行15
表13 2020--行16
'''

'''
2017  表10 2017--行13
表10长安区数据 行2，列5-11
汇总至表15 行16，列5-11
'''
# 索引到第X个工作表
sheetsummarizing=work.sheet_by_index(14)
sheet2007 = work.sheet_by_index(0)

#第二个需要修改的地方：收集哪一年就新建一个变量sheet20XX
#sheet20XX = work.sheet_by_index(编号在上面注释里面一一对应)
#表0 2007：sheet2007 = work.sheet_by_index(0)
#表1 2008：sheet2008 = work.sheet_by_index(1)

xieruhang=3
#第三个需要修改的地方，找汇总表里20XX年长安区的数据在第几行（Excel表格的编号），2007年是第三行，所以xieruhang=3

for i in range(1,169):
    for j in range(4,11):
        a = sheet2007.cell_value(i, j)
        #第四个需要修改的地方，在第二个需要修改的地方，换成新定义的sheet20XX：a = sheet2008.cell_value(i, j)
        print(a)
        ws.write(xieruhang, j, a)
    xieruhang=xieruhang+14
old_content.save(base_dir("sj02.xls"))
#第五个需要修改的地方，修改过的数据存进一个新的文件里面。sj01.xls————sj02.xls，新的文件作为下一次实验的原始文件，下一个循环sj02.xls————sj03.xls
#每改一次就可以把一整年的数据全放进新表里面，循环完所有的年份就行了了



