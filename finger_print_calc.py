
# -*- coding: utf-8 -*-
"""
Created on Sat Feb  6 00:33:47 2016

@author: YuKeji
"""

import re
import datetime
#更改读入txt文件的路径
fd = open('/Users/YuKeji/Desktop/kaoqin/201601-1.txt','r', encoding='gbk')
fd_data = fd.readlines()
all_list = []
#对从txt读出的每行进行操作
for i in fd_data:
    i_split = re.split('\n| ',i) #正则表达式按照空格和回车来拆分字符串
    i_filter = filter(None, i_split) #去除列表中所有空字符
    r = list(i_filter) #转换成列表
    
    if r == []:                           #忽略为空的list
        pass
    else:
        if r[1] == '总时长':    
            all_list.append(r)            #将表头附在all_list上
        else:    
            a = r[1].replace(':','.')     #调整工时的格式
            r[1] = float(a)
            r[2] = datetime.datetime.strptime(r[2],'%Y/%m/%d') #将日期从string变为date格式
            all_list.append(r)            #将工时list附在all_list上
fd.close()

import xlrd
import xlwt
#新建一个excel,用来做数据透视
all_list_excel = xlwt.Workbook()
#在新建的excel中加入一个sheet，命名为all_list
sheet_all_list = all_list_excel.add_sheet('all_list')
for Index, Item in enumerate(all_list):
    for index, item in enumerate(Item):
        sheet_all_list.write(Index, index, item)
all_list_excel.save('/Users/YuKeji/Desktop/kaoqin/all-list-excel.xls')
import pandas as pd
import numpy as np
df = pd.read_excel('/Users/YuKeji/Desktop/kaoqin/all-list-excel.xls')
df_pivot = pd.pivot_table(df,index=['姓名'],values=['总时长'], columns=['考勤日期'],aggfunc=[np.sum])
#from pandas import DataFrame
path_pivot = '/Users/YuKeji/Desktop/kaoqin/pivot.xlsx'
sheet_pivot = 'sheet_pivot'
writer = pd.ExcelWriter(path_pivot)
df_pivot.to_excel(writer,sheet_pivot)
writer.save()


number_of_days = int(input('how many days set this month:')) + 1
number_of_people = int(input('how many people:'))

#左对齐style
alignment = xlwt.Alignment() # Create Alignment
alignment.horz = xlwt.Alignment.HORZ_LEFT # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
#alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
style = xlwt.XFStyle() # Create Style
style.alignment = alignment # Add Alignment to Style

#左对齐+上横线+日期style_1
borders = xlwt.Borders()
borders.top = 1
style_1 = xlwt.XFStyle()
style_1.borders = borders
style_1.alignment = alignment
style_1.num_format_str = 'M/D/YY'

#左对齐+日期 style_2
style_2 = xlwt.XFStyle()
style_2.alignment = alignment
style_2.num_format_str = 'M/D/YY'

#左对齐+上横线 style_3
style_3 = xlwt.XFStyle()
style_3.alignment = alignment
style_3.borders = borders

#新建一个excel
new_excel = xlwt.Workbook()
#在新建的excel中加入一个sheet，命名为kaoqin
sheet_kaoqin = new_excel.add_sheet('kaoqin')

#打开已经经过透视表整理的excel，赋给变量book1
book1 = xlrd.open_workbook(path_pivot)
#找到book1中的Sheet2，这里有考勤的原始数据
sheet1 = book1.sheet_by_name(sheet_pivot)
#初始三个变量，用于写入cell的迭代
row_wt_init = 1 
row_wt_init_name = 1
row_rd_init = 3
#从book1的sheet2中读第一列的名字
names = sheet1.col_values(0, 4, number_of_people+4)
#从第4行读出日期
row_date = sheet1.row_values(row_rd_init-1, 1, number_of_days)
#因为最终要变换为三行，故先将日期列表按照11，11，其余 来切片
row_date_1 = row_date[:11]
row_date_2 = row_date[11:22]
row_date_3 = row_date[22:]
#向新文件kaoqin Sheet第一列写入名字，range中的37可根据总人数来调整
for i in range(0,number_of_people):
    sheet_kaoqin.write(row_wt_init_name, 0, names[i])
    row_wt_init_name += 8
#为第一列（人名）调整列宽，人名最多三个字，每个字算2个字符，256*字符数 为宽度
sheet_kaoqin.col(0).width = 256 * 6
#写入每天的时长，大循环开始；先在原始数据中读出每人每天的时长，依照date的规则切片  
for i in range(row_rd_init+1, number_of_people+4):
    time_everyday = sheet1.row_values(i, 1, number_of_days)
    #统计非空字符，也就是实际到的天数
    days_of_come = number_of_days - time_everyday.count('') - 1
    time_everyday_1 = time_everyday[:11]
    time_everyday_2 = time_everyday[11:22]
    time_everyday_3 = time_everyday[22:]
    
    #嵌套小循环，从第二行开始连写三行（隔行）日期
    for i in range(0,len(row_date_1)):
        sheet_kaoqin.write(row_wt_init, i+1, row_date_1[i], style_1)
        
    for i in range(0,len(row_date_2)):
        sheet_kaoqin.write(row_wt_init+2, i+1, row_date_2[i], style_1)
    for i in range(0,len(row_date_3)):
        sheet_kaoqin.write(row_wt_init+4, i+1, row_date_3[i], style_1)
    #嵌套小循环，从第三行开始连写三行（隔行）每天时长
    for i in range(0,len(row_date_1)):
        sheet_kaoqin.write(row_wt_init+1, i+1, time_everyday_1[i], style)
    for i in range(0,len(row_date_2)):
        sheet_kaoqin.write(row_wt_init+3, i+1, time_everyday_2[i], style)
    for i in range(0,len(row_date_3)):
        sheet_kaoqin.write(row_wt_init+5, i+1, time_everyday_3[i], style)
    #至此row_wt_init的值一直没有变，只有当下面语句生效，才开始从新的cell上开始继续写
    total = 0
    a = number_of_days - 1 
    for i in range(0, len(time_everyday)):
        if time_everyday[i] == '':
            a -= 1
        else :
            add_temp = float(time_everyday[i])
            total += add_temp
    sheet_kaoqin.write(row_wt_init+6, 1, '总时长',style_3)
    sheet_kaoqin.write(row_wt_init+6, 2, total, style_3)
    sheet_kaoqin.write(row_wt_init+6, 3, '实到天数', style_3)
    sheet_kaoqin.write(row_wt_init+6, 4, a,  style_3)
    sheet_kaoqin.write(row_wt_init+6, 5, '日均时长', style_3)
    sheet_kaoqin.write(row_wt_init+6, 6, round(total/days_of_come, 2), style_3)#保留2位小数
        #若不是从第一天开始，就改实际天数为a
    row_wt_init += 8
#更改输出Excel文件的路径
new_excel.save('/Users/YuKeji/Desktop/kaoqin/指纹统计201601.xls')



