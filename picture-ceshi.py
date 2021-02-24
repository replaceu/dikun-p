# -*- coding: utf-8 -*-
"""
Created on Tue Mar 31 10:03:13 2020

@author: Admin
"""

import xlsxwriter
import pandas as pd

workbook = xlsxwriter.Workbook('C://Users//Admin//Desktop//pict.xlsx') #创建一个excel文件
worksheet = workbook.add_worksheet('TEST') #在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
worksheet.set_column('A:D',25)
worksheet.set_row(80)

kind = pd.read_excel(r'C:\Users\Admin\Desktop\kind.xlsx')
kind = pd.DataFrame(kind['kind'])
kind = kind['kind'].tolist()
#kind=['4349','4379','4445','4433','4512','4349','4349','4680','4909','5226']
lth = range((len(kind))+1)
for i,j in zip(kind,lth):
    i =str(i)
    the_img = "//192.168.0.102/2012年生产资料/生产资料/共享/picture/"+i+".jpg"
    j=j+1
    cell ="D"+str(j)
    worksheet.insert_image(cell,the_img, {'x_scale': 0.05, 'y_scale': 0.05}) #插入一张图片
workbook.close()