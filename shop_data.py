# -*- coding: utf-8 -*-
"""
Created on Fri Jun  5 10:20:35 2020

@author: Admin
"""

from datetime import datetime
import pandas as pd
import numpy as np
import xlwt
import xlrd
import xlutils
import matplotlib 
from xlutils.copy import copy

matplotlib.rcParams['font.sans-serif'] = ['SimHei']
matplotlib.rcParams['font.family'] = ['sans-serif']
matplotlib.rcParams['axes.unicode_minus']=False


def customer_data():
    customer =pd.read_excel(r'C:\Users\Admin\Desktop\shop_customer.xlsx')
    customer = pd.DataFrame(customer[['会员卡号','会员名称',
                         '性别','联系电话',
                         '有效期','开卡时间',
                         '大概年龄段','个人描述（样貌）',
                         '编号','微信号',
                         '购买性质']])
    customer = customer.fillna("")

    
    for i in customer['会员卡号'].values:      
        message = customer.loc[customer['会员卡号']==i]
        i = str(i)
        f = open('D:\\我的文件\\Data\\shop_customer\\' + i+'.csv','w',newline='')
        message.to_csv(f,index=False)
        
      #  print(message)  
        
def customer_split():
    
    list_file = pd.read_excel(r'C:\Users\Admin\Desktop\shop_customer.xlsx')
    list_file = pd.DataFrame(list_file)
    list_file = list_file['会员卡号']
    for i in list_file:
        i =str(i)
        each_customer =pd.read_csv(r'D:\\我的文件\\Data\\shop_customer\\' + i+'.csv',encoding='gbk')
        each_customer = pd.DataFrame(each_customer)
        
        card = str(each_customer['会员卡号'].values)
        name = str(each_customer['会员名称'].values)
        sex = str(each_customer['性别'].values)
        start_date = str(each_customer['开卡时间'].values)
        phone = str(each_customer['联系电话'].values)
        deadline = str(each_customer['有效期'].values)
        describe = str(each_customer['个人描述（样貌）'].values)
        Id = str(each_customer['编号'].values)
        age = str(each_customer['大概年龄段'].values)
        wechat = str(each_customer['微信号'].values)
        workbook = xlwt.Workbook('D:\\我的文件\\Data\\shop_customer1\\' + i+'.csv')  #创建一个Excel文件
        
        Style = xlwt.XFStyle()#格式信息
        font = xlwt.Font()#字体基本设置
        font.name = u'微软雅黑'
        font.color_index = 5
        font.height= 220 #字体大小，220就是11号字体，大概就是11*20得来的吧
        Style.font = font
        worksheet = workbook.add_sheet(sheetname=i) 
        
        alignment = xlwt.Alignment() # 设置字体在单元格的位置
        alignment.horz = xlwt.Alignment.HORZ_CENTER #水平方向
        alignment.vert = xlwt.Alignment.VERT_CENTER #竖直方向
        alignment.wrap = 1
        Style.alignment = alignment
        
        border = xlwt.Borders()  #给单元格加框线
        border.left = xlwt.Borders.THIN  #左
        border.top=xlwt.Borders.THIN     #上
        border.right=xlwt.Borders.THIN   #右
        border.bottom=xlwt.Borders.THIN  #下
        border.left_colour = 0x40  #设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
        border.right_colour = 0x40
        border.top_colour = 0x40
        border.bottom_colour = 0x40
        Style.borders = border
        
        tall_style = xlwt.easyxf('font:height 6000')  # 36pt
        first_row = worksheet.row(4)
        first_row.set_style(tall_style)

        worksheet.write(0,0,'会员卡号',Style)
        worksheet.write(0,1,card,Style) 
        
        worksheet.col(0).width = 10000
        worksheet.col(1).width = 10000 # 3333 = 1" (one inch).
        
        worksheet.write(0,2,'会员名称',Style) 
        worksheet.write(0,3,name,Style)
        
        worksheet.write(0,4,'会员编号',Style) 
        worksheet.write(0,5,Id,Style) 
        worksheet.write(1,0,'发卡日期',Style)
        worksheet.write(1,1,start_date,Style)
        worksheet.write(1,2,'性别',Style)
        worksheet.write(1,3,sex,Style)
        worksheet.write(1,4,'年龄段',Style)
        worksheet.write(1,5,age,Style)
        
        worksheet.write(2,0,'结束日期',Style)
        worksheet.write(2,1,deadline,Style)
        worksheet.write(2,2,'电话',Style)
        worksheet.write(2,3,phone,Style)
        worksheet.write(2,4,'微信',Style)
        worksheet.write(2,5,wechat,Style)
        
        worksheet.write_merge(5,5,0,5,'回访内容',Style)

        worksheet.write_merge(4,4,0,5,describe,Style)
        worksheet.write_merge(3,3,0,5,'客户印象',Style)
        
        
        worksheet.col(2).width = 3333
        worksheet.col(3).width = 3333
        worksheet.col(4).width = 3333
        worksheet.col(5).width= 3333
        
        
        workbook.save('D:\\我的文件\\Data\\shop_customer1\\' + i+'.csv')
         

def write_excel_xls_append():
    
    list_file = pd.read_excel(r'C:\Users\Admin\Desktop\shop_customer.xlsx')
    list_file = pd.DataFrame(list_file)
    list_file = list_file['会员卡号']
    
    purchase_history = pd.read_excel(r'C:\Users\Admin\Desktop\shop_sale.xls')
  #  purchase_history = pd.DataFrame(purchase_history)
    purchase_history = pd.DataFrame(purchase_history[['VIP卡号','销售时间','商品名称','颜色','尺码','数量']])
  #  value_list =purchase_history['VIP卡号'].values
   # print(value_list)
    
    for i in list_file:
       # i =str(i)
        value = purchase_history.loc[purchase_history['VIP卡号']==i]
        value = np.array(value)
        print(value)
        
       # print(value,i)
        index = len(value)  # 获取需要写入数据的行数
        workbook = xlrd.open_workbook('D:\\我的文件\\Data\\shop_customer1\\' + str(i)+'.csv',formatting_info=True)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
        for a in range(0, index):
            for j in range(0, len(value[a])):
                Style = xlwt.XFStyle()#格式信息
                font = xlwt.Font()#字体基本设置
                font.name = u'微软雅黑'
                font.color_index = 5
                font.height= 160 #字体大小，220就是11号字体，大概就是11*20得来的吧
                Style.font = font
               # worksheet = workbook.add_sheet(sheetname=i) 
        
                alignment = xlwt.Alignment() # 设置字体在单元格的位置
                alignment.horz = xlwt.Alignment.HORZ_CENTER #水平方向
                alignment.vert = xlwt.Alignment.VERT_CENTER #竖直方向
                alignment.wrap = 1
                Style.alignment = alignment
        
                border = xlwt.Borders()  #给单元格加框线
                border.left = xlwt.Borders.THIN  #左
                border.top=xlwt.Borders.THIN     #上
                border.right=xlwt.Borders.THIN   #右
                border.bottom=xlwt.Borders.THIN  #下
                border.left_colour = 0x40  #设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
                border.right_colour = 0x40
                border.top_colour = 0x40
                border.bottom_colour = 0x40
                Style.borders = border
                new_worksheet.write(a+rows_old, j, value[a][j],Style)
               # new_worksheet.col(6).width = 10000
                tall_style = xlwt.easyxf('font:height 6000')  # 36pt
                first_row = new_worksheet.row(4)
                first_row.set_style(tall_style)
                # 追加写入数据，注意是从i+rows_old行开始写入
                new_workbook.save('D:\\我的文件\\Data\\shop_customer1\\' + str(i)+'.csv')  # 保存工作簿
        #print("xls格式表格【追加】写入数据成功！")

def return_excel_xls_append():
    list_file = pd.read_excel(r'C:\Users\Admin\Desktop\shop_customer.xlsx')
    list_file = pd.DataFrame(list_file)
    list_file = list_file['会员卡号']
    
    record_data = pd.read_excel("D://我的文件//Data//return_record.xlsx")
    record_data = pd.DataFrame(record_data[['会员卡号','日期','回访内容']])
    
    for i in list_file:
       # i =str(i)
        value = record_data.loc[record_data['会员卡号']==i]
        value = np.array(value)
       # print(value,i)
        index = len(value)  # 获取需要写入数据的行数
        workbook = xlrd.open_workbook('D:\\我的文件\\Data\\shop_customer1\\' + str(i)+'.csv',formatting_info=True)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
        for a in range(0, index):
            for j in range(0, len(value[a])):
                Style = xlwt.XFStyle()#格式信息
                font = xlwt.Font()#字体基本设置
                font.name = u'微软雅黑'
                font.color_index = 5
                font.height= 160 #字体大小，220就是11号字体，大概就是11*20得来的吧
                Style.font = font
               # worksheet = workbook.add_sheet(sheetname=i) 
        
                alignment = xlwt.Alignment() # 设置字体在单元格的位置
                alignment.horz = xlwt.Alignment.HORZ_CENTER #水平方向
                alignment.vert = xlwt.Alignment.VERT_CENTER #竖直方向
                alignment.wrap = 1
                Style.alignment = alignment
        
                border = xlwt.Borders()  #给单元格加框线
                border.left = xlwt.Borders.THIN  #左
                border.top=xlwt.Borders.THIN     #上
                border.right=xlwt.Borders.THIN   #右
                border.bottom=xlwt.Borders.THIN  #下
                border.left_colour = 0x40  #设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
                border.right_colour = 0x40
                border.top_colour = 0x40
                border.bottom_colour = 0x40
                Style.borders = border
               # new_worksheet.write_merge(a+rows_old,a+rows_old,2,5, value[a][j-1],Style)
                new_worksheet.write_merge(a+rows_old,a+rows_old,0,5, value[a][j],Style)
               # new_worksheet.col(6).width = 10000
                tall_style = xlwt.easyxf('font:height 6000')  # 36pt
                first_row = new_worksheet.row(4)
                first_row.set_style(tall_style)
                # 追加写入数据，注意是从i+rows_old行开始写入
                new_workbook.save('D:\\我的文件\\Data\\shop_customer1\\' + str(i)+'.csv')  # 保存工作簿
    
    
customer_data()
customer_split()
return_excel_xls_append()
write_excel_xls_append()