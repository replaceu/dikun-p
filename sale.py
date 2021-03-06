# -*- coding: utf-8 -*-
"""
Created on Fri May  8 15:55:12 2020

@author: Admin
"""

import pymysql
import pandas as pd
import xlrd
#先将数据写进到MySQL当中
# 一个根据pandas自动识别type来设定table的type
try:
    db = pymysql.connect(host="127.0.0.1",
                         user="******",
                         passwd="******",
                         db="test",
                         charset='utf8')
except:
    print("could not connect to mysql server")
    
def open_excel():
    try:
        book = xlrd.open_workbook("D:\\我的文件\\Data\\sale\\sale_data.xls")#文件名，把文件与py文件放在同一目录下
    except:
        print("open excel file failed!")
    try:
        sheet = book.sheet_by_name("Sheet1")   #execl里面的worksheet1
        return sheet
    except:
        print("locate worksheet in excel failed!")

def insert_deta():
    sheet = open_excel()

    cursor = db.cursor()
    cursor.execute("truncate TABLE  sale_biao")
    row_num = sheet.nrows
    for i in range(1, row_num):  # 第一行是标题名，对应表中的字段名所以应该从第二行开始，计算机以0开始计数，所以值是1
        row_data = sheet.row_values(i)
        value = (row_data[0],row_data[1],row_data[2],
                 row_data[3],row_data[4],row_data[5])
 
        sql ="""
        INSERT INTO sale_biao
        (goods_name,
        customer,
        saler,
        date,
        number,
        price)
        VALUES
        (%s,%s,%s,%s,%s,%s)
        """
        cursor.execute(sql,value)  # 执行sql语句
        db.commit()
    cursor.close()

insert_deta()
