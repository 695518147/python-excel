# -*- coding:utf-8 -*-
'''
Author: zhangpeiyu
Date: 2020-08-04 23:56:08
LastEditTime: 2020-08-05 00:10:27
Description: 我不是诗人，所以，只能够把爱你写进程序，当作不可解的密码，作为我一个人知道的秘密。
'''
from openpyxl import Workbook
import os


'''
根据数据生成excel
'''
def createExcel(header=[], data=[], wb_name = 'wb-name', sheet_name = 'sheet-name'):
    # 创建Excel工作簿
    wb = Workbook()

    # 激活当前sheet页
    sheet = wb.active

    # 设置sheet页名字
    sheet.title = sheet_name

    # 添加头部数据
    sheet.append(header)

    # 填充具体的数据
    for item in data:
        sheet.append(item)

    filePath = os.path.join(os.path.dirname(__file__), wb_name + ".xlsx")
    wb.save(filePath)
    print(u'最大行数：', sheet.max_row)
    print(u'最大列数：', sheet.max_column)
    print(u'生成的Excel路径是：', filePath)
