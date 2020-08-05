# -*- coding:utf-8 -*-
'''
Author: zhangpeiyu
Date: 2020-08-05 23:52:02
LastEditTime: 2020-08-06 00:30:12
Description: 我不是诗人，所以，只能够把爱你写进程序，当作不可解的密码，作为我一个人知道的秘密。
'''

from openpyxl.formatting.rule import DataBarRule
from openpyxl import Workbook
from datetime import datetime
import os

'''
创建带有数据条件的Excel文件
'''
def create(excelPath, rows):
    # 定义数据条件规则
    rule = DataBarRule(start_type='percentile', start_value=10, end_type='percentile', 
    end_value='90',color="FF638EC6", showValue="None", minLength=None, maxLength=None)

    # 创建工作簿
    wb = Workbook()

    # 获取当前sheet页
    sheet = wb.active

    # 填充数据
    for row in rows:
        sheet.append(row)

    # 设置数据条件规则
    sheet.conditional_formatting.add('B2:B5', rule)

    wb.save(excelPath)
    
    print("生成的Excel路径是：", excelPath)

if __name__ == "__main__":
    rows=[[u'姓名', u'成绩']]
    rows.append([u'张三', 95])
    rows.append([u'李四', 88])
    rows.append([u'王二麻', 67])
    rows.append([u'张三丰', 77])
    rows.append([u'荣妹', 100])

    # 获得当前时间
    now = datetime.now()  # ->这是时间数组格式
    #转换为指定的格式:
    formatTime = now.strftime("%Y-%m-%d")
    create(os.path.join(os.path.dirname(__file__), formatTime + ".xlsx"), rows)


