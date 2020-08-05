# -*- coding:utf-8 -*-
'''
Author: zhangpeiyu
Date: 2020-08-06 00:20:48
LastEditTime: 2020-08-06 01:16:08
Description: 我不是诗人，所以，只能够把爱你写进程序，当作不可解的密码，作为我一个人知道的秘密。
'''
from openpyxl import Workbook
from openpyxl.styles import Alignment
from datetime import datetime
import os

def create(excelPath, rows):
    alignment = Alignment(
        horizontal='center',  #水平对齐('centerContinuous', 'general', 'distributed','left', 'fill', 'center', 'justify', 'right')
        vertical='bottom',     #垂直对齐（'distributed', 'top', 'center', 'justify', 'bottom'）
        text_rotation=45,       #文字旋转
        wrap_text=False,       #自动换行
        shrink_to_fit=False,   #缩小字体填充
        mergeCell=None,        #合并单元格
        indent=0               #缩进
                      )

    # 创建工作簿
    wb = Workbook()

    # 获取当前sheet页
    sheet = wb.active

    # 填充数据
    for row in rows:
        sheet.append(row)

    
    sheet["A7"].value = u'张佩宇'
    tou = (7, 2, 99)
    sheet.cell(*tou)

    print(dir(sheet))
    # obj = {"A8": "宋荣妹", "B8": "98"}
    sheet["A8"].value = u"宋荣妹"
    sheet["B8"].value = 98
    
    # 设置单元格格式
    for col in sheet['A2:B5']:
        print(col)
        col[0].alignment = alignment

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
    create(os.path.join(os.path.dirname(__file__), formatTime + ".alignment.xlsx"), rows)