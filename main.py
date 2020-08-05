# -*- coding:utf-8 -*-
'''
Author: zhangpeiyu
Date: 2020-08-04 23:16:30
LastEditTime: 2020-08-06 00:01:02
Description: 我不是诗人，所以，只能够把爱你写进程序，当作不可解的密码，作为我一个人知道的秘密。
'''
from create_excel import createExcel as create


if __name__ == "__main__":
    # 数据头部
    header = [u'标题1',u'标题2',u'标题三', u'标题4']
    data = []
    for item in range(1,20):
        temp = [item, item+1, item+1, item+1]
        data.append(temp)
    
    create(header, data, u'测试')