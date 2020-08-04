'''
Author: zhangpeiyu
Date: 2020-08-05 00:13:27
LastEditTime: 2020-08-05 00:34:02
Description: 我不是诗人，所以，只能够把爱你写进程序，当作不可解的密码，作为我一个人知道的秘密。
'''
import json
import os

def getJsonOnFile(filePath):
    with open(filePath, "r") as fw:
        load_dict = json.load(fw)
        print(load_dict)
        return load_dict

if __name__ == "__main__":

    jsonPath = os.path.join(os.path.dirname(__file__),"data.json")
    json = getJsonOnFile(jsonPath)
    print(dir(json))
    print(json.keys())
    for item in json:
        print(item)