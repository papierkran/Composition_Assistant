
# -*- coding: utf-8 -*-
import sys
import uuid
import requests
import base64
import hashlib
import json
from imp import reload


import time

reload(sys)

APP_KEY = '1d20af0ef4cab808'
APP_SECRET = 'zOGj0nCldOIsaikU4JLXUywDinVoFMCm'
# YOUDAO_URL = 'https://openapi.youdao.com/ocrapi'

YOUDAO_URL = 'https://openapi.youdao.com/ocr_hand_writing'


def truncate(q):
    if q is None:
        return None
    size = len(q)
    return q if size <= 20 else q[0:10] + str(size) + q[size - 10:size]


def encrypt(signStr):
    hash_algorithm = hashlib.sha256()
    hash_algorithm.update(signStr.encode('utf-8'))
    return hash_algorithm.hexdigest()


def do_request(data):
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    return requests.post(YOUDAO_URL, data=data, headers=headers)


def connect():
    f = open(r'./img_edit/蔡梓妤/f9583df886d1e979bcd7d329509a74c.jpg', 'rb')  # 二进制方式打开图文件
    q = base64.b64encode(f.read()).decode('utf-8')  # 读取文件内容，转换为base64编码
    f.close()

    data = {}
    # data['detectType'] = '识别类型'
    data['imageType'] = '1'
    data['langType'] = 'zh-CHS'
    data['img'] = q
    data['docType'] = 'json'
    data['signType'] = 'v3'
    curtime = str(int(time.time()))
    data['curtime'] = curtime
    salt = str(uuid.uuid1())
    signStr = APP_KEY + truncate(q) + salt + curtime + APP_SECRET
    sign = encrypt(signStr)
    data['appKey'] = APP_KEY
    data['salt'] = salt
    data['sign'] = sign

    response = do_request(data)
    # print(response.content)
    # 将二进制响应内容转换为字符串（JSON）
    try:
        # 方法1：使用 response.json() 直接解析
        result = response.json()
        print("API响应结果（JSON格式）:")
        print(json.dumps(result, ensure_ascii=False, indent=2))

        # 提取识别结果
        if result.get('errorCode') == '0' and 'Result' in result:
            print("\n识别到的文字:")
            regions = result['Result']['regions']
            for region in regions:
                lines = region.get('lines', [])
                for line in lines:
                    text = line.get('text', '')
                    print(text)
        else:
            print(f"识别失败，错误码: {result.get('errorCode')}")
            print(f"错误信息: {result.get('errorMsg', '未知错误')}")

    except json.JSONDecodeError:
        # 如果 response.json() 失败，尝试手动解码
        print("尝试手动解码响应内容...")
        text_response = response.content.decode('utf-8')
        print("原始响应文本:")
        print(text_response)

        # 尝试解析为JSON
        try:
            result = json.loads(text_response)
            print("\n解析后的JSON:")
            print(json.dumps(result, ensure_ascii=False, indent=2))
        except:
            print("无法解析为JSON格式")



if __name__ == '__main__':
    connect()