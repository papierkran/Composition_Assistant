# -*- coding: utf-8 -*-
import base64
import hashlib
import time
import requests
from urllib3 import response
import json

"""
  手写文字识别WebAPI接口调用示例接口文档(必看):https://doc.xfyun.cn/rest_api/%E6%89%8B%E5%86%99%E6%96%87%E5%AD%97%E8%AF%86%E5%88%AB.html
  图片属性：jpg/png/bmp,最短边至少15px，最长边最大4096px,编码后大小不超过4M,识别文字语种：中英文
  webapi OCR服务参考帖子(必看)：http://bbs.xfyun.cn/forum.php?mod=viewthread&tid=39111&highlight=OCR
  (Very Important)创建完webapi应用添加服务之后一定要设置ip白名单，找到控制台--我的应用--设置ip白名单，如何设置参考：http://bbs.xfyun.cn/forum.php?mod=viewthread&tid=41891
  错误码链接：https://www.xfyun.cn/document/error-code (code返回错误码时必看)
  @author iflytek
"""
# OCR手写文字识别接口地址
URL = "https://webapi.xfyun.cn/v1/service/v1/ocr/handwriting"
# URL = "https://cbm01.cn-huabei-1.xf-yun.com/v1/private/se75ocrbm"
# 应用APPID(必须为webapi类型应用,并开通手写文字识别服务,参考帖子如何创建一个webapi应用：http://bbs.xfyun.cn/forum.php?mod=viewthread&tid=36481)
APPID = "ef1639c2"
# 接口密钥(webapi类型应用开通手写文字识别后，控制台--我的应用---手写文字识别---相应服务的apikey)
API_KEY = "1bae243511eca8a05d80ddcba65cac19"
def getHeader():
    curTime = str(int(time.time()))
    param = "{\"language\":\""+language+"\",\"location\":\""+location+"\"}"
    paramBase64 = base64.b64encode(param.encode('utf-8'))

    m2 = hashlib.md5()
    str1 = API_KEY + curTime + str(paramBase64, 'utf-8')
    m2.update(str1.encode('utf-8'))
    checkSum = m2.hexdigest()
	# 组装http请求头
    header = {
        'X-CurTime': curTime,
        'X-Param': paramBase64,
        'X-Appid': APPID,
        'X-CheckSum': checkSum,
        'Content-Type': 'application/x-www-form-urlencoded; charset=utf-8',
    }
    return header
def getBody(filepath):
    with open(filepath, 'rb') as f:
        imgfile = f.read()
    data = {'image': str(base64.b64encode(imgfile), 'utf-8')}
    return data
# 语种设置
language = "cn|en"
# 是否返回文本位置信息
location = "true"
# 图片上传接口地址
picFilePath = "./img_edit/蔡梓妤/f9583df886d1e979bcd7d329509a74c.jpg"
# headers=getHeader(language, location)
r = requests.post(URL, headers=getHeader(), data=getBody(picFilePath))
# print(r)
# print(r.content)
response = r

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