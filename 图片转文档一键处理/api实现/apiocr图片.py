# -*- coding: utf-8 -*-
import os
import uuid
import requests
import base64
import hashlib
import time

# ==== 配置你的有道 key/secret ====
APP_KEY = '1d20af0ef4cab808'
APP_SECRET = 'zOGj0nCldOIsaikU4JLXUywDinVoFMCm'
YOUDAO_URL = 'https://openapi.youdao.com/ocr_hand_writing'

# 支持的图片扩展名
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}

def truncate(q: str) -> str:
    size = len(q)
    if size <= 20:
        return q
    return q[:10] + str(size) + q[-10:]

def encrypt(sign_str: str) -> str:
    return hashlib.sha256(sign_str.encode('utf-8')).hexdigest()

def image_to_base64(image_path: str) -> str | None:
    try:
        with open(image_path, 'rb') as f:
            data = f.read()
            return base64.b64encode(data).decode('utf-8')
    except Exception as e:
        print("⚠️ 无法读取图片为 base64:", image_path, "error:", e)
        return None

def ocr_image_b64(img_b64: str) -> dict | None:
    salt = str(uuid.uuid4())
    curtime = str(int(time.time()))
    sign = encrypt(APP_KEY + truncate(img_b64) + salt + curtime + APP_SECRET)

    data = {
        'appKey': APP_KEY,
        'salt': salt,
        'curtime': curtime,
        'sign': sign,
        'signType': 'v3',
        'imageType': '1',
        'img': img_b64,
        'langType': 'zh-CHS',  # 中文
        'detectType': '10012',
        'docType': 'json',
    }

    try:
        resp = requests.post(YOUDAO_URL, data=data, timeout=60)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print("❌ OCR 请求失败:", e)
        return None

def process_image_file(image_path: str):
    img_b64 = image_to_base64(image_path)
    if img_b64 is None:
        return
    result = ocr_image_b64(img_b64)

    if result is not None:
        print("✅ OCR 结果 —", image_path)
        print(result)  # 或者改成保存到文件
    else:
        print("❌ OCR 失败 —", image_path)

def batch_ocr(root_dir: str):
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for fname in filenames:
            ext = os.path.splitext(fname)[1].lower()
            if ext in IMAGE_EXTS:
                full = os.path.join(dirpath, fname)
                process_image_file(full)

if __name__ == "__main__":
    ROOT = "./img"  # 你要遍历的图片根目录
    batch_ocr(ROOT)
    print("==== 处理结束 ====")
