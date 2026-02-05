# -*- coding: utf-8 -*-
import requests
import base64
import hashlib
import time
import uuid
import os

# ====== 请你在这里填上 ======
APP_KEY = "1d20af0ef4cab808"
APP_SECRET = "zOGj0nCldOIsaikU4JLXUywDinVoFMCm"
API_URL = "https://openapi.youdao.com/ocr_dewarp"
# ============================

def get_base64_image(image_path):
    with open(image_path, "rb") as f:
        img_bytes = f.read()
    return base64.b64encode(img_bytes).decode("utf-8")

def make_request(img_b64: str):
    curtime = str(int(time.time()))
    salt = str(uuid.uuid4())
    # q 参数就是 base64 编码的图片
    q = img_b64
    # sign 签名 = sha256(APP_KEY + q + salt + curtime + APP_SECRET)
    sign_str = APP_KEY + q + salt + curtime + APP_SECRET
    sign = hashlib.sha256(sign_str.encode('utf-8')).hexdigest()

    data = {
        "appKey": APP_KEY,
        "curtime": curtime,
        "salt": salt,
        "q": q,
        "sign": sign,
    }

    resp = requests.post(API_URL, data=data, timeout=60)
    return resp

def save_image_from_base64(b64_str, out_path):
    img_data = base64.b64decode(b64_str)
    with open(out_path, "wb") as f:
        f.write(img_data)

def dewarp_image_file(input_path, output_path):
    img_b64 = get_base64_image(input_path)
    resp = make_request(img_b64)
    if resp.status_code != 200:
        print("❌ 请求失败, status:", resp.status_code, resp.text)
        return False

    try:
        result = resp.json()
    except ValueError:
        print("❌ 响应无法解析为 JSON:", resp.text)
        return False

    if "dewarped_image" in result:
        # result["dewarped_image"] 是 base64 编码后的矫正图像
        save_image_from_base64(result["dewarped_image"], output_path)
        print("✅ 矫正后图片已保存到:", output_path)
        return True
    else:
        print("❌ 矫正失败, 返回内容:", result)
        return False

import os
import cv2  # 如果用 OpenCV 处理图片
# from your_module import your_image_processing_function  # ← 替成你的处理函数

# 支持的图片扩展名（你可以根据需要增删）
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}

def process_image_file(src_path, dst_path):
    """
    这是你的“图片处理 / 矫正 / 保存”逻辑入口。
    src_path: 原图片完整路径
    dst_path: 希望保存处理后图片的完整路径
    """
    img = cv2.imread(src_path)
    if img is None:
        print("⚠️ 无法读取图片，跳过:", src_path)
        return False

    # —— 在这里插入你的图片矫正 / 预处理 / OCR 前处理 代码 ——
    # 比如：
    # processed = your_image_processing_function(img)
    # 这里暂时直接用原图作为示例:
    processed = img

    # 确保目标目录存在
    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
    # 保存图片
    cv2.imwrite(dst_path, processed)
    print("✅ 保存处理后图片:", dst_path)
    return True

def batch_process_images(root_src, root_dst):
    """
    遍历 root_src 下所有子目录／子文件夹，
    对所有图片文件执行 process_image_file，
    并把结果保存到 root_dst，保持子目录结构不变。
    """
    for dirpath, dirnames, filenames in os.walk(root_src):
        # 计算相对路径（相对于 root_src）
        rel_dir = os.path.relpath(dirpath, root_src)
        for fname in filenames:
            ext = os.path.splitext(fname)[1].lower()
            if ext in IMAGE_EXTS:
                src_file = os.path.join(dirpath, fname)
                dst_file = os.path.join(root_dst, rel_dir, fname)
                process_image_file(src_file, dst_file)

if __name__ == "__main__":
    # 修改下面路径为你自己的
    INPUT_ROOT = "../img"        # 存放原始图片（可能有多层子文件夹）的根目录
    OUTPUT_ROOT = "./img_edit"  # 处理后图片输出根目录
    batch_process_images(INPUT_ROOT, OUTPUT_ROOT)


