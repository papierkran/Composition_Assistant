# coding=utf-8
"""百度智能云文档矫正增强 API 客户端。

使用百度 OCR 文档矫正增强接口，对图片进行透视变换和增强处理。
API 文档: https://cloud.baidu.com/doc/OCR/s/Hl4taza5f
"""

from __future__ import annotations

import base64
import logging
from pathlib import Path
from typing import Optional

import requests

_LOGGER = logging.getLogger(__name__)


class BaiduImageCorrector:
    """百度文档矫正增强 API 客户端。"""

    TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"
    CORRECT_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/doc_crop_enhance"

    def __init__(self, api_key: str, secret_key: str):
        """
        初始化客户端。

        Args:
            api_key: 百度智能云 API Key
            secret_key: 百度智能云 Secret Key
        """
        self.api_key = api_key.strip()
        self.secret_key = secret_key.strip()
        self._access_token: Optional[str] = None

    def _get_access_token(self) -> str:
        """获取 access_token（带缓存）。"""
        if self._access_token:
            return self._access_token

        params = {
            "grant_type": "client_credentials",
            "client_id": self.api_key,
            "client_secret": self.secret_key,
        }

        resp = requests.post(self.TOKEN_URL, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        if "access_token" not in data:
            raise RuntimeError(f"获取百度 access_token 失败: {data}")

        self._access_token = data["access_token"]
        _LOGGER.info("百度 access_token 获取成功")
        return self._access_token

    def correct_image(
        self,
        image_path: str,
        enhance_type: int = 1,
    ) -> bytes:
        """
        调用百度文档矫正增强 API 矫正图片。

        Args:
            image_path: 原始图片路径
            enhance_type: 增强类型
                - 0: 不增强
                - 1: 去阴影
                - 2: 增强并锐化
                - 3: 黑白滤镜

        Returns:
            矫正后的图片 bytes（JPEG 格式）
        """
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"图片文件不存在: {image_path}")

        # 读取图片并 base64 编码
        with open(image_path, "rb") as f:
            image_data = f.read()

        image_base64 = base64.b64encode(image_data).decode("utf-8")

        # 获取 access_token
        access_token = self._get_access_token()

        # 构建请求
        url = f"{self.CORRECT_URL}?access_token={access_token}"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "image": image_base64,
            "scan_type": 3,  # 检测并矫正
            "enhance_type": enhance_type,
        }

        # 发送请求
        resp = requests.post(url, data=data, headers=headers, timeout=30)
        resp.raise_for_status()
        result = resp.json()

        # 检查返回结果
        if "image_processed" not in result:
            error_msg = result.get("error_msg", "未知错误")
            error_code = result.get("error_code", "未知")
            raise RuntimeError(
                f"百度文档矫正失败: [{error_code}] {error_msg}"
            )

        # 解码矫正后的图片
        corrected_image = base64.b64decode(result["image_processed"])

        _LOGGER.info(f"图片矫正成功: {image_path.name}")
        return corrected_image


def correct_image_file(
    image_path: str,
    output_path: str,
    api_key: str,
    secret_key: str,
    enhance_type: int = 1,
) -> bool:
    """
    矫正单张图片并保存。

    Args:
        image_path: 输入图片路径
        output_path: 输出图片路径
        api_key: 百度 API Key
        secret_key: 百度 Secret Key
        enhance_type: 增强类型

    Returns:
        是否成功
    """
    try:
        corrector = BaiduImageCorrector(api_key, secret_key)
        corrected_data = corrector.correct_image(image_path, enhance_type)

        # 保存矫正后的图片
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        with open(output, "wb") as f:
            f.write(corrected_data)

        return True
    except Exception as e:
        _LOGGER.error(f"图片矫正失败: {image_path}, 错误: {e}")
        return False
