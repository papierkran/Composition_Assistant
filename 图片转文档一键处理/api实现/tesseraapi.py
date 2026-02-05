import os
from PIL import Image, ImageOps
import pytesseract
from docx import Document

# 如果你的 Tesseract 安装路径未在 PATH 中，在这里设置：
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def ocr_image_to_text(image_path):
    """
    将图片 OCR 识别为文本，并按换行拆分为段落。
    """
    # 打开图片，并做简单预处理（转灰度 + 二值化）
    img = Image.open(image_path)
    img = ImageOps.grayscale(img)

    # OCR 识别，返回整段文字
    text = pytesseract.image_to_string(img, lang='chi_sim+eng')

    # 按换行拆分为“段落”
    paragraphs = [line.strip() for line in text.splitlines() if line.strip()]
    return paragraphs

def save_paragraphs_to_docx(paragraphs, output_file):
    """
    将识别出的段落写入 Word (.docxx) 文件
    """
    doc = Document()
    for para in paragraphs:
        doc.add_paragraph(para)
    doc.save(output_file)

if __name__ == "__main__":
    img_file = "./img_edit/一灏/1.jpg"  # 输入图片路径
    out_doc = "./docxx/ocr_result.docxx"  # 输出 Word

    paras = ocr_image_to_text(img_file)
    save_paragraphs_to_docx(paras, out_doc)

    print(f"识别完成，已保存到：{out_doc}")
