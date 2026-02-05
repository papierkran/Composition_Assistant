# coding=utf-8
# ocr_main.py
import os
import time
import hashlib
import base64
from doctest import debug

import requests
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Cm, Pt
from docx.enum.text import WD_LINE_SPACING
from openai import OpenAI

import json

CONFIG_FILE = "D:\person_data\ocer助手\presson.json"
# CONFIG_FILE = "config.json"

def load_config():
    if not os.path.exists(CONFIG_FILE):
        raise RuntimeError("❌ 未找到 config.json 或相关配置文件，请先配置")

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

config = load_config()



# ========== 讯飞ocr配置区 =================================================
OCR_CONFIG = config.get("OCR", {})

URL = OCR_CONFIG.get("URL")
APPID = OCR_CONFIG.get("APPID")
API_KEY = OCR_CONFIG.get("API_KEY")
language = OCR_CONFIG.get("LANGUAGE", "cn|en")
location = OCR_CONFIG.get("LOCATION", "false")

if not all([URL, APPID, API_KEY]):
    raise RuntimeError("❌ OCR 配置不完整，请检查 config.json")

# ==================================================================


# ========== deepseek_api  配置区 =================================================
DEEPSEEK_CONFIG = config.get("DEEPSEEK", {})

DEEPSEEK_ENABLED = DEEPSEEK_CONFIG.get("ENABLED", False)
DEEPSEEK_API_KEY = DEEPSEEK_CONFIG.get("API_KEY")
DEEPSEEK_MODEL = DEEPSEEK_CONFIG.get("MODEL", "deepseek-chat")

DEEPSEEK_BASE_URL = DEEPSEEK_CONFIG.get("BASE_URL", "https://api.deepseek.com")
DEFAULT_DEEPSEEK_PROMPT = DEEPSEEK_CONFIG.get("PROMPT", (
    "下面是一篇中文文章，请你【只修改错别字和明显的识别错误】。\n"
    "要求：\n"
    "1. 不改变原意\n"
    "2. 不润色文风\n"
    "3. 不增删内容\n"
    "4. 保持原有段落结构\n"
    "5. 只输出修改后的完整文章正文\n"
    "6. 格式应该是  标题  （\\n）下一行  ——xx(替换为姓名)  然后文章内容\n"
    "标题不要出现 ‘题目：’ ‘标题：’等字样\n\n"
    "{text}"
))
client = None

# ==================================================================




# 讯飞ocr相关设置
def getHeader():
    curTime = str(int(time.time()))
    param = json.dumps({"language": language, "location": location})
    paramBase64 = base64.b64encode(param.encode('utf-8')).decode('utf-8')
    checkSum_str = API_KEY + curTime + paramBase64
    checkSum = hashlib.md5(checkSum_str.encode('utf-8')).hexdigest()
    header = {
        'X-CurTime': curTime,
        'X-Param': paramBase64,
        'X-Appid': APPID,
        'X-CheckSum': checkSum,
        'Content-Type': 'application/x-www-form-urlencoded; charset=utf-8',
    }
    return header

# deepseek api 相关设置
def deepseek_fix_typos(prompt_template: str, text: str) -> str:
    """
    调用 DeepSeek，只修改错别字，不改变原意，不润色。
    `prompt_template` 可以包含 `{text}` 占位符；如果没有，会把 `text` 追加到末尾。
    """
    if not prompt_template:
        prompt_template = DEFAULT_DEEPSEEK_PROMPT

    if "{text}" in prompt_template:
        prompt = prompt_template.format(text=text)
    else:
        prompt = prompt_template + "\n\n" + text

    if debug:
        print("\n=================【发送给 DeepSeek 的内容】=================\n")
        print(prompt)
        print("\n============================================================\n")


    response = client.chat.completions.create(
        model=DEEPSEEK_MODEL,
        messages=[
            {"role": "system", "content": "你是一名严谨的中文校对助手"},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        stream=False
    )

    if debug:
        print("\n=================【DeepSeek 原始返回 response】=================\n")
        print(response)
        print("\n==============================================================\n")

    result_text = response.choices[0].message.content.strip()

    if debug:
        print("\n=================【DeepSeek 修改后的正文】=================\n")
        print(result_text)
        print("\n==========================================================\n")

    return result_text


# ocr 并处理自然段
def ocr_and_extract_text(image_path):
    with open(image_path, 'rb') as f:
        imgfile = f.read()

    data = {'image': base64.b64encode(imgfile).decode('utf-8')}
    headers = getHeader()
    resp = requests.post(URL, headers=headers, data=data)
    result = resp.json()
    # print("OCR 原始返回：", result)

    if result.get('code') != '0':
        raise RuntimeError(f"OCR 失败: {result.get('desc')}")

    # ========== ① 提取所有行 ==========
    lines = []
    for block in result.get("data", {}).get("block", []):
        if block.get("type") != "text":
            continue
        for line in block.get("line", []):
            text = "".join(w.get("content", "") for w in line.get("word", []))
            if text.strip():
                lines.append(text)

    # ========== ② 智能合并自然段 ==========
    paragraphs = []
    buffer = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if len(line) <= 4:
            buffer += line
            continue

        buffer += line

        if buffer[-1] in "。！？…" and len(buffer) >= 30:
            paragraphs.append(buffer)
            buffer = ""

    if buffer:
        paragraphs.append(buffer)

    return paragraphs   # ✅ 关键：返回段落列表


#文档处理配置
def create_word(doc_path, all_text_blocks,folder_display_name):

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(12)

    # 修改前
    para_before = doc.add_paragraph("修改前：")
    para_before.paragraph_format.first_line_indent = Cm(0.74)
    para_before.paragraph_format.space_before = Pt(0)
    para_before.paragraph_format.space_after = Pt(0)
    para_before.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    para_before.paragraph_format.line_spacing = Pt(12)

    # 姓名
    para_name = doc.add_paragraph(f"——{folder_display_name}")
    para_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_name.paragraph_format.space_before = Pt(0)
    para_name.paragraph_format.space_after = Pt(0)
    para_name.paragraph_format.line_spacing = Pt(12)

    for para_text in all_text_blocks:
        # OCR 段落
        p = doc.add_paragraph(para_text)
        p.paragraph_format.first_line_indent = Cm(0.74)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        p.paragraph_format.line_spacing = Pt(12)

    doc.add_page_break()
    # 修改后
    para_after = doc.add_paragraph("修改后：")
    para_after.paragraph_format.first_line_indent = Cm(0.74)
    para_after.paragraph_format.space_before = Pt(0)
    para_after.paragraph_format.space_after = Pt(0)
    para_after.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    para_after.paragraph_format.line_spacing = Pt(12)

    doc.save(doc_path)



def has_images(folder_path):
    return any(
        f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))
        for f in os.listdir(folder_path)
    )



# 识别图片
def process_folder(folder_path, log_callback=print, use_deepseek=False, deepseek_api_key=None, deepseek_base_url=None, deepseek_prompt_template=None, use_editor=False, editor_api_key=None, editor_base_url=None, editor_prompt_template=None):
    folder_name = os.path.basename(folder_path)
    all_paragraphs = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
            img_path = os.path.join(folder_path, filename)
            try:
                log_callback(f"正在识别：{img_path}")
                paragraphs = ocr_and_extract_text(img_path)
                all_paragraphs.extend(paragraphs)
            except Exception as e:
                log_callback(f"识别失败：{img_path}，原因：{e}")

    if all_paragraphs:
        doc_path = os.path.join(folder_path, f"{folder_name}.docx")
        log_callback(f"正在生成 Word：{doc_path}")

        # ① 先生成 Word（修改前 / 修改后 框架）
        create_word(doc_path, all_paragraphs, folder_name)

        # ② 再根据开关决定是否调用 DeepSeek
        if use_deepseek:
            log_callback("🤖 正在调用 DeepSeek 进行错别字纠正...")
            fix_docx_with_deepseek(doc_path, deepseek_api_key, base_url=deepseek_base_url, prompt_template=deepseek_prompt_template)
            log_callback("✅ DeepSeek 纠错完成")

        # ③ 如果启用了第二步编辑，再把纠正后的正文送到另一个 API 处理
        if use_editor:
            log_callback("🤖 正在调用 第二步 API 进行作文改写...")
            edit_docx_with_api(doc_path, editor_api_key, base_url=editor_base_url, prompt_template=editor_prompt_template)
            log_callback("✅ 第二步改写完成")

    else:
        log_callback(f"{folder_path} 中没有可识别的图片")

# 寻找‘修改前’与‘修改后’中 的文章
def extract_before_text(doc: Document):
    collecting = False
    paragraphs = []

    for p in doc.paragraphs:
        text = p.text.strip()

        if text == "修改前：":
            collecting = True
            continue

        if text == "修改后：":
            break

        if collecting and text:
            paragraphs.append(text)

    return paragraphs

# 清空「修改前」正文并写入新内容
def clear_before_text(doc: Document):
    start = end = None

    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "修改前：":
            start = i
        elif p.text.strip() == "修改后：" and start is not None:
            end = i
            break

    if start is None or end is None:
        return

    # 删除 start+1 到 end-1
    for i in range(end - 1, start, -1):
        p = doc.paragraphs[i]
        p._element.getparent().remove(p._element)





def insert_before_text(doc: Document, new_paragraphs):
    """
    在“修改前：”下面插入 AI 修正后的正文，并在“修改后：”之前插入分页符。
    如果没找到“修改后：”，则在刚插入的段落之后插入分页符（兜底）。
    """

    # 1) 找到“修改前：”所在位置（insert_index = 下一个段落的索引）
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "修改前：":
            insert_index = i + 1
            break
    else:
        print("⚠️ 未找到「修改前：」，跳过写入")
        return

    # 2) 记录已插入段落数，方便兜底时计算分页位置
    inserted_count = 0

    # 3) 插入 AI 返回的段落（保持顺序）
    for para in new_paragraphs:
        if not para or not para.strip():
            continue
        # 防止 AI 回传意外的标题
        if para.strip() in ("修改前：", "修改后："):
            continue

        new_p = doc.add_paragraph(para)

        # 设置格式：首行缩进、行距等
        fmt = new_p.paragraph_format
        fmt.first_line_indent = Cm(0.74)
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        fmt.line_spacing = Pt(12)

        # 把刚创建在文档末尾的段落移动到指定位置
        doc._body._body.insert(insert_index, new_p._element)
        insert_index += 1
        inserted_count += 1

    # 4) 尝试在“修改后：”之前插入分页符（优先）
    page_break_index = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip() == "修改后：":
            page_break_index = idx
            break

    # 5) 如果没找到“修改后：”，把分页符放在插入段落之后（兜底）
    if page_break_index is None:
        page_break_index = insert_index  # 刚插入段落后的索引

    # 6) 创建一个新的段落并在其 run 上添加分页符（page break）
    #    先 add_paragraph（位于文档末尾），然后把它移动到目标位置
    pb_para = doc.add_paragraph()
    run = pb_para.add_run()
    run.add_break(WD_BREAK.PAGE)

    # 插入分页段落到目标位置（这样分页就在 page_break_index 处）
    doc._body._body.insert(page_break_index, pb_para._element)


def insert_after_text(doc: Document, new_paragraphs):
    """
    在“修改后：”下面插入段落（不移动现有段落），如果未找到则追加到文档末尾。
    """
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "修改后：":
            insert_index = i + 1
            break
    else:
        insert_index = len(doc.paragraphs)

    for para in new_paragraphs:
        if not para or not para.strip():
            continue
        if para.strip() in ("修改前：", "修改后："):
            continue
        new_p = doc.add_paragraph(para)
        fmt = new_p.paragraph_format
        fmt.first_line_indent = Cm(0.74)
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        fmt.line_spacing = Pt(12)
        doc._body._body.insert(insert_index, new_p._element)
        insert_index += 1



# ai 纠错功能
def fix_docx_with_deepseek(docx_path, api_key, base_url=None, prompt_template=None):
    global client

    if not DEEPSEEK_ENABLED:
        print("⚠️ DeepSeek 未启用，跳过纠错")
        return

    if not api_key:
        print("⚠️ 未配置 DeepSeek API Key，跳过纠错")
        return

    client = OpenAI(
        api_key=api_key,
        base_url=(base_url or DEEPSEEK_BASE_URL)
    )

    doc = Document(docx_path)

    before_paragraphs = extract_before_text(doc)
    if not before_paragraphs:
        print("未找到修改前正文，跳过")
        return

    full_text = "\n".join(before_paragraphs)
    fixed_text = deepseek_fix_typos(prompt_template, full_text)

    fixed_paragraphs = [
        p.strip() for p in fixed_text.split("\n") if p.strip()
    ]

    clear_before_text(doc)
    insert_before_text(doc, fixed_paragraphs)

    doc.save(docx_path)


def edit_docx_with_api(docx_path, api_key, base_url=None, prompt_template=None, model=None):
    """
    使用另一个 AI API（与 DeepSeek 调用逻辑一致）对已经被纠错后的文章进行进一步修改，
    并将结果插入到“修改后：”之后。
    """
    global client

    if not api_key:
        print("⚠️ 第二步 API 未配置 API Key，跳过")
        return

    client = OpenAI(
        api_key=api_key,
        base_url=(base_url or DEEPSEEK_BASE_URL)
    )

    doc = Document(docx_path)

    # 读取已纠错的正文（"修改前：" 区域）
    corrected_paragraphs = extract_before_text(doc)
    if not corrected_paragraphs:
        print("未找到可供第二步处理的正文，跳过")
        return

    full_text = "\n".join(corrected_paragraphs)

    # 生成 prompt
    if not prompt_template:
        prompt_template = DEFAULT_DEEPSEEK_PROMPT

    if "{text}" in prompt_template:
        prompt = prompt_template.format(text=full_text)
    else:
        prompt = prompt_template + "\n\n" + full_text

    if debug:
        print("\n=================【发送给 第二步 API 的内容】=================\n")
        print(prompt)
        print("\n============================================================\n")

    response = client.chat.completions.create(
        model=(model or DEEPSEEK_MODEL),
        messages=[
            {"role": "system", "content": "你是一名严谨的中文写作编辑助手"},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        stream=False
    )

    result_text = response.choices[0].message.content.strip()

    edited_paragraphs = [p.strip() for p in result_text.split("\n") if p.strip()]

    insert_after_text(doc, edited_paragraphs)
    doc.save(docx_path)






# 遍历文件夹识别图片
def process_all(root_dir, log_callback=print, use_deepseek=False, deepseek_api_key=None, deepseek_base_url=None, deepseek_prompt_template=None, use_editor=False, editor_api_key=None, editor_base_url=None, editor_prompt_template=None):
    if has_images(root_dir):
        process_folder(root_dir, log_callback, use_deepseek, deepseek_api_key, deepseek_base_url, deepseek_prompt_template, use_editor, editor_api_key, editor_base_url, editor_prompt_template)
    else:
        for sub in os.listdir(root_dir):
            sub_path = os.path.join(root_dir, sub)
            if os.path.isdir(sub_path) and has_images(sub_path):
                process_folder(sub_path, log_callback, use_deepseek, deepseek_api_key, deepseek_base_url, deepseek_prompt_template, use_editor, editor_api_key, editor_base_url, editor_prompt_template)


if __name__ == '__main__':
    ROOT_DIR = input("请输入要处理的文件夹路径：").strip('" ')
    if not os.path.isdir(ROOT_DIR):
        print("无效路径！")
    else:
        process_all(
            ROOT_DIR,
            use_deepseek=DEEPSEEK_ENABLED,
            deepseek_api_key=DEEPSEEK_API_KEY,
            deepseek_base_url=DEEPSEEK_BASE_URL,
            deepseek_prompt_template=DEFAULT_DEEPSEEK_PROMPT
        )
        print("全部处理完成！")
