import base64
import tkinter as tk
from tkinter import filedialog
import json
import threading
from datetime import datetime
import sys
import os
import customtkinter as ctk


CONFIG_FILE = "D:\person_data\ocer助手\presson.json"
# CONFIG_FILE = "config.json"

# ================= 默认配置 =================
DEFAULT_CONFIG = {
    "OCR": {
        "URL": "http://webapi.xfyun.cn/v1/service/v1/ocr/handwriting",
        "APPID": "",
        "API_KEY": "",
        "LANGUAGE": "cn|en",
        "LOCATION": "false"
    },
    "DEEPSEEK": {
        "ENABLED": False,
        "API_KEY": "",
        "MODEL": "deepseek-chat",
        "BASE_URL": "https://api.deepseek.com",
        "PROMPT": "下面是一篇中文文章，请你只修改错别字和明显的错误。\n"
              "要求：\n"
              "1. 不改变原意\n"
              "2. 不润色文风\n"
              "3. 不增删内容\n"
              "4. 保持原有段落结构\n"
              "5. 只输出修改后的文章，不做任何解释，注释，说明\n"
              "6.保留前两行的——及后面的姓名"
              "下面我将发给你文章\n"
                  "{text}"
    },
    "EDITOR": {
        "ENABLED": False,
        "API_KEY": "",
        "MODEL": "deepseek-chat",
        "BASE_URL": "https://api.deepseek.com",
        "PROMPT":"""你现在是一个优秀语文老师的中学语文老师，我将发给你文章，请严格遵循以下规则一直修改作文：
        【规则】
        格式指令：仅返回修改后文章，不包含任何解释与额外说明。
        （重要规则）字数控制：原文字数大于或等于850字，修改后保持在原文字数的50字误差内，字数基本和原文一致；原文小于850字，扩充至820-850字内。所有“字数”均以中文字符为准（不含空格与标点），字数必须保证在650字以上。
        段落结构：全文分为7-8个清晰段落，确保结构分明。
        首尾限定：文章首段与尾段均严格控制在70字以内，若必须收束全文并超出上限，可最多放宽至100字，但必须在结尾明确点题与升华，不能太突兀。
        结尾点题：倒数第二段需要包含文章主题的升华，在点明文章主题，升华主旨，最后一段的结尾在尽量在三句内结束。
        描写量化：细节描写不少于240字，合理分布于全文并至少覆盖3个不同段落；每处细节不少于30字。
        环境点缀：插入2-3处高级环境描写，运用四字词与动宾结构，不得生硬堆砌。
        语言优化：替换基础动词为高级词汇，强化动宾结构，适当融入成语典故，用准确、具体的高级词汇；不得改变原文事实主干与语气，不能太突兀。
        内容忠实：严格保留原文核心内容与主体，不得大幅改动。
        开篇技巧：开头应快速且自然地入题。
        过渡要求：善用关联词与过渡句，确保段落衔接流畅自然。
        详略安排：重点内容具体详写，次要内容简洁略写，要详略得当。
        情感核心：基于文章中的经历与真实感受书写，表达真情实感。
        对话、诗歌与题记：不改变相关内容，原文是什么，修改后就还是什么，不得变动对话、诗歌与题记

        请你严格按照以上【规则】修改
        修改后的文章请直接以以下格式呈现：
        修改后：
        （此处仅为修改后内容）
        当我说  重新修改时  ，你需要重新思考规则并重新按规则修改
        接下来我将发给你文章，请你按照规则修改并输出修改后的文章，记住只输出修改后的文章，不要任何解释和多余的内容。 """
    },
    "APP": {
        "ROOT_DIR": "",
        "DEBUG": False
    }
}

# ================= 配置文件 =================
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()

def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

# ================= 日志 =================
def append_log(message):
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_text.configure(state="normal")
    log_text.insert(tk.END, f"[{timestamp}] {message}\n")
    log_text.see(tk.END)
    log_text.configure(state="disabled")

# ================= 主逻辑 =================
def start_processing():
    def task():
        # ---------- 从 UI 读取 ----------
        config["OCR"]["URL"] = url_entry.get().strip()
        config["OCR"]["APPID"] = appid_entry.get().strip()
        config["OCR"]["API_KEY"] = get_api_key_value('ocr') or ""
        config["APP"]["ROOT_DIR"] = path_entry.get().strip()
        config["DEEPSEEK"]["API_KEY"] = get_api_key_value('deepseek') or ""
        config["DEEPSEEK"]["ENABLED"] = use_deepseek_var.get()
        config["DEEPSEEK"]["BASE_URL"] = deepseek_base_entry.get().strip()
        config["DEEPSEEK"]["PROMPT"] = prompt_text.get("1.0", tk.END).strip()
        # 第二步 API 配置
        config.setdefault("EDITOR", {})
        config["EDITOR"]["ENABLED"] = use_editor_var.get()
        config["EDITOR"]["API_KEY"] = get_api_key_value('editor') or ""
        config["EDITOR"]["BASE_URL"] = editor_base_entry.get().strip()
        config["EDITOR"]["PROMPT"] = editor_prompt_text.get("1.0", tk.END).strip()

        # ---------- 校验 ----------
        if not all([
            config["OCR"]["URL"],
            config["OCR"]["APPID"],
            config["OCR"]["API_KEY"],
            config["APP"]["ROOT_DIR"]
        ]):
            append_log("❌ 请填写完整的 OCR 配置和文件夹路径")
            return

        if config["DEEPSEEK"]["ENABLED"] and not config["DEEPSEEK"]["API_KEY"]:
            append_log("❌ 已启用 DeepSeek，但未填写 API Key")
            return

        if not os.path.isdir(config["APP"]["ROOT_DIR"]):
            append_log("❌ 文件夹路径无效")
            return

        save_config(config)

        # ---------- 日志 ----------
        if config["DEEPSEEK"]["ENABLED"]:
            append_log("🧠 DeepSeek 错别字修正：已启用")
        else:
            append_log("🧠 DeepSeek 错别字修正：未启用")

        try:
            append_log("🚀 开始处理...")
            from ocr_main import process_all

            process_all(
                config["APP"]["ROOT_DIR"],
                log_callback=append_log,
                use_deepseek=config["DEEPSEEK"]["ENABLED"],
                deepseek_api_key=config["DEEPSEEK"]["API_KEY"],
                deepseek_base_url=config["DEEPSEEK"].get("BASE_URL"),
                deepseek_prompt_template=config["DEEPSEEK"].get("PROMPT"),
                use_editor=config.get("EDITOR", {}).get("ENABLED", False),
                editor_api_key=config.get("EDITOR", {}).get("API_KEY"),
                editor_base_url=config.get("EDITOR", {}).get("BASE_URL"),
                editor_prompt_template=config.get("EDITOR", {}).get("PROMPT")
            )

            append_log("✅ 全部处理完成")
        except Exception as e:
            append_log(f"❌ 处理失败：{e}")

    threading.Thread(target=task, daemon=True).start()

# ================= 选择文件夹 =================
def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        path_entry.delete(0, tk.END)
        path_entry.insert(0, folder)

# ================= UI =================
config = load_config()

# 存储已隐藏的 API Key
hidden_api_keys = {}
# 当前活跃的 entry 映射（name->entry 或 None）
entries_map = {}


def make_mask_widget(name, parent):
    frame = ctk.CTkFrame(parent)
    lbl = ctk.CTkLabel(frame, text="API Key 已设置（隐藏）")
    def on_show():
        reveal_entry(name, parent, frame)
    btn = ctk.CTkButton(frame, text="显示", width=60, command=on_show)
    lbl.pack(side="left")
    btn.pack(side="left", padx=8)
    return frame


def reveal_entry(name, parent, mask_frame=None):
    if mask_frame:
        mask_frame.destroy()
    ent = ctk.CTkEntry(parent)
    ent.insert(0, hidden_api_keys.get(name, ""))
    ent.pack(side="left", fill="x", expand=True)
    ent.bind("<FocusOut>", lambda e: on_focus_out_mask(name, ent, parent))
    entries_map[name] = ent


def on_focus_out_mask(name, entry_widget, parent):
    val = entry_widget.get().strip()
    if val:
        hidden_api_keys[name] = val
        entry_widget.destroy()
        mask = make_mask_widget(name, parent)
        mask.pack(side="left")
        entries_map[name] = None


def get_api_key_value(name):
    ent = entries_map.get(name)
    if ent:
        return ent.get().strip()
    return hidden_api_keys.get(name)

icon_base64=""" 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 """

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Composition OCR Assistant v0.5")

# 根据屏幕分辨率按 16:9 比例缩放窗口（占屏幕 40%）
screen_w = root.winfo_screenwidth()
screen_h = root.winfo_screenheight()
max_w = int(screen_w * 0.9)
max_h = int(screen_h * 0.7)
desired_w = min(max_w, int(max_h * 10 / 16))
desired_h = int(desired_w * 16 / 9)
root.geometry(f"{desired_w}x{desired_h}")
root.resizable(False, False)

icon_data = base64.b64decode(icon_base64)
photo = tk.PhotoImage(data=icon_base64)
root.iconphoto(True, photo)


# OCR 配置（标签与输入框同一行）
ocr_frame = ctk.CTkFrame(root)
ocr_frame.pack(padx=10, fill="x")
ctk.CTkLabel(ocr_frame, text="OCR 接口 URL").pack(side="left", padx=(0, 8))
url_entry = ctk.CTkEntry(ocr_frame)
url_entry.insert(0, config["OCR"]["URL"])
url_entry.pack(side="left", fill="x", expand=True)

# APPID 与 API_KEY 同行布局
appid_frame = ctk.CTkFrame(root)
appid_frame.pack(padx=10, fill="x")
ctk.CTkLabel(appid_frame, text="APPID").pack(side="left", padx=(0, 8))
appid_entry = ctk.CTkEntry(appid_frame)
appid_entry.insert(0, config["OCR"]["APPID"])
appid_entry.pack(side="left", fill="x", expand=True)

apikey_frame = ctk.CTkFrame(root)
apikey_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(apikey_frame, text="API_KEY").pack(side="left", padx=(0, 8))
apikey_entry = ctk.CTkEntry(apikey_frame)
apikey_entry.insert(0, config["OCR"]["API_KEY"])
apikey_entry.pack(side="left", fill="x", expand=True)
# 记录并根据已有配置隐藏
entries_map['ocr'] = apikey_entry
if config["OCR"].get("API_KEY"):
    hidden_api_keys['ocr'] = config["OCR"]["API_KEY"]
    apikey_entry.destroy()
    mask = make_mask_widget('ocr', apikey_frame)
    mask.pack(side="left")
    entries_map['ocr'] = None

# DeepSeek
# DeepSeek API Key 与 启用开关 同行
deepseek_frame = ctk.CTkFrame(root)
deepseek_frame.pack(padx=10, fill="x", pady=(10,0))
ctk.CTkLabel(deepseek_frame, text="AI API Key（输入AI的apikey）").pack(side="left", padx=(0,8))
deepseek_entry = ctk.CTkEntry(deepseek_frame)
deepseek_entry.insert(0, config["DEEPSEEK"]["API_KEY"])
deepseek_entry.pack(side="left", fill="x", expand=True)
# 记录并根据已有配置隐藏
entries_map['deepseek'] = deepseek_entry
if config["DEEPSEEK"].get("API_KEY"):
    hidden_api_keys['deepseek'] = config["DEEPSEEK"]["API_KEY"]
    deepseek_entry.destroy()
    mask = make_mask_widget('deepseek', deepseek_frame)
    mask.pack(side="left")
    entries_map['deepseek'] = None

use_deepseek_var = tk.BooleanVar(value=config["DEEPSEEK"]["ENABLED"])
ctk.CTkCheckBox(deepseek_frame, text="启用 AI 错别字自动修正（较慢）", variable=use_deepseek_var).pack(side="left", padx=8)

# DeepSeek Base URL 同行
deepseek_base_frame = ctk.CTkFrame(root)
deepseek_base_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(deepseek_base_frame, text="AI 改错别字（Base URL）").pack(side="left", padx=(0,8))
deepseek_base_entry = ctk.CTkEntry(deepseek_base_frame)
deepseek_base_entry.insert(0, config["DEEPSEEK"].get("BASE_URL", "https://api.deepseek.com"))
deepseek_base_entry.pack(side="left", fill="x", expand=True)

# Prompt（多行）
# DeepSeek Prompt（多行）标签与文本框在同一行（文本框高度固定）
prompt_frame = ctk.CTkFrame(root)
prompt_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(prompt_frame, text="自定义修改错别字提示词").pack(side="left", padx=(0,8), anchor="n")
prompt_text = ctk.CTkTextbox(prompt_frame, height=140)
prompt_text.insert("1.0", config["DEEPSEEK"].get("PROMPT") or DEFAULT_CONFIG["DEEPSEEK"]["PROMPT"])
prompt_text.pack(side="left", fill="x", expand=True)

# ----- 编辑 API（第二步）
use_editor_var = tk.BooleanVar(value=config.get("EDITOR", {}).get("ENABLED", False))
editor_enable_frame = ctk.CTkFrame(root)
editor_enable_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkCheckBox(editor_enable_frame, text="启用 第二步 修改作文", variable=use_editor_var).pack(side="left")

# 第二步 API Key 同行
editor_key_frame = ctk.CTkFrame(root)
editor_key_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(editor_key_frame, text="第二步 AI API Key（输入AI的apikey）").pack(side="left", padx=(0,8))
editor_key_entry = ctk.CTkEntry(editor_key_frame)
editor_key_entry.insert(0, config.get("EDITOR", {}).get("API_KEY", ""))
editor_key_entry.pack(side="left", fill="x", expand=True)
# 记录并根据已有配置隐藏
entries_map['editor'] = editor_key_entry
if config.get("EDITOR", {}).get("API_KEY"):
    hidden_api_keys['editor'] = config.get("EDITOR", {}).get("API_KEY")
    editor_key_entry.destroy()
    mask = make_mask_widget('editor', editor_key_frame)
    mask.pack(side="left")
    entries_map['editor'] = None

# 第二步 Base URL 同行
editor_base_frame = ctk.CTkFrame(root)
editor_base_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(editor_base_frame, text="第二步 API Base URL（可选）").pack(side="left", padx=(0,8))
editor_base_entry = ctk.CTkEntry(editor_base_frame)
editor_base_entry.insert(0, config.get("EDITOR", {}).get("BASE_URL", "https://api.deepseek.com"))
editor_base_entry.pack(side="left", fill="x", expand=True)

# 第二步 Prompt 多行（标签与文本框同一行，文本框固定高度）
editor_prompt_frame = ctk.CTkFrame(root)
editor_prompt_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(editor_prompt_frame, text="第二步 自定义 Prompt（可选）").pack(side="left", padx=(0,8), anchor="n")
editor_prompt_text = ctk.CTkTextbox(editor_prompt_frame, height=140)
editor_prompt_text.insert("1.0", config.get("EDITOR", {}).get("PROMPT") or DEFAULT_CONFIG["EDITOR"]["PROMPT"])
editor_prompt_text.pack(side="left", fill="x", expand=True)

# 路径
# 作文文件夹路径（标签 与 输入框 同行）
path_frame = ctk.CTkFrame(root)
path_frame.pack(padx=10, fill="x", pady=(10,0))
ctk.CTkLabel(path_frame, text="作文文件夹路径").pack(side="left", padx=(0,8))
path_entry = ctk.CTkEntry(path_frame)
path_entry.insert(0, config["APP"]["ROOT_DIR"])
path_entry.pack(side="left", fill="x", expand=True)
ctk.CTkButton(path_frame, text="浏览", command=browse_folder).pack(side="left", padx=8)

# 启动
ctk.CTkButton(root, text="开始处理", fg_color="#4CAF50", text_color="white", height=40, command=start_processing).pack(pady=12)

# 日志
ctk.CTkLabel(root, text="运行日志").pack(anchor="w", padx=10)
log_text = ctk.CTkTextbox(root, width=920, height=200)
log_text.configure(state="disabled")
log_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)

root.mainloop()
