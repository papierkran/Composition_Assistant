import base64
import tkinter as tk
from tkinter import filedialog, messagebox
import json
import threading
from datetime import datetime
import sys
import os
import subprocess
import re
from pathlib import Path
from copy import deepcopy
import customtkinter as ctk
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING, WD_BREAK
from openai import OpenAI
from PIL import Image, ImageTk
import ctypes
import io


# 默认使用项目目录下的 config.json；如需覆盖，可通过环境变量 OCR_CONFIG_FILE 指定。
CONFIG_FILE = Path(os.environ.get("OCR_CONFIG_FILE", "config.json")).expanduser()

# ================= 默认配置（新结构） =================
DEFAULT_CONFIG = {
    "OCR": {
        "PROVIDER": "xfyun_handwriting",
        "XFYUN": {
            "URL": "http://webapi.xfyun.cn/v1/service/v1/ocr/handwriting",
            "APPID": "",
            "API_KEY": "",
            "LANGUAGE": "cn|en",
            "LOCATION": "false",
        },
    },
    "LLM": {
        "PROVIDERS": {
            "deepseek": {
                "API_KEY": "",
                "MODEL": "deepseek-chat",
                "BASE_URL": "https://api.deepseek.com/v1",
            },
            "openai": {
                "API_KEY": "",
                "MODEL": "gpt-4o-mini",
                "BASE_URL": "https://api.openai.com/v1",
            },
            "custom": {
                "API_KEY": "",
                "MODEL": "",
                "BASE_URL": "",
            },
        },
        "TASKS": {
            "typo_fix": {
                "ENABLED": False,
                "PROVIDER": "deepseek",
                "PROMPT": "{text}",
            },
            "editor": {
                "ENABLED": False,
                "PROVIDER": "deepseek",
                "PROMPT": "{text}",
            },
        },
    },
    "APP": {
        "ROOT_DIR": "",
        "DEBUG": False
    }
}

# ================= 配置文件 =================
from config_migrate import ensure_new_schema


def load_config(path: Path = None):
    """
    从 JSON 文件加载配置。
    尝试使用 `path`（或配置的 CONFIG_FILE）。
    如果缺失，则回退到 './config.json'。在任何读取错误的情况下，返回 DEFAULT_CONFIG 的浅拷贝以允许界面启动。
    """
    cfg_path = Path(path or CONFIG_FILE)
    if not cfg_path.exists():
        local = Path("config.json")
        if local.exists():
            cfg_path = local
        else:
            return deepcopy(DEFAULT_CONFIG)

    try:
        with cfg_path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return deepcopy(DEFAULT_CONFIG)

def save_config(config, path: Path = None):
    """Save configuration to file. Ensures parent directory exists."""
    cfg_path = Path(path or CONFIG_FILE)
    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    with cfg_path.open("w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

# ================= 日志 =================
def append_log(message: str):
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_text.configure(state="normal")
    log_text.insert(tk.END, f"[{timestamp}] {message}\n")
    log_text.see(tk.END)
    log_text.configure(state="disabled")

# ================= 配置编辑窗口（表单） =================
from config_editor_ui import open_config_editor_form


def open_config_editor():
    def _on_saved(new_cfg):
        # 更新内存 config，并尽量刷新主页面输入框
        global config
        config = ensure_new_schema(new_cfg)
        try:
            url_entry.delete(0, tk.END)
            url_entry.insert(0, config.get("OCR", {}).get("XFYUN", {}).get("URL", ""))
            appid_entry.delete(0, tk.END)
            appid_entry.insert(0, config.get("OCR", {}).get("XFYUN", {}).get("APPID", ""))
            path_entry.delete(0, tk.END)
            path_entry.insert(0, config.get("APP", {}).get("ROOT_DIR", ""))

            # 开关/提示词
            use_deepseek_var.set(bool((config.get("LLM", {}).get("TASKS", {}).get("typo_fix", {}) or {}).get("ENABLED", False)))
            prompt_text.delete("1.0", tk.END)
            prompt_text.insert("1.0", (config.get("LLM", {}).get("TASKS", {}).get("typo_fix", {}) or {}).get("PROMPT", "{text}"))

            use_editor_var.set(bool((config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("ENABLED", False)))
            editor_prompt_text.delete("1.0", tk.END)
            editor_prompt_text.insert("1.0", (config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("PROMPT", "{text}"))

            # 刷新 provider 下拉，避免删除后主界面仍显示旧 provider
            names = _provider_name_list()
            typo_provider = _normalize_provider_name(
                (config.get("LLM", {}).get("TASKS", {}).get("typo_fix", {}) or {}).get("PROVIDER", "deepseek")
            ) or "deepseek"
            edit_provider = _normalize_provider_name(
                (config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("PROVIDER", "deepseek")
            ) or "deepseek"
            if typo_provider not in names:
                typo_provider = names[0]
            if edit_provider not in names:
                edit_provider = names[0]

            deepseek_provider_menu.configure(values=names)
            deepseek_provider_menu.set(typo_provider)
            _on_deepseek_provider_change(typo_provider)

            editor_provider_menu.configure(values=names)
            editor_provider_menu.set(edit_provider)
            _on_editor_provider_change(edit_provider)

            if "ai_provider_menu" in globals():
                ai_provider_menu.configure(values=names)
                ai_provider = edit_provider if edit_provider in names else names[0]
                ai_provider_menu.set(ai_provider)
                _on_ai_provider_change(ai_provider)
        except Exception:
            pass

    open_config_editor_form(
        root=root,
        config=config,
        config_file=CONFIG_FILE,
        hidden_api_keys=hidden_api_keys,
        make_mask_widget=make_mask_widget,
        reveal_entry=reveal_entry,
        on_saved=_on_saved,
    )


# ================= 主逻辑 =================
def start_processing():
    def task():
        # ---------- 从 UI 读取（新 schema） ----------
        config.setdefault("OCR", {})
        config["OCR"].setdefault("XFYUN", {})
        config["OCR"]["PROVIDER"] = "xfyun_handwriting"
        config["OCR"]["XFYUN"]["URL"] = url_entry.get().strip()
        config["OCR"]["XFYUN"]["APPID"] = appid_entry.get().strip()
        config["OCR"]["XFYUN"]["API_KEY"] = get_api_key_value('ocr') or ""
        config["OCR"]["XFYUN"]["LANGUAGE"] = config["OCR"]["XFYUN"].get("LANGUAGE", "cn|en")
        config["OCR"]["XFYUN"]["LOCATION"] = config["OCR"]["XFYUN"].get("LOCATION", "false")

        config.setdefault("APP", {})
        config["APP"]["ROOT_DIR"] = path_entry.get().strip()

        config.setdefault("LLM", {})
        config["LLM"].setdefault("PROVIDERS", {})
        config["LLM"].setdefault("TASKS", {})

        # typo_fix task/provider
        typo_provider = _ensure_provider_exists(deepseek_provider_menu.get() or "deepseek")

        config["LLM"]["PROVIDERS"].setdefault(typo_provider, {})
        config["LLM"]["PROVIDERS"][typo_provider]["API_KEY"] = get_api_key_value('deepseek') or ""
        config["LLM"]["PROVIDERS"][typo_provider]["BASE_URL"] = deepseek_base_entry.get().strip()
        config["LLM"]["PROVIDERS"][typo_provider].setdefault("MODEL", "deepseek-chat" if typo_provider == "deepseek" else "gpt-4o-mini")

        config["LLM"]["TASKS"].setdefault("typo_fix", {})
        config["LLM"]["TASKS"]["typo_fix"]["ENABLED"] = use_deepseek_var.get()
        config["LLM"]["TASKS"]["typo_fix"]["PROVIDER"] = typo_provider
        config["LLM"]["TASKS"]["typo_fix"]["PROMPT"] = prompt_text.get("1.0", tk.END).strip()

        # editor task/provider
        editor_provider = _ensure_provider_exists(editor_provider_menu.get() or "deepseek")

        config["LLM"]["PROVIDERS"].setdefault(editor_provider, {})
        config["LLM"]["PROVIDERS"][editor_provider]["API_KEY"] = get_api_key_value('editor') or ""
        config["LLM"]["PROVIDERS"][editor_provider]["BASE_URL"] = editor_base_entry.get().strip()
        config["LLM"]["PROVIDERS"][editor_provider].setdefault("MODEL", "deepseek-chat" if editor_provider == "deepseek" else "gpt-4o-mini")

        config["LLM"]["TASKS"].setdefault("editor", {})
        config["LLM"]["TASKS"]["editor"]["ENABLED"] = use_editor_var.get()
        config["LLM"]["TASKS"]["editor"]["PROVIDER"] = editor_provider
        config["LLM"]["TASKS"]["editor"]["PROMPT"] = editor_prompt_text.get("1.0", tk.END).strip()

        # ---------- 校验 ----------
        if not all([
            config.get("OCR", {}).get("XFYUN", {}).get("URL"),
            config.get("OCR", {}).get("XFYUN", {}).get("APPID"),
            config.get("OCR", {}).get("XFYUN", {}).get("API_KEY"),
            config.get("APP", {}).get("ROOT_DIR"),
        ]):
            append_log("❌ 请填写完整的 OCR 配置和文件夹路径")
            return

        # 校验：启用 typo_fix 时必须有对应 provider 的 API Key
        tasks = (config.get("LLM", {}) or {}).get("TASKS", {})
        providers = (config.get("LLM", {}) or {}).get("PROVIDERS", {})
        if bool(tasks.get("typo_fix", {}).get("ENABLED", False)):
            p_name = tasks.get("typo_fix", {}).get("PROVIDER")
            if not (providers.get(p_name, {}) or {}).get("API_KEY"):
                append_log("❌ 已启用 AI 错别字修正，但未填写 API Key")
                return

        if not os.path.isdir(config["APP"]["ROOT_DIR"]):
            append_log("❌ 文件夹路径无效")
            return

        save_config(config)

        # ---------- 日志 ----------
        if bool(tasks.get("typo_fix", {}).get("ENABLED", False)):
            append_log("🧠 AI 错别字修正：已启用")
        else:
            append_log("🧠 AI 错别字修正：未启用")

        try:
            append_log("🚀 开始处理...")
            from ocr_main import process_all

            process_all(
                config["APP"]["ROOT_DIR"],
                log_callback=append_log,
                use_typo_fix=bool(tasks.get("typo_fix", {}).get("ENABLED", False)),
                use_editor=bool(tasks.get("editor", {}).get("ENABLED", False)),
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


def iter_files_limited(folder, max_depth=4):
    """遍历 folder，最多递归 max_depth 层（包含根目录为第0层）。
    返回 (root, files) 与 os.walk 类似，但在达到深度限制时不再向下递归。
    """
    folder = os.path.abspath(folder)
    for root, dirs, files in os.walk(folder, topdown=True):
        rel = os.path.relpath(root, folder)
        if rel == os.curdir:
            depth = 0
        else:
            depth = len(rel.split(os.sep))
        # 深度为 0..(max_depth-1) 可被处理；当达到阈值时阻止继续向下遍历
        if depth >= max_depth - 1:
            dirs[:] = []
        yield root, files

# ================= UI =================
config = ensure_new_schema(load_config())

# 存储已隐藏的 API Key
hidden_api_keys = {}
# 当前活跃的 entry 映射（name->entry 或 None）
entries_map = {}

# 预置 Provider 模板（可快速填充 BASE_URL）
API_PROVIDER_PRESETS = {
    "deepseek": "https://api.deepseek.com/v1",
    "openai": "https://api.openai.com/v1",
}


def _normalize_provider_name(name: str) -> str:
    return (name or "").strip().lower()


def _ensure_provider_exists(name: str):
    p_name = _normalize_provider_name(name)
    if not p_name:
        return ""
    config.setdefault("LLM", {}).setdefault("PROVIDERS", {}).setdefault(p_name, {})
    p_cfg = config["LLM"]["PROVIDERS"][p_name]
    p_cfg.setdefault("API_KEY", "")
    p_cfg.setdefault("BASE_URL", API_PROVIDER_PRESETS.get(p_name, ""))
    p_cfg.setdefault("MODEL", "deepseek-chat" if p_name == "deepseek" else "gpt-4o-mini")
    return p_name


def _provider_name_list():
    providers = (config.get("LLM", {}) or {}).get("PROVIDERS", {}) or {}
    names = {k for k in providers.keys() if isinstance(k, str) and k.strip()}
    names.update(API_PROVIDER_PRESETS.keys())
    return sorted(names)


def _open_new_provider_dialog(on_confirm):
    win = ctk.CTkToplevel(root)
    win.title("新增 AI Provider")
    win.geometry("520x220")
    try:
        win.transient(root)
        win.grab_set()
    except Exception:
        pass

    frm = ctk.CTkFrame(win)
    frm.pack(fill="both", expand=True, padx=12, pady=12)

    ctk.CTkLabel(frm, text="Provider 名称（如 xai / moonshot）").pack(anchor="w")
    name_entry = ctk.CTkEntry(frm)
    name_entry.pack(fill="x", pady=(4, 10))

    ctk.CTkLabel(frm, text="Base URL（可选）").pack(anchor="w")
    base_entry = ctk.CTkEntry(frm)
    base_entry.pack(fill="x", pady=(4, 10))

    def _save():
        p_name = _normalize_provider_name(name_entry.get())
        if not p_name:
            messagebox.showerror("新增失败", "Provider 名称不能为空")
            return
        _ensure_provider_exists(p_name)
        if base_entry.get().strip():
            config["LLM"]["PROVIDERS"][p_name]["BASE_URL"] = base_entry.get().strip()
        save_config(config)
        try:
            on_confirm(p_name)
        finally:
            win.destroy()

    btns = ctk.CTkFrame(frm)
    btns.pack(fill="x", pady=(6, 0))
    ctk.CTkButton(btns, text="取消", width=90, command=win.destroy).pack(side="right")
    ctk.CTkButton(btns, text="保存", width=90, fg_color="#4CAF50", text_color="white", command=_save).pack(side="right", padx=8)


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

icon_base64="""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"""

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Composition OCR Assistant 作文修改助手 v1.0")

# 按基准分辨率 720, 1280 等比缩放窗口布局（根据当前屏幕计算缩放因子）
BASE_W, BASE_H = 720, 1080
screen_w = root.winfo_screenwidth()
screen_h = root.winfo_screenheight()
scale = min(screen_w / BASE_W, screen_h / BASE_H)
try:
    root.tk.call('tk', 'scaling', scale)
except Exception:
    pass
# 将窗口设置为基准尺寸的 90%（再乘以缩放因子），并确保不超过屏幕
win_w = min(int(BASE_W * scale * 0.7), screen_w)
win_h = min(int(BASE_H * scale * 0.6), screen_h)
root.geometry(f"{win_w}x{win_h}")
root.resizable(True, True)

def _resource_path(name: str) -> Path:
    """兼容源码/打包环境的资源路径解析。"""
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / name


def _set_app_icon(win):
    """统一设置窗口左上角与任务栏图标。"""
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CompositionOCRAssistant.App")
    except Exception:
        pass

    ico_candidates = [
        _resource_path("app.ico"),
        Path.cwd() / "app.ico",
        Path(__file__).resolve().parent / "app.ico",
    ]

    for ico_path in ico_candidates:
        try:
            if ico_path.exists():
                abs_ico = str(ico_path.resolve())
                win.iconbitmap(abs_ico)
                # 某些 Windows 环境里，仅 iconbitmap 不会同步到任务栏，补一次 iconphoto。
                img = Image.open(abs_ico)
                photo = ImageTk.PhotoImage(img)
                win.iconphoto(True, photo)
                win._icon_photo = photo
                return
        except Exception:
            continue

    # 回退：使用 base64 内嵌图标，避免完全无图标。
    try:
        icon_data = base64.b64decode(icon_base64.strip())
        image = Image.open(io.BytesIO(icon_data)).resize((32, 32), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(image)
        win.iconphoto(True, photo)
        win._icon_photo = photo
    except Exception:
        try:
            photo = tk.PhotoImage(data=icon_base64.strip())
            win.iconphoto(True, photo)
            win._icon_photo = photo
        except Exception:
            pass


_set_app_icon(root)


# ========== 页面切换框架 ==========
pages = {}
current_page = tk.StringVar(value="ocr")

def show_page(page_name):
    """显示指定页面"""
    current_page.set(page_name)
    for page in pages.values():
        page.pack_forget()
    pages[page_name].pack(fill="both", expand=True)

# 顶部按钮框
top_frame = ctk.CTkFrame(root)
top_frame.pack(side="top", fill="x", padx=5, pady=5)

ctk.CTkLabel(top_frame, text="功能选择:", text_color="gray").pack(side="left", padx=5)
ctk.CTkButton(top_frame, text="图片转作文", width=100, command=lambda: show_page("ocr")).pack(side="left", padx=3)
ctk.CTkButton(top_frame, text="docx作文处理", width=100, command=lambda: show_page("ai")).pack(side="left", padx=3)

# 右上角：配置编辑（打开 JSON 配置的统一编辑窗口）
ctk.CTkButton(top_frame, text="配置编辑", width=100, command=lambda: open_config_editor()).pack(side="right", padx=3)

# ========== PAGE 1: OCR 处理 ==========
page1 = ctk.CTkScrollableFrame(root)
pages["ocr"] = page1

# OCR 配置（标签与输入框同一行）
ocr_frame = ctk.CTkFrame(page1)
ocr_frame.pack(padx=10, fill="x")
ctk.CTkLabel(ocr_frame, text="OCR 接口 URL").pack(side="left", padx=(0, 8))
url_entry = ctk.CTkEntry(ocr_frame)
url_entry.insert(0, config.get("OCR", {}).get("XFYUN", {}).get("URL", ""))
url_entry.pack(side="left", fill="x", expand=True)

# APPID 与 API_KEY 同行布局
appid_frame = ctk.CTkFrame(page1)
appid_frame.pack(padx=10, fill="x")
ctk.CTkLabel(appid_frame, text="APPID").pack(side="left", padx=(0, 8))
appid_entry = ctk.CTkEntry(appid_frame)
appid_entry.insert(0, config.get("OCR", {}).get("XFYUN", {}).get("APPID", ""))
appid_entry.pack(side="left", fill="x", expand=True)

apikey_frame = ctk.CTkFrame(page1)
apikey_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(apikey_frame, text="API_KEY").pack(side="left", padx=(0, 8))
apikey_entry = ctk.CTkEntry(apikey_frame)
apikey_entry.insert(0, config.get("OCR", {}).get("XFYUN", {}).get("API_KEY", ""))
apikey_entry.pack(side="left", fill="x", expand=True)
entries_map['ocr'] = apikey_entry
if config.get("OCR", {}).get("XFYUN", {}).get("API_KEY"):
    hidden_api_keys['ocr'] = config.get("OCR", {}).get("XFYUN", {}).get("API_KEY", "")
    apikey_entry.destroy()
    mask = make_mask_widget('ocr', apikey_frame)
    mask.pack(side="left")
    entries_map['ocr'] = None

# 第一步 AI 改错别字
typo_task_provider = _normalize_provider_name(
    (config.get("LLM", {}).get("TASKS", {}).get("typo_fix", {}) or {}).get("PROVIDER", "deepseek")
) or "deepseek"
editor_task_provider = _normalize_provider_name(
    (config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("PROVIDER", "deepseek")
) or "deepseek"
_ensure_provider_exists(typo_task_provider)
_ensure_provider_exists(editor_task_provider)
provider_names = _provider_name_list()

deepseek_frame = ctk.CTkFrame(page1)
deepseek_frame.pack(padx=10, fill="x", pady=(10,0))
ctk.CTkLabel(deepseek_frame, text="AI API Key（输入AI的apikey）").pack(side="left", padx=(0,8))
deepseek_entry = ctk.CTkEntry(deepseek_frame)
deepseek_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(typo_task_provider, {}) or {}).get("API_KEY", ""))
deepseek_entry.pack(side="left", fill="x", expand=True)
entries_map['deepseek'] = deepseek_entry
# 如果已配置 provider 的 API_KEY，则默认隐藏显示
if (config.get("LLM", {}).get("PROVIDERS", {}).get(typo_task_provider, {}) or {}).get("API_KEY"):
    hidden_api_keys['deepseek'] = (config.get("LLM", {}).get("PROVIDERS", {}).get(typo_task_provider, {}) or {}).get("API_KEY", "")
    deepseek_entry.destroy()
    mask = make_mask_widget('deepseek', deepseek_frame)
    mask.pack(side="left")
    entries_map['deepseek'] = None

use_deepseek_var = tk.BooleanVar(value=bool((config.get("LLM", {}).get("TASKS", {}).get("typo_fix", {}) or {}).get("ENABLED", False)))
ctk.CTkCheckBox(deepseek_frame, text="启用 AI 错别字自动修正（较慢）", variable=use_deepseek_var).pack(side="left", padx=8)

# DeepSeek Base URL 同行
deepseek_base_frame = ctk.CTkFrame(page1)
deepseek_base_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(deepseek_base_frame, text="AI 改错别字（Base URL）").pack(side="left", padx=(0,8))
deepseek_provider_name = typo_task_provider
deepseek_provider_menu = ctk.CTkOptionMenu(deepseek_base_frame, values=provider_names)
deepseek_provider_menu.set(deepseek_provider_name)
deepseek_provider_menu.pack(side="left", padx=(0,8))
ctk.CTkButton(deepseek_base_frame, text="+ 新增", width=68, command=lambda: _open_new_provider_dialog(_on_new_typo_provider)).pack(side="left", padx=(0,8))

deepseek_base_entry = ctk.CTkEntry(deepseek_base_frame)
deepseek_base_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(deepseek_provider_name, {}) or {}).get("BASE_URL", ""))
deepseek_base_entry.pack(side="left", fill="x", expand=True)

def _on_deepseek_provider_change(choice):
    p_name = _ensure_provider_exists(choice)
    deepseek_base_entry.configure(state="normal")
    deepseek_base_entry.delete(0, tk.END)
    deepseek_base_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(p_name, {}) or {}).get("BASE_URL", ""))

def _on_new_typo_provider(p_name):
    names = _provider_name_list()
    deepseek_provider_menu.configure(values=names)
    deepseek_provider_menu.set(p_name)
    _on_deepseek_provider_change(p_name)

deepseek_provider_menu.configure(command=_on_deepseek_provider_change)
_on_deepseek_provider_change(deepseek_provider_name)

# Prompt（多行）
prompt_frame = ctk.CTkFrame(page1)
prompt_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(prompt_frame, text="自定义修改错别字提示词").pack(side="left", padx=(0,8), anchor="n")
prompt_text = ctk.CTkTextbox(prompt_frame, height=140)
prompt_text.insert(
    "1.0",
    (config.get("LLM", {}).get("TASKS", {}).get("typo_fix", {}) or {}).get("PROMPT")
    or DEFAULT_CONFIG["LLM"]["TASKS"]["typo_fix"]["PROMPT"],
)
prompt_text.pack(side="left", fill="x", expand=True)

# 编辑 API（第二步）
use_editor_var = tk.BooleanVar(value=bool((config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("ENABLED", False)))
editor_enable_frame = ctk.CTkFrame(page1)
editor_enable_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkCheckBox(editor_enable_frame, text="启用 第二步 修改作文", variable=use_editor_var).pack(side="left")

# 第二步 API Key 同行
editor_key_frame = ctk.CTkFrame(page1)
editor_key_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(editor_key_frame, text="第二步 AI API Key（输入AI的apikey）").pack(side="left", padx=(0,8))
editor_key_entry = ctk.CTkEntry(editor_key_frame)
editor_key_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(editor_task_provider, {}) or {}).get("API_KEY", ""))
editor_key_entry.pack(side="left", fill="x", expand=True)
entries_map['editor'] = editor_key_entry
if (config.get("LLM", {}).get("PROVIDERS", {}).get(editor_task_provider, {}) or {}).get("API_KEY"):
    hidden_api_keys['editor'] = (config.get("LLM", {}).get("PROVIDERS", {}).get(editor_task_provider, {}) or {}).get("API_KEY", "")
    editor_key_entry.destroy()
    mask = make_mask_widget('editor', editor_key_frame)
    mask.pack(side="left")
    entries_map['editor'] = None

# 第二步 Base URL 同行
editor_base_frame = ctk.CTkFrame(page1)
editor_base_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(editor_base_frame, text="第二步 API Base URL（可选）").pack(side="left", padx=(0,8))
editor_provider_name = editor_task_provider
editor_provider_menu = ctk.CTkOptionMenu(editor_base_frame, values=provider_names)
editor_provider_menu.set(editor_provider_name)
editor_provider_menu.pack(side="left", padx=(0,8))
ctk.CTkButton(editor_base_frame, text="+ 新增", width=68, command=lambda: _open_new_provider_dialog(_on_new_editor_provider)).pack(side="left", padx=(0,8))

editor_base_entry = ctk.CTkEntry(editor_base_frame)
editor_base_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(editor_provider_name, {}) or {}).get("BASE_URL", ""))
editor_base_entry.pack(side="left", fill="x", expand=True)

def _on_editor_provider_change(choice):
    p_name = _ensure_provider_exists(choice)
    editor_base_entry.configure(state="normal")
    editor_base_entry.delete(0, tk.END)
    editor_base_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(p_name, {}) or {}).get("BASE_URL", ""))

def _on_new_editor_provider(p_name):
    names = _provider_name_list()
    editor_provider_menu.configure(values=names)
    editor_provider_menu.set(p_name)
    _on_editor_provider_change(p_name)

editor_provider_menu.configure(command=_on_editor_provider_change)
_on_editor_provider_change(editor_provider_name)

# 第二步 Prompt 多行
editor_prompt_frame = ctk.CTkFrame(page1)
editor_prompt_frame.pack(padx=10, fill="x", pady=(6,0))
ctk.CTkLabel(editor_prompt_frame, text="第二步 自定义 Prompt（可选）").pack(side="left", padx=(0,8), anchor="n")
editor_prompt_text = ctk.CTkTextbox(editor_prompt_frame, height=140)
editor_prompt_text.insert(
    "1.0",
    (config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("PROMPT")
    or DEFAULT_CONFIG["LLM"]["TASKS"]["editor"]["PROMPT"],
)
editor_prompt_text.pack(side="left", fill="x", expand=True)

# 路径
path_frame = ctk.CTkFrame(page1)
path_frame.pack(padx=10, fill="x", pady=(10,0))
ctk.CTkLabel(path_frame, text="作文文件夹路径").pack(side="left", padx=(0,8))
path_entry = ctk.CTkEntry(path_frame)
path_entry.insert(0, config["APP"]["ROOT_DIR"])
path_entry.pack(side="left", fill="x", expand=True)
ctk.CTkButton(path_frame, text="浏览", command=browse_folder).pack(side="left", padx=8)

# 启动
ctk.CTkButton(page1, text="开始处理", fg_color="#4CAF50", text_color="white", height=40, command=start_processing).pack(pady=12)

# 日志
ctk.CTkLabel(page1, text="运行日志").pack(anchor="w", padx=10)
log_text = ctk.CTkTextbox(page1, width=920, height=200)
log_text.configure(state="disabled")
log_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)

# ========== PAGE 2: AI 综合处理 - 可拖动任务流 ==========
page2 = ctk.CTkScrollableFrame(root)
pages["ai"] = page2

# 任务列表配置（顺序：6→1→AI→2→3→5）
task_config = [
    {"id": "6", "name": "6. 转换 DOC → DOCX", "enabled": True, "order": 0},
    {"id": "1", "name": "1. 清除空格", "enabled": True, "order": 1},
    {"id": "AI", "name": "🤖 AI 改作文", "enabled": True, "order": 2},
    {"id": "2", "name": '2. 添加"修改前/后"', "enabled": True, "order": 3},
    {"id": "3", "name": "3. 格式化字体段落", "enabled": True, "order": 4},
    {"id": "5", "name": "5. 修改作者", "enabled": True, "order": 5},
]

# 文件夹选择
ai_path_frame = ctk.CTkFrame(page2)
ai_path_frame.pack(padx=10, fill="x", pady=10)
ctk.CTkLabel(ai_path_frame, text="处理文件夹").pack(side="left", padx=(0, 8))
ai_path_entry = ctk.CTkEntry(ai_path_frame)
ai_path_entry.pack(side="left", fill="x", expand=True)
def browse_ai_folder():
    folder = filedialog.askdirectory()
    if folder:
        ai_path_entry.delete(0, tk.END)
        ai_path_entry.insert(0, folder)
ctk.CTkButton(ai_path_frame, text="浏览", command=browse_ai_folder, width=80).pack(side="left", padx=8)

# AI API Key
ai_task_provider = _normalize_provider_name(
    (config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("PROVIDER", "deepseek")
) or "deepseek"
_ensure_provider_exists(ai_task_provider)

ai_provider_frame = ctk.CTkFrame(page2)
ai_provider_frame.pack(padx=10, fill="x", pady=(6, 0))
ctk.CTkLabel(ai_provider_frame, text="AI Provider").pack(side="left", padx=(0, 8))
ai_provider_menu = ctk.CTkOptionMenu(ai_provider_frame, values=_provider_name_list())
ai_provider_menu.set(ai_task_provider)
ai_provider_menu.pack(side="left", padx=(0, 8))

def _on_ai_provider_change(choice):
    p_name = _ensure_provider_exists(choice)
    provider_cfg = (config.get("LLM", {}).get("PROVIDERS", {}).get(p_name, {}) or {})
    ai_key_entry.delete(0, tk.END)
    ai_key_entry.insert(0, provider_cfg.get("API_KEY", ""))
    ai_url_entry.delete(0, tk.END)
    ai_url_entry.insert(0, provider_cfg.get("BASE_URL", ""))

def _on_new_ai_provider(p_name):
    ai_provider_menu.configure(values=_provider_name_list())
    ai_provider_menu.set(p_name)
    _on_ai_provider_change(p_name)

ctk.CTkButton(ai_provider_frame, text="+ 新增", width=68, command=lambda: _open_new_provider_dialog(_on_new_ai_provider)).pack(side="left")
ai_provider_menu.configure(command=_on_ai_provider_change)

ai_key_frame = ctk.CTkFrame(page2)
ai_key_frame.pack(padx=10, fill="x", pady=(6, 0))
ctk.CTkLabel(ai_key_frame, text="AI API Key").pack(side="left", padx=(0, 8))
ai_key_entry = ctk.CTkEntry(ai_key_frame, show="*")
ai_key_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(ai_task_provider, {}) or {}).get("API_KEY", ""))
ai_key_entry.pack(side="left", fill="x", expand=True)

# Base URL
ai_url_frame = ctk.CTkFrame(page2)
ai_url_frame.pack(padx=10, fill="x", pady=(6, 0))
ctk.CTkLabel(ai_url_frame, text="API Base URL").pack(side="left", padx=(0, 8))
ai_url_entry = ctk.CTkEntry(ai_url_frame)
ai_url_entry.insert(0, (config.get("LLM", {}).get("PROVIDERS", {}).get(ai_task_provider, {}) or {}).get("BASE_URL", ""))
ai_url_entry.pack(side="left", fill="x", expand=True)

# Prompt
ai_prompt_frame = ctk.CTkFrame(page2)
ai_prompt_frame.pack(padx=10, fill="x", pady=(6, 0))
ctk.CTkLabel(ai_prompt_frame, text="AI Prompt").pack(side="left", padx=(0, 8), anchor="n")
ai_prompt_text = ctk.CTkTextbox(ai_prompt_frame, height=80)
ai_prompt_text.insert(
    "1.0",
    (config.get("LLM", {}).get("TASKS", {}).get("editor", {}) or {}).get("PROMPT", "{text}"),
)
ai_prompt_text.pack(side="left", fill="x", expand=True)

# 任务流标签框架 - 可拖动和排序
task_label = ctk.CTkLabel(page2, text="处理流程（勾选/取消步骤，拖动上下移动顺序）", text_color="gray", font=("", 12, "bold"))
task_label.pack(anchor="w", padx=10, pady=(10, 5))

# 创建任务流程容器框
task_container = ctk.CTkFrame(page2, fg_color=("#ffffff", "#2a2a2a"), border_width=2, border_color=("gray50", "gray50"))
task_container.pack(padx=10, pady=(0, 10), fill="both", expand=False)

task_frames_map = {}  # id -> {"frame": frame, "var": var}

def create_task_frame(task):
    """创建单个任务框，支持拖动排序 - 改进版实现"""
    frame = ctk.CTkFrame(task_container, fg_color=("#f0f0f0", "#1f1f1f"), border_width=1, border_color=("gray70", "gray30"))
    frame.pack(padx=10, fill="x", pady=4)
    
    # 左侧：勾选框 + 拖动把手 + 任务名
    left_frame = ctk.CTkFrame(frame, fg_color="transparent")
    left_frame.pack(side="left", padx=10, pady=8, fill="x", expand=True)
    
    var = tk.BooleanVar(value=task["enabled"])
    # 使用 trace 绑定，确保变量变化总是同步到 task 配置
    def _on_var_changed(*_):
        task["enabled"] = bool(var.get())
    var.trace_add("write", _on_var_changed)
    
    # 拖动把手（只在这个 widget 上绑定拖动事件）
    handle_label = ctk.CTkLabel(left_frame, text="☰", text_color="gray", font=("", 14), cursor="hand2")
    handle_label.pack(side="left", padx=(0, 5))
    
    # 拖动状态存储在 frame 的自定义属性中
    frame._drag_state = {
        "start_y": 0,
        "is_dragging": False,
        "initial_color": ("#f0f0f0", "#1f1f1f"),
        "initial_border": ("gray70", "gray30")
    }
    
    # 复选框（使用 trace 处理变量变化，无需额外回调）
    checkbox = ctk.CTkCheckBox(left_frame, text=task["name"], variable=var)
    checkbox.pack(side="left", fill="x", expand=True)
    
    # ========== 拖动事件处理（只绑定在 handle_label） ==========
    def on_handle_press(event):
        """鼠标按下：记录起始位置"""
        frame._drag_state["start_y"] = event.widget.winfo_rooty() + event.y
        frame._drag_state["is_dragging"] = False
        # 视觉提示：改变颜色表示可拖动
        frame.configure(border_color=("cyan", "lightcyan"), fg_color=("#e8f4f8", "#2a3a3a"))
        # 插入一个占位符（用于吸附效果），并暂时将被拖动的 frame 移除
        try:
            placeholder = ctk.CTkFrame(task_container, height=8, fg_color=("gray80", "gray20"))
            # 在当前 frame 前插入占位符以保持原位
            placeholder.pack(padx=10, fill="x", pady=4, before=frame)
            frame._drag_state["placeholder"] = placeholder
            # 把真实 frame 从布局中移除以便拖动并让占位符可见
            frame.pack_forget()
            frame.lift()
        except Exception:
            frame._drag_state.pop("placeholder", None)
    
    def on_handle_drag(event):
        """鼠标移动：检测并执行拖动"""
        current_y = event.widget.winfo_rooty() + event.y
        delta_y = current_y - frame._drag_state["start_y"]

        # 超过阈值才进入拖动状态
        if abs(delta_y) > 30 and not frame._drag_state["is_dragging"]:
            frame._drag_state["is_dragging"] = True
            frame.configure(border_color=("blue", "lightblue"), fg_color=("#d0e8ff", "#1a2a3a"))

        if not frame._drag_state.get("is_dragging"):
            return

        # 计算当前应该插入的位置（按容器内其它任务的中点比较）
        siblings = [f for f in task_container.winfo_children() if isinstance(f, ctk.CTkFrame) and getattr(f, "_drag_state", None) is not None and f != frame]
        if not siblings:
            frame._drag_state["insert_index"] = 0
            return

        insert_index = len(siblings)
        current_y_abs = event.widget.winfo_rooty() + event.y
        for i, other in enumerate(siblings):
            other_mid = other.winfo_rooty() + other.winfo_height() // 2
            if current_y_abs < other_mid:
                insert_index = i
                break

        # 移动占位符到 insert_index 位置（占位符负责视觉吸附）
        placeholder = frame._drag_state.get("placeholder")
        if placeholder:
            try:
                placeholder.pack_forget()
                if insert_index >= len(siblings):
                    last = siblings[-1]
                    placeholder.pack(padx=10, fill="x", pady=4, after=last)
                else:
                    placeholder.pack(padx=10, fill="x", pady=4, before=siblings[insert_index])
                frame._drag_state["insert_index"] = insert_index
            except Exception:
                pass
    
    def on_handle_release(event):
        """鼠标释放：恢复状态"""
        # 恢复原始视觉状态
        frame.configure(
            border_color=frame._drag_state["initial_border"],
            fg_color=frame._drag_state["initial_color"]
        )
        frame._drag_state["is_dragging"] = False

        # 将真实 frame 插入到占位符所在位置（占位符是在容器中的临时占位）
        insert_index = frame._drag_state.get("insert_index")
        placeholder = frame._drag_state.get("placeholder")

        # 清除占位符并在其原位置插入 frame
        try:
            # 列出当前（含占位符）容器子组件顺序，寻找 placeholder 的位置
            children = list(task_container.winfo_children())
            if placeholder in children:
                idx = children.index(placeholder)
            else:
                idx = None
            # 移除占位符视图
            if placeholder:
                placeholder.pack_forget()
        except Exception:
            idx = None

        # 重新获取当前容器内的任务 frames（不包含正在拖动的 frame）
        frames_now = [f for f in task_container.winfo_children() if isinstance(f, ctk.CTkFrame) and getattr(f, "_drag_state", None) is not None and f != frame]

        # 如果没有确定索引，使用 insert_index（拖动计算）作为回退
        if idx is None:
            if insert_index is None or insert_index >= len(frames_now):
                frame.pack(padx=10, fill="x", pady=4)
            else:
                frame.pack(padx=10, fill="x", pady=4, before=frames_now[insert_index])
        else:
            # 按 insert_index 插入
            if insert_index is None or insert_index >= len(frames_now):
                frame.pack(padx=10, fill="x", pady=4)
            else:
                frame.pack(padx=10, fill="x", pady=4, before=frames_now[insert_index])

        # 最终一次性更新所有任务的 order（仅容器内的任务）
        all_frames_current = [f for f in task_container.winfo_children() if isinstance(f, ctk.CTkFrame) and hasattr(f, "_drag_state")]
        for i, f in enumerate(all_frames_current):
            for task_item in task_config:
                if task_frames_map.get(task_item["id"], {}).get("frame") == f:
                    task_item["order"] = i

        # 清理拖动状态
        frame._drag_state.pop("placeholder", None)
        frame._drag_state.pop("insert_index", None)
    
    # 只在拖动把手上绑定事件
    handle_label.bind("<Button-1>", on_handle_press)
    handle_label.bind("<B1-Motion>", on_handle_drag)
    handle_label.bind("<ButtonRelease-1>", on_handle_release)
    
    # 存储任务框信息
    task_frames_map[task["id"]] = {
        "frame": frame,
        "var": var
    }
    
    return frame

# 创建所有任务框
for task in sorted(task_config, key=lambda x: x["order"]):
    create_task_frame(task)

# 启动按钮
def start_ai_workflow():
    def task():
        folder = ai_path_entry.get().strip()
        selected_provider = _ensure_provider_exists(ai_provider_menu.get() or "deepseek")
        api_key = ai_key_entry.get().strip()
        base_url = ai_url_entry.get().strip()
        prompt = ai_prompt_text.get("1.0", tk.END).strip() or None
        
        if not folder or not api_key:
            append_log_ai("❌ 请填写文件夹路径和 API Key")
            return
        
        if not os.path.isdir(folder):
            append_log_ai("❌ 文件夹路径无效")
            return
        
        # 保存配置（写入用户选择的 provider）
        config.setdefault("LLM", {})
        config["LLM"].setdefault("PROVIDERS", {})
        config["LLM"].setdefault("TASKS", {})

        config["LLM"]["PROVIDERS"].setdefault(selected_provider, {})
        config["LLM"]["PROVIDERS"][selected_provider]["API_KEY"] = api_key
        config["LLM"]["PROVIDERS"][selected_provider]["BASE_URL"] = base_url
        config["LLM"]["PROVIDERS"][selected_provider].setdefault("MODEL", "deepseek-chat" if selected_provider == "deepseek" else "gpt-4o-mini")

        config["LLM"]["TASKS"].setdefault("editor", {})
        config["LLM"]["TASKS"]["editor"]["PROMPT"] = prompt or "{text}"
        config["LLM"]["TASKS"]["editor"]["ENABLED"] = True
        config["LLM"]["TASKS"]["editor"]["PROVIDER"] = selected_provider

        save_config(config)
        
        append_log_ai("📋 开始处理流程...")
        
        # 第一步：复制所有原始文件为"改 XXX"（包含 docx 与常见图片格式，最多递归 4 层）
        append_log_ai("【准备】复制原始文件（包含图片）...")
        import shutil
        copied_files = []
        image_exts = (".png", ".jpg", ".jpeg", ".bmp", ".gif")
        for root, files in iter_files_limited(folder, max_depth=4):
            for file in files:
                name_check = file.lstrip()
                name_lower = file.lower()
                if (name_lower.endswith(".docx") or name_lower.endswith(image_exts)) and not name_check.startswith("~$") and not name_check.startswith("改 "):
                    original_path = os.path.join(root, file)
                    new_filename = f"改 {file}"
                    new_path = os.path.join(root, new_filename)
                    try:
                        shutil.copy2(original_path, new_path)
                        copied_files.append(new_filename)
                        append_log_ai(f"  ✓ {new_filename}")
                    except Exception as e:
                        append_log_ai(f"  ✗ {file} 复制失败: {e}")
        
        if not copied_files:
            append_log_ai("❌ 未找到需要处理的文件")
            return
        
        # 获取启用的任务列表
        enabled_tasks = sorted(
            [(t["id"], t["order"]) for t in task_config if t["enabled"]],
            key=lambda x: x[1]
        )
        
        try:
            for task_id, _ in enabled_tasks:
                if task_id == "6":
                    append_log_ai("【步骤 6】转换 DOC → DOCX...")
                    _convert_docs(folder, process_only_modified=True)
                elif task_id == "1":
                    append_log_ai("【步骤 1】清除空格...")
                    _clear_spaces(folder, process_only_modified=True)
                elif task_id == "AI":
                    append_log_ai("【步骤 AI】发送给 AI 修正...")
                    _process_ai(folder, api_key, base_url, prompt, process_only_modified=True)
                elif task_id == "2":
                    append_log_ai("【步骤 2】添加标签...")
                    _add_labels(folder, process_only_modified=True)
                elif task_id == "3":
                    append_log_ai("【步骤 3】格式化...")
                    _format_docs(folder, process_only_modified=True)
                elif task_id == "5":
                    append_log_ai("【步骤 5】修改作者...")
                    _set_author(folder, process_only_modified=True)
            
            append_log_ai("✅ 所有流程完成！")
        except Exception as e:
            append_log_ai(f"❌ 处理失败：{e}")
            import traceback
            traceback.print_exc()
    
    threading.Thread(target=task, daemon=True).start()

def append_log_ai(message: str):
    timestamp = datetime.now().strftime("%H:%M:%S")
    ai_log_text.configure(state="normal")
    ai_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
    ai_log_text.see(tk.END)
    ai_log_text.configure(state="disabled")

# 各步骤的处理函数
def _convert_docs(folder, process_only_modified=False):
    """步骤 6：转换 DOC → DOCX"""
    import subprocess
    for root, files in iter_files_limited(folder, max_depth=4):
        for file in files:
            # 兼容文件名前导空格与中文名
            name_check = file.lstrip()
            name_lower = file.lower()
            # 如果指定只处理修改文件，则跳过不以"改 "开头的文件
            if process_only_modified and not name_check.startswith("改 "):
                continue
            if name_lower.endswith(".doc") and not name_check.startswith("~$"):
                doc_path = os.path.join(root, file)
                try:
                    cmd = ["soffice", "--headless", "--convert-to", "docx", doc_path, "--outdir", root]
                    subprocess.run(cmd, capture_output=True, timeout=30)
                    base_name = os.path.splitext(os.path.basename(doc_path))[0]
                    new_path = os.path.join(root, base_name + ".docx")
                    if os.path.exists(new_path):
                        os.remove(doc_path)
                        append_log_ai(f"  ✓ {base_name}")
                except Exception as e:
                    append_log_ai(f"  ✗ {file}: {e}")

def _clear_spaces(folder, process_only_modified=False):
    """步骤 1：清除空格"""
    for root, files in iter_files_limited(folder, max_depth=4):
        for file in files:
            name_check = file.lstrip()
            name_lower = file.lower()
            # 如果指定只处理修改文件，则跳过不以"改 "开头的文件
            if process_only_modified and not name_check.startswith("改 "):
                continue
            if name_lower.endswith(".docx") and not name_check.startswith("~$"):
                doc_path = os.path.join(root, file)
                try:
                    doc = Document(doc_path)
                    for para in doc.paragraphs:
                        for run in para.runs:
                            run.text = run.text.strip()
                    doc.save(doc_path)
                    append_log_ai(f"  ✓ {file}")
                except Exception as e:
                    append_log_ai(f"  ✗ {file}: {e}")

def _process_ai(folder, api_key, base_url, prompt_template, process_only_modified=False):
    """步骤 AI：AI 处理 - 保留原始内容，追加修改后的内容"""
    if not prompt_template:
        prompt_template = (
            "下面是一篇中文文章，请你【只修改错别字和明显的识别错误】。\n"
            "要求：1. 不改变原意 2. 不润色文风 3. 不增删内容 4. 保持原有段落结构 5. 只输出修改后的完整文章正文\n"
        )
    
    client = OpenAI(api_key=api_key, base_url=base_url)
    for root, files in iter_files_limited(folder, max_depth=4):
        for file in files:
            name_check = file.lstrip()
            name_lower = file.lower()
            # 如果指定只处理修改文件，则跳过不以"改 "开头的文件
            if process_only_modified and not name_check.startswith("改 "):
                continue
            if name_lower.endswith(".docx") and not name_check.startswith("~$"):
                doc_path = os.path.join(root, file)
                try:
                    doc = Document(doc_path)
                    all_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                    
                    if not all_text.strip():
                        append_log_ai(f"  ⊘ {file} (空文档)")
                        continue
                    
                    if "{text}" in prompt_template:
                        full_prompt = prompt_template.format(text=all_text)
                    else:
                        full_prompt = prompt_template + "\n\n" + all_text
                    
                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": "你是一名严谨的中文校对助手"},
                            {"role": "user", "content": full_prompt},
                        ],
                        temperature=0.1,
                        stream=False
                    )
                    
                    ai_result = response.choices[0].message.content.strip()
                    
                    # 在原始内容末尾添加 "修改后：" 标签和空行
                    last_para = doc.paragraphs[-1] if doc.paragraphs else None
                    if last_para:
                        if last_para.runs:
                            last_para.runs[-1].add_break(WD_BREAK.PAGE)
                        else:
                            last_para.add_run().add_break(WD_BREAK.PAGE)
                    
                    # 添加 "修改后：" 标签
                    para_modify_label = doc.add_paragraph("修改后：")
                    para_modify_label.paragraph_format.first_line_indent = Cm(0.74)
                    para_modify_label.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                    para_modify_label.paragraph_format.line_spacing = Pt(12)
                    
                    # 追加 AI 修改后的内容
                    for line in ai_result.split("\n"):
                        if line.strip():
                            p = doc.add_paragraph(line.strip())
                            fmt = p.paragraph_format
                            fmt.first_line_indent = Cm(0.74)
                            fmt.space_before = Pt(0)
                            fmt.space_after = Pt(0)
                            fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                            fmt.line_spacing = Pt(12)
                    
                    doc.save(doc_path)
                    append_log_ai(f"  ✓ {file}")
                except Exception as e:
                    append_log_ai(f"  ✗ {file}: {e}")

def _add_labels(folder, process_only_modified=False):
    """步骤 2：添加标签 - 仅当 AI 步骤未运行时才添加分页符和修改后标签"""
    for root, files in iter_files_limited(folder, max_depth=4):
        for file in files:
            name_check = file.lstrip()
            name_lower = file.lower()
            # 如果指定只处理修改文件，则跳过不以"改 "开头的文件
            if process_only_modified and not name_check.startswith("改 "):
                continue
            if name_lower.endswith(".docx") and not name_check.startswith("~$"):
                doc_path = os.path.join(root, file)
                try:
                    doc = Document(doc_path)
                    if doc.paragraphs:
                        # 检查是否已经有"修改后："标签（说明 AI 步骤已执行）
                        last_para = doc.paragraphs[-1]
                        has_modify_label = last_para.text.strip() == "修改后：" or \
                                         (len(doc.paragraphs) > 1 and doc.paragraphs[-2].text.strip() == "修改后：")
                        
                        # 如果还没有"修改前："标签，则添加
                        if doc.paragraphs[0].text.strip() != "修改前：":
                            doc.paragraphs[0].insert_paragraph_before("修改前：")
                        
                        # 如果还没有"修改后："标签（AI 步骤未运行），则添加分页符和修改后标签
                        if not has_modify_label:
                            last_para = doc.paragraphs[-1]
                            if last_para.runs:
                                last_para.runs[-1].add_break(WD_BREAK.PAGE)
                            else:
                                last_para.add_run().add_break(WD_BREAK.PAGE)
                            para_after = doc.add_paragraph("修改后：")
                            para_after.paragraph_format.first_line_indent = Cm(0.74)
                            para_after.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                            para_after.paragraph_format.line_spacing = Pt(12)
                    doc.save(doc_path)
                    append_log_ai(f"  ✓ {file}")
                except Exception as e:
                    append_log_ai(f"  ✗ {file}: {e}")

def _format_docs(folder, process_only_modified=False):
    """步骤 3：格式化"""
    for root, files in iter_files_limited(folder, max_depth=4):
        for file in files:
            name_check = file.lstrip()
            name_lower = file.lower()
            # 如果指定只处理修改文件，则跳过不以"改 "开头的文件
            if process_only_modified and not name_check.startswith("改 "):
                continue
            if name_lower.endswith(".docx") and not name_check.startswith("~$"):
                doc_path = os.path.join(root, file)
                try:
                    doc = Document(doc_path)
                    style = doc.styles["Normal"]
                    style.font.name = "宋体"
                    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                    style.font.size = Pt(12)
                    for para in doc.paragraphs:
                        para.paragraph_format.first_line_indent = Cm(0.74)
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)
                        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                        para.paragraph_format.line_spacing = Pt(12)
                    doc.save(doc_path)
                    append_log_ai(f"  ✓ {file}")
                except Exception as e:
                    append_log_ai(f"  ✗ {file}: {e}")

def _set_author(folder, process_only_modified=False):
    """步骤 5：修改作者"""
    for root, files in iter_files_limited(folder, max_depth=4):
        for file in files:
            name_check = file.lstrip()
            name_lower = file.lower()
            # 如果指定只处理修改文件，则跳过不以"改 "开头的文件
            if process_only_modified and not name_check.startswith("改 "):
                continue
            if name_lower.endswith(".docx") and not name_check.startswith("~$"):
                doc_path = os.path.join(root, file)
                try:
                    doc = Document(doc_path)
                    doc.core_properties.author = "思睿教育_美丽可爱的尹老师"
                    doc.save(doc_path)
                    append_log_ai(f"  ✓ {file}")
                except Exception as e:
                    append_log_ai(f"  ✗ {file}: {e}")

ctk.CTkButton(page2, text="▶ 开始流程", fg_color="#2196F3", text_color="white", height=40, command=start_ai_workflow).pack(pady=12)

# 日志
ctk.CTkLabel(page2, text="处理日志").pack(anchor="w", padx=10)
ai_log_text = ctk.CTkTextbox(page2, width=920, height=220)
ai_log_text.configure(state="disabled")
ai_log_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)

# 默认显示第一页
show_page("ocr")

root.mainloop()
