# coding=utf-8
"""Form-based config editor UI for CustomTkinter.

Instead of showing raw JSON, this provides a key/value form.
We still keep JSON internally and can optionally expose an "Advanced" JSON view.

This module is designed to be imported by ocr_gui.py.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Optional

import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

from config_migrate import ensure_new_schema


def _get_cfg_path(default_path: Path) -> Path:
    p = Path(default_path)
    if p.exists():
        return p
    local = Path("config.json")
    if local.exists():
        return local
    return p


def _read_cfg_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception:
        return ""


def _write_cfg(path: Path, cfg: Dict[str, Any]):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def open_config_editor_form(
    *,
    root,
    config: Dict[str, Any],
    config_file: Path,
    hidden_api_keys: Dict[str, str],
    make_mask_widget,
    reveal_entry,
    on_saved=None,
):
    """Open a modal config editor (form-based).

    Parameters:
      - root: main CTk root
      - config: current in-memory config (already ensure_new_schema)
      - config_file: CONFIG_FILE Path used by app
      - hidden_api_keys/make_mask_widget/reveal_entry: reuse existing API key masking system
      - on_saved(new_cfg): callback after saving successfully
    """

    win = ctk.CTkToplevel(root)
    win.title("配置编辑")
    win.geometry("980x720")

    try:
        win.transient(root)
        win.grab_set()
    except Exception:
        pass

    cfg_path = _get_cfg_path(config_file)

    # Load from disk as baseline if possible
    disk_raw = _read_cfg_text(cfg_path)
    if disk_raw.strip():
        try:
            disk_cfg = ensure_new_schema(json.loads(disk_raw))
        except Exception:
            disk_cfg = config
    else:
        disk_cfg = config

    cfg = disk_cfg

    # ---- layout ----
    header = ctk.CTkFrame(win)
    header.pack(fill="x", padx=12, pady=(12, 6))
    ctk.CTkLabel(header, text=f"当前配置文件：{cfg_path}", text_color="gray").pack(side="left")

    body = ctk.CTkScrollableFrame(win)
    body.pack(fill="both", expand=True, padx=12, pady=6)

    def section(title: str):
        box = ctk.CTkFrame(body)
        box.pack(fill="x", expand=False, pady=(8, 0))
        ctk.CTkLabel(box, text=title, font=("", 14, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
        inner = ctk.CTkFrame(body)
        inner.pack(fill="x", expand=False)
        return inner

    def row(parent, label: str, widget):
        r = ctk.CTkFrame(parent)
        r.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(r, text=label, width=240, anchor="w").pack(side="left")
        widget.pack(in_=r, side="left", fill="x", expand=True)
        return r

    # ---- OCR (XFYUN) ----
    ocr_sec = section("OCR（讯飞）")
    ocr = (cfg.get("OCR", {}) or {}).get("XFYUN", {})

    ocr_url = ctk.CTkEntry(ocr_sec)
    ocr_url.insert(0, ocr.get("URL", ""))
    row(ocr_sec, "OCR URL", ocr_url)

    ocr_appid = ctk.CTkEntry(ocr_sec)
    ocr_appid.insert(0, ocr.get("APPID", ""))
    row(ocr_sec, "APPID", ocr_appid)

    # API_KEY masked
    ocr_key_frame = ctk.CTkFrame(ocr_sec)
    ocr_key_entry = ctk.CTkEntry(ocr_key_frame)
    ocr_key_entry.insert(0, ocr.get("API_KEY", ""))
    # reuse masking system: name 'ocr'
    # put entry inside frame to allow swap
    ocr_key_entry.pack(side="left", fill="x", expand=True)

    # register by caller's reveal_entry system
    # if already has key, hide it
    if ocr.get("API_KEY"):
        hidden_api_keys['ocr'] = ocr.get("API_KEY", "")
        try:
            ocr_key_entry.destroy()
        except Exception:
            pass
        mask = make_mask_widget('ocr', ocr_key_frame)
        mask.pack(side="left")
    row(ocr_sec, "API Key", ocr_key_frame)

    ocr_lang = ctk.CTkEntry(ocr_sec)
    ocr_lang.insert(0, ocr.get("LANGUAGE", "cn|en"))
    row(ocr_sec, "LANGUAGE", ocr_lang)

    ocr_loc = ctk.CTkEntry(ocr_sec)
    ocr_loc.insert(0, str(ocr.get("LOCATION", "false")))
    row(ocr_sec, "LOCATION (true/false)", ocr_loc)

    # ---- APP ----
    app_sec = section("应用")
    app = cfg.get("APP", {}) or {}

    root_dir = ctk.CTkEntry(app_sec)
    root_dir.insert(0, app.get("ROOT_DIR", ""))
    row(app_sec, "默认处理目录 ROOT_DIR", root_dir)

    debug_var = tk.BooleanVar(value=bool(app.get("DEBUG", False)))
    debug_cb = ctk.CTkCheckBox(app_sec, text="启用 DEBUG（会打印 prompt/response）", variable=debug_var)
    rdbg = ctk.CTkFrame(app_sec)
    rdbg.pack(fill="x", padx=10, pady=6)
    debug_cb.pack(in_=rdbg, side="left")

    # ---- LLM Providers/Tasks ----
    llm_sec = section("LLM（OpenAI-compatible）")
    llm = cfg.get("LLM", {}) or {}
    providers = llm.get("PROVIDERS", {}) or {}
    tasks = llm.get("TASKS", {}) or {}

    provider_choices = ["deepseek", "openai", "custom"]

    def provider_block(title: str, provider_name: str, key_alias: str):
        box = ctk.CTkFrame(llm_sec)
        box.pack(fill="x", padx=10, pady=(6, 0))
        ctk.CTkLabel(box, text=title, font=("", 13, "bold"), text_color="gray").pack(anchor="w", padx=10, pady=(8, 2))

        p = providers.get(provider_name, {}) or {}

        base = ctk.CTkEntry(box)
        base.insert(0, p.get("BASE_URL", ""))
        row(box, "BASE_URL（例：https://api.openai.com/v1）", base)

        model = ctk.CTkEntry(box)
        model.insert(0, p.get("MODEL", ""))
        row(box, "MODEL", model)

        # API key masked
        key_frame = ctk.CTkFrame(box)
        key_entry = ctk.CTkEntry(key_frame)
        key_entry.insert(0, p.get("API_KEY", ""))
        key_entry.pack(side="left", fill="x", expand=True)
        if p.get("API_KEY"):
            hidden_api_keys[key_alias] = p.get("API_KEY", "")
            try:
                key_entry.destroy()
            except Exception:
                pass
            mask = make_mask_widget(key_alias, key_frame)
            mask.pack(side="left")
        row(box, "API_KEY", key_frame)

        return base, model

    deepseek_base, deepseek_model = provider_block("Provider: deepseek", "deepseek", "deepseek")
    openai_base, openai_model = provider_block("Provider: openai", "openai", "openai")
    custom_base, custom_model = provider_block("Provider: custom", "custom", "custom")

    def task_block(task_name: str, title: str):
        t = tasks.get(task_name, {}) or {}
        box = ctk.CTkFrame(llm_sec)
        box.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(box, text=title, font=("", 13, "bold"), text_color="gray").pack(anchor="w", padx=10, pady=(8, 2))

        enabled_var = tk.BooleanVar(value=bool(t.get("ENABLED", False)))
        ctk.CTkCheckBox(box, text="启用", variable=enabled_var).pack(anchor="w", padx=10, pady=(4, 6))

        prov = tk.StringVar(value=(t.get("PROVIDER") or "deepseek"))
        prov_menu = ctk.CTkOptionMenu(box, values=provider_choices, variable=prov)
        pr = ctk.CTkFrame(box)
        pr.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(pr, text="使用 provider", width=240, anchor="w").pack(side="left")
        prov_menu.pack(side="left")

        prompt = ctk.CTkTextbox(box, height=120)
        prompt.insert("1.0", t.get("PROMPT", "{text}"))
        pfr = ctk.CTkFrame(box)
        pfr.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(pfr, text="PROMPT（支持 {text}）", width=240, anchor="w").pack(side="left", anchor="n")
        prompt.pack(in_=pfr, side="left", fill="x", expand=True)

        return enabled_var, prov, prompt

    typo_enabled, typo_provider, typo_prompt = task_block("typo_fix", "任务：错别字修正（typo_fix）")
    edit_enabled, edit_provider, edit_prompt = task_block("editor", "任务：第二步改写（editor）")

    # ---- Advanced JSON (optional) ----
    adv = ctk.CTkFrame(body)
    adv.pack(fill="x", padx=0, pady=(12, 0))
    adv_var = tk.BooleanVar(value=False)

    def _toggle_adv():
        if adv_var.get():
            adv_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        else:
            adv_box.pack_forget()

    ctk.CTkCheckBox(adv, text="高级：显示当前 JSON", variable=adv_var, command=_toggle_adv).pack(anchor="w", padx=10, pady=8)
    adv_box = ctk.CTkTextbox(body, height=180)

    def _refresh_adv_json(new_cfg: Dict[str, Any]):
        adv_box.delete("1.0", tk.END)
        adv_box.insert("1.0", json.dumps(new_cfg, ensure_ascii=False, indent=2))

    _refresh_adv_json(cfg)

    # ---- actions ----
    footer = ctk.CTkFrame(win)
    footer.pack(fill="x", padx=12, pady=(6, 12))

    def on_reload():
        nonlocal cfg
        raw = _read_cfg_text(cfg_path)
        if raw.strip():
            try:
                cfg = ensure_new_schema(json.loads(raw))
            except Exception as e:
                messagebox.showerror("重载失败", f"配置文件 JSON 解析失败：\n{e}")
                return
        _refresh_adv_json(cfg)
        messagebox.showinfo("已重载", "已从配置文件重载到编辑器（当前窗口内控件不会自动回填，建议关闭后重新打开）")

    def on_save():
        # Build new cfg from fields
        new_cfg = ensure_new_schema(cfg)
        new_cfg.setdefault("OCR", {})
        new_cfg["OCR"]["PROVIDER"] = "xfyun_handwriting"
        new_cfg["OCR"].setdefault("XFYUN", {})
        new_cfg["OCR"]["XFYUN"]["URL"] = ocr_url.get().strip()
        new_cfg["OCR"]["XFYUN"]["APPID"] = ocr_appid.get().strip()
        new_cfg["OCR"]["XFYUN"]["API_KEY"] = (hidden_api_keys.get('ocr') or "").strip()
        new_cfg["OCR"]["XFYUN"]["LANGUAGE"] = ocr_lang.get().strip() or "cn|en"
        new_cfg["OCR"]["XFYUN"]["LOCATION"] = ocr_loc.get().strip() or "false"

        new_cfg.setdefault("APP", {})
        new_cfg["APP"]["ROOT_DIR"] = root_dir.get().strip()
        new_cfg["APP"]["DEBUG"] = bool(debug_var.get())

        new_cfg.setdefault("LLM", {})
        new_cfg["LLM"].setdefault("PROVIDERS", {})
        new_cfg["LLM"].setdefault("TASKS", {})

        def set_provider(name: str, base_entry, model_entry, key_alias: str):
            new_cfg["LLM"]["PROVIDERS"].setdefault(name, {})
            new_cfg["LLM"]["PROVIDERS"][name]["BASE_URL"] = base_entry.get().strip()
            new_cfg["LLM"]["PROVIDERS"][name]["MODEL"] = model_entry.get().strip()
            # key stored masked
            new_cfg["LLM"]["PROVIDERS"][name]["API_KEY"] = (hidden_api_keys.get(key_alias) or "").strip()

        set_provider("deepseek", deepseek_base, deepseek_model, "deepseek")
        set_provider("openai", openai_base, openai_model, "openai")
        set_provider("custom", custom_base, custom_model, "custom")

        def set_task(name: str, enabled_var, provider_var, prompt_widget):
            new_cfg["LLM"]["TASKS"].setdefault(name, {})
            new_cfg["LLM"]["TASKS"][name]["ENABLED"] = bool(enabled_var.get())
            new_cfg["LLM"]["TASKS"][name]["PROVIDER"] = provider_var.get().strip() or "deepseek"
            new_cfg["LLM"]["TASKS"][name]["PROMPT"] = prompt_widget.get("1.0", tk.END).strip() or "{text}"

        set_task("typo_fix", typo_enabled, typo_provider, typo_prompt)
        set_task("editor", edit_enabled, edit_provider, edit_prompt)

        # simple validation
        if not new_cfg["OCR"]["XFYUN"]["URL"] or not new_cfg["OCR"]["XFYUN"]["APPID"] or not new_cfg["OCR"]["XFYUN"]["API_KEY"]:
            messagebox.showerror("保存失败", "OCR 配置不完整（URL/APPID/API_KEY 必填）")
            return

        try:
            _write_cfg(cfg_path, new_cfg)
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return

        _refresh_adv_json(new_cfg)
        messagebox.showinfo("已保存", f"配置已保存到：{cfg_path}")
        if on_saved:
            try:
                on_saved(new_cfg)
            except Exception:
                pass

    ctk.CTkButton(footer, text="重载", width=90, command=on_reload).pack(side="left")
    ctk.CTkButton(footer, text="保存", width=90, fg_color="#4CAF50", text_color="white", command=on_save).pack(side="right")
    ctk.CTkButton(footer, text="关闭", width=90, command=win.destroy).pack(side="right", padx=8)

    return win
