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

    def row(parent, label: str):
        r = ctk.CTkFrame(parent)
        r.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(r, text=label, width=240, anchor="w").pack(side="left")
        field = ctk.CTkFrame(r, fg_color="transparent")
        field.pack(side="left", fill="x", expand=True)
        return field

    # ---- OCR (XFYUN) ----
    ocr_sec = section("OCR（讯飞）")
    ocr = (cfg.get("OCR", {}) or {}).get("XFYUN", {})

    ocr_url_row = row(ocr_sec, "OCR URL")
    ocr_url = ctk.CTkEntry(ocr_url_row)
    ocr_url.insert(0, ocr.get("URL", ""))
    ocr_url.pack(fill="x", expand=True)

    ocr_appid_row = row(ocr_sec, "APPID")
    ocr_appid = ctk.CTkEntry(ocr_appid_row)
    ocr_appid.insert(0, ocr.get("APPID", ""))
    ocr_appid.pack(fill="x", expand=True)

    # API_KEY masked
    ocr_key_row = row(ocr_sec, "API Key")
    ocr_key_frame = ctk.CTkFrame(ocr_key_row, fg_color="transparent")
    ocr_key_frame.pack(fill="x", expand=True)
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
    ocr_lang_row = row(ocr_sec, "LANGUAGE")
    ocr_lang = ctk.CTkEntry(ocr_lang_row)
    ocr_lang.insert(0, ocr.get("LANGUAGE", "cn|en"))
    ocr_lang.pack(fill="x", expand=True)

    ocr_loc_row = row(ocr_sec, "LOCATION (true/false)")
    ocr_loc = ctk.CTkEntry(ocr_loc_row)
    ocr_loc.insert(0, str(ocr.get("LOCATION", "false")))
    ocr_loc.pack(fill="x", expand=True)

    # ---- APP ----
    app_sec = section("应用")
    app = cfg.get("APP", {}) or {}

    root_dir_row = row(app_sec, "默认处理目录 ROOT_DIR")
    root_dir = ctk.CTkEntry(root_dir_row)
    root_dir.insert(0, app.get("ROOT_DIR", ""))
    root_dir.pack(fill="x", expand=True)

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
    providers_state = {k: dict(v or {}) for k, v in providers.items() if isinstance(k, str)}
    for built_in in ("deepseek", "openai"):
        providers_state.setdefault(built_in, {})

    provider_widgets: Dict[str, Dict[str, Any]] = {}

    provider_manage = ctk.CTkFrame(llm_sec)
    provider_manage.pack(fill="x", padx=10, pady=(6, 0))
    ctk.CTkLabel(provider_manage, text="Provider 管理", font=("", 13, "bold"), text_color="gray").pack(side="left")

    add_name_entry = ctk.CTkEntry(provider_manage, width=180, placeholder_text="新 provider 名称")
    add_name_entry.pack(side="left", padx=(10, 6))
    add_base_entry = ctk.CTkEntry(provider_manage, width=220, placeholder_text="可选 BASE_URL")
    add_base_entry.pack(side="left", padx=(0, 6))

    providers_container = ctk.CTkFrame(llm_sec)
    providers_container.pack(fill="x", padx=10, pady=(4, 0))

    def _provider_choices() -> list[str]:
        names = sorted({name.strip() for name in providers_state.keys() if name and name.strip()})
        return names or ["deepseek"]

    def _normalize_provider_name(name: str) -> str:
        return (name or "").strip().lower()

    def _is_valid_provider_name(name: str) -> bool:
        for ch in name:
            if not (ch.isalnum() or ch in ("_", "-", ".")):
                return False
        return True

    def _refresh_task_provider_menus():
        names = _provider_choices()
        typo_provider_menu.configure(values=names)
        edit_provider_menu.configure(values=names)
        if typo_provider.get() not in names:
            typo_provider.set(names[0])
        if edit_provider.get() not in names:
            edit_provider.set(names[0])

    def _render_provider_blocks():
        nonlocal provider_widgets
        provider_widgets = {}
        for child in providers_container.winfo_children():
            child.destroy()

        for provider_name in _provider_choices():
            p = providers_state.get(provider_name, {}) or {}
            box = ctk.CTkFrame(providers_container)
            box.pack(fill="x", pady=(6, 0))
            ctk.CTkLabel(
                box,
                text=f"Provider: {provider_name}",
                font=("", 13, "bold"),
                text_color="gray",
            ).pack(anchor="w", padx=10, pady=(8, 2))

            base_row = row(box, "BASE_URL（例：https://api.openai.com/v1）")
            base = ctk.CTkEntry(base_row)
            base.insert(0, p.get("BASE_URL", ""))
            base.pack(fill="x", expand=True)

            model_row = row(box, "MODEL（支持自定义模型 ID）")
            model = ctk.CTkEntry(model_row)
            model.insert(0, p.get("MODEL", ""))
            model.pack(fill="x", expand=True)

            key_row = row(box, "API_KEY")
            key_frame = ctk.CTkFrame(key_row, fg_color="transparent")
            key_frame.pack(fill="x", expand=True)
            key_entry = ctk.CTkEntry(key_frame)
            key_entry.insert(0, p.get("API_KEY", ""))
            key_entry.pack(side="left", fill="x", expand=True)
            if p.get("API_KEY"):
                hidden_api_keys[provider_name] = p.get("API_KEY", "")
                try:
                    key_entry.destroy()
                except Exception:
                    pass
                mask = make_mask_widget(provider_name, key_frame)
                mask.pack(side="left")

            provider_widgets[provider_name] = {
                "base": base,
                "model": model,
                "key": key_entry,
            }

    def _add_provider():
        raw_name = add_name_entry.get().strip()
        name = _normalize_provider_name(raw_name)
        if not name:
            messagebox.showerror("新增失败", "Provider 名称不能为空")
            return
        if not _is_valid_provider_name(name):
            messagebox.showerror("新增失败", "Provider 名称只允许字母/数字/下划线/中划线/点号")
            return
        if name in providers_state:
            messagebox.showinfo("提示", f"Provider '{name}' 已存在")
            return

        providers_state[name] = {
            "BASE_URL": add_base_entry.get().strip(),
            "MODEL": "gpt-4o-mini",
            "API_KEY": "",
        }
        add_name_entry.delete(0, tk.END)
        add_base_entry.delete(0, tk.END)
        _render_provider_blocks()
        _refresh_task_provider_menus()
        names = _provider_choices()
        delete_provider_menu.configure(values=names)
        delete_provider_var.set(name)

    delete_provider_var = tk.StringVar(value=_provider_choices()[0])
    delete_provider_menu = ctk.CTkOptionMenu(provider_manage, values=_provider_choices(), variable=delete_provider_var)
    delete_provider_menu.pack(side="left", padx=(10, 6))

    def _delete_provider():
        name = delete_provider_var.get().strip()
        if not name:
            return
        if name not in providers_state:
            return

        remaining = [p for p in _provider_choices() if p != name]
        if not remaining:
            messagebox.showerror("删除失败", "至少要保留 1 个 provider")
            return
        fallback = "deepseek" if "deepseek" in remaining else remaining[0]

        if not messagebox.askyesno("确认删除", f"确定删除 provider '{name}' 吗？"):
            return

        # 若任务正在引用被删 provider，自动切到兜底 provider。
        if typo_provider.get() == name:
            typo_provider.set(fallback)
        if edit_provider.get() == name:
            edit_provider.set(fallback)

        providers_state.pop(name, None)
        hidden_api_keys.pop(name, None)
        _render_provider_blocks()
        names = _provider_choices()
        delete_provider_var.set(fallback if fallback in names else names[0])
        delete_provider_menu.configure(values=names)
        _refresh_task_provider_menus()

    ctk.CTkButton(provider_manage, text="+ 新增", width=72, command=_add_provider).pack(side="left", padx=(0, 6))
    ctk.CTkButton(provider_manage, text="- 删除", width=72, command=_delete_provider).pack(side="left")

    def task_block(task_name: str, title: str):
        t = tasks.get(task_name, {}) or {}
        box = ctk.CTkFrame(llm_sec)
        box.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(box, text=title, font=("", 13, "bold"), text_color="gray").pack(anchor="w", padx=10, pady=(8, 2))

        enabled_var = tk.BooleanVar(value=bool(t.get("ENABLED", False)))
        ctk.CTkCheckBox(box, text="启用", variable=enabled_var).pack(anchor="w", padx=10, pady=(4, 6))

        default_provider = (t.get("PROVIDER") or "deepseek").strip()
        if default_provider not in providers_state:
            providers_state.setdefault(default_provider, {"BASE_URL": "", "MODEL": "gpt-4o-mini", "API_KEY": ""})

        prov = tk.StringVar(value=default_provider)
        prov_row = row(box, "使用 provider")
        prov_menu = ctk.CTkOptionMenu(prov_row, values=_provider_choices(), variable=prov)
        prov_menu.pack(side="left")

        pfr = row(box, "PROMPT（支持 {text}）")
        prompt = ctk.CTkTextbox(pfr, height=120)
        prompt.insert("1.0", t.get("PROMPT", "{text}"))
        prompt.pack(side="left", fill="x", expand=True)

        return enabled_var, prov, prompt, prov_menu

    typo_enabled, typo_provider, typo_prompt, typo_provider_menu = task_block("typo_fix", "任务：错别字修正（typo_fix）")
    edit_enabled, edit_provider, edit_prompt, edit_provider_menu = task_block("editor", "任务：第二步改写（editor）")
    _render_provider_blocks()

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
        new_cfg["OCR"]["XFYUN"]["API_KEY"] = (hidden_api_keys.get('ocr') or ocr_key_entry.get() or "").strip()
        new_cfg["OCR"]["XFYUN"]["LANGUAGE"] = ocr_lang.get().strip() or "cn|en"
        new_cfg["OCR"]["XFYUN"]["LOCATION"] = ocr_loc.get().strip() or "false"

        new_cfg.setdefault("APP", {})
        new_cfg["APP"]["ROOT_DIR"] = root_dir.get().strip()
        new_cfg["APP"]["DEBUG"] = bool(debug_var.get())

        new_cfg.setdefault("LLM", {})
        # 关键：先清空后按当前界面重建，确保被删除的 provider 不会残留
        new_cfg["LLM"]["PROVIDERS"] = {}
        new_cfg["LLM"].setdefault("TASKS", {})

        def set_provider(name: str, base_entry, model_entry, key_alias: str, key_entry):
            new_cfg["LLM"]["PROVIDERS"].setdefault(name, {})
            new_cfg["LLM"]["PROVIDERS"][name]["BASE_URL"] = base_entry.get().strip()
            # 支持用户自定义任意模型 ID（例如 deepseek-chat / gpt-4o-mini / 自建网关模型名）
            new_cfg["LLM"]["PROVIDERS"][name]["MODEL"] = model_entry.get().strip()
            # key stored masked
            new_cfg["LLM"]["PROVIDERS"][name]["API_KEY"] = (
                hidden_api_keys.get(key_alias) or key_entry.get() or ""
            ).strip()

        for p_name, widgets in provider_widgets.items():
            set_provider(
                p_name,
                widgets["base"],
                widgets["model"],
                p_name,
                widgets["key"],
            )

        def set_task(name: str, enabled_var, provider_var, prompt_widget):
            new_cfg["LLM"]["TASKS"].setdefault(name, {})
            new_cfg["LLM"]["TASKS"][name]["ENABLED"] = bool(enabled_var.get())
            new_cfg["LLM"]["TASKS"][name]["PROVIDER"] = provider_var.get().strip() or "deepseek"
            new_cfg["LLM"]["TASKS"][name]["PROMPT"] = prompt_widget.get("1.0", tk.END).strip() or "{text}"

        set_task("typo_fix", typo_enabled, typo_provider, typo_prompt)
        set_task("editor", edit_enabled, edit_provider, edit_prompt)

        # 允许空配置保存：这里不做必填拦截，仅保持结构化写入。

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
