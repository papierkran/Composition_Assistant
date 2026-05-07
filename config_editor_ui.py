# coding=utf-8
"""Form-based config editor UI for PySide6."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QTextEdit, QCheckBox, QComboBox, QPushButton, QScrollArea, QFrame,
    QMessageBox, QFormLayout, QGroupBox,
)
from PySide6.QtCore import Qt, Signal

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
    parent,
    config: Dict[str, Any],
    config_file: Path,
    hidden_api_keys: Dict[str, str],
    on_saved=None,
):
    """Open a modal config editor (form-based).

    Parameters:
      - parent: main QWidget (parent for the dialog)
      - config: current in-memory config
      - config_file: CONFIG_FILE Path
      - hidden_api_keys: dict of masked API keys
      - on_saved(new_cfg): callback after saving
    """
    cfg_path = _get_cfg_path(config_file)
    disk_raw = _read_cfg_text(cfg_path)
    if disk_raw.strip():
        try:
            cfg = ensure_new_schema(json.loads(disk_raw))
        except Exception:
            cfg = config
    else:
        cfg = config

    win = QWidget(parent)
    win.setWindowTitle("配置编辑")
    win.resize(980, 720)
    win.setWindowFlags(Qt.Window | Qt.Dialog)
    win.setWindowModality(Qt.WindowModal)

    layout = QVBoxLayout(win)

    # Header
    header = QLabel(f"当前配置文件：{cfg_path}")
    header.setStyleSheet("color: gray;")
    layout.addWidget(header)

    # Scroll area for form
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll_widget = QWidget()
    scroll_layout = QVBoxLayout(scroll_widget)

    # ---- OCR ----
    ocr_group = QGroupBox("OCR（讯飞）")
    ocr_form = QFormLayout()
    ocr = (cfg.get("OCR", {}) or {}).get("XFYUN", {})

    ocr_url = QLineEdit(ocr.get("URL", ""))
    ocr_appid = QLineEdit(ocr.get("APPID", ""))
    ocr_key = QLineEdit(ocr.get("API_KEY", ""))
    ocr_key.setEchoMode(QLineEdit.Password)
    ocr_lang = QLineEdit(ocr.get("LANGUAGE", "cn|en"))
    ocr_loc = QLineEdit(str(ocr.get("LOCATION", "false")))

    ocr_form.addRow("URL", ocr_url)
    ocr_form.addRow("APPID", ocr_appid)
    ocr_form.addRow("API Key", ocr_key)
    ocr_form.addRow("LANGUAGE", ocr_lang)
    ocr_form.addRow("LOCATION", ocr_loc)
    ocr_group.setLayout(ocr_form)
    scroll_layout.addWidget(ocr_group)

    # ---- APP ----
    app_group = QGroupBox("应用")
    app_form = QFormLayout()
    app = cfg.get("APP", {}) or {}

    root_dir = QLineEdit(app.get("ROOT_DIR", ""))
    debug_var = QCheckBox("启用 DEBUG")
    debug_var.setChecked(bool(app.get("DEBUG", False)))

    app_form.addRow("默认处理目录 ROOT_DIR", root_dir)
    app_form.addRow(debug_var)
    app_group.setLayout(app_form)
    scroll_layout.addWidget(app_group)

    # ---- LLM Providers ----
    llm_group = QGroupBox("LLM（OpenAI-compatible）")
    llm_layout = QVBoxLayout()
    llm = cfg.get("LLM", {}) or {}
    providers = llm.get("PROVIDERS", {}) or {}
    tasks = llm.get("TASKS", {}) or {}
    providers_state = {k: dict(v or {}) for k, v in providers.items() if isinstance(k, str)}
    for built_in in ("deepseek", "openai"):
        providers_state.setdefault(built_in, {})

    provider_widgets: Dict[str, Dict[str, Any]] = {}

    def _provider_choices() -> list[str]:
        return sorted({n.strip() for n in providers_state.keys() if n and n.strip()}) or ["deepseek"]

    # Provider management
    prov_header = QHBoxLayout()
    prov_header.addWidget(QLabel("Provider 管理"))
    add_name = QLineEdit()
    add_name.setPlaceholderText("新 provider 名称")
    add_name.setFixedWidth(180)
    add_base = QLineEdit()
    add_base.setPlaceholderText("可选 BASE_URL")
    add_base.setFixedWidth(220)
    prov_header.addWidget(add_name)
    prov_header.addWidget(add_base)

    delete_combo = QComboBox()
    delete_combo.addItems(_provider_choices())
    prov_header.addWidget(delete_combo)

    btn_add = QPushButton("+ 新增")
    btn_del = QPushButton("- 删除")
    prov_header.addWidget(btn_add)
    prov_header.addWidget(btn_del)
    llm_layout.addLayout(prov_header)

    providers_container = QVBoxLayout()
    llm_layout.addLayout(providers_container)

    def _render_providers():
        # Clear
        while providers_container.count():
            w = providers_container.takeAt().widget()
            if w:
                w.deleteLater()
        provider_widgets.clear()

        for pname in _provider_choices():
            p = providers_state.get(pname, {}) or {}
            box = QGroupBox(f"Provider: {pname}")
            form = QFormLayout()
            base = QLineEdit(p.get("BASE_URL", ""))
            model = QLineEdit(p.get("MODEL", ""))
            key = QLineEdit(p.get("API_KEY", ""))
            key.setEchoMode(QLineEdit.Password)
            form.addRow("BASE_URL", base)
            form.addRow("MODEL", model)
            form.addRow("API_KEY", key)
            box.setLayout(form)
            providers_container.addWidget(box)
            provider_widgets[pname] = {"base": base, "model": model, "key": key}

    def _on_add():
        name = add_name.text().strip().lower()
        if not name:
            QMessageBox.warning(win, "新增失败", "Provider 名称不能为空")
            return
        if name in providers_state:
            QMessageBox.information(win, "提示", f"Provider '{name}' 已存在")
            return
        providers_state[name] = {"BASE_URL": add_base.text().strip(), "MODEL": "gpt-4o-mini", "API_KEY": ""}
        add_name.clear()
        add_base.clear()
        _render_providers()
        _refresh_delete_combo()

    def _on_delete():
        name = delete_combo.currentText().strip()
        if not name or name not in providers_state:
            return
        remaining = [p for p in _provider_choices() if p != name]
        if not remaining:
            QMessageBox.warning(win, "删除失败", "至少要保留 1 个 provider")
            return
        reply = QMessageBox.question(win, "确认删除", f"确定删除 provider '{name}' 吗？")
        if reply == QMessageBox.Yes:
            providers_state.pop(name, None)
            hidden_api_keys.pop(name, None)
            _render_providers()
            _refresh_delete_combo()

    def _refresh_delete_combo():
        delete_combo.clear()
        delete_combo.addItems(_provider_choices())

    btn_add.clicked.connect(_on_add)
    btn_del.clicked.connect(_on_delete)

    # Tasks
    typo_task = tasks.get("typo_fix", {}) or {}
    edit_task = tasks.get("editor", {}) or {}

    typo_enabled = QCheckBox("启用")
    typo_enabled.setChecked(bool(typo_task.get("ENABLED", False)))
    typo_provider = QComboBox()
    typo_provider.addItems(_provider_choices())
    idx = typo_provider.findText(typo_task.get("PROVIDER", "deepseek"))
    if idx >= 0:
        typo_provider.setCurrentIndex(idx)
    typo_prompt = QTextEdit()
    typo_prompt.setPlainText(typo_task.get("PROMPT", "{text}"))
    typo_prompt.setMaximumHeight(120)

    edit_enabled = QCheckBox("启用")
    edit_enabled.setChecked(bool(edit_task.get("ENABLED", False)))
    edit_provider = QComboBox()
    edit_provider.addItems(_provider_choices())
    idx = edit_provider.findText(edit_task.get("PROVIDER", "deepseek"))
    if idx >= 0:
        edit_provider.setCurrentIndex(idx)
    edit_prompt = QTextEdit()
    edit_prompt.setPlainText(edit_task.get("PROMPT", "{text}"))
    edit_prompt.setMaximumHeight(120)

    typo_group = QGroupBox("任务：错别字修正（typo_fix）")
    typo_form = QFormLayout()
    typo_form.addRow(typo_enabled)
    typo_form.addRow("使用 provider", typo_provider)
    typo_form.addRow("PROMPT", typo_prompt)
    typo_group.setLayout(typo_form)
    llm_layout.addWidget(typo_group)

    edit_group = QGroupBox("任务：第二步改写（editor）")
    edit_form = QFormLayout()
    edit_form.addRow(edit_enabled)
    edit_form.addRow("使用 provider", edit_provider)
    edit_form.addRow("PROMPT", edit_prompt)
    edit_group.setLayout(edit_form)
    llm_layout.addWidget(edit_group)

    llm_group.setLayout(llm_layout)
    scroll_layout.addWidget(llm_group)

    # Advanced JSON
    adv_check = QCheckBox("高级：显示当前 JSON")
    scroll_layout.addWidget(adv_check)
    adv_text = QTextEdit()
    adv_text.setReadOnly(True)
    adv_text.setMaximumHeight(180)
    adv_text.setVisible(False)
    adv_text.setPlainText(json.dumps(cfg, ensure_ascii=False, indent=2))
    scroll_layout.addWidget(adv_text)

    def _toggle_adv():
        adv_text.setVisible(adv_check.isChecked())
    adv_check.toggled.connect(_toggle_adv)

    scroll_layout.addStretch()
    scroll.setWidget(scroll_widget)
    layout.addWidget(scroll, 1)

    # Footer
    footer = QHBoxLayout()
    btn_reload = QPushButton("重载")
    btn_save = QPushButton("保存")
    btn_save.setStyleSheet("background-color: #4CAF50; color: white;")
    btn_close = QPushButton("关闭")
    footer.addWidget(btn_reload)
    footer.addStretch()
    footer.addWidget(btn_close)
    footer.addWidget(btn_save)
    layout.addLayout(footer)

    def _on_save():
        new_cfg = ensure_new_schema(cfg)
        new_cfg.setdefault("OCR", {})
        new_cfg["OCR"]["PROVIDER"] = "xfyun_handwriting"
        new_cfg["OCR"].setdefault("XFYUN", {})
        new_cfg["OCR"]["XFYUN"]["URL"] = ocr_url.text().strip()
        new_cfg["OCR"]["XFYUN"]["APPID"] = ocr_appid.text().strip()
        new_cfg["OCR"]["XFYUN"]["API_KEY"] = ocr_key.text().strip()
        new_cfg["OCR"]["XFYUN"]["LANGUAGE"] = ocr_lang.text().strip() or "cn|en"
        new_cfg["OCR"]["XFYUN"]["LOCATION"] = ocr_loc.text().strip() or "false"

        new_cfg.setdefault("APP", {})
        new_cfg["APP"]["ROOT_DIR"] = root_dir.text().strip()
        new_cfg["APP"]["DEBUG"] = debug_var.isChecked()

        new_cfg.setdefault("LLM", {})
        new_cfg["LLM"]["PROVIDERS"] = {}
        new_cfg["LLM"].setdefault("TASKS", {})

        for pname, w in provider_widgets.items():
            new_cfg["LLM"]["PROVIDERS"][pname] = {
                "BASE_URL": w["base"].text().strip(),
                "MODEL": w["model"].text().strip(),
                "API_KEY": w["key"].text().strip(),
            }

        new_cfg["LLM"]["TASKS"]["typo_fix"] = {
            "ENABLED": typo_enabled.isChecked(),
            "PROVIDER": typo_provider.currentText().strip() or "deepseek",
            "PROMPT": typo_prompt.toPlainText().strip() or "{text}",
        }
        new_cfg["LLM"]["TASKS"]["editor"] = {
            "ENABLED": edit_enabled.isChecked(),
            "PROVIDER": edit_provider.currentText().strip() or "deepseek",
            "PROMPT": edit_prompt.toPlainText().strip() or "{text}",
        }

        try:
            _write_cfg(cfg_path, new_cfg)
        except Exception as e:
            QMessageBox.warning(win, "保存失败", str(e))
            return

        QMessageBox.information(win, "已保存", f"配置已保存到：{cfg_path}")
        if on_saved:
            try:
                on_saved(new_cfg)
            except Exception:
                pass

    def _on_reload():
        nonlocal cfg
        raw = _read_cfg_text(cfg_path)
        if raw.strip():
            try:
                cfg = ensure_new_schema(json.loads(raw))
            except Exception as e:
                QMessageBox.warning(win, "重载失败", f"JSON 解析失败：\n{e}")
                return
        QMessageBox.information(win, "已重载", "配置已重载，建议关闭后重新打开")

    btn_save.clicked.connect(_on_save)
    btn_close.clicked.connect(win.close)
    btn_reload.clicked.connect(_on_reload)

    _render_providers()
    win.show()
    return win
