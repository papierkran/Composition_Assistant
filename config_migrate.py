# coding=utf-8
"""Config migration / compatibility helpers.

Goal:
- Prefer new config schema (LLM.PROVIDERS + LLM.TASKS, OCR.XFYUN)
- If user still has legacy schema (OCR + DEEPSEEK + EDITOR...), map it into
  the new structure in-memory so code paths can be unified.

We DO NOT automatically overwrite the on-disk config unless caller chooses to.
"""

from __future__ import annotations

from typing import Any, Dict


def ensure_new_schema(cfg: Dict[str, Any]) -> Dict[str, Any]:
    if cfg is None:
        cfg = {}

    # If it already looks like new schema, keep it.
    if "LLM" in cfg and isinstance(cfg.get("LLM"), dict):
        # still ensure some defaults exist
        cfg.setdefault("LLM", {}).setdefault("PROVIDERS", {})
        cfg.setdefault("LLM", {}).setdefault("TASKS", {})
        return cfg

    # Legacy -> new mapping
    new_cfg: Dict[str, Any] = {}

    # APP
    new_cfg["APP"] = cfg.get("APP", {}) if isinstance(cfg.get("APP"), dict) else {}

    # OCR
    ocr = cfg.get("OCR", {}) if isinstance(cfg.get("OCR"), dict) else {}
    new_cfg["OCR"] = {
        "PROVIDER": "xfyun_handwriting",
        "XFYUN": {
            "URL": ocr.get("URL", "http://webapi.xfyun.cn/v1/service/v1/ocr/handwriting"),
            "APPID": ocr.get("APPID", ""),
            "API_KEY": ocr.get("API_KEY", ""),
            "LANGUAGE": ocr.get("LANGUAGE", "cn|en"),
            "LOCATION": ocr.get("LOCATION", "false"),
        },
    }

    # LLM providers/tasks
    deepseek = cfg.get("DEEPSEEK", {}) if isinstance(cfg.get("DEEPSEEK"), dict) else {}
    editor = cfg.get("EDITOR", {}) if isinstance(cfg.get("EDITOR"), dict) else {}

    # Legacy DEEPSEEK had no BASE_URL in the small config.json snapshot; assume official.
    deepseek_base = deepseek.get("BASE_URL") or "https://api.deepseek.com/v1"

    new_cfg["LLM"] = {
        "PROVIDERS": {
            "deepseek": {
                "BASE_URL": deepseek_base,
                "API_KEY": deepseek.get("API_KEY", ""),
                "MODEL": deepseek.get("MODEL", "deepseek-chat"),
            },
            "openai": {
                "BASE_URL": "https://api.openai.com/v1",
                "API_KEY": "",
                "MODEL": "gpt-4o-mini",
            },
            "custom": {
                "BASE_URL": "",
                "API_KEY": "",
                "MODEL": "",
            },
        },
        "TASKS": {
            "typo_fix": {
                "ENABLED": bool(deepseek.get("ENABLED", False)),
                "PROVIDER": "deepseek",
                "PROMPT": deepseek.get("PROMPT", "{text}"),
            },
            "editor": {
                "ENABLED": bool(editor.get("ENABLED", False)),
                "PROVIDER": "deepseek",
                "PROMPT": editor.get("PROMPT", "{text}"),
            },
        },
    }

    return new_cfg
