# coding=utf-8
"""LLM client helpers (OpenAI-compatible).

This module centralizes how we create OpenAI SDK clients for different tasks.
We treat providers as OpenAI-compatible endpoints:
- base_url: e.g. https://api.openai.com/v1 or https://api.deepseek.com/v1
- api_key: sent as Bearer token by the OpenAI Python SDK

Config structure (new):
{
  "LLM": {
    "PROVIDERS": {
      "openai": {"BASE_URL": "https://api.openai.com/v1", "API_KEY": "...", "MODEL": "gpt-4o-mini"},
      "deepseek": {"BASE_URL": "https://api.deepseek.com/v1", "API_KEY": "...", "MODEL": "deepseek-chat"}
    },
    "TASKS": {
      "typo_fix": {"ENABLED": false, "PROVIDER": "deepseek", "PROMPT": "...{text}"},
      "editor": {"ENABLED": false, "PROVIDER": "deepseek", "PROMPT": "...{text}"}
    }
  }
}

Legacy config (old) will be mapped by the loaders in GUI/CLI.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple

from openai import OpenAI


@dataclass
class ProviderConfig:
    name: str
    base_url: str
    api_key: str
    model: str


@dataclass
class TaskConfig:
    name: str
    enabled: bool
    provider: str
    prompt: str


def _norm_base_url(url: str) -> str:
    if not url:
        return url
    # OpenAI SDK expects base_url like "https://.../v1" (no trailing slash preferred)
    return url.rstrip("/")


def get_provider_config(config: Dict[str, Any], provider_name: str) -> ProviderConfig:
    llm = config.get("LLM", {})
    providers = llm.get("PROVIDERS", {})
    p = providers.get(provider_name) or {}

    base_url = _norm_base_url(p.get("BASE_URL") or "")
    api_key = (p.get("API_KEY") or "").strip()
    model = (p.get("MODEL") or "").strip()

    return ProviderConfig(name=provider_name, base_url=base_url, api_key=api_key, model=model)


def get_task_config(config: Dict[str, Any], task_name: str) -> TaskConfig:
    llm = config.get("LLM", {})
    tasks = llm.get("TASKS", {})
    t = tasks.get(task_name) or {}
    return TaskConfig(
        name=task_name,
        enabled=bool(t.get("ENABLED", False)),
        provider=(t.get("PROVIDER") or "").strip(),
        prompt=(t.get("PROMPT") or "").strip(),
    )


def make_client(provider: ProviderConfig) -> OpenAI:
    if not provider.api_key:
        raise RuntimeError(f"未配置 API Key（provider={provider.name}）")
    if provider.base_url:
        return OpenAI(api_key=provider.api_key, base_url=provider.base_url)
    # If base_url empty, OpenAI SDK uses default (OpenAI).
    return OpenAI(api_key=provider.api_key)


def resolve_task_client(config: Dict[str, Any], task_name: str) -> Tuple[OpenAI, str, str]:
    """Return (client, model, prompt) for a given task."""
    task = get_task_config(config, task_name)
    if not task.provider:
        raise RuntimeError(f"LLM 任务未指定 provider：{task_name}")

    provider = get_provider_config(config, task.provider)
    if not provider.model:
        raise RuntimeError(f"provider 未配置 MODEL：{provider.name}")

    client = make_client(provider)
    return client, provider.model, task.prompt
