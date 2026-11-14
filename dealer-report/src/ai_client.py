#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI 接入客户端：统一读 ai-config 的配置文件，按 provider 走对应协议。

配置来源优先级：
  1. 环境变量 AI_CONFIG_FILE 指定的 JSON（容器里是这个）
  2. 仓库里 ai-config/config/ai_config.json（本机调试）
  3. 老式环境变量 AI_API_KEY / AI_BASE_URL / AI_MODEL（兜底，兼容未迁移的部署）

配置文件改了立刻生效（按 mtime 判断，不必重启服务）。

对外只暴露三个函数：
  status()  -> 当前生效的 provider/model/来源，给健康检查、capability 用
  chat(prompt, ...) -> 文本；任何失败都返回 ''，让调用方走本地规则兜底
  describe_error() -> 最近一次失败的简短原因（带错误码），排查用

协议：
  openai     {base_url}/chat/completions   Authorization: Bearer <key>
  anthropic  {base_url}/v1/messages        x-api-key + anthropic-version
"""
from __future__ import annotations

import json
import os
import socket
import time
import urllib.error
import urllib.request
from pathlib import Path

PROVIDERS = ("openai", "anthropic")
_CACHE = {"mtime": None, "config": None, "path": None}
_LAST_ERROR = ""


def _candidate_paths() -> list[Path]:
    paths = []
    env_path = os.getenv("AI_CONFIG_FILE")
    if env_path:
        paths.append(Path(env_path))
    here = Path(__file__).resolve().parent          # .../dealer-report/src
    repo = here.parent                              # .../dealer-report
    # 平台侧：配置跟着代码目录走（services/dealer_report 已只读挂进容器，不用再加卷）
    paths.append(here / "ai_config.json")
    paths.append(repo / "config" / "ai_config.json")
    paths.append(repo.parent / "ai-config" / "config" / "ai_config.json")
    return paths


def _env_fallback() -> dict:
    """老部署只有一个 DeepSeek（OpenAI 兼容）key 时，仍然能跑。"""
    return {
        "active": "openai",
        "providers": {
            "openai": {
                "api_key": os.getenv("AI_API_KEY", ""),
                "base_url": os.getenv("AI_BASE_URL", "https://api.deepseek.com/v1"),
                "model": os.getenv("AI_MODEL", "deepseek-v4-flash"),
            },
            "anthropic": {"api_key": "", "base_url": "https://api.anthropic.com", "model": ""},
        },
        "options": {"temperature": 0.4, "max_tokens": 6000, "timeout": 120},
        "_source": "env",
    }


def load_config() -> dict:
    """读配置，带 mtime 缓存；文件不存在时退回环境变量。"""
    for path in _candidate_paths():
        try:
            stat = path.stat()
        except OSError:
            continue
        if _CACHE["path"] == str(path) and _CACHE["mtime"] == stat.st_mtime:
            return _CACHE["config"]
        try:
            cfg = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            print(f"[ai_client] 配置读取失败 {path}: {e}", flush=True)
            continue
        cfg["_source"] = "file"
        cfg["_file"] = str(path)
        _CACHE.update({"mtime": stat.st_mtime, "config": cfg, "path": str(path)})
        return cfg
    return _env_fallback()


def status() -> dict:
    """给健康检查 / capability 用：不返回 key，只说有没有配。"""
    cfg = load_config()
    active = cfg.get("active", "")
    prov = (cfg.get("providers") or {}).get(active, {})
    return {
        "provider": active,
        "model": prov.get("model", ""),
        "base_url": prov.get("base_url", ""),
        "configured": bool(str(prov.get("api_key") or "").strip()),
        "source": cfg.get("_source", "unknown"),
        "file": cfg.get("_file", ""),
    }


def summary() -> dict:
    """平台 capability 用：把 status() 翻译成平台约定的 ai 字段。"""
    st = status()
    return {
        "config_key_present": st["configured"],
        "mode": st["provider"] if st["configured"] else "local-rule-fallback",
        "model": st["model"],
        "note": (f"AI 配置来自 ai-config（{st['source']}）：{st['provider']}/{st['model']}"
                 if st["configured"] else "无 key，分析页走本地规则文本"),
    }


def describe_error() -> str:
    return _LAST_ERROR


def _endpoint(base_url: str, provider: str) -> str:
    base = (base_url or "").strip().rstrip("/")
    if provider == "anthropic":
        return base + "/messages" if base.endswith("/v1") else base + "/v1/messages"
    return base + "/chat/completions"


def _build_request(provider: str, prov: dict, prompt: str,
                   temperature: float, max_tokens: int):
    api_key = str(prov.get("api_key") or "").strip()
    url = _endpoint(prov.get("base_url") or "", provider)
    model = str(prov.get("model") or "").strip()
    if provider == "anthropic":
        payload = {"model": model, "max_tokens": max_tokens, "temperature": temperature,
                   "messages": [{"role": "user", "content": prompt}]}
        headers = {"Content-Type": "application/json; charset=utf-8",
                   "x-api-key": api_key, "anthropic-version": "2023-06-01"}
    else:
        payload = {"model": model, "messages": [{"role": "user", "content": prompt}],
                   "temperature": temperature, "max_tokens": max_tokens}
        headers = {"Content-Type": "application/json; charset=utf-8",
                   "Authorization": f"Bearer {api_key}"}
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    return urllib.request.Request(url, data=body, headers=headers, method="POST")


def _extract_text(provider: str, payload: dict) -> str:
    if provider == "anthropic":
        return "".join(b.get("text", "") for b in payload.get("content", []) if isinstance(b, dict))
    return payload["choices"][0]["message"]["content"]


def chat(prompt: str, temperature: float = None, max_tokens: int = None,
         timeout: int = None) -> str:
    """按当前生效配置请求模型，返回文本；没配 key 或调用失败返回 ''。"""
    global _LAST_ERROR
    cfg = load_config()
    active = cfg.get("active", "openai")
    prov = dict((cfg.get("providers") or {}).get(active, {}))
    options = cfg.get("options") or {}

    if not str(prov.get("api_key") or "").strip():
        _LAST_ERROR = f"AICFG-E101 {active} 没有配置 api_key"
        return ""

    temperature = options.get("temperature", 0.4) if temperature is None else temperature
    max_tokens = int(options.get("max_tokens", 6000)) if max_tokens is None else int(max_tokens)
    timeout = int(options.get("timeout", 120)) if timeout is None else int(timeout)

    started = time.time()
    req = _build_request(active, prov, prompt, float(temperature), max_tokens)
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    try:
        with opener.open(req, timeout=timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            raw = e.read().decode("utf-8", "replace")
            parsed = json.loads(raw)
            err = parsed.get("error")
            detail = err.get("message") if isinstance(err, dict) else (str(err) if err else raw)
        except Exception:
            pass
        _LAST_ERROR = f"AICFG-E102 HTTP {e.code}：{str(detail)[:200]}"
        print(f"[ai_client] {active} 调用失败 {_LAST_ERROR}", flush=True)
        return ""
    except (urllib.error.URLError, TimeoutError, socket.timeout) as e:
        _LAST_ERROR = f"AICFG-E103 连不上上游（{getattr(e, 'reason', e)}）"
        print(f"[ai_client] {active} 调用失败 {_LAST_ERROR}", flush=True)
        return ""
    except json.JSONDecodeError as e:
        _LAST_ERROR = f"AICFG-E104 响应不是合法 JSON（{e}）"
        print(f"[ai_client] {active} 调用失败 {_LAST_ERROR}", flush=True)
        return ""

    try:
        text = _extract_text(active, payload)
    except (KeyError, IndexError, TypeError, AttributeError):
        _LAST_ERROR = "AICFG-E104 响应结构不符合该协议"
        print(f"[ai_client] {active} 调用失败 {_LAST_ERROR} keys={list(payload)[:8]}", flush=True)
        return ""

    _LAST_ERROR = ""
    print(f"[ai_client] {active}/{prov.get('model')} 调用成功 "
          f"{int((time.time() - started) * 1000)}ms，返回 {len(text)} 字", flush=True)
    return text
