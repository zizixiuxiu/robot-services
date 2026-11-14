#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI 接入配置服务（默认 8020）。

只做一件事：把「用哪家模型、key 是什么」变成可运行时修改的配置，
以后换 key / 换模型不用再改 compose、不用重启业务服务。

  GET  /health                 健康检查（不暴露 key）
  GET  /ai-config              读配置（key 默认脱敏；reveal=1 + token 可看明文）
  PUT  /ai-config              改配置（支持部分更新，写盘即生效）
  POST /ai-config/test         用指定 provider 发一条最小请求，验证 key 是否可用

配置落在 AI_CONFIG_FILE（默认 /data/ai_config.json），业务服务直接读同一个文件，
所以本服务挂了也不影响业务，只是改不了配置而已。

支持两种协议：
  openai     POST {base_url}/chat/completions   Authorization: Bearer <key>
             （DeepSeek、通义、Moonshot 等 OpenAI 兼容端点都走这条）
  anthropic  POST {base_url}/v1/messages        x-api-key: <key>
             + anthropic-version: 2023-06-01

错误码：
  AICFG-E001 参数不合法
  AICFG-E002 未知 provider
  AICFG-E003 未授权（token 不对 / 非本机访问）
  AICFG-E004 写配置失败
  AICFG-E101 该 provider 没配 key
  AICFG-E102 上游返回非 2xx
  AICFG-E103 上游连接失败/超时
  AICFG-E104 上游响应无法解析
"""
from __future__ import annotations

import json
import os
import socket
import sys
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

DEFAULT_PORT = int(os.getenv("AI_CONFIG_PORT", "8020"))
CONFIG_FILE = Path(os.getenv("AI_CONFIG_FILE", "/data/ai_config.json"))
ADMIN_TOKEN = os.getenv("AI_CONFIG_TOKEN", "").strip()
LOG_FILE = os.getenv("AI_CONFIG_LOG", "")

PROVIDERS = ("openai", "anthropic")

DEFAULT_CONFIG = {
    "active": "openai",
    "providers": {
        "openai": {
            "api_key": "",
            "base_url": "https://api.openai.com/v1",
            "model": "gpt-4o-mini",
        },
        "anthropic": {
            "api_key": "",
            "base_url": "https://api.anthropic.com",
            "model": "claude-sonnet-4-5-20250929",
        },
    },
    "options": {"temperature": 0.4, "max_tokens": 6000, "timeout": 120},
    "updated_at": "",
}

_lock = threading.Lock()

# 容器默认 UTC，日志/时间戳要和群里、和业务侧对得上，统一按北京时间
if not os.getenv("TZ"):
    os.environ["TZ"] = "Asia/Shanghai"
try:
    time.tzset()
except AttributeError:  # Windows 本机调试没有 tzset
    pass


# ---------------------------------------------------------------------------
# 日志：带时间戳打到 stdout + 可选文件（容器里 docker logs 就够，文件是给 grep 用）
# ---------------------------------------------------------------------------
def log(level: str, msg: str, **fields) -> None:
    extra = " ".join(f"{k}={v}" for k, v in fields.items())
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} [{level}] {msg}" + (f" | {extra}" if extra else "")
    print(line, flush=True)
    if LOG_FILE:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        except OSError:
            pass


def _deep_merge(base: dict, patch: dict) -> dict:
    out = dict(base)
    for key, value in patch.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _deep_merge(out[key], value)
        else:
            out[key] = value
    return out


def load_config(apply_env: bool = True) -> dict:
    """读配置；文件不存在/损坏时返回默认值（不写盘）。

    apply_env=False 用于"准备落盘"的场景：env 兜底来的 key 不写进文件。
    """
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if CONFIG_FILE.exists():
        try:
            cfg = _deep_merge(cfg, json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError) as e:
            log("WARN", "配置文件损坏，改用默认值", file=str(CONFIG_FILE), error=repr(e))
    # 环境变量兜底：文件里没 key 时用 env（不写盘，标记来源）
    if apply_env and not cfg["providers"]["openai"].get("api_key") and os.getenv("AI_API_KEY"):
        cfg["providers"]["openai"]["api_key"] = os.getenv("AI_API_KEY", "")
        cfg["providers"]["openai"]["_source"] = "env:AI_API_KEY"
    return cfg


def save_config(cfg: dict) -> None:
    """原子写：先写临时文件再 replace，避免读到半个文件。"""
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    cfg["updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    tmp = CONFIG_FILE.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    # 0644 而不是 0600：业务容器以非 root 用户运行，只读挂载也要读得动
    try:
        os.chmod(tmp, 0o644)
    except OSError:
        pass
    os.replace(tmp, CONFIG_FILE)


def mask_key(key: str) -> str:
    if not key:
        return ""
    if len(key) <= 10:
        return key[:2] + "***"
    return f"{key[:5]}***{key[-4:]}"


def public_view(cfg: dict, reveal: bool = False) -> dict:
    view = json.loads(json.dumps(cfg))
    for name, prov in view.get("providers", {}).items():
        raw = prov.get("api_key", "")
        prov["api_key"] = raw if reveal else mask_key(raw)
        prov["configured"] = bool(raw)
    view["configured"] = bool(cfg.get("providers", {}).get(cfg.get("active", ""), {}).get("api_key"))
    return view


# ---------------------------------------------------------------------------
# 上游调用：两种协议各一个分支，试配置时跟业务侧走同一套写法
# ---------------------------------------------------------------------------
def _endpoint(base_url: str, provider: str) -> str:
    base = (base_url or "").rstrip("/")
    if provider == "anthropic":
        if base.endswith("/v1"):
            return base + "/messages"
        return base + "/v1/messages"
    return base + "/chat/completions"


def call_provider(provider: str, prov: dict, prompt: str, max_tokens: int = 32,
                  temperature: float = 0.0, timeout: int = 60) -> tuple[bool, str, int]:
    """返回 (成功, 文本或错误信息, 耗时毫秒)。"""
    api_key = (prov.get("api_key") or "").strip()
    if not api_key:
        return False, "AICFG-E101 该 provider 还没配 api_key", 0

    model = (prov.get("model") or "").strip()
    url = _endpoint(prov.get("base_url") or "", provider)
    if provider == "anthropic":
        body = {
            "model": model,
            "max_tokens": max_tokens,
            "temperature": temperature,
            "messages": [{"role": "user", "content": prompt}],
        }
        headers = {
            "Content-Type": "application/json; charset=utf-8",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        }
    else:
        body = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        headers = {
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": f"Bearer {api_key}",
        }

    started = time.time()
    req = urllib.request.Request(url, data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                                 headers=headers, method="POST")
    # 本地服务/上游都不该走系统代理
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    try:
        with opener.open(req, timeout=timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            raw = e.read().decode("utf-8", "replace")
            err = json.loads(raw)
            detail = (err.get("error", {}).get("message") if isinstance(err.get("error"), dict)
                      else err.get("error")) or raw
        except Exception:
            detail = ""
        err_msg = f"AICFG-E102 上游 HTTP {e.code}：{str(detail)[:300]}"
        log("ERROR", "上游返回错误", provider=provider, model=model, url=url, http=e.code, detail=str(detail)[:300])
        return False, err_msg, int((time.time() - started) * 1000)
    except (urllib.error.URLError, TimeoutError, socket.timeout) as e:
        msg = f"AICFG-E103 连不上上游（{getattr(e, 'reason', e)}）"
        log("ERROR", "上游连接失败", provider=provider, url=url, error=repr(e))
        return False, msg, int((time.time() - started) * 1000)
    except json.JSONDecodeError as e:
        log("ERROR", "上游响应无法解析", provider=provider, url=url, error=repr(e))
        return False, "AICFG-E104 上游响应不是合法 JSON", int((time.time() - started) * 1000)

    try:
        if provider == "anthropic":
            text = "".join(block.get("text", "") for block in payload.get("content", [])
                           if isinstance(block, dict))
        else:
            text = payload["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError, AttributeError):
        log("ERROR", "响应结构不符", provider=provider, url=url, keys=list(payload)[:8])
        return False, "AICFG-E104 上游响应结构不符合该协议", int((time.time() - started) * 1000)

    return True, text, int((time.time() - started) * 1000)


# ---------------------------------------------------------------------------
# 参数校验
# ---------------------------------------------------------------------------
def validate(cfg: dict) -> str:
    if cfg.get("active") not in PROVIDERS:
        return f"AICFG-E002 active 只能是 {'/'.join(PROVIDERS)}"
    for name, prov in cfg.get("providers", {}).items():
        if name not in PROVIDERS:
            return f"AICFG-E002 不支持的 provider：{name}"
        base = str(prov.get("base_url") or "")
        if not base.startswith(("http://", "https://")):
            return f"AICFG-E001 {name}.base_url 必须是 http(s) 地址"
        if not str(prov.get("model") or "").strip():
            return f"AICFG-E001 {name}.model 不能为空"
        if prov.get("api_key") is None:
            return f"AICFG-E001 {name}.api_key 不能是 null"
    opts = cfg.get("options", {})
    for field, low, high in (("temperature", 0, 2), ("max_tokens", 1, 200000), ("timeout", 1, 600)):
        if field in opts:
            try:
                value = float(opts[field])
            except (TypeError, ValueError):
                return f"AICFG-E001 options.{field} 必须是数字"
            if not low <= value <= high:
                return f"AICFG-E001 options.{field} 超出范围 {low}-{high}"
    return ""


def apply_patch(patch: dict) -> dict:
    """把请求体并到现有配置上。支持 {providers:{...}, active:...} 和 {openai:{...}} 两种写法。"""
    cfg = load_config(apply_env=False)
    if not isinstance(patch, dict):
        raise ValueError("AICFG-E001 请求体必须是 JSON 对象")

    shorthand = {k: v for k, v in patch.items() if k in PROVIDERS}
    if shorthand:
        patch = {**{k: v for k, v in patch.items() if k not in PROVIDERS},
                 "providers": {**patch.get("providers", {}), **shorthand}}
    if "providers" in patch and not isinstance(patch["providers"], dict):
        raise ValueError("AICFG-E001 providers 必须是对象")

    # GET 返回的是脱敏 key（sk-ab***wxyz），把这些值回传时视为"不改 key"，
    # 否则一次手抖就把真 key 覆盖成星号。
    for prov in (patch.get("providers") or {}).values():
        if isinstance(prov, dict) and "***" in str(prov.get("api_key") or ""):
            prov.pop("api_key", None)

    cfg = _deep_merge(cfg, patch)
    return cfg


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "ai-config/1.0"

    def log_message(self, fmt, *args):  # 静默默认访问日志，自己打结构化日志
        return

    def _send(self, code: int, payload: dict) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _fail(self, code: int, err_code: str, message: str) -> None:
        self._send(code, {"ok": False, "code": err_code, "error": message})

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if not length:
            return {}
        raw = self.rfile.read(length)
        try:
            body = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as e:
            raise ValueError(f"AICFG-E001 请求体不是合法 JSON：{e}")
        if not isinstance(body, dict):
            raise ValueError("AICFG-E001 请求体必须是 JSON 对象")
        return body

    def _authorized(self) -> bool:
        """写操作：配了 AI_CONFIG_TOKEN 就必须带对，没配就放行（内网服务，启动时会告警）。"""
        if ADMIN_TOKEN:
            return (self.headers.get("X-Admin-Token") or "").strip() == ADMIN_TOKEN
        return True

    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/health":
            cfg = load_config()
            active = cfg.get("active")
            prov = cfg.get("providers", {}).get(active, {})
            self._send(200, {"ok": True, "service": "ai-config", "active": active,
                             "model": prov.get("model", ""), "configured": bool(prov.get("api_key")),
                             "file": str(CONFIG_FILE)})
            return
        if path == "/ai-config":
            reveal = "reveal=1" in (self.path.split("?", 1) + [""])[1]
            if reveal and not self._authorized():
                self._fail(403, "AICFG-E003", "看明文 key 需要 X-Admin-Token（或从本机访问）")
                return
            self._send(200, {"ok": True, "config": public_view(load_config(), reveal)})
            return
        self._fail(404, "AICFG-E001", f"未知接口 {path}")

    def do_PUT(self):
        path = self.path.split("?")[0]
        if path != "/ai-config":
            self._fail(404, "AICFG-E001", f"未知接口 {path}")
            return
        if not self._authorized():
            self._fail(403, "AICFG-E003", "改配置需要 X-Admin-Token（或从本机访问）")
            return
        try:
            patch = self._read_json()
            with _lock:
                cfg = apply_patch(patch)
                error = validate(cfg)
                if error:
                    self._fail(400, error.split()[0], error.split(" ", 1)[-1])
                    return
                save_config(cfg)
        except ValueError as e:
            code, _, msg = str(e).partition(" ")
            self._fail(400, code, msg or str(e))
            return
        except OSError as e:
            log("ERROR", "写配置失败", file=str(CONFIG_FILE), error=repr(e))
            self._fail(500, "AICFG-E004", f"写配置失败：{e}")
            return

        view = public_view(load_config())
        log("INFO", "配置已更新", active=view.get("active"),
            openai_model=view["providers"]["openai"]["model"],
            anthropic_model=view["providers"]["anthropic"]["model"])
        self._send(200, {"ok": True, "config": view})

    def do_POST(self):
        path = self.path.split("?")[0]
        if path != "/ai-config/test":
            self._fail(404, "AICFG-E001", f"未知接口 {path}")
            return
        if not self._authorized():
            self._fail(403, "AICFG-E003", "试配置需要 X-Admin-Token（或从本机访问）")
            return
        try:
            body = self._read_json()
        except ValueError as e:
            code, _, msg = str(e).partition(" ")
            self._fail(400, code, msg or str(e))
            return

        cfg = load_config()
        provider = body.get("provider") or cfg.get("active")
        if provider not in PROVIDERS:
            self._fail(400, "AICFG-E002", f"provider 只能是 {'/'.join(PROVIDERS)}")
            return
        prov = dict(cfg.get("providers", {}).get(provider, {}))
        prov.update(body.get("override") or {})  # 允许先用未保存的 key 试
        prompt = body.get("prompt") or "只回复 OK 两个字母，不要别的"
        ok, text, ms = call_provider(provider, prov, prompt,
                                     max_tokens=int(body.get("max_tokens") or 32),
                                     timeout=int(cfg.get("options", {}).get("timeout") or 60))
        log("INFO" if ok else "ERROR", "试配置完成", provider=provider,
            model=prov.get("model"), ok=ok, latency_ms=ms)
        self._send(200 if ok else 502, {"ok": ok, "provider": provider,
                                        "model": prov.get("model"), "latency_ms": ms,
                                        "sample": text[:200] if ok else "", "error": "" if ok else text})


def main() -> None:
    if not CONFIG_FILE.exists():
        try:
            save_config(json.loads(json.dumps(DEFAULT_CONFIG)))
            log("INFO", "配置文件不存在，已生成模板", file=str(CONFIG_FILE))
        except OSError as e:
            log("ERROR", "初始化配置文件失败", file=str(CONFIG_FILE), error=repr(e))

    server = ThreadingHTTPServer(("0.0.0.0", DEFAULT_PORT), Handler)
    cfg = load_config()
    log("INFO", f"ai-config 已启动 http://0.0.0.0:{DEFAULT_PORT}",
        file=str(CONFIG_FILE), active=cfg.get("active"),
        auth="token" if ADMIN_TOKEN else "OPEN(未设置 AI_CONFIG_TOKEN)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        log("INFO", "收到退出信号")
        server.shutdown()


if __name__ == "__main__":
    sys.exit(main())
