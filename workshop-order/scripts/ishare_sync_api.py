# -*- coding: utf-8 -*-
"""iShare 同步触发端口（8016）—— 代替原「每 3 秒自动轮询」的常驻同步进程。

谁录完单，通知后调用一次即可：
    curl -X POST http://<本机IP>:8016/sync
    curl http://<本机IP>:8016/sync          # 浏览器直接开也行

行为：通过 ssh 在服务器上 systemctl start ishare-backup.service（共享盘 /mnt/ishare
-> /home/ops/ishare_backup，只增不删），跑完把 last_run.json 原样返回，
包含 copied / skipped / errors / cost_s / new_files。

其它端点：
    GET  /            老接口兼容，返回 ok
    GET  /healthz     健康 + 上次运行摘要（不触发）
    GET  /status      只读上次运行结果（不触发）
"""
from __future__ import annotations

import json
import os
import subprocess
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

SSH = os.getenv(
    "ISHARE_SSH_BIN",
    r"C:\Users\AI_ProgramFiles\OpenClaw\WorkBuddy\resources\vendor\PortableGit\usr\bin\ssh.exe",
)
SSH_HOST = os.getenv("ISHARE_SSH_HOST", "ops@192.168.6.245")
SSH_PORT = os.getenv("ISHARE_SSH_PORT", "2222")
LISTEN = os.getenv("ISHARE_SYNC_LISTEN", "0.0.0.0:8016")
REMOTE_LAST_RUN = "/home/ops/ishare_backup/last_run.json"
SYNC_UNIT = "ishare-backup.service"
TOKEN = os.getenv("ISHARE_SYNC_TOKEN", "").strip()
LOG_PATH = Path(os.getenv(
    "ISHARE_SYNC_API_LOG",
    r"D:\Services\robot-services\workshop-order\logs\ishare_trigger_8016.log",
))
TIMEOUT_S = float(os.getenv("ISHARE_SYNC_TIMEOUT", "300"))

_lock = threading.Lock()


def log(msg: str, **kw) -> None:
    line = json.dumps({"ts": datetime.now().isoformat(timespec="seconds"),
                       "msg": msg, **kw}, ensure_ascii=False)
    print(line, flush=True)
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def _ssh(cmd: str) -> tuple[int, str, str]:
    p = subprocess.run(
        [SSH, "-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=no",
         "-o", "ConnectTimeout=8", "-p", SSH_PORT, SSH_HOST, cmd],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=TIMEOUT_S,
    )
    return p.returncode, p.stdout or "", p.stderr or ""


def _parse_last_run(text: str) -> dict | None:
    for line in reversed(text.strip().splitlines()):
        line = line.strip()
        if line.startswith("{") and line.endswith("}"):
            try:
                return json.loads(line)
            except Exception:  # noqa: BLE001
                continue
    return None


def read_status() -> dict:
    try:
        rc, out, err = _ssh(f"cat {REMOTE_LAST_RUN}")
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "ssh 超时"}
    if rc != 0:
        return {"ok": False, "error": (err or out).strip()[:300] or f"ssh rc={rc}"}
    data = _parse_last_run(out)
    return data or {"ok": False, "error": "last_run.json 无法解析"}


def run_sync() -> dict:
    """触发一次同步。并发点两次时第二个直接返回 busy，不排队。"""
    if not _lock.acquire(blocking=False):
        return {"ok": False, "error": "busy", "detail": "上一次同步还在跑"}
    t0 = time.time()
    try:
        cmd = f"sudo -n systemctl start {SYNC_UNIT} && cat {REMOTE_LAST_RUN}"
        rc, out, err = _ssh(cmd)
        cost = round(time.time() - t0, 2)
        if rc != 0:
            log("sync_failed", rc=rc, err=err[:300], cost_s=cost)
            return {"ok": False, "error": "触发失败",
                    "detail": (err or out).strip()[:500], "cost_s": cost}
        data = _parse_last_run(out) or {"ok": False, "error": "无 last_run.json"}
        data["cost_s"] = cost
        log("sync_done", copied=data.get("copied"), skipped=data.get("skipped"),
            errors=data.get("errors"), cost_s=cost,
            files=(data.get("new_files") or [])[:10])
        return data
    except subprocess.TimeoutExpired:
        log("sync_timeout", timeout_s=TIMEOUT_S)
        return {"ok": False, "error": f"ssh 超时（>{TIMEOUT_S}s）"}
    finally:
        _lock.release()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):  # 静默默认 stderr 日志
        pass

    def _send(self, code: int, payload) -> None:
        body = payload if isinstance(payload, bytes) else \
            json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8" if not isinstance(payload, bytes) else "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _authorized(self) -> bool:
        if not TOKEN:
            return True
        if self.headers.get("X-Token") == TOKEN:
            return True
        return f"token={TOKEN}" in (self.path or "")

    def do_GET(self):  # noqa: N802
        path = self.path.split("?")[0].rstrip("/") or "/"
        if path in ("/", "/healthz"):
            st = read_status()
            self._send(200, {"ok": True, "service": "ishare-sync-8016",
                             "target": f"{SSH_HOST}:{SYNC_UNIT}", "last_run": st})
            return
        if path == "/status":
            self._send(200, read_status())
            return
        if path == "/sync":
            if not self._authorized():
                self._send(403, {"ok": False, "error": "forbidden"})
                return
            self._send(200, run_sync())
            return
        self._send(404, {"ok": False, "error": "not_found"})

    def do_POST(self):  # noqa: N802
        path = self.path.split("?")[0].rstrip("/") or "/"
        n = int(self.headers.get("Content-Length") or 0)
        if n:
            try:
                self.rfile.read(n)
            except Exception:  # noqa: BLE001
                pass
        if path in ("/sync", "/"):
            if not self._authorized():
                self._send(403, {"ok": False, "error": "forbidden"})
                return
            self._send(200, run_sync())
            return
        self._send(404, {"ok": False, "error": "not_found"})


def main() -> None:
    host, port_s = LISTEN.split(":")
    srv = ThreadingHTTPServer((host, int(port_s)), Handler)
    log("api_up", listen=LISTEN, target=f"{SSH_HOST}:{SYNC_UNIT}",
        token=bool(TOKEN), timeout_s=TIMEOUT_S)
    srv.serve_forever()


if __name__ == "__main__":
    main()
