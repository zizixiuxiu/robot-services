#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 重写服务（宿主机，端口 8017）。

背景：pvc-classify 在 Linux 容器内用 xlwt 生成的 xls，南兴软件不认可；
经 Excel/WPS 打开另存（重写 BIFF 结构）后可正常导入。容器内无法跑 COM，
因此本服务运行在 Windows 宿主机上，通过 PowerShell + Excel COM 完成重写。

接口：
    GET  /health
    POST /resave   {"dir": "D:\\...\\data\\resave_exchange\\<uuid>"}
                   重写目录下所有 .xls（同路径覆盖），返回 {"success": true, "count": N}

安全：仅允许重写 EXCHANGE_ROOT 目录下的文件。
"""

import json
import os
import subprocess
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="ignore")

PORT = int(os.environ.get("RESAVE_PORT", "8017"))
EXCHANGE_ROOT = Path(r"D:\Services\robot-services\pvc-classify\data\resave_exchange").resolve()

_LOCK = threading.Lock()  # Excel COM 串行执行

_PS_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
$dir = $args[0]
$files = Get-ChildItem -Path $dir -Filter *.xls -File
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$count = 0
try {
  foreach ($f in $files) {
    $wb = $excel.Workbooks.Open($f.FullName)
    $tmp = $f.FullName + '.resave_tmp'
    $wb.SaveAs($tmp, 56)   # 56 = xlExcel8 (.xls)
    $wb.Close($false)
    Move-Item -Force $tmp $f.FullName
    $count++
  }
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output "REAVED_COUNT=$count"
"""


def resave_dir(dir_path: Path) -> int:
    """用 Excel COM 重写目录下所有 .xls，返回处理数量。"""
    dir_path = dir_path.resolve()
    if EXCHANGE_ROOT not in dir_path.parents:
        raise ValueError(f"目录不在允许范围内: {dir_path}")
    if not dir_path.is_dir():
        raise ValueError(f"目录不存在: {dir_path}")

    with _LOCK:
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".ps1", delete=False, encoding="gbk"
        ) as f:
            f.write(_PS_TEMPLATE)
            ps1 = f.name
        try:
            result = subprocess.run(
                [
                    "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                    "-File", ps1, str(dir_path),
                ],
                capture_output=True, text=True, encoding="gbk", errors="replace",
                timeout=300,
            )
        finally:
            os.unlink(ps1)

    if result.returncode != 0:
        raise RuntimeError(f"Excel 重写失败: {result.stderr.strip()[:500]}")
    for line in result.stdout.splitlines():
        if line.startswith("REAVED_COUNT="):
            return int(line.split("=", 1)[1])
    raise RuntimeError(f"Excel 重写无输出: stdout={result.stdout[:200]}")


class Handler(BaseHTTPRequestHandler):
    def _reply(self, code: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/health":
            self._reply(200, {"status": "ok", "service": "excel-resave", "port": PORT})
        else:
            self._reply(404, {"error": "not found"})

    def do_POST(self):
        if self.path != "/resave":
            self._reply(404, {"error": "not found"})
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length).decode("utf-8"))
            dir_arg = body.get("dir", "")
            if not dir_arg:
                self._reply(400, {"success": False, "error": "缺少 dir 参数"})
                return
            count = resave_dir(Path(dir_arg))
            print(f"[resave] {dir_arg} -> {count} 个文件")
            self._reply(200, {"success": True, "count": count})
        except Exception as e:
            print(f"[resave] 失败: {e}", file=sys.stderr)
            self._reply(500, {"success": False, "error": str(e)})

    def log_message(self, *args):
        pass


if __name__ == "__main__":
    EXCHANGE_ROOT.mkdir(parents=True, exist_ok=True)
    print(f"Excel 重写服务启动: http://0.0.0.0:{PORT} (exchange: {EXCHANGE_ROOT})")
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
