# -*- coding: utf-8 -*-
"""Mirror workshop-order files onto the Ishare share, and keep a local backup.

Docker Desktop cannot bind a UNC path, so the container writes locally and
this host process:
  1. copies staging -> UNC base (bot output; the year is part of the relative path)
  2. copies UNC <year> -> local backup (including files they put there themselves)
  3. copies staging -> local backup
Layout under both roots is {年份}/{月份}/{姓名}/原文件|处理后, produced by the container
from the processing date, so nothing here needs to know the current year.

Backup never deletes: if someone removes a file on the share, the backup copy stays.
Backup is not copied back onto the share.
"""
from __future__ import annotations

import logging
import os
import shutil
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

SRC = Path(os.getenv(
    "WORKSHOP_ARCHIVE_STAGE",
    r"D:\Services\robot-services\workshop-order\data\ishare-archive",
))
# \u5171\u4eab\u76d8\u57fa\u5740\uff0c\u4e0d\u542b\u5e74\u4efd\uff1a\u5e74\u4efd\u7531\u5bb9\u5668\u5199\u5728\u6682\u5b58\u8def\u5f84\u91cc\uff082026/9/\u59d3\u540d/\u2026\uff09\uff0c\u968f\u76f8\u5bf9\u8def\u5f84\u5e26\u8fc7\u6765\uff0c
# \u8de8\u5e74\u4e0d\u7528\u6539\u914d\u7f6e\u3002\u6ce8\u610f\u57fa\u5740\u4e0b\u8fd8\u6709\u300c\u5df2\u5f55\u300d\u300c\u672a\u5f55\u300d\u300c\u660e\u7ec6\u8868\u300d\u7b49\u522b\u4eba\u7684\u76ee\u5f55\uff0c
# \u6240\u4ee5\u53ea\u80fd\u62ff\u6682\u5b58/\u5e74\u4efd\u5939\u5f53\u6e90\uff0c\u7edd\u4e0d\u80fd\u62ff\u57fa\u5740\u5f53\u6e90\u5f80\u56de\u642c\u3002
UNC_BASE = Path(os.getenv(
    "WORKSHOP_ARCHIVE_UNC",
    "\\\\SKY-20340208RFW\\Ishare" + "\u5f55\u5355\u4e34\u65f6\u6587\u4ef6",
))
BACKUP = Path(os.getenv(
    "WORKSHOP_ARCHIVE_BACKUP",
    r"D:\Ishare_2026_backup",
))
SKIP_PERSONS = {"\u672a\u77e5\u4e0a\u4f20\u8005"}
LISTEN = os.getenv("WORKSHOP_ARCHIVE_SYNC_LISTEN", "127.0.0.1:8016")
LOG_FILE = Path(os.getenv(
    "WORKSHOP_ARCHIVE_SYNC_LOG",
    r"D:\Services\robot-services\workshop-order\logs\ishare_sync.log",
))
INTERVAL = float(os.getenv("WORKSHOP_ARCHIVE_SYNC_INTERVAL", "3"))

LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
logger = logging.getLogger("ishare-sync")


def _is_month(seg: str) -> bool:
    return seg.isdigit() and 1 <= int(seg) <= 12


def _is_year(seg: str) -> bool:
    return seg.isdigit() and 1000 <= int(seg) <= 2999


def _person_of(rel: Path) -> str:
    """相对路径里的上传者姓名。兼容 {姓名}/…、{月}/{姓名}/…、{年}/{月}/{姓名}/… 三种布局。"""
    parts = rel.parts
    i = 0
    if i < len(parts) and _is_year(parts[i]):
        i += 1
    if i < len(parts) and _is_month(parts[i]):
        i += 1
    return parts[i] if i < len(parts) else ""


def copy_newer(src_root: Path, dest_root: Path) -> int:
    """Copy new/newer files. Never delete anything in dest_root."""
    if not src_root.exists():
        return 0
    copied = 0
    for src in src_root.rglob("*"):
        if not src.is_file():
            continue
        rel = src.relative_to(src_root)
        if _person_of(rel) in SKIP_PERSONS:
            continue
        dest = dest_root / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            need = not dest.exists() or src.stat().st_mtime > dest.stat().st_mtime + 1
        except OSError:
            logger.warning("stat failed %s", src, exc_info=True)
            continue
        if not need:
            continue
        try:
            shutil.copy2(src, dest)
        except OSError:
            logger.warning("copy failed %s -> %s", src, dest, exc_info=True)
            continue
        copied += 1
        logger.info("copied %s", dest)
    return copied


def _years_in(root: Path) -> list[str]:
    """暂存顶层里长得像年份的目录名（容器按处理日期写进去的那一层）。"""
    if not root.exists():
        return []
    return sorted(p.name for p in root.iterdir() if p.is_dir() and _is_year(p.name))


def sync_once() -> int:
    # 只处理暂存里按年份分的那些夹 —— 容器写入的就是这个布局。
    # 不能整体镜像暂存顶层：那样暂存的 test/ 会直接落到共享盘基址
    # （Ishare录单临时文件\test\），基址下「已录」「未录」等别人的目录
    # 也会被扫进本机备份。逐年夹走两头都干净。
    copied = 0
    for year in _years_in(SRC):
        copied += copy_newer(SRC / year, UNC_BASE / year)      # 暂存 -> 共享盘
        copied += copy_newer(UNC_BASE / year, BACKUP / year)   # 共享盘 -> 本机备份（含他们自己放的）
        copied += copy_newer(SRC / year, BACKUP / year)        # 暂存 -> 本机备份
    return copied


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        logger.info("%s - %s", self.address_string(), fmt % args)

    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"ok")

    def do_POST(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length:
            self.rfile.read(length)
        n = sync_once()
        body = ("{\"copied\": %d}" % n).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    BACKUP.mkdir(parents=True, exist_ok=True)
    host, port_s = LISTEN.split(":")
    server = ThreadingHTTPServer((host, int(port_s)), Handler)
    logger.info(
        "ishare sync listening on %s, src=%s dst=%s backup=%s",
        LISTEN, SRC, UNC_BASE, BACKUP,
    )
    threading.Thread(target=server.serve_forever, daemon=True).start()
    while True:
        try:
            sync_once()
        except Exception:
            logger.exception("sync failed")
        time.sleep(INTERVAL)


if __name__ == "__main__":
    main()
