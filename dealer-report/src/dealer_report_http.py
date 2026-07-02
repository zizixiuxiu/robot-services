#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dealer report HTTP service.

Port: 8008
API: POST /process
"""
from __future__ import annotations

import argparse
import base64
import json
import logging
import os
import shutil
import sys
import tempfile
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))

from generate_dealer_report import (  # noqa: E402
    build_output_df,
    process_account_data,
    process_designer_data,
    write_formatted_excel,
)


OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", str(WORK_DIR.parent / "data" / "output")))
LOG_DIR = Path(os.getenv("LOG_DIR", str(WORK_DIR.parent / "logs")))
LOG_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOG_DIR / "dealer_report_http.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("dealer-report")


def _detect_file_type(filename: str) -> str:
    name = filename or ""
    if "账号指标" in name or "账号信息" in name:
        return "account"
    if "设计师数据统计" in name or "设计师" in name:
        return "designer"
    return "unknown"


def _safe_stem(filename: str) -> str:
    stem = Path(filename or "dealer-report").stem.strip()
    return stem.strip(" -_") or "dealer-report"


def _decode_to_file(content_b64: str, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(base64.b64decode(content_b64))
    return path


def _encode_output_file(path: Path, filename: str | None = None) -> dict:
    return {
        "path": str(path),
        "filename": filename or path.name,
        "file_content": base64.b64encode(path.read_bytes()).decode("utf-8"),
    }


def run_dealer_report(account_path: Path, designer_path: Path, output_path: Path, title: str) -> dict:
    account_agg = process_account_data(account_path)
    designer_agg = process_designer_data(designer_path)
    out_df = build_output_df(account_agg, designer_agg)
    write_formatted_excel(out_df, output_path, title=title)
    return {
        "rows": int(len(out_df)),
        "account_rows": int(len(account_agg)),
        "designer_rows": int(len(designer_agg)),
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        logger.info("%s - %s", self.address_string(), format % args)

    def _send_json(self, status_code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path not in ("/", "/health"):
            self._send_json(404, {"error": "未知接口"})
            return
        self._send_json(200, {"status": "ok", "service": "dealer-report", "port": 8008})

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path not in ("/process", "/convert"):
            self._send_json(404, {"error": "未知接口，请用 POST /process 或 POST /convert"})
            return

        try:
            content_len = int(self.headers.get("Content-Length", 0))
            req = json.loads(self.rfile.read(content_len).decode("utf-8"))
        except Exception:
            self._send_json(400, {"success": False, "error": "请求体必须是 JSON"})
            return

        result = self._process(req)
        self._send_json(200, result)

    def _process(self, req: dict) -> dict:
        files = req.get("files")
        if not isinstance(files, list) or not files:
            return {"success": False, "error": "需要上传 files 数组，包含账号指标和设计师数据统计两个文件"}

        tmpdir = Path(tempfile.mkdtemp(prefix="dealer_report_"))
        try:
            file_map: dict[str, Path] = {}
            first_filename = ""
            for item in files:
                content = item.get("file_content")
                filename = item.get("filename", "input.xlsx")
                if not first_filename:
                    first_filename = filename
                if not content:
                    continue
                path = _decode_to_file(content, tmpdir / filename)
                file_type = _detect_file_type(filename)
                if file_type != "unknown":
                    file_map[file_type] = path

            missing = [name for name in ("account", "designer") if name not in file_map]
            if missing:
                return {
                    "success": False,
                    "error": f"缺少必要文件: {missing}。文件名需包含：账号指标、设计师数据统计",
                }

            title = req.get("title") or "6月经销商数据"
            output_name = req.get("output_filename") or f"{_safe_stem(first_filename)}_经销商数据.xlsx"
            output_dir = OUTPUT_DIR / f"output_http_{int(time.time())}"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / output_name

            t0 = time.time()
            stats = run_dealer_report(file_map["account"], file_map["designer"], output_path, title)
            cost = round(time.time() - t0, 3)
            if not output_path.exists():
                return {"success": False, "error": "输出文件未生成"}

            logger.info("Dealer report generated: %s, cost=%.3fs", output_path, cost)
            return {
                "success": True,
                "service": "dealer-report",
                "output_dir": str(output_dir),
                "output_filename": output_name,
                "output_files": [_encode_output_file(output_path, output_name)],
                "stats": stats,
                "cost_seconds": cost,
            }
        except Exception as exc:
            import traceback

            logger.exception("Failed to process dealer report")
            return {"success": False, "error": f"服务异常: {exc}", "trace": traceback.format_exc()}
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


def run(port: int = 8008):
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    logger.info("[dealer-report] HTTP 服务启动于 http://0.0.0.0:%d", port)
    server.serve_forever()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8008)
    args = parser.parse_args()
    run(args.port)
