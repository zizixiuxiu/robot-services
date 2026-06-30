#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PVC 自动分类 HTTP 服务（Docker / Linux 兼容）
端口：8005
接口：
  GET  /health
  POST /process
"""
import os
import sys
import json
import time
import base64
import shutil
import tempfile
import logging
import re
import zipfile
from pathlib import Path
from urllib.parse import urlparse
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))

from process_single import extract_date_code, process_input

# 输出根目录：默认在项目 data/output，Docker 中通过环境变量覆盖为 /app/data/output
OUTPUT_BASE = Path(os.getenv("OUTPUT_BASE", str(WORK_DIR.parent / "data" / "output")))

# ==================== 日志配置 ====================
LOG_DIR = Path(os.getenv("LOG_DIR", str(WORK_DIR.parent / "logs")))
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "pvc_classify_http.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("pvc-classify")
# ==================================================


def _resolve_dir_name(filename: str) -> str:
    """根据上传文件名推断输出目录名"""
    base = Path(filename).stem
    # 日期后带 W：6-14W平板 -> W6-14
    m = re.search(r"(\d+-\d+)[^\d-]*W", base, re.IGNORECASE)
    if m:
        return "W" + m.group(1)
    m = re.search(r"(\d+-\d+)", base)
    if m:
        return m.group(1)
    return base


def _safe_zip_name(name: str) -> str:
    return name.replace("/", "_").replace("\\", "_").strip() or "pvc-classify"


def _zip_output_files(file_paths: list[str], zip_path: Path, archive_prefix: str) -> Path:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in file_paths:
            path = Path(file_path)
            if path.exists():
                zf.write(path, f"{archive_prefix}/{path.name}")
    return zip_path


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        logger.info("%s - %s", self.address_string(), format % args)

    def _send_json(self, status_code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _encode_output_file(self, file_path: Path, filename: str = None) -> dict:
        """读取文件并返回 base64 编码结构"""
        try:
            with open(file_path, "rb") as fh:
                b64 = base64.b64encode(fh.read()).decode("utf-8")
            return {
                "filename": filename or file_path.name,
                "file_content": b64,
            }
        except Exception as e:
            logger.error("读取输出文件失败: %s, error=%s", file_path, e)
            return {
                "filename": filename or file_path.name,
                "error": f"读取文件内容失败: {e}",
            }

    def _process_single(self, req: dict) -> dict:
        file_content = req.get("file_content")
        filename = req.get("filename", "input.xls")
        dir_name = req.get("dir_name")

        if not file_content:
            return {"success": False, "error": "缺少 file_content"}

        tmpdir = Path(tempfile.mkdtemp(prefix="pvc_http_in_"))
        work_dir = None
        try:
            input_path = tmpdir / filename
            with open(input_path, "wb") as f:
                f.write(base64.b64decode(file_content))

            if not dir_name:
                dir_name = _resolve_dir_name(filename)

            logger.info("处理文件: %s -> dir_name=%s", filename, dir_name)
            t0 = time.time()
            result = process_input(
                str(input_path),
                dir_name=dir_name,
                output_base=str(OUTPUT_BASE),
                debug_output=False,
            )
            cost = round(time.time() - t0, 3)
            work_dir = Path(result["work_dir"])

            original_output_count = len(result["output_files"])
            zip_base = f"{_safe_zip_name(dir_name)}-自动分类"
            zip_path = work_dir / f"{zip_base}.zip"
            _zip_output_files(result["output_files"], zip_path, zip_base)
            encoded_files = [self._encode_output_file(zip_path)]
            logger.info(
                "转换完成: %s, 输出文件数=%d, cost=%.3fs",
                filename,
                original_output_count,
                cost,
            )

            return {
                "success": True,
                "filename": filename,
                "dir_name": dir_name,
                "count": len(encoded_files),
                "zip_original_count": original_output_count,
                "quantity_total": result.get("quantity_total", 0),
                "quantity_files": result.get("quantity_files", []),
                "output_files": {filename: encoded_files},
                "cost_seconds": cost,
            }
        except Exception as e:
            import traceback
            logger.exception("处理文件失败: %s", filename)
            return {
                "success": False,
                "error": f"服务异常: {str(e)}",
                "trace": traceback.format_exc(),
            }
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
            if work_dir and work_dir.exists():
                shutil.rmtree(work_dir, ignore_errors=True)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/process":
            self._send_json(404, {"error": "未知接口，请用 POST /process"})
            return

        content_len = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_len)

        try:
            req = json.loads(body.decode("utf-8"))
        except Exception:
            self._send_json(400, {"error": "请求体必须是 JSON"})
            return

        # 批量模式
        files = req.get("files")
        if files and isinstance(files, list):
            logger.info("批量处理开始, 文件数=%d", len(files))
            t0 = time.time()
            results = []
            all_output_files = {}
            has_error = False
            for f in files:
                res = self._process_single(f)
                results.append(res)
                if res.get("success"):
                    all_output_files.update(res.get("output_files", {}))
                else:
                    has_error = True
            total_cost = round(time.time() - t0, 3)
            total_output_count = sum(len(v) for v in all_output_files.values())
            quantity_total = sum(float(res.get("quantity_total") or 0) for res in results if res.get("success"))
            quantity_files = []
            for res in results:
                if res.get("success"):
                    quantity_files.extend(res.get("quantity_files") or [])
            logger.info(
                "批量处理完成, 成功=%s, 输出文件数=%d, total_cost=%.3fs",
                not has_error,
                total_output_count,
                total_cost,
            )

            resp = {
                "success": not has_error or len(all_output_files) > 0,
                "batch": True,
                "count": len(files),
                "results": results,
                "output_files": all_output_files,
                "quantity_total": int(quantity_total) if quantity_total.is_integer() else quantity_total,
                "quantity_files": quantity_files,
                "cost_seconds": total_cost,
            }
            if has_error:
                error_msgs = []
                for idx, f in enumerate(files):
                    res = results[idx]
                    if not res.get("success"):
                        err = res.get("error", "未知错误")
                        fn = f.get("filename", "unknown")
                        error_msgs.append(f"[{fn}] {err}")
                if error_msgs:
                    resp["error"] = "; ".join(error_msgs)
            self._send_json(200, resp)
            return

        # 单文件模式
        logger.info("单文件处理请求")
        result = self._process_single(req)
        self._send_json(200, result)

    def do_GET(self):
        self._send_json(200, {"status": "ok", "service": "pvc-classify", "port": 8005})


def run(port=8005):
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    logger.info("[pvc-classify] HTTP 服务启动于 http://0.0.0.0:%d", port)
    server.serve_forever()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8005)
    args = parser.parse_args()
    run(args.port)
