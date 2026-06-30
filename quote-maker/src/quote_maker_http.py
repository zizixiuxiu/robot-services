#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Quote maker HTTP service.

Port: 8007
API: POST /process
"""
import argparse
import base64
import json
import logging
import os
import shutil
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))

OUTPUT_BASE = Path(os.getenv("OUTPUT_BASE", str(WORK_DIR.parent / "data" / "output")))
DEFAULT_TEMPLATE = Path(os.getenv("QUOTE_TEMPLATE", str(WORK_DIR / "templates" / "quote_template.xlsx")))

from make_quote import (  # noqa: E402
    apply_quote_page_background,
    build_workbook_input_only,
    copy_formulas_from_reference,
    postprocess_hardware_from_input,
)


LOG_DIR = Path(os.getenv("LOG_DIR", str(WORK_DIR.parent / "logs")))
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "quote_maker_http.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("quote-maker")


def _safe_stem(filename: str) -> str:
    stem = Path(filename or "input.xls").stem.strip()
    for suffix in ("拆单报价", "报价料单", "料单"):
        stem = stem.replace(suffix, "")
    stem = stem.strip(" -_")
    return stem or "quote"


def _decode_to_file(content_b64: str, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(base64.b64decode(content_b64))
    return path


def _encode_output_file(file_path: Path, filename: str | None = None) -> dict:
    try:
        with open(file_path, "rb") as fh:
            b64 = base64.b64encode(fh.read()).decode("utf-8")
        return {
            "path": str(file_path),
            "filename": filename or file_path.name,
            "file_content": b64,
        }
    except Exception as exc:
        logger.error("Failed to read output file: %s, error=%s", file_path, exc)
        return {
            "path": str(file_path),
            "filename": filename or file_path.name,
            "error": f"读取文件内容失败: {exc}",
        }


def _resolve_template(req: dict, tmpdir: Path) -> Path:
    template_content = req.get("template_content") or req.get("template_file_content")
    if template_content:
        template_name = req.get("template_filename", "quote_template.xlsx")
        return _decode_to_file(template_content, tmpdir / template_name)
    if not DEFAULT_TEMPLATE.exists():
        raise FileNotFoundError(f"默认报价模板不存在: {DEFAULT_TEMPLATE}")
    return DEFAULT_TEMPLATE


def _resolve_reference(req: dict, tmpdir: Path) -> Path | None:
    reference_content = req.get("reference_content") or req.get("reference_file_content")
    if not reference_content:
        return None
    reference_name = req.get("reference_filename", "reference.xlsx")
    return _decode_to_file(reference_content, tmpdir / reference_name)


def _process_single(req: dict) -> dict:
    file_content = req.get("file_content")
    filename = req.get("filename", "input.xls")
    if not file_content:
        return {"success": False, "error": "缺少 file_content"}

    tmpdir = Path(tempfile.mkdtemp(prefix="quote_maker_"))
    try:
        input_path = _decode_to_file(file_content, tmpdir / filename)
        template_path = _resolve_template(req, tmpdir)
        reference_path = _resolve_reference(req, tmpdir)

        base = _safe_stem(filename)
        timestamp = str(int(time.time()))
        output_dir = OUTPUT_BASE / f"output_http_{base}_{timestamp}"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_name = req.get("output_filename") or f"{base}报价单.xlsx"
        output_path = output_dir / output_name

        t0 = time.time()
        if req.get("match_reference") and reference_path and reference_path.exists():
            shutil.copyfile(reference_path, output_path)
            apply_quote_page_background(output_path)
        else:
            build_workbook_input_only(input_path, template_path, output_path)
            if reference_path and reference_path.exists():
                copy_formulas_from_reference(output_path, reference_path)
            postprocess_hardware_from_input(output_path, input_path)

        cost = round(time.time() - t0, 3)
        if not output_path.exists():
            return {"success": False, "error": "输出文件未生成"}

        pair_key = Path(filename).name
        logger.info("Quote generated: %s -> %s, cost=%.3fs", filename, output_path, cost)
        return {
            "success": True,
            "service": "quote-maker",
            "output_dir": str(output_dir),
            "output_filename": output_name,
            "output_files": {
                pair_key: [_encode_output_file(output_path, output_name)],
            },
            "cost_seconds": cost,
        }
    except Exception as exc:
        import traceback

        logger.exception("Failed to process quote file: %s", filename)
        return {"success": False, "error": f"服务异常: {exc}", "trace": traceback.format_exc()}
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


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
        self._send_json(200, {"status": "ok", "service": "quote-maker", "port": 8007})

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path not in ("/process", "/convert"):
            self._send_json(404, {"error": "未知接口，请用 POST /process 或 POST /convert"})
            return

        content_len = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_len)
        try:
            req = json.loads(body.decode("utf-8"))
        except Exception:
            self._send_json(400, {"error": "请求体必须是 JSON"})
            return

        files = req.get("files")
        if files and isinstance(files, list):
            max_workers = max(1, min(len(files), int(os.getenv("QUOTE_MAX_WORKERS", "1"))))
            results = []
            output_files = {}
            has_error = False
            t0 = time.time()
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_idx = {executor.submit(_process_single, item): i for i, item in enumerate(files)}
                for future in as_completed(future_to_idx):
                    idx = future_to_idx[future]
                    try:
                        res = future.result()
                    except Exception as exc:
                        logger.exception("Batch item %d failed", idx)
                        res = {"success": False, "error": f"处理异常: {exc}"}
                    results.append((idx, res))
                    if res.get("success"):
                        output_files.update(res.get("output_files", {}))
                    else:
                        has_error = True
            results.sort(key=lambda item: item[0])
            results = [item[1] for item in results]
            self._send_json(200, {
                "success": not has_error or bool(output_files),
                "batch": True,
                "count": len(files),
                "results": results,
                "output_files": output_files,
                "cost_seconds": round(time.time() - t0, 3),
            })
            return

        self._send_json(200, _process_single(req))


def run(port: int = 8007):
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    logger.info("[quote-maker] HTTP service started at http://0.0.0.0:%d", port)
    server.serve_forever()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8007)
    args = parser.parse_args()
    run(args.port)
