#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
下车间单转换 HTTP 服务
端口：8006
接口：POST /process
"""
import os
import sys
import json
import time
import base64
import shutil
import tempfile
import logging
from pathlib import Path
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))

# 输出根目录：默认在项目 data/output，Docker 中通过环境变量覆盖为 /app/data/output
OUTPUT_BASE = Path(os.getenv("OUTPUT_BASE", str(WORK_DIR.parent / "data" / "output")))

from make_workshop_order import transform, convert_xls_to_xlsx


# ==================== 日志配置 ====================
LOG_DIR = Path(os.getenv("LOG_DIR", str(WORK_DIR.parent / "logs")))
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "workshop_order_http.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("workshop-order")
# ==================================================


def _convert_xls_to_xlsx(input_path: Path) -> Path:
    """将 .xls 转换为 .xlsx（Docker 内无 Excel COM，使用 pandas 转换）。"""
    if input_path.suffix.lower() != ".xls":
        return input_path

    xlsx_path = input_path.with_suffix(".xlsx")
    logger.info("xls 转 xlsx: %s -> %s", input_path, xlsx_path)
    convert_xls_to_xlsx(input_path, xlsx_path)
    return xlsx_path


class Handler(BaseHTTPRequestHandler):
    VALID_ORDER_TYPES = {"auto", "tiepi", "hunyou"}

    def log_message(self, format, *args):
        logger.info("%s - %s", self.address_string(), format % args)

    def _send_json(self, status_code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _encode_output_file(self, file_path: str, filename: str = None) -> dict:
        """读取文件并返回 base64 编码结构"""
        try:
            with open(file_path, "rb") as fh:
                b64 = base64.b64encode(fh.read()).decode("utf-8")
            return {
                "path": file_path,
                "filename": filename or Path(file_path).name,
                "file_content": b64,
            }
        except Exception as e:
            logger.error("读取输出文件失败: %s, error=%s", file_path, e)
            return {
                "path": file_path,
                "filename": filename or Path(file_path).name,
                "error": f"读取文件内容失败: {e}",
            }

    def _process_single(self, req: dict) -> dict:
        file_content = req.get("file_content")
        filename = req.get("filename", "input.xlsx")
        order_type = req.get("order_type", req.get("order-type", "auto"))

        if not file_content:
            return {"success": False, "error": "缺少 file_content"}

        if order_type not in self.VALID_ORDER_TYPES:
            return {
                "success": False,
                "error": f"order_type 必须是 auto/tiepi/hunyou 之一，收到: {order_type}",
            }

        tmpdir = tempfile.mkdtemp()
        try:
            input_path = Path(tmpdir) / filename
            with open(input_path, "wb") as f:
                f.write(base64.b64decode(file_content))

            # .xls 需先转换为 .xlsx
            converted_path = _convert_xls_to_xlsx(input_path)

            base = Path(filename).stem
            timestamp = str(int(time.time()))
            output_dir = OUTPUT_BASE / f"output_http_{base}_{timestamp}"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_name = f"{base}下车间.xlsx"
            output_path = output_dir / output_name

            t0 = time.time()
            stats = transform(
                converted_path,
                output_path,
                discount=float(os.getenv("WORKSHOP_DISCOUNT", "0.85")),
                order_type=order_type,
            )
            cost = round(time.time() - t0, 3)
            logger.info("转换完成, cost=%.3fs, stats=%s", cost, stats)

            if not output_path.exists():
                return {"success": False, "error": "输出文件未生成"}

            pair_key = str(Path(filename).name)
            return {
                "success": True,
                "order_type": order_type,
                "output_dir": str(output_dir),
                "output_files": {
                    pair_key: [self._encode_output_file(str(output_path), output_name)],
                },
                "output_filename": output_name,
                "stats": stats,
                "cost_seconds": cost,
            }
        except Exception as e:
            import traceback
            logger.exception("处理文件失败: %s", filename)
            return {"success": False, "error": f"服务异常: {str(e)}", "trace": traceback.format_exc()}
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

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

        # 批量模式
        files = req.get("files")
        if files and isinstance(files, list):
            logger.info("批量处理开始, 文件数=%d", len(files))
            t0 = time.time()
            results = []
            all_output_files = {}
            has_error = False
            default_max_workers = int(os.getenv("WORKSHOP_MAX_WORKERS", "1"))
            max_workers = max(1, min(len(files), default_max_workers))
            logger.info("批量处理使用 %d 个并发 worker", max_workers)
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_idx = {executor.submit(self._process_single, f): i for i, f in enumerate(files)}
                for future in as_completed(future_to_idx):
                    idx = future_to_idx[future]
                    try:
                        res = future.result()
                    except Exception as e:
                        logger.exception("批量处理中第 %d 个文件异常", idx)
                        res = {"success": False, "error": f"处理异常: {e}"}
                    results.append((idx, res))
                    if res.get("success"):
                        all_output_files.update(res.get("output_files", {}))
                    else:
                        has_error = True
            results.sort(key=lambda x: x[0])
            results = [r[1] for r in results]
            total_cost = round(time.time() - t0, 3)
            total_output_count = sum(len(v) for v in all_output_files.values())
            logger.info("批量处理完成, 成功=%s, 输出对数=%d, 文件数=%d, total_cost=%.3fs",
                        not has_error, len(all_output_files), total_output_count, total_cost)

            resp = {
                "success": not has_error or len(all_output_files) > 0,
                "batch": True,
                "count": len(files),
                "results": results,
                "output_files": all_output_files,
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
        self._send_json(200, {"status": "ok", "service": "workshop-order", "port": 8006})


def run(port=8006):
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    logger.info("[workshop-order] HTTP 服务启动于 http://0.0.0.0:%d", port)
    server.serve_forever()


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8006)
    args = parser.parse_args()
    run(args.port)
