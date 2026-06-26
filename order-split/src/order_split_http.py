#!/usr/bin/env python3
"""
料单拆分 HTTP 服务（无 LLM，直接处理）
端口：8002
接口：POST /process  上传文件，直接返回处理结果
"""
import os
import sys
import json
import time
import shutil
import tempfile
import base64
import zipfile
import logging
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))

# 输出根目录：默认在项目 data/output，Docker 中通过环境变量覆盖为 /app/data/output
OUTPUT_BASE = Path(os.getenv("OUTPUT_BASE", str(WORK_DIR.parent / "data" / "output")))

from bom_utils import detect_file_type, xlsx_to_xls, get_clean_filename


# ==================== 日志配置 ====================
LOG_DIR = Path(os.getenv("LOG_DIR", str(WORK_DIR.parent / "logs")))
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "order_split_http.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("order-split")
# ==================================================


def _zip_output_files(output_dir: str, output_files: list[str], zip_name: str, archive_prefix: str = None) -> str:
    """将多个输出文件打包成 zip，返回 zip 路径。

    archive_prefix 指定 zip 内根文件夹名称；如果不传，使用 zip 文件名（去掉 .zip）。
    """
    zip_path = str(Path(output_dir) / zip_name)
    if archive_prefix is None:
        archive_prefix = Path(zip_name).stem
    logger.info("打包 zip: output_dir=%s, zip_name=%s, archive_prefix=%s, files=%d", output_dir, zip_name, archive_prefix, len(output_files))
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in output_files:
            path = Path(file_path)
            if path.exists():
                # 保留原始文件相对 output_dir 的目录结构，并放入 archive_prefix 下
                try:
                    arcname = str(path.relative_to(output_dir))
                except ValueError:
                    arcname = path.name
                zf.write(path, f"{archive_prefix}/{arcname}")
    return zip_path


def _process_order_split(input_path: str, output_dir: str) -> dict:
    """处理单个料单拆分"""
    logger.info("开始处理文件: %s -> %s", input_path, output_dir)
    ftype = detect_file_type(input_path)
    if ftype != 'order_split':
        msg = ("文件类型错误：这是五金汇总文件，请发到五金汇总群处理。" if ftype == 'hardware'
               else "文件类型错误：无法识别此文件，不是料单拆分格式。")
        logger.warning("文件类型不匹配: %s, type=%s", input_path, ftype)
        return {"success": False, "error": msg}

    if input_path.endswith('.xlsx'):
        xls_path = input_path.rsplit('.', 1)[0] + '.xls'
        xlsx_to_xls(input_path, xls_path)
        input_path = xls_path
        logger.info("xlsx 已转换为 xls: %s", xls_path)

    old_cwd = os.getcwd()
    os.chdir(str(WORK_DIR))
    try:
        from generate_output import generate_all
        t0 = time.time()
        generate_all(input_path, output_dir)
        logger.info("生成料单完成, cost=%.3fs", time.time() - t0)
    except Exception as e:
        import traceback
        logger.exception("处理文件失败: %s", input_path)
        return {"success": False, "error": f"生成失败: {str(e)}", "trace": traceback.format_exc()}
    finally:
        os.chdir(old_cwd)

    out_files = sorted(Path(output_dir).rglob("*.xlsx")) + sorted(Path(output_dir).rglob("*.xls"))
    clean_files = []
    for f in out_files:
        rel = str(f.relative_to(output_dir))
        clean_files.append(os.sep.join(get_clean_filename(p) for p in rel.split(os.sep)))

    result = {
        "success": True,
        "output_dir": output_dir,
        "count": len(out_files),
        "files": clean_files,
        "type": "order_split",
    }
    logger.info("文件处理完成: %s, 输出数=%d", input_path, len(out_files))
    return result


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        logger.info("%s - %s", self.address_string(), format % args)

    def _send_json(self, status_code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _encode_output_file(self, file_path: str, filename: str = None) -> dict:
        """读取文件并返回 base64 编码结构"""
        try:
            with open(file_path, 'rb') as fh:
                b64 = base64.b64encode(fh.read()).decode('utf-8')
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
        input_path = req.get('input_path')
        file_content = req.get('file_content')
        filename = req.get('filename', 'input.xls')

        tmpdir = None
        if file_content and not input_path:
            tmpdir = tempfile.mkdtemp()
            input_path = os.path.join(tmpdir, filename)
            with open(input_path, 'wb') as f:
                f.write(base64.b64decode(file_content))
            logger.info("收到 base64 文件, 已写入临时文件: %s", input_path)

        if not input_path or not os.path.exists(input_path):
            logger.error("文件不存在: %s", input_path)
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)
            return {"success": False, "error": "文件不存在"}

        base = Path(input_path).stem
        timestamp = str(int(time.time()))
        output_dir = str(OUTPUT_BASE / f"output_http_{base}_{timestamp}")
        os.makedirs(output_dir, exist_ok=True)

        t0 = time.time()
        result = _process_order_split(input_path, output_dir)
        result['cost_seconds'] = round(time.time() - t0, 3)

        if result.get('success'):
            output_files = [
                str(f) for f in sorted(Path(output_dir).rglob("*.xlsx")) + sorted(Path(output_dir).rglob("*.xls"))
            ]
            pair_key = str(Path(input_path).name)
            output_files_pair = []
            logger.info("准备构建 output_files, count=%d, output_dir=%s", len(output_files), output_dir)

            if len(output_files) > 2:
                zip_name = f"{get_clean_filename(Path(input_path).name)}_料单拆分.zip"
                zip_path = _zip_output_files(output_dir, output_files, zip_name)
                output_files_pair.append(self._encode_output_file(zip_path, zip_name))
                result['zip_file'] = zip_path
                result['zip_original_count'] = len(output_files)
            else:
                for f in output_files:
                    output_files_pair.append(self._encode_output_file(f))

            result['output_files'] = {pair_key: output_files_pair}

        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)

        return result

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != '/process':
            self._send_json(404, {"error": "未知接口，请用 POST /process"})
            return

        content_len = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_len)

        try:
            req = json.loads(body.decode('utf-8'))
        except Exception:
            self._send_json(400, {"error": "请求体必须是 JSON"})
            return

        # 批量模式：使用线程池并行处理
        files = req.get("files")
        if files and isinstance(files, list):
            logger.info("批量处理开始, 文件数=%d", len(files))
            t0 = time.time()
            results = []
            all_output_files = {}
            has_error = False
            # 批量并行 worker 数：默认 2，可通过环境变量 ORDER_SPLIT_MAX_WORKERS 调整
            default_max_workers = int(os.getenv("ORDER_SPLIT_MAX_WORKERS", "2"))
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
                    if res.get('success'):
                        all_output_files.update(res.get('output_files', {}))
                    else:
                        has_error = True
            # 按原始顺序排列 results
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
        self._send_json(200, {"status": "ok", "service": "order-split", "port": 8002})


def run(port=8002):
    server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    logger.info("[order-split] HTTP 服务启动于 http://0.0.0.0:%d", port)
    server.serve_forever()


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8002)
    args = parser.parse_args()
    run(args.port)
