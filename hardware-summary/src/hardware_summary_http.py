#!/usr/bin/env python3
"""
五金汇总 HTTP 服务（无 LLM，直接处理）
端口：8001
接口：POST /process  上传文件，直接返回处理结果
"""
import os
import sys
import json
import time
import shutil
import tempfile
import base64
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

from bom_utils import detect_file_type, xlsx_to_xls, get_clean_filename, detect_unusual_sheets


# ==================== 日志配置 ====================
LOG_DIR = Path(os.getenv("LOG_DIR", str(WORK_DIR.parent / "logs")))
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "hardware_summary_http.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("hardware-summary")
# ==================================================



def _get_today_date():
    now = datetime.now()
    return f"{now.year}.{now.month}.{now.day}"


def _encode_output_file(file_path: Path, filename: str = None) -> dict:
    try:
        with open(file_path, 'rb') as fh:
            b64 = base64.b64encode(fh.read()).decode('utf-8')
        return {
            "path": str(file_path),
            "filename": filename or file_path.name,
            "file_content": b64,
        }
    except Exception as e:
        logger.error("读取输出文件失败: %s, error=%s", file_path, e)
        return {
            "path": str(file_path),
            "filename": filename or file_path.name,
            "error": f"读取文件内容失败: {e}",
        }


def _encode_output_files(file_paths: list[Path]) -> list[dict]:
    max_workers = max(1, int(os.getenv("HARDWARE_ENCODE_MAX_WORKERS", "2")))
    if len(file_paths) <= 1 or max_workers == 1:
        return [_encode_output_file(path) for path in file_paths]
    with ThreadPoolExecutor(max_workers=min(len(file_paths), max_workers)) as executor:
        return list(executor.map(_encode_output_file, file_paths))


def _process_hardware_summary(input_path: str, output_dir: str, order_date: str = None) -> dict:
    """处理单个五金汇总"""
    logger.info("开始处理文件: %s -> %s", input_path, output_dir)
    ftype = detect_file_type(input_path)
    if ftype != 'hardware':
        err_msg = "文件类型错误：这是料单拆分文件，请发到报价料单处理群处理。" if ftype == 'order_split' else "文件类型错误：无法识别此文件，不是五金汇总格式。"
        logger.warning("文件类型不匹配: %s, type=%s", input_path, ftype)
        return {"success": False, "error": err_msg}

    if input_path.endswith('.xlsx'):
        xls_path = input_path.rsplit('.', 1)[0] + '.xls'
        xlsx_to_xls(input_path, xls_path)
        input_path = xls_path
        logger.info("xlsx 已转换为 xls: %s", xls_path)

    clean_name = get_clean_filename(Path(input_path).name)
    if not order_date:
        order_date = _get_today_date()

    old_cwd = os.getcwd()
    os.chdir(str(WORK_DIR))
    try:
        from hardware.convert import convert_hardware_summary
        from hardware.hide_prices import generate_factory_version
        t0 = time.time()
        convert_hardware_summary(
            input_path,
            os.path.join(output_dir, f"{clean_name}_五金汇总.xlsx")
        )
        logger.info("生成五金汇总完成, cost=%.3fs", time.time() - t0)
        t0 = time.time()
        generate_factory_version(
            input_path,
            os.path.join(output_dir, f"{clean_name}_工厂版.xls"),
            order_date
        )
        logger.info("生成工厂版完成, cost=%.3fs", time.time() - t0)
    except Exception as e:
        import traceback
        logger.exception("处理文件失败: %s", input_path)
        return {"success": False, "error": f"生成失败: {str(e)}", "trace": traceback.format_exc()}
    finally:
        os.chdir(old_cwd)

    result = {
        "success": True,
        "output_dir": output_dir,
        "count": 2,
        "files": [f"{clean_name}_五金汇总.xlsx", f"{clean_name}_工厂版.xls"],
        "type": "hardware",
        "note": f"工厂版下单日期已替换为 {order_date}",
    }

    unusual = detect_unusual_sheets(input_path)
    if unusual:
        names = "、".join(unusual)
        result["warning"] = f"⚠️ 提醒：检测到【{names}】sheet，此类内容未纳入五金汇总，请处理人员单独核对。"
        logger.warning("检测到非常规 sheet: %s", names)

    logger.info("文件处理完成: %s", input_path)
    return result


class Handler(BaseHTTPRequestHandler):
    # 使用 logging 模块记录访问日志，不再静默
    def log_message(self, format, *args):
        logger.info("%s - %s", self.address_string(), format % args)

    def _send_json(self, status_code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _process_single(self, req: dict) -> dict:
        """处理单个文件"""
        input_path = req.get('input_path')
        file_content = req.get('file_content')
        filename = req.get('filename', 'input.xls')
        order_date = req.get('order_date')

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
        result = _process_hardware_summary(input_path, output_dir, order_date)
        result['cost_seconds'] = round(time.time() - t0, 3)

        if result.get('success'):
            pair_key = str(Path(input_path).name)
            output_files_pair = _encode_output_files([Path(output_dir) / f for f in result['files']])
            for item in output_files_pair:
                logger.info("输出文件已编码: %s", item.get("filename"))
            result['output_files'] = {
                pair_key: output_files_pair
            }

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
        files = req.get('files')
        if files and isinstance(files, list):
            logger.info("批量处理开始, 文件数=%d", len(files))
            t0 = time.time()
            results = []
            all_output_files = {}
            has_error = False
            # 批量并行 worker 数：默认 2，可通过环境变量 HARDWARE_MAX_WORKERS 调整
            default_max_workers = int(os.getenv("HARDWARE_MAX_WORKERS", "2"))
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
            self._send_json(200, {
                "success": not has_error or len(all_output_files) > 0,
                "batch": True,
                "count": len(files),
                "results": results,
                "output_files": all_output_files,
                "cost_seconds": total_cost,
            })
            return

        # 单文件模式（兼容旧接口）
        logger.info("单文件处理请求")
        result = self._process_single(req)
        self._send_json(200, result)

    def do_GET(self):
        self._send_json(200, {"status": "ok", "service": "hardware-summary", "port": 8001})


def run(port=8001):
    server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    logger.info("[hardware-summary] HTTP 服务启动于 http://0.0.0.0:%d", port)
    server.serve_forever()


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8001)
    args = parser.parse_args()
    run(args.port)
