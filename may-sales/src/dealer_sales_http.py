#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
5月销售部业绩核对表 HTTP 服务
端口：8003（替换原经销商销售服务）
调用 generate_may_sales_report.js 处理 Excel 文件
支持 Windows 本机运行和 Docker 容器运行
"""
import os
import sys
import json
import time
import shutil
import base64
import logging
import tempfile
import subprocess
from pathlib import Path
from urllib.parse import urlparse
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

# 工作目录
WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
SCRIPT_PATH = WORK_DIR / "generate_may_sales_report.js"
NODE_EXE = "node"

# 默认模板路径；Docker 中通过环境变量 DEFAULT_TEMPLATE 覆盖为 /app/templates/...
DEFAULT_TEMPLATE = Path(os.getenv(
    "DEFAULT_TEMPLATE",
    r'D:\wechat\xwechat_files\wxid_0fh4oxng8dq212_f810\msg\file\2026-06\2026年5月销售部业绩核对表 - 副本.xlsx'
))

# 输出目录；默认在项目 data/output_may_sales，Docker 中通过环境变量覆盖为 /app/output_may_sales
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", str(WORK_DIR.parent / "data" / "output_may_sales")))

# ==================== 日志配置 ====================
LOG_DIR = Path(os.getenv("LOG_DIR", str(WORK_DIR.parent / "logs")))
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "dealer_sales_http.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("dealer-sales")
# ==================================================


def _detect_file_type(filename: str) -> str:
    """根据文件名识别文件类型"""
    name = filename.lower()
    if '综合查询' in name:
        return 'zhcx'
    if '联思' in name:
        return 'liansi'
    if '奢匠' in name or '下单统计' in name:
        return 'shejiang'
    if '核对表' in name or '待核对' in name or '模板' in name:
        return 'template'
    return 'unknown'


def run_may_sales(zhcx_path: str, liansi_path: str, shejiang_path: str, template_path: str, output_path: str) -> dict:
    """调用 generate_may_sales_report.js"""
    if not SCRIPT_PATH.exists():
        logger.error("找不到脚本: %s", SCRIPT_PATH)
        return {"success": False, "error": f"找不到脚本: {SCRIPT_PATH}"}

    # 清理输出文件
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except Exception as e:
            logger.error("清理旧输出文件失败: %s", e)
            return {"success": False, "error": f"清理旧输出文件失败: {e}"}

    cmd = [NODE_EXE, str(SCRIPT_PATH), zhcx_path, liansi_path, shejiang_path, template_path, output_path]
    logger.info("调用 Node.js: %s", " ".join(cmd))
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            cwd=str(WORK_DIR),
        )
    except subprocess.TimeoutExpired:
        logger.error("业绩核对表处理超时")
        return {"success": False, "error": "业绩核对表处理超时（超过 5 分钟）"}
    except Exception as e:
        logger.exception("调用 Node.js 异常")
        return {"success": False, "error": f"调用 Node.js 异常: {str(e)}"}

    if result.returncode != 0:
        logger.error("Node.js 执行失败: code=%d, stderr=%s", result.returncode, result.stderr)
        return {
            "success": False,
            "error": f"Node.js 执行失败 (code={result.returncode})",
            "stderr": result.stderr,
            "stdout": result.stdout,
        }

    if not os.path.exists(output_path):
        logger.error("输出文件未生成")
        return {
            "success": False,
            "error": "输出文件未生成",
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    logger.info("Node.js 执行成功，输出: %s", output_path)
    return {
        "success": True,
        "output_file": output_path,
        "stdout": result.stdout,
        "stderr": result.stderr,
    }


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

    def _process_files(self, req: dict) -> dict:
        files = req.get('files', [])
        logger.info("收到 5月业绩核对请求，文件数=%d", len(files))
        if not files or len(files) < 3:
            return {"success": False, "error": "需要至少 3 个文件：综合查询、联思系统、奢匠下单统计"}

        # 准备临时目录
        tmpdir = tempfile.mkdtemp()
        file_map = {}
        template_path = None

        for f in files:
            file_content = f.get('file_content')
            filename = f.get('filename', 'unknown')
            if not file_content:
                continue
            local_path = os.path.join(tmpdir, filename)
            with open(local_path, 'wb') as fh:
                fh.write(base64.b64decode(file_content))

            ftype = _detect_file_type(filename)
            if ftype == 'template':
                template_path = local_path
            elif ftype != 'unknown':
                file_map[ftype] = local_path

        # 检查必要文件
        missing = []
        for key in ['zhcx', 'liansi', 'shejiang']:
            if key not in file_map:
                missing.append(key)
        if missing:
            logger.warning("缺少必要文件: %s", missing)
            shutil.rmtree(tmpdir, ignore_errors=True)
            return {"success": False, "error": f"缺少必要文件，无法识别: {missing}。文件名需包含：综合查询、联思、奢匠/下单统计"}

        # 模板文件
        if not template_path:
            if DEFAULT_TEMPLATE.exists():
                template_path = str(DEFAULT_TEMPLATE)
                logger.info("使用默认模板: %s", template_path)
            else:
                logger.error("未上传模板文件，且默认模板不存在")
                shutil.rmtree(tmpdir, ignore_errors=True)
                return {"success": False, "error": "未上传模板文件，且默认模板不存在"}

        # 输出路径
        if not OUTPUT_DIR.exists():
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_filename = '2026年5月销售部业绩核对表.xlsx'
        output_path = str(OUTPUT_DIR / output_filename)

        t0 = time.time()
        result = run_may_sales(
            file_map['zhcx'],
            file_map['liansi'],
            file_map['shejiang'],
            template_path,
            output_path,
        )
        result['cost_seconds'] = round(time.time() - t0, 3)

        if result.get('success') and os.path.exists(output_path):
            # 返回路径的同时返回 base64 内容，便于 Docker/跨系统部署
            try:
                with open(output_path, 'rb') as fh:
                    file_content = base64.b64encode(fh.read()).decode('utf-8')
                result['output_files'] = [{
                    "path": output_path,
                    "filename": output_filename,
                    "file_content": file_content,
                }]
                logger.info("输出文件已编码: %s (%d bytes)", output_filename, len(file_content))
            except Exception as e:
                logger.error("读取输出文件内容失败: %s", e)
                result['output_files'] = [{"path": output_path, "filename": output_filename}]
                result['content_warning'] = f"读取输出文件内容失败: {e}"

        shutil.rmtree(tmpdir, ignore_errors=True)
        logger.info("5月业绩核对处理完成，success=%s", result.get('success'))
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

        result = self._process_files(req)
        self._send_json(200, result)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/health':
            self._send_json(200, {"status": "ok", "service": "may-sales", "port": 8003})
            return
        self._send_json(200, {"status": "ok", "service": "may-sales", "port": 8003})


def run(port=8003):
    server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    logger.info("[may-sales] HTTP 服务启动于 http://0.0.0.0:%d", port)
    server.serve_forever()


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8003)
    args = parser.parse_args()
    run(args.port)
