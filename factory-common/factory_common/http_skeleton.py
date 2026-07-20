"""HTTP 服务骨架（http.server 实现）：POST /process + GET /health

精确复刻 hardware-summary 原 Handler 的协议行为，供各微服务复用。

扩展点（均为可选参数，默认值下行为与原版逐字节一致）：
- path_aliases / not_found_error: 额外 POST 路径别名（如 /convert）及 404 文案
- missing_file_error: 无 input_path 且无 file_content 时的错误文案
- get_paths / get_not_found_error: GET 路径白名单（None = 任何 GET 返回 ok）
- batch_error_aggregate: 批量响应顶层 error 聚合各失败项 "[filename] error"
- batch_aggregate_hook: 批量响应发送前回调，可修改响应 dict
- startup_hook: serve_forever 前调用一次（如预热），异常只记日志
- multi_file_fn: {files:[...]} 整体作为一个任务处理（多文件合并单输出场景）
- process_fn 返回 dict 已含 output_files 键时，骨架跳过自动编码
"""
import os
import json
import time
import shutil
import tempfile
import base64
from pathlib import Path
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler


def _encode_output_file(file_path: Path, logger, filename: str = None) -> dict:
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


def _encode_output_files(file_paths: list[Path], logger, encode_workers_env: str = None) -> list[dict]:
    env_name = encode_workers_env or "ENCODE_MAX_WORKERS"
    max_workers = max(1, int(os.getenv(env_name, "2")))
    if len(file_paths) <= 1 or max_workers == 1:
        return [_encode_output_file(path, logger) for path in file_paths]
    with ThreadPoolExecutor(max_workers=min(len(file_paths), max_workers)) as executor:
        return list(executor.map(lambda p: _encode_output_file(p, logger), file_paths))


def run_process_server(*, service_name, default_port, process_fn, logger,
                       max_workers_env, default_max_workers=2,
                       encode_workers_env=None, default_output_base=None,
                       path_aliases=(), not_found_error=None,
                       missing_file_error="文件不存在",
                       get_paths=None, get_not_found_error="未知接口",
                       batch_error_aggregate=False, batch_aggregate_hook=None,
                       startup_hook=None, multi_file_fn=None):
    """启动 /process + /health HTTP 服务（http.server 实现）。

    process_fn(input_path: str, output_dir: str, req: dict) -> dict
    返回 dict 含 success；成功且含 files: [名字] 时骨架负责从 output_dir 编码为 output_files pairs。
    若返回 dict 已含 output_files 键（非 None），骨架跳过自动编码直接使用。

    可选扩展参数见模块 docstring。
    """
    # 输出根目录：默认由调用方传入（一般为 <service>/data/output），可用环境变量 OUTPUT_BASE 覆盖
    output_base = Path(os.getenv("OUTPUT_BASE", str(default_output_base)))

    accepted_post_paths = ('/process',) + tuple(path_aliases)
    if not_found_error is None:
        not_found_error = "未知接口，请用 " + " 或 ".join(f"POST {p}" for p in accepted_post_paths)

    def _process_single(req: dict) -> dict:
        """处理单个文件"""
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
            return {"success": False, "error": missing_file_error}

        base = Path(input_path).stem
        timestamp = str(int(time.time()))
        output_dir = str(output_base / f"output_http_{base}_{timestamp}")
        os.makedirs(output_dir, exist_ok=True)

        t0 = time.time()
        try:
            result = process_fn(input_path, output_dir, req)
            result['cost_seconds'] = round(time.time() - t0, 3)

            if result.get('success'):
                if result.get('output_files') is not None:
                    # process_fn 已自行产出 output_files（如自行打 zip / 自定义编码），骨架跳过
                    logger.info("process_fn 已提供 output_files, 跳过自动编码")
                elif result.get('files'):
                    pair_key = str(Path(input_path).name)
                    output_files_pair = _encode_output_files(
                        [Path(output_dir) / f for f in result['files']], logger, encode_workers_env)
                    for item in output_files_pair:
                        logger.info("输出文件已编码: %s", item.get("filename"))
                    result['output_files'] = {
                        pair_key: output_files_pair
                    }
        finally:
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)

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

        def do_POST(self):
            parsed = urlparse(self.path)
            if parsed.path not in accepted_post_paths:
                self._send_json(404, {"error": not_found_error})
                return

            content_len = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_len)

            try:
                req = json.loads(body.decode('utf-8'))
            except Exception:
                self._send_json(400, {"error": "请求体必须是 JSON"})
                return

            # 多文件整体任务模式：{files:[...]} 作为一个任务交给 multi_file_fn
            files = req.get('files')
            if files and isinstance(files, list) and multi_file_fn is not None:
                logger.info("多文件整体任务开始, 文件数=%d", len(files))
                t0 = time.time()
                work_dir = tempfile.mkdtemp()
                try:
                    response = multi_file_fn(req, work_dir)
                    if 'cost_seconds' not in response:
                        response['cost_seconds'] = round(time.time() - t0, 3)
                except Exception as e:
                    logger.exception("多文件整体任务异常")
                    response = {"success": False, "error": f"处理异常: {e}",
                                "cost_seconds": round(time.time() - t0, 3)}
                finally:
                    shutil.rmtree(work_dir, ignore_errors=True)
                logger.info("多文件整体任务完成, success=%s, cost=%.3fs",
                            response.get('success'), response.get('cost_seconds', 0))
                self._send_json(200, response)
                return

            # 批量模式：使用线程池并行处理
            if files and isinstance(files, list):
                logger.info("批量处理开始, 文件数=%d", len(files))
                t0 = time.time()
                results = []
                all_output_files = {}
                has_error = False
                # 批量并行 worker 数：默认 default_max_workers，可通过环境变量调整
                env_default = int(os.getenv(max_workers_env, str(default_max_workers)))
                max_workers = max(1, min(len(files), env_default))
                logger.info("批量处理使用 %d 个并发 worker", max_workers)
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_idx = {executor.submit(_process_single, f): i for i, f in enumerate(files)}
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
                response = {
                    "success": not has_error or len(all_output_files) > 0,
                    "batch": True,
                    "count": len(files),
                    "results": results,
                    "output_files": all_output_files,
                    "cost_seconds": total_cost,
                }
                if batch_error_aggregate:
                    errors = []
                    for f, r in zip(files, results):
                        if not r.get('success'):
                            fname = f.get('filename', 'input.xls') if isinstance(f, dict) else 'input.xls'
                            errors.append(f"[{fname}] {r.get('error', '处理失败')}")
                    if errors:
                        response["error"] = "; ".join(errors)
                if batch_aggregate_hook is not None:
                    batch_aggregate_hook(response, results, files)
                self._send_json(200, response)
                return

            # 单文件模式（兼容旧接口）
            logger.info("单文件处理请求")
            result = _process_single(req)
            self._send_json(200, result)

        def do_GET(self):
            if get_paths is not None and urlparse(self.path).path not in get_paths:
                self._send_json(404, {"error": get_not_found_error})
                return
            self._send_json(200, {"status": "ok", "service": service_name, "port": port})

    def run(port):
        if startup_hook is not None:
            try:
                startup_hook()
                logger.info("[%s] startup_hook 执行完成", service_name)
            except Exception as e:
                logger.error("[%s] startup_hook 执行失败（不阻断启动）: %s", service_name, e)
        server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
        logger.info("[%s] HTTP 服务启动于 http://0.0.0.0:%d", service_name, port)
        server.serve_forever()

    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=default_port)
    args = parser.parse_args()
    port = args.port
    run(port)
