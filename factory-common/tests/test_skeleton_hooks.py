# -*- coding: utf-8 -*-
"""http_skeleton 扩展钩子自测（公共库首个测试）。

起 3 个子进程服务实例，用 urllib 发请求断言响应：
  - full(18301):  path_aliases / not_found 文案 / missing_file_error /
                  get_paths 白名单 / batch_error_aggregate / batch_aggregate_hook /
                  startup_hook / process_fn 自带 output_files 跳过编码
  - multi(18302): multi_file_fn 整体任务模式（扁平 output_files、cost_seconds 补齐、
                  work_dir 创建与清理、单文件请求仍走 process_fn）
  - plain(18303): 默认参数回归（404 文案、任何 GET ok、"文件不存在"、自动编码）

运行: python test_skeleton_hooks.py   （全过则退出码 0）
"""
import base64
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
import urllib.error
from pathlib import Path

HERE = Path(__file__).resolve().parent
B64 = base64.b64encode(b"dummy").decode()
PORT_FULL, PORT_MULTI, PORT_PLAIN = 18301, 18302, 18303

FAILURES = []


def check(name, cond, detail=""):
    status = "PASS" if cond else "FAIL"
    print(f"[{status}] {name}" + (f"  ({detail})" if detail and not cond else ""))
    if not cond:
        FAILURES.append(f"{name}: {detail}")


def req(method, port, path, payload=None):
    url = f"http://127.0.0.1:{port}{path}"
    data = json.dumps(payload).encode() if payload is not None else None
    r = urllib.request.Request(url, data=data, method=method,
                               headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(r, timeout=10) as resp:
            return resp.status, json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read().decode())


def wait_up(port, timeout=15):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/health", timeout=2):
                return True
        except Exception:
            time.sleep(0.3)
    return False


def main():
    tmp = Path(tempfile.mkdtemp(prefix="skeleton_test_"))
    log_dir = tmp / "logs"
    log_dir.mkdir()
    env = dict(os.environ, LOG_DIR=str(log_dir))
    out_base = str(tmp / "output")
    flag_file = str(tmp / "startup.flag")
    workdir_file = str(tmp / "workdir.json")

    servers = [
        (PORT_FULL, {"port": PORT_FULL, "output_base": out_base,
                     "aliases": ["/convert"], "missing": "缺少 file_content",
                     "get_paths": ["/", "/health"], "batch_agg": True, "hook": True,
                     "startup": True, "startup_flag_file": flag_file}),
        (PORT_MULTI, {"port": PORT_MULTI, "output_base": out_base,
                      "multi": True, "workdir_file": workdir_file}),
        (PORT_PLAIN, {"port": PORT_PLAIN, "output_base": out_base}),
    ]
    procs = []
    try:
        for port, cfg in servers:
            p = subprocess.Popen(
                [sys.executable, str(HERE / "_hook_server.py"), json.dumps(cfg)],
                env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            procs.append(p)
        for port, _ in servers:
            if not wait_up(port):
                print(f"服务 {port} 启动失败")
                return 1

        # ---------- full: path_aliases + 404 文案 ----------
        st, body = req("POST", PORT_FULL, "/convert",
                       {"file_content": B64, "filename": "a.xls"})
        check("path_aliases: POST /convert 可用", st == 200 and body.get("success"))
        st, body = req("POST", PORT_FULL, "/nope", {})
        check("path_aliases: 404 文案自动包含别名",
              st == 404 and body.get("error") == "未知接口，请用 POST /process 或 POST /convert",
              str(body))

        # ---------- full: missing_file_error ----------
        st, body = req("POST", PORT_FULL, "/process", {})
        check("missing_file_error: 文案为 缺少 file_content",
              body.get("error") == "缺少 file_content", str(body))

        # ---------- full: get_paths 白名单 ----------
        st, body = req("GET", PORT_FULL, "/health")
        check("get_paths: /health 200", st == 200 and body.get("status") == "ok")
        st, body = req("GET", PORT_FULL, "/random")
        check("get_paths: 白名单外 404 未知接口",
              st == 404 and body.get("error") == "未知接口", f"st={st} body={body}")

        # ---------- full: startup_hook ----------
        check("startup_hook: 已被调用", Path(flag_file).exists())

        # ---------- full: batch_error_aggregate + batch_aggregate_hook ----------
        payload = {"files": [
            {"file_content": B64, "filename": "good.xls"},
            {"file_content": B64, "filename": "fail.xls"},
            {"filename": "nofile.xls"},
        ]}
        st, body = req("POST", PORT_FULL, "/process", payload)
        exp_err = "[fail.xls] 模拟处理失败; [nofile.xls] 缺少 file_content"
        check("batch_error_aggregate: 顶层 error 格式",
              body.get("error") == exp_err, str(body.get("error")))
        check("batch_error_aggregate: 与现有字段共存",
              body.get("batch") is True and len(body.get("results", [])) == 3)
        check("batch_aggregate_hook: quantity_total/quantity_files",
              body.get("quantity_total") == 1 and body.get("quantity_files") == 3,
              str(body))

        # ---------- full: process_fn 自带 output_files 跳过编码 ----------
        st, body = req("POST", PORT_FULL, "/process",
                       {"file_content": B64, "filename": "a.xls", "self_encode": True})
        check("output_files 跳过编码: 原样透传",
              body.get("output_files") == {"custom": [{"filename": "z.zip",
                                                       "file_content": "WklQ"}]},
              str(body.get("output_files")))

        # ---------- multi: multi_file_fn ----------
        st, body = req("POST", PORT_MULTI, "/process",
                       {"files": [{"file_content": B64, "filename": "a.csv"},
                                  {"file_content": B64, "filename": "b.csv"},
                                  {"file_content": B64, "filename": "c.csv"}]})
        check("multi_file_fn: 返回 dict 直接作为响应",
              body.get("success") and body.get("count") == 3
              and body.get("output_files") == [{"filename": "merged.csv",
                                                "file_content": "QUJD"}],
              str(body))
        check("multi_file_fn: 骨架补 cost_seconds", "cost_seconds" in body)
        wd = json.loads(Path(workdir_file).read_text(encoding="utf-8"))
        check("multi_file_fn: work_dir 调用时存在、善后已清理",
              wd["existed"] and not os.path.exists(wd["work_dir"]), str(wd))
        # 单文件请求仍走 _process_single/process_fn
        st, body = req("POST", PORT_MULTI, "/process",
                       {"file_content": B64, "filename": "single.xls"})
        check("multi_file_fn: 单文件仍走 process_fn 并自动编码",
              body.get("success") and "single.xls" in body.get("output_files", {}),
              str(body.get("output_files")))

        # ---------- plain: 默认行为回归 ----------
        st, body = req("POST", PORT_PLAIN, "/nope", {})
        check("plain: 404 文案不变",
              st == 404 and body.get("error") == "未知接口，请用 POST /process", str(body))
        st, body = req("GET", PORT_PLAIN, "/anything")
        check("plain: 任何 GET 返回 ok", st == 200 and body.get("status") == "ok")
        st, body = req("POST", PORT_PLAIN, "/process", {})
        check("plain: 缺文件文案 文件不存在", body.get("error") == "文件不存在")
        st, body = req("POST", PORT_PLAIN, "/process",
                       {"file_content": B64, "filename": "in.xls"})
        of = body.get("output_files", {}).get("in.xls", [])
        ok = (body.get("success") and len(of) == 1 and of[0]["filename"] == "out.txt"
              and of[0]["file_content"] == base64.b64encode(b"hello").decode())
        check("plain: 自动编码 output_files pairs", ok, str(body.get("output_files")))
        check("plain: 无顶层 error 字段", "error" not in body)
    finally:
        for p in procs:
            p.terminate()
        for p in procs:
            p.wait(timeout=5)

    print()
    if FAILURES:
        print(f"失败 {len(FAILURES)} 项:")
        for f in FAILURES:
            print("  -", f)
        return 1
    print("全部通过 ✔")
    return 0


if __name__ == "__main__":
    sys.exit(main())
