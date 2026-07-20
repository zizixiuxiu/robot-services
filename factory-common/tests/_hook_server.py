# -*- coding: utf-8 -*-
"""test_skeleton_hooks 用的测试服务器启动器（子进程方式运行）。

用法: python _hook_server.py '<json 配置>'
配置字段:
  port, output_base,
  aliases: ["..."]      -> path_aliases
  missing: "..."        -> missing_file_error
  get_paths: [...] | null -> get_paths
  batch_agg: bool       -> batch_error_aggregate
  hook: bool            -> batch_aggregate_hook
  startup: bool         -> startup_hook（写 flag 文件）
  startup_flag_file     -> startup_hook 写入的文件
  multi: bool           -> multi_file_fn
  workdir_file          -> multi_file_fn 记录 work_dir 路径的文件
"""
import json
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))  # factory-common 根目录

from factory_common.logging_utils import setup_logging
from factory_common.http_skeleton import run_process_server

logger = setup_logging("skeleton-test")
cfg = json.loads(sys.argv[1])
sys.argv = [sys.argv[0]]  # 避免骨架内部 argparse 误解析本启动器的 JSON 参数


def startup():
    Path(cfg["startup_flag_file"]).write_text("started", encoding="utf-8")


def process_fn(input_path, output_dir, req):
    name = Path(input_path).name
    if "fail" in name:
        return {"success": False, "error": "模拟处理失败"}
    if req.get("self_encode"):
        # 自行构造 output_files，骨架应跳过自动编码
        return {"success": True,
                "output_files": {"custom": [{"filename": "z.zip", "file_content": "WklQ"}]}}
    out = Path(output_dir) / "out.txt"
    out.write_text("hello", encoding="utf-8")
    return {"success": True, "files": ["out.txt"]}


def hook(response, results, files):
    response["quantity_total"] = sum(1 for r in results if r.get("success"))
    response["quantity_files"] = len(files)


def multi_fn(req, work_dir):
    Path(cfg["workdir_file"]).write_text(
        json.dumps({"work_dir": work_dir, "existed": os.path.isdir(work_dir)}),
        encoding="utf-8")
    # 返回扁平 output_files list，且不返回 cost_seconds（骨架应补上）
    return {"success": True,
            "count": len(req.get("files", [])),
            "output_files": [{"filename": "merged.csv", "file_content": "QUJD"}]}


kwargs = dict(
    service_name="skeleton-test",
    default_port=cfg["port"],
    process_fn=process_fn,
    logger=logger,
    max_workers_env="TEST_MAX_WORKERS",
    default_max_workers=1,
    default_output_base=cfg["output_base"],
)
if cfg.get("aliases"):
    kwargs["path_aliases"] = tuple(cfg["aliases"])
if cfg.get("missing"):
    kwargs["missing_file_error"] = cfg["missing"]
if cfg.get("get_paths") is not None:
    kwargs["get_paths"] = tuple(cfg["get_paths"])
if cfg.get("batch_agg"):
    kwargs["batch_error_aggregate"] = True
if cfg.get("hook"):
    kwargs["batch_aggregate_hook"] = hook
if cfg.get("startup"):
    kwargs["startup_hook"] = startup
if cfg.get("multi"):
    kwargs["multi_file_fn"] = multi_fn

run_process_server(**kwargs)
