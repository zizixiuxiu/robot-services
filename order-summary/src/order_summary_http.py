#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PVC 订单汇总 HTTP 服务（Docker / Linux 兼容）
端口：8011
接口：
  GET  /health
  POST /process
"""
import os
import sys
from pathlib import Path

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))
# 公共库 factory-common（Docker 中通过 ENV PYTHONPATH=/app 提供）
sys.path.insert(0, str(WORK_DIR.parent.parent / "factory-common"))

from summary_core import summarize_order
from factory_common.logging_utils import setup_logging
from factory_common.http_skeleton import run_process_server

# ==================== 日志配置（公共库） ====================
logger = setup_logging("order-summary")
# ==================================================


def process_order_summary(input_path: str, output_dir: str, req: dict) -> dict:
    """骨架适配：单文件处理，输出单个汇总 xlsx，由骨架自动编码返回。"""
    filename = req.get("filename", "input.xls")
    logger.info("处理文件: %s", filename)
    try:
        result = summarize_order(input_path, output_dir)
        logger.info(
            "汇总完成: %s, 组数=%d, 数量合计=%d",
            filename, result["groups"], result["quantity_total"],
        )
        return {
            "success": True,
            "filename": filename,
            "groups": result["groups"],
            "quantity_total": result["quantity_total"],
            "files": [Path(result["output_path"]).name],
        }
    except Exception as e:
        import traceback
        logger.exception("处理文件失败: %s", filename)
        return {
            "success": False,
            "error": f"处理失败: {str(e)}",
            "trace": traceback.format_exc(),
        }


if __name__ == "__main__":
    run_process_server(
        service_name="order-summary",
        default_port=8011,
        process_fn=process_order_summary,
        logger=logger,
        max_workers_env="ORDER_SUMMARY_MAX_WORKERS",
        default_max_workers=1,
        default_output_base=WORK_DIR.parent / "data" / "output",
        missing_file_error="缺少 file_content",
    )
