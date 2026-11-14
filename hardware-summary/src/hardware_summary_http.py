#!/usr/bin/env python3
"""
五金汇总 HTTP 服务（无 LLM，直接处理）
端口：8001
接口：POST /process  上传文件，直接返回处理结果
"""
import os
import sys
import time
from datetime import datetime
from pathlib import Path

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))
# 公共库 factory-common（Docker 中通过 ENV PYTHONPATH=/app 提供）
sys.path.insert(0, str(WORK_DIR.parent.parent / "factory-common"))

from factory_common.logging_utils import setup_logging
from factory_common.bom_utils import detect_file_type, sniff_file_format, xlsx_to_xls, get_clean_filename, detect_unusual_sheets
from factory_common.http_skeleton import run_process_server

logger = setup_logging("hardware-summary")


def _get_today_date():
    now = datetime.now()
    return f"{now.year}.{now.month}.{now.day}"


def _process_hardware_summary(input_path: str, output_dir: str, order_date: str = None) -> dict:
    """处理单个五金汇总"""
    logger.info("开始处理文件: %s -> %s", input_path, output_dir)
    ftype = detect_file_type(input_path)
    if ftype != 'hardware':
        err_msg = "文件类型错误：这是料单拆分文件，请发到报价料单处理群处理。" if ftype == 'order_split' else "文件类型错误：无法识别此文件，不是五金汇总格式。"
        logger.warning("文件类型不匹配: %s, type=%s", input_path, ftype)
        return {"success": False, "error": err_msg}

    original_input = input_path
    src_format = sniff_file_format(input_path)
    if src_format == 'xlsx':
        # 五金汇总仍需要 xls 供 xlrd 提取；工厂版直接用原 xlsx 保留格式
        xls_path = input_path.rsplit('.', 1)[0] + '.xls'
        xlsx_to_xls(input_path, xls_path)
        input_path = xls_path
        logger.info("xlsx 已转换为 xls（仅用于五金汇总提取）: %s", xls_path)

    clean_name = get_clean_filename(Path(input_path).name)
    if not order_date:
        order_date = _get_today_date()

    old_cwd = os.getcwd()
    os.chdir(str(WORK_DIR))
    try:
        from hardware.convert import convert_hardware_summary
        from hardware.hide_prices import generate_factory_version, generate_factory_version_xlsx
        t0 = time.time()
        convert_hardware_summary(
            input_path,
            os.path.join(output_dir, f"{clean_name}_五金汇总.xlsx")
        )
        logger.info("生成五金汇总完成, cost=%.3fs", time.time() - t0)
        t0 = time.time()
        if src_format == 'xlsx':
            factory_filename = f"{clean_name}_工厂版.xlsx"
            generate_factory_version_xlsx(
                original_input,
                os.path.join(output_dir, factory_filename),
                order_date
            )
        else:
            factory_filename = f"{clean_name}_工厂版.xls"
            generate_factory_version(
                original_input,
                os.path.join(output_dir, factory_filename),
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
        "files": [f"{clean_name}_五金汇总.xlsx", factory_filename],
        "type": "hardware",
        "note": f"工厂版下单日期已替换为 {order_date}",
    }

    warnings = []
    unusual = detect_unusual_sheets(input_path)
    if unusual:
        names = "、".join(unusual)
        warnings.append(f"⚠️ 提醒：检测到【{names}】sheet，此类内容未纳入五金汇总，请处理人员单独核对。")
        logger.warning("检测到非常规 sheet: %s", names)
    from hardware.convert import (
        check_hardware_quantity_totals,
        detect_hidden_source_sheets,
        format_hidden_sheet_warning,
        format_quantity_mismatch_warning,
    )
    hidden_msg = format_hidden_sheet_warning(detect_hidden_source_sheets(input_path))
    if hidden_msg:
        warnings.append(hidden_msg)
        logger.warning("隐藏表提醒: %s", hidden_msg)
    qty_msg = format_quantity_mismatch_warning(check_hardware_quantity_totals(input_path))
    if qty_msg:
        warnings.append(qty_msg)
        logger.warning("数量核对失败: %s", qty_msg)
    if warnings:
        result["warnings"] = warnings

    logger.info("文件处理完成: %s", input_path)
    return result


def _process_adapter(input_path: str, output_dir: str, req: dict) -> dict:
    """骨架适配：从请求中取 order_date"""
    return _process_hardware_summary(input_path, output_dir, req.get('order_date'))


if __name__ == '__main__':
    run_process_server(
        service_name="hardware-summary",
        default_port=8001,
        process_fn=_process_adapter,
        logger=logger,
        max_workers_env="HARDWARE_MAX_WORKERS",
        encode_workers_env="HARDWARE_ENCODE_MAX_WORKERS",
        default_output_base=WORK_DIR.parent / "data" / "output",
    )
