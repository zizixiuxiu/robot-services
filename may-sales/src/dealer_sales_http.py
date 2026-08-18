#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
销售部业绩核对表 HTTP 服务（完全接入 factory-common 骨架）
端口：8003（替换原经销商销售服务）
调用 generate_may_sales_report.js 处理 Excel 文件
支持 Windows 本机运行和 Docker 容器运行
"""
import os
import sys
import time
import base64
import subprocess
from pathlib import Path

# 工作目录
WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
SCRIPT_PATH = WORK_DIR / "generate_may_sales_report.js"
NODE_EXE = "node"

# 公共库 factory-common（Docker 中通过 ENV PYTHONPATH=/app 提供）
sys.path.insert(0, str(WORK_DIR.parent.parent / "factory-common"))
from factory_common.logging_utils import setup_logging
from factory_common.http_skeleton import run_process_server

# 默认模板路径；Docker 中通过环境变量 DEFAULT_TEMPLATE 覆盖为 /app/templates/...
DEFAULT_TEMPLATE = Path(os.getenv(
    "DEFAULT_TEMPLATE",
    str(WORK_DIR / "templates" / "2026年5月销售部业绩核对表 - 副本.xlsx")
))

# 奢匠销量情况报表：脚本与固定模板（清空 2026 月度数据、保留公式）
SALES_SCRIPT_PATH = WORK_DIR / "generate_dealer_sales_report.js"
SALES_TEMPLATE = Path(os.getenv(
    "SALES_TEMPLATE",
    str(WORK_DIR / "templates" / "2026年奢匠各经销商销量情况_模板.xlsx")
))
# 销量累积表（2026 月度数据，每月递增，持久化在 data/sales_state）
SALES_ACCUMULATOR = Path(os.getenv(
    "SALES_ACCUMULATOR",
    str(WORK_DIR.parent / "data" / "sales_state" / "销量累计_2026.xlsx")
))

# 输出目录；默认在项目 data/output_may_sales，Docker 中通过环境变量覆盖为 /app/output_may_sales
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", str(WORK_DIR.parent / "data" / "output_may_sales")))

# ==================== 日志配置（公共库） ====================
logger = setup_logging("dealer-sales")
# ==================================================

# 缺文件/非多文件请求的统一错误文案（与原 Handler 一致）
MISSING_FILES_ERROR = "需要至少 3 个文件：综合查询、联思系统、奢匠下单统计"


def _detect_file_type(filename: str) -> str:
    """根据文件名识别文件类型"""
    name = filename.lower()
    if '综合查询' in name:
        return 'zhcx'
    if '联思' in name:
        return 'liansi'
    if '奢匠' in name or '下单统计' in name or '线下' in name:
        return 'shejiang'
    if '核对表' in name or '待核对' in name or '模板' in name:
        return 'template'
    return 'unknown'


def _detect_month_from_filenames(filenames: list) -> str:
    """从文件名中提取月份（如 6月），返回数字字符串"""
    import re
    for fn in filenames:
        m = re.search(r'(\d+)\s*月', fn)
        if m:
            return str(int(m.group(1)))
    return ''


def run_may_sales(zhcx_path: str, liansi_path: str, shejiang_path: str, template_path: str, output_path: str, month: str = '') -> dict:
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
    if month:
        cmd.append(month)
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


def run_dealer_sales(shejiang_path: str, liansi_path: str, zhcx_path: str, month: str, output_dir: str) -> dict:
    """调用 generate_dealer_sales_report.js 生成 汇总表_提取结果 + 奢匠各经销商销量情况"""
    if not SALES_SCRIPT_PATH.exists():
        return {"success": False, "error": f"找不到脚本: {SALES_SCRIPT_PATH}"}
    if not SALES_TEMPLATE.exists():
        return {"success": False, "error": f"销量情况模板不存在: {SALES_TEMPLATE}"}
    if not SALES_ACCUMULATOR.exists():
        return {"success": False, "error": f"销量累积表不存在: {SALES_ACCUMULATOR}"}

    cmd = [NODE_EXE, str(SALES_SCRIPT_PATH), str(month), shejiang_path, liansi_path, zhcx_path,
           str(SALES_TEMPLATE), output_dir, str(SALES_ACCUMULATOR)]
    logger.info("调用 Node.js(销量情况): %s", " ".join(cmd))
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=300, cwd=str(WORK_DIR),
        )
    except subprocess.TimeoutExpired:
        return {"success": False, "error": "销量情况报表处理超时（超过 5 分钟）"}
    except Exception as e:
        logger.exception("调用 Node.js(销量情况) 异常")
        return {"success": False, "error": f"调用 Node.js(销量情况) 异常: {e}"}

    if result.returncode != 0:
        logger.error("销量情况 Node.js 执行失败: code=%d, stderr=%s", result.returncode, result.stderr)
        return {"success": False, "error": f"销量情况报表生成失败 (code={result.returncode}): {result.stderr[-500:]}"}

    outputs = []
    for name in (f"汇总表_提取结果_{month}月.xlsx", f"2026年奢匠各经销商销量情况（{month}月）.xlsx"):
        p = os.path.join(output_dir, name)
        if os.path.exists(p):
            outputs.append({"path": p, "filename": name})
    if len(outputs) != 2:
        return {"success": False, "error": "销量情况输出文件未生成", "stdout": result.stdout, "stderr": result.stderr}
    return {"success": True, "outputs": outputs, "stdout": result.stdout}


def _process_multi(req: dict, work_dir: str) -> dict:
    """多文件整体任务（{files:[...]}）：骨架已创建 work_dir 临时目录并负责清理"""
    files = req.get('files', [])
    order_date = req.get('order_date')
    month = _detect_month_from_filenames([f.get('filename', '') for f in files])
    if not month and order_date:
        try:
            month = str(int(order_date.split('.')[1]))
        except Exception:
            month = '5'
    if not month:
        month = '5'
    logger.info("收到 %s月业绩核对请求，文件数=%d, order_date=%s", month, len(files), order_date)
    if not files or len(files) < 3:
        return {"success": False, "error": MISSING_FILES_ERROR}

    # 临时目录由骨架提供
    tmpdir = work_dir
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
        return {"success": False, "error": f"缺少必要文件，无法识别: {missing}。文件名需包含：综合查询、联思、奢匠/下单统计"}

    # 模板文件
    if not template_path:
        if DEFAULT_TEMPLATE.exists():
            template_path = str(DEFAULT_TEMPLATE)
            logger.info("使用默认模板: %s", template_path)
        else:
            logger.error("未上传模板文件，且默认模板不存在")
            return {"success": False, "error": "未上传模板文件，且默认模板不存在"}

    # 输出路径
    if not OUTPUT_DIR.exists():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    output_filename = f'2026年{month}月销售部业绩核对表.xlsx'
    output_path = str(OUTPUT_DIR / output_filename)

    t0 = time.time()
    result = run_may_sales(
        file_map['zhcx'],
        file_map['liansi'],
        file_map['shejiang'],
        template_path,
        output_path,
        month,
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

        # 追加：奢匠各经销商销量情况报表（同一组输入文件 + 固定模板）
        # 失败不阻断主报表，仅以 warning 形式透传到群
        sales_result = run_dealer_sales(
            file_map['shejiang'], file_map['liansi'], file_map['zhcx'], month, str(OUTPUT_DIR),
        )
        if sales_result.get('success'):
            for item in sales_result['outputs']:
                try:
                    with open(item['path'], 'rb') as fh:
                        item['file_content'] = base64.b64encode(fh.read()).decode('utf-8')
                    result['output_files'].append(item)
                    logger.info("销量情况输出已编码: %s", item['filename'])
                except Exception as e:
                    logger.error("读取销量情况输出失败: %s", e)
                    result['output_files'].append(item)
            if sales_result.get('stdout') and '直营店数据按 0 处理' in sales_result['stdout']:
                result['warning'] = '⚠️ 奢匠明细文件缺少「直营店+电商下单表」Sheet，本月直营店/电商数据未计入销量情况报表'
        else:
            logger.error("销量情况报表生成失败: %s", sales_result.get('error'))
            result['warning'] = f"⚠️ 销量情况报表生成失败：{sales_result.get('error')}（业绩核对表不受影响）"

    logger.info("%s月业绩核对处理完成，success=%s", month, result.get('success'))
    return result


def _process_adapter(input_path: str, output_dir: str, req: dict) -> dict:
    """单文件请求：本服务只支持 {files:[...]} 多文件整体任务，与原 Handler 返回同一错误文案"""
    return {"success": False, "error": MISSING_FILES_ERROR}


if __name__ == '__main__':
    run_process_server(
        service_name="may-sales",
        default_port=8003,
        process_fn=_process_adapter,
        logger=logger,
        max_workers_env="MAY_SALES_MAX_WORKERS",
        default_output_base=OUTPUT_DIR,
        multi_file_fn=_process_multi,
        missing_file_error=MISSING_FILES_ERROR,
    )
