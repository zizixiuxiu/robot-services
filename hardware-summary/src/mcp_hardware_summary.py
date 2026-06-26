"""
五金汇总 MCP Server
支持：单文件处理 + 自动批量处理
通用流程：收到文件 → 校验是否为五金类型 → 处理/报错
生成：完整版五金汇总 + 工厂版（隐藏价格）
"""
import os
import sys
import json
import time
from datetime import datetime
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))

# 输出根目录：默认在项目 data/output，Docker 中通过环境变量覆盖为 /app/data/output
OUTPUT_BASE = Path(os.getenv("OUTPUT_BASE", str(WORK_DIR.parent / "data" / "output")))

from bom_utils import detect_file_type, xlsx_to_xls, get_clean_filename, suppress_output, detect_unusual_sheets
from bom_batch import (
    BomBatchManager, make_batch_tools,
    handle_upload_file, handle_process_batch,
    handle_check_task_status, handle_check_unprocessed_files,
    handle_get_output_files,
)

app = Server("hardware-summary")

# ============== 单文件处理函数 ==============

def _get_today_date():
    """获取当前日期，格式如 2026.5.23"""
    now = datetime.now()
    return f"{now.year}.{now.month}.{now.day}"

def _process_hardware_summary(input_path: str, output_dir: str, order_date: str = None) -> dict:
    """处理单个五金汇总"""
    # 校验文件类型
    ftype = detect_file_type(input_path)
    if ftype != 'hardware':
        if ftype == 'order_split':
            return {"success": False, "error": "文件类型错误：这是料单拆分文件，请发到报价料单处理群处理。"}
        return {"success": False, "error": "文件类型错误：无法识别此文件，不是五金汇总格式。"}

    # xlsx 转 xls
    if input_path.endswith('.xlsx'):
        xls_path = input_path.rsplit('.', 1)[0] + '.xls'
        xlsx_to_xls(input_path, xls_path)
        input_path = xls_path

    clean_name = get_clean_filename(Path(input_path).name)

    # 如果没有提供order_date，使用当前日期
    if not order_date:
        order_date = _get_today_date()

    old_cwd = os.getcwd()
    os.chdir(str(WORK_DIR))
    try:
        with suppress_output():
            from hardware.convert import convert_hardware_summary
            from hardware.hide_prices import generate_factory_version
            convert_hardware_summary(input_path, os.path.join(output_dir, f"{clean_name}_五金汇总.xlsx"))
            generate_factory_version(input_path, os.path.join(output_dir, f"{clean_name}_工厂版.xls"), order_date)
    except Exception as e:
        import traceback
        return {
            "success": False,
            "error": f"生成失败: {str(e)}",
            "trace": traceback.format_exc()
        }
    finally:
        os.chdir(old_cwd)

    result = {
        "success": True,
        "output_dir": output_dir,
        "count": 2,
        "files": [f"{clean_name}_五金汇总.xlsx", f"{clean_name}_工厂版.xls"],
        "type": "hardware",
        "note": f"工厂版下单日期已替换为 {order_date}（根据上传日期）",
        "_instruction": "五金汇总已完成，直接发送以上文件即可。"
    }

    unusual = detect_unusual_sheets(input_path)
    if unusual:
        names = "、".join(unusual)
        result["warning"] = f"⚠️ 提醒：检测到【{names}】sheet，此类内容未纳入五金汇总，请处理人员单独核对。"

    return result


# ============== 批量管理器 ==============

# 为批量模式包装处理函数（不需要 order_date 的版本）
def _batch_process_fn(input_path: str, output_dir: str) -> dict:
    return _process_hardware_summary(input_path, output_dir, order_date=None)

batch_manager = BomBatchManager(
    work_dir=WORK_DIR,
    process_fn=_batch_process_fn,
    max_concurrent=3
)
batch_manager.collection_window = 10


def _validate_hardware(file_path: str) -> tuple:
    """校验文件是否为五金类型"""
    ftype = detect_file_type(file_path)
    if ftype != 'hardware':
        if ftype == 'order_split':
            return False, "文件类型错误：这是料单拆分文件，请发到报价料单处理群处理。"
        return False, "文件类型错误：无法识别此文件，不是五金汇总格式。"
    return True, ""


# ============== 工具定义 ==============

@app.list_tools()
async def list_tools() -> list[Tool]:
    tools = [
        # 单文件处理（保留向后兼容）
        Tool(
            name="generate_hardware_summary",
            description="根据 xls 料单文件生成五金汇总（完整版+工厂版）。自动校验文件类型，非五金文件会报错。",
            inputSchema={
                "type": "object",
                "properties": {
                    "input_path": {
                        "type": "string",
                        "description": "本地 xls 文件绝对路径"
                    },
                    "file_content": {
                        "type": "string",
                        "description": "xls 文件 base64 编码内容（与 input_path 二选一）"
                    },
                    "filename": {
                        "type": "string",
                        "description": "当使用 file_content 时，指定原始文件名"
                    },
                    "order_date": {
                        "type": "string",
                        "description": "下单日期，格式如 2026.5.14，用于替换工厂版中的原日期"
                    }
                },
                "required": []
            }
        ),
    ]

    # 批量处理工具
    batch_tool_defs = make_batch_tools(
        "hardware_summary",
        "上传五金汇总文件（支持批量模式）。系统会自动收集窗口期内的所有文件并批量处理。自动校验文件类型，非五金文件会报错。"
    )
    for bt in batch_tool_defs:
        tools.append(Tool(
            name=bt["name"],
            description=bt["description"],
            inputSchema=bt["inputSchema"]
        ))

    return tools


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    if name == "generate_hardware_summary":
        return await _handle_generate_hardware_summary(arguments)

    # 批量工具路由
    if name == "hardware_summary_upload_file":
        return await handle_upload_file(batch_manager, arguments, _validate_hardware)
    elif name == "hardware_summary_process_batch":
        return await handle_process_batch(batch_manager)
    elif name == "hardware_summary_check_task_status":
        return await handle_check_task_status(batch_manager, arguments)
    elif name == "hardware_summary_check_unprocessed_files":
        return await handle_check_unprocessed_files(batch_manager, arguments)
    elif name == "hardware_summary_get_output_files":
        return await handle_get_output_files(batch_manager)

    return [TextContent(type="text", text=json.dumps({"error": f"未知工具: {name}"}, ensure_ascii=False))]


# ============== 单文件处理（向后兼容） ==============

async def _handle_generate_hardware_summary(arguments: dict) -> list[TextContent]:
    t0 = time.time()

    input_path = arguments.get("input_path")
    file_content = arguments.get("file_content")
    filename = arguments.get("filename", "input.xls")
    order_date = arguments.get("order_date")

    # base64 上传
    tmpdir = None
    if file_content and not input_path:
        import tempfile, base64
        tmpdir = tempfile.mkdtemp()
        input_path = os.path.join(tmpdir, filename)
        with open(input_path, "wb") as f:
            f.write(base64.b64decode(file_content))

    if not input_path or not os.path.exists(input_path):
        if tmpdir:
            import shutil; shutil.rmtree(tmpdir, ignore_errors=True)
        return [TextContent(type="text", text=json.dumps({"success": False, "error": "文件不存在"}, ensure_ascii=False))]

    # 创建输出目录
    base = Path(input_path).stem
    timestamp = str(int(time.time()))
    output_dir = str(OUTPUT_BASE / f"output_{base}_{timestamp}")
    os.makedirs(output_dir, exist_ok=True)

    result = _process_hardware_summary(input_path, output_dir, order_date)

    if tmpdir:
        import shutil; shutil.rmtree(tmpdir, ignore_errors=True)

    cost = time.time() - t0
    if result.get("success"):
        result["cost_seconds"] = round(cost, 3)

    return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
