from __future__ import annotations

import base64
import html
import json
import os
import time
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response

from app.converter import (
    ConversionError,
    EXCLUDED_DOOR_MODELS,
    EXCLUDED_NAME_KEYWORDS,
    convert_excel_to_csv,
)


MAX_UPLOAD_BYTES = int(os.getenv("MAX_UPLOAD_BYTES", str(50 * 1024 * 1024)))

app = FastAPI(title="PMS 优化门扇清单拆分服务", version="1.0.0")


PAGE = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>PMS 门扇清单拆分</title>
  <style>
    body { font-family: Arial, "Microsoft YaHei", sans-serif; background:#f5f6f8; margin:0; }
    main { max-width:680px; margin:64px auto; background:#fff; padding:32px; border:1px solid #ddd; border-radius:12px; }
    h1 { margin-top:0; font-size:26px; }
    p { color:#555; line-height:1.7; }
    input { display:block; width:100%; box-sizing:border-box; padding:12px; border:1px solid #bbb; border-radius:8px; }
    button { margin-top:18px; padding:11px 24px; border:0; border-radius:8px; background:#1677ff; color:#fff; font-size:16px; cursor:pointer; }
    small { display:block; margin-top:20px; color:#777; }
  </style>
</head>
<body><main>
  <h1>PMS 门扇清单拆分</h1>
  <p>上传 PMS 优化门扇清单（.xls 或 .xlsx），系统会按工艺路线拆分皮行（数量翻倍、材料描述=厚度+板材）并下载 CSV。</p>
  <form action="/convert" method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xls,.xlsx,.csv" required>
    <button type="submit">拆分并下载 CSV</button>
  </form>
  <small>文件只在内存中处理，不会保存到服务器。</small>
</main></body></html>"""


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return PAGE


@app.get("/healthz")
@app.get("/health")
def healthz() -> dict[str, str]:
    return {"status": "ok", "service": "pms-door-split", "port": "8013"}


def build_summary_md(stats) -> str:
    normal = stats.normal_output_rows
    split = stats.split_output_rows
    split_lines = "\n".join(f"• {d}" for d in stats.split_details) or "无"
    excluded_count = len(stats.excluded_details)
    excluded_lines = "\n".join(f"• {d}" for d in stats.excluded_details) or "无"
    blank_count = len(stats.blank_thickness_orders)
    blank_lines = "\n".join(f"• {d}" for d in stats.blank_thickness_orders) or "无"
    rule_text = (
        f"共 {len(EXCLUDED_DOOR_MODELS)} 个型号"
        + (f" + 关键词「{'、'.join(EXCLUDED_NAME_KEYWORDS)}」" if EXCLUDED_NAME_KEYWORDS else "")
    )

    summary_md = (
        f"**@胡娅 请核对以下数据，核对无误后点击底部按钮：**\n\n"
        f"📊 **核对汇总**\n"
        f"• 源数据行数：{stats.source_rows}\n"
        f"• 输出数据行数：{stats.output_rows}\n"
        f"• 源数量合计：{stats.input_quantity_sum:g}\n"
        f"• 输出数量合计：{stats.quantity_sum:g}\n"
        f"• 正常行：{stats.source_rows - stats.split_source_rows} 条 → 输出 {normal} 行\n"
        f"• 拆分行：{stats.split_source_rows} 条 → 输出 {split} 行\n"
        f"• 校验：{normal} + {split} = {stats.output_rows} {'✅' if normal + split == stats.output_rows else '❌'}\n\n"
        f"🚫 **本次剔除门型：{excluded_count} 行（未进入生成文件）**\n{excluded_lines}\n\n"
        f"📑 **当前剔除门型表（{rule_text}，命中即不生成）**\n{'、'.join(EXCLUDED_DOOR_MODELS)}\n\n"
        f"⚠️ **工艺路线未写明厚度：{blank_count} 行（材料描述保留原值，请人工确认）**\n{blank_lines}\n\n"
        f"📋 **拆分明细**\n{split_lines}"
    )
    if stats.material_mismatch_details:
        mismatch_lines = "\n".join(f"• {d}" for d in stats.material_mismatch_details)
        summary_md += (
            f"\n\n🔴 **工艺与材料描述不一致：{len(stats.material_mismatch_details)} 行，请人工核对！**\n"
            f"{mismatch_lines}"
        )
    if stats.new_process_details:
        new_lines = "\n".join(f"• {d}" for d in stats.new_process_details)
        summary_md += (
            f"\n\n🔶 **未见过的工艺：{len(stats.new_process_details)} 行，请检查确认**\n"
            f"{new_lines}"
        )
    return summary_md


def build_warnings(stats) -> list[str]:
    warnings = []
    if stats.material_mismatch_details:
        warnings.append(
            f"🔴 工艺与材料描述不一致 {len(stats.material_mismatch_details)} 行：\n"
            + "\n".join(stats.material_mismatch_details)
        )
    if stats.new_process_details:
        warnings.append(
            f"🔶 未见过的工艺 {len(stats.new_process_details)} 行：\n"
            + "\n".join(stats.new_process_details)
        )
    return warnings


@app.post("/process")
async def process(req: Request) -> dict[str, Any]:
    """飞书网关兼容接口：JSON base64 输入 -> JSON base64 CSV 输出。"""
    start = time.time()
    body_bytes = await req.body()
    if not body_bytes:
        return {"success": False, "error": "请求体为空"}
    try:
        payload = json.loads(body_bytes)
    except json.JSONDecodeError:
        return {"success": False, "error": "请求体必须是 JSON"}

    filename = payload.get("filename", "upload.xlsx")
    file_content = payload.get("file_content")
    input_path = payload.get("input_path")

    if input_path:
        content = Path(input_path).read_bytes()
    elif file_content:
        content = base64.b64decode(file_content)
    else:
        return {"success": False, "error": "缺少 file_content 或 input_path"}

    if Path(filename).suffix.lower() not in {".xls", ".xlsx", ".csv"}:
        return {"success": False, "error": "只支持 .xls、.xlsx 或 .csv 文件"}

    try:
        csv_bytes, stats = convert_excel_to_csv(content, filename)
    except ConversionError as exc:
        return {"success": False, "error": html.escape(str(exc))}
    except RuntimeError as exc:
        return {"success": False, "error": html.escape(str(exc))}

    output_name = f"{Path(filename).stem}_拆分.csv"

    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "green",
            "title": {"tag": "plain_text", "content": "PMS 门扇清单拆分完成，请核对"},
        },
        "elements": [
            {
                "tag": "div",
                "text": {"tag": "lark_md", "content": build_summary_md(stats)},
            },
            {
                "tag": "action",
                "actions": [
                    {
                        "tag": "button",
                        "text": {"tag": "plain_text", "content": "核对无误"},
                        "type": "primary",
                        "value": {
                            "action": "confirm",
                            "service": "pms-door-split",
                            "file": filename,
                        },
                    }
                ],
            },
        ],
    }

    return {
        "success": True,
        "card": card,
        "output_files": {
            filename: [
                {
                    "filename": output_name,
                    "file_content": base64.b64encode(csv_bytes).decode("ascii"),
                }
            ]
        },
        "stats": {
            "source_rows": stats.source_rows,
            "output_rows": stats.output_rows,
            "quantity_sum": stats.quantity_sum,
            "split_source_rows": stats.split_source_rows,
            "excluded_rows": len(stats.excluded_details),
        },
        "warnings": build_warnings(stats),
        "cost_seconds": round(time.time() - start, 3),
    }


@app.post("/convert")
@app.post("/api/convert")
async def convert(file: UploadFile = File(...)) -> Response:
    filename = Path(file.filename or "upload.xlsx").name
    if Path(filename).suffix.lower() not in {".xls", ".xlsx", ".csv"}:
        raise HTTPException(status_code=400, detail="只支持 .xls、.xlsx 或 .csv 文件")

    content = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="上传文件过大")
    if not content:
        raise HTTPException(status_code=400, detail="上传文件为空")

    try:
        csv_bytes, stats = convert_excel_to_csv(content, filename)
    except ConversionError as exc:
        raise HTTPException(status_code=422, detail=html.escape(str(exc))) from exc
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=html.escape(str(exc))) from exc

    output_name = f"{Path(filename).stem}_拆分.csv"
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{quote(output_name)}",
        "X-Source-Rows": str(stats.source_rows),
        "X-Output-Rows": str(stats.output_rows),
        "X-Quantity-Sum": format(stats.quantity_sum, "g"),
    }
    return Response(content=csv_bytes, media_type="text/csv; charset=gb18030", headers=headers)
