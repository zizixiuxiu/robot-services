from __future__ import annotations

import base64
import html
import io
import json
import os
import time
import zipfile
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response

from app.converter import ConversionError, ConversionStats, convert_workbook


MAX_UPLOAD_BYTES = int(os.getenv("MAX_UPLOAD_BYTES", str(50 * 1024 * 1024)))

app = FastAPI(title="门套窗套豪迈电子锯转换服务", version="1.0.0")


PAGE = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>门套窗套豪迈转换</title>
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
  <h1>门套窗套豪迈转换</h1>
  <p>上传门套窗套护墙汇总（.xls / .xlsx），系统会把「门套」「窗套」合并成豪迈 CSV，护墙原样另存，手裁丢弃。</p>
  <form action="/convert" method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xls,.xlsx,.csv" required>
    <button type="submit">转换并下载</button>
  </form>
  <small>有护墙时下载 zip（豪迈 CSV + 护墙表）。文件只在内存中处理，不会保存到服务器。</small>
</main></body></html>"""


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return PAGE


@app.get("/healthz")
@app.get("/health")
def healthz() -> dict[str, str]:
    return {"status": "ok", "service": "homag-frame-csv", "port": "8014"}


def build_summary_md(stats: ConversionStats) -> str:
    missing = "、".join(stats.missing_sheets) if stats.missing_sheets else "无"
    huqiang_line = (
        f"• 护墙：原样另存 {stats.huqiang_rows} 行"
        if stats.has_huqiang_file
        else "• 护墙：无（不另存文件）"
    )
    lines = [
        "📊 **核对汇总**",
        f"• 门套输出：{stats.door_rows} 行",
        f"• 窗套输出：{stats.window_rows} 行",
        f"• 合并豪迈 CSV：{stats.output_rows} 行",
        f"• 数量合计：{stats.quantity_sum:g}",
        f"• 校验：{stats.door_rows} + {stats.window_rows} = {stats.output_rows}"
        f" {'✅' if stats.door_rows + stats.window_rows == stats.output_rows else '❌'}",
        "",
        "📎 **其它表**",
        huqiang_line,
        f"• 手裁：丢弃 {stats.discarded_shoucai} 行",
        f"• 空行跳过：{stats.skipped_empty}",
        f"• 颜色无批次前缀（无 /）：{stats.color_no_slash} 行，订单号为空",
        f"• 缺失工作表：{missing}",
    ]
    return "\n".join(lines)


def build_warnings(stats: ConversionStats) -> list[str]:
    warnings: list[str] = []
    if stats.missing_sheets:
        warnings.append("未找到工作表：" + "、".join(stats.missing_sheets))
    if stats.color_no_slash:
        warnings.append(f"{stats.color_no_slash} 行颜色没有「批次/颜色」斜杠，订单号为空")
    return warnings


def _card(stats: ConversionStats) -> dict[str, Any]:
    return {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "green",
            "title": {"tag": "plain_text", "content": "门套窗套豪迈转换完成，请核对"},
        },
        "elements": [
            {
                "tag": "div",
                "text": {"tag": "lark_md", "content": build_summary_md(stats)},
            }
        ],
    }


@app.post("/process")
async def process(req: Request) -> dict[str, Any]:
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
        result = convert_workbook(content, filename)
    except ConversionError as exc:
        return {"success": False, "error": html.escape(str(exc))}
    except RuntimeError as exc:
        return {"success": False, "error": html.escape(str(exc))}

    stem = Path(filename).stem
    files = [
        {
            "filename": f"{stem}_豪迈.csv",
            "file_content": base64.b64encode(result.csv_bytes).decode("ascii"),
        }
    ]
    if result.huqiang_bytes:
        files.append(
            {
                "filename": f"{stem}_护墙.xlsx",
                "file_content": base64.b64encode(result.huqiang_bytes).decode("ascii"),
            }
        )
    stats = result.stats
    return {
        "success": True,
        "card": _card(stats),
        "output_files": {filename: files},
        "stats": {
            "door_rows": stats.door_rows,
            "window_rows": stats.window_rows,
            "output_rows": stats.output_rows,
            "quantity_sum": stats.quantity_sum,
            "huqiang_rows": stats.huqiang_rows,
            "discarded_shoucai": stats.discarded_shoucai,
            "color_no_slash": stats.color_no_slash,
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
        result = convert_workbook(content, filename)
    except ConversionError as exc:
        raise HTTPException(status_code=422, detail=html.escape(str(exc))) from exc
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=html.escape(str(exc))) from exc

    stem = Path(filename).stem
    csv_name = f"{stem}_豪迈.csv"
    stats = result.stats
    extra_headers = {
        "X-Door-Rows": str(stats.door_rows),
        "X-Window-Rows": str(stats.window_rows),
        "X-Output-Rows": str(stats.output_rows),
        "X-Quantity-Sum": format(stats.quantity_sum, "g"),
        "X-Huqiang-Rows": str(stats.huqiang_rows),
    }
    if not result.huqiang_bytes:
        extra_headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(csv_name)}"
        return Response(
            content=result.csv_bytes,
            media_type="text/csv; charset=gb18030",
            headers=extra_headers,
        )

    zip_name = f"{stem}_转换.zip"
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(csv_name, result.csv_bytes)
        zf.writestr(f"{stem}_护墙.xlsx", result.huqiang_bytes)
    extra_headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(zip_name)}"
    return Response(content=buffer.getvalue(), media_type="application/zip", headers=extra_headers)
