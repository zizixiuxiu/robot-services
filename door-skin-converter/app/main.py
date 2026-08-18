from __future__ import annotations

import base64
import html
import json
import os
import re
import time
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response

from app.converter import (
    ConversionError,
    convert_excel_to_csv,
    get_excluded_keywords,
    get_excluded_models,
    load_runtime_exclusions,
    save_runtime_exclusions,
    _normalize_model,
)


MAX_UPLOAD_BYTES = int(os.getenv("MAX_UPLOAD_BYTES", str(50 * 1024 * 1024)))

app = FastAPI(title="门扇清单转换服务", version="1.1.0")


PAGE = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>门扇清单转换</title>
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
  <h1>门扇清单转换</h1>
  <p>上传门扇 Excel（.xls 或 .xlsx），系统会直接下载转换后的 CSV。</p>
  <form action="/convert" method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xls,.xlsx" required>
    <button type="submit">转换并下载 CSV</button>
  </form>
  <small>文件只在内存中处理，不会保存到服务器。</small>
</main></body></html>"""


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return PAGE


@app.get("/healthz")
@app.get("/health")
def healthz() -> dict[str, str]:
    return {"status": "ok", "service": "door-skin-converter", "port": "8010"}


MODEL_CODE_RE = re.compile(r"^[A-Za-z]{1,4}-?\d{2,4}$|^\d{4}$")


def _classify_item(item: str, kind: str) -> str:
    """判断剔除项是门型编号还是关键词（kind=auto 时按格式自动识别）。"""
    if kind in ("model", "keyword"):
        return kind
    return "model" if MODEL_CODE_RE.match(item) else "keyword"


@app.get("/exclusions")
def list_exclusions() -> dict:
    runtime = load_runtime_exclusions()
    return {
        "success": True,
        "models": get_excluded_models(),
        "keywords": get_excluded_keywords(),
        "runtime_models": runtime["models"],
        "runtime_keywords": runtime["keywords"],
    }


@app.post("/exclusions")
async def update_exclusions(req: Request) -> dict:
    """群里命令加/删剔除项。运行时清单只增删自己加过的项，内置型号需改代码发版。"""
    try:
        payload = await req.json()
    except Exception:
        return {"success": False, "error": "请求体必须是 JSON"}
    action = payload.get("action")
    kind = payload.get("kind", "auto")
    items = [str(i).strip() for i in payload.get("items", []) if str(i).strip()]
    if action not in ("add", "remove") or not items:
        return {"success": False, "error": "action 需为 add/remove 且 items 不能为空"}

    runtime = load_runtime_exclusions()
    added: dict[str, list[str]] = {"models": [], "keywords": []}
    removed: list[str] = []
    skipped: list[str] = []

    for item in items:
        if action == "add":
            if _classify_item(item, kind) == "model":
                if _normalize_model(item) in {_normalize_model(m) for m in get_excluded_models()}:
                    skipped.append(f"{item}（型号已在清单）")
                else:
                    runtime["models"].append(item)
                    added["models"].append(item)
            else:
                if item in get_excluded_keywords():
                    skipped.append(f"{item}（关键词已在清单）")
                else:
                    runtime["keywords"].append(item)
                    added["keywords"].append(item)
        else:
            hit = False
            norm = _normalize_model(item)
            for model in list(runtime["models"]):
                if _normalize_model(model) == norm:
                    runtime["models"].remove(model)
                    removed.append(model)
                    hit = True
            for keyword in list(runtime["keywords"]):
                if keyword == item:
                    runtime["keywords"].remove(keyword)
                    removed.append(keyword)
                    hit = True
            if not hit:
                skipped.append(f"{item}（不在可删除清单；内置项需改代码发版）")

    save_runtime_exclusions(runtime)
    return {
        "success": True,
        "action": action,
        "added": added,
        "removed": removed,
        "skipped": skipped,
        "models": get_excluded_models(),
        "keywords": get_excluded_keywords(),
    }


@app.post("/process")
async def process(req: Request) -> dict[str, Any]:
    """飞书网关兼容接口：JSON base64 输入 -> JSON base64 CSV 输出。"""
    start = time.time()
    body_bytes = await req.body()
    if not body_bytes:
        return {"success": False, "error": "请求体为空"}
    try:
        payload = json.loads(body_bytes)
    except json.JSONDecodeError as exc:
        return {"success": False, "error": "请求体必须是 JSON"}

    filename = payload.get("filename", "upload.xls")
    file_content = payload.get("file_content")
    input_path = payload.get("input_path")

    if input_path:
        content = Path(input_path).read_bytes()
    elif file_content:
        content = base64.b64decode(file_content)
    else:
        return {"success": False, "error": "缺少 file_content 或 input_path"}

    if Path(filename).suffix.lower() not in {".xls", ".xlsx"}:
        return {"success": False, "error": "只支持 .xls 或 .xlsx 文件"}

    try:
        csv_bytes, stats = convert_excel_to_csv(content, filename)
    except ConversionError as exc:
        return {"success": False, "error": html.escape(str(exc))}
    except RuntimeError as exc:
        return {"success": False, "error": html.escape(str(exc))}

    output_name = f"{Path(filename).stem}_转换.csv"

    normal = stats.normal_output_rows
    split = stats.split_output_rows
    split_lines = "\n".join(f"• {d}" for d in stats.split_details) or "无"
    excluded_count = len(stats.excluded_details)
    excluded_lines = "\n".join(f"• {d}" for d in stats.excluded_details) or "无"
    excluded_models = get_excluded_models()
    excluded_keywords = get_excluded_keywords()
    excluded_models_text = "、".join(excluded_models)
    excluded_rule_text = (
        f"共 {len(excluded_models)} 个型号"
        + (f" + 关键词「{'、'.join(excluded_keywords)}」" if excluded_keywords else "")
    )
    blank_count = len(stats.blank_thickness_orders)
    blank_lines = "\n".join(f"• {d}" for d in stats.blank_thickness_orders) or "无"

    summary_md = (
        f"**@刘佳 请核对以下数据，核对无误后点击底部按钮：**\n\n"
        f"📊 **核对汇总**\n"
        f"• 源数据行数：{stats.source_rows}\n"
        f"• 输出数据行数：{stats.output_rows}\n"
        f"• 源数量合计：{stats.input_quantity_sum:g}\n"
        f"• 输出数量合计：{stats.quantity_sum:g}\n"
        f"• 正常订单：{stats.source_rows - stats.split_source_rows} 条 → 输出 {normal} 行\n"
        f"• 拆分订单：{stats.split_source_rows} 条 → 输出 {split} 行\n"
        f"• 校验：{normal} + {split} = {stats.output_rows} {'✅' if normal + split == stats.output_rows else '❌'}\n"
        f"• 数量关系：输出 ÷ 源 = {stats.quantity_sum / stats.input_quantity_sum:.0f}\n\n"
        f"🚫 **本次剔除门型：{excluded_count} 行（未进入生成文件）**\n{excluded_lines}\n\n"
        f"📑 **当前剔除门型表（{excluded_rule_text}，命中即不生成）**\n{excluded_models_text}\n\n"
        f"⚠️ **工艺未写明厚度：{blank_count} 行（材料描述已用颜色兜底，请人工确认）**\n{blank_lines}\n\n"
        f"📋 **拆分明细**\n{split_lines}"
    )

    # 工艺列与材料描述不一致时在卡片里追加警告
    if stats.material_mismatch_details:
        mismatch_lines = "\n".join(f"• {d}" for d in stats.material_mismatch_details)
        summary_md += (
            f"\n\n🔴 **工艺与材料描述不一致：{len(stats.material_mismatch_details)} 行，请人工核对！**\n"
            f"{mismatch_lines}"
        )

    # 没见过的工艺也在卡片里提醒（材料描述仍按现有规则生成）
    if stats.new_process_details:
        new_lines = "\n".join(f"• {d}" for d in stats.new_process_details)
        summary_md += (
            f"\n\n🔶 **未见过的工艺：{len(stats.new_process_details)} 行，请检查确认**\n"
            f"{new_lines}"
        )

    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "green",
            "title": {"tag": "plain_text", "content": "门扇转换完成，请核对"},
        },
        "elements": [
            {
                "tag": "div",
                "text": {"tag": "lark_md", "content": summary_md},
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
                            "service": "door-skin-converter",
                            "file": filename,
                        },
                    }
                ],
            },
        ],
    }

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

    resp = {
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
        "warnings": warnings,
        "cost_seconds": round(time.time() - start, 3),
    }
    return resp


@app.post("/convert")
@app.post("/api/convert")
async def convert(file: UploadFile = File(...)) -> Response:
    filename = Path(file.filename or "upload.xls").name
    if Path(filename).suffix.lower() not in {".xls", ".xlsx"}:
        raise HTTPException(status_code=400, detail="只支持 .xls 或 .xlsx 文件")

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

    output_name = f"{Path(filename).stem}_转换.csv"
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{quote(output_name)}",
        "X-Source-Rows": str(stats.source_rows),
        "X-Output-Rows": str(stats.output_rows),
        "X-Quantity-Sum": format(stats.quantity_sum, "g"),
    }
    return Response(content=csv_bytes, media_type="text/csv; charset=gb18030", headers=headers)
