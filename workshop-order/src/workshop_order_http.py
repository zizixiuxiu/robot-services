#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
下车间单转换 HTTP 服务
端口：8006
接口：POST /process（兼容 POST /convert）

完全接入 factory_common.http_skeleton：本文件只保留业务函数、薄适配与
run_process_server(...) 调用。协议差异（/convert 别名、缺文件文案、批量
error 聚合、LibreOffice 预热）全部通过骨架扩展参数表达。

批量：网关 10 秒窗口收集多文件后一次 POST {files:[...]}；成功 ≥2 个时
batch_aggregate_hook 打成一个 zip 回传，单文件仍直接回传 xlsx。
"""
import base64
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from datetime import datetime
from pathlib import Path
from urllib import request as urllib_request

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(WORK_DIR))
# 公共库 factory-common（Docker 中通过 ENV PYTHONPATH=/app 提供）
sys.path.insert(0, str(WORK_DIR.parent.parent / "factory-common"))

from factory_common.logging_utils import setup_logging
from factory_common.http_skeleton import run_process_server, _encode_output_file

from make_workshop_order import transform, convert_xls_to_xlsx, warm_xls_converter
import change_audit
import make_change_notice

logger = setup_logging("workshop-order")

VALID_ORDER_TYPES = {"auto", "tiepi", "hunyou"}
_INVALID_NAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_DUP_SUFFIX = re.compile(r"\s*\(\d+\)$")
_ARCHIVE_START = datetime(2026, 9, 1).date()
# 工分补填原来按表逐条报，改动说明里已经汇总过，群里不必再刷一屏
_WORKPOINT_WARNING = re.compile(r"^工作表 '.*' 补填了 \d+ 个工分公式$")


def archive_root() -> Path | None:
    raw = (os.getenv("WORKSHOP_ARCHIVE_ROOT") or "").strip()
    return Path(raw) if raw else None


def sanitize_uploader_name(name: str) -> str:
    cleaned = _INVALID_NAME_CHARS.sub("", (name or "").strip())
    cleaned = cleaned.strip(" .")
    if cleaned == "未知上传者":
        return ""
    return cleaned


def canonical_archive_name(filename: str) -> str:
    """去掉 Windows/WPS 重传后缀：'下车间 (1).xlsx' / '贴皮(2).xls' → 原名。"""
    path = Path(filename)
    return f"{_DUP_SUFFIX.sub('', path.stem)}{path.suffix}"


def _copy_into(src: Path, dest_dir: Path, filename: str) -> Path | None:
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / canonical_archive_name(filename)
    shutil.copy2(src, dest)
    for leftover in dest_dir.iterdir():
        if leftover.is_file() and leftover != dest and canonical_archive_name(leftover.name) == dest.name:
            leftover.unlink()
    return dest


def archive_workshop_pair(
    original_path: str | Path,
    original_name: str,
    processed_path: str | Path,
    processed_name: str,
    uploader_name: str,
    now: datetime | None = None,
    root: Path | None = None,
) -> Path | None:
    """按年月+上传者姓名归档：{root}/{年份}/{月份}/{姓名}/原文件 、 .../{处理后}。

    年月均取处理日期（不补零，如 2026/9、2027/1）；9.1 起按姓名归档，
    10.4 起加月份层，10.5 起加年份层（这样跨年不用人工改配置）。
    """
    dest_root = root if root is not None else archive_root()
    if dest_root is None:
        return None
    when = (now or datetime.now()).date()
    if when < _ARCHIVE_START:
        return None
    original = Path(original_path)
    processed = Path(processed_path)
    if not original.is_file() or not processed.is_file():
        logger.warning("归档跳过，源文件不存在 original=%s processed=%s", original, processed)
        return None
    person = sanitize_uploader_name(uploader_name)
    if not person:
        logger.info("归档跳过，无飞书上传者姓名 filename=%s", original_name)
        return None
    person_dir = dest_root / str(when.year) / str(when.month) / person
    original_dir = person_dir / "原文件"
    processed_dir = person_dir / "处理后"
    saved_original = _copy_into(original, original_dir, Path(original_name).name or original.name)
    saved_processed = _copy_into(processed, processed_dir, Path(processed_name).name or processed.name)
    logger.info("已归档: %s / %s → %s", saved_original.name, saved_processed.name, person_dir)
    _notify_host_sync()
    return person_dir


def _notify_host_sync() -> None:
    url = (os.getenv("WORKSHOP_ARCHIVE_SYNC_URL") or "").strip()
    if not url:
        return
    try:
        req = urllib_request.Request(
            url,
            data=b"{}",
            method="POST",
            headers={"Content-Type": "application/json"},
        )
        opener = urllib_request.build_opener(urllib_request.ProxyHandler({}))
        with opener.open(req, timeout=5) as resp:
            resp.read()
    except Exception:
        logger.warning("通知宿主机同步共享盘失败", exc_info=True)


def _convert_xls_to_xlsx(input_path: Path) -> Path:
    """将 .xls 转换为 .xlsx（Docker 内使用 LibreOffice 保留格式）。"""
    if input_path.suffix.lower() != ".xls":
        return input_path

    xlsx_path = input_path.with_suffix(".xlsx")
    logger.info("xls 转 xlsx: %s -> %s", input_path, xlsx_path)
    t0 = time.time()
    convert_xls_to_xlsx(input_path, xlsx_path)
    logger.info("xls 转 xlsx 完成, cost=%.3fs", time.time() - t0)
    return xlsx_path


def build_change_notice(baseline: Path, output: Path, filename: str) -> dict:
    """生成「这份文件改了什么」，随结果一起回给群里。

    用转换后的 .xlsx 当基线逐格比对——它就是生成时的实际输入，
    天然同源，不会把文件版本差异误报成改动。

    返回 {text, brief, review}：text 是逐格长说明（日志/清单用），
    brief+review 是给群里的两行摘要（单号 · 改了什么 / ⚠ 核对几处）。
    """
    reports = change_audit.audit(baseline, output)
    payload = change_audit.to_payload(reports)
    items = [c for r in payload for c in r["cell_changes"]]
    stats = make_change_notice.summarize(items, payload)
    hint = make_change_notice.order_hint(Path(filename).stem)
    brief, review = make_change_notice.build_brief(items, stats, hint)
    return {
        "text": make_change_notice.build_text(items, stats, hint),
        "brief": brief,
        "review": review,
    }


def process(input_path: str, output_dir: str, req: dict) -> dict:
    """骨架 process_fn：单文件下车间单转换。

    返回 dict 已含 output_files（自行编码，保持原协议字段），骨架跳过自动编码。
    """
    filename = req.get("filename", "input.xlsx")
    order_type = req.get("order_type", req.get("order-type", "auto"))

    if order_type not in VALID_ORDER_TYPES:
        return {
            "success": False,
            "error": f"order_type 必须是 auto/tiepi/hunyou 之一，收到: {order_type}",
        }

    try:
        # .xls 需先转换为 .xlsx
        converted_path = _convert_xls_to_xlsx(Path(input_path))

        base = Path(filename).stem
        output_name = f"{base}下车间.xlsx"
        output_path = Path(output_dir) / output_name

        t0 = time.time()
        stats = transform(
            converted_path,
            output_path,
            discount=float(os.getenv("WORKSHOP_DISCOUNT", "0.8")),
            order_type=order_type,
        )
        logger.info("转换完成, cost=%.3fs, stats=%s", time.time() - t0, stats)

        if not output_path.exists():
            return {"success": False, "error": "输出文件未生成"}

        # 每处理完一单，自动附一段"改了什么"，让群里不用再人工解释
        notice = ""
        notice_brief = ""
        notice_review = ""
        try:
            t1 = time.time()
            info = build_change_notice(converted_path, output_path, filename)
            notice = info.get("text") or ""
            notice_brief = info.get("brief") or ""
            notice_review = info.get("review") or ""
            logger.info("改动比对完成, cost=%.3fs, brief=%s", time.time() - t1, notice_brief)
        except Exception:
            logger.exception("改动说明生成失败，不影响回传: %s", filename)

        # CHEAPMSG_V1：群里只要 brief 两行，长说明不进 warnings
        # （以前把整段 notice 塞进 warnings，群里一排就是十几行）
        warnings = [w for w in stats.get("warnings", []) if not _WORKPOINT_WARNING.match(w)]
        if not notice_brief:
            warnings.append("⚠️ 没生成改动说明，请人工核对与原表的差异")

        try:
            archive_workshop_pair(
                original_path=input_path,
                original_name=filename,
                processed_path=output_path,
                processed_name=output_name,
                uploader_name=req.get("uploader_name") or "",
            )
        except Exception:
            logger.exception("归档到共享盘失败，不影响回传: %s", filename)

        pair_key = str(Path(input_path).name)
        return {
            "success": True,
            "order_type": order_type,
            "output_dir": str(output_dir),
            "output_files": {
                pair_key: [_encode_output_file(output_path, logger, output_name)],
            },
            "output_filename": output_name,
            "notice": notice,
            "notice_brief": notice_brief,
            "notice_review": notice_review,
            "stats": stats,
            "warnings": warnings,
        }
    except Exception as e:
        import traceback
        logger.exception("处理文件失败: %s", filename)
        return {"success": False, "error": f"服务异常: {str(e)}", "trace": traceback.format_exc()}


def _flatten_output_items(output_files) -> list[dict]:
    """把骨架批量 output_files（dict / list）摊平成 [{filename, file_content, path}, ...]。"""
    items: list[dict] = []
    if isinstance(output_files, dict):
        for pair in output_files.values():
            if isinstance(pair, list):
                items.extend(x for x in pair if isinstance(x, dict))
            elif isinstance(pair, dict):
                items.append(pair)
    elif isinstance(output_files, list):
        items.extend(x for x in output_files if isinstance(x, dict))
    return items


def _unique_arcname(name: str, used: set[str]) -> str:
    """zip 内同名时追加 (2)/(3)...，避免互相覆盖。"""
    if name not in used:
        used.add(name)
        return name
    stem = Path(name).stem
    suffix = Path(name).suffix
    n = 2
    while True:
        candidate = f"{stem}({n}){suffix}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        n += 1


def _prefix_batch_warnings(results: list, files: list) -> None:
    """批量时给每条 warning 带上来源文件名。

    网关只会把这些 warning 原样贴到群里那条「处理完成」后面；一批十几单，
    不带文件名就不知道是哪一单算不出来。
    """
    for idx, res in enumerate(results):
        if not isinstance(res, dict):
            continue
        items = res.get("warnings")
        if not items:
            continue
        name = ""
        if idx < len(files) and isinstance(files[idx], dict):
            name = Path(str(files[idx].get("filename") or "")).stem
        if name:
            res["warnings"] = [f"[{name}] {w}" for w in items]


def _batch_zip_outputs(response: dict, results: list, files: list) -> None:
    """批量成功 ≥2 个输出时，打成一个 zip，替换 response['output_files']。"""
    _prefix_batch_warnings(results, files)
    items = _flatten_output_items(response.get("output_files"))
    if len(items) < 2:
        return

    tmpdir = tempfile.mkdtemp(prefix="workshop_batch_zip_")
    try:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f"下车间批量_{stamp}_{len(items)}个.zip"
        zip_path = Path(tmpdir) / zip_name
        used_names: set[str] = set()
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in items:
                fname = item.get("filename") or "output.xlsx"
                arcname = _unique_arcname(fname, used_names)
                src = item.get("path")
                if src and Path(src).is_file():
                    zf.write(src, arcname)
                elif item.get("file_content"):
                    zf.writestr(arcname, base64.b64decode(item["file_content"]))
                else:
                    logger.warning("批量打包跳过无内容项: %s", fname)
        encoded = _encode_output_file(zip_path, logger, zip_name)
        if encoded.get("error"):
            logger.error("批量 zip 编码失败，保留原多文件回传: %s", encoded.get("error"))
            return
        response["output_files"] = {"batch": [encoded]}
        response["zip_original_count"] = len(items)
        logger.info("批量打包完成: %s, 内含 %d 个下车间文件", zip_name, len(items))
    except Exception:
        logger.exception("批量打包异常，保留原多文件回传")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _warmup_xls_converter():
    """startup_hook：LibreOffice xls 转换器预热（异常由骨架记录，不阻断启动）。"""
    cost = warm_xls_converter()
    logger.info("LibreOffice xls 转换器预热完成, cost=%.3fs", cost)


if __name__ == "__main__":
    run_process_server(
        service_name="workshop-order",
        default_port=8006,
        process_fn=process,
        logger=logger,
        max_workers_env="WORKSHOP_MAX_WORKERS",
        default_max_workers=1,
        default_output_base=WORK_DIR.parent / "data" / "output",
        path_aliases=("/convert",),
        missing_file_error="缺少 file_content",
        batch_error_aggregate=True,
        batch_aggregate_hook=_batch_zip_outputs,
        startup_hook=_warmup_xls_converter,
    )
