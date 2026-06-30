#!/usr/bin/env python3
"""
独立飞书 Bot 网关 — WebSocket 模式（基于 lark_oapi.ws.Client）

绕过 Gateway LLM，直接通过 lark_oapi 的 websocket 客户端接收飞书消息，
文件消息路由到本地 HTTP 服务（8001-8007）处理。

群路由配置（chat_id → 服务端口）：
  oc_f74b3f332d275f70ba22b4332b5b442d → 8002 (order-split)
  oc_52ccbd9aa43c7abcfe9a8039c638e934 → 8001 (hardware-summary)
  oc_09e8345ee873ce43f52ca182770b56a5 → 测试群（同时支持两种，通过文件名判断）
  FEISHU_QUOTE_CHAT_ID                  → 8007 (quote-maker，可选环境变量)

运行：
  python feishu_bot_ws.py
"""
import os
import sys
import json
import time
import base64
import logging
import socket
import tempfile
import threading
import shutil
from datetime import datetime
from pathlib import Path
from urllib import request as urllib_request

import lark_oapi as lark
from lark_oapi.ws import Client as WSClient
from lark_oapi import EventDispatcherHandler


# ---------------------------------------------------------------------------
# 单实例锁：通过绑定固定本地端口，防止同时启动多个网关进程
# ---------------------------------------------------------------------------
_SINGLE_INSTANCE_PORT = 61234
_single_instance_socket = None


def _ensure_single_instance():
    """确保本程序只有一个实例在运行"""
    global _single_instance_socket
    try:
        _single_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        _single_instance_socket.bind(("127.0.0.1", _SINGLE_INSTANCE_PORT))
        # 端口绑定成功，说明没有其它实例在运行
    except socket.error as e:
        logger.error("另一个 feishu_bot_ws.py 实例已在运行（端口 %d 被占用）：%s", _SINGLE_INSTANCE_PORT, e)
        sys.exit(1)


# ---------------------------------------------------------------------------
# 消息去重：防止 WebSocket 重连/重放导致同一消息被处理多次
# ---------------------------------------------------------------------------
_seen_message_ids = {}
DEDUP_WINDOW = 300  # 5 分钟


def _is_duplicate_message(message_id: str) -> bool:
    """判断消息是否重复。file_key 可能相同，message_id 是消息唯一标识。"""
    now = time.time()
    # 清理过期记录
    expired = [mid for mid, ts in _seen_message_ids.items() if now - ts > DEDUP_WINDOW]
    for mid in expired:
        del _seen_message_ids[mid]

    if message_id in _seen_message_ids:
        logger.warning("检测到重复消息 %s，忽略", message_id)
        return True
    _seen_message_ids[message_id] = now
    return False

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
APP_ID = os.getenv("FEISHU_APP_ID", "cli_a96f57b08d3bdbd8")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
if not APP_SECRET:
    raise RuntimeError("FEISHU_APP_SECRET is required")

CHAT_ROUTES = {
    "oc_f74b3f332d275f70ba22b4332b5b442d": {"port": 8002, "name": "报价料单"},
    "oc_52ccbd9aa43c7abcfe9a8039c638e934": {"port": 8001, "name": "五金汇总"},
    "oc_09e8345ee873ce43f52ca182770b56a5": {"port": "auto", "name": "测试群"},
    "oc_29ac7f425833255ff93fcf53f4575a70": {"port": 8003, "name": "5月业绩核对"},
    "oc_43068f21ebba49ac209fbf78e9f86217": {"port": 8004, "name": "CSV板件转换"},
    "oc_51479339eef6b26fe9dcdcb8a5fb0c50": {"port": 8005, "name": "PVC分类"},
    "oc_c0986e7cea619374cfce226cbb199cc4": {"port": 8006, "name": "下车间单转换"},
}

for _chat_id in [x.strip() for x in os.getenv("FEISHU_QUOTE_CHAT_ID", "").split(",") if x.strip()]:
    CHAT_ROUTES[_chat_id] = {"port": 8007, "name": "报价单生成"}

# 经销商群文件配对队列：chat_id -> {"file_path": ..., "file_name": ..., "message_id": ..., "file_key": ..., "time": ...}
_pending_files = {}
_pending_files_locks = {}  # chat_id -> threading.Lock

# 批量收集配置
BATCH_COLLECTION_WINDOW = 10  # 秒
BATCH_MAX_FILES = 20
_batch_queues = {}      # chat_id -> [{file_path, file_name, message_id, file_key, time}]
_batch_timers = {}      # chat_id -> Timer

# WebSocket 假死检测
_last_activity_time = None  # None = 尚未收到过任何消息

HTTP_SERVICE_HOST = os.getenv("HTTP_SERVICE_HOST", "127.0.0.1")

LOG_FILE = Path(os.getenv("FEISHU_LOG_FILE", "/app/logs/feishu_bot_ws.log"))
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
logger = logging.getLogger("feishu-bot-ws")

# ---------------------------------------------------------------------------
# 飞书 API 工具
# ---------------------------------------------------------------------------
_token_cache = {"token": "", "expire": 0}


def _get_tenant_access_token() -> str:
    """获取 tenant_access_token（带缓存）"""
    global _token_cache
    now = time.time()
    if _token_cache["token"] and _token_cache["expire"] > now + 60:
        return _token_cache["token"]
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    data = json.dumps({"app_id": APP_ID, "app_secret": APP_SECRET}).encode("utf-8")
    req = urllib_request.Request(url, data=data, headers={"Content-Type": "application/json"})
    with urllib_request.urlopen(req, timeout=10) as resp:
        res = json.loads(resp.read().decode("utf-8"))
    if res.get("code") != 0:
        raise RuntimeError(f"获取 token 失败: {res}")
    _token_cache["token"] = res["tenant_access_token"]
    _token_cache["expire"] = now + res.get("expire", 7200)
    return _token_cache["token"]


def _download_feishu_file(message_id: str, file_key: str, save_dir: str) -> str:
    """下载飞书消息中的文件"""
    token = _get_tenant_access_token()
    url = f"https://open.feishu.cn/open-apis/im/v1/messages/{message_id}/resources/{file_key}?type=file"
    req = urllib_request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib_request.urlopen(req, timeout=30) as resp:
        body = resp.read()
        cd = resp.headers.get("Content-Disposition", "")
        fname = "download.bin"
        if "filename=" in cd:
            fname = cd.split("filename=")[-1].strip('"').split(";")[0]
        save_path = os.path.join(save_dir, fname)
        with open(save_path, "wb") as f:
            f.write(body)
    return save_path


def _wsl_to_win_path(path: str) -> str:
    r"""把 WSL 路径 /mnt/x/... 转成 Windows 路径 X:\..."""
    p = Path(path)
    parts = p.parts
    # Windows 下 Path('/mnt/d/...').parts[0] 是 '\\' 而不是 '/'，所以只判断 parts[1]=='mnt'
    if len(parts) >= 3 and parts[1] == "mnt" and len(parts[2]) == 1:
        drive = parts[2].upper()
        rest = parts[3:]
        return str(Path(f"{drive}:/").joinpath(*rest))
    return path


def _upload_feishu_file_content(chat_id: str, filename: str, file_data: bytes) -> str:
    """直接上传文件内容到飞书，返回 file_key"""
    token = _get_tenant_access_token()
    boundary = "----FormBoundary" + os.urandom(8).hex()
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file_type"\r\n\r\n'
        f"stream\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file_name"\r\n\r\n'
        f"{filename}\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
        f"Content-Type: application/octet-stream\r\n\r\n"
    ).encode("utf-8") + file_data + f"\r\n--{boundary}--\r\n".encode("utf-8")
    url = "https://open.feishu.cn/open-apis/im/v1/files"
    req = urllib_request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    })
    with urllib_request.urlopen(req, timeout=30) as resp:
        res = json.loads(resp.read().decode("utf-8"))
    if res.get("code") != 0:
        raise RuntimeError(f"上传文件失败: {res}")
    return res["data"]["file_key"]


def _upload_feishu_file(chat_id: str, file_path: str) -> str:
    """上传本地文件到飞书，返回 file_key（兼容旧路径模式）"""
    file_path = _wsl_to_win_path(file_path)
    filename = os.path.basename(file_path)
    with open(file_path, "rb") as f:
        file_data = f.read()
    return _upload_feishu_file_content(chat_id, filename, file_data)


def _send_feishu_message(chat_id: str, msg_type: str, content: dict) -> None:
    """发送飞书消息"""
    token = _get_tenant_access_token()
    receive_id_type = "open_id" if chat_id.startswith("ou_") else "chat_id"
    url = f"https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type={receive_id_type}"
    body = json.dumps({
        "receive_id": chat_id,
        "msg_type": msg_type,
        "content": json.dumps(content, ensure_ascii=False),
    }).encode("utf-8")
    req = urllib_request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    })
    with urllib_request.urlopen(req, timeout=10) as resp:
        res = json.loads(resp.read().decode("utf-8"))
    if res.get("code") != 0:
        logger.warning("发送消息失败: %s", res)


def _send_text(chat_id: str, text: str) -> None:
    _send_feishu_message(chat_id, "text", {"text": text})


def _send_file(chat_id: str, file_key: str) -> None:
    _send_feishu_message(chat_id, "file", {"file_key": file_key})


# ---------------------------------------------------------------------------
# 本地 HTTP 服务调用
# ---------------------------------------------------------------------------

def _call_local_service(port: int, input_path: str, filename: str, order_date: str = None) -> dict:
    """调用本地处理服务"""
    url = f"http://{HTTP_SERVICE_HOST}:{port}/process"
    with open(input_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    payload = {"file_content": b64, "filename": filename}
    if order_date:
        payload["order_date"] = order_date
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib_request.Request(url, data=body, headers={
        "Content-Type": "application/json; charset=utf-8",
    })
    # 本地服务不走代理
    opener = urllib_request.build_opener(urllib_request.ProxyHandler({}))
    with opener.open(req, timeout=120) as resp:
        return json.loads(resp.read().decode("utf-8"))


# ---------------------------------------------------------------------------
# 文件名判断文件类型
# ---------------------------------------------------------------------------

def _detect_type_by_filename(filename: str) -> int:
    """根据文件名判断测试群路由端口。"""
    name_lower = filename.lower()
    if any(k in name_lower for k in ["拆单报价", "quote-maker", "make_quote", "quote"]):
        return 8007
    if any(k in name_lower for k in ["五金", "hardware", "汇总", "马斌星"]):
        return 8001
    if any(k in name_lower for k in ["料单", "order", "split", "马忠义"]):
        return 8002
    if name_lower.startswith("b"):
        return 8001
    if name_lower.startswith("s"):
        return 8002
    return 8002


# ---------------------------------------------------------------------------
# 文件处理
# ---------------------------------------------------------------------------

def _call_dealer_service(chat_id: str, sj_path: str, ls_path: str, sj_name: str, ls_name: str) -> dict:
    """调用经销商销售 HTTP 服务（端口 8003），上传两个文件"""
    url = f"http://{HTTP_SERVICE_HOST}:8003/process"
    with open(sj_path, "rb") as f:
        sj_b64 = base64.b64encode(f.read()).decode("utf-8")
    with open(ls_path, "rb") as f:
        ls_b64 = base64.b64encode(f.read()).decode("utf-8")
    payload = {
        "shejiang_content": sj_b64,
        "liansi_content": ls_b64,
        "shejiang_name": sj_name,
        "liansi_name": ls_name,
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib_request.Request(url, data=body, headers={
        "Content-Type": "application/json; charset=utf-8",
    })
    opener = urllib_request.build_opener(urllib_request.ProxyHandler({}))
    with opener.open(req, timeout=300) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _call_batch_service(port: int, files: list, order_date: str = None) -> dict:
    """批量调用本地处理服务"""
    url = f"http://{HTTP_SERVICE_HOST}:{port}/process"
    payload_files = []
    for f in files:
        with open(f["file_path"], "rb") as fh:
            b64 = base64.b64encode(fh.read()).decode("utf-8")
        payload_files.append({"file_content": b64, "filename": f["file_name"]})
    payload = {"files": payload_files}
    if order_date:
        payload["order_date"] = order_date
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib_request.Request(url, data=body, headers={
        "Content-Type": "application/json; charset=utf-8",
    })
    opener = urllib_request.build_opener(urllib_request.ProxyHandler({}))
    with opener.open(req, timeout=300) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _normalize_output_pairs(output_files):
    """Return output files grouped in the same order as the service response."""
    if isinstance(output_files, dict):
        return list(output_files.values())
    return [[item] for item in output_files]


def _flatten_output_files(output_files):
    if isinstance(output_files, dict):
        flattened = []
        for pair in output_files.values():
            if isinstance(pair, list):
                flattened.extend(pair)
            else:
                flattened.append(pair)
        return flattened
    return output_files


def _send_output_item(chat_id: str, item, content_log: str, path_log: str, warn_missing: bool, exc_info: bool) -> int:
    if isinstance(item, dict):
        out_path = item.get("path", "")
        filename = item.get("filename") or (os.path.basename(out_path) if out_path else "output.xlsx")
        file_content_b64 = item.get("file_content")
    else:
        out_path = str(item)
        filename = os.path.basename(out_path)
        file_content_b64 = None

    try:
        if file_content_b64:
            file_data = base64.b64decode(file_content_b64)
            logger.info("[%s] %s: %s (%d bytes)", chat_id, content_log, filename, len(file_data))
            fk = _upload_feishu_file_content(chat_id, filename, file_data)
        else:
            win_path = _wsl_to_win_path(out_path)
            logger.info("[%s] %s: %s exists=%s", chat_id, path_log, win_path, os.path.exists(win_path))
            if not os.path.exists(win_path):
                if warn_missing:
                    logger.warning("[%s] 输出文件不存在，跳过: %s", chat_id, win_path)
                return 0
            fk = _upload_feishu_file(chat_id, win_path)
            filename = os.path.basename(win_path)
        _send_file(chat_id, fk)
        logger.info("[%s] 已发送文件: %s", chat_id, filename)
        return 1
    except Exception as e:
        logger.error("[%s] 发送文件失败: %s", chat_id, e, exc_info=exc_info)
        return 0


def _send_output_pairs(chat_id: str, output_pairs, content_log: str, path_log: str, warn_missing: bool, exc_info: bool) -> int:
    sent_count = 0
    for pair in output_pairs:
        for item in pair:
            sent_count += _send_output_item(chat_id, item, content_log, path_log, warn_missing, exc_info)
        # Keep a short pause between pairs so Feishu preserves the visible order.
        time.sleep(0.5)
    return sent_count


def _format_quantity_value(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _build_quantity_card(result: dict, service_name: str) -> dict | None:
    if "quantity_total" not in result:
        return None
    value = result.get("quantity_total")
    if value in (None, ""):
        return None
    value = _format_quantity_value(value)
    detail_rows = result.get("quantity_files") or []
    if not detail_rows:
        detail_rows = [{"filename": "合计", "quantity_total": value}]

    rows = [{"filename": "合计", "qty": str(value)}]
    for row in detail_rows:
        filename = row.get("filename", "")
        qty = _format_quantity_value(row.get("quantity_total", 0))
        rows.append({"filename": filename, "qty": str(qty)})

    return {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "green",
            "title": {"tag": "plain_text", "content": f"{service_name}数量明细"},
        },
        "elements": [
            {
                "tag": "table",
                "page_size": min(10, max(1, len(rows))),
                "row_height": "low",
                "freeze_first_column": False,
                "header_style": {
                    "text_align": "left",
                    "text_size": "normal",
                    "background_style": "grey",
                    "text_color": "default",
                    "bold": True,
                    "lines": 1,
                },
                "columns": [
                    {"name": "filename", "display_name": "文件", "data_type": "text", "width": "auto"},
                    {"name": "qty", "display_name": "数量", "data_type": "text", "width": "auto"},
                ],
                "rows": rows,
            }
        ],
    }


def _send_quantity_card(chat_id: str, result: dict, port: int, service_name: str) -> bool:
    if port != 8005:
        return False
    card = _build_quantity_card(result, service_name)
    if not card:
        return False
    try:
        _send_feishu_message(chat_id, "interactive", card)
        return True
    except Exception as e:
        logger.error("[%s] 发送数量表格卡片失败: %s", chat_id, e, exc_info=True)
        value = _format_quantity_value(result.get("quantity_total"))
        _send_text(chat_id, f"本次分类数量合计：{value}")
        return False


def _process_batch(chat_id: str, port: int, service_name: str):
    """处理批量队列"""
    global _batch_queues, _batch_timers
    queue = _batch_queues.pop(chat_id, [])
    _batch_timers.pop(chat_id, None)

    if not queue:
        return

    logger.info("[%s] 批量处理启动，共 %d 个文件", chat_id, len(queue))
    order_date = f"{datetime.now().year}.{datetime.now().month}.{datetime.now().day}"

    try:
        result = _call_batch_service(port, queue, order_date)
        logger.info("[%s] 批量处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            _send_text(chat_id, f"❌ {service_name}批量处理失败：{result.get('error', '未知错误')}")
            return

        output_files = result.get("output_files", [])
        if not output_files:
            _send_text(chat_id, "⚠️ 批量处理完成但未生成文件")
            return

        # 支持成对返回：dict {原始文件名: [file1, file2, ...]} 或旧格式列表
        output_pairs = _normalize_output_pairs(output_files)
        total_count = sum(len(pair) for pair in output_pairs)
        logger.info("[%s] 准备发送 %d 对输出文件，共 %d 个文件", chat_id, len(output_pairs), total_count)
        _send_output_pairs(chat_id, output_pairs, "从返回内容上传文件", "检查输出文件", True, True)

        count = total_count
        msg = f"✅ {service_name}批量处理完成，共处理 {len(queue)} 个文件，生成 {count} 个结果文件，请检查。"
        warnings = []
        for res in result.get("results", []):
            w = res.get("warning")
            if w:
                warnings.append(w)
        if warnings:
            msg += "\n\n" + "\n".join(warnings)
        _send_text(chat_id, msg)
        _send_quantity_card(chat_id, result, port, service_name)
    except Exception as e:
        logger.exception("[%s] 批量处理异常", chat_id)
        _send_text(chat_id, f"❌ {service_name}批量处理异常：{str(e)}")
    finally:
        # 清理临时文件
        for f in queue:
            try:
                if os.path.exists(f["file_path"]):
                    os.remove(f["file_path"])
            except Exception:
                pass


def _start_batch_timer(chat_id: str, port: int, service_name: str):
    """启动批量收集定时器"""
    global _batch_timers
    if chat_id in _batch_timers and _batch_timers[chat_id] is not None:
        return

    def timer_callback():
        _process_batch(chat_id, port, service_name)

    timer = threading.Timer(BATCH_COLLECTION_WINDOW, timer_callback)
    timer.daemon = True
    _batch_timers[chat_id] = timer
    timer.start()


def _detect_may_sales_type(file_name: str) -> str:
    """识别 5月业绩核对文件类型"""
    name = file_name.lower()
    if "综合查询" in name:
        return "zhcx"
    if "联思" in name:
        return "liansi"
    if "奢匠" in name or "下单统计" in name:
        return "shejiang"
    return "unknown"


def _handle_dealer_file(chat_id: str, message_id: str, file_key: str, file_name: str, local_path: str, service_name: str):
    """处理 5月业绩核对群文件：收集 3 个文件后调用 8003"""
    global _pending_files, _pending_files_locks

    ftype = _detect_may_sales_type(file_name)
    if ftype == "unknown":
        _send_text(chat_id, f"⚠️ 无法识别文件「{file_name}」，文件名需包含：综合查询、联思、奢匠/下单统计")
        return

    # 每个 chat_id 一个锁，避免三个文件同时上传时状态竞争
    if chat_id not in _pending_files_locks:
        _pending_files_locks[chat_id] = threading.Lock()
    lock = _pending_files_locks[chat_id]

    with lock:
        # 保存文件到持久化目录
        save_dir = Path(tempfile.gettempdir()) / f"dealer_pending_{chat_id}"
        save_dir.mkdir(exist_ok=True)
        save_path = save_dir / file_name
        shutil.copy(local_path, save_path)

        if chat_id not in _pending_files:
            _pending_files[chat_id] = {}

        _pending_files[chat_id][ftype] = {
            "file_path": str(save_path),
            "file_name": file_name,
            "message_id": message_id,
            "file_key": file_key,
            "time": time.time(),
        }

        # 检查是否凑齐 3 个文件
        required = {"zhcx", "liansi", "shejiang"}
        collected = set(_pending_files[chat_id].keys())
        missing = required - collected

        if missing:
            missing_names = []
            if "zhcx" in missing:
                missing_names.append("综合查询")
            if "liansi" in missing:
                missing_names.append("联思系统")
            if "shejiang" in missing:
                missing_names.append("奢匠下单统计")
            _send_text(chat_id, f"✅ 已收到【{file_name}】，还缺少：{'、'.join(missing_names)}，请继续上传。")
            return

        # 凑齐了，提取文件并删除待处理记录
        files = []
        for key in required:
            files.append(_pending_files[chat_id][key])
        del _pending_files[chat_id]

    # 释放锁后再调用 8003，避免长时间占用锁
    logger.info("[%s] 5月业绩核对文件凑齐，开始处理: %s", chat_id, [f["file_name"] for f in files])
    try:
        result = _call_batch_service(8003, files)
        logger.info("[%s] 处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            _send_text(chat_id, f"❌ 处理失败：{result.get('error', '未知错误')}")
            return

        output_files = result.get("output_files", [])
        if not output_files:
            _send_text(chat_id, "⚠️ 处理完成但未生成文件")
            return

        # 支持成对返回：dict {原始文件名: [file1, file2, ...]} 或旧格式列表
        output_pairs = _normalize_output_pairs(output_files)
        sent_count = _send_output_pairs(chat_id, output_pairs, "从返回内容上传", "检查输出文件", False, False)

        _send_text(chat_id, f"✅ {service_name}处理完成，共 {sent_count} 个文件，请检查。")
    except Exception as e:
        logger.exception("[%s] 处理异常", chat_id)
        _send_text(chat_id, f"❌ 处理异常：{str(e)}")
    finally:
        # 清理持久化文件
        shutil.rmtree(save_dir, ignore_errors=True)

def _process_file(chat_id: str, message_id: str, file_key: str, file_name: str, port: int, service_name: str):
    """后台处理文件"""
    tmpdir = tempfile.mkdtemp()
    try:
        logger.info("[%s] 下载文件: %s", chat_id, file_name)
        local_path = _download_feishu_file(message_id, file_key, tmpdir)
        logger.info("[%s] 下载完成: %s", chat_id, local_path)

        # 经销商群：双文件配对逻辑
        if port == 8003:
            _handle_dealer_file(chat_id, message_id, file_key, file_name, local_path, service_name)
            return

        order_date = f"{datetime.now().year}.{datetime.now().month}.{datetime.now().day}"
        result = _call_local_service(port, local_path, file_name, order_date)
        logger.info("[%s] 处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            error_msg = result.get("error", "未知错误")
            _send_text(chat_id, f"❌ 处理失败：{error_msg}")
            return

        sent_count = 0

        # 优先使用服务返回的 output_content（避免服务端临时目录被提前清理）
        output_content = result.get("output_content")
        output_filename = result.get("output_filename")
        if output_content:
            try:
                out_name = output_filename or f"{Path(file_name).stem}_result.xlsx"
                out_path = Path(tmpdir) / out_name
                out_path.write_bytes(base64.b64decode(output_content))
                logger.info("[%s] 从 output_content 写出文件: %s size=%d", chat_id, out_path, out_path.stat().st_size)
                fk = _upload_feishu_file(chat_id, str(out_path))
                _send_file(chat_id, fk)
                logger.info("[%s] 已发送文件: %s", chat_id, out_name)
                sent_count += 1
            except Exception as e:
                logger.error("[%s] 发送 output_content 文件失败: %s", chat_id, e, exc_info=True)

        # 兼容 output_files 模式（字符串路径、dict 列表，或 {原始文件名: [file, ...]} 字典）
        raw_output_files = result.get("output_files", [])
        output_files = _flatten_output_files(raw_output_files)
        for item in output_files:
            sent_count += _send_output_item(chat_id, item, "从 output_files 内容上传", "检查 output_files", False, True)

        if sent_count == 0 and not output_content and not output_files:
            _send_text(chat_id, "⚠️ 处理完成但未生成文件")
            return

        count = sent_count
        msg = f"✅ {service_name}处理完成，共 {count} 个文件，请检查。"
        warning = result.get("warning")
        if warning:
            msg += f"\n\n{warning}"
        _send_text(chat_id, msg)
        _send_quantity_card(chat_id, result, port, service_name)

    except Exception as e:
        logger.exception("[%s] 处理异常", chat_id)
        _send_text(chat_id, f"❌ 处理异常：{str(e)}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------------------------------------------------------------------
# 事件处理回调
# ---------------------------------------------------------------------------

def _on_message_receive(data) -> None:
    """处理 im.message.receive_v1 事件"""
    global _last_activity_time
    _last_activity_time = time.time()
    try:
        # data 是 P2ImMessageReceiveV1 对象
        msg = data.event.message
        chat_id = msg.chat_id
        message_id = msg.message_id

        # 消息去重
        if _is_duplicate_message(message_id):
            return

        msg_type = msg.message_type
        content = msg.content

        # 只处理文件
        if msg_type not in ("file", "audio", "media"):
            return

        try:
            content_json = json.loads(content) if isinstance(content, str) else content
            file_key = content_json.get("file_key", "")
            file_name = content_json.get("file_name", "unknown")
        except Exception:
            logger.warning("解析文件消息失败")
            return

        if not file_key:
            return

        # 路由判断
        route = CHAT_ROUTES.get(chat_id)
        if not route:
            logger.info("[%s] 未配置的群，忽略", chat_id)
            return

        port = route["port"]
        if port == "auto":
            port = _detect_type_by_filename(file_name)

        if port == 8003:
            # 经销商群：走双文件配对逻辑
            service_name = "经销商销售"
            logger.info("[%s] 收到文件: %s → 路由到 %s (%d)", chat_id, file_name, service_name, port)
            threading.Thread(
                target=_process_file,
                args=(chat_id, message_id, file_key, file_name, port, service_name),
                daemon=True,
            ).start()
            return

        if port in (8001, 8002):
            # 五金汇总 / 报价料单：批量收集逻辑
            service_name = route.get("name", "单文件处理")
            logger.info("[%s] 收到文件: %s → 路由到 %s (%d) 批量收集", chat_id, file_name, service_name, port)

            # 下载文件到持久化目录
            tmpdir = tempfile.mkdtemp()
            local_path = _download_feishu_file(message_id, file_key, tmpdir)
            save_dir = Path(tempfile.gettempdir()) / f"batch_pending_{chat_id}"
            save_dir.mkdir(exist_ok=True)
            save_path = save_dir / file_name
            shutil.copy(local_path, save_path)
            shutil.rmtree(tmpdir, ignore_errors=True)

            # 加入批量队列
            global _batch_queues
            if chat_id not in _batch_queues:
                _batch_queues[chat_id] = []
            _batch_queues[chat_id].append({
                "file_path": str(save_path),
                "file_name": file_name,
                "message_id": message_id,
                "file_key": file_key,
                "time": time.time(),
            })

            queue_len = len(_batch_queues[chat_id])
            if queue_len >= BATCH_MAX_FILES:
                # 达到最大文件数，立即处理
                logger.info("[%s] 批量队列已满 (%d)，立即处理", chat_id, queue_len)
                if chat_id in _batch_timers and _batch_timers[chat_id] is not None:
                    _batch_timers[chat_id].cancel()
                    _batch_timers[chat_id] = None
                threading.Thread(
                    target=_process_batch,
                    args=(chat_id, port, service_name),
                    daemon=True,
                ).start()
            else:
                # 启动或延续收集窗口
                _start_batch_timer(chat_id, port, service_name)
                _send_text(chat_id, f"✅ 已收到第 {queue_len} 个文件「{file_name}」，{BATCH_COLLECTION_WINDOW}秒内继续上传的文件将一起批量处理。")
            return

        # 单文件处理
        service_name = route.get("name", "单文件处理")
        logger.info("[%s] 收到文件: %s → 路由到 %s (%d)", chat_id, file_name, service_name, port)
        threading.Thread(
            target=_process_file,
            args=(chat_id, message_id, file_key, file_name, port, service_name),
            daemon=True,
        ).start()

    except Exception as e:
        logger.exception("处理消息事件异常: %s", e)


# ---------------------------------------------------------------------------
# 主程序
# ---------------------------------------------------------------------------

def _health_check():
    """守护线程：检测 WebSocket 是否假死。

    lark_oapi 的 WSClient 会在连接断开后自动重连，因此这里只负责：
    1. 发现长时间无消息时记录警告
    2. 更新基准时间，避免重复告警
    3. 真正的重启由外部健康检查（Windows 计划任务）完成
    """
    global _last_activity_time

    # 等待首次消息（最多等30分钟）
    wait_start = time.time()
    while _last_activity_time is None:
        if time.time() - wait_start > 1800:
            logger.warning("启动30分钟后仍未收到任何消息，设置基准时间")
            _last_activity_time = time.time()
            break
        time.sleep(60)

    # 开始监控循环
    while True:
        time.sleep(300)  # 每5分钟检查一次

        now = datetime.now()
        # 工作时间：8:00 - 20:00
        if now.hour < 8 or now.hour >= 20:
            continue

        idle = time.time() - _last_activity_time
        if idle > 900:  # 15分钟
            logger.warning(
                "工作时间内 %d 分钟未收到任何消息，WebSocket 可能假死。"
                "WSClient 会自动重连；如长期未恢复，健康检查任务会重启本服务。",
                int(idle / 60),
            )
            # 更新基准时间，避免同一事件反复告警
            _last_activity_time = time.time()


def main():
    # 先初始化日志再检查单实例
    logger.info("=" * 60)
    logger.info("飞书 Bot WebSocket 网关启动 (lark_oapi.ws.Client)")

    # 单实例锁：防止同时运行多个网关进程导致消息重复处理
    _ensure_single_instance()
    logger.info("单实例锁获取成功（端口 %d）", _SINGLE_INSTANCE_PORT)

    logger.info("路由配置:")
    for chat_id, cfg in CHAT_ROUTES.items():
        logger.info("  %s → %s (port=%s)", chat_id, cfg["name"], cfg["port"])
    logger.info("=" * 60)

    import lark_oapi
    lark_oapi.logger.setLevel(logging.INFO)
    lark_oapi.logger.addHandler(logging.StreamHandler(sys.stdout))

    # 启动假死检测守护线程
    threading.Thread(target=_health_check, daemon=True).start()
    logger.info("WebSocket 假死检测已启动（工作时间 8:00-20:00，15分钟无消息自动重启）")

    handler = (
        EventDispatcherHandler.builder("", "")
        .register_p2_im_message_receive_v1(_on_message_receive)
        .build()
    )

    client = WSClient(
        app_id=APP_ID,
        app_secret=APP_SECRET,
        log_level=lark.LogLevel.INFO,
        event_handler=handler,
    )
    logger.info("正在连接飞书 WebSocket...")
    client.start()


if __name__ == "__main__":
    main()
