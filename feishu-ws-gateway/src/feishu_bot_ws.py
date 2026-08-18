#!/usr/bin/env python3
"""
独立飞书 Bot 网关 — WebSocket 模式（基于 lark_oapi.ws.Client）

绕过 Gateway LLM，直接通过 lark_oapi 的 websocket 客户端接收飞书消息，
文件消息路由到本地 HTTP 服务（8001-8010）处理。

群路由配置（chat_id → 服务端口）：
  oc_f74b3f332d275f70ba22b4332b5b442d → 8002 (order-split)
  oc_52ccbd9aa43c7abcfe9a8039c638e934 → 8001 (hardware-summary)
  oc_09e8345ee873ce43f52ca182770b56a5 → 测试群（同时支持两种，通过文件名判断）
  oc_8b2a06d65c0b22fcdb24965898d86290 → 8009 (员工月度考勤)
  FEISHU_QUOTE_CHAT_ID                  → 8007 (quote-maker，可选环境变量)
  FEISHU_DEALER_REPORT_CHAT_ID          → 8008 (dealer-report，可选环境变量)

运行：
  python feishu_bot_ws.py
"""
import os
import sys
import re
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
    "oc_29ac7f425833255ff93fcf53f4575a70": {"port": 8003, "name": "销售部业绩核对"},
    "oc_43068f21ebba49ac209fbf78e9f86217": {"port": 8004, "name": "CSV板件转换"},
    "oc_51479339eef6b26fe9dcdcb8a5fb0c50": {"port": 8005, "name": "PVC分类"},
    "oc_c0986e7cea619374cfce226cbb199cc4": {"port": 8006, "name": "下车间单转换"},
    # 酷家乐月度经销商数据群：固定绑定 8008，不需要根据文件名路由
    "oc_ccb759f87c198521c575984b3f316cb8": {"port": 8008, "name": "酷家乐月度经销商数据"},
    # 员工月度考勤群
    "oc_8b2a06d65c0b22fcdb24965898d86290": {"port": 8009, "name": "员工月度考勤"},
    # 木皮优化群（门扇转换）
    "oc_9067ef06c46495ac35328d89bc16017d": {"port": 8010, "name": "木皮优化"},
    # 联思木皮优化群（PMS 优化门扇清单拆分）
    "oc_a4c44ef470d8a31456667367afff51b8": {"port": 8013, "name": "联思木皮优化"},
    # PVC订单汇总群
    "oc_b34c4d5f830f59b22d5e1a49bfbb630a": {"port": 8011, "name": "PVC订单汇总"},
}

# 木皮优化群 chat_id（该群支持文本命令管理剔除清单）
DOOR_SKIN_CHAT_ID = "oc_9067ef06c46495ac35328d89bc16017d"
DOOR_SKIN_PORT = 8010

# PVC优化群 chat_id（该群支持文本命令触发完整流程）
PVC_OPTIMIZE_CHAT_ID = "oc_51479339eef6b26fe9dcdcb8a5fb0c50"
PVC_TRIGGER_URL = "http://host.docker.internal:8012/process"

# @ 提醒用户映射（中文名 -> open_id）
AT_USER_MAP = {
    "胡娅": "ou_514e834114d0d71a519317197db98308",
    "刘佳": "ou_665c8b46b87c1bb2d213c7644eb3e68b",
}


def _replace_at_mentions(text: str) -> str:
    """把 @中文名 替换为飞书文本消息的 at 标签。"""
    for name, uid in AT_USER_MAP.items():
        text = text.replace(f'@{name}', f'<at user_id="{uid}">{name}</at>')
    return text


def _replace_at_mentions_in_card(obj):
    """递归替换卡片中的 @中文名 为卡片 at 标签。"""
    if isinstance(obj, dict):
        return {k: _replace_at_mentions_in_card(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_replace_at_mentions_in_card(v) for v in obj]
    if isinstance(obj, str):
        for name, uid in AT_USER_MAP.items():
            obj = obj.replace(f'@{name}', f'<at id="{uid}">{name}</at>')
        return obj
    return obj


for _chat_id in [x.strip() for x in os.getenv("FEISHU_QUOTE_CHAT_ID", "").split(",") if x.strip()]:
    CHAT_ROUTES[_chat_id] = {"port": 8007, "name": "报价单生成"}

for _chat_id in [x.strip() for x in os.getenv("FEISHU_DEALER_REPORT_CHAT_ID", "").split(",") if x.strip()]:
    CHAT_ROUTES[_chat_id] = {"port": 8008, "name": "酷家乐月度经销商数据"}

# 经销商群文件配对队列：chat_id -> {"file_path": ..., "file_name": ..., "message_id": ..., "file_key": ..., "time": ...}
_pending_files = {}
_pending_files_locks = {}  # chat_id -> threading.Lock

# 销售部业绩核对（8003）批量收集窗口
_DEALER_SALES_WINDOW = 20       # 每次上传后等待更多文件的窗口
_DEALER_SALES_FINAL_WINDOW = 60 # 最终等待窗口，支持分批一个个上传
_dealer_sales_queues = {}        # chat_id -> [file_info, ...]
_dealer_sales_timers = {}        # chat_id -> Timer (普通窗口)
_dealer_sales_final_timers = {}  # chat_id -> Timer (最终等待窗口)

# 经销商数据报表（8008）：账号指标作为主文件决定月份范围，设计师数据覆盖完整后自动处理
_dealer_report_queues = {}   # chat_id -> [file_info, ...]
_dealer_report_timers = {}   # chat_id -> Timer
_dealer_report_locks = {}    # chat_id -> threading.Lock

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


def _send_feishu_message(chat_id: str, msg_type: str, content: dict) -> str | None:
    """发送飞书消息，返回 message_id（若接口返回）。"""
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
        return None
    return (res.get("data") or {}).get("message_id")


# 缓存发送过的卡片，用于点击后更新
_card_message_cache: dict[str, dict] = {}


def _update_card_message(message_id: str, card: dict) -> None:
    """更新已发送的卡片消息。"""
    if not message_id:
        return
    token = _get_tenant_access_token()
    url = f"https://open.feishu.cn/open-apis/im/v1/messages/{message_id}"
    body = json.dumps({"content": json.dumps(card, ensure_ascii=False)}).encode("utf-8")
    req = urllib_request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }, method="PATCH")
    try:
        with urllib_request.urlopen(req, timeout=10) as resp:
            res = json.loads(resp.read().decode("utf-8"))
        if res.get("code") != 0:
            logger.warning("[%s] 更新卡片失败: %s", message_id, res)
    except Exception as e:
        logger.warning("[%s] 更新卡片异常: %s", message_id, e)


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
# 木皮优化群：文本命令管理剔除清单（8010）
# ---------------------------------------------------------------------------

_EXCLUSION_CMD_USAGE = (
    "支持的命令（@我或直接发均可）：\n"
    "• 加剔除 型号或关键词（可多个，空格分隔，自动识别型号/关键词）\n"
    "• 加型号 YM-999 / 加关键词 铝框门（指定类型）\n"
    "• 删剔除 型号或关键词（只能删群里加的，内置型号需改代码）\n"
    "• 剔除列表（查看当前完整清单）"
)


def _call_door_skin_exclusions(payload: dict = None) -> dict:
    """调用 8010 的剔除清单接口，payload 为 None 时是查询。"""
    url = f"http://{HTTP_SERVICE_HOST}:{DOOR_SKIN_PORT}/exclusions"
    data = None
    headers = {}
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib_request.Request(url, data=data, headers=headers)
    opener = urllib_request.build_opener(urllib_request.ProxyHandler({}))
    with opener.open(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _format_exclusions_reply(result: dict, action_desc: str = "") -> str:
    models = result.get("models", [])
    keywords = result.get("keywords", [])
    lines = []
    if action_desc:
        lines.append(action_desc)
    lines.append(f"📑 当前剔除清单（{len(models)} 个型号 + {len(keywords)} 个关键词，命中即不生成）")
    lines.append("型号：" + "、".join(models))
    lines.append("关键词：" + ("、".join(keywords) if keywords else "无"))
    return "\n".join(lines)


def _handle_pvc_process_command(chat_id: str, raw_text: str) -> None:
    """处理 PVC优化群文本命令：触发完整 PVC 流程"""
    logger.info("[%s] 收到 PVC 处理命令: %s", chat_id, raw_text)
    try:
        body = json.dumps({
            "chat_id": chat_id,
            "text": raw_text,
        }, ensure_ascii=False).encode("utf-8")
        req = urllib_request.Request(
            PVC_TRIGGER_URL,
            data=body,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )
        with urllib_request.urlopen(req, timeout=10) as resp:
            res = json.loads(resp.read().decode("utf-8"))
        if not res.get("success"):
            _send_text(chat_id, f"⚠️ 触发失败：{res.get('error', '未知错误')}")
    except Exception as e:
        logger.exception("[%s] 调用 PVC 触发服务失败", chat_id)
        _send_text(chat_id, f"❌ 触发 PVC 流程失败：{str(e)}")


def _handle_door_skin_text_command(chat_id: str, raw_text: str) -> None:
    """处理木皮优化群里的剔除清单管理命令。"""
    text = re.sub(r"<at[^>]*>.*?</at>", "", raw_text)
    text = re.sub(r"@_user_\d+", "", text).strip()
    if not text:
        return

    try:
        if text in ("剔除列表", "查看剔除", "剔除表", "剔除帮助"):
            if text == "剔除帮助":
                _send_text(chat_id, _EXCLUSION_CMD_USAGE)
                return
            result = _call_door_skin_exclusions()
            _send_text(chat_id, _format_exclusions_reply(result))
            return

        m = re.match(r"^(加剔除|添加剔除|加型号|加关键词)\s*[：: ]?\s*(.+)$", text)
        if m:
            cmd, rest = m.group(1), m.group(2)
            kind = {"加型号": "model", "加关键词": "keyword"}.get(cmd, "auto")
            items = [x for x in re.split(r"[\s,，、]+", rest) if x]
            result = _call_door_skin_exclusions({"action": "add", "items": items, "kind": kind})
            parts = []
            added = result.get("added", {})
            if added.get("models"):
                parts.append("已加型号：" + "、".join(added["models"]))
            if added.get("keywords"):
                parts.append("已加关键词：" + "、".join(added["keywords"]))
            if result.get("skipped"):
                parts.append("跳过：" + "、".join(result["skipped"]))
            desc = "✅ " + "；".join(parts) if parts else "⚠️ 没有新增项"
            _send_text(chat_id, _format_exclusions_reply(result, desc))
            return

        m = re.match(r"^(删剔除|删除剔除|取消剔除)\s*[：: ]?\s*(.+)$", text)
        if m:
            items = [x for x in re.split(r"[\s,，、]+", m.group(2)) if x]
            result = _call_door_skin_exclusions({"action": "remove", "items": items})
            parts = []
            if result.get("removed"):
                parts.append("已删除：" + "、".join(result["removed"]))
            if result.get("skipped"):
                parts.append("跳过：" + "、".join(result["skipped"]))
            desc = "✅ " + "；".join(parts) if parts else "⚠️ 没有删除项"
            _send_text(chat_id, _format_exclusions_reply(result, desc))
            return

        if text.startswith("剔除"):
            _send_text(chat_id, _EXCLUSION_CMD_USAGE)
    except Exception as e:
        logger.exception("[%s] 剔除命令处理失败", chat_id)
        _send_text(chat_id, f"❌ 剔除命令处理失败：{e}")


# ---------------------------------------------------------------------------
# 文件名判断文件类型
# ---------------------------------------------------------------------------

def _detect_type_by_filename(filename: str) -> int:
    """根据文件名判断测试群路由端口。"""
    name_lower = filename.lower()
    if any(k in name_lower for k in ["账号指标", "设计师数据统计"]):
        return 8008
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


def _service_name_for_port(port: int, fallback: str = "单文件处理") -> str:
    return {
        8001: "五金汇总",
        8002: "报价料单",
        8003: "经销商销售",
        8004: "CSV板件转换",
        8005: "PVC分类",
        8006: "下车间单转换",
        8007: "报价单生成",
        8008: "酷家乐月度经销商数据",
        8009: "员工月度考勤",
        8010: "木皮优化",
        8011: "PVC订单汇总",
        8013: "联思木皮优化",
    }.get(port, fallback)


def _detect_dealer_report_type(file_name: str) -> str:
    """识别 经销商数据报表文件类型"""
    name = file_name.lower()
    if "账号指标" in name or "账号信息" in name:
        return "account"
    if "设计师数据统计" in name or "设计师" in name:
        return "designer"
    return "unknown"


def _extract_dealer_report_months(file_name: str) -> set[int]:
    """从文件名中提取月份集合，例如 1-6月、1-3月、6月。"""
    name = file_name.lower()
    months = set()
    for start, end in re.findall(r"(?<!\d)(1[0-2]|0?[1-9])\s*[-~至到]\s*(1[0-2]|0?[1-9])\s*月", name):
        start_m = int(start)
        end_m = int(end)
        if start_m <= end_m:
            months.update(range(start_m, end_m + 1))
        else:
            months.update(range(start_m, 13))
            months.update(range(1, end_m + 1))

    range_spans = list(re.finditer(r"(?<!\d)(1[0-2]|0?[1-9])\s*[-~至到]\s*(1[0-2]|0?[1-9])\s*月", name))
    masked = name
    for span in reversed(range_spans):
        masked = masked[:span.start()] + " " * (span.end() - span.start()) + masked[span.end():]
    for month in re.findall(r"(?<!\d)(1[0-2]|0?[1-9])\s*月", masked):
        months.add(int(month))
    return months


def _format_months(months: set[int]) -> str:
    if not months:
        return "未知月份"
    ordered = sorted(months)
    if ordered == list(range(ordered[0], ordered[-1] + 1)):
        return f"{ordered[0]}月" if len(ordered) == 1 else f"{ordered[0]}-{ordered[-1]}月"
    return "、".join(f"{m}月" for m in ordered)


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


def _call_dealer_report_service(files: list, title: str = None, output_filename: str = None) -> dict:
    """调用经销商数据报表服务（端口 8008），上传账号指标和设计师数据统计文件。"""
    url = f"http://{HTTP_SERVICE_HOST}:8008/process"
    payload_files = []
    for f in files:
        with open(f["file_path"], "rb") as fh:
            b64 = base64.b64encode(fh.read()).decode("utf-8")
        payload_files.append({"file_content": b64, "filename": f["file_name"]})
    payload = {"files": payload_files}
    if title:
        payload["title"] = title
    if output_filename:
        payload["output_filename"] = output_filename
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


def _collect_result_warnings(result: dict) -> list[str]:
    warnings = []
    for key in ("warning", "warnings"):
        value = result.get(key)
        if isinstance(value, list):
            warnings.extend(str(item) for item in value if item)
        elif value:
            warnings.append(str(value))
    for res in result.get("results", []):
        warnings.extend(_collect_result_warnings(res))
    return warnings


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
        warnings = _collect_result_warnings(result)
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
    """识别 销售部业绩核对文件类型"""
    name = file_name.lower()
    if "综合查询" in name:
        return "zhcx"
    if "联思" in name:
        return "liansi"
    if "奢匠" in name or "下单统计" in name or "线下" in name:
        return "shejiang"
    return "unknown"


def _extract_month_from_dealer_sales_files(files):
    """从文件名中提取月份数字"""
    for f in files:
        fn = f.get("file_name", "")
        m = re.search(r"(\d+)\s*月", fn)
        if m:
            return str(int(m.group(1)))
    return None


def _process_dealer_sales_batch(chat_id: str, service_name: str, is_final: bool = False):
    """批量处理 销售部业绩核对队列中的三个文件
    is_final=False: 普通 20 秒窗口触发，如果没齐进入最终 60 秒等待
    is_final=True: 最终 60 秒窗口触发，如果没齐则清空并提示超时
    """
    global _dealer_sales_queues, _dealer_sales_timers, _dealer_sales_final_timers

    # 只有最终定时器触发时才清理 final timer 引用
    if is_final:
        _dealer_sales_final_timers.pop(chat_id, None)
    else:
        _dealer_sales_timers.pop(chat_id, None)

    queue = _dealer_sales_queues.get(chat_id, [])
    if not queue:
        return

    zhcx = [f for f in queue if _detect_may_sales_type(f["file_name"]) == "zhcx"]
    liansi = [f for f in queue if _detect_may_sales_type(f["file_name"]) == "liansi"]
    shejiang = [f for f in queue if _detect_may_sales_type(f["file_name"]) == "shejiang"]

    logger.info("[%s] 销售部业绩核对批量处理触发(is_final=%s)，队列 %d 个，综合查询 %d 个，联思 %d 个，奢匠 %d 个",
                chat_id, is_final, len(queue), len(zhcx), len(liansi), len(shejiang))

    if not zhcx or not liansi or not shejiang:
        if is_final:
            # 最终等待超时，清空队列
            _send_text(chat_id, f"\u274c {service_name}等待超时：未凑齐三类文件（综合查询 {len(zhcx)} 个，联思 {len(liansi)} 个，奢匠 {len(shejiang)} 个），请重新上传。")
            for f in queue:
                try:
                    Path(f["file_path"]).unlink(missing_ok=True)
                except Exception:
                    pass
            _dealer_sales_queues.pop(chat_id, None)
        else:
            # 普通窗口没齐，提示还缺什么，并启动最终等待窗口
            missing_names = []
            if not zhcx:
                missing_names.append("综合查询")
            if not liansi:
                missing_names.append("联思系统")
            if not shejiang:
                missing_names.append("奢匠/下单统计")
            _send_text(chat_id, f"\u26a0\ufe0f 已收到部分文件，还缺少：{'、'.join(missing_names)}。请在 {_DEALER_SALES_FINAL_WINDOW} 秒内继续上传，已收到的文件会保留。")

            # 取消已有的最终定时器
            if chat_id in _dealer_sales_final_timers and _dealer_sales_final_timers[chat_id] is not None:
                _dealer_sales_final_timers[chat_id].cancel()

            timer = threading.Timer(_DEALER_SALES_FINAL_WINDOW, _process_dealer_sales_batch, args=(chat_id, service_name, True))
            timer.daemon = True
            timer.start()
            _dealer_sales_final_timers[chat_id] = timer
        return

    # 凑齐了，处理
    files = [zhcx[0], liansi[0], shejiang[0]]
    for f in queue:
        if f not in files:
            try:
                Path(f["file_path"]).unlink(missing_ok=True)
            except Exception:
                pass
    _dealer_sales_queues.pop(chat_id, None)

    logger.info("[%s] 销售部业绩核对文件凑齐，开始处理: %s", chat_id, [f["file_name"] for f in files])
    try:
        result = _call_batch_service(8003, files)
        logger.info("[%s] 处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            _send_text(chat_id, f"\u274c 处理失败：{result.get('error', '未知错误')}")
            return

        output_files = result.get("output_files", [])
        if not output_files:
            _send_text(chat_id, "\u26a0\ufe0f 处理完成但未生成文件")
            return

        output_pairs = _normalize_output_pairs(output_files)
        sent_count = _send_output_pairs(chat_id, output_pairs, "从返回内容上传", "检查输出文件", False, False)

        month = _extract_month_from_dealer_sales_files(files)
        month_text = f"{month}月" if month else ""
        msg = f"\u2705 {service_name}{month_text}处理完成，共 {sent_count} 个文件，请检查。"
        warnings = _collect_result_warnings(result)
        if warnings:
            msg += "\n\n" + "\n".join(warnings)
        _send_text(chat_id, msg)
    except Exception as e:
        logger.exception("[%s] 处理异常", chat_id)
        _send_text(chat_id, f"\u274c 处理异常：{str(e)}")

def _handle_dealer_file(chat_id: str, message_id: str, file_key: str, file_name: str, local_path: str, service_name: str):
    """处理 销售部业绩核对群文件：支持批量上传或分批一个个上传"""
    global _dealer_sales_queues, _dealer_sales_timers, _dealer_sales_final_timers

    ftype = _detect_may_sales_type(file_name)
    if ftype == "unknown":
        _send_text(chat_id, f"\u26a0\ufe0f 无法识别文件「{file_name}」，文件名需包含：综合查询、联思、奢匠/下单统计")
        return

    save_dir = Path(tempfile.gettempdir()) / f"dealer_sales_pending_{chat_id}"
    save_dir.mkdir(exist_ok=True)
    save_path = save_dir / file_name
    shutil.copy(local_path, save_path)

    if chat_id not in _dealer_sales_queues:
        _dealer_sales_queues[chat_id] = []

    _dealer_sales_queues[chat_id].append({
        "file_path": str(save_path),
        "file_name": file_name,
        "message_id": message_id,
        "file_key": file_key,
        "time": time.time(),
    })

    # 取消已有的普通窗口和最终等待定时器
    if chat_id in _dealer_sales_timers and _dealer_sales_timers[chat_id] is not None:
        _dealer_sales_timers[chat_id].cancel()
    if chat_id in _dealer_sales_final_timers and _dealer_sales_final_timers[chat_id] is not None:
        _dealer_sales_final_timers[chat_id].cancel()

    # 启动新的 20 秒普通窗口
    # 检查是否已经凑齐三类文件，凑齐立即处理
    queue = _dealer_sales_queues[chat_id]
    has_zhcx = any(_detect_may_sales_type(f["file_name"]) == "zhcx" for f in queue)
    has_liansi = any(_detect_may_sales_type(f["file_name"]) == "liansi" for f in queue)
    has_shejiang = any(_detect_may_sales_type(f["file_name"]) == "shejiang" for f in queue)

    if has_zhcx and has_liansi and has_shejiang:
        logger.info("[%s] 销售部业绩核对三类文件已凑齐，立即处理", chat_id)
        _send_text(chat_id, f"✅ 已收到三类文件，立即处理...")
        timer = threading.Timer(0.5, _process_dealer_sales_batch, args=(chat_id, service_name, False))
        timer.daemon = True
        timer.start()
        _dealer_sales_timers[chat_id] = timer
        return

    # 没凑齐，启动新的 20 秒普通窗口
    timer = threading.Timer(_DEALER_SALES_WINDOW, _process_dealer_sales_batch, args=(chat_id, service_name, False))
    timer.daemon = True
    timer.start()
    _dealer_sales_timers[chat_id] = timer

    queue_len = len(_dealer_sales_queues[chat_id])
    logger.info("[%s] 销售部业绩核对文件加入队列: %s，当前 %d 个文件", chat_id, file_name, queue_len)
    _send_text(chat_id, f"\u2705 已收到第 {queue_len} 个文件「{file_name}」，{_DEALER_SALES_WINDOW}秒内上传更多文件会一起批量处理；也可以一个个慢慢传，已收到的文件会保留。")

def _process_dealer_report_queue(chat_id: str, service_name: str):
    """账号指标决定月份范围，设计师统计覆盖完整后生成一份经销商数据报表。"""
    global _dealer_report_queues, _dealer_report_timers
    _dealer_report_timers.pop(chat_id, None)
    queue = _dealer_report_queues.get(chat_id, [])

    accounts = [f for f in queue if _detect_dealer_report_type(f["file_name"]) == "account"]
    designers = [f for f in queue if _detect_dealer_report_type(f["file_name"]) == "designer"]

    logger.info("[%s] 经销商数据报表收集检查，队列 %d 个，账号指标 %d 个，设计师 %d 个",
                chat_id, len(queue), len(accounts), len(designers))

    if not accounts:
        _send_text(chat_id, "⚠️ 请先上传账号指标主文件，文件名需包含月份范围，例如「1-6月账号指标.xlsx」。")
        return

    account = accounts[-1]
    target_months = account.get("months") or _extract_dealer_report_months(account["file_name"])
    if not target_months:
        _send_text(chat_id, f"⚠️ 无法从账号指标「{account['file_name']}」识别月份范围，请把文件名改成类似「1-6月账号指标.xlsx」后重新上传。")
        return

    usable_designers = []
    invalid_designers = []
    designer_months = set()
    for designer in designers:
        months = designer.get("months") or _extract_dealer_report_months(designer["file_name"])
        if months - target_months:
            invalid_designers.append(designer)
            continue
        usable_designers.append(designer)
        designer_months.update(months)

    missing_months = target_months - designer_months
    if missing_months:
        invalid_text = ""
        if invalid_designers:
            invalid_text = "\n以下设计师文件月份超出账号指标范围，暂不参与汇总：" + "、".join(f["file_name"] for f in invalid_designers)
        _send_text(
            chat_id,
            f"✅ 已收到账号指标「{account['file_name']}」，目标范围：{_format_months(target_months)}。\n"
            f"当前设计师数据已覆盖：{_format_months(designer_months)}；还缺：{_format_months(missing_months)}。{invalid_text}"
        )
        return

    process_files = [account] + usable_designers
    sent_count = 0
    errors = []
    try:
        month_text = _format_months(target_months)
        result = _call_dealer_report_service(
            process_files,
            title=f"{month_text}经销商数据",
            output_filename=f"{month_text}经销商数据.xlsx",
        )
        logger.info("[%s] 处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            errors.append(result.get("error", "未知错误"))
        else:
            output_files = result.get("output_files", [])
            if not output_files:
                errors.append("未生成文件")

            for item in output_files:
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
                        logger.info("[%s] 从返回内容上传: %s (%d bytes)", chat_id, filename, len(file_data))
                        fk = _upload_feishu_file_content(chat_id, filename, file_data)
                    else:
                        win_path = _wsl_to_win_path(out_path)
                        if not os.path.exists(win_path):
                            continue
                        fk = _upload_feishu_file(chat_id, win_path)
                        filename = os.path.basename(win_path)
                    _send_file(chat_id, fk)
                    logger.info("[%s] 已发送文件: %s", chat_id, filename)
                    sent_count += 1
                except Exception as e:
                    logger.error("[%s] 发送文件失败: %s", chat_id, e)
                time.sleep(0.5)
    except Exception as e:
        logger.exception("[%s] 处理经销商数据报表异常", chat_id)
        errors.append(str(e))
    finally:
        _dealer_report_queues.pop(chat_id, None)
        for f in queue:
            try:
                Path(f["file_path"]).unlink(missing_ok=True)
            except Exception:
                pass

    if errors:
        _send_text(chat_id, "❌ 酷家乐月度经销商数据处理失败：\n" + "\n".join(errors))
    else:
        _send_text(chat_id, f"✅ {service_name}{_format_months(target_months)}处理完成，账号指标 1 个，设计师数据 {len(usable_designers)} 个，生成 {sent_count} 个结果文件。")


def _handle_dealer_report_file(chat_id: str, message_id: str, file_key: str, file_name: str, local_path: str, service_name: str):
    """处理经销商数据报表群文件：账号指标决定月份范围，设计师统计覆盖完整后调用 8008。"""
    ftype = _detect_dealer_report_type(file_name)
    if ftype == "unknown":
        _send_text(chat_id, f"\u26a0\ufe0f 无法识别文件「{file_name}」，文件名需包含：账号指标、账号信息、设计师数据统计、设计师")
        return
    months = _extract_dealer_report_months(file_name)
    if not months:
        _send_text(chat_id, f"⚠️ 无法从文件名「{file_name}」识别月份，请把文件名改成类似「1-6月账号指标.xlsx」或「1-3月设计师数据统计.xlsx」。")
        return

    save_dir = Path(tempfile.gettempdir()) / f"dealer_report_pending_{chat_id}"
    save_dir.mkdir(exist_ok=True)
    save_path = save_dir / file_name
    shutil.copy(local_path, save_path)

    lock = _dealer_report_locks.setdefault(chat_id, threading.Lock())
    with lock:
        if chat_id not in _dealer_report_queues:
            _dealer_report_queues[chat_id] = []
        if ftype == "account":
            old_accounts = [f for f in _dealer_report_queues[chat_id] if _detect_dealer_report_type(f["file_name"]) == "account"]
            for f in old_accounts:
                try:
                    Path(f["file_path"]).unlink(missing_ok=True)
                except Exception:
                    pass
            _dealer_report_queues[chat_id] = [f for f in _dealer_report_queues[chat_id] if _detect_dealer_report_type(f["file_name"]) != "account"]

        _dealer_report_queues[chat_id].append({
            "file_path": str(save_path),
            "file_name": file_name,
            "message_id": message_id,
            "file_key": file_key,
            "file_type": ftype,
            "months": months,
            "time": time.time(),
        })

        if chat_id in _dealer_report_timers and _dealer_report_timers[chat_id] is not None:
            _dealer_report_timers[chat_id].cancel()

        queue_len = len(_dealer_report_queues[chat_id])
        logger.info("[%s] 经销商数据报表文件加入队列: %s，当前 %d 个文件", chat_id, file_name, queue_len)
        _process_dealer_report_queue(chat_id, service_name)

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

        # 经销商数据报表群：双文件配对逻辑
        if port == 8008:
            _handle_dealer_report_file(chat_id, message_id, file_key, file_name, local_path, service_name)
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
        if result.get("card"):
            card = _replace_at_mentions_in_card(result["card"])
            msg_id = _send_feishu_message(chat_id, "interactive", card)
            if msg_id:
                _card_message_cache[msg_id] = card
        elif result.get("message"):
            _send_text(chat_id, _replace_at_mentions(result["message"]))
        output_files = _flatten_output_files(raw_output_files)
        for item in output_files:
            sent_count += _send_output_item(chat_id, item, "从 output_files 内容上传", "检查 output_files", False, True)

        if sent_count == 0 and not output_content and not output_files:
            _send_text(chat_id, "⚠️ 处理完成但未生成文件")
            return

        count = sent_count
        msg = f"✅ {service_name}处理完成，共 {count} 个文件，请检查。"
        warnings = _collect_result_warnings(result)
        if warnings:
            msg += "\n\n" + "\n".join(warnings)
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

        # 木皮优化群：文本命令管理剔除清单（加剔除/删剔除/剔除列表）
        if msg_type == "text" and chat_id == DOOR_SKIN_CHAT_ID:
            try:
                content_json = json.loads(content) if isinstance(content, str) else content
                raw_text = content_json.get("text", "")
            except Exception:
                raw_text = ""
            if raw_text:
                threading.Thread(
                    target=_handle_door_skin_text_command,
                    args=(chat_id, raw_text),
                    daemon=True,
                ).start()
            return

        # PVC优化群：文本命令触发完整流程（@机器人 + "处理"）
        if msg_type == "text" and chat_id == PVC_OPTIMIZE_CHAT_ID:
            try:
                content_json = json.loads(content) if isinstance(content, str) else content
                raw_text = content_json.get("text", "")
            except Exception:
                raw_text = ""
            if raw_text and "处理" in raw_text:
                threading.Thread(
                    target=_handle_pvc_process_command,
                    args=(chat_id, raw_text),
                    daemon=True,
                ).start()
            return

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
            # 未配置群 fallback：根据文件名识别经销商数据报表
            if _detect_dealer_report_type(file_name) != "unknown":
                port = 8008
                service_name = "酷家乐月度经销商数据"
                logger.info("[%s] 未配置群但文件名匹配经销商数据报表，收到文件: %s → 路由到 %s (%d)", chat_id, file_name, service_name, port)
                threading.Thread(
                    target=_process_file,
                    args=(chat_id, message_id, file_key, file_name, port, service_name),
                    daemon=True,
                ).start()
                return
            logger.info("[%s] 未配置的群，忽略", chat_id)
            return

        port = route["port"]
        if port == "auto":
            port = _detect_type_by_filename(file_name)
            service_name = _service_name_for_port(port, route.get("name", "单文件处理"))
        else:
            service_name = route.get("name", _service_name_for_port(port))

        if port == 8003:
            # 经销商群：走双文件配对逻辑
            logger.info("[%s] 收到文件: %s → 路由到 %s (%d)", chat_id, file_name, service_name, port)
            threading.Thread(
                target=_process_file,
                args=(chat_id, message_id, file_key, file_name, port, service_name),
                daemon=True,
            ).start()
            return

        if port in (8001, 8002):
            # 五金汇总 / 报价料单：批量收集逻辑
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


def _build_confirmed_card(card: dict, confirmed_text: str) -> dict:
    """把原卡片的按钮替换为已核对状态。"""
    new_card = {
        "config": card.get("config", {}),
        "header": card.get("header", {}),
        "elements": [],
    }
    for elem in card.get("elements", []):
        if elem.get("tag") == "action":
            new_card["elements"].append({
                "tag": "div",
                "text": {
                    "tag": "lark_md",
                    "content": confirmed_text,
                },
            })
        else:
            new_card["elements"].append(elem)
    return new_card


def _on_card_action(data) -> dict:
    """处理消息卡片按钮点击事件。"""
    try:
        event = data.event
        action_value = (event.action.value or {}) if event.action else {}
        # 服务 → (绑定群, 核对人, 核对后文案)
        confirm_config = {
            "door-skin-converter": (
                "oc_9067ef06c46495ac35328d89bc16017d",
                "刘佳",
                "✅ **刘佳已核对无误**\n<at id=\"ou_514e834114d0d71a519317197db98308\">胡娅</at> 请接收处理。",
                "@胡娅 刘佳已核对无误，请接收处理。",
            ),
            "pms-door-split": (
                "oc_a4c44ef470d8a31456667367afff51b8",
                "胡娅",
                "✅ **胡娅已核对无误**",
                "胡娅已核对无误。",
            ),
        }
        service = action_value.get("service")
        if action_value.get("action") != "confirm" or service not in confirm_config:
            return {"msg": "success"}
        expected_chat, confirmer_name, confirmed_text, fallback_text = confirm_config[service]
        chat_id = event.context.open_chat_id if event.context else None
        if chat_id != expected_chat:
            return {"msg": "success"}
        operator_id = event.operator.open_id if event.operator else None
        confirmer_id = AT_USER_MAP.get(confirmer_name)
        if operator_id != confirmer_id:
            logger.info("[%s] 非%s点击核对按钮: %s", chat_id, confirmer_name, operator_id)
            return {"msg": "success"}
        # 更新原卡片为已核对状态
        msg_id = event.context.open_message_id if event.context else None
        original_card = _card_message_cache.pop(msg_id, None) if msg_id else None
        if original_card:
            confirmed_card = _build_confirmed_card(original_card, confirmed_text)
            _update_card_message(msg_id, confirmed_card)
            logger.info("[%s] %s点击核对按钮，已更新卡片", chat_id, confirmer_name)
        else:
            # 如果找不到原卡片缓存，回退到发文本
            _send_text(chat_id, _replace_at_mentions(fallback_text))
            logger.info("[%s] %s点击核对按钮（无缓存回退）", chat_id, confirmer_name)
    except Exception as e:
        logger.error("处理卡片动作失败: %s", e, exc_info=True)
    return {"msg": "success"}


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
        .register_p2_card_action_trigger(_on_card_action)
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
