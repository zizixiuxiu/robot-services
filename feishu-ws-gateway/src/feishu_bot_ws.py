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
import uuid
from datetime import datetime
from pathlib import Path
from urllib import request as urllib_request

import lark_oapi as lark
from lark_oapi.ws import Client as WSClient
from lark_oapi import EventDispatcherHandler

# 文件类型按内容识别（sheet 名/表头关键字），文件名只做兜底
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import file_probe
except Exception:                      # noqa: BLE001  缺模块时退回纯文件名识别
    file_probe = None


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
# 飞书 WS 重连后可能重放数小时前的未确认消息（实测 6 分钟 ~ 7 小时后仍有重放），
# 窗口必须覆盖重连重放的延迟；消息量小（几百条/天），内存可忽略。
DEDUP_WINDOW = 86400  # 24 小时


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
    # 门套窗套豪迈电子据优化
    # 2026-11-07 停用：该群及配套服务已不再使用（homag-frame-csv 自 2026-08-26 建库后
    # 零真实调用），容器 homag-frame-csv-8014 已 stop，路由一并注释掉，群内发文件不再响应。
    # 恢复：取消下一行注释 + docker start homag-frame-csv-8014（或指向平台适配器 18014）
    # "oc_98c025defa1eb03c0b7c1a4a805f6a7a": {"port": 8014, "name": "门套窗套豪迈电子据优化"},
}

# 木皮优化群 chat_id（该群支持文本命令管理剔除清单）
DOOR_SKIN_CHAT_ID = "oc_9067ef06c46495ac35328d89bc16017d"
DOOR_SKIN_PORT = 8010

# PVC优化群 chat_id（该群支持文本命令触发完整流程）
PVC_OPTIMIZE_CHAT_ID = "oc_51479339eef6b26fe9dcdcb8a5fb0c50"
PVC_TRIGGER_URL = os.getenv("PVC_TRIGGER_URL", "http://127.0.0.1:8012/process")
# 2026-11-10：8012 自动日报流程废弃（8-10 卡订单通验证码后停用，50 天无人触发）。
# 代码保留不删，便于回滚：把 .env 的 PVC_TRIGGER_MODE 改成 on 并重启 hub 即恢复原行为。
PVC_TRIGGER_MODE = os.getenv("PVC_TRIGGER_MODE", "off").strip().lower()

# PVC2联思群 chat_id（2026-10-20 接入：群里 @机器人 +「处理」跑联思每日流程）
# 触发器跑在宿主机 8015（liansi_batch_service/liansi_trigger_service.py），
# 它自己负责回群消息 —— 网关只转发，不参与联思的文案与发送。
LIANSI_CHAT_ID = "oc_d0ca4c6ad6df4190397012d6a3c6ed5c"
LIANSI_TRIGGER_URL = os.getenv("LIANSI_TRIGGER_URL", "http://127.0.0.1:8015/process")
# 「新增」：当天流程已跑完、订单通又新加了批次时，只补导这一批（只补不重导）
LIANSI_ADD_URL = os.getenv("LIANSI_ADD_URL", "http://127.0.0.1:8015/add")
# 「重导 单号/批次」：退单改完重拆单、或单掉了时只重导这一批（2026-11-09 加）
LIANSI_RE_URL = os.getenv("LIANSI_RE_URL", "http://127.0.0.1:8015/re")

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

# 批量收集配置：窗口从「最后一份入队」起算，中途继续传会重新计时
BATCH_COLLECTION_WINDOW = 10  # 秒
# 按群覆盖收集窗口。下车间单 8006 已于 2026-11-10 改走持久化文件队列（见 _queue_*），
# 不再经过这里的攒批与定时器，所以不再有此群条目。
BATCH_WINDOW_PER_PORT = {}  # 秒；仅对仍走定时器攒批的群生效
BATCH_MAX_FILES = 40    # 一次上传的合理上限，满了立即开跑不再等窗口
_batch_queues = {}      # chat_id -> [{file_path, file_name, message_id, file_key, time}]
_batch_timers = {}      # chat_id -> Timer
_batch_lock = threading.Lock()
_batch_run_locks = {}   # chat_id -> threading.Lock，保证同一群同时只跑一批

# 批量服务调用超时按文件数放宽：单份转换 + 逐格比对实测约 50s（转换 25-35s、
# 改动比对 20-25s），给 75s/份余量；下限保底 600s，上限 1800s ——
# 卡死的请求会堵住后面排队的批次，不能无限放宽。
# 2026-11-06 教训：17 份按旧值 30s/份 只算到 600s 下限，实际跑了 859s 被判超时，
# 服务端白干 14 分钟后结果没人收。改按实测值给足余量。
BATCH_TIMEOUT_PER_FILE = 75   # 秒
BATCH_TIMEOUT_MIN = 600       # 秒
BATCH_TIMEOUT_MAX = 1800      # 秒

# 批量请求的瞬时故障重试。只针对"连不上"这类错误（网络不可达/连接被拒/连接重置），
# 它们意味着请求根本没进到服务里，重试不会造成重复处理。
# 不重试 HTTP 超时：超时说明服务已经在跑，重试会白跑一遍、翻倍占 CPU、还可能产出两份文件。
# 2026-11-07 教训：为切 8008 重启适配器栈，16:53:07 有单撞进 20s 的端口空窗，
# 网关拿到 Errno 101 就直接失败，群里只收到一句报错，业务得手动重发。
BATCH_RETRY_DELAYS = (5, 10, 15)  # 秒，累计约 30s，覆盖服务重启窗口
_RETRYABLE_CONNECT_ERRNOS = {101, 104, 111, 113}

# WebSocket 假死检测
_last_activity_time = None  # None = 尚未收到过任何消息

HTTP_SERVICE_HOST = os.getenv("HTTP_SERVICE_HOST", "127.0.0.1")
_user_name_cache = {}
_chat_members_cache = {}

# 日志时间与追溯号统一用北京时间：容器默认 UTC，业务在群里看到的是 17:xx，
# 日志却写 09:xx，对不上号、没法按时间反查。可用 TZ 环境变量覆盖。
if not os.getenv("TZ"):
    os.environ["TZ"] = "Asia/Shanghai"
try:
    time.tzset()
except AttributeError:  # Windows 本机调试没有 tzset
    pass

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
# 错误码 + 追溯号
#
# 群里只发「错误码/trace + 原因 + 建议」，完整异常链、目标地址、尝试次数全部进日志。
# 业务在群里看到 GW-E1001/1107170240-3f9a，直接
#   grep '1107170240-3f9a' /app/logs/feishu_bot_ws.log
# 就能拿到那一次处理的全部上下文。
#
# 码段划分：
#   GW-E1xxx 网络连接   —— 请求没进到服务，重发大概率能好
#   GW-E2xxx 服务响应   —— 服务接到了但报错（4xx/5xx/返回体坏了）
#   GW-E3xxx 业务结果   —— 服务正常返回但 success=False / 没产物
#   GW-E4xxx 飞书侧     —— 下载、上传、发消息失败
#   GW-E5xxx 本地文件   —— 待处理文件丢了/没权限（2026-10-31 出过：旧批跑完删了新批文件）
#   GW-E9xxx 兜底       —— 未预期异常
# ---------------------------------------------------------------------------
ERROR_CODES = {
    "GW-E1001": ("服务连不上（网络不可达）", "多半是服务正在重启，等 1 分钟重发即可"),
    "GW-E1002": ("服务连不上（连接被拒绝）", "服务没启动，请联系管理员"),
    "GW-E1003": ("服务连接被重置", "等 1 分钟重发；反复出现请联系管理员"),
    "GW-E1004": ("服务所在主机不可达", "网络或服务异常，请联系管理员"),
    "GW-E1005": ("处理超时", "服务可能仍在处理，先确认有没有收到结果，再决定是否重发"),
    "GW-E1006": ("网络异常（未分类）", "等 1 分钟重发；反复出现请联系管理员"),
    "GW-E2001": ("服务拒绝了本次请求", "多为文件内容不符合要求，按原因修改后重发"),
    "GW-E2002": ("服务内部报错", "服务侧问题，请联系管理员并提供错误代码"),
    "GW-E2003": ("服务返回内容无法解析", "服务侧问题，请联系管理员并提供错误代码"),
    "GW-E3001": ("服务返回处理失败", "按原因处理；看不明白请联系管理员并提供错误代码"),
    "GW-E3002": ("处理完成但没有生成结果文件", "确认文件内容是否完整后重发"),
    "GW-E4001": ("飞书文件下载失败", "重发一次；反复出现请联系管理员"),
    "GW-E4002": ("飞书文件上传/发送失败", "重发一次；反复出现请联系管理员"),
    "GW-E5001": ("本地待处理文件读取失败", "文件可能已被清理，请把文件重新发一次"),
    "GW-E9001": ("未预期异常", "请联系管理员并提供错误代码"),
}

_ERRNO_TO_CODE = {101: "GW-E1001", 104: "GW-E1003", 111: "GW-E1002", 113: "GW-E1004"}
_FILE_ERRNOS = {2, 13, 21}  # ENOENT / EACCES / EISDIR

_trace_local = threading.local()


def _new_trace_id():
    """追溯号：月日时分秒 + 4 位随机，带时间便于按时间段反查日志。"""
    return datetime.now().strftime("%m%d%H%M%S") + "-" + uuid.uuid4().hex[:4]


def _begin_trace():
    """给一个处理任务开新的追溯号。处理都跑在自己的线程里，用线程本地保存。"""
    trace_id = _new_trace_id()
    _trace_local.trace_id = trace_id
    return trace_id


def _get_trace():
    trace_id = getattr(_trace_local, "trace_id", None)
    if not trace_id:
        trace_id = _begin_trace()
    return trace_id


class GatewayError(Exception):
    """带错误码和追溯号的失败。日志记全量，群消息只记摘要。"""

    def __init__(self, code, reason, hint=None, detail="", attempts=1, cause=None):
        self.code = code
        self.reason = reason
        self.hint = hint or ERROR_CODES.get(code, ("", ""))[1]
        self.detail = detail
        self.attempts = attempts
        self.trace_id = _get_trace()
        super().__init__(f"[{code}/{self.trace_id}] {reason}")
        if cause is not None:
            self.__cause__ = cause

    @property
    def label(self):
        return f"{self.code}/{self.trace_id}"

    def group_text(self, action):
        lines = [f"❌ {action}失败", f"错误代码：{self.label}", f"原因：{self.reason}"]
        if self.attempts > 1:
            lines.append(f"重试：自动重试 {self.attempts - 1} 次仍未成功")
        if self.hint:
            lines.append(f"处理建议：{self.hint}")
        return "\n".join(lines)


def _exception_detail(exc):
    """把异常链压成一行写日志：每层类型、消息、errno。"""
    parts = []
    for e in _iter_exception_chain(exc):
        piece = type(e).__name__
        text = str(e).strip().replace("\n", " ")
        if text:
            piece += f": {text[:300]}"
        errno = getattr(e, "errno", None)
        if errno is None and isinstance(getattr(e, "reason", None), OSError):
            errno = getattr(e.reason, "errno", None)
        if errno is not None:
            piece += f" [errno={errno}]"
        parts.append(piece)
    return " <- ".join(parts)


def _classify_error(exc):
    """异常 → (错误码, 一句原因)。顺序即优先级：HTTPError 必须在 URLError 之前。"""
    chain = list(_iter_exception_chain(exc))
    for e in chain:
        if isinstance(e, GatewayError):
            return e.code, e.reason
    for e in chain:
        if isinstance(e, urllib_request.HTTPError):
            if e.code >= 500:
                return "GW-E2002", f"服务内部报错（HTTP {e.code}）"
            return "GW-E2001", f"服务拒绝了本次请求（HTTP {e.code}）"
    for e in chain:
        if isinstance(e, json.JSONDecodeError):
            return "GW-E2003", "服务返回内容无法解析"
    for e in chain:
        if isinstance(e, (TimeoutError, socket.timeout)):
            return "GW-E1005", ERROR_CODES["GW-E1005"][0]
        errno = getattr(e, "errno", None)
        if errno is None and isinstance(getattr(e, "reason", None), OSError):
            errno = getattr(e.reason, "errno", None)
        code = _ERRNO_TO_CODE.get(errno)
        if code:
            return code, ERROR_CODES[code][0]
    for e in chain:
        if isinstance(e, (FileNotFoundError, PermissionError, IsADirectoryError)):
            return "GW-E5001", f"{ERROR_CODES['GW-E5001'][0]}（{type(e).__name__}）"
        if isinstance(e, OSError) and getattr(e, "errno", None) in _FILE_ERRNOS:
            return "GW-E5001", ERROR_CODES["GW-E5001"][0]
    for e in chain:
        if isinstance(e, (urllib_request.URLError, OSError)):
            return "GW-E1006", ERROR_CODES["GW-E1006"][0]
    return "GW-E9001", ERROR_CODES["GW-E9001"][0]


def _to_gateway_error(exc, attempts=1):
    """任意异常 → GatewayError（复用已有 trace，不重复开号）。"""
    if isinstance(exc, GatewayError):
        exc.attempts = max(exc.attempts, attempts)
        return exc
    code, reason = _classify_error(exc)
    return GatewayError(code, reason, detail=_exception_detail(exc),
                        attempts=attempts, cause=exc)


def _log_failure(chat_id, action, err, extra=""):
    logger.error("[%s] %s失败 | code=%s | trace=%s | attempts=%d | detail=%s%s",
                 chat_id, action, err.code, err.trace_id, err.attempts,
                 err.detail or err.reason, f" | {extra}" if extra else "")


def _fail_from_exception(chat_id, action, exc, attempts=1, extra="", code=None):
    """异常失败统一出口：日志留全量 + 群里发错误码。返回 GatewayError。

    code 用于把异类环节归到固定码段（例如飞书接口失败统一挂 GW-E4xxx），
    底层原因仍保留在 detail 里，排查不受影响。
    """
    err = _to_gateway_error(exc, attempts=attempts)
    if code and err.code != code:
        err = GatewayError(code, err.reason, detail=err.detail,
                           attempts=err.attempts, cause=exc)
    _log_failure(chat_id, action, err, extra)
    try:
        _send_text(chat_id, err.group_text(action))
    except Exception:
        logger.exception("[%s] 失败提示发送失败 | code=%s | trace=%s", chat_id, err.code, err.trace_id)
    return err


def _fail_with_code(chat_id, action, code, reason, detail="", extra="", reply_to=""):
    """业务失败（服务正常返回但结果不对）统一出口。reply_to 给定时引用回复原上传消息。"""
    err = GatewayError(code, reason, detail=detail or reason)
    _log_failure(chat_id, action, err, extra)
    try:
        _send_text_paired(chat_id, err.group_text(action), reply_to)
    except Exception:
        logger.exception("[%s] 失败提示发送失败 | code=%s | trace=%s", chat_id, err.code, err.trace_id)
    return err


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


def _extract_sender_open_id(data) -> str:
    try:
        sender = data.event.sender
        sender_id = getattr(sender, "sender_id", None)
        return (getattr(sender_id, "open_id", None) or "") if sender_id is not None else ""
    except Exception:
        return ""


def _lookup_user_name_by_contact(open_id: str) -> str:
    if not open_id:
        return ""
    token = _get_tenant_access_token()
    url = (
        "https://open.feishu.cn/open-apis/contact/v3/users/"
        f"{open_id}?user_id_type=open_id"
    )
    req = urllib_request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib_request.urlopen(req, timeout=10) as resp:
        res = json.loads(resp.read().decode("utf-8"))
    if res.get("code") != 0:
        return ""
    user = (res.get("data") or {}).get("user") or {}
    return user.get("name") or user.get("nickname") or ""


def _load_chat_member_names(chat_id: str) -> dict:
    cached = _chat_members_cache.get(chat_id)
    if cached:
        return cached
    token = _get_tenant_access_token()
    names = {}
    page_token = ""
    while True:
        url = (
            f"https://open.feishu.cn/open-apis/im/v1/chats/{chat_id}/members"
            f"?member_id_type=open_id&page_size=100"
        )
        if page_token:
            url += f"&page_token={page_token}"
        req = urllib_request.Request(url, headers={"Authorization": f"Bearer {token}"})
        with urllib_request.urlopen(req, timeout=15) as resp:
            res = json.loads(resp.read().decode("utf-8"))
        if res.get("code") != 0:
            logger.warning("拉取群成员失败 chat=%s: %s", chat_id, res)
            break
        data = res.get("data") or {}
        for item in data.get("items") or []:
            member_id = item.get("member_id") or ""
            name = item.get("name") or ""
            if member_id and name:
                names[member_id] = name
        if not data.get("has_more"):
            break
        page_token = data.get("page_token") or ""
        if not page_token:
            break
    _chat_members_cache[chat_id] = names
    return names


def _resolve_uploader_name(data, chat_id: str) -> str:
    open_id = _extract_sender_open_id(data)
    if open_id and open_id in _user_name_cache:
        return _user_name_cache[open_id]
    name = ""
    if open_id:
        try:
            name = _lookup_user_name_by_contact(open_id)
        except Exception:
            logger.warning("通讯录反查姓名失败 open_id=%s", open_id, exc_info=True)
        if not name:
            try:
                name = _load_chat_member_names(chat_id).get(open_id, "")
            except Exception:
                logger.warning("群成员反查姓名失败 chat=%s open_id=%s", chat_id, open_id, exc_info=True)
    name = (name or "").strip() or "未知上传者"
    if open_id:
        _user_name_cache[open_id] = name
    logger.info("[%s] 上传者: %s (%s)", chat_id, name, open_id or "-")
    return name


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


# ---------------------------------------------------------------------------
# 入群文件存档（旁路，2026-11-06）：下载成功的文件复制一份到挂载盘
# logs/archive/<群名>/<日期>/<时分秒>_<文件名>；同步脚本（Windows 计划任务）
# 另行推送到 Linux 服务器 platform/samples/gateway-archive/。
# 全程 try/except：存档失败只记 WARNING，绝不影响正常派发（路由逻辑零改动）。
# ---------------------------------------------------------------------------
ARCHIVE_ROOT = "/app/logs/archive"


def _archive_file_copy(local_path: str, file_name: str, chat_id: str) -> None:
    """把下载成功的入群文件复制一份到存档目录。失败不影响主流程。"""
    try:
        route = CHAT_ROUTES.get(chat_id) or {}
        group_name = str(route.get("name") or chat_id)
        safe_group = re.sub(r'[\\/:*?"<>|]', "_", group_name)[:60] or "unknown_group"
        day_dir = os.path.join(ARCHIVE_ROOT, safe_group, datetime.now().strftime("%Y-%m-%d"))
        os.makedirs(day_dir, exist_ok=True)
        ts = datetime.now().strftime("%H%M%S")
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", str(file_name))[:120] or "unnamed"
        dest = os.path.join(day_dir, f"{ts}_{safe_name}")
        if os.path.exists(dest):
            dest = os.path.join(day_dir, f"{ts}_{uuid.uuid4().hex[:6]}_{safe_name}")
        shutil.copy2(local_path, dest)
        logger.info("[archive] 已存档入群文件: %s", dest)
    except Exception as e:
        logger.warning("[archive] 存档失败（不影响处理）: %s", e)


def _fetch_message_files(message_id: str) -> list:
    """多文件合并卡片（msg_type=post）：事件 content 没带 files 时，回源 API 取文件清单"""
    token = _get_tenant_access_token()
    url = f"https://open.feishu.cn/open-apis/im/v1/messages/{message_id}"
    req = urllib_request.Request(url, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib_request.urlopen(req, timeout=15) as resp:
            data = json.load(resp)
        items = data.get("data", {}).get("items", [])
        if not items:
            return []
        body_content = items[0].get("body", {}).get("content", "")
        content_json = json.loads(body_content) if isinstance(body_content, str) else body_content
        return [f for f in (content_json.get("files") or []) if not f.get("is_folder")]
    except Exception:
        logger.exception("拉取多文件消息内容失败: %s", message_id)
        return []


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


# PAIRMSG_V1：多份文件同时处理时，结果文件与"这是哪一份的"说明必须绑在一起。
# ① 每条结果都引用回复用户上传的那条原消息，天然配对，不怕并发乱序；
# ② 同一群的"说明 + 文件"用一把锁串起来发，避免两个消费者线程交叉插入。
def _reply_feishu_message(parent_message_id: str, msg_type: str, content: dict) -> str | None:
    """引用回复指定消息，返回新消息 id；失败返回 None（调用方自行回退为普通发送）。"""
    if not parent_message_id:
        return None
    try:
        token = _get_tenant_access_token()
        url = f"https://open.feishu.cn/open-apis/im/v1/messages/{parent_message_id}/reply"
        body = json.dumps({
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
            logger.warning("回复消息失败（回退普通发送）: %s", res)
            return None
        return (res.get("data") or {}).get("message_id")
    except Exception as e:
        logger.warning("回复消息异常（回退普通发送）: %s", e)
        return None


def _send_text_paired(chat_id: str, text: str, parent_message_id: str = "") -> None:
    """成组文本：优先引用回复原消息，失败则普通发送。"""
    if parent_message_id and _reply_feishu_message(parent_message_id, "text", {"text": text}):
        return
    _send_text(chat_id, text)


def _send_file_paired(chat_id: str, file_key: str, parent_message_id: str = "") -> None:
    """成组文件：优先引用回复原消息，失败则普通发送。"""
    if parent_message_id and _reply_feishu_message(parent_message_id, "file", {"file_key": file_key}):
        return
    _send_file(chat_id, file_key)


_group_send_locks: dict = {}
_group_send_locks_guard = threading.Lock()


def _group_send_lock(chat_id: str) -> "threading.Lock":
    """同一群共用一把发送锁：一个文件的「说明 + 文件」必须连续发出，不许别的件插进来。"""
    with _group_send_locks_guard:
        lock = _group_send_locks.get(chat_id)
        if lock is None:
            lock = threading.Lock()
            _group_send_locks[chat_id] = lock
        return lock


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
    try:
        with opener.open(req, timeout=120) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib_request.HTTPError as e:
        # 服务端把真实原因放在响应体里（FastAPI 用 detail；http_skeleton 用 error）。
        # 不读回来的话只会抛出 "HTTP Error 500: Internal Server Error" 这种没用的话。
        # 注意：HTTPError 是 URLError 的子类，本分支必须排在前面。
        detail = ""
        try:
            if e.fp is not None:
                payload = json.loads(e.read().decode("utf-8"))
                raw = payload.get("detail") or payload.get("error") or payload.get("message")
                if raw:
                    # FastAPI 的 422 会把 detail 返回成列表，不能直接当字符串用
                    detail = raw if isinstance(raw, str) else json.dumps(raw, ensure_ascii=False)
        except Exception:
            pass
        finally:
            try:
                e.close()
            except Exception:
                pass
        detail = detail.strip()
        raise RuntimeError(f"{detail}（服务返回 {e.code}）" if detail
                           else f"服务返回 {e.code} {e.reason}") from e
    except (urllib_request.URLError, TimeoutError) as e:
        raise RuntimeError(
            f"无法连接处理服务（端口 {port}）：{getattr(e, 'reason', e)}") from e


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
    if PVC_TRIGGER_MODE != "on":
        logger.info("[%s] PVC 触发已停用（PVC_TRIGGER_MODE=%s），仅回提示语", chat_id, PVC_TRIGGER_MODE)
        _send_text(
            chat_id,
            "ℹ️「处理」这条自动日报流程已下线。\n"
            "请直接上传南兴优化清单 CSV，小虾会自动分类并把 zip 发回来。",
        )
        return
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
        _fail_from_exception(chat_id, "PVC 流程触发", e, extra=f"url={PVC_TRIGGER_URL}")


def _handle_liansi_process_command(chat_id: str, raw_text: str, sender: str,
                                   url: str = LIANSI_TRIGGER_URL) -> None:
    """处理「PVC2联思」群文本命令：把请求转给宿主机 8015 的联思触发器。

    网关只管转发；跑流程、回群消息全在 liansi_trigger_service.py 里
    （liansi_batch_service 的 AGENTS.md 禁止流程本身发飞书，发送代码只准待在那里）。
    url 两种：/process = 跑当天全流程；/add = 只补导订单通上新增的批次（2026-10-24 加）。
    """
    logger.info("[%s] 收到联思命令（%s）: %s (sender=%s)",
                chat_id, url.rsplit("/", 1)[-1], raw_text, sender)
    try:
        body = json.dumps({
            "chat_id": chat_id,
            "text": raw_text,
            "sender": sender,
        }, ensure_ascii=False).encode("utf-8")
        req = urllib_request.Request(
            url,
            data=body,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )
        with urllib_request.urlopen(req, timeout=15) as resp:
            res = json.loads(resp.read().decode("utf-8"))
        if not res.get("success"):
            _send_text(chat_id, f"⚠️ 联思触发失败：{res.get('error', '未知错误')}")
    except Exception as e:
        logger.exception("[%s] 调用联思触发服务失败", chat_id)
        _fail_from_exception(chat_id, "联思流程触发", e, extra=f"url={url}")


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
        _fail_from_exception(chat_id, "剔除命令处理", e)


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
        8014: "门套窗套豪迈电子据优化",
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
        item = {"file_content": b64, "filename": f["file_name"]}
        if f.get("uploader_name"):
            item["uploader_name"] = f["uploader_name"]
        payload_files.append(item)
    payload = {"files": payload_files}
    if order_date:
        payload["order_date"] = order_date
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib_request.Request(url, data=body, headers={
        "Content-Type": "application/json; charset=utf-8",
    })
    opener = urllib_request.build_opener(urllib_request.ProxyHandler({}))
    timeout = min(BATCH_TIMEOUT_MAX,
                  max(BATCH_TIMEOUT_MIN, BATCH_TIMEOUT_PER_FILE * len(files)))
    logger.info("批量请求 %d 个文件，HTTP 超时 %ds", len(files), timeout)
    with opener.open(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _iter_exception_chain(exc: Exception):
    """沿 __cause__ / __context__ 遍历异常链，避免自引用死循环。"""
    seen = set()
    cur = exc
    while cur is not None and id(cur) not in seen:
        seen.add(id(cur))
        yield cur
        cur = cur.__cause__ or cur.__context__


def _is_connect_error(exc: Exception) -> bool:
    """单层判断：是不是"连不上服务"。"""
    if isinstance(exc, urllib_request.HTTPError):
        # 服务已经把响应发回来了（4xx/5xx），不是连接问题
        return False
    if isinstance(exc, (TimeoutError, socket.timeout)):
        return False
    if isinstance(exc, urllib_request.URLError):
        reason = getattr(exc, "reason", None)
        if isinstance(reason, OSError):
            errno = getattr(reason, "errno", None)
            # errno 为 None 的多是 DNS/取址阶段失败，请求同样没发出去
            return errno in _RETRYABLE_CONNECT_ERRNOS or errno is None
        return isinstance(reason, str)
    return isinstance(exc, OSError) and getattr(exc, "errno", None) in _RETRYABLE_CONNECT_ERRNOS


def _is_retryable_connect_error(exc: Exception) -> bool:
    """判断是否属于"连不上服务"这类可以安全重试的故障。

    只看连接建立阶段的错误。HTTP 超时不在此列 —— 超时意味着服务已经开始处理，
    重试会重跑一遍、产出两份文件。
    沿异常链判断：_call_local_service 会把底层网络错误包成 RuntimeError。
    """
    return any(_is_connect_error(e) for e in _iter_exception_chain(exc))


def _friendly_error_text(exc: Exception) -> str:
    """一句话原因（含错误码）。真正发群走 GatewayError.group_text，这里留给脚本/调试用。"""
    code, reason = _classify_error(exc)
    return f"{reason}（{code}）"


def _with_connect_retry(chat_id: str, service_name: str, call, *args):
    """对 call(*args) 重试"连不上"这类瞬时故障，撑过服务重启窗口。

    call 必须是一次请求：失败于连接阶段时，重试不会造成重复处理。
    """
    attempts = len(BATCH_RETRY_DELAYS) + 1
    for attempt in range(1, attempts + 1):
        try:
            return call(*args)
        except Exception as exc:
            retryable = _is_retryable_connect_error(exc)
            if not retryable or attempt >= attempts:
                # 统一带上尝试次数：群里据此判断是不是"刚重启又撞上了"
                raise _to_gateway_error(exc, attempts=attempt) from exc
            delay = BATCH_RETRY_DELAYS[attempt - 1]
            code, reason = _classify_error(exc)
            logger.warning("[%s] %s 第 %d/%d 次请求失败 | code=%s | trace=%s | %s | detail=%s，%d 秒后重试",
                           chat_id, service_name, attempt, attempts,
                           code, _get_trace(), reason, _exception_detail(exc), delay)
            time.sleep(delay)
    raise RuntimeError("重试逻辑异常")


def _call_batch_service_with_retry(chat_id: str, port: int, service_name: str,
                                   files: list, order_date: str = None) -> dict:
    """调用批量服务，遇到"连不上"自动重试几次。"""
    return _with_connect_retry(chat_id, service_name, _call_batch_service,
                               port, files, order_date)


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


def _send_output_item(chat_id: str, item, content_log: str, path_log: str, warn_missing: bool, exc_info: bool,
                      reply_to: str = "") -> int:
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
        _send_file_paired(chat_id, fk, reply_to)
        logger.info("[%s] 已发送文件: %s", chat_id, filename)
        return 1
    except Exception as e:
        logger.error("[%s] 发送文件失败: %s", chat_id, e, exc_info=exc_info)
        return 0


def _send_output_pairs(chat_id: str, output_pairs, content_log: str, path_log: str, warn_missing: bool, exc_info: bool,
                       reply_to: str = "") -> int:
    sent_count = 0
    for pair in output_pairs:
        for item in pair:
            sent_count += _send_output_item(chat_id, item, content_log, path_log, warn_missing, exc_info,
                                            reply_to=reply_to)
        # Keep a short pause between pairs so Feishu preserves the visible order.
        time.sleep(0.5)
    return sent_count


def _output_filenames(output_pairs) -> list:
    """输出对里每个文件的名字，用于在成组说明里回显"生成了什么"。"""
    names = []
    for pair in output_pairs:
        for item in pair:
            if isinstance(item, dict):
                name = item.get("filename") or os.path.basename(item.get("path", "") or "")
            else:
                name = os.path.basename(str(item))
            if name:
                names.append(os.path.basename(name))
    return names


def _format_quantity_value(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _build_quantity_card(result: dict, service_name: str, note: str = None) -> dict | None:
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

    elements = []
    if note:
        elements.append({"tag": "div", "text": {"tag": "lark_md", "content": note}})
    elements.append(
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
    )

    return {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "green",
            "title": {"tag": "plain_text", "content": f"{service_name}数量明细"},
        },
        "elements": elements,
    }


def _send_quantity_card(chat_id: str, result: dict, port: int, service_name: str, note: str = None) -> bool:
    """发送数量卡片。note 提供时合并进卡片（完成文案+警告），卡片失败则合并成一条文本兜底。
    返回 True 表示消息已发出（含兜底），调用方无需再补发 note。"""
    if port != 8005:
        return False
    card = _build_quantity_card(result, service_name, note=note)
    if not card:
        return False
    try:
        _send_feishu_message(chat_id, "interactive", card)
        return True
    except Exception as e:
        logger.error("[%s] 发送数量表格卡片失败: %s", chat_id, e, exc_info=True)
        value = _format_quantity_value(result.get("quantity_total"))
        text = f"本次分类数量合计：{value}"
        if note:
            text = note + "\n\n" + text
        _send_text(chat_id, text)
        return True


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


def _result_field(result: dict, key: str) -> str:
    """取服务响应里的一个文本字段：顶层优先，其次 results[]（单文件批处理时字段在 results 里）。"""
    value = result.get(key)
    if isinstance(value, str) and value.strip():
        return value.strip()
    for res in result.get("results") or []:
        if isinstance(res, dict):
            value = res.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()
    return ""


def _compact_warnings(warnings: list[str], limit: int = 2, width: int = 90) -> list[str]:
    """提醒只留最要紧的几条，一条一行、太长就截断——群里不该让人读长文。"""
    out: list[str] = []
    for item in warnings:
        text = " ".join(str(item).split())
        if not text:
            continue
        out.append(text if len(text) <= width else text[:width] + "…")
        if len(out) >= limit:
            break
    if len(warnings) > len(out):
        out.append(f"…另有 {len(warnings) - len(out)} 条提醒")
    return out


def _batch_file_path(save_dir: Path, file_name: str) -> Path:
    """批量收集落盘用的唯一文件名。

    不能直接用原文件名：同名重传会写同一路径，于是
    ① 上一批处理完在 finally 里按路径删除时，会把刚重传、已被新批排队引用的文件删掉
       → 新批读文件报 FileNotFoundError（2026-10-31 09:26 的
       `S2610-7564-4-4 欧派 OP成都市HZJ261000074YA10-9.xlsx` 即此）；
    ② 即使没删，新内容也会覆盖旧内容，正在跑的旧批会处理错文件。
    对外展示的文件名仍用 file_name（见 _call_batch_service 的 filename 字段）。
    """
    stem = Path(file_name).stem[:80]
    suffix = Path(file_name).suffix
    return save_dir / f"{stem}__{int(time.time() * 1000)}_{uuid.uuid4().hex[:6]}{suffix}"


def _sweep_stale_batch_files(save_dir: Path, max_age_seconds: int = 48 * 3600) -> None:
    """清掉目录里遗留的旧文件（进程异常退出时不会走到 finally 清理）。"""
    cutoff = time.time() - max_age_seconds
    try:
        for item in save_dir.iterdir():
            try:
                if item.is_file() and item.stat().st_mtime < cutoff:
                    item.unlink()
            except OSError:
                continue
    except OSError:
        pass


# ---------------------------------------------------------------------------
# 持久化文件队列（下车间单 8006，2026-11-10；并发 + 漏单兜底 2026-11-10）
# 旧实现是「内存 dict + 定时器」攒批：容器重启/重建时内存队列与容器内 /tmp 里的
# 待处理文件一并消失，用户只看到"已收到"却永远等不到结果；服务返回失败也不重试。
# 现在改为「目录即队列」，落在已挂载的 /app/state/queue：
#   <chat>/pending/     {seq}_{uuid}_{原名} + 同名 .json sidecar（上传人、消息 ID）
#   <chat>/processing/  已被某个消费者原子领走（os.rename 抢占，抢不到的自然跳过）
#   <chat>/done/<日>/   成功件，保留可查、可重放
#   <chat>/failed/<日>/ 重试耗尽的件
#   <chat>/journal/     入队/完成/补发/对账的 jsonl 台账，出问题按件号回查
# 每群 QUEUE_CONCURRENCY 个消费者线程，一次只取一份（绝不合并成批），失败按
# QUEUE_RETRY_DELAYS 重试。另有一个 reconcile 线程专管漏单兜底：
#   1) 启动时 processing 里全部退回 pending（容器重启/重建不丢件）
#   2) processing 中超过 QUEUE_STALE_SECONDS 没有终态 → 判卡死，退回 pending 重跑
#   3) pending 里 sidecar 缺失 → 按队列名还原文件名补回；.part 半成品 → 清理
#   4) pending 有件但消费者线程已退出 → 自动重新拉起
#   5) 每轮打印 pending/processing/今日 done/failed 计数，对账用
# reconcile 只做目录 stat 和 rename，不做重活、不阻塞消费者，不影响出件速度。
# ---------------------------------------------------------------------------
QUEUE_ROOT = Path(os.environ.get("QUEUE_ROOT", "/app/state/queue"))
QUEUE_MAX_ATTEMPTS = int(os.environ.get("QUEUE_MAX_ATTEMPTS", "3"))
QUEUE_RETRY_DELAYS = (5, 20)
QUEUE_POLL_IDLE = 1.0
QUEUE_CONCURRENCY = max(1, int(os.environ.get("QUEUE_CONCURRENCY", "2")))
QUEUE_STALE_SECONDS = int(os.environ.get("QUEUE_STALE_SECONDS", "1800"))
QUEUE_RECONCILE_SECONDS = max(10.0, float(os.environ.get("QUEUE_RECONCILE_SECONDS", "30")))
_queue_workers = {}          # chat_id -> [Thread, ...]
_queue_worker_lock = threading.Lock()
_queue_inflight = {}         # chat_id -> {队列件名: 原始文件名}
_queue_inflight_lock = threading.Lock()
_queue_reconciler = [None]   # 兜底对账线程单例
_queue_reconcile_stats = {"runs": 0, "requeued": 0, "repaired": 0, "spawned": 0, "last": None}


def _queue_dir(chat_id: str, sub: str) -> Path:
    d = QUEUE_ROOT / chat_id / sub
    d.mkdir(parents=True, exist_ok=True)
    return d


def _queue_orig_name(qname: str) -> str:
    """从 {seq}_{uuid}_{原名} 里还原原始文件名（sidecar 丢了时的兜底）。"""
    parts = qname.split("_", 2)
    return parts[2] if len(parts) == 3 else qname


def _queue_items(chat_id: str, sub: str = "pending") -> list:
    """某子目录里待处理的件（排除 sidecar / claim / 半成品），按 seq 升序。"""
    try:
        d = QUEUE_ROOT / chat_id / sub
        if not d.is_dir():
            return []
        return sorted((p for p in d.iterdir()
                       if not p.name.endswith((".json", ".part", ".claim"))),
                      key=lambda p: p.name)
    except OSError:
        return []


def _short_name(file_name: str) -> str:
    """消息里显示的文件名：去扩展名、压缩连续空格，省得一行字太挤。"""
    stem = Path(file_name or "").stem or file_name or "文件"
    return re.sub(r"\s+", " ", stem).strip()


def _queue_inflight_count(chat_id: str) -> int:
    with _queue_inflight_lock:
        return len(_queue_inflight.get(chat_id) or {})


def _queue_journal(chat_id: str, event: str, **fields) -> None:
    """追加一行台账：谁在什么时候入队/完成/被补发，事后能逐件对账。"""
    try:
        d = _queue_dir(chat_id, "journal")
        rec = {"ts": round(time.time(), 3),
               "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
               "event": event}
        rec.update(fields)
        with (d / (datetime.now().strftime("%Y-%m-%d") + ".jsonl")).open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError:
        pass


def _queue_enqueue(chat_id, src_path, file_name, message_id, file_key, uploader_name) -> tuple:
    """落盘入队（原子 rename，崩溃不留半截），返回 (seq, 当前排队数)。"""
    pending = _queue_dir(chat_id, "pending")
    seq = time.time_ns()
    stem = Path(file_name).stem[:80]
    suffix = Path(file_name).suffix
    base = f"{seq:019d}_{uuid.uuid4().hex[:6]}_{stem}{suffix}"

    part = pending / (base + ".part")
    shutil.copy(src_path, part)
    os.replace(part, pending / base)

    meta = {
        "file_name": file_name,
        "message_id": message_id,
        "file_key": file_key,
        "uploader_name": uploader_name,
        "enqueued_at": time.time(),
    }
    meta_part = pending / (base + ".json.part")
    meta_part.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    os.replace(meta_part, pending / (base + ".json"))
    # 排队深度 = 还没取走的 + 正在处理的，并发下才说得准“前面还有几份”
    depth = len(_queue_items(chat_id)) + _queue_inflight_count(chat_id)
    logger.info("[%s] 入队 %s（seq=%d，待处理 %d 份）", chat_id, file_name, seq, depth)
    _queue_journal(chat_id, "enqueue", name=base, file_name=file_name,
                   uploader=uploader_name, message_id=message_id)
    return seq, depth


def _queue_meta(path: Path) -> dict:
    try:
        return json.loads((path.parent / (path.name + ".json")).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        logger.warning("队列 sidecar 缺失或损坏：%s", path.name)
        return {}


def _queue_claim(chat_id: str):
    """原子领取一份：os.rename 到 processing 成功者才算拿到，抢不到自动跳过。"""
    try:
        pending = QUEUE_ROOT / chat_id / "pending"
        proc = _queue_dir(chat_id, "processing")
    except OSError:
        return None
    for path in _queue_items(chat_id):
        dest = proc / path.name
        try:
            os.rename(path, dest)
        except FileNotFoundError:
            continue          # 被别的消费者先领走了
        except OSError:
            logger.exception("[%s] 队列领取失败：%s", chat_id, path.name)
            continue
        side = pending / (path.name + ".json")
        if side.exists():
            try:
                os.replace(side, proc / (path.name + ".json"))
            except OSError:
                pass
        try:
            (proc / (path.name + ".claim")).write_text(
                json.dumps({"claimed_at": time.time(), "thread": threading.current_thread().name}),
                encoding="utf-8")
        except OSError:
            pass
        return dest
    return None


def _queue_alive_workers(chat_id: str) -> int:
    ths = _queue_workers.get(chat_id) or []
    return sum(1 for t in ths if t.is_alive())


def _queue_ensure_workers(chat_id: str, port: int, service_name: str) -> None:
    """确保该群有 QUEUE_CONCURRENCY 个活着的消费者（幂等，可反复调用）。"""
    with _queue_worker_lock:
        ths = [t for t in (_queue_workers.get(chat_id) or []) if t.is_alive()]
        need = QUEUE_CONCURRENCY - len(ths)
        for i in range(need):
            idx = len(ths)
            th = threading.Thread(target=_queue_worker, args=(chat_id, port, service_name, idx),
                                  name="filequeue-%s-%d" % (chat_id[-6:], idx), daemon=True)
            ths.append(th)
            th.start()
        if need:
            _queue_workers[chat_id] = ths
            logger.info("[%s] 文件队列消费者 %d 个在线（新起 %d 个，port=%d, %s）",
                        chat_id, len(ths), need, port, service_name)


def _queue_worker(chat_id: str, port: int, service_name: str, idx: int) -> None:
    while True:
        try:
            path = _queue_claim(chat_id)
            if path is None:
                time.sleep(QUEUE_POLL_IDLE)
                continue
            _queue_process_one(chat_id, port, service_name, path)
        except Exception:
            logger.exception("[%s] 文件队列消费异常（worker #%d），2 秒后继续", chat_id, idx)
            time.sleep(2)


def _queue_process_one(chat_id, port, service_name, path: Path) -> None:
    meta = _queue_meta(path)
    file_name = meta.get("file_name") or _queue_orig_name(path.name)
    with _queue_inflight_lock:
        _queue_inflight.setdefault(chat_id, {})[path.name] = file_name
    try:
        waiting = len(_queue_items(chat_id))
        logger.info("[%s] 队列取件 %s（待处理 %d）", chat_id, file_name, waiting)
        # MSGTRIM_V1：一行一条，不再拼接「后面还有 N 份排队」
        _send_text(chat_id, "\U0001F527 处理中 · %s" % _short_name(file_name))

        file_info = {
            "file_path": str(path),
            "file_name": file_name,
            "message_id": meta.get("message_id", ""),
            "file_key": meta.get("file_key", ""),
            "uploader_name": meta.get("uploader_name", ""),
            "time": meta.get("enqueued_at") or time.time(),
        }

        ok = False
        for attempt in range(1, QUEUE_MAX_ATTEMPTS + 1):
            try:
                ok = _process_batch_locked(chat_id, port, service_name, [file_info], keep_files=True)
            except Exception:
                logger.exception("[%s] 「%s」第 %d 次处理异常", chat_id, file_name, attempt)
                ok = False
            if ok:
                break
            if attempt < QUEUE_MAX_ATTEMPTS:
                delay = QUEUE_RETRY_DELAYS[min(attempt - 1, len(QUEUE_RETRY_DELAYS) - 1)]
                logger.warning("[%s] 「%s」第 %d 次失败，%d 秒后重试", chat_id, file_name, attempt, delay)
                _send_text(chat_id, "⚠️ 「%s」处理失败，%d 秒后自动重试（第 %d/%d 次）"
                           % (file_name, delay, attempt + 1, QUEUE_MAX_ATTEMPTS))
                time.sleep(delay)

        _queue_finalize(chat_id, path, file_name, ok)
    finally:
        with _queue_inflight_lock:
            _queue_inflight.get(chat_id, {}).pop(path.name, None)


def _queue_finalize(chat_id, path: Path, file_name: str, ok: bool) -> None:
    day = datetime.now().strftime("%Y-%m-%d")
    dest = _queue_dir(chat_id, "done" if ok else "failed") / day
    dest.mkdir(parents=True, exist_ok=True)
    for suffix in ("", ".json", ".claim"):
        item = path.parent / (path.name + suffix) if suffix else path
        try:
            if item.exists():
                shutil.move(str(item), str(dest / item.name))
        except OSError:
            logger.exception("[%s] 队列归档失败：%s", chat_id, item.name)
    _queue_journal(chat_id, "done" if ok else "failed", name=path.name, file_name=file_name)
    if ok:
        logger.info("[%s] 「%s」处理完成，已归档", chat_id, file_name)
    else:
        logger.error("[%s] 「%s」重试 %d 次仍失败，已转 failed", chat_id, file_name, QUEUE_MAX_ATTEMPTS)
        _send_text(chat_id, "❌ 「%s」重试 %d 次仍失败，已留在 failed 目录，麻烦重新上传。"
                   % (file_name, QUEUE_MAX_ATTEMPTS))


def _queue_requeue_stale(chat_id: str, proc: Path, now: float) -> int:
    """processing 里卡死（超时没终态）的件退回 pending 重跑，返回退回份数。"""
    moved = 0
    pending = _queue_dir(chat_id, "pending")
    for path in _queue_items(chat_id, "processing"):
        claim = path.parent / (path.name + ".claim")
        try:
            claimed_at = float(json.loads(claim.read_text(encoding="utf-8"))["claimed_at"])
        except (OSError, ValueError, KeyError):
            try:
                claimed_at = path.stat().st_mtime
            except OSError:
                continue
        age = now - claimed_at
        if age <= QUEUE_STALE_SECONDS:
            continue
        try:
            os.rename(path, pending / path.name)
        except OSError:
            logger.exception("[%s] 卡死件退回失败：%s", chat_id, path.name)
            continue
        side = path.parent / (path.name + ".json")
        if side.exists():
            try:
                os.replace(side, pending / side.name)
            except OSError:
                pass
        try:
            claim.unlink()          # 领取标记清掉，下一轮重新计时
        except OSError:
            pass
        moved += 1
        logger.warning("[%s] 「%s」已 %d 秒无终态，判定卡死，退回队列重跑", chat_id,
                       _queue_orig_name(path.name), int(age))
        _queue_journal(chat_id, "requeue_stale", name=path.name, age_s=int(age))
    return moved


def _queue_repair_pending(chat_id: str, now: float) -> int:
    """修 pending：孤儿 sidecar 清理、缺 sidecar 补回、超期 .part 清理。"""
    fixed = 0
    try:
        pending = QUEUE_ROOT / chat_id / "pending"
        if not pending.is_dir():
            return 0
        names = {p.name for p in pending.iterdir()}
        for p in list(pending.iterdir()):
            try:
                if p.name.endswith(".part"):
                    if now - p.stat().st_mtime > 300:      # 入队中断 5 分钟后的半成品
                        p.unlink()
                        fixed += 1
                        logger.warning("[%s] 清理入队半成品 %s", chat_id, p.name)
                    continue
                if p.name.endswith(".json"):
                    body = p.name[:-5]
                    if body not in names and not p.name.endswith(".json.part"):
                        p.unlink()                          # 文件没了，sidecar 也没意义
                        fixed += 1
                        logger.warning("[%s] 清理孤儿 sidecar %s", chat_id, p.name)
                    continue
                if p.name.endswith(".claim"):
                    continue
                if (p.name + ".json") not in names:
                    meta = {"file_name": _queue_orig_name(p.name), "message_id": "",
                            "file_key": "", "uploader_name": "",
                            "enqueued_at": p.stat().st_mtime, "repaired": True}
                    (pending / (p.name + ".json")).write_text(
                        json.dumps(meta, ensure_ascii=False), encoding="utf-8")
                    fixed += 1
                    logger.warning("[%s] 补回缺失 sidecar：%s", chat_id, p.name)
                    _queue_journal(chat_id, "repair_sidecar", name=p.name)
            except OSError:
                continue
    except OSError:
        pass
    return fixed


def _queue_reconcile_once() -> dict:
    """一轮漏单兜底：修 pending、退卡死件、拉活消费者、输出计数。"""
    stats = {"requeued": 0, "repaired": 0, "spawned": 0, "chats": 0}
    if not QUEUE_ROOT.is_dir():
        return stats
    now = time.time()
    day = datetime.now().strftime("%Y-%m-%d")
    for chat_dir in sorted(QUEUE_ROOT.iterdir()):
        if not chat_dir.is_dir():
            continue
        chat_id = chat_dir.name
        cfg = CHAT_ROUTES.get(chat_id)
        if not cfg:
            continue
        stats["chats"] += 1
        stats["repaired"] += _queue_repair_pending(chat_id, now)
        stats["requeued"] += _queue_requeue_stale(chat_id, chat_dir / "processing", now)
        pending_n = len(_queue_items(chat_id))
        proc_n = len(_queue_items(chat_id, "processing"))
        done_n = len(_queue_items(chat_id, "done/" + day))
        fail_n = len(_queue_items(chat_id, "failed/" + day))
        alive = _queue_alive_workers(chat_id)
        if pending_n or proc_n:
            if pending_n and alive < min(QUEUE_CONCURRENCY, pending_n + proc_n):
                before = alive
                _queue_ensure_workers(chat_id, cfg["port"], cfg["name"])
                stats["spawned"] += max(0, _queue_alive_workers(chat_id) - before)
            logger.info("[%s] 队列对账 pending=%d processing=%d 今日done=%d failed=%d 消费者=%d",
                        chat_id, pending_n, proc_n, done_n, fail_n, _queue_alive_workers(chat_id))
        elif alive < QUEUE_CONCURRENCY and _queue_workers.get(chat_id):
            # 群里暂时没件也把消费者补齐，免得下一份上传还要现起线程
            _queue_ensure_workers(chat_id, cfg["port"], cfg["name"])
    return stats


def _queue_reconcile_loop() -> None:
    while True:
        time.sleep(QUEUE_RECONCILE_SECONDS)
        try:
            s = _queue_reconcile_once()
            _queue_reconcile_stats["runs"] += 1
            _queue_reconcile_stats["requeued"] += s["requeued"]
            _queue_reconcile_stats["repaired"] += s["repaired"]
            _queue_reconcile_stats["spawned"] += s.get("spawned", 0)
            _queue_reconcile_stats["last"] = s
            if s["requeued"] or s["repaired"]:
                logger.warning("队列兜底：退回卡死 %d 份、修复 %d 处", s["requeued"], s["repaired"])
        except Exception:
            logger.exception("队列对账异常（下轮继续）")


def _queue_start_reconciler() -> None:
    if _queue_reconciler[0] is not None and _queue_reconciler[0].is_alive():
        return
    th = threading.Thread(target=_queue_reconcile_loop, name="filequeue-reconciler", daemon=True)
    _queue_reconciler[0] = th
    th.start()
    logger.info("文件队列兜底对账已启动（每 %.0f 秒一轮，卡死阈值 %d 秒）",
                QUEUE_RECONCILE_SECONDS, QUEUE_STALE_SECONDS)


def _queue_restore() -> None:
    """启动时恢复：processing 全部退回 pending，再拉起消费者与兜底对账。"""
    try:
        if not QUEUE_ROOT.is_dir():
            logger.info("未发现队列目录 %s，跳过恢复", QUEUE_ROOT)
            _queue_start_reconciler()
            return
        for chat_dir in sorted(QUEUE_ROOT.iterdir()):
            if not chat_dir.is_dir():
                continue
            chat_id = chat_dir.name
            cfg = CHAT_ROUTES.get(chat_id)
            if not cfg:
                logger.warning("队列目录 %s 无对应群路由，跳过恢复", chat_id)
                continue
            # 容器重启/重建：上次处理到一半的件退回 pending 重跑（宁重不漏）
            restored = 0
            proc = chat_dir / "processing"
            if proc.is_dir():
                pending = _queue_dir(chat_id, "pending")
                for path in _queue_items(chat_id, "processing"):
                    side = proc / (path.name + ".json")
                    try:
                        os.rename(path, pending / path.name)
                        restored += 1
                    except OSError:
                        continue
                    if side.exists():
                        try:
                            os.replace(side, pending / side.name)   # sidecar 一起退回，保住上传人/消息ID
                        except OSError:
                            pass
                if restored:
                    for p in proc.glob("*.claim"):
                        try:
                            p.unlink()
                        except OSError:
                            pass
                    logger.warning("[%s] 重启前有 %d 份处理中断，已退回队列重跑", chat_id, restored)
                    _queue_journal(chat_id, "restore", count=restored)
            items = _queue_items(chat_id)
            _queue_ensure_workers(chat_id, cfg["port"], cfg["name"])
            if items:
                logger.info("[%s] 恢复 %d 份待处理文件（%s）", chat_id, len(items), cfg["name"])
                _send_text(chat_id, "♻️ 网关重启，发现 %d 份文件未处理完，已自动续跑。" % len(items))
        _queue_start_reconciler()
    except Exception:
        logger.exception("队列恢复失败（不影响网关启动）")
        _queue_start_reconciler()


def _process_batch(chat_id: str, port: int, service_name: str):
    """处理批量队列"""
    global _batch_queues, _batch_timers
    with _batch_lock:
        queue = _batch_queues.pop(chat_id, [])
        _batch_timers.pop(chat_id, None)

    if not queue:
        return

    # 同一群的批次串行执行：上一批没跑完就在此排队。
    # 否则两批会同时压在一个服务上抢 CPU，单份耗时翻倍，直至撞上 HTTP 超时。
    run_lock = _batch_run_locks.setdefault(chat_id, threading.Lock())
    if run_lock.locked():
        logger.info("[%s] 上一批仍在处理，本批 %d 个文件排队等待", chat_id, len(queue))
    run_lock.acquire()
    try:
        _process_batch_locked(chat_id, port, service_name, queue)
    finally:
        run_lock.release()


def _process_batch_locked(chat_id: str, port: int, service_name: str, queue: list,
                         keep_files: bool = False) -> bool:
    """_process_batch 的实际处理逻辑，调用方须已持有该群的运行锁。

    返回 True 表示已产出结果文件并发送成功，False 表示失败（已向群里报错）。
    keep_files=True 时保留原件不删，供持久化队列重试使用。
    """
    trace_id = _begin_trace()
    logger.info("[%s] 批量处理启动 trace=%s，共 %d 个文件，端口 %d，文件：%s",
                chat_id, trace_id, len(queue), port, [f["file_name"] for f in queue])
    order_date = f"{datetime.now().year}.{datetime.now().month}.{datetime.now().day}"

    # PAIRMSG_V1：下车间单一次只处理一份，结果必须和"这是哪一份的"绑在一起。
    single = len(queue) == 1
    pair_name = (queue[0].get("file_name") or "") if single else ""
    pair_msg = (queue[0].get("message_id") or "") if single else ""

    try:
        result = _call_batch_service_with_retry(chat_id, port, service_name, queue, order_date)
        logger.info("[%s] 批量处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            raw_error = str(result.get("error") or "服务未返回原因").strip()
            _fail_with_code(chat_id, f"{service_name}批量处理", "GW-E3001",
                            f"服务返回处理失败：{raw_error[:160]}",
                            detail=f"success=False raw_error={raw_error!r}",
                            extra=f"files={[f['file_name'] for f in queue]}",
                            reply_to=pair_msg)
            return False

        output_files = result.get("output_files", [])
        if not output_files:
            _fail_with_code(chat_id, f"{service_name}批量处理", "GW-E3002",
                            "处理完成但没有生成结果文件",
                            detail=f"success=True output_files=[] files={[f['file_name'] for f in queue]}",
                            reply_to=pair_msg)
            return False

        # 支持成对返回：dict {原始文件名: [file1, file2, ...]} 或旧格式列表
        output_pairs = _normalize_output_pairs(output_files)
        total_count = sum(len(pair) for pair in output_pairs)
        logger.info("[%s] 准备发送 %d 对输出文件，共 %d 个文件", chat_id, len(output_pairs), total_count)
        warnings = _collect_result_warnings(result)

        if port == 8006 and single:
            # 一份单：一条说明 + 紧跟该份结果文件，整组在群锁内连续发出，
            # 并引用回复用户上传的那条消息。
            # 多份同时处理时各自成组，不会再出现"文件排一串、消息排一串"对不上。
            # CHEAPMSG_V1：群里只留「单号 + 改了什么」，长版逐个格子的说明不再贴群
            # （8006 仍会写进日志和改动清单，车间要抠细节再翻）。
            brief = _result_field(result, "notice_brief")
            review = _result_field(result, "notice_review")
            if brief:
                lines = ["✅ " + brief]
                if review:
                    lines.append(review)
                lines.extend(_compact_warnings(warnings))
            else:
                # 老版本 8006 没给 notice_brief：退回旧格式，不丢信息
                lines = ["✅ %s · 完成" % (_short_name(pair_name) or "文件")]
                names = _output_filenames(output_pairs)
                if names:
                    lines.append("📄 " + "、".join(names))
                lines.extend(warnings)
            msg = "\n".join(lines)
            with _group_send_lock(chat_id):
                _send_text_paired(chat_id, msg, pair_msg)
                _send_output_pairs(chat_id, output_pairs, "从返回内容上传文件", "检查输出文件",
                                   True, True, reply_to=pair_msg)
            return True

        _send_output_pairs(chat_id, output_pairs, "从返回内容上传文件", "检查输出文件", True, True)

        count = total_count
        if port == 8006:
            # DONEMSG_V1：下车间单一次只处理一份，"共处理 1 个文件，生成 1 个结果
            # 文件"纯属废话，直接一行给结论。
            if len(queue) == 1:
                msg = "✅ %s · 完成" % _short_name(queue[0].get("file_name") or "")
            else:
                msg = "✅ %d 份已处理完成" % len(queue)
        else:
            msg = f"✅ {service_name}批量处理完成，共处理 {len(queue)} 个文件，生成 {count} 个结果文件，请检查。"
        if warnings:
            msg += "\n\n" + "\n".join(warnings)
        # 压缩包和对话卡放一起：完成文案合并进数量卡片， zip 与卡片相邻
        if not _send_quantity_card(chat_id, result, port, service_name, note=msg):
            _send_text(chat_id, msg)
        return True
    except Exception as e:
        logger.exception("[%s] 批量处理异常 trace=%s", chat_id, trace_id)
        _fail_from_exception(chat_id, f"{service_name}批量处理", e,
                             extra=f"files={[f['file_name'] for f in queue]}")
        return False
    finally:
        # 清理临时文件（持久化队列需保留原件重试，故 keep_files=True 时跳过）
        if not keep_files:
            for f in queue:
                try:
                    if os.path.exists(f["file_path"]):
                        os.remove(f["file_path"])
                except Exception:
                    pass

def _start_batch_timer(chat_id: str, port: int, service_name: str):
    """每来一份文件就从现在起重新等窗口，避免连传被拆成两批。调用方须已持有 _batch_lock。"""
    global _batch_timers
    old = _batch_timers.get(chat_id)
    if old is not None:
        old.cancel()
        _batch_timers[chat_id] = None

    def timer_callback():
        _process_batch(chat_id, port, service_name)

    window = BATCH_WINDOW_PER_PORT.get(port, BATCH_COLLECTION_WINDOW)
    timer = threading.Timer(window, timer_callback)
    timer.daemon = True
    _batch_timers[chat_id] = timer
    timer.start()


def _detect_may_sales_type(file_name: str, file_path: str = "") -> str:
    """识别 销售部业绩核对文件类型：先看文件内容（sheet 名/表头），再退回文件名。

    文件名是人起的，改名就认不出（2026-10 联思导出改成「经销商财务统计」、
    群里随手存成 xxx(1).xlsx 都踩过），所以只要有本地路径就先读内容判断。
    """
    if file_path and file_probe is not None:
        try:
            kind = file_probe.detect_by_content(file_path)
        except Exception:              # noqa: BLE001  读不了就当认不出
            kind = ""
        if kind:
            logger.info("文件内容识别「%s」→ %s", file_name, kind)
            return kind
    name = file_name.lower()
    if "综合查询" in name:
        return "zhcx"
    if "联思" in name:
        return "liansi"
    # 2026-10 起联思导出改叫「经销商财务统计」，文件名里没有「联思」
    if "财务统计" in name:
        return "liansi"
    if "奢匠" in name or "下单统计" in name or "线下" in name:
        return "shejiang"
    # 自己传上来的核对表/业绩数据表：只作参照，不参与计算（版式和脚本模板不同）
    if "业绩数据表" in name or "核对表" in name or "模板" in name:
        return "template"
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

    zhcx = [f for f in queue if _detect_may_sales_type(f["file_name"], f.get("file_path", "")) == "zhcx"]
    liansi = [f for f in queue if _detect_may_sales_type(f["file_name"], f.get("file_path", "")) == "liansi"]
    shejiang = [f for f in queue if _detect_may_sales_type(f["file_name"], f.get("file_path", "")) == "shejiang"]

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

    trace_id = _begin_trace()
    logger.info("[%s] 销售部业绩核对文件凑齐，开始处理 trace=%s: %s",
                chat_id, trace_id, [f["file_name"] for f in files])
    try:
        result = _call_batch_service(8003, files)
        logger.info("[%s] 处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            raw_error = str(result.get("error") or "服务未返回原因").strip()
            _fail_with_code(chat_id, f"{service_name}处理", "GW-E3001",
                            f"服务返回处理失败：{raw_error[:160]}",
                            detail=f"success=False raw_error={raw_error!r}",
                            extra=f"files={[f['file_name'] for f in files]} port=8003")
            return

        output_files = result.get("output_files", [])
        if not output_files:
            _fail_with_code(chat_id, f"{service_name}处理", "GW-E3002",
                            "处理完成但没有生成结果文件",
                            detail=f"success=True output_files=[] port=8003")
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
        logger.exception("[%s] 处理异常 trace=%s", chat_id, trace_id)
        _fail_from_exception(chat_id, f"{service_name}处理", e,
                             extra=f"files={[f['file_name'] for f in files]} port=8003")

def _handle_dealer_file(chat_id: str, message_id: str, file_key: str, file_name: str, local_path: str, service_name: str):
    """处理 销售部业绩核对群文件：支持批量上传或分批一个个上传"""
    global _dealer_sales_queues, _dealer_sales_timers, _dealer_sales_final_timers

    ftype = _detect_may_sales_type(file_name, local_path)
    if ftype == "template":
        # 自己传的核对表只作参照：脚本用的是固定模板（Sheet1/Sheet1 (2) 版式），
        # 传上来的这张版式不同，强行当模板会直接报错，所以静默跳过、不占三类名额
        _send_text(chat_id, f"\u2139\ufe0f 「{file_name}」当作核对参照，不参与计算；"
                            f"还差的三类文件继续传即可。")
        return
    if ftype == "unknown":
        _send_text(chat_id, f"\u26a0\ufe0f 无法识别文件「{file_name}」（已按内容和文件名一起判断），"
                            f"请确认它是综合查询 / 联思财务统计 / 奢匠下单统计之一。")
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
    has_zhcx = any(_detect_may_sales_type(f["file_name"], f.get("file_path", "")) == "zhcx" for f in queue)
    has_liansi = any(_detect_may_sales_type(f["file_name"], f.get("file_path", "")) == "liansi" for f in queue)
    has_shejiang = any(_detect_may_sales_type(f["file_name"], f.get("file_path", "")) == "shejiang" for f in queue)

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
    errors = []  # list[GatewayError]，每个都带错误码/trace
    trace_id = _get_trace()  # 可能来自 _process_file，沿用同一条追溯链
    logger.info("[%s] 经销商数据报表开始处理 trace=%s，文件：%s",
                chat_id, trace_id, [f["file_name"] for f in process_files])
    try:
        month_text = _format_months(target_months)
        result = _call_dealer_report_service(
            process_files,
            title=f"{month_text}经销商数据",
            output_filename=f"{month_text}经销商数据.xlsx",
        )
        logger.info("[%s] 处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            raw_error = str(result.get("error") or "服务未返回原因").strip()
            errors.append(GatewayError("GW-E3001", f"服务返回处理失败：{raw_error[:160]}",
                                       detail=f"success=False raw_error={raw_error!r}"))
        else:
            output_files = result.get("output_files", [])
            if not output_files:
                errors.append(GatewayError("GW-E3002", "处理完成但没有生成结果文件",
                                           detail="success=True output_files=[] port=8008"))

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
                    logger.exception("[%s] 发送文件失败 trace=%s file=%s", chat_id, trace_id, filename)
                    errors.append(GatewayError("GW-E4002", f"结果文件上传/发送失败（{filename}）",
                                               detail=_exception_detail(e), cause=e))
                time.sleep(0.5)
    except Exception as e:
        logger.exception("[%s] 处理经销商数据报表异常 trace=%s", chat_id, trace_id)
        errors.append(_to_gateway_error(e))
    finally:
        _dealer_report_queues.pop(chat_id, None)
        for f in queue:
            try:
                Path(f["file_path"]).unlink(missing_ok=True)
            except Exception:
                pass

    if errors:
        action = "酷家乐月度经销商数据处理"
        extra = f"files={[f['file_name'] for f in process_files]} port=8008"
        for err in errors:
            _log_failure(chat_id, action, err, extra)
        lines = [f"❌ {action}失败"]
        for err in errors:
            lines.append(f"• [{err.label}] {err.reason}")
        hint = next((err.hint for err in errors if err.hint), "")
        if hint:
            lines.append(f"处理建议：{hint}")
        try:
            _send_text(chat_id, "\n".join(lines))
        except Exception:
            logger.exception("[%s] 失败提示发送失败 trace=%s", chat_id, trace_id)
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
    trace_id = _begin_trace()
    tmpdir = tempfile.mkdtemp()
    try:
        logger.info("[%s] 下载文件: %s（trace=%s，端口 %d，message_id=%s）",
                    chat_id, file_name, trace_id, port, message_id)
        try:
            local_path = _download_feishu_file(message_id, file_key, tmpdir)
        except Exception as e:
            logger.exception("[%s] 下载文件失败 trace=%s", chat_id, trace_id)
            _fail_from_exception(chat_id, f"{service_name}文件下载", e, code="GW-E4001",
                                 extra=f"message_id={message_id} file_key={file_key}")
            return
        logger.info("[%s] 下载完成: %s", chat_id, local_path)
        _archive_file_copy(local_path, file_name, chat_id)

        # 经销商群：双文件配对逻辑
        if port == 8003:
            _handle_dealer_file(chat_id, message_id, file_key, file_name, local_path, service_name)
            return

        # 经销商数据报表群：双文件配对逻辑
        if port == 8008:
            _handle_dealer_report_file(chat_id, message_id, file_key, file_name, local_path, service_name)
            return
        order_date = f"{datetime.now().year}.{datetime.now().month}.{datetime.now().day}"
        result = _with_connect_retry(chat_id, service_name, _call_local_service,
                                     port, local_path, file_name, order_date)
        logger.info("[%s] 处理结果: %s", chat_id, result.get("success"))

        if not result.get("success"):
            raw_error = str(result.get("error") or "服务未返回原因").strip()
            _fail_with_code(chat_id, f"{service_name}处理", "GW-E3001",
                            f"服务返回处理失败：{raw_error[:160]}",
                            detail=f"success=False raw_error={raw_error!r}",
                            extra=f"file={file_name} port={port}")
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
            _fail_with_code(chat_id, f"{service_name}处理", "GW-E3002",
                            "处理完成但没有生成结果文件",
                            detail=f"success=True output_content=None output_files=[] file={file_name}")
            return

        count = sent_count
        msg = f"✅ {service_name}处理完成，共 {count} 个文件，请检查。"
        warnings = _collect_result_warnings(result)
        if warnings:
            msg += "\n\n" + "\n".join(warnings)
        # 压缩包和对话卡放一起：完成文案合并进数量卡片， zip 与卡片相邻
        if not _send_quantity_card(chat_id, result, port, service_name, note=msg):
            _send_text(chat_id, msg)

    except Exception as e:
        logger.exception("[%s] 处理异常 trace=%s", chat_id, trace_id)
        _fail_from_exception(chat_id, f"{service_name}处理", e,
                             extra=f"file={file_name} port={port} message_id={message_id}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _route_file_message(data, chat_id: str, message_id: str, file_key: str, file_name: str) -> None:
    """按群路由分发单个文件（多文件合并卡片会逐文件调用）"""
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

    if port in (8001, 8002, 8006):
        # 五金汇总 / 报价料单 / 下车间单：批量收集逻辑
        logger.info("[%s] 收到文件: %s → 路由到 %s (%d) 批量收集", chat_id, file_name, service_name, port)
        uploader_name = _resolve_uploader_name(data, chat_id) if port == 8006 else ""

        # 下载文件到持久化目录
        tmpdir = tempfile.mkdtemp()
        local_path = _download_feishu_file(message_id, file_key, tmpdir)
        _archive_file_copy(local_path, file_name, chat_id)
        if port == 8006:
            # 下车间单：走持久化文件队列（一次一份、重启续跑、失败重试）
            try:
                _, depth = _queue_enqueue(chat_id, local_path, file_name,
                                          message_id, file_key, uploader_name)
            finally:
                shutil.rmtree(tmpdir, ignore_errors=True)
            _queue_ensure_workers(chat_id, port, service_name)
            # MSGTRIM_V1：并发之内（depth <= 消费者数）马上就会开工，不必再刷一
            # 条「已收到」，直接等「处理中」那一条就行；只有真要等才提示一次。
            if depth > QUEUE_CONCURRENCY:
                _send_text(chat_id, "\u23F3 排队中 · %s（前面还有 %d 份）"
                           % (_short_name(file_name), depth - 1))
            return

        save_dir = Path(tempfile.gettempdir()) / f"batch_pending_{chat_id}"
        save_dir.mkdir(exist_ok=True)
        _sweep_stale_batch_files(save_dir)
        save_path = _batch_file_path(save_dir, file_name)
        shutil.copy(local_path, save_path)
        shutil.rmtree(tmpdir, ignore_errors=True)

        # 加入批量队列
        global _batch_queues
        with _batch_lock:
            if chat_id not in _batch_queues:
                _batch_queues[chat_id] = []
            _batch_queues[chat_id].append({
                "file_path": str(save_path),
                "file_name": file_name,
                "message_id": message_id,
                "file_key": file_key,
                "uploader_name": uploader_name,
                "time": time.time(),
            })
            queue_len = len(_batch_queues[chat_id])
            queue_full = queue_len >= BATCH_MAX_FILES
            if queue_full:
                old = _batch_timers.get(chat_id)
                if old is not None:
                    old.cancel()
                    _batch_timers[chat_id] = None
            else:
                _start_batch_timer(chat_id, port, service_name)

        if queue_full:
            logger.info("[%s] 批量队列已满 (%d)，立即处理", chat_id, queue_len)
            threading.Thread(
                target=_process_batch,
                args=(chat_id, port, service_name),
                daemon=True,
            ).start()
        else:
            if port == 8006:
                _send_text(chat_id, f"✅ 已收到「{file_name}」，正在单独处理…")
            else:
                _send_text(chat_id, f"✅ 已收到第 {queue_len} 个文件「{file_name}」，{BATCH_COLLECTION_WINDOW}秒内继续上传的文件将一起批量处理。")
        return

    # 单文件处理
    logger.info("[%s] 收到文件: %s → 路由到 %s (%d)", chat_id, file_name, service_name, port)
    threading.Thread(
        target=_process_file,
        args=(chat_id, message_id, file_key, file_name, port, service_name),
        daemon=True,
    ).start()


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

        # PVC2联思群：@机器人 + 「处理」触发联思每日流程（2026-10-20 接入）
        if msg_type == "text" and chat_id == LIANSI_CHAT_ID:
            try:
                content_json = json.loads(content) if isinstance(content, str) else content
                raw_text = content_json.get("text", "")
            except Exception:
                raw_text = ""
            # 「新增」优先：当天流程跑完后订单通又新加了批次（例 10.24美心加急，排产表批次名
            # 不带「联思」每日流程看不见）⇒ 只补导这一批，不重跑整天
            if raw_text and "新增" in raw_text:
                threading.Thread(
                    target=_handle_liansi_process_command,
                    args=(chat_id, raw_text, _extract_sender_open_id(data), LIANSI_ADD_URL),
                    daemon=True,
                ).start()
            # 「重导」：退单改完 / 单掉了，只重导点名的单或批次（不读排产表、不重跑整天）
            elif raw_text and "重导" in raw_text:
                threading.Thread(
                    target=_handle_liansi_process_command,
                    args=(chat_id, raw_text, _extract_sender_open_id(data), LIANSI_RE_URL),
                    daemon=True,
                ).start()
            elif raw_text and ("处理" in raw_text or "状态" in raw_text or "进度" in raw_text):
                threading.Thread(
                    target=_handle_liansi_process_command,
                    args=(chat_id, raw_text, _extract_sender_open_id(data)),
                    daemon=True,
                ).start()
            return

        # 多文件合并卡片（一张卡片多个文件）：msg_type=post，文件在 content.files
        if msg_type == "post":
            try:
                content_json = json.loads(content) if isinstance(content, str) else content
            except Exception:
                logger.warning("解析多文件消息失败")
                return
            files = [f for f in (content_json.get("files") or []) if not f.get("is_folder")]
            if not files:
                # 事件 content 没带 files 时回源 API 取
                files = _fetch_message_files(message_id)
            if not files:
                return
            if len(files) > 1:
                logger.info("[%s] 收到多文件卡片，共 %d 个文件", chat_id, len(files))
            for f in files:
                _route_file_message(data, chat_id, message_id, f.get("file_key", ""), f.get("file_name", "unknown"))
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

        _route_file_message(data, chat_id, message_id, file_key, file_name)

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

    _queue_restore()

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
