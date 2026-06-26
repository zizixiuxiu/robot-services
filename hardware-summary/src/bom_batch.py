"""
BOM 批量处理共享模块
支持：文件收集窗口、并发处理、自动漏处理检测、状态查询
BOM文件不需要配对，每个文件独立处理
"""

import asyncio
import json
import os
import time
import uuid
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Dict, List, Optional, Callable, Any
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum

from mcp.types import TextContent


class TaskStatus(Enum):
    PENDING = "pending"
    COLLECTING = "collecting"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


@dataclass
class BomTask:
    """BOM处理任务"""
    task_id: str
    input_path: str
    status: TaskStatus = TaskStatus.PENDING
    created_at: float = field(default_factory=time.time)
    completed_at: Optional[float] = None
    result: Optional[Dict] = None
    error: Optional[str] = None
    chat_id: Optional[str] = None
    message_id: Optional[str] = None


class BomBatchManager:
    """BOM批量任务管理器"""

    def __init__(self, work_dir: Path, process_fn: Callable[[str, str], Dict],
                 max_concurrent: int = 3):
        """
        Args:
            work_dir: 工作目录
            process_fn: 单文件处理函数 (input_path, output_dir) -> Dict
            max_concurrent: 最大并发数
        """
        self.work_dir = work_dir
        self.process_fn = process_fn
        self.tasks: Dict[str, BomTask] = {}
        self.lock = threading.Lock()
        self.executor = ThreadPoolExecutor(max_workers=max_concurrent)
        self._batch_timer: Optional[asyncio.Task] = None
        self._batch_start_time: Optional[float] = None
        self._auto_check_timer: Optional[asyncio.Task] = None

        # 配置
        self.collection_window = 10   # 收集窗口（秒）
        self.auto_check_interval = 30  # 自动检查间隔
        self.task_timeout = 300        # 任务超时（秒）
        self.max_task_age = 3600       # 最大任务保留时间（秒）

    def add_file(self, input_path: str, chat_id: str = None,
                 message_id: str = None) -> BomTask:
        """添加文件到批量队列"""
        task_id = str(uuid.uuid4())[:8]
        task = BomTask(
            task_id=task_id,
            input_path=input_path,
            chat_id=chat_id,
            message_id=message_id,
            status=TaskStatus.COLLECTING
        )
        with self.lock:
            self.tasks[task_id] = task
        return task

    def get_ready_tasks(self) -> List[BomTask]:
        """获取可处理的任务（COLLECTING状态）"""
        with self.lock:
            return [t for t in self.tasks.values()
                    if t.status == TaskStatus.COLLECTING]

    def get_unprocessed_tasks(self) -> List[BomTask]:
        """获取未处理的任务"""
        with self.lock:
            return [t for t in self.tasks.values()
                    if t.status in (TaskStatus.COLLECTING, TaskStatus.PENDING)]

    def get_orphaned_tasks(self) -> List[BomTask]:
        """获取超时未处理的任务"""
        now = time.time()
        with self.lock:
            return [t for t in self.tasks.values()
                    if t.status == TaskStatus.COLLECTING
                    and (now - t.created_at) > self.task_timeout]

    def get_task(self, task_id: str) -> Optional[BomTask]:
        return self.tasks.get(task_id)

    def _process_sync(self, task: BomTask) -> Dict:
        """同步处理单个任务（在线程池中运行）"""
        try:
            task.status = TaskStatus.PROCESSING

            # 创建输出目录
            base = Path(task.input_path).stem
            timestamp = str(int(time.time()))
            output_dir = str(self.work_dir / f"output_{base}_{timestamp}")
            os.makedirs(output_dir, exist_ok=True)

            # 调用处理函数
            result = self.process_fn(task.input_path, output_dir)

            task.status = TaskStatus.COMPLETED
            task.completed_at = time.time()
            task.result = result
            return result

        except Exception as e:
            import traceback
            task.status = TaskStatus.FAILED
            task.error = str(e)
            return {
                "success": False,
                "error": str(e),
                "trace": traceback.format_exc(),
                "task_id": task.task_id
            }

    async def process_task(self, task: BomTask) -> Dict:
        """异步处理单个任务"""
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(self.executor, self._process_sync, task)
        return result

    async def process_all_ready(self) -> List[Dict]:
        """并发处理所有就绪任务"""
        ready = self.get_ready_tasks()
        if not ready:
            return []

        loop = asyncio.get_event_loop()
        futures = [
            loop.run_in_executor(self.executor, self._process_sync, task)
            for task in ready
        ]
        results = await asyncio.gather(*futures, return_exceptions=True)

        processed = []
        for task, result in zip(ready, results):
            if isinstance(result, Exception):
                task.status = TaskStatus.FAILED
                task.error = str(result)
                processed.append({
                    "success": False, "error": str(result),
                    "task_id": task.task_id
                })
            else:
                processed.append(result)

        return processed

    async def start_batch_collection(self):
        """启动批量收集窗口"""
        if self._batch_timer and not self._batch_timer.done():
            return  # 已有窗口在运行

        self._batch_start_time = time.time()

        async def timer():
            await asyncio.sleep(self.collection_window)
            await self.process_all_ready()

        self._batch_timer = asyncio.create_task(timer())

    def cleanup_old_tasks(self):
        """清理过期任务"""
        now = time.time()
        with self.lock:
            old = [tid for tid, t in self.tasks.items()
                   if now - t.created_at > self.max_task_age]
            for tid in old:
                del self.tasks[tid]


# ============== MCP 工具定义 ==============

BATCH_TOOLS = [
    {
        "name": "upload_file",
        "description": "上传文件（支持批量模式）。系统会自动收集窗口期内的所有文件并批量处理",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "本地文件绝对路径"
                },
                "batch_mode": {
                    "type": "boolean",
                    "description": "是否启用批量模式（默认true）",
                    "default": True
                },
                "chat_id": {
                    "type": "string",
                    "description": "来源群ID（自动填充）"
                },
                "message_id": {
                    "type": "string",
                    "description": "来源消息ID（自动填充）"
                }
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "process_batch",
        "description": "立即处理所有已收集的文件",
        "inputSchema": {"type": "object", "properties": {}}
    },
    {
        "name": "check_task_status",
        "description": "检查任务状态",
        "inputSchema": {
            "type": "object",
            "properties": {
                "task_id": {
                    "type": "string",
                    "description": "任务ID（可选，不传则返回所有任务）"
                }
            }
        }
    },
    {
        "name": "check_unprocessed_files",
        "description": "检查是否有未处理的文件（自动漏处理检测）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "chat_id": {
                    "type": "string",
                    "description": "指定群ID（可选，不传则检查所有群）"
                }
            }
        }
    },
    {
        "name": "get_output_files",
        "description": "获取最新的输出文件列表",
        "inputSchema": {"type": "object", "properties": {}}
    },
]


def make_batch_tools(prefix: str, description: str) -> list:
    """为指定服务生成批量工具定义"""
    tools = []
    for t in BATCH_TOOLS:
        tool = dict(t)
        if tool["name"] == "upload_file":
            tool["name"] = f"{prefix}_upload_file"
            tool["description"] = description
        else:
            tool["name"] = f"{prefix}_{tool['name']}"
        tools.append(tool)
    return tools


def _json_response(data: Any) -> list:
    return [TextContent(type="text", text=json.dumps(data, ensure_ascii=False, indent=2))]


async def handle_upload_file(
    manager: BomBatchManager,
    arguments: dict,
    validate_fn: Callable[[str], tuple[bool, str]],
) -> list:
    """处理文件上传（通用）"""
    file_path = arguments.get("file_path", "")
    batch_mode = arguments.get("batch_mode", True)
    chat_id = arguments.get("chat_id", "default")
    message_id = arguments.get("message_id", "")

    if not file_path or not os.path.exists(file_path):
        return _json_response({"success": False, "error": "文件不存在"})

    # 文件类型校验
    ok, msg = validate_fn(file_path)
    if not ok:
        return _json_response({"success": False, "error": msg})

    # 添加到批量队列
    task = manager.add_file(file_path, chat_id, message_id)

    if batch_mode:
        await manager.start_batch_collection()
        return _json_response({
            "success": True,
            "status": "collecting",
            "task_id": task.task_id,
            "hint": f"文件已加入批量队列，{manager.collection_window}秒内上传的文件将一起处理"
        })
    else:
        # 单文件立即处理
        result = await manager.process_task(task)
        return _json_response(result)


async def handle_process_batch(manager: BomBatchManager) -> list:
    """手动触发批量处理"""
    # 取消收集定时器
    if manager._batch_timer and not manager._batch_timer.done():
        manager._batch_timer.cancel()

    results = await manager.process_all_ready()

    if not results:
        return _json_response({"success": True, "message": "没有待处理的任务"})

    success_count = sum(1 for r in results if r.get("success"))
    fail_count = len(results) - success_count

    # 汇总所有输出文件
    all_files = []
    for r in results:
        if r.get("success") and r.get("output_dir"):
            from pathlib import Path as P
            out_dir = P(r["output_dir"])
            for f in sorted(out_dir.rglob("*.xlsx")) + sorted(out_dir.rglob("*.xls")):
                all_files.append(str(f))

    return _json_response({
        "success": True,
        "total": len(results),
        "success_count": success_count,
        "fail_count": fail_count,
        "files": all_files,
        "results": results
    })


async def handle_check_task_status(manager: BomBatchManager, arguments: dict) -> list:
    """检查任务状态"""
    task_id = arguments.get("task_id")

    if task_id:
        task = manager.get_task(task_id)
        if not task:
            return _json_response({"success": False, "error": f"任务 {task_id} 不存在"})

        info = {
            "task_id": task.task_id,
            "status": task.status.value,
            "input_path": task.input_path,
            "created_at": datetime.fromtimestamp(task.created_at).strftime("%Y-%m-%d %H:%M:%S"),
        }
        if task.completed_at:
            info["completed_at"] = datetime.fromtimestamp(task.completed_at).strftime("%Y-%m-%d %H:%M:%S")
        if task.result:
            info["result"] = task.result
        if task.error:
            info["error"] = task.error

        return _json_response(info)
    else:
        # 返回所有任务
        all_tasks = []
        for task in manager.tasks.values():
            t = {
                "task_id": task.task_id,
                "status": task.status.value,
                "input": os.path.basename(task.input_path),
            }
            if task.error:
                t["error"] = task.error
            all_tasks.append(t)

        return _json_response({
            "total_tasks": len(all_tasks),
            "tasks": all_tasks
        })


async def handle_check_unprocessed_files(manager: BomBatchManager, arguments: dict) -> list:
    """检查未处理的文件"""
    chat_id = arguments.get("chat_id")

    unprocessed = manager.get_unprocessed_tasks()
    if chat_id:
        unprocessed = [t for t in unprocessed if t.chat_id == chat_id]

    if not unprocessed:
        return _json_response({
            "success": True,
            "total_unprocessed": 0,
            "hint": "所有文件已处理完毕"
        })

    # 按群分组
    by_chat: Dict[str, List] = {}
    for t in unprocessed:
        cid = t.chat_id or "unknown"
        by_chat.setdefault(cid, []).append({
            "task_id": t.task_id,
            "file": os.path.basename(t.input_path),
            "age_seconds": int(time.time() - t.created_at)
        })

    return _json_response({
        "success": True,
        "total_unprocessed": len(unprocessed),
        "by_chat": by_chat,
        "hint": f"发现 {len(unprocessed)} 个未处理文件，调用 process_batch 立即处理"
    })


async def handle_get_output_files(manager: BomBatchManager) -> list:
    """获取最新的输出文件列表"""
    output_files = []
    for task in manager.tasks.values():
        if task.status == TaskStatus.COMPLETED and task.result:
            if task.result.get("output_dir"):
                from pathlib import Path as P
                out_dir = P(task.result["output_dir"])
                for f in sorted(out_dir.rglob("*.xlsx")) + sorted(out_dir.rglob("*.xls")):
                    output_files.append({
                        "task_id": task.task_id,
                        "file": f.name,
                        "path": str(f),
                    })

    return _json_response({
        "success": True,
        "total": len(output_files),
        "files": output_files
    })
