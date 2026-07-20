# factory-common

robot-services 微服务公共库，消除各服务重复代码。

## 模块

| 模块 | 说明 |
| --- | --- |
| `factory_common.logging_utils` | `setup_logging(service_name)`：统一日志配置（LOG_DIR 环境变量、文件+控制台、utf-8） |
| `factory_common.bom_utils` | BOM 公共工具：`detect_file_type` / `xlsx_to_xls` / `get_clean_filename` / `suppress_output` / `detect_unusual_sheets` |
| `factory_common.http_skeleton` | `run_process_server(...)`：http.server 版 `POST /process` + `GET /health` 骨架（单文件/批量、base64 编解码、临时目录、输出编码） |

## 用法

```python
import sys
from pathlib import Path

WORK_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(WORK_DIR.parent.parent / "factory-common"))  # 本地运行
# Docker 中通过 ENV PYTHONPATH=/app 提供

from factory_common.logging_utils import setup_logging
from factory_common.http_skeleton import run_process_server

logger = setup_logging("my-service")

def process(input_path, output_dir, req):
    ...  # 业务处理，返回 {"success": True, "files": [...]} 

if __name__ == "__main__":
    run_process_server(
        service_name="my-service",
        default_port=800X,
        process_fn=process,
        logger=logger,
        max_workers_env="MYSERVICE_MAX_WORKERS",   # 批量并发数环境变量
        encode_workers_env="MYSERVICE_ENCODE_MAX_WORKERS",  # 输出编码并发数环境变量
        default_output_base=WORK_DIR.parent / "data" / "output",
    )
```

## run_process_server 可选扩展参数

以下参数全部有默认值，默认时行为与原版完全一致（hardware-summary / csv-board 无需改动）：

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `path_aliases` | `()` | 除 `/process` 外额外接受的 POST 路径，如 `("/convert",)`。404 文案自动变为 `未知接口，请用 POST /process 或 POST /convert` |
| `not_found_error` | `None` | 完全自定义 POST 404 文案（覆盖 path_aliases 自动生成值） |
| `missing_file_error` | `"文件不存在"` | 无 input_path 且无 file_content（或文件不存在）时的错误文案，可传 `"缺少 file_content"` |
| `get_paths` | `None` | GET 路径白名单，如 `("/", "/health")`；`None` = 任何 GET 返回 ok（现状）。白名单外返回 404 |
| `get_not_found_error` | `"未知接口"` | GET 白名单外路径的 404 文案 |
| `batch_error_aggregate` | `False` | `True` 时批量响应有失败项则附加顶层 `"error"`：`"; "` 拼接 `"[<filename>] <error>"`（按原始 idx 顺序） |
| `batch_aggregate_hook` | `None` | 批量响应发送前回调 `hook(response_dict, results, files)`，可直接修改 response_dict（如加 quantity_total） |
| `startup_hook` | `None` | serve_forever 前调用一次（如 LibreOffice 预热）；异常只记日志不阻断启动 |
| `multi_file_fn` | `None` | 提供后 `{files:[...]}` 整体作为一个任务调用 `multi_file_fn(req, work_dir) -> dict`（work_dir 临时目录由骨架创建/清理），返回 dict 直接作为响应（骨架补 cost_seconds）。单文件请求仍走 process_fn |

另外：单文件模式下，若 `process_fn` 返回的 dict 已含 `output_files` 键（非 None），骨架跳过自动编码直接使用（服务自行打 zip / 自定义编码时用）。

### 扩展示例（quote-maker 风格）

```python
run_process_server(
    service_name="quote-maker",
    default_port=8007,
    process_fn=process,            # 成功时自行返回 {"success": True, "output_files": {...}}
    logger=logger,
    max_workers_env="QUOTE_MAX_WORKERS",
    default_max_workers=1,
    default_output_base=WORK_DIR.parent / "data" / "output",
    path_aliases=("/convert",),
    missing_file_error="缺少 file_content",
    get_paths=("/", "/health"),
)
```

## Docker 接入

公共库在仓库根目录、服务构建上下文之外，用 compose `additional_contexts` 引入：

```yaml
build:
  context: ../..
  dockerfile: deploy/docker/Dockerfile
  additional_contexts:
    factory-common: ../../..   # 仓库根目录
```

```dockerfile
COPY --from=factory-common factory-common/factory_common /app/factory_common
ENV PYTHONPATH=/app
```

compose 挂载（与 src 一致，宿主机改代码即生效；注意相对 compose 文件需三级 `../../../`）：

```yaml
- ../../../factory-common/factory_common:/app/factory_common:ro
```
