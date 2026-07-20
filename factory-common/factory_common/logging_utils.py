"""公共日志配置：复刻各微服务现有日志块行为"""
import os
import sys
import inspect
import logging
from pathlib import Path


def setup_logging(service_name: str) -> logging.Logger:
    """创建服务 logger。

    行为与各服务原日志块一致：
    - 日志目录取 LOG_DIR 环境变量；缺省为 <调用方文件所在目录的上一级>/logs
      （调用方通常在 <service>/src/ 下，即 <service>/logs）
    - 日志文件名 <service_name 中 - 换成 _>_http.log
    - FileHandler(utf-8) + StreamHandler(stdout)
    - 格式 "%(asctime)s [%(levelname)s] %(message)s"
    """
    caller_file = Path(inspect.stack()[1].filename).resolve()
    default_log_dir = caller_file.parent.parent / "logs"

    log_dir = Path(os.getenv("LOG_DIR", str(default_log_dir)))
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{service_name.replace('-', '_')}_http.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger(service_name)
