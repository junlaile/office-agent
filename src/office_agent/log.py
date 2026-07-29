"""统一日志管理模块。

全项目的日志配置与 logger 获取都收敛到这里:

    from office_agent.log import get_logger

    logger = get_logger(__name__)

设计要点:
    - ``setup_logging()`` 只由进程入口显式调用（CLI ``cli.main.run``、
      API ``api.app``、Worker ``api.worker``），不在 import 时自动执行，
      避免"import 本包就改掉宿主进程的全局日志配置"这种副作用。
    - 日志级别来自配置项 ``LOG_LEVEL``（环境变量 / .env / pyproject.toml），
      默认 INFO。
    - 可选文件输出:配置项 ``LOG_FILE`` 非空时，额外挂一个
      ``RotatingFileHandler``（单文件 10MB，保留 3 份备份），与控制台
      使用同一格式，便于线上排查。
    - 幂等:重复调用 ``setup_logging()`` 不会叠加 handler。
    - 统一压低第三方库（httpx/urllib3 等）的日志噪音到 WARNING。
"""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from office_agent.config import settings

# 全项目统一的日志格式:时间 [级别] 模块名: 消息
_LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_DATE_FORMAT = "%H:%M:%S"

# 文件轮转参数:单文件 10MB、保留 3 份备份，防止长期运行撑爆磁盘
_MAX_BYTES = 10 * 1024 * 1024
_BACKUP_COUNT = 3

# 这些第三方库在 DEBUG/INFO 级别会刷大量请求细节，统一压到 WARNING
_NOISY_LOGGERS = ("httpx", "httpcore", "urllib3", "openai")

# 幂等标记:setup_logging 已执行过则直接返回
_configured = False


def setup_logging(level: str | None = None) -> None:
    """按配置初始化 root logger（幂等，重复调用不叠加 handler）。

    Args:
        level: 显式指定级别（如 ``"DEBUG"``）；缺省时读 ``settings.log_level``。
    """
    global _configured
    if _configured:
        return
    _configured = True

    level_name = (level or settings.log_level).strip().upper()
    resolved = getattr(logging, level_name, logging.INFO)

    formatter = logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT)

    handlers: list[logging.Handler] = []

    # 控制台输出（stderr），始终开启
    console = logging.StreamHandler()
    console.setFormatter(formatter)
    handlers.append(console)

    # 可选文件输出:LOG_FILE 非空时启用轮转文件
    if settings.log_file:
        log_path = Path(settings.log_file)
        if not log_path.is_absolute():
            log_path = (settings.project_root / log_path).resolve()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=_MAX_BYTES,
            backupCount=_BACKUP_COUNT,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        handlers.append(file_handler)

    root = logging.getLogger()
    root.setLevel(resolved)
    for handler in handlers:
        root.addHandler(handler)

    # 压低第三方库噪音，避免 DEBUG 级别时刷屏
    for name in _NOISY_LOGGERS:
        logging.getLogger(name).setLevel(logging.WARNING)


def get_logger(name: str) -> logging.Logger:
    """统一的 logger 获取入口。

    各模块统一用 ``get_logger(__name__)``，保证日志名与模块层级一致
    （如 ``office_agent.api.app``），便于按前缀过滤/调级。
    """
    return logging.getLogger(name)
