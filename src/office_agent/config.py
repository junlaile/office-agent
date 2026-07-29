"""集中加载环境配置。

配置优先级（高 → 低）:
    1. 真实环境变量
    2. .env 文件（python-dotenv 加载到环境变量）
    3. pyproject.toml 的 [tool.office-agent] 段（写死的工程默认值）

这样工程开箱即用（默认值在 pyproject.toml），又能用环境变量临时覆盖。
"""

from __future__ import annotations

import logging
import os
import tomllib
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_PYPROJECT = _PROJECT_ROOT / "pyproject.toml"


def _load_pyconfig() -> dict[str, str]:
    """读取 pyproject.toml 的 [tool.office-agent] 段。"""
    if not _PYPROJECT.exists():
        return {}
    try:
        with _PYPROJECT.open("rb") as f:
            data = tomllib.load(f)
    except Exception:  # noqa: BLE001
        return {}
    section = data.get("tool", {}).get("office-agent", {})
    # 统一成字符串
    return {k: str(v) for k, v in section.items()}


@dataclass(frozen=True)
class Settings:
    # LLM
    llm_base_url: str
    llm_api_key: str
    llm_model: str
    llm_request_timeout: int
    llm_temperature: float

    # OfficeCLI
    officecli_bin: str  # 可空：空则 resolve_bin 解析
    officecli_timeout: int

    # 输出
    output_dir: Path

    # Agent
    recursion_limit: int  # LangGraph 超级步上限（agent↔tools 往返算 2 步）

    # Session store
    session_backend: str  # memory/mysql
    session_ttl_seconds: int
    mysql_dsn: str
    mysql_host: str
    mysql_port: int
    mysql_user: str
    mysql_password: str
    mysql_database: str
    mysql_connect_timeout: int
    mysql_read_timeout: int
    mysql_write_timeout: int

    # Transport
    transport_mode: str  # inproc/rabbitmq_stomp
    rabbitmq_host: str
    rabbitmq_port: int
    rabbitmq_login: str
    rabbitmq_passcode: str
    rabbitmq_vhost: str
    stomp_inbound_destination: str
    stomp_outbound_destination: str
    stomp_dlq_destination: str
    stomp_exchange: str
    stomp_routing_key: str
    stomp_heartbeat_ms: int
    stomp_max_retries: int
    stomp_retry_delay_ms: int
    stomp_use_memory_broker: bool

    # 日志
    log_level: str  # DEBUG/INFO/WARNING/ERROR

    @property
    def project_root(self) -> Path:
        return _PROJECT_ROOT


def _load() -> Settings:
    # .env 在工程根目录；load_dotenv 不会覆盖已存在的环境变量
    load_dotenv(_PROJECT_ROOT / ".env")
    pyconfig = _load_pyconfig()

    def env(key: str, default: str = "") -> str:
        # 优先级：环境变量 > pyproject.toml > default。
        # 显式设成空串的环境变量也算"已设置"（可用于清空某项配置）。
        val = os.environ.get(key)
        if val is not None:
            return val
        return pyconfig.get(key, default)

    def int_env(key: str, default: int) -> int:
        raw = env(key, str(default)).strip()
        if not raw:
            return default
        try:
            return int(raw)
        except ValueError:
            raise SystemExit(f"配置 {key} 的值不是整数: {raw!r}（检查环境变量 / .env / pyproject.toml）") from None

    def float_env(key: str, default: float) -> float:
        raw = env(key, str(default)).strip()
        if not raw:
            return default
        try:
            return float(raw)
        except ValueError:
            raise SystemExit(f"配置 {key} 的值不是数字: {raw!r}（检查环境变量 / .env / pyproject.toml）") from None

    output_dir_raw = env("OUTPUT_DIR", "./output") or "./output"
    output_dir = Path(output_dir_raw)
    if not output_dir.is_absolute():
        output_dir = (_PROJECT_ROOT / output_dir).resolve()

    return Settings(
        llm_base_url=env("LLM_BASE_URL", "https://api.deepseek.com"),
        llm_api_key=env("LLM_API_KEY", ""),
        llm_model=env("LLM_MODEL", "deepseek-v4-flash"),
        llm_request_timeout=int_env("LLM_REQUEST_TIMEOUT", 120),
        llm_temperature=float_env("LLM_TEMPERATURE", 0.5),
        officecli_bin=env("OFFICECLI_BIN", ""),
        officecli_timeout=int_env("OFFICECLI_TIMEOUT", 120),
        output_dir=output_dir,
        recursion_limit=int_env("RECURSION_LIMIT", 100),
        session_backend=env("SESSION_BACKEND", "memory"),
        session_ttl_seconds=int_env("SESSION_TTL_SECONDS", 86400),
        mysql_dsn=env("MYSQL_DSN", ""),
        mysql_host=env("MYSQL_HOST", "127.0.0.1"),
        mysql_port=int_env("MYSQL_PORT", 3306),
        mysql_user=env("MYSQL_USER", ""),
        mysql_password=env("MYSQL_PASSWORD", ""),
        mysql_database=env("MYSQL_DATABASE", "office_agent"),
        mysql_connect_timeout=int_env("MYSQL_CONNECT_TIMEOUT", 5),
        mysql_read_timeout=int_env("MYSQL_READ_TIMEOUT", 15),
        mysql_write_timeout=int_env("MYSQL_WRITE_TIMEOUT", 15),
        transport_mode=env("TRANSPORT_MODE", "inproc"),
        rabbitmq_host=env("RABBITMQ_HOST", "127.0.0.1"),
        rabbitmq_port=int_env("RABBITMQ_PORT", 61613),
        rabbitmq_login=env("RABBITMQ_LOGIN", "guest"),
        rabbitmq_passcode=env("RABBITMQ_PASSCODE", "guest"),
        rabbitmq_vhost=env("RABBITMQ_VHOST", "/"),
        stomp_inbound_destination=env(
            "STOMP_INBOUND_DESTINATION", "/queue/office-agent.inbound"
        ),
        stomp_outbound_destination=env(
            "STOMP_OUTBOUND_DESTINATION", "/queue/office-agent.outbound"
        ),
        stomp_dlq_destination=env(
            "STOMP_DLQ_DESTINATION", "/queue/office-agent.dlq"
        ),
        stomp_exchange=env("STOMP_EXCHANGE", "office-agent"),
        stomp_routing_key=env("STOMP_ROUTING_KEY", "session"),
        stomp_heartbeat_ms=int_env("STOMP_HEARTBEAT_MS", 10000),
        stomp_max_retries=int_env("STOMP_MAX_RETRIES", 3),
        stomp_retry_delay_ms=int_env("STOMP_RETRY_DELAY_MS", 200),
        stomp_use_memory_broker=env("STOMP_USE_MEMORY_BROKER", "").lower()
        in ("1", "true", "yes"),
        log_level=env("LOG_LEVEL", "INFO"),
    )


# 模块级单例：import 一次即生效
settings = _load()


def setup_logging() -> None:
    """按 settings.log_level 配置 root logger。

    幂等：重复调用不会叠加 handler（basicConfig 已自带此保证）。
    由 CLI 入口（cli.main.run）显式调用——不在 import 时自动执行，
    避免"import 本包就改掉宿主进程的全局日志配置"这种副作用。
    """
    level_name = settings.log_level.strip().upper()
    level = getattr(logging, level_name, logging.INFO)
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def assert_llm_ready() -> None:
    """启动前检查 LLM 配置是否齐全，给出友好提示。"""
    missing = []
    if not settings.llm_api_key:
        missing.append("LLM_API_KEY")
    if not settings.llm_base_url:
        missing.append("LLM_BASE_URL")
    if not settings.llm_model:
        missing.append("LLM_MODEL")
    if missing:
        keys = ", ".join(missing)
        raise SystemExit(
            f"缺少 LLM 配置: {keys}\n"
            f"请复制 .env.example 为 .env 并填写（位于 {_PROJECT_ROOT / '.env'}）。"
        )
