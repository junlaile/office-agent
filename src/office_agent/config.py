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
    max_iterations: int
    recursion_limit: int  # LangGraph 超级步上限（agent↔tools 往返算 2 步）

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
        # 优先级：环境变量 > pyproject.toml > default
        val = os.environ.get(key)
        if val:
            return val
        return pyconfig.get(key, default)

    output_dir_raw = env("OUTPUT_DIR", "./output")
    output_dir = Path(output_dir_raw)
    if not output_dir.is_absolute():
        output_dir = (_PROJECT_ROOT / output_dir).resolve()

    return Settings(
        llm_base_url=env("LLM_BASE_URL", "https://api.deepseek.com"),
        llm_api_key=env("LLM_API_KEY", ""),
        llm_model=env("LLM_MODEL", "deepseek-v4-flash"),
        llm_request_timeout=int(env("LLM_REQUEST_TIMEOUT", "120") or 120),
        llm_temperature=float(env("LLM_TEMPERATURE", "0.5") or 0.5),
        officecli_bin=env("OFFICECLI_BIN", ""),
        officecli_timeout=int(env("OFFICECLI_TIMEOUT", "120") or 120),
        output_dir=output_dir,
        max_iterations=int(env("MAX_ITERATIONS", "2") or 2),
        recursion_limit=int(env("RECURSION_LIMIT", "100") or 100),
        log_level=env("LOG_LEVEL", "INFO"),
    )


# 模块级单例：import 一次即生效
settings = _load()


def setup_logging() -> None:
    """按 settings.log_level 配置 root logger。

    幂等：重复调用不会叠加 handler（basicConfig 已自带此保证）。
    任何 `import office_agent.*` 都会经过本模块，因此自动配好日志。
    """
    level_name = settings.log_level.strip().upper()
    level = getattr(logging, level_name, logging.INFO)
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


# 模块加载即配好日志，使用方零配置
setup_logging()


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
