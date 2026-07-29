"""LLM 工厂：构造 OpenAI 兼容的 ChatOpenAI。

复用 settings.llm_* 配置；支持 GLM/DeepSeek/OpenAI/通义等任何兼容接口。
"""

from __future__ import annotations

from langchain_openai import ChatOpenAI

from office_agent.config import settings


def get_llm(*, streaming: bool = False) -> ChatOpenAI:
    """构造一个 ChatOpenAI 实例。

    Args:
        streaming: 是否启用流式输出（SSE 场景传 True）。
    """
    from pydantic import SecretStr

    # api_key 用 SecretStr 包装，避免日志/repr 意外泄露明文密钥
    return ChatOpenAI(
        base_url=settings.llm_base_url,
        api_key=SecretStr(settings.llm_api_key),
        model=settings.llm_model,
        temperature=settings.llm_temperature,
        timeout=settings.llm_request_timeout,
        streaming=streaming,
    )
