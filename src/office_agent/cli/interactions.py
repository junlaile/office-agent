"""用户交互适配器接口。

Graph 只产生结构化 InteractionRequest；CLI/Web 等前端各自实现 collect，
并返回统一 InteractionResponse。这样新增交互工具不需要依赖终端实现。
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import Any, Protocol

from office_agent.agent.state import InteractionResponse


class InteractionAdapter(Protocol):
    def collect(self, request: Mapping[str, Any]) -> InteractionResponse: ...


@dataclass(slots=True)
class CLIInteractionAdapter:
    """复用 CLI 现有渲染函数的适配器。"""

    collect_form: Callable[[dict], dict[str, str]]
    collect_question: Callable[[dict], str]

    def collect(self, request: Mapping[str, Any]) -> InteractionResponse:
        payload = dict(request)
        request_id = str(payload.get("request_id", ""))
        kind = payload.get("kind")

        if kind == "confirmation":
            value = self.collect_question(payload)
            accepted = value.strip().lower() in {
                "确认",
                "是",
                "同意",
                "继续",
                "yes",
                "y",
                "ok",
                "true",
            }
            return InteractionResponse(
                request_id=request_id,
                accepted=accepted,
                value=value,
            )

        if payload.get("fields"):
            return InteractionResponse(
                request_id=request_id,
                answers=self.collect_form(payload),
            )

        return InteractionResponse(
            request_id=request_id,
            value=self.collect_question(payload),
        )
