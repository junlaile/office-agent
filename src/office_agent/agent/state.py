"""LangGraph 状态定义（ReAct agent 极简版）。

ReAct agent 的核心是 messages 列表（LLM 与工具的消息往返）。
外加少量控制字段：
    - doc_path: 会话文档路径（main.py 注入，供工具读取）
    - done: finish 工具触发后置 true，路由到 END
"""

from __future__ import annotations

from typing import Annotated, Any, Literal, TypedDict

from langgraph.graph.message import add_messages


class InteractionField(TypedDict, total=False):
    key: str
    label: str
    required: bool
    options: list[str]
    hint: str


class InteractionRequest(TypedDict, total=False):
    """与具体 UI 无关的用户交互请求。"""

    request_id: str
    tool_call_id: str
    tool_name: str
    kind: Literal["form", "question", "confirmation"]
    title: str
    description: str
    question: str
    fields: list[InteractionField]
    options: list[str]
    tool_args: dict[str, Any]


class InteractionResponse(TypedDict, total=False):
    request_id: str
    accepted: bool
    answers: dict[str, str]
    value: str


class ToolExecutionRecord(TypedDict, total=False):
    tool_call_id: str
    tool_name: str
    status: Literal["completed", "cancelled", "failed"]
    result: Any


class ToolResult(TypedDict, total=False):
    """推荐给新增工具使用的统一结果信封。"""

    ok: bool
    code: str
    message: str
    data: Any
    retryable: bool


class AgentState(TypedDict, total=False):
    """ReAct agent 状态。total=False 让节点可只更新部分字段。"""

    # 消息往返：LLM 输出(AIMessage) + 工具结果(ToolMessage) 累积于此
    messages: Annotated[list, add_messages]
    # 会话文档路径（main.py 启动时注入，工具内部读取）
    doc_path: str
    # 是否完成（finish 工具置 true）
    done: bool
    # finish 时 LLM 给的总结
    summary: str
    # 当前待处理交互；由 prepare_interaction 节点写入，恢复后清空
    pending_interaction: InteractionRequest | None
    # 已完成调用记录；用于审计，并在 checkpoint 恢复时避免重复执行
    executed_calls: dict[str, ToolExecutionRecord]
