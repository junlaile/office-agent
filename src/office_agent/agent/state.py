"""LangGraph 状态定义（ReAct agent 极简版）。

ReAct agent 的核心是 messages 列表（LLM 与工具的消息往返）。
外加少量控制字段：
    - doc_path: 会话文档路径（main.py 注入，供工具读取）
    - done: finish 工具触发后置 true，路由到 END
"""

from __future__ import annotations

import operator
from typing import Annotated, TypedDict

from langgraph.graph.message import add_messages


class AgentState(TypedDict, total=False):
    """ReAct agent 状态。total=False 让节点可只更新部分字段。"""

    # 消息往返：LLM 输出(AIMessage) + 工具结果(ToolMessage) 累积于此
    messages: Annotated[list, add_messages]
    # 会话文档路径（main.py 启动时注入，工具内部读取）
    doc_path: str
    # agent 节点已执行的轮数（每次 LLM 调用 +1）。
    # 软收尾用它对比 recursion_limit（一轮 agent+tools ≈ 2 个超级步），
    # 不再用 len(messages) 估算——并行工具调用一轮就加十几条消息，会误触发。
    steps: Annotated[int, operator.add]
    # 是否完成（finish 工具置 true）
    done: bool
    # finish 时 LLM 给的总结
    summary: str
