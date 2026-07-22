"""ReAct agent 图装配（LangGraph 原生写法 + DeepSeek tool_calls）。

结构:
    START → agent → tools → (路由) → agent ↻ ... → END

    - agent 节点: LLM.bind_tools(ALL_TOOLS)，决定调哪个工具
    - tools 节点: 手写（不用 ToolNode），因为要处理
        · ask_user 的 interrupt 挂起
        · finish 的短路完成
    - 路由: tools 之后，若 finish 已调用 → END；
            否则回到 agent 继续（agent 没再调工具时自然到 END）。

防卡死: agent 节点在接近 recursion_limit 软阈值时会注入"尽快 finish"提示，
        引导 LLM 主动收尾，避免撞硬限制报 GraphRecursionError。

参考: https://api-docs.deepseek.com/zh-cn/guides/tool_calls
（DeepSeek 支持 OpenAI 标准 tools；用 bind_tools，不强制 tool_choice）
"""

from __future__ import annotations

import logging
from typing import Any

from langchain_core.messages import AIMessage, SystemMessage, ToolMessage
from langgraph.checkpoint.memory import MemorySaver
from langgraph.errors import GraphBubbleUp
from langgraph.graph import END, START, StateGraph

from .config import settings
from .llm import get_llm
from .prompts import build_system_prompt
from .state import AgentState
from .tools import ALL_TOOLS, TOOL_BY_NAME

logger = logging.getLogger(__name__)

# 软收尾阈值：当剩余步数 < 限额的此比例时，触发"尽快 finish"提醒。
# 0.7 意味着用掉 70% 预算后开始催促收尾，留 30% 余量完成 view_text+finish。
SOFT_FINISH_RATIO = 0.7


def _agent_node_factory(doc_path: str, doc_type: str | None = None):
    """构造 agent 节点。doc_path 用于系统提示词，doc_type 触发公文模式分支。"""

    llm_with_tools = get_llm().bind_tools(ALL_TOOLS)
    system_msg = SystemMessage(content=build_system_prompt(doc_path, doc_type))

    def agent_node(state: AgentState) -> dict[str, Any]:
        messages = state.get("messages", [])
        limit = settings.recursion_limit
        # 用 messages 数估算已用步数：每轮 agent+tools 约加 2 条消息。
        # 粗略但足够触发软收尾。
        used = len(messages)
        soft_threshold = int(limit * SOFT_FINISH_RATIO)

        prompt_messages: list = [system_msg, *messages]
        # 软收尾提醒：接近预算时催促 finish
        if used >= soft_threshold:
            remaining = max(0, limit - used)
            logger.debug("触发软收尾提醒: used=%d, remaining=%d", used, remaining)
            reminder = SystemMessage(
                content=(
                    f"【系统提醒】已用约 {used // 2} 轮工具调用，剩余预算约 {remaining // 2} 轮。"
                    f"请【立即停止添加新内容】，执行 view_text 自查（若还没查），"
                    f"然后调用 finish 宣告完成。不要再调用 add_* 工具。"
                )
            )
            prompt_messages.append(reminder)

        ai_msg = llm_with_tools.invoke(prompt_messages)
        return {"messages": [ai_msg]}

    return agent_node


def _tools_node(state: AgentState) -> dict[str, Any]:
    """执行 agent 要求的工具调用。

    特殊处理:
        - ask_user: 内部 interrupt 挂起，resume 后返回值即用户答案，
          包成 ToolMessage 回传。
        - finish: 标记 done，短路结束。
    其余 officecli 工具直接 .invoke 执行。
    """
    messages = state.get("messages", [])
    last_ai: AIMessage = messages[-1]
    tool_messages: list[ToolMessage] = []
    updates: dict[str, Any] = {"messages": tool_messages}
    tool_calls = last_ai.tool_calls or []

    # ask_user 必须独占执行：它内部 interrupt 会挂起整个节点，若同批还有
    # 其他工具，会破坏"tool_calls 必须紧跟对应 ToolMessage"的消息约束
    # （DeepSeek 报 400 insufficient tool messages）。
    # 处理：检测到同批含 ask_user 时，整批都不执行，给每个 tool_call 回一条
    # ToolMessage 提示 LLM "请单独调用 ask_user"，下一轮 LLM 会重新决策。
    # 这样绝不触发 interrupt，消息配对始终完整。
    if any(tc.get("name") == "ask_user" for tc in tool_calls) and len(tool_calls) > 1:
        for tc in tool_calls:
            if tc.get("name") == "ask_user":
                tool_messages.append(
                    ToolMessage(
                        content=(
                            "本批工具调用中同时包含了 ask_user 和其他工具。"
                            "ask_user 会暂停整个流程等待用户输入，必须【单独】调用。"
                            "请在下一次回复里【只】调用 ask_user，不要附带其他工具。"
                        ),
                        tool_call_id=tc.get("id", ""),
                    )
                )
            else:
                tool_messages.append(
                    ToolMessage(
                        content=(
                            f"该 {tc.get('name')} 调用已取消（本批同时调用了 ask_user）。"
                            f"请在 ask_user 完成后重新发起。"
                        ),
                        tool_call_id=tc.get("id", ""),
                    )
                )
        return updates

    for tc in tool_calls:
        name = tc.get("name")
        args = tc.get("args") or {}
        tc_id = tc.get("id", "")
        tool = TOOL_BY_NAME.get(name)

        if name == "finish":
            # 短路完成：把 summary 存入 state，并回传 ToolMessage
            summary = args.get("summary", "")
            updates["done"] = True
            updates["summary"] = summary
            tool_messages.append(
                ToolMessage(
                    content=f"已完成: {summary}",
                    tool_call_id=tc_id,
                )
            )
            continue

        if tool is None:
            tool_messages.append(
                ToolMessage(
                    content=f"错误: 未知工具 '{name}'",
                    tool_call_id=tc_id,
                )
            )
            continue

        # ask_user 会触发 interrupt（抛 GraphInterrupt/GraphBubbleUp）；
        # 这种 LangGraph 控制流信号【必须】向上抛出给运行时处理，绝不能吞掉，
        # 否则 agent 会误以为工具出错而放弃提问。
        try:
            result = tool.invoke(args)
        except GraphBubbleUp:
            # interrupt 等控制流：原样上抛，由 LangGraph 挂起/恢复
            raise
        except Exception as e:  # noqa: BLE001
            logger.exception("工具 %s 执行失败", name)
            result = f"工具执行出错({name}): {e}"
        tool_messages.append(
            ToolMessage(
                content=str(result),
                tool_call_id=tc_id,
            )
        )

    return updates


def _route_after_tools(state: AgentState) -> str:
    """tools 之后：finish 完成 → END；否则回 agent。"""
    if state.get("done"):
        return END
    return "agent"


def _route_after_agent(state: AgentState) -> str:
    """agent 之后：有 tool_calls → tools；无 → END（LLM 直接给出最终答复）。"""
    messages = state.get("messages", [])
    if not messages:
        return END
    last = messages[-1]
    if isinstance(last, AIMessage) and last.tool_calls:
        return "tools"
    return END


def build_graph(doc_path: str, doc_type: str | None = None):
    """构建并编译 ReAct agent 图。

    doc_path: 本会话的文档输出路径（注入到系统提示词和工具会话）。
    doc_type: 法定公文文种名（如 '通知'）。非空时启用公文模式——
        提示词走公文分支，指导 LLM 编辑模板正文而非从零拼接。
        None 表示普通模式（Word/Excel/PowerPoint 自由生成）。
    """
    builder = StateGraph(AgentState)

    builder.add_node("agent", _agent_node_factory(doc_path, doc_type))
    builder.add_node("tools", _tools_node)

    builder.add_edge(START, "agent")
    builder.add_conditional_edges("agent", _route_after_agent, {"tools": "tools", END: END})
    builder.add_conditional_edges("tools", _route_after_tools, {"agent": "agent", END: END})

    return builder.compile(checkpointer=MemorySaver())
