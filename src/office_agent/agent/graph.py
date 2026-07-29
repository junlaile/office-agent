"""ReAct agent 图装配（LangGraph 原生写法 + DeepSeek tool_calls）。

结构:
    START → agent → tools → agent ↻ ... → END
                   ↘ prepare_interaction → interaction → agent
                   ↘ nudge → agent（空转纠偏，最多 _MAX_NUDGE 次）

    - agent 节点: 按文档类型筛选工具后调用 LLM.bind_tools，决定调哪个工具
    - tools 节点: 批量执行普通工具，并按元数据处理终止工具
    - interaction 节点: 独占处理表单/确认，直接调用 LangGraph interrupt
    - nudge 节点: agent 只输出文字、没发 tool_calls（"光说不练"）时注入纠偏提示，
        回 agent 重试。避免 DeepSeek 用自然语言描述意图却不调工具、首轮直接判死。
    - 路由: agent 之后，有 tool_calls → tools；有内容但无 tool_calls 且未达上限
            → nudge；否则 END。tools 之后 finish → END，否则回 agent。

防卡死: ① nudge 最多 _MAX_NUDGE 次（超过放行 END）；
        ② agent 节点在接近 recursion_limit 软阈值时会注入"尽快 finish"提示，
          引导 LLM 主动收尾，避免撞硬限制报 GraphRecursionError。

参考: https://api-docs.deepseek.com/zh-cn/guides/tool_calls
（DeepSeek 支持 OpenAI 标准 tools；用 bind_tools，不强制 tool_choice）
"""

from __future__ import annotations

import json
import logging
from collections.abc import Mapping
from typing import Any, Literal, cast

from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langgraph.checkpoint.memory import MemorySaver
from langgraph.errors import GraphBubbleUp
from langgraph.graph import END, START, StateGraph
from langgraph.types import interrupt

from office_agent.cli.user_input import (
    PREFIX_FORCE,
    PREFIX_SUPPLEMENT,
    get_bridge,
)
from office_agent.config import settings
from office_agent.domain.format import kind_from_path
from office_agent.tools import (
    SPEC_BY_NAME,
    TOOL_BY_NAME,
    ExecutionMode,
    SideEffect,
    tools_for_kind,
    tools_for_doc_path as _tools_for_doc_path,
)
from office_agent.tools.batching import (
    BATCH_FALLBACK_PREFIX,
    execute_batched,
    is_batchable,
    take_batch_fallback_reason,
)

from .llm import get_llm
from .prompts import build_system_prompt
from .state import AgentState, InteractionField, InteractionRequest, ToolExecutionRecord

logger = logging.getLogger(__name__)
tools_for_doc_path = _tools_for_doc_path

# 软收尾阈值：当剩余步数 < 限额的此比例时，触发"尽快 finish"提醒。
# 0.7 意味着用掉 70% 预算后开始催促收尾，留 30% 余量完成 view_text+finish。
SOFT_FINISH_RATIO = 0.7

# 空转纠偏上限：LLM 连续输出"有内容但无 tool_calls"的轮数达到此值后，
# 不再 nudge 重试，放行 END（交给 main.py 的"文档未生成"兜底 + recursion_limit
# 硬上限双保险），避免死循环。实测 DeepSeek 在 nudge 后绝大多数会改用 tool_calls。
_MAX_NUDGE = 2


def _count_idle_turns(messages: list) -> int:
    """统计 agent 已连续空转（"有内容但无 tool_calls"）的轮数。

    从消息末尾往前扫，数连续的"无 tool_calls 的 AIMessage"。nudge 节点注入的
    SystemMessage（纠偏提示）会被跳过——它不是 agent 输出，插在两次空转 AIMessage
    之间不应中断计数。遇到 Human/Tool 或有 tool_calls 的 AIMessage 即停。
    """
    count = 0
    for m in reversed(messages or []):
        if isinstance(m, SystemMessage):
            continue  # nudge 注入的纠偏提示，跳过
        if isinstance(m, AIMessage) and not m.tool_calls:
            count += 1
        else:
            break
    return count


def _nudge_node(state: AgentState) -> dict[str, Any]:
    """空转纠偏节点：agent 只输出文字、没调工具时，注入提示后回 agent 重试。

    触发场景（DeepSeek 等）：LLM 用自然语言描述"我先问您几个问题"，
    却不发 ask_user 的 tool_call → 原路由会直接 END、任务判死。
    本节点注入一条 SystemMessage 明确要求"用工具行动"，给 LLM 再来一次的机会。
    """
    hint = SystemMessage(
        content=(
            "【系统纠偏】你上一条只输出了文字、没有调用任何工具，但任务尚未完成。"
            "请直接用工具调用行动（例如 create_doc 创建文档、ask_user 向用户提问、"
            "view_text 查看文档等），不要用自然语言描述你打算做什么。"
            "只输出文字而不发工具调用 = 任务失败。"
        )
    )
    return {"messages": [hint]}


def _agent_node_factory(
    doc_path: str,
    doc_type: str | None = None,
    *,
    template_text: str = "",
    approved_outline: str = "",
    vehicle_mode: bool = False,
):
    """构造 agent 节点。doc_path 用于系统提示词，doc_type 触发公文模式分支。

    template_text：公文模式下已读取的模板正文（view_text 输出），注入系统
    提示词让 LLM 第一轮即看到段落结构，无需自调 view_text。
    approved_outline：用户已批准的 Markdown 大纲。
    """

    kind = kind_from_path(doc_path)
    selected_tools = (
        tools_for_kind(kind, include_vehicle=True) if vehicle_mode else tools_for_doc_path(doc_path)
    )
    llm_with_tools = get_llm().bind_tools(selected_tools)
    system_msg = SystemMessage(
        content=build_system_prompt(
            doc_path,
            doc_type,
            template_text=template_text,
            approved_outline=approved_outline,
        )
    )

    def agent_node(state: AgentState) -> dict[str, Any]:
        messages = list(state.get("messages", []))

        # 节点边界：注入忙时用户补充 / 强制打断（CLI 主循环之外的第二道保险）
        injected: list[HumanMessage] = []
        bridge = get_bridge()
        if bridge is not None:
            force = bridge.consume_force()
            if force:
                injected.append(HumanMessage(content=f"{PREFIX_FORCE}{force}"))
            for text in bridge.drain_soft():
                injected.append(HumanMessage(content=f"{PREFIX_SUPPLEMENT}{text}"))

        limit = settings.recursion_limit
        # 用 messages 数估算已用步数：每轮 agent+tools 约加 2 条消息。
        # 粗略但足够触发软收尾。
        used = len(messages) + len(injected)
        soft_threshold = int(limit * SOFT_FINISH_RATIO)

        prompt_messages: list = [system_msg, *messages, *injected]
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
        # 把注入的 HumanMessage 也写回 state，保证 checkpoint / 后续轮次可见
        return {"messages": [*injected, ai_msg]}

    return agent_node


def _tools_node(state: AgentState) -> dict[str, Any]:
    """执行工具调用，支持独占工具与批处理回退。"""
    messages = state.get("messages", [])
    last_ai: AIMessage = messages[-1]
    tool_messages: list[ToolMessage] = []
    updates: dict[str, Any] = {"messages": tool_messages}
    tool_calls = last_ai.tool_calls or []
    executed = dict(state.get("executed_calls", {}))

    exclusive_names = {
        spec.name
        for tc in tool_calls
        if (spec := SPEC_BY_NAME.get(str(tc.get("name") or ""))) is not None
        and not spec.can_batch
    }
    for tc in tool_calls:
        if str(tc.get("name") or "") in {"ask_user", "finish"}:
            exclusive_names.add(str(tc.get("name") or ""))
    if exclusive_names and len(tool_calls) > 1:
        exclusive_text = "、".join(sorted(exclusive_names))
        for tc in tool_calls:
            name = str(tc.get("name") or "")
            tc_id = str(tc.get("id") or "")
            if name in exclusive_names:
                content = (
                    f"本批包含必须独占执行的工具 {exclusive_text}。"
                    f"请在下一次回复里单独调用 {name}，不要附带其他工具。"
                )
            else:
                content = (
                    f"该 {name} 调用已取消（本批同时包含独占工具 {exclusive_text}），"
                    "请在交互完成后重新发起。"
                )
            tool_messages.append(ToolMessage(content=content, tool_call_id=tc_id))
            executed[tc_id] = ToolExecutionRecord(
                tool_call_id=tc_id,
                tool_name=name,
                argument_keys=sorted(str(key) for key in (tc.get("args") or {})),
                status="cancelled",
                result=content,
            )
        updates["executed_calls"] = executed
        return updates

    def _record(tc_id: str, name: str, args: dict[str, Any], status: Literal["completed", "cancelled", "failed"], result: Any) -> None:
        executed[tc_id] = ToolExecutionRecord(
            tool_call_id=tc_id,
            tool_name=name,
            argument_keys=sorted(str(key) for key in args),
            status=status,
            result=result,
        )

    def _exec_single(tc: Mapping[str, Any]) -> None:
        name = str(tc.get("name") or "")
        args = tc.get("args") or {}
        tc_id = str(tc.get("id") or "")

        if tc_id in executed:
            record = executed[tc_id]
            tool_messages.append(
                ToolMessage(content=str(record.get("result", "")), tool_call_id=tc_id)
            )
            return

        tool = TOOL_BY_NAME.get(name)
        if tool is None:
            result = f"错误: 未知工具 '{name}'"
            tool_messages.append(ToolMessage(content=result, tool_call_id=tc_id))
            _record(tc_id, name, args, "failed", result)
            return

        status: Literal["completed", "cancelled", "failed"] = "completed"
        if name == "finish":
            try:
                result = tool.invoke(args)
            except GraphBubbleUp:
                # finish 在部分流程中走独立确认节点；若工具层仍触发 interrupt，
                # 这里兼容为“确认通过”，避免再追加一次 resume。
                result = f"FINISHED:{args.get('summary', '')}"
            except Exception as e:  # noqa: BLE001
                logger.exception("工具 %s 执行失败", name)
                result = f"工具执行出错({name}): {e}"
                status = "failed"
        else:
            try:
                result = tool.invoke(args)
            except GraphBubbleUp:
                raise
            except Exception as e:  # noqa: BLE001
                logger.exception("工具 %s 执行失败", name)
                result = f"工具执行出错({name}): {e}"
                status = "failed"
        result_str = str(result)
        if name == "finish" and status == "completed":
            if result_str.startswith("FINISHED:"):
                updates["done"] = True
                updates["summary"] = str(args.get("summary") or "")
                result_str = f"已完成: {args.get('summary', '')}"
            else:
                updates.pop("done", None)
                updates.pop("summary", None)
        tool_messages.append(ToolMessage(content=result_str, tool_call_id=tc_id))
        _record(tc_id, name, args, status, result_str)

    pending_batch: list[Mapping[str, Any]] = []

    def _flush_batch() -> None:
        if not pending_batch:
            return
        batch_results = execute_batched([dict(tc) for tc in pending_batch]) if len(pending_batch) >= 2 else None
        if batch_results is not None:
            for tc_id, content in batch_results:
                tool_messages.append(ToolMessage(content=content, tool_call_id=tc_id))
                tc = next((x for x in pending_batch if str(x.get("id") or "") == tc_id), None)
                name = str(tc.get("name") or "") if tc else ""
                args = (tc.get("args") or {}) if tc else {}
                _record(tc_id, name, args, "completed", content)
        else:
            fallback_reason = take_batch_fallback_reason()
            start = len(tool_messages)
            for tc in pending_batch:
                _exec_single(tc)
            if fallback_reason and len(tool_messages) > start:
                first = tool_messages[start]
                patched = ToolMessage(
                    content=(
                        f"{BATCH_FALLBACK_PREFIX} 批量写入失败已回退为逐条执行（{fallback_reason}）\n"
                        f"{first.content}"
                    ),
                    tool_call_id=first.tool_call_id,
                )
                tool_messages[start] = patched
                if first.tool_call_id in executed:
                    executed[first.tool_call_id]["result"] = patched.content
        pending_batch.clear()

    for tc in tool_calls:
        name = str(tc.get("name") or "")
        if name not in {"ask_user", "finish"} and is_batchable(dict(tc)):
            pending_batch.append(tc)
        else:
            _flush_batch()
            _exec_single(tc)
    _flush_batch()

    updates["executed_calls"] = executed
    return updates


def _prepare_interaction_node(state: AgentState) -> dict[str, Any]:
    """把独占工具调用转换为与 UI 无关的交互请求。"""
    last: AIMessage = state.get("messages", [])[-1]
    tc = (last.tool_calls or [])[0]
    name = str(tc.get("name") or "")
    tc_id = str(tc.get("id") or "")
    args = tc.get("args") or {}
    spec = SPEC_BY_NAME.get(name)

    if spec is None:
        return {
            "messages": [
                ToolMessage(content=f"错误: 工具 {name} 不是交互工具", tool_call_id=tc_id)
            ],
            "pending_interaction": None,
        }

    if spec.execution_mode is ExecutionMode.DIRECT and spec.side_effect is not SideEffect.TERMINAL:
        return {
            "messages": [
                ToolMessage(content=f"错误: 工具 {name} 不是交互工具", tool_call_id=tc_id)
            ],
            "pending_interaction": None,
        }

    if spec.execution_mode is ExecutionMode.INTERACTION:
        try:
            payload = spec.tool.invoke(args)
        except Exception as e:  # noqa: BLE001
            logger.exception("构造交互请求失败: %s", name)
            return {
                "messages": [
                    ToolMessage(
                        content=f"工具执行出错({name}): {e}",
                        tool_call_id=tc_id,
                    )
                ],
                "pending_interaction": None,
            }
        raw = dict(payload) if isinstance(payload, Mapping) else {"question": str(payload)}
        raw_fields = raw.get("fields")
        fields = cast(
            list[InteractionField],
            raw_fields if isinstance(raw_fields, list) else [],
        )
        raw_options = raw.get("options")
        options = cast(list[str], raw_options if isinstance(raw_options, list) else [])
        kind: Literal["form", "question", "confirmation"] = "form" if fields else "question"
        request = InteractionRequest(
            request_id=f"interaction-{tc_id}",
            tool_call_id=tc_id,
            tool_name=name,
            kind=kind,
            title=str(raw.get("title", "信息采集")),
            description=str(raw.get("description", "")),
            question=str(raw.get("question", "")),
            fields=fields,
            options=options,
            tool_args=args,
        )
    else:
        request = InteractionRequest(
            request_id=f"confirmation-{tc_id}",
            tool_call_id=tc_id,
            tool_name=name,
            kind="confirmation",
            title=spec.confirmation_title or f"确认执行 {name}",
            description="该操作会修改文档，请确认是否继续。",
            question=f"确认执行工具 {name}？",
            options=["确认", "取消"],
            tool_args=args,
        )
    return {"pending_interaction": request}


def _normalize_interaction_answer(answer: Any) -> dict[str, Any]:
    if isinstance(answer, Mapping):
        if "answers" in answer and isinstance(answer["answers"], Mapping):
            return dict(answer["answers"])
        if "value" in answer and len(answer) <= 3:
            value = str(answer.get("value", "")).strip()
            return {"value": value} if value else {}
        return dict(answer)
    if isinstance(answer, str) and answer.strip():
        return {"value": answer.strip()}
    return {}


def _answer_is_accepted(answer: Any) -> bool:
    if isinstance(answer, Mapping) and "accepted" in answer:
        return bool(answer["accepted"])
    if isinstance(answer, bool):
        return answer
    value = str(answer.get("value", "") if isinstance(answer, Mapping) else answer)
    return value.strip().lower() in {"确认", "是", "同意", "继续", "yes", "y", "ok", "true"}


def _interaction_node(state: AgentState) -> dict[str, Any]:
    """挂起等待交互；恢复后回传 ToolMessage，确认型工具才在确认后执行。"""
    request = state.get("pending_interaction")
    if not request:
        return {}

    tc_id = request.get("tool_call_id", "")
    name = request.get("tool_name", "")
    executed = dict(state.get("executed_calls", {}))
    if tc_id in executed:
        record = executed[tc_id]
        return {
            "messages": [
                ToolMessage(content=str(record.get("result", "")), tool_call_id=tc_id)
            ],
            "pending_interaction": None,
        }

    answer = interrupt(dict(request))
    status: Literal["completed", "cancelled", "failed"] = "completed"

    if request.get("kind") == "confirmation":
        if not _answer_is_accepted(answer):
            status = "cancelled"
            result: Any = {
                "ok": False,
                "code": "user_cancelled",
                "message": f"用户取消执行 {name}",
                "data": None,
                "retryable": False,
            }
        else:
            if name == "finish":
                summary = str((request.get("tool_args") or {}).get("summary") or "")
                result = {
                    "ok": True,
                    "code": "ok",
                    "message": f"已完成: {summary}",
                    "data": {"summary": summary},
                    "retryable": False,
                }
                status = "completed"
            else:
                tool = TOOL_BY_NAME.get(name)
                try:
                    value = tool.invoke(request.get("tool_args") or {}) if tool else None
                    result = {
                        "ok": tool is not None,
                        "code": "ok" if tool is not None else "unknown_tool",
                        "message": str(value) if tool is not None else f"未知工具 {name}",
                        "data": value,
                        "retryable": False,
                    }
                except Exception as e:  # noqa: BLE001
                    logger.exception("确认工具 %s 执行失败", name)
                    status = "failed"
                    result = {
                        "ok": False,
                        "code": "tool_error",
                        "message": str(e),
                        "data": None,
                        "retryable": True,
                    }
    else:
        result = {
            "ok": True,
            "code": "interaction_completed",
            "message": "已收到用户输入",
            "data": _normalize_interaction_answer(answer),
            "retryable": False,
        }

    content = json.dumps(result, ensure_ascii=False, default=str)
    executed[tc_id] = ToolExecutionRecord(
        tool_call_id=tc_id,
        tool_name=name,
        argument_keys=sorted(str(key) for key in (request.get("tool_args") or {})),
        status=status,
        result=content,
    )
    updates: dict[str, Any] = {
        "messages": [ToolMessage(content=content, tool_call_id=tc_id)],
        "pending_interaction": None,
        "executed_calls": executed,
    }
    if request.get("kind") == "confirmation" and name == "finish" and status == "completed":
        summary = str((request.get("tool_args") or {}).get("summary") or "")
        updates["done"] = True
        updates["summary"] = summary
    return updates


def _route_after_prepare_interaction(state: AgentState) -> str:
    return "interaction" if state.get("pending_interaction") else "agent"


def _route_after_tools(state: AgentState) -> str:
    """tools 之后：finish 完成 → END；否则回 agent。"""
    if state.get("done"):
        return END
    return "agent"


def _route_after_agent(state: AgentState) -> str:
    """agent 之后的路由: tools / prepare_interaction / nudge / END。

    - 单个交互/确认工具 → prepare_interaction
    - 普通工具或混合批次 → tools（混合批次由 tools 节点完整配对并拒绝）
    - AIMessage 有内容但【无 tool_calls】且未连续空转过多次 → nudge（纠偏重试）。
      覆盖 LLM"光说不练"（只描述意图、不发工具调用）导致首轮直接判死的失败模式。
    - 其它（无消息 / 非 AIMessage / 纯空串 / 已达 _MAX_NUDGE 上限）→ END。
    """
    messages = state.get("messages", [])
    if not messages:
        return END
    last = messages[-1]
    if isinstance(last, AIMessage) and last.tool_calls:
        if len(last.tool_calls) == 1:
            spec = SPEC_BY_NAME.get(str(last.tool_calls[0].get("name") or ""))
            if spec is not None and spec.execution_mode is not ExecutionMode.DIRECT:
                return "prepare_interaction"
        return "tools"
    # AIMessage 但无 tool_calls：有实质内容且未达空转上限 → 纠偏重试
    if (
        isinstance(last, AIMessage)
        and str(last.content).strip()
        and _count_idle_turns(messages) <= _MAX_NUDGE
    ):
        return "nudge"
    return END


def build_graph(
    doc_path: str,
    doc_type: str | None = None,
    *,
    template_text: str = "",
    approved_outline: str = "",
    vehicle_mode: bool = False,
):
    """构建并编译 ReAct agent 图。

    doc_path: 本会话的文档输出路径（注入到系统提示词和工具会话）。
    doc_type: 法定公文文种名（如 '通知'）。非空时启用公文模式——
        提示词走公文分支，指导 LLM 编辑模板正文而非从零拼接。
        None 表示普通模式（Word/Excel/PowerPoint 自由生成）。
    template_text: 公文模式下已读取的模板正文（view_text 输出，带路径标注），
        注入提示词让 LLM 照路径编辑，避免跳过读模板直接瞎改。
    approved_outline: 用户已批准的 Markdown 大纲，注入系统提示词。
    """
    builder = StateGraph(AgentState)

    builder.add_node(
        "agent",
        _agent_node_factory(
            doc_path,
            doc_type,
            template_text=template_text,
            approved_outline=approved_outline,
            vehicle_mode=vehicle_mode,
        ),
    )
    builder.add_node("tools", _tools_node)
    builder.add_node("prepare_interaction", _prepare_interaction_node)
    builder.add_node("interaction", _interaction_node)
    builder.add_node("nudge", _nudge_node)

    builder.add_edge(START, "agent")
    builder.add_conditional_edges(
        "agent",
        _route_after_agent,
        {
            "tools": "tools",
            "prepare_interaction": "prepare_interaction",
            "nudge": "nudge",
            END: END,
        },
    )
    builder.add_conditional_edges("tools", _route_after_tools, {"agent": "agent", END: END})
    builder.add_conditional_edges(
        "prepare_interaction",
        _route_after_prepare_interaction,
        {"interaction": "interaction", "agent": "agent"},
    )
    builder.add_edge("interaction", "agent")
    # nudge 注入纠偏提示后回 agent 重试
    builder.add_edge("nudge", "agent")

    return builder.compile(checkpointer=MemorySaver())
