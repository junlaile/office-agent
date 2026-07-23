"""office-agent 终端入口（ReAct agent 版）。

交互流程:
    1. 启动 UserInputBridge（单一 stdin 读线程）
    2. 读取配置，检查 officecli
    3. 读取用户需求，确定输出文档路径
    4. Word：大纲预览 → 用户批准/修改/取消（批准前不写 .docx）
    5. 公文：批准后才 _prepare_official_doc；再 build graph
    6. graph.stream 驱动 agent；忙时可补充 / 强制打断 / 继续 / 退出
    7. 遇到 ask_user 触发的 interrupt 时，打印问题、读用户输入、resume
    8. finish 或 LLM 终止时，打印生成的文档路径

用法:
    uv run office-agent
    uv run python -m office_agent "写一份项目周报"
    python main.py  # 兼容 shim
"""

from __future__ import annotations

import logging
import uuid
from pathlib import Path
from typing import Any

from langchain_core.messages import AIMessage, HumanMessage, ToolMessage

from office_agent.cli.ui import (
    _BOLD,
    _CYAN,
    _DIM,
    _GREEN,
    _RED,
    _RESET,
    _YELLOW,
    _banner,
    _check_officecli,
    _derive_doc_path,
    _handle_interrupt,
    _prepare_official_doc,
    _print_agent_step,
    _print_tool_results,
    _read_requirement,
    _readline,
    _run_outline_approval_loop,
)
from office_agent.cli.user_input import (
    CONTINUE_PROMPT,
    PREFIX_FORCE,
    PREFIX_SUPPLEMENT,
    UserInputBridge,
    set_bridge,
)
from office_agent.config import assert_llm_ready, settings
from office_agent.domain.templates import detect_doc_type
from office_agent.tools import set_session_doc

logger = logging.getLogger("office_agent.cli.main")


def main() -> None:
    """Console script / ``python -m office_agent`` 入口。"""
    raise SystemExit(run())


def run() -> int:
    bridge = UserInputBridge()
    set_bridge(bridge)
    bridge.start()
    bridge.set_busy(False)

    try:
        return _run_with_bridge(bridge)
    finally:
        bridge.stop()
        set_bridge(None)


def _run_with_bridge(bridge: UserInputBridge) -> int:
    _banner()
    assert_llm_ready()
    if not _check_officecli():
        return 1

    requirement = _read_requirement()

    # 先识别公文文种：命中则强制 .docx 并跳过文档类型询问
    doc_type = detect_doc_type(requirement)
    if doc_type:
        print(f"{_GREEN}✓ 识别为法定公文【{doc_type}】{_RESET}")
    doc_path = _derive_doc_path(requirement, doc_type=doc_type)
    print(f"{_DIM}输出文件: {doc_path}{_RESET}")

    approved_outline = ""
    template_text = ""

    # Word：批准大纲前不落盘；Excel/PPT 跳过预览门控
    if doc_path.lower().endswith(".docx"):
        bridge.set_busy(False)
        approved = _run_outline_approval_loop(requirement, doc_type)
        if approved is None:
            return 0
        approved_outline = approved
        # 公文模板合并推迟到大纲批准之后
        if doc_type:
            doc_type, template_text = _prepare_official_doc(doc_type, doc_path)
            doc_type = doc_type or None
    print()

    from office_agent.agent.graph import build_graph

    set_session_doc(doc_path)
    graph = build_graph(
        doc_path,
        doc_type=doc_type,
        template_text=template_text,
        approved_outline=approved_outline,
    )
    thread_id = str(uuid.uuid4())
    config = {
        "configurable": {"thread_id": thread_id},
        "recursion_limit": settings.recursion_limit,
    }

    initial: dict[str, Any] = {
        "messages": [],
        "doc_path": doc_path,
    }
    if approved_outline:
        first_content = f"{requirement}\n\n【已批准的文档大纲】\n{approved_outline}"
    else:
        first_content = requirement
    first_input = [HumanMessage(content=first_content)]

    print(f"{_DIM}开始生成…（Agent 会自主调用工具，缺关键信息时会问你）{_RESET}")
    print(
        f"{_DIM}忙时输入: 普通文字=补充  !内容=强制打断  继续=接着做  退出=结束{_RESET}\n"
    )

    try:
        bridge.set_busy(True)
        _stream(graph, initial, first_input, config)

        while True:
            pending_rc = _handle_pending_inputs(graph, config, bridge)
            if pending_rc is not None:
                return pending_rc

            # ask_user interrupt：空闲读 stdin
            bridge.set_busy(False)
            cmd = _handle_interrupt(graph, config)
            if cmd is None:
                break
            bridge.set_busy(True)
            _stream(graph, None, None, config, command=cmd)

    except KeyboardInterrupt:
        print(f"\n{_YELLOW}已中断。{_RESET}")
        return 130
    except Exception as e:  # noqa: BLE001
        from langgraph.errors import GraphRecursionError

        if isinstance(e, GraphRecursionError):
            print(
                f"\n{_YELLOW}⚠ Agent 达到步数上限（{settings.recursion_limit}）"
                f"未显式完成。{_RESET}"
            )
            if Path(doc_path).exists():
                print(f"{_GREEN}已生成部分文档（未显式 finish）:{_RESET} {doc_path}")
                return 0
            print(f"{_RED}未能生成文档。{_RESET}")
            return 1
        logger.error("运行出错: %s", e)
        logger.exception("agent 运行异常")
        return 1

    return _print_final_result(graph, config, doc_path)


def _handle_pending_inputs(graph, config: dict, bridge: UserInputBridge) -> int | None:
    """处理忙时队列：QUIT 返回退出码；注入消息后继续 stream；无 pending 返回 None。"""
    while True:
        if bridge.consume_quit():
            print(f"\n{_YELLOW}已退出。{_RESET}")
            return 0

        if bridge.consume_soft_pause():
            _soft_pause_repl(bridge)
            continue

        msgs: list[HumanMessage] = []
        force = bridge.consume_force()
        if force:
            msgs.append(HumanMessage(content=f"{PREFIX_FORCE}{force}"))
            print(f"{_YELLOW}⚡ 强制打断:{_RESET} {force}")
        for text in bridge.drain_soft():
            msgs.append(HumanMessage(content=f"{PREFIX_SUPPLEMENT}{text}"))
            print(f"{_DIM}✎ 用户补充:{_RESET} {text}")
        if bridge.consume_continue():
            msgs.append(HumanMessage(content=CONTINUE_PROMPT))
            print(f"{_DIM}→ 继续生成…{_RESET}")

        if not msgs:
            return None

        graph.update_state(config, {"messages": msgs})
        bridge.set_busy(True)
        _stream(graph, None, None, config)
        # 再检查是否又有新 pending（循环由 while True 处理）


def _soft_pause_repl(bridge: UserInputBridge) -> None:
    """软暂停 REPL：等待「继续」或新指令后再恢复。"""
    bridge.set_busy(False)
    print(
        f"\n{_YELLOW}⏸ 已暂停。输入「继续」恢复；"
        f"也可输入补充 / !强制 / 退出{_RESET}"
    )
    while True:
        try:
            line = _readline(f"{_CYAN}暂停 ❯ {_RESET}")
        except (EOFError, KeyboardInterrupt):
            bridge.submit("退出")
            return
        item = bridge.submit(line)
        if item is None:
            continue
        # submit 已投递到队列；外层 _handle_pending_inputs 会消费
        if item.kind.value in ("continue", "force", "supplement", "quit"):
            return


def _print_final_result(graph, config: dict, doc_path: str) -> int:
    snapshot = graph.get_state(config)
    vals = snapshot.values if snapshot else {}
    done = vals.get("done")
    summary = vals.get("summary", "")
    print(f"\n{_BOLD}{_CYAN}════════════════════════════════════════════{_RESET}")
    if done:
        print(f"{_GREEN}{_BOLD}✓ 文档已生成:{_RESET} {doc_path}")
        if summary:
            print(f"{_DIM}总结: {summary}{_RESET}")
        return 0
    if Path(doc_path).exists():
        print(f"{_YELLOW}文档已生成（未显式 finish）:{_RESET} {doc_path}")
        return 0
    print(f"{_RED}未能生成文档。{_RESET}")
    return 1


def _stream(graph, initial, first_input, config, *, command=None) -> None:
    """stream 一段。展示 agent 工具调用与工具结果。"""
    if command is not None:
        stream_input = command
    elif initial is not None:
        stream_input = {**initial, "messages": first_input}
    else:
        # 仅 update_state 后继续：传 None 让 LangGraph 从 checkpoint 恢复
        stream_input = None

    for chunk in graph.stream(stream_input, config=config, stream_mode="updates"):
        for _node, updates in chunk.items():
            if not isinstance(updates, dict):
                continue
            new_msgs = updates.get("messages") or []
            for m in new_msgs:
                if isinstance(m, AIMessage):
                    _print_agent_step(m)
                elif isinstance(m, ToolMessage):
                    _print_tool_results(m)


if __name__ == "__main__":
    main()
