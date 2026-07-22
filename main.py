"""office-agent 终端入口（ReAct agent 版）。

交互流程:
    1. 读取配置，检查 officecli
    2. 读取用户需求，确定输出文档路径
    3. build graph（注入 doc_path）+ 设会话文档路径
    4. graph.stream 驱动 agent；展示每一步工具调用
    5. 遇到 ask_user 触发的 interrupt 时，打印问题、读用户输入、resume
    6. finish 或 LLM 终止时，打印生成的 docx 路径

用法:
    python main.py
    python main.py "写一份项目周报，包含本周进展、风险、下周计划"

UI 辅助函数（需求读取、文档类型推断、工具调用展示、interrupt 表单收集等）
已下沉到 :mod:`office_agent.cli_ui`，本文件只保留 run() 主流程编排。
"""

from __future__ import annotations

import logging
import sys
import uuid
from pathlib import Path
from typing import Any

# src 布局：把 src/ 加入 sys.path，让 `python main.py` 能直接跑
_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from langchain_core.messages import AIMessage, HumanMessage, ToolMessage  # noqa: E402

from office_agent.cli_ui import (  # noqa: E402
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
)
from office_agent.config import assert_llm_ready, settings  # noqa: E402
from office_agent.templates import detect_doc_type  # noqa: E402
from office_agent.tools import set_session_doc  # noqa: E402

logger = logging.getLogger("office_agent.main")


def run() -> int:
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

    # 公文模式：从模板预创建文档并预填版头（失败则回退普通模式）。
    # _prepare_official_doc 同时预读模板正文（template_text），注入提示词，
    # 让 LLM 第一轮就看到段落结构，避免跳过读模板直接瞎改。
    template_text = ""
    if doc_type:
        doc_type, template_text = _prepare_official_doc(doc_type, doc_path)
        doc_type = doc_type or None
    print()

    # 延迟 import（确保 officecli 已就绪）
    from office_agent.graph import build_graph

    set_session_doc(doc_path)
    graph = build_graph(doc_path, doc_type=doc_type, template_text=template_text)
    thread_id = str(uuid.uuid4())
    config = {
        "configurable": {"thread_id": thread_id},
        "recursion_limit": settings.recursion_limit,
    }

    initial: dict[str, Any] = {
        "messages": [],  # 第一条 HumanMessage 由下面加入
        "doc_path": doc_path,
    }
    first_input = [HumanMessage(content=requirement)]

    print(f"{_DIM}开始生成…（Agent 会自主调用工具，缺关键信息时会问你）{_RESET}\n")

    try:
        # 首轮：带初始 HumanMessage 启动
        _stream(graph, initial, first_input, config)

        # interrupt 循环
        while True:
            cmd = _handle_interrupt(graph, config)
            if cmd is None:
                break
            _stream(graph, None, None, config, command=cmd)

    except KeyboardInterrupt:
        print(f"\n{_YELLOW}已中断。{_RESET}")
        return 130
    except Exception as e:  # noqa: BLE001
        # 友好处理递归超限：尽力交付已生成的文档
        from langgraph.errors import GraphRecursionError

        if isinstance(e, GraphRecursionError):
            print(
                f"\n{_YELLOW}⚠ Agent 达到步数上限（{settings.recursion_limit}）未显式完成。{_RESET}"
            )
            if Path(doc_path).exists():
                print(f"{_GREEN}已生成部分文档（未显式 finish）:{_RESET} {doc_path}")
                return 0
            print(f"{_RED}未能生成文档。{_RESET}")
            return 1
        logger.error("运行出错: %s", e)
        logger.exception("agent 运行异常")
        return 1

    # 最终结果
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
    # 没显式 finish 但已结束
    if Path(doc_path).exists():
        print(f"{_YELLOW}文档已生成（未显式 finish）:{_RESET} {doc_path}")
        return 0
    print(f"{_RED}未能生成文档。{_RESET}")
    return 1


def _stream(graph, initial, first_input, config, *, command=None) -> None:
    """stream 一段。展示 agent 工具调用与工具结果。"""
    if command is not None:
        # resume 场景：用 Command 继续
        # command 内含 resume 值，graph 会把它喂给挂起的 ask_user 工具
        stream_input = command
    else:
        stream_input = {**initial, "messages": first_input}

    for chunk in graph.stream(stream_input, config=config, stream_mode="updates"):
        # chunk: {node_name: state_update}
        for node, updates in chunk.items():
            if not isinstance(updates, dict):
                continue
            new_msgs = updates.get("messages") or []
            for m in new_msgs:
                if isinstance(m, AIMessage):
                    _print_agent_step(m)
                elif isinstance(m, ToolMessage):
                    _print_tool_results(m)


if __name__ == "__main__":
    raise SystemExit(run())
