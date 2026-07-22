"""office-agent 终端入口（ReAct agent 版）。

交互流程:
    1. 读取配置，检查 officecli
    2. 读取用户需求，确定输出文档路径
    3. build graph（注入 doc_path）+ 设会话文档路径 + 启动忙时输入桥
    4. graph.stream 驱动 agent；展示每一步工具调用
    5. 忙时可打字补充；! / 强制 打断；Ctrl+C 软暂停；继续指令恢复
    6. 遇到 ask_user 触发的 interrupt 时，打印问题、读用户输入、resume
    7. finish 或用户退出时，打印生成的 docx 路径

用法:
    python main.py
    python main.py "写一份项目周报，包含本周进展、风险、下周计划"
"""

from __future__ import annotations

import re
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

# src 布局：把 src/ 加入 sys.path，让 `python main.py` 能直接跑
_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from langchain_core.messages import AIMessage, HumanMessage, ToolMessage
from langgraph.types import Command

from office_agent.config import assert_llm_ready, settings
from office_agent.format_detect import detect_format, format_label
from office_agent.officecli import OfficeCLIError, resolve_bin
from office_agent.tools import set_session_doc, set_session_format
from office_agent.user_input import (
    CONTINUE_PROMPT,
    PREFIX_FORCE,
    PREFIX_SUPPLEMENT,
    InputKind,
    UserInputBridge,
    classify,
    set_bridge,
)

# ANSI 颜色（Windows 10+ 终端支持）
_CYAN = "\033[36m"
_YELLOW = "\033[33m"
_GREEN = "\033[32m"
_RED = "\033[31m"
_MAGENTA = "\033[35m"
_DIM = "\033[2m"
_BOLD = "\033[1m"
_RESET = "\033[0m"


def _banner() -> None:
    print(f"\n{_BOLD}{_CYAN}╔══════════════════════════════════════════╗\n"
          f"║   Office Agent — Word / Excel / PPT       ║\n"
          f"║   ReAct + DeepSeek tool_calls + OfficeCLI ║\n"
          f"╚══════════════════════════════════════════╝{_RESET}\n")
    print(f"{_DIM}按需求自动识别格式（报表→Excel，PPT→演示，默认 Word）{_RESET}")
    print(f"{_DIM}忙时可直接输入补充；!内容 或 强制:内容 打断；"
          f"继续 / continue 恢复；Ctrl+C 软暂停{_RESET}\n")


def _check_officecli() -> bool:
    try:
        bin_path = resolve_bin()
        print(f"{_DIM}[ok] officecli: {bin_path}{_RESET}")
        return True
    except OfficeCLIError as e:
        print(f"{_RED}[error] {e}{_RESET}")
        print(f"{_YELLOW}请先运行: python scripts/fetch_officecli.py{_RESET}")
        return False


def _read_requirement(bridge: UserInputBridge) -> str:
    if len(sys.argv) > 1:
        req = " ".join(sys.argv[1:]).strip()
        if req:
            print(f"{_DIM}需求: {req}{_RESET}\n")
            return req
    print(f"{_BOLD}请描述你要生成的文档（回车提交）:{_RESET}")
    print(f"{_DIM}例如: 写一份项目周报 / 做一份销售月度报表 / 做一份项目汇报 PPT{_RESET}\n")
    while True:
        try:
            req = bridge.blocking_readline(f"{_CYAN}❯ {_RESET}").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{_YELLOW}已取消。{_RESET}")
            sys.exit(0)
        if req:
            return req
        print(f"{_YELLOW}需求不能为空，请重新输入。{_RESET}")


def _derive_doc_path(requirement: str, fmt: str = "docx") -> str:
    """从需求与格式生成输出文档路径。"""
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    # 提取需求里的中文/字母数字作文件名
    safe = re.sub(r'[\\/:*?"<>|]', "", requirement).strip()
    safe = re.sub(r"\s+", "_", safe)
    # 截断并兜底
    safe = safe[:30] or "document"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    ext = fmt if fmt in ("docx", "xlsx", "pptx") else "docx"
    return str((settings.output_dir / f"{safe}_{stamp}.{ext}").resolve())


def _indent(text: str, pad: str = "    ") -> str:
    return "\n".join(pad + line for line in text.splitlines())


def _format_tool_call(name: str, args: dict) -> str:
    """把工具调用格式化成一行可读文字。"""
    if name == "add_title":
        return f'add_title("{args.get("text", "")}")'
    if name == "add_heading":
        return f'add_heading("{args.get("text", "")}", level={args.get("level", 1)})'
    if name == "add_paragraph":
        bold = args.get("bold", False)
        italic = args.get("italic", False)
        align = args.get("align") or ""
        extras = []
        if bold:
            extras.append("bold")
        if italic:
            extras.append("italic")
        if align:
            extras.append(f'align="{align}"')
        extra = (", " + ", ".join(extras)) if extras else ""
        text = args.get("text", "")
        preview = text[:30] + ("…" if len(text) > 30 else "")
        return f'add_paragraph("{preview}"{extra})'
    if name == "add_list_item":
        ordered = args.get("ordered", False)
        level = args.get("level", 0)
        lvl = f", level={level}" if level else ""
        return f'add_list_item("{args.get("text", "")}"{", ordered" if ordered else ""}{lvl})'
    if name == "add_table":
        data = args.get("data") or []
        rows = len(data)
        cols = max((len(r) for r in data), default=0)
        style = args.get("style", "medium2")
        return f'add_table({rows}×{cols}, style="{style}")'
    if name == "add_image":
        cap = args.get("caption", "")
        cap_str = f', caption="{cap}"' if cap else ""
        h = args.get("height") or ""
        h_str = f', height="{h}"' if h else ""
        src = str(args.get("url_or_path", ""))
        src_short = src if len(src) <= 30 else src[:27] + "..."
        return f'add_image("{src_short}", width="{args.get("width", "8cm")}"{h_str}{cap_str})'
    if name == "add_header":
        pn = ", page_numbers" if args.get("page_numbers") else ""
        return f'add_header("{args.get("text", "")}"{pn})'
    if name == "add_footer":
        pn = ", page_numbers" if args.get("page_numbers") else ""
        return f'add_footer("{args.get("text", "")}"{pn})'
    if name == "add_page_break":
        return "add_page_break()"
    if name == "replace_text":
        return f'replace_text("{args.get("find", "")}" → "{args.get("replace", "")}")'
    if name == "batch_add":
        ops = args.get("ops") or []
        return f"batch_add({len(ops)} ops)"
    if name == "view_outline":
        return "view_outline()"
    if name == "create_workbook":
        return "create_workbook()"
    if name == "add_sheet":
        return f'add_sheet("{args.get("name", "")}")'
    if name == "write_range":
        data = args.get("data") or []
        rows = len(data)
        cols = max((len(r) for r in data), default=0)
        return (
            f'write_range("{args.get("sheet", "Sheet1")}", '
            f'"{args.get("start", "A1")}", {rows}×{cols})'
        )
    if name == "write_cell":
        fml = args.get("formula") or ""
        if fml:
            return (
                f'write_cell("{args.get("sheet", "")}", '
                f'"{args.get("ref", "")}", formula="{fml}")'
            )
        return (
            f'write_cell("{args.get("sheet", "")}", '
            f'"{args.get("ref", "")}", "{args.get("value", "")}")'
        )
    if name == "view_sheet":
        return "view_sheet()"
    if name == "create_presentation":
        return "create_presentation()"
    if name == "add_slide":
        return f'add_slide("{args.get("title", "")}")'
    if name == "add_bullets":
        items = args.get("items") or []
        return f'add_bullets(slide={args.get("slide_index", 1)}, {len(items)} items)'
    if name == "add_slide_table":
        data = args.get("data") or []
        rows = len(data)
        cols = max((len(r) for r in data), default=0)
        return f'add_slide_table(slide={args.get("slide_index", 1)}, {rows}×{cols})'
    if name == "add_slide_image":
        src = str(args.get("url_or_path", ""))
        src_short = src if len(src) <= 30 else src[:27] + "..."
        return (
            f'add_slide_image(slide={args.get("slide_index", 1)}, '
            f'"{src_short}", width="{args.get("width", "12cm")}")'
        )
    if name == "query_vehicle":
        return f'query_vehicle("{args.get("plate_number", "")}")'
    if name == "ask_user":
        fields = args.get("fields", []) or []
        title = args.get("title", "")
        if fields:
            return f'ask_user("{title}", {len(fields)}个字段)'
        return f'ask_user("{title}")'
    if name == "finish":
        return f'finish("{args.get("summary", "")}")'
    # create_doc / view_text / validate_doc 等无参工具
    return f"{name}()"


def _print_agent_step(message: AIMessage) -> None:
    """展示 agent 的一步：思考内容（若有）+ 工具调用。"""
    if message.content:
        text = str(message.content).strip()
        if text:
            print(f"{_MAGENTA}{_DIM}[思考]{_RESET} {_DIM}{text[:200]}{_RESET}")
    for tc in message.tool_calls or []:
        name = tc.get("name", "")
        args = tc.get("args") or {}
        call_str = _format_tool_call(name, args)
        icon = "🔧"
        if name == "ask_user":
            icon = "❓"
        elif name == "finish":
            icon = "✅"
        print(f"  {icon} {_CYAN}{call_str}{_RESET}")


def _print_tool_results(message: ToolMessage) -> None:
    """展示工具返回结果（简短）。"""
    content = str(message.content).strip()
    if not content:
        return
    # 截断长输出（如 view_text）
    preview = content[:150] + ("…" if len(content) > 150 else "")
    # 多行的话只显示首行
    first_line = preview.split("\n")[0]
    print(f"     {_DIM}↳ {first_line}{_RESET}")


def _readline(bridge: UserInputBridge, prompt: str) -> str:
    try:
        return bridge.blocking_readline(prompt).strip()
    except KeyboardInterrupt:
        raise
    except EOFError:
        print(f"\n{_YELLOW}已取消。{_RESET}")
        sys.exit(0)


def _handle_interrupt(graph, config: dict, bridge: UserInputBridge) -> Command | None:
    """若当前挂在 interrupt，按 payload 形态渲染并收集用户输入。

    支持两种 payload:
      - 表单卡片（新）: {title, description, fields:[{key,label,required,options,hint}]}
      - 单问题（旧）  : {question, options}  —— 向后兼容
    """
    snapshot = graph.get_state(config)
    if snapshot is None:
        return None
    interrupts = []
    for t in snapshot.tasks or []:
        if hasattr(t, "interrupts") and t.interrupts:
            interrupts.extend(t.interrupts)
    if not interrupts:
        return None

    payload = interrupts[0].value or {}
    bridge.set_busy(False)

    # 分支：表单卡片 vs 单问题
    if isinstance(payload, dict) and payload.get("fields"):
        resume_value = _collect_form(payload, bridge)
    else:
        resume_value = _collect_single_question(payload, bridge)

    print(f"\n{_DIM}已收到，继续生成…{_RESET}")
    return Command(resume=resume_value)


def _collect_single_question(payload: dict, bridge: UserInputBridge) -> str:
    """旧格式：单问题自由输入 + 可选候选。"""
    question = payload.get("question", "请确认")
    options = payload.get("options", []) or []

    print(f"\n{_BOLD}{_YELLOW}❓ Agent 需要你确认:{_RESET}")
    print(f"{_indent(question)}")
    if options:
        print(f"{_DIM}可选答案（输入序号或文字，也可自行输入）:{_RESET}")
        for i, o in enumerate(options, 1):
            print(f"  {i}. {o}")
    print()
    while True:
        try:
            ans = _readline(bridge, f"{_CYAN}答 ❯ {_RESET}")
        except KeyboardInterrupt:
            print(f"\n{_YELLOW}已取消。{_RESET}")
            sys.exit(0)
        if not ans:
            print(f"{_YELLOW}请输入答案（或 Ctrl-C 取消）。{_RESET}")
            continue
        if options and ans.isdigit():
            idx = int(ans) - 1
            if 0 <= idx < len(options):
                ans = options[idx]
        break
    return ans


def _collect_form(payload: dict, bridge: UserInputBridge) -> dict:
    """新格式：多字段表单卡片，逐字段收集 + 必填校验。"""
    title = payload.get("title", "信息采集")
    description = payload.get("description", "")
    fields = payload.get("fields", []) or []

    # 卡片头
    print(f"\n{_BOLD}{_YELLOW}┌─ ❓ {title} ", end="")
    # 标题右侧补齐横线（粗略对齐）
    pad = max(0, 44 - len(title) - 6)  # 6 = " ❓ " 的视觉宽度估算
    print("─" * pad + f"┐{_RESET}")
    if description:
        print(f"{_YELLOW}│{_RESET} {_DIM}{description}{_RESET}")
    print(f"{_YELLOW}│{_RESET} {_DIM}带 * 为必填；有候选的可输序号或自由输入；可选字段可回车跳过{_RESET}")
    print(f"{_YELLOW}│{_RESET}")

    answers: dict[str, str] = {}
    for i, f in enumerate(fields, 1):
        key = f.get("key", f"field{i}")
        label = f.get("label", key)
        required = bool(f.get("required", False))
        options = f.get("options", []) or []
        hint = f.get("hint", "")

        star = f"{_RED}*{_RESET}" if required else f"{_DIM} {_RESET}"
        print(f"{_YELLOW}│{_RESET} {_BOLD}{i}. {label}{_RESET} {star}")
        if hint:
            print(f"{_YELLOW}│{_RESET}   {_DIM}提示: {hint}{_RESET}")
        if options:
            opts_str = "  ".join(
                f"{_CYAN}[{n}]{_RESET} {o}" for n, o in enumerate(options, 1)
            )
            print(f"{_YELLOW}│{_RESET}   {_DIM}候选: {_RESET}{opts_str}")

        # 逐字段输入循环（必填校验）
        while True:
            try:
                val = _readline(bridge, f"{_YELLOW}│{_RESET} {_CYAN}❯ {_RESET}")
            except KeyboardInterrupt:
                print(f"\n{_YELLOW}已取消。{_RESET}")
                sys.exit(0)
            if not val and required:
                print(f"{_YELLOW}│{_RESET}   {_RED}此项必填，请输入（或 Ctrl-C 取消）。{_RESET}")
                continue
            # 候选序号映射
            if val and options and val.isdigit():
                idx = int(val) - 1
                if 0 <= idx < len(options):
                    val = options[idx]
            break
        answers[key] = val

    print(f"{_YELLOW}└" + "─" * 52 + f"┘{_RESET}")
    return answers


def _graph_has_interrupt(graph, config: dict) -> bool:
    snapshot = graph.get_state(config)
    if snapshot is None:
        return False
    for t in snapshot.tasks or []:
        if hasattr(t, "interrupts") and t.interrupts:
            return True
    return False


def _graph_explicitly_done(graph, config: dict) -> bool:
    snapshot = graph.get_state(config)
    if snapshot is None:
        return False
    return bool((snapshot.values or {}).get("done"))


def _graph_idle(graph, config: dict) -> bool:
    """无后续节点且无 ask_user interrupt。"""
    snapshot = graph.get_state(config)
    if snapshot is None:
        return True
    next_nodes = getattr(snapshot, "next", None) or ()
    if next_nodes:
        return False
    return not _graph_has_interrupt(graph, config)


def _inject_and_clear_done(graph, config: dict, messages: list[HumanMessage]) -> None:
    """注入用户消息并清 done，便于续跑或追加修改。"""
    if not messages:
        return
    graph.update_state(config, {"messages": messages, "done": False})


def _pending_to_messages(bridge: UserInputBridge) -> list[HumanMessage]:
    """消费 bridge 上已排队的 force/soft/continue，转为 HumanMessage。"""
    msgs: list[HumanMessage] = []
    force = bridge.consume_force()
    if force:
        msgs.append(HumanMessage(content=f"{PREFIX_FORCE}{force}"))
    for text in bridge.drain_soft():
        msgs.append(HumanMessage(content=f"{PREFIX_SUPPLEMENT}{text}"))
    if bridge.consume_continue():
        msgs.append(HumanMessage(content=CONTINUE_PROMPT))
    return msgs


def _ensure_resumable(graph, config: dict) -> None:
    """若图已在 END，把续跑入口拨回 agent（via as_node=tools → 下一跳 agent）。"""
    snapshot = graph.get_state(config)
    if snapshot is None:
        return
    next_nodes = getattr(snapshot, "next", None) or ()
    if next_nodes:
        return
    # 已结束：伪装一次 tools 更新，路由回到 agent
    graph.update_state(config, values={}, as_node="tools")


def _safe_to_inject(graph, config: dict) -> bool:
    """避免在 AIMessage.tool_calls 与对应 ToolMessage 之间插入 HumanMessage。"""
    snapshot = graph.get_state(config)
    if snapshot is None:
        return True
    msgs = (snapshot.values or {}).get("messages") or []
    if not msgs:
        return True
    last = msgs[-1]
    if isinstance(last, AIMessage) and last.tool_calls:
        return False
    return True


def _soft_pause_repl(bridge: UserInputBridge) -> list[HumanMessage] | None:
    """软暂停交互。返回待注入消息；None 表示用户退出。"""
    bridge.set_busy(False)
    bridge.consume_soft_pause()
    # 先消化暂停前已排队的输入
    queued = _pending_to_messages(bridge)
    if queued:
        return queued

    print(f"\n{_YELLOW}⏸ 已暂停。{_RESET} "
          f"{_DIM}输入「继续」恢复；直接打字补充；「退出」结束{_RESET}")
    while True:
        if bridge.consume_quit():
            return None
        # 忙时线程可能又投递了（若曾误设 busy）
        queued = _pending_to_messages(bridge)
        if queued:
            return queued
        try:
            line = _readline(bridge, f"{_CYAN}暂停 ❯ {_RESET}")
        except KeyboardInterrupt:
            print(f"\n{_YELLOW}再次中断，退出。{_RESET}")
            return None
        item = classify(line)
        if item is None:
            continue
        if item.kind == InputKind.QUIT:
            return None
        if item.kind == InputKind.CONTINUE:
            return [HumanMessage(content=CONTINUE_PROMPT)]
        if item.kind == InputKind.FORCE:
            return [HumanMessage(content=f"{PREFIX_FORCE}{item.text}")]
        if item.kind == InputKind.SUPPLEMENT:
            return [HumanMessage(content=f"{PREFIX_SUPPLEMENT}{item.text}")]


def _print_final(graph, config: dict, doc_path: str) -> int:
    snapshot = graph.get_state(config)
    vals = snapshot.values if snapshot else {}
    done = vals.get("done")
    summary = vals.get("summary", "")
    print(f"\n{_BOLD}{_CYAN}════════════════════════════════════════════{_RESET}")
    if done:
        print(f"{_GREEN}{_BOLD}✓ 已生成:{_RESET} {doc_path}")
        if summary:
            print(f"{_DIM}总结: {summary}{_RESET}")
        return 0
    from pathlib import Path as _P
    if _P(doc_path).exists():
        print(f"{_YELLOW}已生成（未显式 finish）:{_RESET} {doc_path}")
        return 0
    print(f"{_RED}未能生成文件。{_RESET}")
    return 1


def run() -> int:
    _banner()
    assert_llm_ready()
    if not _check_officecli():
        return 1

    bridge = UserInputBridge()
    set_bridge(bridge)
    bridge.start()

    try:
        return _run_with_bridge(bridge)
    finally:
        bridge.stop()
        set_bridge(None)


def _run_with_bridge(bridge: UserInputBridge) -> int:
    requirement = _read_requirement(bridge)
    fmt = detect_format(requirement)
    doc_path = _derive_doc_path(requirement, fmt)
    print(f"{_DIM}识别格式: {format_label(fmt)}（.{fmt}）{_RESET}")
    print(f"{_DIM}输出文件: {doc_path}{_RESET}\n")

    from office_agent.graph import build_graph

    set_session_doc(doc_path)
    set_session_format(fmt)
    graph = build_graph(doc_path, fmt)
    thread_id = str(uuid.uuid4())
    config = {
        "configurable": {"thread_id": thread_id},
        "recursion_limit": settings.recursion_limit,
    }

    initial: dict[str, Any] = {
        "messages": [],
        "doc_path": doc_path,
    }
    first_input = [HumanMessage(content=requirement)]

    print(f"{_DIM}开始生成…（忙时可输入补充；! 强制打断；缺信息时会问你）{_RESET}\n")

    try:
        # 首轮
        reason = _stream(graph, initial, first_input, config, bridge)

        while True:
            # 1) 软暂停（若卡在 tool_calls 之后，先续跑完 tools 再进 REPL）
            if reason == "pause" or bridge.peek_soft_pause():
                while (
                    not _safe_to_inject(graph, config)
                    and not _graph_has_interrupt(graph, config)
                    and not _graph_idle(graph, config)
                ):
                    reason = _stream(graph, None, None, config, bridge)
                injected = _soft_pause_repl(bridge)
                if injected is None:
                    return _print_final(graph, config, doc_path)
                _inject_and_clear_done(graph, config, injected)
                _ensure_resumable(graph, config)
                print(f"\n{_DIM}已收到，继续生成…{_RESET}\n")
                reason = _stream(graph, None, None, config, bridge)
                continue

            # 2) ask_user interrupt
            if _graph_has_interrupt(graph, config):
                cmd = _handle_interrupt(graph, config, bridge)
                if cmd is None:
                    break
                reason = _stream(graph, None, None, config, bridge, command=cmd)
                continue

            # 3) 强制 / 已排队补充 / 继续
            if bridge.has_pending():
                if bridge.consume_quit():
                    return _print_final(graph, config, doc_path)
                if bridge.consume_soft_pause():
                    reason = "pause"
                    continue
                injected = _pending_to_messages(bridge)
                if injected:
                    print(f"\n{_YELLOW}⚡ 已打断，正在理解你的补充…{_RESET}\n")
                    _inject_and_clear_done(graph, config, injected)
                    _ensure_resumable(graph, config)
                    reason = _stream(graph, None, None, config, bridge)
                    continue

            # 4) 显式 finish 且空闲 → 有排队则追加，否则结束
            if _graph_idle(graph, config) and _graph_explicitly_done(graph, config):
                if bridge.has_pending() and not bridge.consume_quit():
                    injected = _pending_to_messages(bridge)
                    if injected:
                        _inject_and_clear_done(graph, config, injected)
                        _ensure_resumable(graph, config)
                        reason = _stream(graph, None, None, config, bridge)
                        continue
                break

            # 5) 空闲但未 finish → 软暂停，允许「继续」或补充
            if _graph_idle(graph, config):
                injected = _soft_pause_repl(bridge)
                if injected is None:
                    return _print_final(graph, config, doc_path)
                _inject_and_clear_done(graph, config, injected)
                _ensure_resumable(graph, config)
                print(f"\n{_DIM}已收到，继续生成…{_RESET}\n")
                reason = _stream(graph, None, None, config, bridge)
                continue

            # 6) 仍有 next 节点（例如 force 后未抽干）→ 继续 stream
            reason = _stream(graph, None, None, config, bridge)

    except Exception as e:  # noqa: BLE001
        from langgraph.errors import GraphRecursionError
        if isinstance(e, GraphRecursionError):
            print(f"\n{_YELLOW}⚠ Agent 达到步数上限（{settings.recursion_limit}）未显式完成。{_RESET}")
            from pathlib import Path as _P
            if _P(doc_path).exists():
                print(f"{_GREEN}已生成部分文档（未显式 finish）:{_RESET} {doc_path}")
                return 0
            print(f"{_RED}未能生成文档。{_RESET}")
            return 1
        print(f"\n{_RED}运行出错:{_RESET} {e}")
        import traceback
        traceback.print_exc()
        return 1

    return _print_final(graph, config, doc_path)


def _stream(
    graph,
    initial,
    first_input,
    config,
    bridge: UserInputBridge,
    *,
    command=None,
) -> str:
    """stream 一段。返回结束原因: ok | force | pause。"""
    if command is not None:
        stream_input = command
    elif initial is None and first_input is None:
        stream_input = None
    else:
        stream_input = {**(initial or {}), "messages": first_input or []}

    bridge.set_busy(True)
    reason = "ok"
    try:
        for chunk in graph.stream(stream_input, config=config, stream_mode="updates"):
            for node, updates in chunk.items():
                if not isinstance(updates, dict):
                    continue
                new_msgs = updates.get("messages") or []
                for m in new_msgs:
                    if isinstance(m, AIMessage):
                        _print_agent_step(m)
                    elif isinstance(m, ToolMessage):
                        _print_tool_results(m)
                    elif isinstance(m, HumanMessage):
                        preview = str(m.content)[:80]
                        print(f"  {_YELLOW}📝 {preview}{_RESET}")

            # 仅在可安全注入时打断，避免拆开 tool_calls ↔ ToolMessage
            if bridge.peek_soft_pause() or bridge.has_force():
                if not _safe_to_inject(graph, config):
                    continue
                if bridge.peek_soft_pause():
                    reason = "pause"
                    print(f"\n{_YELLOW}收到中断信号，当前步骤后暂停…{_RESET}")
                    break
                reason = "force"
                print(f"\n{_YELLOW}⚡ 收到强制推送，当前步骤后打断…{_RESET}")
                break
    except KeyboardInterrupt:
        bridge.request_soft_pause()
        reason = "pause"
        print(f"\n{_YELLOW}已暂停。{_RESET}")
    finally:
        bridge.set_busy(False)

    return reason


if __name__ == "__main__":
    raise SystemExit(run())
