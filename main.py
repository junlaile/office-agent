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

from langchain_core.messages import AIMessage, ToolMessage
from langgraph.types import Command

from office_agent.config import assert_llm_ready, settings
from office_agent.officecli import OfficeCLIError, resolve_bin
from office_agent.tools import set_session_doc

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
    print(f"\n{_BOLD}{_CYAN}╔══════════════════════════════════════════════╗\n"
          f"║   Office Agent — 交互式 Office 文档生成器     ║\n"
          f"║   Word / Excel / PowerPoint · OfficeCLI       ║\n"
          f"╚══════════════════════════════════════════════╝{_RESET}\n")


def _check_officecli() -> bool:
    try:
        bin_path = resolve_bin()
        print(f"{_DIM}[ok] officecli: {bin_path}{_RESET}")
        return True
    except OfficeCLIError as e:
        print(f"{_RED}[error] {e}{_RESET}")
        print(f"{_YELLOW}请先运行: python scripts/fetch_officecli.py{_RESET}")
        return False


def _read_requirement() -> str:
    if len(sys.argv) > 1:
        req = " ".join(sys.argv[1:]).strip()
        if req:
            print(f"{_DIM}需求: {req}{_RESET}\n")
            return req
    print(f"{_BOLD}请描述你要生成的 Office 文档（回车提交）:{_RESET}")
    print(f"{_DIM}支持 Word / Excel / PowerPoint。例如:{_RESET}")
    print(f"{_DIM}  · 写一份项目周报，包含本周进展、风险、下周计划{_RESET}")
    print(f"{_DIM}  · 做一份季度销售数据的 Excel 表格，含图表{_RESET}")
    print(f"{_DIM}  · 做一个 10 页的产品介绍 PPT{_RESET}\n")
    while True:
        try:
            req = input(f"{_CYAN}❯ {_RESET}").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{_YELLOW}已取消。{_RESET}")
            sys.exit(0)
        if req:
            return req
        print(f"{_YELLOW}需求不能为空，请重新输入。{_RESET}")


# 文档类型推断：关键词 → 扩展名。命中数越多置信度越高。
_XLSX_KEYWORDS = ["excel", "表格", "报表", "数据表", "工作表", "spreadsheet",
                  "财务模型", "预算表", "销售数据", "库存表", "工资表"]
_PPTX_KEYWORDS = ["ppt", "pptx", "幻灯片", "演示", "汇报", "演讲", "宣讲",
                  "课件", "路演", "deck", "slides", "presentation", "powerpoint"]
_DOCX_KEYWORDS = ["word", "doc", "文档", "报告", "说明", "方案", "总结",
                  "周报", "月报", "通知", "规章", "制度", "文章", "论文"]


def _infer_doc_kind(requirement: str) -> tuple[str, int]:
    """从需求关键词推断文档类型。返回 (kind, 命中数)。

    kind ∈ {'docx','xlsx','pptx'}；命中数越高置信度越高（0 表示无明确线索）。
    平局时优先级 xlsx > pptx > docx（因为 docx 是默认值，能往后让）。
    """
    text = requirement.lower()
    scores = {
        "xlsx": sum(1 for k in _XLSX_KEYWORDS if k in text),
        "pptx": sum(1 for k in _PPTX_KEYWORDS if k in text),
        "docx": sum(1 for k in _DOCX_KEYWORDS if k in text),
    }
    kind = max(scores, key=lambda k: (scores[k], {"xlsx": 2, "pptx": 1, "docx": 0}[k]))
    return kind, scores[kind]


def _ask_doc_kind() -> str:
    """无法从需求推断文档类型时，交互问用户。"""
    print(f"{_BOLD}{_YELLOW}需要确认要生成的文档类型:{_RESET}")
    print(f"{_DIM}  1. Word 文档（报告/说明/方案）{_RESET}")
    print(f"{_DIM}  2. Excel 表格（数据/报表）{_RESET}")
    print(f"{_DIM}  3. PowerPoint 演示（汇报/演讲）{_RESET}")
    while True:
        try:
            ans = input(f"{_CYAN}选 [1-3] ❯ {_RESET}").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{_YELLOW}已取消。{_RESET}")
            sys.exit(0)
        if ans in ("1", "word", "w"):
            return "docx"
        if ans in ("2", "excel", "e"):
            return "xlsx"
        if ans in ("3", "ppt", "pptx", "p"):
            return "pptx"
        print(f"{_YELLOW}请输入 1/2/3 或 word/excel/ppt。{_RESET}")


def _derive_doc_path(requirement: str) -> str:
    """从需求生成输出文档路径（含扩展名推断）。"""
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    kind, score = _infer_doc_kind(requirement)
    if score == 0:
        # 没有明确线索，问用户
        kind = _ask_doc_kind()
    # 提取需求里的中文/字母数字作文件名
    safe = re.sub(r'[\\/:*?"<>|]', "", requirement).strip()
    safe = re.sub(r"\s+", "_", safe)
    # 截断并兜底
    safe = safe[:30] or "document"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return str((settings.output_dir / f"{safe}_{stamp}.{kind}").resolve())


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
        extra = ", bold" if bold else ""
        text = args.get("text", "")
        preview = text[:30] + ("…" if len(text) > 30 else "")
        return f'add_paragraph("{preview}"{extra})'
    if name == "add_list_item":
        ordered = args.get("ordered", False)
        return f'add_list_item("{args.get("text", "")}"{", ordered" if ordered else ""})'
    if name == "add_table" or name == "add_slide_table":
        data = args.get("data") or []
        rows = len(data)
        cols = max((len(r) for r in data), default=0)
        return f"{name}({rows}×{cols})"
    if name == "add_image":
        cap = args.get("caption", "")
        cap_str = f', caption="{cap}"' if cap else ""
        src = str(args.get("url_or_path", ""))
        # 截断长 URL
        src_short = src if len(src) <= 30 else src[:27] + "..."
        return f'add_image("{src_short}"{cap_str})'
    if name == "add_sheet":
        return f'add_sheet("{args.get("name", "")}")'
    if name == "set_cell":
        return (f'set_cell("{args.get("sheet", "")}", '
                f'{args.get("ref", "")}, {args.get("value", "")!r})')
    if name == "set_cells":
        data = args.get("data") or []
        rows = len(data)
        cols = max((len(r) for r in data), default=0)
        return (f'set_cells("{args.get("sheet", "")}", {rows}×{cols}, '
                f'start="{args.get("start", "A1")}")')
    if name == "set_formula":
        return (f'set_formula("{args.get("sheet", "")}", '
                f'{args.get("ref", "")}, ={args.get("formula", "")})')
    if name == "add_excel_chart":
        return (f'add_excel_chart("{args.get("sheet", "")}", '
                f'{args.get("chart_type", "")}, {args.get("data_range", "")})')
    if name == "add_slide":
        title = args.get("title", "")
        body = args.get("body_text", "") or ""
        # 统计正文行数（便于看出 LLM 是否真的写了正文）
        body_lines = [l for l in body.split("\n") if l.strip()]
        body_hint = f", 正文{len(body_lines)}行" if body_lines else ", ⚠️无正文"
        return f'add_slide(title="{title}"{body_hint})'
    if name == "add_textbox":
        text = args.get("text", "")
        preview = text[:25] + ("…" if len(text) > 25 else "")
        return f'add_textbox("{preview}")'
    if name == "add_slide_image":
        src = str(args.get("url_or_path", ""))
        src_short = src if len(src) <= 30 else src[:27] + "..."
        return f'add_slide_image("{src_short}")'
    if name == "query_vehicle":
        return f'query_vehicle("{args.get("plate_number", "")}")'
    if name == "ask_user":
        # 表单模式显示字段数，单问题显示问题
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


def _handle_interrupt(graph, config: dict) -> Command | None:
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

    # 分支：表单卡片 vs 单问题
    if isinstance(payload, dict) and payload.get("fields"):
        resume_value = _collect_form(payload)
    else:
        resume_value = _collect_single_question(payload)

    print(f"\n{_DIM}已收到，继续生成…{_RESET}")
    return Command(resume=resume_value)


def _collect_single_question(payload: dict) -> str:
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
            ans = input(f"{_CYAN}答 ❯ {_RESET}").strip()
        except (EOFError, KeyboardInterrupt):
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


def _collect_form(payload: dict) -> dict:
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
                val = input(f"{_YELLOW}│{_RESET} {_CYAN}❯ {_RESET}").strip()
            except (EOFError, KeyboardInterrupt):
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


def run() -> int:
    _banner()
    assert_llm_ready()
    if not _check_officecli():
        return 1

    requirement = _read_requirement()
    doc_path = _derive_doc_path(requirement)
    print(f"{_DIM}输出文件: {doc_path}{_RESET}\n")

    # 延迟 import（确保 officecli 已就绪）
    from office_agent.graph import build_graph

    set_session_doc(doc_path)
    graph = build_graph(doc_path)
    thread_id = str(uuid.uuid4())
    config = {
        "configurable": {"thread_id": thread_id},
        "recursion_limit": settings.recursion_limit,
    }

    initial: dict[str, Any] = {
        "messages": [],  # 第一条 HumanMessage 由下面加入
        "doc_path": doc_path,
    }
    from langchain_core.messages import HumanMessage
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
    from pathlib import Path as _P
    if _P(doc_path).exists():
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
