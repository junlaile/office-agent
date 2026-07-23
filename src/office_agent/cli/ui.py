"""终端 UI 辅助：需求读取、文档类型推断、工具调用展示、interrupt 交互。

从 main.py 下沉而来，让 cli.main 只保留 run() 编排。这些函数中：
    - 纯函数（_infer_doc_kind / _format_tool_call / _indent / _derive_doc_path /
      _prepare_official_doc）可独立单元测试。
    - UI 函数（_banner / _print_* / _collect_* / _handle_interrupt）
      涉及 input/print，测试需 capsys/monkeypatch。
    - 所有交互输入经 ``_readline``：有 UserInputBridge 时走 blocking_readline，
      避免与忙时读线程抢 stdin。
"""

from __future__ import annotations

import logging
import re
import sys
from datetime import datetime

from langchain_core.messages import AIMessage, ToolMessage
from langgraph.types import Command

from office_agent.cli.user_input import get_bridge
from office_agent.config import settings
from office_agent.domain.format import infer_doc_kind
from office_agent.domain.templates import (
    default_merge_data,
    template_path,
)
from office_agent.office.doc import DocTool
from office_agent.officecli import OfficeCLIError, merge_template, resolve_bin

logger = logging.getLogger("office_agent.cli.ui")


def _readline(prompt: str = "") -> str:
    """统一读一行：优先走 UserInputBridge，否则回退 builtin input。"""
    bridge = get_bridge()
    if bridge is not None:
        return bridge.blocking_readline(prompt)
    return input(prompt)

# ANSI 颜色（Windows 10+ 终端支持）
_CYAN = "[36m"
_YELLOW = "[33m"
_GREEN = "[32m"
_RED = "[31m"
_MAGENTA = "[35m"
_DIM = "[2m"
_BOLD = "[1m"
_RESET = "[0m"


def _banner() -> None:
    print(
        f"\n{_BOLD}{_CYAN}╔══════════════════════════════════════════════╗\n"
        f"║   Office Agent — 交互式 Office 文档生成器     ║\n"
        f"║   Word / Excel / PowerPoint · OfficeCLI       ║\n"
        f"╚══════════════════════════════════════════════╝{_RESET}\n"
    )


def _check_officecli() -> bool:
    try:
        bin_path = resolve_bin()
        logger.info("officecli: %s", bin_path)
        return True
    except OfficeCLIError as e:
        logger.error("officecli 解析失败: %s", e)
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
            req = _readline(f"{_CYAN}❯ {_RESET}").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{_YELLOW}已取消。{_RESET}")
            sys.exit(0)
        if req:
            return req
        print(f"{_YELLOW}需求不能为空，请重新输入。{_RESET}")


def _infer_doc_kind(requirement: str) -> tuple[str, int]:
    """从需求关键词推断文档类型（委托 domain.format）。"""
    return infer_doc_kind(requirement)


def _ask_doc_kind() -> str:
    """无法从需求推断文档类型时，交互问用户。"""
    print(f"{_BOLD}{_YELLOW}需要确认要生成的文档类型:{_RESET}")
    print(f"{_DIM}  1. Word 文档（报告/说明/方案）{_RESET}")
    print(f"{_DIM}  2. Excel 表格（数据/报表）{_RESET}")
    print(f"{_DIM}  3. PowerPoint 演示（汇报/演讲）{_RESET}")
    while True:
        try:
            ans = _readline(f"{_CYAN}选 [1-3] ❯ {_RESET}").strip()
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


def _derive_doc_path(requirement: str, doc_type: str | None = None) -> str:
    """从需求生成输出文档路径（含扩展名推断）。

    doc_type 非空表示已识别为公文 → 强制 .docx 扩展名，跳过类型询问。
    """
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    if doc_type:
        # 公文模式：一定是 Word
        kind = "docx"
    else:
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


def _prepare_official_doc(doc_type: str, doc_path: str) -> tuple[str | None, str]:
    """从公文模板预创建文档并预填版头槽位，返回 (文种名, 模板正文)。

    失败时返回 (None, "")，由调用方回退到普通生成流程。

    调用前提: doc_type 已由 detect_doc_type 识别为合法文种。
    做的事:
        1. 定位模板文件 template/word/NN-{doc_type}.docx。
        2. merge 模板到 doc_path（一步完成复制 + 版头槽位预填）。
        3. 立即读一次 view_text，拿到带路径标注的模板正文，回传给
           build_system_prompt 注入提示词——让 LLM 第一轮就"看到"段落结构，
           不必依赖它自己调 view_text（核心：解决"没读模板就瞎改"问题）。
        4. 打印预创建结果。

    template_text 读取失败（officecli 异常等）不阻断主流程：退化为空串，
    提示词回退到软指令"先 view_text"，LLM 仍可自行读取兜底。
    """
    tmpl = template_path(doc_type)
    if not tmpl.exists():
        logger.warning("公文模板缺失，回退普通模式: %s", tmpl)
        return None, ""

    # 预填数据：用户没提供的版头槽位用默认占位值
    merge_data = default_merge_data(doc_type)
    try:
        merge_template(str(tmpl), doc_path, merge_data)
    except OfficeCLIError as e:
        logger.warning("公文模板预填失败，回退普通模式: %s", e)
        return None, ""

    # 预读模板正文（带路径标注），注入提示词，避免 LLM 跳过读模板
    template_text = ""
    try:
        template_text = DocTool(doc_path).view_text()
    except OfficeCLIError as e:
        # 预读失败不致命：退化为空，LLM 自调 view_text 兜底
        logger.warning("模板正文预读失败，回退到 LLM 自行 view_text: %s", e)

    print(f"{_GREEN}✓ 已从 GB/T 9704 模板创建{_RESET}")
    print(f"{_DIM}  模板: {tmpl.name} | 版头槽位已预填占位值，正文待 agent 编辑{_RESET}")
    return doc_type, template_text


def _print_outline_preview(outline: str) -> None:
    """终端展示 Markdown 大纲预览。"""
    print(f"\n{_BOLD}{_CYAN}┌─ 文档大纲预览 ", end="")
    pad = max(0, 44 - 10)
    print("─" * pad + f"┐{_RESET}")
    for line in (outline or "").splitlines() or ["（空大纲）"]:
        print(f"{_CYAN}│{_RESET} {line}")
    print(f"{_CYAN}└" + "─" * 52 + f"┘{_RESET}")


def _collect_outline_decision() -> tuple[str, str]:
    """收集用户对大纲的决定。

    返回 ``(action, feedback)``，其中 action ∈ approve|revise|cancel；
    仅 revise 时 feedback 非空（可能仍为空串若用户回车）。
    """
    print(f"\n{_BOLD}请确认大纲:{_RESET}")
    print(f"{_DIM}  1. 批准生成 Word{_RESET}")
    print(f"{_DIM}  2. 提出修改意见（修订大纲后再预览）{_RESET}")
    print(f"{_DIM}  3. 取消{_RESET}")
    while True:
        try:
            ans = _readline(f"{_CYAN}选 [1-3] ❯ {_RESET}").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{_YELLOW}已取消。{_RESET}")
            return "cancel", ""
        low = ans.lower()
        if ans in ("1",) or low in ("批准", "approve", "y", "yes", "ok"):
            return "approve", ""
        if ans in ("2",) or low in ("修改", "revise", "改"):
            try:
                fb = _readline(f"{_CYAN}修改意见 ❯ {_RESET}").strip()
            except (EOFError, KeyboardInterrupt):
                print(f"\n{_YELLOW}已取消。{_RESET}")
                return "cancel", ""
            return "revise", fb
        if ans in ("3",) or low in ("取消", "cancel", "n", "no", "q"):
            return "cancel", ""
        print(f"{_YELLOW}请输入 1/2/3 或 批准/修改/取消。{_RESET}")


def _run_outline_approval_loop(
    requirement: str,
    doc_type: str | None = None,
) -> str | None:
    """生成并展示大纲，循环至用户批准或取消。

    返回批准后的大纲文本；取消返回 None。
    """
    from office_agent.agent.outline import generate_outline

    feedback = ""
    while True:
        print(f"\n{_DIM}正在生成文档大纲预览…{_RESET}")
        try:
            outline = generate_outline(
                requirement, feedback=feedback, doc_type=doc_type
            )
        except Exception as e:  # noqa: BLE001
            logger.exception("大纲生成失败")
            print(f"{_RED}大纲生成失败: {e}{_RESET}")
            print(f"{_DIM}可输入修改意见重试，或取消。{_RESET}")
            outline = f"（生成失败）{e}"
        _print_outline_preview(outline)
        action, feedback = _collect_outline_decision()
        if action == "approve":
            if outline.startswith("（生成失败）"):
                print(f"{_YELLOW}大纲无效，请选择修改重试或取消。{_RESET}")
                continue
            print(f"{_GREEN}✓ 大纲已批准，开始生成 Word…{_RESET}")
            return outline
        if action == "cancel":
            print(f"{_YELLOW}已取消，未生成文档。{_RESET}")
            return None
        # revise: feedback 已填，继续循环
        if not feedback:
            print(f"{_YELLOW}未填写修改意见，将按原需求重新生成。{_RESET}")


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
        return (
            f'set_cell("{args.get("sheet", "")}", {args.get("ref", "")}, {args.get("value", "")!r})'
        )
    if name == "set_cells":
        data = args.get("data") or []
        rows = len(data)
        cols = max((len(r) for r in data), default=0)
        return (
            f'set_cells("{args.get("sheet", "")}", {rows}×{cols}, '
            f'start="{args.get("start", "A1")}")'
        )
    if name == "set_formula":
        return (
            f'set_formula("{args.get("sheet", "")}", '
            f"{args.get('ref', '')}, ={args.get('formula', '')})"
        )
    if name == "add_excel_chart":
        return (
            f'add_excel_chart("{args.get("sheet", "")}", '
            f"{args.get('chart_type', '')}, {args.get('data_range', '')})"
        )
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
    if name == "start_from_template":
        dt = args.get("doc_type", "")
        org = args.get("org", "")
        extra = f', org="{org}"' if org else ""
        return f'start_from_template("{dt}"{extra})'
    if name == "update_paragraph":
        text = args.get("text", "")
        preview = text[:25] + ("…" if len(text) > 25 else "")
        return f'update_paragraph("{args.get("path", "")}", "{preview}")'
    if name == "replace_text":
        return f'replace_text("{args.get("find", "")}" → "{args.get("replace", "")}")'
    if name == "remove_paragraph":
        return f'remove_paragraph("{args.get("path", "")}")'
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
        elif name == "start_from_template":
            icon = "📋"
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
    resume_value: dict[str, str] | str
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
            ans = _readline(f"{_CYAN}答 ❯ {_RESET}").strip()
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
    print(
        f"{_YELLOW}│{_RESET} {_DIM}带 * 为必填；有候选的可输序号或自由输入；可选字段可回车跳过{_RESET}"
    )
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
            opts_str = "  ".join(f"{_CYAN}[{n}]{_RESET} {o}" for n, o in enumerate(options, 1))
            print(f"{_YELLOW}│{_RESET}   {_DIM}候选: {_RESET}{opts_str}")

        # 逐字段输入循环（必填校验）
        while True:
            try:
                val = _readline(f"{_YELLOW}│{_RESET} {_CYAN}❯ {_RESET}").strip()
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
