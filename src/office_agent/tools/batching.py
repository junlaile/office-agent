"""同批工具调用合并：连续的"末尾追加"类写调用 → 一次 officecli batch。

提示词鼓励 LLM 在一次回复里并行发多个工具调用（add_paragraph × N /
add_slide × N）。逐个执行时每个调用一个 subprocess，且 officecli 每条命令
都完整 open/save 文档——N 段就是 N 次全量读写。本模块把可安全合并的连续
调用翻译成一次 ``batch`` 命令（单次 open/save、原子回滚），失败时由调用方
回退逐个执行，外部行为不变。

只合并【无顺序依赖的末尾追加】操作：
    - docx: add_title / add_heading / add_paragraph / add_list_item
      （都 append 到 /body，合并后顺序与逐个执行一致；props 定义复用
      DocTool 的 *_props 构造器，与单发路径同源）
    - pptx: add_slide（append 到 /，同理；props 复用 PptxTool.slide_props）
不合并的：
    - 编辑/删除类（update/replace/remove）——依赖段落索引，语义随执行顺序变化；
    - 需要读回结果做后续定位的（add_table 建表后查索引、add_image 建载段）；
    - Excel 的 set_cells 本身已是 batch，无需二次合并。
"""

from __future__ import annotations

import json
import logging
from typing import Any

from office_agent.office.doc import DocTool
from office_agent.office.pptx import PptxTool
from office_agent.office.runner import OfficeCLIError, get_runner

from .session import session_doc_kind, session_doc_path

logger = logging.getLogger(__name__)

# 最近一次 batch 失败原因；graph 回退后读取并标注到 ToolMessage，随后清空。
_LAST_BATCH_FALLBACK_REASON: str | None = None

BATCH_FALLBACK_PREFIX = "[batch-fallback]"


def take_batch_fallback_reason() -> str | None:
    """取出并清空最近一次 batch 失败原因（供 graph 回退标注用）。"""
    global _LAST_BATCH_FALLBACK_REASON
    reason = _LAST_BATCH_FALLBACK_REASON
    _LAST_BATCH_FALLBACK_REASON = None
    return reason


def _set_batch_fallback_reason(reason: str) -> None:
    global _LAST_BATCH_FALLBACK_REASON
    _LAST_BATCH_FALLBACK_REASON = reason


# 与 tools/pptx.py add_slide 的"只有标题没正文"警告保持一致的口径
_SLIDE_NO_BODY_WARNING = (
    "⚠️ 本页只有标题没有正文（body_text 为空）。"
    "若这是封面/章节分隔页可以；若是【内容页】，请重新调用本工具并"
    "【补上 body_text】写入要点内容——否则这页会是空的。\n"
)

# 兜底的成功文案（batch 输出解析不到时用；正常路径用 officecli 的逐条 output）
_GENERIC_OK = {
    "add_title": "已添加主标题",
    "add_heading": "已添加标题",
    "add_paragraph": "已添加段落",
    "add_list_item": "已添加列表项",
    "add_slide": "已添加幻灯片",
}


def _stringify(props: dict) -> dict[str, str]:
    return {k: str(v) for k, v in props.items() if v is not None}


def _op_from_call(name: str, args: dict, kind: str) -> dict | None:
    """把一个工具调用翻译成 batch op；不可合并返回 None。"""
    if kind == "docx":
        if name == "add_title":
            props = DocTool.title_props(str(args.get("text", "")))
        elif name == "add_heading":
            props = DocTool.heading_props(str(args.get("text", "")), int(args.get("level") or 1))
        elif name == "add_paragraph":
            props = DocTool.paragraph_props(
                str(args.get("text", "")),
                bold=bool(args.get("bold", False)),
                italic=bool(args.get("italic", False)),
            )
        elif name == "add_list_item":
            props = DocTool.list_item_props(
                str(args.get("text", "")), ordered=bool(args.get("ordered", False))
            )
        else:
            return None
        return {"command": "add", "path": "/body", "type": "paragraph", "props": _stringify(props)}

    if kind == "pptx" and name == "add_slide":
        props = PptxTool.slide_props(
            title=str(args.get("title") or ""),
            text=str(args.get("body_text") or ""),
            layout=str(args.get("layout") or ""),
        )
        return {"command": "add", "path": "/", "type": "slide", "props": _stringify(props)}

    return None


def is_batchable(tool_call: dict) -> bool:
    """判断一个 tool_call 是否可参与 batch 合并（会话未初始化时一律 False）。"""
    try:
        kind = session_doc_kind()
        op = _op_from_call(tool_call.get("name") or "", tool_call.get("args") or {}, kind)
    except Exception:  # noqa: BLE001  # 参数异常等 → 走逐个执行，让原路径报错
        return False
    return op is not None


def _parse_batch_outputs(raw: Any) -> dict[int, str]:
    """从 batch 的 JSON 输出里取逐 op 的 output 文案（与逐个执行的返回一致）。"""
    try:
        payload = json.loads(raw) if isinstance(raw, str) else raw
        results = payload.get("data", {}).get("results", [])
        return {
            int(r["index"]): str(r.get("output") or "")
            for r in results
            if isinstance(r, dict) and r.get("index") is not None
        }
    except Exception:  # noqa: BLE001
        return {}


def execute_batched(tool_calls: list[dict]) -> list[tuple[str, str]] | None:
    """把一段连续可合并的 tool_calls 打成一次 batch 执行。

    成功返回 ``[(tool_call_id, content), ...]``（与 tool_calls 一一对应，
    content 是给 LLM 的 ToolMessage 文案）；不可合并或 batch 失败返回 None，
    由调用方回退逐个执行（batch 原子回滚，不会残留半截内容）。
    """
    # 每次尝试前清空，避免上一次未消费的原因污染本次成功路径
    global _LAST_BATCH_FALLBACK_REASON
    _LAST_BATCH_FALLBACK_REASON = None

    if len(tool_calls) < 2:
        return None
    try:
        kind = session_doc_kind()
        doc_path = session_doc_path()
    except OfficeCLIError:
        return None
    if not doc_path:
        return None

    ops: list[dict] = []
    for tc in tool_calls:
        try:
            op = _op_from_call(tc.get("name") or "", tc.get("args") or {}, kind)
        except Exception:  # noqa: BLE001
            op = None
        if op is None:
            return None
        ops.append(op)

    payload = json.dumps(ops, ensure_ascii=False)
    try:
        raw = get_runner().run(["batch", doc_path, "--commands", payload, "--json"])
    except OfficeCLIError as e:
        # 保留回退能力，但把失败原因留给 graph 标注到 ToolMessage，
        # 避免"最终成功"掩盖中间 batch 失败。
        _set_batch_fallback_reason(str(e))
        logger.warning(
            "batch 合并执行失败（%d 个 %s 写调用），将回退逐个执行: %s",
            len(ops),
            kind,
            e,
        )
        return None

    outputs = _parse_batch_outputs(raw)
    results: list[tuple[str, str]] = []
    for i, tc in enumerate(tool_calls):
        name = tc.get("name") or ""
        args = tc.get("args") or {}
        content = outputs.get(i) or _GENERIC_OK.get(name, f"已执行 {name}")
        # 与 tools/pptx.add_slide 的"标题无正文"警告保持一致
        if (
            name == "add_slide"
            and args.get("title")
            and not str(args.get("body_text") or "").strip()
        ):
            content = _SLIDE_NO_BODY_WARNING + content
        results.append((tc.get("id", ""), content))

    logger.debug("已合并执行 %d 个 %s 写调用（单次 batch）", len(ops), kind)
    return results
