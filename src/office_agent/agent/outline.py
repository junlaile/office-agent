"""Word 文档大纲生成（预览批准门控用）。

在正式 create_doc / 公文模板落盘之前，用无工具绑定的 LLM 生成结构化
Markdown 大纲，供用户批准或提出修改意见。
"""

from __future__ import annotations

import logging
import re

from langchain_core.messages import HumanMessage, SystemMessage

from .llm import get_llm

logger = logging.getLogger(__name__)

_OUTLINE_SYSTEM = """\
你是文档结构规划助手。根据用户需求输出【结构化 Markdown 大纲】，供用户预览批准。

硬性约束:
1. 只输出 Markdown 大纲，不要寒暄、不要解释、不要用代码围栏。
2. 结构：一级标题用 #，章节用 ## / ###，每节下列 2～6 个要点（- 开头）。
3. 这是大纲不是成稿：每条要点一句话概括写什么，不要写成完整段落。
4. 用户未提供的关键事实（人名、日期、机关名、数据）标为「待确认」，禁止臆造。
5. 总长度控制在约 400～800 汉字（要点合计），宁短勿长。
"""

_OFFICIAL_HINT = """\
本次为法定公文《{doc_type}》。大纲按该文种常见结构组织，例如：
- # 标题（事由概括）
- ## 主送机关（可写「待确认」）
- ## 正文要点（分条）
- ## 结语 / 落款说明
不要写版头假数据（发文字号、印发日期等留给模板）。
"""


def generate_outline(
    requirement: str,
    *,
    feedback: str = "",
    doc_type: str | None = None,
) -> str:
    """根据需求（及可选修改意见）生成 Markdown 大纲。

    不 bind tools；失败时抛出底层异常由调用方处理。
    """
    req = (requirement or "").strip()
    if not req:
        raise ValueError("requirement 不能为空")

    parts = [_OUTLINE_SYSTEM]
    if doc_type:
        parts.append(_OFFICIAL_HINT.format(doc_type=doc_type))
    system = "\n".join(parts)

    user_lines = [f"用户需求：\n{req}"]
    fb = (feedback or "").strip()
    if fb:
        user_lines.append(f"用户对上一版大纲的修改意见：\n{fb}\n请据此修订大纲。")
    user_lines.append("请直接输出修订后的 Markdown 大纲。")

    llm = get_llm()
    resp = llm.invoke(
        [
            SystemMessage(content=system),
            HumanMessage(content="\n\n".join(user_lines)),
        ]
    )
    text = _normalize_outline(str(getattr(resp, "content", "") or ""))
    if not text:
        raise RuntimeError("大纲生成结果为空")
    logger.info("已生成大纲，约 %d 字", len(text))
    return text


def _normalize_outline(raw: str) -> str:
    """去掉模型偶发的代码围栏与首尾空白。"""
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:markdown|md)?\s*", "", text, count=1, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    return text.strip()
