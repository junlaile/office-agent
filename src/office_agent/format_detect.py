"""根据用户需求自动识别输出 Office 格式。"""

from __future__ import annotations

import re
from typing import Literal

OfficeFormat = Literal["docx", "xlsx", "pptx"]

_FORMAT_LABELS = {
    "docx": "Word",
    "xlsx": "Excel",
    "pptx": "PowerPoint",
}

# pptx 优先于 xlsx（「汇报 PPT」不应落到报表）
_PPTX_RE = re.compile(
    r"ppt|pptx|powerpoint|幻灯|演示文稿|汇报片|\bslides?\b",
    re.IGNORECASE,
)
_XLSX_RE = re.compile(
    r"excel|xlsx|工作簿|工作表|台账|报表|数据表|spreadsheet",
    re.IGNORECASE,
)


def detect_format(requirement: str) -> OfficeFormat:
    """从自然语言需求推断输出格式。默认 docx。

    仅出现「表格」不足以判为 Excel（Word 也常用表格）。
    """
    text = (requirement or "").strip()
    if not text:
        return "docx"
    if _PPTX_RE.search(text):
        return "pptx"
    if _XLSX_RE.search(text):
        return "xlsx"
    return "docx"


def format_label(fmt: OfficeFormat) -> str:
    return _FORMAT_LABELS.get(fmt, fmt)
