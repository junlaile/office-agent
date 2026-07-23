"""根据用户需求推断输出 Office 格式（docx / xlsx / pptx）。"""

from __future__ import annotations

from typing import Literal

OfficeFormat = Literal["docx", "xlsx", "pptx"]

_FORMAT_LABELS = {
    "docx": "Word",
    "xlsx": "Excel",
    "pptx": "PowerPoint",
}

# 关键词 → 扩展名。命中数越多置信度越高。
_XLSX_KEYWORDS = [
    "excel",
    "表格",
    "报表",
    "数据表",
    "工作表",
    "spreadsheet",
    "财务模型",
    "预算表",
    "销售数据",
    "库存表",
    "工资表",
]
_PPTX_KEYWORDS = [
    "ppt",
    "pptx",
    "幻灯片",
    "演示",
    "汇报",
    "演讲",
    "宣讲",
    "课件",
    "路演",
    "deck",
    "slides",
    "presentation",
    "powerpoint",
]
_DOCX_KEYWORDS = [
    "word",
    "doc",
    "文档",
    "报告",
    "说明",
    "方案",
    "总结",
    "周报",
    "月报",
    "通知",
    "规章",
    "制度",
    "文章",
    "论文",
]


def infer_doc_kind(requirement: str) -> tuple[OfficeFormat, int]:
    """从需求关键词推断文档类型。返回 (kind, 命中数)。

    kind ∈ {'docx','xlsx','pptx'}；命中数越高置信度越高（0 表示无明确线索）。
    平局时优先级 xlsx > pptx > docx（因为 docx 是默认值，能往后让）。
    """
    text = (requirement or "").lower()
    scores: dict[OfficeFormat, int] = {
        "xlsx": sum(1 for k in _XLSX_KEYWORDS if k in text),
        "pptx": sum(1 for k in _PPTX_KEYWORDS if k in text),
        "docx": sum(1 for k in _DOCX_KEYWORDS if k in text),
    }
    kind = max(
        scores,
        key=lambda k: (scores[k], {"xlsx": 2, "pptx": 1, "docx": 0}[k]),
    )
    return kind, scores[kind]


def format_label(fmt: OfficeFormat | str) -> str:
    return _FORMAT_LABELS.get(fmt, str(fmt))
