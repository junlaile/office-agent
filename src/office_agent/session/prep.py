"""会话准备纯函数：路径推导、公文版头字段、模板 merge（无终端 I/O）。"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any

from office_agent.config import settings
from office_agent.domain.format import OfficeFormat, infer_doc_kind
from office_agent.domain.templates import (
    default_merge_data,
    is_upward,
    template_path,
)
from office_agent.log import get_logger
from office_agent.office.doc import DocTool
from office_agent.officecli import OfficeCLIError, merge_template

logger = get_logger(__name__)


def resolve_output_kind(
    requirement: str,
    doc_type: str | None = None,
) -> tuple[OfficeFormat | None, int]:
    """推断输出扩展名。

    返回 ``(kind, score)``：
      - 公文命中 → ``("docx", 高分)``
      - 关键词命中 → ``(kind, score>0)``
      - 无线索 → ``(None, 0)``，调用方应向用户询问类型
    """
    if doc_type:
        return "docx", 999
    kind, score = infer_doc_kind(requirement)
    if score == 0:
        return None, 0
    return kind, score


def build_doc_path(requirement: str, *, kind: OfficeFormat) -> str:
    """从需求文本与已确定的扩展名生成输出路径。"""
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>|]', "", requirement).strip()
    safe = re.sub(r"\s+", "_", safe)
    safe = safe[:30] or "document"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return str((settings.output_dir / f"{safe}_{stamp}.{kind}").resolve())


def official_header_fields(doc_type: str) -> list[dict[str, Any]]:
    """公文版头采集字段定义（与 CLI 表单一致，供 API 下发）。"""
    fields: list[dict[str, Any]] = [
        {
            "key": "org",
            "label": "发文机关",
            "required": True,
            "hint": "如：市公安局（写入版头红字与落款）",
        },
    ]
    if is_upward(doc_type):
        signer_label = "签发人（报告人）" if doc_type == "报告" else "签发人"
        fields.append(
            {
                "key": "signer",
                "label": signer_label,
                "required": True,
                "hint": "姓名，勿留空",
            }
        )
    fields.extend(
        [
            {
                "key": "doc_no",
                "label": "发文字号",
                "required": False,
                "hint": "如：X公发〔2026〕1号；不确定可跳过",
            },
            {
                "key": "date_cn",
                "label": "成文日期",
                "required": False,
                "hint": "如：二〇二六年三月三十一日；不确定可跳过",
            },
        ]
    )
    return fields


def merge_official_doc(
    doc_type: str,
    doc_path: str,
    user_header: dict[str, str] | None = None,
) -> tuple[str | None, str]:
    """从公文模板 merge 到 ``doc_path``，返回 ``(文种名, 模板正文)``。

    ``user_header`` 为用户已填版头；未填项用占位。失败返回 ``(None, "")``。
    """
    tmpl = template_path(doc_type)
    if not tmpl.exists():
        logger.warning("公文模板缺失，回退普通模式: %s", tmpl)
        return None, ""

    cleaned = {
        k: v.strip()
        for k, v in (user_header or {}).items()
        if isinstance(v, str) and v.strip()
    }
    merge_data = default_merge_data(doc_type, **cleaned)
    try:
        merge_template(str(tmpl), doc_path, merge_data)
    except OfficeCLIError as e:
        logger.warning("公文模板预填失败，回退普通模式: %s", e)
        return None, ""

    template_text = ""
    try:
        template_text = DocTool(doc_path).view_text()
    except OfficeCLIError as e:
        logger.warning("模板正文预读失败，回退到 LLM 自行 view_text: %s", e)

    return doc_type, template_text
