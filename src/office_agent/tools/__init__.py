"""暴露给 LLM 的工具集（@tool 装饰器）包。

设计:
    - 会话级文档路径注入：main.py 启动 agent 前调用 set_session_doc(path)，
      所有工具内部读取该路径，LLM 不需要传路径参数（避免出错）。
    - 扩展名路由：路径后缀决定文档类型（.docx / .xlsx / .pptx），
      _tool() 工厂按 kind 返回 DocTool / ExcelTool / PptxTool。
    - 通用工具（create_doc / add_table / view_text / validate_doc / finish）
      三种格式都支持；docx/xlsx/pptx 专属工具在其他格式下给出明确提示。
    - ask_user 工具内部用 LangGraph interrupt 挂起，等用户输入后作为
      ToolMessage 回传给 agent。
    - finish 工具让 LLM 显式宣告完成。

结构:
    - 本 ``__init__.py``: 会话状态基础设施 + 聚合 ALL_TOOLS/TOOL_BY_NAME。
    - ``common.py``: 通用工具（三格式共用）+ 控制（ask_user/finish）+ 公文/查询。
    - ``doc.py``:    Word 专属工具。
    - ``excel.py``:  Excel 专属工具。
    - ``pptx.py``:   PowerPoint 专属工具。

子模块通过 ``from office_agent.tools import _tool, session_doc_kind, ...``
反向引用本包的会话基础设施——因此这些符号必须在 ``__init__.py`` 定义，
且子模块的 import 在 ``__init__.py`` 底部执行（此时基础设施已就绪）。
"""

from __future__ import annotations

import logging

from office_agent.officecli import DocTool, ExcelTool, OfficeCLIError, PptxTool

logger = logging.getLogger(__name__)

# ============================================================
# 会话状态（模块级，agent 运行期间唯一）
# ============================================================
_session_doc_path: str | None = None


def set_session_doc(path: str | None) -> None:
    """main.py 在启动 agent 前调用，设定本会话的文档路径。

    传 None 清空（供测试 teardown）。
    """
    global _session_doc_path
    _session_doc_path = path


def session_doc_path() -> str | None:
    return _session_doc_path


def session_doc_kind() -> str:
    """返回当前会话的文档类型: 'docx' | 'xlsx' | 'pptx'。

    路径未初始化时抛错；扩展名无法识别时默认 'docx'。
    """
    if not _session_doc_path:
        raise OfficeCLIError("会话文档路径未初始化（请先调用 set_session_doc）")
    p = _session_doc_path.lower()
    if p.endswith(".xlsx"):
        return "xlsx"
    if p.endswith(".pptx"):
        return "pptx"
    return "docx"


def _tool():
    """工厂：按当前会话扩展名返回对应的 Tool 实例。"""
    if not _session_doc_path:
        raise OfficeCLIError("会话文档路径未初始化（请先调用 set_session_doc）")
    kind = session_doc_kind()
    if kind == "xlsx":
        return ExcelTool(_session_doc_path)
    if kind == "pptx":
        return PptxTool(_session_doc_path)
    return DocTool(_session_doc_path)


def _doc() -> DocTool:
    """向后兼容：旧代码可能引用 _doc()。"""
    return _tool()  # type: ignore[return-value]


def _wrong_kind_msg(tool_name: str, expected: str, hint: str = "") -> str:
    """当前会话格式与工具不符时的友好提示（而非 AttributeError 崩溃）。"""
    actual = session_doc_kind()
    logger.warning("工具与文档类型不匹配: %s 需要 %s，当前 %s", tool_name, expected, actual)
    hint_str = f" {hint}" if hint else ""
    return (f"{tool_name} 是 {expected.upper()} 专属工具，当前文档是 {actual}。{hint_str}").strip()


# ============================================================
# 从子模块 import 所有 @tool 工具，聚合 ALL_TOOLS
# ============================================================
# 必须放在会话基础设施定义【之后】：子模块 import 本包时要用到上面的符号。
from .common import (  # noqa: E402
    add_image,
    add_table,
    ask_user,
    create_doc,
    finish,
    query_vehicle,
    set_doc_properties,
    start_from_template,
    validate_doc,
    view_text,
)
from .doc import (  # noqa: E402
    add_footer,
    add_header,
    add_heading,
    add_hyperlink,
    add_list_item,
    add_page_number,
    add_paragraph,
    add_section_break,
    add_title,
    add_toc,
    add_word_chart,
    remove_paragraph,
    replace_text,
    update_paragraph,
)
from .excel import (  # noqa: E402
    add_color_scale,
    add_data_bar,
    add_dropdown,
    add_excel_chart,
    add_list_table,
    add_pivot_table,
    add_sheet,
    autofit_column,
    highlight_cells,
    merge_cells,
    rename_sheet,
    set_autofilter,
    set_cell,
    set_cells,
    set_column_width,
    set_formula,
    sort_sheet,
)
from .pptx import (  # noqa: E402
    add_slide,
    add_slide_image,
    add_slide_table,
    add_textbox,
    set_slide_notes,
    set_slide_transition,
    set_theme_colors,
    set_theme_fonts,
)

# ============================================================
# 工具清单（按文档类型绑定；ALL_TOOLS 保留作注册表）
# ============================================================
COMMON_TOOLS = [
    # 三种格式均支持
    create_doc,
    add_table,
    view_text,
    validate_doc,
    set_doc_properties,
]

IMAGE_TOOLS = [add_image]

WORD_TOOLS = [
    start_from_template,
    add_title,
    add_heading,
    add_paragraph,
    add_list_item,
    add_toc,
    add_page_number,
    add_header,
    add_footer,
    add_hyperlink,
    add_word_chart,
    add_section_break,
    # Word 编辑（改/删/替换）——公文模式主力
    update_paragraph,
    replace_text,
    remove_paragraph,
]

EXCEL_TOOLS = [
    add_sheet,
    set_cell,
    set_cells,
    set_formula,
    add_excel_chart,
    sort_sheet,
    set_autofilter,
    highlight_cells,
    add_color_scale,
    add_data_bar,
    add_pivot_table,
    add_list_table,
    add_dropdown,
    merge_cells,
    set_column_width,
    autofit_column,
    rename_sheet,
]

PPTX_TOOLS = [
    add_slide,
    add_textbox,
    add_slide_image,
    add_slide_table,
    set_slide_transition,
    set_slide_notes,
    set_theme_colors,
    set_theme_fonts,
]

CONTROL_TOOLS = [
    # 业务专项 + 控制（三种格式均可使用）
    query_vehicle,
    ask_user,
    finish,
]


def tools_for_doc_path(doc_path: str) -> list:
    """按输出文件扩展名返回应暴露给 LLM 的工具。

    未知扩展名与会话路由保持一致，默认按 Word 处理。返回新列表，避免调用方
    修改模块级工具集合。
    """
    path = doc_path.lower()
    if path.endswith(".xlsx"):
        specific = EXCEL_TOOLS
        image_tools = []
    elif path.endswith(".pptx"):
        specific = PPTX_TOOLS
        image_tools = IMAGE_TOOLS
    else:
        specific = WORD_TOOLS
        image_tools = IMAGE_TOOLS
    return [*COMMON_TOOLS, *image_tools, *specific, *CONTROL_TOOLS]


# 全量注册表仅供工具执行节点按名称分发，以及兼容既有导入。
ALL_TOOLS = [
    *COMMON_TOOLS,
    *IMAGE_TOOLS,
    *WORD_TOOLS,
    *EXCEL_TOOLS,
    *PPTX_TOOLS,
    *CONTROL_TOOLS,
]

# 工具名 -> 工具对象，便于 tools 节点按名分发
TOOL_BY_NAME = {t.name: t for t in ALL_TOOLS}

__all__ = [
    # 会话基础设施
    "set_session_doc",
    "session_doc_path",
    "session_doc_kind",
    "_tool",
    "_doc",
    "_wrong_kind_msg",
    "_session_doc_path",
    # 聚合
    "ALL_TOOLS",
    "TOOL_BY_NAME",
    "COMMON_TOOLS",
    "IMAGE_TOOLS",
    "WORD_TOOLS",
    "EXCEL_TOOLS",
    "PPTX_TOOLS",
    "CONTROL_TOOLS",
    "tools_for_doc_path",
    # 所有工具（供 from office_agent.tools import X）
    *[t.name for t in ALL_TOOLS],
]
