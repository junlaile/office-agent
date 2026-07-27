"""暴露给 LLM 的工具集（@tool 装饰器）包。

设计:
    - 会话级文档路径注入：main.py 启动 agent 前调用 set_session_doc(path)，
      所有工具内部读取该路径，LLM 不需要传路径参数（避免出错）。
    - 扩展名路由：路径后缀决定文档类型（.docx / .xlsx / .pptx），
      _tool() 工厂按 kind 返回 DocTool / ExcelTool / PptxTool。
    - 按会话类型裁剪：graph.py 用 ``tools_for_kind(kind)`` 只把当前会话
      用得到的工具绑定给 LLM（通用 + 对应格式专属 + 控制），而非全量 49 个——
      每轮请求少发几十个无关工具的 JSON schema，也减少 LLM 误调用。
    - ask_user 工具内部用 LangGraph interrupt 挂起，等用户输入后作为
      ToolMessage 回传给 agent。
    - finish 工具让 LLM 显式宣告完成。

结构:
    - ``session.py``: 会话状态基础设施（set_session_doc / session_doc_kind /
      _tool 工厂），子模块从它 import，无循环依赖。
    - 本 ``__init__.py``: 聚合 ALL_TOOLS / TOOL_BY_NAME / tools_for_kind，
      并 re-export session 符号（向后兼容）。
    - ``common.py``: 通用工具（三格式共用）+ 控制（ask_user/finish）+ 公文/查询。
    - ``doc.py``:    Word 专属工具。
    - ``excel.py``:  Excel 专属工具。
    - ``pptx.py``:   PowerPoint 专属工具。
    - ``batching.py``: 同批"末尾追加"类调用 → 一次 officecli batch 的翻译。
"""

from __future__ import annotations

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

# 会话基础设施（re-export，保持 from office_agent.tools import X 兼容）
from .session import (
    _doc,
    _tool,
    _wrong_kind_msg,
    doc_tool,
    excel_tool,
    pptx_tool,
    session_doc_kind,
    session_doc_path,
    set_session_doc,
)

# ============================================================
# 按会话类型的工具子集（graph.py 用 tools_for_kind 绑定）
# ============================================================
# 三格式共用
_COMMON_TOOLS = [
    create_doc,
    view_text,
    validate_doc,
    set_doc_properties,
]
# 控制类（每个会话都要）
_CONTROL_TOOLS = [
    ask_user,
    finish,
]
_WORD_ONLY = [
    start_from_template,
    add_table,
    add_image,
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
_EXCEL_ONLY = [
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
_PPTX_ONLY = [
    add_table,
    add_image,
    add_slide,
    add_textbox,
    add_slide_image,
    add_slide_table,
    set_slide_transition,
    set_slide_notes,
    set_theme_colors,
    set_theme_fonts,
]

_KIND_TOOLS = {
    "docx": _WORD_ONLY,
    "xlsx": _EXCEL_ONLY,
    "pptx": _PPTX_ONLY,
}


def tools_for_kind(kind: str, *, include_vehicle: bool = False) -> list:
    """返回某会话类型应绑定给 LLM 的工具子集。

    kind: 'docx' | 'xlsx' | 'pptx'（未知值按 docx 处理，与 session_doc_kind 一致）。
    include_vehicle: 需求与车辆/交通相关时为 True，附加 query_vehicle。

    只影响 bind_tools 暴露给 LLM 的清单；执行分发仍走全量 TOOL_BY_NAME，
    即使 LLM 幻觉调用了未绑定的工具也能得到友好错误而非崩溃。
    """
    specific = _KIND_TOOLS.get(kind, _WORD_ONLY)
    tools = [*_COMMON_TOOLS, *specific]
    if include_vehicle:
        tools.append(query_vehicle)
    tools.extend(_CONTROL_TOOLS)
    return tools


# ============================================================
# 全量工具清单（执行分发用；tools_for_kind 是它的子集视图）
# ============================================================
ALL_TOOLS = [
    # 通用
    create_doc,
    start_from_template,
    add_table,
    view_text,
    validate_doc,
    add_image,
    set_doc_properties,
    # Word 专属
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
    # Excel 专属
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
    # PowerPoint 专属
    add_slide,
    add_textbox,
    add_slide_image,
    add_slide_table,
    set_slide_transition,
    set_slide_notes,
    set_theme_colors,
    set_theme_fonts,
    # 业务专项 + 控制
    query_vehicle,
    ask_user,
    finish,
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
    "doc_tool",
    "excel_tool",
    "pptx_tool",
    "_wrong_kind_msg",
    # 聚合
    "ALL_TOOLS",
    "TOOL_BY_NAME",
    "tools_for_kind",
    # 所有工具（供 from office_agent.tools import X）
    *[t.name for t in ALL_TOOLS],
]
