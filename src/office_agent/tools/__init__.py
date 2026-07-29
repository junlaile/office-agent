"""暴露给 LLM 的工具集（@tool 装饰器）包。

设计:
    - 会话级文档路径注入：main.py 启动 agent 前调用 set_session_doc(path)，
      所有工具内部读取该路径，LLM 不需要传路径参数（避免出错）。
    - 扩展名路由：路径后缀决定文档类型（.docx / .xlsx / .pptx），
      _tool() 工厂按 kind 返回 DocTool / ExcelTool / PptxTool。
    - 通用工具（create_doc / add_table / view_text / validate_doc / finish）
      三种格式都支持；docx/xlsx/pptx 专属工具在其他格式下给出明确提示。
    - ask_user 只构造结构化交互请求；LangGraph interaction 节点负责
      interrupt/resume 并把答案作为 ToolMessage 回传给 agent。
    - finish 工具让 LLM 显式宣告完成。

结构:
    - 本 ``__init__.py``: 会话状态基础设施 + ToolRegistry 注册。
    - ``common.py``: 通用工具（三格式共用）+ 控制（ask_user/finish）+ 公文/查询。
    - ``doc.py``:    Word 专属工具。
    - ``excel.py``:  Excel 专属工具。
    - ``pptx.py``:   PowerPoint 专属工具。

子模块通过 ``from office_agent.tools import _tool, session_doc_kind, ...``
反向引用本包的会话基础设施——因此这些符号必须在 ``__init__.py`` 定义，
且子模块的 import 在 ``__init__.py`` 底部执行（此时基础设施已就绪）。
"""

from __future__ import annotations

from .session import (  # noqa: E402
    _doc,
    _tool,
    _wrong_kind_msg,
    session_doc_kind,
    session_doc_path,
    set_session_doc,
)


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
from .registry import (  # noqa: E402
    ALL_DOCUMENT_KINDS,
    DOCX_KINDS,
    PPTX_KINDS,
    XLSX_KINDS,
    ExecutionMode,
    SideEffect,
    ToolRegistry,
    ToolSpec,
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


TOOL_SPECS = [
    ToolSpec(create_doc, ALL_DOCUMENT_KINDS, side_effect=SideEffect.INIT),
    ToolSpec(add_table, ALL_DOCUMENT_KINDS),
    ToolSpec(view_text, ALL_DOCUMENT_KINDS, side_effect=SideEffect.READ),
    ToolSpec(validate_doc, ALL_DOCUMENT_KINDS, side_effect=SideEffect.READ),
    ToolSpec(set_doc_properties, ALL_DOCUMENT_KINDS),
    ToolSpec(add_image, frozenset({"docx", "pptx"})),
    ToolSpec(start_from_template, DOCX_KINDS, side_effect=SideEffect.INIT),
    *(ToolSpec(tool, DOCX_KINDS) for tool in WORD_TOOLS if tool is not start_from_template),
    *(ToolSpec(tool, XLSX_KINDS) for tool in EXCEL_TOOLS),
    *(ToolSpec(tool, PPTX_KINDS) for tool in PPTX_TOOLS),
    ToolSpec(query_vehicle, ALL_DOCUMENT_KINDS, side_effect=SideEffect.NONE),
    ToolSpec(
        ask_user,
        ALL_DOCUMENT_KINDS,
        execution_mode=ExecutionMode.INTERACTION,
        side_effect=SideEffect.HUMAN,
        can_batch=False,
    ),
    ToolSpec(finish, ALL_DOCUMENT_KINDS, side_effect=SideEffect.TERMINAL),
]

REGISTRY = ToolRegistry(TOOL_SPECS)


def tools_for_kind(kind: str, *, include_vehicle: bool = False):
    """按文档类型返回绑定给 LLM 的工具子集。

    默认不暴露 query_vehicle（仅在交通类任务显式开启）。
    历史兼容：未知 kind 回退到 docx。
    """
    normalized = str(kind or "").lower()
    if normalized not in ALL_DOCUMENT_KINDS:
        normalized = "docx"

    tools = []
    for spec in TOOL_SPECS:
        if normalized not in spec.document_kinds:
            continue
        # Excel 场景由 set_cells 系列负责表格写入，避免 add_table 误调用。
        if normalized == "xlsx" and spec.name == "add_table":
            continue
        if not include_vehicle and spec.name == "query_vehicle":
            continue
        tools.append(spec.tool)
    return tools

# 兼容既有导入；实际来源统一为注册表。
ALL_TOOLS = REGISTRY.all_tools
TOOL_BY_NAME = REGISTRY.tool_by_name
SPEC_BY_NAME = REGISTRY.spec_by_name
tools_for_doc_path = REGISTRY.bindable_tools

__all__ = [
    # 会话基础设施
    "set_session_doc",
    "session_doc_path",
    "session_doc_kind",
    "_tool",
    "_doc",
    "_wrong_kind_msg",
    # 聚合
    "ALL_TOOLS",
    "TOOL_BY_NAME",
    "TOOL_SPECS",
    "SPEC_BY_NAME",
    "REGISTRY",
    "ToolSpec",
    "ToolRegistry",
    "ExecutionMode",
    "SideEffect",
    "COMMON_TOOLS",
    "IMAGE_TOOLS",
    "WORD_TOOLS",
    "EXCEL_TOOLS",
    "PPTX_TOOLS",
    "CONTROL_TOOLS",
    "tools_for_doc_path",
    "tools_for_kind",
    # 所有工具（供 from office_agent.tools import X）
    *[t.name for t in ALL_TOOLS],
]
