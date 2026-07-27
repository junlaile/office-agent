"""Word（.docx）专属 LLM 工具。"""

from __future__ import annotations

import logging

from langchain_core.tools import tool

from office_agent.officecli import (
    OfficeCLIError,
)
from office_agent.tools.session import (
    _wrong_kind_msg,
    doc_tool,
    session_doc_kind,
)

logger = logging.getLogger(__name__)


@tool
def add_title(text: str) -> str:
    """【Word 专属】添加文档主标题（居中、大字号、加粗）。整篇文档只调一次。

    在 Excel/PowerPoint 下不要调用本工具：
        - PowerPoint 用 add_slide(title=...) 或 add_textbox 做标题。
        - Excel 没有标题概念，用 set_cells 写表头行。

    参数:
        text: 主标题文字，如"2024年第三季度销售报告"。"""
    kind = session_doc_kind()
    if kind != "docx":
        return (
            f"add_title 是 Word 专属工具，当前文档是 {kind}。"
            f"请改用：pptx→add_slide(title=...) 或 add_textbox；"
            f"xlsx→set_cells 写表头行。"
        )
    return doc_tool().add_title(text)


@tool
def add_heading(text: str, level: int = 1) -> str:
    """【Word 专属】添加一个章节标题。

    在 Excel/PowerPoint 下不要调用：pptx 用 add_textbox(size=28, bold=true)，
    xlsx 没有章节概念。

    参数:
        text: 标题文字（不要带编号，如用"项目背景"而非"一、项目背景"）。
        level: 标题层级 1-9，1=一级标题（最大），2=二级子标题。默认 1。"""
    kind = session_doc_kind()
    if kind != "docx":
        return (
            f"add_heading 是 Word 专属工具，当前文档是 {kind}。"
            f"pptx 用 add_textbox 当标题；xlsx 用 set_cells 写分组行。"
        )
    return doc_tool().add_heading(text, level=level)


@tool
def add_paragraph(text: str, bold: bool = False, italic: bool = False) -> str:
    """【Word 专属】添加一段正文。

    在 Excel/PowerPoint 下不要调用：pptx 用 add_textbox；xlsx 用 set_cell/set_cells。

    参数:
        text: 段落正文（纯文本，不要 markdown 标记，不要在一段里塞多个主题）。
        bold: 是否整段加粗（默认 false，仅强调时用）。
        italic: 是否斜体（默认 false）。"""
    kind = session_doc_kind()
    if kind != "docx":
        return (
            f"add_paragraph 是 Word 专属工具，当前文档是 {kind}。"
            f"pptx 用 add_textbox；xlsx 用 set_cell/set_cells。"
        )
    return doc_tool().add_paragraph(text, bold=bold, italic=italic)


@tool
def add_list_item(text: str, ordered: bool = False) -> str:
    """【Word 专属】添加一个列表项（项目符号或编号）。

    在 Excel/PowerPoint 下不要调用：pptx 用 add_textbox（多行文本即可当列表），
    xlsx 用 set_cells 写多行。

    参数:
        text: 单个列表项的文字。
        ordered: true=有序编号列表(1. 2. 3.)，false=无序项目符号(•)。默认 false。
    连续多次调用本工具即可构成一个完整列表。"""
    kind = session_doc_kind()
    if kind != "docx":
        return (
            f"add_list_item 是 Word 专属工具，当前文档是 {kind}。"
            f"pptx 用 add_textbox 多行；xlsx 用 set_cells。"
        )
    return doc_tool().add_list_item(text, ordered=ordered)


@tool
def add_toc(levels: str = "1-3", title: str = "目录") -> str:
    """【Word 专属】插入目录（Table of Contents）。

    生成多章节报告时【强烈建议】在文档开头或结尾加目录。
    目录自动收录所有用 add_heading 添加的标题。
    打开文档时不会弹「是否更新域」；若目录仍是占位文本，可在 Word 中右键目录
    → 更新域 / 按 F9 手动刷新。

    参数:
        levels: 收录的标题层级范围，如 '1-3'（收录 1-3 级标题）。默认 '1-3'。
        title: 目录上方的标题文字（如"目录"/"Contents"）。默认"目录"。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("add_toc", "docx", "目录只在 Word 文档中可用")
    try:
        return doc_tool().add_toc(levels=levels, title=title)
    except OfficeCLIError as e:
        return f"添加目录失败: {e}"


@tool
def add_page_number(align: str = "center") -> str:
    """【Word 专属】在页脚居中（或指定位置）加页码。

    参数:
        align: 位置 'center'(居中,默认)/'left'/'right'。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("add_page_number", "docx", "页码只在 Word 文档中可用")
    try:
        return doc_tool().add_footer(field="page", align=align)
    except OfficeCLIError as e:
        return f"添加页码失败: {e}"


@tool
def add_header(text: str, align: str = "right") -> str:
    """【Word 专属】加页眉文字（如文档标题、机密标识）。

    参数:
        text: 页眉文字。
        align: 对齐 'left'/'center'/'right'（默认 right）。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("add_header", "docx", "页眉只在 Word 文档中可用")
    try:
        return doc_tool().add_header(text, align=align)
    except OfficeCLIError as e:
        return f"添加页眉失败: {e}"


@tool
def add_footer(text: str, align: str = "center") -> str:
    """【Word 专属】加页脚文字（如版权、作者）。

    参数:
        text: 页脚文字。
        align: 对齐 'left'/'center'/'right'（默认 center）。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("add_footer", "docx", "页脚只在 Word 文档中可用")
    try:
        return doc_tool().add_footer(text, align=align)
    except OfficeCLIError as e:
        return f"添加页脚失败: {e}"


@tool
def add_hyperlink(text: str, url: str, tooltip: str = "") -> str:
    """【Word 专属】加一个超链接段落（链接到外部网址）。

    参数:
        text: 显示文字（如"点击访问官网"）。
        url: 链接地址（如 'https://example.com'）。
        tooltip: 鼠标悬停提示（可空）。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg(
            "add_hyperlink", "docx", "超链接只在 Word 文档中可用（pptx 用 add_slide）"
        )
    try:
        return doc_tool().add_hyperlink(text, url, tooltip=tooltip)
    except OfficeCLIError as e:
        return f"添加超链接失败: {e}"


@tool
def add_word_chart(chart_type: str, data: str, categories: str = "", title: str = "") -> str:
    """【Word 专属】在文档末尾加嵌入式图表（Word 图表自带数据）。

    参数:
        chart_type: 'column'(柱形)/'bar'(条形)/'line'(折线)/'pie'(饼)/'doughnut'(环形)/'area'(面积)。
        data: 内联数据，格式 '系列名:值,值,值'，多系列用分号分隔。
              例: 'Sales:10,20,30;Cost:5,8,12'
        categories: 分类标签，逗号分隔，如 'Q1,Q2,Q3,Q4'。数据有几个值就给几个标签。
        title: 图表标题。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg(
            "add_word_chart", "docx", "Word 图表只在 Word 文档中可用（Excel 用 add_excel_chart）"
        )
    try:
        return doc_tool().add_chart(
            chart_type,
            data,
            categories=categories,
            title=title,
        )
    except OfficeCLIError as e:
        return f"添加图表失败: {e}"


@tool
def add_section_break(orientation: str = "portrait") -> str:
    """【Word 专属】插入分节符（用于切换页面方向或分栏）。

    参数:
        orientation: 新节的页面方向 'portrait'(纵向,默认)/'landscape'(横向)。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("add_section_break", "docx", "分节只在 Word 文档中可用")
    try:
        return doc_tool().add_section(orientation=orientation)
    except OfficeCLIError as e:
        return f"添加分节失败: {e}"


@tool
def update_paragraph(path: str, text: str) -> str:
    """【Word 专属】修改指定段落的整段文字。

    【公文模式常用】把模板里的范例标题/主送/某段正文整段换成真实内容。
    先 view_text 看 /body/p[N] 路径，再用本工具改对应段。

    参数:
        path: 段落路径，view_text 输出里方括号标注的那个，如 '/body/p[4]'。
        text: 新的整段文字（纯文本，不要 markdown 标记）。

    ⚠️ 会重置段内的字体/字号到默认（段落对齐/缩进/行距保留）。
    对公文标题段、正文段，若想保留小标宋/仿宋字体，改几个字时优先用
    replace_text；只有"整段全换"时才用本工具。改完后建议 view_text 复查字体。

    例: update_paragraph(path='/body/p[4]', text='市公安局关于做好防汛工作的通知')
        —— 把范例标题换成真实标题。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("update_paragraph", "docx", "Excel/PowerPoint 没有段落概念")
    try:
        return doc_tool().set_paragraph_text(path, text)
    except OfficeCLIError as e:
        return f"修改段落失败: {e}"


@tool
def replace_text(find: str, replace: str, path: str = "") -> str:
    """【Word 专属】文本替换（保留字体格式，公文模式编辑正文的首选）。

    【公文模式主力工具】把模板正文里的 'XX'、'XX工作'、'XX机关' 等占位
    换成真实内容，同时【保留】原段落的字体/字号（小标宋、仿宋等）。
    比 update_paragraph 更安全——只动匹配的文字，不碰段内其余内容和格式。

    参数:
        find: 要查找的文字（字面子串，区分大小写）。
              如 'XX工作' 会匹配所有含 'XX工作' 的段落。
        replace: 替换成的文字。如 '防汛工作'。
        path: 作用域。留空=全文 body 范围替换（最常用）；
              '/body/p[N]' = 仅替换该段内的匹配。

    例: replace_text(find='XX', replace='防汛')  —— 把全文 'XX' 换成 '防汛'
        replace_text(find='XX局', replace='应急局', path='/body/p[5]')  —— 只改主送段

    注意: 'XX' 是常见占位但可能误伤（如 'XX市' 里的 'XX' 也会被换）。
    用更长的匹配串（'XX工作' 比 'XX' 精确）减少误伤。替换后 view_text 复查。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg(
            "replace_text", "docx", "Excel 用 set_cell；PowerPoint 用 add_textbox"
        )
    try:
        scope = path if path else "/body"
        return doc_tool().find_replace(find, replace, path=scope)
    except OfficeCLIError as e:
        return f"文本替换失败: {e}"


@tool
def remove_paragraph(path: str) -> str:
    """【Word 专属】删除指定段落（或表格/图片等其他元素）。

    【公文模式常用】删掉模板里多余的范例段，精简成用户实际需要的内容。

    参数:
        path: 要删除元素的路径，view_text 标注的，如 '/body/p[10]'。

    ⚠️ 删段后，后续段落索引立即前移（删 p[5] 后，原 p[6] 变成新 p[5]）。
    连续删多段时：【从后往前删】（先删 p[10] 再删 p[5]，顺序不变），
    或每删一段后重新 view_text 确认新索引。不要用旧索引连删。

    例: remove_paragraph(path='/body/p[10]')  —— 删第10段
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("remove_paragraph", "docx", "Excel/PowerPoint 暂不支持按路径删元素")
    try:
        return doc_tool().remove(path)
    except OfficeCLIError as e:
        return f"删除段落失败: {e}"
