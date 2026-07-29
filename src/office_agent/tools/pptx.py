"""PowerPoint（.pptx）专属 LLM 工具。"""

from __future__ import annotations

from langchain_core.tools import tool

from office_agent.log import get_logger
from office_agent.officecli import (
    OfficeCLIError,
)
from office_agent.tools.common import _validate_image_source
from office_agent.tools.session import (
    _wrong_kind_msg,
    pptx_tool,
    session_doc_kind,
)

logger = get_logger(__name__)


@tool
def add_slide(title: str = "", body_text: str = "", layout: str = "") -> str:
    """【PowerPoint 专属】添加一张幻灯片。生成 PPT 时每加一页都调本工具。

    【排版模式：占位符】每页的内容【只通过 title + body_text 写入】:
      - title 非空 → 自动建标题占位符（页面顶部）。
      - body_text 非空 → 自动建正文占位符（页面中下部）。
    正文占位符会按换行 \\n 自动拆成多段，支持要点符号、空行分段。

    【body_text 是内容页的必填项】只有标题没正文的页是废页。
      - 封面/章节分隔页：可以只传 title。
      - 【所有内容页必须传 body_text】，不能为空。
        ✗ add_slide(title='核心功能')
        ✓ add_slide(title='核心功能', body_text='· 要点1\\n· 要点2')

    【绝不要】在同一页再调 add_textbox / add_slide_table / add_slide_image——
    它们会叠加在已有占位符之上导致文字重叠。一页放不下就拆成多页。
    （这些工具是高级用法，仅用于 add_slide() 不带任何参数的纯空白页，
    常规 PPT 用不到。）

    参数:
        title: 幻灯片标题。
        body_text: 【内容页必填】正文。多行用 \\n 分隔。排版建议:
            - 每行一个要点，前缀用 '·'/'-'/'•'。
            - 不同小节之间加一个空行（即 \\n\\n）。
            - 单页 5-8 行为宜；超了就拆页。
            例: '· 要点一\\n· 要点二\\n\\n补充说明:\\n· 细节A\\n· 细节B'
        layout: 可选，通常【留空】。只是元数据标签，不影响占位符创建。

    典型流程: create_doc → add_slide(封面) → add_slide(每页内容)... → finish。
    """
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg(
            "add_slide", "pptx", "docx 用 add_heading/add_paragraph；xlsx 用 add_sheet"
        )
    # 内容页必须有正文：只传 title 不传 body_text 通常是 LLM 漏写正文，
    # 返回明确警告让它补上（封面/章节页可无视，重试一次 body_text 任意内容即可）。
    # 不能直接拒绝——封面页确实只需 title。所以用"警告但放行 + 强提示"策略。
    result = ""
    warning = ""
    if title and not body_text.strip():
        warning = (
            "⚠️ 本页只有标题没有正文（body_text 为空）。"
            "若这是封面/章节分隔页可以；若是【内容页】，请重新调用本工具并"
            "【补上 body_text】写入要点内容——否则这页会是空的。\n"
        )
    try:
        result = pptx_tool().add_slide(title=title, text=body_text, layout=layout)
        return warning + result
    except OfficeCLIError as e:
        return f"添加幻灯片失败: {e}"


@tool
def add_textbox(
    text: str,
    x: str = "1cm",
    y: str = "2cm",
    width: str = "22cm",
    height: str = "2cm",
    size: float = 18,
    bold: bool = False,
    color: str = "",
    fill: str = "",
    align: str = "left",
) -> str:
    """【PowerPoint 专属·高级】在【最新一张】幻灯片上加自由文本框。

    ⚠️ 高级工具，常规 PPT 【用不到】。常规排版请用 add_slide(title, body_text)
    把内容写进正文占位符即可。本工具【仅】用于这种场景：
    你调过 add_slide() 【不带 title 和 body_text】得到一张纯空白页，需要在上面
    精细摆放多个文本块。
    【绝不要】在已用 title/body_text 创建占位符的页上再调本工具——会与正文
    占位符重叠，导致文字叠在一起。

    参数:
        text: 文本内容（可含换行 \\n 当多行）。
        x/y: 文本框左上角位置（如 '2cm'）。幻灯片画布 33.87cm × 19.05cm（16:9）。
        width/height: 文本框宽高。
        size: 字号 pt（默认 18）。标题建议 28-36，正文 16-20，注释 12-14。
        bold: 是否加粗。
        color: 文字颜色，6位十六进制（如 '4472C4' 蓝、'FFFFFF' 白）。
        fill: 文本框背景色（如 'FFFF00' 黄底）。可空。
        align: 水平对齐 'left'/'center'/'right'/'justify'。"""
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg("add_textbox", "pptx", "docx 用 add_paragraph；xlsx 用 set_cell")
    try:
        pptx = pptx_tool()
        slide_index = pptx.last_slide_index() or 1
        return pptx.add_textbox(
            slide_index,
            text,
            x=x,
            y=y,
            width=width,
            height=height,
            size=size,
            bold=bold,
            color=color,
            fill=fill,
            align=align,
        )
    except OfficeCLIError as e:
        return f"添加文本框失败: {e}"


@tool
def add_slide_image(url_or_path: str, x: str = "2cm", y: str = "2cm", width: str = "15cm") -> str:
    """【PowerPoint 专属·高级】在【最新一张】幻灯片上插入图片。

    ⚠️ 仅用于纯空白页（add_slide() 不带 title/body_text）。若本页已有正文占位符，
    图片可能与文字重叠——常规内容页要配图时，优先把图片作为独立一页，
    或调整 x/y/width 让图片落在正文占位符之外的区域（占位符约位于 y=5cm~17cm、
    x=0.9cm~30cm，避开这片区域）。

    参数:
        url_or_path: 图片来源（本地路径 / URL / data URI）。
        x/y: 图片左上角位置。画布 33.87cm × 19.05cm（16:9）。
        width: 显示宽度（高度按图片原始比例自动算）。"""
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg(
            "add_slide_image", "pptx", "docx 用 add_image；xlsx 单元格内图片暂不支持"
        )
    # 预校验图片来源：坏源在碰文档前跳过，不嵌入、不浪费 officecli 调用。
    reason = _validate_image_source(url_or_path)
    if reason:
        return (
            f"⚠️ 跳过插入图片（{reason}）。"
            f"不要重试这张图，继续生成幻灯片其他内容。"
        )
    try:
        pptx = pptx_tool()
        slide_index = pptx.last_slide_index() or 1
        return pptx.add_image(slide_index, url_or_path, x=x, y=y, width=width)
    except OfficeCLIError as e:
        return f"插入图片失败: {e}"


@tool
def add_slide_table(
    data: list[list], has_header: bool = True, x: str = "1cm", y: str = "4cm", width: str = "22cm"
) -> str:
    """【PowerPoint 专属·高级】在【最新一张】幻灯片上加表格。

    ⚠️ 仅用于纯空白页（add_slide() 不带 title/body_text）。若本页已有正文占位符，
    表格会与之重叠。常规做法：需要表格的页，用 add_slide(title='xxx') 只设标题
    （不设 body_text），再调本工具加表格——但要注意表格的 y 起点要在标题下方
    （标题约占 y=0~3.7cm，表格 y 建议从 5cm 起）。

    参数:
        data: 二维数组（外层行、内层单元格）。
        has_header: 第一行作为表头（加粗）。
        x/y: 表格左上角位置。画布 33.87cm × 19.05cm（16:9）。
        width: 表格总宽。

    表格列宽自动均分。建议表格不超过 6 列，超过会显挤。"""
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg("add_slide_table", "pptx", "docx 用 add_table；xlsx 用 set_cells")
    try:
        clean = []
        for row in data or []:
            if row is None:
                continue
            clean.append([("" if c is None else str(c)) for c in row])
        if not clean:
            return "添加表格失败: 数据为空"
        pptx = pptx_tool()
        slide_index = pptx.last_slide_index() or 1
        return pptx.add_table(slide_index, clean, has_header=has_header, x=x, y=y, width=width)
    except OfficeCLIError as e:
        return f"添加表格失败: {e}"


# ============================================================
# Word 进阶工具（docx）
# ============================================================


@tool
def set_slide_transition(
    slide_index: int, transition: str, auto_advance_seconds: float | None = None
) -> str:
    """【PowerPoint 专属】设置幻灯片切换效果（页与页之间的过渡动画）。

    参数:
        slide_index: 幻灯片 1-based 序号。
        transition: 切换效果。常用:
            'fade'        淡入（最常用，柔和）
            'push-right'  从右推移
            'wipe'        擦除
            'morph'       平滑变形（相邻页有同名形状时效果惊艳，推荐）
            'none'        无切换
            'zoom'/'cover'/'split'/'dissolve'/'circle'/'wheel' 等
        auto_advance_seconds: 自动换页秒数（如 5.0=5秒）。None=需点击换页。
    """
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg("set_slide_transition", "pptx", "切换效果只在 PPT 中可用")
    try:
        advance_ms = int(auto_advance_seconds * 1000) if auto_advance_seconds else None
        return pptx_tool().set_transition(
            slide_index,
            transition,
            advance_time_ms=advance_ms,
        )
    except OfficeCLIError as e:
        return f"设置切换失败: {e}"


@tool
def set_slide_notes(slide_index: int, notes: str) -> str:
    """【PowerPoint 专属】设置幻灯片的演讲者备注（放映时仅演讲者可见）。

    参数:
        slide_index: 幻灯片 1-based 序号。
        notes: 备注全文（讲稿要点、提示词，可含换行）。
    """
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg("set_slide_notes", "pptx", "备注只在 PPT 中可用")
    try:
        return pptx_tool().set_notes(slide_index, notes)
    except OfficeCLIError as e:
        return f"设置备注失败: {e}"


@tool
def set_theme_colors(
    accent1: str = "", accent2: str = "", accent3: str = "", hyperlink: str = ""
) -> str:
    """【PowerPoint 专属】自定义主题颜色（影响整个 deck 的强调色）。

    只传需要改的颜色（6位十六进制如 '4472C4'），其余保持不变。
    建议在 create_doc 后、add_slide 前调用，让所有幻灯片都使用新主题色。

    参数:
        accent1: 主强调色（影响标题、形状默认色）。如 '4472C4' 蓝。
        accent2: 次强调色。如 'ED7D31' 橙。
        accent3: 第三强调色。
        hyperlink: 超链接颜色。如 '0563C1'。
    """
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg("set_theme_colors", "pptx", "主题色只在 PPT 中可用")
    try:
        return pptx_tool().set_theme_colors(
            accent1=accent1,
            accent2=accent2,
            accent3=accent3,
            hyperlink=hyperlink,
        )
    except OfficeCLIError as e:
        return f"设置主题色失败: {e}"


@tool
def set_theme_fonts(heading_font: str = "", body_font: str = "") -> str:
    """【PowerPoint 专属】自定义主题字体（影响整个 deck）。

    建议在 create_doc 后、add_slide 前调用。

    参数:
        heading_font: 标题字体（如 '微软雅黑'）。
        body_font: 正文字体（如 '微软雅黑'）。
    """
    if session_doc_kind() != "pptx":
        return _wrong_kind_msg("set_theme_fonts", "pptx", "主题字体只在 PPT 中可用")
    try:
        return pptx_tool().set_theme_fonts(
            heading_font=heading_font,
            body_font=body_font,
        )
    except OfficeCLIError as e:
        return f"设置主题字体失败: {e}"


# ============================================================
# 车辆号牌查询（交通类文档专项）
# ============================================================
