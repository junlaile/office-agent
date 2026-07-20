"""暴露给 LLM 的工具集（@tool 装饰器）。

设计:
    - 会话级文档路径注入：main.py 启动 agent 前调用 set_session_doc(path)，
      所有工具内部读取该路径，LLM 不需要传路径参数（避免出错）。
    - 扩展名路由：路径后缀决定文档类型（.docx / .xlsx / .pptx），
      _tool() 工厂按 kind 返回 DocTool / ExcelTool / PptxTool。
      通用工具（create_doc / add_table / view_text / validate_doc / finish）
      三种格式都支持；docx 专属工具（add_title 等）在其他格式下给出
      明确提示，引导 LLM 改用对应格式的专属工具。
    - ask_user 工具内部用 LangGraph interrupt 挂起，等用户输入后作为
      ToolMessage 回传给 agent。
    - finish 工具让 LLM 显式宣告完成。

每个工具的 docstring 是 LLM 判断"何时调用"的依据，写得具体清晰。
"""

from __future__ import annotations

from typing import Annotated, Any

from langchain_core.tools import tool
from langgraph.types import interrupt
from pydantic import BaseModel, Field

from .officecli import DocTool, ExcelTool, OfficeCLIError, PptxTool

# ============================================================
# 会话状态（模块级，agent 运行期间唯一）
# ============================================================
_session_doc_path: str | None = None


def set_session_doc(path: str) -> None:
    """main.py 在启动 agent 前调用，设定本会话的文档路径。"""
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


# 向后兼容：旧代码可能引用 _doc()
def _doc() -> DocTool:
    return _tool()  # type: ignore[return-value]


def _wrong_kind_msg(tool_name: str, expected: str, hint: str = "") -> str:
    """当前会话格式与工具不符时的友好提示（而非 AttributeError 崩溃）。"""
    actual = session_doc_kind()
    hint_str = f" {hint}" if hint else ""
    return (f"{tool_name} 是 {expected.upper()} 专属工具，当前文档是 {actual}。{hint_str}").strip()


# ============================================================
# 通用工具：三种格式都适用
# ============================================================
@tool
def create_doc() -> str:
    """创建一个新的空白 Office 文档。这是生成文档时【必须第一个调用】的工具，
    会覆盖同名旧文档。调用一次即可，之后用对应格式的添加工具往里加内容。

    文档类型由会话决定（Word/Excel/PowerPoint），你无需关心，工具会自动选对。
    """
    return _tool().create()


@tool
def add_table(data: list[list], has_header: bool = True) -> str:
    """添加一个表格。适合展示对比、数据、结构化信息。三种文档类型都支持。

    参数:
        data: 二维数组，外层是行、内层是单元格。每行长度应一致。
              单元格可以是字符串或数字（内部自动转字符串）。
              例: [["姓名","成绩"],["张三",95],["李四",88]]。
        has_header: 第一行是否作为表头（加粗）。默认 true。

    各格式行为:
        - Word: 在文档末尾插入表格。
        - Excel: 在当前工作表末尾追加（建议优先用专门的 set_cells，可控起始位置）。
        - PowerPoint: 加到最新一张幻灯片（建议先 add_slide 再加表格）。
    """
    kind = session_doc_kind()
    try:
        clean = []
        for row in (data or []):
            if row is None:
                continue
            clean.append([("" if c is None else str(c)) for c in row])
        if not clean:
            return "添加表格失败: 数据为空"

        t = _tool()
        if kind == "pptx":
            # PptxTool.add_table 需要 slide_index，加到最新幻灯片
            pptx: PptxTool = t  # type: ignore[assignment]
            slide_index = pptx._last_slide_index() or 1
            pptx.add_table(slide_index, clean, has_header=has_header)
        else:
            t.add_table(clean, has_header=has_header)  # type: ignore[attr-defined]
        rows = len(clean)
        cols = max(len(r) for r in clean) if clean else 0
        return f"已添加 {rows} 行 × {cols} 列的表格"
    except OfficeCLIError as e:
        return f"添加表格失败: {e}"


@tool
def view_text() -> str:
    """读取当前文档的全部纯文本内容。

    建议在调 finish 之前调用一次，自查文档结构和内容是否正确、完整。

    各格式输出:
        - Word: 段落路径 + 文本
        - Excel: 每个 sheet 的 A1=value 制表符分隔
        - PowerPoint: 每张幻灯片的文本（按 slide 分段）
    """
    return _tool().view_text()


@tool
def validate_doc() -> str:
    """校验当前文档是否符合 OpenXML 规范。返回校验结果。
    一般无需调用；若怀疑文档有问题时可用来确认。"""
    return _tool().validate()


# ============================================================
# Word 专属工具（docx）
# ============================================================
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
        return (f"add_title 是 Word 专属工具，当前文档是 {kind}。"
                f"请改用：pptx→add_slide(title=...) 或 add_textbox；"
                f"xlsx→set_cells 写表头行。")
    return _tool().add_title(text)  # type: ignore[union-attr]


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
        return (f"add_heading 是 Word 专属工具，当前文档是 {kind}。"
                f"pptx 用 add_textbox 当标题；xlsx 用 set_cells 写分组行。")
    return _tool().add_heading(text, level=level)  # type: ignore[union-attr]


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
        return (f"add_paragraph 是 Word 专属工具，当前文档是 {kind}。"
                f"pptx 用 add_textbox；xlsx 用 set_cell/set_cells。")
    return _tool().add_paragraph(text, bold=bold, italic=italic)  # type: ignore[union-attr]


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
        return (f"add_list_item 是 Word 专属工具，当前文档是 {kind}。"
                f"pptx 用 add_textbox 多行；xlsx 用 set_cells。")
    return _tool().add_list_item(text, ordered=ordered)  # type: ignore[union-attr]


@tool
def add_image(url_or_path: str, width: str = "8cm", caption: str = "") -> str:
    """【Word/PowerPoint 通用】插入一张图片。

    参数:
        url_or_path: 图片来源。支持本地文件路径、HTTP/HTTPS URL、data URI。
        width: 显示宽度，如 '8cm'/'400px'/'3in'（默认 8cm）。
        caption: 【仅 Word】可选图注文字。非空时在图片下方显示。

    - Word: 在文档末尾插入。
    - PowerPoint: 默认加到【最新一张幻灯片】。要加到特定幻灯片请先 add_slide 再调本工具。
    """
    kind = session_doc_kind()
    try:
        t = _tool()
        if kind == "pptx":
            pptx: PptxTool = t  # type: ignore[assignment]
            slide_index = pptx._last_slide_index() or 1
            return pptx.add_image(slide_index, url_or_path, width=width,
                                  alt=caption or "图片")
        # docx
        return t.add_image(url_or_path, width=width,  # type: ignore[union-attr]
                           alt=caption or "图片", caption=caption)
    except OfficeCLIError as e:
        return f"插入图片失败: {e}"


# ============================================================
# Excel 专属工具（xlsx）
# ============================================================
@tool
def add_sheet(name: str, tab_color: str = "") -> str:
    """【Excel 专属】添加一张工作表（tab）。

    新建的 xlsx 默认自带一张名为 'Sheet1' 的空工作表。本工具再加新的工作表。
    典型流程: create_doc → add_sheet('销售数据') → set_cells('销售数据', ...)。

    参数:
        name: 工作表名（如 '销售数据'）。后续 set_cell/set_cells/add_excel_chart
              都要传这个名字定位工作表。
        tab_color: 可选标签色，6位十六进制（如 '4472C4' 蓝色）。美化用。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_sheet", "xlsx",
                               "docx 用 add_table；pptx 无工作表概念")
    try:
        return _tool().add_sheet(name, tab_color=tab_color)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"添加工作表失败: {e}"


@tool
def set_cell(sheet: str, ref: str, value: Any,
             bold: bool = False, fill: str = "",
             number_format: str = "") -> str:
    """【Excel 专属】写单个单元格。最基础、最灵活的写入工具。

    参数:
        sheet: 工作表名（如 'Sheet1' 或 '销售数据'）。
        ref: 单元格地址，列字母+行号，如 'A1' / 'B2' / 'AA10'。
        value: 单元格值。数字会存为数字（参与计算），字符串存为文本。
               注意：纯数字字符串如电话号 '01234' 想保留前导 0，
               要传 number_format='@' 强制文本。
        bold: 是否加粗（表头常用）。
        fill: 背景色，6位十六进制（如 'FFFF00' 黄色高亮）。
        number_format: 数字格式代码。常用:
            '@' 文本（保留前导0）
            '#,##0' 千分位整数
            '#,##0.00' 千分位两位小数
            '0%' 百分比
            'yyyy-mm-dd' 日期
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_cell", "xlsx",
                               "docx 用 add_table/add_paragraph；pptx 用 add_textbox")
    try:
        return _tool().set_cell(  # type: ignore[union-attr]
            sheet, ref, value, bold=bold, fill=fill, number_format=number_format,
        )
    except OfficeCLIError as e:
        return f"写入单元格失败: {e}"


@tool
def set_cells(sheet: str, data: list[list], start: str = "A1",
              has_header: bool = False) -> str:
    """【Excel 专属】批量写二维数据到工作表。写表格数据【首选】本工具，比逐个 set_cell 高效得多。

    参数:
        sheet: 目标工作表名。
        data: 二维数组，外层=行、内层=单元格。数字/字符串混合均可。
              例: [["月份","销售额"],["1月",12000],["2月",15000]]
        start: 起始单元格地址（默认 'A1'）。内部按行列递增自动算出每个 cell 的地址。
               如 start='B2' 则 data[0][0] 写到 B2、data[0][1] 写到 C2、data[1][0] 写到 B3...
        has_header: 第一行是否作为表头（加粗）。默认 false。

    一次调用写入一整片区域，适合写完整的表格/报表数据。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_cells", "xlsx",
                               "docx 用 add_table；pptx 用 add_slide_table")
    try:
        return _tool().set_cells(sheet, data, start_ref=start,  # type: ignore[union-attr]
                                 has_header=has_header)
    except OfficeCLIError as e:
        return f"批量写入失败: {e}"


@tool
def set_formula(sheet: str, ref: str, formula: str) -> str:
    """【Excel 专属】在单元格写公式（不带前导 =）。

    参数:
        sheet: 工作表名。
        ref: 单元格地址（如 'D2'）。
        formula: 公式表达式，【不要】带前导 '='。
              例: 'SUM(B2:B10)' / 'AVERAGE(C2:C5)' / 'B2-C2' / 'B2*C2*0.1'。

    公式里引用的单元格必须已先用 set_cell/set_cells 写入值，否则结果是 0 或错误。"""
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_formula", "xlsx", "公式只在 Excel 表格中可用")
    try:
        return _tool().set_formula(sheet, ref, formula)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"写公式失败: {e}"


@tool
def add_excel_chart(sheet: str, chart_type: str, data_range: str,
                    title: str = "", categories: str = "") -> str:
    """【Excel 专属】在工作表上加图表（基于已写入的单元格数据）。

    前提: 数据必须已先用 set_cells 写入工作表，本工具只是引用这些单元格画图。

    参数:
        sheet: 图表要放在哪张工作表（图浮在该 sheet 上）。
        chart_type: 图表类型。常用:
            'column' 柱形图（默认推荐，对比类目数据）
            'bar'    条形图（横向柱状图）
            'line'   折线图（趋势）
            'pie'    饼图（占比）
            'doughnut' 环形图
            'area'   面积图
            'scatter' 散点图（相关性）
        data_range: 数据源区域，【必须带工作表名前缀】。
              格式 '工作表名!起始:结束'，如 '销售数据!B1:C4'。
              默认【首列当分类轴、其余列当数据系列】。
              若想让每列都是系列，请用 categories 参数单独指定分类轴。
        title: 图表标题（可空）。
        categories: 可选。分类轴数据，两种写法:
              - 区域引用: '销售数据!A2:A4'（推荐，引用已写入的类目标签）
              - 逗号分隔: 'Q1,Q2,Q3,Q4'

    示例: 数据写在 销售数据!A1:C4（A列月份/B列销售/C列成本），
          调 add_excel_chart('销售数据','column','销售数据!B1:C4',
                             title='季度销售vs成本', categories='销售数据!A2:A4')
          → 画出 B/C 两列为系列、A 列月份为分类的柱形图。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_excel_chart", "xlsx",
                               "docx/pptx 的图表暂不支持数据区域引用")
    try:
        return _tool().add_chart(  # type: ignore[union-attr]
            sheet, chart_type, data_range, categories=categories, title=title,
        )
    except OfficeCLIError as e:
        return f"添加图表失败: {e}"


# ============================================================
# PowerPoint 专属工具（pptx）
# ============================================================
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
        return _wrong_kind_msg("add_slide", "pptx",
                               "docx 用 add_heading/add_paragraph；xlsx 用 add_sheet")
    # 内容页必须有正文：只传 title 不传 body_text 通常是 LLM 漏写正文，
    # 返回明确警告让它补上（封面/章节页可无视，重试一次 body_text 任意内容即可）。
    # 不能直接拒绝——封面页确实只需 title。所以用"警告但放行 + 强提示"策略。
    result = ""
    warning = ""
    if title and not body_text.strip():
        warning = ("⚠️ 本页只有标题没有正文（body_text 为空）。"
                   "若这是封面/章节分隔页可以；若是【内容页】，请重新调用本工具并"
                   "【补上 body_text】写入要点内容——否则这页会是空的。\n")
    try:
        result = _tool().add_slide(title=title, text=body_text, layout=layout)  # type: ignore[union-attr]
        return warning + result
    except OfficeCLIError as e:
        return f"添加幻灯片失败: {e}"


@tool
def add_textbox(text: str, x: str = "1cm", y: str = "2cm",
                width: str = "22cm", height: str = "2cm",
                size: float = 18, bold: bool = False, color: str = "",
                fill: str = "", align: str = "left") -> str:
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
        return _wrong_kind_msg("add_textbox", "pptx",
                               "docx 用 add_paragraph；xlsx 用 set_cell")
    try:
        pptx: PptxTool = _tool()  # type: ignore[assignment]
        slide_index = pptx._last_slide_index() or 1
        return pptx.add_textbox(
            slide_index, text, x=x, y=y, width=width, height=height,
            size=size, bold=bold, color=color, fill=fill, align=align,
        )
    except OfficeCLIError as e:
        return f"添加文本框失败: {e}"


@tool
def add_slide_image(url_or_path: str, x: str = "2cm", y: str = "2cm",
                    width: str = "15cm") -> str:
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
        return _wrong_kind_msg("add_slide_image", "pptx",
                               "docx 用 add_image；xlsx 单元格内图片暂不支持")
    try:
        pptx: PptxTool = _tool()  # type: ignore[assignment]
        slide_index = pptx._last_slide_index() or 1
        return pptx.add_image(slide_index, url_or_path, x=x, y=y, width=width)
    except OfficeCLIError as e:
        return f"插入图片失败: {e}"


@tool
def add_slide_table(data: list[list], has_header: bool = True,
                    x: str = "1cm", y: str = "4cm", width: str = "22cm") -> str:
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
        return _wrong_kind_msg("add_slide_table", "pptx",
                               "docx 用 add_table；xlsx 用 set_cells")
    try:
        clean = []
        for row in (data or []):
            if row is None:
                continue
            clean.append([("" if c is None else str(c)) for c in row])
        if not clean:
            return "添加表格失败: 数据为空"
        pptx: PptxTool = _tool()  # type: ignore[assignment]
        slide_index = pptx._last_slide_index() or 1
        return pptx.add_table(slide_index, clean, has_header=has_header,
                              x=x, y=y, width=width)
    except OfficeCLIError as e:
        return f"添加表格失败: {e}"


# ============================================================
# Word 进阶工具（docx）
# ============================================================
@tool
def add_toc(levels: str = "1-3", title: str = "目录") -> str:
    """【Word 专属】插入目录（Table of Contents）。

    生成多章节报告时【强烈建议】在文档开头或结尾加目录。
    目录自动收录所有用 add_heading 添加的标题。

    参数:
        levels: 收录的标题层级范围，如 '1-3'（收录 1-3 级标题）。默认 '1-3'。
        title: 目录上方的标题文字（如"目录"/"Contents"）。默认"目录"。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("add_toc", "docx", "目录只在 Word 文档中可用")
    try:
        return _tool().add_toc(levels=levels, title=title)  # type: ignore[union-attr]
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
        return _tool().add_footer(field="page", align=align)  # type: ignore[union-attr]
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
        return _tool().add_header(text, align=align)  # type: ignore[union-attr]
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
        return _tool().add_footer(text, align=align)  # type: ignore[union-attr]
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
        return _wrong_kind_msg("add_hyperlink", "docx",
                               "超链接只在 Word 文档中可用（pptx 用 add_slide）")
    try:
        return _tool().add_hyperlink(text, url, tooltip=tooltip)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"添加超链接失败: {e}"


@tool
def add_word_chart(chart_type: str, data: str, categories: str = "",
                   title: str = "") -> str:
    """【Word 专属】在文档末尾加嵌入式图表（Word 图表自带数据）。

    参数:
        chart_type: 'column'(柱形)/'bar'(条形)/'line'(折线)/'pie'(饼)/'doughnut'(环形)/'area'(面积)。
        data: 内联数据，格式 '系列名:值,值,值'，多系列用分号分隔。
              例: 'Sales:10,20,30;Cost:5,8,12'
        categories: 分类标签，逗号分隔，如 'Q1,Q2,Q3,Q4'。数据有几个值就给几个标签。
        title: 图表标题。
    """
    if session_doc_kind() != "docx":
        return _wrong_kind_msg("add_word_chart", "docx",
                               "Word 图表只在 Word 文档中可用（Excel 用 add_excel_chart）")
    try:
        return _tool().add_chart(  # type: ignore[union-attr]
            chart_type, data, categories=categories, title=title,
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
        return _tool().add_section(orientation=orientation)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"添加分节失败: {e}"


@tool
def set_doc_properties(title: str = "", author: str = "",
                       subject: str = "", keywords: str = "") -> str:
    """【Word/Excel/PowerPoint 通用】设置文档核心属性（文件信息里显示）。

    参数都可空，只传需要设置的:
        title: 文档标题。
        author: 作者。
        subject: 主题。
        keywords: 关键词（逗号分隔）。
    """
    try:
        t = _tool()
        kind = session_doc_kind()
        if kind == "pptx":
            return t.set_presentation_props(  # type: ignore[union-attr]
                title=title, author=author, subject=subject,
            )
        # docx / xlsx 都用 set on /
        return t.set_doc_properties(  # type: ignore[union-attr]
            title=title, author=author, subject=subject, keywords=keywords,
        )
    except OfficeCLIError as e:
        return f"设置文档属性失败: {e}"


# ============================================================
# Excel 进阶工具（xlsx）
# ============================================================
@tool
def sort_sheet(sheet: str, keys: str, has_header: bool = True) -> str:
    """【Excel 专属】对工作表数据排序（就地修改行顺序）。

    参数:
        sheet: 工作表名。
        keys: 排序键，格式 '列字母 [asc|desc]'，多键逗号分隔。
              例: 'B desc'（按 B 列降序）/ 'A asc, B desc'（先 A 升再 B 降）/ 'C'（C 列升序）。
        has_header: 首行是否为表头（不参与排序）。默认 true。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("sort_sheet", "xlsx", "排序只在 Excel 中可用")
    try:
        return _tool().sort(sheet, keys, has_header=has_header)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"排序失败: {e}"


@tool
def set_autofilter(sheet: str, cell_range: str = "") -> str:
    """【Excel 专属】开启自动筛选（表头出现下拉箭头，可筛选）。

    参数:
        sheet: 工作表名。
        cell_range: 筛选区域如 'A1:D10'。留空=自动识别已用区域。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_autofilter", "xlsx", "自动筛选只在 Excel 中可用")
    try:
        return _tool().set_autofilter(sheet, cell_range)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"开启筛选失败: {e}"


@tool
def highlight_cells(sheet: str, cell_range: str, operator: str,
                    value: str, fill: str = "FFFF00") -> str:
    """【Excel 专属】条件格式：高亮满足条件的单元格（如大于某值）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'C2:C100'。
        operator: 比较运算符 'greaterThan'/'lessThan'/'equal'/'between' 等。
        value: 比较值。between 时用 '低值,高值'（此时会自动用 value2）。
        fill: 高亮背景色，6位十六进制（默认 'FFFF00' 黄色）。

    例: 高亮销售额>10000 的单元格
        highlight_cells('销售','C2:C100','greaterThan','10000','FF0000')
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("highlight_cells", "xlsx", "条件格式只在 Excel 中可用")
    try:
        props: dict = {"operator": operator, "fill": fill}
        if "," in value and operator in ("between", "notBetween"):
            v1, v2 = value.split(",", 1)
            props["value"] = v1.strip()
            props["value2"] = v2.strip()
        else:
            props["value"] = value
        return _tool().add_conditional_format(  # type: ignore[union-attr]
            sheet, "cellIs", cell_range, **props,
        )
    except OfficeCLIError as e:
        return f"添加条件格式失败: {e}"


@tool
def add_color_scale(sheet: str, cell_range: str,
                    min_color: str = "F8696B",
                    mid_color: str = "FFEB84",
                    max_color: str = "63BE7B") -> str:
    """【Excel 专属】条件格式：3色渐变（红-黄-绿，数据热力图）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'C2:C100'。
        min_color: 最小值颜色（默认红 F8696B）。
        mid_color: 中间值颜色（默认黄 FFEB84）。
        max_color: 最大值颜色（默认绿 63BE7B）。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_color_scale", "xlsx", "条件格式只在 Excel 中可用")
    try:
        return _tool().add_conditional_format(  # type: ignore[union-attr]
            sheet, "colorScale", cell_range,
            minColor=min_color, midColor=mid_color, maxColor=max_color,
        )
    except OfficeCLIError as e:
        return f"添加色阶失败: {e}"


@tool
def add_data_bar(sheet: str, cell_range: str,
                 color: str = "638EC6") -> str:
    """【Excel 专属】条件格式：数据条（单元格内显示横向条形）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'C2:C100'。
        color: 条形颜色（默认蓝 638EC6）。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_data_bar", "xlsx", "条件格式只在 Excel 中可用")
    try:
        return _tool().add_conditional_format(  # type: ignore[union-attr]
            sheet, "dataBar", cell_range, color=color,
        )
    except OfficeCLIError as e:
        return f"添加数据条失败: {e}"


@tool
def add_pivot_table(sheet: str, source: str, rows: str, values: str,
                    cols: str = "", filters: str = "",
                    position: str = "") -> str:
    """【Excel 专属】加数据透视表（汇总分析的利器）。

    参数:
        sheet: 透视表放在哪张工作表。
        source: 源数据区域，如 'Sheet1!A1:D100'（必须含表头行，区域要带工作表名前缀）。
        rows: 行字段，逗号分隔，如 '区域,产品'。
        values: 值字段及聚合方式，格式 '字段:agg'，多个逗号分隔。
                agg ∈ sum/avg/count/max/min。例: '销售额:sum,数量:count'。
        cols: 列字段，如 '季度'（可空）。
        filters: 筛选页字段，如 '年份'（可空）。
        position: 透视表左上角位置，如 'H1'。留空自动放在源数据旁。

    例: 按 region 汇总 sales 总和
        add_pivot_table('Sheet1','Sheet1!A1:D100','region','sales:sum',position='F1')
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_pivot_table", "xlsx", "透视表只在 Excel 中可用")
    try:
        return _tool().add_pivot_table(  # type: ignore[union-attr]
            sheet, source, rows=rows, cols=cols, values=values,
            filters=filters, position=position,
        )
    except OfficeCLIError as e:
        return f"添加透视表失败: {e}"


@tool
def add_list_table(sheet: str, cell_range: str, style: str = "medium2",
                   total_row: bool = False) -> str:
    """【Excel 专属】把单元格区域转成真正的 Excel 表格（带样式、筛选按钮、结构化引用）。

    与普通写数据的区别: 真 Excel 表格自动有蓝色条纹样式、表头筛选按钮、
    可用结构化引用（如 Table1[销售额]），还能一键汇总。

    参数:
        sheet: 工作表名。
        cell_range: 区域，如 'A1:C10'（首行需是表头）。
        style: 表样式，如 'medium2'(默认,蓝)/'medium4'(绿)/'light1'(浅)。
        total_row: 是否显示汇总行。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_list_table", "xlsx", "Excel 表格只在 Excel 中可用")
    try:
        return _tool().add_list_table(  # type: ignore[union-attr]
            sheet, cell_range, style=style, total_row=total_row,
        )
    except OfficeCLIError as e:
        return f"添加表格失败: {e}"


@tool
def add_dropdown(sheet: str, cell_range: str, options: str) -> str:
    """【Excel 专属】给单元格加下拉列表（数据验证）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'B2:B100'。
        options: 选项，逗号分隔，如 '是,否,待定'。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_dropdown", "xlsx", "下拉列表只在 Excel 中可用")
    try:
        return _tool().add_validation(  # type: ignore[union-attr]
            sheet, cell_range, "list", formula1=options, in_cell_dropdown=True,
        )
    except OfficeCLIError as e:
        return f"添加下拉列表失败: {e}"


@tool
def merge_cells(sheet: str, cell_range: str) -> str:
    """【Excel 专属】合并单元格。cell_range 如 'A1:D1'（左上格为锚点）。"""
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("merge_cells", "xlsx", "合并单元格只在 Excel 中可用")
    try:
        return _tool().merge_cells(sheet, cell_range)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"合并单元格失败: {e}"


@tool
def set_column_width(sheet: str, col: str, width: float) -> str:
    """【Excel 专属】设置列宽。

    参数:
        sheet: 工作表名。
        col: 列字母（如 'A'）。
        width: 宽度（字符单位，如 20）。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_column_width", "xlsx", "列宽只在 Excel 中可设")
    try:
        return _tool().set_column_width(sheet, col, width)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"设置列宽失败: {e}"


@tool
def autofit_column(sheet: str, col: str) -> str:
    """【Excel 专属】自动调整列宽以适应内容。

    参数:
        sheet: 工作表名。
        col: 列字母（如 'A'）。也可传 'A:C' 一次调整多列。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("autofit_column", "xlsx", "自动列宽只在 Excel 中可用")
    try:
        return _tool().autofit_column(sheet, col)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"自动列宽失败: {e}"


@tool
def rename_sheet(old_name: str, new_name: str) -> str:
    """【Excel 专属】重命名工作表。"""
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("rename_sheet", "xlsx", "工作表只在 Excel 中有")
    try:
        return _tool().rename_sheet(old_name, new_name)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"重命名失败: {e}"


# ============================================================
# PowerPoint 进阶工具（pptx）—— 均与占位符模式兼容，不会重叠
# ============================================================
@tool
def set_slide_transition(slide_index: int, transition: str,
                         auto_advance_seconds: float | None = None) -> str:
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
        return _tool().set_transition(  # type: ignore[union-attr]
            slide_index, transition, advance_time_ms=advance_ms,
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
        return _tool().set_notes(slide_index, notes)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"设置备注失败: {e}"


@tool
def set_theme_colors(accent1: str = "", accent2: str = "",
                     accent3: str = "", hyperlink: str = "") -> str:
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
        return _tool().set_theme_colors(  # type: ignore[union-attr]
            accent1=accent1, accent2=accent2, accent3=accent3, hyperlink=hyperlink,
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
        return _tool().set_theme_fonts(  # type: ignore[union-attr]
            heading_font=heading_font, body_font=body_font,
        )
    except OfficeCLIError as e:
        return f"设置主题字体失败: {e}"


# ============================================================
# 车辆号牌查询（交通类文档专项）
# ============================================================
@tool
def query_vehicle(plate_number: str) -> dict:
    """根据车牌号查询车辆的详细信息（基本信息/所有人/图片/事故/违法）。

    【何时调用】生成交通事故报告、车辆评估、车险理赔等交通类文档时，
    在用 ask_user 收集到车牌号后，【逐个】调用本工具查询每辆车的详细信息，
    再把查回的信息写入文档。

    返回 dict，status 字段标识查询结果（关键，据此决定后续动作）:
      - status="ok": 唯一匹配。含 vehicle(基本信息+所有人)、image_url(车辆照片URL)、
                     accidents(事故记录列表)、violations(违法记录列表)、stats(统计)。
                     → 直接用这些信息写文档；image_url 可传给 add_image 插入照片。
      - status="multiple": 多辆匹配。含 candidates(候选清单，每项有 id/plate/owner/brand)。
                     → 需调 ask_user 让用户从候选中选择（options 用候选的简要描述），
                       用户选定后，用所选候选的完整信息继续。
      - status="not_found": 无匹配。→ 告知用户该车牌无记录，可让其重新提供。

    参数:
        plate_number: 车牌号（如 "京A12345"）。
    """
    from .vehicle_data import query
    return query(plate_number)


# ============================================================
# 交互与控制（非常规 officecli 操作）
# ============================================================
class AskField(BaseModel):
    """表单中的一个字段。"""

    key: str = Field(
        ..., description="字段标识，英文蛇形（如 time/location），回传答案用此 key"
    )
    label: str = Field(
        ..., description="字段的中文显示标签（如 '事故时间'）"
    )
    required: bool = Field(
        False, description="是否必填。关键信息设 true，可推断或缺省的字段设 false。"
    )
    options: list[str] = Field(
        default_factory=list,
        description="候选选项（0-4 个）。枚举型字段（如责任认定）尽量提供；"
        "自由文本字段（如事故经过）留空。",
    )
    hint: str = Field(
        "", description="输入提示/示例（如 '如 2025年6月10日 14:30'）。可空。"
    )


@tool
def ask_user(
    title: str,
    fields: list[AskField],
    description: str = "",
) -> dict:
    """当缺少【关键】信息、存在歧义、或需要用户做选择时，向用户采集信息并等待回答。

    【何时调用】仅当某个信息缺失会导致文档无法生成或严重偏离预期时才用。
    能用合理假设推断的（如没说字数就用默认篇幅），【不要】调本工具。

    【表单模式（推荐）】一次提交多个相关字段，用户逐个填写，体验好。
    例如写交通事故报告缺信息时：
        title="交通事故信息采集"
        fields=[
            {key:"time", label:"事故时间", required:true, hint:"如 2025-06-10 14:30"},
            {key:"location", label:"事故地点", required:true},
            {key:"vehicles", label:"涉事车辆(车型/车牌/驾驶人)", required:true,
             hint:"可多辆，换行分隔"},
            {key:"injury", label:"人员伤亡", required:false, options:["无","轻伤","重伤","死亡"]},
            {key:"liability", label:"责任认定", required:false,
             options:["全责","主责","同责","次责","无责"]},
        ]
    【单问题模式】只问一个问题：fields 只放一个字段即可。

    【字段设计原则】
    - 枚举型字段（责任认定、严重程度、优先级...）尽量给 options，减少用户打字。
    - 自由文本字段（经过描述、地址...）options 留空。
    - 必填(true)仅用于缺失会导致文档偏离的字段；能缺省的字段设 false。

    参数:
        title: 卡片标题（简洁，如 "交通事故信息采集"）。
        fields: 字段列表（1-8 个）。
        description: 卡片说明（可选，简短解释为何需要这些信息）。

    返回: dict，key 是字段 key，value 是用户输入（候选选项已映射为文本）。"""
    payload = {
        "title": title,
        "description": description,
        "fields": [f.model_dump() for f in fields],
    }
    # interrupt 挂起执行；main.py 收集用户输入后 Command(resume=dict) 恢复，
    # resume 的 dict 会作为本函数的返回值传回给 agent。
    answer = interrupt(payload)
    # 兜底：旧版可能返回 str，统一成 dict
    if isinstance(answer, str):
        return {"value": answer.strip()} if answer.strip() else {}
    return answer if isinstance(answer, dict) else {}


@tool
def finish(summary: str) -> str:
    """宣告文档生成完成。当你确认文档结构完整、内容已写好、且自查无误后调用本工具。
    参数:
        summary: 一句话总结你生成了什么文档（会展示给用户）。"""
    return f"FINISHED: {summary}"


# ============================================================
# 工具清单（供 graph.py 绑定）
# ============================================================
ALL_TOOLS = [
    # 通用
    create_doc,
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
