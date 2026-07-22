"""暴露给 LLM 的工具集（@tool 装饰器）。

设计:
    - 会话级文档路径注入：main.py 启动 agent 前调用 set_session_doc(path)，
      所有工具内部读取该路径，LLM 不需要传路径参数（避免出错）。
    - 工具方法直接复用 officecli.DocTool（已实测验证的正确命令）。
    - ask_user 工具内部用 LangGraph interrupt 挂起，等用户输入后作为
      ToolMessage 回传给 agent。
    - finish 工具让 LLM 显式宣告完成。

每个工具的 docstring 是 LLM 判断"何时调用"的依据，写得具体清晰。
"""

from __future__ import annotations

from typing import Any

from langchain_core.tools import tool
from langgraph.types import interrupt
from pydantic import BaseModel, Field

from .format_detect import OfficeFormat
from .officecli import DocTool, OfficeCLIError, SheetTool, SlideTool

# ============================================================
# 会话状态（模块级，agent 运行期间唯一）
# ============================================================
_session_doc_path: str | None = None
_session_format: OfficeFormat = "docx"


def set_session_doc(path: str) -> None:
    """main.py 在启动 agent 前调用，设定本会话的文档路径。"""
    global _session_doc_path
    _session_doc_path = path


def set_session_format(fmt: OfficeFormat) -> None:
    """设定本会话输出格式（docx / xlsx / pptx）。"""
    global _session_format
    _session_format = fmt


def session_format() -> OfficeFormat:
    return _session_format


def _require_path() -> str:
    if _session_doc_path is None:
        raise OfficeCLIError("会话文档路径未初始化（请先调用 set_session_doc）")
    return _session_doc_path


def _doc() -> DocTool:
    return DocTool(_require_path())


def _sheet() -> SheetTool:
    return SheetTool(_require_path())


def _slide() -> SlideTool:
    return SlideTool(_require_path())


def session_doc_path() -> str | None:
    return _session_doc_path


def _cli_call(action: str, fn) -> str:
    """统一捕获 OfficeCLIError，返回可读字符串给 LLM。"""
    try:
        result = fn()
        return result if isinstance(result, str) else str(result)
    except OfficeCLIError as e:
        return f"{action}失败: {e}"


# ============================================================
# 文档生命周期
# ============================================================
@tool
def create_doc() -> str:
    """创建一个新的空白 Word 文档（.docx）。这是生成文档时【必须第一个调用】的工具，
    会覆盖同名旧文档。调用一次即可，之后用 add_* 工具往里加内容。"""
    return _cli_call("创建文档", lambda: _doc().create())


@tool
def add_title(text: str) -> str:
    """添加文档主标题（居中、大字号、加粗）。整篇文档只调一次。
    参数:
        text: 主标题文字，如"2024年第三季度销售报告"。"""
    return _cli_call("添加标题", lambda: _doc().add_title(text))


@tool
def add_heading(text: str, level: int = 1) -> str:
    """添加一个章节标题。
    参数:
        text: 标题文字（不要带编号，如用"项目背景"而非"一、项目背景"）。
        level: 标题层级 1-9，1=一级标题（最大），2=二级子标题。默认 1。"""
    return _cli_call("添加章节标题", lambda: _doc().add_heading(text, level=level))


@tool
def add_paragraph(
    text: str,
    bold: bool = False,
    italic: bool = False,
    align: str = "",
) -> str:
    """添加一段正文。
    参数:
        text: 段落正文（纯文本，不要 markdown 标记，不要在一段里塞多个主题）。
        bold: 是否整段加粗（默认 false，仅强调时用）。
        italic: 是否斜体（默认 false）。
        align: 对齐方式，可选 left/center/right/justify；空=默认左对齐。"""
    return _cli_call(
        "添加段落",
        lambda: _doc().add_paragraph(
            text, bold=bold, italic=italic, align=align or None,
        ),
    )


@tool
def add_list_item(text: str, ordered: bool = False, level: int = 0) -> str:
    """添加一个列表项（项目符号或编号）。
    参数:
        text: 单个列表项的文字。
        ordered: true=有序编号列表(1. 2. 3.)，false=无序项目符号(•)。默认 false。
        level: 嵌套层级 0-8，0=顶层。默认 0。
    连续多次调用本工具即可构成一个完整列表。"""
    return _cli_call(
        "添加列表项",
        lambda: _doc().add_list_item(text, ordered=ordered, level=level),
    )


@tool
def add_table(
    data: list[list],
    has_header: bool = True,
    style: str = "medium2",
) -> str:
    """添加一个表格。适合展示对比、数据、结构化信息。
    参数:
        data: 二维数组，外层是行、内层是单元格。每行长度应一致。
              单元格可以是字符串或数字（内部自动转字符串）。
              例: [["姓名","成绩"],["张三",95],["李四",88]]。
        has_header: 第一行是否作为表头（加粗）。默认 true。
        style: 表格样式名，如 medium2/light1/dark1；传空字符串表示不用样式。默认 medium2。"""
    def _run() -> str:
        clean = []
        for row in (data or []):
            if row is None:
                continue
            clean.append([("" if c is None else str(c)) for c in row])
        if not clean:
            raise OfficeCLIError("数据为空")
        return _doc().add_table(
            clean,
            has_header=has_header,
            style=style or None,
        )

    return _cli_call("添加表格", _run)


@tool
def add_image(
    url_or_path: str,
    width: str = "8cm",
    height: str = "",
    caption: str = "",
) -> str:
    """在文档末尾插入一张图片（如车辆照片）。

    参数:
        url_or_path: 图片来源。支持本地文件路径、HTTP/HTTPS URL、data URI。
        width: 显示宽度，必须带单位，如 '8cm'/'400px'/'3in'（默认 8cm）。
        height: 可选高度，同样必须带单位；空=按比例缩放。
        caption: 可选图注文字。非空时在图片下方居中显示。

    用途: 生成交通事故/车辆评估等文档时，用 query_vehicle 查回的 image_url
    调本工具插入车辆照片。"""
    return _cli_call(
        "插入图片",
        lambda: _doc().add_image(
            url_or_path,
            width=width,
            height=height or "",
            alt=caption or "图片",
            caption=caption,
        ),
    )


@tool
def add_header(
    text: str = "",
    align: str = "center",
    page_numbers: bool = False,
) -> str:
    """添加页眉。正式报告可在 create_doc 之后尽早调用一次。
    参数:
        text: 页眉文字（page_numbers=true 时作为前缀，可空）。
        align: 对齐 left/center/right，默认 center。
        page_numbers: true 时在页眉加入「第 X 页 / 共 Y 页」。"""
    return _cli_call(
        "添加页眉",
        lambda: _doc().add_header(
            text, align=align, page_numbers=page_numbers,
        ),
    )


@tool
def add_footer(
    text: str = "",
    align: str = "center",
    page_numbers: bool = False,
) -> str:
    """添加页脚。正式报告建议开启 page_numbers。
    参数:
        text: 页脚文字（page_numbers=true 时作为前缀，可空）。
        align: 对齐 left/center/right，默认 center。
        page_numbers: true 时在页脚加入「第 X 页 / 共 Y 页」。"""
    return _cli_call(
        "添加页脚",
        lambda: _doc().add_footer(
            text, align=align, page_numbers=page_numbers,
        ),
    )


@tool
def add_page_break() -> str:
    """在文档当前位置插入分页符。大章节开始前可调用，让新章从新页起。"""
    return _cli_call("插入分页符", lambda: _doc().add_page_break())


@tool
def replace_text(find: str, replace: str) -> str:
    """在正文中查找并替换文字（字面量匹配）。
    【何时调用】用户补充/纠正用词、称呼、数据时，用本工具局部修改，
    【不要】整章删掉重写。
    参数:
        find: 要查找的原文（不能为空）。
        replace: 替换成的新文字（可空，表示删除匹配）。"""
    return _cli_call(
        "查找替换",
        lambda: _doc().replace_text(find, replace),
    )


@tool
def batch_add(ops: list[dict]) -> str:
    """一次性批量写入同一章节的多段内容（一次 subprocess，更快）。
    【优先】写同一章的标题+多段+列表时用本工具，而不是多次单独 add_*。

    参数:
        ops: 操作列表。每项是对象，必须含 op 字段：
          - {"op":"title","text":"..."}
          - {"op":"heading","text":"...","level":1}
          - {"op":"paragraph","text":"...","bold":false,"italic":false,"align":""}
          - {"op":"list_item","text":"...","ordered":false,"level":0}
          - {"op":"page_break"}
        表格和图片请仍用 add_table / add_image（不能放进本工具）。"""
    return _cli_call("批量写入", lambda: _doc().batch_add(ops or []))


# ============================================================
# 文档读取（供 LLM 自查）
# ============================================================
@tool
def view_outline() -> str:
    """读取当前文档大纲（标题/幻灯片结构）。自查时【优先】用本工具，比 view_text 更轻更快。
    若大纲看起来缺章/乱序，再调 view_text 看正文细节。"""
    def _run() -> str:
        if session_format() == "pptx":
            return _slide().view_outline()
        return _doc().view_outline()

    return _cli_call("读取大纲", _run)


@tool
def view_text() -> str:
    """读取当前文档的全部纯文本内容（带路径标注）。
    自查细节、核对措辞时使用；日常结构检查优先 view_outline（Word/PPT）。"""
    def _run() -> str:
        fmt = session_format()
        if fmt == "pptx":
            return _slide().view_text()
        if fmt == "xlsx":
            return _sheet().view_text()
        return _doc().view_text()

    return _cli_call("读取正文", _run)


@tool
def validate_doc() -> str:
    """校验当前文档是否符合 OpenXML 规范。返回校验结果。
    一般无需调用；若怀疑文档有问题时可用来确认。"""
    def _run() -> str:
        fmt = session_format()
        if fmt == "pptx":
            return _slide().validate()
        if fmt == "xlsx":
            return _sheet().validate()
        return _doc().validate()

    return _cli_call("校验文档", _run)


# ============================================================
# Excel MVP
# ============================================================
@tool
def create_workbook() -> str:
    """创建一个新的空白 Excel 工作簿（.xlsx）。生成表格时【必须第一个调用】，只调一次。"""
    return _cli_call("创建工作簿", lambda: _sheet().create())


@tool
def add_sheet(name: str) -> str:
    """新增一个工作表（sheet）。
    参数:
        name: 工作表名称，如"销售汇总"。新建工作簿默认已有 Sheet1，需要更多表时再调。"""
    return _cli_call("添加工作表", lambda: _sheet().add_sheet(name))


@tool
def write_range(
    sheet: str,
    start: str,
    data: list[list],
    header_bold: bool = True,
) -> str:
    """从指定单元格起写入一块二维数据（批量）。适合表头+多行数据。
    参数:
        sheet: 工作表名，默认常用 "Sheet1"。
        start: 起始单元格，如 "A1"。
        data: 二维数组，外层行、内层单元格。例: [["区域","金额"],["华东",100],["华北",80]]。
        header_bold: 第一行是否加粗（表头）。默认 true。"""
    def _run() -> str:
        clean = []
        for row in (data or []):
            if row is None:
                continue
            clean.append([("" if c is None else str(c)) for c in row])
        if not clean:
            raise OfficeCLIError("数据为空")
        return _sheet().write_range(
            sheet or "Sheet1",
            start or "A1",
            clean,
            header_bold=header_bold,
        )

    return _cli_call("写入区域", _run)


@tool
def write_cell(
    sheet: str,
    ref: str,
    value: str = "",
    formula: str = "",
    bold: bool = False,
) -> str:
    """写入单个单元格的值或公式。
    参数:
        sheet: 工作表名，如 "Sheet1"。
        ref: 单元格引用，如 "B10"。
        value: 字面量（与 formula 二选一；有 formula 时忽略 value）。
        formula: Excel 公式，可不带前导 =，如 "SUM(B2:B9)"。
        bold: 是否加粗。"""
    return _cli_call(
        "写入单元格",
        lambda: _sheet().write_cell(
            sheet or "Sheet1",
            ref,
            value=value,
            formula=formula,
            bold=bold,
        ),
    )


@tool
def view_sheet() -> str:
    """读取当前工作簿的文本视图（各 sheet 内容）。写完数据后建议自查一次再 finish。"""
    return _cli_call("读取工作簿", lambda: _sheet().view_text())


# ============================================================
# PowerPoint MVP
# ============================================================
@tool
def create_presentation() -> str:
    """创建一个新的空白 PowerPoint 演示文稿（.pptx）。生成 PPT 时【必须第一个调用】，只调一次。"""
    return _cli_call("创建演示文稿", lambda: _slide().create())


@tool
def add_slide(title: str = "", body: str = "", layout: str = "") -> str:
    """新增一页幻灯片。
    参数:
        title: 页面标题。
        body: 正文/概述（可选；要点列表更推荐随后用 add_bullets）。
        layout: 可选布局名（如 "Title Slide"）；空=默认。"""
    return _cli_call(
        "添加幻灯片",
        lambda: _slide().add_slide(title=title, body=body, layout=layout or ""),
    )


@tool
def add_bullets(slide_index: int, items: list[str]) -> str:
    """在指定页添加要点列表（项目符号）。
    参数:
        slide_index: 幻灯片序号，从 1 开始（第一页为 1）。
        items: 要点文字列表，如 ["进展A", "风险B"]。"""
    return _cli_call(
        "添加要点",
        lambda: _slide().add_bullets(slide_index, items or []),
    )


@tool
def add_slide_table(slide_index: int, data: list[list]) -> str:
    """在指定页插入表格。
    参数:
        slide_index: 幻灯片序号（从 1 开始）。
        data: 二维数组，如 [["指标","数值"],["完成率","92%"]]。"""
    def _run() -> str:
        clean = []
        for row in (data or []):
            if row is None:
                continue
            clean.append([("" if c is None else str(c)) for c in row])
        if not clean:
            raise OfficeCLIError("数据为空")
        return _slide().add_table(slide_index, clean)

    return _cli_call("添加幻灯片表格", _run)


@tool
def add_slide_image(
    slide_index: int,
    url_or_path: str,
    width: str = "12cm",
) -> str:
    """在指定页插入图片。
    参数:
        slide_index: 幻灯片序号（从 1 开始）。
        url_or_path: 本地路径或 URL。
        width: 宽度，必须带单位（如 12cm / 400px）。"""
    return _cli_call(
        "添加幻灯片图片",
        lambda: _slide().add_image(slide_index, url_or_path, width=width),
    )


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
# 工具清单（按格式绑定；TOOL_BY_NAME 含全部以便分发）
# ============================================================
_SHARED = [query_vehicle, ask_user, finish, validate_doc]

WORD_TOOLS = [
    create_doc,
    add_title,
    add_heading,
    add_paragraph,
    add_list_item,
    add_table,
    add_image,
    add_header,
    add_footer,
    add_page_break,
    replace_text,
    batch_add,
    view_outline,
    view_text,
    *_SHARED,
]

EXCEL_TOOLS = [
    create_workbook,
    add_sheet,
    write_range,
    write_cell,
    view_sheet,
    *_SHARED,
]

PPT_TOOLS = [
    create_presentation,
    add_slide,
    add_bullets,
    add_slide_table,
    add_slide_image,
    view_outline,
    view_text,
    *_SHARED,
]

# 向后兼容：默认 Word 工具集
ALL_TOOLS = WORD_TOOLS


def tools_for_format(fmt: OfficeFormat | str) -> list:
    """按会话格式返回可绑定的工具列表。"""
    if fmt == "xlsx":
        return EXCEL_TOOLS
    if fmt == "pptx":
        return PPT_TOOLS
    return WORD_TOOLS


# 工具名 -> 工具对象（三组合并，供 tools 节点按名分发）
TOOL_BY_NAME = {
    t.name: t for t in [*WORD_TOOLS, *EXCEL_TOOLS, *PPT_TOOLS]
}
