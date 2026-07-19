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

from typing import Annotated, Any

from langchain_core.tools import tool
from langgraph.types import interrupt
from pydantic import BaseModel, Field

from .officecli import DocTool, OfficeCLIError

# ============================================================
# 会话状态（模块级，agent 运行期间唯一）
# ============================================================
_session_doc_path: str | None = None


def set_session_doc(path: str) -> None:
    """main.py 在启动 agent 前调用，设定本会话的文档路径。"""
    global _session_doc_path
    _session_doc_path = path


def _doc() -> DocTool:
    if _session_doc_path is None:
        raise OfficeCLIError("会话文档路径未初始化（请先调用 set_session_doc）")
    return DocTool(_session_doc_path)


def session_doc_path() -> str | None:
    return _session_doc_path


# ============================================================
# 文档生命周期
# ============================================================
@tool
def create_doc() -> str:
    """创建一个新的空白 Word 文档（.docx）。这是生成文档时【必须第一个调用】的工具，
    会覆盖同名旧文档。调用一次即可，之后用 add_* 工具往里加内容。"""
    return _doc().create()


@tool
def add_title(text: str) -> str:
    """添加文档主标题（居中、大字号、加粗）。整篇文档只调一次。
    参数:
        text: 主标题文字，如"2024年第三季度销售报告"。"""
    return _doc().add_title(text)


@tool
def add_heading(text: str, level: int = 1) -> str:
    """添加一个章节标题。
    参数:
        text: 标题文字（不要带编号，如用"项目背景"而非"一、项目背景"）。
        level: 标题层级 1-9，1=一级标题（最大），2=二级子标题。默认 1。"""
    return _doc().add_heading(text, level=level)


@tool
def add_paragraph(text: str, bold: bool = False, italic: bool = False) -> str:
    """添加一段正文。
    参数:
        text: 段落正文（纯文本，不要 markdown 标记，不要在一段里塞多个主题）。
        bold: 是否整段加粗（默认 false，仅强调时用）。
        italic: 是否斜体（默认 false）。"""
    return _doc().add_paragraph(text, bold=bold, italic=italic)


@tool
def add_list_item(text: str, ordered: bool = False) -> str:
    """添加一个列表项（项目符号或编号）。
    参数:
        text: 单个列表项的文字。
        ordered: true=有序编号列表(1. 2. 3.)，false=无序项目符号(•)。默认 false。
    连续多次调用本工具即可构成一个完整列表。"""
    return _doc().add_list_item(text, ordered=ordered)


@tool
def add_table(data: list[list], has_header: bool = True) -> str:
    """添加一个表格。适合展示对比、数据、结构化信息。
    参数:
        data: 二维数组，外层是行、内层是单元格。每行长度应一致。
              单元格可以是字符串或数字（内部自动转字符串）。
              例: [["姓名","成绩"],["张三",95],["李四",88]]。
        has_header: 第一行是否作为表头（加粗）。默认 true。"""
    try:
        # 防御性清洗：null 转空串、数字转字符串、确保二维结构
        clean = []
        for row in (data or []):
            if row is None:
                continue
            clean.append([("" if c is None else str(c)) for c in row])
        if not clean:
            return "添加表格失败: 数据为空"
        _doc().add_table(clean, has_header=has_header)
        rows = len(clean)
        cols = max(len(r) for r in clean) if clean else 0
        return f"已添加 {rows} 行 × {cols} 列的表格"
    except OfficeCLIError as e:
        return f"添加表格失败: {e}"


# ============================================================
# 文档读取（供 LLM 自查）
# ============================================================
@tool
def view_text() -> str:
    """读取当前文档的全部纯文本内容（带段落路径标注）。
    建议在调 finish 之前调用一次，自查文档结构和内容是否正确、完整。"""
    return _doc().view_text()


@tool
def validate_doc() -> str:
    """校验当前文档是否符合 OpenXML 规范。返回校验结果。
    一般无需调用；若怀疑文档有问题时可用来确认。"""
    return _doc().validate()


@tool
def add_image(url_or_path: str, width: str = "8cm", caption: str = "") -> str:
    """在文档末尾插入一张图片（如车辆照片）。

    参数:
        url_or_path: 图片来源。支持本地文件路径、HTTP/HTTPS URL、data URI。
        width: 显示宽度，如 '8cm'/'400px'/'3in'（默认 8cm）。
        caption: 可选图注文字。非空时在图片下方显示。

    用途: 生成交通事故/车辆评估等文档时，用 query_vehicle 查回的 image_url
    调本工具插入车辆照片。"""
    try:
        return _doc().add_image(url_or_path, width=width, alt=caption or "图片", caption=caption)
    except OfficeCLIError as e:
        return f"插入图片失败: {e}"


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
    create_doc,
    add_title,
    add_heading,
    add_paragraph,
    add_list_item,
    add_table,
    add_image,
    view_text,
    validate_doc,
    query_vehicle,
    ask_user,
    finish,
]

# 工具名 -> 工具对象，便于 tools 节点按名分发
TOOL_BY_NAME = {t.name: t for t in ALL_TOOLS}
