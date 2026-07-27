"""通用工具（三格式共用）+ 控制（ask_user/finish）+ 公文（start_from_template/query_vehicle）。"""

from __future__ import annotations

import logging
import os
from urllib.error import URLError
from urllib.request import Request, urlopen

from langchain_core.tools import tool
from pydantic import BaseModel, Field

from office_agent.domain.templates import (
    DOC_TYPE_NAMES,
    default_merge_data,
    template_exists,
    template_path,
)
from office_agent.officecli import (
    OfficeCLIError,
    PptxTool,
    merge_template,
)
from office_agent.tools import (
    _tool,
    session_doc_kind,
    session_doc_path,
)

logger = logging.getLogger(__name__)


# 图片来源预校验超时（秒）。仅用于 HEAD 探测 URL 是否可达，不影响 officecli 实际下载。
_IMAGE_HEAD_TIMEOUT = 8


def _validate_image_source(src: str) -> str | None:
    """校验图片来源是否可用，在碰文档前拦掉坏源。

    返回 None 表示有效（可继续插入）；返回 str 表示失败原因（供工具层反馈给 LLM）。

    设计原则【宁可放行，不可误拦】：
        - 本地路径：os.path.exists 硬检查，可靠无歧义 → 坚决拦。
        - data URI：内联数据，总是可用 → 放行。
        - HTTP/HTTPS：发 HEAD 探测。明确 404/410、连接失败/DNS 失败/超时 → 拦；
          405（方法不允许）或其它不确定响应 → 放行交给 officecli 实际 GET
          （有些服务器不支持 HEAD 但 GET 能取，避免误伤）。

    这样主流坏源（本地不存在、404、域名打不开）在碰文档前被拦掉、零副作用；
    边缘情况（HEAD 不准）由 officecli 兜底，失败时 doc_tool 的事后清理兜底会删载段。
    """
    if not src or not src.strip():
        return "图片来源为空"
    s = src.strip()

    # data URI：内联，总是可用
    if s.startswith("data:"):
        return None

    # HTTP/HTTPS：HEAD 探测可达性
    if s.startswith(("http://", "https://")):
        try:
            req = Request(s, method="HEAD")
            with urlopen(req, timeout=_IMAGE_HEAD_TIMEOUT) as resp:
                code = getattr(resp, "status", None) or resp.getcode()
                # 明确的"不存在"类状态码 → 拦
                if code in (404, 410):
                    return f"图片不存在（HTTP {code}）"
                # 其它（2xx/3xx/405/...）→ 放行
                return None
        except URLError as e:
            # HTTPError（4xx/5xx 有响应）也是 URLError 子类
            code = getattr(e, "code", None)
            if code in (404, 410):
                return f"图片不存在（HTTP {code}）"
            if code in (405,) or "Method Not Allowed" in str(e):
                # HEAD 不被支持 → 放行，让 officecli 实际 GET
                return None
            # 连接失败 / DNS 失败 / 超时 / 拒绝 → 拦
            return f"图片不可访问（{e.reason}）"
        except Exception as e:  # noqa: BLE001
            # 未知异常 → 放行（不误拦），交给 officecli 兜底
            logger.debug("图片源 HEAD 探测异常（放行交给 officecli）: %s", e)
            return None

    # 本地路径：硬检查存在性
    if not os.path.exists(s):
        return f"文件不存在: {s}"
    return None


class AskField(BaseModel):
    """表单中的一个字段。"""

    key: str = Field(..., description="字段标识，英文蛇形（如 time/location），回传答案用此 key")
    label: str = Field(..., description="字段的中文显示标签（如 '事故时间'）")
    required: bool = Field(
        False, description="是否必填。关键信息设 true，可推断或缺省的字段设 false。"
    )
    options: list[str] = Field(
        default_factory=list,
        description="候选选项（0-4 个）。枚举型字段（如责任认定）尽量提供；"
        "自由文本字段（如事故经过）留空。",
    )
    hint: str = Field("", description="输入提示/示例（如 '如 2025年6月10日 14:30'）。可空。")


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
        for row in data or []:
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


@tool
def start_from_template(
    doc_type: str,
    org: str = "",
    doc_no: str = "",
    signer: str = "",
    date_cn: str = "",
    title: str = "",
    addressee: str = "",
) -> str:
    """【公文模式专用】从法定公文模板创建会话文档，替代 create_doc。

    【何时调用】用户要写法定公文时调用本工具，【不要】调 create_doc。
    模板已含 GB/T 9704 标准版头（发文机关标志/红色分隔线/发文字号/签发人/
    标题/主送/正文范例/落款/版记/页码），调用后只需编辑正文范例文字即可，
    不必从零拼接（从零拼接会丢失规范版式）。

    【可选文种 doc_type 取值】15 个法定文种：决议、决定、命令（令）、公报、
    公告、通告、意见、通知、通报、报告、请示、批复、议案、函、纪要。
    （完整适用情形见系统提示词的公文文种清单。）

    参数:
        doc_type: 文种名，必须是上面 15 个之一（如 '通知'）。
        org: 发文机关（版头红字 + 落款署名），如 '北京市公安局'。
        doc_no: 发文字号，如 '京公发〔2026〕12号'。命令（令）用 '第 X 号'。
        signer: 签发人姓名。【仅上行文】（请示/报告/议案）需要，其他文种忽略。
        date_cn: 成文日期（中文数字），如 '二〇二六年三月三十一日'。
        title: 公文标题。【不参与 merge 预填】——模板里的范例标题需你调用本工具后
               用 update_paragraph(path='/body/p[4]', text=<这个 title>) 手动替换。
               留空则保留范例标题（之后也可改）。
        addressee: 主送机关。【不参与 merge】——同理用 update_paragraph('/body/p[5]', ...)
               替换范例主送。留空则保留范例主送。

    【关于 title/addressee】这俩不在 merge 槽位里（模板里是范例文字，供你参考
    公文标题的拟写格式）。之所以不自动替换，是因为不同文种标题段位置略有差异，
    自动定位不可靠——由你 view_text 确认路径后用 update_paragraph 改更稳。

    【调用后必做】
    1. view_text 查看模板结构（确认标题段 /body/p[4]、主送段 /body/p[5] 位置）
    2. update_paragraph 改标题、主送（若 title/addressee 传了）
    3. replace_text 把正文里的 'XX' 占位换成真实内容（保字体）
    4. remove_paragraph 删多余范例段、add_paragraph 补新段
    5. view_text 自查 → finish

    【注意】只在 Word 会话下可用。模板正文里的 'XX' 是范例占位，
    需要你逐处替换成具体内容；版头槽位由本工具一次性预填好。
    """
    kind = session_doc_kind()
    if kind != "docx":
        return f"start_from_template 只在 Word 会话下可用，当前是 {kind}。公文模板只支持 Word。"

    if not session_doc_path():
        return "会话文档路径未初始化。"

    # 1) 校验文种 + 模板存在
    if not template_exists(doc_type):
        return f"文种 '{doc_type}' 不存在或模板缺失。合法文种: {DOC_TYPE_NAMES}"
    tmpl = str(template_path(doc_type))

    # 2) 构造 merge 数据（overrides > 文种默认 > 全局默认）
    overrides = {
        "org": org,
        "doc_no": doc_no,
        "signer": signer,
        "date_cn": date_cn,
        "signer_org": org,  # 落款默认与发文机关一致
    }
    overrides = {k: v for k, v in overrides.items() if v and v.strip()}
    try:
        merge_data = default_merge_data(doc_type, **overrides)
    except ValueError as e:
        return f"文种数据构造失败: {e}"

    # 3) merge 模板 → 会话输出路径（一步完成复制 + 预填）
    try:
        result = merge_template(tmpl, session_doc_path() or "", merge_data)
    except OfficeCLIError as e:
        return f"从模板创建失败: {e}"

    replaced = sum(1 for v in merge_data.values() if v)
    # title/addressee 提示：明确告知路径，引导 LLM 下一步用 update_paragraph
    hints = ["下一步：view_text 看结构"]
    if title:
        hints.append(f"用 update_paragraph('/body/p[4]', '{title}') 替换范例标题")
    if addressee:
        hints.append(f"用 update_paragraph('/body/p[5]', '{addressee}') 替换范例主送")
    if not title and not addressee:
        hints.append("用 replace_text/​update_paragraph 编辑正文范例")
    return (
        f"已从《{doc_type}》模板创建会话文档: {session_doc_path()}\n"
        f"预填 {replaced} 个版头槽位。{result.strip()}\n"
        f"{'; '.join(hints)}。"
    )


# ============================================================
# Word 专属工具（docx）
# ============================================================


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
    # 预校验图片来源：坏源（404/文件不存在/域名打不开）在碰文档前跳过，
    # 不嵌入、不留空段、不浪费 officecli 调用。
    reason = _validate_image_source(url_or_path)
    if reason:
        return (
            f"⚠️ 跳过插入图片（{reason}）。"
            f"不要重试这张图，继续生成文档其他内容。"
        )
    kind = session_doc_kind()
    try:
        t = _tool()
        if kind == "pptx":
            pptx: PptxTool = t  # type: ignore[assignment]
            slide_index = pptx._last_slide_index() or 1
            return pptx.add_image(slide_index, url_or_path, width=width, alt=caption or "图片")
        # docx
        return t.add_image(
            url_or_path,
            width=width,  # type: ignore[union-attr]
            alt=caption or "图片",
            caption=caption,
        )
    except OfficeCLIError as e:
        return f"插入图片失败: {e}"


# ============================================================
# Excel 专属工具（xlsx）
# ============================================================


@tool
def set_doc_properties(
    title: str = "", author: str = "", subject: str = "", keywords: str = ""
) -> str:
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
                title=title,
                author=author,
                subject=subject,
            )
        # docx / xlsx 都用 set on /
        return t.set_doc_properties(  # type: ignore[union-attr]
            title=title,
            author=author,
            subject=subject,
            keywords=keywords,
        )
    except OfficeCLIError as e:
        return f"设置文档属性失败: {e}"


# ============================================================
# Excel 进阶工具（xlsx）
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
    from office_agent.domain.vehicle_data import query

    return query(plate_number)


# ============================================================
# 交互与控制（非常规 officecli 操作）
# ============================================================
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

    返回: 结构化交互请求。Graph 会挂起并把用户答案以 dict 形式回传给 Agent。"""
    return {
        "title": title,
        "description": description,
        "fields": [f.model_dump() for f in fields],
    }


@tool
def finish(summary: str) -> str:
    """宣告文档生成完成。当你确认文档结构完整、内容已写好、且自查无误后调用本工具。
    参数:
        summary: 一句话总结你生成了什么文档（会展示给用户）。"""
    return f"FINISHED: {summary}"


# ============================================================
# 工具清单（供 graph.py 绑定）
# ============================================================
