"""法定公文（15 文种）模板管理：文种识别、模板路径解析、merge 数据。

依据《党政机关公文处理工作条例》第八条规定的 15 个法定文种：
    决议、决定、命令（令）、公报、公告、通告、意见、通知、通报、
    报告、请示、批复、议案、函、纪要

模板文件位于 ``<project_root>/template/word/NN-文种.docx``，由
``scripts/gen_official_templates.py`` 生成，版头固定槽位使用 ``{{key}}``
占位（供 officecli merge 预填），正文为范例文字（供 LLM 编辑替换）。

设计:
    - ``OFFICIAL_DOCS``: 15 文种元数据（dataclass 列表），是唯一数据源。
    - ``detect_doc_type(requirement)``: 从用户需求关键词推断文种。
    - ``template_path(doc_type)``: 返回模板 .docx 路径。
    - ``default_merge_data(doc_type, **overrides)``: 构造 merge 用 JSON，
      未提供的槽位用合理默认值（绝不留 ``{{key}}`` 字面量）。
    - ``is_upward(doc_type)``: 是否上行文（决定是否需要签发人）。
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

from office_agent.config import settings

logger = logging.getLogger(__name__)

# 模板目录：工程根 / template / word
TEMPLATE_DIR: Path = settings.project_root / "template" / "word"


# ============================================================
# 文种元数据
# ============================================================
@dataclass(frozen=True)
class OfficialDocType:
    """单个法定文种的元数据。"""

    name: str  # 文种名（与文件名、用户输入一致），如 "通知"
    index: int  # 1-based 序号，对应文件名前缀 01/02/...
    direction: str  # 行向: 'upward'(上行)/'downward'(下行)/'parallel'(平行)/'meeting'(会议)
    summary: str  # 适用情形简述（给 LLM 看的提示）
    # 关键词：detect_doc_type 用，命中任一即判定为该文种。
    # 用词根避免误命中（如 "决定" 易与 "决定要做" 混淆，所以加 "作出决定" 等限定）。
    keywords: tuple[str, ...] = ()
    # 默认 merge 数据：用户没提供时用这些值填充版头槽位。
    # signer_org 默认与 org 一致（在 default_merge_data 里兜底）。
    defaults: dict[str, str] = field(default_factory=dict)

    @property
    def filename(self) -> str:
        return f"{self.index:02d}-{self.name}.docx"

    @property
    def path(self) -> Path:
        return TEMPLATE_DIR / self.filename


# ---- 各文种的默认版头数据 ----
# 统一占位机关名/字号/日期，用户没指定时用这些。
_DEF_ORG = "XX市XX机关"
_DEF_DOC_NO = "XX〔2026〕X号"
_DEF_DATE_CN = "二〇二六年X月X日"
_DEF_ISSUER = "XX市XX机关办公室"
_DEF_ISSUE_DATE = "2026年X月X日印发"
_DEF_SIGNER = "XXX"  # 上行文签发人占位


OFFICIAL_DOCS: list[OfficialDocType] = [
    OfficialDocType(
        name="决议",
        index=1,
        direction="meeting",
        summary="经会议讨论通过的重大决策事项",
        keywords=("决议",),
    ),
    OfficialDocType(
        name="决定",
        index=2,
        direction="downward",
        summary="重要事项作出决策和部署、奖惩有关单位和人员",
        # "决定"太通用，用 "作出决定/发布决定/关于...的决定" 限定
        keywords=("作出决定", "发布决定", "的决定", "给予处分", "给予嘉奖"),
    ),
    OfficialDocType(
        name="命令（令）",
        index=3,
        direction="downward",
        summary="公布行政法规和规章、宣布施行重大强制性措施、嘉奖",
        keywords=("命令", "发布令", "主席令", "市长令", "省长令", "公布令"),
    ),
    OfficialDocType(
        name="公报",
        index=4,
        direction="downward",
        summary="公布重要决定或者重大事项",
        keywords=("公报",),
    ),
    OfficialDocType(
        name="公告",
        index=5,
        direction="downward",
        summary="向国内外宣布重要事项或者法定事项",
        keywords=("公告",),
    ),
    OfficialDocType(
        name="通告",
        index=6,
        direction="downward",
        summary="在一定范围内公布应当遵守或者周知的事项",
        keywords=("通告",),
    ),
    OfficialDocType(
        name="意见",
        index=7,
        direction="downward",
        summary="对重要问题提出见解和处理办法",
        # "意见"太通用，限定为公文语境
        keywords=("实施意见", "指导意见", "处理意见", "的意见"),
    ),
    OfficialDocType(
        name="通知",
        index=8,
        direction="downward",
        summary="发布传达要求下级执行；批转转发公文；任免人员",
        keywords=("通知",),
    ),
    OfficialDocType(
        name="通报",
        index=9,
        direction="downward",
        summary="表彰先进、批评错误、传达重要精神和告知情况",
        keywords=("通报", "表彰", "表扬"),
    ),
    OfficialDocType(
        name="报告",
        index=10,
        direction="upward",
        summary="向上级汇报工作、反映情况、回复上级询问",
        # "报告"易与"调研报告/分析报告/实验报告"等非公文报告混淆，
        # 用 "向上级报告/工作报告/情况报告/汇报工作" 等公文语境限定。
        # 用户若想生成公文报告，请明确说"向上级的工作报告/情况报告"。
        keywords=("向上级报告", "工作报告", "情况报告", "汇报工作", "向上级汇报", "述职报告"),
    ),
    OfficialDocType(
        name="请示",
        index=11,
        direction="upward",
        summary="向上级请求指示、批准",
        keywords=("请示", "申请批准", "请求批准", "恳请", "妥否，请批示"),
    ),
    OfficialDocType(
        name="批复",
        index=12,
        direction="downward",
        summary="答复下级请示事项",
        # "批复/答复"出现时通常就是要写批复（答复某请示），即使文本里也有"请示"。
        # detect_doc_type 对"批复"做特殊优先处理。
        keywords=("批复", "答复请示", "回复请示", "同意请示", "答复下级"),
    ),
    OfficialDocType(
        name="议案",
        index=13,
        direction="upward",
        summary="各级人民政府按法律程序向同级人大或其常委会提请审议",
        keywords=("议案", "提请审议", "人大"),
    ),
    OfficialDocType(
        name="函",
        index=14,
        direction="parallel",
        summary="不相隶属机关之间商洽工作、询问和答复问题",
        keywords=("商请函", "公函", "询函", "复函", "商洽函", "的函"),
    ),
    OfficialDocType(
        name="纪要",
        index=15,
        direction="meeting",
        summary="记载会议主要情况和议定事项",
        keywords=("会议纪要", "纪要"),
    ),
]

# 文种名 → 元数据，O(1) 查找
DOC_BY_NAME: dict[str, OfficialDocType] = {d.name: d for d in OFFICIAL_DOCS}

# 所有合法文种名（含别名，给工具 docstring 用）
DOC_TYPE_NAMES: list[str] = [d.name for d in OFFICIAL_DOCS]

# 行向常量
UPWARD_TYPES = {d.name for d in OFFICIAL_DOCS if d.direction == "upward"}
MEETING_TYPES = {d.name for d in OFFICIAL_DOCS if d.direction == "meeting"}


def is_upward(doc_type: str) -> bool:
    """是否上行文（请示、报告、议案）。上行文版头需显示签发人。"""
    return doc_type in UPWARD_TYPES


def is_meeting(doc_type: str) -> bool:
    """是否会议类文种（决议、纪要）。通常无主送、无版记。"""
    return doc_type in MEETING_TYPES


# ============================================================
# 文种识别
# ============================================================
def _normalize(text: str) -> str:
    """简单归一化：转小写、折叠空白。中文不分词，直接子串匹配。"""
    return "".join(text.lower().split())


def detect_doc_type(requirement: str) -> str | None:
    """从用户需求关键词推断公文文种。

    返回文种名（如 "通知"）或 None（未命中公文关键词）。

    规则:
        - 命中文种关键词 → 返回该文种。
        - 多文种同时命中：优先级 上行文 > 会议文种 > 下行文 > 平行文
          （上行文最严格，命中即采信；"通知/通报" 等下行文最通用，让位）。
        - 必须显式命中公文关键词；"写份文档" 这类不命中。
    """
    text = _normalize(requirement)
    if not text:
        return None

    # 特殊规则：若文本含"批复"相关词，直接判批复。
    # 因为"批复下级的请示""答复xx的请示"等场景里，"请示"是宾语而非要写的文种，
    # 用户真正要写的是批复。优先于普通关键词匹配，避免被请示抢判。
    pifu_kws = DOC_BY_NAME["批复"].keywords
    if any(kw in text for kw in pifu_kws):
        return "批复"

    # 优先级权重（高→低）
    priority = {
        "upward": 4,  # 请示/报告/议案
        "meeting": 3,  # 决议/纪要
        "parallel": 2,  # 函
        "downward": 1,  # 通知/决定/命令...（最通用，最低）
    }

    hits: list[tuple[int, OfficialDocType]] = []
    for d in OFFICIAL_DOCS:
        if any(kw in text for kw in d.keywords):
            hits.append((priority[d.direction], d))

    if not hits:
        return None

    # 取优先级最高者；同优先级取列表顺序（即 index 小的）
    hits.sort(key=lambda x: (-x[0], x[1].index))
    return hits[0][1].name


# ============================================================
# 模板路径与 merge 数据
# ============================================================
def template_path(doc_type: str) -> Path:
    """返回指定文种的模板 .docx 路径。文种不存在抛 ValueError。"""
    d = DOC_BY_NAME.get(doc_type)
    if d is None:
        raise ValueError(f"未知文种 '{doc_type}'。合法文种: {DOC_TYPE_NAMES}")
    return d.path


def template_exists(doc_type: str) -> bool:
    """模板文件是否存在。"""
    try:
        return template_path(doc_type).exists()
    except ValueError:
        return False


def default_merge_data(doc_type: str, **overrides: str) -> dict[str, str]:
    """构造 merge 用的 JSON 数据。

    优先级: 显式 overrides > 文种 defaults > 全局默认值。
    保证返回字典里绝不出现 ``{{key}}`` 字面量（否则 merge 后文档会残留占位符）。

    槽位:
        org, doc_no, signer, signer_org, date_cn, cc, issuer, issue_date
    """
    d = DOC_BY_NAME.get(doc_type)
    if d is None:
        raise ValueError(f"未知文种 '{doc_type}'。合法文种: {DOC_TYPE_NAMES}")

    # 会议类文种（决议/纪要/公报）通常没有印发栏和落款日期
    # 但模板里仍可能有 {{issuer}} 等，统一填空串避免残留
    base: dict[str, str] = {
        "org": _DEF_ORG,
        "doc_no": _DEF_DOC_NO,
        "date_cn": _DEF_DATE_CN,
        "signer_org": _DEF_ORG,  # 默认与 org 一致
        "issuer": _DEF_ISSUER,
        "issue_date": _DEF_ISSUE_DATE,
        "cc": "",  # 默认无抄送
    }
    # 上行文才需要签发人；非上行文 signer 填空串（模板里也不会出现 {{signer}}）
    base["signer"] = _DEF_SIGNER if is_upward(doc_type) else ""

    # 会议类：无落款、无版记印发栏
    if is_meeting(doc_type):
        base["signer_org"] = ""
        base["date_cn"] = ""
        base["issuer"] = ""
        base["issue_date"] = ""
    # 纪要特别处理（meeting 已覆盖）

    # 命令（令）发文字号格式特殊
    if doc_type == "命令（令）":
        base["doc_no"] = "第 X 号"

    # 议案落款是"市长 XX"风格
    if doc_type == "议案":
        base["signer_org"] = "市长  XX"

    # 应用文种自带 defaults（覆盖 base）
    base.update(d.defaults)
    # 应用调用方 overrides（最高优先级），过滤空值
    for k, v in overrides.items():
        if v and v.strip():
            base[k] = v.strip()

    return base


def list_templates() -> list[OfficialDocType]:
    """返回所有文种元数据（按 index 排序）。"""
    return list(OFFICIAL_DOCS)


def format_doc_type_list() -> str:
    """格式化文种清单为多行文本，供 LLM 工具描述使用。"""
    lines = []
    for d in OFFICIAL_DOCS:
        lines.append(f"  - {d.name}：{d.summary}")
    return "\n".join(lines)
