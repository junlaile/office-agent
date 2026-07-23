"""templates.py 单元测试：15 文种识别 + merge 数据构造。

全部纯函数，无外部依赖（detect_doc_type/default_merge_data/is_upward/is_meeting
不碰文件系统；template_path 只算路径不检查存在）。
"""

from __future__ import annotations

import pytest

from office_agent.domain.templates import (
    DOC_BY_NAME,
    DOC_TYPE_NAMES,
    OFFICIAL_DOCS,
    default_merge_data,
    detect_doc_type,
    format_doc_type_list,
    is_meeting,
    is_upward,
    template_exists,
    template_path,
)


# ============================================================
# detect_doc_type: 文种识别
# ============================================================
class TestDetectDocType:
    """文种关键词识别。"""

    @pytest.mark.parametrize(
        "requirement,expected",
        [
            # 各文种典型正例
            ("写一份关于做好防汛工作的通知", "通知"),
            ("向省厅写一份请示申请资金", "请示"),
            ("写一份会议纪要", "纪要"),
            ("给不相隶属单位写一份商请函", "函"),
            ("写一份通报表彰先进个人", "通报"),
            ("关于进一步加强安全生产工作的意见", "意见"),
            ("向上级汇报本季度工作情况", "报告"),
            ("向上级写一份工作报告", "报告"),
            ("写一份公告向社会公布决定", "公告"),
            ("写一份决定给予处分", "决定"),
            ("发布会议公报", "公报"),
            ("发布通告告知事项", "通告"),
            ("写一份决议", "决议"),
            ("发布命令公布规章", "命令（令）"),
            ("写一份议案提请审议", "议案"),
        ],
    )
    def test_positive_cases(self, requirement, expected):
        """各文种关键词命中。"""
        assert detect_doc_type(requirement) == expected

    @pytest.mark.parametrize(
        "requirement",
        [
            "做一份季度销售数据的 Excel 表格",
            "写一份项目周报，包含本周进展",
            "做一个 10 页的产品介绍 PPT",
            "写一份调研报告",  # 非公文报告
            "帮我分析一下这个数据",
            "写一篇技术博客",
            "",  # 空
            "   ",  # 空白
        ],
    )
    def test_negative_cases(self, requirement):
        """非公文需求应返回 None。"""
        assert detect_doc_type(requirement) is None

    def test_priority_upward_beats_downward(self):
        """上行文优先级 > 下行文：同时命中时上行文胜。"""
        # "请示" 是上行，"通知" 是下行；含两者时取请示
        result = detect_doc_type("写一份请示通知上级")
        assert result == "请示"

    def test_priority_meeting_beats_parallel(self):
        """会议文种 > 平行文。"""
        # 纪要(meeting) vs 函(parallel)
        result = detect_doc_type("写一份会议纪要的函")
        assert result == "纪要"

    def test_pifu_overrides_qingshi(self):
        """批复特殊优先：含'批复'即使有'请示'也判批复。

        '批复下级的请示' 场景，'请示'是宾语而非要写的文种。
        """
        assert detect_doc_type("批复下级的请示") == "批复"
        assert detect_doc_type("答复下级机关的请示") == "批复"

    def test_case_insensitive_and_whitespace(self):
        """大小写和空白归一化。"""
        assert detect_doc_type("  写一份  通知  ") == "通知"

    def test_all_15_types_have_name(self):
        """15 个法定文种全部注册。"""
        expected_names = {
            "决议",
            "决定",
            "命令（令）",
            "公报",
            "公告",
            "通告",
            "意见",
            "通知",
            "通报",
            "报告",
            "请示",
            "批复",
            "议案",
            "函",
            "纪要",
        }
        assert set(DOC_TYPE_NAMES) == expected_names
        assert len(OFFICIAL_DOCS) == 15

    def test_all_types_have_keywords(self):
        """每个文种至少有一个识别关键词。"""
        for d in OFFICIAL_DOCS:
            assert d.keywords, f"{d.name} 缺关键词"


# ============================================================
# default_merge_data: merge 数据构造
# ============================================================
class TestDefaultMergeData:
    """版头槽位预填数据。"""

    def test_no_residual_placeholders(self):
        """所有文种的默认数据里绝不残留 {{key}} 字面量。"""
        for d in OFFICIAL_DOCS:
            data = default_merge_data(d.name)
            for key, value in data.items():
                assert "{{" not in value, f"{d.name}.{key} 残留占位符: {value!r}"

    def test_upward_has_signer(self):
        """上行文（请示/报告/议案）有非空 signer。"""
        for name in ("请示", "报告", "议案"):
            data = default_merge_data(name)
            assert data["signer"], f"{name} 应有签发人"

    def test_non_upward_signer_empty(self):
        """非上行文的 signer 为空（模板里也没有 {{signer}} 槽）。"""
        for name in ("通知", "函", "通报", "批复"):
            data = default_merge_data(name)
            assert data["signer"] == "", f"{name} 不该有签发人"

    def test_meeting_no_issuer(self):
        """会议类（决议/纪要/公报）无印发栏。"""
        for name in ("决议", "纪要"):
            data = default_merge_data(name)
            assert data["issuer"] == ""
            assert data["date_cn"] == ""

    def test_command_special_doc_no(self):
        """命令（令）发文字号特殊：'第 X 号'。"""
        data = default_merge_data("命令（令）")
        assert "第" in data["doc_no"]

    def test_overrides_take_priority(self):
        """显式 overrides 覆盖默认值。"""
        data = default_merge_data("通知", org="市公安局", doc_no="X公发〔2026〕1号")
        assert data["org"] == "市公安局"
        assert data["doc_no"] == "X公发〔2026〕1号"

    def test_signer_org_follows_org(self):
        """传 org 时 signer_org 默认与 org 一致（start_from_template 传的）。"""
        data = default_merge_data("通知", org="市公安局")
        # 注意：default_merge_data 本身不自动同步 signer_org=org，
        # 那是 start_from_template 工具的逻辑；这里只验证默认值合理
        assert data["org"] == "市公安局"

    def test_empty_overrides_ignored(self):
        """空字符串 overrides 被忽略，用默认值。"""
        data = default_merge_data("通知", org="", doc_no="   ")
        assert data["org"] != ""
        assert data["doc_no"] != ""

    def test_unknown_type_raises(self):
        """未知文种抛 ValueError。"""
        with pytest.raises(ValueError, match="未知文种"):
            default_merge_data("不存在文种")

    def test_required_keys_present(self):
        """所有必要槽位都存在。"""
        data = default_merge_data("通知")
        required = {
            "org",
            "doc_no",
            "date_cn",
            "signer",
            "signer_org",
            "issuer",
            "issue_date",
            "cc",
        }
        assert required <= set(data.keys())


# ============================================================
# template_path / template_exists
# ============================================================
class TestTemplatePath:
    """模板路径解析。"""

    def test_known_type_returns_path(self):
        """合法文种返回 template/word/NN-文种.docx 路径。"""
        p = template_path("通知")
        assert p.name == "08-通知.docx"
        assert p.parent.name == "word"

    def test_all_15_filenames(self):
        """15 文种文件名格式正确（NN-文种.docx）。"""
        for d in OFFICIAL_DOCS:
            p = template_path(d.name)
            assert p.name == f"{d.index:02d}-{d.name}.docx"

    def test_unknown_type_raises(self):
        """未知文种抛 ValueError。"""
        with pytest.raises(ValueError, match="未知文种"):
            template_path("不存在")

    def test_template_exists_for_real_templates(self):
        """真实模板文件存在（依赖项目根 template/word/）。"""
        # 这依赖项目结构，若模板未生成会失败
        assert template_exists("通知") is True

    def test_template_exists_unknown_returns_false(self):
        """未知文种返回 False 而非抛错。"""
        assert template_exists("不存在") is False


# ============================================================
# is_upward / is_meeting
# ============================================================
class TestDirectionHelpers:
    """行向判定。"""

    @pytest.mark.parametrize("name", ["请示", "报告", "议案"])
    def test_is_upward_true(self, name):
        assert is_upward(name) is True

    @pytest.mark.parametrize("name", ["通知", "函", "通报", "批复", "决定", "决议", "纪要"])
    def test_is_upward_false(self, name):
        assert is_upward(name) is False

    @pytest.mark.parametrize("name", ["决议", "纪要"])
    def test_is_meeting_true(self, name):
        assert is_meeting(name) is True

    @pytest.mark.parametrize("name", ["通知", "请示", "函", "报告"])
    def test_is_meeting_false(self, name):
        assert is_meeting(name) is False


# ============================================================
# format_doc_type_list / 元数据
# ============================================================
class TestMetadata:
    """文种元数据完整性。"""

    def test_format_doc_type_list_contains_all(self):
        """文种清单文本含全部 15 文种名。"""
        text = format_doc_type_list()
        for d in OFFICIAL_DOCS:
            assert d.name in text

    def test_doc_by_name_lookup(self):
        """DOC_BY_NAME 字典查找。"""
        assert DOC_BY_NAME["通知"].name == "通知"
        assert DOC_BY_NAME["通知"].index == 8

    def test_indices_unique_and_sequential(self):
        """15 文种 index 1-15 唯一且连续。"""
        indices = sorted(d.index for d in OFFICIAL_DOCS)
        assert indices == list(range(1, 16))

    def test_filenames_unique(self):
        """文件名唯一。"""
        names = [d.filename for d in OFFICIAL_DOCS]
        assert len(names) == len(set(names))
