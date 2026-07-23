"""cli_ui.py 单元测试：从 main.py 下沉的纯函数 + UI 函数。

纯函数（_infer_doc_kind / _format_tool_call / _indent / _derive_doc_path）
直接测；UI 函数（_banner / _print_* / _collect_*）用 capsys/monkeypatch。
"""

from __future__ import annotations

import dataclasses
from pathlib import Path

import pytest

from office_agent.cli import ui as cli_ui
from office_agent.cli.ui import (
    _derive_doc_path,
    _format_tool_call,
    _indent,
    _infer_doc_kind,
)
from office_agent.domain.format import (
    _DOCX_KEYWORDS,
    _PPTX_KEYWORDS,
    _XLSX_KEYWORDS,
)


@pytest.fixture
def output_to_tmp(monkeypatch, tmp_path):
    """把 cli_ui.settings.output_dir 指向临时目录。

    Settings 是 frozen dataclass，用 dataclasses.replace 造新对象替换。
    """
    new_settings = dataclasses.replace(cli_ui.settings, output_dir=tmp_path)
    monkeypatch.setattr(cli_ui, "settings", new_settings)
    return tmp_path


# ============================================================
# _infer_doc_kind: 文档类型关键词推断
# ============================================================
class TestInferDocKind:
    """关键词命中与优先级。"""

    @pytest.mark.parametrize(
        "req,expected_kind",
        [
            ("做一份销售数据的 Excel 表格", "xlsx"),
            ("写一份工作表", "xlsx"),
            ("做一个产品介绍 PPT", "pptx"),
            ("做一份幻灯片", "pptx"),
            ("写一份 word 文档", "docx"),
            ("写一份项目报告", "docx"),
            ("写一份方案", "docx"),
        ],
    )
    def test_keyword_match(self, req, expected_kind):
        kind, score = _infer_doc_kind(req)
        assert kind == expected_kind
        assert score >= 1

    def test_no_match_returns_zero(self):
        """无任何关键词命中时 score=0（kind 仍是默认 docx）。"""
        kind, score = _infer_doc_kind("随便写点东西")
        assert score == 0

    def test_tie_priority_xlsx_beats_pptx_beats_docx(self):
        """平局时优先级 xlsx > pptx > docx。"""
        # 构造同时命中 xlsx 和 docx 的需求
        kind, _ = _infer_doc_kind("excel 报告")  # xlsx(excel) + docx(报告) 各 1
        assert kind == "xlsx"

    def test_case_insensitive(self):
        """关键词大小写不敏感。"""
        kind, score = _infer_doc_kind("做一份 EXCEL")
        assert kind == "xlsx"
        assert score >= 1

    def test_keywords_nonempty(self):
        """三个关键词表都不空。"""
        assert _XLSX_KEYWORDS
        assert _PPTX_KEYWORDS
        assert _DOCX_KEYWORDS


# ============================================================
# _format_tool_call: 工具调用展示格式化
# ============================================================
class TestFormatToolCall:
    """每个工具分支的格式化。"""

    def test_add_title(self):
        s = _format_tool_call("add_title", {"text": "标题"})
        assert "标题" in s
        assert "add_title" in s

    def test_add_heading(self):
        s = _format_tool_call("add_heading", {"text": "章节", "level": 2})
        assert "章节" in s
        assert "level=2" in s

    def test_add_paragraph_truncated(self):
        """长文本截断。"""
        long = "x" * 100
        s = _format_tool_call("add_paragraph", {"text": long})
        assert "…" in s

    def test_add_paragraph_bold_flag(self):
        s = _format_tool_call("add_paragraph", {"text": "t", "bold": True})
        assert "bold" in s

    def test_add_list_item_ordered(self):
        s = _format_tool_call("add_list_item", {"text": "项", "ordered": True})
        assert "ordered" in s

    def test_add_table_dimensions(self):
        s = _format_tool_call("add_table", {"data": [["a", "b"], ["c", "d"]]})
        assert "2×2" in s

    def test_add_slide_with_body(self):
        s = _format_tool_call("add_slide", {"title": "T", "body_text": "a\nb\nc"})
        assert "T" in s
        assert "3行" in s  # 3 行正文

    def test_add_slide_no_body_warns(self):
        """无正文的幻灯片提示警告。"""
        s = _format_tool_call("add_slide", {"title": "T", "body_text": ""})
        assert "无正文" in s

    def test_add_image_with_caption(self):
        s = _format_tool_call("add_image", {"url_or_path": "http://x", "caption": "图"})
        assert "图" in s

    def test_set_cells(self):
        s = _format_tool_call("set_cells", {"sheet": "S", "data": [["a"]], "start": "B2"})
        assert "S" in s
        assert "B2" in s

    def test_set_formula(self):
        s = _format_tool_call("set_formula", {"sheet": "S", "ref": "D2", "formula": "SUM(B2:B10)"})
        assert "=SUM" in s

    def test_start_from_template(self):
        s = _format_tool_call("start_from_template", {"doc_type": "通知", "org": "局"})
        assert "通知" in s
        assert "局" in s

    def test_update_paragraph(self):
        s = _format_tool_call("update_paragraph", {"path": "/body/p[4]", "text": "新"})
        assert "/body/p[4]" in s
        assert "新" in s

    def test_replace_text(self):
        s = _format_tool_call("replace_text", {"find": "XX", "replace": "新"})
        assert "XX" in s
        assert "新" in s
        assert "→" in s

    def test_remove_paragraph(self):
        s = _format_tool_call("remove_paragraph", {"path": "/body/p[10]"})
        assert "/body/p[10]" in s

    def test_ask_user_with_fields(self):
        s = _format_tool_call("ask_user", {"title": "采集", "fields": [{"k": "v"}, {"k": "w"}]})
        assert "采集" in s
        assert "2个字段" in s

    def test_finish(self):
        s = _format_tool_call("finish", {"summary": "完成"})
        assert "完成" in s

    def test_unknown_tool_fallback(self):
        """未知工具回退到 name() 格式。"""
        s = _format_tool_call("some_new_tool", {})
        assert "some_new_tool()" in s

    def test_no_arg_tools(self):
        """无参工具（create_doc/view_text/validate_doc）格式为 name()。"""
        for name in ["create_doc", "view_text", "validate_doc"]:
            s = _format_tool_call(name, {})
            assert s == f"{name}()"


# ============================================================
# _indent: 多行缩进
# ============================================================
class TestIndent:
    def test_single_line(self):
        assert _indent("hello") == "    hello"

    def test_multi_line(self):
        assert _indent("a\nb\nc") == "    a\n    b\n    c"

    def test_custom_pad(self):
        assert _indent("x", pad="  ") == "  x"

    def test_empty(self):
        assert _indent("") == ""

    def test_preserves_internal_whitespace(self):
        assert _indent("  indented") == "      indented"


# ============================================================
# _derive_doc_path: 文档路径推导
# ============================================================
class TestDeriveDocPath:
    """路径推导（含扩展名 + 时间戳）。"""

    def test_returns_docx_for_docx_keyword(self, output_to_tmp):
        """含 word 关键词 → .docx 扩展名。"""
        p = _derive_doc_path("写一份 word 报告")
        assert p.endswith(".docx")

    def test_returns_xlsx_for_excel_keyword(self, output_to_tmp):
        p = _derive_doc_path("做一份 excel 表格")
        assert p.endswith(".xlsx")

    def test_official_mode_forces_docx(self, output_to_tmp):
        """doc_type 非空时强制 .docx，即使需求无关键词。"""
        p = _derive_doc_path("随便写点", doc_type="通知")
        assert p.endswith(".docx")

    def test_path_in_output_dir(self, output_to_tmp):
        """路径在 output_dir 下。"""
        p = _derive_doc_path("写报告")
        assert str(output_to_tmp) in p

    def test_filename_contains_requirement_text(self, output_to_tmp):
        """文件名含需求文字（清理后）。"""
        p = _derive_doc_path("项目周报")
        assert "项目周报" in p

    def test_filename_sanitizes_special_chars(self, output_to_tmp):
        """特殊字符被清理（用 doc_type 强制 docx，避免触发类型询问交互）。"""
        p = _derive_doc_path('a/b\\c:*?"<>|', doc_type="通知")
        name = Path(p).name
        for ch in '/\\:*?"<>|':
            assert ch not in name


# ============================================================
# 大纲预览决策
# ============================================================
class TestOutlineDecision:
    def test_approve(self, monkeypatch):
        from office_agent.cli import ui

        monkeypatch.setattr(ui, "_readline", lambda prompt="": "1")
        action, fb = ui._collect_outline_decision()
        assert action == "approve"
        assert fb == ""

    def test_revise_with_feedback(self, monkeypatch):
        from office_agent.cli import ui

        answers = iter(["2", "加上风险一节"])
        monkeypatch.setattr(ui, "_readline", lambda prompt="": next(answers))
        action, fb = ui._collect_outline_decision()
        assert action == "revise"
        assert "风险" in fb

    def test_cancel(self, monkeypatch):
        from office_agent.cli import ui

        monkeypatch.setattr(ui, "_readline", lambda prompt="": "取消")
        action, fb = ui._collect_outline_decision()
        assert action == "cancel"

    def test_approval_loop_approve(self, monkeypatch, capsys):
        from office_agent.cli import ui

        monkeypatch.setattr(
            "office_agent.agent.outline.generate_outline",
            lambda *a, **k: "# 大纲\n## A\n- 1",
        )
        monkeypatch.setattr(ui, "_readline", lambda prompt="": "1")
        result = ui._run_outline_approval_loop("写周报")
        assert result is not None
        assert "# 大纲" in result
        assert "已批准" in capsys.readouterr().out

    def test_approval_loop_cancel(self, monkeypatch):
        from office_agent.cli import ui

        monkeypatch.setattr(
            "office_agent.agent.outline.generate_outline",
            lambda *a, **k: "# X",
        )
        monkeypatch.setattr(ui, "_readline", lambda prompt="": "3")
        assert ui._run_outline_approval_loop("写周报") is None
