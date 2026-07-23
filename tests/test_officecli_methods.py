"""DocTool/ExcelTool/PptxTool 更多方法的 argv 测试（扩充覆盖率）。

补充 test_officecli_argv.py 未覆盖的方法，用 FakeRunner 断言 argv。
"""

from __future__ import annotations

import pytest

from office_agent.office.doc import DocTool
from office_agent.office.excel import ExcelTool
from office_agent.office.pptx import PptxTool


def joined_args(call):
    """把 argv list 拼成字符串便于子串断言。"""
    return " ".join(call)


# ============================================================
# DocTool 剩余方法
# ============================================================
class TestDocToolMore:
    def test_add_image(self, fake_runner, tmp_path):
        fake_runner.responses = {"add": "Added paragraph at /body/p[1]"}
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_image("http://x/a.png", width="10cm", alt="图", caption="注")
        # 应有多次 add（段落 + picture + 图注）
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert len(add_calls) >= 2
        # 含 picture 和 src
        assert any(
            "picture" in joined_args(c) and "src=http://x/a.png" in joined_args(c)
            for c in add_calls
        )

    def test_add_toc(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_toc(levels="1-3", title="目录")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("toc" in joined_args(c) for c in add_calls)

    def test_add_header(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_header("页眉", align="right")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("header" in joined_args(c) and "text=页眉" in joined_args(c) for c in add_calls)

    def test_add_footer(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_footer("页脚", field="page")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("footer" in joined_args(c) for c in add_calls)

    def test_add_hyperlink(self, fake_runner, tmp_path):
        fake_runner.responses = {"add": "Added paragraph at /body/p[1]"}
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_hyperlink("链接", url="http://x", tooltip="提示")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any(
            "hyperlink" in joined_args(c) and "url=http://x" in joined_args(c) for c in add_calls
        )

    def test_add_chart(self, fake_runner, tmp_path):
        fake_runner.responses = {"add": "Added paragraph at /body/p[1]"}
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_chart("column", "Sales:1,2,3", categories="A,B,C", title="图")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any(
            "chart" in joined_args(c) and "chartType=column" in joined_args(c) for c in add_calls
        )

    def test_add_section(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_section(orientation="landscape")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("section" in joined_args(c) for c in add_calls)

    def test_set_doc_properties(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.set_doc_properties(title="T", author="A")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("title=T" in joined_args(c) for c in set_calls)

    def test_set_doc_properties_no_args_raises(self, fake_runner, tmp_path):
        from office_agent.office.runner import OfficeCLIError

        tool = DocTool(str(tmp_path / "t.docx"))
        with pytest.raises(OfficeCLIError):
            tool.set_doc_properties()

    def test_add_styled_paragraph(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_styled_paragraph("x", "Heading1")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("style=Heading1" in joined_args(c) for c in add_calls)

    def test_add_style(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_style("My", "我的", size=14, bold=True)
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("id=My" in joined_args(c) for c in add_calls)

    def test_view_text(self, fake_runner, tmp_path):
        fake_runner.default_stdout = "[/body/p[1]] 内容"
        tool = DocTool(str(tmp_path / "t.docx"))
        result = tool.view_text()
        assert "内容" in result

    def test_validate(self, fake_runner, tmp_path):
        fake_runner.default_stdout = "Validation passed"
        tool = DocTool(str(tmp_path / "t.docx"))
        assert "passed" in tool.validate().lower()

    def test_close(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.close()
        close_calls = [c for c in fake_runner.calls if c[0] == "close"]
        assert close_calls


# ============================================================
# ExcelTool 剩余方法
# ============================================================
class TestExcelToolMore:
    def test_set_cells_with_header(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_cells("S", [["h1", "h2"], ["v1", "v2"]], start_ref="A1", has_header=True)
        batch_calls = [c for c in fake_runner.calls if c[0] == "batch"]
        assert batch_calls

    def test_set_column_width(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_column_width("S", "A", 20)
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("width=20" in joined_args(c) for c in set_calls)

    def test_set_row_height(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_row_height("S", 1, 30)
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("height=30" in joined_args(c) for c in set_calls)

    def test_autofit_column(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.autofit_column("S", "A")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("autofit=true" in joined_args(c) for c in set_calls)

    def test_merge_cells(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.merge_cells("S", "A1:D1")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("merge=A1:D1" in joined_args(c) for c in set_calls)

    def test_add_chart(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.add_chart("S", "column", "S!B1:C4", title="图")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any(
            "chart" in joined_args(c) and "chartType=column" in joined_args(c) for c in add_calls
        )

    def test_sort(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.sort("S", "B desc", has_header=True)
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("sort=B desc" in joined_args(c) for c in set_calls)

    def test_set_autofilter(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_autofilter("S", "A1:D10")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("autoFilter=A1:D10" in joined_args(c) for c in set_calls)

    def test_add_conditional_format_cellis(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.add_conditional_format(
            "S", "cellIs", "C2:C100", operator="greaterThan", value="50", fill="FF0000"
        )
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("cellIs" in joined_args(c) for c in add_calls)

    def test_add_pivot_table(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.add_pivot_table("S", "Sheet1!A1:D100", rows="区域", values="sales:sum")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("pivottable" in joined_args(c) for c in add_calls)

    def test_add_list_table(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.add_list_table("S", "A1:C10", style="medium2")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("table" in joined_args(c) for c in add_calls)

    def test_add_validation_list(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.add_validation("S", "B2:B100", "list", formula1="是,否")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("validation" in joined_args(c) for c in add_calls)

    def test_rename_sheet(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.rename_sheet("old", "new")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("name=new" in joined_args(c) for c in set_calls)

    def test_set_sheet_color(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_sheet_color("S", "FF0000")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("tabColor=FF0000" in joined_args(c) for c in set_calls)

    def test_add_named_range(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.add_named_range("Rev", "Sheet1!A1:A10")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("namedrange" in joined_args(c) for c in add_calls)

    def test_get_cell(self, fake_runner, tmp_path):
        fake_runner.responses = {"get": {"data": {"results": [{"value": "x"}]}}}
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.get_cell("S", "A1")
        # 应有 get 调用
        get_calls = [c for c in fake_runner.calls if c[0] == "get"]
        assert get_calls

    def test_view_text(self, fake_runner, tmp_path):
        fake_runner.default_stdout = "A1=x"
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        assert "x" in tool.view_text()


# ============================================================
# PptxTool 剩余方法
# ============================================================
class TestPptxToolMore:
    def test_add_shape(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.add_shape(1, "文字", geometry="rect", fill="FF0000")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any(
            "shape" in joined_args(c) and "geometry=rect" in joined_args(c) for c in add_calls
        )

    def test_add_image(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.add_image(1, "http://x/a.png", width="10cm")
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("picture" in joined_args(c) for c in add_calls)

    def test_add_table(self, fake_runner, tmp_path):
        fake_runner.responses = {"get": {"data": {"results": [{"children": []}]}}}
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.add_table(1, [["a", "b"]], has_header=True)
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("table" in joined_args(c) for c in add_calls)

    def test_set_transition(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.set_transition(1, "fade", advance_time_ms=5000)
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("transition=fade" in joined_args(c) for c in set_calls)

    def test_set_slide_hidden(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.set_slide_hidden(2, True)
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("hidden=true" in joined_args(c) for c in set_calls)

    def test_set_theme_colors(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.set_theme_colors(accent1="4472C4", hyperlink="0563C1")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("accent1=4472C4" in joined_args(c) for c in set_calls)

    def test_set_theme_fonts(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.set_theme_fonts(heading_font="雅黑", body_font="雅黑")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("headingFont=雅黑" in joined_args(c) for c in set_calls)

    def test_set_presentation_props(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.set_presentation_props(title="T", author="A")
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("title=T" in joined_args(c) for c in set_calls)

    def test_set_presentation_props_no_args_raises(self, fake_runner, tmp_path):
        from office_agent.office.runner import OfficeCLIError

        tool = PptxTool(str(tmp_path / "t.pptx"))
        with pytest.raises(OfficeCLIError):
            tool.set_presentation_props()

    def test_set_theme_colors_no_args_raises(self, fake_runner, tmp_path):
        from office_agent.office.runner import OfficeCLIError

        tool = PptxTool(str(tmp_path / "t.pptx"))
        with pytest.raises(OfficeCLIError):
            tool.set_theme_colors()

    def test_view_text(self, fake_runner, tmp_path):
        fake_runner.default_stdout = "slide1: 内容"
        tool = PptxTool(str(tmp_path / "t.pptx"))
        assert "内容" in tool.view_text()

    def test_close(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.close()
        assert [c for c in fake_runner.calls if c[0] == "close"]


# ============================================================
# cli_runner 的 raw 与错误路径
# ============================================================
class TestCliRunner:
    def test_raw_whitelisted(self, fake_runner):
        from office_agent.office.runner import raw

        raw(["view", "t.docx", "text"])
        assert fake_runner.calls

    def test_raw_non_whitelisted_raises(self, fake_runner):
        from office_agent.office.runner import OfficeCLIError, raw

        with pytest.raises(OfficeCLIError):
            raw(["dangerous", "cmd"])

    def test_raw_empty_raises(self, fake_runner):
        from office_agent.office.runner import OfficeCLIError, raw

        with pytest.raises(OfficeCLIError):
            raw([])
