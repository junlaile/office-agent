"""officecli 封装的 argv 断言测试：注入 FakeRunner，验证生成的命令参数数组。

不真调 officecli.exe。FakeRunner 把每次 run(args) 的 args 存下来供断言。
"""

from __future__ import annotations

import json

import pytest

from office_agent.doc_tool import DocTool
from office_agent.excel_tool import ExcelTool
from office_agent.pptx_tool import PptxTool


# ============================================================
# 辅助：从 FakeRunner.calls 里找含某子串的调用
# ============================================================
def find_call(fake, sub_command: str) -> list[str]:
    """找第一条以 sub_command 开头的调用 argv。"""
    for call in fake.calls:
        if call and call[0] == sub_command:
            return call
    pytest.fail(f"未找到子命令 '{sub_command}' 的调用，所有调用: {fake.calls}")


def has_prop(args: list[str], key_value: str) -> bool:
    """argv 里是否含 '--prop key=value' 形式的某 prop。"""
    return key_value in args


# ============================================================
# DocTool 写入命令
# ============================================================
class TestDocToolArgv:
    def test_create_skips_if_exists(self, fake_runner, tmp_path):
        """create() 文件已存在时跳过（不调 runner）。"""
        existing = tmp_path / "exists.docx"
        existing.write_text("dummy")
        tool = DocTool(str(existing))
        result = tool.create()
        assert "跳过" in result
        assert fake_runner.calls == []  # 没调 runner

    def test_create_calls_force_when_absent(self, fake_runner, tmp_path):
        """create() 文件不存在时调 create --force。"""
        tool = DocTool(str(tmp_path / "new.docx"))
        tool.create()
        call = find_call(fake_runner, "create")
        assert "--force" in call

    def test_add_title_argv(self, fake_runner, tmp_path):
        """add_title: size=26 bold align=center。"""
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_title("标题")
        call = find_call(fake_runner, "add")
        assert call[1] == tool.doc_path
        assert "/body" in call
        assert "--type" in call and "paragraph" in call
        assert has_prop(call, "text=标题")
        assert has_prop(call, "size=26")
        assert has_prop(call, "bold=true")
        assert has_prop(call, "align=center")

    def test_add_heading_size_by_level(self, fake_runner, tmp_path):
        """add_heading: level 决定 size（L1=22, L2=18, L3=15）。"""
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_heading("章", level=1)
        call = find_call(fake_runner, "add")
        assert has_prop(call, "size=22")
        assert has_prop(call, "text=章")

    @pytest.mark.parametrize("level,expected_size", [(1, 22), (2, 18), (3, 15), (4, 13.5)])
    def test_heading_levels(self, fake_runner, tmp_path, level, expected_size):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_heading("x", level=level)
        call = find_call(fake_runner, "add")
        assert has_prop(call, f"size={expected_size}")

    def test_add_paragraph_basic(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_paragraph("正文内容")
        call = find_call(fake_runner, "add")
        assert has_prop(call, "text=正文内容")
        # 无 bold/italic 时不附加
        assert not any("bold" in p for p in call if isinstance(p, str))

    def test_add_paragraph_bold_italic(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_paragraph("x", bold=True, italic=True)
        call = find_call(fake_runner, "add")
        assert has_prop(call, "bold=true")
        assert has_prop(call, "italic=true")

    def test_add_list_item_bullet(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_list_item("项", ordered=False)
        call = find_call(fake_runner, "add")
        assert has_prop(call, "listStyle=bullet")

    def test_add_list_item_ordered(self, fake_runner, tmp_path):
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_list_item("项", ordered=True)
        call = find_call(fake_runner, "add")
        assert has_prop(call, "listStyle=ordered")

    def test_set_paragraph_text(self, fake_runner, tmp_path):
        """set_paragraph_text: set path --prop text=..."""
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.set_paragraph_text("/body/p[4]", "新内容")
        call = find_call(fake_runner, "set")
        assert call[2] == "/body/p[4]"
        assert has_prop(call, "text=新内容")

    def test_set_paragraph_text_empty_raises(self, fake_runner, tmp_path):
        """空 path/text 抛错。"""
        from office_agent.cli_runner import OfficeCLIError

        tool = DocTool(str(tmp_path / "t.docx"))
        with pytest.raises(OfficeCLIError):
            tool.set_paragraph_text("", "x")
        with pytest.raises(OfficeCLIError):
            tool.set_paragraph_text("/body/p[1]", "")

    def test_find_replace_basic(self, fake_runner, tmp_path):
        """find_replace: --find X --replace Y。"""
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.find_replace("XX", "新", path="/body")
        call = find_call(fake_runner, "set")
        assert "--find" in call and "XX" in call
        assert "--replace" in call and "新" in call
        assert "/body" in call

    def test_find_replace_default_path_is_body(self, fake_runner, tmp_path):
        """不传 path 默认 /body。"""
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.find_replace("a", "b")
        call = find_call(fake_runner, "set")
        assert "/body" in call

    def test_find_replace_regex_prefix(self, fake_runner, tmp_path):
        """regex=True 时 find 加 r"..." 前缀。"""
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.find_replace(r"\d+", "2026", regex=True)
        call = find_call(fake_runner, "set")
        # find 值应是 r"\d+"（officecli 正则语法）
        idx = call.index("--find")
        assert call[idx + 1] == 'r"\\d+"'

    def test_remove(self, fake_runner, tmp_path):
        """remove: remove doc path。"""
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.remove("/body/p[10]")
        call = find_call(fake_runner, "remove")
        assert call[1] == tool.doc_path
        assert call[2] == "/body/p[10]"

    def test_add_table_creates_table_and_writes_cells(self, fake_runner, tmp_path):
        """add_table: 先 add table 建空表，再 batch 写单元格。"""
        fake_runner.default_stdout = "Added table at /body/tbl[1]"
        tool = DocTool(str(tmp_path / "t.docx"))
        tool.add_table([["h1", "h2"], ["v1", "v2"]], has_header=True)
        # 应有 add table + batch
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        batch_calls = [c for c in fake_runner.calls if c[0] == "batch"]
        assert any("--type" in c and "table" in c for c in add_calls)
        assert len(batch_calls) >= 1


# ============================================================
# ExcelTool 命令
# ============================================================
class TestExcelToolArgv:
    def test_set_cell_basic(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_cell("Sheet1", "A1", "值")
        call = find_call(fake_runner, "set")
        assert "/Sheet1/A1" in call
        assert has_prop(call, "value=值")

    def test_set_cell_bold_fill(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_cell("S", "B2", "x", bold=True, fill="FFFF00")
        call = find_call(fake_runner, "set")
        assert has_prop(call, "bold=true")
        assert has_prop(call, "fill=FFFF00")

    def test_set_formula(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_formula("S", "D2", "SUM(B2:B10)")
        call = find_call(fake_runner, "set")
        assert has_prop(call, "formula=SUM(B2:B10)")

    def test_add_sheet(self, fake_runner, tmp_path):
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.add_sheet("数据", tab_color="4472C4")
        call = find_call(fake_runner, "add")
        assert has_prop(call, "name=数据")
        assert has_prop(call, "tabColor=4472C4")

    def test_set_cells_batch(self, fake_runner, tmp_path):
        """set_cells 用 batch 逐 cell 写。"""
        tool = ExcelTool(str(tmp_path / "t.xlsx"))
        tool.set_cells("S", [["a", "b"], ["c", "d"]], start_ref="A1")
        batch_calls = [c for c in fake_runner.calls if c[0] == "batch"]
        assert len(batch_calls) >= 1


# ============================================================
# PptxTool 命令
# ============================================================
class TestPptxToolArgv:
    def test_add_slide_with_title_text(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.add_slide(title="标题", text="正文")
        call = find_call(fake_runner, "add")
        assert has_prop(call, "title=标题")
        assert has_prop(call, "text=正文")

    def test_add_textbox(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.add_textbox(1, "内容", x="2cm", y="3cm", size=18)
        call = find_call(fake_runner, "add")
        assert "/slide[1]" in call
        assert has_prop(call, "text=内容")
        assert has_prop(call, "x=2cm")
        assert has_prop(call, "size=18")

    def test_set_notes(self, fake_runner, tmp_path):
        tool = PptxTool(str(tmp_path / "t.pptx"))
        tool.set_notes(2, "备注内容")
        call = find_call(fake_runner, "set")
        assert "/slide[2]" in call
        assert has_prop(call, "notes=备注内容")


# ============================================================
# merge_template 命令
# ============================================================
class TestMergeTemplate:
    def test_merge_argv(self, fake_runner):
        """merge_template: merge tmpl output --data <json> --force。"""
        from office_agent.cli_runner import merge_template

        merge_template("/tmpl.docx", "/out.docx", {"org": "局"})
        call = find_call(fake_runner, "merge")
        assert call[1] == "/tmpl.docx"
        assert call[2] == "/out.docx"
        assert "--data" in call
        # data 是 JSON 字符串
        idx = call.index("--data")
        data = json.loads(call[idx + 1])
        assert data == {"org": "局"}
        assert "--force" in call

    def test_merge_chinese_not_escaped(self, fake_runner):
        """JSON 用 ensure_ascii=False，中文不转义。"""
        from office_agent.cli_runner import merge_template

        merge_template("/t.docx", "/o.docx", {"org": "公安局"})
        call = find_call(fake_runner, "merge")
        idx = call.index("--data")
        assert "公安局" in call[idx + 1]  # 中文直接出现，非 \uXXXX


# ============================================================
# 错误传播
# ============================================================
class TestErrorPropagation:
    def test_runner_error_raised(self, tmp_path):
        """FakeRunner 模拟失败时，DocTool 方法应抛 OfficeCLIError。"""
        from office_agent import cli_runner
        from office_agent.cli_runner import OfficeCLIError

        class FailingRunner:
            def run(self, args, **kwargs):
                raise OfficeCLIError("模拟失败", cmd=args)

        cli_runner._runner = FailingRunner()
        try:
            tool = DocTool(str(tmp_path / "t.docx"))
            with pytest.raises(OfficeCLIError, match="模拟失败"):
                tool.add_title("x")
        finally:
            cli_runner._runner = None
