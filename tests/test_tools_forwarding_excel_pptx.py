"""Excel / PowerPoint @tool 函数 → ExcelTool / PptxTool 的 argv 转发测试。

与 test_tools_forwarding.py（Word）同一思路：fake_runner 断言 subprocess 层
实际发出的命令，补齐 tools/excel.py 与 tools/pptx.py 的转发覆盖。
"""

from __future__ import annotations

import json

from office_agent.tools import TOOL_BY_NAME


class TestExcelToolsForwarding:
    def test_add_sheet_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["add_sheet"].invoke({"name": "销售数据", "tab_color": "4472C4"})
        add_calls = fake_runner.calls_of("add")
        assert add_calls
        assert any("name=销售数据" in c and "tabColor=4472C4" in c for c in add_calls)

    def test_set_cell_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["set_cell"].invoke({"sheet": "S", "ref": "A1", "value": "标题", "bold": True})
        set_calls = fake_runner.calls_of("set")
        assert any("/S/A1" in c and "value=标题" in c and "bold=true" in c for c in set_calls)

    def test_set_cells_uses_batch(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["set_cells"].invoke(
            {
                "sheet": "S",
                "data": [["月份", "销售额"], ["1月", 12000]],
                "start": "A1",
                "has_header": True,
            }
        )
        batch_calls = fake_runner.calls_of("batch")
        assert len(batch_calls) == 1
        payload = json.loads(batch_calls[0][batch_calls[0].index("--commands") + 1])
        assert payload[0]["path"] == "/S/A1"
        assert payload[0]["props"]["bold"] == "true"  # 表头加粗
        assert payload[3]["path"] == "/S/B2"
        assert payload[3]["props"]["value"] == "12000"

    def test_set_formula_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["set_formula"].invoke({"sheet": "S", "ref": "D2", "formula": "SUM(B2:B10)"})
        set_calls = fake_runner.calls_of("set")
        assert any("/S/D2" in c and "formula=SUM(B2:B10)" in c for c in set_calls)

    def test_sort_sheet_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["sort_sheet"].invoke({"sheet": "S", "keys": "B desc"})
        set_calls = fake_runner.calls_of("set")
        assert any("sort=B desc" in c for c in set_calls)

    def test_set_autofilter_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["set_autofilter"].invoke({"sheet": "S", "cell_range": "A1:D10"})
        set_calls = fake_runner.calls_of("set")
        assert any("autoFilter=A1:D10" in c for c in set_calls)

    def test_merge_cells_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["merge_cells"].invoke({"sheet": "S", "cell_range": "A1:C1"})
        set_calls = fake_runner.calls_of("set")
        assert any("merge=A1:C1" in c for c in set_calls)

    def test_rename_sheet_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["rename_sheet"].invoke({"old_name": "Sheet1", "new_name": "数据"})
        set_calls = fake_runner.calls_of("set")
        assert any("/Sheet1" in c and "name=数据" in c for c in set_calls)


class TestPptxToolsForwarding:
    def test_add_slide_forwards(self, fake_runner, pptx_session):
        TOOL_BY_NAME["add_slide"].invoke({"title": "封面", "body_text": "副标题"})
        add_calls = fake_runner.calls_of("add")
        assert any("title=封面" in c and "text=副标题" in c for c in add_calls)

    def test_add_slide_title_only_warns(self, fake_runner, pptx_session):
        result = TOOL_BY_NAME["add_slide"].invoke({"title": "内容页"})
        assert "body_text" in result  # 只有标题没正文 → 警告

    def test_add_textbox_targets_last_slide(self, fake_runner, pptx_session):
        TOOL_BY_NAME["add_textbox"].invoke({"text": "文本", "size": 20})
        add_calls = fake_runner.calls_of("add")
        # 无幻灯片时 last_slide_index 返回 0 → 兜底 slide[1]
        assert any("/slide[1]" in c and "text=文本" in c for c in add_calls)

    def test_set_slide_notes_forwards(self, fake_runner, pptx_session):
        TOOL_BY_NAME["set_slide_notes"].invoke({"slide_index": 1, "notes": "开场要点"})
        assert any("开场要点" in " ".join(c) for c in fake_runner.calls)

    def test_set_theme_colors_forwards(self, fake_runner, pptx_session):
        TOOL_BY_NAME["set_theme_colors"].invoke({"accent1": "4472C4"})
        assert any("accent1=4472C4" in " ".join(c) for c in fake_runner.calls)

    def test_wrong_kind_rejected(self, doc_session):
        result = TOOL_BY_NAME["add_slide"].invoke({"title": "x"})
        assert "pptx" in result.lower() or "PPTX" in result
