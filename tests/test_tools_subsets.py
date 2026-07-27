"""tools_for_kind：按会话类型裁剪绑定给 LLM 的工具子集。"""

from __future__ import annotations

from office_agent.tools import ALL_TOOLS, tools_for_kind


def _names(tools) -> set[str]:
    return {t.name for t in tools}


class TestToolsForKind:
    def test_docx_subset(self):
        names = _names(tools_for_kind("docx"))
        # 通用 + Word 专属 + 控制
        assert {"create_doc", "view_text", "add_title", "add_heading", "update_paragraph"} <= names
        assert "add_table" in names
        assert {"ask_user", "finish"} <= names
        assert "start_from_template" in names  # 公文模板只在 Word 会话可用
        # 不含 Excel / PPT 专属
        assert "set_cells" not in names
        assert "add_slide" not in names

    def test_xlsx_subset(self):
        names = _names(tools_for_kind("xlsx"))
        assert {"set_cell", "set_cells", "set_formula", "add_pivot_table"} <= names
        assert "add_paragraph" not in names
        assert "add_slide" not in names
        assert "add_image" not in names  # add_image 仅 Word/PPT
        assert "add_table" not in names  # Excel 写表格用 set_cells

    def test_pptx_subset(self):
        names = _names(tools_for_kind("pptx"))
        assert {"add_slide", "add_textbox", "set_theme_colors", "add_image"} <= names
        assert "add_heading" not in names
        assert "set_cells" not in names

    def test_vehicle_flag_appends_query_vehicle(self):
        assert "query_vehicle" not in _names(tools_for_kind("docx"))
        assert "query_vehicle" in _names(tools_for_kind("docx", include_vehicle=True))

    def test_subset_is_strict_subset_of_all(self):
        all_names = _names(ALL_TOOLS)
        for kind in ("docx", "xlsx", "pptx"):
            subset = _names(tools_for_kind(kind, include_vehicle=True))
            assert subset < all_names

    def test_unknown_kind_falls_back_to_docx(self):
        assert _names(tools_for_kind("unknown")) == _names(tools_for_kind("docx"))

    def test_every_kind_has_control_tools(self):
        for kind in ("docx", "xlsx", "pptx"):
            names = _names(tools_for_kind(kind))
            assert {"ask_user", "finish", "create_doc", "view_text", "validate_doc"} <= names
