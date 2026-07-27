"""tools 包测试：session 路由 + 工具转发 + start_from_template 编排。

通过 set_session_doc 注入会话路径，monkeypatch merge_template / DocTool 方法
断言转发逻辑，不碰真实 officecli.exe。
"""

from __future__ import annotations

import pytest

from office_agent.tools import (
    REGISTRY,
    SPEC_BY_NAME,
    TOOL_BY_NAME,
    ExecutionMode,
    SideEffect,
    session_doc_kind,
    session_doc_path,
    set_session_doc,
    tools_for_doc_path,
)


# ============================================================
# session 路由
# ============================================================
class TestSessionRouting:
    def test_docx_kind(self, doc_session):
        assert session_doc_kind() == "docx"

    def test_xlsx_kind(self, xlsx_session):
        assert session_doc_kind() == "xlsx"

    def test_pptx_kind(self, pptx_session):
        assert session_doc_kind() == "pptx"

    def test_unknown_extension_defaults_docx(self, tmp_path):
        set_session_doc(str(tmp_path / "x.unknown"))
        assert session_doc_kind() == "docx"

    def test_uninitialized_raises(self):
        """未初始化 session 时 session_doc_kind 抛错。"""
        from office_agent.office.runner import OfficeCLIError

        set_session_doc(None)
        with pytest.raises(OfficeCLIError):
            session_doc_kind()

    def test_session_path_getter(self, doc_session):
        assert session_doc_path() == doc_session


# ============================================================
# 工具的格式守卫（wrong kind）
# ============================================================
class TestWrongKindGuards:
    """Word 工具在 xlsx/pptx 会话下应返回 wrong_kind 提示。"""

    def test_update_paragraph_in_xlsx(self, xlsx_session, capsys):
        result = TOOL_BY_NAME["update_paragraph"].invoke({"path": "/body/p[1]", "text": "x"})
        assert "DOCX 专属" in result or "docx" in result.lower()

    def test_replace_text_in_xlsx(self, xlsx_session):
        result = TOOL_BY_NAME["replace_text"].invoke({"find": "a", "replace": "b"})
        assert "docx" in result.lower() or "DOCX" in result

    def test_remove_paragraph_in_pptx(self, pptx_session):
        result = TOOL_BY_NAME["remove_paragraph"].invoke({"path": "/x"})
        assert "docx" in result.lower() or "DOCX" in result

    def test_excel_tool_in_docx(self, doc_session):
        """Excel 工具在 docx 会话下拒绝。"""
        result = TOOL_BY_NAME["set_cell"].invoke({"sheet": "S", "ref": "A1", "value": "x"})
        assert "xlsx" in result.lower() or "XLSX" in result


# ============================================================
# start_from_template 编排（mock merge + 校验）
# ============================================================
class TestStartFromTemplate:
    def test_wrong_kind_in_xlsx(self, xlsx_session):
        """非 docx 会话拒绝。"""
        result = TOOL_BY_NAME["start_from_template"].invoke({"doc_type": "通知"})
        assert "Word" in result

    def test_unknown_doc_type(self, doc_session):
        """未知文种返回错误。"""
        result = TOOL_BY_NAME["start_from_template"].invoke({"doc_type": "不存在文种"})
        assert "不存在" in result or "合法文种" in result

    def test_calls_merge_with_overrides(self, doc_session, monkeypatch):
        """调用 merge_template 并传入含 overrides 的 merge_data。"""
        calls = []

        def fake_merge(tmpl, out, data):
            calls.append((tmpl, out, data))
            return "Merged OK"

        monkeypatch.setattr("office_agent.tools.common.merge_template", fake_merge)
        result = TOOL_BY_NAME["start_from_template"].invoke(
            {
                "doc_type": "通知",
                "org": "市公安局",
                "doc_no": "X公发〔2026〕1号",
            }
        )
        assert len(calls) == 1
        _, _, data = calls[0]
        assert data["org"] == "市公安局"
        assert data["doc_no"] == "X公发〔2026〕1号"
        # 无 {{ 残留
        assert all("{{" not in v for v in data.values())
        assert "已从《通知》模板创建" in result

    def test_title_addressee_hint_in_result(self, doc_session, monkeypatch):
        """传 title/addressee 时返回里含路径提示。"""
        monkeypatch.setattr("office_agent.tools.common.merge_template", lambda *a, **k: "OK")
        result = TOOL_BY_NAME["start_from_template"].invoke(
            {
                "doc_type": "通知",
                "title": "真实标题",
                "addressee": "各分局：",
            }
        )
        assert "/body/p[4]" in result  # 标题路径提示
        assert "/body/p[5]" in result  # 主送路径提示

    def test_merge_failure_returns_error(self, doc_session, monkeypatch):
        """merge 失败时返回错误而非崩溃。"""
        from office_agent.office.runner import OfficeCLIError

        def failing_merge(*a, **k):
            raise OfficeCLIError("merge 失败")

        monkeypatch.setattr("office_agent.tools.common.merge_template", failing_merge)
        result = TOOL_BY_NAME["start_from_template"].invoke({"doc_type": "通知"})
        assert "失败" in result


# ============================================================
# 工具集完整性
# ============================================================
class TestToolRegistry:
    def test_all_tools_count(self):
        """ALL_TOOLS 含 49 个工具。"""
        from office_agent.tools import ALL_TOOLS

        assert len(ALL_TOOLS) == 49

    def test_tool_by_name_complete(self):
        """TOOL_BY_NAME 含全部工具。"""
        from office_agent.tools import ALL_TOOLS

        names = {t.name for t in ALL_TOOLS}
        assert set(TOOL_BY_NAME.keys()) == names

    def test_critical_tools_present(self):
        """关键工具都在。"""
        must_have = {
            "create_doc",
            "start_from_template",
            "view_text",
            "validate_doc",
            "finish",
            "ask_user",
            # Word
            "add_title",
            "add_paragraph",
            "add_heading",
            "update_paragraph",
            "replace_text",
            "remove_paragraph",
            # Excel
            "set_cell",
            "set_cells",
            # PPT
            "add_slide",
        }
        assert must_have <= set(TOOL_BY_NAME.keys())

    def test_tool_names_unique(self):
        """工具名唯一。"""
        from office_agent.tools import ALL_TOOLS

        names = [t.name for t in ALL_TOOLS]
        assert len(names) == len(set(names))

    @pytest.mark.parametrize(
        ("path", "included", "excluded"),
        [
            (
                "report.docx",
                {"create_doc", "add_title", "update_paragraph", "add_image", "finish"},
                {"set_cells", "add_slide"},
            ),
            (
                "report.xlsx",
                {"create_doc", "set_cells", "add_excel_chart", "finish"},
                {"add_title", "add_slide", "add_image", "start_from_template"},
            ),
            (
                "report.pptx",
                {"create_doc", "add_slide", "add_image", "finish"},
                {"add_title", "set_cells", "start_from_template"},
            ),
        ],
    )
    def test_tools_selected_by_document_type(self, path, included, excluded):
        """每种文档只向 LLM 暴露相关工具。"""
        names = {tool.name for tool in tools_for_doc_path(path)}
        assert included <= names
        assert names.isdisjoint(excluded)
        assert {"query_vehicle", "ask_user", "finish"} <= names

    @pytest.mark.parametrize("path", ["report.docx", "report.xlsx", "report.pptx"])
    def test_selected_tool_names_unique(self, path):
        tools = tools_for_doc_path(path)
        names = [tool.name for tool in tools]
        assert len(names) == len(set(names))

    def test_unknown_extension_defaults_to_word_tools(self):
        names = {tool.name for tool in tools_for_doc_path("report.unknown")}
        assert {"add_title", "update_paragraph", "add_image"} <= names
        assert "set_cells" not in names
        assert "add_slide" not in names

    def test_registry_metadata_matches_tools(self):
        assert len(REGISTRY.specs) == 49
        assert set(SPEC_BY_NAME) == set(TOOL_BY_NAME)

    def test_interaction_tools_are_exclusive(self):
        interaction_specs = {
            spec.name: spec
            for spec in REGISTRY.specs
            if spec.execution_mode is not ExecutionMode.DIRECT
        }
        assert set(interaction_specs) == {"ask_user"}
        assert all(not spec.can_batch for spec in interaction_specs.values())
        assert interaction_specs["ask_user"].side_effect is SideEffect.HUMAN

    def test_side_effect_metadata(self):
        assert SPEC_BY_NAME["create_doc"].side_effect is SideEffect.INIT
        assert SPEC_BY_NAME["start_from_template"].side_effect is SideEffect.INIT
        assert SPEC_BY_NAME["view_text"].side_effect is SideEffect.READ
        assert SPEC_BY_NAME["validate_doc"].side_effect is SideEffect.READ
        assert SPEC_BY_NAME["query_vehicle"].side_effect is SideEffect.NONE
        assert SPEC_BY_NAME["finish"].side_effect is SideEffect.TERMINAL
