"""tools 包测试：session 路由 + 工具转发 + start_from_template 编排。

通过 set_session_doc 注入会话路径，monkeypatch merge_template / DocTool 方法
断言转发逻辑，不碰真实 officecli.exe。
"""

from __future__ import annotations

import pytest

from office_agent.tools import (
    TOOL_BY_NAME,
    session_doc_kind,
    session_doc_path,
    set_session_doc,
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

    def test_add_table_in_xlsx_redirects_to_set_cells(self, xlsx_session):
        """Excel 会话下 add_table 引导改用 set_cells（历史版本会 AttributeError）。"""
        result = TOOL_BY_NAME["add_table"].invoke({"data": [["a", "b"], ["1", "2"]]})
        assert "set_cells" in result


class TestTypedAccessors:
    """doc_tool/excel_tool/pptx_tool：类型化访问器的运行时校验。"""

    def test_doc_tool_in_docx(self, doc_session):
        from office_agent.office.doc import DocTool
        from office_agent.tools.session import doc_tool

        assert isinstance(doc_tool(), DocTool)

    def test_doc_tool_in_xlsx_raises(self, xlsx_session):
        from office_agent.office.runner import OfficeCLIError
        from office_agent.tools.session import doc_tool

        with pytest.raises(OfficeCLIError):
            doc_tool()

    def test_excel_tool_in_pptx_raises(self, pptx_session):
        from office_agent.office.runner import OfficeCLIError
        from office_agent.tools.session import excel_tool

        with pytest.raises(OfficeCLIError):
            excel_tool()

    def test_pptx_tool_in_pptx(self, pptx_session):
        from office_agent.office.pptx import PptxTool
        from office_agent.tools.session import pptx_tool

        assert isinstance(pptx_tool(), PptxTool)


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
