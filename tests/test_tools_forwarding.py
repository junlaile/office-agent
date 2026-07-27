"""tools 包的转发测试：@tool 函数 → DocTool/ExcelTool/PptxTool 方法的 argv 转发。

补足 test_officecli_argv（测 *Tool 类）和 test_tools_session（测守卫）之间的空白：
验证 LLM 工具层 @tool 函数成功调用底层 *Tool 方法（mock runner 断言 argv）。
"""

from __future__ import annotations

from office_agent.tools import TOOL_BY_NAME


class TestDocToolsForwarding:
    """Word @tool 函数成功转发到 DocTool。"""

    def test_add_title_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_title"].invoke({"text": "标题"})
        # 应有 add 调用含 text=标题 size=26
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert add_calls
        assert any("text=标题" in c for c in add_calls)

    def test_add_paragraph_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_paragraph"].invoke({"text": "正文"})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("text=正文" in c for c in add_calls)

    def test_add_heading_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_heading"].invoke({"text": "章", "level": 2})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("text=章" in c and "size=18" in c for c in add_calls)

    def test_add_list_item_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_list_item"].invoke({"text": "项", "ordered": True})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("listStyle=ordered" in c for c in add_calls)

    def test_update_paragraph_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["update_paragraph"].invoke({"path": "/body/p[4]", "text": "新"})
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("text=新" in c and "/body/p[4]" in c for c in set_calls)

    def test_replace_text_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["replace_text"].invoke({"find": "XX", "replace": "新", "path": ""})
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("--find" in c and "XX" in c for c in set_calls)

    def test_replace_text_with_path(self, fake_runner, doc_session):
        TOOL_BY_NAME["replace_text"].invoke({"find": "a", "replace": "b", "path": "/body/p[5]"})
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("/body/p[5]" in c for c in set_calls)

    def test_remove_paragraph_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["remove_paragraph"].invoke({"path": "/body/p[10]"})
        remove_calls = [c for c in fake_runner.calls if c[0] == "remove"]
        assert any("/body/p[10]" in c for c in remove_calls)

    def test_add_toc_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_toc"].invoke({"levels": "1-3", "title": "目录"})
        # add toc（随后会清 updateFields；FakeRunner 无真实 zip，仅断言 toc 调用）
        all_calls = fake_runner.calls
        assert any("toc" in str(c) for c in all_calls)

    def test_add_page_number_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_page_number"].invoke({"align": "center"})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("footer" in str(c).lower() for c in add_calls)

    def test_add_header_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_header"].invoke({"text": "页眉", "align": "right"})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("header" in str(c).lower() and "页眉" in str(c) for c in add_calls)

    def test_add_footer_forwards(self, fake_runner, doc_session):
        TOOL_BY_NAME["add_footer"].invoke({"text": "页脚"})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("footer" in str(c).lower() for c in add_calls)

    def test_view_text_forwards(self, fake_runner, doc_session):
        fake_runner.default_stdout = "[/body/p[1]] 内容"
        result = TOOL_BY_NAME["view_text"].invoke({})
        assert "内容" in result
        view_calls = [c for c in fake_runner.calls if c[0] == "view"]
        assert view_calls

    def test_validate_doc_forwards(self, fake_runner, doc_session):
        fake_runner.default_stdout = "Validation passed"
        result = TOOL_BY_NAME["validate_doc"].invoke({})
        assert "passed" in result.lower()

    def test_create_doc_forwards(self, fake_runner, doc_session, tmp_path):
        """create_doc：doc_session 路径的文件可能不存在 → 调 create --force。"""
        # doc_session 路径在 tmp_path/test.docx，不存在
        TOOL_BY_NAME["create_doc"].invoke({})
        create_calls = [c for c in fake_runner.calls if c[0] == "create"]
        assert create_calls

    def test_add_image_valid_source_forwards(self, fake_runner, doc_session):
        """有效图片源（data URI）→ 转发到 officecli（调 runner）。"""
        result = TOOL_BY_NAME["add_image"].invoke(
            {"url_or_path": "data:image/png;base64,iVBORw0KGgo=", "width": "8cm"}
        )
        # 应有 add paragraph（载段）+ add picture 两条调用
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("picture" in str(c) for c in add_calls)
        # data URI 通过预校验，返回值不是跳过提示
        assert "跳过" not in result

    def test_add_image_bad_source_skips_without_touching_doc(self, fake_runner, doc_session):
        """坏图片源（本地不存在）→ 预校验拦住，不碰文档（不调 runner）。"""
        result = TOOL_BY_NAME["add_image"].invoke(
            {"url_or_path": "/definitely/nonexistent/image.png"}
        )
        # 返回跳过提示
        assert "跳过" in result
        assert "不存在" in result
        assert "不要重试" in result
        # 关键：完全没碰文档——runner 零调用（连载段都没建）
        assert fake_runner.calls == []


class TestExcelToolsForwarding:
    def test_set_cell_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["set_cell"].invoke({"sheet": "S", "ref": "A1", "value": "x", "bold": True})
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        assert any("/S/A1" in c and "value=x" in c for c in set_calls)

    def test_add_sheet_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["add_sheet"].invoke({"name": "数据"})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("name=数据" in c for c in add_calls)

    def test_set_formula_forwards(self, fake_runner, xlsx_session):
        TOOL_BY_NAME["set_formula"].invoke({"sheet": "S", "ref": "D2", "formula": "SUM(A1:A2)"})
        set_calls = [c for c in fake_runner.calls if c[0] == "set"]
        # argv 是 list，formula 在 "--prop" "formula=..." 元素里
        assert any("formula=SUM(A1:A2)" in " ".join(c) for c in set_calls)


class TestPptxToolsForwarding:
    def test_add_slide_forwards(self, fake_runner, pptx_session):
        TOOL_BY_NAME["add_slide"].invoke({"title": "T", "body_text": "B"})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("title=T" in c and "text=B" in c for c in add_calls)

    def test_add_textbox_forwards(self, fake_runner, pptx_session):
        TOOL_BY_NAME["add_textbox"].invoke({"text": "内容", "size": 18})
        add_calls = [c for c in fake_runner.calls if c[0] == "add"]
        assert any("textbox" in str(c).lower() and "内容" in str(c) for c in add_calls)


class TestFinishAndAskUser:
    def test_finish_returns_summary(self):
        """finish 不依赖 session，直接返回。"""
        result = TOOL_BY_NAME["finish"].invoke({"summary": "完成总结"})
        assert "完成总结" in result

    def test_query_vehicle_returns_dict(self):
        """query_vehicle 返回 dict（确定性）。"""
        result = TOOL_BY_NAME["query_vehicle"].invoke({"plate_number": "京A12345"})
        assert isinstance(result, dict)
        assert "status" in result
