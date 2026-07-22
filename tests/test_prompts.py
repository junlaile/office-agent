"""prompts.py 单元测试：build_system_prompt 分支选择。

纯函数（依赖 templates.is_upward），不碰外部。
"""

from __future__ import annotations

import pytest

from office_agent.prompts import build_system_prompt


class TestBuildSystemPromptNormal:
    """普通模式（无 doc_type）按扩展名选分支。"""

    def test_docx_branch(self):
        p = build_system_prompt("output/report.docx")
        assert "WORD" in p
        assert "公文模式" not in p  # 普通模式
        assert "add_paragraph" in p  # Word 工作流

    def test_xlsx_branch(self):
        p = build_system_prompt("output/data.xlsx")
        assert "EXCEL" in p
        assert "set_cells" in p
        assert "公文模式" not in p

    def test_pptx_branch(self):
        p = build_system_prompt("output/deck.pptx")
        assert "PPTX" in p
        assert "add_slide" in p
        assert "公文模式" not in p

    def test_default_to_word(self):
        """无扩展名或未知扩展名默认 Word。"""
        p = build_system_prompt("output/file")
        assert "WORD" in p

    def test_contains_doc_path(self):
        """提示词末尾含当前会话文档路径。"""
        p = build_system_prompt("/tmp/my.docx")
        assert "/tmp/my.docx" in p

    def test_contains_common_rules(self):
        """含通用规则段（ask_user 约束、完成标准等）。"""
        p = build_system_prompt("output/x.docx")
        assert "通用规则" in p
        assert "ask_user" in p
        assert "finish" in p


class TestBuildSystemPromptOfficial:
    """公文模式（doc_type 非空）。"""

    def test_official_mode_marker(self):
        """公文模式含【公文】标记。"""
        p = build_system_prompt("output/x.docx", doc_type="通知")
        assert "【公文】通知" in p
        assert "公文模式" in p

    def test_official_editing_tools(self):
        """公文分支介绍三大编辑工具。"""
        p = build_system_prompt("output/x.docx", doc_type="通知")
        assert "update_paragraph" in p
        assert "replace_text" in p
        assert "remove_paragraph" in p

    def test_official_closing_phrases(self):
        """公文分支含各文种结语规范。"""
        p = build_system_prompt("output/x.docx", doc_type="通知")
        assert "请认真贯彻执行" in p
        assert "妥否，请批示" in p  # 请示结语
        assert "此复" in p  # 批复结语

    def test_official_skip_create_doc(self):
        """公文分支提示跳过 create_doc。"""
        p = build_system_prompt("output/x.docx", doc_type="通知")
        assert "create_doc" in p
        assert "跳过" in p or "不要" in p

    def test_official_mode_no_create_doc_strict_rule(self):
        """公文模式【不】输出"create_doc 必须最先单独调用一次"这条强制规则
        （它与公文分支"跳过 create_doc、先 view_text"直接冲突，曾导致 LLM
        跳过读模板直接 add_paragraph 追加到模板末尾）。"""
        p = build_system_prompt("output/x.docx", doc_type="通知")
        assert "create_doc 必须【最先单独调用一次】" not in p

    def test_normal_mode_keeps_create_doc_strict_rule(self):
        """普通模式仍保留 create_doc 强制规则（从零生成需要先建空文档）。"""
        p = build_system_prompt("output/x.docx")
        assert "create_doc 必须【最先单独调用一次】" in p

    def test_template_text_injected_into_official_prompt(self):
        """公文模式下，template_text 非空时其内容被注入提示词——
        让 LLM 第一轮就看到段落路径，不必自调 view_text。"""
        sample = "[/body/p[4]] 关于做好XX工作的通知\n[/body/p[5]] 各区公安分局："
        p = build_system_prompt("output/x.docx", doc_type="通知", template_text=sample)
        assert sample in p
        # 注入引导语
        assert "已为你读取的模板正文" in p
        assert "编辑前【不必】再调 view_text" in p

    def test_template_text_ignored_in_normal_mode(self):
        """普通模式忽略 template_text（它只对公文模式有意义）。"""
        sample = "[/body/p[4]] 范例"
        p = build_system_prompt("output/x.docx", template_text=sample)
        assert sample not in p

    def test_empty_template_text_not_injected(self):
        """template_text 为空时公文分支不追加注入段（保持原提示词形态）。"""
        p_empty = build_system_prompt("output/x.docx", doc_type="通知", template_text="")
        p_blank = build_system_prompt("output/x.docx", doc_type="通知", template_text="   ")
        assert "已为你读取的模板正文" not in p_empty
        assert "已为你读取的模板正文" not in p_blank

    def test_upward_note_present(self):
        """上行文含'上行文特别提示'。"""
        p = build_system_prompt("output/x.docx", doc_type="请示")
        assert "上行文特别提示" in p

    def test_non_upward_note(self):
        """非上行文含'非上行文'说明。"""
        p = build_system_prompt("output/x.docx", doc_type="通知")
        assert "非上行文" in p

    @pytest.mark.parametrize("doc_type", ["通知", "请示", "函", "纪要", "批复", "公告"])
    def test_each_doc_type_renders(self, doc_type):
        """各文种都能正常渲染（不抛错）。"""
        p = build_system_prompt("output/x.docx", doc_type=doc_type)
        assert doc_type in p
        assert len(p) > 500  # 内容充实

    def test_official_vs_normal_different(self):
        """公文模式与普通模式提示词明显不同。"""
        normal = build_system_prompt("output/x.docx")
        official = build_system_prompt("output/x.docx", doc_type="通知")
        assert normal != official
        assert len(official) > len(normal)
