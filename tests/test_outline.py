"""agent.outline 单元测试（mock LLM）。"""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from office_agent.agent import outline as outline_mod


class TestGenerateOutline:
    def test_returns_text(self, monkeypatch):
        class FakeLLM:
            def invoke(self, messages):
                return SimpleNamespace(content="# 标题\n## 第一节\n- 要点 A")

        monkeypatch.setattr(outline_mod, "get_llm", lambda: FakeLLM())
        text = outline_mod.generate_outline("写一份周报")
        assert "# 标题" in text
        assert "要点 A" in text

    def test_strips_code_fence(self, monkeypatch):
        class FakeLLM:
            def invoke(self, messages):
                return SimpleNamespace(content="```markdown\n# A\n```")

        monkeypatch.setattr(outline_mod, "get_llm", lambda: FakeLLM())
        assert outline_mod.generate_outline("需求") == "# A"

    def test_feedback_in_prompt(self, monkeypatch):
        captured: list = []

        class FakeLLM:
            def invoke(self, messages):
                captured.extend(messages)
                return SimpleNamespace(content="# 修订版")

        monkeypatch.setattr(outline_mod, "get_llm", lambda: FakeLLM())
        outline_mod.generate_outline("周报", feedback="增加风险一节")
        user = str(captured[-1].content)
        assert "增加风险一节" in user
        assert "修改意见" in user

    def test_doc_type_hint(self, monkeypatch):
        captured: list = []

        class FakeLLM:
            def invoke(self, messages):
                captured.extend(messages)
                return SimpleNamespace(content="# 通知大纲")

        monkeypatch.setattr(outline_mod, "get_llm", lambda: FakeLLM())
        outline_mod.generate_outline("发个通知", doc_type="通知")
        system = str(captured[0].content)
        assert "通知" in system
        assert "法定公文" in system

    def test_empty_requirement_raises(self):
        with pytest.raises(ValueError):
            outline_mod.generate_outline("  ")

    def test_empty_response_raises(self, monkeypatch):
        class FakeLLM:
            def invoke(self, messages):
                return SimpleNamespace(content="   ")

        monkeypatch.setattr(outline_mod, "get_llm", lambda: FakeLLM())
        with pytest.raises(RuntimeError):
            outline_mod.generate_outline("周报")
