"""user_input 分类与 UserInputBridge 单元测试。"""

from __future__ import annotations

import pytest
from langchain_core.messages import AIMessage, HumanMessage

from office_agent.cli.user_input import (
    CONTINUE_PROMPT,
    PREFIX_FORCE,
    PREFIX_SUPPLEMENT,
    InputKind,
    UserInputBridge,
    classify,
    get_bridge,
    set_bridge,
)


class TestClassify:
    def test_empty_returns_none(self):
        assert classify("") is None
        assert classify("   ") is None

    def test_quit(self):
        item = classify("退出")
        assert item is not None
        assert item.kind == InputKind.QUIT

    def test_continue(self):
        item = classify("继续")
        assert item is not None
        assert item.kind == InputKind.CONTINUE
        assert item.text == CONTINUE_PROMPT

    def test_force_bang(self):
        item = classify("!改成红色标题")
        assert item is not None
        assert item.kind == InputKind.FORCE
        assert item.text == "改成红色标题"

    def test_force_prefix(self):
        item = classify("强制:立刻加目录")
        assert item is not None
        assert item.kind == InputKind.FORCE
        assert "目录" in item.text

    def test_supplement(self):
        item = classify("再加一节风险")
        assert item is not None
        assert item.kind == InputKind.SUPPLEMENT
        assert item.text == "再加一节风险"


class TestUserInputBridge:
    def test_submit_and_consume(self):
        b = UserInputBridge()
        b.submit("补充内容")
        b.submit("!强制内容")
        b.submit("继续")
        assert b.drain_soft() == ["补充内容"]
        assert b.consume_force() == "强制内容"
        assert b.consume_continue() is True
        assert b.consume_force() is None
        assert b.consume_continue() is False

    def test_quit_and_soft_pause(self):
        b = UserInputBridge()
        b.submit("退出")
        assert b.consume_quit() is True
        b.request_soft_pause()
        assert b.peek_soft_pause() is True
        assert b.consume_soft_pause() is True
        assert b.consume_soft_pause() is False

    def test_blocking_readline_from_raw(self):
        b = UserInputBridge()
        b._raw.put("hello")
        assert b.blocking_readline() == "hello"

    def test_blocking_readline_eof(self):
        b = UserInputBridge()
        b._raw.put(None)
        with pytest.raises(EOFError):
            b.blocking_readline()

    def test_set_get_bridge(self):
        b = UserInputBridge()
        set_bridge(b)
        try:
            assert get_bridge() is b
        finally:
            set_bridge(None)
        assert get_bridge() is None


class TestAgentNodeInjectsPending:
    """agent 节点边界消费 soft/force。"""

    def test_injects_force_and_soft(self, monkeypatch):
        from office_agent.agent import graph as graph_mod

        bridge = UserInputBridge()
        bridge.submit("补充A")
        bridge.submit("!强制B")
        set_bridge(bridge)

        captured: list = []

        class FakeLLM:
            def bind_tools(self, tools):
                return self

            def invoke(self, messages):
                captured.extend(messages)
                return AIMessage(content="ok", tool_calls=[])

        monkeypatch.setattr(graph_mod, "get_llm", lambda: FakeLLM())
        monkeypatch.setattr(graph_mod, "ALL_TOOLS", [])
        monkeypatch.setattr(
            graph_mod,
            "build_system_prompt",
            lambda *a, **k: "sys",
        )

        try:
            node = graph_mod._agent_node_factory("out.docx")
            result = node({"messages": [HumanMessage(content="原始需求")]})
        finally:
            set_bridge(None)

        msgs = result["messages"]
        assert any(
            isinstance(m, HumanMessage) and m.content.startswith(PREFIX_SUPPLEMENT)
            for m in msgs
        )
        assert any(
            isinstance(m, HumanMessage) and m.content.startswith(PREFIX_FORCE)
            for m in msgs
        )
        assert isinstance(msgs[-1], AIMessage)

        # prompt 侧也应看到注入
        assert any(
            isinstance(m, HumanMessage) and PREFIX_FORCE in m.content for m in captured
        )
