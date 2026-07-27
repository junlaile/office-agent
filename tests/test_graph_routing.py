"""graph.py 路由与工具节点测试（不碰 LLM）。

构造假 AIMessage state，验证 _route_after_agent / _route_after_tools / _tools_node
/ _nudge_node 的路由与短路逻辑。
"""

from __future__ import annotations

from langchain_core.messages import AIMessage, SystemMessage, ToolMessage
from langgraph.graph import END

from office_agent.agent.graph import (
    _MAX_NUDGE,
    _agent_node_factory,
    _count_idle_turns,
    _nudge_node,
    _route_after_agent,
    _route_after_tools,
    _tools_node,
)


# ============================================================
# _count_idle_turns: 连续空转计数
# ============================================================
class TestCountIdleTurns:
    def test_single_idle_ai_message(self):
        assert _count_idle_turns([AIMessage(content="我先问问")]) == 1

    def test_stops_at_tool_call_ai(self):
        """有 tool_calls 的 AIMessage 中断计数。"""
        msgs = [
            AIMessage(content="", tool_calls=[{"name": "x", "args": {}, "id": "1"}]),
            AIMessage(content="又空转了"),
        ]
        assert _count_idle_turns(msgs) == 1

    def test_stops_at_tool_message(self):
        """ToolMessage 中断计数。"""
        msgs = [
            ToolMessage(content="r", tool_call_id="1"),
            AIMessage(content="空转"),
        ]
        assert _count_idle_turns(msgs) == 1

    def test_skips_nudge_system_message(self):
        """nudge 注入的 SystemMessage 不中断计数（它夹在两次空转之间）。"""
        msgs = [
            AIMessage(content="空转1"),
            SystemMessage(content="【系统纠偏】请用工具"),
            AIMessage(content="空转2"),
        ]
        assert _count_idle_turns(msgs) == 2

    def test_empty_or_none(self):
        assert _count_idle_turns([]) == 0
        assert _count_idle_turns(None) == 0


class TestAgentToolBinding:
    @staticmethod
    def _bound_tool_names(monkeypatch, doc_path):
        captured = []

        class FakeLLM:
            def bind_tools(self, tools):
                captured.extend(tools)
                return self

        monkeypatch.setattr("office_agent.agent.graph.get_llm", lambda: FakeLLM())
        _agent_node_factory(doc_path)
        return {tool.name for tool in captured}

    def test_docx_binds_word_tools_only(self, monkeypatch):
        names = self._bound_tool_names(monkeypatch, "report.docx")
        assert {"add_title", "update_paragraph", "finish"} <= names
        assert names.isdisjoint({"set_cells", "add_slide"})

    def test_xlsx_binds_excel_tools_only(self, monkeypatch):
        names = self._bound_tool_names(monkeypatch, "report.xlsx")
        assert {"set_cells", "add_excel_chart", "finish"} <= names
        assert names.isdisjoint({"add_title", "add_slide", "add_image"})

    def test_pptx_binds_presentation_tools_only(self, monkeypatch):
        names = self._bound_tool_names(monkeypatch, "report.pptx")
        assert {"add_slide", "add_image", "finish"} <= names
        assert names.isdisjoint({"add_title", "set_cells"})


# ============================================================
# _route_after_agent: agent 之后的路由
# ============================================================
class TestRouteAfterAgent:
    def test_with_tool_calls_goes_to_tools(self):
        """AI 消息含 tool_calls → tools。"""
        msg = AIMessage(content="", tool_calls=[{"name": "create_doc", "args": {}, "id": "1"}])
        assert _route_after_agent({"messages": [msg]}) == "tools"

    def test_empty_messages_goes_to_end(self):
        """无消息 → END。"""
        assert _route_after_agent({"messages": []}) == END

    def test_non_ai_message_goes_to_end(self):
        """最后一条非 AIMessage → END。"""
        tm = ToolMessage(content="result", tool_call_id="x")
        assert _route_after_agent({"messages": [tm]}) == END

    def test_idle_with_content_goes_to_nudge(self):
        """AIMessage 有内容、无 tool_calls、未达上限 → nudge（核心：纠偏而非判死）。"""
        msg = AIMessage(content="我先问您几个关键问题")  # 无 tool_calls
        assert _route_after_agent({"messages": [msg]}) == "nudge"

    def test_empty_content_ai_goes_to_end(self):
        """AIMessage 纯空串（无内容无 tool_calls）→ END（不无意义纠偏）。"""
        msg = AIMessage(content="")
        assert _route_after_agent({"messages": [msg]}) == END

    def test_idle_at_max_nudge_goes_to_end(self):
        """连续空转达 _MAX_NUDGE 上限后 → END（防死循环，放行兜底）。"""
        # 构造 _MAX_NUDGE 条连续空转 AIMessage（夹着 nudge SystemMessage）
        msgs = []
        for i in range(_MAX_NUDGE):
            msgs.append(AIMessage(content=f"空转{i + 1}"))
            msgs.append(SystemMessage(content="【系统纠偏】请用工具"))
        # 最后又空转一次：此时 _count_idle_turns = _MAX_NUDGE + 1 > 上限
        msgs.append(AIMessage(content=f"空转{_MAX_NUDGE + 1}"))
        assert _route_after_agent({"messages": msgs}) == END

    def test_nudge_then_tool_call_resets_to_tools(self):
        """nudge 后 LLM 改用 tool_calls → 回到正常 tools 路径（纠偏生效）。"""
        msgs = [
            AIMessage(content="我先问问"),  # 空转1
            SystemMessage(content="【系统纠偏】请用工具"),
            AIMessage(
                content="", tool_calls=[{"name": "ask_user", "args": {}, "id": "1"}]
            ),
        ]
        assert _route_after_agent({"messages": msgs}) == "tools"


# ============================================================
# _nudge_node: 空转纠偏节点
# ============================================================
class TestNudgeNode:
    def test_returns_system_message(self):
        """nudge 节点返回一条 SystemMessage。"""
        result = _nudge_node({"messages": [AIMessage(content="我先问问")]})
        msgs = result["messages"]
        assert len(msgs) == 1
        assert isinstance(msgs[0], SystemMessage)

    def test_hint_content_addresses_no_tool_calls(self):
        """纠偏提示明确指出"没有调用任何工具"。"""
        result = _nudge_node({"messages": []})
        content = result["messages"][0].content
        assert "没有调用任何工具" in content
        assert "工具调用" in content  # 强调要用工具

    def test_hint_suggests_actionable_tools(self):
        """纠偏提示给出可行动的工具示例。"""
        content = _nudge_node({"messages": []})["messages"][0].content
        assert "create_doc" in content
        assert "ask_user" in content


# ============================================================
# _route_after_tools: tools 之后的路由
# ============================================================
class TestRouteAfterTools:
    def test_done_goes_to_end(self):
        """done=True → END。"""
        assert _route_after_tools({"done": True}) == END

    def test_not_done_goes_to_agent(self):
        """done=False → agent。"""
        assert _route_after_tools({"done": False}) == "agent"

    def test_no_done_key_goes_to_agent(self):
        """无 done 键默认 → agent。"""
        assert _route_after_tools({}) == "agent"


# ============================================================
# _tools_node: 工具执行节点
# ============================================================
class TestToolsNode:
    def _ai_with_calls(self, calls):
        return AIMessage(content="", tool_calls=calls)

    def test_finish_short_circuits(self):
        """finish 调用：done=True，存 summary，回 ToolMessage。"""
        ai = self._ai_with_calls([{"name": "finish", "args": {"summary": "完成总结"}, "id": "f1"}])
        result = _tools_node({"messages": [ai]})
        assert result["done"] is True
        assert result["summary"] == "完成总结"
        msgs = result["messages"]
        assert len(msgs) == 1
        assert "完成总结" in msgs[0].content

    def test_unknown_tool_returns_error_message(self):
        """未知工具 → 错误 ToolMessage（不崩溃）。"""
        ai = self._ai_with_calls([{"name": "不存在的工具", "args": {}, "id": "u1"}])
        result = _tools_node({"messages": [ai]})
        msgs = result["messages"]
        assert len(msgs) == 1
        assert "未知工具" in msgs[0].content or "错误" in msgs[0].content

    def test_ask_user_alone_with_others_cancels_batch(self):
        """ask_user 与其他工具同批时，整批取消并引导单独调用 ask_user。"""
        ai = self._ai_with_calls(
            [
                {"name": "create_doc", "args": {}, "id": "c1"},
                {"name": "ask_user", "args": {"title": "问", "fields": []}, "id": "a1"},
            ]
        )
        result = _tools_node({"messages": [ai]})
        msgs = result["messages"]
        # 两条 ToolMessage（每个 tool_call 一条），都提示重新单独调用
        assert len(msgs) == 2
        contents = " ".join(m.content for m in msgs)
        assert "ask_user" in contents
        assert "单独" in contents
        # 没有 done
        assert not result.get("done")

    def test_normal_tool_executes(self):
        """正常工具（如 view_text）执行并回 ToolMessage。

        需要 session 才能调真工具——这里用一个会抛错的简单验证：
        未设 session 时工具转发出错，_tools_node 把异常包成 ToolMessage。
        """
        ai = self._ai_with_calls([{"name": "view_text", "args": {}, "id": "v1"}])
        # 不设 session → 工具内部抛 OfficeCLIError → 被捕获成 ToolMessage
        from office_agent.tools import set_session_doc

        set_session_doc(None)
        result = _tools_node({"messages": [ai]})
        msgs = result["messages"]
        assert len(msgs) == 1
        # 应是 ToolMessage（错误信息）
        assert isinstance(msgs[0], ToolMessage)
