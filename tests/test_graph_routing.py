"""graph.py 路由与工具节点测试（不碰 LLM）。

构造假 AIMessage state，验证 _route_after_agent / _route_after_tools / _tools_node
的路由与短路逻辑。
"""

from __future__ import annotations

from langchain_core.messages import AIMessage, ToolMessage
from langgraph.graph import END

from office_agent.graph import (
    _route_after_agent,
    _route_after_tools,
    _tools_node,
)


# ============================================================
# _route_after_agent: agent 之后的路由
# ============================================================
class TestRouteAfterAgent:
    def test_with_tool_calls_goes_to_tools(self):
        """AI 消息含 tool_calls → tools。"""
        msg = AIMessage(content="", tool_calls=[{"name": "create_doc", "args": {}, "id": "1"}])
        assert _route_after_agent({"messages": [msg]}) == "tools"

    def test_without_tool_calls_goes_to_end(self):
        """AI 消息无 tool_calls → END。"""
        msg = AIMessage(content="完成了")
        assert _route_after_agent({"messages": [msg]}) == END

    def test_empty_messages_goes_to_end(self):
        """无消息 → END。"""
        assert _route_after_agent({"messages": []}) == END

    def test_non_ai_message_goes_to_end(self):
        """最后一条非 AIMessage → END。"""
        tm = ToolMessage(content="result", tool_call_id="x")
        assert _route_after_agent({"messages": [tm]}) == END


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
