"""cli_ui.py 的 UI 输出函数测试（用 capsys 捕获 print）。

这些函数（_banner / _print_agent_step / _print_tool_results / _check_officecli）
都往 stdout 输出，用 capsys 断言。
"""

from __future__ import annotations

from langchain_core.messages import AIMessage, ToolMessage

from office_agent.cli import ui as cli_ui
from office_agent.cli.ui import (
    _banner,
    _check_officecli,
    _print_agent_step,
    _print_tool_results,
)


class TestBanner:
    def test_banner_prints(self, capsys):
        _banner()
        out = capsys.readouterr().out
        assert "Office Agent" in out
        assert "Word" in out
        assert "Excel" in out


class TestPrintAgentStep:
    def test_prints_tool_calls(self, capsys):
        """含 tool_calls 的 AIMessage 打印工具调用。"""
        msg = AIMessage(
            content="思考中",
            tool_calls=[{"name": "add_paragraph", "args": {"text": "x"}, "id": "1"}],
        )
        _print_agent_step(msg)
        out = capsys.readouterr().out
        assert "思考" in out
        assert "add_paragraph" in out

    def test_ask_user_icon(self, capsys):
        """ask_user 用 ❓ 图标。"""
        msg = AIMessage(
            content="",
            tool_calls=[{"name": "ask_user", "args": {"title": "问", "fields": []}, "id": "1"}],
        )
        _print_agent_step(msg)
        out = capsys.readouterr().out
        assert "❓" in out

    def test_finish_icon(self, capsys):
        """finish 用 ✅ 图标。"""
        msg = AIMessage(
            content="", tool_calls=[{"name": "finish", "args": {"summary": "完"}, "id": "1"}]
        )
        _print_agent_step(msg)
        out = capsys.readouterr().out
        assert "✅" in out

    def test_start_from_template_icon(self, capsys):
        """start_from_template 用 📋 图标。"""
        msg = AIMessage(
            content="",
            tool_calls=[{"name": "start_from_template", "args": {"doc_type": "通知"}, "id": "1"}],
        )
        _print_agent_step(msg)
        out = capsys.readouterr().out
        assert "📋" in out

    def test_no_content_no_error(self, capsys):
        """无 content 不打印思考行。"""
        msg = AIMessage(content="", tool_calls=[])
        _print_agent_step(msg)  # 不应抛错
        capsys.readouterr()  # 清空


class TestPrintToolResults:
    def test_prints_content(self, capsys):
        msg = ToolMessage(content="工具返回结果", tool_call_id="x")
        _print_tool_results(msg)
        out = capsys.readouterr().out
        assert "工具返回结果" in out

    def test_empty_content_no_print(self, capsys):
        msg = ToolMessage(content="", tool_call_id="x")
        _print_tool_results(msg)
        out = capsys.readouterr().out
        assert out == ""

    def test_long_content_truncated(self, capsys):
        """长输出截断。"""
        msg = ToolMessage(content="x" * 200, tool_call_id="x")
        _print_tool_results(msg)
        out = capsys.readouterr().out
        assert "…" in out


class TestCheckOfficecli:
    def test_returns_true_when_resolvable(self):
        """能解析到 officecli.exe 时返回 True。"""
        # 项目内 bin/officecli.exe 存在
        assert _check_officecli() is True

    def test_returns_false_when_missing(self, monkeypatch, capsys):
        """无法解析时返回 False 并打印提示。"""
        from office_agent.officecli import OfficeCLIError

        def raise_error():
            raise OfficeCLIError("找不到")

        monkeypatch.setattr(cli_ui, "resolve_bin", raise_error)
        assert _check_officecli() is False
        out = capsys.readouterr().out
        assert "fetch_officecli" in out
