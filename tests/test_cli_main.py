"""cli.main 辅助逻辑单元测试（不跑完整 agent）。"""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

import office_agent.cli.main as cli_main_mod
from office_agent.cli.user_input import (
    CONTINUE_PROMPT,
    PREFIX_FORCE,
    PREFIX_SUPPLEMENT,
    UserInputBridge,
)


class TestHandlePendingInputs:
    def test_quit_returns_zero(self, capsys):
        bridge = UserInputBridge()
        bridge.submit("退出")
        graph = MagicMock()
        rc = cli_main_mod._handle_pending_inputs(graph, {}, bridge)
        assert rc == 0
        graph.update_state.assert_not_called()

    def test_injects_force_soft_continue(self, monkeypatch):
        bridge = UserInputBridge()
        bridge.submit("补充X")
        bridge.submit("!强制Y")
        bridge.submit("继续")

        graph = MagicMock()
        streamed: list = []

        def fake_stream(g, initial, first_input, config, *, command=None):
            streamed.append(True)

        monkeypatch.setattr(cli_main_mod, "_stream", fake_stream)

        # 第一次注入后 stream；第二次循环无 pending → None
        rc = cli_main_mod._handle_pending_inputs(graph, {"configurable": {}}, bridge)
        assert rc is None
        graph.update_state.assert_called_once()
        msgs = graph.update_state.call_args[0][1]["messages"]
        assert any(PREFIX_SUPPLEMENT in m.content for m in msgs)
        assert any(PREFIX_FORCE in m.content for m in msgs)
        assert any(m.content == CONTINUE_PROMPT for m in msgs)
        assert streamed == [True]


class TestPrintFinalResult:
    def test_done(self, tmp_path, capsys):
        doc = tmp_path / "a.docx"
        doc.write_text("x", encoding="utf-8")
        graph = MagicMock()
        graph.get_state.return_value = SimpleNamespace(
            values={"done": True, "summary": "完成"}
        )
        assert cli_main_mod._print_final_result(graph, {}, str(doc)) == 0
        assert "文档已生成" in capsys.readouterr().out

    def test_exists_without_finish(self, tmp_path, capsys):
        doc = tmp_path / "b.docx"
        doc.write_text("x", encoding="utf-8")
        graph = MagicMock()
        graph.get_state.return_value = SimpleNamespace(values={"done": False})
        assert cli_main_mod._print_final_result(graph, {}, str(doc)) == 0
        assert "未显式 finish" in capsys.readouterr().out

    def test_missing(self, tmp_path, capsys):
        graph = MagicMock()
        graph.get_state.return_value = SimpleNamespace(values={})
        assert cli_main_mod._print_final_result(graph, {}, str(tmp_path / "no.docx")) == 1


class TestReadline:
    def test_fallback_input(self, monkeypatch):
        from office_agent.cli import ui

        monkeypatch.setattr(ui, "get_bridge", lambda: None)
        monkeypatch.setattr("builtins.input", lambda prompt="": "hi")
        assert ui._readline("p") == "hi"

    def test_via_bridge(self, monkeypatch):
        from office_agent.cli import ui

        b = UserInputBridge()
        b._raw.put("via-bridge")
        monkeypatch.setattr(ui, "get_bridge", lambda: b)
        assert ui._readline() == "via-bridge"


class TestFormatLabel:
    def test_labels(self):
        from office_agent.domain.format import format_label, infer_doc_kind

        assert format_label("docx") == "Word"
        assert format_label("xlsx") == "Excel"
        kind, score = infer_doc_kind("做一份 Excel")
        assert kind == "xlsx"
        assert score >= 1


class TestOutlineGateInRun:
    """_run_with_bridge 对 Word 走大纲门控，Excel 跳过。"""

    def _stub_common(self, monkeypatch, tmp_path, *, doc_path: str):
        monkeypatch.setattr(cli_main_mod, "_banner", lambda: None)
        monkeypatch.setattr(cli_main_mod, "assert_llm_ready", lambda: None)
        monkeypatch.setattr(cli_main_mod, "_check_officecli", lambda: True)
        monkeypatch.setattr(cli_main_mod, "_read_requirement", lambda: "需求")
        monkeypatch.setattr(cli_main_mod, "detect_doc_type", lambda r: None)
        monkeypatch.setattr(cli_main_mod, "_derive_doc_path", lambda r, doc_type=None: doc_path)
        monkeypatch.setattr(cli_main_mod, "set_session_doc", lambda p: None)
        monkeypatch.setattr(cli_main_mod, "_stream", lambda *a, **k: None)
        monkeypatch.setattr(cli_main_mod, "_handle_pending_inputs", lambda *a, **k: None)
        monkeypatch.setattr(cli_main_mod, "_handle_interrupt", lambda *a, **k: None)
        monkeypatch.setattr(
            cli_main_mod, "_print_final_result", lambda *a, **k: 0
        )

        fake_graph = MagicMock()
        fake_graph.get_state.return_value = SimpleNamespace(values={"done": True})

        import sys
        from types import ModuleType

        # build_graph 是函数内延迟 import
        calls: list = []

        def fake_build_graph(*args, **kwargs):
            calls.append(kwargs)
            return fake_graph

        fake_mod = ModuleType("office_agent.agent.graph")
        fake_mod.build_graph = fake_build_graph
        monkeypatch.setitem(sys.modules, "office_agent.agent.graph", fake_mod)
        return calls

    def test_docx_calls_outline_then_build(self, monkeypatch, tmp_path):
        doc = str(tmp_path / "a.docx")
        calls = self._stub_common(monkeypatch, tmp_path, doc_path=doc)
        outline_called = []

        def fake_loop(req, doc_type=None):
            outline_called.append(True)
            return "# 批准大纲"

        monkeypatch.setattr(cli_main_mod, "_run_outline_approval_loop", fake_loop)
        monkeypatch.setattr(cli_main_mod, "_prepare_official_doc", lambda *a, **k: (None, ""))

        bridge = UserInputBridge()
        rc = cli_main_mod._run_with_bridge(bridge)
        assert rc == 0
        assert outline_called == [True]
        assert calls and calls[0].get("approved_outline") == "# 批准大纲"

    def test_docx_cancel_skips_build(self, monkeypatch, tmp_path):
        doc = str(tmp_path / "a.docx")
        calls = self._stub_common(monkeypatch, tmp_path, doc_path=doc)
        monkeypatch.setattr(
            cli_main_mod, "_run_outline_approval_loop", lambda *a, **k: None
        )
        bridge = UserInputBridge()
        rc = cli_main_mod._run_with_bridge(bridge)
        assert rc == 0
        assert calls == []

    def test_xlsx_skips_outline(self, monkeypatch, tmp_path):
        doc = str(tmp_path / "a.xlsx")
        calls = self._stub_common(monkeypatch, tmp_path, doc_path=doc)
        outline_called = []

        def fake_loop(*a, **k):
            outline_called.append(True)
            return "# should not run"

        monkeypatch.setattr(cli_main_mod, "_run_outline_approval_loop", fake_loop)
        bridge = UserInputBridge()
        rc = cli_main_mod._run_with_bridge(bridge)
        assert rc == 0
        assert outline_called == []
        assert calls and calls[0].get("approved_outline") == ""
