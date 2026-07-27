"""batch 合并执行：连续"末尾追加"类调用 → 一次 officecli batch。

覆盖 tools/batching.py 的翻译逻辑与 graph._tools_node 的合并/回退路径。
用 fake_runner 断言 subprocess 层实际发出的命令。
"""

from __future__ import annotations

import json

from langchain_core.messages import AIMessage, ToolMessage

from office_agent.agent.graph import _tools_node
from office_agent.tools.batching import (
    _op_from_call,
    execute_batched,
    is_batchable,
)


def _tc(name: str, args: dict, tc_id: str) -> dict:
    return {"name": name, "args": args, "id": tc_id}


# ============================================================
# _op_from_call: 工具调用 → batch op 翻译
# ============================================================
class TestOpFromCall:
    def test_docx_add_paragraph(self):
        op = _op_from_call("add_paragraph", {"text": "正文", "bold": True}, "docx")
        assert op == {
            "command": "add",
            "path": "/body",
            "type": "paragraph",
            "props": {"text": "正文", "bold": "true"},
        }

    def test_docx_add_heading_maps_size_and_outline(self):
        op = _op_from_call("add_heading", {"text": "章", "level": 2}, "docx")
        assert op["props"]["size"] == "18"
        assert op["props"]["outlineLvl"] == "1"

    def test_docx_add_title(self):
        op = _op_from_call("add_title", {"text": "标题"}, "docx")
        assert op["props"]["align"] == "center"
        assert op["props"]["size"] == "26"

    def test_docx_add_list_item(self):
        op = _op_from_call("add_list_item", {"text": "项", "ordered": True}, "docx")
        assert op["props"]["listStyle"] == "ordered"

    def test_pptx_add_slide(self):
        op = _op_from_call("add_slide", {"title": "页", "body_text": "· 要点"}, "pptx")
        assert op == {
            "command": "add",
            "path": "/",
            "type": "slide",
            "props": {"title": "页", "text": "· 要点"},
        }

    def test_edit_tools_not_batchable(self):
        assert (
            _op_from_call("update_paragraph", {"path": "/body/p[1]", "text": "x"}, "docx") is None
        )
        assert _op_from_call("remove_paragraph", {"path": "/body/p[1]"}, "docx") is None
        assert _op_from_call("add_table", {"data": [["a"]]}, "docx") is None

    def test_kind_mismatch_not_batchable(self):
        assert _op_from_call("add_paragraph", {"text": "x"}, "pptx") is None
        assert _op_from_call("add_slide", {"title": "x"}, "docx") is None


# ============================================================
# is_batchable: 会话感知
# ============================================================
class TestIsBatchable:
    def test_true_for_docx_paragraph(self, doc_session):
        assert is_batchable(_tc("add_paragraph", {"text": "x"}, "1")) is True

    def test_false_without_session(self):
        from office_agent.tools import set_session_doc

        set_session_doc(None)
        assert is_batchable(_tc("add_paragraph", {"text": "x"}, "1")) is False

    def test_false_for_non_appending_tool(self, doc_session):
        assert is_batchable(_tc("view_text", {}, "1")) is False


# ============================================================
# execute_batched: 合并执行 + 输出解析
# ============================================================
class TestExecuteBatched:
    def test_single_call_not_batched(self, fake_runner, doc_session):
        assert execute_batched([_tc("add_paragraph", {"text": "x"}, "1")]) is None
        assert fake_runner.calls == []

    def test_two_paragraphs_one_batch_command(self, fake_runner, doc_session):
        results = execute_batched(
            [
                _tc("add_paragraph", {"text": "一"}, "1"),
                _tc("add_paragraph", {"text": "二"}, "2"),
            ]
        )
        assert results is not None
        assert [r[0] for r in results] == ["1", "2"]
        batch_calls = fake_runner.calls_of("batch")
        assert len(batch_calls) == 1
        # payload 里两条 add op，顺序与调用一致
        payload = json.loads(batch_calls[0][batch_calls[0].index("--commands") + 1])
        assert [op["props"]["text"] for op in payload] == ["一", "二"]

    def test_uses_officecli_outputs_when_parsable(self, doc_session, monkeypatch):
        from office_agent.office import runner as runner_module
        from tests.conftest import FakeRunner

        stdout = json.dumps(
            {
                "success": True,
                "data": {
                    "results": [
                        {"index": 0, "success": True, "output": "Added paragraph at /body/p[1]"},
                        {"index": 1, "success": True, "output": "Added paragraph at /body/p[2]"},
                    ]
                },
            }
        )
        monkeypatch.setattr(runner_module, "_runner", FakeRunner(default_stdout=stdout))
        results = execute_batched(
            [
                _tc("add_paragraph", {"text": "一"}, "1"),
                _tc("add_paragraph", {"text": "二"}, "2"),
            ]
        )
        assert results[0][1] == "Added paragraph at /body/p[1]"
        assert results[1][1] == "Added paragraph at /body/p[2]"

    def test_batch_failure_returns_none(self, doc_session, monkeypatch):
        from office_agent.office import runner as runner_module
        from tests.conftest import FakeRunner

        monkeypatch.setattr(runner_module, "_runner", FakeRunner(error_on="batch"))
        assert (
            execute_batched(
                [
                    _tc("add_paragraph", {"text": "一"}, "1"),
                    _tc("add_paragraph", {"text": "二"}, "2"),
                ]
            )
            is None
        )

    def test_slide_without_body_gets_warning(self, fake_runner, pptx_session):
        results = execute_batched(
            [
                _tc("add_slide", {"title": "封面", "body_text": "副标题"}, "1"),
                _tc("add_slide", {"title": "内容页"}, "2"),
            ]
        )
        assert "⚠️" not in results[0][1]
        assert "body_text" in results[1][1]  # 只有标题没正文 → 警告


# ============================================================
# _tools_node 集成：合并 + 回退 + 混合批次
# ============================================================
class TestToolsNodeCoalescing:
    def _state(self, calls):
        return {"messages": [AIMessage(content="", tool_calls=calls)]}

    def test_consecutive_adds_single_batch(self, fake_runner, doc_session):
        result = _tools_node(
            self._state(
                [
                    _tc("add_heading", {"text": "一、背景", "level": 1}, "1"),
                    _tc("add_paragraph", {"text": "正文一"}, "2"),
                    _tc("add_paragraph", {"text": "正文二"}, "3"),
                ]
            )
        )
        msgs = result["messages"]
        assert len(msgs) == 3
        assert all(isinstance(m, ToolMessage) for m in msgs)
        assert [m.tool_call_id for m in msgs] == ["1", "2", "3"]
        # 三个调用只发一条 batch 命令
        assert len(fake_runner.calls_of("batch")) == 1
        assert len(fake_runner.calls_of("add")) == 0

    def test_non_batchable_breaks_run(self, fake_runner, doc_session):
        """view_text 夹在中间：前后各自处理，顺序保持。"""
        result = _tools_node(
            self._state(
                [
                    _tc("add_paragraph", {"text": "一"}, "1"),
                    _tc("view_text", {}, "2"),
                    _tc("add_paragraph", {"text": "二"}, "3"),
                ]
            )
        )
        msgs = result["messages"]
        assert [m.tool_call_id for m in msgs] == ["1", "2", "3"]
        # 单独的 add_paragraph 不 batch（<2 个），走普通 add
        assert len(fake_runner.calls_of("batch")) == 0
        assert len(fake_runner.calls_of("add")) == 2
        assert len(fake_runner.calls_of("view")) == 1

    def test_batch_failure_falls_back_to_singles(self, doc_session, monkeypatch):
        from office_agent.office import runner as runner_module
        from office_agent.tools.batching import BATCH_FALLBACK_PREFIX
        from tests.conftest import FakeRunner

        fake = FakeRunner(error_on="batch")
        monkeypatch.setattr(runner_module, "_runner", fake)
        result = _tools_node(
            self._state(
                [
                    _tc("add_paragraph", {"text": "一"}, "1"),
                    _tc("add_paragraph", {"text": "二"}, "2"),
                ]
            )
        )
        msgs = result["messages"]
        assert [m.tool_call_id for m in msgs] == ["1", "2"]
        # batch 失败后回退：逐个 add 成功
        assert len(fake.calls_of("add")) == 2
        # 首条结果应带 batch-fallback 可观测标记，避免“偶发失败却最终成功”被静默
        assert BATCH_FALLBACK_PREFIX in msgs[0].content
        assert "回退" in msgs[0].content
        # 第二条不必重复前缀
        assert BATCH_FALLBACK_PREFIX not in msgs[1].content

    def test_finish_after_adds(self, fake_runner, doc_session):
        """add × 2 + finish：add 合并，finish 正常短路。"""
        result = _tools_node(
            self._state(
                [
                    _tc("add_paragraph", {"text": "一"}, "1"),
                    _tc("add_paragraph", {"text": "二"}, "2"),
                    _tc("finish", {"summary": "完成"}, "3"),
                ]
            )
        )
        assert result["done"] is True
        assert len(result["messages"]) == 3
        assert len(fake_runner.calls_of("batch")) == 1
