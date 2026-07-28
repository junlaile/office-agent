"""AgentSession 状态机测试（不碰真 LLM / officecli）。"""

from __future__ import annotations

import dataclasses
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

import office_agent.session.runner as runner_mod
from office_agent.session.runner import AgentSession, SessionPhase


@pytest.fixture
def output_dir(monkeypatch, tmp_path):
    from office_agent import config
    from office_agent.session import prep

    new_settings = dataclasses.replace(config.settings, output_dir=tmp_path)
    monkeypatch.setattr(config, "settings", new_settings)
    monkeypatch.setattr(prep, "settings", new_settings)
    return tmp_path


def _events(it):
    return list(it)


class TestAgentSessionStart:
    def test_empty_requirement(self):
        s = AgentSession()
        evs = _events(s.start("  "))
        assert s.phase == SessionPhase.ERROR
        assert any(e["type"] == "error" for e in evs)

    def test_need_kind(self):
        s = AgentSession()
        evs = _events(s.start("弄点东西"))
        assert s.phase == SessionPhase.AWAITING_KIND
        assert any(e["type"] == "need_kind" for e in evs)

    def test_outline_flow(self, output_dir, monkeypatch):
        def fake_outline(req, feedback="", doc_type=None):
            return "# 大纲\n- a"

        monkeypatch.setattr(
            "office_agent.agent.outline.generate_outline", fake_outline
        )
        s = AgentSession()
        evs = _events(s.start("写一份项目周报文档"))
        assert s.phase == SessionPhase.AWAITING_OUTLINE
        assert any(e["type"] == "outline" and "大纲" in e["outline"] for e in evs)

        evs2 = _events(
            s.handle(
                {"type": "outline_decision", "action": "revise", "feedback": "加风险"}
            )
        )
        assert s.phase == SessionPhase.AWAITING_OUTLINE
        assert any(e["type"] == "outline" for e in evs2)

        s2 = AgentSession()
        _events(s2.start("写一份项目周报文档"))
        evs3 = _events(s2.handle({"type": "outline_decision", "action": "cancel"}))
        assert s2.phase == SessionPhase.CANCELLED
        assert any(e["type"] == "cancelled" for e in evs3)


class TestChooseKindAndExcelStart:
    def test_choose_kind_then_agent(self, output_dir, monkeypatch):
        graph = MagicMock()
        graph.stream.return_value = iter([])
        graph.get_state.return_value = SimpleNamespace(
            values={"done": True, "summary": "ok"},
            tasks=[],
        )

        monkeypatch.setattr(
            "office_agent.agent.graph.build_graph", lambda *a, **k: graph
        )
        monkeypatch.setattr(runner_mod, "pending_interrupt", lambda g, c: None)

        s = AgentSession()
        _events(s.start("弄点东西"))
        assert s.phase == SessionPhase.AWAITING_KIND
        evs = _events(s.handle({"type": "choose_kind", "kind": "xlsx"}))
        assert any(e["type"] == "doc_ready" and e["kind"] == "xlsx" for e in evs)
        assert any(e["type"] == "done" for e in evs)
        assert s.phase == SessionPhase.DONE


class TestInterruptResume:
    def test_resume_answers(self, output_dir, monkeypatch):
        graph = MagicMock()
        graph.stream.return_value = iter([])
        graph.get_state.return_value = SimpleNamespace(
            values={"done": True, "summary": "done"},
            tasks=[],
        )

        calls = {"n": 0}

        def pending(g, c):
            calls["n"] += 1
            if calls["n"] == 1:
                return {
                    "title": "采集",
                    "fields": [{"key": "name", "label": "姓名", "required": True}],
                }
            return None

        monkeypatch.setattr(
            "office_agent.agent.graph.build_graph", lambda *a, **k: graph
        )
        monkeypatch.setattr(runner_mod, "pending_interrupt", pending)

        s = AgentSession()
        _events(s.start("弄点东西"))
        _events(s.handle({"type": "choose_kind", "kind": "xlsx"}))
        assert s.phase == SessionPhase.AWAITING_INTERRUPT

        evs = _events(s.handle({"type": "resume", "answers": {"name": "张三"}}))
        assert any(e["type"] == "done" for e in evs)
        assert s.phase == SessionPhase.DONE
        assert graph.stream.call_count >= 2


class TestOfficialHeaderGate:
    def test_approve_outline_asks_header(self, output_dir, monkeypatch):
        monkeypatch.setattr(
            "office_agent.agent.outline.generate_outline",
            lambda *a, **k: "# 通知大纲",
        )
        monkeypatch.setattr(
            "office_agent.domain.templates.detect_doc_type",
            lambda req: "通知",
        )

        s = AgentSession()
        _events(s.start("写一份关于加班的通知"))
        assert s.phase == SessionPhase.AWAITING_OUTLINE
        evs = _events(s.handle({"type": "outline_decision", "action": "approve"}))
        assert s.phase == SessionPhase.AWAITING_OFFICIAL_HEADER
        assert any(e["type"] == "need_official_header" for e in evs)
