"""OpenAI 兼容协议单元测试。"""

from __future__ import annotations

import dataclasses
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest
from fastapi.testclient import TestClient

from office_agent.api.app import create_app
from office_agent.api.openai_compat import (
    ChatMessage,
    extract_session_id_from_messages,
    format_event,
    session_marker,
    user_text_to_message,
)
from office_agent.session.runner import AgentSession, SessionPhase


@pytest.fixture
def client(tmp_path, monkeypatch):
    from office_agent import config
    from office_agent.session import prep

    new_settings = dataclasses.replace(
        config.settings, output_dir=tmp_path, llm_api_key="test-key"
    )
    monkeypatch.setattr(config, "settings", new_settings)
    monkeypatch.setattr(prep, "settings", new_settings)
    app = create_app()
    with TestClient(app) as c:
        yield c


class TestHelpers:
    def test_extract_session_id(self):
        sid = "12345678-1234-1234-1234-123456789abc"
        msgs = [
            ChatMessage(role="user", content="hi"),
            ChatMessage(
                role="assistant",
                content=f"{session_marker(sid)}\nhello",
            ),
            ChatMessage(role="user", content="next"),
        ]
        assert extract_session_id_from_messages(msgs) == sid

    def test_user_text_outline_approve(self):
        s = AgentSession()
        s.phase = SessionPhase.AWAITING_OUTLINE
        assert user_text_to_message(s, "批准")["action"] == "approve"
        assert user_text_to_message(s, "修改 加一节风险")["feedback"] == "加一节风险"

    def test_user_text_interrupt_json(self):
        s = AgentSession()
        s.phase = SessionPhase.AWAITING_INTERRUPT
        msg = user_text_to_message(s, '{"decision":"确认生成"}')
        assert msg["type"] == "resume"
        assert msg["answers"]["decision"] == "确认生成"

    def test_format_need_kind(self):
        text = format_event(
            {
                "type": "need_kind",
                "options": [{"id": "docx", "label": "Word"}],
            }
        )
        assert "docx" in text


class TestOpenAIEndpoints:
    def test_models(self, client):
        r = client.get("/v1/models")
        assert r.status_code == 200
        ids = [m["id"] for m in r.json()["data"]]
        assert "office-agent" in ids

    def test_chat_completions_need_kind(self, client, monkeypatch):
        # start ambiguous → need_kind，不碰 LLM graph
        r = client.post(
            "/v1/chat/completions",
            json={
                "model": "office-agent",
                "messages": [{"role": "user", "content": "弄点东西"}],
                "stream": False,
            },
            headers={"Authorization": "Bearer test"},
        )
        assert r.status_code == 200
        body = r.json()
        content = body["choices"][0]["message"]["content"]
        assert "office-agent-session:" in content
        assert "docx" in content or "文档类型" in content
        assert body.get("session_id")
        assert r.headers.get("x-session-id") == body["session_id"]

    def test_chat_completions_continue_choose_kind(self, client, monkeypatch):
        graph = MagicMock()
        graph.stream.return_value = iter([])
        graph.get_state.return_value = SimpleNamespace(
            values={"done": True, "summary": "ok"},
            tasks=[],
        )
        monkeypatch.setattr(
            "office_agent.agent.graph.build_graph", lambda *a, **k: graph
        )
        monkeypatch.setattr(
            "office_agent.session.runner.pending_interrupt", lambda g, c: None
        )

        r1 = client.post(
            "/v1/chat/completions",
            json={
                "model": "office-agent",
                "messages": [{"role": "user", "content": "弄点东西"}],
            },
        )
        assert r1.status_code == 200
        sid = r1.json()["session_id"]
        assistant = r1.json()["choices"][0]["message"]["content"]

        r2 = client.post(
            "/v1/chat/completions",
            json={
                "model": "office-agent",
                "messages": [
                    {"role": "user", "content": "弄点东西"},
                    {"role": "assistant", "content": assistant},
                    {"role": "user", "content": "xlsx"},
                ],
            },
            headers={"X-Session-Id": sid},
        )
        assert r2.status_code == 200
        content = r2.json()["choices"][0]["message"]["content"]
        assert "文档已生成" in content or "download" in content

    def test_stream(self, client, monkeypatch):
        r = client.post(
            "/v1/chat/completions",
            json={
                "model": "office-agent",
                "stream": True,
                "messages": [{"role": "user", "content": "弄点东西"}],
            },
        )
        assert r.status_code == 200
        assert "text/event-stream" in r.headers["content-type"]
        raw = r.text
        assert "data: " in raw
        assert "[DONE]" in raw
