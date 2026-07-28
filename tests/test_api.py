"""FastAPI HTTP / WebSocket 协议测试。"""

from __future__ import annotations

import dataclasses
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest
from fastapi.testclient import TestClient

from office_agent.api import manager as session_manager
from office_agent.api.app import create_app
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
        yield c, tmp_path


class TestHealthAndSessionHttp:
    def test_health(self, client):
        c, _ = client
        r = c.get("/health")
        assert r.status_code == 200
        body = r.json()
        assert body["status"] == "ok"
        assert "llm_configured" in body

    def test_session_not_found(self, client):
        c, _ = client
        r = c.get("/api/v1/sessions/no-such")
        assert r.status_code == 404

    def test_download(self, client):
        c, tmp_path = client
        doc = tmp_path / "demo.docx"
        doc.write_bytes(b"PK\x03\x04fake")
        s = AgentSession(session_id="sid-1")
        s.doc_path = str(doc)
        s.phase = SessionPhase.DONE
        session_manager.register(s)

        r = c.get("/api/v1/sessions/sid-1/download")
        assert r.status_code == 200
        assert r.content.startswith(b"PK")

        info = c.get("/api/v1/sessions/sid-1")
        assert info.status_code == 200
        assert info.json()["session_id"] == "sid-1"

        session_manager.remove("sid-1")

    def test_download_path_traversal_blocked(self, client):
        c, _ = client
        outside = Path("/tmp/evil.docx")
        s = AgentSession(session_id="sid-2")
        s.doc_path = str(outside)
        s.phase = SessionPhase.DONE
        session_manager.register(s)
        r = c.get("/api/v1/sessions/sid-2/download")
        assert r.status_code in (403, 404)
        session_manager.remove("sid-2")


class TestWebSocketProtocol:
    def test_need_kind_then_done(self, client, monkeypatch):
        c, _ = client
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

        with c.websocket_connect("/api/v1/ws") as ws:
            ws.send_text(json.dumps({"type": "start", "requirement": "弄点东西"}))
            types = []
            while True:
                ev = ws.receive_json()
                types.append(ev["type"])
                if ev["type"] in ("need_kind", "error"):
                    break
            assert "need_kind" in types

            ws.send_text(json.dumps({"type": "choose_kind", "kind": "xlsx"}))
            done = None
            for _ in range(20):
                ev = ws.receive_json()
                if ev["type"] == "done":
                    done = ev
                    break
            assert done is not None
            assert "download_url" in done

    def test_reconnect_with_same_session_id(self, client, monkeypatch):
        c, _ = client
        graph = MagicMock()
        graph.stream.return_value = iter([])
        graph.get_state.return_value = SimpleNamespace(
            values={"done": False, "summary": ""},
            tasks=[],
        )
        monkeypatch.setattr(
            "office_agent.agent.graph.build_graph", lambda *a, **k: graph
        )
        monkeypatch.setattr(
            "office_agent.session.runner.pending_interrupt", lambda g, c: None
        )

        sid = "reconnect-sid"
        with c.websocket_connect("/api/v1/ws") as ws1:
            ws1.send_text(
                json.dumps(
                    {
                        "type": "start",
                        "session_id": sid,
                        "requirement": "弄点东西",
                    }
                )
            )
            first = ws1.receive_json()
            assert first["type"] == "session"
            assert first["session_id"] == sid

        with c.websocket_connect("/api/v1/ws") as ws2:
            ws2.send_text(json.dumps({"type": "start", "session_id": sid}))
            ev = ws2.receive_json()
            assert ev["type"] == "session"
            assert ev["reconnected"] is True
            assert ev["session_id"] == sid

    def test_rabbitmq_transport_mode_with_fake_bridge(self, tmp_path, monkeypatch):
        from office_agent import config
        from office_agent.session import prep

        new_settings = dataclasses.replace(
            config.settings,
            output_dir=tmp_path,
            llm_api_key="test-key",
            transport_mode="rabbitmq_stomp",
        )
        monkeypatch.setattr(config, "settings", new_settings)
        monkeypatch.setattr(prep, "settings", new_settings)

        class _FakeBridge:
            def __init__(self):
                self.inbound = []
                self.outbound = {}

            def publish_inbound(self, envelope):
                self.inbound.append(envelope)

            def process_next(self):
                if not self.inbound:
                    return False
                env = self.inbound.pop(0)
                self.outbound.setdefault(env.session_id, []).append(
                    {"type": "session", "session_id": env.session_id}
                )
                self.outbound[env.session_id].append(
                    {"type": "done", "session_id": env.session_id}
                )
                return True

            async def consume_outbound(self, session_id, timeout_s=0.2):
                _ = timeout_s
                return list(self.outbound.pop(session_id, []))

            def close(self):
                return None

        monkeypatch.setattr("office_agent.api.app.build_stomp_bridge", _FakeBridge)
        app = create_app()
        with TestClient(app) as c:
            with c.websocket_connect("/api/v1/ws") as ws:
                ws.send_text(json.dumps({"type": "start", "requirement": "测试"}))
                first = ws.receive_json()
                second = ws.receive_json()
                assert first["type"] == "session"
                assert second["type"] == "done"
