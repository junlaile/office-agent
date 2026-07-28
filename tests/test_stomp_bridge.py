from __future__ import annotations

from office_agent.api.stomp_bridge import StompBrokerBridge, StompConfig
from office_agent.api.transport import MessageEnvelope


class _FakeConn:
    def __init__(self):
        self.sent: list[tuple[str, str, dict]] = []

    def send(self, destination, body, headers):
        self.sent.append((destination, body, headers))

    def disconnect(self):
        return None


def _bridge(monkeypatch):
    monkeypatch.setattr(StompBrokerBridge, "_connect", lambda self: _FakeConn())
    cfg = StompConfig(
        host="127.0.0.1",
        port=61613,
        login="guest",
        passcode="guest",
        vhost="/",
        inbound_destination="/queue/in",
        outbound_destination="/queue/out",
        heartbeat_ms=10000,
    )
    return StompBrokerBridge(cfg)


def test_stomp_bridge_deduplicates_by_message_id(monkeypatch):
    bridge = _bridge(monkeypatch)
    calls = {"n": 0}

    def _fake_process(env):
        calls["n"] += 1
        return [{"type": "ok", "session_id": env.session_id}]

    monkeypatch.setattr(bridge, "_process_envelope", _fake_process)

    e1 = MessageEnvelope.from_client_message(
        session_id="s1", message={"type": "start", "requirement": "x"}
    )
    e2 = MessageEnvelope.from_dict(e1.to_dict())

    bridge.publish_inbound(e1)
    bridge.publish_inbound(e2)
    bridge.process_next()
    bridge.process_next()

    assert calls["n"] == 1


def test_stomp_bridge_preserves_message_order(monkeypatch):
    bridge = _bridge(monkeypatch)
    monkeypatch.setattr(
        bridge,
        "_process_envelope",
        lambda env: [{"type": "seq", "n": env.payload["n"], "session_id": env.session_id}],
    )

    e1 = MessageEnvelope.from_client_message(
        session_id="s2", message={"type": "supplement", "n": 1}
    )
    e2 = MessageEnvelope.from_client_message(
        session_id="s2", message={"type": "supplement", "n": 2}
    )
    bridge.publish_inbound(e1)
    bridge.publish_inbound(e2)
    bridge.process_next()
    bridge.process_next()

    out = bridge._drain_outbound("s2", 0.01)
    assert [x["n"] for x in out] == [1, 2]
