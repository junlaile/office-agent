from __future__ import annotations

from office_agent.api.stomp_bridge import MemoryBroker, StompBrokerBridge, StompConfig
from office_agent.api.transport import MessageEnvelope


def _cfg(**overrides) -> StompConfig:
    base = dict(
        host="127.0.0.1",
        port=61613,
        login="guest",
        passcode="guest",
        vhost="/",
        inbound_destination="/queue/office-agent.inbound",
        outbound_destination="/queue/office-agent.outbound",
        dlq_destination="/queue/office-agent.dlq",
        exchange="office-agent",
        routing_key="session",
        heartbeat_ms=10000,
        max_retries=2,
        retry_delay_ms=0,
        use_memory_broker=True,
    )
    base.update(overrides)
    return StompConfig(**base)


def _bridge(**cfg_overrides) -> StompBrokerBridge:
    broker = MemoryBroker()
    return StompBrokerBridge(_cfg(**cfg_overrides), broker=broker)


def test_session_worker_idempotency(monkeypatch):
    from office_agent.api import manager as session_manager
    from office_agent.api.session_worker import process_envelope

    seen = {"n": 0}

    def fake_dispatch(envelope, **kwargs):
        seen["n"] += 1
        return [{"type": "ok", "session_id": envelope.session_id}]

    monkeypatch.setattr("office_agent.api.session_worker._dispatch", fake_dispatch)
    env = MessageEnvelope.from_client_message(
        session_id="idem-1", message={"type": "start", "requirement": "x"}
    )
    assert process_envelope(env)
    assert process_envelope(env) == []
    assert seen["n"] == 1
    session_manager.remove("idem-1")


def test_stomp_bridge_preserves_message_order(monkeypatch):
    bridge = _bridge()
    monkeypatch.setattr(
        "office_agent.api.stomp_bridge.process_envelope",
        lambda env: [
            {"type": "seq", "n": env.payload["n"], "session_id": env.session_id}
        ],
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


def test_stomp_bridge_retry_then_dlq(monkeypatch):
    bridge = _bridge(max_retries=2, retry_delay_ms=0)
    attempts = {"n": 0}

    def boom(env):
        attempts["n"] += 1
        raise RuntimeError("boom")

    monkeypatch.setattr("office_agent.api.stomp_bridge.process_envelope", boom)

    env = MessageEnvelope.from_client_message(
        session_id="s3", message={"type": "start", "requirement": "x"}
    )
    bridge.publish_inbound(env)

    assert bridge.process_next() is True
    assert bridge.process_next() is True
    assert bridge.process_next() is True

    assert isinstance(bridge._broker, MemoryBroker)
    assert len(bridge._broker.dlq) == 1
    out = bridge._drain_outbound("s3", 0.01)
    assert out and out[0]["type"] == "error" and out[0].get("dlq") is True
    assert attempts["n"] == 3
