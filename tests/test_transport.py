from __future__ import annotations

from office_agent.api.transport import MessageEnvelope


def test_message_envelope_shape():
    env = MessageEnvelope.from_client_message(
        session_id="sid-1",
        message={"type": "supplement", "text": "hello"},
        session_version=3,
    )
    data = env.to_dict()
    assert data["session_id"] == "sid-1"
    assert data["message_type"] == "supplement"
    assert data["payload"]["text"] == "hello"
    assert data["session_version"] == 3
    assert data["version"] == 1
    assert data["trace_id"]
    assert data["message_id"]
