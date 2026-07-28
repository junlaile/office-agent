from __future__ import annotations

from office_agent.api.session_store import InMemorySessionStore
from office_agent.session.runner import AgentSession, SessionPhase


def test_in_memory_store_roundtrip():
    store = InMemorySessionStore()
    session = AgentSession(session_id="sid-1")
    session.phase = SessionPhase.RUNNING
    session.requirement = "写周报"

    store.register(session)
    loaded = store.get("sid-1")

    assert loaded is session
    assert loaded is not None
    assert loaded.phase == SessionPhase.RUNNING
    assert loaded.requirement == "写周报"


def test_in_memory_store_remove():
    store = InMemorySessionStore()
    session = AgentSession(session_id="sid-2")
    store.register(session)

    store.remove("sid-2")

    assert store.get("sid-2") is None
