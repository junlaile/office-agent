from __future__ import annotations

import dataclasses
import time

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


def test_in_memory_store_ttl_expiry(monkeypatch):
    from office_agent import config
    from office_agent.api import session_store as ss

    new_settings = dataclasses.replace(config.settings, session_ttl_seconds=1)
    monkeypatch.setattr(ss, "settings", new_settings)

    store = InMemorySessionStore()
    session = AgentSession(session_id="sid-ttl")
    store.register(session)
    assert store.get("sid-ttl") is session

    with store._lock:
        store._sessions["sid-ttl"] = (session, time.time() - 1)
    assert store.get("sid-ttl") is None


def test_in_memory_idempotency():
    store = InMemorySessionStore()
    assert (
        store.is_duplicate_message(
            session_id="s", message_id="m1", session_version=1
        )
        is False
    )
    store.mark_message_processed(
        session_id="s", message_id="m1", session_version=1
    )
    assert (
        store.is_duplicate_message(
            session_id="s", message_id="m1", session_version=1
        )
        is True
    )
