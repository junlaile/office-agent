"""会话注册表（支持内存 / MySQL 后端）。"""

from __future__ import annotations

import threading
from typing import TYPE_CHECKING

from office_agent.api.session_store import SessionStore, build_session_store

if TYPE_CHECKING:
    from office_agent.session.runner import AgentSession

_lock = threading.Lock()
_sessions: dict[str, AgentSession] = {}
_store: SessionStore = build_session_store()


def register(session: AgentSession) -> AgentSession:
    with _lock:
        _sessions[session.session_id] = session
    _store.register(session)
    return session


def get(session_id: str) -> AgentSession | None:
    with _lock:
        session = _sessions.get(session_id)
    if session is not None:
        return session
    session = _store.get(session_id)
    if session is not None:
        with _lock:
            _sessions[session.session_id] = session
    return session


def remove(session_id: str) -> None:
    with _lock:
        _sessions.pop(session_id, None)
    _store.remove(session_id)


def is_duplicate_message(
    *, session_id: str, message_id: str, session_version: int
) -> bool:
    return _store.is_duplicate_message(
        session_id=session_id,
        message_id=message_id,
        session_version=session_version,
    )


def mark_message_processed(
    *, session_id: str, message_id: str, session_version: int
) -> None:
    _store.mark_message_processed(
        session_id=session_id,
        message_id=message_id,
        session_version=session_version,
    )
