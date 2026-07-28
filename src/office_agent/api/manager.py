"""进程内会话注册表。"""

from __future__ import annotations

import threading
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_agent.session.runner import AgentSession

_lock = threading.Lock()
_sessions: dict[str, AgentSession] = {}


def register(session: AgentSession) -> AgentSession:
    with _lock:
        _sessions[session.session_id] = session
    return session


def get(session_id: str) -> AgentSession | None:
    with _lock:
        return _sessions.get(session_id)


def remove(session_id: str) -> None:
    with _lock:
        _sessions.pop(session_id, None)
