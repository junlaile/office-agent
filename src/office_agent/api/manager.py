"""会话注册表（支持内存 / MySQL 后端）。

模块级单例，进程内所有请求共享:
    - _sessions: 进程内一级缓存，保存"活的"AgentSession 对象
      （含 LangGraph 运行时状态），命中时无需反序列化。
    - _store:    可插拔的二级存储（内存/MySQL），负责持久化快照与
      消息幂等记录;跨进程恢复的会话由快照重建。
"""

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
    """登记会话:写入进程内缓存并持久化快照。"""
    with _lock:
        _sessions[session.session_id] = session
    _store.register(session)
    return session


def get(session_id: str) -> AgentSession | None:
    """查会话:优先进程内缓存，未命中再从存储恢复并回填缓存。"""
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
    """删除会话:缓存与持久化存储同时清理。"""
    with _lock:
        _sessions.pop(session_id, None)
    _store.remove(session_id)


def is_duplicate_message(
    *, session_id: str, message_id: str, session_version: int
) -> bool:
    """判断消息是否已处理过（幂等去重，委托给存储层）。"""
    return _store.is_duplicate_message(
        session_id=session_id,
        message_id=message_id,
        session_version=session_version,
    )


def mark_message_processed(
    *, session_id: str, message_id: str, session_version: int
) -> None:
    """标记消息已处理（幂等去重，委托给存储层）。"""
    _store.mark_message_processed(
        session_id=session_id,
        message_id=message_id,
        session_version=session_version,
    )
