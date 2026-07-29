"""会话 Worker：消费入站信封并产出事件。"""

from __future__ import annotations

import logging
from typing import Any

from office_agent.api import manager as session_manager
from office_agent.api.transport import MessageEnvelope
from office_agent.session.runner import AgentSession, SessionPhase

logger = logging.getLogger(__name__)


def process_envelope(envelope: MessageEnvelope) -> list[dict[str, Any]]:
    """处理一条入站消息，返回应推给客户端的事件列表。"""
    sid = envelope.session_id
    msg_type = envelope.message_type
    payload = dict(envelope.payload)

    if session_manager.is_duplicate_message(
        session_id=sid,
        message_id=envelope.message_id,
        session_version=envelope.session_version,
    ):
        logger.info(
            "幂等跳过 sid=%s msg_id=%s version=%s",
            sid,
            envelope.message_id,
            envelope.session_version,
        )
        return []

    try:
        events = _dispatch(envelope, sid=sid, msg_type=msg_type, payload=payload)
    except Exception:
        # 失败不标记，允许重试
        raise

    session_manager.mark_message_processed(
        session_id=sid,
        message_id=envelope.message_id,
        session_version=envelope.session_version,
    )
    return events


def _dispatch(
    envelope: MessageEnvelope,
    *,
    sid: str,
    msg_type: str,
    payload: dict[str, Any],
) -> list[dict[str, Any]]:
    session = session_manager.get(sid)
    if session is None and msg_type == "start":
        session = AgentSession(session_id=sid)
        session_manager.register(session)
        req = str(payload.get("requirement") or "")
        events = list(session.start(req))
        session_manager.register(session)
        return events

    if session is None:
        return [
            {
                "type": "error",
                "message": "session not found",
                "session_id": sid,
            }
        ]

    if msg_type == "start":
        if session.phase != SessionPhase.CREATED:
            return [
                {
                    "type": "session",
                    "session_id": session.session_id,
                    "phase": str(session.phase),
                    "reconnected": True,
                    "doc_type": session.doc_type,
                    "kind": session.kind,
                    "doc_path": session.doc_path,
                    "summary": session.summary,
                    "error": session.error,
                }
            ]
        req = str(payload.get("requirement") or "")
        events = list(session.start(req))
        session_manager.register(session)
        return events

    payload.pop("session_id", None)
    events = list(session.handle({"type": msg_type, **payload}))
    session_manager.register(session)
    return events
