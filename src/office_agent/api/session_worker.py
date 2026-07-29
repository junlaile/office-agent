"""会话 Worker：消费入站信封并产出事件。

传输层无关的会话消息处理核心:WebSocket 直连（app.py）和
STOMP/RabbitMQ 桥（stomp_bridge.py）最终都调用 process_envelope。

处理流程:
    1. 幂等检查:message_id 已处理过则直接跳过（消息队列可能重复投递）
    2. 分发:start 消息创建/恢复会话，其余消息交给会话状态机 handle
    3. 成功后标记消息已处理并持久化会话快照
"""

from __future__ import annotations

from typing import Any

from office_agent.api import manager as session_manager
from office_agent.api.transport import MessageEnvelope
from office_agent.log import get_logger
from office_agent.session.runner import AgentSession, SessionPhase

logger = get_logger(__name__)


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
        # 处理失败不标记已处理，让消息队列重试/进 DLQ
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
    """按消息类型分发到会话状态机，返回待推送事件。"""
    session = session_manager.get(sid)
    # 新会话:start 消息触发创建并立即启动
    if session is None and msg_type == "start":
        session = AgentSession(session_id=sid)
        session_manager.register(session)
        req = str(payload.get("requirement") or "")
        events = list(session.start(req))
        # start 会推进会话状态，再注册一次以持久化最新快照
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
        # 已有会话再次收到 start:视为断线重连，返回当前状态而非重新启动
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

    # 其余消息类型（supplement/approve/finish 等）交给会话状态机处理
    payload.pop("session_id", None)
    events = list(session.handle({"type": msg_type, **payload}))
    session_manager.register(session)
    return events
