"""传输层抽象与消息信封。

把"客户端 ↔ 会话 Worker"之间的消息统一为 MessageEnvelope，
并用 TransportAdapter 协议屏蔽底层传输差异:
    - InProcessWebSocketTransport: 进程内直连 WebSocket（默认 inproc 模式）
    - STOMP/RabbitMQ 路径见 stomp_bridge.py（rabbitmq_stomp 模式）

这样会话处理逻辑（session_worker.process_envelope）不感知消息
是来自 WebSocket 还是消息队列。
"""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any, Protocol

from fastapi import WebSocket


@dataclass
class MessageEnvelope:
    """跨传输层的统一消息信封。

    字段说明:
        session_id:      会话标识，路由到对应 AgentSession
        message_type:    消息类型（start/supplement/approve 等）
        payload:         除 type 外的业务字段
        trace_id:        链路追踪 id（每条消息独立生成）
        timestamp:       消息生成时间（UTC ISO 格式）
        version:         信封结构版本，便于将来演进
        message_id:      消息唯一 id，配合存储层做幂等去重
        session_version: 会话版本号，用于识别过期消息
    """

    session_id: str
    message_type: str
    payload: dict[str, Any]
    trace_id: str
    timestamp: str
    version: int = 1
    message_id: str = ""
    session_version: int = 0

    @classmethod
    def from_client_message(
        cls, *, session_id: str, message: dict[str, Any], session_version: int = 0
    ) -> MessageEnvelope:
        """把客户端原始 JSON 消息包装成信封（自动生成 trace/message id）。"""
        msg_type = str(message.get("type") or "").strip()
        # type 字段提升为 message_type，其余字段原样进 payload
        payload = {k: v for k, v in message.items() if k != "type"}
        return cls(
            session_id=session_id,
            message_type=msg_type,
            payload=payload,
            trace_id=uuid.uuid4().hex,
            timestamp=datetime.now(UTC).isoformat(),
            message_id=uuid.uuid4().hex,
            session_version=session_version,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "session_id": self.session_id,
            "message_type": self.message_type,
            "payload": self.payload,
            "trace_id": self.trace_id,
            "timestamp": self.timestamp,
            "message_id": self.message_id,
            "session_version": self.session_version,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> MessageEnvelope:
        """从字典还原信封;缺失字段取安全默认值（兼容旧版本消息）。"""
        return cls(
            version=int(data.get("version") or 1),
            session_id=str(data.get("session_id") or ""),
            message_type=str(data.get("message_type") or ""),
            payload=dict(data.get("payload") or {}),
            trace_id=str(data.get("trace_id") or uuid.uuid4().hex),
            timestamp=str(data.get("timestamp") or datetime.now(UTC).isoformat()),
            message_id=str(data.get("message_id") or uuid.uuid4().hex),
            session_version=int(data.get("session_version") or 0),
        )

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False)


class TransportAdapter(Protocol):
    """传输适配器协议:收消息、发事件、确认、报错四个动作。"""

    async def receive(self) -> dict[str, Any]: ...
    async def send(self, event: dict[str, Any]) -> None: ...
    async def ack(self, envelope: MessageEnvelope) -> None: ...
    async def error(self, message: str) -> None: ...


class InProcessWebSocketTransport:
    """原生 WebSocket 适配器（inproc 模式:消息不经过消息队列）。"""

    def __init__(self, ws: WebSocket) -> None:
        self._ws = ws

    async def receive(self) -> dict[str, Any]:
        """读取并解析一条客户端 JSON 消息;非法输入抛 ValueError。"""
        raw = await self._ws.receive_text()
        try:
            message = json.loads(raw)
        except json.JSONDecodeError as e:
            raise ValueError("无效 JSON") from e
        if not isinstance(message, dict):
            raise ValueError("消息须为 JSON 对象")
        return message

    async def send(self, event: dict[str, Any]) -> None:
        await self._ws.send_json(event)

    async def ack(self, envelope: MessageEnvelope) -> None:
        # WebSocket 直连无需显式确认（不同于 STOMP 的 ack 语义）
        _ = envelope

    async def error(self, message: str) -> None:
        await self._ws.send_json({"type": "error", "message": message})
