"""传输层抽象与消息信封。"""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any, Protocol

from fastapi import WebSocket


@dataclass
class MessageEnvelope:
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
    ) -> "MessageEnvelope":
        msg_type = str(message.get("type") or "").strip()
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
    def from_dict(cls, data: dict[str, Any]) -> "MessageEnvelope":
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
    async def receive(self) -> dict[str, Any]: ...
    async def send(self, event: dict[str, Any]) -> None: ...
    async def ack(self, envelope: MessageEnvelope) -> None: ...
    async def error(self, message: str) -> None: ...


class InProcessWebSocketTransport:
    """原生 WebSocket 适配器。"""

    def __init__(self, ws: WebSocket) -> None:
        self._ws = ws

    async def receive(self) -> dict[str, Any]:
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
        _ = envelope

    async def error(self, message: str) -> None:
        await self._ws.send_json({"type": "error", "message": message})
