"""RabbitMQ STOMP 桥接（Gateway + 同进程 Worker）。"""

from __future__ import annotations

import asyncio
import json
import logging
import queue
import threading
from collections import defaultdict
from dataclasses import dataclass
from typing import Any

from office_agent.api import manager as session_manager
from office_agent.api.transport import MessageEnvelope
from office_agent.config import settings
from office_agent.session.runner import AgentSession

logger = logging.getLogger(__name__)


@dataclass
class StompConfig:
    host: str
    port: int
    login: str
    passcode: str
    vhost: str
    inbound_destination: str
    outbound_destination: str
    heartbeat_ms: int


class StompBrokerBridge:
    """用 STOMP broker 转发请求与事件。

    说明：当前实现保留 Worker 于同进程，便于与现有 AgentSession 兼容；
    后续可将 ``process_next`` 迁移至独立 worker 服务。
    """

    def __init__(self, cfg: StompConfig) -> None:
        self._cfg = cfg
        self._lock = threading.Lock()
        self._inbound_q: queue.Queue[MessageEnvelope] = queue.Queue()
        self._outbound_q: dict[str, queue.Queue[dict[str, Any]]] = defaultdict(queue.Queue)
        self._dedupe_ids: dict[str, set[str]] = defaultdict(set)
        self._conn = self._connect()

    def _connect(self):  # type: ignore[no-untyped-def]
        try:
            import stomp
        except ImportError as e:  # pragma: no cover
            raise RuntimeError("TRANSPORT_MODE=rabbitmq_stomp 需要安装 stomp.py") from e
        conn = stomp.Connection12([(self._cfg.host, self._cfg.port)])
        conn.connect(
            login=self._cfg.login,
            passcode=self._cfg.passcode,
            wait=True,
            headers={"host": self._cfg.vhost},
        )
        return conn

    def publish_inbound(self, envelope: MessageEnvelope) -> None:
        self._conn.send(
            destination=self._cfg.inbound_destination,
            body=envelope.to_json(),
            headers={"persistent": "true", "content-type": "application/json"},
        )
        self._inbound_q.put(envelope)

    def process_next(self) -> bool:
        try:
            envelope = self._inbound_q.get_nowait()
        except queue.Empty:
            return False
        sid = envelope.session_id
        if envelope.message_id in self._dedupe_ids[sid]:
            logger.info("忽略重复消息 sid=%s msg_id=%s", sid, envelope.message_id)
            return True
        self._dedupe_ids[sid].add(envelope.message_id)
        events = self._process_envelope(envelope)
        for ev in events:
            self._conn.send(
                destination=self._cfg.outbound_destination,
                body=json.dumps(ev, ensure_ascii=False),
                headers={
                    "persistent": "true",
                    "content-type": "application/json",
                    "session_id": sid,
                },
            )
            self._outbound_q[sid].put(ev)
        return True

    def _process_envelope(self, envelope: MessageEnvelope) -> list[dict[str, Any]]:
        sid = envelope.session_id
        msg_type = envelope.message_type
        payload = envelope.payload
        session = session_manager.get(sid)
        if session is None and msg_type == "start":
            session = AgentSession(session_id=sid)
            session_manager.register(session)
            req = str(payload.get("requirement") or "")
            events = list(session.start(req))
            session_manager.register(session)
            return events
        if session is None:
            return [{"type": "error", "message": "session not found", "session_id": sid}]
        events = list(session.handle({"type": msg_type, **payload}))
        session_manager.register(session)
        return events

    async def consume_outbound(
        self, session_id: str, timeout_s: float = 0.2
    ) -> list[dict[str, Any]]:
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(None, self._drain_outbound, session_id, timeout_s)

    def _drain_outbound(self, session_id: str, timeout_s: float) -> list[dict[str, Any]]:
        q = self._outbound_q[session_id]
        out: list[dict[str, Any]] = []
        try:
            first = q.get(timeout=timeout_s)
        except queue.Empty:
            return out
        out.append(first)
        while True:
            try:
                out.append(q.get_nowait())
            except queue.Empty:
                return out

    def close(self) -> None:
        with self._lock:
            try:
                self._conn.disconnect()
            except Exception:  # noqa: BLE001
                pass


def build_stomp_bridge() -> StompBrokerBridge:
    cfg = StompConfig(
        host=settings.rabbitmq_host,
        port=settings.rabbitmq_port,
        login=settings.rabbitmq_login,
        passcode=settings.rabbitmq_passcode,
        vhost=settings.rabbitmq_vhost,
        inbound_destination=settings.stomp_inbound_destination,
        outbound_destination=settings.stomp_outbound_destination,
        heartbeat_ms=settings.stomp_heartbeat_ms,
    )
    return StompBrokerBridge(cfg)
