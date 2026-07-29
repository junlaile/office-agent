"""RabbitMQ STOMP 桥接（Gateway + Worker，含重试与 DLQ）。

TRANSPORT_MODE=rabbitmq_stomp 时的消息路径:

    WebSocket 客户端
        │ (Gateway: app.py 收到消息后 publish_inbound)
        ▼
    inbound 队列 ──▶ SessionWorker（process_next / run_forever）
        │                 │ 处理失败:按 STOMP_MAX_RETRIES 重试，
        │                 │ 超限后写入 DLQ 并回推错误事件
        ▼                 ▼
    outbound 队列 ◀── process_envelope 产出的事件
        │ (Gateway: consume_outbound 拉取后经 WebSocket 推给客户端)
        ▼
    WebSocket 客户端

两种 broker 实现:
    - StompBrokerClient: 真实 RabbitMQ STOMP 连接（需要 stomp.py）
    - MemoryBroker:      进程内队列模拟，供单测与无 RabbitMQ 联调
"""

from __future__ import annotations

import asyncio
import json
import queue
import threading
import time
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Protocol

from office_agent.api.session_worker import process_envelope
from office_agent.api.transport import MessageEnvelope
from office_agent.config import settings
from office_agent.log import get_logger

logger = get_logger(__name__)


@dataclass
class StompConfig:
    """STOMP 连接与队列配置（由 settings 的 RABBITMQ_*/STOMP_* 组装）。"""

    host: str
    port: int
    login: str
    passcode: str
    vhost: str
    inbound_destination: str
    outbound_destination: str
    dlq_destination: str
    exchange: str
    routing_key: str
    heartbeat_ms: int
    max_retries: int
    retry_delay_ms: int
    use_memory_broker: bool = False


class BrokerClient(Protocol):
    """broker 最小抽象:只需要"发消息"和"断开"两个能力。"""

    def send(self, destination: str, body: str, headers: dict[str, str]) -> None: ...
    def disconnect(self) -> None: ...


class MemoryBroker:
    """进程内 broker，便于单测与无 RabbitMQ 时联调。"""

    def __init__(self) -> None:
        self.inbound: queue.Queue[tuple[str, dict[str, str]]] = queue.Queue()
        self.outbound: dict[str, queue.Queue[dict[str, Any]]] = defaultdict(queue.Queue)
        self.dlq: list[tuple[str, dict[str, str]]] = []
        self.sent: list[tuple[str, str, dict[str, str]]] = []

    def send(self, destination: str, body: str, headers: dict[str, str]) -> None:
        """按目的地名称路由到对应内存队列（dlq / outbound / inbound）。"""
        self.sent.append((destination, body, headers))
        dest = destination.lower()
        if dest.endswith(".dlq") or "/dlq" in dest or dest.endswith("dlq"):
            self.dlq.append((body, headers))
            return
        if "outbound" in dest:
            sid = headers.get("session_id") or ""
            try:
                event = json.loads(body)
            except json.JSONDecodeError:
                event = {"type": "error", "message": "invalid outbound body"}
            self.outbound[sid].put(event)
            return
        self.inbound.put((body, headers))

    def disconnect(self) -> None:
        return None


class StompBrokerClient:
    """真实 RabbitMQ STOMP 客户端（薄封装 stomp.py 的 Connection12）。"""

    def __init__(self, cfg: StompConfig) -> None:
        # stomp.py 是可选依赖:只有启用 rabbitmq_stomp 模式时才要求安装
        try:
            import stomp
        except ImportError as e:  # pragma: no cover
            raise RuntimeError("TRANSPORT_MODE=rabbitmq_stomp 需要安装 stomp.py") from e
        self._conn = stomp.Connection12([(cfg.host, cfg.port)])
        self._conn.connect(
            login=cfg.login,
            passcode=cfg.passcode,
            wait=True,
            headers={"host": cfg.vhost},
        )

    def send(self, destination: str, body: str, headers: dict[str, str]) -> None:
        self._conn.send(destination=destination, body=body, headers=headers)

    def disconnect(self) -> None:
        try:
            self._conn.disconnect()
        except Exception:  # noqa: BLE001
            pass


@dataclass
class PendingMessage:
    """待消费的入站消息:信封 + 已重试次数 + 原始 STOMP 头。"""

    envelope: MessageEnvelope
    attempt: int = 0
    headers: dict[str, str] = field(default_factory=dict)


class StompBrokerBridge:
    """Gateway 发布入站；Worker 消费处理并回写出站；失败重试后进 DLQ。"""

    def __init__(self, cfg: StompConfig, broker: BrokerClient | None = None) -> None:
        self._cfg = cfg
        self._lock = threading.Lock()
        # 真实 STOMP 模式下同进程 embed worker 的本地待消费队列
        self._pending: queue.Queue[PendingMessage] = queue.Queue()
        # 真实 STOMP 模式下按 session 分桶的本地出站队列（供 Gateway 拉取）
        self._outbound_local: dict[str, queue.Queue[dict[str, Any]]] = defaultdict(
            queue.Queue
        )
        self._broker = broker or self._build_broker()
        self._memory = isinstance(self._broker, MemoryBroker)

    def _build_broker(self) -> BrokerClient:
        if self._cfg.use_memory_broker:
            return MemoryBroker()
        return StompBrokerClient(self._cfg)

    def publish_inbound(self, envelope: MessageEnvelope) -> None:
        """Gateway 侧:把客户端消息发布到 inbound 队列（持久化投递）。"""
        headers = {
            "persistent": "true",
            "content-type": "application/json",
            "session_id": envelope.session_id,
            "message_id": envelope.message_id,
            "session_version": str(envelope.session_version),
            "retry_count": "0",
        }
        if self._cfg.exchange:
            headers["exchange"] = self._cfg.exchange
        if self._cfg.routing_key:
            headers["routing_key"] = self._cfg.routing_key
        self._broker.send(
            destination=self._cfg.inbound_destination,
            body=envelope.to_json(),
            headers=headers,
        )
        # 真实 STOMP：同进程 embed worker 用本地 pending；memory broker 从 inbound 拉
        if not self._memory:
            self._pending.put(
                PendingMessage(envelope=envelope, attempt=0, headers=headers)
            )

    def process_next(self) -> bool:
        """处理一条待消费消息。返回是否处理了消息。"""
        pending = self._take_pending()
        if pending is None:
            return False
        return self._handle_pending(pending)

    def run_forever(self, *, poll_interval_s: float = 0.05) -> None:
        """独立 worker 循环。"""
        logger.info(
            "SessionWorker 启动 inbound=%s outbound=%s dlq=%s",
            self._cfg.inbound_destination,
            self._cfg.outbound_destination,
            self._cfg.dlq_destination,
        )
        while True:
            handled = self.process_next()
            if not handled:
                time.sleep(poll_interval_s)

    def _take_pending(self) -> PendingMessage | None:
        """取一条待处理消息;memory broker 从 inbound 队列拉，否则用本地 pending。"""
        if self._memory and isinstance(self._broker, MemoryBroker):
            try:
                body, headers = self._broker.inbound.get_nowait()
            except queue.Empty:
                return None
            envelope = MessageEnvelope.from_dict(json.loads(body))
            attempt = int(headers.get("retry_count") or 0)
            return PendingMessage(envelope=envelope, attempt=attempt, headers=headers)
        try:
            return self._pending.get_nowait()
        except queue.Empty:
            return None

    def _handle_pending(self, pending: PendingMessage) -> bool:
        """处理一条消息:成功则把事件写出站队列，失败走重试/DLQ。"""
        envelope = pending.envelope
        sid = envelope.session_id
        try:
            events = process_envelope(envelope)
        except Exception as e:  # noqa: BLE001
            logger.exception(
                "处理失败 sid=%s msg_id=%s attempt=%s",
                sid,
                envelope.message_id,
                pending.attempt,
            )
            return self._retry_or_dlq(pending, error=str(e))

        for ev in events:
            headers = {
                "persistent": "true",
                "content-type": "application/json",
                "session_id": sid,
                "message_id": envelope.message_id,
            }
            body = json.dumps(ev, ensure_ascii=False)
            self._broker.send(
                destination=self._cfg.outbound_destination,
                body=body,
                headers=headers,
            )
            if not self._memory:
                self._outbound_local[sid].put(ev)
        return True

    def _retry_or_dlq(self, pending: PendingMessage, *, error: str) -> bool:
        """失败处理:未超过重试上限则延迟后重新入队，否则写 DLQ 并回推错误事件。"""
        next_attempt = pending.attempt + 1
        if next_attempt <= self._cfg.max_retries:
            delay = self._cfg.retry_delay_ms / 1000.0
            logger.warning(
                "消息重试 sid=%s msg_id=%s attempt=%s delay=%.3fs",
                pending.envelope.session_id,
                pending.envelope.message_id,
                next_attempt,
                delay,
            )
            if delay > 0:
                time.sleep(delay)
            headers = dict(pending.headers)
            headers["retry_count"] = str(next_attempt)
            headers["last_error"] = error[:500]
            self._broker.send(
                destination=self._cfg.inbound_destination,
                body=pending.envelope.to_json(),
                headers=headers,
            )
            if not self._memory:
                self._pending.put(
                    PendingMessage(
                        envelope=pending.envelope,
                        attempt=next_attempt,
                        headers=headers,
                    )
                )
            return True

        logger.error(
            "消息进入 DLQ sid=%s msg_id=%s error=%s",
            pending.envelope.session_id,
            pending.envelope.message_id,
            error,
        )
        dlq_body = json.dumps(
            {
                "envelope": pending.envelope.to_dict(),
                "error": error,
                "attempts": pending.attempt,
            },
            ensure_ascii=False,
        )
        self._broker.send(
            destination=self._cfg.dlq_destination,
            body=dlq_body,
            headers={
                "persistent": "true",
                "content-type": "application/json",
                "session_id": pending.envelope.session_id,
                "message_id": pending.envelope.message_id,
            },
        )
        fail_ev = {
            "type": "error",
            "message": f"消息处理失败并进入 DLQ: {error}",
            "session_id": pending.envelope.session_id,
            "dlq": True,
        }
        if self._memory and isinstance(self._broker, MemoryBroker):
            self._broker.outbound[pending.envelope.session_id].put(fail_ev)
        else:
            self._outbound_local[pending.envelope.session_id].put(fail_ev)
        return True

    async def consume_outbound(
        self, session_id: str, timeout_s: float = 0.5
    ) -> list[dict[str, Any]]:
        """Gateway 侧:拉取指定会话的出站事件（阻塞取放到线程池以免卡事件循环）。"""
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(
            None, self._drain_outbound, session_id, timeout_s
        )

    def _drain_outbound(self, session_id: str, timeout_s: float) -> list[dict[str, Any]]:
        """排空会话的出站队列:第一条最多等 timeout_s，其余立即取完。"""
        if self._memory and isinstance(self._broker, MemoryBroker):
            q = self._broker.outbound[session_id]
        else:
            q = self._outbound_local[session_id]
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
                self._broker.disconnect()
            except Exception:  # noqa: BLE001
                pass


def build_stomp_config() -> StompConfig:
    """从全局 settings 组装 STOMP 配置。"""
    return StompConfig(
        host=settings.rabbitmq_host,
        port=settings.rabbitmq_port,
        login=settings.rabbitmq_login,
        passcode=settings.rabbitmq_passcode,
        vhost=settings.rabbitmq_vhost,
        inbound_destination=settings.stomp_inbound_destination,
        outbound_destination=settings.stomp_outbound_destination,
        dlq_destination=settings.stomp_dlq_destination,
        exchange=settings.stomp_exchange,
        routing_key=settings.stomp_routing_key,
        heartbeat_ms=settings.stomp_heartbeat_ms,
        max_retries=settings.stomp_max_retries,
        retry_delay_ms=settings.stomp_retry_delay_ms,
        use_memory_broker=settings.stomp_use_memory_broker,
    )


def build_stomp_bridge() -> StompBrokerBridge:
    """构建桥接实例（app.py 的 rabbitmq_stomp 模式与独立 worker 共用）。"""
    return StompBrokerBridge(build_stomp_config())
