"""会话存储后端（内存 / MySQL），含 TTL 与消息幂等。"""

from __future__ import annotations

import json
import threading
import time
from dataclasses import asdict, dataclass
from typing import Protocol
from urllib.parse import urlparse

from office_agent.config import settings
from office_agent.log import get_logger
from office_agent.session.runner import AgentSession, SessionPhase

logger = get_logger(__name__)


class SessionStore(Protocol):
    def register(self, session: AgentSession) -> AgentSession: ...
    def get(self, session_id: str) -> AgentSession | None: ...
    def remove(self, session_id: str) -> None: ...

    def is_duplicate_message(
        self, *, session_id: str, message_id: str, session_version: int
    ) -> bool: ...

    def mark_message_processed(
        self, *, session_id: str, message_id: str, session_version: int
    ) -> None: ...


@dataclass
class SessionSnapshot:
    session_id: str
    phase: str
    requirement: str
    doc_type: str | None
    doc_path: str | None
    kind: str | None
    approved_outline: str
    template_text: str
    summary: str | None
    error: str | None
    session_version: int = 0
    expires_at: float | None = None


def _to_snapshot(session: AgentSession) -> SessionSnapshot:
    ttl = max(0, int(settings.session_ttl_seconds))
    expires_at = (time.time() + ttl) if ttl else None
    version = int(getattr(session, "session_version", 0) or 0)
    return SessionSnapshot(
        session_id=session.session_id,
        phase=str(session.phase),
        requirement=session.requirement,
        doc_type=session.doc_type,
        doc_path=session.doc_path,
        kind=session.kind,
        approved_outline=session.approved_outline,
        template_text=session.template_text,
        summary=session.summary,
        error=session.error,
        session_version=version,
        expires_at=expires_at,
    )


def _from_snapshot(snapshot: SessionSnapshot) -> AgentSession:
    session = AgentSession(session_id=snapshot.session_id)
    try:
        session.phase = SessionPhase(snapshot.phase)
    except ValueError:
        session.phase = SessionPhase.ERROR
        session.error = f"无效会话 phase: {snapshot.phase}"
    session.requirement = snapshot.requirement
    session.doc_type = snapshot.doc_type
    session.doc_path = snapshot.doc_path
    session.kind = snapshot.kind
    session.approved_outline = snapshot.approved_outline
    session.template_text = snapshot.template_text
    session.summary = snapshot.summary
    if snapshot.error:
        session.error = snapshot.error
    session.session_version = snapshot.session_version
    return session


class InMemorySessionStore:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._sessions: dict[str, tuple[AgentSession, float | None]] = {}
        self._processed: set[tuple[str, str]] = set()

    def register(self, session: AgentSession) -> AgentSession:
        snap = _to_snapshot(session)
        with self._lock:
            self._sessions[session.session_id] = (session, snap.expires_at)
        return session

    def get(self, session_id: str) -> AgentSession | None:
        with self._lock:
            item = self._sessions.get(session_id)
            if item is None:
                return None
            session, expires_at = item
            if expires_at is not None and time.time() > expires_at:
                self._sessions.pop(session_id, None)
                return None
            return session

    def remove(self, session_id: str) -> None:
        with self._lock:
            self._sessions.pop(session_id, None)

    def is_duplicate_message(
        self, *, session_id: str, message_id: str, session_version: int
    ) -> bool:
        _ = session_version
        with self._lock:
            return (session_id, message_id) in self._processed

    def mark_message_processed(
        self, *, session_id: str, message_id: str, session_version: int
    ) -> None:
        _ = session_version
        with self._lock:
            self._processed.add((session_id, message_id))


class MySQLSessionStore:
    TABLE_NAME = "office_agent_sessions"
    MSG_TABLE = "office_agent_processed_messages"

    def __init__(self, dsn: str) -> None:
        self._dsn = dsn
        self._pymysql = self._import_driver()
        self._ensure_schema()

    def _import_driver(self):  # type: ignore[no-untyped-def]
        try:
            import pymysql
        except ImportError as e:  # pragma: no cover
            raise RuntimeError("SESSION_BACKEND=mysql 需要安装 pymysql。") from e
        return pymysql

    def _connect(self):  # type: ignore[no-untyped-def]
        cfg = {
            "host": settings.mysql_host,
            "port": settings.mysql_port,
            "user": settings.mysql_user,
            "password": settings.mysql_password,
            "database": settings.mysql_database,
        }
        if self._dsn:
            parsed = urlparse(self._dsn)
            if parsed.hostname:
                cfg["host"] = parsed.hostname
            if parsed.port:
                cfg["port"] = parsed.port
            if parsed.username:
                cfg["user"] = parsed.username
            if parsed.password:
                cfg["password"] = parsed.password
            if parsed.path and parsed.path != "/":
                cfg["database"] = parsed.path.lstrip("/")

        return self._pymysql.connect(
            host=cfg["host"],
            port=cfg["port"],
            user=cfg["user"],
            password=cfg["password"],
            database=cfg["database"],
            charset="utf8mb4",
            autocommit=True,
            connect_timeout=settings.mysql_connect_timeout,
            read_timeout=settings.mysql_read_timeout,
            write_timeout=settings.mysql_write_timeout,
        )

    def _ensure_schema(self) -> None:
        session_sql = f"""
        CREATE TABLE IF NOT EXISTS {self.TABLE_NAME} (
            session_id VARCHAR(64) PRIMARY KEY,
            payload JSON NOT NULL,
            expires_at TIMESTAMP NULL,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                ON UPDATE CURRENT_TIMESTAMP,
            INDEX idx_expires_at (expires_at)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        msg_sql = f"""
        CREATE TABLE IF NOT EXISTS {self.MSG_TABLE} (
            session_id VARCHAR(64) NOT NULL,
            message_id VARCHAR(64) NOT NULL,
            session_version INT NOT NULL DEFAULT 0,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (session_id, message_id)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(session_sql)
                cur.execute(msg_sql)

    def register(self, session: AgentSession) -> AgentSession:
        snap = _to_snapshot(session)
        payload = json.dumps(asdict(snap), ensure_ascii=False)
        expires_at = (
            time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(snap.expires_at))
            if snap.expires_at
            else None
        )
        sql = f"""
        INSERT INTO {self.TABLE_NAME} (session_id, payload, expires_at)
        VALUES (%s, CAST(%s AS JSON), %s)
        ON DUPLICATE KEY UPDATE
            payload = VALUES(payload),
            expires_at = VALUES(expires_at)
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session.session_id, payload, expires_at))
        return session

    def get(self, session_id: str) -> AgentSession | None:
        sql = f"""
        SELECT payload, expires_at FROM {self.TABLE_NAME}
        WHERE session_id=%s
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session_id,))
                row = cur.fetchone()
        if not row:
            return None
        raw, expires_at = row[0], row[1]
        if expires_at is not None:
            # pymysql may return datetime
            ts = expires_at.timestamp() if hasattr(expires_at, "timestamp") else None
            if ts is not None and time.time() > ts:
                self.remove(session_id)
                return None
        data = json.loads(raw) if isinstance(raw, str) else raw
        snap = SessionSnapshot(**data)
        return _from_snapshot(snap)

    def remove(self, session_id: str) -> None:
        sql = f"DELETE FROM {self.TABLE_NAME} WHERE session_id=%s"
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session_id,))

    def is_duplicate_message(
        self, *, session_id: str, message_id: str, session_version: int
    ) -> bool:
        _ = session_version
        sql = f"""
        SELECT 1 FROM {self.MSG_TABLE}
        WHERE session_id=%s AND message_id=%s
        LIMIT 1
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session_id, message_id))
                return cur.fetchone() is not None

    def mark_message_processed(
        self, *, session_id: str, message_id: str, session_version: int
    ) -> None:
        sql = f"""
        INSERT IGNORE INTO {self.MSG_TABLE}
            (session_id, message_id, session_version)
        VALUES (%s, %s, %s)
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session_id, message_id, session_version))


def build_session_store() -> SessionStore:
    backend = settings.session_backend.strip().lower()
    if backend == "mysql":
        try:
            return MySQLSessionStore(settings.mysql_dsn)
        except Exception:  # noqa: BLE001
            logger.exception("MySQL 会话存储初始化失败，回退内存存储")
    return InMemorySessionStore()
