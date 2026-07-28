"""会话存储后端（内存 / MySQL）。"""

from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass
from typing import TYPE_CHECKING, Protocol
from urllib.parse import urlparse

from office_agent.config import settings
from office_agent.session.runner import AgentSession, SessionPhase

if TYPE_CHECKING:
    import pymysql

logger = logging.getLogger(__name__)


class SessionStore(Protocol):
    def register(self, session: AgentSession) -> AgentSession: ...
    def get(self, session_id: str) -> AgentSession | None: ...
    def remove(self, session_id: str) -> None: ...


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


def _to_snapshot(session: AgentSession) -> SessionSnapshot:
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
    return session


class InMemorySessionStore:
    def __init__(self) -> None:
        import threading

        self._lock = threading.Lock()
        self._sessions: dict[str, AgentSession] = {}

    def register(self, session: AgentSession) -> AgentSession:
        with self._lock:
            self._sessions[session.session_id] = session
        return session

    def get(self, session_id: str) -> AgentSession | None:
        with self._lock:
            return self._sessions.get(session_id)

    def remove(self, session_id: str) -> None:
        with self._lock:
            self._sessions.pop(session_id, None)


class MySQLSessionStore:
    TABLE_NAME = "office_agent_sessions"

    def __init__(self, dsn: str) -> None:
        self._dsn = dsn
        self._pymysql = self._import_driver()
        self._ensure_schema()

    def _import_driver(self):  # type: ignore[no-untyped-def]
        try:
            import pymysql
        except ImportError as e:  # pragma: no cover
            raise RuntimeError(
                "SESSION_BACKEND=mysql 需要安装 pymysql。"
            ) from e
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
        sql = f"""
        CREATE TABLE IF NOT EXISTS {self.TABLE_NAME} (
            session_id VARCHAR(64) PRIMARY KEY,
            payload JSON NOT NULL,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                ON UPDATE CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql)

    def register(self, session: AgentSession) -> AgentSession:
        snap = _to_snapshot(session)
        payload = json.dumps(asdict(snap), ensure_ascii=False)
        sql = f"""
        INSERT INTO {self.TABLE_NAME} (session_id, payload)
        VALUES (%s, CAST(%s AS JSON))
        ON DUPLICATE KEY UPDATE payload = VALUES(payload)
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session.session_id, payload))
        return session

    def get(self, session_id: str) -> AgentSession | None:
        sql = f"SELECT payload FROM {self.TABLE_NAME} WHERE session_id=%s"
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session_id,))
                row = cur.fetchone()
        if not row:
            return None
        raw = row[0]
        data = json.loads(raw) if isinstance(raw, str) else raw
        snap = SessionSnapshot(**data)
        return _from_snapshot(snap)

    def remove(self, session_id: str) -> None:
        sql = f"DELETE FROM {self.TABLE_NAME} WHERE session_id=%s"
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session_id,))


def build_session_store() -> SessionStore:
    backend = settings.session_backend.strip().lower()
    if backend == "mysql":
        try:
            return MySQLSessionStore(settings.mysql_dsn)
        except Exception:  # noqa: BLE001
            logger.exception("MySQL 会话存储初始化失败，回退内存存储")
    return InMemorySessionStore()
