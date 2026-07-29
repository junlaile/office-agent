"""会话存储后端（内存 / MySQL），含 TTL 与消息幂等。

提供两个实现:
    - InMemorySessionStore: 进程内字典存储，零依赖，适合单机开发/测试。
    - MySQLSessionStore:   会话快照持久化到 MySQL，进程重启/多 Worker 场景
      可恢复会话；同时用独立表记录已处理消息实现"至少一次投递"下的幂等。

设计取舍:
    - 只持久化可序列化的"快照"字段（SessionSnapshot），LangGraph 的
      运行时状态（内存 checkpointer）不落库——跨进程恢复的会话会丢失
      图执行进度，只能基于快照字段继续对话。
    - MySQL 后端每次操作都新建连接（autocommit），牺牲一点性能换实现
      简单与无连接池状态；量大时可再引入连接池。
"""

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
    """会话存储协议:注册/查询/删除 + 消息幂等去重。"""

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
    """AgentSession 的可序列化快照（仅业务字段，不含 LangGraph 运行时状态）。"""

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
    """把运行中的会话转成快照；TTL=0 表示永不过期。"""
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
    """从快照重建会话对象（跨进程恢复时使用）。"""
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
    """进程内会话存储:dict + 锁，get 时惰性清理过期会话。"""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        # session_id -> (会话对象, 过期时间戳或 None=永不过期)
        self._sessions: dict[str, tuple[AgentSession, float | None]] = {}
        # 已处理消息集合 (session_id, message_id)，用于幂等去重
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
            # 惰性过期:读到已过期的会话时顺手删除
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
    """MySQL 会话存储:快照存 JSON 列，消息幂等记录存独立表。

    初始化时自动建表（CREATE TABLE IF NOT EXISTS），无需手工迁移。
    """

    TABLE_NAME = "office_agent_sessions"
    MSG_TABLE = "office_agent_processed_messages"

    def __init__(self, dsn: str) -> None:
        self._dsn = dsn
        self._pymysql = self._import_driver()
        self._ensure_schema()

    def _import_driver(self):  # type: ignore[no-untyped-def]
        # pymysql 是可选依赖:只有选用 mysql 后端时才要求安装
        try:
            import pymysql
        except ImportError as e:  # pragma: no cover
            raise RuntimeError("SESSION_BACKEND=mysql 需要安装 pymysql。") from e
        return pymysql

    def _connect(self):  # type: ignore[no-untyped-def]
        """建立 MySQL 连接;MYSQL_DSN 中的字段优先于 MYSQL_HOST 等零散配置。"""
        cfg = {
            "host": settings.mysql_host,
            "port": settings.mysql_port,
            "user": settings.mysql_user,
            "password": settings.mysql_password,
            "database": settings.mysql_database,
        }
        if self._dsn:
            # DSN 形如 mysql://user:pass@host:3306/dbname，逐字段覆盖默认配置
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
        """幂等建表:会话快照表 + 已处理消息表。"""
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
        """插入或更新会话快照（UPSERT），同时刷新过期时间。"""
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
            # pymysql 可能把 TIMESTAMP 列返回为 datetime 对象
            ts = expires_at.timestamp() if hasattr(expires_at, "timestamp") else None
            # 惰性过期:读到已过期的会话时顺手删除
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
        # INSERT IGNORE:并发重复标记时静默跳过（主键去重）
        sql = f"""
        INSERT IGNORE INTO {self.MSG_TABLE}
            (session_id, message_id, session_version)
        VALUES (%s, %s, %s)
        """
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (session_id, message_id, session_version))


def build_session_store() -> SessionStore:
    """按 SESSION_BACKEND 构建存储后端;MySQL 初始化失败时降级到内存存储。"""
    backend = settings.session_backend.strip().lower()
    if backend == "mysql":
        try:
            return MySQLSessionStore(settings.mysql_dsn)
        except Exception:  # noqa: BLE001
            # 连接失败/驱动缺失等都不阻塞启动，降级为内存存储并记录日志
            logger.exception("MySQL 会话存储初始化失败，回退内存存储")
    return InMemorySessionStore()
