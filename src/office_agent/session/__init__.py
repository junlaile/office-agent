"""可复用的会话编排（CLI / Web API 共用）。"""

from office_agent.session.prep import (
    build_doc_path,
    merge_official_doc,
    official_header_fields,
    resolve_output_kind,
)

__all__ = [
    "AgentSession",
    "SessionPhase",
    "build_doc_path",
    "merge_official_doc",
    "official_header_fields",
    "resolve_output_kind",
]


def __getattr__(name: str):
    # 延迟导入，避免 session.prep ← cli.ui ← session.runner 循环
    if name in ("AgentSession", "SessionPhase"):
        from office_agent.session.runner import AgentSession, SessionPhase

        return {"AgentSession": AgentSession, "SessionPhase": SessionPhase}[name]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
