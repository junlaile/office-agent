"""Agent 会话状态机：完整对话生命周期（无终端 I/O）。

事件以 dict 形式 yield，供 WebSocket / 其它前端适配层序列化。
忙时输入复用 ``UserInputBridge``（不启动 stdin 读线程，仅用 submit/consume）。
"""

from __future__ import annotations

import threading
import uuid
from collections.abc import Iterator
from enum import StrEnum
from pathlib import Path
from typing import Any

from langchain_core.messages import AIMessage, HumanMessage, ToolMessage
from langgraph.errors import GraphRecursionError
from langgraph.types import Command

from office_agent.cli.ui import _format_tool_call
from office_agent.cli.user_input import (
    CONTINUE_PROMPT,
    PREFIX_FORCE,
    PREFIX_SUPPLEMENT,
    UserInputBridge,
    set_bridge,
)
from office_agent.config import settings
from office_agent.domain.templates import detect_doc_type
from office_agent.domain.vehicle_data import is_vehicle_related
from office_agent.log import get_logger
from office_agent.session.interrupt_util import pending_interrupt
from office_agent.session.prep import (
    build_doc_path,
    merge_official_doc,
    official_header_fields,
    resolve_output_kind,
)
from office_agent.tools import set_session_doc

logger = get_logger(__name__)

# 工具层 ``set_session_doc`` 有进程级回退；同进程多会话时串行执行 graph 步骤。
_EXEC_LOCK = threading.Lock()


class SessionPhase(StrEnum):
    CREATED = "created"
    AWAITING_KIND = "awaiting_kind"
    AWAITING_OUTLINE = "awaiting_outline"
    AWAITING_OFFICIAL_HEADER = "awaiting_official_header"
    RUNNING = "running"
    AWAITING_INTERRUPT = "awaiting_interrupt"
    DONE = "done"
    CANCELLED = "cancelled"
    ERROR = "error"


def _tool_call_dicts(message: AIMessage) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for tc in message.tool_calls or []:
        name = tc.get("name", "")
        args = tc.get("args") or {}
        out.append(
            {
                "name": name,
                "args": args,
                "display": _format_tool_call(name, args),
            }
        )
    return out


class AgentSession:
    """一次文档生成会话。

    典型用法（WebSocket）::

        session = AgentSession()
        for ev in session.start(requirement):
            send(ev)
        for ev in session.handle(client_msg):
            send(ev)
    """

    def __init__(self, session_id: str | None = None) -> None:
        self.session_id = session_id or str(uuid.uuid4())
        self.phase = SessionPhase.CREATED
        self.requirement = ""
        self.doc_type: str | None = None
        self.doc_path: str | None = None
        self.kind: str | None = None
        self.approved_outline = ""
        self.template_text = ""
        self.session_version = 0
        self._outline_feedback = ""
        self._last_outline = ""
        self._graph: Any = None
        self._config: dict[str, Any] | None = None
        self.bridge = UserInputBridge()
        self.error: str | None = None
        self.summary: str | None = None

    # ── 对外入口 ──────────────────────────────────────────

    def start(self, requirement: str) -> Iterator[dict[str, Any]]:
        """开始会话：推断类型 / 请求用户选类型 / 进入大纲或生成。"""
        req = (requirement or "").strip()
        if not req:
            self.phase = SessionPhase.ERROR
            self.error = "需求不能为空"
            yield self._event("error", message=self.error)
            return

        self.requirement = req
        self.doc_type = detect_doc_type(req)
        kind, score = resolve_output_kind(req, self.doc_type)
        yield self._event(
            "session",
            session_id=self.session_id,
            doc_type=self.doc_type,
            inferred_kind=kind,
            inferred_score=score,
        )

        if kind is None:
            self.phase = SessionPhase.AWAITING_KIND
            yield self._event(
                "need_kind",
                options=[
                    {"id": "docx", "label": "Word 文档（报告/说明/方案）"},
                    {"id": "xlsx", "label": "Excel 表格（数据/报表）"},
                    {"id": "pptx", "label": "PowerPoint 演示（汇报/演讲）"},
                ],
            )
            return

        self.kind = kind
        self.doc_path = build_doc_path(req, kind=kind)  # type: ignore[arg-type]
        yield self._event("doc_ready", doc_path=self.doc_path, kind=self.kind)
        yield from self._after_path_ready()

    def handle(self, message: dict[str, Any]) -> Iterator[dict[str, Any]]:
        """处理前端消息，按当前 phase 分发。"""
        msg_type = str(message.get("type") or "").strip()
        if not msg_type:
            yield self._event("error", message="消息缺少 type")
            return

        # 终态
        if self.phase in (SessionPhase.DONE, SessionPhase.CANCELLED, SessionPhase.ERROR):
            yield self._event("error", message=f"会话已结束（{self.phase}）")
            return

        # 忙时控制：任意 RUNNING / AWAITING_INTERRUPT 阶段可接受
        if msg_type in ("supplement", "force", "continue", "quit", "pause"):
            yield from self._handle_busy_control(msg_type, message)
            return

        if msg_type == "choose_kind":
            yield from self._handle_choose_kind(message)
            return
        if msg_type == "outline_decision":
            yield from self._handle_outline_decision(message)
            return
        if msg_type == "official_header":
            yield from self._handle_official_header(message)
            return
        if msg_type == "resume":
            yield from self._handle_resume(message)
            return

        yield self._event("error", message=f"未知或当前阶段不支持的消息类型: {msg_type}")

    # ── 阶段处理 ──────────────────────────────────────────

    def _after_path_ready(self) -> Iterator[dict[str, Any]]:
        assert self.doc_path is not None
        if self.doc_path.lower().endswith(".docx"):
            yield from self._generate_outline()
            return
        yield from self._start_agent()

    def _generate_outline(self) -> Iterator[dict[str, Any]]:
        from office_agent.agent.outline import generate_outline

        self.phase = SessionPhase.AWAITING_OUTLINE
        yield self._event("status", phase=self.phase, message="正在生成文档大纲…")
        try:
            outline = generate_outline(
                self.requirement,
                feedback=self._outline_feedback,
                doc_type=self.doc_type,
            )
        except Exception as e:  # noqa: BLE001
            logger.exception("大纲生成失败")
            outline = f"（生成失败）{e}"
        self._last_outline = outline
        yield self._event("outline", outline=outline)

    def _handle_choose_kind(self, message: dict[str, Any]) -> Iterator[dict[str, Any]]:
        if self.phase != SessionPhase.AWAITING_KIND:
            yield self._event("error", message="当前不在选择文档类型阶段")
            return
        kind = str(message.get("kind") or "").strip().lower()
        mapping = {
            "1": "docx",
            "word": "docx",
            "w": "docx",
            "docx": "docx",
            "2": "xlsx",
            "excel": "xlsx",
            "e": "xlsx",
            "xlsx": "xlsx",
            "3": "pptx",
            "ppt": "pptx",
            "pptx": "pptx",
            "p": "pptx",
        }
        resolved = mapping.get(kind)
        if not resolved:
            yield self._event("error", message="无效文档类型，请传 docx/xlsx/pptx")
            return
        self.kind = resolved
        self.doc_path = build_doc_path(self.requirement, kind=resolved)  # type: ignore[arg-type]
        yield self._event("doc_ready", doc_path=self.doc_path, kind=self.kind)
        yield from self._after_path_ready()

    def _handle_outline_decision(self, message: dict[str, Any]) -> Iterator[dict[str, Any]]:
        if self.phase != SessionPhase.AWAITING_OUTLINE:
            yield self._event("error", message="当前不在大纲确认阶段")
            return
        action = str(message.get("action") or "").strip().lower()
        feedback = str(message.get("feedback") or "").strip()
        # 允许前端覆盖 outline；默认用会话内最近一次生成结果
        outline = str(message.get("outline") or self._last_outline or "")

        if action in ("cancel", "取消", "3"):
            self.phase = SessionPhase.CANCELLED
            yield self._event("cancelled", message="已取消，未生成文档")
            return
        if action in ("revise", "修改", "2"):
            self._outline_feedback = feedback
            yield from self._generate_outline()
            return
        if action in ("approve", "批准", "1", "ok", "yes", "y"):
            if not outline.strip() or outline.startswith("（生成失败）"):
                yield self._event("error", message="大纲无效，请修改重试或取消")
                return
            self.approved_outline = outline
            yield self._event("status", message="大纲已批准")
            if self.doc_type:
                self.phase = SessionPhase.AWAITING_OFFICIAL_HEADER
                fields = official_header_fields(self.doc_type)
                yield self._event(
                    "need_official_header",
                    title=f"《{self.doc_type}》版头信息",
                    description="请填写公文版头/落款信息；必填项勿留空。",
                    fields=fields,
                    doc_type=self.doc_type,
                )
                return
            yield from self._start_agent()
            return
        yield self._event("error", message="action 须为 approve/revise/cancel")

    def _handle_official_header(self, message: dict[str, Any]) -> Iterator[dict[str, Any]]:
        if self.phase != SessionPhase.AWAITING_OFFICIAL_HEADER:
            yield self._event("error", message="当前不在公文版头采集阶段")
            return
        assert self.doc_path is not None and self.doc_type is not None
        answers = message.get("answers") or {}
        if not isinstance(answers, dict):
            yield self._event("error", message="answers 须为对象")
            return
        # 校验必填
        fields = official_header_fields(self.doc_type)
        for f in fields:
            if f.get("required") and not str(answers.get(f["key"], "")).strip():
                yield self._event("error", message=f"必填字段缺失: {f['label']}")
                return
        str_answers = {str(k): str(v) for k, v in answers.items()}
        resolved_type, template_text = merge_official_doc(
            self.doc_type, self.doc_path, str_answers
        )
        if resolved_type is None:
            # 回退普通模式
            self.doc_type = None
            self.template_text = ""
            yield self._event("status", message="公文模板失败，回退普通 Word 生成")
        else:
            self.doc_type = resolved_type
            self.template_text = template_text
            yield self._event("status", message="公文模板已创建")
        yield from self._start_agent()

    def _start_agent(self) -> Iterator[dict[str, Any]]:
        from office_agent.agent.graph import build_graph

        assert self.doc_path is not None
        # 非公文 Word：无版头；公文已在上一步 merge
        set_session_doc(self.doc_path)
        set_bridge(self.bridge)
        self.bridge.set_busy(True)

        self._graph = build_graph(
            self.doc_path,
            doc_type=self.doc_type,
            template_text=self.template_text,
            approved_outline=self.approved_outline,
            vehicle_mode=is_vehicle_related(self.requirement),
        )
        thread_id = self.session_id
        self._config = {
            "configurable": {"thread_id": thread_id},
            "recursion_limit": settings.recursion_limit,
        }

        initial: dict[str, Any] = {
            "messages": [],
            "doc_path": self.doc_path,
        }
        if self.approved_outline:
            first_content = (
                f"{self.requirement}\n\n【已批准的文档大纲】\n{self.approved_outline}"
            )
        else:
            first_content = self.requirement
        first_input = [HumanMessage(content=first_content)]

        self.phase = SessionPhase.RUNNING
        yield self._event(
            "status",
            phase=self.phase,
            message="开始生成（缺关键信息时会询问；可用 supplement/force/continue/quit）",
            doc_path=self.doc_path,
        )
        yield from self._stream(initial=initial, first_input=first_input)
        yield from self._after_stream()

    def _handle_resume(self, message: dict[str, Any]) -> Iterator[dict[str, Any]]:
        if self.phase != SessionPhase.AWAITING_INTERRUPT:
            yield self._event("error", message="当前不在等待用户回答阶段")
            return
        if self._graph is None or self._config is None:
            yield self._event("error", message="会话 graph 未初始化")
            return

        answers = message.get("answers")
        text = message.get("text")
        if answers is not None and isinstance(answers, dict):
            resume_value: dict[str, str] | str = {
                str(k): str(v) for k, v in answers.items()
            }
        elif text is not None:
            resume_value = str(text)
        else:
            yield self._event(
                "error",
                message="resume 需提供 answers(对象) 或 text(字符串)",
            )
            return

        self.phase = SessionPhase.RUNNING
        self.bridge.set_busy(True)
        yield self._event("status", phase=self.phase, message="已收到，继续生成…")
        yield from self._stream(command=Command(resume=resume_value))
        yield from self._after_stream()

    def _handle_busy_control(
        self, msg_type: str, message: dict[str, Any]
    ) -> Iterator[dict[str, Any]]:
        if self.phase not in (
            SessionPhase.RUNNING,
            SessionPhase.AWAITING_INTERRUPT,
        ):
            # quit 在大纲/版头阶段也允许取消
            if msg_type == "quit":
                self.phase = SessionPhase.CANCELLED
                yield self._event("cancelled", message="已退出")
                return
            yield self._event("error", message=f"当前阶段不支持 {msg_type}")
            return

        if msg_type == "quit":
            self.bridge.submit("退出")
            self.phase = SessionPhase.CANCELLED
            yield self._event("cancelled", message="已退出")
            return

        if msg_type == "pause":
            self.bridge.request_soft_pause()
            yield self._event("status", message="已请求暂停；发送 continue / force / supplement 继续")
            return

        if msg_type == "continue":
            self.bridge.submit("继续")
        elif msg_type == "force":
            text = str(message.get("text") or "").strip()
            if not text:
                yield self._event("error", message="force 需要 text")
                return
            self.bridge.submit(f"!{text}")
        elif msg_type == "supplement":
            text = str(message.get("text") or "").strip()
            if not text:
                yield self._event("error", message="supplement 需要 text")
                return
            self.bridge.submit(text)

        # 若正在等待 interrupt，忙时消息只入队，等 resume 后或下一轮边界注入
        if self.phase == SessionPhase.AWAITING_INTERRUPT:
            yield self._event("status", message=f"已记录 {msg_type}，请先完成当前问答或继续")
            return

        # RUNNING：stream 间隙注入后继续
        if self._graph is not None and self._config is not None:
            yield from self._after_stream()

    def _after_stream(self) -> Iterator[dict[str, Any]]:
        """一段 stream 结束后：消化 pending → interrupt → 或完成。"""
        if self.phase == SessionPhase.CANCELLED:
            return
        assert self._graph is not None and self._config is not None

        while True:
            if self.phase == SessionPhase.CANCELLED:
                return

            if self.bridge.consume_quit():
                self.phase = SessionPhase.CANCELLED
                yield self._event("cancelled", message="已退出")
                return

            if self.bridge.consume_soft_pause():
                yield self._event(
                    "status",
                    phase="paused",
                    message="已暂停。发送 continue / supplement / force / quit",
                )
                return

            msgs: list[HumanMessage] = []
            force = self.bridge.consume_force()
            if force:
                msgs.append(HumanMessage(content=f"{PREFIX_FORCE}{force}"))
                yield self._event("status", message=f"强制打断: {force}")
            for text in self.bridge.drain_soft():
                msgs.append(HumanMessage(content=f"{PREFIX_SUPPLEMENT}{text}"))
                yield self._event("status", message=f"用户补充: {text}")
            if self.bridge.consume_continue():
                msgs.append(HumanMessage(content=CONTINUE_PROMPT))
                yield self._event("status", message="继续生成…")

            if msgs:
                with _EXEC_LOCK:
                    set_session_doc(self.doc_path)
                    set_bridge(self.bridge)
                    self._graph.update_state(self._config, {"messages": msgs})
                self.bridge.set_busy(True)
                self.phase = SessionPhase.RUNNING
                yield from self._stream()
                continue

            break

        payload = pending_interrupt(self._graph, self._config)
        if payload is not None:
            self.phase = SessionPhase.AWAITING_INTERRUPT
            self.bridge.set_busy(False)
            yield self._event("interrupt", payload=payload)
            return

        yield from self._emit_final()

    def _stream(
        self,
        *,
        initial: dict[str, Any] | None = None,
        first_input: list | None = None,
        command: Command | None = None,
    ) -> Iterator[dict[str, Any]]:
        assert self._graph is not None and self._config is not None
        if command is not None:
            stream_input: Any = command
        elif initial is not None:
            stream_input = {**initial, "messages": first_input}
        else:
            stream_input = None

        try:
            with _EXEC_LOCK:
                set_session_doc(self.doc_path)
                set_bridge(self.bridge)
                for chunk in self._graph.stream(
                    stream_input, config=self._config, stream_mode="updates"
                ):
                    for _node, updates in chunk.items():
                        if not isinstance(updates, dict):
                            continue
                        for m in updates.get("messages") or []:
                            if isinstance(m, AIMessage):
                                content = str(m.content or "").strip()
                                yield self._event(
                                    "agent_step",
                                    content=content[:500] if content else "",
                                    tool_calls=_tool_call_dicts(m),
                                )
                            elif isinstance(m, ToolMessage):
                                text = str(m.content or "").strip()
                                preview = text[:300] + ("…" if len(text) > 300 else "")
                                yield self._event(
                                    "tool_result",
                                    content=preview,
                                    name=getattr(m, "name", "") or "",
                                )
        except GraphRecursionError:
            logger.warning("达到 recursion_limit=%s", settings.recursion_limit)
            yield self._event(
                "status",
                message=f"Agent 达到步数上限（{settings.recursion_limit}）",
            )
        except Exception as e:  # noqa: BLE001
            if self.phase == SessionPhase.CANCELLED:
                return
            logger.exception("agent stream 异常")
            self.phase = SessionPhase.ERROR
            self.error = str(e)
            yield self._event("error", message=self.error)

    def _emit_final(self) -> Iterator[dict[str, Any]]:
        assert self._graph is not None and self._config is not None
        assert self.doc_path is not None
        snapshot = self._graph.get_state(self._config)
        vals = snapshot.values if snapshot else {}
        done = bool(vals.get("done"))
        summary = str(vals.get("summary") or "")
        self.summary = summary or None
        exists = Path(self.doc_path).exists()

        if done or exists:
            self.phase = SessionPhase.DONE
            yield self._event(
                "done",
                doc_path=self.doc_path,
                summary=summary,
                finished=done,
                download_url=f"/api/v1/sessions/{self.session_id}/download",
            )
            return

        self.phase = SessionPhase.ERROR
        self.error = "未能生成文档"
        yield self._event("error", message=self.error)

    def _event(self, type_: str, **payload: Any) -> dict[str, Any]:
        ev = {"type": type_, "session_id": self.session_id, "phase": str(self.phase)}
        ev.update(payload)
        return ev
