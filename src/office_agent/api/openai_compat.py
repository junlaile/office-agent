"""OpenAI 兼容 Chat Completions 协议适配。

支持:
  - ``POST /v1/chat/completions``（stream / 非 stream）
  - ``GET /v1/models``

多轮会话通过 assistant 回复中的 ``<!--office-agent-session:UUID-->`` 标记，
或请求头 ``X-Session-Id`` 续接；也可用非标准字段 ``session_id``。
"""

from __future__ import annotations

import json
import re
import time
import uuid
from collections.abc import Iterator
from typing import Any, Literal

from pydantic import BaseModel

from office_agent.api import manager as session_manager
from office_agent.session.runner import AgentSession, SessionPhase

MODEL_ID = "office-agent"
SESSION_MARKER_RE = re.compile(
    r"<!--\s*office-agent-session:([0-9a-fA-F-]{36})\s*-->"
)


# ── OpenAI 请求/响应模型（子集）──────────────────────────


class ChatMessage(BaseModel):
    role: Literal["system", "user", "assistant", "tool", "function"]
    content: str | None = None
    name: str | None = None


class ChatCompletionRequest(BaseModel):
    model: str = MODEL_ID
    messages: list[ChatMessage]
    stream: bool = False
    temperature: float | None = None
    max_tokens: int | None = None
    user: str | None = None
    # 非标准扩展：显式会话 ID
    session_id: str | None = None


# ── 文本 / 事件格式化 ────────────────────────────────────


def session_marker(session_id: str) -> str:
    return f"<!--office-agent-session:{session_id}-->"


def extract_session_id_from_messages(messages: list[ChatMessage]) -> str | None:
    """从历史 assistant 消息中找最近的 session 标记。"""
    for msg in reversed(messages):
        if msg.role != "assistant" or not msg.content:
            continue
        m = SESSION_MARKER_RE.search(msg.content)
        if m:
            return m.group(1)
    return None


def last_user_content(messages: list[ChatMessage]) -> str:
    for msg in reversed(messages):
        if msg.role == "user" and msg.content:
            return msg.content.strip()
    return ""


def format_event(ev: dict[str, Any]) -> str:
    """把内部事件转成可读助手文本片段。"""
    t = ev.get("type")
    if t == "session":
        doc_type = ev.get("doc_type")
        kind = ev.get("inferred_kind")
        parts = ["已创建会话。"]
        if doc_type:
            parts.append(f"识别公文文种：{doc_type}。")
        if kind:
            parts.append(f"推断文档类型：{kind}。")
        return " ".join(parts)
    if t == "need_kind":
        opts = ev.get("options") or []
        lines = ["请选择要生成的文档类型（回复 1/2/3 或 docx/xlsx/pptx）："]
        for o in opts:
            lines.append(f"- {o.get('id')}: {o.get('label')}")
        return "\n".join(lines)
    if t == "doc_ready":
        return f"输出路径：{ev.get('doc_path')}（{ev.get('kind')}）"
    if t == "outline":
        return (
            "【文档大纲预览】\n"
            f"{ev.get('outline')}\n\n"
            "请回复：批准 / 修改 <意见> / 取消"
        )
    if t == "need_official_header":
        fields = ev.get("fields") or []
        lines = [
            f"【{ev.get('title') or '公文版头'}】",
            str(ev.get("description") or ""),
            "请用 JSON 回复，例如：",
            '{"org":"市公安局","signer":"张三"}',
            "字段：",
        ]
        for f in fields:
            star = "*" if f.get("required") else ""
            hint = f" — {f.get('hint')}" if f.get("hint") else ""
            lines.append(f"- {f.get('key')}{star}: {f.get('label')}{hint}")
        return "\n".join(lines)
    if t == "status":
        return str(ev.get("message") or "")
    if t == "agent_step":
        bits: list[str] = []
        content = (ev.get("content") or "").strip()
        if content:
            bits.append(f"[思考] {content[:300]}")
        for tc in ev.get("tool_calls") or []:
            bits.append(f"🔧 {tc.get('display') or tc.get('name')}")
        return "\n".join(bits)
    if t == "tool_result":
        return f"↳ {ev.get('content') or ''}"
    if t == "interrupt":
        payload = ev.get("payload") or {}
        return _format_interrupt(payload)
    if t == "done":
        url = ev.get("download_url") or ""
        summary = ev.get("summary") or ""
        path = ev.get("doc_path") or ""
        lines = ["✓ 文档已生成。"]
        if summary:
            lines.append(f"总结：{summary}")
        if path:
            lines.append(f"路径：{path}")
        if url:
            lines.append(f"下载：{url}")
        return "\n".join(lines)
    if t == "cancelled":
        return str(ev.get("message") or "已取消")
    if t == "error":
        return f"错误：{ev.get('message') or 'unknown'}"
    return ""


def _format_interrupt(payload: dict[str, Any]) -> str:
    if payload.get("type") == "confirm_finish" or payload.get("content_preview"):
        lines = [
            f"【{payload.get('title') or '文档内容确认'}】",
            str(payload.get("description") or ""),
            "—— 内容预览 ——",
            str(payload.get("content_preview") or "（无预览）"),
            "——",
            '请回复 JSON：{"decision":"确认生成"} 或 '
            '{"decision":"需要修改","feedback":"..."}',
        ]
        return "\n".join(lines)

    fields = payload.get("fields") or []
    if fields:
        lines = [
            f"❓ {payload.get('title') or '请填写'}",
            str(payload.get("description") or ""),
            "请用 JSON 按字段 key 回复，例如：",
        ]
        example = {f.get("key", f"f{i}"): "" for i, f in enumerate(fields, 1)}
        lines.append(json.dumps(example, ensure_ascii=False))
        for f in fields:
            star = "*" if f.get("required") else ""
            opts = f.get("options") or []
            opt_s = f" 候选={opts}" if opts else ""
            lines.append(f"- {f.get('key')}{star}: {f.get('label')}{opt_s}")
        return "\n".join(lines)

    q = payload.get("question") or payload.get("title") or "请确认"
    opts = payload.get("options") or []
    lines = [f"❓ {q}"]
    if opts:
        lines.append("可选：")
        for i, o in enumerate(opts, 1):
            lines.append(f"  {i}. {o}")
    return "\n".join(lines)


# ── 用户文本 → 会话消息 ──────────────────────────────────


def _try_json_obj(text: str) -> dict[str, Any] | None:
    raw = text.strip()
    if not raw.startswith("{"):
        # 也支持 ```json ... ```
        m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw, re.DOTALL)
        if not m:
            return None
        raw = m.group(1)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return None
    return data if isinstance(data, dict) else None


def _parse_kv_lines(text: str) -> dict[str, str]:
    answers: dict[str, str] = {}
    for line in text.splitlines():
        if ":" in line:
            k, _, v = line.partition(":")
        elif "：" in line:
            k, _, v = line.partition("：")
        elif "=" in line:
            k, _, v = line.partition("=")
        else:
            continue
        key = k.strip()
        val = v.strip()
        if key and val:
            answers[key] = val
    return answers


def user_text_to_message(session: AgentSession, text: str) -> dict[str, Any]:
    """按当前 phase 把自然语言/JSON 映射为 handle() 消息。"""
    t = (text or "").strip()
    phase = session.phase

    if t in ("退出", "quit", "exit", "q"):
        return {"type": "quit"}
    if t in ("继续", "continue"):
        return {"type": "continue"}
    if t.startswith("!") or t.startswith("强制:") or t.startswith("强制："):
        body = t[1:] if t.startswith("!") else t.split(":", 1)[-1].split("：", 1)[-1]
        return {"type": "force", "text": body.strip()}
    if t in ("暂停", "pause"):
        return {"type": "pause"}

    if phase == SessionPhase.AWAITING_KIND:
        return {"type": "choose_kind", "kind": t}

    if phase == SessionPhase.AWAITING_OUTLINE:
        low = t.lower()
        if low in ("1", "批准", "approve", "y", "yes", "ok") or t.startswith("批准"):
            return {"type": "outline_decision", "action": "approve"}
        if low in ("3", "取消", "cancel", "n", "no", "q") or t.startswith("取消"):
            return {"type": "outline_decision", "action": "cancel"}
        if low in ("2", "修改", "revise") or t.startswith("修改"):
            fb = t
            for prefix in ("修改", "revise", "2"):
                if fb.lower().startswith(prefix):
                    fb = fb[len(prefix) :].lstrip(" :：")
                    break
            return {"type": "outline_decision", "action": "revise", "feedback": fb}
        # 默认当作修改意见
        return {"type": "outline_decision", "action": "revise", "feedback": t}

    if phase == SessionPhase.AWAITING_OFFICIAL_HEADER:
        data = _try_json_obj(t) or _parse_kv_lines(t)
        return {"type": "official_header", "answers": data}

    if phase == SessionPhase.AWAITING_INTERRUPT:
        data = _try_json_obj(t)
        if data is not None:
            return {"type": "resume", "answers": data}
        return {"type": "resume", "text": t}

    # running / paused 等：软补充
    return {"type": "supplement", "text": t}


# ── 驱动一轮 completion ──────────────────────────────────


def run_turn(
    *,
    messages: list[ChatMessage],
    session_id: str | None = None,
) -> tuple[AgentSession, list[dict[str, Any]]]:
    """执行一轮：新建或续接会话，返回 (session, events)。"""
    sid = session_id or extract_session_id_from_messages(messages)
    user_text = last_user_content(messages)
    if not user_text:
        raise ValueError("messages 中缺少 user 内容")

    session: AgentSession | None = session_manager.get(sid) if sid else None

    if session is None:
        session = AgentSession(session_id=sid)
        session_manager.register(session)
        events = list(session.start(user_text))
        return session, events

    # 终态后新开一轮（同一 conversation 再提需求）
    if session.phase in (
        SessionPhase.DONE,
        SessionPhase.CANCELLED,
        SessionPhase.ERROR,
    ):
        session = AgentSession()
        session_manager.register(session)
        events = list(session.start(user_text))
        return session, events

    msg = user_text_to_message(session, user_text)
    events = list(session.handle(msg))
    return session, events


def events_to_assistant_text(session: AgentSession, events: list[dict[str, Any]]) -> str:
    chunks = [format_event(ev) for ev in events]
    body = "\n\n".join(c for c in chunks if c)
    marker = session_marker(session.session_id)
    if body:
        return f"{marker}\n{body}"
    return marker


def build_completion_response(
    *,
    session: AgentSession,
    text: str,
    model: str,
    completion_id: str | None = None,
) -> dict[str, Any]:
    cid = completion_id or f"chatcmpl-{uuid.uuid4().hex[:24]}"
    finish = "stop"
    if session.phase in (
        SessionPhase.AWAITING_KIND,
        SessionPhase.AWAITING_OUTLINE,
        SessionPhase.AWAITING_OFFICIAL_HEADER,
        SessionPhase.AWAITING_INTERRUPT,
    ):
        # 仍在等人：对客户端仍是 stop（下一轮 user 继续）
        finish = "stop"
    return {
        "id": cid,
        "object": "chat.completion",
        "created": int(time.time()),
        "model": model or MODEL_ID,
        "choices": [
            {
                "index": 0,
                "message": {"role": "assistant", "content": text},
                "finish_reason": finish,
            }
        ],
        "usage": {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
        },
        # 非标准扩展，方便原生客户端
        "session_id": session.session_id,
        "phase": str(session.phase),
    }


def iter_sse_chunks(
    *,
    session: AgentSession,
    events: list[dict[str, Any]],
    model: str,
) -> Iterator[str]:
    """产出 OpenAI SSE 行（含最终 [DONE]）。"""
    cid = f"chatcmpl-{uuid.uuid4().hex[:24]}"
    created = int(time.time())

    def chunk(delta: dict[str, Any], finish_reason: str | None = None) -> str:
        payload = {
            "id": cid,
            "object": "chat.completion.chunk",
            "created": created,
            "model": model or MODEL_ID,
            "choices": [
                {
                    "index": 0,
                    "delta": delta,
                    "finish_reason": finish_reason,
                }
            ],
        }
        return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

    yield chunk({"role": "assistant", "content": ""})
    # 先发 session marker，便于客户端续接
    yield chunk({"content": session_marker(session.session_id) + "\n"})

    for ev in events:
        text = format_event(ev)
        if not text:
            continue
        yield chunk({"content": text + "\n\n"})

    yield chunk({}, finish_reason="stop")
    yield "data: [DONE]\n\n"


def models_list_response() -> dict[str, Any]:
    return {
        "object": "list",
        "data": [
            {
                "id": MODEL_ID,
                "object": "model",
                "created": 0,
                "owned_by": "office-agent",
            }
        ],
    }


# 供 OpenAPI / 路由使用的请求体字段说明
class ChatCompletionRequestDoc(ChatCompletionRequest):
    """文档用别名。"""

    pass


__all__ = [
    "MODEL_ID",
    "ChatCompletionRequest",
    "ChatMessage",
    "build_completion_response",
    "events_to_assistant_text",
    "extract_session_id_from_messages",
    "iter_sse_chunks",
    "models_list_response",
    "run_turn",
    "session_marker",
    "user_text_to_message",
]
