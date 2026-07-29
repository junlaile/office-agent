"""FastAPI 应用：健康检查、会话查询、文档下载、WebSocket 对话。"""

from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Header, HTTPException, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse

from office_agent import config as app_config
from office_agent.api import manager as session_manager
from office_agent.api.stomp_bridge import build_stomp_bridge
from office_agent.api.transport import InProcessWebSocketTransport, MessageEnvelope, TransportAdapter
from office_agent.api.openai_compat import (
    ChatCompletionRequest,
    build_completion_response,
    events_to_assistant_text,
    iter_sse_chunks,
    models_list_response,
    run_turn,
)
from office_agent.config import assert_llm_ready
from office_agent.log import get_logger, setup_logging
from office_agent.officecli import OfficeCLIError, resolve_bin
from office_agent.session.runner import AgentSession, SessionPhase

logger = get_logger(__name__)


def create_app() -> FastAPI:
    setup_logging()
    app = FastAPI(
        title="Office Agent API",
        description=(
            "交互式 Office 文档生成 Agent 的后端 API。"
            "OpenAI 兼容 `/v1/chat/completions` + WebSocket `/api/v1/ws`；"
            "支持大纲/版头/ask_user/finish 确认/忙时补充与文档下载。"
        ),
        version="0.1.0",
    )

    origins = _cors_origins()
    app.add_middleware(
        CORSMiddleware,
        allow_origins=origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.get("/health")
    def health() -> dict[str, Any]:
        officecli_ok = False
        officecli_path = ""
        try:
            officecli_path = str(resolve_bin())
            officecli_ok = True
        except OfficeCLIError:
            pass
        return {
            "status": "ok",
            "llm_configured": bool(app_config.settings.llm_api_key),
            "officecli_ok": officecli_ok,
            "officecli": officecli_path,
            "transport_mode": app_config.settings.transport_mode,
            "session_backend": app_config.settings.session_backend,
        }

    @app.get("/api/v1/sessions/{session_id}")
    def get_session(session_id: str) -> dict[str, Any]:
        session = session_manager.get(session_id)
        if session is None:
            raise HTTPException(status_code=404, detail="session not found")
        return {
            "session_id": session.session_id,
            "phase": str(session.phase),
            "requirement": session.requirement,
            "doc_type": session.doc_type,
            "doc_path": session.doc_path,
            "kind": session.kind,
            "summary": session.summary,
            "error": session.error,
            "download_url": (
                f"/api/v1/sessions/{session.session_id}/download"
                if session.doc_path
                else None
            ),
        }

    @app.get("/api/v1/sessions/{session_id}/download")
    def download(session_id: str) -> FileResponse:
        session = session_manager.get(session_id)
        if session is None:
            raise HTTPException(status_code=404, detail="session not found")
        if not session.doc_path:
            raise HTTPException(status_code=404, detail="document path not set")
        path = Path(session.doc_path)
        # 仅允许下载输出目录内文件，防路径穿越
        try:
            path.resolve().relative_to(app_config.settings.output_dir.resolve())
        except ValueError as e:
            raise HTTPException(status_code=403, detail="forbidden path") from e
        if not path.is_file():
            raise HTTPException(status_code=404, detail="document not found on disk")
        media = _media_type(path.suffix.lower())
        return FileResponse(
            path,
            media_type=media,
            filename=path.name,
        )


    @app.get("/v1/models")
    def list_models() -> dict[str, Any]:
        """OpenAI 兼容：模型列表。"""
        return models_list_response()

    @app.get("/v1/models/{model_id}")
    def get_model(model_id: str) -> dict[str, Any]:
        data = models_list_response()["data"]
        for m in data:
            if m["id"] == model_id:
                return m
        raise HTTPException(status_code=404, detail="model not found")

    @app.post("/v1/chat/completions")
    async def chat_completions(
        body: ChatCompletionRequest,
        x_session_id: str | None = Header(default=None, alias="X-Session-Id"),
        authorization: str | None = Header(default=None),
    ):
        """OpenAI 兼容 Chat Completions（支持 stream）。

        Authorization Bearer 任意非空即可（兼容常见客户端校验）；
        会话续接优先 ``X-Session-Id`` / body.session_id，其次历史消息中的
        ``<!--office-agent-session:UUID-->`` 标记。
        """
        # 兼容客户端：有 Authorization 头时不强制校验 key 内容
        _ = authorization
        try:
            assert_llm_ready()
        except SystemExit as e:
            raise HTTPException(status_code=503, detail=str(e)) from e

        sid = x_session_id or body.session_id

        def _run() -> tuple[AgentSession, list[dict[str, Any]]]:
            return run_turn(messages=body.messages, session_id=sid)

        try:
            session, events = await asyncio.to_thread(_run)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e
        except Exception as e:  # noqa: BLE001
            logger.exception("chat.completions 失败")
            raise HTTPException(status_code=500, detail=str(e)) from e

        if body.stream:
            return StreamingResponse(
                iter_sse_chunks(
                    session=session, events=events, model=body.model
                ),
                media_type="text/event-stream",
                headers={
                    "Cache-Control": "no-cache",
                    "Connection": "keep-alive",
                    "X-Session-Id": session.session_id,
                },
            )

        text = events_to_assistant_text(session, events)
        payload = build_completion_response(
            session=session, text=text, model=body.model
        )
        return JSONResponse(
            payload,
            headers={"X-Session-Id": session.session_id},
        )

    @app.websocket("/api/v1/ws")
    async def websocket_chat(ws: WebSocket) -> None:
        await ws.accept()
        transport: TransportAdapter = InProcessWebSocketTransport(ws)
        stomp_bridge = None
        if app_config.settings.transport_mode == "rabbitmq_stomp":
            try:
                stomp_bridge = build_stomp_bridge()
            except Exception as e:  # noqa: BLE001
                logger.exception("STOMP 初始化失败，回退 inproc")
                await transport.error(f"STOMP 不可用，回退 inproc: {e}")
        session: AgentSession | None = None
        try:
            # 启动前轻量检查（不强制 officecli——集成环境可能后装）
            try:
                assert_llm_ready()
            except SystemExit as e:
                await transport.error(str(e))
                await ws.close()
                return

            while True:
                try:
                    message = await transport.receive()
                except ValueError as e:
                    await transport.error(str(e))
                    continue

                msg_type = str(message.get("type") or "").strip()
                if msg_type == "ping":
                    await transport.send({"type": "pong"})
                    continue

                if session is None:
                    if msg_type != "start":
                        await transport.send(
                            {
                                "type": "error",
                                "message": "请先发送 {type:start, requirement:...}",
                            }
                        )
                        continue
                    requested_sid = str(message.get("session_id") or "").strip()
                    existing = (
                        session_manager.get(requested_sid) if requested_sid else None
                    )
                    if existing is not None:
                        session = existing
                        await transport.send(_session_reconnected_event(session))
                        continue
                    session = AgentSession(session_id=requested_sid or None)
                    session_manager.register(session)
                    session.session_version += 1
                    envelope = MessageEnvelope.from_client_message(
                        session_id=session.session_id,
                        message=message,
                        session_version=session.session_version,
                    )
                    events = await _dispatch_message(
                        session=session,
                        message=message,
                        envelope=envelope,
                        transport=transport,
                        stomp_bridge=stomp_bridge,
                    )
                    for ev in events:
                        await transport.send(ev)
                    await transport.ack(envelope)
                    continue

                # 已有会话
                if msg_type == "start":
                    await transport.send(
                        {
                            "type": "error",
                            "message": "本连接已有会话；请新开 WebSocket 开始新会话",
                        }
                    )
                    continue

                session.session_version += 1
                envelope = MessageEnvelope.from_client_message(
                    session_id=session.session_id,
                    message=message,
                    session_version=session.session_version,
                )
                events = await _dispatch_message(
                    session=session,
                    message=message,
                    envelope=envelope,
                    transport=transport,
                    stomp_bridge=stomp_bridge,
                )
                for ev in events:
                    await transport.send(ev)
                await transport.ack(envelope)

                if session.phase in (
                    SessionPhase.DONE,
                    SessionPhase.CANCELLED,
                    SessionPhase.ERROR,
                ):
                    # 终态后仍保留 session 供下载；连接可继续 ping 或关闭
                    pass

        except WebSocketDisconnect:
            logger.info(
                "WebSocket 断开 session=%s",
                session.session_id if session else None,
            )
        except Exception as e:  # noqa: BLE001
            logger.exception("WebSocket 异常")
            try:
                await transport.error(str(e))
            except Exception:  # noqa: BLE001
                pass
        finally:
            if stomp_bridge is not None:
                stomp_bridge.close()

    return app


async def _dispatch_message(
    *,
    session: AgentSession,
    message: dict[str, Any],
    envelope: MessageEnvelope,
    transport: TransportAdapter,
    stomp_bridge: Any,
) -> list[dict[str, Any]]:
    _ = transport
    if stomp_bridge is None:
        if message.get("type") == "start":
            if session.phase != SessionPhase.CREATED:
                return [_session_reconnected_event(session)]
            req = str(message.get("requirement") or "")

            def _start() -> list[dict[str, Any]]:
                events = list(session.start(req))
                session_manager.register(session)
                return events

            return await asyncio.to_thread(_start)

        def _handle() -> list[dict[str, Any]]:
            events = list(session.handle(message))
            session_manager.register(session)
            return events

        return await asyncio.to_thread(_handle)

    await asyncio.to_thread(stomp_bridge.publish_inbound, envelope)
    await asyncio.to_thread(stomp_bridge.process_next)
    events = await stomp_bridge.consume_outbound(envelope.session_id)
    return events


def _session_reconnected_event(session: AgentSession) -> dict[str, Any]:
    return {
        "type": "session",
        "session_id": session.session_id,
        "phase": str(session.phase),
        "reconnected": True,
        "doc_type": session.doc_type,
        "kind": session.kind,
        "doc_path": session.doc_path,
        "summary": session.summary,
        "error": session.error,
    }


def _cors_origins() -> list[str]:
    import os

    raw = os.environ.get("API_CORS_ORIGINS", "*").strip()
    if not raw or raw == "*":
        return ["*"]
    return [o.strip() for o in raw.split(",") if o.strip()]


def _media_type(suffix: str) -> str:
    return {
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }.get(suffix, "application/octet-stream")


app = create_app()
