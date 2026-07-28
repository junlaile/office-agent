"""FastAPI 应用：健康检查、会话查询、文档下载、WebSocket 对话。"""

from __future__ import annotations

import asyncio
import json
import logging
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from office_agent import config as app_config
from office_agent.api import manager as session_manager
from office_agent.config import assert_llm_ready, setup_logging
from office_agent.officecli import OfficeCLIError, resolve_bin
from office_agent.session.runner import AgentSession, SessionPhase

logger = logging.getLogger("office_agent.api")


def create_app() -> FastAPI:
    setup_logging()
    app = FastAPI(
        title="Office Agent API",
        description=(
            "交互式 Office 文档生成 Agent 的后端 API。"
            "通过 WebSocket `/api/v1/ws` 完成完整对话（大纲/版头/ask_user/"
            "finish 确认/忙时补充），HTTP 提供健康检查与文档下载。"
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

    @app.websocket("/api/v1/ws")
    async def websocket_chat(ws: WebSocket) -> None:
        await ws.accept()
        session: AgentSession | None = None
        try:
            # 启动前轻量检查（不强制 officecli——集成环境可能后装）
            try:
                assert_llm_ready()
            except SystemExit as e:
                await ws.send_json({"type": "error", "message": str(e)})
                await ws.close()
                return

            while True:
                raw = await ws.receive_text()
                try:
                    message = json.loads(raw)
                except json.JSONDecodeError:
                    await ws.send_json({"type": "error", "message": "无效 JSON"})
                    continue
                if not isinstance(message, dict):
                    await ws.send_json({"type": "error", "message": "消息须为 JSON 对象"})
                    continue

                msg_type = str(message.get("type") or "").strip()
                if msg_type == "ping":
                    await ws.send_json({"type": "pong"})
                    continue

                if session is None:
                    if msg_type != "start":
                        await ws.send_json(
                            {
                                "type": "error",
                                "message": "请先发送 {type:start, requirement:...}",
                            }
                        )
                        continue
                    session = AgentSession()
                    session_manager.register(session)
                    req = str(message.get("requirement") or "")
                    sess0 = session

                    def _start(s=sess0, r=req) -> list:
                        return list(s.start(r))

                    events = await asyncio.to_thread(_start)
                    for ev in events:
                        await ws.send_json(ev)
                    continue

                # 已有会话
                if msg_type == "start":
                    await ws.send_json(
                        {
                            "type": "error",
                            "message": "本连接已有会话；请新开 WebSocket 开始新会话",
                        }
                    )
                    continue

                sess = session
                msg = message

                def _handle(s=sess, m=msg) -> list[dict[str, Any]]:
                    return list(s.handle(m))

                events = await asyncio.to_thread(_handle)
                for ev in events:
                    await ws.send_json(ev)

                if sess.phase in (
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
                await ws.send_json({"type": "error", "message": str(e)})
            except Exception:  # noqa: BLE001
                pass

    return app


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
