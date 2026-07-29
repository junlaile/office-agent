"""``python -m office_agent.api`` 启动 uvicorn。

监听地址/端口/热重载由环境变量控制:
    API_HOST（默认 0.0.0.0）、API_PORT（默认 8000）、
    API_RELOAD（1/true/yes 开启，开发时使用）。
"""

from __future__ import annotations

import os


def main() -> None:
    # uvicorn 延迟导入:避免 CLI 场景 import 本包时强依赖 web 组件
    import uvicorn

    host = os.environ.get("API_HOST", "0.0.0.0")
    port = int(os.environ.get("API_PORT", "8000"))
    uvicorn.run(
        "office_agent.api.app:app",
        host=host,
        port=port,
        reload=os.environ.get("API_RELOAD", "").lower() in ("1", "true", "yes"),
    )


if __name__ == "__main__":
    main()
