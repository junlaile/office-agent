"""``python -m office_agent.api`` 启动 uvicorn。"""

from __future__ import annotations

import os


def main() -> None:
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
