"""``python -m office_agent.api.worker`` 启动独立 SessionWorker。"""

from __future__ import annotations

from office_agent.api.stomp_bridge import build_stomp_bridge
from office_agent.config import setup_logging


def main() -> None:
    setup_logging()
    bridge = build_stomp_bridge()
    try:
        bridge.run_forever()
    finally:
        bridge.close()


if __name__ == "__main__":
    main()
