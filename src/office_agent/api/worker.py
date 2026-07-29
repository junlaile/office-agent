"""``python -m office_agent.api.worker`` 启动独立 SessionWorker。

rabbitmq_stomp 模式下的独立消费进程:轮询 inbound 队列，
调用会话状态机处理消息，事件写回 outbound 队列（Gateway 拉取
后推给 WebSocket 客户端）。可与 API 进程分开部署、水平扩展。
"""

from __future__ import annotations

from office_agent.api.stomp_bridge import build_stomp_bridge
from office_agent.log import setup_logging


def main() -> None:
    """Worker 入口:初始化日志后进入消费循环，退出时断开 broker。"""
    setup_logging()
    bridge = build_stomp_bridge()
    try:
        bridge.run_forever()
    finally:
        bridge.close()


if __name__ == "__main__":
    main()
