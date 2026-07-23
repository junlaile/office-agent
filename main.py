"""兼容入口：``python main.py`` → ``office_agent.cli.main``。

推荐用法：``uv run office-agent`` 或 ``uv run python -m office_agent``。
"""

from __future__ import annotations

from office_agent.cli.main import main

if __name__ == "__main__":
    main()
