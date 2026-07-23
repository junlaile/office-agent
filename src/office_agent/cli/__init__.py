"""终端入口与交互：主循环、UI、忙时输入桥。

入口请用::

    from office_agent.cli.main import main, run
    # 或: uv run office-agent / python -m office_agent

勿在本包 ``__init__`` 中 ``from .main import main``——会覆盖子模块
``cli.main``，导致 ``import office_agent.cli.main`` 拿到函数而非模块。
"""
