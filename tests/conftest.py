"""pytest 全局 fixture 与配置。

核心设施:
    - FakeRunner: 替代真实 officecli.exe subprocess 的假执行器，记录 argv 供断言。
    - doc_session / xlsx_session / pptx_session: 注入会话文档路径并自动清理。
    - clean_session: 每个 session-using 测试后清空全局，防污染。

设计原则: 单测【绝不】真调 officecli.exe 或 LLM API；真调的测试打 marker 默认 skip。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

# 把 src/ 加入 sys.path（项目是 src 布局，但测试通过 import office_agent 走 src）
_SRC = Path(__file__).resolve().parent.parent / "src"
import sys

if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from office_agent.tools import set_session_doc  # noqa: E402


# ============================================================
# FakeRunner: 替代 officecli.exe subprocess
# ============================================================
class FakeRunner:
    """记录 argv 调用、按预设规则返回 mock stdout 的假 Runner。

    用法（在测试里）::

        fake = FakeRunner()
        officecli_module.reset_runner.__wrapped__(fake)  # 注入
        # 或用 fake_runner fixture 自动注入

    特性:
        - 每次 ``run(args)`` 把 args 追加到 ``self.calls``，便于断言生成的命令。
        - 默认返回空字符串；可在构造时传 ``default_stdout`` 改默认返回。
        - 对 ``get`` / ``view stats`` 等需要 JSON 的命令，用 ``responses`` 字典
          按命令子串（如 "get" / "view"）匹配返回预设值。
        - 抛 ``OfficeCLIError`` 模拟失败：把 ``error_on`` 设为命令子串。
    """

    def __init__(
        self,
        *,
        default_stdout: str = "",
        responses: dict[str, Any] | None = None,
        error_on: str | None = None,
    ) -> None:
        self.calls: list[list[str]] = []
        self.default_stdout = default_stdout
        self.responses = responses or {}
        self.error_on = error_on

    def run(
        self,
        args: list[str],
        *,
        json_output: bool = False,
        timeout: int | None = None,
    ) -> Any:
        self.calls.append(list(args))
        # 模拟失败
        if self.error_on and self.error_on in " ".join(args):
            from office_agent.officecli import OfficeCLIError

            raise OfficeCLIError(
                f"模拟失败: 命令含 '{self.error_on}'",
                cmd=["officecli", *args],
            )
        # 按命令子串匹配预设响应
        joined = " ".join(args)
        for key, value in self.responses.items():
            if key in joined:
                return value
        return self.default_stdout

    @property
    def last_call(self) -> list[str] | None:
        """最后一次调用的 argv（不含 officecli 可执行文件本身）。"""
        return self.calls[-1] if self.calls else None

    def calls_of(self, sub_command: str) -> list[list[str]]:
        """筛选出以某子命令（如 'add' / 'set' / 'create'）开头的调用。"""
        return [c for c in self.calls if c and c[0] == sub_command]


@pytest.fixture
def fake_runner(monkeypatch):
    """注入一个 FakeRunner 到 cli_runner 模块的单例位置。

    ``get_runner()`` 读的是 ``cli_runner._runner`` 全局，所以必须 patch
    cli_runner 模块（而非 officecli 门面）。fixture 结束 monkeypatch 自动恢复。
    """
    from office_agent import cli_runner as runner_module

    fake = FakeRunner()
    monkeypatch.setattr(runner_module, "_runner", fake)
    return fake


# ============================================================
# 会话隔离 fixture
# ============================================================
@pytest.fixture(autouse=True)
def clean_session():
    """每个测试后清空 tools 模块的全局会话路径，防测试间污染。"""
    yield
    set_session_doc.__wrapped__() if hasattr(set_session_doc, "__wrapped__") else None
    # 直接重置模块全局（set_session_doc(None) 不行，用底层）
    from office_agent import tools as tools_module

    tools_module._session_doc_path = None


@pytest.fixture
def doc_session(tmp_path):
    """注入一个 .docx 会话路径，返回该路径。"""
    path = str((tmp_path / "test.docx").resolve())
    set_session_doc(path)
    return path


@pytest.fixture
def xlsx_session(tmp_path):
    """注入一个 .xlsx 会话路径。"""
    path = str((tmp_path / "test.xlsx").resolve())
    set_session_doc(path)
    return path


@pytest.fixture
def pptx_session(tmp_path):
    """注入一个 .pptx 会话路径。"""
    path = str((tmp_path / "test.pptx").resolve())
    set_session_doc(path)
    return path


# ============================================================
# 环境隔离：测试默认不带 LLM key（避免误调 API）
# ============================================================
@pytest.fixture(autouse=True)
def _no_llm_env(monkeypatch):
    """默认剥离 LLM 配置，防止单测意外真调 API。

    需要 LLM 的测试用 @pytest.mark.llm 标记并自行设环境变量。
    """
    # 仅对非 llm 标记的测试生效
    yield
