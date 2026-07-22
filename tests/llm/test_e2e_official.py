"""LLM 端到端测试：真跑 agent 生成公文（耗 token、需 API key）。

标记 @pytest.mark.llm，默认 skip。
显式运行：``uv run pytest -m llm``

需 .env 配置有效的 LLM_API_KEY / LLM_BASE_URL / LLM_MODEL。
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

from office_agent.config import settings

pytestmark = [pytest.mark.llm, pytest.mark.integration]


@pytest.fixture(scope="module")
def llm_ready():
    """确认 LLM 配置就绪，否则 skip。"""
    if not settings.llm_api_key:
        pytest.skip("未配置 LLM_API_KEY")


def test_generate_notice_e2e(llm_ready, tmp_path):
    """端到端生成通知公文。

    用 subprocess 跑 `python main.py "..."`，避免 LangGraph 在测试进程内的复杂性。
    """
    project_root = Path(__file__).resolve().parents[2]
    req = "写一份关于做好2026年防汛工作的通知，发文机关市公安局，100字以内"
    # 跑 main.py（非交互：argv 传需求）
    result = subprocess.run(
        [sys.executable, "main.py", req],
        cwd=project_root,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=180,
    )
    assert result.returncode == 0, f"main.py 失败:\n{result.stderr[-500:]}"

    # 输出里应有成功标记
    assert "文档已生成" in result.stdout or "✓" in result.stdout

    # 产物文件存在
    output_files = list((project_root / "output").glob("*通知*.docx"))
    assert output_files, "未找到生成的通知文档"

    # 校验最新产物
    latest = max(output_files, key=lambda p: p.stat().st_mtime)
    bin_path = settings.project_root / "bin" / "officecli.exe"
    v = subprocess.run(
        [str(bin_path), "validate", str(latest)],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert "passed" in v.stdout.lower() or "no errors" in v.stdout.lower(), f"校验失败: {v.stdout}"

    # 清理产物
    try:
        latest.unlink()
    except OSError:
        pass
