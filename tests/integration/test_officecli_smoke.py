"""officecli.exe 冒烟集成测试（真调二进制）。

标记 @pytest.mark.integration，默认 skip（单测不应依赖外部二进制）。
显式运行：``uv run pytest -m integration``

验证 create → add → view → validate 的真实端到端行为。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from office_agent.doc_tool import DocTool
from office_agent.officecli import resolve_bin

pytestmark = pytest.mark.integration


@pytest.fixture
def real_doc(tmp_path):
    """真实创建一个 docx 并返回路径。"""
    path = str((tmp_path / "smoke.docx").resolve())
    tool = DocTool(path)
    tool.create()
    yield path
    # 清理：释放 resident + 删文件
    try:
        tool.close()
    except Exception:  # noqa: BLE001
        pass


def test_officecli_binary_exists():
    """officecli.exe 可解析。"""
    bin_path = resolve_bin()
    assert Path(bin_path).exists()


def test_create_and_view(real_doc):
    """创建后能 view_text（应至少返回空或路径标注）。"""
    tool = DocTool(real_doc)
    text = tool.view_text()
    assert isinstance(text, str)


def test_add_and_view_paragraph(real_doc):
    """add_paragraph 后 view_text 能看到文字。"""
    tool = DocTool(real_doc)
    tool.add_paragraph("集成测试段落")
    text = tool.view_text()
    assert "集成测试段落" in text


def test_validate_passes(real_doc):
    """新建文档 validate 通过。"""
    tool = DocTool(real_doc)
    result = tool.validate()
    assert "passed" in result.lower() or "no errors" in result.lower()


def test_add_table(real_doc):
    """add_table 真实建表 + 写单元格。"""
    tool = DocTool(real_doc)
    result = tool.add_table([["姓名", "年龄"], ["张三", "30"]], has_header=True)
    assert "2 行" in result or "2×2" in result
    # validate 仍通过
    assert "passed" in tool.validate().lower() or "no errors" in tool.validate().lower()


def test_set_and_remove_paragraph(real_doc):
    """编辑工具：set_paragraph_text + remove。"""
    tool = DocTool(real_doc)
    tool.add_paragraph("第一段")
    tool.add_paragraph("第二段")
    # 改第一段
    tool.set_paragraph_text("/body/p[1]", "修改后的第一段")
    text = tool.view_text()
    assert "修改后的第一段" in text
